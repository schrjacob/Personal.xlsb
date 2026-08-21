Attribute VB_Name = "PERSONAL_ShareCleanup"
Option Explicit

' ============================================================
'  Cascade Energy - Share Cleanup
'
'  Converts every cell containing a Cascade custom formula into
'  its evaluated value, flags converted cells with a fill color,
'  and logs each converted block to an audit sheet with a
'  representative formula. Operates on a SAVED COPY only; the
'  original workbook on disk is never modified.
'
'  Blocks are grouped by which custom function(s) a cell uses
'  (a "signature"), so:
'    - a uniform filled-down column collapses to one audit row
'      even when its R1C1 text varies row to row;
'    - a column whose formula changes partway down splits into
'      separate audit rows, each with a correct representative.
'
'  Spill (dynamic-array) formulas: only the anchor cell holds the
'  formula, so the WHOLE spill range is captured and written back
'  as static values, and the audit logs the full spilled extent
'  with a flag marking it as spill-derived.
'
'  Does NOT require any Trust Center changes.
' ============================================================

Private Const AUDIT_SHEET_NAME As String = "Formulas Removed"

' Fill color applied to converted cells (RGB).
Private Const FLAG_COLOR_R As Long = 255
Private Const FLAG_COLOR_G As Long = 242
Private Const FLAG_COLOR_B As Long = 204   ' pale amber

' Global structures for Color Conversions - Added 2026.06.13
Type RGBColor
    r As Byte
    g As Byte
    b As Byte
End Type

Type XYZColor
    X As Double
    Y As Double
    Z As Double
End Type

Type LabColor
    L As Double
    A As Double
    b As Double
End Type

' Returns the list of custom function names.
' This is the ONLY thing that changes between overhauls.
' It is populated from an external manifest (see Python tooling).
Private Function CustomFunctionNames() As Variant
    ' --- BEGIN GENERATED LIST ---
    CustomFunctionNames = Array( _
        "acfm_to_scfm", "acfm1_to_acfm2", "AirDensity", "AirLeak", "BilinInterp", _
        "BinMaker", "CalculateCoefficients", "Compressor_kWPerTon", _
        "CompressorCondenser_KWperTon", "CompressorIncrementalEfficiency", _
        "Condenser_kWperTon", "CondenserIncrementalEfficiency", _
        "CoolingTowerCapacityFactor", "DB_DP2WB", "DB_RH2DP", "DB_RH2H", _
        "DB_RH2WB", "DB_WB2H", "DB_WB2RH", "Defrost_false_load_MS", _
        "Defrost_false_load_SS", "Door_Infiltration", "EvaporatorFan_kWperTon", _
        "EvaporatorIncrementalEfficiency", "Fifteenterm", "FindAnomaly", _
        "Inflow_Infiltration", "inHg_to_psia", "kW", "LinInterp", _
        "MotorEfficiency", "NH3_p2hf", "NH3_p2hg", "NH3_p2sf", "NH3_p2sg", _
        "NH3_p2t", "NH3_p2vf", "NH3_p2vg", "NH3_t2p", "Nineterm", "Patm", _
        "Polynomial", "PowerFactor", "Pressure_to_Temp", "Saturation_Pressure", _
        "scfm_to_acfm", "Slide_to_Capacity", "SmallScroll_kWperTon", _
        "SystemIncrementalEfficiency", "Temp_to_Pressure", "Time_interval", _
        "TwentyFiveTerm", "VFDEfficiency", "Vi" _
    )    ' --- END GENERATED LIST ---
End Function

' ============================================================
'  Entry point
' ============================================================

Public Sub ConvertCustomFormulasToValues()
    Dim Wb As Workbook
    Dim ws As Worksheet
    Dim names As Variant
    Dim convertedCount As Long
    Dim resp As VbMsgBoxResult
    Dim auditWs As Worksheet
    Dim auditData As Collection
    Dim newPath As Variant

    ' Saved Application states, restored in both success and failure paths.
    Dim savedScreen As Boolean, savedEvents As Boolean
    Dim savedAlerts As Boolean, savedCalc As XlCalculation

    ' --- Sanity check: a real, non-add-in workbook must be active ---
    ' Runs before everything else so the macro fails clearly rather than
    ' acting on Nothing or on the add-in itself.
    If ActiveWorkbook Is Nothing Then
        MsgBox "No workbook is active. Open the workbook you want to share, " & _
               "make it the active window, then run the macro.", _
               vbExclamation, "Cannot proceed - no workbook"
        Exit Sub
    End If
    If ActiveWorkbook Is ThisWorkbook Then
        MsgBox "The active workbook is the add-in itself. Switch to the data " & _
               "workbook you want to share, then run the macro.", _
               vbExclamation, "Cannot proceed - add-in active"
        Exit Sub
    End If

    resp = MsgBox("This will save a COPY of the workbook, then convert all " & _
                  "Cascade custom-formula cells to values in that copy." & vbCrLf & vbCrLf & _
                  "The conversion cannot be undone. Continue?", _
                  vbExclamation + vbYesNo, "Confirm conversion")
    If resp <> vbYes Then Exit Sub

    ' --- Workbook-structure protection still hard-stops ---
    ' Structure protection blocks adding the log sheet. Unlike sheet
    ' protection (handled below), there is no clean per-sheet save/restore,
    ' so this remains a refuse-and-exit case.
    If ActiveWorkbook.ProtectStructure Then
        MsgBox "This workbook's structure is protected, so the macro cannot " & _
               "add the '" & AUDIT_SHEET_NAME & "' sheet." & vbCrLf & vbCrLf & _
               "No changes have been made." & vbCrLf & vbCrLf & _
               "To proceed: Review tab > Protect Workbook > turn OFF " & _
               "'Protect Workbook Structure' (you may need the password), " & _
               "then run the macro again.", _
               vbExclamation, "Cannot proceed - workbook protected"
        Exit Sub
    End If

    ' --- Sheet protection: detect and HALT, up front, before any file is created ---
    ' The macro cannot write to a protected sheet, and it will NOT attempt to
    ' unprotect sheets itself: a bare Unprotect on a password-protected sheet
    ' makes Excel show its own modal password dialog, which cannot be suppressed
    ' or trapped from VBA. So if ANY worksheet is protected, stop and require the
    ' user to unprotect manually first. No copy is created and nothing is changed.
    Dim protectedSheets As String
    protectedSheets = ProtectedSheetList(ActiveWorkbook)

    If Len(protectedSheets) > 0 Then
        MsgBox "One or more worksheets in this workbook are protected:" & _
               vbCrLf & vbCrLf & protectedSheets & vbCrLf & _
               "The macro cannot convert formulas on a protected sheet, and it " & _
               "will not handle sheet passwords." & vbCrLf & vbCrLf & _
               "No changes have been made." & vbCrLf & vbCrLf & _
               "To proceed: on each listed sheet, Review tab > Unprotect Sheet " & _
               "(enter the password if prompted), then run the macro again.", _
               vbExclamation, "Cannot proceed - sheet(s) protected"
        Exit Sub
    End If

    ' --- Forced Save As before any destructive action ---
    newPath = Application.GetSaveAsFilename( _
                  InitialFileName:=SuggestedCopyName(ActiveWorkbook), _
                  FileFilter:="Excel Workbook (*.xlsx), *.xlsx," & _
                              "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
                  Title:="Save a shareable COPY before converting")
    If newPath = False Then Exit Sub   ' user cancelled

    ' Capture current Application state.
    savedScreen = Application.ScreenUpdating
    savedEvents = Application.EnableEvents
    savedAlerts = Application.DisplayAlerts
    savedCalc = Application.Calculation

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Save under the new name; the saved copy becomes the active workbook,
    ' so the original on disk is untouched.
    ActiveWorkbook.SaveAs fileName:=CStr(newPath), _
                          FileFormat:=FileFormatFromPath(CStr(newPath))
    Set Wb = ActiveWorkbook

    names = CustomFunctionNames()
    convertedCount = 0
    Set auditData = New Collection

    ' By this point the up-front check has guaranteed no worksheet is protected,
    ' so the macro can write to every sheet directly.
    For Each ws In Wb.Worksheets
        If ws.Name <> AUDIT_SHEET_NAME Then
            ResetUsedRange ws
            convertedCount = convertedCount + _
                ProcessSheet(ws, names, auditData)
        End If
    Next ws

    ' Build / refresh the audit sheet from accumulated entries.
    Set auditWs = CreateAuditSheet(Wb)
    WriteAuditSheet auditWs, auditData
    FinalizeAuditSheet auditWs

    Application.Calculation = savedCalc
    Application.DisplayAlerts = savedAlerts
    Application.EnableEvents = savedEvents
    Wb.Save
    Application.ScreenUpdating = savedScreen

    MsgBox convertedCount & " cell(s) converted and flagged, across " & _
           auditData.count & " block(s) logged to '" & AUDIT_SHEET_NAME & "'." & _
           vbCrLf & vbCrLf & "Saved copy: " & Wb.fullName, _
           vbInformation, "Done"
    Exit Sub

CleanFail:
    Application.Calculation = savedCalc
    Application.DisplayAlerts = savedAlerts
    Application.EnableEvents = savedEvents
    Application.ScreenUpdating = savedScreen
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' ============================================================
'  Per-sheet / per-area processing
' ============================================================

Private Function ProcessSheet(ByVal ws As Worksheet, ByVal names As Variant, _
                              ByVal auditData As Collection) As Long
    Dim formulaRng As Range
    Dim area As Range
    Dim count As Long

    ' The up-front check guarantees no sheet is protected, so SpecialCells
    ' failing here just means the sheet has no formula cells. Skip it quietly.
    On Error Resume Next
    Set formulaRng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If formulaRng Is Nothing Then Exit Function

    ' A SpecialCells result may be several non-contiguous areas.
    For Each area In formulaRng.Areas
        count = count + ProcessArea(ws, area, names, auditData)
    Next area

    ProcessSheet = count
End Function

Private Function ProcessArea(ByVal ws As Worksheet, ByVal area As Range, _
                             ByVal names As Variant, _
                             ByVal auditData As Collection) As Long
    Dim fArr As Variant            ' A1-style formulas, bulk-read
    Dim r As Long, c As Long
    Dim nRows As Long, nCols As Long
    Dim firstRowAbs As Long, firstColAbs As Long
    Dim count As Long
    Dim sig As String

    ' groups:     signature -> Range (union of cells sharing that signature)
    ' repFormula: signature -> representative A1 formula (first cell seen)
    Dim groups As Object, repFormula As Object
    Set groups = CreateObject("Scripting.Dictionary")
    Set repFormula = CreateObject("Scripting.Dictionary")

    ' Bulk-read formulas into an in-memory array (one boundary crossing).
    fArr = area.Formula
    firstRowAbs = area.Row
    firstColAbs = area.Column

    If Not IsArray(fArr) Then
        ' Single-cell area: .Formula returns a scalar, not a 2-D array.
        sig = FormulaSignature(CStr(fArr), names)
        If Len(sig) > 0 Then
            AddToGroup groups, repFormula, sig, CStr(fArr), area
            count = 1
        End If
    Else
        nRows = UBound(fArr, 1)
        nCols = UBound(fArr, 2)
        For r = 1 To nRows
            For c = 1 To nCols
                sig = FormulaSignature(CStr(fArr(r, c)), names)
                If Len(sig) > 0 Then
                    AddToGroup groups, repFormula, sig, CStr(fArr(r, c)), _
                               ws.Cells(firstRowAbs + r - 1, firstColAbs + c - 1)
                    count = count + 1
                End If
            Next c
        Next r
    End If

    If count = 0 Then Exit Function

    ' Convert, flag, and log -- per signature group, per contiguous block.
    ' CollapseBlock returns the FULL range actually converted (anchors +
    ' spilled-into cells) and reports whether any spill was expanded.
    Dim sigKey As Variant
    Dim grpRange As Range, blk As Range, convBlk As Range, subBlk As Range
    Dim flagAll As Range          ' accumulates every cell actually converted
    Dim blockHadSpill As Boolean
    For Each sigKey In groups.keys
        Set grpRange = groups(sigKey)

        For Each blk In grpRange.Areas
            blockHadSpill = False
            Set convBlk = CollapseBlock(blk, flagAll, blockHadSpill)
            If Not convBlk Is Nothing Then
                ' Log the full converted extent, split into contiguous
                ' sub-blocks for clean addresses.
                For Each subBlk In convBlk.Areas
                    auditData.Add AuditEntry(ws.Name, _
                                             subBlk.Address(False, False), _
                                             subBlk.Cells.count, _
                                             CStr(repFormula(sigKey)), _
                                             blockHadSpill)
                Next subBlk
            End If
        Next blk
    Next sigKey

    ' Flag everything that was converted, including spilled-into cells.
    If Not flagAll Is Nothing Then
        flagAll.Interior.Color = RGB(FLAG_COLOR_R, FLAG_COLOR_G, FLAG_COLOR_B)
    End If

    ProcessArea = count
End Function

' Adds a cell to the union for its signature, recording the first-seen
' A1 formula as that signature's representative.
Private Sub AddToGroup(ByVal groups As Object, ByVal repFormula As Object, _
                       ByVal sig As String, ByVal a1Formula As String, _
                       ByVal cell As Range)
    If groups.Exists(sig) Then
        Set groups(sig) = Union(groups(sig), cell)
    Else
        groups.Add sig, cell
        repFormula.Add sig, a1Formula
    End If
End Sub

' ============================================================
'  Conversion (handles ordinary cells and spill anchors)
' ============================================================

' Collapses one contiguous block of matched (anchor) cells to static
' values. If a cell is a spill anchor, its ENTIRE spill range is captured
' and written back, so the spilled-into cells survive as values. Returns
' the full range actually converted (anchors + spilled cells), accumulates
' the same into flagAll for later coloring, and sets blockHadSpill True if
' any spill was expanded in this block.
Private Function CollapseBlock(ByVal blk As Range, ByRef flagAll As Range, _
                               ByRef blockHadSpill As Boolean) As Range
    Dim cell As Range
    Dim spillRng As Range
    Dim vals As Variant
    Dim converted As Range          ' full extent converted in THIS block
    Dim hasSpill As Boolean

    For Each cell In blk.Cells
        hasSpill = False
        On Error Resume Next
        hasSpill = cell.hasSpill     ' False/err on Excel without dynamic arrays
        On Error GoTo 0

        If hasSpill Then
            Set spillRng = cell.SpillingToRange    ' full nx2 (etc.) extent
            ' Capture spilled values BEFORE clearing the anchor (clearing
            ' the anchor destroys the spill), then stamp static values back.
            vals = spillRng.Value
            cell.ClearContents
            spillRng.Value = vals
            AccumulateRange flagAll, spillRng
            AccumulateRange converted, spillRng
            blockHadSpill = True
        Else
            cell.Value = cell.Value                ' ordinary single-cell collapse
            AccumulateRange flagAll, cell
            AccumulateRange converted, cell
        End If
    Next cell

    Set CollapseBlock = converted
End Function

Private Sub AccumulateRange(ByRef target As Range, ByVal addition As Range)
    If target Is Nothing Then
        Set target = addition
    Else
        Set target = Union(target, addition)
    End If
End Sub

' ============================================================
'  Formula scanning (single pass: match + signature together)
' ============================================================

' Returns a signature built from the set of custom-function names the
' formula uses, sorted and joined with "+". Returns "" when the formula
' uses no custom function -- so the caller treats "" as "no match" and a
' non-empty result as both "matched" and "its grouping key", in one pass.
Private Function FormulaSignature(ByVal f As String, ByVal names As Variant) As String
    Dim i As Long
    Dim upperF As String
    Dim hits As Collection

    If Len(f) = 0 Then Exit Function

    upperF = UCase$(f)
    Set hits = New Collection
    For i = LBound(names) To UBound(names)
        If ContainsFunctionCall(upperF, UCase$(CStr(names(i)))) Then
            hits.Add UCase$(CStr(names(i)))
        End If
    Next i

    FormulaSignature = JoinSortedCollection(hits)
End Function

Private Function JoinSortedCollection(ByVal c As Collection) As String
    Dim arr() As String
    Dim i As Long, j As Long, tmp As String
    If c.count = 0 Then Exit Function

    ReDim arr(1 To c.count)
    For i = 1 To c.count
        arr(i) = c(i)
    Next i
    ' Tiny set (distinct custom funcs within one formula); simple sort.
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
    JoinSortedCollection = Join(arr, "+")
End Function

' A function call is the name immediately followed by "(", preceded by a
' non-identifier character (or the start of the string). Avoids matching
' a name as a substring of a longer name or inside another token.
Private Function ContainsFunctionCall(ByVal hay As String, _
                                      ByVal needle As String) As Boolean
    Dim pos As Long, startAt As Long
    Dim before As String
    startAt = 1
    Do
        pos = InStr(startAt, hay, needle & "(")
        If pos = 0 Then Exit Do
        If pos = 1 Then
            before = ""
        Else
            before = Mid$(hay, pos - 1, 1)
        End If
        If Not IsIdentifierChar(before) Then
            ContainsFunctionCall = True
            Exit Function
        End If
        startAt = pos + 1
    Loop
End Function

Private Function IsIdentifierChar(ByVal Ch As String) As Boolean
    If Len(Ch) = 0 Then
        IsIdentifierChar = False
    Else
        IsIdentifierChar = (Ch Like "[A-Za-z0-9_.]")
    End If
End Function

' ============================================================
'  Audit sheet
' ============================================================

Private Function AuditEntry(ByVal sheetName As String, ByVal addr As String, _
                            ByVal cellCount As Long, ByVal repFormula As String, _
                            ByVal wasSpill As Boolean) As Variant
    AuditEntry = Array(sheetName, addr, cellCount, repFormula, _
                       IIf(wasSpill, "Yes", ""), Now)
End Function

Private Function CreateAuditSheet(ByVal Wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Wb.Worksheets(AUDIT_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = Wb.Worksheets.Add(After:=Wb.Worksheets(Wb.Worksheets.count))
        ws.Name = AUDIT_SHEET_NAME
    Else
        ws.Cells.Clear
    End If

    ws.Range("A1:F1").Value = Array("Sheet", "Range (copied as values)", _
                                    "Cell Count", _
                                    "Representative Formula (first cell of block)", _
                                    "Spill Range?", "Converted At")
    ws.Range("A1:F1").Font.Bold = True
    Set CreateAuditSheet = ws
End Function

Private Sub WriteAuditSheet(ByVal auditWs As Worksheet, ByVal auditData As Collection)
    Dim n As Long, i As Long
    Dim outArr() As Variant
    Dim entry As Variant

    n = auditData.count
    If n = 0 Then Exit Sub

    ReDim outArr(1 To n, 1 To 6)
    For i = 1 To n
        entry = auditData(i)
        outArr(i, 1) = entry(0)
        outArr(i, 2) = entry(1)
        outArr(i, 3) = entry(2)
        outArr(i, 4) = "'" & entry(3)   ' leading apostrophe: store formula as inert text
        outArr(i, 5) = entry(4)
        outArr(i, 6) = entry(5)
    Next i

    ' Single bulk write of the whole audit body.
    auditWs.Range("A2").Resize(n, 6).Value = outArr
End Sub

Private Sub FinalizeAuditSheet(ByVal auditWs As Worksheet)
    auditWs.Columns("A:F").Autofit
    auditWs.Columns("F").NumberFormat = "yyyy-mm-dd hh:mm"
    ' Cap the formula column so a long formula does not produce an unusable width.
    If auditWs.Columns("D").ColumnWidth > 80 Then
        auditWs.Columns("D").ColumnWidth = 80
    End If
End Sub

' ============================================================
'  Protection scan
' ============================================================

' Returns a newline-separated list of every protected worksheet (Review >
' Protect Sheet), regardless of whether it contains custom formulas. Empty
' string means none are protected. Used both to decide whether to prompt
' the user and to show which sheets are affected.
Private Function ProtectedSheetList(ByVal Wb As Workbook) As String
    Dim ws As Worksheet
    Dim out As String
    For Each ws In Wb.Worksheets
        If ws.Name <> AUDIT_SHEET_NAME Then
            If ws.ProtectContents Then
                out = out & "  - " & ws.Name & vbCrLf
            End If
        End If
    Next ws
    ProtectedSheetList = out
End Function

' ============================================================
'  Used-range reset
' ============================================================

' Referencing UsedRange.Address forces Excel to recompute the true used
' range, discarding phantom extent from stray formatting in distant cells.
' This is the safe (non-destructive) form. If files have severely bloated
' used ranges that this does not fix, a destructive variant that deletes
' trailing empty rows/columns is possible -- and acceptable here because
' the macro only ever operates on a saved copy.
Private Sub ResetUsedRange(ByVal ws As Worksheet)
    Dim dummy As String
    dummy = ws.UsedRange.Address
End Sub

' ============================================================
'  Save-As helpers
' ============================================================

Private Function SuggestedCopyName(ByVal Wb As Workbook) As String
    Dim base As String, dotPos As Long
    base = Wb.Name
    dotPos = InStrRev(base, ".")
    If dotPos > 0 Then base = Left$(base, dotPos - 1)
    SuggestedCopyName = base & "_SHARE_" & Format(Now, "yyyymmdd_hhnn")
End Function

Private Function FileFormatFromPath(ByVal p As String) As XlFileFormat
    If LCase$(Right$(p, 5)) = ".xlsm" Then
        FileFormatFromPath = xlOpenXMLWorkbookMacroEnabled
    Else
        FileFormatFromPath = xlOpenXMLWorkbook
    End If
End Function


Sub FindLeastUsedBackground()
    Dim ws As Worksheet
    Dim cell As Range
    Dim usedColors As Object
    Set usedColors = CreateObject("Scripting.Dictionary")
    
    ' 1. Scan workbook for background colors in use
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        For Each cell In ws.UsedRange
            ' Filter out no fill (-4142) and pure white (16777215)
            If cell.Interior.ColorIndex <> xlNone And cell.Interior.Color <> 16777215 Then
                usedColors(cell.Interior.Color) = True
            End If
        Next cell
    Next ws
    On Error GoTo 0
    
    ' Convert collected workbook colors to CIELAB
    Dim usedLab() As LabColor
    Dim usedCount As Long
    usedCount = usedColors.count
    
    If usedCount > 0 Then
        ReDim usedLab(1 To usedCount)
        Dim k As Long: k = 1
        Dim key As Variant
        Dim rgbVal As RGBColor
        For Each key In usedColors.keys
            rgbVal = LongToRGB(CLng(key))
            usedLab(k) = RGBToLab(rgbVal)
            k = k + 1
        Next key
    End If
    
    ' 2. Generate and test pastel/low-saturation color grid
    Dim bestRGB As RGBColor
    Dim maxMinDeltaE As Double: maxMinDeltaE = -1
    Dim currentCandidate As RGBColor
    Dim candLab As LabColor
    
    Dim rLoop As Long, gLoop As Long, bLoop As Long
    ' Loop using step 15 across soft, bright pastel ranges (high values mean light background)
    For rLoop = 220 To 255 Step 15
        For gLoop = 220 To 255 Step 15
            For bLoop = 220 To 255 Step 15
                
                ' Skip pure white
                If rLoop = 255 And gLoop = 255 And bLoop = 255 Then GoTo NextLoop
                
                currentCandidate.r = CByte(rLoop)
                currentCandidate.g = CByte(gLoop)
                currentCandidate.b = CByte(bLoop)
                
                candLab = RGBToLab(currentCandidate)
                
                ' Evaluate distance against all active workbook colors
                Dim minDeltaE As Double: minDeltaE = 999999
                
                If usedCount = 0 Then
                    ' Default behavior if no background colors exist yet
                    minDeltaE = candLab.L
                Else
                    Dim i As Long
                    For i = 1 To usedCount
                        Dim dE As Double
                        dE = CalculateDeltaE(candLab, usedLab(i))
                        If dE < minDeltaE Then minDeltaE = dE
                    Next i
                End If
                
                ' Maximize the minimum separation
                If minDeltaE > maxMinDeltaE Then
                    maxMinDeltaE = minDeltaE
                    bestRGB = currentCandidate
                End If
NextLoop:
            Next bLoop
        Next gLoop
    Next rLoop
    
    ' 3. Display result and color the active cell as a preview
    Dim hexColor As String
    hexColor = Right("0" & Hex(bestRGB.r), 2) & Right("0" & Hex(bestRGB.g), 2) & Right("0" & Hex(bestRGB.b), 2)
    
    MsgBox "Optimal Background Found!" & vbCrLf & _
           "Red: " & bestRGB.r & vbCrLf & _
           "Green: " & bestRGB.g & vbCrLf & _
           "Blue: " & bestRGB.b & vbCrLf & _
           "Hex: #" & hexColor, vbInformation, "Color Recommendation"
           
    If Not ActiveCell Is Nothing Then
        ActiveCell.Interior.Color = RGB(bestRGB.r, bestRGB.g, bestRGB.b)
    End If
End Sub

' Helper function: Extract components from standard Long Color
Function LongToRGB(ByVal val As Long) As RGBColor
    LongToRGB.r = val Mod 256
    LongToRGB.g = (val \ 256) Mod 256
    LongToRGB.b = (val \ 65536) Mod 256
End Function

' Helper function: Full color conversion path (RGB -> XYZ -> Lab)
Function RGBToLab(rgbVal As RGBColor) As LabColor
    Dim rLinear As Double, gLinear As Double, bLinear As Double
    rLinear = rgbVal.r / 255#
    gLinear = rgbVal.g / 255#
    bLinear = rgbVal.b / 255#
    
    ' Inverse Gamma Correction
    If rLinear > 0.04045 Then rLinear = ((rLinear + 0.055) / 1.055) ^ 2.4 Else rLinear = rLinear / 12.92
    If gLinear > 0.04045 Then gLinear = ((gLinear + 0.055) / 1.055) ^ 2.4 Else gLinear = gLinear / 12.92
    If bLinear > 0.04045 Then bLinear = ((bLinear + 0.055) / 1.055) ^ 2.4 Else bLinear = bLinear / 12.92
    
    rLinear = rLinear * 100
    gLinear = gLinear * 100
    bLinear = bLinear * 100
    
    ' To XYZ space (Assuming standard D65 illuminant)
    Dim xyzVal As XYZColor
    xyzVal.X = rLinear * 0.4124564 + gLinear * 0.3575761 + bLinear * 0.1804375
    xyzVal.Y = rLinear * 0.2126729 + gLinear * 0.7151522 + bLinear * 0.072175
    xyzVal.Z = rLinear * 0.0193339 + gLinear * 0.119192 + bLinear * 0.9503041
    
    ' To CIELAB space
    Dim xr As Double, yr As Double, zr As Double
    xr = xyzVal.X / 95.047
    yr = xyzVal.Y / 100#
    zr = xyzVal.Z / 108.883
    
    Dim fx As Double, fy As Double, fz As Double
    Const epsilon As Double = 0.008856
    Const kappa As Double = 903.3
    
    If xr > epsilon Then fx = xr ^ (1 / 3#) Else fx = (kappa * xr + 16) / 116
    If yr > epsilon Then fy = yr ^ (1 / 3#) Else fy = (kappa * yr + 16) / 116
    If zr > epsilon Then fz = zr ^ (1 / 3#) Else fz = (kappa * zr + 16) / 116
    
    RGBToLab.L = (116 * fy) - 16
    RGBToLab.A = 500 * (fx - fy)
    RGBToLab.b = 200 * (fy - fz)
End Function

' Helper function: Euclidean distance calculation (CIE76 Delta E)
Function CalculateDeltaE(c1 As LabColor, c2 As LabColor) As Double
    CalculateDeltaE = Sqr((c1.L - c2.L) ^ 2 + (c1.A - c2.A) ^ 2 + (c1.b - c2.b) ^ 2)
End Function

