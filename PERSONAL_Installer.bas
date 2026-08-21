Attribute VB_Name = "PERSONAL_Installer"
Option Explicit

' =============================================================================
' PERSONAL MACRO PACKAGE INSTALLER
' =============================================================================
' Downloads and installs all modules from GitHub into PERSONAL.XLSB,
' then assigns keyboard shortcuts to each macro.
'
' HOW TO USE:
'   1. Open Excel and press Alt+F11 to open the Visual Basic Editor.
'   2. In the Project Explorer, right-click any module under
'      VBAProject (PERSONAL.XLSB) -> Import File -> select this file.
'   3. Press F5 (or Run -> Run Macro) and select InstallPackage.
'   4. After installation completes, delete this module:
'      right-click PERSONAL_Installer in the Project Explorer -> Remove.
'
' REQUIREMENTS:
'   - PERSONAL.XLSB must exist. If it doesn't, record any macro with
'     "Personal Macro Workbook" as the destination, delete the recorded
'     macro, then try again.
'   - Macro security must allow programmatic access to the VBA project.
'     File -> Options -> Trust Center -> Trust Center Settings ->
'     Macro Settings -> check "Trust access to the VBA project object model".
' =============================================================================

Private Const REPO_RAW_BASE As String = _
    "https://raw.githubusercontent.com/jacob-armin-schroeder/Personal.xlsb/main/"

Private Const README_URL As String = _
    "https://github.com/jacob-armin-schroeder/Personal.xlsb/blob/main/README.md"

' List of modules to install: "Filename|ModuleName"
' ModuleName must match the Attribute VB_Name value inside each file.
Private Const MODULE_LIST As String = _
    "PERSONAL_AutoFill.bas|PERSONAL_AutoFill," & _
    "PERSONAL_CellFormatting.bas|PERSONAL_Borders," & _
    "PERSONAL_FindChanges.bas|PERSONAL_FindChanges," & _
    "PERSONAL_NumberFormats.bas|PERSONAL_NumberFormats," & _
    "PERSONAL_RowColumnSize.bas|PERSONAL_RowColumnSize," & _
    "PERSONAL_ShareCleanup.bas|PERSONAL_ShareCleanup"


Public Sub InstallPackage()

    ' --- Verify VBA project access is enabled ---
    If Not VBAProjectAccessEnabled() Then
        MsgBox "Installer requires access to the VBA project object model." & vbCrLf & vbCrLf & _
               "Go to: File -> Options -> Trust Center -> Trust Center Settings ->" & vbCrLf & _
               "Macro Settings -> check 'Trust access to the VBA project object model'." & vbCrLf & vbCrLf & _
               "Then try again.", vbCritical, "Access Required"
        Exit Sub
    End If

    ' --- Locate PERSONAL.XLSB ---
    Dim personalWB As Workbook
    Set personalWB = GetPersonalWorkbook()
    If personalWB Is Nothing Then
        MsgBox "PERSONAL.XLSB was not found." & vbCrLf & vbCrLf & _
               "To create it: record any macro with 'Personal Macro Workbook' " & _
               "as the destination, then delete the recorded macro and try again.", _
               vbCritical, "PERSONAL.XLSB Not Found"
        Exit Sub
    End If

    ' --- Create temp folder for downloaded files ---
    Dim tempFolder As String
    tempFolder = Environ("TEMP") & "\PersonalMacroInstall\"
    If Dir(tempFolder, vbDirectory) = "" Then MkDir tempFolder

    ' --- Process each module ---
    Dim modules() As String
    modules = Split(MODULE_LIST, ",")

    Dim installedReport As String, skipped As String, failed As String
    installedReport = "": skipped = "": failed = ""

    ' Track installed module names for scoped shortcut assignment
    Dim installedModules() As String
    ReDim installedModules(UBound(modules))
    Dim installedCount As Integer
    installedCount = 0

    Dim i As Integer
    For i = 0 To UBound(modules)
        Dim parts() As String
        parts = Split(modules(i), "|")
        Dim fileName As String, moduleName As String
        fileName = parts(0)
        moduleName = parts(1)

        ' Download file
        Dim localPath As String
        localPath = tempFolder & fileName
        If Not DownloadFile(REPO_RAW_BASE & fileName, localPath) Then
            failed = failed & "  - " & fileName & " (download failed)" & vbCrLf
            GoTo NextModule
        End If

        ' Check for existing module
        If ModuleExists(personalWB, moduleName) Then
            Dim answer As Integer
            answer = MsgBox("Module '" & moduleName & "' already exists in PERSONAL.XLSB." & vbCrLf & vbCrLf & _
                            "Overwrite it?", vbYesNo + vbQuestion, "Module Already Exists")
            If answer = vbNo Then
                skipped = skipped & "  - " & moduleName & vbCrLf
                GoTo NextModule
            End If
            ' Remove existing module before importing
            personalWB.VBProject.VBComponents.Remove _
                personalWB.VBProject.VBComponents(moduleName)
        End If

        ' Import module and record it as installed
        personalWB.VBProject.VBComponents.Import localPath
        installedReport = installedReport & "  - " & moduleName & vbCrLf
        installedModules(installedCount) = moduleName
        installedCount = installedCount + 1

NextModule:
    Next i

    ' --- Optionally assign keyboard shortcuts ---
    Dim assignKeys As Boolean
    assignKeys = False

    If installedCount > 0 Then
        Dim keyAnswer As Integer
        keyAnswer = MsgBox( _
            "Would you like to assign the recommended keyboard shortcuts?" & vbCrLf & vbCrLf & _
            "Click Yes to assign shortcuts automatically." & vbCrLf & _
            "Click No to skip -- you will be responsible for assigning" & vbCrLf & _
            "all shortcuts manually." & vbCrLf & vbCrLf & _
            "The full shortcut list is in the README on GitHub.", _
            vbYesNo + vbQuestion, "Assign Keyboard Shortcuts?")
        If keyAnswer = vbYes Then
            AssignShortcuts installedModules, installedCount
            assignKeys = True
        End If
    End If

    ' --- Clean up temp files ---
    On Error Resume Next
    Dim f As String
    f = Dir(tempFolder & "*.bas")
    Do While f <> ""
        Kill tempFolder & f
        f = Dir()
    Loop
    RmDir tempFolder
    On Error GoTo 0

    ' --- Save PERSONAL.XLSB ---
    personalWB.Save

    ' --- Report results ---
    Dim msg As String
    msg = "Installation complete." & vbCrLf & vbCrLf

    If installedReport <> "" Then msg = msg & "Installed:" & vbCrLf & installedReport & vbCrLf
    If skipped <> "" Then msg = msg & "Skipped (kept existing):" & vbCrLf & skipped & vbCrLf
    If failed <> "" Then msg = msg & "Failed:" & vbCrLf & failed & vbCrLf

    If installedCount > 0 Then
        If assignKeys Then
            msg = msg & "Keyboard shortcuts have been assigned." & vbCrLf & vbCrLf
        Else
            msg = msg & "Keyboard shortcuts were NOT assigned. See the README" & vbCrLf & _
                        "for the full shortcut list." & vbCrLf & vbCrLf
        End If
    End If

    msg = msg & "You can now delete the PERSONAL_Installer module." & vbCrLf & vbCrLf & _
                "Open the README on GitHub?"

    Dim readmeAnswer As Integer
    readmeAnswer = MsgBox(msg, vbYesNo + vbInformation, "Install Complete")
    If readmeAnswer = vbYes Then
        Shell "explorer.exe " & README_URL
    End If

End Sub


Private Sub AssignShortcuts(ByRef installedModules() As String, ByVal installedCount As Integer)
' Assigns keyboard shortcuts only for modules that were actually installed.
' Shortcut key values: lowercase = Ctrl+key, uppercase = Ctrl+Shift+key.
'
' Each entry: "ModuleName|MacroName|ShortcutKey"

    Dim shortcuts As Variant
    shortcuts = Array( _
        "PERSONAL_AutoFill|AutoFill|d", _
        "PERSONAL_Borders|Border_Table_Heading|H", _
        "PERSONAL_Borders|VerticalLines|e", _
        "PERSONAL_Borders|HorizontalLines|E", _
        "PERSONAL_Borders|Border_Outline|O", _
        "PERSONAL_Borders|Clear_Formatting|N", _
        "PERSONAL_Borders|FillBright|B", _
        "PERSONAL_Borders|FillDark|D", _
        "PERSONAL_FindChanges|FindNextChange|m", _
        "PERSONAL_FindChanges|FindPrevChange|M", _
        "PERSONAL_NumberFormats|NumberFormatDecimal|A", _
        "PERSONAL_NumberFormats|NumberFormatPercentage|P", _
        "PERSONAL_NumberFormats|NumberFormatCurrency|C", _
        "PERSONAL_NumberFormats|NumberFormatDateTime|T", _
        "PERSONAL_RowColumnSize|Autofit|W", _
        "PERSONAL_RowColumnSize|ColumnWidthIncrease|q", _
        "PERSONAL_RowColumnSize|ColumnWidthDecrease|Q", _
        "PERSONAL_RowColumnSize|RowHeightIncrease|r", _
        "PERSONAL_RowColumnSize|RowHeightDecrease|R")

    Dim j As Integer
    For j = 0 To UBound(shortcuts)
        Dim entry() As String
        entry = Split(shortcuts(j), "|")
        Dim entryModule As String, macroName As String, shortcutKey As String
        entryModule = entry(0)
        macroName = entry(1)
        shortcutKey = entry(2)

        ' Only assign if this macro's module was installed in this session
        If ModuleWasInstalled(entryModule, installedModules, installedCount) Then
            Application.MacroOptions macro:="PERSONAL.XLSB!" & macroName, _
                                     shortcutKey:=shortcutKey
        End If
    Next j

End Sub


Private Function ModuleWasInstalled(ByVal moduleName As String, _
                                    ByRef installedModules() As String, _
                                    ByVal installedCount As Integer) As Boolean
    Dim k As Integer
    For k = 0 To installedCount - 1
        If installedModules(k) = moduleName Then
            ModuleWasInstalled = True
            Exit Function
        End If
    Next k
    ModuleWasInstalled = False
End Function


Private Function DownloadFile(ByVal Url As String, ByVal localPath As String) As Boolean
    On Error GoTo DownloadFailed
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", Url, False
    http.Send

    If http.status <> 200 Then GoTo DownloadFailed

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 1 ' Binary
    stream.Write http.responseBody
    stream.SaveToFile localPath, 2 ' Overwrite
    stream.Close

    DownloadFile = True
    Exit Function

DownloadFailed:
    DownloadFile = False
End Function


Private Function ModuleExists(Wb As Workbook, ByVal moduleName As String) As Boolean
    On Error Resume Next
    Dim m As Object
    Set m = Wb.VBProject.VBComponents(moduleName)
    ModuleExists = Not (m Is Nothing)
    On Error GoTo 0
End Function


Private Function GetPersonalWorkbook() As Workbook
    Dim Wb As Workbook
    For Each Wb In Application.Workbooks
        If InStr(1, UCase(Wb.Name), "PERSONAL") > 0 And _
           InStr(1, UCase(Wb.Name), ".XLS") > 0 Then
            Set GetPersonalWorkbook = Wb
            Exit Function
        End If
    Next Wb
    Set GetPersonalWorkbook = Nothing
End Function


Private Function VBAProjectAccessEnabled() As Boolean
    On Error Resume Next
    Dim test As Object
    Set test = ThisWorkbook.VBProject.VBComponents
    VBAProjectAccessEnabled = (Err.Number = 0)
    On Error GoTo 0
End Function
