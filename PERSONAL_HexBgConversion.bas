Attribute VB_Name = "PERSONAL_HexBgConversion"
Sub HexToBackground()
    Dim cell As Range
    Dim hexStr As String
    Dim rVal As Long, gVal As Long, bVal As Long
    
    ' Check if cells are selected
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    ' Optimize macro execution speed
    Application.ScreenUpdating = False
    
    For Each cell In Selection
        ' Remove spaces and the # symbol if present
        hexStr = Replace(Trim(cell.Value), "#", "")
        
        ' Ensure the hex string is exactly 6 characters long
        If Len(hexStr) = 6 Then
            On Error Resume Next
            ' Convert 2-character hex segments into decimal RGB values
            rVal = CLng("&H" & Mid(hexStr, 1, 2))
            gVal = CLng("&H" & Mid(hexStr, 3, 2))
            bVal = CLng("&H" & Mid(hexStr, 5, 2))
            
            ' Apply the color to the cell background
            If Err.Number = 0 Then
                cell.Interior.Color = RGB(rVal, gVal, bVal)
            End If
            On Error GoTo 0
        End If
    Next cell
    
    Application.ScreenUpdating = True
End Sub

Sub BackgroundToHex()
    Dim cell As Range
    Dim colorLong As Long
    Dim rVal As Long, gVal As Long, bVal As Long
    Dim hexStr As String
    Dim hasContent As Boolean
    Dim response As VbMsgBoxResult
    
    ' Check if cells are selected
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    ' First pass: check whether any cell in the selection already has content,
    ' including formulas that evaluate to an empty string
    hasContent = False
    For Each cell In Selection
        If Len(Trim(cell.Value)) > 0 Or cell.HasFormula Then
            hasContent = True
            Exit For
        End If
    Next cell
    
    ' Warn the user before overwriting existing values/formulas
    If hasContent Then
        response = MsgBox("The selected range contains existing values or formulas." & vbCrLf & _
                           "Running this macro will overwrite them with hex color codes." & vbCrLf & vbCrLf & _
                           "Continue?", vbYesNo + vbExclamation, "Confirm Overwrite")
        If response = vbNo Then Exit Sub
    End If
    
    ' Optimize macro execution speed
    Application.ScreenUpdating = False
    
    For Each cell In Selection
        ' Skip cells with no fill applied
        If cell.Interior.Pattern <> xlNone Then
            colorLong = cell.Interior.Color
            
            ' Decompose the Long into R, G, B components
            rVal = colorLong Mod 256
            gVal = (colorLong \ 256) Mod 256
            bVal = (colorLong \ 65536) Mod 256
            
            ' Build the hex string, padding each byte to 2 digits
            hexStr = Right("0" & Hex(rVal), 2) & _
                      Right("0" & Hex(gVal), 2) & _
                      Right("0" & Hex(bVal), 2)
            
            cell.Value = "#" & hexStr
        End If
    Next cell
    
    Application.ScreenUpdating = True
End Sub

