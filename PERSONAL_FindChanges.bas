Attribute VB_Name = "PERSONAL_FindChanges"
Option Explicit

Sub FindNextChange()
Attribute FindNextChange.VB_ProcData.VB_Invoke_Func = "m\n14"
<<<<<<< Updated upstream
' Shortcut: Ctrl+M
=======
' Recommended Shortcut: Ctrl+M
>>>>>>> Stashed changes
' Navigates DOWN the active column to the next cell containing a value
' different from the active cell. Useful for stepping through any column
' where values change in blocks (e.g., True/False flags, category codes,
' numeric indicators, or status fields).
'
' If no different value exists within the used range, lands on the last
' cell evaluated (i.e., the last occupied cell in the column).
    NavigateChange 1
End Sub

Sub FindPrevChange()
Attribute FindPrevChange.VB_ProcData.VB_Invoke_Func = "M\n14"
<<<<<<< Updated upstream
' Shortcut: Ctrl+Shift+M
=======
' Recommended Shortcut: Ctrl+Shift+M
>>>>>>> Stashed changes
' Navigates UP the active column to the previous cell containing a value
' different from the active cell. See FindNextChange for full description.
'
' If no different value exists within the used range, lands on the last
' cell evaluated (i.e., the first occupied cell in the column).
    NavigateChange -1
End Sub


Private Sub NavigateChange(ByVal direction As Integer)
    Application.ScreenUpdating = False

    Dim startCell As Range
    Set startCell = ActiveCell

    Dim startVal As String
    startVal = SafeCellString(startCell)

    Dim firstRow As Long, lastRow As Long
    firstRow = ActiveSheet.UsedRange.Row
    lastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.count - 1

    Dim nextRow As Long
    nextRow = startCell.Row + direction

    Dim lastEvaluated As Range
    Set lastEvaluated = startCell

    Dim targetCell As Range

    Do While nextRow >= firstRow And nextRow <= lastRow
        Set lastEvaluated = ActiveSheet.Cells(nextRow, startCell.Column)

        If SafeCellString(lastEvaluated) <> startVal Then
            Set targetCell = lastEvaluated
            Exit Do
        End If

        nextRow = nextRow + direction
    Loop

    If targetCell Is Nothing Then Set targetCell = lastEvaluated

    Application.ScreenUpdating = True
    targetCell.Activate
    FlashCell targetCell

End Sub


Private Sub FlashCell(ByVal c As Range)
' Briefly highlights the destination cell, then restores original formatting.
' Handles three fill states: no fill, direct RGB color, and theme color.

    Dim wasNoFill As Boolean
    wasNoFill = (c.Interior.Pattern = xlNone)

    Dim wasThemeColor As Boolean
    Dim savedColor As Long
    Dim savedThemeColor As Long
    Dim savedTintShade As Double

    If Not wasNoFill Then
        If c.Interior.ThemeColor = 0 Then
            wasThemeColor = False
            savedColor = c.Interior.Color
        Else
            wasThemeColor = True
            savedThemeColor = c.Interior.ThemeColor
            savedTintShade = c.Interior.TintAndShade
        End If
    End If

    c.Interior.Color = RGB(255, 210, 40)

    Dim startTime As Double
    startTime = Timer
<<<<<<< Updated upstream
    Do While Timer >= startTime And Timer < startTime + 0.3
=======
    Do While (Timer - startTime) < 0.3 And (Timer - startTime) >= 0
>>>>>>> Stashed changes
        DoEvents
    Loop

    If wasNoFill Then
        c.Interior.Pattern = xlNone
    ElseIf wasThemeColor Then
        c.Interior.ThemeColor = savedThemeColor
        c.Interior.TintAndShade = savedTintShade
    Else
        c.Interior.Color = savedColor
    End If

End Sub


Private Function SafeCellString(ByVal c As Range) As String
' Safely converts any cell value — including errors — to a string for comparison.
' Error values are prefixed with "ERR:" so distinct error types are treated as distinct.
    If IsError(c.Value) Then
        SafeCellString = "ERR:" & CStr(c.Value)
    Else
        SafeCellString = CStr(c.Value)
    End If
End Function

