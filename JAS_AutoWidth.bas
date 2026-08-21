<<<<<<< Updated upstream:JAS_AutoWidth.bas
' Created: 4/3/2018 by Jacob Schroeder
' Revised: 4/23/2026 by Jacob Schroeder

Sub AutoWidth()
    With Selection
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
    End With
End Sub

Sub WidthIncrease()
    AdjustColumnWidth 1
End Sub

Sub WidthDecrease()
    AdjustColumnWidth -1
End Sub

Sub HeightIncrease()
    AdjustRowHeight 5
End Sub

Sub HeightDecrease()
    AdjustRowHeight -5
End Sub

Private Sub AdjustColumnWidth(delta As Double)
    Dim w As Double
    w = ActiveCell.ColumnWidth
    Selection.ColumnWidth = Application.Round(Application.Max(w + delta, 1), 0)
End Sub

Private Sub AdjustRowHeight(delta As Double)
    Dim h As Double
    h = ActiveCell.RowHeight
    Selection.RowHeight = Application.Round(Application.Max(h + delta, 5), 0)
End Sub
=======
Attribute VB_Name = "PERSONAL_RowColumnSize"
Option Explicit

Sub Autofit()
Attribute Autofit.VB_ProcData.VB_Invoke_Func = "W\n14"
' Recommended Shortcut: Ctrl+Shift+W
    With Selection
        .EntireColumn.Autofit
        .EntireRow.Autofit
    End With
End Sub

Sub ColumnWidthIncrease()
Attribute ColumnWidthIncrease.VB_ProcData.VB_Invoke_Func = "q\n14"
' Recommended Shortcut: Ctrl+Q
    AdjustColumnWidth 1
End Sub

Sub ColumnWidthDecrease()
Attribute ColumnWidthDecrease.VB_ProcData.VB_Invoke_Func = "Q\n14"
' Recommended Shortcut: Ctrl+Shift+Q
    AdjustColumnWidth -1
End Sub

Sub RowHeightIncrease()
Attribute RowHeightIncrease.VB_ProcData.VB_Invoke_Func = "r\n14"
' Recommended Shortcut: Ctrl+R
    AdjustRowHeight 5
End Sub

Sub RowHeightDecrease()
Attribute RowHeightDecrease.VB_ProcData.VB_Invoke_Func = "R\n14"
' Recommended Shortcut: Ctrl+Shift+R
    AdjustRowHeight -5
End Sub

Private Sub AdjustColumnWidth(delta As Double)
    Dim w As Double
    w = ActiveCell.ColumnWidth
    Selection.ColumnWidth = Application.Round(Application.Max(w + delta, 1), 0)
End Sub

Private Sub AdjustRowHeight(delta As Double)
    Dim h As Double
    h = ActiveCell.RowHeight
    Selection.RowHeight = Application.Round(Application.Max(h + delta, 5), 0)
End Sub
>>>>>>> Stashed changes:PERSONAL_RowColumnSize.bas
