Attribute VB_Name = "PERSONAL_NumberFormats"
Option Explicit

Sub NumberFormatDecimal()
Attribute NumberFormatDecimal.VB_ProcData.VB_Invoke_Func = "A\n14"
<<<<<<< Updated upstream
' Shortcut: Ctrl+Shift+A
=======
' Recommended Shortcut: Ctrl+Shift+A
>>>>>>> Stashed changes
' Cycles: #,##0 > #,##0.0 > #,##0.00 > #,##0.000 > (repeat)
' Applies right alignment and turns off wrap text.
    CycleNumberFormat Array( _
        "#,##0", _
        "#,##0.0", _
        "#,##0.00", _
        "#,##0.000", _
        "#,##0.0000", _
        "#,##0.00000", _
        "#,##0.000000")
End Sub

Sub NumberFormatPercentage()
Attribute NumberFormatPercentage.VB_ProcData.VB_Invoke_Func = "P\n14"
<<<<<<< Updated upstream
' Shortcut: Ctrl+Shift+P
=======
' Recommended Shortcut: Ctrl+Shift+P
>>>>>>> Stashed changes
' Cycles: #,##0% > #,##0.0% > #,##0.00% > #,##0.000% > (repeat)
' Applies right alignment and turns off wrap text.
    CycleNumberFormat Array( _
        "#,##0%", _
        "#,##0.0%", _
        "#,##0.00%", _
        "#,##0.000%")
End Sub

Sub NumberFormatCurrency()
Attribute NumberFormatCurrency.VB_ProcData.VB_Invoke_Func = "C\n14"
<<<<<<< Updated upstream
' Shortcut: (assign as needed)
=======
' Recommended Shortcut: Ctrl+Shift+C
>>>>>>> Stashed changes
' Cycles through simple, red-negative, and accounting variants at 0 and 2 decimal places.
' Applies right alignment and turns off wrap text.
    CycleNumberFormat Array( _
        "$#,##0", _
        "$#,##0.00", _
        "$#,##0_);[Red]($#,##0)", _
        "$#,##0.00_);[Red]($#,##0.00)", _
        "_($* _(#,##0_);_($* (#,##0);_($* ""-""??_);_(@_)", _
        "_($* _(#,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)", _
        "_($* _(#,##0_);[Red]_($* (#,##0);_($* ""-""??_);_(@_)", _
        "_($* _(#,##0.00_);[Red]_($* (#,##0.00);_($* ""-""??_);_(@_)")
End Sub

Sub NumberFormatDateTime()
Attribute NumberFormatDateTime.VB_ProcData.VB_Invoke_Func = "T\n14"
<<<<<<< Updated upstream
' Shortcut: Ctrl+Shift+T
=======
' Recommended Shortcut: Ctrl+Shift+T
>>>>>>> Stashed changes
' Cycles through common date and time formats. Leaves alignment unchanged.
' Turns off wrap text.
    CycleNumberFormat Array( _
        "m/d/yy", _
        "m/d/yyyy", _
        "mm/dd/yyyy", _
        "hh:mm", _
        "hh:mm:ss", _
        "m/d/yy h:mm", _
        "m/d/yyyy hh:mm", _
        "mm/dd/yyyy hh:mm", _
        "yyyy-mm-dd hh:mm", _
        "yyyy-mm-dd hh:mm:ss"), _
        preserveAlignment:=True
        
    Selection.EntireColumn.Autofit
End Sub

Private Sub CycleNumberFormat(ByVal formats As Variant, Optional ByVal preserveAlignment As Boolean = False)
' Matches the active cell's current format against the provided array.
' Applies the next format in the sequence, wrapping back to the first
' when the end of the array is reached or no match is found.
' Unless preserveAlignment is True, right-aligns the selection and turns off wrap text.

    Dim X As Variant
    X = Application.Match(ActiveCell.NumberFormat, formats, False)
    If IsError(X) Then X = 0

    Selection.NumberFormat = formats(X Mod (UBound(formats) + 1))

    Selection.WrapText = False

    If Not preserveAlignment Then
        Selection.HorizontalAlignment = xlRight
    End If

End Sub

