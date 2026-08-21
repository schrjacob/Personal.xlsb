Attribute VB_Name = "PERSONAL_Shortcuts"
Sub TogglePageBreaks()
Attribute TogglePageBreaks.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' TogglePageBreaks Macro
' Keyboard Shortcut: Ctrl+p
' This simple macro toggles the current DisplayPageBreaks status.
    
    ActiveSheet.DisplayPageBreaks = Not (ActiveSheet.DisplayPageBreaks)

End Sub

Sub ToggleGridlines()
Attribute ToggleGridlines.VB_ProcData.VB_Invoke_Func = "I\n14"
'
' TogglePageBreaks Macro
' Keyboard Shortcut: Ctrl+Shift+I
' This simple macro toggles the current DisplayGridlines status.

    ActiveWindow.DisplayGridlines = Not (ActiveWindow.DisplayGridlines)
    
End Sub
