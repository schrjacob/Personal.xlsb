Attribute VB_Name = "PERSONAL_NewWorksheets"
Sub NewWorksheets()
Dim NameCell, NameRange As Range, List As String, i As Integer, HomeSheet As Worksheet
On Error Resume Next
xTitleId = "New Tab Names"
Set NameRange = Application.Selection
Set NameRange = Application.InputBox("Select range with new tab names:", xTitleId, NameRange.Address, Type:=8)

'First define a variable that is the active sheet. This allows the macro to know which
'sheet to return to when it has finished running

    Set HomeSheet = ActiveSheet

    CutCopyMode = False 'Deselects any range that may be copied, in order to prevent macro errors
    Application.ScreenUpdating = False  'Improves macro speed by preventing the screen from updating while the macro runs

'Creating the new worksheets

    i = 0   'Sets up an index variable that will count the number of new worksheets created
    List = NameList.Value   'where "Namelist" is the range of values selected in the UserForm
    Set NameRange = Range(List) 'Defines a new range with the values contained in List
   
        For Each NameCell In NameRange
            
            'If a cell we selected is blank, the macro does nothing for that cell
                
                If NameCell = "" Then
            
            'For all cells that are NOT blank, the macro creates a new worksheet named by the cell contents
                
                Else
                    Sheets.Add After:=Sheets(Sheets.count)      'creates a new worksheet at the END of the workbook
                    Sheets(Sheets.count).Name = NameCell.Value  'renames the new worksheet with the "NameCell" value
                    i = i + 1   'Raises the count of new worksheets created by one
                End If
                
        Next NameCell   'Moves to the next NameCell in NameRange
        
'Reselecting the worskheet we began on, and then re-enabling screen updating
    HomeSheet.Activate
    Application.ScreenUpdating = True

'The before/after screen doesn't change, so a message box gives confirmation that the macro ran correctly.
'The i index also allows it to report the total number of new worksheets created.

    MsgBox (i & " new worksheets were added to the end of the workbook.")

End Sub
