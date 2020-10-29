Attribute VB_Name = "WorksheetsTools"
Function sheetExistInCurrentWorkbook(ByVal candidateName As String)
    Dim boolSheetExist As Boolean
    
    For Each SheetItem In ActiveWorkbook.Sheets
        'Debug.Print SheetItem.Name
        If InStr(1, SheetItem.Name, candidateName) > 0 Then
            'Debug.Print "Sheet name " + candidateName + " has been found"
            sheetExistInCurrentWorkbook = True
            Exit For
        Else
            sheetExistInCurrentWorkbook = False
        End If
    Next
End Function
'create copy of current sheet (e.g. for backup)
Sub subToolSheetCopy(Optional ByVal sheetNameToCopy As String = "ACT_SHEET", Optional ByVal newSheetName As String = "ADD_CUR_DATE")
    Dim ws1 As Worksheet
    If sheetNameToCopy = "ACT_SHEET" Then
        Set ws1 = ActiveWorkbook.ActiveSheet
    Else
        Set ws1 = ActiveWorkbook.Sheets(sheetNameToCopy)
    End If
    
    If ws1.Visible = xlSheetHidden Then
        ws1.Visible = xlSheetVisible
        ws1.Copy after:=ActiveWorkbook.ActiveSheet
        ws1.Visible = xlSheetHidden
    Else
        ws1.Copy after:=ActiveWorkbook.ActiveSheet
    End If
    
    If newSheetName = "ADD_CUR_DATE" Then
        ActiveSheet.Name = ws1.Name & Date
    Else
        ActiveSheet.Name = newSheetName
    End If
Set ws1 = Nothing
End Sub

Sub AscendingSortOfWorksheets()
disableDisplayOfUpdates (0)
'Sort worksheets in a workbook in ascending order

Dim SCount, i, j As Integer

'For disabling screen updates

'Getting total no. of worsheets in workbook
SCount = Worksheets.Count

'Checking condition whether count of worksheets is greater than 1, If count is one then exit the procedure
If SCount = 1 Then Exit Sub

'Using Bubble sort as sorting algorithm
'Looping through all worksheets
For i = 1 To SCount - 1

    'Making comparison of selected sheet name with other sheets for moving selected sheet to appropriate position
    For j = i + 1 To SCount
        If Worksheets(j).Name < Worksheets(i).Name Then
            Worksheets(j).Move Before:=Worksheets(i)
        End If
    Next j
Next i

ActiveWorkbook.Sheets("srvc_project").Move Before:=Worksheets(1)
ActiveWorkbook.Sheets("Bug reports").Move Before:=Worksheets(1)
ActiveWorkbook.Sheets(CStr(Application.Range("TEST_CASES_SHEET"))).Move Before:=Worksheets(1)

disableDisplayOfUpdates (1)
End Sub

Sub DecendingSortOfWorksheets()
disableDisplayOfUpdates (0)

'Sort worksheets in a workbook in descending order

Dim SCount, i, j As Integer

'For disabling screen updates

'Getting total no. of worsheets in workbook
SCount = Worksheets.Count

'Checking condition whether count of worksheets is greater than 1, If count is one then exit the procedure
If SCount = 1 Then Exit Sub

'Using Bubble sort as sorting algorithm
'Looping through all worksheets
For i = 1 To SCount - 1

    'Making comparison of selected sheet name with other sheets for moving selected sheet to appropriate position
    For j = i + 1 To SCount
        If Worksheets(j).Name > Worksheets(i).Name Then
            Worksheets(j).Move Before:=Worksheets(i)
        End If
    Next j
Next i
ActiveWorkbook.Sheets("srvc_project").Move Before:=Worksheets(1)
ActiveWorkbook.Sheets("Bug reports").Move Before:=Worksheets(1)
ActiveWorkbook.Sheets(CStr(Application.Range("TEST_CASES_SHEET"))).Move Before:=Worksheets(1)



disableDisplayOfUpdates (1)
End Sub
