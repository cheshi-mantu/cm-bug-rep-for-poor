Attribute VB_Name = "WorksheetsTools"
Sub AscendingSortOfWorksheets()

'Sort worksheets in a workbook in ascending order

Dim SCount, i, j As Integer

'For disabling screen updates
Application.ScreenUpdating = False

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


End Sub

Sub DecendingSortOfWorksheets()

'Sort worksheets in a workbook in descending order

Dim SCount, i, j As Integer

'For disabling screen updates
Application.ScreenUpdating = False

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


End Sub
