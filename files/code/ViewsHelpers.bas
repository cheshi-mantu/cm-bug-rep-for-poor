Attribute VB_Name = "ViewsHelpers"
Sub applyWorkModeView()
    Dim strSheetName As String
    strSheetName = Application.Range("TEST_CASES_SHEET").Text
    'Debug.Print strSheetName
    ActiveWorkbook.Sheets(strSheetName).Rows("3:1001").Select
    Selection.RowHeight = 14
    Cells(1, 1).Select
End Sub
Sub applyViewerModeView()
    Dim strSheetName As String
    strSheetName = Application.Range("TEST_CASES_SHEET").Text
    ActiveWorkbook.Sheets(strSheetName).Rows("3:1001").Select
    
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Rows.AutoFit
    End With

    Cells(1, 1).Select
    
End Sub

