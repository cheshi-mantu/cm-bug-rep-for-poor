Attribute VB_Name = "ViewsHelpers"
Sub applyWorkModeView()
disableDisplayOfUpdates (0)
    Dim strSheetName As String
    strSheetName = Application.Range("TEST_CASES_SHEET").Text
    'Debug.Print strSheetName
    ActiveWorkbook.Sheets(strSheetName).Rows("3:1001").Select
    Selection.RowHeight = 14
    Rows("1:1").EntireRow.Hidden = False
    Cells(1, 1).Select
    disableDisplayOfUpdates (1)
End Sub
Sub applyViewerModeView()
disableDisplayOfUpdates (0)
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
    Rows("1:1").EntireRow.Hidden = True
    Cells(1, 1).Select
disableDisplayOfUpdates (1)
End Sub

