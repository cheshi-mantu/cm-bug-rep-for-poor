Attribute VB_Name = "QaReportHelpers"
Function getColumnNumber(strInHeaderToSearch As String, worksheetNameToSearch As String)
    Dim SearchForRange As Range
    With ActiveWorkbook.Sheets(worksheetNameToSearch).Range("A1:Z1")
        Set SearchForRange = .Find(What:=strInHeaderToSearch, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False)
    End With
    If Not SearchForRange Is Nothing Then
        getColumnNumber = SearchForRange.Column
    Else
        getColumnNumber = -1
    End If
End Function
Sub locateAndUpdate(ByVal sheetName As String, parameterToLocate As String, cellOffsetCol As Integer, textForUpdate As String)
    Dim CellToLocate As Range
    Dim CellToUpdate As Range
    Dim RangeToSearch As Range
    
    Set RangeToSearch = Application.ActiveWorkbook.Sheets(sheetName).Range("A1:A20")
    
    With RangeToSearch
        Set CellToLocate = .Find(What:=parameterToLocate, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False)
    End With
        
    If Not CellToLocate Is Nothing Then
        Set CellToUpdate = CellToLocate.offset(0, cellOffsetCol)
    End If
    
    CellToUpdate.Value = textForUpdate
    Set CellToLocate = Nothing
    Set CellToUpdate = Nothing
    Set RangeToSearch = Nothing
    
    
End Sub
Function locateAndGetValue(ByVal sheetName As String, parameterToLocate As String, cellOffsetCol As Integer)
    Dim CellToLocate As Range
    Dim CellToGetFrom As Range
    Dim RangeToSearch As Range
    
    Set RangeToSearch = Application.ActiveWorkbook.Sheets(sheetName).Range("A1:A20")
    
    With RangeToSearch
        Set CellToLocate = .Find(What:=parameterToLocate, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False)
    End With
        
    If Not CellToLocate Is Nothing Then
        Set CellToGetFrom = CellToLocate.offset(0, cellOffsetCol)
    End If
    
    locateAndGetValue = CellToGetFrom.Text
    Set CellToLocate = Nothing
    Set CellToGetFrom = Nothing
    Set RangeToSearch = Nothing
    
End Function

Sub setStatus(strStatus As String)
    Dim strCurrentWorkSheetName As String
    Dim strSearchString As String
    Dim SearchForRange As Range
    Dim StatusUpdateCell As Range
    
    strCurrentWorkSheetName = Application.ActiveSheet.Name
    strSearchString = Right(strCurrentWorkSheetName, Len(strCurrentWorkSheetName) - 3)
    'Debug.Print "'" + strSearchString + "'"
    
    With ActiveWorkbook.Sheets("test cases").Range("A1:A1000")
        Set SearchForRange = .Find(What:=strSearchString, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False)
    End With
    
    If Not SearchForRange Is Nothing Then
        Set StatusUpdateCell = SearchForRange.offset(0, getColumnNumber("Status", "test cases") - 1)
        StatusUpdateCell.Value = strStatus
    Else
        MsgBox "No Id found for " + strSearchString, vbOKOnly, "Something is wrong"
    End If
End Sub
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
Function checkFileExists(fullFilePath As String)
    Dim ObjFS, ObjFile
    Set ObjFS = CreateObject("Scripting.FileSystemObject")
    
    If ObjFS.FileExists(fullFilePath) Then
        checkFileExists = True
    Else
        checkFileExists = False
    End If
End Function
Sub addItemNumbersToCellLines()
Dim strCellContent As String
Dim arrTextLines
Dim intCounter As Integer

intCounter = 0

strCellContent = Application.Selection
    
    arrTextLines = Split(strCellContent, vbLf)
    
    For Each arrItem In arrTextLines
        arrTextLines(intCounter) = CStr(intCounter + 1) + ". " + arrItem
        intCounter = intCounter + 1
    Next
    
    strCellContent = Join(arrTextLines, vbLf)
    Application.Selection.Value = strCellContent

End Sub
Sub runFieldContentEditor()
    frmAddTextToSelectedCell.Show
End Sub
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
Sub applyWorkModeView()
    Dim strSheetName As String
    strSheetName = Application.Range("TEST_CASES_SHEET").Text
    'Debug.Print strSheetName
    ActiveWorkbook.Sheets(strSheetName).Rows("2:1001").Select
    Selection.RowHeight = 14
    Cells(1, 1).Select
End Sub
Sub applyViewerModeView()
    Dim strSheetName As String
    strSheetName = Application.Range("TEST_CASES_SHEET").Text
    ActiveWorkbook.Sheets(strSheetName).Rows("2:1001").Select
    
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

Sub applyBordersForSelection()

    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub
Sub formatTableHeader()
        With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    With Selection.Font
        .Bold = True
        .Name = "Calibri"
        .Size = 12
        .Color = -10477568
        .TintAndShade = 0
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0

    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6299648
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
































