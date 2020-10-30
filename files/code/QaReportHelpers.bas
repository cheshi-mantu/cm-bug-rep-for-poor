Attribute VB_Name = "QaReportHelpers"
Function getColumnNumber(strInHeaderToSearch As String, worksheetNameToSearch As String)
    Dim SearchForRange As Range
    With ActiveWorkbook.Sheets(worksheetNameToSearch).Range("A1:Z1")
        Set SearchForRange = .Find(What:=strInHeaderToSearch, LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False)
'        Set SearchForRange = .Find(What:=strInHeaderToSearch, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False)
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
        Set SearchForRange = .Find(What:=strSearchString, LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False)
'        Set SearchForRange = .Find(What:=strSearchString, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False)
    End With
    
    If Not SearchForRange Is Nothing Then
        Set StatusUpdateCell = SearchForRange.offset(0, getColumnNumber("HDR_BR_STATUS", "test cases") - 1)
        StatusUpdateCell.Value = strStatus
    Else
        MsgBox "No Id found for " + strSearchString, vbOKOnly, "Something is wrong"
    End If
End Sub
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
Function getThisPath()
Dim FilePath, FileOnly, PathOnly As String
    FilePath = ActiveWorkbook.FullName
    FileOnly = ActiveWorkbook.Name
    PathOnly = Left(FilePath, Len(FilePath) - Len(FileOnly))
    getThisPath = PathOnly
End Function

Sub disableDisplayOfUpdates(EnableDisable As Boolean)
'if EnableDisable = 0, then disable notifications and updates
'if EnableDisable = 1, then enable notifications and updates
If EnableDisable = False Then
    Application.ScreenUpdating = False 'don't update screet in order to speed up the process
    Application.DisplayAlerts = False 'suppress all dialogues application could raise
    Application.Calculation = xlManual 'temporary disable automatic calculation
    Application.StatusBar = "Notifications and calculations are suppressed"
End If
If EnableDisable = True Then
    Application.Calculate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.StatusBar = "Notifications and calculations are enabled"
End If
End Sub



























