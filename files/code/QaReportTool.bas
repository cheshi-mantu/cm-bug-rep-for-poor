Attribute VB_Name = "QaReportTool"

Sub getCellValueCreateSheet()

    Dim nameForNewSheet
    Dim NewSheet As Worksheet
    Dim SelectedCell As Range
    Dim CurrentWorksheet As Worksheet


    Set SelectedCell = Application.ActiveCell
    Set CurrentWorksheet = Application.ActiveSheet
    
    If Selection.Column > 1 Then
        Cells(Selection.Row, 1).Select
    End If
    
    nameForNewSheet = Selection.Value
    
    Set NewSheet = ActiveWorkbook.Sheets.Add(Before:=ActiveWorkbook.Worksheets("dict"))
        NewSheet.Paste

    NewSheet.Name = nameForNewSheet
    
    Application.Sheets(CurrentWorksheet.Name).Activate
    
    CurrentWorksheet.Hyperlinks.Add Anchor:=SelectedCell, Address:="", SubAddress:="'" + nameForNewSheet + "'" + "!A1", TextToDisplay:=SelectedCell.Value

Set SelectedCell = Nothing
Set NewSheet = Nothing
Set CurrentWorksheet = Nothing

End Sub

Sub createNewBugReportSheet()
    Dim nameForNewSheet
    Dim SelectedCell As Range
    Dim CurrentWorksheet As Worksheet
    Dim CellToUpdate As Range
    Dim ObjTestId As Range
    Dim CellToGetInfoFrom As Range
    Dim strDateTime As String
    
    Set SelectedCell = Application.ActiveCell
    Set CurrentWorksheet = Application.ActiveSheet

    If Application.Selection.Column > 1 Then
        Set ObjTestId = ActiveWorkbook.ActiveSheet.Cells(Application.Selection.Row, 1)
    ElseIf Application.Selection.Column = 1 Then
        Set ObjTestId = Application.Selection
    End If
    
    ObjTestId.Select

    If Selection.Column > 1 Then
        Cells(Selection.Row, 1).Select
    End If
            
    nameForNewSheet = "BR" + Selection.Value
    
    If Not sheetExistInCurrentWorkbook(nameForNewSheet) Then
        subToolSheetCopy "bug-report template", nameForNewSheet
    End If
        
    
    '-------------
    Application.Sheets(CurrentWorksheet.Name).Activate
    '-------------
        
    Set CellToUpdate = Cells(SelectedCell.Row, getColumnNumber("HDR_COM_LNK", Application.Range("TEST_CASES_SHEET").Text))
    
    CurrentWorksheet.Hyperlinks.Add Anchor:=CellToUpdate, Address:="", SubAddress:="'" + nameForNewSheet + "'" + "!A1", TextToDisplay:="Bug report " + nameForNewSheet
    
    locateAndUpdate nameForNewSheet, "Status", 1, "New"
    
    Set CellToGetInfoFrom = ActiveSheet.Cells(ObjTestId.Row, getColumnNumber("HDR_RESULT", "test cases"))
    locateAndUpdate nameForNewSheet, "Brief description", 1, CellToGetInfoFrom.Text
    locateAndUpdate nameForNewSheet, "Actual Result", 1, CellToGetInfoFrom.Text
    
        
    Set CellToGetInfoFrom = ActiveSheet.Cells(ObjTestId.Row, getColumnNumber("HDR_EXPECT", "test cases"))
    locateAndUpdate nameForNewSheet, "Expected result", 1, CellToGetInfoFrom.Text
    
    Set CellToGetInfoFrom = ActiveSheet.Cells(ObjTestId.Row, getColumnNumber("HDR_ACTION", "test cases"))
    locateAndUpdate nameForNewSheet, "Steps to reproduce", 1, CellToGetInfoFrom.Text
    Set CellToGetInfoFrom = ActiveSheet.Cells(ObjTestId.Row, getColumnNumber("HDR_SITE_PAGE", "test cases"))
    locateAndUpdate nameForNewSheet, "App component", 1, CellToGetInfoFrom.Text
    
    strDateTime = CStr(Date) + " " + CStr(Time())
    
    locateAndUpdate nameForNewSheet, "Bug report creation", 1, strDateTime
    Selection.NumberFormat = "dd/mm/yy h:mm;@"
    
    Sheets(nameForNewSheet).Rows("1:14").EntireRow.AutoFit
    
    
    ActiveWorkbook.Save
    Application.StatusBar = "Saved..."
    
    
    Set SelectedCell = Nothing
    Set NewSheet = Nothing
    Set CurrentWorksheet = Nothing
    Set CellToUpdate = Nothing
    Set ObjTestId = Nothing
    Set CellToGetInfo = Nothing

End Sub

Sub saveBugReportAsPdf()

Dim pathOnly As String

pathOnly = getThisPath
    
    Rows("3:10002").Select
    
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
    End With
    
    Rows("1:14").EntireRow.AutoFit
    Cells(1, 1).Select

With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, Filename:=pathOnly + ActiveSheet.Name, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
     End With
End Sub

Sub sendBugReportPdf()
    'On Error Resume Next
    Dim strHTMLBody As String
    Dim severityText As Range
    Dim strRecipient As String
    Dim strAttachPath As String
    Dim myAttachments As Outlook.Attachments
    Dim strHTMLBodyHeader, strHTMLBodyFooter
     
    
        
    strAttachPath = getThisPath
    If checkFileExists(strAttachPath + ActiveSheet.Name + ".pdf") Then
        strAttachPath = strAttachPath + ActiveSheet.Name + ".pdf"
    ElseIf checkFileExists(strAttachPath + ActiveSheet.Name + ".xlsx") Then
        strAttachPath = strAttachPath + ActiveSheet.Name + ".xlsx"
    Else
        strAttachPath = ""
        MsgBox "Bug report is not saved as separate file, please export the bug report and try again", vbOKOnly, "Attachment is not found"
        Exit Sub
    End If

    
    strHTMLBodyHeader = "<html><body><div style='font-family:Arial Unicode MS;font-size:12pt'>"
    strHTMLBodyFooter = "</div></body></html>"
    
    strRecipient = Application.Range("CLIENT_EMAIL").Text
    
    
    Application.ActiveWorkbook.Sheets("email_templates").Range("BUG_REP_ID").Value = ActiveSheet.Name
    Application.ActiveWorkbook.Sheets("email_templates").Range("ISSUE_BRIEF_DESCRIPTION").Value = locateAndGetValue(ActiveSheet.Name, "Brief description", 1)
    
    
    With ActiveSheet.Range("A1:A100")
        Set ThisPos = .Find(What:="Severity", LookAt:=xlPart, MatchCase:=False, SearchFormat:=False)
    End With
    
    If Not ThisPos Is Nothing Then
        Set severityText = ThisPos.offset(0, 1)
    End If

    
    Application.ActiveWorkbook.Sheets("email_templates").Range("MSG_SUBJECT").Value = "Bug report - Severity " + Left(severityText.Text, 2) + " - " + ActiveSheet.Name + " - for " + Application.Range("APP_NAME")
    
    strHTMLBody = Application.ActiveWorkbook.Sheets("email_templates").Range("HEADER_MSG").Text + Application.ActiveWorkbook.Sheets("email_templates").Range("BODY_MSG").Text + Application.ActiveWorkbook.Sheets("email_templates").Range("FOOTER_MSG").Text
    strHTMLBody = Replace(strHTMLBody, vbLf, "<br/>")
    strHTMLBody = strHTMLBodyHeader + strHTMLBody + strHTMLBodyFooter
    
    
    If checkFileExists(strAttachPath) Then
        'Debug.Print strAttachPath + " exists"
        Set oApp = CreateObject("Outlook.Application")
        Set oMail = oApp.CreateItem(olMailItem)

        oMail.To = strRecipient
        oMail.HTMLBody = strHTMLBody
        oMail.Subject = Application.ActiveWorkbook.Sheets("email_templates").Range("MSG_SUBJECT").Text
        
        Set myAttachments = oMail.Attachments
        myAttachments.Add strAttachPath
        
        oMail.Display
        'oMail.Save
        'oMail.Send
        setStatus ("Sent")
    Else
        MsgBox "There is no such file" + strAttachPath, vbOKOnly, "File not found, bro"
    End If
    
    'clean up
    Set oApp = Nothing
    Set oMail = Nothing
    Set myAttachments = Nothing

End Sub
Sub bugReportsTOC()
Dim strTestCasesSheet As String
Dim strBugRepSheetName As String
Dim bugReportsWS As Worksheet
Dim intLastLine As Integer
    disableDisplayOfUpdates (0)

    strBugRepSheetName = "Bug reports"
    strTestCasesSheet = Application.Range("TEST_CASES_SHEET")
 
    If Not sheetExistInCurrentWorkbook(strBugRepSheetName) Then
        Application.ActiveWorkbook.Sheets.Add Before:=Sheets(strTestCasesSheet)
        ActiveSheet.Name = strBugRepSheetName
    End If
    Set bugReportsWS = Sheets(strBugRepSheetName)
    bugReportsWS.Range("A1:B2000").Clear

    bugReportsWS.Cells(1, 1).Value = "Bug report ID"
    bugReportsWS.Cells(1, 2).Value = "Bug report brief description"

    
    For Each SheetItem In ActiveWorkbook.Sheets
        intLastLine = bugReportsWS.Range("A1048576").End(xlUp).Row
        
        If InStr(SheetItem.Name, "BR") > 0 Then
            'SheetItem.Activate
            'bugReportsWS.Cells(intLastLine + 1, 1).Value = SheetItem.Name
            bugReportsWS.Hyperlinks.Add Anchor:=bugReportsWS.Cells(intLastLine + 1, 1), Address:="", SubAddress:="'" + SheetItem.Name + "'" + "!A1", TextToDisplay:=SheetItem.Name
            bugReportsWS.Cells(intLastLine + 1, 2).Value = locateAndGetValue(SheetItem.Name, "Brief description", 1)
        End If
    Next
    intLastLine = bugReportsWS.Range("A1048576").End(xlUp).Row
    
    
    bugReportsWS.Columns(1).Select
    
    With Selection
        .Columns.AutoFit
    End With
    
    
    bugReportsWS.Columns(2).Select
    
    With Selection
        .ColumnWidth = 100
    End With
    
    
    bugReportsWS.Range("A2:B" + CStr(intLastLine)).Select
            
    bugReportsWS.Sort.SortFields.Clear
    bugReportsWS.Sort.SortFields.Add2 Key:=Range("B" + CStr(intLastLine) _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With bugReportsWS.Sort
        .SetRange Range("A2:B" + CStr(intLastLine))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
            
            
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .AddIndent = False
        .ReadingOrder = xlContext
    End With
    
    applyBordersForSelection
    
    
    bugReportsWS.Range("A1:B1").Select
    ' TODO: hide into a subroutine in helpers
    formatTableHeader
    disableDisplayOfUpdates (1)
End Sub
Sub saveWorkSheetAsSeparateFile()
disableDisplayOfUpdates (0)
Dim strPathToSave As String
Dim ThisWS As Worksheet
    Set ThisWS = ActiveWorkbook.ActiveSheet
Dim NewWB As Workbook
    strPathToSave = getThisPath + ThisWS.Name + ".xlsx"

Set NewWB = ThisWS.Application.Workbooks.Add
    ThisWS.Copy Before:=NewWB.Sheets(1)
    NewWB.SaveAs strPathToSave, xlOpenXMLWorkbook
    NewWB.Close
Set NewWB = Nothing
Set ThisWS = Nothing
disableDisplayOfUpdates (1)
End Sub


Sub exportCurrentWorksheet()
    Dim strExpFrmt As String
    strExpFrmt = Application.Range("BR_EXPORT_FORMAT").Text
    If strExpFrmt = "XLSX" Then
        saveWorkSheetAsSeparateFile
    ElseIf strExpFrmt = "PDF" Then
        saveBugReportAsPdf
    Else
        MsgBox "Export format is not set in the sheet 'srvc_project'", vbOKOnly, "Please complete the settings for this project"
    End If
End Sub


