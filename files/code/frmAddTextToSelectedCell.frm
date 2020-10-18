VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddTextToSelectedCell 
   Caption         =   "Add the description"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   12480
   OleObjectBlob   =   "frmAddTextToSelectedCell.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddTextToSelectedCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSaveChanges_Click()
Dim intRowNumber As Integer
Dim strRowNumber As String
    Application.Selection.Value = txtBxPrimitive.Value
    intRowNumber = Application.Selection.Row
    strRowNumber = CStr(intRowNumber)
    strRowNumber = strRowNumber + ":" + strRowNumber
    Rows(strRowNumber).Select
    Selection.RowHeight = 14
    frmAddTextToSelectedCell.Hide
    Unload Me
End Sub

Private Sub txtBxPrimitive_Change()

End Sub

Private Sub UserForm_Activate()
    txtBxPrimitive.Value = Application.Selection.Value
End Sub

