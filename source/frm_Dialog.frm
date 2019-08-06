VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Dialog 
   Caption         =   "Случайное замощение"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3915
   OleObjectBlob   =   "frm_Dialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isOk As Boolean

Private Sub UserForm_Initialize()
  isOk = False
End Sub

Private Sub btnOK_Click()
  isOk = True
  Unload Me
End Sub

Private Sub ElementH_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyNumbers KeyAscii
End Sub

Private Sub ElementW_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyNumbers KeyAscii
End Sub

Private Sub RowsNum_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyNumbersInt KeyAscii
End Sub

Private Sub SelNum_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyNumbersInt KeyAscii
End Sub

Private Function OnlyNumbers(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc(",")
    Case Else
      KeyAscii = 0
  End Select
End Function

Private Function OnlyNumbersInt(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Else
      KeyAscii = 0
  End Select
End Function
