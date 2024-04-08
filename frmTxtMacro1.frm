VERSION 5.00
Object = "{4016B910-CCE8-4B27-95FA-006C7152BC93}#2.16#0"; "MacabiShared.ocx"
Begin VB.Form frmTxtMacro1 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin MacabiShared.FreeTextTemplateCtrl TxtFreeText 
      Height          =   5175
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9128
   End
End
Attribute VB_Name = "frmTxtMacro1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Initialize(strRemark As String)
6090  On Error GoTo ERR_Initialize



16430         TxtFreeText.InitContent = strRemark





16520        Call TxtFreeText.UpdateRTF

63120      Exit Sub
ERR_Initialize:
63130  MsgBox "Error on line:" & Erl & " in  Initialize" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub strTxt_Change()

End Sub

Private Sub TxtFreeText_DblClick()

End Sub
