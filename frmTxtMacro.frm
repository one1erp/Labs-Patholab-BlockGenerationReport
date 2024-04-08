VERSION 5.00
Object = "{4016B910-CCE8-4B27-95FA-006C7152BC93}#2.16#0"; "MacabiShared.ocx"
Begin VB.Form frmTxtMacro 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin MacabiShared.FreeTextTemplateCtrl TxtFreeText 
      Height          =   4812
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9612
      _ExtentX        =   16960
      _ExtentY        =   8493
   End
End
Attribute VB_Name = "frmTxtMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Initialize(pFld As String)
6090  On Error GoTo ERR_Initialize


 'If Not RtfResult.EOF Then
                  'TxtFreeText.TextRTF = ReadClob(RtfResult("RTF_TEXT"))
17790             TxtFreeText.InitContent = pFld ' ReadClob(rtf) 'RtfResult("RTF_TEXT"))
17800             TxtFreeText.FontName = "Arial"
17810         '        TxtFreeText.rightMargin = nInch * 6
17820         '    TxtFreeText.Left = 120
17830         '    TxtFreeText.Width = MAX_WIDTH

17840

17880             Call TxtFreeText.UpdateRTF
17890



6120      Exit Sub
ERR_Initialize:
6130  MsgBox "Error on line:" & Erl & " in  Initialize" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub FreeTextTemplateCtrl1_DblClick()

End Sub

