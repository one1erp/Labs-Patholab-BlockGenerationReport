VERSION 5.00
Begin VB.Form frmExtraRequest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "בקשה חוזרת"
   ClientHeight    =   3504
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8772
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3504
   ScaleWidth      =   8772
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtReembeddingReason 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox x 
      Alignment       =   1  'Right Justify
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   6615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "הערות"
      Height          =   255
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "סיבה להעמדה חוזרת"
      Height          =   375
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmExtraRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Initialize(strReembeddingReason As String, strRemark As String)
6090  On Error GoTo ERR_Initialize

6100      txtReembeddingReason.Text = strReembeddingReason
6110      txtRemark.Text = strRemark

6120      Exit Sub
ERR_Initialize:
6130  MsgBox "Error on line:" & Erl & " in  Initialize" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub Form_Click()
6140  On Error GoTo ERR_Form_Click
6150      Call Me.Hide
          
6160      Exit Sub
ERR_Form_Click:
6170  MsgBox "Error on line:" & Erl & " in  Form_Click" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


