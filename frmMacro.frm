VERSION 5.00
Object = "{6D226E96-1587-48EA-9797-5080D13F7CE2}#364.17#0"; "MacroHis.ocx"
Begin VB.Form frmMacro 
   Caption         =   "Form1"
   ClientHeight    =   10560
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   18915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10560
   ScaleWidth      =   18915
   StartUpPosition =   3  'Windows Default
   Begin MacroHis.MacroHisCtrl MacroHisCtrl 
      Height          =   10095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18615
      _ExtentX        =   32835
      _ExtentY        =   17806
   End
End
Attribute VB_Name = "frmMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Activate()
          'MsgBox MacroHisCtrl.Width & " current= " & Me.Width
6180      Me.Width = MacroHisCtrl.Width + 100
           'MsgBox MacroHisCtrl.Height & " current= " & Me.Height
6190      Me.Height = MacroHisCtrl.Height + 450
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
6200      Cancel = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
6160      Cancel = 0
End Sub

Private Sub MacroHisCtrl_CloseClicked()
6170      Me.Hide
End Sub



