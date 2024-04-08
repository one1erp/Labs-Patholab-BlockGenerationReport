VERSION 5.00
Object = "{30CB9C1A-EE46-4D4C-BBDE-1D306015D2DD}#47.8#0"; "RequestRemark.ocx"
Object = "{E40B1134-8362-494C-99D9-AB6AD0E21EB5}#6.22#0"; "Organ.ocx"
Begin VB.UserControl BlockCtrl 
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   ScaleHeight     =   6996
   ScaleMode       =   0  'User
   ScaleWidth      =   11796
   Begin VB.CommandButton cmdMacrotxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6480
      Picture         =   "BlockGenerationReportCtrl.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   5940
      Width           =   1092
   End
   Begin VB.OptionButton K_option 
      Caption         =   "קבלנות"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   120
      Width           =   1035
   End
   Begin VB.OptionButton R_option 
      Caption         =   "רוטינה"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmdMacro 
      Caption         =   "מאקרו"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5040
      Picture         =   "BlockGenerationReportCtrl.ctx":0442
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "לחץ לפתיחת מאקרו(MACRO)"
      Top             =   5940
      Width           =   1335
   End
   Begin Organ.OrganCtrl OrganCtrl 
      Height          =   372
      Left            =   7680
      TabIndex        =   40
      Top             =   6024
      Visible         =   0   'False
      Width           =   1092
      _ExtentX        =   1931
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdExtraRequest 
      BackColor       =   &H000000FF&
      Caption         =   "בקשה חוזרת: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5899
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Frame FrameDetails 
      Height          =   2688
      Left            =   0
      TabIndex        =   14
      Top             =   3240
      Width           =   11652
      Begin VB.CommandButton OKButton2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   624
         Left            =   10440
         Picture         =   "BlockGenerationReportCtrl.ctx":0884
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "לחץ לאישור"
         Top             =   120
         Width           =   1212
      End
      Begin VB.PictureBox Picture1 
         Height          =   1800
         Left            =   0
         ScaleHeight     =   1740
         ScaleWidth      =   11430
         TabIndex        =   19
         Top             =   720
         Width           =   11484
         Begin VB.VScrollBar VScroll 
            Height          =   1800
            Left            =   11160
            TabIndex        =   25
            Top             =   0
            Visible         =   0   'False
            Width           =   384
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   500
            Left            =   0
            ScaleHeight     =   495
            ScaleWidth      =   14175
            TabIndex        =   20
            Top             =   120
            Width           =   14175
            Begin VB.ComboBox CmbReason 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   336
               Index           =   0
               Left            =   5280
               RightToLeft     =   -1  'True
               Sorted          =   -1  'True
               TabIndex        =   21
               Top             =   50
               Visible         =   0   'False
               Width           =   3972
            End
            Begin VB.Label LblCreatedOn 
               Alignment       =   2  'Center
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   0
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   -72
               Visible         =   0   'False
               Width           =   2472
            End
            Begin VB.Label LblRowNum 
               Alignment       =   2  'Center
               BackColor       =   &H80000018&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   348
               Index           =   0
               Left            =   9360
               TabIndex        =   23
               Top             =   70
               Visible         =   0   'False
               Width           =   972
            End
            Begin VB.Label LblCreatedBy 
               Alignment       =   2  'Center
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   0
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   -72
               Visible         =   0   'False
               Width           =   2472
            End
         End
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "בוצע ע""י \ בתאריך"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2124
         TabIndex        =   18
         Top             =   240
         Width           =   2088
      End
      Begin VB.Label LblReasonTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "סיבה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8640
         TabIndex        =   17
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "מס' שיקוע"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9120
         TabIndex        =   16
         Top             =   240
         Width           =   1248
      End
   End
   Begin VB.CommandButton CmdCancel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   9000
      Picture         =   "BlockGenerationReportCtrl.ctx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "לחץ לביטול"
      Top             =   5959
      Width           =   495
   End
   Begin VB.Frame FrameRemark 
      Height          =   1400
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   11532
      Begin VB.ComboBox CmbAliquotRemark 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   10.5
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   10
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   456
         Width           =   4392
      End
      Begin RequestRemark.RequestRemarkCtrl RequestRemarkCtrl 
         Height          =   492
         Left            =   3600
         TabIndex        =   1
         Top             =   816
         Visible         =   0   'False
         Width           =   852
         _ExtentX        =   1508
         _ExtentY        =   873
      End
      Begin VB.Label lblPathologMacro 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   480
         Width           =   2412
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "פתולוג המאקרו"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   204
         Left            =   5640
         TabIndex        =   37
         Top             =   204
         Width           =   1452
      End
      Begin VB.Label lblPriorityTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "דחיפות"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   204
         Left            =   10608
         TabIndex        =   36
         Top             =   204
         Width           =   792
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   744
         Left            =   10200
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   456
         Width           =   1152
      End
      Begin VB.Label LblNumOfPiecesTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "קטעי רקמה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   204
         Left            =   7440
         TabIndex        =   33
         Top             =   204
         Width           =   1260
      End
      Begin VB.Label LblNumOfPieces 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label lblSettingUp 
         Caption         =   "לא"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   372
         Left            =   9480
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   492
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "העמדה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   204
         Left            =   9240
         TabIndex        =   15
         Top             =   204
         Width           =   780
      End
      Begin VB.Label LblAliquotRemarkTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "הערה לבלוק"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   204
         Left            =   2900
         TabIndex        =   9
         Top             =   204
         Width           =   1380
      End
   End
   Begin VB.Frame FrameHeader 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   400
      Width           =   11628
      Begin VB.TextBox txtOrgan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4440
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   240
         Width           =   1332
      End
      Begin VB.TextBox txtProcedure 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4428
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   720
         Width           =   1320
      End
      Begin VB.TextBox TxtBlockID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8880
         TabIndex        =   0
         Top             =   720
         Width           =   2472
      End
      Begin VB.Label LblInternalNbr 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8880
         TabIndex        =   44
         Top             =   360
         Width           =   1632
      End
      Begin VB.Label lblNumOfBlocksPerSdg 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   600
         Width           =   552
      End
      Begin VB.Label lblNumBlocksPerSdg 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "בלוקים למקרה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   840
         TabIndex        =   30
         Top             =   720
         Width           =   1416
      End
      Begin VB.Label lblNumOfBlocksPerSample 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   120
         Width           =   552
      End
      Begin VB.Label lblNumOfBlocks 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "בלוקים לדגימה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   960
         TabIndex        =   28
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label LblInternalNbrTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "מס. פתולאב"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   7392
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   840
      End
      Begin VB.Label LblLocation 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2760
         TabIndex        =   12
         Top             =   600
         Width           =   1392
      End
      Begin VB.Label LblLocationTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "איפיון הבלוק"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   2760
         TabIndex        =   11
         Top             =   240
         Width           =   1284
      End
      Begin VB.Label LblOrganTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "איבר"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblPatholbNbr 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6240
         TabIndex        =   8
         Top             =   600
         Width           =   2472
      End
      Begin VB.Label LblBlockIDTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "מס. בלוק"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10560
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   876
      End
   End
   Begin VB.CommandButton OkButton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   864
      Left            =   9600
      Picture         =   "BlockGenerationReportCtrl.ctx":1108
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "לחץ לאישור"
      Top             =   5959
      Width           =   2052
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "סך הכל בלוקים"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   3264
      TabIndex        =   46
      Top             =   6100
      Width           =   1368
   End
   Begin VB.Label lblCounter 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   6000
      Width           =   792
   End
   Begin VB.Label LblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "עמדת שיקוע"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6228
      TabIndex        =   26
      Top             =   11
      Width           =   2676
   End
End
Attribute VB_Name = "BlockCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements LSExtensionWindowLib.IExtensionWindow
Implements LSExtensionWindowLib.IExtensionWindow2

Option Explicit

'rtf box
 Private Declare Function PeekMessageW Lib "user32" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Const WM_KEYFIRST = &H100
 Const WM_KEYLAST = &H108
 Private Type POINTAPI
    x As Long
    y As Long
End Type
 Private Type Msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Dim UKEY$


'הגדרת צבעים גלובליים
Private Const RED = &HFF&
Private Const WHITE = &HFFFFFF
Private Const BLUE = &H80000013
Private Const BLACK = &H80000008
Private Const MARK_SELECTED = &H80000001
Private Const MARK_APPROVED = &H80000018
Private Const LIGHT_BLUE = &HFFC0C0
Private Const LIGHT_YELLOW = &H80000018
Private Const csHeBrEw As String = "iso-8859-8" ' Hebrew character set
Private Const nInch = 1440

'define roles (patholog / doctor / ...)
Private Const DOCTOR_ROLE = "63"
 
Private ProcessXML As LSSERVICEPROVIDERLib.NautilusProcessXML
Private NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
Private NtlsSite As LSExtensionWindowLib.IExtensionWindowSite2
Private NtlsUser As LSSERVICEPROVIDERLib.NautilusUser
Private sp As LSSERVICEPROVIDERLib.NautilusServiceProvider

Private Con As ADODB.Connection
Private Block As ADODB.Recordset
Private BlockFlag As Boolean
Private OperatorName As String
Private OperatorID As String
Private BlockID As String
Private SdgID As Long
Private TestID As String
Private ResultID As String
Private SampleType As String
Private sdg_log As New SdgLog.CreateLog
Private sdg_log_desc As String
Private strExternalReference As String
Private FirstResult As Boolean
Private WorkFolder As String
Private DisplayRemarks As String
Private PathologCodesNumberToName As Scripting.Dictionary
Private PriorityDic As Scripting.Dictionary
Private strExtraRequestDetails As String
Private strExtraRequestDescription As String
Private strExtraRequestDataId As String
Private strExtraRequestCreatedBy As String
Private strExtraRequestCreatedOn As String
Private OpenedRequest As Boolean


Private isCytology As Boolean


'Private Sub ChckSettingUp_Click()
'
'    If ChckSettingUp.Value = 0 Then
'        ChckSettingUp.ForeColor = &HFF0000
'        ChckSettingUp.Caption = "לא"
'    Else
'        ChckSettingUp.ForeColor = &HFF&
'        ChckSettingUp.Caption = "כן"
'    End If
'End Sub

Private Sub CmbAliquotRemark_GotFocus()
10        Call zLang.Hebrew
End Sub

Private Sub CmbAliquotRemark_LostFocus()
20        Call zLang.SetOrigLang
End Sub

Private Sub CmdCancel_Click()
          Dim MBRes As VbMsgBoxResult

30        MBRes = MsgBox("? האם את/ה בטוח שברצונך לצאת מממסך זה", vbYesNo, "Nautilus - דיווח הכנת בלוק")
40        If MBRes = vbNo Then Exit Sub

50        Call RemoveAll
60        Call zLang.SetOrigLang
If Not NtlsSite Is Nothing Then
70        Call NtlsSite.CloseWindow
End If
End Sub

Private Sub CmdCancel_GotFocus()
80        OkButton.TabIndex = CmdCancel.TabIndex + 1
End Sub



Private Sub cmdExtraRequest_Click()
90    On Error GoTo ERR_cmdExtraRequest_Click

100       Call frmExtraRequest.Initialize(strExtraRequestDetails, strExtraRequestDescription)
110       Call frmExtraRequest.Show(vbModal)

120       Exit Sub
ERR_cmdExtraRequest_Click:
130   MsgBox "Error on line:" & Erl & " in  cmdExtraRequest_Click" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub cmdMacro_Click()
       Dim RequestNBR As String
          Dim strParameters As String
          Dim rst As ADODB.Recordset
          Dim frs As New frmMacro
          Dim sql As String
          
          
140       RequestNBR = Trim(LblInternalNbr.Caption)
150       If RequestNBR = "" Then Exit Sub
           
           
160        frs.MacroHisCtrl.RunFromWindow = True
170        Call frs.MacroHisCtrl.IExtensionWindow_SetSite(NtlsSite)

180        Call frs.MacroHisCtrl.IExtensionWindow_SetServiceProvider(sp)
190        frs.MacroHisCtrl.IExtensionWindow_Internationalise

200        Set rst = Con.Execute("select PARAMETER2 from lims_sys.command " & _
               "where name = 'Macro Histology'")
210        If Not rst.EOF Then
220            strParameters = Trim(nte(rst("PARAMETER2")))
230            Call frs.MacroHisCtrl.IExtensionWindow_SetParameters(strParameters)
240        End If

250        rst.Close

260        frs.MacroHisCtrl.IExtensionWindow_PreDisplay

270        frs.MacroHisCtrl.IExtensionWindow_GetButtons

280        frs.MacroHisCtrl.IExtensionWindow_Setup
           
290       sql = " select s.name "
300       sql = sql & " from lims_sys.sample s , lims_sys.aliquot a "
310       sql = sql & " where "
320       sql = sql & " a.name='" & RequestNBR & "'"
330       sql = sql & " and s.SAMPLE_ID=a.SAMPLE_ID"
340        Set rst = Con.Execute(sql)
350        If Not rst.EOF Then
360            frs.MacroHisCtrl.InitiateSample (UCase(nte(rst(0))))
370            frs.Show vbModal
380         End If
390        rst.Close
400        frs.MacroHisCtrl.IExtensionWindow_CloseQuery
410        Unload frs
420        Set frs = Nothing
End Sub

Private Sub cmdMacrotxt_Click()

    Dim ResID As String
    Dim RtfResult As New ADODB.Recordset

    ResID = GetResultID("Histology Macro text")

    If Trim(ResID) <> "" Then
        Set RtfResult = New ADODB.Recordset
        Call RtfResult.Open("select rtf_text from lims_sys.rtf_result where rtf_result_id = " & ResID, Con, adOpenStatic, adLockOptimistic)
        If Not RtfResult.EOF Then
            frmTxtMacro.Initialize ReadClob(RtfResult("RTF_TEXT"))
           frmTxtMacro.Show vbModal
      
        End If
    End If
End Sub

Private Function ReadClob(pFld As ADODB.Field) As String

          ' Function read a the clob data from the field
          ' using the stream object of the ADODB library

          Dim lStream As ADODB.Stream
          Dim lstData As String

21020     Set lStream = New ADODB.Stream
21030     lStream.Charset = csHeBrEw
21040     lStream.Type = adTypeText
21050     lStream.Open

21060     lStream.WriteText nte(pFld.Value)
21070     lStream.position = 0
21080     lstData = lStream.ReadText

21090     lStream.Close
21100     Set lStream = Nothing
          
21110     ReadClob = lstData
          
End Function


Private Function GetResultID(ResultName As String) As String
          Dim ResultRec As Recordset
          Dim ResultStr As String
          
      
21120     ResultStr = "select result.result_id from lims_sys.result, " & _
                          "lims_sys.test, lims_sys.aliquot, lims_sys.sample " & _
                          "where result.name = '" & ResultName & "' And " & _
                          "sample.sdg_id = " & SdgID & " and " & _
                          "aliquot.sample_id = sample.sample_id and " & _
                          "test.aliquot_id = aliquot.aliquot_id and " & _
                          "result.test_id = test.test_id and " & _
                          "result.status <> 'X' "
                          
21130     Set ResultRec = Con.Execute(ResultStr)

21140     If ResultRec.EOF Then
21150         GetResultID = ""
21160     Else
21170         GetResultID = ResultRec("RESULT_ID")
21180     End If
End Function

Private Function IExtensionWindow_CloseQuery() As Boolean
          'Happens when the user close the window
430       Set Block = Nothing
440       IExtensionWindow_CloseQuery = True
End Function

Private Function IExtensionWindow_DataChange() As LSExtensionWindowLib.WindowRefreshType
450       IExtensionWindow_DataChange = windowRefreshNone
End Function

Private Function IExtensionWindow_GetButtons() As LSExtensionWindowLib.WindowButtonsType
460       IExtensionWindow_GetButtons = windowButtonsNone
End Function

Private Sub IExtensionWindow_Internationalise()
End Sub

Private Sub IExtensionWindow_PreDisplay()
470       On Error GoTo ErrEnd
          Dim constr As String

480       Set Block = New ADODB.Recordset
490       Set Con = New ADODB.Connection

500       constr = "Provider=OraOLEDB.Oracle" & _
              ";Data Source=" & NtlsCon.GetServerDetails & _
              ";User ID=" & NtlsCon.GetUsername & _
              ";Password=" & NtlsCon.GetPassword
              
         If NtlsCon.GetServerIsProxy Then
            constr = "Provider=OraOLEDB.Oracle;Data Source=" & _
            NtlsCon.GetServerDetails & ";User id=/;Persist Security Info=True;"
          End If


510       Con.Open constr
520       Con.CursorLocation = adUseClient

530       Con.Execute "SET ROLE LIMS_USER"
540       Call ConnectSameSession(CDbl(NtlsCon.GetSessionId))

550       VScroll.LargeChange = 20    ' Cross in 5 clicks.
560       VScroll.SmallChange = 5     ' Cross in 20 clicks.
570       Picture2.Container = Picture1

580       Call RequestRemarkCtrl.InitializeConnection(Con)
'MsgBox "opid " & NtlsUser.GetOperatorId
590       Call RequestRemarkCtrl.GetOperatorId(NtlsUser.GetOperatorId)
'MsgBox "opid " & NtlsUser.GetRoleName()
600       Call GetOperatorDetails(NtlsUser.GetOperatorId)
610       RequestRemarkCtrl.Visible = False

620       Set sdg_log.Con = Con
630       sdg_log.session = CDbl(NtlsCon.GetSessionId)
          
640       OrganCtrl.Connection = Con
650       OrganCtrl.OperatorName = NtlsUser.GetOperatorName
660       OrganCtrl.SessionId = NtlsCon.GetSessionId

670       Call EnabledButton(False)
680       Call EnabledFields(False)
          'ChckSettingUp.Enabled = False

690       BlockFlag = False
700       Exit Sub

ErrEnd:
710      MsgBox "Error on line:" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub IExtensionWindow_refresh()
    'Code for refreshing the window
End Sub

Private Sub IExtensionWindow_RestoreSettings(ByVal hKey As Long)
End Sub

Private Function IExtensionWindow_SaveData() As Boolean
End Function

Private Sub IExtensionWindow_SaveSettings(ByVal hKey As Long)
End Sub

Private Sub IExtensionWindow_SetParameters(ByVal parameters As String)
720       DisplayRemarks = "F"
730       If parameters <> "" Then
740           DisplayRemarks = Trim(parameters)
750       End If
End Sub

Private Sub IExtensionWindow_SetServiceProvider(ByVal serviceProvider As Object)
          'Dim sp As LSSERVICEPROVIDERLib.NautilusServiceProvider
760       Set sp = serviceProvider
770       Set ProcessXML = sp.QueryServiceProvider("ProcessXML")
780       Set NtlsCon = sp.QueryServiceProvider("DBConnection")
790       Set NtlsUser = sp.QueryServiceProvider("User")
End Sub

Private Sub IExtensionWindow_SetSite(ByVal Site As Object)
800       Set NtlsSite = Site
810       NtlsSite.SetWindowInternalName ("MacabiBlockGenerationReport")
820       NtlsSite.SetWindowRegistryName ("MacabiBlockGenerationReport")
830       Call NtlsSite.SetWindowTitle("דיווח הכנת בלוק")
End Sub

Private Sub IExtensionWindow_Setup()
840       On Error GoTo ErrEnd
          Dim phrase As ADODB.Recordset
          Dim Organ As ADODB.Recordset
          Dim Mark As ADODB.Recordset
          Dim MarkYN As ADODB.Recordset
          Dim SampleFix As ADODB.Recordset
          Dim LeftOver As ADODB.Recordset
          Dim CassetteFix As ADODB.Recordset
          Dim ResultDesc As ADODB.Recordset
          Dim sql As String



          'Init the Aliquot Remark... combo
850       Set phrase = Con.Execute("select phrase_description, phrase_name from lims_sys.phrase_entry " & _
              "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
              "name = 'Aliquot Remark') " & _
              "order by order_number")

860       CmbAliquotRemark.List(0) = ""
870       CmbAliquotRemark.Text = ""
880       Do Until phrase.EOF
890           CmbAliquotRemark.List(CmbAliquotRemark.ListCount) = phrase("PHRASE_DESCRIPTION")
900           phrase.MoveNext
910       Loop
920       phrase.Close

          'Init the sample suspension combo
930       Set phrase = Con.Execute("select phrase_description, phrase_name from lims_sys.phrase_entry " & _
              "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
              "name = 'Embedding') " & _
              "order by order_number")

940       CmbReason(0).List(0) = "None"
950       Do Until phrase.EOF
960           CmbReason(0).List(CmbReason(0).ListCount) = phrase("PHRASE_DESCRIPTION")
970           phrase.MoveNext
980       Loop
990       phrase.Close
          
          
          'Set phrase = Con.Execute("select ou.u_hebrew_name, ou.operator_id " & _
                            "from lims_sys.operator o, lims_sys.operator_user ou " & _
                            "where o.operator_id = ou.operator_id " & _
                            "and o.role_id = " & DOCTOR_ROLE)
                            
1000      sql = " select o.OPERATOR_ID,"
1010      sql = sql & "        ou.U_HEBREW_NAME"
1020      sql = sql & " from lims_sys.operator o, "
1030      sql = sql & "      lims_sys.operator_user ou"
1040      sql = sql & " where ou.OPERATOR_ID=o.OPERATOR_ID"
1050      sql = sql & " and   ou.U_PATHOLOG_MACRO = 'T'"
1060      sql = sql & " order by ou.U_ORDER"

1070      Set phrase = Con.Execute(sql)

1080      Set PathologCodesNumberToName = New Scripting.Dictionary
1090      Do Until phrase.EOF
1100          Call PathologCodesNumberToName.Add(CStr(phrase("OPERATOR_ID").Value), _
                                                 CStr(phrase("U_HEBREW_NAME").Value))
1110          phrase.MoveNext
1120      Loop
          
          
              Set phrase = Con.Execute("select phrase_description, phrase_name from lims_sys.phrase_entry " & _
              "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
              "name = 'Priority') " & _
              "order by order_number")
                            



      Set PriorityDic = New Scripting.Dictionary
      Do Until phrase.EOF
          Call PriorityDic.Add(CStr(phrase("phrase_name").Value), _
                                                 CStr(phrase("phrase_description").Value))
          phrase.MoveNext
      Loop
          

1130      WorkFolder = ""
1140      WorkFolder = xmlManager.GetDefaultFolderFromWorkStation(NtlsUser.GetWorkstationId, Con)
1150      If Trim(WorkFolder) <> "" Then
1160          xmlManager.XmlFolder = WorkFolder & "\BlockGenerationReport\"
1170      End If

1180      Call zLang.English
1190      TxtBlockID.Alignment = vbLeftJustify
1200      TxtBlockID.RightToLeft = False
1210      Call TxtBlockID.SetFocus

1220      SdgID = 0
1230      TestID = ""
1240      BlockID = ""
1250      ResultID = ""
1260      sdg_log_desc = ""
1270      FirstResult = False
1280      Exit Sub

ErrEnd:
1290      MsgBox "setup..." & vbCrLf & _
                   "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Function IExtensionWindow_ViewRefresh() As LSExtensionWindowLib.WindowRefreshType
1300      IExtensionWindow_ViewRefresh = windowRefreshNone
End Function

Private Sub ConnectSameSession(ByVal aSessionID)
          Dim aProc As New ADODB.Command
          Dim aSession As New ADODB.Parameter
          
1310      aProc.ActiveConnection = Con
1320      aProc.CommandText = "lims.lims_env.connect_same_session"
1330      aProc.CommandType = adCmdStoredProc

1340      aSession.Type = adDouble
1350      aSession.direction = adParamInput
1360      aSession.Value = aSessionID
1370      aProc.parameters.Append aSession

1380      aProc.Execute
1390      Set aSession = Nothing
1400      Set aProc = Nothing
End Sub

Private Sub IExtensionWindow2_Close()
End Sub

Private Function nte(e As Variant) As Variant
1410      nte = IIf(IsNull(e), "", e)
End Function

Public Function ntz(e As Variant) As Variant
1420      ntz = IIf(IsNull(e), 0, e)
End Function

 

Private Sub K_option_Click()
lblCounter.Caption = 0
End Sub

Private Sub OkButton_Click()
1430      Call OKButtonClick
End Sub

Private Sub OKButtonClick()
1440      On Error GoTo ErrEnd
1450      If Not BlockFlag Then Exit Sub
1460      If Trim(CmbReason(CmbReason.Count - 1).Text) = "None" Or _
             Trim(CmbReason(CmbReason.Count - 1).Text) = "" Then
1470          CmbReason(CmbReason.Count - 1).BackColor = RED
1480          MsgBox "יש לבחור סיבת שיקוע בשורה מס': " & (CmbReason.Count - 1), vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "Nautilus - Block Generation Report"
1490          CmbReason(CmbReason.Count - 1).BackColor = WHITE
1500          Call CmbReason(CmbReason.Count - 1).SetFocus
1510          Exit Sub
1520      End If
          Dim strRemark As String
          Dim ProgramName As String
          Dim SlideTitle As String
          Dim strDetails As String
          Dim ResultVal As String
          Dim NewResultId As String

1530      NewResultId = ""
1540      ProgramName = "התוכנית " & " BlockGenerationReport "
1550      SlideTitle = " לבלוק מס.: " & LblInternalNbr.Caption & " "
1560      strDetails = " בוצע אישור "
1570      strRemark = ProgramName & SlideTitle & _
                      " הופעלה על-ידי המשתמש: " & _
                      OperatorID & " - " & OperatorName & " " & _
                      "בתאריך: " & _
                      Now & " " & _
                      strDetails

1580      ResultVal = Trim(CmbReason(CmbReason.Count - 1).Text)
1590      BlockFlag = False
1600      Call UpdateBlockDetails(BlockID)
1610      Call UpdateBlockStatus(BlockID)
1620      Call UpdateAliquotTrace(BlockID)
1625      Call UpdateBlockAgree(BlockID)
1630      Call UpdateExtraRequestStatus(strExtraRequestDataId)
1640      Call ResetExtraRequestData
1650      cmdExtraRequest.Visible = False
          
1660      NewResultId = ResultID
1670      If Not FirstResult Then
1680          NewResultId = TriggerTestEvent("Add Embedding", TestID)
1690      End If
1700      Call AddEmbedding(NewResultId, ResultVal)

1710      If DisplayRemarks = "T" Then
1720          Call AddRequestRemark(strRemark)
1730      End If

          'the description in the sdg_log table
          'will be the name of the block:
1740      sdg_log_desc = LblInternalNbr.Caption
1750      Call sdg_log.InsertLog(SdgID, "BGR.UPD", sdg_log_desc)
1760      Call EnabledButton(False)
1770      Call EnabledFields(False)
          'ChckSettingUp.Enabled = False

          'show that this line was approved:
1780      CmbReason(CmbReason.Count - 1).Enabled = False
1790      CmbReason(CmbReason.Count - 1).BackColor = MARK_APPROVED
1800      LblRowNum(CmbReason.Count - 1).BackColor = MARK_APPROVED
1810      LblCreatedBy(CmbReason.Count - 1).BackColor = MARK_APPROVED
1820      LblCreatedOn(CmbReason.Count - 1).BackColor = MARK_APPROVED

1830      Call ClearPage

1840      Call zLang.English
1850      TxtBlockID.Alignment = vbLeftJustify
1860      TxtBlockID.RightToLeft = False
1870      Call TxtBlockID.SetFocus

1880      BlockFlag = False
          lblCounter.Caption = lblCounter.Caption + 1
        OpenedRequest = False
1890      Exit Sub

ErrEnd:
1900      MsgBox "OkButton_Click... " & vbCrLf & _
                   "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description

End Sub

Private Sub ClearPage()
Call InitScreenFields

End Sub
Private Sub UpdateBlockAgree(BlockID As String)
19101      On Error GoTo ErrEnd
          Dim strSql As String
                 Dim agrtype As String
          Dim BlockRs As ADODB.Recordset

     agrtype = IIf(K_option.Value, "1", "2")

19201    strSql = "update lims_sys.aliquot_user " & _
                   "set u_agree_type = '" & agrtype & "' " & _
                   " where aliquot_id = '" & BlockID & "'"
                   
                '   MsgBox strSql
19301     Set BlockRs = Con.Execute(strSql)
19401      Exit Sub

ErrEnd:
19501      MsgBox "UpdateBlockAgree... " & vbCrLf & _
                   "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub
Private Sub UpdateBlockStatus(BlockID As String)
1910      On Error GoTo ErrEnd
          Dim strSql As String
          Dim BlockRs As ADODB.Recordset

1920      strSql = "update lims_sys.aliquot " & _
                   "set status = 'V' " & _
                   "where aliquot_id = '" & BlockID & "'"
1930      Set BlockRs = Con.Execute(strSql)
1940      Exit Sub

ErrEnd:
1950      MsgBox "UpdateBlockStatus... " & vbCrLf & _
                   "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub UpdateBlockDetails(BlockID As String)
1960      On Error GoTo ErrEnd
          Dim strSql As String
          Dim BlockRs As ADODB.Recordset
          'Dim strSettingUp As String

          'strSettingUp = "F"
          'If ChckSettingUp.Value = 1 Then
          '    strSettingUp = "T"
              
          '    lblSettingUp.Caption = "כן"
          'End If

1970      strSql = "update lims_sys.aliquot_user " & _
                   "set U_ALIQUOT_REMARK = '" & CmbAliquotRemark & "' " & _
                   "where aliquot_id = '" & BlockID & "'"

      '    strSql = "update lims_sys.aliquot_user " & _
                   "set U_ALIQUOT_REMARK = '" & CmbAliquotRemark & "', " & _
                   "U_SETTING_UP = '" & strSettingUp & "' " & _
                   "where aliquot_id = '" & BlockID & "'"
1980      Set BlockRs = Con.Execute(strSql)
1990      Exit Sub

ErrEnd:
2000      MsgBox "UpdateBlockDetails... " & vbCrLf & _
                   "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub GetOperatorDetails(Operator_ID As String)
          Dim OperatorRST As ADODB.Recordset

2010      Set OperatorRST = Con.Execute("select OPERATOR_ID, U_HEBREW_NAME from lims_sys.operator_user " & _
                                              "where operator_id = '" & Operator_ID & "'")
2020      If Not OperatorRST.EOF Then
2030          OperatorID = Trim(nte(OperatorRST("OPERATOR_ID")))
2040          OperatorName = Trim(nte(OperatorRST("U_HEBREW_NAME")))
2050      End If
End Sub

Private Sub AddRequestRemark(strRemark As String)
2060      RequestRemarkCtrl.AddRemark (strRemark)
2070      RequestRemarkCtrl.Visible = True
2080      Call RequestRemarkCtrl.GetsdgName(strExternalReference)
2090      RequestRemarkCtrl.Refresh
End Sub

Private Sub EnabledFields(EnabledFlag As Boolean)
2100      If EnabledFlag Then
2110          CmbAliquotRemark.Enabled = True
2120      Else
2130          CmbAliquotRemark.Enabled = False
2140      End If
End Sub

Private Sub EnabledButton(EnabledFlag As Boolean)
2150      If EnabledFlag Then
2160          OkButton.Enabled = True
2170          OKButton2.Enabled = True
2180      Else
2190          OkButton.Enabled = False
2200          OKButton2.Enabled = False
2210      End If
End Sub

Private Sub RemoveAll()
    Dim i As Integer
End Sub

Private Sub OkButton_GotFocus()
2220      CmdCancel.TabIndex = OkButton.TabIndex + 1
End Sub



Private Sub OKButton2_Click()
2230       Call OKButtonClick
End Sub



Public Function GetLastKeyPressed() As Long
      Dim Message As Msg
2240      If PeekMessageW(Message, 0, WM_KEYFIRST, WM_KEYLAST, 0) Then
2250          GetLastKeyPressed = Message.wParam
2260      Else
2270          GetLastKeyPressed = -1
2280      End If
2290      Exit Function
End Function

Private Sub SampleRmrk_GotFocus()
2300      Call zLang.Hebrew
End Sub

Private Sub SampleRmrk_LostFocus()
2310      Call zLang.SetOrigLang
End Sub




Private Sub R_option_Click()
lblCounter.Caption = 0
End Sub

Private Sub TxtBlockID_KeyDown(KeyCode As Integer, Shift As Integer)
2320      On Error GoTo ErrEnd
2330      If Not KeyCode = vbKeyReturn Then Exit Sub
2340      If Trim(TxtBlockID.Text) = "" Then Exit Sub

          Dim Hi As Long
          Dim i As Integer
          Dim strSql As String
          Dim StrTitile As String
          Dim NumOfResult As Integer
          Dim strRequestPriority As String

2350      SdgID = 0
2360      TestID = ""
2370      BlockID = ""
2380      ResultID = ""
2390      sdg_log_desc = ""
2400      FirstResult = False
15510     If OpenedRequest Then
15520         If MsgBox("הבלוק לא עודכן." & vbCrLf & _
                      " ? האם אתה בטוח שברצונך להמשיך ", vbYesNo + vbDefaultButton2) = vbNo Then
                      TxtBlockID.Text = ""
15530             Exit Sub
15540         End If
15550      End If

2410      Call InitScreenFields
2420      Call UnloadFileds
2430      Call EnabledButton(False)
2440      Call EnabledFields(False)

        If R_option.Value = False And K_option.Value = False Then
            MsgBox "! אנא בחר קבלנות או רוטינה", vbOKOnly, "Nautilus - דיווח הכנת בלוק"
            TxtBlockID.Text = ""
             Exit Sub
        End If
        



2450      strSql = "select " & _
                      "d.SDG_ID, " & _
                      "d.EXTERNAL_REFERENCE, du.u_priority,  au.U_PATHOLAB_ALIQUOT_NAME, " & _
                      "s.SAMPLE_TYPE, s.SAMPLE_ID, " & _
                      "su.U_ORGAN, " & _
                      "su.U_TOPOGRAPHY, su.u_patholog_macro, " & _
                      "s.DESCRIPTION SAMPLE_DESCRIPTION, " & _
                      "a.aliquot_id BLOCK_ID, a.NAME, a.STATUS, " & _
                      "au.U_LOCATION, au.U_IS_CELL_BLOCK, " & _
                      "au.U_NUM_OF_TISSUES, " & _
                      "au.U_ALIQUOT_REMARK, " & _
                      "au.U_SETTING_UP, " & _
                      "t.test_id, " & _
                      "r.created_on, r.completed_on, r.formatted_result, r.status RESULT_STAT, " & _
                      "r.result_id, r.completed_by, au.U_ALIQUOT_STATION staion " & _
                   "from lims_sys.sdg d, lims_sys.sdg_user du,lims_sys.sample_user su, lims_sys.sample s, " & _
                      "lims_sys.aliquot a, lims_sys.aliquot_user au, lims_sys.test t, " & _
                      "lims_sys.result r " & _
                   "where d.sdg_id = du.sdg_id and s.sdg_id = d.sdg_id and s.sample_id = su.sample_id and " & _
                      "a.aliquot_id = au.aliquot_id and a.sample_id = s.sample_id and " & _
                      "a.aliquot_id = t.aliquot_id and " & _
                      " t.test_id = r.test_id " & _
                      " and t.name = 'Embedding' and " & _
                      "a.name = '" & UCase(TxtBlockID.Text) & "' " & _
                   "order by r.result_id"

      'o.full_name created_by
      ', lims_sys.operator o
      '" and r.created_by = o.operator_id   " & _

2460      Set Block = Con.Execute(strSql)

2470      If Block.EOF Then
2480          TxtBlockID.BackColor = RED
2490          MsgBox "! (" & Trim(UCase(TxtBlockID.Text)) & ") הבלוק אינו קיים במערכת ", , "Nautilus - דיווח הכנת בלוק"
2500          Call TxtBlockID.SetFocus
2510          TxtBlockID.BackColor = WHITE
2520          TxtBlockID.Text = ""
2530          Exit Sub
2540      End If

2550      If Block("STATUS") = "X" Or Block("STATUS") = "R" Or Block("STATUS") = "S" Then
2560          TxtBlockID.BackColor = RED
2570          MsgBox "! (" & Block("STATUS") & ") סטטוס הבלוק שגוי ", , "Nautilus - דיווח הכנת בלוק"
2580          Call TxtBlockID.SetFocus
2590          TxtBlockID.BackColor = WHITE
2600          TxtBlockID.Text = ""
2610          Exit Sub
2620      End If

2630      If Not CInt(ntz(Block("staion"))) >= 2 Then
2640          TxtBlockID.BackColor = RED
2650          MsgBox "! " & "בלוק לא עבר מאקרו", , "Nautilus - דיווח הכנת בלוק"
2660          Call TxtBlockID.SetFocus
2670          TxtBlockID.BackColor = WHITE
2680          TxtBlockID.Text = ""
2690          Exit Sub
2700      End If

 
          OpenedRequest = True
2710      FirstResult = False
2720      If Block.RecordCount = 1 And Trim(nte(Block("RESULT_STAT"))) = "V" Then
2730          FirstResult = True
2740      End If
        
2750      BlockID = nte(Block("BLOCK_ID"))
2760      SdgID = nte(Block("SDG_ID"))
2770      TestID = nte(Block("TEST_ID"))
2780      ResultID = nte(Block("RESULT_ID"))
2790      SampleType = nte(Block("SAMPLE_TYPE"))
2800      strExternalReference = nte(Block("EXTERNAL_REFERENCE"))
2810      LblInternalNbr.Caption = nte(Block("NAME"))
2820      lblPatholbNbr.Caption = nte(Block("U_PATHOLAB_ALIQUOT_NAME"))
                
2830      If Block("U_IS_CELL_BLOCK") = "T" Then
2840          isCytology = True
2850      Else
2860          isCytology = False
2870      End If
          
          'OrganCtrl.SdgID = nte(Block("SDG_ID"))
2880      OrganCtrl.SampleId = nte(Block("SAMPLE_ID"))
2890      OrganCtrl.Initialize

2900      ReadOrganDetails
          'lblOrgan.Caption = OrganCtrl.Organ
          'LblTopography.Caption = OrganCtrl.Topography
         
2910      LblLocation.Caption = nte(Block("U_LOCATION"))
2920      LblNumOfPieces.Caption = nte(Block("U_NUM_OF_TISSUES"))

2940      CmbAliquotRemark.Text = nte(Block("U_ALIQUOT_REMARK"))
2950      strRequestPriority = nte(Block("u_priority"))
2960      lblPathologMacro = PathologCodesNumberToName(nte(Block("u_patholog_macro")))
          
          'ChckSettingUp.Value = 0
2970      If nte(Block("U_SETTING_UP")) = "T" Then
          '    ChckSettingUp.Value = 1
              
2980          lblSettingUp.Caption = "כן"
2990          lblSettingUp.ForeColor = &HFF&
3000      Else
3010          lblSettingUp.Caption = ""
      '        lblSettingUp.Caption = "לא"
      '        lblSettingUp.ForeColor = &HFF0000
3020      End If
        
3030      lblSettingUp.Visible = True

      '    If ChckSettingUp.Value = 0 Then
      '        ChckSettingUp.ForeColor = &HFF0000
      '        ChckSettingUp.Caption = "לא"
      '    Else
      '        ChckSettingUp.ForeColor = &HFF&
      '        ChckSettingUp.Caption = "כן"
      '    End If
          
          
          'ChckSettingUp.Enabled = True
          'If Trim(nte(Block("U_SETTING_UP"))) <> "" Then
          '    ChckSettingUp.Enabled = False
          'End If

3040      NumOfResult = 0
3050      While Not Block.EOF
3060          NumOfResult = NumOfResult + 1
              'existing results that are already reported - exist flag = true
              'not yet exist or not yet reported          - exist flag = false
              'a different caption and color is shown in each of the two cases:
3070          Hi = InitResultScreen(Block, NumOfResult, True And nte(Block("formatted_result")) <> "")
3080          Block.MoveNext
3090      Wend
3100      Block.Close

3110      If Not FirstResult Then
3120          NumOfResult = NumOfResult + 1
3130          Hi = InitResultScreen(Block, NumOfResult, False)
3140      End If

3150      Picture2.Height = Hi + 170
3160      VScroll.Max = (Picture2.Height - Picture1.Height) / ScaleHeight * 100
3170      If VScroll.Max < 0 Then
3180          VScroll.Visible = False
3190      Else
3200          VScroll.Visible = True
3210      End If

3220      RequestRemarkCtrl.Visible = True
3230      Call RequestRemarkCtrl.GetsdgName(strExternalReference)
3240      RequestRemarkCtrl.Refresh

3250      StrTitile = " דיווח הכנת בלוק " & Trim(UCase(TxtBlockID.Text))
         If Not NtlsSite Is Nothing Then
3260      Call NtlsSite.SetWindowTitle(StrTitile)
End If
3270      TxtBlockID.Text = ""


          'show the numbers of blocks per sample / sdg:
          Dim strSampleId As String
          Dim rs As Recordset
          Dim rsNumBlocks As ADODB.Recordset
          Dim sql As String

3280      Set rs = Con.Execute("select sample_id from lims_sys.aliquot " & _
                               "where aliquot_id= " & BlockID & " ")
                  
3290      strSampleId = rs(0)
          
3300      If Not isCytology Then
          
3310          sql = " select count(*) from lims_sys.aliquot_user, lims_sys.aliquot "
3320          sql = sql & " where aliquot.aliquot_id = aliquot_user.aliquot_id and "
3330          sql = sql & " status not in ('R', 'U', 'X') and sample_id = " & strSampleId & " and "
3340          sql = sql & " not exists ( select 1 from lims_sys.aliquot_formulation "
3350          sql = sql & " where aliquot_formulation.child_aliquot_id = aliquot.aliquot_id ) "
                  
3360          Set rsNumBlocks = Con.Execute(sql)
          
3370          If Not rsNumBlocks.EOF Then
3380              lblNumOfBlocksPerSample.Caption = nte(rsNumBlocks(0))
3390          End If
          
3400          sql = " select count(*) from lims_sys.aliquot_user, lims_sys.aliquot "
3410          sql = sql & " where aliquot.aliquot_id = aliquot_user.aliquot_id and "
3420          sql = sql & " status not in ('R', 'U', 'X') and sample_id in "
3430          sql = sql & " (select sample_id from lims_sys.sample where sdg_id=" & SdgID & ") and "
3440          sql = sql & " not exists ( select 1 from lims_sys.aliquot_formulation "
3450          sql = sql & " where aliquot_formulation.child_aliquot_id = aliquot.aliquot_id ) "
                      
3460          Set rsNumBlocks = Con.Execute(sql)
          
3470          If Not rsNumBlocks.EOF Then
3480              lblNumOfBlocksPerSdg.Caption = nte(rsNumBlocks(0))
3490          End If
          
3500      Else ' i.e. it is a cell block - therefore only cell block aliqout should be shown
               
3510          sql = " select count(*) from lims_sys.aliquot_user, lims_sys.aliquot "
3520          sql = sql & " where aliquot.aliquot_id = aliquot_user.aliquot_id and "
3530          sql = sql & " aliquot_user.U_IS_CELL_BLOCK = 'T' and"
3540          sql = sql & " status not in ('R', 'U', 'X') and sample_id = " & strSampleId & " and "
3550          sql = sql & " not exists ( select 1 from lims_sys.aliquot_formulation "
3560          sql = sql & " where aliquot_formulation.child_aliquot_id = aliquot.aliquot_id ) "
                  
3570          Set rsNumBlocks = Con.Execute(sql)
          
3580          If Not rsNumBlocks.EOF Then
3590              lblNumOfBlocksPerSample.Caption = nte(rsNumBlocks(0))
3600          End If
          
3610          sql = " select count(*) from lims_sys.aliquot_user, lims_sys.aliquot "
3620          sql = sql & " where aliquot.aliquot_id = aliquot_user.aliquot_id and "
3630          sql = sql & " aliquot_user.U_IS_CELL_BLOCK = 'T' and"
3640          sql = sql & " status not in ('R', 'U', 'X') and sample_id in "
3650          sql = sql & " (select sample_id from lims_sys.sample where sdg_id=" & SdgID & ") and "
3660          sql = sql & " not exists ( select 1 from lims_sys.aliquot_formulation "
3670          sql = sql & " where aliquot_formulation.child_aliquot_id = aliquot.aliquot_id ) "
                      
3680          Set rsNumBlocks = Con.Execute(sql)
          
3690          If Not rsNumBlocks.EOF Then
3700              lblNumOfBlocksPerSdg.Caption = nte(rsNumBlocks(0))
3710          End If
          
3720      End If
          
          'set the priority label
3730      lblPriority.Caption = PriorityDic(strRequestPriority)
3740      If CInt(strRequestPriority) > 2 Then
3750          lblPriority.BackColor = RED
3760      Else
3770          lblPriority.BackColor = LIGHT_YELLOW
3780      End If

3790      Call EnabledButton(True)
3800      Call EnabledFields(True)
3810      Call TxtBlockID.SetFocus
3820      Call zLang.SetOrigLang
3830      BlockFlag = True
          
          
3840      If ExtraRequest(LblInternalNbr.Caption) = True Then
3850          cmdExtraRequest.Caption = "בקשות חוזרות: " & _
                                         GetOperatorName(strExtraRequestCreatedBy) & ", " & _
                                         strExtraRequestCreatedOn
              
3860          cmdExtraRequest.Visible = True
3870      Else
3880          cmdExtraRequest.Visible = False
3890      End If
          
3900      Exit Sub

ErrEnd:
3910      MsgBox "txtBlockID_KeyDown" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub VScroll_Change()
3920      Picture2.Top = -(VScroll.Value / 100) * ScaleHeight + 50
End Sub

Private Sub InitScreenFields()
3930      LblInternalNbr.Caption = ""
3940      lblPatholbNbr.Caption = ""
3950      txtOrgan.Text = ""
3960      txtProcedure.Text = ""
          'LblTopography.Caption = ""
3970      LblLocation.Caption = ""
3980      LblNumOfPieces.Caption = ""

4000      CmbAliquotRemark.Text = ""
4010    lblPriority.Caption = ""
4020    lblPathologMacro.Caption = ""
4030    lblNumOfBlocksPerSample.Caption = ""
4040    lblNumOfBlocksPerSdg.Caption = ""



          'ChckSettingUp.Value = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
          Dim strVer As String

4050      If KeyCode = vbKeyF10 And Shift = 1 Then
4060          strVer = "Name: " & App.EXEName & vbCrLf & vbCrLf & _
                       "Path: " & App.Path & vbCrLf & vbCrLf & _
                       "Version: " & "[" & App.Major & "." & App.Minor & "." & App.Revision & "]" & vbCrLf & vbCrLf & _
                       "Company: One Software Technologies (O.S.T) Ltd."
4070          MsgBox strVer, vbInformation, "Nautilus - Project Properties"
4080          Call TxtBlockID.SetFocus
4090      End If
End Sub

Private Function InitResultScreen(ResultRec As Recordset, _
                                  NumOfResult As Integer, _
                                  ExistsFlag As Boolean) As Long
4100  On Error GoTo ERR_InitResultScreen

          Dim i As Integer

4110      Load LblRowNum(NumOfResult)
4120      With LblRowNum(NumOfResult)
4130          .Top = 500 * (NumOfResult - 1) + 10
4140          .Caption = NumOfResult
4150          .Visible = True
4160          InitResultScreen = .Top + .Height
4170          If (Not ExistsFlag) Then 'Or nte(ResultRec("FORMATTED_RESULT")) = "" Then
4180              .Caption = "<- " & NumOfResult
4190              .BackColor = LIGHT_BLUE
                  '.BackColor = &H80000018
4200          End If
4210      End With

          'completed-by instead of created-by;
          'we are interested in the operator who REPORTED the embedding
          'and if no one yet reported should get a empty string
4220      Load LblCreatedBy(NumOfResult)
4230      With LblCreatedBy(NumOfResult)
4240          .Top = 500 * (NumOfResult - 1) + 10
4250          .Caption = ""
4260          If ExistsFlag Then
4270              .Caption = GetOperatorName(nte(ResultRec("COMPLETED_BY")))
                  '.Caption = nte(ResultRec("COMPLETED_BY"))
                  '.Caption = nte(ResultRec("CREATED_BY"))
4280          End If
4290          .Visible = True
4300      End With

4310      Load LblCreatedOn(NumOfResult)
4320      With LblCreatedOn(NumOfResult)
4330          .Top = 500 * (NumOfResult - 1) + 10
4340          .Caption = ""
4350          If ExistsFlag Then
4360              .Caption = nte(ResultRec("COMPLETED_ON"))
                  '.Caption = nte(ResultRec("CREATED_ON"))
4370          End If
4380          .Visible = True
4390      End With

4400      Load CmbReason(NumOfResult)
4410      With CmbReason(NumOfResult)
4420          .Top = 500 * (NumOfResult - 1) + 10
4430          .Visible = True
4440          For i = 0 To CmbReason(0).ListCount - 1
4450              CmbReason(NumOfResult).List(i) = CmbReason(0).List(i)
4460          Next i
4470          If ExistsFlag Then
4480              .Enabled = False
4490              If FirstResult Then
4500                  If ResultRec("RESULT_STAT") = "V" Then
4510                      .Enabled = True
4520                      .ListIndex = 1
4530                  End If
4540              End If
4550              If Trim(nte(ResultRec("FORMATTED_RESULT"))) <> "" Then
4560                  .Text = nte(ResultRec("FORMATTED_RESULT"))
4570              End If
4580          Else
4590              .Enabled = True
                  
                  'show "embedding 1" for the 1st result (choose it for the user) to be made;
                  'show "None" for any additional embedding yet to be made:
4600              If NumOfResult = 1 Then
4610                  .Text = .List(1)
4620              Else
4630                  .Text = .List(0)
4640              End If
                  '.Text = "None"
4650          End If
4660      End With
          
4670      Exit Function
ERR_InitResultScreen:
4680  MsgBox "Error on line:" & Erl & " in  InitResultScreen" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function

Private Sub UnloadFileds()
          Dim i As Integer

4690      For i = 1 To LblRowNum.Count - 1
4700          Unload LblRowNum(i)
4710      Next

4720      For i = 1 To LblCreatedBy.Count - 1
4730          Unload LblCreatedBy(i)
4740      Next

4750      For i = 1 To LblCreatedOn.Count - 1
4760          Unload LblCreatedOn(i)
4770      Next

4780      For i = 1 To CmbReason.Count - 1
4790          Unload CmbReason(i)
4800      Next
End Sub

Private Sub AddEmbedding(ResultID As String, ResultVal As String)
4810      On Error GoTo ErrEnd
          Dim Xmldoc As New DOMDocument
          Dim Xmlres As New DOMDocument
          Dim XmlELimsReq As IXMLDOMElement
          Dim XmlEResultReq As IXMLDOMElement
          Dim XmlELoad As IXMLDOMElement
          Dim XmlEResultEntry As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String

4820      Set XmlELimsReq = Xmldoc.createElement("lims-request")
4830      XmlELimsReq.setAttribute "version", "1"
4840      Set XmlEResultReq = Xmldoc.createElement("result-request")
4850      XmlEResultReq.setAttribute "version", "1"
4860      Set XmlELoad = Xmldoc.createElement("load")
4870      XmlELoad.setAttribute "entity", "SDG"
4880      XmlELoad.setAttribute "id", SdgID
4890      XmlELoad.setAttribute "mode", "entry"

4900      Set XmlEResultEntry = Xmldoc.createElement("result-entry")
4910      XmlEResultEntry.setAttribute "result-id", ResultID
4920      XmlEResultEntry.setAttribute "original-result", ResultVal
4930      Call XmlELoad.appendChild(XmlEResultEntry)

4940      Call XmlEResultReq.appendChild(XmlELoad)
4950      Call XmlELimsReq.appendChild(XmlEResultReq)
4960      Call Xmldoc.appendChild(XmlELimsReq)

4970      If Trim(WorkFolder) <> "" Then
4980          FileName = "BlockGenerationReport_" & "AddEmbedding_" & SdgID & "_DOC1"
4990          Call xmlManager.SaveXmlFile(Xmldoc, FileName)
5000      End If

5010      RetError = ProcessXML.ProcessXMLWithResponse(Xmldoc, Xmlres)
5020      If Trim(RetError) <> "" Then
5030          MsgBox "Error occurred while trying process xml file. (AddEmbedding) " & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Block Generation Report"
5040      End If

5050      If Trim(WorkFolder) <> "" Then
5060          FileName = "BlockGenerationReport_" & "AddEmbedding_" & SdgID & "_RES1"
5070          Call xmlManager.SaveXmlFile(Xmlres, FileName)
5080      End If
5090      Exit Sub

ErrEnd:
5100     MsgBox "Error on line:" & Erl & " AddEmbedding ... " & vbCrLf & _
                  "Result ID: " & ResultID & vbCrLf & _
                  "Value: " & ResultVal & vbCrLf & _
                   "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Function TriggerTestEvent(EventName As String, TestID As String) As String
5110      On Error GoTo ErrEnd
          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlSdg As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String

5120      Set e = doc.createElement("lims-request")
5130      Call doc.appendChild(e)
5140      Set xmlLogin = doc.createElement("login-request")
5150      Call e.appendChild(xmlLogin)
5160      Set xmlSdg = doc.createElement("TEST")
5170      Call xmlLogin.appendChild(xmlSdg)
5180      Set element = doc.createElement("find-by-id")
5190      element.Text = TestID
5200      Call xmlSdg.appendChild(element)
5210      Set element = doc.createElement("fire-event")
5220      element.Text = EventName
5230      Call xmlSdg.appendChild(element)

5240      If Trim(WorkFolder) <> "" Then
5250          FileName = "BlockGenerationReport_" & EventName & "_" & TestID & "_DOC2"
5260          Call xmlManager.SaveXmlFile(doc, FileName)
5270      End If

5280      RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
5290      If Trim(RetError) <> "" Then
5300          MsgBox "Error occurred while trying process xml file. (TriggerTestEvent) " & vbCrLf & _
                     "Test ID: " & TestID & vbCrLf & _
                     "Event Name: " & EventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Block Generation Report"
5310      End If

5320      If Trim(WorkFolder) <> "" Then
5330          FileName = "BlockGenerationReport_" & EventName & "_" & TestID & "_RES2"
5340          Call xmlManager.SaveXmlFile(res, FileName)
5350      End If

5360      TriggerTestEvent = res.selectSingleNode("//RESULT_ID").Text
5370      Exit Function

ErrEnd:
5380      MsgBox "Error on line:" & Erl & " TriggerTestEvent... " & vbCrLf & _
                  "Test ID = " & TestID & vbCrLf & _
                  "Event Name = " & EventName & vbCrLf & _
                   "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function

'update the aliquot record to show that it
'was in this station.
'if it already was here, do not update
Private Sub UpdateAliquotTrace(strAliquotId As String)
5390  On Error GoTo REE_UpdateAliquotTrace

          Dim rs As Recordset
          Dim strSql As String
          Dim StrStationName As String
      '    Dim strOldStation
          
          'get this station's name from the phrase:
5400      Set rs = Con.Execute("select phrase_name from lims_sys.phrase_entry " & _
              "where phrase_description = 'block generation' and " & _
              "phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
              "name = 'AliquotStationTrace') " & _
              "order by order_number")
          
5410      StrStationName = rs("phrase_name")
          
          'check if this aliquot was in this station:
      '    strSql = "select u_old_aliquot_station " & _
                   "from lims_sys.aliquot_user " & _
                   "where aliquot_id = " & strAliquotId
                   
      '    Set rs = Con.Execute(strSql)
      '    strOldStation = nte(rs("u_old_aliquot_station"))
          
          'update pass through this station if needed:
      '    If InStr(1, strOldStation, StrStationName, vbTextCompare) = 0 Then
         
5420          strSql = " update lims_sys.aliquot_user set " & _
                       " u_old_aliquot_station = u_old_aliquot_station || u_aliquot_station , " & _
                       " u_aliquot_station = '" & StrStationName & "' " & _
                       " where aliquot_id = " & strAliquotId
         
         '     strSql = "update lims_sys.aliquot_user " & _
                       "set u_aliquot_station = '" & StrStationName & "', " & _
                       "u_old_aliquot_station = u_old_aliquot_station || '" & StrStationName & "' " & _
                       "where aliquot_id = " & strAliquotId
                       
5430          Call Con.Execute(strSql)
      '    End If
                   
5440      Exit Sub
REE_UpdateAliquotTrace:
5450  MsgBox "Error on line:" & Erl & " UpdateAliquotTrace" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Function GetOperatorName(strOperatorId As String) As String
5460  On Error GoTo ERR_GetOperatorName
          Dim rs As Recordset
          Dim sql As String
          
5470      If strOperatorId = "" Then Exit Function
          
5480      sql = " select o.FULL_NAME"
5490      sql = sql & " from lims_sys.operator o"
5500      sql = sql & " where o.OPERATOR_ID=" & strOperatorId

5510      Set rs = Con.Execute(sql)

5520      If Not rs.EOF Then GetOperatorName = nte(rs("FULL_NAME"))

5530      Exit Function
ERR_GetOperatorName:
5540  MsgBox "Error on line:" & Erl & " in  GetOperatorName" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function

'get the last reembedding extra request for this block;
'return false if no reembeding request exists, true otherwise;
'set the values for the variables used in initializing of the extra request form
Private Function ExtraRequest(strBlockName As String) As Boolean
5550  On Error GoTo ERR_ExtraRequest
          Dim rs As Recordset
          Dim sql As String

5560      ExtraRequest = False
5570      strExtraRequestDataId = ""

5580      sql = "  select rdu.U_REQUEST_DETAILS, rd.DESCRIPTION, rd.U_EXTRA_REQUEST_DATA_ID, "
5590      sql = sql & " ru.u_created_by, ru.u_created_on "
5600      sql = sql & "  from lims_sys.u_extra_request r,"
5610      sql = sql & "       lims_sys.u_extra_request_user ru,"
5620      sql = sql & "       lims_sys.u_extra_request_data rd, "
5630      sql = sql & "       lims_sys.u_extra_request_data_user rdu"
5640      sql = sql & "  where rd.U_EXTRA_REQUEST_DATA_ID=rdu.U_EXTRA_REQUEST_DATA_ID"
5650      sql = sql & "  and   r.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
5660      sql = sql & "  and   ru.U_EXTRA_REQUEST_ID=r.U_EXTRA_REQUEST_ID"
5670      sql = sql & "  and   r.NAME like 'Re Embedding;%'"
5680      sql = sql & "  and   rd.NAME like '" & strBlockName & ";%'"
5690      sql = sql & "  and   rdu.U_STATUS='P' "
5700      sql = sql & "  order by r.U_EXTRA_REQUEST_ID desc"

5710      Set rs = Con.Execute(sql)
          
5720      If rs.EOF = True Then
5730          Exit Function
5740      End If

          'init the Extra Request Data:
5750      strExtraRequestDetails = nte(rs("U_REQUEST_DETAILS"))
5760      strExtraRequestDescription = nte(rs("DESCRIPTION"))
5770      strExtraRequestDataId = nte(rs("U_EXTRA_REQUEST_DATA_ID"))
5780      strExtraRequestCreatedBy = nte(rs("U_CREATED_BY"))
5790      strExtraRequestCreatedOn = nte(rs("U_CREATED_ON"))

5800      ExtraRequest = True
          
5810      Exit Function
ERR_ExtraRequest:
End Function

'update the status of the extra request data
'to show it was reported in the inlab process
Private Sub UpdateExtraRequestStatus(strExtraRequestDataId As String)
5820  On Error GoTo ERR_UpdateExtraRequestStatus
          Dim sql As String
          
5830      If strExtraRequestDataId = "" Then Exit Sub

5840      sql = " update lims_sys.u_extra_request_data_user"
5850      sql = sql & " set u_status = 'L',"
5860      sql = sql & " u_lab_on = to_char (sysdate) "
5870      sql = sql & " where u_extra_request_data_id = '" & strExtraRequestDataId & "'"

5880      Call Con.Execute(sql)

5890      Exit Sub
ERR_UpdateExtraRequestStatus:
5900  MsgBox "Error on line:" & Erl & " in  UpdateExtraRequestStatus" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

'reset global data so on next block input
'it will not see it for an extra request by mistake:
Private Sub ResetExtraRequestData()
5910  On Error GoTo ERR_ResetExtraRequestData
5920      strExtraRequestDetails = ""
5930      strExtraRequestDescription = ""
5940      strExtraRequestDataId = ""
5950      strExtraRequestCreatedBy = ""
5960      strExtraRequestCreatedOn = ""
          
5970      Exit Sub
ERR_ResetExtraRequestData:
5980  MsgBox "Error on line:" & Erl & " in  ResetExtraRequestData" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub ReadOrganDetails()
5990  On Error GoTo ERR_ReadOrganDetails

          Dim strProcedure As String
          
6000      strProcedure = OrganCtrl.ProcedureCode
6010      Select Case strProcedure
              Case ""
              
6020          Case Else
6030              strProcedure = strProcedure & " - " & OrganCtrl.ProcedureName
6040      End Select
          
6050      txtOrgan.Text = OrganCtrl.Organ
6060      txtProcedure.Text = strProcedure

6070      Exit Sub
ERR_ReadOrganDetails:
6080  MsgBox "Error on line:" & Erl & " in  ReadOrganDetails" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub
