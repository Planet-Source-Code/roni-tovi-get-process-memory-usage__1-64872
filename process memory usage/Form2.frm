VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
   LinkTopic       =   "Form2"
   ScaleHeight     =   2775
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.FileListBox File1 
         Height          =   870
         Left            =   1920
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1215
         Left            =   1560
         TabIndex        =   10
         Top             =   1320
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   0
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3600
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4440
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Shape Shape1 
         Height          =   1575
         Left            =   3240
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
