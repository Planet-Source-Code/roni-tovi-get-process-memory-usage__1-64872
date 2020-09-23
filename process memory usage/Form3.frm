VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   Caption         =   "Form3"
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2100
   LinkTopic       =   "Form3"
   ScaleHeight     =   1230
   ScaleWidth      =   2100
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Move Me.Left, Me.Top, Form1.myPic.Width, Form1.myPic.Height
    Set Picture1.Picture = Form1.myPic
End Sub

Private Sub Form_Resize()
    Picture1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
