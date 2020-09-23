VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Get memory usage of any running  process"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Other processes"
      ForeColor       =   &H00000080&
      Height          =   2055
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Width           =   4935
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   1695
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "These simple memory usage actions are for test purpose"
      ForeColor       =   &H00000080&
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton Command10 
         Caption         =   "Clear Picture"
         Height          =   375
         Left            =   2520
         TabIndex        =   18
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Unload file"
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Unload all forms"
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Free all arrays"
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Refresh list"
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Draw picture"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Load File"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Load bitmap from file"
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Create ~400K array"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load 100 forms"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "This process"
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   4935
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   3840
         Top             =   240
      End
      Begin MSComDlg.CommonDialog common 
         Left            =   4320
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Current PageFile Usage: "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblPF 
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblUsage 
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Current Memory Usage: "
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPic As IPictureDisp
Private FormNo As Integer

Private Sub Command1_Click()
    On Error GoTo err_catch
    FormNo = FormNo + 100
    ReDim Forms(1 To FormNo) As Form2
    Dim x As Integer
    For x = 1 To UBound(Forms)
        Set Forms(x) = New Form2
    Next
    lblStatus.Caption = "Loaded " & IIf(UBound(Forms) > 100, "another ", "") & "100 forms" & IIf(UBound(Forms) > 100, " (" & Format(UBound(Forms), "###,###") & " forms now)", "")
    Exit Sub
err_catch:
    MsgBox "Can't load form" & String(2, vbNewLine) & Err.Number & ":" & Err.Description, vbCritical, "Load Error"
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Command10_Click()
    If myPic Is Nothing Then lblStatus.Caption = "Picture already cleared or never loaded!": Exit Sub
    Unload Form3
    Set Form3 = Nothing
    Set myPic = Nothing
    lblStatus.Caption = "Picture cleared"
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    PersonCount = PersonCount + 10000
    ReDim Persons(1 To PersonCount)
    If Err Then
        MsgBox "Can't load arrays" & String(2, vbNewLine) & Err.Number & ":" & Err.Description, vbCritical, "Load Error"
    Else
        lblStatus.Caption = "Created " & IIf(UBound(Persons) > 10000, "another ", "") & "~400K arrays" & IIf(UBound(Persons) > 10000, " (" & Format(UBound(Persons), "###,###") & " arrays now)", "")
    End If
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Dim sFile As String
    common.InitDir = "C:\"
    common.ShowOpen
    sFile = common.FileName
    If Not Exists(sFile) Then Exit Sub
    Set myPic = LoadPicture(sFile)
    If Err Then
        MsgBox "Can't load bitmap" & String(2, vbNewLine) & Err.Number & ":" & Err.Description, vbCritical, "Load Error"
    Else
        lblStatus.Caption = "Loaded " & FormatUsage(FileLen(sFile) / 1024) & " Kb (" & Round(FileLen(sFile) / 1048576, 1) & " Mb)  of picture"
    End If
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    Dim sFile As String
        
    common.InitDir = "C:\"
    common.ShowOpen
    sFile = common.FileName
    If Not Exists(sFile) Then Exit Sub
    ReDim fArray(FileLen(sFile) - 1) As Byte
    
    Dim fFile As Integer
    fFile = FreeFile
    
    Open sFile For Binary As fFile
    Get #fFile, , fArray
    Close fFile
    If Err Then
        FileLoaded = 0
        MsgBox "Can't load file" & String(2, vbNewLine) & Err.Number & ":" & Err.Description, vbCritical, "Load Error"
    Else
        FileLoaded = 1
        lblStatus.Caption = "Loaded " & FormatUsage(UBound(fArray) / 1024) & " Kb (" & Round(UBound(fArray) / 1048576, 1) & " Mb)  of data"
    End If
End Sub

Private Sub Command5_Click()
    If myPic Is Nothing Then
        MsgBox "Try load a picture first", vbCritical
    Else
        Unload Form3
        Form3.Show
        Me.Show
    End If
End Sub

Private Sub Command6_Click()
    ListAllProcesses Me.Text1
    lblStatus.Caption = "Memory usage list refreshed"
End Sub

Private Sub Command7_Click()
    If PersonCount = 0 Then lblStatus.Caption = "Arrays already cleared or never been created!": Exit Sub
    ReDim Persons(0 To 0)
    PersonCount = 0
    lblStatus.Caption = "Arrays cleared"
End Sub

Private Sub Command8_Click()
    If FormNo = 0 Then lblStatus.Caption = "Forms already unloaded or never get loaded!": Exit Sub
    Dim x As Integer
    For x = 1 To UBound(Forms)
        Unload Forms(x)
        Set Forms(x) = Nothing
    Next
    FormNo = 0
    lblStatus.Caption = "Forms unloaded"
End Sub

Private Sub Command9_Click()
    If Not FileLoaded Then lblStatus.Caption = "File already unloded or never get loaded!": Exit Sub
    ReDim fArray(0 To 0)
    FileLoaded = 0
    lblStatus.Caption = "File unloaded"
End Sub

Private Sub Form_Load()
    If InStr(1, App.EXEName, "RAM_usage", vbTextCompare) = 0 Then
        MsgBox "You should compile this project to .EXE and run it", vbInformation
        End
    End If
    ListAllProcesses Me.Text1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Please vote for me if you like it!!", vbExclamation
    End
End Sub

Private Sub Timer1_Timer()
    Dim myUsage As Double
    myUsage = RamUsage
    lblUsage.Caption = FormatUsage(myUsage) & "K (" & FormatUsage(myUsage / 1024) & " Mb)"
    myUsage = PFUsage
    lblPF.Caption = FormatUsage(myUsage) & "K (" & FormatUsage(myUsage / 1024) & " Mb)"
End Sub



