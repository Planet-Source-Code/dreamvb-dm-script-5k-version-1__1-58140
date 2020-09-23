VERSION 5.00
Begin VB.Form frmproj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Project"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   645
      Left            =   30
      ScaleHeight     =   585
      ScaleWidth      =   3825
      TabIndex        =   9
      Top             =   3360
      Width           =   3885
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   30
         TabIndex        =   10
         Top             =   60
         Width           =   3780
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4125
      TabIndex        =   5
      Top             =   1455
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4125
      TabIndex        =   4
      Top             =   885
      Width           =   1215
   End
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   780
      Width           =   1140
      _extentx        =   2011
      _extenty        =   53
   End
   Begin VB.PictureBox pictop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   5430
      TabIndex        =   1
      Top             =   0
      Width           =   5430
      Begin VB.Image imgicon 
         Height          =   300
         Left            =   150
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DM++ Script"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   690
         TabIndex        =   2
         Top             =   105
         Width           =   4245
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   30
      ScaleHeight     =   2415
      ScaleWidth      =   3810
      TabIndex        =   0
      Top             =   885
      Width           =   3870
      Begin VB.Label lblA 
         BackStyle       =   0  'Transparent
         Caption         =   "Dialog Script"
         Height          =   495
         Index           =   2
         Left            =   2445
         TabIndex        =   8
         Top             =   720
         Width           =   1020
      End
      Begin VB.Image img1 
         Height          =   480
         Index           =   2
         Left            =   2655
         Picture         =   "frmproj.frx":0000
         Top             =   165
         Width           =   480
      End
      Begin VB.Label lblA 
         BackStyle       =   0  'Transparent
         Caption         =   "Console App"
         Height          =   495
         Index           =   1
         Left            =   1215
         TabIndex        =   7
         Top             =   705
         Width           =   1020
      End
      Begin VB.Image img1 
         Height          =   480
         Index           =   1
         Left            =   1455
         Picture         =   "frmproj.frx":0346
         Top             =   135
         Width           =   480
      End
      Begin VB.Label lblA 
         BackStyle       =   0  'Transparent
         Caption         =   "Script File"
         Height          =   495
         Index           =   0
         Left            =   225
         TabIndex        =   6
         Top             =   720
         Width           =   1020
      End
      Begin VB.Image img1 
         Height          =   480
         Index           =   0
         Left            =   315
         Picture         =   "frmproj.frx":07EC
         Top             =   165
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmproj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ChangeFColor(index As Integer)
Dim I As Integer

    For I = 0 To lblA.Count - 1
        lblA(I).ForeColor = vbBlack
    Next
    I = 0
    If index = -1 Then Exit Sub
    lblA(index).ForeColor = vbBlue

End Sub

Sub ProjectTypeInfo(index As Integer)
    Select Case index
        Case 0
            lblInfo.Caption = "Create a new D++ script file."
        Case 1
            lblInfo.Caption = "Create a new D++ Console Project." & vbCrLf & "Not available in this version."
            cmdok.Enabled = False
        Case 2
            lblInfo.Caption = "Create a new D++ script Dialog Project." & vbCrLf & "Not available in this version."
            cmdok.Enabled = False
        Case Else
            lblInfo = ""
    End Select
End Sub

Private Sub cmdCancel_Click()
    ButtonPressed = 0
    Unload frmproj
End Sub

Private Sub cmdok_Click()
    ButtonPressed = 1
    Unload frmproj
End Sub

Private Sub Form_Load()
    imgicon.Picture = frmMain.Icon
End Sub

Private Sub Form_Resize()
    Line3D1.Width = frmproj.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmproj = Nothing
End Sub

Private Sub img1_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdok.Enabled = True
    img1(index).Top = img1(index).Top + 40
    ProjectTypeInfo index
End Sub

Private Sub img1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ChangeFColor index
End Sub

Private Sub img1_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    img1(index).Top = img1(index).Top - 40
    ProjectType = index
End Sub

Private Sub lblA_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    img1_MouseDown index, Button, Shift, 0, 0
End Sub

Private Sub lblA_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    img1_MouseMove index, Button, Shift, 0, 0
End Sub

Private Sub lblA_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    img1_MouseUp index, Button, Shift, 0, 0
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ChangeFColor -1
End Sub
