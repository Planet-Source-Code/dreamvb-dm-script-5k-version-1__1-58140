VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3285
      TabIndex        =   4
      Top             =   1995
      Width           =   825
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   1755
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   7665
      TabIndex        =   0
      Top             =   0
      Width           =   7725
      Begin VB.Image Image1 
         Height          =   495
         Left            =   60
         Top             =   90
         Width           =   480
      End
      Begin VB.Label lblFVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4710
         TabIndex        =   5
         Top             =   285
         Width           =   105
      End
      Begin VB.Image Image2 
         Height          =   750
         Left            =   660
         Picture         =   "frmAbout.frx":0000
         Top             =   210
         Width           =   3825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2003 - 2005 Ben Jones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   4710
         TabIndex        =   3
         Top             =   1365
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "THIS PROGRAM IS FREEWARE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4710
         TabIndex        =   2
         Top             =   705
         Width           =   2520
      End
      Begin VB.Label lblIdeVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   720
         TabIndex        =   1
         Top             =   930
         Width           =   105
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    Unload frmAbout
End Sub

Private Sub Form_Load()
    Image1.Picture = frmMain.Icon
    lblIdeVer.Caption = "Ver " & App.Major & "." & App.Minor & "." & App.Revision
    lblFVer.Caption = "Framework Version " & FrameWorkVer
    frmAbout.Caption = Right(frmMain.mnuabout.Caption, Len(frmMain.mnuabout.Caption) - 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbout = Nothing
End Sub
