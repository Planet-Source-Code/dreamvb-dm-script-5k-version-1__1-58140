VERSION 5.00
Begin VB.Form frmtip 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quick Help"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3645
      TabIndex        =   3
      Top             =   2895
      Width           =   825
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2610
      Left            =   120
      ScaleHeight     =   2550
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   105
      Width           =   4335
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   195
         X2              =   4065
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   675
         TabIndex        =   2
         Top             =   195
         Width           =   675
      End
      Begin VB.Image img1 
         Height          =   345
         Left            =   120
         Picture         =   "frmtip.frx":0000
         Top             =   105
         Width           =   360
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   180
         TabIndex        =   1
         Top             =   810
         Width           =   4050
      End
   End
End
Attribute VB_Name = "frmtip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload frmtip
End Sub

