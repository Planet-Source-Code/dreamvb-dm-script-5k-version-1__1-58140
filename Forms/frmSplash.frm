VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr 
      Interval        =   1600
      Left            =   3555
      Top             =   2145
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Please Wait....."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   2460
      Width           =   2265
   End
   Begin VB.Label l3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3030
      TabIndex        =   3
      Top             =   990
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "++"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2190
      TabIndex        =   2
      Top             =   225
      Width           =   270
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   123
      X2              =   296
      Y1              =   45
      Y2              =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scripting language"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1815
      TabIndex        =   1
      Top             =   435
      Width           =   2160
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1830
      TabIndex        =   0
      Top             =   225
      Width           =   360
   End
   Begin VB.Shape Bottom 
      BackColor       =   &H009B6B38&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Left            =   0
      Top             =   2415
      Width           =   1215
   End
   Begin VB.Image imglogo 
      Height          =   735
      Left            =   255
      Top             =   300
      Width           =   825
   End
   Begin VB.Shape border 
      BorderColor     =   &H00000000&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    imglogo.Picture = frmMain.Icon
    border.Width = (frmSplash.ScaleWidth)
    border.Height = (frmSplash.ScaleHeight)
    Bottom.Top = (frmSplash.ScaleHeight - Bottom.Height)
    Bottom.Width = (frmSplash.ScaleWidth)
    l3.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSplash = Nothing
End Sub

Private Sub tmr_Timer()
    Unload frmSplash
    frmMain.Show
End Sub
