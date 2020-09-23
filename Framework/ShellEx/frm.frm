VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DM++ Script Runner v1.0"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1380
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   390
      Left            =   2865
      TabIndex        =   1
      Top             =   870
      Width           =   795
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Always show error messages."
      Height          =   585
      Left            =   135
      TabIndex        =   0
      Top             =   615
      Width           =   2490
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   75
      Top             =   90
      Width           =   660
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    Unload Form1
End Sub

Private Sub Form_Load()
    Image1.Picture = Form1.Icon
    Check1.Value = Val(GetSetting("dmScript", "Main", "OnError", "1"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "dmScript", "Main", "OnError", CStr(Check1.Value)
    Set Form1 = Nothing
End Sub
