VERSION 5.00
Begin VB.Form frmAddins 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/ Remove Add-ons"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   350
      Left            =   2985
      TabIndex        =   4
      Top             =   1530
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "R&emove"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2985
      TabIndex        =   3
      Top             =   1035
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   350
      Left            =   2985
      TabIndex        =   2
      Top             =   510
      Width           =   1215
   End
   Begin VB.ListBox lstAddons 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   2565
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Add-ons"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   225
      Width           =   1140
   End
End
Attribute VB_Name = "frmAddins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vStr As Variant

Public Sub UpDateList()
On Error Resume Next
    Dim X As Integer
    lstAddons.Clear
    
    
    For X = LBound(t_Plugins.PlgStr) To UBound(t_Plugins.PlgStr)
        vLst = Split(t_Plugins.PlgStr(X), ";")
        lstAddons.AddItem vLst(1)
    Next
    
    Erase vLst
    X = 0
End Sub

Private Sub cmdAdd_Click()
    frmInst.Show vbModal, frmAddins
End Sub

Private Sub cmdClose_Click()
    lstAddons.Clear
    Unload frmAddins
End Sub

Private Sub cmdRemove_Click()
Dim S As String, X As Integer, nFile As Long, StrB As String

    If MsgBox("Do you want to remove the refernce for this add-on?", vbYesNo Or vbQuestion, "Remove Add-on") = vbNo Then Exit Sub
    t_Plugins.PlgStr(lstAddons.ListIndex) = "": cmdRemove.Enabled = False
    
    ' save the addons list
    For X = LBound(t_Plugins.PlgStr) To UBound(t_Plugins.PlgStr)
        If Len(t_Plugins.PlgStr(X)) > 0 Then
            S = S & t_Plugins.PlgStr(X) & vbCrLf
        End If
    Next
    
    lstAddons.RemoveItem lstAddons.ListIndex

    X = 0
    
    StrB = "#DM++ Script Add-ons Configuration." & vbCrLf
    StrB = StrB & vbCrLf
    StrB = StrB & S
    
    nFile = FreeFile
    Open Plg_IniFile For Output As #nFile
        Print #nFile, StrB;
    Close #nFile
    
    StrB = "": S = ""
    t_Plugins.PlgCount = -1
    frmMain.ProcessAddonMenu
    
End Sub

Private Sub Form_Load()
    If t_Plugins.PlgCount = -1 Then Exit Sub
    UpDateList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAddins = Nothing
End Sub

Private Sub lstAddons_Click()
    cmdRemove.Enabled = True
End Sub
