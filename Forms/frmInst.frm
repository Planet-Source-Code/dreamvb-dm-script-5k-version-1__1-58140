VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInst 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New Add-on"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog TDialog 
      Left            =   240
      Top             =   2775
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   2790
      Width           =   915
   End
   Begin VB.CommandButton cmInst 
      Caption         =   "&Install"
      Height          =   375
      Left            =   1575
      TabIndex        =   8
      Top             =   2790
      Width           =   960
   End
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   165
      TabIndex        =   7
      Top             =   2595
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "...."
      Height          =   345
      Left            =   3075
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtDlllPath 
      Height          =   330
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2040
      Width           =   2880
   End
   Begin VB.TextBox txtClsName 
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   1215
      Width           =   3435
   End
   Begin VB.TextBox txtMnuText 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   450
      Width           =   3435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "myplug.main"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1845
      TabIndex        =   10
      Top             =   930
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Link to Add-on Dynamic Libary:"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   1725
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add-on Class Name: eg"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   930
      Width           =   1665
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Item Name:"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   195
      Width           =   1260
   End
End
Attribute VB_Name = "frmInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCan_Click()
    Unload frmInst
End Sub

Private Sub cmdOpen_Click()
On Error GoTo CanErr:
    With TDialog
        .Filename = ""
        .InitDir = AbsPathRoot & "Add-ons\"
        .DialogTitle = "Install Add-ons"
        .Filter = "Addon Files(*.dll)|*.dll|"
        .ShowOpen
        
        If Len(.Filename) = 0 Then Exit Sub
        txtDlllPath.Text = .Filename
        Exit Sub
    End With
CanErr:
    If Err = cdlCancel Then Err.Clear
    
End Sub

Private Sub cmInst_Click()
Dim xCount As Integer
On Error Resume Next
    If Len(Trim(txtMnuText.Text)) = 0 Then
        MsgBox "You need to enter some text for the add-ons menu text", vbInformation, frmInst.Caption
        Exit Sub
    ElseIf Len(Trim(txtClsName.Text)) = 0 Then
        MsgBox "You need to include the Class name of your add-on eg Calulator.main", vbInformation, frmInst.Caption
        Exit Sub
    ElseIf Len(txtDlllPath.Text) = 0 Then
        MsgBox "You need to select your add-on filename", vbInformation, frmInst.Caption
        Exit Sub
    Else
        xCount = UBound(t_Plugins.PlgStr) + 1
        ReDim Preserve t_Plugins.PlgStr(xCount)
        t_Plugins.PlgStr(xCount) = "plgRef ;" & txtMnuText.Text & ";" & txtClsName.Text & ";" & txtDlllPath.Text
        xCount = 0
    End If
    
    For xCount = LBound(t_Plugins.PlgStr) To UBound(t_Plugins.PlgStr)
        If Len(t_Plugins.PlgStr(xCount)) > 0 Then
            S = S & t_Plugins.PlgStr(xCount) & vbCrLf
        End If
    Next
    
    StrB = "#DM++ Script Add-ons Configuration." & vbCrLf
    StrB = StrB & vbCrLf
    StrB = StrB & S
    
    nFile = FreeFile
    Open Plg_IniFile For Output As #nFile
        Print #nFile, StrB;
    Close #nFile
    
    StrB = ""
    S = ""
    
    MsgBox "The new add-on has now been installed.", vbInformation, frmInst.Caption
    frmAddins.UpDateList
    frmMain.ProcessAddonMenu
    Unload frmInst
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmInst = Nothing
End Sub
