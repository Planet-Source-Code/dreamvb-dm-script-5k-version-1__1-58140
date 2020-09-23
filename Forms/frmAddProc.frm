VERSION 5.00
Begin VB.Form frmAddProc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Procedure"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton c2 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   3255
      TabIndex        =   5
      Top             =   750
      Width           =   1000
   End
   Begin VB.CommandButton c1 
      Caption         =   "OK"
      Height          =   360
      Left            =   3255
      TabIndex        =   4
      Top             =   165
      Width           =   1000
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   750
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   720
      TabIndex        =   1
      Top             =   210
      Width           =   2340
   End
   Begin VB.Label lb2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   810
      Width           =   405
   End
   Begin VB.Label lb1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   255
      Width           =   465
   End
End
Attribute VB_Name = "frmAddProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CboTmp As String

Function isAlpha(lpString As String) As Boolean
Dim I As Integer
    isAlpha = True
    
    For I = 1 To Len(lpString)
        If Not Mid(lpString, I, 1) Like "*[A-Za-z]*" Then
            isAlpha = False
            Exit For
        End If
    Next
    I = 0
    
End Function

Private Sub c1_Click()
Dim sProc As String

    If Len(Trim(txtName.Text)) = 0 Then
        MsgBox "Procedure name required", vbCritical, frmAddProc.Caption
        Exit Sub
    ElseIf isAlpha(txtName.Text) <> True Then
        MsgBox "Procedure name may only contain alpha characters A-Z a-z", vbCritical, frmAddProc.Caption
        Exit Sub
    ElseIf LCase(txtName.Text) = "procedure" Or LCase(txtName.Text) = "function" Then
        MsgBox "Invaild Procedure or Function name.", vbCritical, "Inavild Name"
        Exit Sub
    Else
        sProc = vbCrLf & CboTmp & " " & txtName.Text & "();" & vbCrLf & vbCrLf & "End " & txtName.Text & ";" & vbCrLf
        frmMain.txtCode.SelText = sProc
        sProc = ""
        c2_Click
    End If
    
End Sub

Private Sub c2_Click()
    txtName.Text = ""
    Unload frmAddProc
End Sub

Private Sub cboType_Change()
    cboType.Text = CboTmp
End Sub

Private Sub cboType_Click()
    CboTmp = cboType.Text
End Sub

Private Sub Form_Load()
    frmAddProc.Icon = Nothing
    cboType.AddItem "Procedure"
    cboType.AddItem "Function"
    cboType.ListIndex = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmAddProc = Nothing
End Sub
