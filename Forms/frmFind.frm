VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Text"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtfind 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   915
      TabIndex        =   0
      Top             =   172
      Width           =   3060
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "&Find"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4125
      TabIndex        =   1
      Top             =   135
      Width           =   885
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4125
      TabIndex        =   2
      Top             =   675
      Width           =   885
   End
   Begin VB.CheckBox chkmatch 
      Caption         =   "Match Case"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   975
      Width           =   1320
   End
   Begin VB.Label lblfind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find Text:"
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   225
      Width           =   705
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pos As Integer

Private Sub cmdCancel_Click()
    txtfind.Text = "" ' clear the contents of the textbox
    Unload frmFind  ' unload the form
End Sub

Private Sub cmdfind_Click()
Dim Compare As Integer

    If chkmatch Then Compare = 0 Else Compare = 1
    
    Pos = InStr(Pos + 1, clsTextBox.Text, txtfind.Text, Compare)
    If Pos > 0 Then
        clsTextBox.SelStart = (Pos - 1)
        clsTextBox.SelLength = Len(txtfind.Text)
        clsTextBox.SetFocus
    Else
        MsgBox "The string " & Chr(34) & txtfind.Text & Chr(34) & " was not found.", vbExclamation, frmFind.Caption
    End If
    
    Compare = 0
    ipos = 0
End Sub

Private Sub Form_Load()
    frmFind.Icon = Nothing ' Remove the forms icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFind = Nothing
End Sub

Private Sub txtfind_Change()
    cmdfind.Enabled = Len(Trim(txtfind.Text)) <> 0
End Sub

