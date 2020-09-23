VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test Example - Plug-in"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   450
      Left            =   4215
      TabIndex        =   3
      Top             =   2955
      Width           =   1095
   End
   Begin VB.TextBox txtCode 
      Height          =   2670
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   135
      Width           =   5625
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Framework Info"
      Height          =   450
      Left            =   2325
      TabIndex        =   1
      Top             =   2955
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Editor Text"
      Height          =   435
      Left            =   150
      TabIndex        =   0
      Top             =   2955
      Width           =   1980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' the public members are object from the DM++ IDE thay are the only members used
Public Editor As Object ' Gives assess to the Code Editor
Public DevIDE As Object ' Gives access to the DM++ IDE includeing editors forms, menus ets
Public DevFrameWork As Object ' Access to the DM++ Script Framework

Private Sub Command1_Click()
    txtCode.Text = Editor.Text ' get the code editors text
End Sub

Private Sub Command2_Click()
Dim s As String
Dim vList As Variant, I As Integer, sLst As String

    s = "DM++ Script Framework"
    s = s & "Version: " & DevFrameWork.version & vbCrLf
    s = s & "Install Path: " & DevFrameWork.Framework_Path & vbCrLf
    s = s & "InBuilt Functions: " & vbCrLf & vbCrLf
    
    vList = Split(DevFrameWork.Function_List, ",")
    
    For I = LBound(vList) To UBound(vList)
        If Len(vList(I)) > 0 Then
            sLst = sLst & vList(I) & vbCrLf
        End If
    Next
    
    s = s & sLst
    
    txtCode.Text = s
    
    I = 0: s = "": Erase vList: sLst = ""
    
End Sub

Private Sub Command3_Click()
    Unload Form1 ' unloads this form
End Sub


Private Sub Form_Load()

End Sub
