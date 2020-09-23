VERSION 5.00
Begin VB.Form frmConvert 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dec2Bin / Bin2Dec Convetor"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdabout 
      Caption         =   "About"
      Height          =   350
      Left            =   1635
      TabIndex        =   9
      Top             =   2475
      Width           =   960
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   350
      Left            =   2760
      TabIndex        =   8
      Top             =   2475
      Width           =   1000
   End
   Begin VB.CommandButton cmdconv 
      Caption         =   "Convert"
      Enabled         =   0   'False
      Height          =   350
      Left            =   420
      TabIndex        =   7
      Top             =   2475
      Width           =   1000
   End
   Begin VB.TextBox txtResult 
      Height          =   300
      Left            =   795
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1860
      Width           =   2910
   End
   Begin VB.Frame f1 
      Caption         =   "Options"
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   675
      Width           =   3660
      Begin VB.OptionButton op1 
         Caption         =   "Binary To Decimal"
         Height          =   270
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   615
         Width           =   1785
      End
      Begin VB.OptionButton op1 
         Caption         =   "Decimal To Binary"
         Height          =   270
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   285
         Value           =   -1  'True
         Width           =   1785
      End
   End
   Begin VB.TextBox txtValue 
      Height          =   300
      Left            =   780
      TabIndex        =   1
      Top             =   165
      Width           =   2910
   End
   Begin VB.Label l2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output:"
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   1913
      Width           =   525
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value:"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   218
      Width           =   450
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Op As Integer

Function Bin2Dec(BinNumber As String) As Long
Dim I As Long, Tmp As Long, sDec As Long
   ' Returns the integer value of a binary string
    Tmp = 1
    
    For I = Len(BinNumber) To 1 Step -1
        If CStr(Mid(BinNumber, I, 1)) = 1 Then
            sDec = sDec + Tmp
        End If
        Tmp = Tmp * 2
    Next
    
    Bin2Dec = sDec
    
End Function

Function Dec2Bin(IntDec As Long) As String
Dim sBinCheck As Boolean
Dim sBin As String
  ' Returns the binary string of an integer
    Do While IntDec <> 0
        sBinCheck = IntDec Mod 2
        If sBinCheck Then
            sBin = "1" & sBin
        Else
            sBin = "0" & sBin
        End If
        IntDec = IntDec \ 2
    Loop
    If sBin = "" Then
        Dec2Bin = "0"
    Else
        Dec2Bin = sBin
    End If
    sBin = ""
    
End Function

Private Sub cmdabout_Click()
    MsgBox "Decimal to Binary" & vbCrLf & "Binary to Decimal Convertor", vbInformation, frmConvert.Caption
End Sub

Private Sub cmdclose_Click()
    txtValue.Text = "": txtResult.Text = "": Op = 0
    Unload frmConvert
End Sub

Private Sub cmdconv_Click()
    If Op = 0 Then
        txtResult.Text = Dec2Bin(Val(txtValue.Text))
    Else
        txtResult.Text = Bin2Dec(Val(txtValue.Text))
    End If
End Sub

Private Sub Form_Load()
    op1_Click 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmConvert = Nothing
End Sub


Private Sub op1_Click(Index As Integer)
    Op = Index
End Sub

Private Sub txtValue_Change()
    cmdconv.Enabled = Trim(Len(txtValue.Text)) <> 0
End Sub
