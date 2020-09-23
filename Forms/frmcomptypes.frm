VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmcomptypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compile Options"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   405
      Top             =   4785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Application Security"
      Height          =   1170
      Left            =   135
      TabIndex        =   12
      Top             =   3450
      Width           =   6675
      Begin VB.CheckBox chkAdmin 
         Caption         =   "Must have admin rights"
         Height          =   270
         Left            =   210
         TabIndex        =   15
         Top             =   705
         Width           =   1980
      End
      Begin VB.CheckBox chkEncrypt 
         Caption         =   "Encrypt Appliaction source"
         Height          =   270
         Left            =   225
         TabIndex        =   13
         Top             =   345
         Width           =   2235
      End
      Begin VB.Label lblInfoLink 
         AutoSize        =   -1  'True
         Caption         =   "Need More Info ?"
         MouseIcon       =   "frmcomptypes.frx":0000
         Height          =   195
         Index           =   4
         Left            =   2625
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Tag             =   "0"
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label lblInfoLink 
         AutoSize        =   -1  'True
         Caption         =   "Need More Info ?"
         Height          =   195
         Index           =   3
         Left            =   2625
         MouseIcon       =   "frmcomptypes.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Tag             =   "0"
         Top             =   390
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5565
      TabIndex        =   5
      Top             =   4725
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4035
      TabIndex        =   4
      Top             =   4725
      Width           =   1215
   End
   Begin VB.Frame Fraoutput 
      Caption         =   "Output"
      Height          =   1875
      Left            =   135
      TabIndex        =   7
      Top             =   1485
      Width           =   6675
      Begin VB.CommandButton cmdFolName 
         Caption         =   "...."
         Height          =   330
         Left            =   5925
         TabIndex        =   6
         Top             =   630
         Width           =   510
      End
      Begin VB.TextBox txtOutput 
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   630
         Width           =   5595
      End
      Begin VB.Label lblInfoLink 
         AutoSize        =   -1  'True
         Caption         =   "Need More Info ?"
         Height          =   195
         Index           =   2
         Left            =   5205
         MouseIcon       =   "frmcomptypes.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   315
         Width           =   1245
      End
      Begin VB.Label lblout 
         AutoSize        =   -1  'True
         Caption         =   "Output Location:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.Frame frmcomp 
      Caption         =   "Compile Options:"
      Height          =   1185
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   6675
      Begin VB.OptionButton OptCompType 
         Caption         =   "Create DM++ Standalon Appliaction."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Tag             =   "5"
         Top             =   795
         Width           =   2880
      End
      Begin VB.OptionButton OptCompType 
         Caption         =   "Create DM++ Script runtime appliaction. "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Tag             =   "1"
         Top             =   375
         Width           =   3120
      End
      Begin VB.Label lblInfoLink 
         AutoSize        =   -1  'True
         Caption         =   "Need More Info ?"
         Height          =   195
         Index           =   1
         Left            =   3495
         MouseIcon       =   "frmcomptypes.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label lblInfoLink 
         AutoSize        =   -1  'True
         Caption         =   "Need More Info ?"
         Height          =   195
         Index           =   0
         Left            =   3495
         MouseIcon       =   "frmcomptypes.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Tag             =   "0"
         Top             =   375
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmcomptypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ExeObjFileName As String

Sub UpdateInfoLabel(index, SkipUpDate As Boolean)
Dim X As Integer
    For X = 0 To lblInfoLink.Count - 1
        lblInfoLink(X).ForeColor = vbBlack
        lblInfoLink(X).FontUnderline = vbBlack
    Next
    X = 0

    If SkipUpDate Then Exit Sub
    lblInfoLink(index).ForeColor = vbBlue
    lblInfoLink(index).FontUnderline = True
End Sub

Sub UpAndShowInfoHelp(ID As Integer, Button)
On Error Resume Next
Dim A As String

    If Button <> vbLeftButton Then Exit Sub
    A = InfoHelper.InfoTopics(ID)
    A = Replace(A, "\n", vbCrLf)
    
    If InfoHelper.HasError Or Len(A) = 0 Then
        MsgBox "There was an error while finding the topic", vbCritical, "Topic Not Found"
        Exit Sub
    Else
        frmtip.lblInfo.Caption = A
        frmtip.Show vbModal, frmcomptypes
        A = ""
    End If
End Sub

Private Sub chkAdmin_Click()
    ScriptAppliaction.Security.IsAdminExe = chkAdmin
End Sub

Private Sub chkEncrypt_Click()
    ScriptAppliaction.Security.ExeEncrypted = chkEncrypt
End Sub

Private Sub cmdCancel_Click()
    txtOutput.Text = ""
    Unload frmcomptypes
End Sub

Private Sub cmdFolName_Click()
On Error GoTo SaveError:
    With CDialog
        .Filename = ""
        .CancelError = True
        .DialogTitle = "Make"
        .Filter = "EXE Files(*.exe)|*.exe|"
        .ShowSave
        If Len(.Filename) = 0 Then Exit Sub
        txtOutput.Text = .Filename
        ExeObjFileName = RemoveFileExt(.FileTitle) & "o"
    End With
    
SaveError:
    If Err = cdlCancel Then Err.Clear
    
End Sub

Private Sub cmdok_Click()
Dim SrcBuff As String, StrResponseFile As String, ObjLinker As Object

    On Error Resume Next
    
    If Not IsFileHere(TFrameWork.Framework_Path & "Link.dll") Then
        ' check that the linker is found
        MsgBox "Link.dll was not found." _
        & vbCrLf & "Please reinstall the appliaction and try agian.", vbInformation, "DM++ Framework missing component"
        Exit Sub
    Else
        'create the linker object
        Set ObjLinker = CreateObject("dmLink.clsLinker")
        If Err Then MsgBox Err.Description: Exit Sub
    End If
    
    If Len(Trim(txtOutput.Text)) = 0 Then
        MsgBox "You must Specify were your compiled appliaction will be saved to.", vbInformation, "Filename Not Found"
        Exit Sub
    End If
    
    ScriptAppliaction.ScriptSIG = &H25424 ' Magic number
    SrcBuff = StripModules(TFrameWork.codebackbuff)

    If ScriptAppliaction.Security.ExeEncrypted Then
        ' Encrypt the source script if encryption is set
        SrcBuff = Encrypt(SrcBuff)
    Else
        SrcBuff = Replace(SrcBuff, vbCrLf, vbLf)
    End If

    LinkToObjectFile = FixPath(GetTempPathA) & ExeObjFileName
    ScriptAppliaction.Script_Source_Code = SrcBuff
    ' dump all our data to a object file to be linked up
    WriteObjFile LinkToObjectFile
    ' Build the response data
    StrResponseFile = "#DM++" & vbCrLf
    StrResponseFile = StrResponseFile & "/d:exe %" & ScriptAppliaction.ExeType & vbCrLf
    StrResponseFile = StrResponseFile & "/obj:" & LinkToObjectFile & vbCrLf
    StrResponseFile = StrResponseFile & "/icon:-1" & vbCrLf
    StrResponseFile = StrResponseFile & "/out:" & txtOutput.Text
    ' send the respose data to the linker object
    ObjLinker.ResponseData = StrResponseFile
    ObjLinker.LinkAll
    ' check for any linker errors
    If Len(ObjLinker.Error_CallBack) <> 0 Then
        MsgBox "Error while linking :" & vbCrLf & vbCrLf & ObjLinker.Error_CallBack, vbInformation, "Error"
    Else
        MsgBox "Your appliaction has now been compiled to:" & vbCrLf _
        & txtOutput.Text
    End If
    
    If IsFileHere(LinkToObjectFile) Then
        Kill LinkToObjectFile ' kill the object file
    End If
    
    Unload frmcomptypes
    LinkToObjectFile = ""
    StrResponseFile = ""
    SrcBuff = ""
    Set ObjLinker = Nothing
    
End Sub

Private Sub Form_Load()
    frmcomptypes.Icon = Nothing
    OptCompType_Click 0
    chkEncrypt_Click
    chkAdmin_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmcomp_MouseMove Button, Shift, X, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmcomptypes = Nothing
    Set frmFind = Nothing
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmcomp_MouseMove Button, Shift, X, Y
End Sub

Private Sub Fraoutput_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmcomp_MouseMove Button, Shift, X, Y
End Sub

Private Sub frmcomp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateInfoLabel -1, True
End Sub

Private Sub lblInfoLink_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateInfoLabel index, False
End Sub

Private Sub lblInfoLink_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpAndShowInfoHelp index + 1, Button
End Sub

Private Sub OptCompType_Click(index As Integer)
    ScriptAppliaction.ExeType = Val(OptCompType(index).Tag)
End Sub
