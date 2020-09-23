VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "DM++ Script Environment"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8955
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -15
      Top             =   4755
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1582
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F78
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":261C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":296E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3012
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicCodeEditor 
      BorderStyle     =   0  'None
      Height          =   1860
      Left            =   720
      ScaleHeight     =   1860
      ScaleWidth      =   1590
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   1590
      Begin VB.ListBox LstOutput 
         Height          =   645
         Left            =   15
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   1050
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   15
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   0
         Width           =   990
      End
   End
   Begin VB.PictureBox PicStartPage 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1290
      Left            =   720
      ScaleHeight     =   1290
      ScaleWidth      =   1440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   885
      Width           =   1440
      Begin SHDocVwCtl.WebBrowser WebV 
         Height          =   1155
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1335
         ExtentX         =   2355
         ExtentY         =   2037
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   5325
      Left            =   690
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   9393
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Start"
            Key             =   "START_PAGE"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6075
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9887
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Ln 0"
            TextSave        =   "Ln 0"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Col 0"
            TextSave        =   "Col 0"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   570
      Top             =   4635
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicBar 
      Height          =   3615
      Left            =   30
      ScaleHeight     =   3555
      ScaleWidth      =   510
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   510
      Width           =   570
   End
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   53
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_NEW"
            Object.ToolTipText     =   "New..."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_OPEN"
            Object.ToolTipText     =   "Open..."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_SAVE"
            Object.ToolTipText     =   "Save..."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_CUT"
            Object.ToolTipText     =   "Cut..."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_COPY"
            Object.ToolTipText     =   "Copy..."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_PASTE"
            Object.ToolTipText     =   "Paste..."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_FIND"
            Object.ToolTipText     =   "Find..."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_RUN"
            Object.ToolTipText     =   "Run..."
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.PictureBox p1 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   3135
         ScaleHeight     =   315
         ScaleWidth      =   2145
         TabIndex        =   11
         Top             =   15
         Visible         =   0   'False
         Width           =   2145
         Begin VB.Label lbCol 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Col 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   660
            TabIndex        =   13
            Top             =   45
            Width           =   360
         End
         Begin VB.Label lbLine 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ln 0 ,"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   75
            TabIndex        =   12
            Top             =   45
            Width           =   405
         End
      End
   End
   Begin Project1.Line3D Line3D2 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   450
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   53
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuopen 
         Caption         =   "O&pen"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnublank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save..."
      End
      Begin VB.Menu mnublank4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMakeExe 
         Caption         =   "Make"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnublank3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnucut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnucopy 
         Caption         =   "C&opy"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnupaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuselall 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnumessage 
         Caption         =   "Clear &Messages"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnufind 
         Caption         =   "&Find Text"
      End
      Begin VB.Menu mnuinsert 
         Caption         =   "&Insert"
         Begin VB.Menu mnuopcom 
            Caption         =   "Comm&ent"
         End
         Begin VB.Menu mnuendblock 
            Caption         =   "&End Block"
         End
         Begin VB.Menu mnuDate 
            Caption         =   "&Date"
         End
         Begin VB.Menu mnuTime 
            Caption         =   "&Time"
         End
      End
   End
   Begin VB.Menu mnucompile 
      Caption         =   "&Run"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuaddprodc 
         Caption         =   "&Add Procedure"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuoptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuAddons 
      Caption         =   "Add-Ons"
      Begin VB.Menu mnuAddRem 
         Caption         =   "&Add/ Remove Add-ons"
      End
      Begin VB.Menu mnublank5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuplg 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnucontents 
         Caption         =   "&Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuref 
         Caption         =   "&DM++ Function Reference..."
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About DM++ Script..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartPage As String
Dim Changed As Boolean
Private CanMakeExe As Boolean

Enum mSetupOption
    New_Project = 1
    Open_Project = 2
End Enum

Public Sub ProcessAddonMenu()
Dim StrA As String, nFile As Long, xCount As Integer
Dim vStr As Variant
On Error Resume Next
    
    Erase t_Plugins.PlgStr()
    
    For xCount = 0 To mnuplg.Count
        mnuplg(xCount).Caption = ""
        mnuplg(xCount).Visible = False
        Unload mnuplg(xCount)
    Next
    
    xCount = -1
    nFile = FreeFile
    If IsFileHere(Plg_IniFile) = False Then Exit Sub
    Open Plg_IniFile For Input As #nFile
        Do While Not EOF(nFile)
            Input #nFile, StrA
            StrA = Trim(StrA)

            If LCase(Left(StrA, 6)) = "plgref" Then
                xCount = xCount + 1
                ReDim Preserve t_Plugins.PlgStr(xCount)
                t_Plugins.PlgStr(xCount) = StrA
                t_Plugins.PlgCount = xCount
                vStr = Split(StrA, ";")
                Load mnuplg(xCount)
                mnuplg(xCount).Caption = vStr(1) ' add menu item name
                mnuplg(xCount).Visible = True
            End If
            DoEvents
        Loop
    Close #nFile
    
    StrA = ""
    Erase vStr

End Sub
Private Sub DisableItems(mDisable As Boolean)
    mnuselall.Enabled = mDisable
    mnumessage.Enabled = mDisable
    mnufind.Enabled = mDisable
    mnuinsert.Enabled = mDisable
    mnucut.Enabled = mDisable
    mnucopy.Enabled = mDisable
    mnupaste.Enabled = mDisable
    
    Toolbar1.Buttons(5).Enabled = mDisable
    Toolbar1.Buttons(6).Enabled = mDisable
    Toolbar1.Buttons(7).Enabled = mDisable
    Toolbar1.Buttons(8).Enabled = mDisable
    
End Sub

Private Sub EnableItems()
    mnusave.Enabled = True
    mnuStart.Enabled = True
    mnuselall.Enabled = True
    mnumessage.Enabled = True
    mnufind.Enabled = True
    mnuinsert.Enabled = True
    
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    Toolbar1.Buttons(8).Enabled = True
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(10).Enabled = True
    
End Sub
Private Sub OpenProject()

    If ButtonPressed = 0 Then Exit Sub
    mnuMakeExe.Visible = True
    mnublank3.Visible = True
    
    StatusBar1.Panels(1).Text = ""
    
    If Tab1.Tabs.Count >= 2 Then Exit Sub
    Tab1.Tabs.Add , "M_CODE", "Project1"
    Tab1.Tabs(2).Selected = True
    EnableItems
End Sub

Private Function KillWebTemp()
    WebTempPath = StartPagePath & "webtemp.htm"
    If IsFileHere(WebTempPath) <> False Then Kill WebTempPath
End Function

Private Sub CleanUp()
    KillWebTemp
    Set TFrameWork = Nothing
    Set clsTextBox = Nothing
    Set clsPlugIFace = Nothing
    Set txtClass = Nothing
    Changed = False
    FileCount = 0
    Plg_IniFile = ""
    SetupExeName = ""
    BuildOutDir = ""
    txtCode.Text = ""
    FrameWorkVer = ""
    RefHelpFile = ""
    AbsPathRoot = ""
    t_Plugins.PlgCount = 0
    Erase t_Plugins.PlgStr()
    Erase InfoHelper.InfoTopics()
    AppConfig.EditorBackColor = 0
    AppConfig.EditorFontName = ""
    AppConfig.EditorFontSize = 0
    AppConfig.EditorForeColor = 0
    AppConfig.EditorTabSize = 0
    LstOutput.Clear
    StartPagePath = ""
    Unload frmMain
End Sub

Private Sub DoOpen()
On Error GoTo DlgError
    With CDialog
        .CancelError = True ' Turn on error checking
        .DialogTitle = "Open DM Install Script."
        .Filter = "DM Small Installer Script(*.dms)|*.dms|"
        .InitDir = FixPath(App.Path)
        .ShowOpen ' show the save dialog
         
         If Not LCase(Right(.Filename, 3)) = "dms" Then
            MsgBox "This is not a valid DM Small Installer Script." _
            & vbCrLf & vbCrLf & "Please try selecting a different filename.", vbCritical, frmMain.Caption
            Exit Sub
        Else
            ButtonPressed = 1 ' Button open was pressed
            txtCode.Text = OpenFile(.Filename)
            Changed = False
            StatusBar1.Panels(1).Text = ""
        End If
        Exit Sub
DlgError:
        ButtonPressed = 0 ' Cancel button was pressed
        If Err = cdlCancel Then Err.Clear
    End With

End Sub

Private Sub DoSave()
On Error GoTo DlgError
    With CDialog
        .CancelError = True ' Turn on error checking
        .DialogTitle = "Save DM Install Script." ' dialog title
        .Filter = "DM Small Installer Script(*.dms)|*.dms|" ' Filter file type
        .ShowSave ' show the save dialog
         SaveFile .Filename, txtCode.Text ' Save the data in the editor
         Changed = False ' set Changed back to fase
         Exit Sub
DlgError:
        If Err = cdlCancel Then Err.Clear
    End With

End Sub

Private Sub DoBevel(PanIdx As Integer)
    If StatusBar1.Panels(PanIdx).Bevel = sbrRaised Then
        StatusBar1.Panels(PanIdx).Bevel = sbrInset
    Else
        StatusBar1.Panels(PanIdx).Bevel = sbrRaised
    End If
End Sub

Private Sub EnableEditMenu()
    ' menu items
    mnucut.Enabled = clsTextBox.EnableCutPaste ' cut menu command
    mnucopy.Enabled = clsTextBox.EnableCutPaste ' copy menu command
    mnupaste.Enabled = Not clsTextBox.IsClipEmpty ' paste menu command
    ' toolbar buttons
    Toolbar1.Buttons(5).Enabled = mnucut.Enabled ' cut button
    Toolbar1.Buttons(6).Enabled = mnucopy.Enabled ' copy button
    Toolbar1.Buttons(7).Enabled = mnupaste.Enabled ' paste button
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 20
            DoBevel 3
        Case 144
            DoBevel 2
         Case 45
            DoBevel 4
    End Select
End Sub

Private Sub Form_Load()
Dim nVer As Object
On Error Resume Next
'top = 885' left = 720
    t_Plugins.PlgCount = -1
    
    AbsPathRoot = FixPath(App.Path)
    Plg_IniFile = AbsPathRoot & "add-on.ini"

    frmMain.MousePointer = vbHourglass
    clsTextBox.TextBox = txtCode
    Changed = False ' Set the editors textbox chnaged state to False
    EnableEditMenu ' enable the edit menu
    
    Set TFrameWork = CreateObject("dmFramework.host")
    
    If Err.Number = -2147024770 Then
        MsgBox "Unable to locate DM++ Script Framework." _
        & vbCrLf & "Please make sure the DM++ Framework is installed.", vbCritical, "File Not Found"
        End
    End If

    FrameWorkVer = TFrameWork.Version ' Get the Framework Version
    LoadInfoHelp AbsPathRoot & "InfoTips.txt"
    
    StartPagePath = AbsPathRoot & "startpage\"
    RefTempPage = StartPagePath & "base.htm"
    RefHelpFile = StartPagePath & "help.ref"
    StartPage = StartPagePath & "startpage.html"
    
    WebV.Navigate StartPage
    
    mnusave.Enabled = False
    mnuStart.Enabled = False
    mnuselall.Enabled = False
    mnumessage.Enabled = False
    mnufind.Enabled = False
    mnuinsert.Enabled = False
    mnucut.Enabled = False
    mnucopy.Enabled = False
    mnupaste.Enabled = False
    
    Toolbar1.Buttons(3).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(8).Enabled = False
    Toolbar1.Buttons(9).Enabled = False
    Toolbar1.Buttons(10).Enabled = False

    KillWebTemp
    GlobalEditorUpDate
    ProcessAddonMenu
    frmMain.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
On Error Resume Next
    Line3D1.Width = frmMain.Width: Line3D2.Width = frmMain.Width
    PicBar.Height = (frmMain.Height - StatusBar1.Height - Toolbar1.Height - PicBar.Top - 360)

    Tab1.Width = (frmMain.Width - Tab1.Left - 180)
    Tab1.Height = PicBar.Height
    
    If Tab1.SelectedItem.Key = "START_PAGE" Then
        PicStartPage.Width = (Tab1.Width - 110)
        PicStartPage.Height = (Tab1.Height - Tab1.Top + 50)
        WebV.Width = PicStartPage.Width
        WebV.Height = PicStartPage.Height
    End If
    
    If Tab1.SelectedItem.Key = "M_CODE" Then
        PicCodeEditor.Width = (Tab1.Width - 110)
        PicCodeEditor.Height = (Tab1.Height - Tab1.Top + 50)
        txtCode.Width = PicCodeEditor.Width
        txtCode.Height = (PicCodeEditor.Height - LstOutput.Height - 80)
        LstOutput.Width = (PicCodeEditor.Width - LstOutput.Left)
        LstOutput.Top = (PicCodeEditor.Height - LstOutput.Height)
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbout = Nothing
    Set frmAddins = Nothing
    Set frmAddProc = Nothing
    Set frmcomptypes = Nothing
    Set frmFind = Nothing
    Set frmInst = Nothing
    Set frmoptions = Nothing
    Set frmproj = Nothing
    Set frmtip = Nothing
    Set frmMain = Nothing
    KillWebTemp
    End
End Sub
Private Sub mnuabout_Click()
    frmAbout.Show vbModal, frmMain
End Sub

Private Sub mnuaddprodc_Click()
    frmAddProc.Show vbModal, frmMain
End Sub

Private Sub mnuAddRem_Click()
    frmAddins.Show vbModal
End Sub

Private Sub mnucontents_Click()
    ShellExecute frmMain.hwnd, "open", AbsPathRoot & "help.chm", "", "", 1
End Sub

Private Sub mnucopy_Click()
    clsTextBox.EditMenu vsCOPY
End Sub

Private Sub mnucut_Click()
    clsTextBox.EditMenu vsCUT
End Sub

Private Sub mnuDate_Click()
    txtCode.SelText = Date
End Sub

Private Sub mnuendblock_Click()
    clsTextBox.SelText = ";"
End Sub

Private Sub mnuexit_Click()
Dim ans As Integer

    If Not Changed Then CleanUp
    
    ans = MsgBox("You have made changes to your work." _
    & vbCrLf & vbCrLf & "Would you like to save the changes now?", vbYesNo Or vbQuestion, frmMain.Caption)
    
    If ans = vbNo Then CleanUp: Exit Sub
    
    DoSave
    CleanUp
    
End Sub

Private Sub mnufind_Click()
    frmFind.Show , frmMain
End Sub

Private Sub mnuMakeExe_Click()
    If CanMakeExe <> True Then
        MsgBox "There have been changes made to your work." _
        & vbCrLf & "Please run the project, to check for errors before compileing.", vbInformation, frmMain.Caption
        Exit Sub
    Else
        frmcomptypes.Show vbModal, frmMain
    End If
    
End Sub

Private Sub mnumessage_Click()
    LstOutput.Clear
End Sub

Public Sub mnunew_Click()
Dim ans As Integer

    If Not Changed Or Len(txtCode.Text) = 0 Then
        SetupIDE New_Project
        Exit Sub
    End If
    
    ans = MsgBox("You have made chnages to your work." _
    & vbCrLf & vbCrLf & "Would you to like to save the chnages now?", vbYesNo Or vbQuestion, frmMain.Caption)
    
    If ans = vbNo Then
        SetupIDE New_Project
        Exit Sub
    Else
        DoSave
    End If
    
End Sub

Private Sub SetupIDE(mOption As mSetupOption)
    
    If mOption = Open_Project Then
        OpenProject
        Exit Sub
    End If
    
    frmproj.Show vbModal, frmMain
    If ButtonPressed = 0 Then Exit Sub

    txtCode.Text = ""
    StatusBar1.Panels(1).Text = ""
    
    Select Case ProjectType
        Case 0
            txtCode.Text = AddScriptBlock
            If Tab1.Tabs.Count >= 2 Then Exit Sub
            Tab1.Tabs.Add , "M_CODE", "Project1"
            Tab1.Tabs(2).Selected = True
            mnuMakeExe.Visible = True
            mnublank3.Visible = True
    End Select
    
    txtCode.SelStart = Len(txtCode.Text)
    txtCode.SetFocus
    EnableItems
End Sub

Private Sub mnuopcom_Click()
    clsTextBox.SelText = " //"
End Sub

Private Sub mnuopen_Click()
Dim ans As Integer

    If Not Changed Then
        DoOpen
        SetupIDE Open_Project
        Exit Sub
    End If
    
    ans = MsgBox("Your have made chnages to your work." _
    & vbCrLf & vbCrLf & "Do you want to save the chnages now?", vbYesNo Or vbQuestion, frmMain.Caption)
        
    If ans = vbNo Then
        DoOpen
        SetupIDE Open_Project
        Exit Sub
    Else
        DoSave
        DoOpen
        SetupIDE Open_Project
    End If

End Sub

Private Sub mnuoptions_Click()
    frmoptions.Show vbModal, frmMain
End Sub

Private Sub mnupaste_Click()
    clsTextBox.EditMenu vsPASTE
End Sub

Private Sub mnuplg_Click(index As Integer)
Dim vPlgInfo As Variant, sPlg As String
Dim clsPlg As New clsPlugIFace
Dim PlgObject As Object
On Error GoTo PlugErr

    vPlgInfo = Split(t_Plugins.PlgStr(index), ";")
    
    If UBound(vPlgInfo) < 3 Then
        Err.Number = -1
        GoTo PlugErr
        Exit Sub
    Else
        sPlg = CStr(vPlgInfo(3))
        RegisterActiveX sPlg, Register
        
        Set PlgObject = CreateObject(vPlgInfo(2))
        Set clsPlg.Dev_IDE_Window = frmMain
        Set clsPlg.Dev_Frame_Work = modMain.TFrameWork

        PlgObject.RunPlugin clsPlg

        Set PlgObject = Nothing
        Set clsPlg.Dev_IDE_Window = Nothing
        Set clsPlg.Dev_Frame_Work = Nothing
        Exit Sub
    End If
    
PlugErr:
    If Err Then
        MsgBox "Unable to load add-on" & vbCrLf & Err.Description, vbCritical, "Error_" & Err.Number
    End If
    
End Sub

Private Sub mnuref_Click()
    Tab1.Tabs(1).Selected = True
    WebV.Navigate StartPagePath & "functions.html"
End Sub

Private Sub mnusave_Click()
    DoSave
End Sub

Private Sub mnuselall_Click()
    clsTextBox.EditMenu vsSELALL
End Sub

Private Sub mnuStart_Click()
On Error Resume Next

    LstOutput.Clear ' Clear output box
    LstOutput.AddItem "Compileing..."
    TFrameWork.sc_File = "" ' Clear script buffer
    TFrameWork.sc_File = txtCode.Text ' Add script buffer from textbox
    TFrameWork.Execute ' Execute the code
    
    mnuMakeExe.Enabled = True
    CanMakeExe = mnuMakeExe.Enabled
    
    If TFrameWork.CompileError Then ' check for any errors
        Beep
        LstOutput.AddItem TFrameWork.ErrorString ' show error results
    End If
    
End Sub

Private Sub mnuTime_Click()
    txtCode.SelText = Time
End Sub

Private Sub Tab1_Click()
    Form_Resize
    mnuaddprodc.Enabled = False
    
    p1.Visible = (Tab1.SelectedItem.index - 1)
    StatusBar1.Panels(5).Visible = p1.Visible: StatusBar1.Panels(6).Visible = p1.Visible

    Select Case Tab1.SelectedItem.Key
        Case "START_PAGE"
            DisableItems False
            PicStartPage.Visible = True
            PicCodeEditor.Visible = False
        Case "M_CODE"
            DisableItems True
            EnableEditMenu
            PicStartPage.Visible = False
            PicCodeEditor.Visible = True
            mnuaddprodc.Enabled = True
            txtCode.SetFocus
    End Select
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "M_NEW"
            mnunew_Click
        Case "M_OPEN"
            mnuopen_Click
        Case "M_SAVE"
            mnusave_Click
        Case "M_CUT"
            mnucut_Click
        Case "M_COPY"
            mnucopy_Click
        Case "M_PASTE"
            mnupaste_Click
        Case "M_FIND"
            mnufind_Click
        Case "M_RUN"
            mnuStart_Click
        Case "M_STOP"
    End Select
End Sub

Private Sub txtCode_Change()
    CanMakeExe = False
    If Changed Then Exit Sub
    Changed = True
    StatusBar1.Panels(1).Text = "Modified"
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyTab Then
        txtCode.SelText = Space(AppConfig.EditorTabSize)
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbLine.Caption = "Ln " & clsTextBox.GetCurrentLineNumber & " ,"
    lbCol.Caption = "Col " & clsTextBox.GetCurrentLineLength
    
    StatusBar1.Panels(5).Text = Left(lbLine.Caption, Len(lbLine.Caption) - 1)
    StatusBar1.Panels(6).Text = lbCol.Caption

End Sub

Private Sub txtCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    EnableEditMenu
End Sub

Private Sub WebV_TitleChange(ByVal Text As String)
On Error Resume Next
Dim iStart, iEnd As Long, ipos As Long, RefPage As String, iReturn As Integer
    
   iStart = InStr(1, Text, "ref:", vbTextCompare)
   ipos = InStr(1, Text, "help:", vbTextCompare)
   iEnd = InStr(iStart + 1, Text, "(", vbTextCompare)
   
   Select Case UCase(Mid(Text, iStart + 4, iEnd - iStart - 4))
        Case "HOME"
            WebV.Navigate StartPage
        Case "OPENFILE"
            WebV.Navigate StartPage
            mnuopen_Click
        Case "NEWFILE"
            WebV.Navigate StartPage
            mnunew_Click
        Case "CLOSE"
            mnuexit_Click
   End Select
   iStart = 0
   iEnd = 0
   
   If ipos Then
        KillWebTemp
        RefPage = LCase(Trim(Mid(Text, ipos + 5, Len(Text) - ipos)))
        iReturn = BuildHelpTopic(RefPage)
        If (iReturn <> 0) Then WebV.Navigate WebTempPath
   End If
   
End Sub
