Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Public clsTextBox As New txtClass

Private Type Plugins
    PlgCount As Integer
    PlgStr() As String
End Type

Private Type tConfig
    EditorBackColor As Long
    EditorForeColor As Long
    EditorTabSize As Integer
    EditorFontName As String
    EditorFontSize As Integer
    EditorMarginSize As Integer
End Type

Private Type InfoHelp
    HasError As Boolean
    InfoTopics() As String
End Type

Enum ProjType
    Script = 0
    ScriptExe = 1
    ScripDialog = 2
End Enum

Private Type HlpPak
    nSig As Long
    NoTopic As Integer
    nFuncName() As String
    nPageData() As String
End Type

Public t_Plugins As Plugins, Plg_IniFile As String
Private dmRefHelp As HlpPak

Public AppConfig As tConfig
Public InfoHelper As InfoHelp

Public TFrameWork As Object ' Link to the DM++ Framework Object
Public FrameWorkVer As String ' Version of the DM++ Framework

Public ProjectType As Integer
Public ButtonPressed As Integer

Public RefHelpFile As String, RefTempPage As String
Public AbsPathRoot As String
Public StartPagePath As String
Public WebTempPath As String

Public Sub LoadConfigData()
' Loads the apps configuration
    AppConfig.EditorFontName = GetSetting("DmScript", "Options", "FontName", "Courier")
    AppConfig.EditorFontSize = GetSetting("DmScript", "Options", "FontSize", "10")
    AppConfig.EditorForeColor = Val(GetSetting("DmScript", "Options", "ForeColor", &H80000012))
    AppConfig.EditorBackColor = Val(GetSetting("DmScript", "Options", "BackColor", &HFFFFFF))
    AppConfig.EditorTabSize = GetSetting("DmScript", "Options", "Indent", "4")
    AppConfig.EditorMarginSize = GetSetting("DmScript", "Options", "Margin", "5")
End Sub

Public Sub GlobalEditorUpDate()
    LoadConfigData
    ' This sub will update the editor window if it is open
    frmMain.txtCode.FontName = AppConfig.EditorFontName
    frmMain.txtCode.FontSize = AppConfig.EditorFontSize
    frmMain.txtCode.ForeColor = AppConfig.EditorForeColor
    frmMain.txtCode.BackColor = AppConfig.EditorBackColor
    clsTextBox.MarginSize = AppConfig.EditorMarginSize
End Sub

Public Function RemoveFileExt(lzFileName As String) As String
    RemoveFileExt = Mid(lzFileName, 1, InStr(1, lzFileName, ".", vbTextCompare))
End Function

Public Function GetTempPathA() As String
Dim RetVal As Long, TmpStr As String

    TmpStr = Space(216)
    RetVal = GetTempPath(Len(TmpStr), TmpStr)
    If RetVal <> 0 Then GetTempPathA = Left(TmpStr, RetVal)
    RetVal = 0: TmpStr = ""
    
End Function

Function FixPath(lzPath As String) As String
   If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function IsFileHere(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Function SaveFile(lzFile As String, FileData As String)
Dim iFile As Long
    iFile = FreeFile
    Open lzFile For Binary As #iFile
        Put #iFile, , FileData
    Close #iFile
    
    FileData = ""
End Function

Public Function OpenFile(Filename As String) As String
Dim iFile As Long, StrB As String
    iFile = FreeFile
    Open Filename For Binary As #iFile
        StrB = Space(LOF(iFile))
        Get #iFile, , StrB
    Close #iFile
    
    OpenFile = StrB
    StrB = ""
    
End Function

Function AddScriptBlock() As String
Dim StrA As String
    StrA = "Procedure Main();" & vbCrLf
    StrA = StrA & "// Add your code here" & vbCrLf & vbCrLf
    StrA = StrA & "End Main;" & vbCrLf
    AddScriptBlock = StrA
    StrA = ""
End Function

Function BuildHelpTopic(TopicID As String) As Integer
Dim nFile As Long, I As Integer, xFunName As String, TopicIdx As Integer
Dim BuffTemp As String
Dim StrBuff As String

    nFile = FreeFile
    Open RefHelpFile For Binary As #nFile
        Get #nFile, , dmRefHelp
    Close #nFile

    If dmRefHelp.nSig <> &HA0FFDBFF Then BuildHelpTopic = 0: Exit Function
    
    TopicIdx = -1
    For I = LBound(dmRefHelp.nFuncName) To UBound(dmRefHelp.nFuncName)
        ' serech all the function names
        xFunName = LCase(dmRefHelp.nFuncName(I))
        If xFunName = TopicID Then
            TopicIdx = I
            Exit For
        End If
    Next
    I = 0

    BuffTemp = OpenFile(RefTempPage) ' Open Temp holder HTML
    
    If TopicIdx = -1 Then
        StrBuff = dmRefHelp.nPageData(dmRefHelp.NoTopic)
        BuffTemp = Replace(BuffTemp, "<!--CODE -->", StrBuff)
        ' write new temp file
        WebTempPath = StartPagePath & "webtemp.htm"
        SaveFile WebTempPath, BuffTemp
        BuildHelpTopic = 1
        BuffTemp = "": StrBuff = "": xFunName = ""
        Erase dmRefHelp.nFuncName()
        Erase dmRefHelp.nPageData()
        dmRefHelp.NoTopic = -1
        dmRefHelp.nSig = 0
        Exit Function
    Else
        StrBuff = dmRefHelp.nPageData(TopicIdx)
        BuffTemp = Replace(BuffTemp, "<!--CODE -->", StrBuff)
        ' write new temp file
        WebTempPath = StartPagePath & "webtemp.htm"
        SaveFile WebTempPath, BuffTemp
        BuildHelpTopic = 1
        BuffTemp = "": StrBuff = "": xFunName = "": TopicIdx = 0
        Erase dmRefHelp.nFuncName()
        Erase dmRefHelp.nPageData()
        dmRefHelp.NoTopic = -1
        dmRefHelp.nSig = 0
        Exit Function
    End If

End Function

Public Function LoadInfoHelp(lzFile As String) As String
Dim nFile As Long, ipos As Integer, iCounter As Integer
Dim StrA As String

    InfoHelper.HasError = False
    If IsFileHere(lzFile) = False Then InfoHelper.HasError = True: Exit Function
    
    nFile = FreeFile
    Open lzFile For Input As #nFile
        Do While Not EOF(nFile)
            Input #nFile, StrA
            ipos = InStr(1, StrA, "=", vbTextCompare)
            StrA = Replace(StrA, "..", ",")
            If ipos <> 0 Then
                iCounter = iCounter + 1
                ReDim Preserve InfoHelper.InfoTopics(0 To iCounter)
                InfoHelper.InfoTopics(iCounter) = Right(StrA, Len(StrA) - 3)
            End If
            DoEvents
        Loop
    Close #nFile
    
    StrA = "": ipos = 0: iCounter = 0
    
End Function
