Attribute VB_Name = "ModExe2"
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Type DMSecurity
    ExeEncrypted As Boolean
    IsAdminExe As Boolean
End Type

Private Type ScriptPE
    ScriptSIG As Long ' Indentier for the script &H25424
    ExeType As Integer ' 1 = framework runtime 5 Standalone
    Security As DMSecurity ' Security if any is assigned see DMSecurity
    Script_Source_Code As String ' Main script source of the appliaction
End Type

Private sc_File As String
Public ScriptAppliaction As ScriptPE

Public Sub DmsRun()
Dim StrMain As String, MainFound As Boolean

    MainFound = ProcedureExists("main") ' Allways check for the main function
    
    If MainFound = False Then
        LastError 13, 0, "Required: Procedure Main();"
    End If
    
    StrMain = GetProcedureCode("main") ' Get the code from the main function
    ExecuteCode StrMain ' Execute the code
    
End Sub

Public Sub AddCode(ExpCode As String)
Dim ThisCode As String, StrLibs As String
    
    Reset
    ThisCode = "": StrLibs = "": CodeBackBuff = ""
    ThisCode = ExpCode
    StrLibs = GetModules(ThisCode)
    ThisCode = StrLibs & ThisCode
    ThisCode = FormatStr(ThisCode)
    GetEnums ThisCode
    ThisCode = RemoveComments(ThisCode)
    GetProcedures ThisCode
    
End Sub

Public Sub ExecuteEx()
    AddCode sc_File
    If FoundError Then Exit Sub
    DmsRun
End Sub

Private Function GetTempPathA() As String
Dim TmpResult As Long, TmpStrBuff As String
    TmpStrBuff = Space(216)
    TmpResult = GetTempPath(Len(TmpStrBuff), TmpStrBuff)
    If TmpResult <> 0 Then GetTempPathA = Left(TmpStrBuff, TmpResult)
    TmpResult = 0: TmpStrBuff = ""
End Function

Private Function WriteFile(lzObjFileName As String, fData As String)
Dim nFile As Long
    nFile = FreeFile

    Open lzObjFileName For Binary As #nFile
        Put #nFile, , fData
    Close #nFile
    
End Function

Private Function OpenObjectFile(dwFile As String) As Integer
Dim iFile As Long
On Error Resume Next

    iFile = FreeFile
    OpenObjectFile = 1
    Open dwFile For Binary As #iFile
        Get #iFile, , ScriptAppliaction
    Close #iFile
    
    Exit Function
    
    If Err Then
        OpenObjectFile = 0
        Err.Clear
    End If
    
End Function

Function Encrypt(lzStr As String) As String
Dim sByte() As Byte
On Error Resume Next
' small simple function to encrypt / decrypt text
    sByte() = StrConv(lzStr, vbFromUnicode)
    
    For I = LBound(sByte) To UBound(sByte)
        sByte(I) = sByte(I) Xor 159
    Next
    
    Encrypt = StrConv(sByte, vbUnicode)
    I = 0
    Erase sByte()
    
    If Err Then MsgBox "Error while decodeing code block", vbCritical, "Decode Error": End
    
End Function

Sub Main()
Dim StrBuffer As String, Code_Off_Set As Long, StrObjectFileBuff As String
Dim StrObjFileName  As String, MsgErr As String, MsgTitle As String, lzFile As String
    On Error Resume Next
    
    MsgErr = "There was an error while loading the appliaction."
    MsgTitle = "Read Stream Error"
    
    scriptFullPath = FixPath(App.Path)
    lzFile = scriptFullPath & App.EXEName & ".exe"
    
    Open lzFile For Binary As #1
        StrBuffer = Space(LOF(1))
        Get #1, , StrBuffer
    Close #1

    Code_Off_Set = InStr(1, StrBuffer, "[DM]" & Chr(241), vbTextCompare)
    
    If Code_Off_Set = 0 Then
        MsgBox MsgErr, vbCritical, MsgTitle
        GoTo CleanNow:
    End If
    
    StrObjectFileBuff = Mid(StrBuffer, Code_Off_Set + 5, Len(StrBuffer)) ' Object file with it's data
    StrBuffer = "" ' finsihed with this now
    StrObjFileName = FixPath(GetTempPathA) & App.EXEName & ".o"  ' Object file to open
    WriteFile StrObjFileName, StrObjectFileBuff
    
    If OpenObjectFile(StrObjFileName) <> 1 Then
        MsgBox MsgErr, vbCritical, MsgTitle
        GoTo CleanNow:
    End If
    
    Kill StrObjFileName ' we don't need this anymore
    StrObjectFileBuff = "": StrObjFileName = ""
    ' check that we have a vaild sig
    If ScriptAppliaction.ScriptSIG <> &H25424 Then
        MsgBox MsgErr, vbCritical, MsgTitle
        GoTo CleanNow:
    End If
    ' is the script encryped
    If ScriptAppliaction.Security.ExeEncrypted Then
        ScriptAppliaction.Script_Source_Code = Encrypt(ScriptAppliaction.Script_Source_Code)
    Else
        ScriptAppliaction.Script_Source_Code = Replace(ScriptAppliaction.Script_Source_Code, vbLf, vbCrLf)
    End If
    ' is the script only runnable by admins
    If (ScriptAppliaction.Security.IsAdminExe) And isAdmin = 0 Then
        MsgBox "You need to be an administrator to run this application." _
        & vbCrLf & vbCrLf & "Please contact your local administrator for more details.", vbInformation, "Operation denined"
        GoTo CleanNow:
    End If
    ' Now we can start to run the app
    
    sc_File = ScriptAppliaction.Script_Source_Code

    StrCommandLine = Command$
    StrCommandLine = Replace(StrCommandLine, Chr(34), "")
    ExecuteEx
    
    If FoundError Then
        MsgBox ErrorStr, vbCritical, "DM++ Script"
        GoTo CleanNow:
    End If
    
CleanNow:
    Reset
    MsgTitle = ""
    StrObjFileName = ""
    MsgErr = ""
    StrBuffer = ""
    Code_Off_Set = 0
    ScriptAppliaction.ExeType = 0
    ScriptAppliaction.Script_Source_Code = ""
    ScriptAppliaction.ScriptSIG = 0
    ScriptAppliaction.Security.ExeEncrypted = False
    ScriptAppliaction.Security.IsAdminExe = False
    CommandParms = ""
    sc_File = ""
    End
End Sub
