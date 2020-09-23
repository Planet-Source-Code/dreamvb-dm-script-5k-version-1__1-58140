Attribute VB_Name = "ModExeHead"
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long

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

Private TFrameWork As Object ' dm++ framework object
Public ScriptAppliaction As ScriptPE

Function isAdmin() As Long
    ' Test to see if the user is of admin
    isAdmin = IsNTAdmin(ByVal 0&, ByVal 0&)
End Function

Function FixPath(lzPath As String) As String
   If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

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
    
    For i = LBound(sByte) To UBound(sByte)
        sByte(i) = sByte(i) Xor 159
    Next
    
    Encrypt = StrConv(sByte, vbUnicode)
    i = 0
    Erase sByte()
    
    If Err Then MsgBox "Error while decodeing code block", vbCritical, "Decode Error": End
    
End Function

Sub Main()
Dim StrBuffer As String, Code_Off_Set As Long, StrObjectFileBuff As String, sCommand As String, lzFile As String
Dim StrObjFileName  As String, MsgErr As String, MsgTitle As String, TheFileName As String
    On Error Resume Next
    
    MsgErr = "There was an error while loading the appliaction."
    MsgTitle = "Read Stream Error"
    TheFileName = FixPath(App.Path) & App.EXEName & ".exe"
    lzFile = FixPath(App.Path)
    
    Open TheFileName For Binary As #1
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
    
    Select Case ScriptAppliaction.ExeType
        Case 1 ' Framework exe
            ' Create the DM++ framework object
            Set TFrameWork = CreateObject("dmFramework.host")
            If Err.Number = -2147024770 Then
                MsgBox "Unable to locate appliaction Framework." _
                & vbCrLf & vbCrLf & "Please make sure the DM++ Framework is installed, on this computer before running this appliaction.", vbCritical, "File Not Found"
                GoTo CleanNow:
            Else
                TFrameWork.sc_File = ScriptAppliaction.Script_Source_Code
                TFrameWork.bScriptFile = TheFileName
                sCommand = Command$
                sCommand = Replace(sCommand, Chr(34), "")
                TFrameWork.bScriptCommand = sCommand
                TFrameWork.bScriptFile = lzFile
     
                
                TFrameWork.Execute
                
                If TFrameWork.CompileError Then
                    MsgBox TFrameWork.ErrorString, vbCritical, "DM++ Script"
                End If
                
                Set TFrameWork = Nothing
                GoTo CleanNow:
            End If
        Case 5 ' Standalone exe
            
        Case Else
            MsgBox MsgErr, vbCritical, MsgTitle
            GoTo CleanNow:
    End Select
    
CleanNow:
    TheFileName = ""
    sCommand = ""
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
    End
End Sub
