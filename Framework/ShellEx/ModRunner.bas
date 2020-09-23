Attribute VB_Name = "ModRunner"
Private Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public ScriptTimeOut As Integer
Private Function GetCommLine() As String
    Dim RetStr As Long, SLen As Long
    Dim Buffer As String
    'Get a pointer to a string, which contains the command line
    RetStr = GetCommandLine
    'Get the length of that string
    SLen = lstrlen(RetStr)
    If SLen > 0 Then
        'Create a buffer
        GetCommLine = Space$(SLen)
        'Copy to the buffer
        CopyMemory ByVal GetCommLine, ByVal RetStr, SLen
    End If
End Function

Sub Main()
Dim DM_Script As Object
Dim sCommand As String, FileExt As String * 3, ErrorState As Integer

    ErrorState = Val(GetSetting("dmScript", "Main", "OnError", "1"))
    sCommand = Command$ ' Get commmand line parms
    If Len(sCommand) = 0 Then Form1.Show: Exit Sub
    ' Remove char34 from left and right side of sCommand
    If Left(sCommand, 1) = Chr(34) Or Right(sCommand, 1) = Chr(34) Then
        sCommand = Right(sCommand, Len(sCommand) - 1)
        sCommand = Left(sCommand, Len(sCommand) - 1)
    End If
    
    FileExt = UCase(Mid(sCommand, InStrRev(sCommand, ".", Len(sCommand), vbTextCompare) + 1))
    ' Line above get's the file ext
    
    If Not FileExt = "DMS" Then ' check if this is a vaild DM Script file DMS
        MsgBox "This is not a vaild DM++ Script File", vbCritical, "File Not Supported"
        GoTo Endn:
    End If
    
    Set DM_Script = CreateObject("dmFramework.host") ' Create the DM++ script Framework
    DM_Script.ScriptFile = sCommand ' Script file to Execute
    DM_Script.Execute ' Execute and run the script file
    
    If DM_Script.CompileError Then
        ' inform if any errors have been found
        If ErrorState <> 0 Then ' check if use has error warrings turned on see form1
            MsgBox DM_Script.ErrorString, vbCritical, "DM++ Script"
            GoTo Endn:
        Else
            GoTo Endn:
        End If
    End If
    
    Set DM_Script = Nothing ' destroy framework object
    ErrorState = 0
    sCommand = ""
    FileExt = ""
    
    GoTo Endn:
    
Endn:
    End
    
End Sub
