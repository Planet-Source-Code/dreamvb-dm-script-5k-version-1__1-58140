' This script will register the DM++ Script framework on your PC.

Dim sCmd, WshShell, fso, Ans, Installed
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")

Function GetPath()
    Dim path
    path = WScript.ScriptFullName
    GetPath = Left(path, InStrRev(path, "\"))
End Function

Sub Main()
On Error Resume Next
    
    Installed = WshShell.RegRead("HKLM\Software\DmScript\Framework\Installed")
    If Installed = 1 Then
        Ans = MsgBox("DM ++ framework has already been installed on this system." _
        & vbCrLf & vbCrLf & "Whould like you to reinstall agian ?", vbYesNo Or vbQuestion)
        If Ans = vbNo Then Exit Sub
    End If
    
    If Not fso.FileExists(GetPath & "dmFramework.dll") Then
        MsgBox "Unable to locate dmFramework.dll", vbCritical, "Register"
        Exit Sub
    Else
        sCmd = "regsvr32.exe dmFramework.dll /S"
        WshShell.Run sCmd
	sCmd = "regsvr32.exe Link.dll /S"
	WshShell.Run sCmd
        WshShell.RegWrite "HKLM\Software\DmScript\Framework\Host", GetPath & "dmFramework.dll"
        Ans = MsgBox("Whould you like to Associate DM++Script Files with Windows ?", vbYesNo Or vbQuestion)
        
        If Ans = vbNo Then Exit Sub
        
        If Not fso.FileExists(GetPath & "dmScript.exe") Then
            MsgBox "Unable to locate dmScript.exe", vbCritical, "Register"
            Exit Sub
        Else
            WshShell.RegWrite "HKCR\.dms\", "DMSFile"
	    WshShell.RegWrite "HKCR\.dms\Content Type", "text/plain"
            WshShell.RegWrite "HKCR\DMSFile\", "DM++ Script"
            WshShell.RegWrite "HKCR\DMSFile\DefaultIcon\", GetPath & "dmScript.exe,1"
            WshShell.RegWrite "HKCR\DMSFile\Shell\Open\Command\", GetPath & "dmScript.exe %1"
            WshShell.RegWrite "HKCR\DMSFile\Shell\Edit\Command\", "Notepad.exe %1"
	    WshShell.RegWrite "HKCR\DMSFile\Shell\Print\Command\", "Notepad.exe /p %1"
            WshShell.RegWrite "HKLM\Software\DmScript\Framework\Installed", "1"
            MsgBox "DM++ Framework has now been registered on your system." _
            & vbCrLf & vbCrLf & "You may need to restart your system for some setting to take effect.", vbInformation, "Register"
        End If
    End If
End Sub

Call Main

