Welcome
Running DM++ Script with VBScript.
Add the code in to a new .vbs file

' Cut From Here
Dim dmScript
Set dmScript = CreateObject("dmFramework.host")
dmScript.ScriptFile = "Name-and-path-of-ScriptFile"
dmScript.Execute

If dmScript.CompileError Then
   MsgBox dmScript.ErrorString, vbCritical, "DM++ Script"
End If
