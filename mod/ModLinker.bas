Attribute VB_Name = "ModLinker"
Option Explicit

Private Type DMSecurity
    ExeEncrypted As Boolean
    IsAdminExe As Boolean
End Type

Private Type ScriptPE
    ScriptSIG As Long ' Indentier for the script &H25424
    ExeType As Integer ' 1 = framework runtime 5 = Standalone
    Security As DMSecurity ' Security if any is assigned see DMSecurity
    Script_Source_Code As String ' Main script source of the appliaction
End Type

Public ScriptAppliaction As ScriptPE
Public LinkToObjectFile As String

Function StripModules(CodeBufferEX As String) As String
Dim StrCodeB As String, vStr As Variant, StrLine As String, AddCr As Boolean, X As Long
    ' This used to strip away module names in the code
    StrCodeB = CodeBufferEX
    
    vStr = Split(StrCodeB, vbCrLf)
    StrCodeB = "" ' Clean it up
    For X = LBound(vStr) To UBound(vStr)
        StrLine = LCase(Trim(vStr(X)))
        StrLine = Replace(StrLine, vbTab, "")
        
        If Left(StrLine, 6) = "module" Then
            StrLine = ""
            AddCr = False
        Else
            AddCr = True
        End If
       
        If AddCr Then
            StrCodeB = StrCodeB & StrLine & vbCrLf
        End If
    Next
    
    X = 0
    StripModules = StrCodeB
    Erase vStr
    StrLine = ""
    StrCodeB = ""
    CodeBufferEX = ""
    
End Function
Function Encrypt(lzStr As String) As String
Dim sByte() As Byte, i As Integer
On Error Resume Next
    ' Function to encrypt / decrypt a string. using byte arrays faster than strings concat
    sByte() = StrConv(lzStr, vbFromUnicode)
    
    For i = LBound(sByte) To UBound(sByte)
        sByte(i) = sByte(i) Xor 159
    Next
    
    Encrypt = StrConv(sByte, vbUnicode)
    i = 0
    Erase sByte()
    
    If Err Then MsgBox "Error while decodeing code block", vbCritical, "Decode Error": End
    
End Function

Function WriteObjFile(lzObjFileName As String)
Dim nFile As Long
    nFile = FreeFile

    Open lzObjFileName For Binary As #nFile
        Put #nFile, , ScriptAppliaction
    Close #nFile
    
End Function

Function CleanPackagePE()
    ScriptAppliaction.ScriptSIG = 0
    ScriptAppliaction.ExeType = 0
    ScriptAppliaction.Security.ExeEncrypted = False
    ScriptAppliaction.Security.IsAdminExe = False
    ScriptAppliaction.Script_Source_Code = ""
End Function
