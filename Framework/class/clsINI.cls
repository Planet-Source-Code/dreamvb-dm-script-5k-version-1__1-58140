VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public iniFile As String
Private m_ReturnCode As Integer

#If Win32 Then
    ' Profile String functions:
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
    ' Profile String functions:
    Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
#End If

Private Function CheckKey(Selection As String, KeyName As String) As String
Dim S As String, I As Long
    S = Space(255)
    I = GetPrivateProfileString(Selection, KeyName, "ERR", S, 255, iniFile)
    CheckKey = Trim(Left(S, I))
    S = ""
    I = 0
End Function
Public Function CheckIni() As Long
    CheckIni = 1
    If IsFileHere(iniFile) = 0 Then
        CheckIni = 0
        Exit Function
    End If
    
    If GetAttr(iniFile) = 33 Then
        CheckIni = 0
        Exit Function
    End If
End Function

Public Function INIReadKeyValue(ByVal Selection As String, ByVal KeyName As String, Optional ByVal DefaultKey As String) As String
Dim Strbuff As String, iRet As Long

    Strbuff = Space(256) ' create a buffer
    iRet = GetPrivateProfileString(Selection, KeyName, DefaultKey, Strbuff, 256, iniFile)
    If iRet > 0 Then
        INIReadKeyValue = Left(Strbuff, iRet)
        iRet = 0
        Strbuff = ""
    End If
End Function

Public Function INIWriteKeyValue(ByVal Selection As String, ByVal KeyName As String, sKeyValue As String) As Long
    INIWriteKeyValue = WritePrivateProfileString(Selection, KeyName, sKeyValue, iniFile)
End Function

Public Function INIDeleteKeyValue(ByVal Selection As String, ByVal KeyName As String) As Long
    Dim S As String
    S = Space(255)
    If CheckIni = 0 Then
        INIDeleteKeyValue = 0
        Exit Function
    End If
    
    If Not CheckKey(Selection, KeyName) <> "ERR" Then
        INIDeleteKeyValue = 0
        Exit Function
    End If
    
    INIDeleteKeyValue = WritePrivateProfileString(Selection, KeyName, "", iniFile)
End Function

Public Function INIDeleteKey(ByVal Selection As String, ByVal KeyName As String) As Long
    Dim S As String
    S = Space(255)
    If CheckIni = 0 Then
        INIDeleteKey = 0
        Exit Function
    End If
    
    If Not CheckKey(Selection, KeyName) <> "ERR" Then
        INIDeleteKey = 0
        Exit Function
    End If
    
    INIDeleteKey = WritePrivateProfileString(Selection, KeyName, 0&, iniFile)
End Function

Function INIDeleteSelection(ByVal Selection As String) As Long
    INIDeleteSelection = WritePrivateProfileString(Selection, 0&, 0&, iniFile)
End Function
