Attribute VB_Name = "modReg"
' Registry module for reading and writeing to the Registry made for DM++ Script
' at moment only returns and saves DWORDS and STRINGS

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Const KEY_QUERY_VALUE As Long = &H1
Private Const ERROR_SUCCESS As Long = 0&

Enum KeyType
    REG_EXPAND_SZ = 2
    REG_DWORD = 4
End Enum

Enum RegDelete
    REG_KEY = 1
    REG_VALUE = 2
End Enum

Enum tHKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

Function RegKeyDelete(hKey As tHKey, keyPath As String) As Long
    RegKeyDelete = RegDeleteKey(hKey, keyPath)
End Function

Function RegReadString(hKey As tHKey, keyPath As String, KeyName As String, RegType As KeyType) As Variant
Dim sBuffer As String, sBufferLen As Long, sRegKey As Long
Dim RegDWord As Long, tBinary() As Byte, TKeyType As KeyType

    TKeyType = RegType
    
    If RegOpenKeyEx(hKey, keyPath, 0&, KEY_QUERY_VALUE, sRegKey) <> ERROR_SUCCESS Then
        RegReadString = ""
        Exit Function
    ElseIf RegQueryValueEx(sRegKey, KeyName, 0, RegType, ByVal 0&, sBufferLen) <> ERROR_SUCCESS Then
        RegReadString = ""
        Exit Function
    Else
        Select Case TKeyType
            Case REG_EXPAND_SZ
                sBuffer = Space(sBufferLen - 1)
                RegQueryValueEx sRegKey, KeyName, 0&, TKeyType, ByVal sBuffer, sBufferLen
            Case REG_DWORD
                sBufferLen = 4
                RegQueryValueEx sRegKey, KeyName, 0&, TKeyType, RegDWord, sBufferLen
                sBuffer = RegDWord
        End Select
        RegReadString = sBuffer
    End If
    
    RegCloseKey sRegKey
    
    sBuffer = ""
    sBufferLen = 0
    RegDWord = 0
    sRegKey = 0
    
End Function

Public Function RegSaveValue(hKey As tHKey, keyPath As String, KeyName As String, RegType As KeyType, Optional sRegKeyData As Variant) As Long
Dim sRegKey As Long
Dim lResult As Long

    If RegCreateKey(hKey, keyPath, sRegKey) <> ERROR_SUCCESS Then
        RegSaveValue = 0
        Exit Function
    End If
    
    If RegType = REG_EXPAND_SZ Then
        lResult = RegSetValueEx(sRegKey, KeyName, 0, RegType, ByVal CStr(sRegKeyData), Len(sRegKeyData))
    End If
    
    If RegType = REG_DWORD Then
        lResult = RegSetValueEx(sRegKey, KeyName, 0&, RegType, CLng(sRegKeyData), 4)
    End If
    
    lResult = RegCloseKey(sRegKey)
    
    If lResult <> ERROR_SUCCESS Then
        RegSaveValue = 0
        Exit Function
    Else
        RegSaveValue = 1
    End If
    
End Function

Public Function RegDeleteValueEx(hKey As tHKey, keyPath As String, KeyValue As String) As Long
Dim sRegKey As Long
Dim lResult As Long

    If RegOpenKey(hKey, keyPath, sRegKey) <> ERROR_SUCCESS Then
        RegDeleteValueEx = 2
        Exit Function
    End If
    
    lResult = RegDeleteValue(sRegKey, KeyValue)
    RegDeleteValueEx = lResult
    
    RegCloseKey sRegKey
    
End Function

