VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function FlashWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function ShellExecuteA Lib "shell32.dll" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Function DriveType(ParmArr As Variant) As Long
On Error GoTo ErrorCheck:
Dim a As String
    
    a = ParmArr(0)

    DriveType = GetDriveType(CStr(ReturnData(a)))

ErrorCheck:
    If Err Then
        If Err Then LastError 13, CurrentLinePos, Err.Description
    End If
    
End Function

Public Function GetWindowsDir(ParmArr As Variant) As Long
On Error GoTo ErrorCheck:
Dim a As String, b As Long

    a = ParmArr(0)
    b = ReturnData(CStr(ParmArr(1)))
    
    If Not VariableExists(CStr(ParmArr(0))) Then
        GetWindowsDir -1
        Exit Function
    Else
        GetWindowsDir = GetWindowsDirectory(Modvariables.Variables.Variabledata(VariableIndex(CStr(ParmArr(0)))), b)
    End If
    
    a = ""
    b = 0
    Erase ParmArr
ErrorCheck:
    If Err Then
        If Err Then LastError 13, CurrentLinePos, Err.Description
    End If
End Function


Public Function ShellEx(ParmArr As Variant) As Long
    ShellEx = ShellExecute(ReturnData(CStr(ParmArr(0))), ReturnData(CStr(ParmArr(1))), _
    ReturnData(CStr(ParmArr(2))), ReturnData(CStr(ParmArr(3))), ReturnData(CStr(ParmArr(4))), _
    ReturnData(CStr(ParmArr(5))))
End Function
Public Function BeepEx(ParmArr As Variant) As Long
Dim a As Long, b As Long
On Error GoTo ErrorCheck:

    a = ReturnData(CStr(ParmArr(0)))
    b = ReturnData(CStr(ParmArr(1)))
    
    BeepEx = Beep(a, b)
    
ErrorCheck:
    If Err Then
        If Err Then LastError 13, CurrentLinePos, Err.Description
    End If
End Function


Public Function WinExe(ParmArr) As Long
On Error GoTo ErrorCheck:
Dim a As String, b As Long

    a = ReturnData(CStr(ParmArr(0)))
    b = ReturnData(CStr(ParmArr(1)))
    WinExe = WinExec(a, b)
    a = ""
    b = 0
    Erase ParmArr
ErrorCheck:
    If Err Then
        If Err Then LastError 13, CurrentLinePos, Err.Description
    End If
End Function

Public Function GetCompName(ParmArr) As Long
On Error GoTo ErrorCheck:
Dim a As String, b As Long

    a = ParmArr(0)
    b = ReturnData(CStr(ParmArr(1)))
    
    If Not VariableExists(CStr(ParmArr(0))) Then
        GetCompName -1
        Exit Function
    Else
        GetCompName = GetComputerName(Modvariables.Variables.Variabledata(VariableIndex(CStr(ParmArr(0)))), b)
    End If
    
    a = ""
    b = 0
    Erase ParmArr
ErrorCheck:
    If Err Then
        If Err Then LastError 13, CurrentLinePos, Err.Description
    End If
End Function

Public Function FlashWnd(ParmArr As Variant) As Long
Dim lzReturn As Long
'On Error GoTo ErrState

    lzReturn = FlashWindow(CLng(ReturnData(CVar(ParmArr(0)))), CLng(ReturnData(CVar(ParmArr(1)))))
    FlashWnd = lzReturn
    
End Function

Public Function GetWndText(ParmArr As Variant) As Long
On Error GoTo ErrorCheck:
Dim a As Long, c As Long, sBuff As String

    a = ReturnData(CStr(ParmArr(0)))
    b = ParmArr(1)
    c = ReturnData(CStr(ParmArr(2)))
    
    If Not VariableExists(CStr(ParmArr(1))) Then
        GetWndText = -1
        Exit Function
    Else
        GetWndText = GetWindowText(a, Modvariables.Variables.Variabledata(VariableIndex(CStr(ParmArr(1)))), c)
    End If
    
    a = 0
    c = 0
    Erase ParmArr
ErrorCheck:
    If Err Then
        If Err Then LastError 13, CurrentLinePos, Err.Description
    End If
    
End Function

Public Function SetWndText(ParmArr As Variant) As Long
On Error GoTo ErrorCheck:
Dim a As Long, b As String

    a = ReturnData(CStr(ParmArr(0)))
    b = ReturnData(CStr(ParmArr(1)))
    SetWndText = SetWindowText(a, b)
    
    a = 0
    c = 0
    Erase ParmArr
ErrorCheck:
    If Err Then
        If Err Then LastError 13, CurrentLinePos, Err.Description
    End If
    
End Function
