Attribute VB_Name = "ModActiveXReg"
Option Explicit

Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Enum RegOp
    Register = 1
    UnRegister
End Enum

Public Function RegisterActiveX(lzAxDll As String, mRegOption As RegOp) As Boolean
Dim mLib As Long, DllProcAddress As Long
Dim mThread
Dim sWait As Long
Dim mExitCode As Long
Dim lpThreadID As Long

    mLib = LoadLibrary(lzAxDll)
    
    If mLib = 0 Then RegisterActiveX = False: Exit Function
    
    If mRegOption = Register Then
        DllProcAddress = GetProcAddress(mLib, "DllRegisterServer")
    Else
        DllProcAddress = GetProcAddress(mLib, "DllUnregisterServer")
    End If
    
    If DllProcAddress = 0 Then
        RegisterActiveX = False
        Exit Function
    Else
        mThread = CreateThread(ByVal 0, 0, ByVal DllProcAddress, ByVal 0, 0, lpThreadID)
        
        If mThread = 0 Then
            FreeLibrary mLib
            RegisterActiveX = False
            Exit Function
        Else
            sWait = WaitForSingleObject(mThread, 10000)
            If sWait <> 0 Then
                FreeLibrary mLib
                mExitCode = GetExitCodeThread(mThread, mExitCode)
                ExitThread mExitCode
                Exit Function
            Else
                FreeLibrary mLib
                CloseHandle mThread
            End If
        End If
    End If
    
    RegisterActiveX = True
    
End Function
