Attribute VB_Name = "Win32_API"
' this mod is were I added some API stuff

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
' Network API stuff
Public Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long
' Drives and files
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
' Broswe folder API
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

' Browse folder consts
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Enum EnumIni
    T_INIReadkeyValue = 1
    T_INIWriteKeyValue = 2
    T_INIDeleteKeyValue = 3
    T_INIDeleteSelection = 4
End Enum

Function DoINI(mParm As String, ParmCount As Integer, IniState As EnumIni) As Variant
Dim vINIList As Variant
Dim A As String, B As String, C As String, D As String

    If FoundError Then GoTo ErrorCheck: Exit Function

    vINIList = Split(mParm, ",")
    
    If UBound(vINIList) < ParmCount Then
        LastError 7, CurrentLinePos
        Exit Function
    End If
    
    A = ReturnData(CStr(vINIList(0))) ' Ini Filename
    B = ReturnData(CStr(vINIList(1))) ' Ini Selection
    
    If IniState = T_INIDeleteKeyValue Or IniState = T_INIReadkeyValue Or IniState = T_INIWriteKeyValue Then
        C = ReturnData(CStr(vINIList(2))) ' INI Key name
    End If
    
    If IniState = T_INIReadkeyValue Or IniState = T_INIWriteKeyValue Then
        D = ReturnData(CStr(vINIList(3))) ' INI Default return data or key value data
    End If
    
    TIniFile.iniFile = A ' Full path and file name of the INI file to open
    
    ' Reads an INI key value
    If IniState = T_INIReadkeyValue Then
        If TIniFile.CheckIni <> 1 Then DoINI = "": GoTo INICleanUp:
        DoINI = TIniFile.INIReadKeyValue(B, C, D)
    End If
    
    ' Writes a New INI Key Value
    If IniState = T_INIWriteKeyValue Then
        If TIniFile.CheckIni <> 1 Then DoINI = 0: GoTo INICleanUp:
        DoINI = TIniFile.INIWriteKeyValue(B, C, D)
    End If
    
    ' Deletes an INI Key Value
    If IniState = T_INIDeleteKeyValue Then
        If TIniFile.CheckIni <> 1 Then DoINI = 0: GoTo INICleanUp:
        DoINI = TIniFile.INIDeleteKey(B, C)
    End If
    
    If IniState = T_INIDeleteSelection Then
        If TIniFile.CheckIni <> 1 Then DoINI = 0: GoTo INICleanUp:
        DoINI = TIniFile.INIDeleteSelection(B)
    End If
    
INICleanUp:
    Erase vINIList
    A = "": B = "": C = "": D = ""
    Exit Function
    
ErrorCheck:
    If Err Then LastError 13, CurrentLinePos, Err.Description

End Function

Function GetHwnd(lzWinStr As String) As Long
Dim dwWnd As Long
    If Len(Trim(lzWinStr)) = 0 Then GetHwnd = 0: Exit Function
    dwWnd = FindWindow(vbNullString, lzWinStr)
    GetHwnd = dwWnd
End Function

Function isAdmin() As Long
    ' Test to see if the user is of
    isAdmin = IsNTAdmin(ByVal 0&, ByVal 0&)
End Function

Function GetSpFolder(lzStrFolName As String) As String
    ' Get and return special folder locations
    GetSpFolder = RegReadString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", lzStrFolName, REG_EXPAND_SZ)
End Function

Function RegDeleteValueOrKey(mParm As String, DelRegOption As RegDelete) As Long
Dim A As String, B As String, iPos As Integer, hPos As Integer, RegRootFlag As tHKey

    ' used to delete values and Keys form the registry
    
    A = ReturnData(mParm)
    
    iPos = InStr(1, A, "\", vbTextCompare)
    
    If (Not iPos) And (iPos <> 5) Then
        LastError 13, CurrentLinePos, "Syntax error in statement"
        A = ""
        Exit Function
    End If
    
    B = Mid(A, iPos + 1, Len(A)) ' Get the keyname or value name

    Select Case UCase(Left(A, iPos - 1)) ' setup the reg key to open
        Case "HKCR"
            RegRootFlag = HKEY_CLASSES_ROOT
        Case "HKCU"
            RegRootFlag = HKEY_CURRENT_USER
        Case "HKLM"
            RegRootFlag = HKEY_LOCAL_MACHINE
        Case "HKUE"
            RegRootFlag = HKEY_USERS
        Case "HKCO"
            RegRootFlag = HKEY_CURRENT_CONFIG
        End Select
        
        If DelRegOption = REG_KEY Then
            ' delete a Registry key
            RegDeleteValueOrKey = RegKeyDelete(RegRootFlag, B)
        End If
        
        If DelRegOption = REG_VALUE Then
            ' delete a Registry value form a key
            hPos = InStr(iPos + 1, A, "\", vbTextCompare)
            A = Mid(A, iPos + 1, hPos - iPos)
            B = Mid(B, (hPos - iPos) + 1, Len(B) - iPos)
            RegDeleteValueOrKey = RegDeleteValueEx(RegRootFlag, A, B)
        End If
        
        B = ""
        A = ""
        iPos = 0
        hPos = 0
        
End Function

Function RegReadWrite(mParm As String, mParmCount As Integer) As String
Dim vParms As Variant, iPos As Integer
Dim A As String, B As String, C As String, D As Variant, RegRootFlag As tHKey, RegTypeFlag As KeyType
On Error GoTo ErrorCheck:

    If FoundError Then GoTo ErrorCheck: Exit Function
    
    vParms = Split(mParm, ",")

    If UBound(vParms) < mParmCount Then
        LastError 7, CurrentLinePos
        Erase vParms
        mParm = ""
        Exit Function
    End If
    
    A = ReturnData(CStr(vParms(0))) ' Root Key
    B = ReturnData(CStr(vParms(1))) ' KeyName
    C = ReturnData(CStr(vParms(2))) ' KeyType
    
    If mParmCount > 2 Then
        D = ReturnData(CStr(vParms(3))) ' Regdata
    End If
    
    iPos = InStr(1, A, "\", vbBinaryCompare)
    
    ' The code above checks for the first slash in the regpath
    ' eg HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion
    ' This will then return the point to the Root Key Name seee below
    
    If (Not iPos) And (iPos <> 5) Then
        LastError 13, CurrentLinePos, "Syntax error in statement"
        A = "": B = "": C = "": D = "": Erase vParms
        Exit Function
    Else
        Select Case UCase(Left(A, iPos - 1))
            Case "HKCR"
                RegRootFlag = HKEY_CLASSES_ROOT
            Case "HKCU"
                RegRootFlag = HKEY_CURRENT_USER
            Case "HKLM"
                RegRootFlag = HKEY_LOCAL_MACHINE
            Case "HKUE"
                RegRootFlag = HKEY_USERS
            Case "HKCO"
                RegRootFlag = HKEY_CURRENT_CONFIG
        End Select
    End If
    
    Select Case UCase(C)
        Case "REG_SZ"
            RegTypeFlag = REG_EXPAND_SZ
        Case "REG_DWORD"
            RegTypeFlag = REG_DWORD
    End Select
    
    ' Strip of the root name of the path
    A = Right(A, Len(A) - iPos)
    
    If mParmCount = 2 Then
        RegReadWrite = RegReadString(RegRootFlag, A, B, RegTypeFlag)
        A = "": B = "": C = "": iPos = 0: Erase vParms
    End If
    
    If mParmCount = 3 Then
        RegReadWrite = RegSaveValue(RegRootFlag, A, B, RegTypeFlag, D)
        A = "": B = "": C = "": iPos = 0: Erase vParms
    End If
    
ErrorCheck:
    If Err Then LastError 13, CurrentLinePos, Err.Description
End Function

Function GetDriveCount() As Integer
Dim DrvBuff As String
Dim iRet As Long, vList() As String, iCount As Integer

    iCount = -1
    
    ' This returns a list of Locial Drives on the system
    DrvBuff = Space$(128)
    iRet = GetLogicalDriveStrings(256, DrvBuff)
    
    If iRet = 0 Then GetDriveCount = 0: DrvBuff = "": Exit Function
    DrvBuff = Left$(DrvBuff, iRet)
    
    Do While I < Len(DrvBuff)
        I = I + 1
        If Asc(Mid(DrvBuff, I, 1)) <> 0 Then
            S = S & Chr$(Asc(Mid$(DrvBuff, I, 1)))
            If Len(S) >= 3 Then
                iCount = iCount + 1
                S = ""
            End If
        End If
    Loop
    I = 0
    GetDriveCount = iCount
    iCount = 0
    
End Function

Function GetDriveList(num As Integer) As String
Dim DrvBuff As String
Dim iRet As Long, vList() As String, iCount As Integer
On Error GoTo DrvError:
    iCount = -1
    
    ' This returns a list of Locial Drives on the system
    DrvBuff = Space$(128)
    iRet = GetLogicalDriveStrings(256, DrvBuff)
    
    If iRet = 0 Then GetDriveList = "": DrvBuff = "": Exit Function
    DrvBuff = Left$(DrvBuff, iRet)
    
    Do While I < Len(DrvBuff)
        I = I + 1
        If Asc(Mid(DrvBuff, I, 1)) <> 0 Then
            S = S & Chr$(Asc(Mid$(DrvBuff, I, 1)))
            If Len(S) >= 3 Then
                iCount = iCount + 1
                ReDim Preserve vList(iCount)
                vList(iCount) = S
                S = ""
            End If
        End If
    Loop
    
    If (num < 0) And (num > iCount) Then
        LastError 13, CurrentLinePos, "Subscript out of range"
        GoTo CleanUp:
        Exit Function
    End If
    
    GetDriveList = vList(num)
    GoTo CleanUp:
    
CleanUp:
    I = 0
    iCount = 0
    DrvBuff = ""
    iRet = 0
    Erase vList()
    
DrvError:
    If Err Then LastError 13, CurrentLinePos, Err.Description
    
End Function

Function BrowseForFolder(lParm As String)
Dim lzTitle As String
    On Error Resume Next
    
    lzTitle = ReturnData(lParm)
    
    If FoundError Then
        LastError 13, CurrentLinePos, ErrorStr
        Exit Function
    End If

    BrowseForFolder = GetFolder(lzTitle)
    lzTitle = ""
    If n <> 0 Then Erase vList
    
End Function

Function GetFolder(Optional mTitle As String = "Look in:") As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim OffSet As Integer

    bInf.hOwner = 0
    bInf.lpszTitle = mTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE

    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    
    If RetVal Then
        OffSet = InStr(RetPath, Chr$(0))
        GetFolder = Left$(RetPath, OffSet - 1)
    End If
    
End Function

