Attribute VB_Name = "modFunctions"
' This mod is used to house some of the built in functions for DM++
' Please note I not comment much of this project as I whould be here till next year
' And wanted to get the project ready for you all
' I will hover ever try and add a comment line to the top of each function to say what it's doing

Enum StringOptions ' Used to tell us what function we working on
    DoLeft = 1
    DoRight = 2
    doFillChar
    isSame
    tadd
    tSub
    tMult
    tDiv
    tMod
    tCharPos
    tSetClipboard
    tSendKeys
    tDir
    tAppActivate
    tMonthName
    tWeekName
End Enum

Enum FileAccess ' Used to tell us what file access to use
    dnBinary = 1
    dmOutput = 2
    dmAppend = 3
    dmInput = 4
End Enum

Enum Conditional ' used to tell us what Conditiona to preform
    EqualTo = 1
    LessThan
    MoreThan
    NoEqual
    MoreThanEqual
    LessThanEqual
    BoolCheck
    IfLike
End Enum

Dim IFTest As Conditional

Function CreateFile(lzFile As String) As Long
On Error GoTo FileCreateError:
Dim nFile As Long
    
    nFile = FreeFile
    CreateFile = 1
    
    Open lzFile For Binary As #nFile
    Close #nFile
    
    Exit Function
FileCreateError:
    If Err Then CreateFile = 0

End Function
Function EnvCount() As Integer
Dim m As String
Dim iCount As Integer
' return system Environment variable count
Do
    iCount = iCount + 1
    S = Environ(iCount)
    If Len(S) = 0 Then Exit Do
    l = l + 1
Loop

EnvCount = (iCount - 1)
iCount = 0
S = ""

End Function

Function dmVar(Expression) As Variant
    If IsNumeric(Expression) Then dmVar = Val(Expression) Else dmVar = Expression
End Function

Function Sum(lzStr As String) As Double
On Error GoTo ErrCheck:
Dim vStr As Variant, I As Integer, hResult As Long

    ' This function is used to return the sum of numbers like you see in excel
    '$myvar = @Sum(10,10,10,10,etc)
    
    vStr = Split(lzStr, ",")
    
    If UBound(vStr) <= -1 Then LastError 7, CurrentLinePos: Exit Function
        
        For I = LBound(vStr) To UBound(vStr)
            hResult = hResult + ReturnData(CStr(vStr(I)))
        Next
        
        Erase vStr
        Sum = hResult
        hResult = 0
        I = 0
        Exit Function
ErrCheck:
    If Err Then LastError 13, CurrentLinePos, Err.Description
    
End Function

Function Average(lzStr As String) As Double
On Error GoTo ErrCheck:
Dim vStr As Variant, I As Integer, hResult As Long

    ' This function is used to return the average of whole or decimal numbers
    '$myvar = @Average(10,3.6,10,5.5,etc)
    
    vStr = Split(lzStr, ",")
    
    If UBound(vStr) <= -1 Then LastError 7, CurrentLinePos: Exit Function
        
        For I = LBound(vStr) To UBound(vStr)
            hResult = hResult + ReturnData(CStr(vStr(I)))
        Next

        Erase vStr
        Average = hResult / I
        hResult = 0
        I = 0
        Exit Function
ErrCheck:
    If Err Then LastError 13, CurrentLinePos, Err.Description
    
End Function

Function StrReplace(lzStr As String) As String
Dim A As String, B As String, C As String, D As Long
On Error GoTo RepError
    ' function used to replace a string what something else
    ' This does the same as it whould to in VB
    
    nArr = Split(lzStr, ",", , vbBinaryCompare)
    If UBound(nArr) < 2 Then LastError 7, CurrentLinePos: Exit Function
    
    ' In Comming Args
    A = nArr(0)
    B = nArr(1)
    C = nArr(2)
    
    ' return the data of the Args
    A = ReturnData(A)
    B = ReturnData(B)
    C = ReturnData(C)
   
    If UBound(nArr) = 2 Then
        StrReplace = Replace(A, B, C) ' send back result of the replace function
        Erase nArr
        A = "": B = "": C = ""
        Exit Function
    Else
        D = ReturnData(CStr(nArr(3)))
        StrReplace = Replace(A, B, C, , , D) ' send back result of the replace function
        Erase nArr
        A = "": B = "": C = "": D = 0
    End If
    
RepError:
    If Err Then LastError 13, CurrentLinePos, Err.Description

End Function

Function DoForLoop(mData As String)
Dim vHolder As Variant, sTmp(1 To 6) As String, sLine As String
Dim j As Integer

    ' Used to get all the information need to setup a forloop
    ' Does not support Step yet
    
    ' Syntex For varname = start to end
             'Next
    
    vHolder = Split(mData, " ") ' Split for loop by white spaces
    For I = LBound(vHolder) To UBound(vHolder)
        If Len(vHolder(I)) > 0 Then
            j = j + 1
            sTmp(j) = LCase(Trim(vHolder(I)))
        End If
    Next
    
    If I < 6 Then ' Stop if ubound vHolder is less than 6
        LastError 13, CurrentLinePos, "Syntext error in for loop"
        Exit Function
    End If
    
    Erase vHolder
    I = 0
    j = 0
    
    If Not VariableExists(sTmp(2)) Then ' check for a variable name
        LastError 13, CurrentLinePos, "Required: Variable"
        Erase sTmp
        Exit Function
    ElseIf Not sTmp(3) = "=" Then ' check for the equals sign
        LastError 13, CurrentLinePos, "Required ="
        Erase sTmp
        Exit Function
    ElseIf Not sTmp(5) = "to" Then ' check were we counting to
        LastError 13, CurrentLinePos, "Required ="
        Erase sTmp
        Exit Function
    Else
        ' Store all the forloop data
        ForLoopCounter = Val(ReturnData(sTmp(2)))
        ForLoopStart = Val(ReturnData(sTmp(4)))
        ForLoopEnd = Val(ReturnData(sTmp(6)))
        ForLoopVar = sTmp(2)
        Erase sTmp
    End If

End Function

Function DoMid(mData As String) As String
On Error GoTo MidError:
Dim nArr As Variant, A As String, B As Variant, D As Variant
  
    ' Returns the same as a normal mid function in VB
    
    nArr = Split(mData, ",", , vbBinaryCompare)
    If UBound(nArr) < 1 Then LastError 7, CurrentLinePos: Exit Function
    
    A = nArr(0)
    B = nArr(1)
    
    A = ReturnData(A)
    B = ReturnData(CStr(B))
   
    If UBound(nArr) = 1 Then
        DoMid = Mid(A, Val(B))
        Erase nArr
        A = "": B = ""
        Exit Function
    Else
        D = nArr(2)
        D = ReturnData(CStr(D))
        
        If Val(B) > 0 Then
            DoMid = Mid(A, Val(B), Val(D))
        End If
        
        Erase nArr
        A = "": B = ""
    End If
    
    Exit Function
MidError:

    If Err Then LastError 13, CurrentLinePos, Err.Description

End Function

Function ConCat(mStr As String) As String
Dim vCanCat As Variant
Dim sBuffer As String

    ' Function used to join one or more strings
    vCanCat = Split(mStr, ",")
    
    If UBound(vCanCat) = -1 Then LastError 7, CurrentLinePos: Exit Function
    
    For I = LBound(vCanCat) To UBound(vCanCat)
        sBuffer = sBuffer & ReturnData(CStr(vCanCat(I)))
    Next
    
    ConCat = sBuffer
    sBuffer = ""
    I = 0
End Function

Function Dec(lzStr As String)
Dim sHex As String
    ' Returns the same result as vb's CDec
    sHex = lzStr
    sHex = Replace(sHex, "&H", "", , , vbTextCompare)
    
    sHex = ReturnData(sHex)
    Dec = CDec("&H" & sHex)
    
    sHex = ""
End Function

Function DoProcCall(mData As String) As String
Dim iPos As Integer, lpos As String, ProcName As String, Parms As Variant
Dim HasParms As Boolean, ProcParms As Variant, Strbuff As String, j As Long
Dim VarTmp As String
On Error Resume Next
  
  
  ' Used to call a procedure in DM++
  ' Syntex
  ' Call MyFunc
  ' Call MyFunc(args)
  
    iPos = InStr(1, mData, "(", vbTextCompare)
    If iPos > 0 Then
        ProcName = GetFunctionName(mData)
        Parms = Split(GetBetweenBrackets(mData), ",")
        HasParms = True
    Else
        ProcName = mData
        HasParms = False
    End If

    ProcParms = Split(ProcedureGetParms(ProcName), ",")

    If Not HasParms And UBound(ProcParms) = -1 Then
        DoProcCall = ProcName
        GoTo CleanUp:
        Exit Function
    End If
    
    If Not UBound(ProcParms) = UBound(Parms) Then
        LastError 13, CurrentLinePos, "Wrong Number of arguments"
        GoTo CleanUp:
        Exit Function
    End If
    
    For j = LBound(Parms) To UBound(Parms)
        Strbuff = ReturnData(CStr(Parms(j)))
        SetVariableData CStr(ProcParms(j)), Strbuff
    Next
        
    DoProcCall = ProcName
    Erase Parms
    Erase ProcParms
    GoTo CleanUp:
    
CleanUp:
    HasParms = False
    j = 0
    iPos = 0
    lpos = 0
    ProcName = ""
    VarTmp = ""
    Strbuff = ""

End Function

Function Bin2Dec(BinNumber As String) As Long
Dim I As Long, Tmp As Long, sDec As Long
  
   ' Returns the integer value of a binary string
    Tmp = 1
    
    For I = Len(BinNumber) To 1 Step -1
        If CStr(Mid(BinNumber, I, 1)) = 1 Then
            sDec = sDec + Tmp
        End If
        Tmp = Tmp * 2
    Next
    
    Bin2Dec = sDec
    
End Function

Function Dec2Bin(IntDec As Long) As String
Dim sBinCheck As Boolean
Dim sBin As String
  ' Returns the binary string of an integer
  
    Do While IntDec <> 0
        sBinCheck = IntDec Mod 2
        If sBinCheck Then
            sBin = "1" & sBin
        Else
            sBin = "0" & sBin
        End If
        IntDec = IntDec \ 2
    Loop
    If sBin = "" Then
        Dec2Bin = "0"
    Else
        Dec2Bin = sBin
    End If
    sBin = ""
    
End Function

Function TestIF(mData As String) As Integer
Dim StrA As String, StrB As String, vStr As Variant, sLine As String, iPos As Integer
Dim I As Integer
Dim Strbuff As String
' this function does have bugs in it just never got around to fixing them
' will have to sort this out in the next version

' Function to preform IF statements eg $myvar1 <> $myvar2

On Error Resume Next
    vStr = Split(mData, " ")
    
    For I = LBound(vStr) To UBound(vStr)
        sLine = CStr(vStr(I))
        
        Select Case LCase(sLine)
            Case "=="
                iPos = InStr(1, mData, sLine, vbTextCompare)
                IFTest = EqualTo
                Exit For
            Case "<"
                iPos = InStr(1, mData, sLine, vbTextCompare)
                IFTest = LessThan
                Exit For
            Case ">"
                iPos = InStr(1, mData, sLine, vbTextCompare)
                IFTest = MoreThan
                Exit For
            Case "<="
                iPos = InStr(1, mData, sLine, vbTextCompare)
                IFTest = LessThanEqual
                Exit For
            Case ">="
                iPos = InStr(1, mData, sLine, vbTextCompare)
                IFTest = MoreThanEqual
                Exit For
            Case "<>"
                iPos = InStr(1, mData, sLine, vbTextCompare)
                IFTest = NoEqual
                Exit For
            Case "like"
                iPos = InStr(1, mData, sLine, vbTextCompare)
                IFTest = IfLike
                Exit For
        End Select
    Next
    I = 0
    Erase vStr
    
    Strbuff = Trim(StrFromBrackets(mData))
    
    If Len(Strbuff) = 0 Then
        Exit Function
    End If
    
    vStr = Split(Strbuff, sLine)
    If UBound(vStr) = 0 Then IFTest = BoolCheck
    
    StrA = ReturnData(Trim(CStr(vStr((0)))))
    StrB = ReturnData(Trim(CStr(vStr((1)))))
    
    Select Case IFTest
        Case EqualTo
            TestIF = Abs(StrA = StrB)
        Case LessThan
             TestIF = Abs(StrA < StrB)
        Case MoreThan
            TestIF = Abs(StrA > StrB)
        Case NoEqual
            TestIF = Abs(StrA <> StrB)
        Case LessThanEqual
            TestIF = Abs(StrA <= StrB)
        Case MoreThanEqual
            TestIF = Abs(StrA >= StrB)
        Case BoolCheck
            TestIF = Abs(CBool(ReturnData(Strbuff)))
        Case IfLike
            TestIF = Abs(StrA Like StrB)
    End Select
    
    Erase vStr
    iPos = 0
    StrA = ""
    StrB = ""
    Strbuff = ""
    sLine = ""
    
End Function

Function TFormat(mData As String) As String
Dim vStr As Variant, A As String, B As String
On Error Resume Next

   ' return a formated string eg for date and time
    vStr = Split(mData, ",")
    If UBound(vStr) < 0 Then LastError 7, CurrentLinePos: Exit Function
    
    A = ReturnData(CStr(vStr(0)))
    B = ReturnData(CStr(vStr(1)))
    
    If Err.Description Then Err.Clear
    TFormat = Format(A, B)
    Erase vStr
    A = ""
    B = ""
    
End Function

Function DirectoryExists(lzDir As String) As Integer
    ' is a folder found result is -1 or 0
    DirectoryExists = Abs(LenB(Dir(lzDir, vbDirectory)) > 1)
End Function

Function sBooltoInt(lz As String)
Dim A As String
    ' just returns the integer of a bool -1 or 0
    A = LCase(lz)
    If A = "true" Then sBooltoInt = "-1" Else sBooltoInt = "0"
    A = ""
End Function

Function TCopyFile(mData As String)
Dim vData As Variant, A As String, B As String
On Error GoTo CopyErr:

    ' Prforms a filecopy
    vData = Split(mData, ",")
    If UBound(vData) < 1 Then LastError 7, CurrentLinePos: Exit Function
    
    A = ReturnData(Trim(vData(0)))
    B = ReturnData(Trim(vData(1)))
    
    FileCopy A, B
    
    A = ""
    B = ""
    Erase vData
    
CopyErr:
If Err Then LastError 13, CurrentLinePos, Err.Description
End Function

Function RenameFile(mData As String) As Integer
Dim iPos As Integer, OldFile As String, NewFile As String
On Error GoTo RnFileErr:
    iPos = InStr(1, mData, ",", vbTextCompare)
    
    If iPos = 0 Then LastError 7, CurrentLinePos, "RenameFile": Exit Function
    
    OldFile = ReturnData(Left(mData, iPos - 1))
    NewFile = ReturnData(Right(mData, Len(mData) - iPos))
    
    If FoundError Then GoTo RnFileErr: Exit Function
    
    Name OldFile As NewFile
    RenameFile = 1
    
    OldFile = "": NewFile = "": iPos = 0

RnFileErr:
    If Err Then
        RenameFile = -1
    End If
    
    OldFile = "": NewFile = "": iPos = 0
    
End Function
Function isEqual(ifirst, inext) As String
    ' tests to see if ifirst is the same as inext
    ' no compare method is used add next time
    If ifirst = inext Then isEqual = "1" Else isEqual = "0"
End Function

Function ReturnData(lzData As String) As String
Dim StrA As String
    StrA = lzData
   'This function seems now to be the main part of the program
   ' used to return the data of expressions
   
    Select Case VarOrFunction(StrA)
        Case sFunction
            If Not Right(StrA, 1) = ")" Then StrA = StrA & ")"
            StrA = GetDataFormParam(StrA) ' function
        Case sTextOnly
            StrA = GetDataFormParam(StrA) ' text
        Case sVariable
            StrA = GetDataFormParam(StrA) ' variable
        Case Else
            StrA = Chr(34) & StrA & Chr(34)
            StrA = GetDataFormParam(StrA)
    End Select
    
    ReturnData = StrA
    StrA = ""
    
End Function

Function GetFileAcessType(lzType As Variant) As FileAccess
    ' find out the file access type we working with
    Select Case LCase(lzType)
        Case "binary"
            GetFileAcessType = dnBinary
        Case "output"
            GetFileAcessType = dmOutput
        Case "append"
            GetFileAcessType = dmAppend
        Case Else
            GetFileAcessType = -1
    End Select
    
End Function

Function DoFSO(lzData As String)
Dim fsoA As Variant, lzFileName As String, aTmp As String
Dim lpos As Integer, hPos As Integer
On Error GoTo FileError:
   
    aTmp = lzData
    aTmp = LTrim(Right(aTmp, Len(aTmp) - 4)) ' Strip away the open command
    
    lpos = InStr(1, aTmp, "[", vbBinaryCompare) ' find the first square bracket [
    hPos = InStrRev(aTmp, "]", Len(aTmp), vbBinaryCompare)
    
    If lpos = 0 Then LastError 12, CurrentLinePos, "[": Exit Function
    If hPos = 0 Then LastError 12, CurrentLinePos, "]": Exit Function
    
    aTmp = Mid(aTmp, lpos + 1, hPos - lpos - 1) ' extract the data between [ and ]
    lzFileName = GetDataFormParam(aTmp)
    aTmp = "": lpos = 0: hPos = 0
    
    fsoA = Split(lzData, Chr(32))
    hPos = UBound(fsoA) ' Now we need to find out the file access type eg binary,append or output
    
    If GetFileAcessType(fsoA(hPos)) = -1 Then LastError 11, CurrentLinePos: Exit Function
    If Not LCase(fsoA(hPos - 1)) = "for" Then LastError 12, CurrentLinePos, "For": Exit Function
    
    nFilePointer = FreeFile
    
    Select Case GetFileAcessType(fsoA(hPos))
        Case dnBinary
            nFileReadWriteAccess = dnBinary
            Open lzFileName For Binary As #1
        Case dmOutput
            nFileReadWriteAccess = dmOutput
            Open lzFileName For Output As #1
        Case dmAppend
            nFileReadWriteAccess = dmAppend
            Open lzFileName For Append As #1
        Case dmInput
            nFileReadWriteAccess = dmInput
            Open lzFileName For Input As #1
    End Select
    
    Erase fsoA
    lzFileName = ""
    hPos = 0
    
FileError:
    If Err Then
        LastError 13, CurrentLinePos, Err.Description
    End If
    
End Function

Function FsoDoRead(lzData As String)
Dim vParm As Variant
Dim StrA As String, StrB As String, FileBuff As String

    vParm = Split(lzData, ",")
    If UBound(vParm) <> 2 Then LastError 13, CurrentLinePos, "Required: List Separator ,"
    
    StrA = Trim(vParm(1))
    StrB = vParm(2)
    Erase vParm
    
      Select Case VarOrFunction(StrA)
        Case sFunction
            StrA = GetDataFormParam(StrA)
        Case sTextOnly
            StrA = GetDataFormParam(StrA)
        Case sVariable
            StrA = GetDataFormParam(StrA)
        Case Else
            StrA = Chr(34) & StrA & Chr(34)
            StrA = GetDataFormParam(StrA)
    End Select
    
    If Not VariableExists(StrB) Then
        LastError 13, CurrentLinePos, "Variable Required"
        Exit Function
    End If
    
    
    FileBuff = Space(Len(GetVariableData(StrB)))

    Select Case nFileReadWriteAccess
        Case dnBinary
        If Len(StrA) = 0 Then
            Get #nFilePointer, , FileBuff
            SetVariableData StrB, FileBuff
            FileBuff = ""
        Else
            Get #nFilePointer, Val(StrA), FileBuff
            SetVariableData StrB, FileBuff
            FileBuff = ""
        End If
        
    End Select
    
End Function

Function FsoDoWrite(lzData As String)
Dim vParm As Variant
Dim StrA As String, StrB As String

    vParm = Split(lzData, ",")
    If UBound(vParm) <> 2 Then LastError 13, CurrentLinePos, "Required: List Separator ,"
    
    StrA = Trim(vParm(1))
    StrB = vParm(2)
    Erase vParm
    
    Select Case VarOrFunction(StrA)
        Case sFunction
            StrA = GetDataFormParam(StrA)
        Case sTextOnly
            StrA = GetDataFormParam(StrA)
        Case sVariable
            StrA = GetDataFormParam(StrA)
        Case Else
            StrA = Chr(34) & StrA & Chr(34)
            StrA = GetDataFormParam(StrA)
    End Select

    Select Case VarOrFunction(StrB)
        Case sFunction
            StrB = GetDataFormParam(StrB)
        Case sTextOnly
            StrB = GetDataFormParam(StrB)
        Case sVariable
            StrB = GetDataFormParam(StrB)
        Case Else
            StrB = Chr(34) & StrB & Chr(34)
            StrB = GetDataFormParam(StrB)
    End Select
    
    Select Case nFileReadWriteAccess
        Case dmAppend
            If Len(StrA) = 0 Then
                Print #nFilePointer, StrB
            Else
                Print #nFilePointer, Val(StrA), StrB
            End If
        Case dnBinary
            If Len(StrA) = 0 Then
                Put #nFilePointer, , StrB
            Else
                Put #nFilePointer, Val(StrA), StrB
            End If
        Case dmOutput
            If Len(StrA) = 0 Then
                Print #nFilePointer, StrB
            Else
                Print #nFilePointer, Val(StrA), StrB
            End If
    End Select
    
    StrA = ""
    StrB = ""
    
End Function

Public Sub Pause(Optional interval)
    'pause execution
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub

Function DeleteFile(lzFile As String) As Integer
On Error GoTo DeleteErr:
    DeleteFile = 1
    Kill lzFile
Exit Function
DeleteErr:
    If Err Then LastError 13, CurrentLinePos, Err.Description
End Function

Public Function MakeDir(lzDirName As String) As Long
On Error GoTo FileIoErr:
    MakeDir = 1
    MkDir ReturnData(lzDirName)
    Exit Function
FileIoErr:
    MakeDir = 0
End Function

Public Sub RemoveDir(lzDirName As String)
On Error GoTo FileIoErr:
    RmDir ReturnData(lzDirName)
Exit Sub
FileIoErr:
    If Err Then LastError 13, CurrentLinePos, Err.Description
End Sub

Public Sub ChangeDir(lzDirName As String)
On Error GoTo FileIoErr:
    ChDir ReturnData(lzDirName)
Exit Sub
FileIoErr:
    If Err Then LastError 13, CurrentLinePos, Err.Description
End Sub

Function StrFunction(mData As String, StrOption As StringOptions, Optional isOptional As Boolean) As String
Dim A As String, B As String, iPos As Integer
On Error Resume Next

    iPos = InStr(1, mData, ",", vbTextCompare)
    
    If Not isOptional Then
        If iPos = 0 Then LastError 7, CurrentLinePos: Exit Function
    End If
    
    
    A = GetDataFormParam(Left(mData, iPos - 1))
    B = Right(mData, Len(mData) - iPos)
    B = ReturnData(B) 'new code

    
    If Not VariableExists(B) Then
        B = Chr(34) & B & Chr(34): B = FixQuote(B)
        B = GetDataFormParam(B)
    Else
        B = GetDataFormParam(B)
    End If
    
    Select Case StrOption
        Case tAppActivate
            If Len(A) = 0 Then AppActivate B, 0 Else AppActivate A, Int(B)
            GoTo ErrorCheck:
        Case DoLeft
            StrFunction = Left(A, Val(B))
        Case DoRight
            StrFunction = Right(A, Val(B))
        Case doFillChar
           StrFunction = String(Val(B), A)
        Case isSame
            StrFunction = CStr(isEqual(A, B))
        Case tadd
            StrFunction = CStr(Val(A) + Val(B))
        Case tSub
            StrFunction = CStr(Val(A) - Val(B))
        Case tMonthName
            If Len(A) = 0 Then StrFunction = MonthName(Int(B)) Else StrFunction = MonthName(Int(A), Int(B))
        Case tWeekName
           If Len(A) = 0 Then StrFunction = WeekdayName(Int(B)) Else StrFunction = WeekdayName(Int(A), Int(B))
        Case tMult
            StrFunction = CStr(Val(A) * Val(B))
        Case tDiv
            StrFunction = CStr(Val(A) / Val(B))
        Case tMod
            StrFunction = CStr(Val(A) Mod Val(B))
        Case tCharPos
            StrFunction = InStr(1, A, B, vbTextCompare)
        Case tSetClipboard
            Clipboard.SetText A, B
        Case tSendKeys
           SendKeys B, Val(A)
        Case tDir
            If Len(A) = 0 Then StrFunction = Dir(B, Val(A)): Exit Function
            StrFunction = Dir(A, Val(B))
            
    End Select
    A = "": B = "": iPos = 0
    
ErrorCheck:
    If Err Then
        If Err Then LastError 13, CurrentLinePos, Err.Description
    End If
    
End Function

Function PromptBox(mData As String) As String
Dim vStr As Variant
Dim StrA As String, StrB As String
On Error Resume Next

    vStr = Split(mData, ",")
    
    If UBound(vStr) = 0 Then
        StrA = vStr(0)
        sInputBoxPrompt = GetDataFormParam(StrA)
    End If
    
    If UBound(vStr) = 1 Then
        StrA = vStr(0)
        StrB = vStr(1)
        sInputBoxPrompt = GetDataFormParam(StrA)
        sInputBoxTitle = GetDataFormParam(StrB)
    End If
    
    PromptBox = InputBox(sInputBoxPrompt, sInputBoxTitle)
    
    ' Clean up used variables

    
End Function

Function HasBrackets(lzData As String) As Boolean
Dim lpos As Integer, hPos As Integer
    HasBrackets = False
    lpos = InStr(1, lzData, "(", vbTextCompare)
    hPos = InStr(lpos + 1, lzData, ")", vbTextCompare)
    If (lpos > 0 And hPos > 0) Then HasBrackets = True
    lpos = 0: hPos = 0
End Function

Function GetDataFromStr(sData As String) As String
Dim StrB As String
Dim sFuncName As String
Dim TempA As String, TempB As String, TempC As String, TempInt As Integer, VarB As Variant

    Select Case VarOrFunction(sData)
         Case sFunction
            sFuncName = GetFunctionName(sData) ' Get the in comming function name
            StrB = DoFunction(sFuncName, sData) ' return the functions data
  
         Case sVariable
     
            If Not VariableExists(sData) Then
                If Not HasBrackets(sData) Then
                    LastError 6, CurrentLinePos, sData
                    Exit Function
                Else
                    TempB = Mid(sData, 1, GetOpenBracket(sData) - 1)
                    If Not VariableExists(TempB) Then
                        LastError 6, CurrentLinePos, sData
                        Exit Function
                    Else
                        'TempC = GetVariableData(TempB)
                        'MsgBox TempC
                        
                        VarB = Split(GetVariableData(TempB), ",", , vbTextCompare)
                        TempInt = Int(ReturnData(StrFromBrackets(sData)))
                        If (TempInt > UBound(VarB) Or TempInt < LBound(VarB)) Then LastError 13, CurrentLinePos, "Subscript out of range": Exit Function
                        sData = Mid(sData, 1, GetOpenBracket(sData) - 1)
                        If TempInt = 0 Then
                            TempB = Trim(Mid(VarB(0), 9, Len(VarB(0))))
                        Else
                            TempB = CStr(VarB(TempInt))
                        End If
                        GetDataFromStr = ReturnData(TempB)
                        Exit Function
                    End If
                   End If
                    Exit Function
                End If

            If Not HasBrackets(sData) Then 'check if we are dealing with an array datatype
                StrB = GetVariableData(sData)
                
                If Left(StrB, 8) = "{ARRAY}-" Then ' new code
                    LastError 10, CurrentLinePos, sData
                    Exit Function
                End If
                'end of new code
            Else
                ' yes we found an array datatype
                StrB = ReturnArrayData(GetArrayNameIndex(sData, 1), GetArrayNameIndex(sData, 2))
            End If
            
         Case sTextOnly ' plain text
         
         If IsNumeric(sData) Then sData = Chr(34) & sData & Chr(34) ' new code
         If ConstExists(sData) Then sData = GetConstValue(sData) ' check if it's a const
           ' Code used to detect enums
            If GetOpenBracket(sData) Then
                TempA = LCase(RTrim(Left(sData, GetOpenBracket(sData) - 1)))
                If EnumExists(TempA) Then
                    TempB = GetBetweenBrackets(sData)
                    If Not EnumGetItemIndex(TempA, TempB) <> vbNullString Then
                        LastError 12, CurrentLinePos, "Expression"
                        TempB = ""
                        TempA = ""
                        Exit Function
                    Else
                         TempA = ReturnData(EnumGetItemIndex(TempA, TempB))
                         TempB = ""
                         sData = Chr(34) & TempA & Chr(34)
                         TempA = ""
                    End If
                End If
            End If
            'end if enum code
            
            If Not (Left(sData, 1) = Chr(34) And Right(sData, 1) = Chr(34)) Then ' new code
                LastError 10, CurrentLinePos, Chr(34) ' new code
                Exit Function ' new code
            Else ' new code
                StrB = GetStrFromQuot(sData) ' old code
            End If ' new code
    End Select
    
    TempA = ""
    TempB = ""
    sFuncName = ""
    GetDataFromStr = StrB
    StrB = "" ' clear buffer
End Function

Function Echo(mData As String) As String
Dim vStr As Variant, iPos As Integer, fData As String
Dim FuncName As String
Dim StrA As String, StrB As String, StrC As String
On Error Resume Next

    sInputBoxTitle = "DM++ Script"
    
    If Left(mData, 1) = "@" Then
        ' This block of code may or not be complate yet I only used it for testing purposes
        ' is used to return a functions return data eg @right("ben",1) ' return 'n'
        iPos = GetPositionOfChar(mData, ")")
        FuncName = GetFunctionName(mData)
        fData = Mid(mData, Len(FuncName) + 1, iPos - Len(FuncName))
        fData = DoFunction(FuncName, fData)
        mData = Chr(34) & fData & Chr(34) & Right(mData, Len(mData) - iPos)
    End If
    
    If FoundError Then Exit Function
    
    vStr = Split(mData, ",")
    
    If UBound(vStr) = 0 Then
        StrA = vStr(0)
        sInputBoxPrompt = GetDataFormParam(StrA)
    End If
    
    If UBound(vStr) = 1 Then
        StrA = vStr(0)
        StrB = vStr(1)
        sInputBoxPrompt = GetDataFormParam(StrA)
        sInputBoxButtons = Val(GetDataFormParam(StrB))
    End If
    
    If UBound(vStr) = 2 Then
        StrA = vStr(0)
        StrB = vStr(1)
        StrC = vStr(2)
        
        sInputBoxPrompt = GetDataFormParam(StrA)
        sInputBoxButtons = Val(GetDataFormParam(StrB))
        sInputBoxTitle = GetDataFormParam(StrC)
    End If
    
    If FoundError Then Exit Function
        Echo = MsgBox(sInputBoxPrompt, sInputBoxButtons, sInputBoxTitle)
        
CleanUp:
    sInputBoxPrompt = ""
    sInputBoxButtons = 0
    sInputBoxTitle = ""
    Erase vStr
End Function

Function GetDataFormParam(lzStr As String) As String
On Error Resume Next
Dim tStr As Variant, j As Integer
Dim StrA As String
Dim StrData As String


        tStr = Split(lzStr, "&")
        
        For j = LBound(tStr) To UBound(tStr)
           StrA = Trim(tStr(j))
           'removes the last bracked at the end of a string eg ) only when a string is assign
           If Right(StrA, 2) = Chr(34) & ")" And Not Left(StrA, 1) = "@" Then StrA = Left(StrA, Len(StrA) - 1)
           StrData = StrData & GetDataFromStr(StrA)
        Next
        
        If FoundError Then Exit Function
        GetDataFormParam = StrData
        
        StrData = ""
        lzStr = ""
        StrA = ""
        Erase tStr
        j = 0
End Function

Public Function Rand(num) As Integer
    Randomize
    Rand = Int(num * Rnd)
End Function
