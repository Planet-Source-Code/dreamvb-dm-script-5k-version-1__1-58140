Attribute VB_Name = "ModTools"
Public FoundError As Boolean
Public ErrorStr As String

Enum VarFuncCheck
    [sVariable] = 1
    [sFunction] = 2
    [sTextOnly] = 3
End Enum

Public Function GetArrayNameIndex(lzStr As String, Op As Integer) As Variant
Dim a As Long, b As Long, StrA As String
    
    a = InStr(1, lzStr, "(", vbTextCompare)
    b = InStr(a + 1, lzStr, ")", vbTextCompare)

    If Not (a > 0 And b > 0) Then
        LastError 5, CurrentLinePos, lzStr
        Exit Function
    End If
    
    If Op = 1 Then ' Get array name
        GetArrayNameIndex = Trim(Mid(lzStr, 2, a - 2))
    ElseIf Op = 2 Then ' get array index
        GetArrayNameIndex = ReturnData(Trim(Mid(lzStr, a + 1, b - a - 1)))
    Else
        StrA = GetFromAssignPos(lzStr, GetPositionOfChar(lzStr, "=") + 1)
        StrA = StripEOL(StrA)
        GetArrayNameIndex = StrA
        StrA = ""
    End If
    
    a = 0: b = 0
    
End Function

Public Function IsFileHere(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Function FixPath(lzPath As String) As String
    If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Function GetDirective(lzStr As String) As String
Dim a As String
    a = lzStr: a = Replace(a, ">", ""): a = Replace(a, "<", ""): GetDirective = a
    a = ""
End Function
Function GetFunctionName(lzStr As String) As String
Dim StrA As String
    StrA = LCase(Mid(lzStr, 1, GetPositionOfChar(lzStr, "(") - 1))
    StrA = CleanWhiteSpace(StrA)
    GetFunctionName = StrA
End Function

Public Function GetStrFromQuot(lzStr As String) As String
Dim hPos As Integer, lpos As Integer
    lpos = InStr(1, lzStr, Chr(34), vbTextCompare)
    hPos = InStr(lpos + 1, lzStr, Chr(34), vbTextCompare)
    If (lpos > 0) And (hPos > 0) Then
        GetStrFromQuot = Mid(lzStr, lpos + 1, hPos - lpos - 1)
    Else
        GetStrFromQuot = lzStr
    End If
End Function

Public Function GetAssignPos(lzStr As String, AssignChar As String) As Long
    GetAssignPos = InStr(1, lzStr, AssignChar, vbTextCompare)
End Function

Public Function GetOpenBracket(lzStr As String) As Long
    GetOpenBracket = InStr(1, lzStr, "(", vbTextCompare)
End Function

Public Sub LastError(nErrorNo As Integer, LineNo As Long, Optional ExtraInfo As String)
    Select Case nErrorNo
        Case 0
           ErrorStr = "End terminater missing ; " & ExtraInfo & " at line: " & LineNo
        Case 1
            ErrorStr = "Start of statement missing ( " & "at line: " & LineNo
        Case 2
            ErrorStr = "End of statement missing ) " & "at line: " & LineNo
        Case 3
            ErrorStr = "Expected identfier at line: " & LineNo
        Case 4
            ErrorStr = "Expected End " & ExtraInfo & "; at line: " & LineNo
        Case 5
            ErrorStr = "Syntax error Unknown Data Type " & ExtraInfo & " at line: " & LineNo
        Case 6
           ErrorStr = "Undeclared identifier:  " & ExtraInfo & " at line: " & LineNo
        Case 7
            ErrorStr = "Argument not optional at line: " & LineNo
        Case 8
            ErrorStr = "Illegal character at line: " & LineNo
        Case 9
            ErrorStr = "Procedure or Function not defined " & ExtraInfo & " at line: " & LineNo
        Case 10
            ErrorStr = "Type mismatch " & ExtraInfo & " at line: " & LineNo
        Case 11
            ErrorStr = "Expected Output,Binary or Append at line: " & LineNo
        Case 12
           ErrorStr = "Expected " & ExtraInfo & " at line: " & LineNo
        Case 13
           ErrorStr = ExtraInfo & " at line: " & LineNo
    End Select
    
    FoundError = True
    CurrentLinePos = LineNo
    
End Sub

Function GetPositionOfChar(lzStr As String, nChar As String) As Integer
Dim iPos As Integer
    iPos = InStr(1, lzStr, nChar, vbTextCompare)
    If iPos > 0 Then GetPositionOfChar = iPos
    iPos = 0
End Function

Function GetBetweenBrackets(lzStr As String) As String
Dim iPos As Integer, lpos As Integer
    iPos = InStr(1, lzStr, "(", vbTextCompare)
    lpos = InStr(iPos + 1, lzStr, ")", vbTextCompare)
    
    If (iPos > 0) And (lpos > 0) Then
        GetBetweenBrackets = Trim(Mid(lzStr, iPos + 1, lpos - iPos - 1))
    Else
        GetBetweenBrackets = lzStr
    End If
End Function

Function GetArrayData(lzStr As String) As String
Dim iPos As Integer, lpos As Integer
    iPos = InStr(1, lzStr, "(", vbTextCompare)
    lpos = InStr(iPos + 1, lzStr, ")", vbTextCompare)
    
    If (iPos > 0) And (lpos > 0) Then
        GetArrayData = Trim(Mid(lzStr, iPos + 1, lpos - iPos - 1))
    Else
        GetArrayData = lzStr
    End If
End Function

Function GetEOL(lzStr As String) As Integer
    GetEOL = InStr(1, lzStr, ";", vbTextCompare)
End Function

Function RemoveComments(lzStr As String) As String

Dim sText As Variant, I As Long, iPos As Long, b As String, StrLine As String

    sText = Split(lzStr, vbCrLf)
    
    For I = LBound(sText) To UBound(sText)
        StrLine = sText(I)
        iPos = InStr(1, StrLine, "//", vbTextCompare)
        If iPos <> 0 Then
             sText(I) = Mid(StrLine, 1, iPos - 1)
        End If
        
        b = b & sText(I) & vbCrLf
    Next
    
    RemoveComments = b
    Erase sText
    b = ""
    StrLine = ""
    I = 0
    iPos = 0

End Function

Function RemoveTabs(S As String) As String
    RemoveTabs = Replace(S, vbTab, "")
End Function

Function GetVariableType(sVar As String) As String

    Select Case Left(sVar, 1)
        Case "$"
            GetVariableType = vbString
        Case "%"
            GetVariableType = vbInteger
        Case "&"
            GetVariableType = vbLong
        Case "!"
            GetVariableType = vbSingle
        Case "#"
            GetVariableType = vbArray
        Case Else
            GetVariableType = vbNullChar
    End Select
    
End Function

Function IsVariable(lzStr As String) As Boolean
    IsVariable = False
    If lzStr = "$" Or lzStr = "%" Or lzStr = "!" Or lzStr = "#" Then IsVariable = True
End Function

Function VarOrFunction(scheck As String) As VarFuncCheck

    Select Case Left(scheck, 1)
        Case "$", "%", "!", "#"
            VarOrFunction = sVariable ' it's a vairable
        Case "@"
            VarOrFunction = sFunction ' it's a function
        Case Chr(34)
            VarOrFunction = sTextOnly ' it text
        Case Else
            VarOrFunction = sTextOnly
    End Select
    
End Function

Function CleanWhiteSpace(lzStr As String) As String
Dim S As String

    S = lzStr
    S = Replace(S, " ", "")
    S = Replace(S, vbTab, "")
    CleanWhiteSpace = S
    S = ""
End Function

Function FormatStr(lzStr As String) As String
Dim StrA As String
' this function is used to house all the built in consts of dm++

    StrA = lzStr
    ' String Consts
    StrA = Replace(StrA, "dmLF", "@chr(10)", , , vbTextCompare)
    'StrA = Replace(StrA, "dmCR", "@chr(13)", , , vbTextCompare)
    StrA = Replace(StrA, "dmCrLf", "@chr(13) & @chr(10)", , , vbTextCompare)
    StrA = Replace(StrA, "dmFormFeed", "@chr(12)", , , vbTextCompare)
    StrA = Replace(StrA, "dmBack", "@chr(8)", , , vbTextCompare)
    StrA = Replace(StrA, "dmNewLine", "@chr(13) & @chr(10)", , , vbTextCompare)
    StrA = Replace(StrA, "dmNull", "", , , vbTextCompare)
    StrA = Replace(StrA, "dmTab", "@chr(9)", , , vbTextCompare)
    StrA = Replace(StrA, "dmVerticalTab", "@chr(11)", , , vbTextCompare)
    StrA = Replace(StrA, "dmSpace", "@chr(32)", , , vbTextCompare)
    StrA = Replace(StrA, "dmQuote", Chr(34) & Chr(34) & " & " & "@chr(34)", , , vbTextCompare)
    
    'Message box Types consts
    StrA = Replace(StrA, "dmAbortRetryIgnore", "2", , , vbTextCompare)
    StrA = Replace(StrA, "mbInformation", "64", , , vbTextCompare)
    StrA = Replace(StrA, "mbCritical", "16", , , vbTextCompare)
    StrA = Replace(StrA, "mbExclamation", "48", , , vbTextCompare)
    StrA = Replace(StrA, "mbQuestion", "32", , , vbTextCompare)
    StrA = Replace(StrA, "dmRetryCancel", "5", , , vbTextCompare)
    StrA = Replace(StrA, "dmApplicationModal", "0)", , , vbTextCompare)
    StrA = Replace(StrA, "dmSystemModal", "4096", , , vbTextCompare)
    ' Message Box button consts
    StrA = Replace(StrA, "dmOK", "1", , , vbTextCompare)
    StrA = Replace(StrA, "dmCancel", "2", , , vbTextCompare)
    StrA = Replace(StrA, "dmAbort", "3", , , vbTextCompare)
    StrA = Replace(StrA, "dmRetry", "4", , , vbTextCompare)
    StrA = Replace(StrA, "dmIgnore", "5", , , vbTextCompare)
    StrA = Replace(StrA, "dmYes", "6", , , vbTextCompare)
    StrA = Replace(StrA, "dmNo", "7", , , vbTextCompare)
    StrA = Replace(StrA, "mbYesNo", "4", , , vbTextCompare)
    
    ' Clicpboard consts
    StrA = Replace(StrA, "dmText", "1", , , vbTextCompare)
    StrA = Replace(StrA, "dmRTF", "-16639", , , vbTextCompare)
    ' DM++ datatype consts
    StrA = Replace(StrA, "dmString", "8", , , vbTextCompare)
    StrA = Replace(StrA, "dmInteger", "2", , , vbTextCompare)
    
    ' Boolean consts
    StrA = Replace(StrA, "dmTrue", "-1", , , vbTextCompare)
    StrA = Replace(StrA, "dmFalse", "0", , , vbTextCompare)
    
    'DM++ Window style consts
    StrA = Replace(StrA, "SW_HIDE", "0", , , vbTextCompare)
    StrA = Replace(StrA, "SW_NORMAL", "1", , , vbTextCompare)
    StrA = Replace(StrA, "SW_MAXSIZE", "3", , , vbTextCompare)
    StrA = Replace(StrA, "SW_MINSIZE", "6", , , vbTextCompare)
    
    'Dir Consts
    'dmFolder
    StrA = Replace(StrA, "dmFolder", "18", , , vbTextCompare)
    StrA = Replace(StrA, "dmArchive", "32", , , vbTextCompare)
    StrA = Replace(StrA, "dmHidden", "2", , , vbTextCompare)
    StrA = Replace(StrA, "dmReadOnly", "1", , , vbTextCompare)
    StrA = Replace(StrA, "dmSystem", "4", , , vbTextCompare)
    
    ' Drive Type Consts
    
    StrA = Replace(StrA, "dwRemovable", "2", , , vbTextCompare)
    StrA = Replace(StrA, "dwFixed", "3", , , vbTextCompare)
    StrA = Replace(StrA, "dwRemote", "4", , , vbTextCompare)
    StrA = Replace(StrA, "dwCDRom", "5", , , vbTextCompare)
    StrA = Replace(StrA, "dwRamDisk", "6", , , vbTextCompare)
    
    ' Date consts weekdays
    StrA = Replace(StrA, "dmMonday", "2", , , vbTextCompare)
    StrA = Replace(StrA, "dmTuesday", "3", , , vbTextCompare)
    StrA = Replace(StrA, "dmWednesday", "4", , , vbTextCompare)
    StrA = Replace(StrA, "dmThursday", "5", , , vbTextCompare)
    StrA = Replace(StrA, "dmFriday", "6", , , vbTextCompare)
    StrA = Replace(StrA, "dmSaturday", "7", , , vbTextCompare)
    StrA = Replace(StrA, "dmSunday", "1", , , vbTextCompare)
    
    ' Compare Conts
    StrA = Replace(StrA, "dwBinCompare", "0", , , vbTextCompare)
    StrA = Replace(StrA, "dwStrCompare", "1", , , vbTextCompare)
    
    ' Version Version
    StrA = Replace(StrA, "Script.VersionBuild", App.Major & "." & App.Minor, , , vbTextCompare)
    StrA = Replace(StrA, "Script.PathFileName", Chr(34) & scriptFullPath & Chr(34), , , vbTextCompare)
    
    
    FormatStr = StrA
    StrA = ""
    
End Function

Function GetFromAssignPos(StrString As String, AssignPos As Integer) As String
    If AssignPos = 0 Then GetFromAssignPos = "": Exit Function
    GetFromAssignPos = LTrim(Mid(StrString, AssignPos, Len(StrString)))
End Function

Function StripEOL(lzStr As String) As String
    If Right(lzStr, 1) = ";" Then
        StripEOL = Left(lzStr, Len(lzStr) - 1)
    Else
        StripEOL = lzStr
    End If
    
End Function

Function StrFromBrackets(lzStr As String) As String
Dim iPos As Integer, lpos As Integer, I As Integer
    iPos = InStr(1, lzStr, "(", vbTextCompare)
    
    If Not iPos > 0 Then
        StrFromBrackets = ""
        Exit Function
    End If
    
    For I = 1 To Len(lzStr)
        If Mid(lzStr, I, 1) = ")" Then lpos = I
    Next
    
    If Not lpos > 0 Then StrFromBrackets = "": Exit Function
    StrFromBrackets = Trim(Mid(lzStr, iPos + 1, lpos - iPos - 1))
    
    I = 0
    lpos = 0
    iPos = 0
    
End Function

Public Function DoFunction(nFunctionName As String, Optional Function_Data As String, Optional FindBrackets As Boolean) As String
Dim aTmp As String, StrAPIname As String
Dim bTmp As String, StrTemp As String, args As Variant, CustomFunction As Boolean
Dim tIdx As Integer, tArraySize As Integer
Dim iCounter As Integer
Dim Func_Data As String
On Error Resume Next


    If Not FindBrackets Then
        aTmp = StrFromBrackets(Function_Data)
    Else
        aTmp = Function_Data
    End If
  
    Select Case LCase(nFunctionName)
        Case "@average"
            DoFunction = Average(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@atn"
            DoFunction = Atn(ReturnData(aTmp))
        Case "@app.path"
            If Len(scriptFullPath) = 0 Then
                DoFunction = App.Path
            Else
                DoFunction = scriptFullPath
            End If
            
            GoTo ErrorCheck
        Case "@add"
            DoFunction = StrFunction(aTmp, tadd)
            GoTo ErrorCheck
        Case "@array"
            DoFunction = "{ARRAY}-" & aTmp
        Case "@abs"
            DoFunction = Abs(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@asc"
            aTmp = ReturnData(aTmp)
            If aTmp = "" Then Else DoFunction = Asc(aTmp)
            GoTo ErrorCheck
        Case "@browseforfolder"
            DoFunction = BrowseForFolder(aTmp)
            GoTo ErrorCheck:
        Case "@bin2dec"
            aTmp = ReturnData(aTmp)
            DoFunction = Bin2Dec(aTmp)
            GoTo ErrorCheck
        Case "@booltoint"
            DoFunction = sBooltoInt((ReturnData(aTmp)))
            GoTo ErrorCheck
        Case "@command"
            DoFunction = StrCommandLine
        Case "@clipboard.gettext"
             aTmp = ReturnData(aTmp)
             If Len(aTmp) > 0 Then
                DoFunction = Clipboard.GetText(aTmp)
            Else
                DoFunction = Clipboard.GetText()
            End If
            GoTo ErrorCheck
        Case "@cos"
            DoFunction = Cos(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@clipboard.settext"
            DoFunction = StrFunction(aTmp, tSetClipboard)
            GoTo ErrorCheck
        Case "@chrpos"
            DoFunction = StrFunction(aTmp, tCharPos)
        Case "@curdir"
            DoFunction = CurDir(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@concat"
            DoFunction = ConCat(aTmp)
            GoTo ErrorCheck
        Case "@destroy"
            aTmp = Right(aTmp, Len(aTmp) - 1)
            tIdx = GetArryIndexByName(aTmp)
            If tIdx = -1 Then LastError 12, CurrentLinePos, "Array": Exit Function
            DoFunction = EraseArray(tIdx)
        Case "@dec"
            DoFunction = Dec(aTmp)
            GoTo ErrorCheck
        Case "@dec2bin"
             aTmp = ReturnData(aTmp)
             DoFunction = Dec2Bin(CStr(aTmp))
             GoTo ErrorCheck
        Case "@drives.count"
            DoFunction = GetDriveCount
            GoTo ErrorCheck
        Case "@drivelist"
            DoFunction = GetDriveList(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@directoryexists"
            DoFunction = DirectoryExists(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@deletefile"
            DoFunction = DeleteFile(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@day"
            DoFunction = Day(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@div"
            DoFunction = StrFunction(aTmp, tDiv)
            GoTo ErrorCheck
        Case "@date"
            DoFunction = Date
        Case "@dir"
            DoFunction = StrFunction(aTmp, tDir, True)
        Case "@eval"
            DoFunction = modEVal.Evaluate(aTmp)
        Case "@exp"
            DoFunction = CDbl(Exp(ReturnData(aTmp)))
            GoTo ErrorCheck
        Case "@format"
            DoFunction = TFormat(aTmp)
        Case "@filedatetime"
            DoFunction = FileDateTime(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@filesize"
            DoFunction = FileLen(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@fillstr"
            DoFunction = StrFunction(aTmp, doFillChar)
        Case "@fixnum"
            DoFunction = Fix(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@hex"
            DoFunction = Hex(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@getspecialfolder"
            DoFunction = GetSpFolder(ReturnData(aTmp))
        Case "@gethwnd"
            DoFunction = CDbl(GetHwnd(ReturnData(aTmp)))
            GoTo ErrorCheck
        Case "@hour"
            DoFunction = Hour(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@inideleteselection"
            DoFunction = DoINI(aTmp, 1, T_INIDeleteSelection)
        Case "@inicreatefile"
            DoFunction = CreateFile(ReturnData(aTmp))
            GoTo ErrorCheck:
        Case "@inideletekeyvalue"
            DoFunction = DoINI(aTmp, 2, T_INIDeleteKeyValue)
            GoTo ErrorCheck
        Case "@iniwritekey"
            DoFunction = DoINI(aTmp, 3, T_INIWriteKeyValue)
            GoTo ErrorCheck
        Case "@inireadkey"
            DoFunction = DoINI(aTmp, 3, T_INIReadkeyValue)
            GoTo ErrorCheck
        Case "@isequal"
            DoFunction = StrFunction(aTmp, isSame)
        Case "@int"
            DoFunction = Int(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@isnum"
            DoFunction = IsNumeric(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@inttobool"
            DoFunction = CBool(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@getvartype"
            If Not VariableExists(aTmp) Then
                LastError 5, CurrentLinePos, "Required: Variable Name"
                Exit Function
            Else
                DoFunction = GetVariableType(aTmp)
            End If
        Case "@getenv.count"
            DoFunction = EnvCount
            GoTo ErrorCheck:
        Case "@getenv"
            aTmp = ReturnData(aTmp)
            DoFunction = Environ(dmVar(aTmp))
            GoTo ErrorCheck
        Case "@log"
            DoFunction = Log(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@lcase"
            DoFunction = LCase(ReturnData(aTmp))
           ' MsgBox ReturnData(aTmp)
           ' GoTo ErrorCheck:
        Case "@ucase"
            DoFunction = UCase(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@regdeletevalue"
            DoFunction = RegDeleteValueOrKey(aTmp, REG_VALUE)
            GoTo ErrorCheck:
        Case "@regdeletekey"
            DoFunction = RegDeleteValueOrKey(aTmp, REG_KEY)
            GoTo ErrorCheck:
        Case "@regread"
            DoFunction = RegReadWrite(aTmp, 2)
            GoTo ErrorCheck
        Case "@regwritedata"
            DoFunction = RegReadWrite(aTmp, 3)
            GoTo ErrorCheck
        Case "@right" ' Do Right string function
            DoFunction = StrFunction(aTmp, DoRight)
            GoTo ErrorCheck
        Case "@trim"
           DoFunction = Trim(ReturnData(aTmp))
           GoTo ErrorCheck
        Case "@renamefile"
            DoFunction = RenameFile(aTmp)
        Case "@rtrim"
            DoFunction = RTrim(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@ltrim"
            DoFunction = LTrim(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@left"
            DoFunction = StrFunction(aTmp, DoLeft)
            GoTo ErrorCheck
        Case "@len"
            DoFunction = Len(ReturnData(aTmp))
        Case "@prompt"
            DoFunction = PromptBox(aTmp)
            GoTo ErrorCheck
        Case "@echo"
            DoFunction = Echo(aTmp)
            GoTo ErrorCheck
        Case "@chr"
            aTmp = ReturnData(aTmp)
            DoFunction = Chr(aTmp)
            GoTo ErrorCheck:
        Case "@sizeof"
            aTmp = Right(aTmp, Len(aTmp) - 1)
            tIdx = GetArryIndexByName(aTmp)
            If tIdx = -1 Then LastError 12, CurrentLinePos, "Array": Exit Function
            DoFunction = GetArryUBound(tIdx)
            tIdx = "": aTmp = ""
        Case "@sin"
            DoFunction = Sin(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@second"
            DoFunction = Second(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@sum"
            DoFunction = Sum(aTmp)
            GoTo ErrorCheck
        Case "@shell"
            DoFunction = Shell(ReturnData(aTmp))
        Case "@space"
             DoFunction = Space(ReturnData(aTmp))
             GoTo ErrorCheck
        Case "@sub"
            DoFunction = StrFunction(aTmp, tSub)
            GoTo ErrorCheck
        Case "@time"
            DoFunction = Time
        Case "@rand"
            DoFunction = Rand(ReturnData(aTmp))
            GoTo ErrorCheck:
        Case "@lof", "@lof(" ' gets the length of a open file
            If nFilePointer <= 0 Then LastError 13, CurrentLinePos, "Bad file name or file accesses unavailable": Exit Function
            DoFunction = LOF(nFilePointer)
        Case "@settime"
            Time = ReturnData(aTmp)
            GoTo ErrorCheck
        Case "@setdate"
            Date = ReturnData(aTmp)
            GoTo ErrorCheck
        Case "@sqr"
             DoFunction = Sqr(ReturnData(aTmp))
        Case "@strweekname"
            DoFunction = StrFunction(aTmp, tWeekName, True)
        Case "@strmonthname"
            DoFunction = StrFunction(aTmp, tMonthName, True)
        Case "@strcpy"
            DoFunction = DoMid(aTmp)
            GoTo ErrorCheck
        Case "@sgn"
            DoFunction = Sgn(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@str"
            aTmp = ReturnData(aTmp)
            DoFunction = Str(aTmp)
            GoTo ErrorCheck:
        'Case "@strsplit"' Split I not sorted this out yet
            'DoFunction = "{SPLIT}-" & aTmp
        Case "@strreplace"
            DoFunction = StrReplace(aTmp)
            GoTo ErrorCheck
        Case "@mult"
            DoFunction = StrFunction(aTmp, tMult)
            GoTo ErrorCheck
        Case "@mod"
            DoFunction = StrFunction(aTmp, tMod)
            GoTo ErrorCheck
        Case "@year"
            DoFunction = Year(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@month"
            DoFunction = Month(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@minute"
            DoFunction = Minute(ReturnData(aTmp))
            GoTo ErrorCheck
       Case "@mkdir"
            DoFunction = MakeDir(aTmp)
            GoTo ErrorCheck
        Case "@oct"
            DoFunction = Oct(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@tan"
            DoFunction = Tan(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@redim"
            tArraySize = CInt(GetArrayNameIndex(aTmp, 2))
            aTmp = GetArrayNameIndex(aTmp, 1)
            tIdx = GetArryIndexByName(aTmp)
            If tIdx = -1 Then LastError 12, CurrentLinePos, "Array": Exit Function
            ReDim Preserve dwArray.mArrayItemData(tIdx).mArrayBoundData(tArraySize)
            tArraySize = 0
            GoTo ErrorCheck
        Case "@round"
            DoFunction = Round(ReturnData(aTmp))
            GoTo ErrorCheck
        Case "@pi"
            DoFunction = 3.14159265358979
            GoTo ErrorCheck
        Case "@win32"
            StrAPIname = Left(aTmp, GetPositionOfChar(aTmp, ",") - 1)
            StrAPIname = Trim(ReturnData(StrAPIname))
            aTmp = Trim(Mid(aTmp, GetPositionOfChar(aTmp, ",") + 1, Len(aTmp) - GetPositionOfChar(aTmp, ",")))
            args = Split(aTmp, ",", , vbTextCompare)
            DoFunction = CallByName(WinAPI, StrAPIname, VbMethod, args)
            GoTo ErrorCheck
        Case "@vartype"
            If Not VariableExists(aTmp) Then
                LastError 13, CurrentLinePos, "Required Variable Name"
                Exit Function
            Else
                DoFunction = Int(GetVariableType(aTmp))
            End If
        Case "@isarray"
            If Not IsVariable(Left(aTmp, 1)) Then LastError 13, CurrentLinePos, "Required: Variable Type": Exit Function
            aTmp = Right(aTmp, Len(aTmp) - 1)
            DoFunction = isdmArray(aTmp)
            GoTo ErrorCheck
        Case "@isntadmin"
            DoFunction = isAdmin
        Case "@isdate"
            DoFunction = Int(IsDate(ReturnData(aTmp)))
            GoTo ErrorCheck
        Case "@isvarempty"
            If Not VariableExists(aTmp) Then
                LastError 13, CurrentLinePos, "Required: Variable Name"
                Exit Function
            Else
                DoFunction = Not Int(Len(ReturnData(aTmp)) <> 0)
            End If
        Case "@val"
            DoFunction = Val(ReturnData(aTmp))
            GoTo ErrorCheck
        Case Else
        
            Func_Data = Function_Data
            
            StrTemp = nFunctionName
            StrTemp = Right(StrTemp, Len(StrTemp) - 1) ' remove @ from the string

            If Left(Func_Data, 1) = "@" Then
                Func_Data = Right(Func_Data, Len(Func_Data) - 1)
            Else
                Func_Data = StrTemp & Func_Data
            End If
           
            If Not ProcedureExists(StrTemp) Then
                StrTemp = ""
                LastError 9, CurrentLinePos, nFunctionName
                Exit Function
            Else
                If LenB(DoProcCall(Func_Data)) <> 0 Then
                    ReturnFuncData = ExecuteCode(GetProcedureCode(StrTemp))
                    DoFunction = ReturnFuncData
                End If
                Exit Function
            End If
    End Select

Exit Function

ErrorCheck:
    If Err Then
        If Err Then LastError 13, CurrentLinePos, Err.Description
    End If
    
End Function

Function FixQuote(StrString As String) As String
Dim StrA As String
    StrA = StrString

    If Right(StrA, 2) = Chr(34) & Chr(34) Or Left(StrA, 2) = Chr(34) & Chr(34) Then
        StrA = Right(StrA, Len(StrA) - 1)
        StrA = Left(StrA, Len(StrA) - 1)
    End If
    
   FixQuote = StrA
   StrA = ""
   
End Function

Function OpenFile(lzFile As String) As String
Dim ByteBuff() As Byte
    nFile = FreeFile
    Open lzFile For Binary As #nFile
        ReDim ByteBuff(LOF(1) - 1)
        Get #nFile, , ByteBuff()
    Close #nFile
    
    OpenFile = StrConv(ByteBuff(), vbUnicode)
    Erase ByteBuff
    
End Function

