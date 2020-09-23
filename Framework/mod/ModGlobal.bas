Attribute VB_Name = "ModGlobal"
Public sInputBoxPrompt As String ' used for input box and message box
Public sInputBoxTitle As String ' used for inputbox and message box
Public sInputBoxButtons As Integer ' inputbox or messagebox button tyles
Public FunctionMembers() As String
Public sReturnData As String
Public StrCommandLine As String

Public TIniFile As New clsINI

Public CurrentLinePos As Long
Global IncCounter As Integer

Public nFilePointer As Long
Public nFileReadWriteAccess As FileAccess
Public scriptFullPath As String 'new

Public GotoLabel As String
Public SwitchLabel As String
Public ReturnFuncData As String

Public ForLoopStart As Integer, ForLoopEnd As Integer, ForLoopCounter As Integer, ForLoopVar As String

Private Type TProcedure
    Count As Integer
    TName() As String
    TCode() As String
    TParms() As String
    TReturnData() As String
End Type

Public WinAPI As New clsAPI
Public Procedures As TProcedure
Public IsFunction As Boolean

Public Sub GetProcedures(StrCodeBuff As String)
Dim I As Long, j As Long, vLines As Variant, ProgStart As String, sLine As String
Dim sPram As String, StrTemp As String, VarTmp As Variant, TmpVarType As String

On Error Resume Next

Dim lpos As Integer, hPos As Integer, dPos As Long, sPos As Long
Dim a As Integer, b As Integer

    Procedures.Count = 0
    ReDim Preserve Procedures.TCode(0)
    ReDim Preserve Procedures.TName(0)
    
    If Len(Trim(StrCodeBuff)) = 0 Then Exit Sub
    
    vLines = Split(StrCodeBuff, vbCrLf)
    
    For I = LBound(vLines) To UBound(vLines)
        sLine = Trim(vLines(I))

        If LCase(Left(sLine, 9)) = "procedure" Or LCase(Left(sLine, 8)) = "function" Then
            iPos = GetPositionOfChar(sLine, Chr(32))
            If iPos > 0 Then ProgStart = Trim(Mid(sLine, iPos, Len(sLine)))
            
            If Not CBool(GetEOL(ProgStart)) Then
                LastError 0, I
                GoTo DoCelanup:
                Exit Sub
            End If
            
            sPram = Trim(GetBetweenBrackets(ProgStart))
            dPos = GetPositionOfChar(ProgStart, "(")
            epos = GetPositionOfChar(ProgStart, ")")
            
            If Not dPos > 0 Then
                LastError 1, I
                GoTo DoCelanup:
                Exit Sub
            End If
            
            If Not epos > 0 Then
                GoTo DoCelanup:
                LastError 2, I
                Exit Sub
            End If
            
            StrTemp = Trim(Mid(ProgStart, 1, dPos - 1))
            
            If Len(StrTemp) = 0 Then LastError 3, I: GoTo DoCelanup: Exit Sub
            
            lpos = InStr(1, StrCodeBuff, ProgStart, vbTextCompare)
            hPos = InStr(lpos + 1, StrCodeBuff, "end " & StrTemp & ";", vbTextCompare)
            
            If Not hPos > 0 Then
                LastError 4, I, StrTemp
                Exit Sub
            End If
            
            a = Len(ProgStart) + 2
            b = Len("end " & StrTemp & ";")
            
            
            Procedures.Count = Procedures.Count + 1
            ReDim Preserve Procedures.TCode(Procedures.Count)
            ReDim Preserve Procedures.TName(Procedures.Count)
            ReDim Preserve Procedures.TParms(Procedures.Count)
            Procedures.TName(Procedures.Count) = StrTemp
            Procedures.TCode(Procedures.Count) = Trim(Mid(StrCodeBuff, lpos + a, hPos - lpos - a))
            Procedures.TParms(Procedures.Count) = sPram
            Procedures.TReturnData(Procedures.Count) = ""
            
            If Len(sPram) > 0 Then
                VarTmp = Split(sPram, ",")
                
                For j = LBound(VarTmp) To UBound(VarTmp)
                    TmpVarType = Left(VarTmp(j), 1)
                    
                    Variables.Count = Variables.Count + 1
                    ReDim Preserve Variables.VariableName(Variables.Count)
                    ReDim Preserve Variables.VariableType(Variables.Count)
                    ReDim Preserve Variables.Variabledata(Variables.Count)
                    
                    Variables.VariableName(Variables.Count) = VarTmp(j)
                    Variables.VariableType(Variables.Count) = GetVariableType(TmpVarType)
                    Variables.Variabledata(Variables.Count) = ""
                Next
            End If
        End If
    Next
    
DoCelanup:
    
    Erase vLines
    Erase VarTmp
    TmpVarType = ""
    sLine = ""
    iPos = 0
    I = 0
    j = 0
    StrTemp = ""
    sLine = ""
    
End Sub

Public Function GetModules(StrCodeBuff As String) As String
Dim I As Integer, vList As Variant, sLine As String
Dim ModFile As String, AbsPathFile As String
Dim sBuffer As String

    vList = Split(StrCodeBuff, vbCrLf)
    For I = LBound(vList) To UBound(vList)
        sLine = vList(I)
        If LCase(Left(sLine, 6)) = "module" Then
            sLine = Trim(Right(sLine, Len(sLine) - 6))
            
            If (Left(sLine, 1) = "<" And Right(sLine, 1) = ">") = False Then
                FoundError = True
                ModFile = GetDirective(sLine)
                ErrorStr = "Unable to include file: " & ModFile & " inavild directive format."
                GoTo CleanUp:
                Exit Function
            Else
                ModFile = GetDirective(sLine)
                If InStr(1, ModFile, "\") Then
                    AbsPathFile = ModFile
                Else
                    AbsPathFile = FixPath(App.Path) & "modules\" & ModFile
                End If
                
                If Not IsFileHere(AbsPathFile) Then
                    FoundError = True
                    ModFile = AbsPathFile
                    ErrorStr = "Unable to open file: " & AbsPathFile
                    GoTo CleanUp:
                    Exit Function
                Else
                    sBuffer = sBuffer & OpenFile(AbsPathFile)
                End If
            End If
        End If
    Next
    
    GetModules = sBuffer
    
CleanUp:
    I = 0
    ModFile = ""
    sLine = ""
    sBuffer = ""
    AbsPathFile = ""
    Erase vList
    
    
End Function

Public Sub GetEnums(StrCodeBuff As String)
Dim vLine As Variant, slnCheck As String, iPos As String
Dim iEnd As Integer, iStart As Integer, S As String
Dim EnumName As String, EnumDataList As Variant, j As Integer, Strln As String
Dim AssignPos As Integer
Dim StrA As String, StrB As String

Dim I As Long

    vLine = Split(StrCodeBuff, vbCrLf)
    
    For I = LBound(vLine) To UBound(vLine)
        iPos = GetPositionOfChar(CStr(vLine(I)), " ")
        If iPos > 0 Then
            slnCheck = Trim(Left(vLine(I), iPos))
            If slnCheck = "enum" Then iStart = iPos
        End If

        iEnd = GetPositionOfChar(CStr(vLine(I)), "}")
        
        If iStart > 0 Then
            S = S & vLine(I) & vbCrLf
        End If
        
        If iStart = 0 And Len(S) > 1 Then
            S = LTrim(Right(S, Len(S) - 4))
            ' check for open bracket {
            If GetPositionOfChar(S, "{") <= 0 Then
                LastError 10, I, "missing {"
                GoTo TidyUp:
                Exit Sub
            End If
            
            EnumName = LCase(Left(S, GetPositionOfChar(S, "{") - 1))
            EnumName = RTrim(EnumName)
        
            S = Right(S, Len(S) - Len(EnumName))
            S = Replace(S, "{", "")
            S = Replace(S, "}", "")
            S = Trim(S)
            
            If Left(S, 2) = vbCrLf Then S = Right(S, Len(S) - 2)
            If Right(S, 2) = vbCrLf Then S = Left(S, Len(S) - 2)
            
            If Len(S) = 0 Then
                LastError 13, I, "Enum members required"
                GoTo TidyUp:
                Exit Sub
            Else
                dmEnum.Count = dmEnum.Count + 1 ' Add one to our Counter
                ReDim Preserve dmEnum.EnumNames(dmEnum.Count) ' Resize Enum Count
                dmEnum.EnumNames(dmEnum.Count) = EnumName ' Enum Name
                
                EnumDataList = Split(S, vbCrLf)
                For j = LBound(EnumDataList) To UBound(EnumDataList)
                    Strln = Trim(EnumDataList(j))
                    If Len(Strln) > 0 Then
                        If Not GetEOL(Strln) <> 0 Then LastError 0, I, Strln: Exit Sub
                        AssignPos = GetPositionOfChar(Strln, "=") 'enum item value
                        
                        If AssignPos Then
                            StrA = Trim(Left(Strln, AssignPos - 1)) ' Enum Listitem Name
                            StrB = Trim(Mid(Strln, AssignPos + 1, Len(Strln) - AssignPos - 1))
                            If Len(StrB) = 0 Then StrB = vbNullString
                        Else
                            StrA = RTrim(Left(Strln, Len(Strln) - 1)) ' Enum Listitem Name
                            StrB = CStr(j) 'enum item value
                        End If
                        
                       ReDim Preserve dmEnum.ItemData(dmEnum.Count)
                       
                       dmEnum.ItemData(dmEnum.Count).Count = dmEnum.ItemData(dmEnum.Count).Count + 1
                        
                       ReDim Preserve dmEnum.ItemData(dmEnum.Count).EnumItemName(dmEnum.ItemData(dmEnum.Count).Count)
                       ReDim Preserve dmEnum.ItemData(dmEnum.Count).EnumDataValue(dmEnum.ItemData(dmEnum.Count).Count)
                      
                       dmEnum.ItemData(dmEnum.Count).EnumItemName(dmEnum.ItemData(dmEnum.Count).Count) = StrA
                       dmEnum.ItemData(dmEnum.Count).EnumDataValue(dmEnum.ItemData(dmEnum.Count).Count) = StrB
                    End If
                Next
            End If
            
            S = ""
            iStart = 0
        End If
        
        If iEnd > 0 Then iStart = 0
    Next
    
TidyUp:
    Erase vLine
    I = 0
    StrA = ""
    StrB = ""
    iEnd = 0
    iStart = 0
    iPos = 0
    slnCheck = ""
    S = ""

End Sub

Public Function ProcedureIndex(ProcedureName As String) As Integer
Dim n As Integer
    ProcedureIndex = -1
    If Procedures.Count = 0 Then Exit Function
    For n = 0 To Procedures.Count
        If LCase(Procedures.TName(n)) = LCase(ProcedureName) Then
            ProcedureIndex = n
            Exit For
        End If
    Next
    n = 0
End Function

Public Function ProcedureGotParms(ProcedureName As String) As Boolean
Dim iCheck As Integer
    iCheck = ProcedureIndex(ProcedureName)
    If iCheck = -1 Then Exit Function
    ProcedureGotParms = LenB(Procedures.TParms(iCheck)) <> 0
End Function

Public Function ProcedureGetParms(ProcedureName As String) As String
Dim iCheck As Integer
    ProcedureGetParms = ""
    iCheck = ProcedureIndex(ProcedureName)
    If iCheck = -1 Then Exit Function
    ProcedureGetParms = Procedures.TParms(iCheck)
End Function

Function ProcedureExists(ProcedureName As String) As Boolean
    ProcedureExists = False
    ProcedureExists = ProcedureIndex(ProcedureName) <> -1
End Function

Function GetProcedureCode(ProcedureName As String) As String
    If Procedures.Count = 0 Then Exit Function
    If ProcedureIndex(ProcedureName) <> -1 Then
        GetProcedureCode = Procedures.TCode(ProcedureIndex(ProcedureName))
    End If
End Function

Public Sub LoadFunctionList()
ReDim FunctionMembers(106)
' All the code can be remove from this function if you want

' I added this at the start as a function check but did not use it
' But Know I am thinking that I may keep it as I can access this
' From my IDE so I may use it as some sort of Object Broswer like in VB

    'in built functions list

    FunctionMembers(0) = "@chr"
    FunctionMembers(1) = "@space"
    FunctionMembers(2) = "@ucase"
    FunctionMembers(3) = "@lcase"
    FunctionMembers(4) = "@trim"
    FunctionMembers(5) = "@rtrim"
    FunctionMembers(6) = "@ltrim"
    FunctionMembers(7) = "@len"
    FunctionMembers(8) = "@prompt"
    FunctionMembers(9) = "@echo"
    FunctionMembers(10) = "@space"
    FunctionMembers(11) = "@asc"
    FunctionMembers(12) = "@date"
    FunctionMembers(13) = "@time"
    FunctionMembers(14) = "@array"
    FunctionMembers(15) = "@rand"
    FunctionMembers(16) = "@right"
    FunctionMembers(17) = "@left"
    FunctionMembers(18) = "@FillStr"
    FunctionMembers(19) = "@str"
    FunctionMembers(20) = "@int"
    FunctionMembers(21) = "@IntToBool"
    FunctionMembers(22) = "@LOF"
    FunctionMembers(23) = "@getenv"
    FunctionMembers(24) = "@curdir"
    FunctionMembers(25) = "@isEqual"
    FunctionMembers(26) = "@Add"
    FunctionMembers(27) = "@sub"
    FunctionMembers(28) = "@Mult"
    FunctionMembers(29) = "@div"
    FunctionMembers(30) = "@exp"
    FunctionMembers(31) = "@isNum"
    FunctionMembers(32) = "@BoolToInt"
    FunctionMembers(33) = "@mod"
    FunctionMembers(34) = "@app.path"
    FunctionMembers(35) = "@chrPos"
    FunctionMembers(36) = "@day"
    FunctionMembers(37) = "@Year"
    FunctionMembers(38) = "@Month"
    FunctionMembers(39) = "@Hour"
    FunctionMembers(40) = "@Minute"
    FunctionMembers(41) = "@Second"
    FunctionMembers(42) = "@FileSize"
    FunctionMembers(43) = "@DeleteFile"
    FunctionMembers(44) = "@DirectoryExists"
    FunctionMembers(45) = "@Format"
    FunctionMembers(46) = "@hex"
    FunctionMembers(47) = "@Clipboard.GetText"
    FunctionMembers(48) = "@Clipboard.SetText"
    FunctionMembers(49) = "@Dec2Bin"
    FunctionMembers(50) = "@Bin2Dec"
    FunctionMembers(51) = "@shell"
    FunctionMembers(52) = "@command"
    FunctionMembers(53) = "@dir"
    FunctionMembers(54) = "@GetHwnd"
    FunctionMembers(55) = "@sin"
    FunctionMembers(56) = "@cos"
    FunctionMembers(57) = "@GetVarType"
    FunctionMembers(58) = "@sqr"
    FunctionMembers(59) = "@oct"
    FunctionMembers(60) = "@tan"
    'FunctionMembers(61) = "@countif"
    FunctionMembers(62) = "@log"
    FunctionMembers(63) = "@round"
    FunctionMembers(64) = "@IsVarEmpty"
    FunctionMembers(65) = "@winapi"
    FunctionMembers(66) = "@val"
    FunctionMembers(67) = "@renamefile"
    FunctionMembers(68) = "@eval"
    FunctionMembers(69) = "@abs"
    FunctionMembers(70) = "@mkdir"
    FunctionMembers(71) = "@dec"
    FunctionMembers(72) = "@pi"
    FunctionMembers(73) = "@concat"
    FunctionMembers(74) = "@vartype"
    FunctionMembers(75) = "@strcpy"
    FunctionMembers(76) = "@isdate"
    FunctionMembers(77) = "@atn"
    FunctionMembers(78) = "@Sgn"
    FunctionMembers(79) = "strReplace"
    FunctionMembers(80) = "@sum"
    FunctionMembers(81) = "@settime"
    FunctionMembers(82) = "@setdate"
    FunctionMembers(83) = "@average"
    FunctionMembers(84) = "@FixNum"
    FunctionMembers(85) = "@StrMonthName"
    FunctionMembers(86) = "@StrWeekName"
    FunctionMembers(87) = "@FileDateTime"
    FunctionMembers(88) = "@SizeOf"
    FunctionMembers(89) = "@Destroy"
    FunctionMembers(90) = "@IsNTAdmin"
    FunctionMembers(91) = "@isArray"
    FunctionMembers(92) = "@strsplit"
    FunctionMembers(93) = "@GetSpecialFolder"
    FunctionMembers(94) = "@regread"
    FunctionMembers(95) = "@DriveList"
    FunctionMembers(96) = "@GetDriveCount"
    FunctionMembers(97) = "@BrowseForFolder"
    FunctionMembers(98) = "@getenv.count"
    FunctionMembers(99) = "@RegWriteData"
    FunctionMembers(100) = "@RegDeleteKey"
    FunctionMembers(101) = "@regdeletevalue"
    FunctionMembers(102) = "@INIReadKey"
    FunctionMembers(103) = "@IniWriteKey"
    FunctionMembers(104) = "@IniDeleteKeyValue"
    FunctionMembers(105) = "@IniCreatefile"
    FunctionMembers(106) = "@INIDeleteSelection"
    
End Sub

Public Sub Reset()
    Set WinAPI = Nothing
    
    FoundError = False
    IncCounter = 0
    CurrentLinePos = 0
    sInputBoxButtons = 0
    sInputBoxPrompt = ""
    sInputBoxTitle = ""
    sReturnData = ""
    GotoLabel = ""
    SwitchLabel = ""
    ForLoopStart = 0
    ForLoopEnd = 0
    ForLoopVar = ""
    ForLoopCounter = 0
    ErrorStr = ""
    sReturnData = ""
    dmEnum.Count = 0
    dmConsts.Count = 0
    Procedures.Count = 0
    Variables.Count = 0
    
    dwArray.Count = 0
    Erase dwArray.mArrayItemData()
    Erase dwArray.mArrayNames()
    
    Erase dmEnum.EnumNames()
    Erase dmEnum.ItemData()
    Erase dmEnum.EnumNames()
    Erase dmConsts.ConstName()
    Erase dmConsts.ConstValue()
    Erase Procedures.TCode()
    Erase Procedures.TName()
    Erase Procedures.TParms()
    Erase Variables.Variabledata()
    Erase Variables.VariableName()
    Erase Variables.VariableType()
    
End Sub
