Attribute VB_Name = "ModExecute"

Public Function ExecuteCode(Procedure As String) As String
Dim n As Long, vLines As Variant, Strln As String, StrKeyWord As String
Dim AssignData As String, StrTempHolder As String, StrMacro As String
Dim VarHolder As Variant, sCmd As String, sConst As String
Dim sConstName As String, sConstValue As String, IfSkip As Boolean, Skip As Boolean
Dim x As Long, a As String, b As String, tGotoSubList As Variant
Dim nStrLin As String
Dim FuncName As String, sTmp As String, vTmp As Variant
Dim iPos As Long, lpos As Long, VarPart As String, lVarName As String, nVarType As String * 1
Dim TheArrayName As String, TheArrayData, TheArraySize As Integer, TheArrayIndex As Integer
Dim e As Integer, D As Integer, kkIndex As Integer

    n = 0
    kkIndex = 0
    
    vLines = Split(Procedure, vbCrLf)

    For n = LBound(vLines) To UBound(vLines)
    
        CurrentLinePos = n
        Strln = Trim(vLines(n))
        
        Strln = RemoveTabs(Strln)
        IsFunction = False
        
        If GotoLabel <> "" Then
            If LCase(GotoLabel) = LCase(Strln) Then GotoLabel = ""
            GoTo nStop
        End If

        If SwitchLabel <> "" Then
            If SwitchLabel = Strln Then SwitchLabel = ""
            GoTo nStop
        End If
        
        If Skip = True Then
            If LCase(Strln) = "end if" Or LCase(Strln) = "endif" Then
                Skip = False
            End If
            GoTo nStop
        End If
        
        If IfSkip = True Then
            Select Case LCase(Strln)
                Case "end if", "endif"
                    IfSkip = False
                Case "else"
                    IfSkip = False
            End Select
            GoTo nStop
        Else
            If LCase(Strln) = "else" Then
                Skip = True
                GoTo nStop
            End If
        End If
        
        If Len(nStrLin) > 0 And LCase(Strln) = "next" Then
            ForLoopStart = ForLoopStart + 1
            If (ForLoopStart <= ForLoopEnd) Then
                n = CInt(Mid(nStrLin, 1, InStr(1, nStrLin, ":", vbTextCompare) - 1))
                SetVariableData ForLoopVar, CStr(ForLoopStart)
            End If
        End If
        
        If Len(Strln) > 0 Then
        
            iPos = GetPositionOfChar(Strln, Chr(32))
            
            StrKeyWord = Trim(LCase(Mid(Strln, 1, iPos)))
            
            
            Select Case LCase(StrKeyWord)
                Case "for" ' Used to ident a for loop
                    Strln = StripEOL(Strln)
                    DoForLoop Strln
                    nStrLin = n & ":" & Strln
                Case "appactivate"
                    Strln = StripEOL(Strln)
                    Strln = Right(Strln, Len(Strln) - 11)
                    If Not Len(Strln) > 1 Then LastError 7, n, "AppActivate": Exit Function
                    StrFunction Strln, tAppActivate
                Case "if" ' Used to ident an IF Statement
                    Strln = Right(Strln, Len(Strln) - 2)
                    If Not LCase(Right(Strln, 4)) = "then" Then
                        LastError 13, n, "Expected: Then"
                        GoTo CleanUpNow:
                        Exit Function
                    Else
                        Strln = Trim(Left(Strln, Len(Strln) - 4)) ' strip of then
                        If Not TestIF(Strln) = 1 Then IfSkip = True
                    End If
                    
                Case "const"
                    sConst = StripEOL(Strln)
                    sConst = Right(sConst, Len(sConst) - 5)
                    If GetPositionOfChar(sConst, "=") <= 0 Then
                        LastError 13, n, "Required: = "
                        GoTo CleanUpNow:
                        Exit Function
                    End If
                    
                    sConstName = Trim(Mid(sConst, 1, GetPositionOfChar(sConst, "=") - 1))
                    sConstValue = Trim(Mid(sConst, GetPositionOfChar(sConst, "=") + 1, Len(sConst)))
                    
                    If Len(sConstName) = 0 Then
                        LastError 13, n, "Required: Const Name"
                        GoTo CleanUpNow:
                        Exit Function
                    ElseIf Len(sConstValue) = 0 Then
                        LastError 13, n, "Required Const Value"
                        GoTo CleanUpNow:
                        Exit Function
                    ElseIf ConstExists(sConstName) Then
                        LastError 13, n, "Duplicate Const Name in current Scope"
                        Exit Function
                    Else
                        sConstValue = Chr(34) & sConstValue & Chr(34)
                        sConstValue = FixQuote(sConstValue)
                        
                        dmConsts.Count = dmConsts.Count + 1
                        ReDim Preserve dmConsts.ConstName(dmConsts.Count)
                        ReDim Preserve dmConsts.ConstValue(dmConsts.Count)
                        dmConsts.ConstName(dmConsts.Count) = sConstName
                        dmConsts.ConstValue(dmConsts.Count) = sConstValue

                        sConstName = ""
                        sConstValue = ""
                        sConst = ""
                    End If
  
                Case "copyfile"
                    Strln = StripEOL(Strln)
                    Strln = Right(Strln, Len(Strln) - 8)
                    TCopyFile Strln
                Case "call"
                    Strln = StripEOL(Strln)
                    
                    StrTempHolder = DoProcCall(Trim(Right(Strln, Len(Strln) - 4)))
                    If LenB(StrTempHolder) <> 0 Then
                        ExecuteCode GetProcedureCode(StrTempHolder)
                        StrTempHolder = ""
                    End If
                Case "destroy"
                    Strln = StripEOL(Strln)
                    Strln = Right(Strln, Len(Strln) - 7)
                    DoFunction "@destroy", Trim(Strln), True
                Case "var" ' we found a variable
                   
                    If Not CBool(GetEOL(Strln)) Then ' check that we have end of line ;
                        LastError 0, n
                        Exit Function
                    Else
                        ' extract variable name and type
                        lVarName = Trim(Mid(Strln, iPos + 1, Len(Strln) - iPos - 1))
                        If Not Len(lVarName) > 1 Then LastError 3, n: GoTo CleanUpNow: Exit Function
       
                        VarHolder = Split(lVarName, ",")

                        For x = LBound(VarHolder) To UBound(VarHolder)
                            nVarType = Left(VarHolder(x), 1) ' Get the variable type
                            
                            If GetVariableType(nVarType) = vbNullChar Then LastError 5, n, nVarType:  Exit Function ' exit if variable type not found
                            ' new code
                           If HasBrackets(CStr(VarHolder(x))) Then
                           
                                sTmp = VarHolder(x)
                                sTmp = Right(sTmp, Len(sTmp) - 1)
                                e = GetPositionOfChar(sTmp, "("): D = GetPositionOfChar(sTmp, ")")
                                
                                If Not CBool(e > 0 And D > 0) Then
                                    LastError 5, n, sTmp
                                    Exit Function
                                Else
                                    TheArrayName = Trim(Mid(sTmp, 1, e - 1))
                                    If GetArrayNameIndex(sTmp, 2) = "" Then
                                        TheArraySize = -1
                                    Else
                                        TheArraySize = CInt(Trim(Mid(sTmp, e + 1, D - e - 1)))
                                    End If
                                    
                                    dwArray.Count = dwArray.Count + 1
                                    ReDim Preserve dwArray.mArrayNames(dwArray.Count)
                                    ReDim Preserve dwArray.mArrayItemData(dwArray.Count)
                                    
                                    If Not TheArraySize = -1 Then ' check if the array has been resized
                                        ReDim Preserve dwArray.mArrayItemData(dwArray.Count).mArrayBoundData(TheArraySize)
                                    End If
                                    
                                    ' If Err Then LastError 13, n, Err.Description: Exit Function
                                    dwArray.mArrayNames(dwArray.Count) = TheArrayName
                                    e = 0: D = 0
                                    sTmp = ""
                                    TheArrayName = ""
                                    TheArraySize = 0
                                    
                                End If
                                ' end of new code
                            Else
                                Variables.Count = Variables.Count + 1 ' Add one to out variable count
                                ' Time to resize some arrays
                                ReDim Preserve Variables.VariableName(Variables.Count)
                                ReDim Preserve Variables.VariableType(Variables.Count)
                                ReDim Preserve Variables.Variabledata(Variables.Count)
                                
                                Dim VarAssign As String
                                
                                If GetAssignPos(CStr(VarHolder(x)), "=") <> 0 Then
                                    VarAssign = GetFromAssignPos(CStr(VarHolder(x)), GetAssignPos(CStr(VarHolder(x)), "="))
                                    VarAssign = Right(VarAssign, Len(VarAssign) - 1)
                                    VarAssign = Trim(VarAssign)
                                    Variables.VariableName(Variables.Count) = Trim(Mid(CStr(VarHolder(x)), 1, GetAssignPos(CStr(VarHolder(x)), "=") - 1))
                                Else
                                    Variables.VariableName(Variables.Count) = VarHolder(x)
                                End If
                                
                                Variables.VariableType(Variables.Count) = GetVariableType(nVarType)
                            
                                If GetVariableType(nVarType) = vbInteger Then
                                    If VarAssign = "" Then VarAssign = 1
                                    If Not IsNumeric(VarAssign) Then LastError 10, n: GoTo CleanUpNow: Exit Function
                                    Variables.Variabledata(Variables.Count) = CInt(VarAssign) ' set integer types to 1
                                    VarAssign = ""
                                Else
                                    Variables.Variabledata(Variables.Count) = ReturnData(VarAssign)
                                    VarAssign = ""
                                End If
                            End If
                        Next
                        
                        x = 0
                        Erase VarHolder
                    End If
                    VarAssign = ""
                    
                    lVarName = ""
                    nVarType = ""
                ' still in case
                Case "deletefile"
                    Strln = StripEOL(Strln)
                    DoFunction "@DeleteFile", Trim(Right(Strln, Len(Strln) - 10)), True
                Case "echo"
            
                    If Not Len(Trim(Mid(Strln, 5, Len(Strln)))) > 1 Then
                        LastError 7, n, "echo"
                        GoTo CleanUpNow:
                        Exit Function
                    Else
                        If Right(Strln, 1) = ";" Then Strln = Left(Strln, Len(Strln) - 1)
                    End If
                    
                    Echo Trim(Mid(Strln, 5, Len(Strln)))

                Case "pause"
                    If Not Len(Trim(Mid(Strln, 5, Len(Strln)))) > 2 Then
                        LastError 7, n, "pause"
                        Exit Function
                    Else
                        Strln = StripEOL(Strln)
                        StrTempHolder = GetDataFormParam(Trim(Mid(Strln, 6, Len(Strln))))
                        
                        If Not IsNumeric(StrTempHolder) Then
                            LastError 10, n, StrTempHolder
                            Exit Function
                        Else
                            Pause Val(StrTempHolder)
                            StrTempHolder = ""
                        End If
                    End If
                
                Case "on"
                    Strln = StripEOL(Strln)
                    StrTempHolder = LTrim(Right(Strln, Len(Strln) - 2))
                    lVarName = Mid(StrTempHolder, 1, GetPositionOfChar(StrTempHolder, Chr(32)) - 1)
                    
                    If Not VariableExists(lVarName) Then
                        LastError 13, n, "Required: Variable"
                        Exit Function
                    Else
                        sCmd = Trim(Right(StrTempHolder, Len(StrTempHolder) - Len(lVarName)))
                        StrTempHolder = sCmd
                        sCmd = RTrim(Mid(sCmd, 1, GetPositionOfChar(sCmd, " ")))
                        StrTempHolder = LTrim(Right(StrTempHolder, Len(StrTempHolder) - Len(sCmd)))
                        If LCase(sCmd) = "goto" Or LCase(sCmd) = "gosub" Then
                            tGotoSubList = Split(StrTempHolder, ",")
                            On Error Resume Next
                        
                            If Err.Number = 9 Then
                                LastError 13, n, Err.Description
                                Exit Function
                            Else
                                GotoLabel = Trim(tGotoSubList(GetVariableData(lVarName) - 1) & ":")
                                If GotoLabel = "" Then GotoLabel = "default:"
                                Erase tGotoSubList
                            End If
                        Else
                            LastError 13, n, "Required: GoTo or GoSub"
                            Exit Function
                        End If
                    End If

                Case "open"
                    Strln = StripEOL(Strln)
                    DoFSO Strln
                Case "goto"
                    GotoLabel = LTrim(Right(Strln, Len(Strln) - 4))
                Case "switch"
                   ' First check for the end of switch
                    If Not InStr(1, Procedure, "end switch", vbTextCompare) > 0 Then
                        LastError 13, n, "Switch without End Switch"
                        Exit Function
                    End If
 
                   SwitchLabel = StripEOL(LTrim(Right(Strln, Len(Strln) - 6)))
                   SwitchLabel = ReturnData(CStr(SwitchLabel))
                   SwitchLabel = "case " & SwitchLabel & ":"
                Case "shell"
                    Strln = StripEOL(Strln)
                    Strln = Right(Strln, Len(Strln) - 5)
                    If Not Len(Strln) > 1 Then LastError 7, n, "shell": Exit Function
                    DoFunction "@shell", Strln, True
                Case "sendkeys"
                    Strln = StripEOL(Strln)
                    Strln = Right(Strln, Len(Strln) - 8)
                    If Not Len(Strln) > 1 Then LastError 7, n, "shell": Exit Function
                    StrFunction Strln, tSendKeys, True
                Case "settime"
                    Strln = StripEOL(Strln)
                    DoFunction "@settime", Trim(Right(Strln, Len(Strln) - 7)), True
                Case "setdate"
                    Strln = StripEOL(Strln)
                    DoFunction "@setdate", Trim(Right(Strln, Len(Strln) - 7)), True
                Case "mkdir"
                    MakeDir StripEOL(Trim(Right(Strln, Len(Strln) - 5)))
                Case "rmdir"
                    RemoveDir StripEOL(Trim(Right(Strln, Len(Strln) - 5)))
                Case "chdir"
                   ChangeDir StripEOL(Trim(Right(Strln, Len(Strln) - 5)))
                Case "clipboard.settext"
                    Strln = StripEOL(Strln)
                    a = "@" & Trim(Left(Strln, 18))
                    b = Right(Strln, Len(Strln) - 18)
                    DoFunction a, b, True
                    a = ""
                    b = ""
                Case "return"
                    Strln = StripEOL(Strln)
                    a = Strln
                    a = Trim(Right(a, Len(a) - 6))
                    a = ReturnData(a)
                    ExecuteCode = a
                    a = ""
                Case "redim"
                    Strln = StripEOL(Strln)
                    Strln = Right(Strln, Len(Strln) - 5)
                    DoFunction "@redim", Trim(Strln), True
                Case "regwritedata"
                    Strln = StripEOL(Strln)
                    Strln = Right(Strln, Len(Strln) - 12)
                    DoFunction "@regwritedata", Trim(Strln), True
                Case "regdeletekey"
                    Strln = StripEOL(Strln)
                    Strln = Right(Strln, Len(Strln) - 12)
                    DoFunction "@regdeletekey", Trim(Strln), True
                Case "regdeletevalue"
                    Strln = StripEOL(Strln)
                    Strln = Right(Strln, Len(Strln) - 14)
                    DoFunction "@regdeletevalue", Trim(Strln), True
                Case "iniwritekey"
                    Strln = StripEOL(Strln)
                    Strln = Right(Strln, Len(Strln) - 11)
                    DoFunction "@iniwritekey", Trim(Strln), True
                Case "inicreatefile"
                    Strln = StripEOL(Strln)
                    Strln = Right(Strln, Len(Strln) - 13)
                    DoFunction "@inicreatefile", Trim(Strln), True
                Case "inideleteselection"
                    Strln = StripEOL(Strln)
                    Strln = Right(Strln, Len(Strln) - 18)
                    DoFunction "@inideleteselection", Trim(Strln), True
                    
            End Select

         If VarOrFunction(Strln) = sVariable Then
            lVarName = Trim(StrKeyWord)
            StrTempHolder = GetFromAssignPos(Strln, GetAssignPos(Strln, "=") + 1)
            ' next we check if a variable or function or plain text is been assigned

            Select Case VarOrFunction(StrTempHolder)
            
                Case sFunction ' found a function
                    FuncName = GetFunctionName(StrTempHolder) ' Extract the function name
                    sReturnData = DoFunction(FuncName, StripEOL(StrTempHolder)) ' Send function name and the functions data to be processed

       
                    If Not VariableExists(lVarName) Then
                        LastError 6, CurrentLinePos, lVarName
                        GoTo CleanUpNow:
                        Exit Function
                    End If
                    
                    If Not HasBrackets(lVarName) Then
                        'Store the variables new data
                        SetVariableData lVarName, sReturnData
                        sReturnData = ""
                        AssignData = ""
                   Else
                        TheArrayName = GetArrayNameIndex(Strln, 1)
                        TheArrayIndex = Val(GetArrayNameIndex(Strln, 2))
                        TheArrayData = ReturnData(GetArrayNameIndex(Strln, 3))
                        
                        kkIndex = GetArryIndexByName(TheArrayName)
                        dwArray.mArrayItemData(kkIndex).mArrayBoundData(TheArrayIndex) = TheArrayData
                        kkIndex = 0
                        If Err Then
                            LastError 13, n, Err.Description
                            Exit Function
                        End If
                   End If
                   

                Case sTextOnly ' an assignment of text was found eg $myvar = "some text"
                    
                
                    AssignData = GetFromAssignPos(Strln, GetAssignPos(Strln, "=") + 1)
                    AssignData = StripEOL(AssignData) ' Get the assigned data
                    sReturnData = GetDataFormParam(AssignData) ' Process the extracted assign data
                    
                    ' check if the variable is in the variables list
                    
                    
                    If Not VariableExists(lVarName) Then
                        ' theres lines of code check if a variable has been added
                        LastError 6, CurrentLinePos, lVarName
                        GoTo CleanUpNow:
                        Exit Function
                    End If
                    
                    On Error Resume Next
                    If GetVariableType(lVarName) = vbInteger Then
                        sReturnData = CDbl(sReturnData)
                        If Err Then
                            LastError 13, n, Err.Description
                            Exit Function
                        End If
                    End If
       
       
                    If Not HasBrackets(lVarName) Then
                        'Store the variables new data
                        SetVariableData lVarName, sReturnData
                        sReturnData = ""
                        AssignData = ""
                    Else
                        
                        TheArrayName = GetArrayNameIndex(Strln, 1)
                        TheArrayIndex = Val(GetArrayNameIndex(Strln, 2))
                        TheArrayData = ReturnData(GetArrayNameIndex(Strln, 3))
                        
                        kkIndex = GetArryIndexByName(TheArrayName)
                        dwArray.mArrayItemData(kkIndex).mArrayBoundData(TheArrayIndex) = TheArrayData
                        kkIndex = 0
                        If Err Then
                            LastError 13, n, Err.Description
                            Exit Function
                        End If
                    End If
                    
                    
                Case sVariable 'an assignment of variable type was found eg $myvar = $newvar
                    'sReturnData = ""
                    AssignData = GetFromAssignPos(Strln, GetAssignPos(Strln, "=") + 1)
                    AssignData = StripEOL(AssignData)
                    
                    ' INC code
                    If Right(AssignData, 2) = "++" Then
                        k = 0
                        lVarName = Trim(Left(AssignData, Len(AssignData) - 2))
                        IncCounter = IncCounter + 1
                        k = Val(GetVariableData(lVarName)) + IncCounter
                        AssignData = Chr(34) & CStr(k) & Chr(34)
                    End If
                    
                    If Right(AssignData, 2) = "--" Then
                        k = 0
                        lVarName = Trim(Left(AssignData, Len(AssignData) - 2))
                        IncCounter = IncCounter + 1
                        k = Val(GetVariableData(lVarName)) - IncCounter
                        AssignData = Chr(34) & CStr(k) & Chr(34)
                    End If
                    ' End INC Code

                    If GetVariableType(lVarName) = vbInteger Then
                        sReturnData = GetDataFormParam(AssignData)
                    ElseIf LCase(Left(AssignData, Len(lVarName))) = LCase(lVarName) Then
                        sReturnData = sReturnData & GetDataFormParam(AssignData)
                    Else
                        sReturnData = ReturnData(AssignData)
                    End If

                    ' check if the variable is in the variables list
                    If Not VariableExists(lVarName) Then
                        LastError 6, CurrentLinePos, lVarName
                        GoTo CleanUpNow:
                        Exit Function
                    End If
                    
                    If Not HasBrackets(lVarName) Then
                        'Store the variables new data
                        SetVariableData lVarName, sReturnData
                        AssignData = ""
                        k = 0
                    Else
                        TheArrayName = GetArrayNameIndex(Strln, 1)
                        TheArrayIndex = Val(GetArrayNameIndex(Strln, 2))
                        TheArrayData = ReturnData(GetArrayNameIndex(Strln, 3))
                        
                        kkIndex = GetArryIndexByName(TheArrayName)
                        dwArray.mArrayItemData(kkIndex).mArrayBoundData(TheArrayIndex) = TheArrayData
                        kkIndex = 0
                        If Err Then
                            LastError 13, n, Err.Description
                            Exit Function
                        End If
                    End If
                    
            End Select
        '//************************************************************************************************************************
         End If
         
        ' this select case I used for picking out small items eg beep,break,closefile,next
        StrMacro = LCase(StripEOL(Strln))

        Select Case StrMacro
            Case "beep"
                Beep ' Do I need to explain this one
            Case "break" ' Stop the code at a block like vbs exit sub
                Exit Function
            Case "clipboard.reset"
            
                Clipboard.Clear ' clear the contents of the clipboard
            Case "closefile" ' used for closeing a file
                If nFilePointer > 0 Then Close #nFilePointer: nFileReadWriteAccess = -1
            Case "end"
                'End ' ends the program
                'Unload
        End Select
        
        If GetPositionOfChar(Strln, ",") > 0 Then
            sCmd = Trim(Mid(Strln, 1, GetPositionOfChar(Strln, ",") - 1))
        End If
    
        Select Case LCase(sCmd)
            Case "write" ' used for writeing data to a open file
                If nFilePointer <= 0 Then LastError 13, CurrentLinePos, "Bad file name or file accesses unavailable": Exit Function
                Strln = StripEOL(Strln)
                ' strip of the command name
                Strln = Right(Strln, Len(Strln) - 5)
                FsoDoWrite Strln
            Case "get" ' used to read data from a file
                If nFilePointer <= 0 Then LastError 13, CurrentLinePos, "Bad file name or file accesses unavailable": Exit Function
                Strln = StripEOL(Strln)
                Strln = Right(Strln, Len(Strln) - 3)
                FsoDoRead Strln
                
        End Select

        ' still incode block
        sCmd = ""
        Strln = ""
        StrTempHolder = ""
        End If
nStop:
    Next n
    
CleanUpNow:
On Error Resume Next
n = 0
x = 0
iPos = 0
lpos = 0

Erase vLines
Erase vTmp
Strln = ""
sTmp = ""
VarPart = ""
lVarName = ""
StrKeyWord = ""
AssignData = ""
StrTempHolder = ""
StrMacro = ""
sConstName = ""
sConstValue = ""
FuncName = ""

End Function

