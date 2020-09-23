Attribute VB_Name = "modEVal"
Dim Result As Long

Private Function IsOperator(StrOp As String) As Boolean
' we use this to check for Operator types
    Select Case StrOp
        Case "&" 'Concatenation string operator
            IsOperator = True
        Case "+", "-", "*", "/", "\", "^", "&" 'Arithmetic operators 1
            IsOperator = True
        Case "<", ">", "=" 'Relational operators
             IsOperator = True
        Case "<=", ">=", "<>", "or"
            IsOperator = True
        Case "not", "and", "xor", "shl", "shr", "mod", "div" 'Logical bitwise operators
            IsOperator = True
        Case "like"
            IsOperator = True
        Case Else
            IsOperator = False
    End Select

End Function

Function Evaluate(Expression As String) As Variant
Dim iCounter As Integer
Dim StrCh1 As String * 1, StrCh2 As String * 3, StrCh3 As String * 2, StrCh4 As String * 4
Dim Result As Variant, sNum As Long
Dim Operator As String
On Error Resume Next

    iCounter = 1
    
    Do Until iCounter > Len(Expression) ' loop until we hit the expressions length
        
        StrCh1 = Mid(Expression, iCounter, 1)
        StrCh2 = LCase(Mid(Expression, iCounter, 3))
        StrCh3 = LCase(Mid(Expression, iCounter, 2))
        StrCh4 = LCase(Mid(Expression, iCounter, 4))
        
        If IsOperator(StrCh1) Then
            Operator = Mid(Expression, iCounter, 1)
            iCounter = iCounter + 1
        ElseIf IsOperator(StrCh2) Then
            Operator = Mid(Expression, iCounter, 3)
            iCounter = iCounter + 3
        ElseIf IsOperator(StrCh3) Then
            Operator = Mid(Expression, iCounter, 2)
            iCounter = iCounter + 2
        ElseIf IsOperator(StrCh4) Then
            Operator = Mid(Expression, iCounter, 4)
            iCounter = iCounter + 4
        End If
        
        Select Case LCase(Operator) ' check to see what Operator was found
            Case "" ' no choise was made so just return the data found
                Result = Trim(GetResultEx(Expression, iCounter))
            Case "+" 'preform addition
                Result = Result + GetResultEx(Expression, iCounter)
            Case "-" ' preform subtraction
                Result = Result - GetResultEx(Expression, iCounter)
            Case "/" ' preform a real division will include real numbers eg 5.2 3.666
                Result = Result / GetResultEx(Expression, iCounter)
            Case "\" ' preform Integer division returns real numbers eg 10 20 40
                Result = Result \ GetResultEx(Expression, iCounter)
            Case "^" ' preform Power
                Result = Result ^ GetResultEx(Expression, iCounter)
            Case "*" ' preform multiplication
                Result = Result * GetResultEx(Expression, iCounter)
            Case "<" ' less than
                sNum = Trim(GetResultEx(Expression, iCounter))
                Result = Result < sNum
            Case ">" ' more than
                sNum = Trim(GetResultEx(Expression, iCounter))
                Result = Result > sNum
            Case "=" ' equals
                Result = Result = Trim(GetResultEx(Expression, iCounter))
            Case "&" ' String concatenation
                Result = Result & GetResultEx(Expression, iCounter)
            Case "div" ' preform Integer division returns real numbers eg 10 20 40
                Result = Result \ GetResultEx(Expression, iCounter)
            Case "and" 'Logical bitwise and
                Result = Result And GetResultEx(Expression, iCounter)
            Case "mod" ' mod remainder
                Result = Result Mod GetResultEx(Expression, iCounter)
            Case "xor" ' exclusive or
                Result = Result Xor GetResultEx(Expression, iCounter)
                If Err Then Err.Clear
            Case "or"   ' Logical bitwise or
                Result = Result Or GetResultEx(Expression, iCounter)
            Case "not"
               Result = Not GetResultEx(Expression, iCounter)
            Case "like"
                Result = Result Like GetResultEx(Expression, iCounter)
            Case Else ' something else other than above was found
                Result = 0 ' just return zero becuase we don't know what it is
                Exit Do ' exit the loop so we don;t get stuck in here
            End Select
            DoEvents ' let other things process
        Loop
        


        Evaluate = Result ' Return the value
        
        ' Clean up uses vars
        iCounter = 0
        StrCh1 = ""
        StrCh2 = ""
        StrCh3 = ""
        StrCh4 = ""
        Operator = ""
        Result = 0
    
End Function

Private Function GetResultEx(Expression As String, iCounter) As Variant
Dim StrCh1 As String * 1, StrCh2 As String * 3, StrCh3 As String * 2, StrCh4 As String, sData As String

    Do Until iCounter > Len(Expression) ' loop until we hit the end of the expression
        StrCh1 = Mid(Expression, iCounter, 1)
        StrCh2 = LCase(Mid(Expression, iCounter, 3))
        StrCh3 = LCase(Mid(Expression, iCounter, 2))
        StrCh4 = LCase(Mid(Expression, iCounter, 4))
        If IsOperator(StrCh1) Then
            Exit Do
        ElseIf IsOperator(StrCh2) Then
            Exit Do
        ElseIf IsOperator(StrCh3) Then
            If Left(Expression, 1) = Chr(34) Then sData = Expression
            Exit Do
        ElseIf IsOperator(StrCh4) Then
            Exit Do
        Else
            sData = sData & StrCh1
            iCounter = iCounter + 1 ' update our counter
        End If
    Loop ' finish
    
    
    sData = ReturnData(Trim(sData))
    
    If IsNumeric(sData) Then
        GetResultEx = Val(sData) ' yes it is so pass it back
    Else
        GetResultEx = sData
    End If
        
    ' clean up
    StrCh1 = ""
    StrCh2 = ""
    StrCh3 = ""
    StrCh4 = ""
    sData = ""
End Function
