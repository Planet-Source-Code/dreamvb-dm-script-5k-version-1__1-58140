Attribute VB_Name = "Modvariables"
Private Type TVar ' Used to hold all the variable data
    Count As Integer ' number of variables
    VariableName() As String ' variable name
    VariableType() As String ' variable type
    Variabledata() As String ' the data of an variable
End Type

Private Type TConsts ' used to hold consts data
    Count As Integer ' number of consts
    ConstName() As String ' consts name
    ConstValue() As String ' consts value
End Type

Private Type EnumItems ' enum holder
    Count As Integer ' number of enums in code
    EnumItemName() As String ' enum name
    EnumDataValue() As Variant ' enum value
End Type

Private Type TEnum
    Count As Integer ' number of items in enum list
    EnumNames() As String ' the name and of the enum item in list
    ItemData() As EnumItems ' the data of the enum item in the list
End Type

Private Type TArrayItems
    mArrayBoundData() As Variant ' used for array the data of the item in the array see TArray
End Type

Private Type TArray ' array holder
    Count As Integer ' number of arrays
    mArrayNames() As String ' the array name
    mArrayItemData() As TArrayItems ' link to arry items
End Type

Public dwArray As TArray ' see TArray and TArrayItems

Public dmEnum As TEnum ' see TEnum
Public Variables As TVar ' see TVar
Public dmConsts As TConsts ' see TConsts

Public Function GetArryUBound(dwIndex) As Long
On Error Resume Next
    ' return the upper bound of an array in other words UBound(arraName)
    GetArryUBound = UBound(dwArray.mArrayItemData(dwIndex).mArrayBoundData)
    If Err Then GetArryUBound = -1
End Function

Public Function EraseArray(dwIndex) As Long
On Error Resume Next
    ' Destroys an arrays contents
    Erase dwArray.mArrayItemData(dwIndex).mArrayBoundData()
    EraseArray = -1
    Exit Function
    If Err Then EraseArray = 0
End Function

Public Function ReturnArrayData(lzArrName As String, ArrItemIndex As Integer)
On Error Resume Next
Dim nIndx As Integer
    'returns the data from an array
    ReturnArrayData = ""
    nIndx = GetArryIndexByName(lzArrName)
    If nIndx = -1 Then ReturnArrayData = "": Exit Function
    ReturnArrayData = dwArray.mArrayItemData(nIndx).mArrayBoundData(ArrItemIndex)
    
    If Err Then LastError 13, 0, Err.Description
    
End Function

Public Function GetArryIndexByName(lzArrName As String) As Integer
' new code
Dim I As Integer
Dim nIdx As Integer
    ' returns the index of an array by using it's name
    nIdx = -1
    For I = 0 To dwArray.Count
        If LCase(dwArray.mArrayNames(I)) = LCase(lzArrName) Then
            nIdx = I
            Exit For
        End If
    Next
    
    I = 0
    GetArryIndexByName = nIdx
    nIdx = 0
    
End Function

Function isdmArray(lzArrayName As String) As Integer
Dim I As Integer, iCheck As Integer
    If dwArray.Count = 0 Then isdmArray = 0: Exit Function
    ' check if an array is in the array list if so it's an array
    For I = 0 To dwArray.Count
        If LCase(dwArray.mArrayNames(I)) = LCase(lzArrayName) Then
            iCheck = -1
            Exit For
        End If
    Next
    
    I = 0
    isdmArray = iCheck
    iCheck = False
    
End Function

Function EnumExists(dwEnumName As String) As Boolean
    ' check if an Enum Exists
    EnumExists = EnumIndex(dwEnumName) <> -1
End Function

Function EnumItemIndexByName(dwEnumIndex As Integer, EnumItemName As String) As Integer
Dim n As Integer
Dim nIndex As Integer
    ' returns the item index of a Enum
    EnumItemIndexByName = -1
    For n = 0 To dmEnum.ItemData(dwEnumIndex).Count
        If LCase(dmEnum.ItemData(dwEnumIndex).EnumItemName(n)) = LCase(EnumItemName) Then
            nIndex = n
            Exit For
        End If
    Next
    
    EnumItemIndexByName = nIndex
    nIndex = 0
    n = 0
    
End Function

Function EnumGetItemIndex(dwEnumName As String, EnumItemName As String) As String
Dim bIndex As Integer, ItemIndex As Integer
    
    EnumGetItemIndex = vbNullString
    bIndex = EnumIndex(dwEnumName)
    ' Get the Enums Index
    ItemIndex = EnumItemIndexByName(bIndex, EnumItemName)
    EnumGetItemIndex = dmEnum.ItemData(bIndex).EnumDataValue(ItemIndex)
    
End Function

Public Function EnumIndex(dwEnumName As String) As Integer
Dim n As Integer
On Error Resume Next

    If dmEnum.Count = 0 Then EnumIndex = -1: Exit Function
    For n = 0 To dmEnum.Count
        If LCase(dwEnumName) = LCase(dmEnum.EnumNames(n)) Then
            EnumIndex = n
            Exit For
        End If
    Next
    n = 0
End Function

Public Function GetConstValue(dwConstName As String)
    If dmConsts.Count = 0 Then Exit Function
    ' return the consts data
    GetConstValue = dmConsts.ConstValue(ConstIndex(dwConstName))
End Function

Public Function ConstIndex(dwConstName As String) As Integer
Dim n As Integer
On Error Resume Next
    ' returns the index of the consts
    If dmConsts.Count = 0 Then Exit Function
    For n = 0 To dmConsts.Count
        If LCase(dwConstName) = LCase(dmConsts.ConstName(n)) Then
            ConstIndex = n
            Exit For
        End If
    Next
    n = 0
End Function

Public Function ConstExists(dwConstName As String) As Boolean
Dim n As Integer
On Error Resume Next
    ' check if a Const Exists in the const list
    ConstExists = False
    If dmConsts.Count = 0 Then Exit Function
    For n = 0 To dmConsts.Count
        If LCase(dwConstName) = LCase(dmConsts.ConstName(n)) Then
            ConstExists = True
            Exit For
        End If
    Next
    n = 0
End Function

Public Function VariableExists(dwVarname As String) As Boolean
Dim n As Integer
    ' check to see if a variable Exists checks for both normal variables and arrays
    ' check if we are dealing with an array datatype or something else
    If Not HasBrackets(dwVarname) Then ' we dealing with a normal array
        VariableExists = False
        If Variables.Count = 0 Then Exit Function
        For n = 0 To Variables.Count
            If LCase(dwVarname) = LCase(Variables.VariableName(n)) Then
                VariableExists = True
                Exit For
            End If
        Next
        n = 0
        Exit Function
    Else ' hay we got an array here
        VariableExists = False
        If dwArray.Count = 0 Then Exit Function
        For n = 0 To dwArray.Count
            If LCase(dwArray.mArrayNames(n)) = LCase(GetArrayNameIndex(dwVarname, 1)) Then
                VariableExists = True
                Exit For
            End If
        Next
        n = 0
    End If
    
End Function

Public Function SetVariableData(dwVarname As String, dwVarData As String)
    If Variables.Count = 0 Then Exit Function
    ' sets the variable data
    Variables.Variabledata(VariableIndex(dwVarname)) = dwVarData
End Function

Public Function GetVariableData(dwVarname As String)
    If Variables.Count = 0 Then Exit Function
    ' returns a variables data based on it's name
    GetVariableData = Variables.Variabledata(VariableIndex(dwVarname))
End Function

Public Function VariableIndex(dwVarname As String) As Integer
Dim n As Integer
    ' returns the index of an variable
    If Variables.Count = 0 Then Exit Function
    For n = 0 To Variables.Count
        If LCase(dwVarname) = LCase(Variables.VariableName(n)) Then
            VariableIndex = n
            Exit For
        End If
    Next
    n = 0
End Function

