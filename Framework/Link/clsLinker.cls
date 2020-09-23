VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public ResponseData As String
Public Error_CallBack As String
Dim m_FileName As String

Public Function IsFileHere(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Private Function OpenFile(lzFile As String) As String
Dim ByteBuff() As Byte
    nFile = FreeFile
    Open lzFile For Binary As #nFile
        ReDim ByteBuff(1 To LOF(nFile))
        Get #nFile, , ByteBuff()
    Close #nFile
    
    OpenFile = StrConv(ByteBuff(), vbUnicode)
    Erase ByteBuff
    
End Function

Public Sub LinkAll()
Dim x As Integer
Dim vLink As Variant, StrLine As String, iPos As Integer, Parm As String
Dim lPos As Integer, TypeX As String, ObjFile As String
Dim ErrMsg1 As String, ErrMsg2 As String
Dim ExeFileToUse As String, OutputFile As String, dwFile As Long

    If Len(ResponseData) = 0 Then
        Error_CallBack = "You must supply some parameters."
        Exit Sub
    End If
    
    ErrMsg1 = "There was a problum while linking the file. " ' Error message
    ErrMsg2 = "The object file was not found:" & vbCrLf & "<OBJ>" & vbCrLf ' Error message
    
    Error_CallBack = ""
    vLink = Split(ResponseData, vbCrLf)
    
    If Not LCase(vLink(0)) = "#dm++" Then ' Must always be this and the first line
        Error_CallBack = ErrMsg1
        Exit Sub
    End If
    
    ' phase out the response data
    For x = LBound(vLink) To UBound(vLink)
        StrLine = LCase(Trim(vLink(x)))
        iPos = InStr(1, StrLine, ":", vbTextCompare)
        If iPos <> 0 Then Parm = UCase(Mid(StrLine, 1, iPos))
        
        If Len(StrLine) > 0 Then
            '//////////////////////////////////////////////////////////////////////////////
            Select Case Parm
                Case "/D:" ' used for the EXE Types
                    lPos = InStr(1, StrLine, " ", vbTextCompare)
                    If lPos = 0 Then Error_CallBack = ErrMsg1: Exit For
                    TypeX = UCase(Trim(Mid(StrLine, lPos + 1, Len(StrLine)))) ' will return %1 or %5
                Case "/OBJ:" ' Path and filename to the object file to be linked
                    ObjFile = Right(StrLine, Len(StrLine) - 5)
                    If Not Dir(ObjFile) <> "" Then
                        ErrMsg2 = Replace(ErrMsg2, "<OBJ>", ObjFile, , , vbTextCompare)
                        Error_CallBack = ErrMsg2
                        Exit For
                    End If
                    '//////////////////////////////////////////////////////////////////////////////
                    Select Case TypeX
                        Case "%1" ' Framework exe
                            ExeFileToUse = App.Path & "\exec\exe1.exe"
                        Case "%5" ' Standlone exe
                            ExeFileToUse = App.Path & "\exec\exe2.exe"
                        Case Else ' it ant in the list it an error
                            Error_CallBack = ErrMsg1 & " Unkown switch /d:exe" & TypeX
                            Exit For
                    End Select
                    '//////////////////////////////////////////////////////////////////////////////
                Case "/OUT:" ' Used for the output location of the compiled appliaction
                    If Not Dir(ExeFileToUse) <> "" Then
                        Error_CallBack = ErrMsg1 & "  File not found " & ExeFileToUse
                        Exit For
                    End If
                    
                    OutputFile = Right(StrLine, Len(StrLine) - 5)
                    FileCopy ExeFileToUse, OutputFile ' create a copy of the exe above
                    ' compile the object file into the new copyed exe
                    dwFile = FreeFile
                    Open OutputFile For Binary As #dwFile
                        Put #dwFile, LOF(1) + 1, "[DM]" & Chr(241) & OpenFile(ObjFile)
                    Close #dwFile
                    Parm = ""
                Case Else
                    StrLine = ""
                    Parm = ""
            End Select
            '//////////////////////////////////////////////////////////////////////////////
        End If
    Next
    
    x = 0: iPos = 0: lPos = 0
    Erase vLink
    StrLine = "": Parm = "": TypeX = "": ObjFile = "": ErrMsg1 = "": ErrMsg2 = "": ExeFileToUse = ""
    
End Sub

Public Property Get ResponseFile() As String
    ResponseFile = m_FileName
End Property

Public Property Let ResponseFile(ByVal vNewValue As String)
    If Not IsFileHere(vNewValue) Then
        Error_CallBack = "File not found: " & vbNewLine & vbNewLine & vNewValue
        Exit Property
    End If
    
    m_FileName = OpenFile(vNewValue)
End Property

