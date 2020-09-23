VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "host"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public sc_File As String
Public Framework_Path As String ' The path to the frameworks loaction
Private m_def_Error As Boolean
Public CodeBackBuff As String

Public Sub DmsRun()
Dim StrMain As String, MainFound As Boolean

    MainFound = ProcedureExists("main") ' Allways check for the main function
    
    If MainFound = False Then
        LastError 13, 0, "Required: Procedure Main();"
    End If
    
    StrMain = GetProcedureCode("main") ' Get the code from the main function
    ExecuteCode StrMain ' Execute the code
    
   ' If FoundError Then
     '   MsgBox ErrorStr, vbExclamation, App.Title
   ' End If
End Sub

Public Sub AddCode(ExpCode As String)
Dim ThisCode As String, StrLibs As String
    
    Reset
    ThisCode = "": StrLibs = "": CodeBackBuff = ""
    ThisCode = ExpCode ' Get the main code to run
    StrLibs = GetModules(ThisCode) ' Get and store all the include modules if any
    ThisCode = StrLibs & ThisCode  ' Add modules code to main code
    ThisCode = FormatStr(ThisCode) ' Add special formatting this is for the consts
    GetEnums ThisCode ' Get all enums
    ThisCode = RemoveComments(ThisCode) ' Remove any comments from the code
    GetProcedures ThisCode ' get all Procedures
    
    CodeBackBuff = ThisCode
End Sub

Private Function OpenFile(lzFile As String) As String
Dim ByteBuff() As Byte
    nFile = FreeFile
    scriptFullPath = lzFile
    Open lzFile For Binary As #nFile
        ReDim ByteBuff(LOF(1) - 1)
        Get #nFile, , ByteBuff()
    Close #nFile
    
    OpenFile = StrConv(ByteBuff(), vbUnicode)
    Erase ByteBuff
    
End Function

Public Property Let ScriptFile(ByVal vNewValue As String)
    If Not IsFileHere(vNewValue) Then
        FoundError = True
        ErrorStr = "File not found: " & vbNewLine & vbNewLine & vNewValue
        Exit Property
    End If
    
    sc_File = OpenFile(vNewValue)
End Property

Public Sub Execute()
    sc_File = Replace(sc_File, " _" & vbCrLf, " ", , , vbTextCompare)
    AddCode sc_File
    If FoundError Then Exit Sub
    DmsRun
End Sub

Public Property Get Version() As String
    Version = "v" & App.Major & "." & App.Minor
End Property

Public Property Get ErrorString() As String
   ErrorString = ErrorStr
End Property

Public Property Get CompileError() As Boolean
    CompileError = FoundError
End Property

Public Property Get Function_List() As String
Dim I As Integer, StrBuffer As String
    LoadFunctionList
            
    For I = LBound(FunctionMembers) To UBound(FunctionMembers)
        StrBuffer = StrBuffer & FunctionMembers(I) & ","
    Next
    
    I = 0
    Function_List = StrBuffer
    StrBuffer = ""
    
End Property

Private Sub Class_Initialize()
    Framework_Path = FixPath(App.Path)
End Sub

Public Property Let bScriptFile(ByVal vNewValue As String)
    scriptFullPath = vNewValue
End Property

Public Property Let bScriptCommand(ByVal vNewValue As String)
    StrCommandLine = vNewValue
End Property
