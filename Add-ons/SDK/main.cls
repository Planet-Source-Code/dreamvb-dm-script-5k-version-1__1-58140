VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' You can use this project as a base to build your plug-ins

' mObject.DevCodeWindow will give you access the code editor
' mObject.DevIDE will give you access to the IDE it's self

' or the plug-ins will not work

Public Sub RunPlugin(mObject As Object)
    ' note this sub is required to run the plugins
    
    Set Form1.Editor = mObject.Dev_IDE_Window.txtCode
    Set Form1.DevIDE = mObject.Dev_IDE_Window
    Set Form1.DevFrameWork = mObject.Dev_Frame_Work
    Form1.Show
    
End Sub
