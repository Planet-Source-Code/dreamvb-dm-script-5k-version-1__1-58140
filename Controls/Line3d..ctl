VERSION 5.00
Begin VB.UserControl Line3D 
   AutoRedraw      =   -1  'True
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1905
   ScaleHeight     =   90
   ScaleWidth      =   1905
End
Attribute VB_Name = "Line3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub DrawLine()
    UserControl.Line (0, 15)-(ScaleWidth, 15), vbWhite
    UserControl.Line (0, 0)-(ScaleWidth, 0), vbButtonShadow
End Sub
Private Sub UserControl_Initialize(): DrawLine: End Sub

Private Sub UserControl_Resize()
 On Error Resume Next
    UserControl.Height = 30: DrawLine: If Err Then Err.Clear
End Sub

