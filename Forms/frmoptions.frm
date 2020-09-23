VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmoptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4095
      TabIndex        =   15
      Top             =   3645
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Cdialog 
      Left            =   1515
      Top             =   195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   540
      Top             =   3645
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   2610
      TabIndex        =   8
      Top             =   3645
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Editor"
      Height          =   3255
      Left            =   105
      TabIndex        =   0
      Top             =   255
      Width           =   5190
      Begin VB.TextBox txtMargin 
         Height          =   285
         Left            =   3870
         TabIndex        =   17
         Text            =   "5"
         Top             =   645
         Width           =   765
      End
      Begin VB.TextBox txtTabSize 
         Height          =   285
         Left            =   3900
         TabIndex        =   14
         Text            =   "4"
         Top             =   1395
         Width           =   765
      End
      Begin VB.PictureBox PicBack 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1290
         ScaleHeight     =   135
         ScaleWidth      =   810
         TabIndex        =   12
         Top             =   1560
         Width           =   870
      End
      Begin VB.PictureBox picFore 
         BackColor       =   &H00000000&
         Height          =   195
         Left            =   255
         ScaleHeight     =   135
         ScaleWidth      =   810
         TabIndex        =   10
         Top             =   1560
         Width           =   870
      End
      Begin VB.PictureBox PicPreView 
         BackColor       =   &H00FFFFFF&
         Height          =   810
         Left            =   150
         ScaleHeight     =   750
         ScaleWidth      =   2025
         TabIndex        =   5
         Top             =   2160
         Width           =   2085
         Begin VB.Label lblexample 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "IF @len($Str) <> 0 then"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   7
            Top             =   255
            Width           =   2760
         End
      End
      Begin VB.ComboBox cboSize 
         Height          =   315
         Left            =   2370
         TabIndex        =   4
         Top             =   645
         Width           =   990
      End
      Begin VB.ComboBox cbofonts 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   645
         Width           =   2040
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margin Size:"
         Height          =   195
         Left            =   3870
         TabIndex        =   16
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tab Size:"
         Height          =   195
         Left            =   3900
         TabIndex        =   13
         Top             =   1110
         Width           =   675
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   3615
         X2              =   3615
         Y1              =   405
         Y2              =   3090
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   3600
         X2              =   3600
         Y1              =   405
         Y2              =   3090
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Background:"
         Height          =   195
         Left            =   1290
         TabIndex        =   11
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Foreground:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preview"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   1935
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font Size:"
         Height          =   195
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblfont 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font:"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   360
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempA As String, TempB As String ' Holders for the combo boxes

Sub UpdatePicBoxCol(mPicBox As PictureBox)
On Error GoTo CanError:
    ' show colour dialog and update picture box with selected color
    CDialog.CancelError = True
    CDialog.ShowColor
    mPicBox.BackColor = CDialog.Color
    Exit Sub
    
CanError:
    If Err = cdlCancel Then Err.Clear: Exit Sub
    
End Sub

Function GetIndexFromComboBox(ItemToFind As String, nComboBoxName As ComboBox)
Dim nIndex As Integer
    ' returns the index of an item in a combo box
    For i = 0 To nComboBoxName.ListCount
        If LCase(nComboBoxName.List(i)) = LCase(ItemToFind) Then
            nIndex = i
            Exit For
        End If
    Next

    GetIndexFromComboBox = i
    
End Function

Private Sub cbofonts_Change()
    cbofonts.Text = TempA
End Sub

Private Sub cbofonts_Click()
    TempA = cbofonts.Text
    lblexample.FontName = cbofonts.Text
End Sub

Private Sub cboSize_Change()
    cboSize.Text = TempB
End Sub

Private Sub cboSize_Click()
    TempB = cboSize.Text
    lblexample.FontSize = Val(cboSize.Text)
    lblexample.Font.Weight = 400
End Sub

Private Sub cmdCancel_Click()
    Unload frmoptions 'unload this form
End Sub

Private Sub cmdUpdate_Click()
    ' Save the settings to the registery
    SaveSetting "DmScript", "Options", "FontName", cbofonts.Text
    SaveSetting "DmScript", "Options", "FontSize", cboSize.Text
    SaveSetting "DmScript", "Options", "ForeColor", "&H" & Hex(picFore.BackColor) & "&"
    SaveSetting "DmScript", "Options", "BackColor", "&H" & Hex(PicBack.BackColor) & "&"
    SaveSetting "DmScript", "Options", "Margin", txtMargin.Text
    SaveSetting "DmScript", "Options", "Indent", txtTabSize.Text
    
    GlobalEditorUpDate
    cmdCancel_Click ' Call cancel button and close this form
End Sub

Private Sub Form_Load()
    Dim i As Integer
    frmoptions.Icon = Nothing
    
    For i = 0 To Screen.FontCount - 1
        cbofonts.AddItem Screen.Fonts(i)
    Next
    
    cbofonts.ListIndex = GetIndexFromComboBox(AppConfig.EditorFontName, cbofonts)
    
    cboSize.AddItem "7"
    cboSize.AddItem "8"
    cboSize.AddItem "9"
    cboSize.AddItem "10"
    cboSize.AddItem "12"
    cboSize.AddItem "14"
    cboSize.AddItem "16"
    cboSize.AddItem "18"
    cboSize.AddItem "20"
    cboSize.AddItem "22"
    cboSize.AddItem "24"
    cboSize.AddItem "26"
    
    cboSize.ListIndex = GetIndexFromComboBox(CStr(AppConfig.EditorFontSize), cboSize)
    txtTabSize.Text = AppConfig.EditorTabSize
    txtMargin.Text = AppConfig.EditorMarginSize
    lblexample.Left = (AppConfig.EditorMarginSize * 13)
    
    picFore.BackColor = AppConfig.EditorForeColor
    PicBack.BackColor = AppConfig.EditorBackColor
    lblexample.ForeColor = AppConfig.EditorForeColor
    PicPreView.BackColor = AppConfig.EditorBackColor
    
End Sub

Private Sub PicBack_Click()
    UpdatePicBoxCol PicBack
    PicPreView.BackColor = PicBack.BackColor
End Sub

Private Sub picFore_Click()
    UpdatePicBoxCol picFore
    lblexample.ForeColor = picFore.BackColor
End Sub
Private Sub txtMargin_LostFocus()
    If Len(txtMargin.Text) = 0 Or Not IsNumeric(txtMargin.Text) Then txtMargin.Text = 5
    lblexample.Left = Val(txtMargin.Text * 13)
End Sub

Private Sub txtTabSize_LostFocus()
    If Len(txtTabSize.Text) = 0 Or Not IsNumeric(txtTabSize.Text) Then txtTabSize.Text = 4
End Sub
