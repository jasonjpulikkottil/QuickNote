VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "QuickNote"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3960
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Interval        =   700
      Left            =   3480
      Top             =   1320
   End
   Begin RichTextLib.RichTextBox Text 
      Height          =   3615
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6376
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":259E8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   255
      Left            =   4080
      Shape           =   5  'Rounded Square
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   5050
      Width           =   375
   End
   Begin VB.Label ClrBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Tag             =   "0"
      Top             =   960
      Width           =   255
   End
   Begin VB.Label StrikeBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1115
      TabIndex        =   14
      Tag             =   "0"
      Top             =   960
      Width           =   255
   End
   Begin VB.Label UnderLineBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&U"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   780
      TabIndex        =   13
      Tag             =   "0"
      Top             =   960
      Width           =   255
   End
   Begin VB.Label BoldBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Tag             =   "0"
      Top             =   960
      Width           =   255
   End
   Begin VB.Label ItalicsBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   450
      TabIndex        =   11
      Tag             =   "0"
      Top             =   960
      Width           =   255
   End
   Begin VB.Label SettingsBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1425
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.Label OpenBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   75
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Label SaveBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   750
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Label ClearBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2340
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label MinBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3880
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.Label CloseBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4200
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.Label DateAndTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1150
      TabIndex        =   3
      ToolTipText     =   "Date and Time"
      Top             =   5040
      Width           =   2960
   End
   Begin VB.Label CharInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Character count"
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QuickNote"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   105
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   5415
      Left            =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Mover 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C0C0&
      X1              =   0
      X2              =   4560
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim CHX, CHY As Integer





Public Function SaveFile()

    With CommonDialog1
            .DialogTitle = "Save - QuickNote"
            .CancelError = False
            .Filter = "Text Files (*.txt)|*.txt|Rich Text Format (*.rtf)|*.rtf|All Files (*.*)|*.*"
           .ShowSave
           
            If Len(.FileName) = 0 Then
        Return
            End If
                sfile = .FileName
            End With
        
     extension = LCase(Right(sfile, 3))
      
    
        If extension = "rtf" Then
         Text.SaveFile sfile
      Else
       
       
      
       
       Open sfile For Output As #1
        Print #1, Text.Text
        Close #1
        
      End If
       
       
       
End Function

Public Function OpenFile()

       With CommonDialog1
        .DialogTitle = "Open - QuickNote"
        .CancelError = False
                .Filter = "Text Files (*.txt)|*.txt|Rich Text Format(*.rtf)|*.rtf|All Files(*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
           Return
        End If
        sfile = .FileName
    End With
 Text.LoadFile sfile
   
       
End Function











Private Sub BoldBTN_Click()

If Val(BoldBTN.Tag) = 0 Then

Text.SelBold = True
BoldBTN.BackColor = &HFF80FF
BoldBTN.Tag = 1

Else

Text.SelBold = False
BoldBTN.BackColor = RGB(255, 255, 255)
BoldBTN.Tag = 0

End If

End Sub









Private Sub ClrBTN_Click()
 With CommonDialog1
 
        .DialogTitle = "Font"
        .CancelError = False
        .ShowColor
         ClrBTN.ForeColor = .Color
         Text.SelColor = .Color
         
End With

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 19 Then SaveFile
If KeyAscii = 15 Then OpenFile
End Sub

Private Sub Form_Load()
Me.Hide
Splash.Show
Me.Label2.ForeColor = RGB(0, 0, 200)


Main.BackColor = Settings.ColorC.FillColor
Main.Text.Font.Name = Settings.FontNameTXT
Main.Text.Font.Size = Settings.FontSizeTXT

If Settings.F1.Text = "B" Then Main.Text.Font.Bold = True Else Main.Text.Font.Bold = False
If Settings.F2.Text = "I" Then Main.Text.Font.Italic = True Else Main.Text.Font.Italic = False
If Settings.F1.Text = "U" Then Main.Text.Font.Underline = True Else Main.Text.Font.Underline = False
If Settings.F1.Text = "S" Then Main.Text.Font.Strikethrough = True Else Main.Text.Font.Strikethrough = False


End Sub



Private Sub ItalicsBTN_Click()


If Val(ItalicsBTN.Tag) = 0 Then

Text.SelItalic = True
ItalicsBTN.BackColor = &HFF80FF
ItalicsBTN.Tag = 1

Else

Text.SelItalic = False
ItalicsBTN.BackColor = RGB(255, 255, 255)
ItalicsBTN.Tag = 0

End If


End Sub








Private Sub Label1_Click()
About.Show
End Sub


Private Sub Mover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 On Error Resume Next
 If Button = 1 Then
        Me.Left = Me.Left + x
        Me.Top = Me.Top + y
    End If
End Sub





Private Sub CloseBTN_Click()
Closewindow.Show
End Sub

Private Sub CloseBTN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
CloseBTN.BorderStyle = 1
CloseBTN.BackColor = &H80&
CloseBTN.ForeColor = RGB(200, 250, 255)
End Sub

Private Sub CloseBTN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
CloseBTN.BorderStyle = 0
CloseBTN.BackColor = RGB(255, 20, 20)
CloseBTN.ForeColor = RGB(0, 20, 250)
End Sub

Private Sub MinBTN_Click()
Me.WindowState = 1
End Sub
Private Sub MinBTN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MinBTN.BackColor = RGB(50, 50, 255)
MinBTN.ForeColor = &HFFFF&
End Sub

Private Sub MinBTN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MinBTN.BackColor = &HFFFF00
MinBTN.ForeColor = &HC00000
End Sub



Private Sub ClearBTN_Click()
Text.Text = ""
End Sub

Private Sub ClearBTN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ClearBTN.BackColor = RGB(0, 205, 0)
End Sub

Private Sub ClearBTN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
ClearBTN.BackColor = &HFF00&
End Sub





Private Sub SaveBTN_Click()
On Error Resume Next
     SaveFile
End Sub

Private Sub SaveBTN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
SaveBTN.BackColor = RGB(0, 205, 0)
End Sub

Private Sub SaveBTN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
SaveBTN.BackColor = &HFF00&
End Sub




Private Sub OpenBTN_Click()
On Error Resume Next
    OpenFile
 End Sub
Private Sub OpenBTN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
OpenBTN.BackColor = RGB(0, 205, 0)
End Sub

Private Sub OpenBTN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
OpenBTN.BackColor = &HFF00&
End Sub
















Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Shape1.BorderColor = RGB(255, 0, 0)
End Sub


Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Shape1.BorderColor = RGB(0, 25, 250)
End Sub



Private Sub SettingsBTN_Click()
Settings.Show
End Sub

Private Sub SettingsBTN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
SettingsBTN.BackColor = RGB(0, 205, 0)
End Sub

Private Sub SettingsBTN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
SettingsBTN.BackColor = &HFF00&
End Sub



Private Sub StrikeBTN_Click()




If Val(StrikeBTN.Tag) = 0 Then

Text.SelStrikeThru = True
StrikeBTN.BackColor = &HFF80FF
StrikeBTN.Tag = 1

Else

Text.SelStrikeThru = False
StrikeBTN.BackColor = RGB(255, 255, 255)
StrikeBTN.Tag = 0

End If



End Sub

Private Sub Text_Change()
CharInfo.Caption = Len(Text.Text)
End Sub


Private Sub Text_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 19 Then SaveFile
If KeyAscii = 15 Then OpenFile
End Sub


Private Sub Timer1_Timer()
DateAndTime.Caption = Date & "  " & Time
If Val(CharInfo.Caption) > 0 Then ClearBTN.Visible = True Else ClearBTN.Visible = False


End Sub

Private Sub Timer2_Timer()
If Label1.ForeColor = &H8000& Then
Label1.ForeColor = &H80&
Else
Label1.ForeColor = &H8000&
End If
End Sub

Private Sub UnderLineBTN_Click()


If Val(UnderLineBTN.Tag) = 0 Then

Text.SelUnderline = True
UnderLineBTN.BackColor = &HFF80FF
UnderLineBTN.Tag = 1

Else

Text.SelUnderline = False
UnderLineBTN.BackColor = RGB(255, 255, 255)
UnderLineBTN.Tag = 0

End If


End Sub
