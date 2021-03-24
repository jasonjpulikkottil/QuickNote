VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Settings 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "QuickNote"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   FillStyle       =   2  'Horizontal Line
   BeginProperty Font 
      Name            =   "Segoe UI Semibold"
      Size            =   9
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox F2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox F3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox F4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox F1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox FontSizeTXT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Text            =   "9"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox FontNameTXT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Segoe UI"
      Top             =   2400
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CMN 
      Left            =   2880
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Mover 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   0
      TabIndex        =   16
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
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
      Left            =   600
      TabIndex        =   15
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   600
      TabIndex        =   14
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   600
      TabIndex        =   13
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label OKBTN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                         OK"
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   1320
      TabIndex        =   12
      Top             =   3930
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Shape ColorC 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1200
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label CH1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label CH2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   2760
      TabIndex        =   5
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label CloseBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
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
      Left            =   3500
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Font"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Skin Color"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CH1_Click()
  With CMN
        .DialogTitle = "Font"
        .CancelError = False
        .ShowColor
    ColorC.FillColor = .Color
        End With
End Sub

Private Sub CH1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

CH1.ForeColor = &HFFFF&
CH1.BorderStyle = 1
End Sub

Private Sub CH1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
CH1.ForeColor = &HFF00&
CH1.BorderStyle = 0
End Sub










Private Sub CH2_Click()
Dim FType As String

    With CMN
   
        .DialogTitle = "Font"
        .CancelError = False
        .ShowFont
    FontNameTXT.Text = .FontName
    FontSizeTXT.Text = .FontSize
    
        If .FontBold = True Then F1.Text = "B"
        If .FontItalic = True Then F2.Text = "I"
        If .FontUnderline = True Then F3.Text = "U"
        If .FontStrikethru = True Then F4.Text = "S"
        

End With
End Sub

Private Sub CH2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
CH2.BorderStyle = 1
CH2.ForeColor = &HFFFF&
End Sub

Private Sub CH2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
CH2.ForeColor = &HFF00&
CH2.BorderStyle = 0
End Sub










Private Sub Mover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
        Me.Left = Me.Left + x
        Me.Top = Me.Top + y
    End If
End Sub



Private Sub CloseBTN_Click()
Me.Hide
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

Private Sub OKBTN_Click()
Main.BackColor = Settings.ColorC.FillColor

Main.Text.Font.Name = Settings.FontNameTXT
Main.Text.Font.Size = Settings.FontSizeTXT

If Settings.F1.Text = "B" Then Main.Text.Font.Bold = True Else Main.Text.Font.Bold = False
If Settings.F2.Text = "I" Then Main.Text.Font.Italic = True Else Main.Text.Font.Italic = False
If Settings.F1.Text = "U" Then Main.Text.Font.Underline = True Else Main.Text.Font.Underline = False
If Settings.F1.Text = "S" Then Main.Text.Font.Strikethrough = True Else Main.Text.Font.Strikethrough = False


Me.Hide



End Sub
Private Sub OKBTN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Shape2.FillColor = &HFF0000
End Sub
Private Sub OKBTN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Shape2.FillColor = &H800000
End Sub
