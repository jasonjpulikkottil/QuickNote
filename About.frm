VERSION 5.00
Begin VB.Form About 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   800
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label CloseBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1060
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2415
      Left            =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JStar© Inc."
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   930
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quicknote"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Jason J Pulikkottil"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label mover 
      BackStyle       =   0  'Transparent
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function MoveEffect(Ob1 As Label, Ob2 As Label)
Dim Xa, Xb, Ya, Yb
With Ob1
Xa = .Left + 10
Xb = .Left + .Width - 10
Ya = .Top + 10
Yb = .Top + .Height - 10


If (x > Xa Or x < Xb) And (y > Ya Or y < Yb) Then

.ForeColor = &HFF00&
Ob2.Caption = "yes"

End If

If (x < Xa Or x > Xb) And (y < Ya Or y > Yb) Then

.ForeColor = &H0&
Ob2.Caption = " no"

End If

End With

    
End Function


Private Sub Form_Load()
Label4.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & "  Beta"
End Sub

Private Sub Mover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
        Me.Left = Me.Left + x
        Me.Top = Me.Top + y
    End If
End Sub



Private Sub CloseBTN_Click()
Main.Show
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


