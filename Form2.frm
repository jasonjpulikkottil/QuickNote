VERSION 5.00
Begin VB.Form Closewindow 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "QuickNote"
   ClientHeight    =   1215
   ClientLeft      =   8010
   ClientTop       =   9000
   ClientWidth     =   2535
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Mover 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   675
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   675
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1320
      Top             =   600
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   240
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure to exit ?"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Closewindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label4_Click()
Me.Hide
End Sub

Private Sub Label5_Click()
End
End Sub


Private Sub Mover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
        Me.Left = Me.Left + x
        Me.Top = Me.Top + y
    End If
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Shape3.FillColor = RGB(0, 200, 0)
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Shape3.FillColor = &H80FF80
End Sub





Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Shape2.FillColor = RGB(200, 0, 0)
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Shape2.FillColor = &H8080FF
End Sub


