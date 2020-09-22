VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   1800
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   10
      Height          =   135
      Left            =   120
      ScaleHeight     =   10
      ScaleMode       =   0  'User
      ScaleWidth      =   35
      TabIndex        =   4
      Top             =   2040
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Game"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   2295
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "KamaKazi - Escape"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   30
      Width           =   4470
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Shake1 As Boolean, Shake2 As Boolean, Shake3 As Boolean
Private Sub Command1_Click()
Me.Hide: Form3.Show: Pl.PlaneTYPE = 1
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shake2 = True
Shake3 = False
Shake1 = False
End Sub

Private Sub Command2_Click()
Unload Form1
Unload Form3
Unload Me
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shake3 = True
Shake2 = False
Shake1 = False
End Sub

Private Sub Form_Load()
Me.Show: Me.Refresh
Load Form1: Form1.Hide
DS.Initialize_Engine Form1.Hwnd
DS.LoadWavToChannel 1, App.Path & "\peng.wav"
DS.LoadWavToChannel 2, App.Path & "\ceng.wav"
DS.LoadWavToChannel 3, App.Path & "\c.wav"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shake1 = False
Shake2 = False
Shake3 = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
DS.Terminate_Engine
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shake1 = True
Shake2 = False
Shake3 = False
End Sub

Private Sub Timer1_Timer()
If Shake1 Then
Label2.Top = RndRange(10, 50)
Label2.Left = RndRange(100, 140)
End If
If Shake2 Then
Command1.Top = RndRange(580, 620)
Command1.Left = RndRange(100, 140)
End If
If Shake3 Then
Command2.Top = RndRange(1180, 1220)
Command2.Left = RndRange(100, 140)
End If
End Sub

