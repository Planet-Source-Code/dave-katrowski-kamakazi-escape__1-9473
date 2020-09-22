VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Choose Plane"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "I Want #1"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1800
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   3
      Left            =   1920
      Picture         =   "Form3.frx":0C42
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   2
      Left            =   1320
      Picture         =   "Form3.frx":1884
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   1
      Left            =   720
      Picture         =   "Form3.frx":24C6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   0
      Left            =   120
      Picture         =   "Form3.frx":3108
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "         has selected:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1640
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   2055
      Left            =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   2520
      X2              =   0
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Me.Hide: Form1.Show
Level = 1
SetXY 1, 20, 20
SetXY 2, 200, 200
Form1.Timer1.Enabled = True
Form1.Label5 = "120"
Form1.Timer3.Enabled = True
DS.PlaySound 1, True
DS.PlaySound 2, True
End Sub

Private Sub Form_Load()
I = Int((Rnd * 4) + 0.5) + 1
Comp.PlaneTYPE = I
Select Case I
Case 1
Picture2.Picture = Form1.Picture4.Picture
Case 2
Picture2.Picture = Form1.Picture6.Picture
Case 3
Picture2.Picture = Form1.Picture7.Picture
Case Else
Picture2.Picture = Form1.Picture8.Picture
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Show
End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Caption = "I Want #" & Index + 1
Pl.PlaneTYPE = Index + 1
End Sub
