VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form4"
   ScaleHeight     =   1215
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lets Go"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Your total score is:"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Are you ready for level 1?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Me.Hide: Form1.Show
SetXY 1, 20, 20
SetXY 2, RndRange(200, 400), RndRange(200, 400)
For i = 1 To 2
p(i).a = 0
p(i).v = 0
Next
Form1.Label5 = "120"
Form1.Timer3.Enabled = True
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub
