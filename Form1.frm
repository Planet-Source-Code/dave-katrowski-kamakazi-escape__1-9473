VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "KamaKazi - Escape"
   ClientHeight    =   8415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   195
      Left            =   3000
      TabIndex        =   209
      Top             =   6720
      Width           =   1455
   End
   Begin VB.PictureBox Picture19 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3900
      Left            =   3480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   208
      Top             =   7200
      Width           =   3900
   End
   Begin VB.PictureBox Picture17 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3900
      Left            =   1560
      Picture         =   "Form1.frx":10604
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   207
      Top             =   7440
      Width           =   3900
   End
   Begin VB.PictureBox Picture16 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3900
      Left            =   120
      Picture         =   "Form1.frx":20C08
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   206
      Top             =   7800
      Width           =   3900
   End
   Begin VB.PictureBox Picture15 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3900
      Left            =   1080
      Picture         =   "Form1.frx":3120C
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   205
      Top             =   7680
      Width           =   3900
   End
   Begin VB.PictureBox Picture14 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3900
      Left            =   600
      Picture         =   "Form1.frx":41810
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   204
      Top             =   7680
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.PictureBox Picture13 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3900
      Left            =   3840
      Picture         =   "Form1.frx":51E14
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   203
      Top             =   7080
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   1680
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2160
      Top             =   1320
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   35
      Left            =   4800
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   185
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   35
      Left            =   4800
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   149
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   35
      Left            =   4800
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   186
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   34
      Left            =   4680
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   184
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   33
      Left            =   4560
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   183
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   32
      Left            =   4440
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   182
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   31
      Left            =   4320
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   181
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   30
      Left            =   4200
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   180
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   29
      Left            =   4080
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   179
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   28
      Left            =   3960
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   178
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   27
      Left            =   3840
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   177
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   26
      Left            =   3720
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   176
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   25
      Left            =   3600
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   175
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   24
      Left            =   3480
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   174
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   23
      Left            =   3360
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   173
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   22
      Left            =   3240
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   172
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   21
      Left            =   3120
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   171
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   20
      Left            =   3000
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   170
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   19
      Left            =   2880
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   169
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   18
      Left            =   2760
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   168
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   17
      Left            =   2640
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   167
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   16
      Left            =   2520
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   166
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   15
      Left            =   2400
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   165
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   14
      Left            =   2280
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   164
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   13
      Left            =   2160
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   163
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   12
      Left            =   2040
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   162
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   11
      Left            =   1920
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   161
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   10
      Left            =   1800
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   160
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   9
      Left            =   1680
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   159
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   8
      Left            =   1560
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   158
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   7
      Left            =   1440
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   157
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   6
      Left            =   1320
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   156
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   5
      Left            =   1200
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   155
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   4
      Left            =   1080
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   154
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   3
      Left            =   960
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   153
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   2
      Left            =   840
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   152
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   1
      Left            =   720
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   151
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   0
      Left            =   600
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   150
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   34
      Left            =   4680
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   148
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   33
      Left            =   4560
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   147
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   32
      Left            =   4440
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   146
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   31
      Left            =   4320
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   145
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   30
      Left            =   4200
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   144
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   29
      Left            =   4080
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   143
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   28
      Left            =   3960
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   142
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   27
      Left            =   3840
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   141
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   26
      Left            =   3720
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   140
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   25
      Left            =   3600
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   139
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   24
      Left            =   3480
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   138
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   23
      Left            =   3360
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   137
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   22
      Left            =   3240
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   136
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   21
      Left            =   3120
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   135
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   20
      Left            =   3000
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   134
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   19
      Left            =   2880
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   133
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   18
      Left            =   2760
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   132
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   17
      Left            =   2640
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   131
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   16
      Left            =   2520
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   130
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   15
      Left            =   2400
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   129
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   14
      Left            =   2280
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   128
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   13
      Left            =   2160
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   127
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   12
      Left            =   2040
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   126
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   11
      Left            =   1920
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   125
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   10
      Left            =   1800
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   124
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   9
      Left            =   1680
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   123
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   8
      Left            =   1560
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   122
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   7
      Left            =   1440
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   121
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   6
      Left            =   1320
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   120
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   5
      Left            =   1200
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   119
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   4
      Left            =   1080
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   118
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   3
      Left            =   960
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   117
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   2
      Left            =   840
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   116
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   1
      Left            =   720
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   115
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   0
      Left            =   600
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   114
      Top             =   8520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   34
      Left            =   4680
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   113
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   33
      Left            =   4560
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   112
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   32
      Left            =   4440
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   111
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   31
      Left            =   4320
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   110
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   30
      Left            =   4200
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   109
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   29
      Left            =   4080
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   108
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   28
      Left            =   3960
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   107
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   27
      Left            =   3840
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   106
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   26
      Left            =   3720
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   105
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   25
      Left            =   3600
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   104
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   24
      Left            =   3480
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   103
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   23
      Left            =   3360
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   102
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   22
      Left            =   3240
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   101
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   21
      Left            =   3120
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   100
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   20
      Left            =   3000
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   99
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   19
      Left            =   2880
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   98
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   18
      Left            =   2760
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   97
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   17
      Left            =   2640
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   96
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   16
      Left            =   2520
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   95
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   15
      Left            =   2400
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   94
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   14
      Left            =   2280
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   93
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   13
      Left            =   2160
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   92
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   12
      Left            =   2040
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   91
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   11
      Left            =   1920
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   90
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   10
      Left            =   1800
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   89
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   9
      Left            =   1680
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   88
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   8
      Left            =   1560
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   87
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   7
      Left            =   1440
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   86
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   6
      Left            =   1320
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   85
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   5
      Left            =   1200
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   84
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   4
      Left            =   1080
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   83
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   3
      Left            =   960
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   82
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   2
      Left            =   840
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   81
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   1
      Left            =   720
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   80
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   0
      Left            =   600
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   79
      Top             =   8400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture9 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3900
      Left            =   120
      Picture         =   "Form1.frx":62418
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   78
      Top             =   7080
      Width           =   3900
   End
   Begin VB.PictureBox Picture8 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "Form1.frx":72A1C
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   77
      Top             =   8640
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture7 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "Form1.frx":7365E
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   76
      Top             =   8520
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "Form1.frx":742A0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   75
      Top             =   8400
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   35
      Left            =   4800
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   74
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   34
      Left            =   4680
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   73
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   33
      Left            =   4560
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   72
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   32
      Left            =   4440
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   71
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   31
      Left            =   4320
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   70
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   30
      Left            =   4200
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   69
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   29
      Left            =   4080
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   68
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   28
      Left            =   3960
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   67
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   27
      Left            =   3840
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   66
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   26
      Left            =   3720
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   65
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   25
      Left            =   3600
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   64
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   24
      Left            =   3480
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   63
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   23
      Left            =   3360
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   62
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   22
      Left            =   3240
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   61
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   21
      Left            =   3120
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   60
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   20
      Left            =   3000
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   59
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   19
      Left            =   2880
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   58
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   18
      Left            =   2760
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   57
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   17
      Left            =   2640
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   56
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   16
      Left            =   2520
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   55
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   15
      Left            =   2400
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   54
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   14
      Left            =   2280
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   53
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   13
      Left            =   2160
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   52
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   12
      Left            =   2040
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   51
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   11
      Left            =   1920
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   50
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   10
      Left            =   1800
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   49
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   9
      Left            =   1680
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   48
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   8
      Left            =   1560
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   47
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   7
      Left            =   1440
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   46
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   6
      Left            =   1320
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   45
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   5
      Left            =   1200
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   44
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   4
      Left            =   1080
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   43
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   3
      Left            =   960
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   42
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   2
      Left            =   840
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   41
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   1
      Left            =   720
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   40
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Index           =   0
      Left            =   600
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   39
      Top             =   8280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "Form1.frx":74EE2
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   38
      Top             =   8280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "Form1.frx":75B24
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   37
      Top             =   8160
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   35
      Left            =   4800
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   36
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   34
      Left            =   4680
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   35
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   33
      Left            =   4560
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   34
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   32
      Left            =   4440
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   33
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   31
      Left            =   4320
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   32
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   30
      Left            =   4200
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   31
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   29
      Left            =   4080
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   30
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   28
      Left            =   3960
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   29
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   27
      Left            =   3840
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   28
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   26
      Left            =   3720
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   27
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   25
      Left            =   3600
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   26
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   24
      Left            =   3480
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   25
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   23
      Left            =   3360
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   24
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   22
      Left            =   3240
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   23
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   21
      Left            =   3120
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   22
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   20
      Left            =   3000
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   21
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   19
      Left            =   2880
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   20
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   18
      Left            =   2760
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   19
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   17
      Left            =   2640
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   18
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   16
      Left            =   2520
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   17
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   15
      Left            =   2400
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   16
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   14
      Left            =   2280
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   15
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   13
      Left            =   2160
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   14
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   12
      Left            =   2040
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   13
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   11
      Left            =   1920
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   12
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   10
      Left            =   1800
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   11
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   9
      Left            =   1680
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   10
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   8
      Left            =   1560
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   9
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   7
      Left            =   1440
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   8
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   6
      Left            =   1320
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   5
      Left            =   1200
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   4
      Left            =   1080
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   3
      Left            =   960
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   2
      Left            =   840
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   3
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   1
      Left            =   720
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Index           =   0
      Left            =   600
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   8160
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1680
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5535
      Left            =   120
      ScaleHeight     =   365
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   477
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "seconds to go..."
      Height          =   255
      Left            =   3000
      TabIndex        =   202
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "120"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   201
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   6000
      TabIndex        =   200
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   255
      Index           =   4
      Left            =   6000
      TabIndex        =   199
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   198
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Distance"
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   197
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Heading"
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   196
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Air Speed"
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   195
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   194
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   193
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   192
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   191
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Heading"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   190
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Air Speed"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   189
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enemy"
      Height          =   255
      Left            =   4560
      TabIndex        =   188
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player1"
      Height          =   255
      Left            =   120
      TabIndex        =   187
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   6975
      Left            =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim an As Single, sX As Single, sY As Single, sA As Single, sV As Long, Dis1 As Double, Dis2 As Double, Dis3 As Double, Dis4 As Double

Private Sub Command1_Click()
Me.Hide: Form2.Show
Form2.Command1.Visible = False
End Sub

Private Sub Form_Load()
Math_BTT
f = 25
For I = 0 To 35
For ii = 0 To 35: DoEvents
Form2.Picture1.Line (I, 0)-(I, 10), RGB(0, 0, I * 15)
Form2.Label1.ForeColor = RGB(I * 15, 0, 0)
Next
If -(I And 1) Then
Form2.Label1.Visible = True
Else
Form2.Label1.Visible = False
End If
bmp_rotate Picture3, Picture2(I), I * 10 * (3.14 / 180)
bmp_rotate Picture4, Picture5(I), I * 10 * (3.14 / 180)
bmp_rotate Picture6, Picture10(I), I * 10 * (3.14 / 180)
bmp_rotate Picture7, Picture11(I), I * 10 * (3.14 / 180)
bmp_rotate Picture8, Picture12(I), I * 10 * (3.14 / 180)
Next

Form2.Command1.Visible = True
Form2.Command2.Visible = True
Form2.Label1.Visible = False
Form2.Picture1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Command1.Visible = False
Form2.Show
End Sub

Private Sub Timer1_Timer()
Keys(0) = GetKeyState(&H26)
Keys(1) = GetKeyState(&H28)
Keys(2) = GetKeyState(&H25)
Keys(3) = GetKeyState(&H27)

If Keys(0) < 0 Then Particle.Move 1, "Forward", 5, 2, 1
If Keys(1) < 0 Then Particle.Move 1, "Reverse", 5, 2, 1
If Keys(2) < 0 Then Particle.Move 1, "Left", 5, 2, 1
If Keys(3) < 0 Then Particle.Move 1, "Right", 5, 2, 1
Particle.Move 1, "", 5, 2, 1

sX = p(2).X: sY = p(2).Y: sA = p(2).a: sV = p(2).v: Dis1 = Sqr(((sX - p(1).X) ^ 2) + ((sY - p(1).Y) ^ 2))

Particle.Move 2, "Left", 4, 1, 1
Particle.Move 2, "Forward", 4, 1, 1
Dis2 = Sqr(((p(2).X - p(1).X) ^ 2) + ((p(2).Y - p(1).Y) ^ 2))
SetXY 2, sX, sY: p(2).v = sV: p(2).a = sA

Particle.Move 2, "Right", 4, 1, 1
Particle.Move 2, "Forward", 4, 1, 1
Dis3 = Sqr(((p(2).X - p(1).X) ^ 2) + ((p(2).Y - p(1).Y) ^ 2))
SetXY 2, sX, sY: p(2).v = sV: p(2).a = sA

Particle.Move 2, "Forward", 4, 1, 1
Dis4 = Sqr(((p(2).X - p(1).X) ^ 2) + ((p(2).Y - p(1).Y) ^ 2))


SetXY 2, sX, sY: p(2).v = sV: p(2).a = sA
If Dis2 < Dis3 And Dis2 < Dis4 And Not Dis1 < Dis4 Then
Particle.Move 2, "Left", 3 + Level / 2, 1, 1
Particle.Move 2, "Forward", 3 + Level / 2, 1, 1
If Dis2 < 200 And Dis2 > 100 And p(2).v > 100 And p(1).v < p(2).v Then p(2).v = p(2).v - 2
If Dis2 < 100 And p(2).v > 100 And p(1).v < p(2).v Then p(2).v = p(2).v - 4
ElseIf Dis3 < Dis2 And Dis3 < Dis4 And Not Dis1 < Dis4 Then
Particle.Move 2, "Right", 3 + Level / 2, 1, 1
Particle.Move 2, "Forward", 3 + Level / 2, 1, 1
If Dis3 < 200 And Dis3 > 100 And p(2).v > 100 And p(1).v < p(2).v Then p(2).v = p(2).v - 2
If Dis3 < 100 And p(2).v > 100 And p(1).v < p(2).v Then p(2).v = p(2).v - 4
ElseIf Dis4 < Dis2 And Dis4 < Dis3 And Not Dis1 < Dis4 Then
Particle.Move 2, "Forward", 3 + Level / 2, 1, 1
If Dis4 < 200 And Dis4 > 100 And p(2).v > 100 And p(1).v < p(2).v Then p(2).v = p(2).v - 2
If Dis4 < 100 And p(2).v > 100 And p(1).v < p(2).v Then p(2).v = p(2).v - 4
Else 'If Dis1 < Dis2 And Dis1 < Dis3 And Dis1 < Dis4 Then
If Dis2 < Dis3 Then
Particle.Move 2, "Left", 3 + Level / 2, 1, 1
Particle.Move 2, "Forward", 3 + Level / 2, 1, 1
If Dis1 < 200 And Dis1 > 100 And p(2).v > 100 And p(1).v < p(2).v Then p(2).v = p(2).v - 2
If Dis1 < 100 And p(2).v > 100 And p(1).v < p(2).v Then p(2).v = p(2).v - 4
Else
Particle.Move 2, "Right", 3 + Level / 2, 1, 1
Particle.Move 2, "Forward", 3 + Level / 2, 1, 1
If Dis1 < 200 And Dis1 > 100 And p(2).v > 100 And p(1).v < p(2).v Then p(2).v = p(2).v - 2
If Dis1 < 100 And p(2).v > 100 And p(1).v < p(2).v Then p(2).v = p(2).v - 4
End If
End If

For I = 1 To 2
If p(I).X < 10 Then p(I).X = 10
If p(I).X > 466 Then p(I).X = 466
If p(I).Y < 1 Then p(I).Y = 10
If p(I).Y > 354 Then p(I).Y = 354
Next

Picture1.Cls
Select Case Level
Case 1
TileBack Picture1, Picture9, 0, 0
Case 2
TileBack Picture1, Picture13, 0, 0
Case 3
TileBack Picture1, Picture14, 0, 0
Case 4
TileBack Picture1, Picture15, 0, 0
Case 5
TileBack Picture1, Picture16, 0, 0
Case 6
TileBack Picture1, Picture17, 0, 0
Case 7
TileBack Picture1, Picture19, 0, 0
End Select
Select Case Pl.PlaneTYPE
Case 1
DrawPIC 1, Picture5(p(1).a).hDC, Picture2(p(1).a).hDC, Picture1.hDC, Picture2(0).ScaleWidth, Picture2(0).ScaleHeight, -(Picture2(0).ScaleWidth / 2), -(Picture2(0).ScaleHeight / 2)
Case 2
DrawPIC 1, Picture10(p(1).a).hDC, Picture2(p(1).a).hDC, Picture1.hDC, Picture2(0).ScaleWidth, Picture2(0).ScaleHeight, -(Picture2(0).ScaleWidth / 2), -(Picture2(0).ScaleHeight / 2)
Case 3
DrawPIC 1, Picture11(p(1).a).hDC, Picture2(p(1).a).hDC, Picture1.hDC, Picture2(p(2).a).ScaleWidth, Picture2(p(1).a).ScaleHeight, -(Picture2(0).ScaleWidth / 2), -(Picture2(0).ScaleHeight / 2)
Case 4
DrawPIC 1, Picture12(p(1).a).hDC, Picture2(p(1).a).hDC, Picture1.hDC, Picture2(0).ScaleWidth, Picture2(0).ScaleHeight, -(Picture2(0).ScaleWidth / 2), -(Picture2(0).ScaleHeight / 2)
End Select
Select Case Comp.PlaneTYPE
Case 1
DrawPIC 2, Picture5(p(2).a).hDC, Picture2(p(2).a).hDC, Picture1.hDC, Picture2(p(2).a).ScaleWidth, Picture2(p(2).a).ScaleHeight, -(Picture2(0).ScaleWidth / 2), -(Picture2(0).ScaleHeight / 2)
Case 2
DrawPIC 2, Picture10(p(2).a).hDC, Picture2(p(2).a).hDC, Picture1.hDC, Picture2(p(2).a).ScaleWidth, Picture2(p(2).a).ScaleHeight, -(Picture2(0).ScaleWidth / 2), -(Picture2(0).ScaleHeight / 2)
Case 3
DrawPIC 2, Picture11(p(2).a).hDC, Picture2(p(2).a).hDC, Picture1.hDC, Picture2(p(2).a).ScaleWidth, Picture2(p(2).a).ScaleHeight, -(Picture2(0).ScaleWidth / 2), -(Picture2(0).ScaleHeight / 2)
Case Else
DrawPIC 2, Picture12(p(2).a).hDC, Picture2(p(2).a).hDC, Picture1.hDC, Picture2(p(2).a).ScaleWidth, Picture2(p(2).a).ScaleHeight, -(Picture2(0).ScaleWidth / 2), -(Picture2(0).ScaleHeight / 2)
End Select

Pl.score = Pl.score + 1

DS.SetFrequency 1, 11025 + (p(1).v * 100)
DS.SetFrequency 2, 11025 + (p(2).v * 150)
If Form1.Visible = False Then
DS.StopSound 1
DS.StopSound 2
Else
DS.PlaySound 1, True
DS.PlaySound 2, True
End If
End Sub
Sub bmp_rotate(pic1 As PictureBox, pic2 As PictureBox, ByVal theta As Single)

    'Rotate the image in a picture box.

    'pic1 is the picture box with the bitmap to rotate

    'pic2 is the picture box to receive the rotated bitmap

    'theta is the angle of rotation

    
    Dim c1x As Integer, c1y As Integer
    Dim c2x As Integer, c2y As Integer
    Dim a As Single
    Dim p1x As Integer, p1y As Integer
    Dim p2x As Integer, p2y As Integer
    Dim n As Integer, r As Integer
    Dim c0 As Long, c1 As Long, C2 As Long, c3 As Long
    
    c1x = pic1.ScaleWidth \ 2
    c1y = pic1.ScaleHeight \ 2
    c2x = pic2.ScaleWidth \ 2
    c2y = pic2.ScaleHeight \ 2
    If c2x < c2y Then n = c2y Else n = c2x
    n = n - 1
    pic1hDC = pic1.hDC
    pic2hDC = pic2.hDC


    For p2x = 0 To n / 2 ': DoEvents


        For p2y = 0 To n / 2
            If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x)
            r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
            p1x = r * Cos(a + theta!)
            p1y = r * Sin(a + theta!)
            c0& = pic1.Point(c1x + p1x, c1y + p1y)
            c1& = pic1.Point(c1x - p1x, c1y - p1y)
            C2& = pic1.Point(c1x + p1y, c1y - p1x)
            c3& = pic1.Point(c1x - p1y, c1y + p1x)
            If c0& <> -1 Then pic2.PSet (c2x + p2x, c2y + p2y), c0&
            If c1& <> -1 Then pic2.PSet (c2x - p2x, c2y - p2y), c1&
            If C2& <> -1 Then pic2.PSet (c2x + p2y, c2y - p2x), C2&
            If c3& <> -1 Then pic2.PSet (c2x - p2y, c2y + p2x), c3&
        Next

        
        
    Next

End Sub

Private Sub Timer2_Timer()
Label4(0) = Int((p(1).v / 5) + 0.5)
Label4(1) = Int((p(1).a * 10) + 0.5)
Label4(2) = Pl.score
Label4(3) = Int((p(2).v / 5) + 0.5)
Label4(4) = Int((p(2).a * 10) + 0.5)
Label4(5) = Int((Dis1) + 0.5)

End Sub

Private Sub Timer3_Timer()
Label5 = Label5 - 1
If Label5 = 0 Then
Level = Level + 1
If Level = 8 Then
Form4.Label1 = "You have won all levels!!"
Else
Form4.Label1 = "Are you ready for level " & Level & "?"
End If
Form4.Label3 = Pl.score
Form1.Hide: Form4.Show
Timer3.Enabled = False
End If
End Sub
