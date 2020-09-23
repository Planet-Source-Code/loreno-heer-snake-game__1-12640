VERSION 5.00
Begin VB.Form frmHighscore 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Highscores"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3465
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   231
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   232
      X2              =   0
      Y1              =   25
      Y2              =   25
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   232
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Label Score 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      Caption         =   "Score #1"
      Height          =   165
      Index           =   1
      Left            =   2520
      TabIndex        =   9
      Top             =   480
      Width           =   840
   End
   Begin VB.Label HName 
      AutoSize        =   -1  'True
      Caption         =   "Name #1"
      Height          =   165
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Score 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      Caption         =   "Score #2"
      Height          =   165
      Index           =   2
      Left            =   2520
      TabIndex        =   7
      Top             =   720
      Width           =   840
   End
   Begin VB.Label HName 
      AutoSize        =   -1  'True
      Caption         =   "Name #2"
      Height          =   165
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Score 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      Caption         =   "Score #3"
      Height          =   165
      Index           =   3
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   840
   End
   Begin VB.Label HName 
      AutoSize        =   -1  'True
      Caption         =   "Name #3"
      Height          =   165
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Score 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      Caption         =   "Score #4"
      Height          =   165
      Index           =   4
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label HName 
      AutoSize        =   -1  'True
      Caption         =   "Name #4"
      Height          =   165
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Score"
      Height          =   165
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   165
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmHighscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MsgBox("Are you Sure?", vbYesNo, "Delete Highscores") = vbYes Then
    SaveSetting "Snake", "Highscore", "Name1", "-"
    SaveSetting "Snake", "Highscore", "Name2", "-"
    SaveSetting "Snake", "Highscore", "Name3", "-"
    SaveSetting "Snake", "Highscore", "Name4", "-"
    SaveSetting "Snake", "Highscore", "Score1", 0
    SaveSetting "Snake", "Highscore", "Score2", 0
    SaveSetting "Snake", "Highscore", "Score3", 0
    SaveSetting "Snake", "Highscore", "Score4", 0
    For n = 1 To 4
        Score(n).Caption = ""
        HName(n).Caption = "-"
    Next
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
For n = 1 To 4
    HName(n).Caption = GetSetting("Snake", "Highscore", ("Name" & n), "")
    Score(n).Caption = GetSetting("Snake", "Highscore", ("Score" & n), 0)
Next
End Sub
