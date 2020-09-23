VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Options"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4035
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   96
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   269
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Play"
      Default         =   -1  'True
      Height          =   255
      Left            =   2760
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Highscoore"
      Height          =   255
      Left            =   2760
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox CSpeed 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      ItemData        =   "Form2.frx":0000
      Left            =   1560
      List            =   "Form2.frx":002F
      Style           =   2  'Dropdown-Liste
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.CheckBox ShowPoints 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "Yes"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1155
      Value           =   1  'Aktiviert
      Width           =   615
   End
   Begin VB.TextBox TPPFood 
      Alignment       =   1  'Rechts
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   225
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "5"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox TSize 
      Alignment       =   1  'Rechts
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   1560
      TabIndex        =   4
      Text            =   "5"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "About"
      Height          =   255
      Left            =   2760
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Points:"
      Height          =   165
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PPFood:"
      Height          =   165
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      Height          =   165
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   630
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom: (Default=5)"
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1155
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmPlayfield.Show
Me.Hide
End Sub

Private Sub Command2_Click()
frmHighscore.Show 1, Me
End Sub

Private Sub Command3_Click()
frmAbout.Show 1, Me
End Sub

Private Sub CSpeed_Click()
TPPFood.Text = CSpeed.ListIndex + 1
End Sub

Private Sub Form_Load()
If GetSetting("Snake", "Options", "Played", False) = False Then
    SaveSetting "Snake", "Options", "HardCore", 2
    SaveSetting "Snake", "Options", "Played", True
    CSpeed.ListIndex = 0
    SaveSetting "Snake", "Options", "Speed", CSpeed.ListIndex
    TPPFood.Text = CSpeed.ListIndex + 1
    SaveSetting "Snake", "Options", "PPFood", TPPFood.Text
    TSize.Text = 5
    SaveSetting "Snake", "Options", "Size", TSize.Text
    'Highscore
    SaveSetting "Snake", "Highscore", "Name1", ""
    SaveSetting "Snake", "Highscore", "Name2", ""
    SaveSetting "Snake", "Highscore", "Name3", ""
    SaveSetting "Snake", "Highscore", "Name4", ""
    SaveSetting "Snake", "Highscore", "Score1", "0"
    SaveSetting "Snake", "Highscore", "Score2", "0"
    SaveSetting "Snake", "Highscore", "Score3", "0"
    SaveSetting "Snake", "Highscore", "Score4", "0"
Else
    CSpeed.ListIndex = GetSetting("Snake", "Options", "Speed", 0)
    TPPFood.Text = GetSetting("Snake", "Options", "PPFood", 1)
    TSize.Text = GetSetting("Snake", "Options", "Size", 5)
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
SaveSetting "Snake", "Options", "Size", TSize.Text
SaveSetting "Snake", "Options", "PPFood", TPPFood.Text
SaveSetting "Snake", "Options", "Speed", CSpeed.ListIndex
End Sub
