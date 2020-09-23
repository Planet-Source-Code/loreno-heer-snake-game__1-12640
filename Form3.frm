VERSION 5.00
Begin VB.Form frmInput 
   ClientHeight    =   825
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3630
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   825
   ScaleWidth      =   3630
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Text            =   "Player 1"
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You've made a new Highscore! Please enter your name"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3600
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub

