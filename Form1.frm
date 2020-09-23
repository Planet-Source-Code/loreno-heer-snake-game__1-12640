VERSION 5.00
Begin VB.Form frmPlayfield 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Snake - By Loreno Heer [v0.01]"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   668
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "frmPlayfield"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Direction of the Snake
Dim sDirection As Integer
Const dUp = 0
Const dLeft = 1
Const dDown = 2
Const dRight = 3
'Other information
Dim sLenght As Integer
Dim sPosX As Integer
Dim sPosY As Integer
Dim sSize As Integer
Dim sAllX(512) As Integer
Dim sAllY(512) As Integer
Dim sCurrent As Integer
'Variables
Dim n As Integer
Dim rndX As Integer
Dim rndY As Integer
Dim sKey As Boolean
Dim sB As Integer
'Speed
Dim sSpeed As Integer
Const sp1 = 400
Const sp2 = 300
Const sp3 = 250
Const sp4 = 200
Const sp5 = 150
Const sp6 = 125
Const sp7 = 100
Const sp8 = 75
Const sp9 = 50
'Points
Dim sPPFood As Integer
Dim sPoints As Integer
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If sKey = False Then
If KeyCode = vbKeyUp And sDirection <> dDown And sDirection <> dUp Then
    sDirection = dUp
    sKey = True
ElseIf KeyCode = vbKeyLeft And sDirection <> dRight And sDirection <> dLeft Then
    sDirection = dLeft
    sKey = True
ElseIf KeyCode = vbKeyDown And sDirection <> dUp And sDirection <> dDown Then
    sDirection = dDown
    sKey = True
ElseIf KeyCode = vbKeyRight And sDirection <> dLeft And sDirection <> dRight Then
    sDirection = dRight
    sKey = True
End If
End If
End Sub

Private Sub Form_Load()
Randomize Timer
Me.Caption = "Snake - By Loreno Heer [v" & App.Major & "." & App.Minor & App.Revision & "]"
Label1.Visible = frmStart.ShowPoints.Value
sSize = frmStart.TSize.Text
sSpeed = frmStart.CSpeed.ItemData(frmStart.CSpeed.ListIndex)
sPPFood = frmStart.TPPFood.Text
sDirection = dRight
sPosX = 2 * sSize
sPosY = 2 * sSize
sLenght = 5
sCurrent = 0
sPoints = 0
sB = 0
Label1.Caption = 0
Label1.Left = sSize * 19
If sSize <= 7 Then
    Label1.Top = -5
ElseIf sSize >= 8 Then
    Label1.Top = -8
End If
Label1.FontSize = (sSize + 3)
Timer1.Interval = sSpeed
DrawPlayField
Food
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmStart
End Sub

Private Sub Timer1_Timer()
If sDirection = dUp Then
    sPosY = sPosY - sSize
    If Point(sPosX + 1, sPosY + 1) = vbBlack And Not (sPosX = sAllX(1) And sPosY = sAllY(1)) Then
        sB = sB + 1
        If sB = GetSetting("Snake", "Options", "HardCore", 2) Then
            MsgBox "Game Over", , "Snake"
            CheckScore sPoints
            Restart
        Else
            sPosY = sPosY + sSize
            Exit Sub
        End If
    End If
ElseIf sDirection = dLeft Then
    sPosX = sPosX - sSize
        If Point(sPosX + 1, sPosY + 1) = vbBlack And Not (sPosX = sAllX(1) And sPosY = sAllY(1)) Then
        sB = sB + 1
        If sB = GetSetting("Snake", "Options", "HardCore", 2) Then
            MsgBox "Game Over", , "Snake"
            CheckScore sPoints
            Restart
        Else
            sPosX = sPosX + sSize
            Exit Sub
        End If
    End If
ElseIf sDirection = dDown Then
    sPosY = sPosY + sSize
        If Point(sPosX + 1, sPosY + 1) = vbBlack And Not (sPosX = sAllX(1) And sPosY = sAllY(1)) Then
        sB = sB + 1
        If sB = GetSetting("Snake", "Options", "HardCore", 2) Then
            MsgBox "Game Over", , "Snake"
            CheckScore sPoints
            Restart
        Else
            sPosY = sPosY - sSize
            Exit Sub
        End If
    End If
ElseIf sDirection = dRight Then
    sPosX = sPosX + sSize
        If Point(sPosX + 1, sPosY + 1) = vbBlack And Not (sPosX = sAllX(1) And sPosY = sAllY(1)) Then
        sB = sB + 1
        If sB = GetSetting("Snake", "Options", "HardCore", 2) Then
            MsgBox "Game Over", , "Snake"
            CheckScore sPoints
            Restart
        Else
            sPosX = sPosX - sSize
            Exit Sub
        End If
    End If
End If
sCurrent = sCurrent + 1
If sCurrent > sLenght Then
    Line (sAllX(1), sAllY(1))-(sAllX(1) + sSize - 1, sAllY(1) + sSize - 1), Me.BackColor, BF
    For n = 1 To sLenght - 1
        sAllX(n) = sAllX(n + 1)
    Next
    For n = 1 To sLenght - 1
        sAllY(n) = sAllY(n + 1)
    Next
    sCurrent = sCurrent - 1
End If
sAllX(sCurrent) = sPosX
sAllY(sCurrent) = sPosY
If Point(sPosX + 1, sPosY + 1) = vbBlack Then
    'Do Nothing
ElseIf Point(sPosX + 1, sPosY + 1) = vbGreen Then
    sLenght = sLenght + 1
    sPoints = sPoints + sPPFood
    Label1.Caption = sPoints
    Food
    Line (sPosX, sPosY)-(sPosX + sSize - 1, sPosY + sSize - 1), , BF
Else
    Line (sPosX, sPosY)-(sPosX + sSize - 1, sPosY + sSize - 1), , BF
    sB = 0
End If
sKey = False
End Sub
Sub Restart()
Cls
Form_Load
End Sub
Function Food()
Randomize Timer
Do
rndX = (Int(Rnd * 20) + 1) * sSize
rndY = (Int(Rnd * 12) + 2) * sSize
Loop Until Point(rndX + 1, rndY + 1) <> vbBlack And Point(rndX + 1, rndY + 1) <> vbGreen
Line (rndX, rndY)-(rndX + sSize - 1, rndY + sSize - 1), vbGreen, BF
End Function
Sub DrawPlayField()
For n = 0 To 21
Line (n * sSize, 1 * sSize)-(n * sSize + (sSize - 1), 1 * sSize + (sSize - 1)), , BF
Next
For n = 0 To 21
Line (n * sSize, 13 * sSize)-(n * sSize + (sSize - 1), 13 * sSize + (sSize - 1)), , BF
Next
For n = 0 To 12
Line (0 * sSize, (n + 1) * sSize)-(0 * sSize + (sSize - 1), (n + 1) * sSize + (sSize - 1)), , BF
Next
For n = 0 To 12
Line (21 * sSize, (n + 1) * sSize)-(21 * sSize + (sSize - 1), (n + 1) * sSize + (sSize - 1)), , BF
Next
Do Until Me.ScaleWidth <= sSize * 22
Me.Width = Me.Width - 10
Loop
Do Until Me.ScaleHeight <= sSize * 14
Me.Height = Me.Height - 10
Loop
End Sub
Function CheckScore(Points As Integer)
For n = 1 To 4
    If Points > GetSetting("Snake", "Highscore", ("Score" & n), 0) Then
        frmInput.Show 1, Me
        SaveSetting "Snake", "Highscore", ("Score" & n), Points
        SaveSetting "Snake", "Highscore", ("Name" & n), frmInput.Text1.Text
        Unload frmInput
        Exit For
    End If
Next
End Function
