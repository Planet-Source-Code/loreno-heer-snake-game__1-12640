VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "About Snake"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   382
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Systeminfo..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   3075
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Innen ausgefüllt
      Index           =   1
      X1              =   6
      X2              =   376.933
      Y1              =   163
      Y2              =   163
   End
   Begin VB.Label lblDescription 
      Caption         =   "mailto:fatfish@bluewin.ch"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   3
      Top             =   1125
      Width           =   3915
   End
   Begin VB.Label lblTitle 
      Caption         =   "This Snakegame is programmed by Loreno Heer and published as Open-Source Freeware."
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   7
      X2              =   376.933
      Y1              =   164
      Y2              =   164
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":0884
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   4
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Registrierungsschlüssel-Sicherheitsoptionen...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Registrierungsschlüssel-Stammtypen...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Null-terminierte Unicode-Zeichenfolge
Const REG_DWORD = 4                      ' 32-Bit-Zahl

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title & " - By Loreno Heer"
    lblDescription.Caption = _
    "mailto:fatfish@bluewin.ch" & vbCrLf & _
    "" & vbCrLf & _
    "I've programed this game for Fabiola ;)"
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Versuchen, den Systeminfo-Programmpfad/-namen aus der Registrierung abzurufen...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Versuchen, nur den Systeminfo-Programmpfad aus der Registrierung abzurufen...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Überprüfen, ob bekannte 32-Dateiversion vorhanden ist
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Fehler - Datei wurde nicht gefunden...
        Else
            GoTo SysInfoErr
        End If
    ' Fehler - Registrierungseintrag wurde nicht gefunden...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "Systeminformationen sind momentan nicht verfügbar", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Schleifenzähler
    Dim rc As Long                                          ' Rückgabe-Code
    Dim hKey As Long                                        ' Zugriffsnummer für einen offenen Registrierungsschlüssel
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Datentyp eines Registrierungsschlüssels
    Dim tmpVal As String                                    ' Temporärer Speicher eines Registrierungsschlüsselwertes
    Dim KeyValSize As Long                                  ' Größe der Registrierungsschlüsselvariablen
    '------------------------------------------------------------
    ' Registrierungsschlüssel unter KeyRoot {HKEY_LOCAL_MACHINE...} öffnen
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Registrierungsschlüssel öffnen
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln...
    
    tmpVal = String$(1024, 0)                             ' Platz für Variable reservieren
    KeyValSize = 1024                                       ' Größe der Variable markieren
    
    '------------------------------------------------------------
    ' Registrierungsschlüsselwert abrufen...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Schlüsselwert abrufen/erstellen
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 fügt null-terminierte Zeichenfolge hinzu...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null gefunden, aus Zeichenfolge extrahieren
    Else                                                    ' Keine null-terminierte Zeichenfolge für WinNT...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null nicht gefunden, nur Zeichenfolge extrahieren
    End If
    '------------------------------------------------------------
    ' Schlüsselwerttyp für Konvertierung bestimmen...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Datentypen durchsuchen...
    Case REG_SZ                                             ' Zeichenfolge für Registrierungsschlüsseldatentyp
        KeyVal = tmpVal                                     ' Zeichenfolgenwert kopieren
    Case REG_DWORD                                          ' Registrierungsschlüsseldatentyp DWORD
        For i = Len(tmpVal) To 1 Step -1                    ' Jedes Bit konvertieren
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Wert Zeichen für Zeichen erstellen
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' DWORD in Zeichenfolge konvertieren
    End Select
    
    GetKeyValue = True                                      ' Erfolgreiche Ausführung zurückgeben
    rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
    Exit Function                                           ' Beenden
    
GetKeyError:      ' Bereinigen, nachdem ein Fehler aufgetreten ist...
    KeyVal = ""                                             ' Rückgabewert auf leere Zeichenfolge setzen
    GetKeyValue = False                                     ' Fehlgeschlagene Ausführung zurückgeben
    rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
End Function
