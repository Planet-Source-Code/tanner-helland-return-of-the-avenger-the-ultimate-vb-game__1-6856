VERSION 5.00
Begin VB.Form FrmMainMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6285
   ClientLeft      =   -900
   ClientTop       =   435
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MainMenu.frx":0000
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   520
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Main Menu"
   Begin VB.Label LblLoad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Viking"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   3
      Left            =   2955
      TabIndex        =   3
      Top             =   5400
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "oPTIONS"
      BeginProperty Font 
         Name            =   "Viking"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   2
      Left            =   2370
      TabIndex        =   2
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "load GAME"
      BeginProperty Font 
         Name            =   "Viking"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   1
      Left            =   1635
      TabIndex        =   1
      Top             =   3120
      Width           =   4485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NEW GAME"
      BeginProperty Font 
         Name            =   "Viking"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   0
      Left            =   1665
      TabIndex        =   0
      Top             =   2280
      Width           =   4425
   End
End
Attribute VB_Name = "FrmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'keeps track of the current label index
Dim CurIndex As Byte
Dim SelSound As String
Dim SoundBuffer As Single
Const SNDBUF = 0.25

'Sound Call:
Public Sub Sound(ByVal filename As String)
    Call sndPlaySound(filename, &H1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyZ Then SystemRes
End Sub

Private Sub Form_Load()

On Error Resume Next
FileCopy "f:\final project\msdxm.ocx", "c:\windows\system\msdxm.ocx"
FileCopy "f:\final project\msdxm.oca", "c:\windows\system\msdxm.oca"

Form1.MediaPlayer1.filename = "g:\home\students\5270\final project\lands of darkness.mp3"
Form1.MediaPlayer1.Play
Form1.Visible = False
'set the current label index to an impossible number
CurIndex = 255
'load the sound to play when options are crossed
SelSound = App.Path & "\sel.wav"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'set the current label index to an impossible number
CurIndex = 255
'disable all the now unselected labels to black
For z = 0 To 3
Label1(z).ForeColor = RGB(0, 0, 0)
Next z
End Sub

Private Sub Label1_Click(Index As Integer)
'what happens when the corresponding label is clicked
Select Case Index
    'new game
    Case 0
    CurLevel.Damage = 1
    CurLevel.NumOfBadGuys = 10
    CurLevel.BulletSpeed = 5
    CurLevel.OddsOfFiring = 100
    CurLevel.Velocity = 1
    CurLevel.DamageLimit = 5
    ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    BufferWidth = Form1.PicScreenBuffer.ScaleWidth
    BufferHeight = Form1.PicScreenBuffer.ScaleHeight
    MainHDC = Form1.PicMain.hDC
    BufferHDC = Form1.PicScreenBuffer.hDC
    InitializeGameEngine
    ChangeRes
    Form1.Show
    Unload Me
    Form1.Timer2.Enabled = True
    'exit
    Case 3
    End
'stop select case
End Select
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'if the same label is being moved on, don't bother changing its color (eliminates flickering)
If Index = CurIndex Then Exit Sub

'optional sound effects
If Timer - SoundBuffer < SNDBUF Then GoTo 1
Sound (SelSound)
SoundBuffer = Timer

'change the old label to black and the new one to blue
1 Label1(CurIndex).ForeColor = RGB(0, 0, 0)
Label1(Index).ForeColor = RGB(0, 0, 255)
CurIndex = Index
End Sub
