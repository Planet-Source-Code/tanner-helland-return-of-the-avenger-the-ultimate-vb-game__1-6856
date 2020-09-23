VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicBBulletM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   120
      Left            =   480
      Picture         =   "Starfield.frx":0000
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   24
      Top             =   3240
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox PicBBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   120
      Left            =   240
      Picture         =   "Starfield.frx":0074
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   855
   End
   Begin VB.PictureBox PicExplode 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   120
      Picture         =   "Starfield.frx":00E8
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1067
      TabIndex        =   20
      Top             =   6240
      Visible         =   0   'False
      Width           =   16065
   End
   Begin VB.PictureBox PicExplodeM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   120
      Picture         =   "Starfield.frx":1102C
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1067
      TabIndex        =   21
      Top             =   5880
      Visible         =   0   'False
      Width           =   16065
   End
   Begin VB.PictureBox PicScreenBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6900
      Left            =   1080
      ScaleHeight     =   456
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   456
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   6900
   End
   Begin VB.PictureBox PicHealth 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6975
      Left            =   8400
      ScaleHeight     =   461
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   15
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00000000&
      Height          =   6900
      Left            =   1080
      ScaleHeight     =   456
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   456
      TabIndex        =   14
      Top             =   0
      Width           =   6900
   End
   Begin VB.PictureBox PicTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1020
      Left            =   240
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicBulletM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   120
      Left            =   360
      Picture         =   "Starfield.frx":21F70
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox PicBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   120
      Left            =   120
      Picture         =   "Starfield.frx":21FE4
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox PicBM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   4680
      Picture         =   "Starfield.frx":22058
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   3600
      Picture         =   "Starfield.frx":23CC7
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicTM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   8520
      Picture         =   "Starfield.frx":25E70
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   7920
      Picture         =   "Starfield.frx":27BB5
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicRM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   6840
      Picture         =   "Starfield.frx":29A98
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicR 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   5760
      Picture         =   "Starfield.frx":2B4C6
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicLM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   4680
      Picture         =   "Starfield.frx":2D0FB
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   3600
      Picture         =   "Starfield.frx":2EB8D
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   4080
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   12060
      Left            =   480
      Picture         =   "Starfield.frx":3079A
      ScaleHeight     =   800
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   2
      Top             =   8520
      Visible         =   0   'False
      Width           =   12060
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   1200
      Picture         =   "Starfield.frx":36609
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   120
      Picture         =   "Starfield.frx":3834E
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   1020
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   615
      Left            =   0
      TabIndex        =   25
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   3
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VELOCITY"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   960
   End
   Begin VB.Label LblVel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   9960
      TabIndex        =   16
      Top             =   120
      Width           =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'©2000 Tanner Helland

Private Sub Command1_Click()
    GameActive = 0
    RestoreRes
    End
End Sub

Private Sub Form_Load()
    'InitializeGameEngine
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
    sLeft = 1
    Picture1.Picture = PicL.Picture
    Picture2.Picture = PicLM.Picture
End If
If KeyCode = vbKeyRight Then
    sRight = 1
    Picture1.Picture = PicR.Picture
    Picture2.Picture = PicRM.Picture
End If
If KeyCode = vbKeyUp Then sUp = 1
If KeyCode = vbKeyDown Then sDown = 1

If KeyCode = vbKeySpace Then
    Firing = 1
End If

If KeyCode = vbKeySubtract Then
    Health = Health - 1
    UpdateHealth
End If

If KeyCode = vbKeyEscape Then
    GameActive = 0
    RestoreRes
    End
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then sLeft = 0
If KeyCode = vbKeyRight Then sRight = 0
'If KeyCode = vbKeyUp Then sUp = 0
'If KeyCode = vbKeyDown Then sDown = 0
If KeyCode = vbKeySpace Then Firing = 0
End Sub

Private Sub Timer2_Timer()

OldShipX = ShipX
OldShipY = ShipY

'Do all of the game engine stuff...
DrawStars
VelocityCode
BadGuyStuff
FireBullets
GoodGuyStuff

'Change the velocity caption
LblVel.Caption = 100 - sVel & " kps"

End Sub
