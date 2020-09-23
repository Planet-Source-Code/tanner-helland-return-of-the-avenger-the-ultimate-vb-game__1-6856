Attribute VB_Name = "Miscellaneous_Module"
Global MainHDC As Long
Global BufferHDC As Long

'ALL API CALLS:
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'ALL CONSTANTS:
Global Const NumOfStars = 250
Global Const BulletSpeed = 20
Global Const HAcc As Single = 2
Global Const HDel As Single = 2
Global Const VAcc As Single = 2
Global Const VDel As Single = 2
Global Const BadBulletSpeed = 5
Global Const DamageLimit = 15
Global Const KeySpeed = 10
Global Const NumOfBullets = 50

'ALL TYPES:
Public Type Star
    X As Integer
    Y As Integer
    bright As Byte
    speed As Byte
End Type

Public Type Bullet
    X As Integer
    Y As Integer
    Velocity As Integer
    Activated As Byte
End Type

Public Type BadBullet
    X As Integer
    Y As Integer
    Velocity As Integer
    Activated As Integer
End Type

Public Type BadGuy
    X As Integer
    Y As Integer
    OldX As Integer
    OldY As Integer
    Exploding As Byte
    ExplodingFrame As Byte
    Activated As Byte
    DstX As Integer
    DstY As Integer
    Damage As Byte
    Velocity As Byte
    Bullets(0 To 25) As BadBullet
    BulletsActivated As Byte
    Firing As Byte
End Type

Public Type Level
    NumOfBadGuys As Byte
    Damage As Byte
    DamageLimit As Byte
    Velocity As Byte
    OddsOfFiring As Byte
    BulletSpeed As Byte
End Type

Public Type PointXY
    X As Integer
    Y As Integer
End Type

'ALL GLOBAL VARIABLES:
Global Health As Single
Global BufferWidth As Long
Global BufferHeight As Long

Public Sub UpdateHealth()
BitBlt Form1.PicHealth.hDC, 0, 0, Form1.PicHealth.ScaleWidth, ((100 - Health) / 100) * Form1.PicHealth.ScaleHeight, Form1.Picture3.hDC, 0, 0, vbSrcCopy
Form1.PicHealth.Refresh
End Sub

Public Sub DrawHealthBar()
'calculation variables for r,g,b gradiency
Dim VR, VG, VB As Single
'colors of the picture boxes
Dim Color1, Color2 As Long
'r,g,b variables for each picture box
Dim R, G, B, R2, G2, B2 As Integer
'calculation variable for extracting the rgb values
Dim temp As Long

Color1 = RGB(255, 255, 0)
Color2 = RGB(255, 0, 0)

'extract the r,g,b values from the first picture box
temp = (Color1 And 255)
R = temp And 255
temp = Int(Color1 / 256)
G = temp And 255
temp = Int(Color1 / 65536)
B = temp And 255
temp = (Color2 And 255)
R2 = temp And 255
temp = Int(Color2 / 256)
G2 = temp And 255
temp = Int(Color2 / 65536)
B2 = temp And 255

'create a calculation variable for determining the step between
'each level of the gradient; this also allows the user to create
'a perfect gradient regardless of the form size
VR = Abs(R - R2) / Form1.PicHealth.ScaleHeight
VG = Abs(G - G2) / Form1.PicHealth.ScaleHeight
VB = Abs(B - B2) / Form1.PicHealth.ScaleHeight
'if the second value is lower then the first value, make the step
'negative
If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < B Then VB = -VB
'run a loop through the form height, incrementing the gradient color
'according to the height of the line being drawn
For Y = 0 To Form1.PicHealth.ScaleHeight
R2 = R + VR * Y
G2 = G + VG * Y
B2 = B + VB * Y
'draw the line and continue through the loop
Form1.PicHealth.Line (0, Y)-(Form1.PicHealth.ScaleWidth, Y), RGB(R2, G2, B2)
Next Y

End Sub
