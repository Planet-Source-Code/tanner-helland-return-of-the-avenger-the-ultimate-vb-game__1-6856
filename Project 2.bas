Attribute VB_Name = "Game_Logic_Module"
'ALL VARIABLES:
Global Bad1 As BadGuy
Global ShipX, ShipY As Integer
Global OldShipX, OldShipY
Global KeyState As Byte
Global sLeft As Byte
Global sRight As Byte
Global sUp As Byte
Global sDown As Byte
Global sVelocity As Single
Global sVel As Single
Global Firing As Byte
Global BulletsActivated As Byte
Global CurLevel As Level
Global CurrentLevel As Byte
Global CurTime As Long
Global OriginalTime As Long
Global Exploding As Byte
Global ExplodingFrame As Byte
Global Explosions(0 To 5) As PointXY
Global BadGuys() As BadGuy
Global BadGuyNum As Byte
Global Bullets(0 To NumOfBullets) As Bullet
Global StarArray(0 To NumOfStars) As Star

Public Sub InitializeGameEngine()

    Form1.WindowState = 2
    Randomize
    
    For X = 0 To NumOfStars
        BuildNewStar1 (X)
    Next X

    OldShipX = 0
    OldShipY = 0
    ShipX = Form1.PicMain.ScaleWidth / 2 - 32
    ShipY = Form1.PicMain.ScaleHeight - 64
    sVelocity = 10
    BulletsActivated = 0
    OriginalTime = Timer
    Health = 100
    CurrentLevel = 1
    BuildBadGuys
    DrawHealthBar
    Form1.Show
End Sub


Public Sub DrawStars()

Form1.PicScreenBuffer.Cls
'Draw the stars to their buffer
For X = 0 To NumOfStars
    StarArray(X).Y = StarArray(X).Y + StarArray(X).speed
    If StarArray(X).Y > Form1.PicMain.ScaleHeight Then BuildNewStar (X)
    SetPixelV Form1.PicScreenBuffer.hDC, StarArray(X).X, StarArray(X).Y, RGB(StarArray(X).bright, StarArray(X).bright, StarArray(X).bright)
Next X

End Sub

Public Sub FireBullets()
'Bullets
For X = 0 To NumOfBullets

If Firing = 1 And BulletsActivated <= 1 And Bullets(X).Activated = 0 Then
    Bullets(X).Activated = 1
    BulletsActivated = BulletsActivated + 1
    Bullets(X).X = ShipX + 30
    Bullets(X).Y = ShipY
End If

If Bullets(X).Activated = 1 Then
'BitBlt PicMain.hDC, Bullets(X).X, Bullets(X).Y, 4, 4, Picture3.hDC, 0, 0, vbSrcCopy
    Bullets(X).Y = Bullets(X).Y - BulletSpeed
    If Bullets(X).Y < -7 Then Bullets(X).Activated = 0
    BitBlt Form1.PicScreenBuffer.hDC, Bullets(X).X, Bullets(X).Y, 4, 4, Form1.PicBulletM.hDC, 0, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hDC, Bullets(X).X, Bullets(X).Y, 4, 4, Form1.PicBullet.hDC, 0, 0, vbSrcAnd

    For Y = 0 To CurLevel.NumOfBadGuys
        If BadGuys(Y).Activated = 1 And BadGuys(Y).Exploding = 0 And Bullets(X).Activated = 1 Then
            If Abs((BadGuys(Y).X + 32) - (Bullets(X).X + 2)) < 32 And Abs((BadGuys(Y).Y + 32) - (Bullets(X).Y + 2)) < 32 Then
                BadGuys(Y).Damage = BadGuys(Y).Damage + 1
                Bullets(X).Activated = 0
            End If
        End If
    Next Y
    
End If
11 Next X

'Don't allow any more bullets to be created
BulletsActivated = 0

End Sub

Public Sub ScrollShip()
    'Scrolling across edges of screen
    If ShipX > PicMain.ScaleWidth Then ShipX = 0
    If ShipX < -64 Then ShipX = PicMain.ScaleWidth
    If ShipY > PicMain.ScaleHeight Then ShipY = 0
    If ShipY < -64 Then ShipY = PicMain.ScaleHeight
End Sub

Public Sub VelocityCode()
'Movement Up
'If sUp = 1 Then
'sVel = sVel - VAcc
'ShipY = ShipY + sVel
'End If

'Movement Down
'If sDown = 1 Then
'sVel = sVel + VAcc
'ShipY = ShipY + sVel
'End If

'Vertical Deceleration
'If sUp = 0 And sDown = 0 And sVel <> 0 Then
'If sVel > 0 Then
'sVel = sVel - VDel
'If sVel <= 0 Then sVel = 0
'Else
'sVel = sVel + VDel
'If sVel >= 0 Then sVel = 0
'End If
'ShipY = ShipY + sVel
'End If

'Movement Left
If sLeft = 1 Then
    sVelocity = sVelocity - HAcc
    ShipX = ShipX + sVelocity
End If

'Movement Right
If sRight = 1 Then
    sVelocity = sVelocity + HAcc
    ShipX = ShipX + sVelocity
End If

If ShipX >= Form1.PicMain.ScaleWidth - 64 Then
    ShipX = Form1.PicMain.ScaleWidth - 64
    sVelocity = 0
End If

If ShipX <= 0 Then
    ShipX = 0
    sVelocity = 0
End If


'Horizontal Deceleration
If sRight = 0 And sLeft = 0 And sVelocity <> 0 Then
    If sVelocity > 0 Then
        sVelocity = sVelocity - HDel
        If sVelocity <= 0 Then sVelocity = 0
    Else
        sVelocity = sVelocity + HDel
        If sVelocity >= 0 Then sVelocity = 0
    End If
    ShipX = ShipX + sVelocity
End If

If sVelocity = 0 Then
    Form1.Picture1.Picture = Form1.PicT.Picture
    Form1.Picture2.Picture = Form1.PicTM.Picture
End If
End Sub

Public Sub BuildNewStar1(ByVal ArrayVal As Integer)
    StarArray(ArrayVal).X = Rnd * Form1.PicMain.ScaleWidth
    StarArray(ArrayVal).Y = Rnd * Form1.PicMain.ScaleHeight
    StarArray(ArrayVal).bright = Rnd * 255
    StarArray(ArrayVal).speed = Rnd * 10 + 2
End Sub

Public Sub BuildNewStar(ByVal ArrayVal As Integer)
    StarArray(ArrayVal).X = Rnd * Form1.PicMain.ScaleWidth
    StarArray(ArrayVal).Y = 0
    StarArray(ArrayVal).bright = Rnd * 200 + 55
    StarArray(ArrayVal).speed = Rnd * 10 + 1
End Sub

Public Sub BadGuyStuff()

For X = 0 To CurLevel.NumOfBadGuys
If BadGuys(X).Activated = 0 Then GoTo 10
    BadGuys(X).OldX = BadGuys(X).X
    BadGuys(X).OldY = BadGuys(X).Y
    If BadGuys(X).Activated = 1 And BadGuys(X).Exploding = 0 Then
    If BadGuys(X).DstX < BadGuys(X).X Then
        BadGuys(X).X = BadGuys(X).X - BadGuys(X).Velocity
    Else
        BadGuys(X).X = BadGuys(X).X + BadGuys(X).Velocity
    End If
    If BadGuys(X).DstY < BadGuys(X).Y Then
        BadGuys(X).Y = BadGuys(X).Y - BadGuys(X).Velocity
    Else
        BadGuys(X).Y = BadGuys(X).Y + BadGuys(X).Velocity
    End If
    If Abs(BadGuys(X).X - BadGuys(X).DstX) < CurLevel.Velocity + 1 Then BadGuys(X).DstX = Rnd * (Form1.PicMain.ScaleWidth + 72) - 72
    If Abs(BadGuys(X).Y - BadGuys(X).DstY) < CurLevel.Velocity + 1 Then BadGuys(X).DstY = Rnd * (Form1.PicMain.ScaleHeight - 200)
    BitBlt Form1.PicScreenBuffer.hDC, BadGuys(X).X, BadGuys(X).Y, 64, 64, Form1.PicBM.hDC, 0, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hDC, BadGuys(X).X, BadGuys(X).Y, 64, 64, Form1.PicB.hDC, 0, 0, vbSrcAnd
    End If
    If BadGuys(X).Damage > CurLevel.DamageLimit Then BadGuys(X).Exploding = 1
    If BadGuys(X).Exploding = 1 And BadGuys(X).Activated = 1 Then
    BitBlt Form1.PicScreenBuffer.hDC, BadGuys(X).X, BadGuys(X).Y, 75, 64, Form1.PicExplodeM.hDC, 77 * BadGuys(X).ExplodingFrame, 0, vbSrcPaint
    BitBlt Form1.PicScreenBuffer.hDC, BadGuys(X).X, BadGuys(X).Y, 75, 64, Form1.PicExplode.hDC, 77 * BadGuys(X).ExplodingFrame, 0, vbSrcInvert
    BadGuys(X).ExplodingFrame = BadGuys(X).ExplodingFrame + 1
    If BadGuys(X).ExplodingFrame = 13 Then BadGuys(X).Activated = 0
    End If

BadGuys(X).Firing = Int(Rnd * CurLevel.OddsOfFiring)

'Firing bullets
For Y = 0 To 25

If BadGuys(X).Firing = 1 And BadGuys(X).BulletsActivated <= 1 And BadGuys(X).Bullets(Y).Activated = 0 Then
BadGuys(X).Bullets(Y).Activated = 1
BadGuys(X).BulletsActivated = BadGuys(X).BulletsActivated + 1
BadGuys(X).Bullets(Y).X = BadGuys(X).X + 30
BadGuys(X).Bullets(Y).Y = BadGuys(X).Y
End If

If BadGuys(X).Bullets(Y).Activated = 1 Then
BadGuys(X).Bullets(Y).Y = BadGuys(X).Bullets(Y).Y + CurLevel.BulletSpeed
If BadGuys(X).Bullets(Y).Y > Form1.PicMain.ScaleHeight Then BadGuys(X).Bullets(Y).Activated = 0
BitBlt Form1.PicScreenBuffer.hDC, BadGuys(X).Bullets(Y).X, BadGuys(X).Bullets(Y).Y, 4, 4, Form1.PicBBulletM.hDC, 0, 0, vbMergePaint
BitBlt Form1.PicScreenBuffer.hDC, BadGuys(X).Bullets(Y).X, BadGuys(X).Bullets(Y).Y, 4, 4, Form1.PicBBullet.hDC, 0, 0, vbSrcAnd
If Abs((BadGuys(X).Bullets(Y).X - 32) - ShipX) < 25 And Abs((BadGuys(X).Bullets(Y).Y - 32) - ShipY) < 25 Then
Health = Health - BadGuys(Y).Damage
UpdateHealth
BadGuys(X).Bullets(Y).Activated = 0
End If

End If
11 Next Y
'Don't allow any more bullets to be created
BadGuys(X).BulletsActivated = 0
10 Next X

End Sub

Public Sub GoodGuyStuff()

    If Health <= 0 Then Exploding = 1
    If Exploding = 1 And ExplodingFrame = 0 Then
        For q = 0 To 5
            Explosions(q).X = ShipX + Int(Rnd * 64)
            Explosions(q).Y = ShipY + Int(Rnd * 64)
        Next q
        ExplodingFrame = 1
    End If
    If Exploding = 1 And ExplodingFrame <> 0 Then
        For q = 0 To 5
            BitBlt Form1.PicScreenBuffer.hDC, Explosions(q).X, Explosions(q).Y, 75, 64, Form1.PicExplodeM.hDC, 77 * ExplodingFrame, 0, vbSrcPaint
            BitBlt Form1.PicScreenBuffer.hDC, Explosions(q).X, Explosions(q).Y, 75, 64, Form1.PicExplode.hDC, 77 * ExplodingFrame, 0, vbSrcInvert
        Next q
        ExplodingFrame = ExplodingFrame + 1
    End If

    If ExplodingFrame >= 14 Then
        MsgBox "Game Over."
        RestoreRes
        End
    End If

    If Exploding = 0 Then
        BitBlt Form1.PicScreenBuffer.hDC, ShipX, ShipY, 64, 64, Form1.Picture2.hDC, 0, 0, vbMergePaint
        BitBlt Form1.PicScreenBuffer.hDC, ShipX, ShipY, 64, 64, Form1.Picture1.hDC, 0, 0, vbSrcAnd
    End If
    
    BitBlt Form1.PicMain.hDC, 0, 0, BufferWidth, BufferHeight, Form1.PicScreenBuffer.hDC, 0, 0, vbSrcCopy

    Dim TempCalc As Byte

    For X = 0 To CurLevel.NumOfBadGuys
        If BadGuys(X).Activated = 0 Then TempCalc = TempCalc + 1
    Next X

    If TempCalc = CurLevel.NumOfBadGuys + 1 Then
        MsgBox "You have beaten Level " & CurrentLevel
        For X = 0 To NumOfBullets
        Bullets(X).Activated = 0
        Next X
        CurrentLevel = CurrentLevel + 1
        BuildBadGuys
    End If
End Sub


Public Sub BuildBadGuys()

Firing = 0

'If CurrentLevel = 1 Then
'        CurLevel.Damage = 1
'        CurLevel.NumOfBadGuys = 10
'        CurLevel.OddsOfFiring = 40
'        CurLevel.Velocity = 1
'        CurLevel.DamageLimit = 5
'End If

'If CurrentLevel = 2 Then
'        CurLevel.Damage = 20
'        CurLevel.NumOfBadGuys = 200
'        CurLevel.OddsOfFiring = 2
'        CurLevel.Velocity = 2
'        CurLevel.DamageLimit = 0
'End If

CurLevel.Damage = CurLevel.Damage + 1
CurLevel.NumOfBadGuys = CurLevel.NumOfBadGuys + 1
CurLevel.BulletSpeed = CurLevel.BulletSpeed + 1
CurLevel.OddsOfFiring = CurLevel.OddsOfFiring * 0.8
CurLevel.Velocity = CurLevel.Velocity + 1
CurLevel.DamageLimit = CurLevel.DamageLimit + 1



ReDim BadGuys(0 To CurLevel.NumOfBadGuys) As BadGuy

For X = 0 To CurLevel.NumOfBadGuys
Randomize
     BadGuys(X).Activated = 1
     BadGuys(X).X = Rnd * Form1.PicMain.ScaleWidth
     BadGuys(X).Y = Rnd * (Form1.PicMain.ScaleHeight - 200)
     BadGuys(X).DstX = Rnd * Form1.PicMain.ScaleWidth
     BadGuys(X).DstY = Rnd * (Form1.PicMain.ScaleHeight - 200)
     BadGuys(X).Velocity = CurLevel.Velocity
     BadGuys(X).Damage = CurLevel.Damage
Next X


End Sub

