Option Explicit
Dim girllights, gl, maturepictures, alternateflippers

'PLAYBOY (STERN 2002): BY RASCAL AND HIREZ00

'GAME OPTIONS BELOW
'SET OPTIONS TO TRUE OR FALSE
'-------------------------------------------------------------------------------------------------------------
girllights=True     '<--- Change True to False to DISABLE extra Centerfold playfield girl lights.
maturepictures=False  '<--- True = Full Nude/Mature images. False = Non-Nude images (default)
alternateflippers=False '<--- True = Alternative Flipper Decals, False = Factory Default Flipper Decals.
'-------------------------------------------------------------------------------------------------------------
'END GAME OPTIONS

' Thalamus 2020 April : Improved directional sounds
' !! NOTE : Table not verified yet !!
' Options
' Volume devided by - lower gets higher sound
Const VolDiv = 5000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolFlip   = 1    ' Flipper volume.
Const VolFlipD  = 1    ' Flipper Down volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolGates  = .5   ' Gates volume.
Const VolSling  = 1    ' Slingshot volume.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim UseVPMDMD
UseVPMDMD=True

LoadVPM "01560000", "SEGA.VBS", 3.26

changeflippers
Sub changeflippers
  If alternateflippers=True Then
    PLeftFlipper.Image="HiRez00_left-flipper-alt"
    PRightFlipper.Image="HiRez00_right-flipper-alt"
  Else
    PLeftFlipper.Image="lflipper AP"
    PRightFlipper.Image="rflipper AP"
  End If
End Sub

changegl
Sub changegl
  If girllights=True then
    For Each gl in monthgirls
      gl.Visible=True
    Next
  Else
    For Each gl in monthgirls
      gl.Visible=False
    Next
  End If
End Sub

Dim walldrop, dmdhide
If Table1.ShowDT = True Then
  dmdhide=1
  DMD.Visible=True
  DMD1.Visible=False
  For Each walldrop in cabwalls
    walldrop.Visible=True
  Next
  For Each walldrop in dtdisplay
    walldrop.Visible=True
  Next
Else
  dmdhide=1
  DMD.Visible=False
  DMD1.Visible=True
  For Each walldrop in cabwalls
    walldrop.Visible=False
  Next
  For Each walldrop in dtdisplay
    walldrop.Visible=False
  Next
End If

If B2SOn = True then
  dmdhide=0
  DMD.Visible=False
  DMD1.Visible=False
End If

Const UseSolenoids  = True
Const UseLamps    = True
Const UseSync   = True
Const UseGI     = False     'Only WPC games have special GI circuit.
Const SSolenoidOn = "SolOn"       'Solenoid activates
Const SSolenoidOff  = "SolOff"      'Solenoid deactivates
Const SFlipperOn  = "FlipperUp"   'Flipper activated
Const SFlipperOff = "FlipperDown" 'Flipper deactivated
Const SCoin     = "coin3"     'Coin inserted

Dim bubble
Set bubble=Kicker003.CreateSizedBall(6)
bubble.ReflectionEnabled=False
Kicker003.Kick 0, 0
Kicker003.Enabled=False
bubble.Image="bubble"

Sub table1_KeyDown(ByVal keycode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = LeftMagnaSave then maturepictures = not maturepictures:mature
  If keycode = RightMagnaSave then girllights = not girllights:changegl
  If keycode = PlungerKey Then Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal keycode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol SoundFX("Plunger",DOFContactors), Plunger, 1
End Sub

SolCallBack(1)      = "SolTrough"
SolCallBack(2)    = "vpmSolAutoPlunges Plunger1,SoundFX(""SolOn"",DOFContactors),0,"
SolCallBack(3)    = "leftramplockpost"
SolCallback(4)    = "LeftOrbitPost"
SolCallback(5)    = "dtDrop.SolDropUp"
SolCallback(6)    = "SolCLane"
SolCallback(7)    = "SolPeekLeft"'7=Bead Screen Left (peekaboo)
SolCallback(8)    = "SolPeekRight"'8=Bead Screen Right (peekaboo)
SolCallBack(12)   = "bsGrotto.SolOut"
SolCallback(13)   = "bsVUK.SolOut"
SolCallback(14)     = "magopenclose" '14=Magazine Post
'SolCallback(19)    = "SolTease1"'19=Drop Screen Stepper 1
'SolCallback(20)    = "SolTease2"'20=Drop Screen Stepper 2
'SolCallback(21)    = "SolTease3"'21=Drop Screen Stepper 3
SolCallback(22)   = "splashmotor"
'SolCallback(23)    = "SolTease4"'23=Drop Screen Stepper 4
solcallback(25)="splashflash"
SolCallBack(26) ="mirrorflash"
SolCallback(27)="RearFlashers"
SolCallback(28) = "LeftSlingFlasher"
SolCallBack(29) = "RightSlingFlasher"
'SolCallBack(30) = "vpmflasher f30,"
'31=Centerfold On/Off
'32=Centerfold Open/Close

'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"
'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
    ' PlaySoundAtVol SoundFX("fx_flipperup",DOFFlippers), LeftFlipper, VolFlip
    RandomSoundFlipperUpLeft
    LeftFlipper.RotateToEnd
    Else
    'PlaySoundAtVol SoundFX("fx_flipperdown",DOFFlippers), LeftFlipper, VolFlip
    RandomSoundFlipperDownLeft
    LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
    ' PlaySoundAtVol SoundFX("fx_flipperup",DOFFlippers), RightFlipper, VolFlip
    RandomSoundFlipperUpRight
    RightFlipper.RotateToEnd
  Else
    'PlaySoundAtVol SoundFX("fx_flipperdown",DOFFlippers), RightFlipper, VolFlip
    RandomSoundFlipperDownRight
    RightFlipper.RotateToStart
  End If
End Sub

Sub AutoPlunger(Enabled)
  If Enabled Then
    Plunger1.Fire
  End If
End Sub

Set GICallback = GetRef("GIUpdate")
Dim xx
Sub GIUpdate(no, Enabled)
    For each xx in GiLights
        xx.State = ABS(Enabled)
    Next
'    For each xx in GiFlashers
'        xx.visible = Enabled
'    Next
  striplightstimer.enabled = Light010.state
End Sub

Sub leftramplockpost(Enabled)
  If Enabled then
    balllockpost.IsDropped=True:PlaysoundAtVol SoundFX("solenoid",DOFContactors), balllockpost, 1
  Else
    balllockpost.IsDropped=False:PlaysoundAtVol SoundFX("solenoid",DOFContactors), balllockpost, 1
  End If
End Sub

Dim mirrorrnd
Sub mirrorflash(Enabled)
  If Enabled then
    'If maturepictures=True Then
      Randomize
      mirrorrnd=int(rnd*5)+1
      Pmirror.Image="mirror-nude-"&mirrorrnd
    'Else
      'Pmirror.Image="mirror-no-nude"
    'End If
  Else
    Pmirror.Image="mirrornotlit"
  End If
End Sub

mature
Dim bulkrnd
Sub mature
  If maturepictures=True Then
    Randomize
    bulkrnd=int(rnd*5)+1
    PCFTop.Image="centerfold-top-nude-"&bulkrnd:PCFMiddle.Image="centerfold-middle-nude-"&bulkrnd:PCFBottom.Image="centerfold-bottom-nude-"&bulkrnd
  Else
    PCFTop.Image="centerfold-top-non-nude":PCFMiddle.Image="centerfold-middle-non-nude":PCFBottom.Image="centerfold-bottom-no-nude"
  End If
  If maturepictures=True Then
    Randomize
    bulkrnd=int(rnd*5)+1
    Pmaginside.Image="maginside-nude-"&bulkrnd
  Else
    Pmaginside.Image="maginside-non-nude"
  End If
  If maturepictures=True Then
    bg1.Image="pb-background-1-off":bg2.Image="pb-background-1-off"
  Else
    bg1.Image="pb-background-2-off":bg2.Image="pb-background-2-off"
  End If
  If maturepictures=True Then
    Randomize
    bulkrnd=int(rnd*5)+1
    Ppeekaboo.Image="peek-nude-"&bulkrnd
  Else
    Ppeekaboo.Image="peek-no-nude-1"
  End If
  If maturepictures=True Then
    Randomize
    bulkrnd=int(rnd*5)+1
    Pteasepic.Image="tease-nude-"&bulkrnd
  Else
    Pteasepic.Image="tease-non-nude"
  End If
  If maturepictures=True Then
    Randomize
    bulkrnd=int(rnd*5)+1
    Psplashtriangle001.Image="splashtriangle-nude-"&bulkrnd
  Else
    Psplashtriangle001.Image="splashtriangle-non-nude"
  End If
End Sub

Dim splashdir
Sub splashmotor(Enabled)
  If Enabled then splashtimer.Enabled=True:splashdir=1 Else splashtimer.Enabled=True:splashdir=-1
End Sub

Sub splashtimer_Timer()
  Psplashtriangle001.RotY=Psplashtriangle001.RotY+splashdir
  If Psplashtriangle001.RotY=360 then Psplashtriangle001.RotY=0
  If Psplashtriangle001.RotY=200 or Psplashtriangle001.RotY=-160 then splashtimer.Enabled=False
End Sub

Dim magmov
Sub magopenclose(Enabled)
  If Enabled then magmov=1 Else magmov=-1
  magmove.Enabled=True
End Sub

Sub magmove_Timer()
  Pmagcover.RotZ=Pmagcover.RotZ+magmov
  If Pmagcover.RotZ = 0 or Pmagcover.RotZ = 100 then magmove.Enabled=False
End Sub

Sub RearFlashers(Enabled)
  If Enabled then f27a.Visible=True:f27b.Visible=True:f27c.Visible=True:Primitive033.DisableLighting=1:Primitive034.DisableLighting=1:f27d.State=False Else f27a.Visible=False:f27b.Visible=False:f27c.Visible=False:Primitive033.DisableLighting=0.2:Primitive034.DisableLighting=0.2:f27d.State=True
End Sub

Sub LeftSlingFlasher(Enabled)
  If Enabled then f28.Visible=True:Primitive007.DisableLighting=1:f28a.State=False Else f28.Visible=False:Primitive007.DisableLighting=0.2:f28a.State=True
End Sub

Sub RightSlingFlasher(Enabled)
  If Enabled then f29.Visible=True:Primitive011.DisableLighting=1:f29a.State=False:f29b.State=False:f29c.State=False:f29d.State=1 Else f29.Visible=False:Primitive011.DisableLighting=0.6:f29a.State=True:f29b.State=True:f29c.State=True:f29d.State=0
End Sub

Sub splashflash(Enabled)
  If Enabled Then Flasher25.Visible=True Else Flasher25.Visible=False
End Sub

Dim chainleft
Sub SolPeekLeft(Enabled)
  'If Enabled Then PeekReel.SetValue 0
  If Enabled Then
    For Each chainleft in chainopenleft
      chainleft.Visible=True
    Next
    For Each chainleft in chainclosedleft
      chainleft.Visible=False
    Next
  Else
    For Each chainleft in chainopenleft
      chainleft.Visible=False
    Next
    For Each chainleft in chainclosedleft
      chainleft.Visible=True
    Next
    cball.VelX=3:cball.VelY=3
  End If
End Sub

Dim chainright
Sub SolPeekRight(Enabled)
  'If Enabled Then PeekReel.SetValue 1
  If Enabled Then
    For Each chainright in chainopenright
      chainright.Visible=True
    Next
    For Each chainright in chainclosedright
      chainright.Visible=False
    Next
  Else
    For Each chainright in chainopenright
      chainright.Visible=False
    Next
    For Each chainright in chainclosedright
      chainright.Visible=True
    Next
    cball.VelX=3:cball.VelY=3
  End If
End Sub

Sub SolTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 15
  End If
End Sub

Sub LeftOrbitPost(Enabled)
  If Enabled Then
    LO.IsDropped=0
  Else
    LO.IsDropped=1
  End If
End Sub

Sub SolCLane(Enabled)
  If Enabled Then
    CLane.IsDropped=0
  Else
    CLane.IsDropped=1
  End If
End Sub

Dim bsTrough,bsVUK,dtDrop,bsGrotto,mTease,mCenterfold
Const cRegistryName = "Playboy"
Const cOptionsName = "Options"
Dim vpmDips
Dim PlayboyOptions
Dim OptRom
Dim OptYield
Dim OptSkip
Sub PlayboyShowDips
  If Not IsObject(vpmDips) Then
    Set vpmDips = New cvpmDips
    With vpmDips
      .AddForm 100, 100, "Playboy Game Settings"
      .AddFrameExtra 0, 0, 250, "ROM Version", &Hf0, Array("International Display (default)", &H00, "France Display", &H10, "German Display", &H20, "Italy Display", &H40, "Spain Display", &H80)
        .AddChkExtra 7,100,250,Array("Enable YieldTime (for slow computers)", &H100)
      .AddChkExtra 7,120,250,Array("Do not show this menu again at startup", &H200)
      .AddLabel 7,140,250,20,"Press F6 during play to bring up this menu."
      .AddLabel 7,160,250,20,"Quit and restart game for changes to take effect."
    End With
  End If
  PlayboyOptions = vpmDips.ViewDipsExtra(PlayboyOptions)
  SaveValue cRegistryName,cOptionsName,PlayboyOptions
  PlayboySetOptions
End Sub

Set vpmShowDips = GetRef("PlayboyShowDips")
Sub PlayboySetOptions
  OptRom = "playboys"
  OptYield = False
  OptSkip = False
  If PlayboyOptions And &H10 Then OptRom = "playboyf"
  If PlayboyOptions And &H20 Then OptRom = "playboyg"
  If PlayboyOptions And &H40 Then OptRom = "playboyi"
  If PlayboyOptions And &H80 Then OptRom = "playboyl"
  If PlayboyOptions And &H100 Then OptYield = True
  If PlayboyOptions And &H200 Then OptSkip = True
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw and Herweh
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' *******************************************************************************************************
' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Set position as table object (Use object or light but NOT wall) and Vol to 1
Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

' set position as table object and Vol + RndPitch manually
Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.
Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.
Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3
Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound
Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume

Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
  AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
  BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
  BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
  VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************
Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling
Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b
    BOT = GetBalls
  ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

  ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 Then 'AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
    End If

    ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************
Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub


' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Different sound
' Sub Gates_Hit (idx)
'   PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
' End Sub

' Sub Rubbers_Hit(idx)
'   dim finalspeed
'     finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
'   If finalspeed > 20 then
'     PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'   End if
'   If finalspeed >= 6 AND finalspeed <= 20 then
'     RandomSoundRubber()
'   End If
' End Sub
'
' Sub RandomSoundRubber()
'   Select Case Int(Rnd*3)+1
'     Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'     Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'     Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'   End Select
' End Sub

Sub table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

 '  SaveValue cRegistryName,cOptionsName,0  ' This will clear your Registry settings if uncommented

  PlayboyOptions = CInt("0" & LoadValue(cRegistryName,cOptionsName))

  PlayboySetOptions

  If optSkip = False Then vpmShowDips

  PlayboySetOptions
  With Controller
    If optYield = True Then : If .Version >= "01500000" Then Table1.YieldTime = 2 : End If
    .GameName=OptRom
    Plunger1.PullBack
    table1.Yieldtime=1
    .SplashInfoLine = "Playboy - Stern 2002" & vbNewLine & "VPX version by HiRez00 and Rascal"
    .ShowTitle = False
    .ShowDMDOnly = 1
    .Hidden = dmdhide
    .ShowFrame = False
    .HandleMechanics=0
    '.SetDisplayPosition 13,229
  .DIP(0)=&H00
  On Error Resume Next
  .Run
    If Err Then MsgBox Err.Description
  On Error Goto 0
  PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch = 56
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)
  vpmMapLights clights

    Set bsTrough = New cvpmBallStack
  bsTrough.InitSw 0,14,13,12,11,0,0,0
  bsTrough.InitKick BallRelease,90,3
  bsTrough.InitEntrySnd "Solenoid", "Solenoid"
  bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors), SoundFX("Solon",DOFContactors)
  bsTrough.Balls=4

  Set bsVUK=New cvpmBallStack
  bsVUK.InitSw 0,35,0,0,0,0,0,0
  bsVUK.InitKick Kicker2,270,8
  bsVUK.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("solon",DOFContactors)

  Set bsGrotto=New cvpmBallStack
  bsGrotto.InitSw 0,34,0,0,0,0,0,0
  bsGrotto.InitKick Kicker4,40,6
  bsGrotto.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("solon",DOFContactors)

  set dtDrop=new cvpmDropTarget
  dtDrop.InitDrop wall5,26
  dtDrop.InitSnd SoundFX("flapclos",DOFDropTargets),SoundFX("flapopen",DOFDropTargets)

  Set mTease=New cvpmMech
  mTease.MType=vpmMechStepSol+vpmMechReverse+vpmMechLinear
  mTease.Sol1=19
  mTease.Sol2=20
  mTease.Length=60
  mTease.Steps=120
  mTease.AddSw 52,0,0
  mTease.Callback=GetRef("UpdateTease")
  mTease.Start

  Set mCenterfold=New cvpmMech
  mCenterfold.MType=vpmMechStepSol+vpmMechReverse+vpmMechLinear
  mCenterfold.Sol1=31
  mCenterfold.Sol2=32
  mCenterfold.Length=165
  mCenterfold.Steps=170
  mCenterfold.AddSw 43,0,0
  mCenterfold.Callback=GetRef("UpdateCenterfold")
  mCenterfold.Start

 End with
End Sub

Sub UpdateTease(aNewPos,aSpeed,aLastPos)
  If aNewPos>0 And aNewPos<121 Then
    Pteasedoor.Z=120-aNewPos
  End If
End Sub

 Sub UpdateCenterfold(aNewPos,aSpeed,aLastPos)
  If aNewPos>-1 And aNewPos<170 Then
    If aNewPos<=5 then aNewPos=0
    If aNewPos<85 then PCFTop.ObjRotX=aNewPos*2
    If aNewPos>=85 then PCFBottom.ObjRotX=-(aNewPos-85)*2 Else PCFBottom.ObjRotX=0
  End If
End Sub

Sub Trigger13_Hit:Controller.Switch(9)=1:End Sub        '9
Sub Trigger13_unHit:Controller.Switch(9)=0:End Sub
Sub Trigger4_Hit:Controller.Switch(10)=1:End Sub        '10
Sub Trigger4_unHit:Controller.Switch(10)=0:End Sub
Sub Drain_Hit:bsTrough.AddBall Me:End Sub           '11
Sub Trigger6_Hit:Controller.Switch(16)=1:End Sub        '16
Sub Trigger6_unHit:Controller.Switch(16)=0:End Sub
Sub Trigger7_Hit:Controller.Switch(17)=1:End Sub        '17
Sub Trigger7_unHit:Controller.Switch(17)=0:End Sub
Sub Trigger8_Hit:Controller.Switch(18)=1:End Sub        '18
Sub Trigger8_unHit:Controller.Switch(18)=0:End Sub
Sub Trigger14_Hit:Controller.Switch(19)=1:End Sub       '19
Sub Trigger14_unHit:Controller.Switch(19)=0:End Sub
Sub Trigger9_Hit:Controller.Switch(21)=1:End Sub        '21
Sub Trigger9_unHit:Controller.Switch(21)=0:End Sub
Sub Trigger1_Hit:Controller.Switch(22)=1:End Sub        '22
Sub Trigger1_unHit:Controller.Switch(22)=0:End Sub
Sub Trigger2_Hit:Controller.Switch(23)=1:End Sub        '23
Sub Trigger2_unHit:Controller.Switch(23)=0:End Sub
Sub Trigger3_Hit:Controller.Switch(24)=1:End Sub        '24
Sub Trigger3_unHit:Controller.Switch(24)=0:End Sub
Sub Trigger12_Hit:Controller.Switch(25)=1:End Sub       '25
Sub Trigger12_unHit:Controller.Switch(25)=0:End Sub
Sub Wall5_Hit:dtDrop.Hit 1:End Sub                '26
Sub Trigger5_Hit:Controller.Switch(27)=1:End Sub        '27
Sub Trigger5_unHit:Controller.Switch(27)=0:End Sub
'29=Triangle Mech 1 (Right) on right ramp
'30=Triangle Mech 2 (Left) on right ramp
Sub T32_Hit:vpmTimer.PulseSw 32:End Sub             '32
Sub Target001_Hit:vpmTimer.PulseSw 33:End Sub         '33
Sub Kicker4_Hit:bsGrotto.AddBall Me:End Sub           '34
Sub Kicker1_Hit:bsVUK.AddBall Me:End Sub            '35
Sub Trigger17_Hit:Controller.Switch(38)=1:End Sub       '38
Sub Trigger17_unHit:Controller.Switch(38)=0:End Sub
Sub Trigger15_Hit:Controller.Switch(39)=1:End Sub       '39
Sub Trigger15_unHit:Controller.Switch(39)=0:End Sub
Sub Trigger16_Hit:Controller.Switch(40)=1:End Sub       '40
Sub Trigger16_unHit:Controller.Switch(40)=0:End Sub

Dim MyBall
Sub Trigger10_Hit
  Set MyBall=ActiveBall
  If MyBall.VelY>0 Then MyBall.VelX=MyBall.VelX+5
  Controller.Switch(41)=1
End Sub       '41

Sub Trigger10_unHit:Controller.Switch(41)=0:End Sub
Sub Trigger11_Hit:Controller.Switch(42)=1:End Sub       '42
Sub Trigger11_unHit:Controller.Switch(42)=0:End Sub
'43=Centerfold 1 (Closed)
'44=Centerfold 2 (Open)
Sub Bumper1_Hit:vpmTimer.PulseSw 49:bumpershake:RandomSoundBumperB:End Sub      '49
Sub Bumper2_Hit:vpmTimer.PulseSw 50:bumpershake:RandomSoundBumperB:End Sub      '50
Sub Bumper3_Hit:vpmTimer.PulseSw 51:bumpershake:RandomSoundBumperB:End Sub      '51
Dim chainmovx:chainmovx=1
Sub bumpershake()
  cball.VelY=-.5
  If chainmovx=1 then cball.VelX=-.5:chainmovx=2 Else cball.VelX=.5:chainmovx=1
End Sub
'52=Tease Screw Limit on backsplash board
'53=Tournament Button
Sub LeftSlingshot_Slingshot:vpmTimer.PulseSw 59:RandomSoundSlingshotLeft:End Sub      '59
                                '56 Tilt
Sub LeftOutlane_Hit:Controller.Switch(57)=1:End Sub       '57
Sub LeftOutlane_unHit:Controller.Switch(57)=0:End Sub
Sub LeftInlane_Hit:Controller.Switch(58)=1:End Sub        '58
Sub LeftInlane_unHit:Controller.Switch(58)=0:End Sub
Sub RightOutlane_Hit:Controller.Switch(60)=1:End Sub      '60
Sub RightOutlane_unHit:Controller.Switch(60)=0:End Sub
Sub RightInlane_Hit:Controller.Switch(61)=1:End Sub       '61
Sub RightInlane_unHit:Controller.Switch(61)=0:End Sub
Sub RightSlingshot_Slingshot:vpmTimer.PulseSw 62:RandomSoundSlingshotRight:End Sub    '62

Sub Kicker7_Hit:Me.DestroyBall:Set mball=Kicker004.CreateBall:Kicker8.TimerEnabled=True:End Sub
Sub Kicker8_Timer()
  Kicker004.DestroyBall
  Set mball=Kicker8.CreateBall
  Kicker8.Kick 180, 0
  Kicker8.TimerEnabled=False
End Sub
Sub Trigger18_Hit:ActiveBall.VelY=2:End Sub
Sub Trigger19_Hit:ActiveBall.VelY=2:End Sub

Sub Trigger001_Hit()
  ActiveBall.VelX=0
End Sub

Sub Kicker001_Hit()
  Kicker002.Enabled=True
  Kicker001.TimerEnabled=True
End Sub

Sub Kicker001_Timer()
  Kicker001.DestroyBall
  Set mball=Kicker002.CreateBall
  Kicker002.Kick 180, 3
  Kicker002.Enabled=False
  Kicker001.TimerEnabled=False
End Sub

Sub Timer002_Timer()
  If Light74.State = 1 then Flasher002.Visible=True Else Flasher002.Visible=False
  If Light75.State = 1 then Flasher001.Visible=True Else Flasher001.Visible=False
End Sub

Sub flippertimer_Timer()
  PLeftFlipper.RotY=LeftFlipper.CurrentAngle
  PRightFlipper.RotY=RightFlipper.CurrentAngle
End Sub

Dim slswitch:slswitch=0
Dim slitems
Sub striplightstimer_Timer()
  If slswitch = 0 Then
    slswitch=1
    For Each slitems in striplights1
      slitems.Visible=True
    Next
    For Each slitems in striplights2
      slitems.Visible=False
    Next
  Else
    slswitch=0
    For Each slitems in striplights1
      slitems.Visible=False
    Next
    For Each slitems in striplights2
      slitems.Visible=True
    Next
  End If
End Sub

Sub mirrorballtrigger_Hit()
  mirrorballtimer.Enabled=True
End Sub

Sub mirrorballtrigger_Unhit()
  mirrorballtimer.Enabled=False
End Sub

Sub mirrorballtimer_Timer()
  On Error Resume Next
  Pmirrorball.Z=mball.Y-20
  Pmirrorball.X=mball.X
End Sub

Dim mball
Sub mirrorballcreate_Hit()
  mirrorballcreate.DestroyBall
  Set mball=mirrorballcreate.CreateBall
  mirrorballcreate.kick 0, 0
  mirrorballcreate.Enabled = False
End Sub

Sub Gate1_Hit()
  mirrorballcreate.Enabled = True
End Sub

' *** Nudge
Dim mMagnet, cBall

Sub WobbleMagnet_Init
   Set mMagnet = new cvpmMagnet
   With mMagnet
    .InitMagnet WobbleMagnet, .9
    .Size = 150
    .CreateEvents mMagnet
    .MagnetOn = True
    .GrabCenter = False
   End With
  Set cBall = ckicker.CreateBall
' Set cBall = ckicker.CreateSizedBallWithMass(25, 1)
  ckicker.Kick 0,0:mMagnet.addball cball
End Sub

Dim chain
Sub chaintimer_Timer()
  For Each chain in chainsclosed
    chain.RotX=cBall.X-ckicker.X:chain.RotZ=cBall.Y-ckicker.Y
  Next
End Sub

Sub light45timer_Timer()
  If Light45.State=1 then grottolight.Visible=True Else grottolight.Visible=False
  If Light67.State=1 then L67a.Visible=True:L67b.Visible=True Else L67a.Visible=False:L67b.Visible=False
  If Light68.State=1 then L68.Visible=True Else L68.Visible=False
  If Light69.State=1 then L69.Visible=True Else L69.Visible=False
End Sub

Sub ballhitchains_Hit()
  cball.VelX=-4:cball.VelY=-4
  PlaySoundAtVol "chain-01a", ActiveBall, 1
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong()
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong()
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

Sub Gates_Hit(idx)
  PlaySoundAtVol "GateWire", ActiveBall, VolGates
End Sub

Sub RandomSoundRubberStrong()
debug.print "rubber strong ok"
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtVol "TOM_Rubber_Flipper_Slingshot_Strong_1", ActiveBall, VolRH
    Case 2 : PlaySoundAtVol "TOM_Rubber_Flipper_Slingshot_Strong_2", ActiveBall, VolRH
    Case 3 : PlaySoundAtVol "TOM_Rubber_Flipper_Slingshot_Strong_3", ActiveBall, VolRH
    Case 4 : PlaySoundAtVol "TOM_Rubber_Flipper_Slingshot_Strong_4", ActiveBall, VolRH
    Case 5 : PlaySoundAtVol "TOM_Rubber_Flipper_Slingshot_Strong_5", ActiveBall, VolRH
    Case 6 : PlaySoundAtVol "TOM_Rubber_Flipper_Slingshot_Strong_6", ActiveBall, VolRH
    Case 7 : PlaySoundAtVol "TOM_Rubber_Flipper_Slingshot_Strong_7", ActiveBall, VolRH
    Case 8 : PlaySoundAtVol "TOM_Rubber_Flipper_Slingshot_Strong_8", ActiveBall, VolRH
    Case 9 : PlaySoundAtVol "TOM_Rubber_Flipper_Slingshot_Strong_9", ActiveBall, VolRH
    Case 10 : PlaySoundAtVol "TOM_Rubber_1_Hard", ActiveBall, VolRH
  End Select
End Sub

Sub RandomSoundRubberWeak()
debug.print "rubber weak ok"
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtVol "TOM_Rubber_1", ActiveBall, VolRH * 0.5
    Case 2 : PlaySoundAtVol "TOM_Rubber_2", ActiveBall, VolRH * 0.5
    Case 3 : PlaySoundAtVol "TOM_Rubber_3", ActiveBall, VolRH * 0.5
    Case 4 : PlaySoundAtVol "TOM_Rubber_3", ActiveBall, VolRH * 0.5
    Case 5 : PlaySoundAtVol "TOM_Rubber_5", ActiveBall, VolRH * 0.5
    Case 6 : PlaySoundAtVol "TOM_Rubber_6", ActiveBall, VolRH * 0.5
    Case 7 : PlaySoundAtVol "TOM_Rubber_7", ActiveBall, VolRH * 0.5
    Case 8 : PlaySoundAtVol "TOM_Rubber_8", ActiveBall, VolRH * 0.5
    Case 9 : PlaySoundAtVol "TOM_Rubber_9", ActiveBall, VolRH * 0.5
  End Select
End Sub

Sub RandomSoundRollover()
debug.print "rollover ok"
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtVol "TOM_Rollover_1", ActiveBall, VolRol
    Case 2 : PlaySoundAtVol "TOM_Rollover_2", ActiveBall, VolRol
    Case 3 : PlaySoundAtVol "TOM_Rollover_3", ActiveBall, VolRol
    Case 4 : PlaySoundAtVol "TOM_Rollover_4", ActiveBall, VolRol
  End Select
End Sub

'********** FLIPPER - UP - LEFT - DOF

Sub RandomSoundFlipperUpLeft()
  Select Case Int(Rnd*11)+1
    Case 1 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L01",DOFFlippers), LeftFlipper, VolFlip
    Case 2 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L02",DOFFlippers), LeftFlipper, VolFlip
    Case 3 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L07",DOFFlippers), LeftFlipper, VolFlip
    Case 4 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L08",DOFFlippers), LeftFlipper, VolFlip
    Case 5 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L09",DOFFlippers), LeftFlipper, VolFlip
    Case 6 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L10",DOFFlippers), LeftFlipper, VolFlip
    Case 7 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L12",DOFFlippers), LeftFlipper, VolFlip
    Case 8 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L14",DOFFlippers), LeftFlipper, VolFlip
    Case 9 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L18",DOFFlippers), LeftFlipper, VolFlip
    Case 10 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L20",DOFFlippers), LeftFlipper, VolFlip
    Case 11 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L26",DOFFlippers), LeftFlipper, VolFlip
  End Select
End Sub

'********** FLIPPER - DOWN - LEFT - DOF

Sub RandomSoundFlipperDownLeft()
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Left_Down_1_Bass",DOFFlippers), LeftFlipper, VolFlipD
    Case 2 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Left_Down_2_Bass",DOFFlippers), LeftFlipper, VolFlipD
    Case 3 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Left_Down_3_Bass",DOFFlippers), LeftFlipper, VolFlipD
    Case 4 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Left_Down_4_Bass",DOFFlippers), LeftFlipper, VolFlipD
    Case 5 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Left_Down_5_Bass",DOFFlippers), LeftFlipper, VolFlipD
    Case 6 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Left_Down_6_Bass",DOFFlippers), LeftFlipper, VolFlipD
    Case 7 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Left_Down_7_Bass",DOFFlippers), LeftFlipper, VolFlipD
  End Select
End Sub

'********** FLIPPER - UP - RIGHT - DOF

Sub RandomSoundFlipperUpRight()
  Select Case Int(Rnd*11)+1
    Case 1 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R01",DOFFlippers), RightFlipper, VolFlip
    Case 2 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R02",DOFFlippers), RightFlipper, VolFlip
    Case 3 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R03",DOFFlippers), RightFlipper, VolFlip
    Case 4 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R04",DOFFlippers), RightFlipper, VolFlip
    Case 5 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R05",DOFFlippers), RightFlipper, VolFlip
    Case 6 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R06",DOFFlippers), RightFlipper, VolFlip
    Case 7 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R07",DOFFlippers), RightFlipper, VolFlip
    Case 8 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R08",DOFFlippers), RightFlipper, VolFlip
    Case 9 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R09",DOFFlippers), RightFlipper, VolFlip
    Case 10 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R10",DOFFlippers), RightFlipper, VolFlip
    Case 11 : PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R11",DOFFlippers), RightFlipper, VolFlip
  End Select
End Sub

'********** FLIPPER - DOWN - RIGHT - DOF

Sub RandomSoundFlipperDownRight()
  Select Case Int(Rnd*8)+1
    Case 1 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Right_Down_1_Bass",DOFFlippers), RightFlipper, VolFlipD
    Case 2 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Right_Down_2_Bass",DOFFlippers), RightFlipper, VolFlipD
    Case 3 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Right_Down_3_Bass",DOFFlippers), RightFlipper, VolFlipD
    Case 4 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Right_Down_4_Bass",DOFFlippers), RightFlipper, VolFlipD
    Case 5 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Right_Down_5_Bass",DOFFlippers), RightFlipper, VolFlipD
    Case 6 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Right_Down_6_Bass",DOFFlippers), RightFlipper, VolFlipD
    Case 7 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Right_Down_7_Bass",DOFFlippers), RightFlipper, VolFlipD
    Case 8 : PlaySoundAtVol SoundFX("WD_TOM_Flipper_Right_Down_8_Bass",DOFFlippers), RightFlipper, VolFlipD
  End Select
End Sub

'********* SLINGSHOT - LEFT

Sub RandomSoundSlingshotLeft()
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_L1_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 2 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_L2_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 3 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_L3_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 4 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_L4_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 5 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_L5_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 6 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_L6_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 7 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_L7_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 8 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_L8_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 9 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_L9_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 10 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_L10_Strong_Layered",DOFContactors), ActiveBall, VolSling
  End Select
End Sub

'********* SLINGSHOT - RIGHT

Sub RandomSoundSlingshotRight()
  Select Case Int(Rnd*8)+1
    Case 1 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_R1_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 2 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_R2_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 3 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_R3_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 4 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_R4_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 5 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_R5_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 6 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_R6_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 7 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_R7_Strong_Layered",DOFContactors), ActiveBall, VolSling
    Case 8 : PlaySoundAtVol SoundFX("TOM_Calle_Sling_Rework_R9_Strong_Layered",DOFContactors), ActiveBall, VolSling
  End Select
End Sub

Sub RandomSoundBumperB()
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtVol SoundFX("TOM_Bumpers_Reworked_v2_Middle_1",DOFContactors), ActiveBall, VolBump
    Case 2 : PlaySoundAtVol SoundFX("TOM_Bumpers_Reworked_v2_Middle_2",DOFContactors), ActiveBall, VolBump
    Case 3 : PlaySoundAtVol SoundFX("TOM_Bumpers_Reworked_v2_Middle_3",DOFContactors), ActiveBall, VolBump
    Case 4 : PlaySoundAtVol SoundFX("TOM_Bumpers_Reworked_v2_Middle_4",DOFContactors), ActiveBall, VolBump
    Case 5 : PlaySoundAtVol SoundFX("TOM_Bumpers_Reworked_v2_Middle_5",DOFContactors), ActiveBall, VolBump
  End Select
End Sub

