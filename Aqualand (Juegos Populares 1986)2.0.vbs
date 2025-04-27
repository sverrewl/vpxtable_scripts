option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cgamename = "aqualand", UseSolenoids=1, UseLamps=1,UseGI=0, SCoin="fx_coin"

LoadVPM "01520000", "juegos.vbs", 3.1

Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
End if


If Table1.ShowDT = false then
    for each xx in aBackdrop: xx.visible = 0:next
  else
    for each xx in aBackdrop: xx.visible = 1:next
End If

'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------

SolCallback(8)="Flasher8.state="
'SolCallback(9)="vpmSolSound SoundFX(""left_slingshot""),"
'SolCallback(10)="vpmSolSound SoundFX(""right_slingshot""),"
'SolCallback(11)="vpmSolSound SoundFX(""fx_bumper3""),"
'SolCallback(12)="vpmSolSound SoundFX(""fx_bumper3""),"
SolCallback(13)= "dtDrop.SolDropUp" 'Drop Targets
SolCallback(14)="bsSaucer.SolOut"
SolCallback(15)= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(16)="bsTrough.SolOut"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
       PlaySoundAtVol SoundFX("fx_FlipperUp",DOFFlippers), LeftFlipper, 1 : LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_FlipperDown",DOFFlippers), LeftFlipper, 1 : LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
       PlaySoundAtVol SoundFX("fx_FlipperUp",DOFFlippers), RightFlipper, 1 : RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_FlipperDown",DOFFlippers), RightFlipper, 1: RightFlipper.RotateToStart
     End If
End Sub

'Playfield GI
'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'************
' Table init.
'************
Dim bsTrough,dtDrop,bsSaucer

Sub Table1_Init

  vpminit me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = ""&chr(13)&"You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 1
        .Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=31
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj = Array(RightSlingShot,LeftSlingShot,Bumper1,Bumper2) '


  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,25,0,0,0,0,0,0
  bsTrough.InitKick BallRelease,90,3
  bsTrough.InitExitSnd SoundFX("fx_ballrel",DOFContactors),SoundFX("fx_solenoid",DOFContactors)
  bsTrough.Balls=1

  Set dtDrop=New cvpmDropTarget
  dtDrop.InitDrop Array(S8,S7,S6),Array(8,7,6)
  dtDrop.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)

  Set bsSaucer=New cvpmBallStack
  bsSaucer.InitSaucer S12,12,283,15
  bsSaucer.InitExitSnd SoundFX("Popper_ball",DOFContactors),SoundFX("fx_solenoid",DOFContactors)

  vpmMapLights Goodlights
  vpmMapLights aLights

  VariTarget_Init

Controller.Games(cGameName).Settings.Value("volume") = -12.0

End Sub

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal Keycode)
  If keycode = LeftMagnaSave Then bLutActive = True
  If keycode = RightMagnaSave Then
    If bLutActive Then NextLUT: End If
    End If
    If KeyCode=LeftFlipperKey Then Controller.Switch(84)=1
  If KeyCode=RightFlipperKey Then Controller.Switch(82)=1
  If KeyCode=PlungerKey Then Plunger.Pullback:PlaySoundAtVol"plungerpull",Plunger,1
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
    If keycode = LeftMagnaSave Then bLutActive = False
  If KeyCode=LeftFlipperKey Then Controller.Switch(84)=0
  If KeyCode=RightFlipperKey Then Controller.Switch(82)=0
  If KeyCode=PlungerKey Then PlaySoundAtVol"plunger", Plunger, 1:Plunger.Fire
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 13
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), ActiveBall, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
' gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 14
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), ActiveBall, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
' gi3.State = 0:Gi4.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

'------------------------------
'------  Trough Handler  ------
'------------------------------

Sub S25_Hit:bsTrough.AddBall me:PlaySoundAtVol"Drain", s25, 1:End Sub                  '25
Sub S12_Hit:bsSaucer.AddBall 0:PlaySoundAtVol "popper_ball", s12, 1: End Sub       '12

'**** Target

Sub S11_Hit:vpmTimer.PulseSw 11: PlaySoundAtVol "target", ActiveBall, 1:End Sub '11

'**** RollOvers

Sub S1_Hit:Controller.Switch(1)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub          '1
Sub S1_unHit:Controller.Switch(1)=0:End Sub
Sub S2_Hit:Controller.Switch(2)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub          '2
Sub S2_unHit:Controller.Switch(2)=0:End Sub
Sub S3_Hit:Controller.Switch(3)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub          '3
Sub S3_unHit:Controller.Switch(3)=0:End Sub
Sub S4_Hit:Controller.Switch(4)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub          '4
Sub S4_unHit:Controller.Switch(4)=0:End Sub
Sub S5_Hit:Controller.Switch(5)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub          '5
Sub S5_unHit:Controller.Switch(5)=0:End Sub
Sub S9_Hit:Controller.Switch(9)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub          '9
Sub S9_unHit:Controller.Switch(9)=0:End Sub
Sub S10_Hit:Controller.Switch(10)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub        '10
Sub S10_unHit:Controller.Switch(10)=0:End Sub
Sub S33_Hit:Controller.Switch(33)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub        '33
Sub S33_unHit:Controller.Switch(33)=0:End Sub
Sub S34_Hit:Controller.Switch(34)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub        '34
Sub S34_unHit:Controller.Switch(34)=0:End Sub
Sub S35_Hit:Controller.Switch(35)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub        '35
Sub S35_unHit:Controller.Switch(35)=0:End Sub
Sub S36_Hit:Controller.Switch(36)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub        '36
Sub S36_unHit:Controller.Switch(36)=0:End Sub
Sub S37_Hit:Controller.Switch(37)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub        '37
Sub S37_unHit:Controller.Switch(37)=0:End Sub

'**** Drop Targets

Sub S6_Hit:dtDrop.Hit 3:me.timerenabled=1:End Sub     '6
Sub S7_Hit:dtDrop.Hit 2:me.timerenabled=1:End Sub     '7
Sub S8_Hit:dtDrop.Hit 1:me.timerenabled=1:End Sub     '8

'****Spinner

Sub S17_Spin:vpmTimer.PulseSw 17:PlaySoundAtVol "fx_spinner", S17, 1:End Sub         '17

'****Rubber Walls

Sub S18A_Hit:vpmTimer.PulseSw 18:End Sub          '18
Sub S18B_Hit:vpmTimer.PulseSw 18:End Sub
Sub S18C_Hit:vpmTimer.PulseSw 18:End Sub


'**** Gates

Sub Gate2_Hit:PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub
Sub Gate3_Hit:PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub
Sub Gate4_Hit:PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub
Sub Gate5_Hit:PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub
Sub Gate5_Hit:PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub

'------------------------------
'------  Bumpers  ------
'------------------------------

Sub Bumper1_hit : vpmTimer.PulseSw(15) : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors),ActiveBall, 1: End Sub
Sub Bumper2_hit : vpmTimer.PulseSw(16) : PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors),ActiveBall, 1: End Sub


Dim Old28,New28,Old97,New97
Old28=0:New28=0:Old97=0:New97=0

Set LampCallback=GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
  New28=G28.State
    If New28<>Old28 Then
      If New28=1 Then
        'VariReset
        VariTargetTimer.Enabled = True
      End If
    End If
  Old28=New28

  New97=G97.State
    If New97<>Old97 Then
      If New97=1 Then
        Flipper1.RotateToEnd
      Else
        Flipper1.RotateToStart
      End If
    End If
  Old97=New97

End Sub

'***********************************************************************************
'****                 VariTarget Handling                   ****
'***********************************************************************************
Dim VariNewPos, VariOPos, VariSwitch, i
Dim VariSwitches: VariSwitches = Array(24, 23, 22, 21, 20, 19)
Dim VariPositions: VariPositions = Array(0, 4, 8, 12, 16, 20)
Const VT_Delay_Factor = 0.87                  'used to slow down the ball when hitting the vari target
dim switchset

Sub vtn_Hit(vidx)
  if switchset=0 then
  VariNewPos=vidx
  If ((ActiveBall.VelY < 0) AND (vidx >= VariOPos)) Then
    ActiveBall.VelY = ActiveBall.VelY * VT_Delay_Factor
    PlaySoundAtVol SoundFX("fx_solenoidoff",DOFContactors), ActiveBall, 1
    DOF 101, 2
    For i = 0 to 5:If vidx >= VariPositions(i) Then VariSwitch = VariSwitches(i):End If:Next
    VariOPos=vidx: VariTargetP.TransZ = (VariOPos * 10)
  End If
  If ActiveBall.VelY >= 0 Then
    controller.switch(variswitch)=1:switchset=1
  End If
  end if
End Sub

Sub VariTargetTimer_Timer
  If VariTargetP.TransZ > 0 Then
    VariTargetP.TransZ = VariTargetP.TransZ - 5
  Else
    VariOPos = 0: Me.Enabled = 0:controller.switch(variswitch)=0:switchset = 0':textbox1.text="WIEDER AUS"
  End If
End Sub

Sub VariTarget_Init
  VariNewPos=0: VariOPos=0: VariSwitch = 0
End Sub


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(32)
Digits(0)=Array(Light1,Light2,Light3,Light4,Light5,Light6,Light7,Light8)
Digits(1)=Array(Light15,Light9,Light10,Light11,Light12,Light13,Light14)
Digits(2)=Array(Light22,Light16,Light17,Light18,Light19,Light20,Light21)
Digits(3)=Array(Light29,Light23,Light24,Light25,Light26,Light27,Light28,Light30)
Digits(4)=Array(Light37,Light31,Light32,Light33,Light34,Light35,Light36)
Digits(5)=Array(Light44,Light38,Light39,Light40,Light41,Light42,Light43)
Digits(6)=Array(Light51,Light45,Light46,Light47,Light48,Light49,Light50)
Digits(7)=Array(Light103,Light104,Light105,Light106,Light107,Light108,Light109,Light110)
Digits(8)=Array(Light117,Light111,Light112,Light113,Light114,Light115,Light116)
Digits(9)=Array(Light124,Light118,Light119,Light120,Light121,Light122,Light123)
Digits(10)=Array(Light131,Light125,Light126,Light127,Light128,Light129,Light130,Light132)
Digits(11)=Array(Light139,Light133,Light134,Light135,Light136,Light137,Light138)
Digits(12)=Array(Light146,Light140,Light141,Light142,Light143,Light144,Light145)
Digits(13)=Array(Light153,Light147,Light148,Light149,Light150,Light151,Light152)
Digits(14)=Array(Light52,Light53,Light54,Light55,Light56,Light57,Light58,Light59)
Digits(15)=Array(Light66,Light60,Light61,Light62,Light63,Light64,Light65)
Digits(16)=Array(Light73,Light67,Light68,Light69,Light70,Light71,Light72)
Digits(17)=Array(Light80,Light74,Light75,Light76,Light77,Light78,Light79,Light81)
Digits(18)=Array(Light88,Light82,Light83,Light84,Light85,Light86,Light87)
Digits(19)=Array(Light95,Light89,Light90,Light91,Light92,Light93,Light94)
Digits(20)=Array(Light102,Light96,Light97,Light98,Light99,Light100,Light101)
Digits(21)=Array(Light154,Light155,Light156,Light157,Light158,Light159,Light160,Light161)
Digits(22)=Array(Light168,Light162,Light163,Light164,Light165,Light166,Light167)
Digits(23)=Array(Light175,Light169,Light170,Light171,Light172,Light173,Light174)
Digits(24)=Array(Light182,Light176,Light177,Light178,Light179,Light180,Light181,Light183)
Digits(25)=Array(Light190,Light184,Light185,Light186,Light187,Light188,Light189)
Digits(26)=Array(Light197,Light191,Light192,Light193,Light194,Light195,Light196)
Digits(27)=Array(Light204,Light198,Light199,Light200,Light201,Light202,Light203)
Digits(28)=Array(Light211,Light205,Light206,Light207,Light208,Light209,Light210)
Digits(29)=Array(Light218,Light212,Light213,Light214,Light215,Light216,Light217)
Digits(30)=Array(Light225,Light219,Light220,Light221,Light222,Light223,Light224)
Digits(31)=Array(Light232,Light226,Light227,Light228,Light229,Light230,Light231)
Digits(32)=Array(Light239,Light233,Light234,Light235,Light236,Light237,Light238,Light240)

Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLED=Controller.ChangedLEDs(&Hffffffff,&Hffffffff)
  If Not IsEmpty(ChgLED) Then
    For ii=0 To UBound(chgLED)
      num=chgLED(ii,0):chg=chgLED(ii,1):stat=chgLED(ii,2)
      For Each obj In Digits(num)
        If chg And 1 Then obj.State=stat And 1
        chg=chg\2:stat=stat\2
      Next
    Next
  End If
End Sub

 'Aqualand
Sub editDips
Dim vpmDips:Set vpmDips=New cvpmDips
With vpmDips
.AddForm 700,600,"Aqualand - DIP switches"
.AddFrame 0,0,190,"25 psts coin slot - credits",5,Array("2:1",5,"1:1",0,"1:2",4,"1:3",1)'vpm dip 1 and 3      '80
.AddFrame 220,0,190,"100 psts coin slot - credits",80,Array("1:1",0,"1:2",64,"1:3",16,"1:10",80)'vpm dip 5 and 7  '80
.AddFrame 0,80,190,"Balls per game",1024,Array("3 balls",0,"5 balls",1024)'vpm dip 11               '46
.AddFrame 220,80,150,"Match Feature",32768,Array("ON",32768,"OFF",0)'vpm dip 16                   '46
.AddLabel 106,130,270,20,"Disable score replays by checking both of these"                      '15
.AddChk 126,145,200,Array("Ignore first replay if under 1.000.000",32)'vpm dip 6                  '15
.AddChk 126,160,200,Array("Only allow 1 score replay per game",64)'vpm dip 7                    '15
.AddChk 21,185,190,Array("Enable Extra Ball via score award:",8)'vpm dip 4                      '15
.AddLabel 160,200,280,20,"600.000"
.AddLabel 160,214,280,20,"700.000"
.AddLabel 160,228,280,20,"900.000"
.AddLabel 160,242,280,20,"800.000"
.AddFrame 220,186,190,"Score Replay Levels",768,Array("800.000 - 1.000,000 - 1.200.000",0,"900.000 - 1.000,000 - 1.200.000",256,"1.000.000 - 1.500.000 - 2.000.000",768,"1.000.000 - 1.000.000 - 1.400.000",512)'vpm dip 9 and 10 '80
.AddFrame 160,280,100,"Bookkeeping",524288,Array("bookkeeping off",0,"coin audits",262144,"play audits",524288)'vpm dip 19 and 20   '75
.AddLabel 80,350,280,20,"After hitting OK, press F3 to reset game with new settings."
.ViewDips
End With
End Sub
Set vpmShowDips=GetRef("editDips")

' dip switches

 'Coins 1 (4 on keyboard) = dip 0 switch 1/dip 0 switch 3
 'OFF OFF = 1:1   OFF ON =1:2
 'ON ON   = 2:1   ON OFF =1:3
 'Coins 2 (5 on keyboard) = dip 0 switch 5/dip 0 switch 7
 'OFF OFF = 1:1   OFF ON =1:2
 'ON ON   = 1:10  ON OFF =1:3

 'Extra Ball Award = dip 0 switch 4
 'Replay Level A = ON = 600.000 EB
 'Replay Level B = ON = 700.000 EB
 'Replay Level C = ON = 900.000 EB
 'Replay Level D = ON = 800.000 EB

 'Replay Levels = dip 1 switch 1/dip 1 switch 2
 'Replay Level A = OFF OFF=800.000/1.000.000/1.200.000
'Replay Level B = ON OFF=900.000/1.000.000/1.200.000
 'Replay Level C = ON ON =1.000.000/1.500.000/2.000.000
 'Replay Level D = OFF ON=1.000.000/1.000.000/1.400.000

 'BALLS = Controller.Dip 1 Switch 3 - off=3/on=5
'Ignore first replay level if under 1.000.000= dip 0 switch 6
 'ON=Ignore, OFF=Award Free Game
 'Only allow 1 score-based replay per game - to disable score based replays, enable 'ignore first replay level' feature
 'dip 0 switch 7
 'MATCH = Controller.Dip 1 Switch 8 - off=disabled/on=enabled

  'Coin Audits = Controller.Dip 2 Switch 3 - off=disabled/on=enabled
'Display Credit audit at end of each game = dip 2 switch 4
 'ON=Enabled, OFF=Disabled

'                    *           *     *     *      *      *      *
'Controller.Dip(0)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '01-08
'                    *     *     *                                *
'Controller.Dip(1)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '09-16
'                                *     *
'Controller.Dip(2)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) 'A-D

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
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
'            Supporting Ball, Sound Functions and Math
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
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
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function VolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  VolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
End Function

Function DVolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  DVolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
  debug.print DVolMulti
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
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
				  ' PlaySound("fx_ballrolling" & b), -1, BallRollVol(BOT(b) )*.8, AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
				  ' PlaySound("fx_ballrolling" & b), -1, BallRollVol(BOT(b) )*.2, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
        End If
      End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()

  pleftFlipper.objrotz=leftFlipper.CurrentAngle
  prightFlipper.objrotz=rightFlipper.CurrentAngle

    diverterP.RotZ = Flipper1.CurrentAngle
End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT)<(tnob-1) Then
        For b = (UBound(BOT) + 1) to (tnob-1)
            BallShadow(b).visible = 0
        Next
    End If
    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'*************
'   JP'S LUT
'*************

Dim bLutActive, LUTImage
Sub LoadLUT
Dim x
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
  UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT: LUTImage = (LUTImage +1 ) MOD 10: UpdateLUT: SaveLUT: End Sub

Sub UpdateLUT
Select Case LutImage
Case 0: table1.ColorGradeImage = "LUT0"
Case 1: table1.ColorGradeImage = "LUT1"
Case 2: table1.ColorGradeImage = "LUT2"
Case 3: table1.ColorGradeImage = "LUT3"
Case 4: table1.ColorGradeImage = "LUT4"
Case 5: table1.ColorGradeImage = "LUT5"
Case 6: table1.ColorGradeImage = "LUT6"
Case 7: table1.ColorGradeImage = "LUT7"
Case 8: table1.ColorGradeImage = "LUT8"
Case 9: table1.ColorGradeImage = "LUT9"

End Select
End Sub

'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


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

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "Goma", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
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

Sub Table1_Exit()
  Controller.Pause = False
  Controller.Stop
End Sub
