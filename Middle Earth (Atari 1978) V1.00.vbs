Option Explicit


Const BallMass = 1.1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01200000", "ATARI1b.VBS", 3.1

Dim BallSound, GIURC, lgiurc,bgiurc,lbgiurc

Const cGameName     = "midearth"   ' PinMAME short name
Const cCredits      = "VPX table by BalateR"
Const UseSolenoids  = True
Const UseLamps      = True
Const UseGI         = False

' Standard Sounds

Const SFlipperOn    = "FlipperUp"
Const SFlipperOff   = "FlipperDown"
Const SCoin         = "coin3"
Const SSolenoidOn   = "solon"
Const SSolenoidOff  = "soloff"



'Solenoid Definitions
Const sOuthole    = 2'1
Const sRSling   = 9'2
Const sRReset   = 10'3
Const sBumper2    = 1'4
Const sBumper1    = 4'5
Const sLReset   = 11'6
Const sLSling   = 13'7
Const sKnocker    = 7


'Solenoid Callbacks
SolCallback(sOutHole)   = "bsTrough.SolOut"
SolCallback(sRSling)    = "vpmSolSound SoundFX(""sling"",DOFContactors),"
SolCallback(sRReset)    = "dtR.SolDropUp"
SolCallback(sBumper2)   = "vpmSolSound SoundFX(""Jet3"",DOFContactors),"
SolCallback(sBumper1)   = "vpmSolSound SoundFX(""Jet3"",DOFContactors),"
SolCallback(sLReset)    = "dtL.SolDropUp"
SolCallback(sLSling)    = "vpmSolSound SoundFX(""sling"",DOFContactors),"

Dim bsTrough,dtL,dtR

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

    On Error Resume Next
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = cCredits
    .HandleMechanics = 0
    .ShowDMDOnly = 1
    .ShowFrame = False
    .ShowTitle = False
    .Run
    .Hidden=1
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
    PinMAMETimer.Enabled=True
    ' Nudging
  vpmNudge.TiltSwitch = 18
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj=Array(RightSlingshot,LeftSlingshot)


  Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 0,41,0,0,0,0,0,0
    bsTrough.InitKick BallRelease,80,5
    bsTrough.InitExitSnd SoundFX("Ballrel",DOFContactors),SoundFX("solon",DOFContactors)
    bsTrough.Balls=1

  Set dtL=New cvpmDropTarget
  dtL.InitDrop Array(sw57,sw58,sw59,sw60,sw53),Array(57,58,59,60,53)
  dtL.InitSnd SoundFX("fx_target",DOFContactors),SoundFX("fx_resetdrop",DOFContactors)

  Set dtR=New cvpmDropTarget
  dtR.InitDrop Array(sw39,sw40,sw52,sw50,sw56),Array(39,40,52,50,56)
  dtR.InitSnd SoundFX("fx_target",DOFContactors),SoundFX("fx_resetdrop",DOFContactors)

End Sub


Sub Table1_KeyDown(ByVal KeyCode)
 If KeyCode=LeftFlipperKey And tilt.state=0 And BIP.state=1 Then
   Controller.Switch(81)=1
   LeftFlipper.RotateToEnd
    LeftFlipperU.RotateToEnd
    PlaySoundAtVol"flipperup", LeftFlipper, 1
  End If
  If KeyCode=RightFlipperKey And tilt.state=0 And BIP.state=1 Then
    Controller.Switch(82)=1
   RightFlipper.RotateToEnd
    RightFlipperU.RotateToEnd
   PlaySoundAtVol"flipperup", RightFlipper, 1
  End If
  If vpmKeyDown(KeyCode) Then Exit Sub
  If KeyCode=PlungerKey Then Plunger.PullBack
End  Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode=LeftFlipperKey And tilt.state=0 And BIP.state=1 Then
   Controller.Switch(81)=0
   LeftFlipper.RotateToStart
    LeftFlipperU.RotateToStart
    PlaySoundAtVol"flipperdown", LeftFlipper, 1
  End If
  If KeyCode=RightFlipperKey And tilt.state=0 And BIP.state=1 Then
    Controller.Switch(82)=0
   RightFlipper.RotateToStart
    RightFlipperU.RotateToStart
   PlaySoundAtVol"flipperdown", RightFlipper, 1
  End If
  If vpmKeyUp(KeyCode) Then Exit Sub
  If KeyCode=PlungerKey Then PlaySoundAtVol"plunger", Plunger, 1:Plunger.Fire
End Sub


Sub sw21_Hit:Controller.Switch(21)=1:End Sub      '21
Sub sw21_unHit:Controller.Switch(21)=0:End Sub
Sub sw22_Hit:Controller.Switch(22)=1:End Sub      '22
Sub sw22_unHit:Controller.Switch(22)=0:End Sub
Sub sw33_Hit:Controller.Switch(33)=1:End Sub      '33
Sub sw33_unHit:Controller.Switch(33)=0:End Sub
Sub sw34_Hit:Controller.Switch(34)=1:End Sub      '34
Sub sw34_unHit:Controller.Switch(34)=0:End Sub
Sub sw35_Hit:Controller.Switch(35)=1:End Sub      '35
Sub sw35_unHit:Controller.Switch(35)=0:End Sub
Sub sw36_Hit:Controller.Switch(36)=1:End Sub      '36
Sub sw36_unHit:Controller.Switch(36)=0:End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:End Sub        '37
Sub sw38_Hit:vpmTimer.PulseSw 38:End Sub        '38
Sub sw39_Dropped:dtR.Hit 1:End Sub            '39
Sub sw40_Dropped:dtR.Hit 2:End Sub            '40
Sub Drain_Hit:bsTrough.AddBall Me:End Sub       '41
Sub sw44_Hit:Controller.Switch(44)=1:End Sub      '44
Sub sw44_unHit:Controller.Switch(44)=0:End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:End Sub        '45
Sub sw46_Hit:vpmTimer.PulseSw 46:End Sub        '46
Sub sw47_Hit:vpmTimer.PulseSw 47:End Sub        '47
Sub sw48_Hit:Controller.Switch(48)=1:End Sub      '48
Sub sw48_unHit:Controller.Switch(48)=0:End Sub
Sub sw50_Dropped:dtR.Hit 4:End Sub            '50
Sub sw52_Dropped:dtR.Hit 3:End Sub            '52
Sub sw53_Dropped:dtL.Hit 5:End Sub            '53
Sub sw56_Dropped:dtR.Hit 5:End Sub            '56
Sub sw57_Dropped:dtL.Hit 1:End Sub            '57
Sub sw58_Dropped:dtL.Hit 2:End Sub            '58
Sub sw59_Dropped:dtL.Hit 3:End Sub            '59
Sub sw60_Dropped:dtL.Hit 4:End Sub            '60
Sub Bumper1_Hit : vpmTimer.PulseSw(23) :  playsoundAtVol "fx_bumper1", ActiveBall, 1:End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(24) :  playsoundAtVol "fx_bumper1", ActiveBall, 1:End Sub
Sub sw49_slingshot(idx):playsoundAtVol"sling", ActiveBall, 1:vpmTimer.PulseSw 49:End Sub      '49
Sub sw54_slingshot(idx):playsoundAtVol"sling", ActiveBall, 1:vpmTimer.PulseSw 54:End Sub      '54
Sub Sp51_Spin:PlaysoundAtVol "droptarget2", sp51, 1:vpmTimer.PulseSw 51:DOF 101, 2:End Sub  '51
Sub Sp55_Spin:PlaysoundAtVol "droptarget2", sp55, 1:vpmTimer.PulseSw 55:DOF 101, 2:End Sub  '55


Dim RStep, Lstep

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSwitch (43), 0, ""
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSwitch (42), 0, ""
  PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub


Set Lights(2)=L2
Set Lights(5)=TILT
Set Lights(6)=L6
Set Lights(7)=L7
Set Lights(10)=L10
Set Lights(11)=L11
Set Lights(13)=BIP
Set Lights(14)=L14
Set Lights(15)=L15
Set Lights(17)=L17
Set Lights(18)=L18
Set Lights(19)=L19
Set Lights(20)=L20
Set Lights(21)=L21
Set Lights(22)=L22
Set Lights(23)=L23
Set Lights(25)=L25
Set Lights(26)=L26
Set Lights(27)=L27
Set Lights(29)=L29
Set Lights(30)=L30
Set Lights(31)=L31
Set Lights(33)=L33
Set Lights(34)=L34
Set Lights(35)=L35
Set Lights(37)=L37
Set Lights(38)=L38
Set Lights(39)=L39
Set Lights(41)=L41
Set Lights(42)=L42
Set Lights(43)=L43
Set Lights(45)=L45
Set Lights(46)=GO
Set Lights(47)=L47
Set Lights(49)=L49
Set Lights(50)=L50
Set Lights(51)=L51
Set Lights(52)=L52
Set Lights(53)=L53
Set Lights(54)=L54
Set Lights(55)=L55
Set Lights(57)=L57
Set Lights(58)=L58
Set Lights(59)=L59
Set Lights(61)=L61
Set Lights(62)=L62
Set Lights(63)=L63

lights(129)= array(light001,light002,light003,light004,light005)
lights(130)= array(Light006,Light007,light013,light014,light015)
lights(131)= array(Light008,Light009,Light016,Light017,Light018)
lights(132)= array(Light010,Light011,Light012,Light019,Light020)

Dim msg(38)
Dim char(9999)
char(0) = " "
char(63) = "0"
char(6) = "1"
char(91) = "2"
char(79) = "3"
char(102) = "4"
char(109) = "5"
char(124) = "6"
char(7) = "7"
char(127) = "8"
char(103) = "9"

'1 on the center line (for Gottlieb) - $CA (202)
'1 on the center with comma - $CB (203)
'6 without top line (Atari, Gottlieb) - $CC (204)
'6 without top line with comma - $CD (205)
'9 without bottom line (Atari, Gottlieb) - $CE (206)
'9 without bottom line with comma - $CF (207)
'char(768) = chr(&HCA) ' 1 on Gottlieb tables
'char(896) = chr(&HCB) ' 1, on Gottlieb tables
char(124) = chr(&HCC) ' 6 without top line
char(252) = chr(&HCD) ' 6, without top line
char(103) = chr(&HCE) ' 9 without bottom line
char(231) = chr(&HCF) ' 9, without bottom line


'Atari Middle Earth
'added by Inkochnito

Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 700,400,"Middle Earth - DIP switches"
    .AddFrame 0,5,190,"Coins per credit",&H000000C3,Array("1 coin 1 credit",0,"1 coin 2 credits",&H00000002,"1 coin 3 credits",&H00000001,"1 coin 4 credits",&H00000003,"2 coins 1 credits",&H00000080)'SW2-3&SW2-4&SW2-5&SW2-6 (dip 2&1&8&7)
    .AddFrame 0,95,190,"Maximum credits",49152,Array("5 credits",0,"10 credits",32768,"15 credits",&H00004000,"20 credits",49152)'SW1-5&SW1-6 (dip 16&15)
    .AddFrame 0,172,190,"Bonus memory feature",&H00000100,Array("normal bonus",0,"previous bonus held on last ball",&H00000100)'SW1-4 (dip 9)
    .AddFrame 0,220,190,"Balls per game",&H00000008,Array("3 balls",0,"5 balls",&H00000008)'SW2-1 (dip 4)
    .AddFrame 0,268,190,"Replay award",&H00001000,Array("extra ball",0,"replay",&H00001000)'SW1-8 (dip 13)
    .AddFrame 210,5,190,"Score threshold level",&H000F0000,Array("no replay",0,"70,000 and 100,000 points",&H00040000,"100,000 and 130,000 points",&H00070000,"140,000 and 170,000 points",&H000B0000,"220,000 and 270,000 points",&H000F0000)'Rotary (dip 17&18&19&20)
    .AddFrame 210,95,190,"Special award",&H00000030,Array("200,000 points",0,"10,000 points",&H00000010,"extra ball",&H00000020,"replay",&H00000030)'SW2-7&SW2-8 (dip 6&5)
    .AddFrame 210,172,190,"Extra ball award",&H00000400,Array("extra ball",0,"10,000 points",&H00000400)'SW1-2 (dip 11)
    .AddFrame 210,220,190,"Replay limit",&H00000200,Array("1 replay per ball",0,"2 replays per ball",&H00000200)'SW1-3 (dip 10)
    .AddFrame 210,268,190,"Score adjust",&H00002000,Array("all scores 5X",0,"normal scoring",&H00002000)'SW1-7 (dip 14)
    .AddChk 0,320,190,Array("Match feature",&H00000004)'SW2-2 (dip 3)
    .AddChk 210,320,190,Array("Diagnostic mode (must be off)",&H00000800)'SW1-1 (dip 12)
    .AddLabel 50,350,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub

Set vpmShowDips=GetRef("editDips")

 Dim Digits(27)
Digits(0)=Array(Light1,Light2,Light3,Light4,Light5,Light6,Light7)
Digits(1)=Array(Light8,Light9,Light10,Light11,Light12,Light13,Light14)
Digits(2)=Array(Light15,Light16,Light17,Light18,Light19,Light20,Light21)
Digits(3)=Array(Light22,Light23,Light24,Light25,Light26,Light27,Light28)
Digits(4)=Array(Light29,Light30,Light31,Light32,Light33,Light34,Light35)
Digits(5)=Array(Light36,Light37,Light38,Light39,Light40,Light41,Light42)
Digits(6)=Array(Light43,Light44,Light45,Light46,Light47,Light48,Light49)
Digits(7)=Array(Light50,Light51,Light52,Light53,Light54,Light55,Light56)
Digits(8)=Array(Light57,Light58,Light59,Light60,Light61,Light62,Light63)
Digits(9)=Array(Light64,Light65,Light66,Light67,Light68,Light69,Light70)
Digits(10)=Array(Light71,Light72,Light73,Light74,Light75,Light76,Light77)
Digits(11)=Array(Light78,Light79,Light80,Light81,Light82,Light83,Light84)
Digits(12)=Array(Light85,Light86,Light87,Light88,Light89,Light90,Light91)
Digits(13)=Array(Light92,Light93,Light94,Light95,Light96,Light97,Light98)
Digits(14)=Array(Light99,Light100,Light101,Light102,Light103,Light104,Light105)
Digits(15)=Array(Light106,Light107,Light108,Light109,Light110,Light111,Light112)
Digits(16)=Array(Light113,Light114,Light115,Light116,Light117,Light118,Light119)
Digits(17)=Array(Light120,Light121,Light122,Light123,Light124,Light125,Light126)
Digits(18)=Array(Light168,Light167,Light166,Light165,Light164,Light163,Light162)
Digits(19)=Array(Light161,Light160,Light159,Light158,Light157,Light156,Light155)
Digits(20)=Array(Light154,Light153,Light152,Light151,Light150,Light149,Light148)
Digits(21)=Array(Light147,Light146,Light145,Light144,Light143,Light142,Light141)
Digits(22)=Array(Light140,Light139,Light138,Light137,Light136,Light135,Light134)
Digits(23)=Array(Light133,Light132,Light131,Light130,Light129,Light128,Light127)
Digits(24)=Array(Light196,Light195,Light194,Light193,Light192,Light191,Light190)
Digits(25)=Array(Light189,Light188,Light187,Light186,Light185,Light184,Light183)
Digits(26)=Array(Light182,Light181,Light180,Light179,Light178,Light177,Light176) 'CREDITS X10
Digits(27)=Array(Light175,Light174,Light173,Light172,Light171,Light170,Light169) 'CREDITS X1

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

GIURC=1+L49.state+L50.state+L51.state+l53.state+l57.state+l61.state+l54.state+l58.state+l61.state+l62.state
lgiurc=0.1+((1.6-giurc/10)/2)
bgiurc=lgiurc*3+5
lbgiurc=lgiurc*8

FGO.Visible=GO.state
FBall.Visible=BIP.state
FMatch.Visible=1- BIP.state
FTilt.Visible=TILT.state
If light176.state=1 and light169.state=1 or light176.state=0 and light177.state=1 and light178.state=1 and light179.state=1 and light180.state=1 and light181.state=1 and light182.state=1 and light169.state=0 and light170.state=1 and light171.state=1 and light172.state=1 and light173.state=1 and light174.state=1 and light175.state=1 then
fcredit.visible=0
else
fcredit.visible=1
End if

FlashL13.intensity=lGIURC
FlashL14.intensity=lGIURC
FlashL15.intensity=lGIURC
FlashL5.intensity=lGIURC
FlashL6.intensity=lGIURC
FlashL003.intensity=lGIURC
FlashL10.intensity=lGIURC
FlashL11.intensity=lGIURC
FlashL12.intensity=lGIURC
FlashL004.intensity=lGIURC
FlashL007.intensity=lGIURC
FlashL008.intensity=lGIURC
FlashL001.intensity=lGIURC
FlashL002.intensity=lGIURC
Light045.intensity=lbgiurc
Light021.intensity=lbgiurc
Light022.intensity=lbgiurc
Light023.intensity=lbgiurc
light024.intensity=lbgiurc
Light026.intensity=lbgiurc
Light027.intensity=lbgiurc
Light028.intensity=lbgiurc
Light030.intensity=lbgiurc
Light031.intensity=lbgiurc
Light032.intensity=lbgiurc
Light033.intensity=lbgiurc
Light034.intensity=lbgiurc
Light036.intensity=lbgiurc
Light041.intensity=lbgiurc
Light042.intensity=lbgiurc
Light045.intensity=bgiurc
P1.visible=l35.state
P2.visible=L39.state
P3.visible=L43.state
P4.visible=L47.state
End Sub

Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

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

Const tnob = 2 ' total number of balls
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*10, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

Sub Rubbers_Hit(idx)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub Metals_Hit(idx)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "metal_hit1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "metal_hit2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "metal_hit3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub Flipper_hit(idx)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub rollingplastic_Hit (idx)
  PlaySound "fx_plasticrolling0", 0, VolMulti(ActiveBall,VolPlast), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Gate_hit()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "metal_hit1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "metal_hit2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "metal_hit3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub
