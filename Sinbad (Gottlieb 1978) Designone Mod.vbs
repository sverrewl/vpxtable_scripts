On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName="sinbad",UseSolenoids=1,UseLamps=0,UseGI=0,UseSync=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown"
Const SCoin="fx_coin3",cCredits=""

LoadVPM"01150000","GTS1.VBS",3.22

DesktopMode = Table1.ShowDT

Dim VarHidden

If Table1.ShowDT = true then 'desktop view
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
else 'FS view
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
    lrail.Visible = 0
    rrail.Visible = 0
end if

if B2SOn = true then VarHidden = 1


Dim counter
Dim bump1, bump2, x
Dim DesktopMode

Const WYTar=6
Const PTar=7
Const RTar=8

Sub WYRaised(enabled)
  debug.print "WYRaised"
  If enabled Then
    WYReset.interval=500
    WYReset.enabled=True
  End If
End Sub

Sub WYReset_Timer()
  debug.print "Raise White and Yellow Targets"
  GIPL37.State=0
  WYReset.enabled=False
  dtDropW.DropSol_On
End Sub

Sub PRaised(enabled)
  debug.print "PRaised"
  If enabled Then
    PReset.interval=500
    PReset.enabled=True
  End If
End Sub

Sub PReset_Timer()
  debug.print "Raise Purple Targets"
  GIPL38.State=0:GIPL39.State=0:GIPL40.State=0
  PReset.enabled=False
  dtDropP.DropSol_On
End Sub

Sub RRaised(enabled)
  debug.print "RRaised"
  If enabled Then
    RReset.interval=500
    RReset.enabled=True
  End If
End Sub

Sub RReset_Timer()
  debug.print "Raise Red Targets"
  GIPL30.State=0:GIPL31.State=0:GIPL32.State=0:GIPL33.State=0
  RReset.enabled=False
  dtDropR.DropSol_On
End Sub

'*********
'   LUT
'*********

Dim bLutActive, LUTImage
Sub LoadLUT
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "")Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage + 1)MOD 10:UpdateLUT:SaveLUT:End Sub

Sub UpdateLUT
    Select Case LutImage
       Case 0:table1.ColorGradeImage = "LUT0"
        Case 1:table1.ColorGradeImage = "LUT1"
        Case 2:table1.ColorGradeImage = "LUT2"
        Case 3:table1.ColorGradeImage = "LUT3"
        Case 4:table1.ColorGradeImage = "LUT4"
        Case 5:table1.ColorGradeImage = "LUT5"
        Case 6:table1.ColorGradeImage = "LUT6"
        Case 7:table1.ColorGradeImage = "LUT7"
        Case 8:table1.ColorGradeImage = "LUT8"
        Case 9:table1.ColorGradeImage = "LUT9"
        Case 10:table1.ColorGradeImage = "LUT10"
    End Select
End Sub

SolCallback(1)="bsTrough.SolOut"
SolCallback(2)="VpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(3)="vpmSolSound SoundFX(""10pts"",DOFChimes),"
SolCallback(4)="vpmSolSound SoundFX(""100pts"",DOFChimes),"
SolCallback(5)="vpmSolSound SoundFX(""1000pts"",DOFChimes),"
SolCallback(sLLFlipper)="SolLFlipper"
SolCallback(sLRFlipper)="SolRFlipper"
SolCallback(WYTar)="WYRaised" 'dtDropW.SolDropUp
SolCallBack(PTar)="PRaised"   'dtDropP.SolDropUp
SolCallback(RTar)="RRaised"   'dtDropR.SolDropUp


Sub SolLFlipper(Enabled)
   If Enabled Then
      PlaySound "flipperup", DOFFlippers
      LeftFlipper.RotateToEnd
    LeftFlipper1.RotateToEnd
    Else
      PlaySound "flipperdown", DOFFlippers
      LeftFlipper.RotateToStart
    LeftFlipper1.RotateToStart
   End If
End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
        PlaySound "Flipperup",DOFFlippers
    RightFlipper.RotateToEnd
    RightFlipper1.RotateToEnd
Else
        PlaySound "Flipperdown",DOFFlippers
    RightFlipper.RotateToStart
    RightFlipper1.RotateToStart
     End If
End Sub

Sub flippertimer_timer()
dim PI:PI=3.1415926
  SpinnerP.Rotz = Spinner1.CurrentAngle
  SpinnerRod.TransZ = sin( (Spinner1.CurrentAngle+180) * (2*PI/360)) * 5
  SpinnerRod.TransX = -1*(sin( (Spinner1.CurrentAngle- 90) * (2*PI/360)) * 5)


  LFlip.RotZ = LeftFlipper.CurrentAngle +240
  RFlip.RotZ = RightFlipper.CurrentAngle +120
  Lflip1.RotZ = LeftFlipper1.CurrentAngle +240
  RFlip1.RotZ = RightFlipper1.CurrentAngle +120

  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  FlipperLSh1.RotZ = LeftFlipper1.currentangle
  FlipperRSh1.RotZ = RightFlipper1.currentangle
end sub


Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub


Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = LeftMagnaSave Then bLutActive = True
    If keycode = RightMagnaSave Then
        If bLutActive Then NextLUT:End If
    End If
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = LeftMagnaSave Then bLutActive = False
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

Dim rdwastep, rdwbstep, rdwcstep, rdwdstep, rdwestep
Dim bsTrough,dtDropW,dtDropP,dtDropR,bsSaucer, xx
Dim BPG, hsaward



LS74d.State = 1   'lightstate.state --> 0 = off, 1 = on
LS74c.State = 1
LS74b.State = 1
LS74a.State = 1



Sub Table1_Init
Controller.Games(cGameName).Settings.Value("dmd_red")=0 '** change the dmd colour to dark blue
Controller.Games(cGameName).Settings.Value("dmd_green")=128
Controller.Games(cGameName).Settings.Value("dmd_blue")=255
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Sinbad (Gottlieb 1978)"&chr(13)&""
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .Hidden = VarHidden
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
    Controller.SolMask(0)=0
  vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=4
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2)

     vpmMapLights aGILights
       ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"
    GIUpdateTimer.Enabled = 1
    LoadLUT
    FindDips    'find balls per game and high score reward
    vpmMapLights AllLights
   End Sub

'************TROUGH************************
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 66, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd  SoundFX("fx_ballrelease",DOFContactors) ,SoundFX("fx_ballrelease",DOFContactors)
        .Balls = 1
    End With

'**************SPINNER************

Sub Spinner1_Spin():vpmTimer.PulseSw 10:PlaySoundAt "fx_spinner", Spinner1 :End Sub

'*******GATES ******************

Sub Gate_Hit(): PlaySoundAt "fx_Gate", ActiveBall : End Sub

'********************DROP TARGETS******************

Set dtDropW=New cvpmDropTarget'
  dtDropW.InitDrop Array(TW1,TY1,TY2),Array(20,21,24)
  dtDropW.InitSnd SoundFX("Drop",DOFContactors),SoundFX("reset",DOFContactors)

  Set dtDropP=New cvpmDropTarget'
  dtDropP.InitDrop Array(TP1,TP2,TP3),Array(30,31,34)
  dtDropP.InitSnd SoundFX("Drop",DOFContactors),SoundFX("reset",DOFContactors)

  Set dtDropR=New cvpmDropTarget'
  dtDropR.InitDrop Array(TR1,TR2,TR3,TR4),Array(50,51,60,61)
  dtDropR.InitSnd SoundFX("Drop",DOFContactors),SoundFX("reset",DOFContactors)


' Drain &

Sub Drain_Hit:bsTrough.AddBall Me:PlaysoundAt "Drain", ActiveBall:End Sub

'Pop Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 14 : playsoundAt SoundFX("fx_bumper",DOFContactors), ActiveBall
: B1L1.State = 1:B1L2. State = 1 : Me.TimerEnabled = 1 : End Sub
Sub Bumper1_Timer : B1L1.State = 0:B1L2. State = 0 : Me.Timerenabled = 0 : End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 13 : playsoundAt SoundFX("fx_bumper",DOFContactors), ActiveBall
: B2L1.State = 1:B2L2. State = 1 : Me.TimerEnabled = 1 : End Sub
Sub Bumper2_Timer : B2L1.State = 0:B2L2. State = 0 : Me.Timerenabled = 0 : End Sub

'**top rollover switches

Sub SW70_Hit():Controller.Switch(70)=1:End Sub
Sub SW70_UnHit():Controller.Switch(70)=0:End Sub
Sub SW71_Hit():Controller.Switch(71)=1:End Sub
Sub SW71_UnHit():Controller.Switch(71)=0:End Sub
Sub SW40_Hit():Controller.Switch(40)=1:DOF 101, DOFOn:End Sub
Sub SW40_UnHit():Controller.Switch(40)=0:DOF 101, DOFOff:End Sub
Sub SW41_Hit():Controller.Switch(41)=1:DOF 105, DOFOn:End Sub
Sub SW41_UnHit():Controller.Switch(41)=0:DOF 105, DOFOff:End Sub

'**left green star rollover switches
Sub SW74d_Hit():LS74d.State = 0:Controller.Switch(74)=1:DOF 111, DOFOn:End Sub
Sub SW74d_UnHit():LS74d.State = 1:Controller.Switch(74)=0:DOF 111, DOFOff:End Sub
Sub SW74c_Hit():LS74c.State = 0:Controller.Switch(74)=1:DOF 110, DOFOn:End Sub
Sub SW74c_UnHit():LS74c.State = 1:Controller.Switch(74)=0:DOF 110, DOFOff:End Sub
Sub SW74b_Hit():LS74b.State = 0:Controller.Switch(74)=1:DOF 109, DOFOn:End Sub
Sub SW74b_UnHit():LS74b.State = 1:Controller.Switch(74)=0:DOF 109, DOFOff:End Sub
Sub SW74a_Hit():LS74a.State = 0:Controller.Switch(74)=1:DOF 108, DOFOn:End Sub
Sub SW74a_UnHit():LS74a.State = 1:Controller.Switch(74)=0:DOF 108, DOFOff:End Sub

'****************Target********************

Sub TargetSW41a_hit:vpmTimer.PulseSw(41):PlaysoundAt SoundFXDOF("target",106,DOFPulse,DOFContactors),ActiveBall:GIPL34.State=0:TargetSW41a.TimerEnabled=True:End Sub
Sub TargetSW41a_Timer:TargetSW41a.TimerEnabled=False:GIPL34.State=1:End Sub
Sub TargetSW41c_hit:vpmTimer.PulseSw(41):PlaysoundAt SoundFXDOF("target",107,DOFPulse,DOFContactors),ActiveBall:End Sub

'**right lane switch
Sub SW44_Hit():Controller.Switch(44)=1:End Sub
Sub SW44_UnHit():Controller.Switch(44)=0:Playsound "":End Sub

Sub LeftSlingshot_slingshot():vpmtimer.pulsesw 60:PlaysoundAt SoundFX("slingshot",DOFContactors),ActiveBall:End Sub
Sub RightSlingshot_Slingshot():vpmtimer.pulsesw 60:PlaysoundAt SoundFX("slingshot",DOFContactors),ActiveBall:End Sub
Sub TargetExtraBall_Hit():vpmtimer.pulsesw 64:playsoundAt SoundFX("spothit",DOFContactors),ActiveBall:TargetExtraBall.Isdropped=1:TargetExtraBall2.Isdropped=0:TargetExtraBall.Timerenabled=1:End Sub
Sub TargetExtraBall_Timer:TargetExtraBall.TimerEnabled=0:TargetExtraBall.isdropped=0:TargetExtraBall2.isdropped=1:End Sub

'**you hit the left or right outlane, bad luck, try again!
Sub SW40b_Hit():Controller.Switch(40)=1:DOF 102, DOFOn:End Sub
Sub SW40b_Unhit():Controller.Switch(40)=0:DOF 102, DOFOff:End Sub
Sub SW44b_Hit():Controller.Switch(44)=1:End Sub
Sub SW44b_Unhit():Controller.Switch(44)=0:End Sub

'**drop target bank
Sub TR1_Hit():dtDropR.Hit 1:GIPL30.State=1:End Sub
Sub TR2_Hit():dtDropR.Hit 2:GIPL31.State=1:End Sub
Sub TR3_Hit():dtDropR.Hit 3:GIPL32.State=1:End Sub
Sub TR4_Hit():dtDropR.HIt 4:GIPL33.State=1:End Sub

'**white & yellow drop target bank
Sub TW1_Hit():dtDropW.Hit 1:GIPL41.State=1:End Sub
Sub TY1_Hit():dtDropW.Hit 2:GIPL42.State=1:End Sub
Sub TY2_Hit():dtDropW.Hit 3:GIPL37.State=1:End Sub

'**purple drop target bank
Sub TP1_Hit():dtDropP.Hit 1:GIPL38.State=1:End Sub
Sub TP2_Hit():dtDropP.Hit 2:GIPL39.State=1:End Sub
Sub TP3_Hit():dtDropP.Hit 3:GIPL40.State=1:End Sub

'**triggers for leaf switches linked to switch 54
Sub SW54a_Hit():vpmTimer.PulseSw(54): PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub SW54b_Hit():vpmTimer.PulseSw(54): PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub SW54c_Hit():vpmTimer.PulseSw(54): PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub SW54d_Hit():vpmTimer.PulseSw(54): PlaySoundAt "fx_sensor",ActiveBall:End Sub
'Sub SW54d_Timer:Sw54d.TimerEnabled=False:End Sub
Sub SW54e_Hit():vpmTimer.PulseSw(54):End Sub
Sub SW54f_Hit():vpmTimer.PulseSw(54):End Sub
Sub SW54g_Hit():vpmTimer.PulseSw(54):End Sub
Sub SW54h_Hit():vpmTimer.PulseSw(54):End Sub
Sub SW54m_Hit():vpmTimer.PulseSw(54):End Sub

Dim LampState(200)


LampTimer.Interval = 40
LampTimer.Enabled = 1

Sub LampTimer_Timer()
  Dim chgLamp, num, chg, ii
  chgLamp = Controller.ChangedLamps
  If Not IsEmpty(chgLamp) Then
    For ii = 0 To UBound(chgLamp)
      LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
    Next
  End If

  UpdateLamps

End Sub

Sub UpdateLamps

NFadeL 4, L4
NFadeL 5, L5
NFadeL 6, L6
NFadeL 7, L7
NFadeL 8, L8
NFadeL 9, L9
NFadeL 10, L10
NFadeL 11, L11
'NFadeL 12, L12
'NFadeL 13, L13
NFadeL 14, L14
NFadeL 15, L15
NFadeL 16, L16
NFadeL 17, L17
NFadeL 18, L18
NFadeL 19, L19
NFadeL 20, L20
NFadeL 21, L21
NFadeL 22, L22
NFadeL 23, L23
NFadeL 24, L24
NFadeL 25, L25
NFadeL 26, L26
NFadeL 27, L27
NFadeL 28, L28
NFadeL 29, L29
NFadeL 30, L30
NFadeL 31, L31
NFadeL 32, L32
NFadeL 33, L33
NFadeL 34, L34

End SUB

Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub

Sub SetLamp(nr, value):LampState(nr) = abs(value) + 4:End Sub

Sub NFadeL(nr, a)
  Select Case LampState(nr)
    Case 4:a.state = 0:LampState(nr) = 0
    Case 5:a.State = 1:LampState(nr) = 1
  End Select
End Sub


'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'      Gottlieb led Patterns
'************************************

Dim Digits(28)
Dim Patterns(11) 'normal numbers
Dim Patterns2(11) 'numbers with a comma
Dim Patterns3(11) 'the credits and ball in play

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 768   '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 124   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 103  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 896  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 231 '9

Patterns3(0) = 0     'empty
Patterns3(1) = 63    '0
Patterns3(2) = 6     '1
Patterns3(3) = 91    '2
Patterns3(4) = 79    '3
Patterns3(5) = 102   '4
Patterns3(6) = 109   '5
Patterns3(7) = 124   '6
Patterns3(8) = 7     '7
Patterns3(9) = 127   '8
Patterns3(10) = 103  '9


Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5

Set Digits(6) = b0
Set Digits(7) = b1
Set Digits(8) = b2
Set Digits(9) = b3
Set Digits(10) = b4
Set Digits(11) = b5


Set Digits(12) = c0
Set Digits(13) = c1
Set Digits(14) = c2
Set Digits(15) = c3
Set Digits(16) = c4
Set Digits(17) = c5


Set Digits(18) = d0
Set Digits(19) = d1
Set Digits(20) = d2
Set Digits(21) = d3
Set Digits(22) = d4
Set Digits(23) = d5


Set Digits(24) = e2
Set Digits(25) = e3

Set Digits(26) = e0
Set Digits(27) = e1


Sub DisplayTimer_Timer
 On Error Resume Next
    dim oldstat
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
if oldstat <> STAT then
debug.print STAT
oldstat = stat
end if
                If(stat = Patterns(jj) ) OR (stat = Patterns2(jj) ) OR (stat = Patterns3(jj) ) then Digits(chgLED(ii, 0) ).SetValue jj
            Next
        Next
    End IF
End Sub

Set MotorCallback=GetRef("UpdateMultipleLamps")

Dim  N1

Sub UpdateMultipleLamps

 N1=Controller.Lamp(1)
  If N1 Then
    GOBox.text= ""
        BIPBox.text="Ball in Play"
    Else
        GOBox.text="Game Over"
        BIPBox.text=""

  End if

N1=Controller.Lamp(3)
  If N1 Then
    HStoDateBox.text="High Score"
    Else
    HStoDateBox.text=""
  End if

N1=Controller.Lamp(2) '
  If N1 Then
    TILTBox.text="TILT"
    Else
    TILTBox.text=""
  End If

N1=Controller.Lamp(4)
  If N1 Then
    SABox.text="Shoot Again"
    Else
    SABox.text=""
  End If

End Sub

'**dipswitch menu - press F6 to display it

Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"System 1 - DIP switches"
    .AddFrame 0,0,190,"Coin chutecontrol",&H00040000,Array("seperate",0,"same",&H00040000)'dip 19
    .AddFrame 0,46,190,"Game mode",&H00000400,Array("extraball",0,"replay",&H00000400)'dip 11
    .AddFrame 0,92,190,"High game to date awards",&H00200000,Array("noaward",0,"5 replays",&H00200000)'dip 22
    .AddFrame 0,138,190,"Balls per game",&H00000100,Array("5 balls",0,"5balls",&H00000100)'dip 9
    .AddFrame 0,184,190,"Tilt effect",&H00000800,Array("game over",0,"ball in play only",&H00000800)'dip 12
    .AddFrame 205,0,190,"Maximum credits",&H00030000,Array("5 credits",0,"8 credits",&H00020000,"10 credits",&H00010000,"15 credits",&H00030000)'dip 17&18
    .AddFrame 205,76,190,"Soundsettings",&H80000000,Array("sounds",0,"tones",&H80000000)'dip 32
    .AddFrame 205,122,190,"Attract tune",&H10000000,Array("no attract tune",0,"attract tune played every 6 minutes",&H10000000)'dip 29
    .AddChk 205,175,190,Array("Match feature",&H00000200)'dip 10
    .AddChk 205,190,190,Array("Credits displayed",&H00001000)'dip 13
    .AddChk 205,205,190,Array("Play credit button tune",&H00002000)'dip 14
    .AddChk 205,220,190,Array("Play tones when scoring",&H00080000)'dip 20
    .AddChk 205,235,190,Array("Play coin switch tune",&H00400000)'dip 23
    .AddChk 205,250,190,Array("High game to date displayed",&H00100000)'dip 21
    .AddLabel 50,280,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips = GetRef("editDips")

'*******************************************
'DIPS code
'*******************************************
'Finding an individual dip state based on scapino's Strikes and spares dip code - from unclewillys pinball pool, added another section to get high score award to set replay cards

Dim TheDips(32)
Sub FindDips
  Dim DipsNumber
  DipsNumber = Controller.Dip(1)
  TheDips(16) = Int(DipsNumber/128)
  If TheDips(16) = 1 then DipsNumber = DipsNumber - 128 end if
  TheDips(15) = Int(DipsNumber/64)
  If TheDips(15) = 1 then DipsNumber = DipsNumber - 64 end if
  TheDips(14) = Int(DipsNumber/32)
  If TheDips(14) = 1 then DipsNumber = DipsNumber - 32 end if
  TheDips(13) = Int(DipsNumber/16)
  If TheDips(13) = 1 then DipsNumber = DipsNumber - 16 end if
  TheDips(12) = Int(DipsNumber/8)
  If TheDips(12) = 1 then DipsNumber = DipsNumber - 8 end if
  TheDips(11) = Int(DipsNumber/4)
  If TheDips(11) = 1 then DipsNumber = DipsNumber - 4 end if
  TheDips(10) = Int(DipsNumber/2)
  If TheDips(10) = 1 then DipsNumber = DipsNumber - 2 end if
  TheDips(9) = Int(DipsNumber)
  DipsNumber = Controller.Dip(2)
  TheDips(24) = Int(DipsNumber/128)
  If TheDips(24) = 1 then DipsNumber = DipsNumber - 128 end if
  TheDips(23) = Int(DipsNumber/64)
  If TheDips(23) = 1 then DipsNumber = DipsNumber - 64 end if
  TheDips(22) = Int(DipsNumber/32)
  If TheDips(22) = 1 then DipsNumber = DipsNumber - 32 end if
  TheDips(21) = Int(DipsNumber/16)
  If TheDips(21) = 1 then DipsNumber = DipsNumber - 16 end if
  TheDips(20) = Int(DipsNumber/8)
  If TheDips(20) = 1 then DipsNumber = DipsNumber - 8 end if
  TheDips(19) = Int(DipsNumber/4)
  If TheDips(19) = 1 then DipsNumber = DipsNumber - 4 end if
  TheDips(18) = Int(DipsNumber/2)
  If TheDips(18) = 1 then DipsNumber = DipsNumber - 2 end if
  TheDips(17) = Int(DipsNumber)
  DipsNumber = Controller.Dip(3)
  TheDips(26) = Int(DipsNumber/128)
  If TheDips(26) = 1 then DipsNumber = DipsNumber - 2 end if
  DipsTimer.Enabled=1
End Sub

 Sub DipsTimer_Timer()
  hsaward = TheDips(22)
  BPG = TheDips(9)
  dim ebplay: ebplay= TheDips(11)
  If BPG = 1 then
    if ebplay = 1 then
      instcard.image="InstCard3Balls"
      else
      instcard.image="InstCard3BallsEB"
    end if
    Else
    if ebplay = 1 then
      instcard.image="InstCard5Balls"
      else
      instcard.image="InstCard5BallsEB"
    end if
  End if
  repcard.image="replaycard"&hsaward
  DipsTimer.enabled=0
 End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_metalhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aRubber_Pegs_Hit(idx):PlaySound "fx_rubber_peg", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub

'*********************************************************
' Real Time Flipper adjustments - by JLouLouLou & JPSalas
'        (to enable flipper tricks)
'*********************************************************

Dim FlipperPower
Dim FlipperElasticity
Dim SOSTorque, SOSAngle
Dim FullStrokeEOS_Torque, LiveStrokeEOS_Torque
Dim LeftFlipperOn
Dim RightFlipperOn

Dim LLiveCatchTimer
Dim RLiveCatchTimer
Dim LiveCatchSensivity


FlipperPower = 5000
FlipperElasticity = 0.85
FullStrokeEOS_Torque = 0.3  ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.2  ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10

SOSTorque = 0.1
SOSAngle = 6

LiveCatchSensivity = 10

LLiveCatchTimer = 0
RLiveCatchTimer = 0

LeftFlipper.TimerInterval = 1
LeftFlipper.TimerEnabled = 1

Sub LeftFlipper_Timer 'flipper's tricks timer
'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If LeftFlipper.CurrentAngle >= LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If LeftFlipperOn = 1 Then
    If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle then
      LeftFlipper.EOSTorque = FullStrokeEOS_Torque
      LLiveCatchTimer = LLiveCatchTimer + 1
      If LLiveCatchTimer < LiveCatchSensivity Then
        LeftFlipper.Elasticity = 0
      Else
        LeftFlipper.Elasticity = FlipperElasticity
        LLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    LeftFlipper.Elasticity = FlipperElasticity
    LeftFlipper.EOSTorque = LiveStrokeEOS_Torque
    LLiveCatchTimer = 0
  End If


'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If RightFlipperOn = 1 Then
    If RightFlipper.CurrentAngle = RightFlipper.EndAngle Then
      RightFlipper.EOSTorque = FullStrokeEOS_Torque
      RLiveCatchTimer = RLiveCatchTimer + 1
      If RLiveCatchTimer < LiveCatchSensivity Then
        RightFlipper.Elasticity = 0
      Else
        RightFlipper.Elasticity = FlipperElasticity
        RLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    RightFlipper.Elasticity = FlipperElasticity
    RightFlipper.EOSTorque = LiveStrokeEOS_Torque
    RLiveCatchTimer = 0
  End If
End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 200
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp> 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.1, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'****************************************************
'   JP's VPX Rolling Sounds with ball speed control
'****************************************************

Const tnob = 2    'total number of balls
Const lob = 0     'number of locked balls
Const maxvel = 30 'max ball velocity 25 -50
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
    RollingTimer.Enabled = 1
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <0 Then
                ballpitch = Pitch(BOT(b) ) - 5000 'decrease the pitch under the playfield
                ballvol = Vol(BOT(b) )
            ElseIf BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 25000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 3
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' dropping sounds
        If BOT(b).VelZ <-1 Then
            'from ramp
            If BOT(b).z <55 and BOT(b).z> 27 Then PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
            'down a hole
            If BOT(b).z <10 and BOT(b).z> -10 Then PlaySound "fx_hole_enter", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        End If

        ' jps ball speed & spin control
        BOT(b).AngMomZ = BOT(b).AngMomZ * 0.95
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactory
                BOT(b).VelY = BOT(b).VelY * speedfactory
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

'***********
' GI lights
'***********

Dim oldGiState
oldGiState = 0

Sub GiOn 'enciEnde las luces GI
    Dim bulb
    PlaySound"fx_gion"
    For each bulb in aGILights
        bulb.State = 1
    Next
End Sub

Sub GiOff 'apaga las luces GI
    Dim bulb
    PlaySound"fx_gioff"
    For each bulb in aGILights
        bulb.State = 0
    Next
End Sub

Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then '-1 means no balls, 0 is the first captive ball, 1 is the second captive ball...)
            GiOff               ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
        Else
            Gion
        End If
    End If
End Sub

'******************Dingwall ************************************

sub DingwallA_hit
  vpmTimer.PulseSw 54
  PlaySoundat "target", ActiveBall
  rdwa.visible=0
  RDWA1.visible=1
  rdwAstep=1
  DingwallA.timerenabled=1

end sub


sub DingwallA_timer
  select case rdwAstep
    Case 1: RDWA1.visible=0: rdwa.visible=1
    case 2: rdwa.visible=0: rdwa2.visible=1
    Case 3: rdwa2.visible=0: rdwa.visible=1: DingwallA.timerenabled=0
  end Select
  rdwAstep=rdwAstep+1
end sub

sub DingwallB_hit
    vpmTimer.PulseSw 54
  PlaySoundat "target", ActiveBall
  rdwB.visible=0
  RDWB1.visible=1
  rdwbstep=1
  DingwallB.timerenabled=1

end sub

sub DingwallB_timer
  select case rdwbstep
    Case 1: RDWb1.visible=0: rdwb.visible=1
    case 2: rdwb.visible=0: rdwb2.visible=1
    Case 3: rdwb2.visible=0: rdwb.visible=1: DingwallB.timerenabled=0
  end Select
  rdwbstep=rdwbstep+1
end sub

sub DingwallC_hit
    vpmTimer.PulseSw 54
  PlaySoundat "target", ActiveBall
  rdwC.visible=0
  RDWc1.visible=1
  rdwCstep=1
  DingwallC.timerenabled=1

end sub

sub DingwallC_timer
  select case rdwCstep
    Case 1: RDWC1.visible=0: rdwc.visible=1
    case 2: rdwC.visible=0: rdwc2.visible=1
    Case 3: rdwc2.visible=0: rdwc.visible=1: DingwallC.timerenabled=0
  end Select
  rdwCstep=rdwCstep+1
end sub


sub DingwallD_hit
    vpmTimer.PulseSw 54
  PlaySoundat "target", ActiveBall
  rdwd.visible=0
  RDWd1.visible=1
  rdwDstep=1
  DingwallD.timerenabled=1

end sub


sub DingwallD_timer
  select case rdwDstep
    Case 1: RDWd1.visible=0: rdwd.visible=1
    case 2: rdwd.visible=0: rdwd2.visible=1
    Case 3: rdwd2.visible=0: rdwd.visible=1: DingwallD.timerenabled=0
  end Select
  rdwDstep=rdwDstep+1
end sub

sub DingwallE_hit
    vpmTimer.PulseSw 54
  PlaySoundat "target", ActiveBall
  rdwE.visible=0
  RDWE1.visible=1
  rdwEstep=1
  DingwallE.timerenabled=1
end sub

sub DingwallE_timer
  select case rdwEstep
    Case 1: RDWE1.visible=0: rdwe.visible=1
    case 2: rdwE.visible=0: rdwe2.visible=1
    Case 3: rdwE2.visible=0: rdwe.visible=1: DingwallE.timerenabled=0
  end Select
  rdwEstep=rdwEstep+1
end sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub

