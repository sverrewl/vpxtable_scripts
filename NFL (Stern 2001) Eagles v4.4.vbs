Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="nfl_phi",UseSolenoids=1,UseLamps=0,UseGI=0,UseSync=1, SCoin="coin"

LoadVPM "01520000","Sega.VBS",3.02

'************
'Team Colors DMD with Original Plastics
'  0= Regular Orange  1= Team Colors DMD
CDMDWOP=1 'Team Colors DMD with Original Plastics ON by Default
'************
TPlastics=1  'Change to =0 for Non-Transparent Plastics
'************
SoundMod=0   'Change to =0 for Regular Rom sounds
'************
PupM=0       'Change to =1 for use with PupPacks that have music.

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)= "SolTrough"
SolCallback(2)= "AutoPlunger"
SolCallback(3)= "VukLEFTPop"
SolCallback(4)= "uVUK"

SolCallback(6)= "dtDrop.SolDropUp"
SolCallback(7)= "dtDrop.SolHit 1,"
SolCallback(8)= "StLock"

SolCallback(12)= "Magnet1.MagnetOn="
'SolCallback(13)= "Magnet2.MagnetOn="

SolCallback(19)= "SolBallDeflector"
SolCallback(20)= "SetLamp 120," 'Stadium X4
SolCallback(21)= "dtDrop.SolHit 2,"
SolCallback(22)= "dtDrop.SolHit 3,"
SolCallback(23)= "dtDrop.SolHit 4,"

'SolCallback(25)= "" 'Goalie
'SolCallback(26)= "" 'Goalie
SolCallBack(27)= "FlashYellow" 'Upper Flipper X1
SolCallBack(28)= "FlashRed" 'Spinner X2
SolCallBack(29)= "FlashWhite" 'Rampp X1
SolCallBack(30)= "SetLamp 130," 'Back Panel X4
SolCallback(31)= "SetLamp 131," 'Pops X4
SolCallback(32)= "SetLamp 132," 'Slingshots X4

'Aux board UK only
'SolCallback(33)= "SolLeftPost"
'SolCallback(34)= "SolMidPost"
'SolCallback(35)= "SolRightPost"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd:a01.enabled=False
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, 1:RightFlipper.RotateToEnd:a01.enabled=False
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, 1:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 15
  End If
End Sub

Sub AutoPlunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub StLock(enabled)
  if Enabled Then
    Stadiumlock.isDropped = 1
        PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), Stadiumlock, 1
  Else
    Stadiumlock.isDropped = 0
  end If
End Sub

Sub SolBallDeflector(Enabled)
  If Enabled Then
    Deflector.IsDropped=0
        PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), Deflector, 1
  Else
    Deflector.IsDropped=1
  End If
End Sub

set GICallback = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
    SetLamp 134, Enabled  'Backwall bulbs
  If Enabled Then
    dim xx
    For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"
    DOF 101, DOFOn
  Else For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"
    DOF 101, DOFOff
  End If
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, dtDrop, mGoalie, Magnet1, Magnet2,VLLock

Sub Table1_Init
  vpmInit Me:LoadPLA04:LoadBWL04
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "NFL (Stern 2001) Eagles v4.4"&chr(13)&"Team Mod by Xenonph"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 0
         If CDMD=0 Then
         Controller.Games("nfl").Settings.Value("DMD_Colorize") =0
         End If
         If CDMD=1 Then
         Controller.Games("nfl").Settings.Value("DMD_Colorize") =1
                   .GameName = "nfl"
         NVOffset (4)
                   .GameName = "nfl_phi"
         Controller.Games("nfl").Settings.Value("dmd_red") = 255
         Controller.Games("nfl").Settings.Value("dmd_green") = 255
         Controller.Games("nfl").Settings.Value("dmd_blue") = 255
         Controller.Games("nfl").Settings.Value("dmd_red66") = 186
         Controller.Games("nfl").Settings.Value("dmd_green66") = 196
         Controller.Games("nfl").Settings.Value("dmd_blue66") = 201
         Controller.Games("nfl").Settings.Value("dmd_red33") = 0
         Controller.Games("nfl").Settings.Value("dmd_green33") = 76
         Controller.Games("nfl").Settings.Value("dmd_blue33") = 84
         Controller.Games("nfl").Settings.Value("dmd_red0") = 0
         Controller.Games("nfl").Settings.Value("dmd_green0") = 0
         Controller.Games("nfl").Settings.Value("dmd_blue0") = 0
         End If
         If SoundMod=1 Then
         Controller.Games("nfl").Settings.Value("sound_mode") =1
         End If
         If PupM=1 Then
         Controller.Games("nfl").Settings.Value("sound_mode") =1
         End If
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1
    If SoundMod=1 And PupM=0 Then a02.enabled=true:a02.interval=8000:End If
  vpmNudge.TiltSwitch=56:
  vpmNudge.Sensitivity=2
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 0,14,13,12,11,0,0,0
    bsTrough.InitKick BallRelease,95,4
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls=4

  Set dtDrop=New cvpmDropTarget
    dtDrop.InitDrop Array(sw20,sw19,sw18,sw17),Array(20,19,18,17)
    dtDrop.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

    Set Magnet1=New cvpmMagnet
  with Magnet1
    .InitMagnet Trigger8,60
    .CreateEvents "Magnet1"
    .Grabcenter = True
  end With

    Set Magnet2=New cvpmMagnet
  with magnet2
    .InitMagnet RampMag,50
    .Solenoid=13
    .CreateEvents "Magnet2"
    .GrabCenter = 0
  End With

  Set mGoalie=New cvpmMech
    mGoalie.MType=vpmMechOneDirSol+vpmMechReverse+vpmMechLinear
    mGoalie.Sol1=25
    mGoalie.Sol2=26
    mGoalie.Length=20
    mGoalie.Steps=13
    mGoalie.AddSw 41,0,0
    mGoalie.AddSw 47,3,4
    mGoalie.AddSw 42,13,13
    mGoalie.Callback=GetRef("UpdateGoalie")
    mGoalie.Start

  For X=0 To 12:LBPlace(X).IsDropped=1:Next

  Deflector.IsDropped=1
LoadLUT04
LoadLBB04
LoadPT04
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  'If KeyCode=LeftFlipper Then Controller.Switch(1)=1  'UK only
  'If KeyCode=RightFlipper Then Controller.Switch(8)=1  'UK only
  If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAtVol"plungerpull", Plunger, 1
  If keycode = 2 And BIP = 0 Then AM=1
  If keycode = 3 And BIP = 0 Then NextLut04
  If keycode = 6 And SoundMod=1 Then c01.enabled=True
  If keycode = 17 Then NextLBB04
    If keycode = 38 And BIP=0 Then NextPT04
  If keycode = LeftMagnaSave And BIP=0 Then NextPLA04
  If keycode = RightMagnaSave And BIP=0 Then NextBWL04: End If
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  'If KeyCode=LeftFlipper Then Controller.Switch(1)=0
  'If KeyCode=RightFlipper Then Controller.Switch(8)=0
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

  Dim plungerIM
    Const IMPowerSetting = 135 'Plunger Power
    Const IMTime = 1.1       ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 16
        .InitExitSnd  SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .CreateEvents "plungerIM"
    End With

'**********************************************************************************************************
Dim BIP, BC, BL, AM
BIP=0
BC=0
BL=0

Sub BallRelease_unhit()
    BIP=1
    BC=BC+1
    If PupM=1 Then Exit Sub
    If AM=1 And SoundMod=1 Then m00.enabled=true:End If
    a01.enabled=false:a02.enabled=false:a03.enabled=false:a04.enabled=false
End Sub

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : PlaySoundAtVol"drain", Drain, 1
    BIP=0
    If AM=1 Then Exit Sub
    BC=BC-1
    If PupM=1 Then Exit Sub
    If SoundMod=1 then a01.enabled=True:a01.interval=15000
End Sub

'GOALIE Subway
Sub sw44a_Hit:vpmTimer.PulseSw 44: PlaySoundAtVol "subway", ActiveBall, 1: End Sub
Sub sw44b_Hit:vpmTimer.PulseSw 44: PlaySoundAtVol "subway", ActiveBall, 1: End Sub
Sub sw44c_Hit:vpmTimer.PulseSw 44: PlaySoundAtVol "subway", ActiveBall, 1: End Sub
Sub sw44d_Hit:vpmTimer.PulseSw 44: PlaySoundAtVol "subway", ActiveBall, 1: End Sub
Sub sw44e_Hit:vpmTimer.PulseSw 44: PlaySoundAtVol "subway", ActiveBall, 1: End Sub


 '***********************************
 'Left Raising VUK
 '***********************************
 'Variables used for VUK
 Dim raiseballsw, raiseball
 Sub sw45_Hit()
  sw45.Enabled=FALSE
  Controller.switch (45) = True
  PlaySoundAtVol"popper_ball", sw45, 1
 End Sub

 Sub VukLEFTPop(enabled)
  if(enabled and Controller.switch (45)) then
    PlaySoundAtVol SoundFX("Popper",DOFContactors), sw45, 1
    sw45.DestroyBall
    Set raiseball = sw45.CreateBall
    raiseballsw = True
    sw45raiseballtimer.Enabled = True 'Added by Rascal
    sw45.Enabled=TRUE
    Controller.switch (45) = False
  else
    PlaySoundAtVol "Popper", sw45, 1
  end if
End Sub

 Sub sw45raiseballtimer_Timer()
  If raiseballsw = True then
    raiseball.z = raiseball.z + 10
    If raiseball.z > 180 then
      sw45.Kick 180, 3
      Set raiseball = Nothing
      sw45raiseballtimer.Enabled = False
      raiseballsw = False
      PlaySoundAtVol "Wire Ramp", sw45, 1
    End If
  End If
 End Sub


 '***********************************
 'Top Raising VUK
 '***********************************
Sub sw46_Hit:Controller.Switch(46)=1: PlaySoundAtVol "popper_ball", ActiveBall, 1: End Sub

Sub sw46_Timer
  Controller.Switch(46) = 0
  sw46.Timerenabled = 0
End Sub

Sub Kicker5_Timer
  Kicker5.Enabled = 1
  Kicker5.TimerEnabled = 0
End Sub

Sub uVUK(Enabled)
  If Enabled Then
  Kicker5.Enabled = 0
    Kicker5.timerEnabled = 1
  sw46.Kickz 0, 40,89, -20
  'TardisEntrance.KickZ 180, 35, 92, 0
  PlaySoundAtVol SoundFX("Solenoid",DOFContactors), sw46, 1
  sw46.TimerEnabled = 1
  End If
End Sub

 '***********************************
 '***********************************

'Drop Targets
 Sub Sw20_Dropped:dtDrop.Hit 1 :End Sub
 Sub Sw19_Dropped:dtDrop.Hit 2 :End Sub
 Sub Sw18_Dropped:dtDrop.Hit 3 :End Sub
 Sub Sw17_Dropped:dtDrop.Hit 4 :End Sub

'Spinners
Sub sw9_Spin:vpmTimer.PulseSw 9 : PlaySoundAtVol"fx_spinner", sw9, 1 : End Sub

'Wire Triggers
Sub sw16_Hit: PlaySoundAtVol"rollover", ActiveBall, 1 : End Sub 'coded to impulse plunger
'Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub
Sub sw31_Hit:Controller.Switch(31)=1 : PlaySoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw31_UnHit:Controller.Switch(31)=0:End Sub
Sub sw32_Hit:Controller.Switch(32)=1 : PlaySoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw32_UnHit:Controller.Switch(32)=0:End Sub
Sub sw33_Hit:Controller.Switch(33)=1 : PlaySoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw33_UnHit:Controller.Switch(33)=0:End Sub
Sub sw34_Hit:Controller.Switch(34)=1 : PlaySoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw34_UnHit:Controller.Switch(34)=0:End Sub
Sub sw35_Hit:Controller.Switch(35)=1 : PlaySoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw35_UnHit:Controller.Switch(35)=0:End Sub
Sub sw48_Hit:Controller.Switch(48)=1 : PlaySoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw48_UnHit:Controller.Switch(48)=0:End Sub
Sub sw57_Hit:Controller.Switch(57)=1 : PlaySoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw57_UnHit:Controller.Switch(57)=0:End Sub
Sub sw58_Hit:Controller.Switch(58)=1 : PlaySoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw58_UnHit:Controller.Switch(58)=0:End Sub
Sub sw60_Hit:Controller.Switch(60)=1 : PlaySoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw60_UnHit:Controller.Switch(60)=0:End Sub
Sub sw61_Hit:Controller.Switch(61)=1 : PlaySoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw61_UnHit:Controller.Switch(61)=0:End Sub

'Lock Triggers
Sub SW21_Hit:Controller.Switch(21)=1:End Sub
Sub SW21_unHit:Controller.Switch(21)=0:End Sub
Sub SW22_Hit:Controller.Switch(22)=1:End Sub
Sub SW22_unHit:Controller.Switch(22)=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1:End Sub
Sub SW23_unHit:Controller.Switch(23)=0:End Sub
Sub SW24_Hit:Controller.Switch(24)=1:BL=1:End Sub
Sub SW24_unHit:Controller.Switch(24)=0:BL=0:End Sub

 'Stand Up Targets
Sub sw26_Hit:vpmTimer.PulseSw 26:End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:End Sub
Sub sw52_Hit:vpmTimer.PulseSw 52:End Sub

'Ramp Triggers
Sub sw38_Hit:vpmTimer.pulseSw 38 : End Sub
Sub sw39_Hit:vpmTimer.pulseSw 39 : End Sub

'Goalie Opto Trigger
Sub SW25_Hit:Controller.Switch(25)= 1:End Sub
Sub SW25_unHit:Controller.Switch(25)= 0:End Sub

'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 49 : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50 : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51 : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub


 '***********************************
'GOALIE animation
 '***********************************

Dim LBPlace,X
LBPlace=Array(GOALIE0,GOALIE1,GOALIE2,GOALIE3,GOALIE4,GOALIE5,GOALIE6,GOALIE7,GOALIE8,GOALIE9,GOALIE10,GOALIE11,GOALIE12)

Sub UpdateGoalie(aNewPos,aSpeed,aLastPos)
If aNewPos>-1 And aNewPos<13 Then For X=0 To 12:LBPlace(X).IsDropped=1:Next
  Select Case aNewPos
    Case 0:Magnet1.X=159:Magnet1.Y=282:LBPlace(0).IsDropped=0:Goalie.ObjRotZ = 38
    Case 1:Magnet1.X=175:Magnet1.Y=293:LBPlace(1).IsDropped=0:Goalie.ObjRotZ = 26
    Case 2:Magnet1.X=194:Magnet1.Y=302:LBPlace(2).IsDropped=0:Goalie.ObjRotZ = 15
    Case 3:Magnet1.X=212:Magnet1.Y=304:LBPlace(3).IsDropped=0:Goalie.ObjRotZ = 0
    Case 4:Magnet1.X=234:Magnet1.Y=303:LBPlace(4).IsDropped=0:Goalie.ObjRotZ = -15
    Case 5:Magnet1.X=252:Magnet1.Y=295:LBPlace(5).IsDropped=0:Goalie.ObjRotZ = -26
    Case 6:Magnet1.X=272:Magnet1.Y=287:LBPlace(6).IsDropped=0:Goalie.ObjRotZ = -38
    Case 7:Magnet1.X=286:Magnet1.Y=274:LBPlace(7).IsDropped=0:Goalie.ObjRotZ = -48
    Case 8:Magnet1.X=297:Magnet1.Y=262:LBPlace(8).IsDropped=0:Goalie.ObjRotZ = -55
    Case 9:Magnet1.X=305:Magnet1.Y=246:LBPlace(9).IsDropped=0:Goalie.ObjRotZ = -65
    Case 10:Magnet1.X=310:Magnet1.Y=230:LBPlace(10).IsDropped=0:Goalie.ObjRotZ = -75
    Case 11:Magnet1.X=313:Magnet1.Y=213:LBPlace(11).IsDropped=0:Goalie.ObjRotZ = -85
    Case 12:Magnet1.X=312:Magnet1.Y=198:LBPlace(12).IsDropped=0:Goalie.ObjRotZ = -95
  End Select
  Magnet1.Size=75
End Sub


Sub SW43_Hit:Controller.Switch(43)=1:End Sub 'switch 43 is goalie magnet
Sub SW43_unHit:Controller.Switch(43)=0:End Sub
Sub Trigger8_Hit:Magnet1.AddBall ActiveBall:End Sub
Sub Trigger8_UnHit:Magnet1.RemoveBall ActiveBall:End Sub


'Generic Sounds
Sub Trigger001_Hit: PlaySoundAtVol"Wire Ramp", ActiveBall, 1:End Sub
Sub Trigger2_Hit:stopSound "Wire Ramp":End Sub
Sub Trigger3_Hit:stopSound "Wire Ramp":End Sub

'******************************
' PLASTICS TRANSPARENCY CHANGE

Dim bPTActive04, PTImage04
Dim xn
Sub LoadPT04
  bPTActive04 = False
    xn = LoadValue(cGameName, "PTImage04")
    If(xn <> "") Then PTImage04 = xn Else PTImage04 = 0
    UpdatePT04
End Sub

Sub SavePT04
    SaveValue cGameName, "PTImage04", PTImage04
End Sub

Sub NextPT04: PTImage04 = (PTImage04 +1 ) MOD 2: UpdatePT04: SavePT04: End Sub

Sub UpdatePT04
    If PTImage04=0 Then
  For each walls in PlasticsSTAT: walls.TopMaterial = "Plastic with a Light": Next
  W13.TopMaterial = "Plastic with a Light"
  W12.TopMaterial = "Plastic with a Light"
  Primitive002.Material = "Plastic with a Light"
  Primitive88.Material = "Plastic with a Light"
  Primitive9.Material = "Plastic with a Light"
  Primitive52.Material = "Plastic with a Light"
  Primitive63.Material = "Plastic with a Light"
  Primitive10.Material = "Plastic with a Light"
    End If
    If PTImage04=1 Then
     For each walls in PlasticsSTAT: walls.TopMaterial = "Plastic with an Image": Next
     W13.TopMaterial = "Plastic with an Image"
   W12.TopMaterial = "Plastic with an Image"
   Primitive002.Material = "Plastic with an Image"
   Primitive88.Material = "Plastic with an Image"
   Primitive9.Material = "Plastic with an Image"
   Primitive52.Material = "Plastic with an Image"
   Primitive63.Material = "Plastic with an Image"
   Primitive10.Material = "Plastic with an Image"
    End If
End Sub

Dim TPlastics
    If TPlastics=0 Then
    For each walls in PlasticsSTAT: walls.TopMaterial = "Plastic with an Image": Next
    W13.TopMaterial = "Plastic with an Image"
    W12.TopMaterial = "Plastic with an Image"
    Primitive002.Material = "Plastic with an Image"
    Primitive88.Material = "Plastic with an Image"
    Primitive9.Material = "Plastic with an Image"
    Primitive52.Material = "Plastic with an Image"
    Primitive63.Material = "Plastic with an Image"
    Primitive10.Material = "Plastic with an Image"
    Else
    For each walls in PlasticsSTAT: walls.TopMaterial = "Plastic with a Light": Next
    W13.TopMaterial = "Plastic with a Light"
    W12.TopMaterial = "Plastic with a Light"
    Primitive002.Material = "Plastic with a Light"
    Primitive88.Material = "Plastic with a Light"
    Primitive9.Material = "Plastic with a Light"
    Primitive52.Material = "Plastic with a Light"
    Primitive63.Material = "Plastic with a Light"
    Primitive10.Material = "Plastic with a Light"
End If

Dim CDMD
Dim CDMDWOP

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
 For each BWLA in BackWL: BWLA.height = 280: Next
End If

'*************
'   TEAM PLASTICS CHANGE
'*************

Dim bPlaActive04, PLAImage04
Dim xb
Dim walls
Sub LoadPLA04
  bPlaActive04 = False
    xb = LoadValue(cGameName, "PLAImage04")
    If(xb <> "") Then PLAImage04 = xb Else PLAImage04 = 0
    UpdatePLA04
End Sub

Sub SavePLA04
    SaveValue cGameName, "PLAImage04", PLAImage04
End Sub

Sub NextPLA04: PLAImage04 = (PLAImage04 +1 ) MOD 5: UpdatePLA04: SavePLA04: End Sub

Sub UpdatePLA04
    If PLAImage04=0 And CDMDWOP=0 Then
    If LBA=0 Then Goalie.image = "DefenderMap A":End If
    If LBA=1 Then Goalie.image = "DefenderMap B":End If
    For each walls in PlasticsSTAT: walls.image = "Plastics": Next
    W13.image = "Plastics2"
    W12.image = "Plastics2"
    Primitive002.image = "Plastic4"
    Primitive88.image = "Plastic3"
    Ramp001.image = "Ramp1"
    Primitive9.image = "StadiumMap1"
    sw9.image = "Spinner"
    Primitive52.image = "bumpercap"
    Primitive63.image = "bumpercap"
    Primitive10.image = "bumpercap"
    Wall49.image = "plastics"
        Primitive86.image = "Stern Apron Round"
    RFLogo.image = "flipper-r2"
    LFLogo.image = "flipper-l2"
    LFLogo1.image = "flipper-l3"
        CDMD=0
    End If
    If PLAImage04=0 And CDMDWOP=1 Then
    If LBA=0 Then Goalie.image = "DefenderMap A":End If
    If LBA=1 Then Goalie.image = "DefenderMap B":End If
    For each walls in PlasticsSTAT: walls.image = "Plastics": Next
    W13.image = "Plastics2"
    W12.image = "Plastics2"
    Primitive002.image = "Plastic4"
    Primitive88.image = "Plastic3"
    Ramp001.image = "Ramp1"
    Primitive9.image = "StadiumMap1"
    sw9.image = "Spinner"
    Primitive52.image = "bumpercap"
    Primitive63.image = "bumpercap"
    Primitive10.image = "bumpercap"
    Wall49.image = "plastics"
        Primitive86.image = "Stern Apron Round"
    RFLogo.image = "flipper-r2"
    LFLogo.image = "flipper-l2"
    LFLogo1.image = "flipper-l3"
        CDMD=1
    End If
    If PLAImage04=1 Then
    If LBA=0 Then Goalie.image = "DefenderMap-A2":End If
    If LBA=1 Then Goalie.image = "DefenderMap-B2":End If
    For each walls in PlasticsSTAT: walls.image = "Plastics-2": Next
    W13.image = "Plastics2-2"
    W12.image = "Plastics2-2"
    Primitive002.image = "Plastic4-2"
    Primitive88.image = "Plastic3-2"
    Ramp001.image = "Ramp1-2"
    Primitive9.image = "StadiumMap1-2"
    sw9.image = "Spinner-2"
    Primitive52.image = "bumpercap-2"
    Primitive63.image = "bumpercap-2"
    Primitive10.image = "bumpercap-2"
    Wall49.image = "plastics-2"
        Primitive86.image = "Stern Apron Round-2"
    RFLogo.image = "flipper-r2-2"
    LFLogo.image = "flipper-l2-2"
    LFLogo1.image = "flipper-l3-2"
        CDMD=1
    End If
    If PLAImage04=2 Then
    If LBA=0 Then Goalie.image = "DefenderMap-A3":End If
    If LBA=1 Then Goalie.image = "DefenderMap-B3":End If
    For each walls in PlasticsSTAT: walls.image = "Plastics-3": Next
    W13.image = "Plastics2-3"
    W12.image = "Plastics2-3"
    Primitive002.image = "Plastic4-3"
    Primitive88.image = "Plastic3-3"
    Ramp001.image = "Ramp1-3"
    Primitive9.image = "StadiumMap1-3"
    sw9.image = "Spinner-3"
    Primitive52.image = "bumpercap-3"
    Primitive63.image = "bumpercap-3"
    Primitive10.image = "bumpercap-3"
    Wall49.image = "plastics-3"
        Primitive86.image = "Stern Apron Round-3"
    RFLogo.image = "flipper-r2-3"
    LFLogo.image = "flipper-l2-3"
    LFLogo1.image = "flipper-l3-3"
        CDMD=1
    End If
    If PLAImage04=3 Then
    If LBA=0 Then Goalie.image = "DefenderMap-A4":End If
    If LBA=1 Then Goalie.image = "DefenderMap-B4":End If
    For each walls in PlasticsSTAT: walls.image = "Plastics-4": Next
    W13.image = "Plastics2-4"
    W12.image = "Plastics2-4"
    Primitive002.image = "Plastic4-4"
    Primitive88.image = "Plastic3-4"
    Ramp001.image = "Ramp1-4"
    Primitive9.image = "StadiumMap1-4"
    sw9.image = "Spinner-4"
    Primitive52.image = "bumpercap-4"
    Primitive63.image = "bumpercap-4"
    Primitive10.image = "bumpercap-4"
    Wall49.image = "plastics-4"
        Primitive86.image = "Stern Apron Round-4"
    RFLogo.image = "flipper-r2-4"
    LFLogo.image = "flipper-l2-4"
    LFLogo1.image = "flipper-l3-4"
        CDMD=1
    End If
    If PLAImage04=4 Then
    If LBA=0 Then Goalie.image = "DefenderMap-A5":End If
    If LBA=1 Then Goalie.image = "DefenderMap-B5":End If
    For each walls in PlasticsSTAT: walls.image = "Plastics-5": Next
    W13.image = "Plastics2-5"
    W12.image = "Plastics2-5"
    Primitive002.image = "Plastic4-5"
    Primitive88.image = "Plastic3-5"
    Ramp001.image = "Ramp1-5"
    Primitive9.image = "StadiumMap1-5"
    sw9.image = "Spinner-5"
    Primitive52.image = "bumpercap-5"
    Primitive63.image = "bumpercap-5"
    Primitive10.image = "bumpercap-5"
    Wall49.image = "plastics-5"
        Primitive86.image = "Stern Apron Round-5"
    RFLogo.image = "flipper-r2-5"
    LFLogo.image = "flipper-l2-5"
    LFLogo1.image = "flipper-l3-5"
        CDMD=1
    End If
End Sub

'*************
'   LINEBACKER COLOR CHANGE
'*************

Dim bLBBActive04, LBBImage04
Dim xd
Dim LBA
Dim LBB
Sub LoadLBB04
  bLBBActive04 = False
    xd = LoadValue(cGameName, "LBBImage04")
    If(xd <> "") Then LBBImage04 = xd Else LBBImage04 = 0
    UpdateLBB04
End Sub

Sub SaveLBB04
    SaveValue cGameName, "LBBImage04", LBBImage04
End Sub

Sub NextLBB04: LBBImage04 = (LBBImage04 +1 ) MOD 2: UpdateLBB04: SaveLBB04: End Sub

Sub UpdateLBB04
    If LBBImage04=0 And PLAImage04=0 Then Goalie.image = "DefenderMap B":LBA=1:End If
    If LBBImage04=0 And PLAImage04=1 Then Goalie.image = "DefenderMap-B2":LBA=1:End If
    If LBBImage04=0 And PLAImage04=2 Then Goalie.image = "DefenderMap-B3":LBA=1:End If
    If LBBImage04=0 And PLAImage04=3 Then Goalie.image = "DefenderMap-B4":LBA=1:End If
    If LBBImage04=0 And PLAImage04=4 Then Goalie.image = "DefenderMap-B5":LBA=1:End If
    If LBBImage04=0 And PLAImage04=5 Then Goalie.image = "DefenderMap-B6":LBA=1:End If
    If LBBImage04=1 And PLAImage04=0 Then Goalie.image = "DefenderMap A":LBA=0:End If
    If LBBImage04=1 And PLAImage04=1 Then Goalie.image = "DefenderMap-A2":LBA=0:End If
    If LBBImage04=1 And PLAImage04=2 Then Goalie.image = "DefenderMap-A3":LBA=0:End If
    If LBBImage04=1 And PLAImage04=3 Then Goalie.image = "DefenderMap-A4":LBA=0:End If
    If LBBImage04=1 And PLAImage04=4 Then Goalie.image = "DefenderMap-A5":LBA=0:End If
    If LBBImage04=1 And PLAImage04=5 Then Goalie.image = "DefenderMap-A6":LBA=0:End If
End Sub

'*************
'   BACKWALL LIGHTS
'*************

Dim bBWLActive04, BWLImage04
Dim xc
Dim BWLA
Dim BWLB
Sub LoadBWL04
  bBWLActive04 = False
    xc = LoadValue(cGameName, "BWLImage04")
    If(xc <> "") Then BWLImage04 = xc Else BWLImage04 = 0
    UpdateBWL04
End Sub

Sub SaveBWL04
    SaveValue cGameName, "BWLImage04", BWLImage04
End Sub

Sub NextBWL04: BWLImage04 = (BWLImage04 +1 ) MOD 4: UpdateBWL04: SaveBWL04: End Sub

Sub UpdateBWL04
    If BWLImage04=0 Then
       Flasher1.imagea="FC_Yellow":Flasher1.imageb="FC_Yellow":Flasher1.color = RGB(255,255,0):Flasher1.ModulateVsAdd = 0.5
       Flasher2.imagea="FC_Whitered":Flasher2.imageb="FC_Whitered":Flasher2.color = RGB(255,0,0):Flasher2.ModulateVsAdd = 0.1
       Flasher3.imagea="FC_Yellow":Flasher3.imageb="FC_Yellow":Flasher3.color = RGB(255,255,0):Flasher3.ModulateVsAdd = 0.5
       Flasher4.imagea="FC_Whitered":Flasher4.imageb="FC_Whitered":Flasher4.color = RGB(255,0,0):Flasher4.ModulateVsAdd = 0.1
       Flasher5.imagea="FC_White":Flasher5.imageb="FC_White":Flasher5.color = RGB(255,255,255):Flasher5.ModulateVsAdd = 0.1
       Flasher6.imagea="FC_White":Flasher6.imageb="FC_White":Flasher6.color = RGB(255,255,255):Flasher6.ModulateVsAdd = 0.1
       Flasher7.imagea="FC_Whitered":Flasher7.imageb="FC_Whitered":Flasher7.color = RGB(255,0,0):Flasher7.ModulateVsAdd = 0.1
       Flasher8.imagea="FC_Yellow":Flasher8.imageb="FC_Yellow":Flasher8.color = RGB(255,255,0):Flasher8.ModulateVsAdd = 0.5
       Flasher9.imagea="FC_Whitered":Flasher9.imageb="FC_Whitered":Flasher9.color = RGB(255,0,0):Flasher9.ModulateVsAdd = 0.1
       Flasher10.imagea="FC_Yellow":Flasher10.imageb="FC_Yellow":Flasher10.color = RGB(255,255,0):Flasher10.ModulateVsAdd = 0.5
    End If
    If BWLImage04=1 Then
       Flasher1.imagea="FC_Whitegreen":Flasher1.imageb="FC_Whitegreen":Flasher1.color = RGB(0,255,0):Flasher1.ModulateVsAdd = 0.5
       Flasher2.imagea="FC_White":Flasher2.imageb="FC_White":Flasher2.color = RGB(255,255,255):Flasher2.ModulateVsAdd = 0.1
       Flasher3.imagea="FC_Whitegreen":Flasher3.imageb="FC_Whitegreen":Flasher3.color = RGB(0,255,0):Flasher3.ModulateVsAdd = 0.5
       Flasher4.imagea="FC_White":Flasher4.imageb="FC_White":Flasher4.color = RGB(255,255,255):Flasher4.ModulateVsAdd = 0.1
       Flasher5.imagea="FC_Whitegreen":Flasher5.imageb="FC_Whitegreen":Flasher5.color = RGB(0,255,0):Flasher5.ModulateVsAdd = 0.5
       Flasher6.imagea="FC_White":Flasher6.imageb="FC_White":Flasher6.color = RGB(255,255,255):Flasher6.ModulateVsAdd = 0.1
       Flasher7.imagea="FC_Whitegreen":Flasher7.imageb="FC_Whitegreen":Flasher7.color = RGB(0,255,0):Flasher7.ModulateVsAdd = 0.5
       Flasher8.imagea="FC_White":Flasher8.imageb="FC_White":Flasher8.color = RGB(255,255,255):Flasher8.ModulateVsAdd = 0.1
       Flasher9.imagea="FC_Whitegreen":Flasher9.imageb="FC_Whitegreen":Flasher9.color = RGB(0,255,0):Flasher9.ModulateVsAdd = 0.5
       Flasher10.imagea="FC_White":Flasher10.imageb="FC_White":Flasher10.color = RGB(255,255,255):Flasher10.ModulateVsAdd = 0.1
    End If
    If BWLImage04=2 Then
       For each BWLA in BackWL: BWLA.imagea = "FC_White": Next
       For each BWLA in BackWL: BWLA.imageb = "FC_White": Next
       For each BWLA in BackWL: BWLA.color = RGB(255,255,255): Next
       For each BWLA in BackWL: BWLA.ModulateVsAdd = 0.1: Next
    End If
    If BWLImage04=3 Then
       For each BWLA in BackWL: BWLA.imagea = "FC_Whitegreen": Next
       For each BWLA in BackWL: BWLA.imageb = "FC_Whitegreen": Next
       For each BWLA in BackWL: BWLA.color = RGB(0,255,0): Next
       For each BWLA in BackWL: BWLA.ModulateVsAdd = 0.5: Next
    End If
End Sub

'*************
'   JP'S LUT
'*************

Dim bLutActive04, LUTImage04
Dim xa
Sub LoadLUT04
  bLutActive04 = False
    xa = LoadValue(cGameName, "LUTImage04")
    If(xa <> "") Then LUTImage04 = xa Else LUTImage04 = 0
  UpdateLUT04
End Sub

Sub SaveLUT04
    SaveValue cGameName, "LUTImage04", LUTImage04
End Sub

Sub NextLUT04: LUTImage04 = (LUTImage04 +1 ) MOD 13: UpdateLUT04: SaveLUT04: End Sub

Sub UpdateLUT04
Select Case LutImage04
Case 0: Table1.ColorGradeImage = "LUT0"
Case 1: Table1.ColorGradeImage = "LUT1"
Case 2: Table1.ColorGradeImage = "LUT2"
Case 3: Table1.ColorGradeImage = "LUT3"
Case 4: Table1.ColorGradeImage = "LUT4"
Case 5: Table1.ColorGradeImage = "LUT5"
Case 6: Table1.ColorGradeImage = "LUT6"
Case 7: Table1.ColorGradeImage = "LUT7"
Case 8: Table1.ColorGradeImage = "LUT8"
Case 9: Table1.ColorGradeImage = "LUT9"
Case 10: Table1.ColorGradeImage = "LUT10"
Case 11: Table1.ColorGradeImage = "LUT11"
Case 12: Table1.ColorGradeImage = "LUT12"
End Select
End Sub

'**************
' MUSIC TIMERS
Dim SoundMod, PupM

Sub m00_Timer
    Dim zi
    zi = INT(3 * RND(1) )
    Select Case zi
    Case 0:m01.enabled=true:AM=0
    Case 1:m02.enabled=true:AM=0
    Case 2:m03.enabled=true:AM=0
    End Select
m00.enabled=False
end sub

Sub m01_Timer
    Dim zb
    zb = INT(3 * RND(1) )
    Select Case zb
    Case 0:PlayMusic"NFL STERN\Main 01.ogg":m02.enabled=true:m02.interval=151000
    Case 1:PlayMusic"NFL STERN\Main 02.ogg":m02.enabled=true:m02.interval=118500
    Case 2:PlayMusic"NFL STERN\Main 03.ogg":m02.enabled=true:m02.interval=130250
    End Select
m01.enabled=False
end sub

Sub m02_Timer
    Dim zc
    zc = INT(3 * RND(1) )
    Select Case zc
    Case 0:PlayMusic"NFL STERN\Main 04.ogg":m03.enabled=true:m03.interval=115000
    Case 1:PlayMusic"NFL STERN\Main 05.ogg":m03.enabled=true:m03.interval=108250
    Case 2:PlayMusic"NFL STERN\Main 06.ogg":m03.enabled=true:m03.interval=152500
    End Select
m02.enabled=False
end sub

Sub m03_Timer
    Dim zd
    zd = INT(4 * RND(1) )
    Select Case zd
    Case 0:PlayMusic"NFL STERN\Main 07.ogg":m01.enabled=true:m01.interval=156000
    Case 1:PlayMusic"NFL STERN\Main 08.ogg":m01.enabled=true:m01.interval=104500
    Case 2:PlayMusic"NFL STERN\Main 09.ogg":m01.enabled=true:m01.interval=116500
    Case 3:PlayMusic"NFL STERN\Main 10.ogg":m01.enabled=true:m01.interval=106000
    End Select
m03.enabled=False
end sub

'********
'ATTRACT MUSIC TIMERS

Sub a01_Timer
    m01.enabled=False:m02.enabled=False:m03.enabled=False
    Dim ze
    ze = INT(10 * RND(1) )
    Select Case ze
    Case 0:PlayMusic"NFL STERN\Main 11.ogg":a02.enabled=true:a02.interval=24500:AM=1
    Case 1:PlayMusic"NFL STERN\Main 12.ogg":a02.enabled=true:a02.interval=34500:AM=1
    Case 2:PlayMusic"NFL STERN\Main 13.ogg":a02.enabled=true:a02.interval=36700:AM=1
    Case 3:PlayMusic"NFL STERN\Main 14.ogg":a02.enabled=true:a02.interval=15700:AM=1
    Case 4:PlayMusic"NFL STERN\Main 15.ogg":a02.enabled=true:a02.interval=29000:AM=1
    Case 5:PlayMusic"NFL STERN\Main 16.ogg":a02.enabled=true:a02.interval=20750:AM=1
    Case 6:PlayMusic"NFL STERN\Main 17.ogg":a02.enabled=true:a02.interval=18300:AM=1
    Case 7:PlayMusic"NFL STERN\Main 18.ogg":a02.enabled=true:a02.interval=18000:AM=1
    Case 8:PlayMusic"NFL STERN\Main 19.ogg":a02.enabled=true:a02.interval=28600:AM=1
    Case 9:PlayMusic"NFL STERN\Main 20.ogg":a02.enabled=true:a02.interval=28050:AM=1
    End Select
a01.enabled=False
end sub

Sub a02_Timer
    Dim zf
    zf = INT(10 * RND(1) )
    Select Case zf
    Case 0:PlayMusic"NFL STERN\Main 32.ogg":a03.enabled=true:a03.interval=29100:AM=1
    Case 1:PlayMusic"NFL STERN\Main 33.ogg":a03.enabled=true:a03.interval=28850:AM=1
    Case 2:PlayMusic"NFL STERN\Main 34.ogg":a03.enabled=true:a03.interval=29100:AM=1
    Case 3:PlayMusic"NFL STERN\Main 35.ogg":a03.enabled=true:a03.interval=28825:AM=1
    Case 4:PlayMusic"NFL STERN\Main 36.ogg":a03.enabled=true:a03.interval=28800:AM=1
    Case 5:PlayMusic"NFL STERN\Main 37.ogg":a03.enabled=true:a03.interval=28570:AM=1
    Case 6:PlayMusic"NFL STERN\Main 38.ogg":a03.enabled=true:a03.interval=28870:AM=1
    Case 7:PlayMusic"NFL STERN\Main 39.ogg":a03.enabled=true:a03.interval=28760:AM=1
    Case 8:PlayMusic"NFL STERN\Main 40.ogg":a03.enabled=true:a03.interval=27510:AM=1
    Case 9:PlayMusic"NFL STERN\Main 41.ogg":a03.enabled=true:a03.interval=29050:AM=1
    End Select
a02.enabled=False
end sub

Sub a03_Timer
    Dim zg
    zg = INT(11 * RND(1) )
    Select Case zg
    Case 0:PlayMusic"NFL STERN\Main 21.ogg":a04.enabled=true:a04.interval=21800:AM=1
    Case 1:PlayMusic"NFL STERN\Main 22.ogg":a04.enabled=true:a04.interval=32100:AM=1
    Case 2:PlayMusic"NFL STERN\Main 23.ogg":a04.enabled=true:a04.interval=38550:AM=1
    Case 3:PlayMusic"NFL STERN\Main 24.ogg":a04.enabled=true:a04.interval=28460:AM=1
    Case 4:PlayMusic"NFL STERN\Main 25.ogg":a04.enabled=true:a04.interval=21570:AM=1
    Case 5:PlayMusic"NFL STERN\Main 26.ogg":a04.enabled=true:a04.interval=30790:AM=1
    Case 6:PlayMusic"NFL STERN\Main 27.ogg":a04.enabled=true:a04.interval=12800:AM=1
    Case 7:PlayMusic"NFL STERN\Main 28.ogg":a04.enabled=true:a04.interval=60100:AM=1
    Case 8:PlayMusic"NFL STERN\Main 29.ogg":a04.enabled=true:a04.interval=53550:AM=1
    Case 9:PlayMusic"NFL STERN\Main 30.ogg":a04.enabled=true:a04.interval=60250:AM=1
    Case 10:PlayMusic"NFL STERN\Main 31.ogg":a04.enabled=true:a04.interval=27460:AM=1
    End Select
a03.enabled=False
end sub

Sub a04_Timer
    Dim zh
    zh = INT(7 * RND(1) )
    Select Case zh
    Case 0:PlayMusic"NFL STERN\Eagles 01.ogg":a01.enabled=true:a01.interval=35000:AM=1
    Case 1:PlayMusic"NFL STERN\Eagles 02.ogg":a01.enabled=true:a01.interval=104000:AM=1
    Case 2:PlayMusic"NFL STERN\Eagles 03.ogg":a01.enabled=true:a01.interval=25000:AM=1
    Case 3:PlayMusic"NFL STERN\Eagles 04.ogg":a01.enabled=true:a01.interval=12000:AM=1
    Case 4:PlayMusic"NFL STERN\Eagles 05.ogg":a01.enabled=true:a01.interval=79000:AM=1
    Case 5:PlayMusic"NFL STERN\Eagles 06.ogg":a01.enabled=true:a01.interval=14000:AM=1
    Case 6:PlayMusic"NFL STERN\Eagles 07.ogg":a01.enabled=true:a01.interval=22000:AM=1
    End Select
a04.enabled=False
end sub

'***********
'  COIN TIMERS
Dim Coin
Coin=0

Sub c01_Timer
    If Coin=0 Then
    Dim zx
    zx = INT(14 * RND(1) )
    Select Case zx
    Case 0:PlaySound "0td" :c02.enabled=true:c02.interval=2500:Coin=1
    Case 1:PlaySound "0td2" :c02.enabled=true:c02.interval=2000:Coin=1
    Case 2:PlaySound "0dc" :c02.enabled=true:c02.interval=7000:Coin=1
    Case 3:PlaySound "0fg 01" :c02.enabled=true:c02.interval=2000:Coin=1
    Case 4:PlaySound "0fg 02" :c02.enabled=true:c02.interval=2000:Coin=1
    Case 5:PlaySound "0fg 03" :c02.enabled=true:c02.interval=2000:Coin=1
    Case 6:PlaySound "0whistle" :c02.enabled=true:c02.interval=3000:Coin=1
    Case 7:PlaySound "0td3" :c02.enabled=true:c02.interval=2500:Coin=1
    Case 8:PlaySound "0td4" :c02.enabled=true:c02.interval=4000:Coin=1
    Case 9:PlaySound "0td5" :c02.enabled=true:c02.interval=2500:Coin=1
    Case 10:PlaySound "0Eagles" :c02.enabled=true:c02.interval=4000:Coin=1
    Case 11:PlaySound "0Eaglesb" :c02.enabled=true:c02.interval=4000:Coin=1
    Case 12:PlaySound "0Eaglesc" :c02.enabled=true:c02.interval=6500:Coin=1
    Case 13:PlaySound "0Eaglesd" :c02.enabled=true:c02.interval=5500:Coin=1
    End Select
    End If
c01.enabled=False
end sub

Sub c02_Timer
    Coin=0
c02.enabled=False
end sub


'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 5 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step

     'Special Handling
     'If chgLamp(ii,0) = 2 Then solTrough chgLamp(ii,1)
     'If chgLamp(ii,0) = 4 Then PFGI chgLamp(ii,1)

        Next
    End If
    UpdateLamps
End Sub


Sub UpdateLamps()
NFadeL 1, L1
NFadeL 2, L2
NFadeL 3, L3
NFadeL 4, L4
NFadeL 5, L5
NFadeL 6, L6
NFadeL 7, L7
NFadeL 8, L8
NFadeL 9, L9
NFadeL 10, L10
NFadeL 11, L11
NFadeL 12, L12
NFadeL 13, L13
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
Flash 25, F25 'Backwall
Flash 26, F26 'Backwall
Flash 27, F27 'Backwall

NFadeL 29, L29
NFadeL 30, L30
NFadeL 31, L31
NFadeL 32, L32
NFadeL 33, L33
NFadeL 34, L34
NFadeL 35, L35

NFadeLm 38, L38 'Bumper 1
NFadeL 38, L38b
NFadeLm 39, L39 'Bumper 2
NFadeL 39, L39b
NFadeLm 40, L40 'Bumper 3
NFadeL 40, L40b

NFadeL 42, L42
NFadeL 43, L43
NFadeL 44, L44
NFadeL 45, L45
NFadeL 46, L46
NFadeL 47, L47
NFadeL 48, L48
NFadeL 49, L49
NFadeL 50, L50
NFadeL 51, L51
NFadeL 52, L52
NFadeL 53, L53
NFadeL 54, L54
NFadeL 55, L55
NFadeL 56, L56
NFadeL 57, L57
NFadeL 58, L58
NFadeL 59, L59
NFadeL 60, L60
NFadeL 61, L61

NFadeL 64, L64
NFadeL 65, L65
NFadeL 66, L66
NFadeL 67, L67
NFadeL 68, L68
NFadeL 69, L69
NFadeL 70, L70
NFadeL 71, L71
NFadeL 72, L72

'Solenoid Controlled

NFadeLm 120, S20a
NFadeLm 120, S20b
NFadeLm 120, S20c
NFadeLm 120, S20d
NFadeLm 120, S20e
NFadeLm 120, S20f
NFadeLm 120, S20g
NFadeL 120, S20h


Flash 130, S30 'BackPanel Goal

NFadeLm 131, S31a
NFadeLm 131, S31b
NFadeLm 131, S31c
NFadeL 131, S31d

NFadeLm 132, S32a
NFadeLm 132, S32b
NFadeLm 132, S32c
NFadeL 132, S32d

'Backwall Bulbs controlled by GI string

Flashm 134, Flasher1
Flashm 134, Flasher2
Flashm 134, Flasher3
Flashm 134, Flasher4
Flashm 134, Flasher5
Flashm 134, Flasher6
Flashm 134, Flasher7
Flashm 134, Flasher8
Flashm 134, Flasher9
Flash 134, Flasher10

End Sub


' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

' #####################################
' ###### Fluppers Domes V2.2 #####
' #####################################

 Sub FlashYellow(flstate)
  If Flstate Then
    Objlevel(1) = 1 : FlasherFlash1_Timer
  End If
 End Sub

 Sub FlashWhite(flstate)
  If Flstate Then
    Objlevel(2) = 1 : FlasherFlash2_Timer
  End If
 End Sub

 Sub FlashRed(flstate)
  If Flstate Then
    Objlevel(3) = 1 : FlasherFlash3_Timer
  End If
 End Sub


Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 1   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 1   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objlight(20)
Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "yellow" :
InitFlasher 2, "White" :
InitFlasher 3, "red"



Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ =  atn( (tablewidth/2 - objbase(nr).x) / (objbase(nr).y - tableheight*1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ : objflasher(nr).height = objbase(nr).z + 60
  End If
  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0 : objlit(nr).visible = 0 : objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX : objlit(nr).RotY = objbase(nr).RotY : objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX : objlit(nr).ObjRotY = objbase(nr).ObjRotY : objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x : objlit(nr).y = objbase(nr).y : objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness
  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4)
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4)
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)
  If not objflasher(nr).TimerEnabled Then objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objlit(nr).visible = 1 : End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objlit(nr).visible = 0 : End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub
Sub FlasherFlash5_Timer() : FlashFlasher(5) : End Sub
Sub FlasherFlash6_Timer() : FlashFlasher(6) : End Sub
Sub FlasherFlash7_Timer() : FlashFlasher(7) : End Sub
Sub FlasherFlash8_Timer() : FlashFlasher(8) : End Sub
Sub FlasherFlash9_Timer() : FlashFlasher(9) : End Sub
Sub FlasherFlash10_Timer() : FlashFlasher(10) : End Sub
Sub FlasherFlash11_Timer() : FlashFlasher(11) : End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 62
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), ActiveBall, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 59
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), ActiveBall, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
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


'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

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

    For b = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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


'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  FlipperLSh1.RotZ = LeftFlipper1.currentangle

    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
    LFLogo1.RotY = LeftFlipper1.CurrentAngle


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
  PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
         Controller.Games("nfl").Settings.Value("DMD_Colorize") =0
         Controller.Games("nfl").Settings.Value("sound_mode") =0
         Controller.Games("nfl").Settings.Value("dmd_red") = 255
         Controller.Games("nfl").Settings.Value("dmd_green") = 88
         Controller.Games("nfl").Settings.Value("dmd_blue") = 32
         Controller.Games("nfl").Settings.Value("dmd_red66") = 155
         Controller.Games("nfl").Settings.Value("dmd_green66") = 39
         Controller.Games("nfl").Settings.Value("dmd_blue66") = 0
         Controller.Games("nfl").Settings.Value("dmd_red33") = 100
         Controller.Games("nfl").Settings.Value("dmd_green33") = 26
         Controller.Games("nfl").Settings.Value("dmd_blue33") = 0
         Controller.Games("nfl").Settings.Value("dmd_red0") = 0
         Controller.Games("nfl").Settings.Value("dmd_green0") = 0
         Controller.Games("nfl").Settings.Value("dmd_blue0") = 0
         Controller.Stop
End Sub
