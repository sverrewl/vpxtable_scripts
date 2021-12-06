  Option Explicit
    Randomize


'  LoadVPM "01560000", "sam.VBS", 3.10
  LoadVPM "01120100", "de.vbs", 3.02

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


 '******************* Options *********************
' DMD/Backglass Controller Setting
Const cController = 3   '0=Use value defined in cController.txt, 1=VPinMAME, 2=UVP server, 3=B2S server, 4=B2S with DOF (disable VP mech sounds)
'*************************************************

Dim cNewController
Sub LoadVPM(VPMver, VBSfile, VBSver)
  Dim FileObj, ControllerFile, TextStr

  On Error Resume Next
  If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
  ExecuteGlobal GetTextFile(VBSfile)
  If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description

  cNewController = 1
  If cController = 0 then
    Set FileObj=CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) then
      Msgbox "Visual Pinball\User directory does not exist. Defaulting to vPinMame"
    ElseIf Not FileObj.FileExists(UserDirectory & "cController.txt") then
      Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
      ControllerFile.WriteLine 1: ControllerFile.Close
    Else
      Set ControllerFile=FileObj.GetFile(UserDirectory & "cController.txt")
      Set TextStr=ControllerFile.OpenAsTextStream(1,0)
      If (TextStr.AtEndOfStream=True) then
        Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
        ControllerFile.WriteLine 1: ControllerFile.Close
      Else
        cNewController=Textstr.ReadLine: TextStr.Close
      End If
    End If
  Else
    cNewController = cController
  End If

  Select Case cNewController
    Case 1
      Set Controller = CreateObject("VPinMAME.Controller")
      If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
      If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
      If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
    Case 2
      Set Controller = CreateObject("UltraVP.BackglassServ")
    Case 3,4
      Set Controller = CreateObject("B2S.Server")
  End Select
  On Error Goto 0
End Sub

'*************************************************************
'Toggle DOF sounds on/off based on cController value
'*************************************************************
Dim ToggleMechSounds
Function SoundFX (sound)
    If cNewController= 4 and ToggleMechSounds = 0 Then
        SoundFX = ""
    Else
        SoundFX = sound
    End If
End Function

Sub DOF(dofevent, dofstate)
  If cNewController>2 Then
    If dofstate = 2 Then
      Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
    Else
      Controller.B2SSetData dofevent, dofstate
    End If
  End If
End Sub

'******************************
'***START DONOR TABLE SCRIPT***
'******************************




 ' Based on Tales from the Crypt by Data East 1993 / IPD No. 2493 / December, 1993 / 4 Players
 ' VP9 v2.0 by JPSalas 2009
 ' Used Lander's table as a reference
 ' and I also used some of his subs

 Dim bsTrough, bsBallRelease, dtBank, bsRVuk, bsLVuk, bsPScoop, mTombStone, cbCaptive, plungerIM, bsVuk
 Dim x, bump1, bump2, bump3, BallInVuk, ShakeDir
 Dim RipStones
 RipStones = Array(rip, rip1, rip2, rip3, rip4, rip5, rip6, rip7, rip8, rip9, rip10, rip11, rip12, rip13, rip14, rip15, rip16)

 Const UseSolenoids = 1
 Const UseLamps = 0
 Const UseGI = 0
 Const UseSync = 0
 Const HandleMech = 0

 ' Standard Sounds
 Const SSolenoidOn = "Solenoid"
 Const SSolenoidOff = ""
 Const SFlipperOn = "FlipperUp"
 Const SFlipperOff = "FlipperDown"
 Const SCoin = "coin"

 '************
 ' Table init.
 '************

 Sub table1_Init
       Dim cGameName
     vpmInit Me
     ' SaveValue cRegistryName, "Options", 0: Exit Sub  ' This will clear your Registry settings if uncommented
     JPTCOptions = CInt("0" & LoadValue(cRegistryName, "Options") )
     With Controller
         cGameName = Array(Romset1, Romset2, Romset3, Romset4) ((JPTCOptions And(5 * cOptRom) ) \ cOptRom)
           .GameName = cGameName
         .SplashInfoLine = "Based on Tales from the Crypt, Data East 1993" & vbNewLine & "VP9-VPM table by JPSalas v.2.0"
          'DMD position and size for 1400x1050
         '.Games(cGameName).Settings.Value("dmd_pos_x")=500
         '.Games(cGameName).Settings.Value("dmd_pos_y")=2
         '.Games(cGameName).Settings.Value("dmd_width")=400
         '.Games(cGameName).Settings.Value("dmd_height")=92
         '.Games(cGameName).Settings.Value("rol")=0
         .HandleMechanics = 0
         .HandleKeyboard = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .ShowTitle = 0
         .Hidden = 0
         If Err Then MsgBox Err.Description
     End With
     On Error Goto 0
     Controller.SolMask(0) = 0
     vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
     Controller.Run

     ' Nudging
     vpmNudge.TiltSwitch = 1
     vpmNudge.Sensitivity = 5
     vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

     ' Trough & Ball Release
     Set bsTrough = New cvpmBallStack
     With bsTrough
         .InitSw 0, 14, 13, 12, 11, 10, 9, 0
         .Balls = 6
     End With

     Set bsBallRelease = New cvpmBallStack
     With bsBallRelease
         .InitSaucer BallRelease, 15, 80, 6
         .InitEntrySnd "Solenoid", "Solenoid"
         .InitExitSnd "ballrelease", "Solenoid"
     End With

     ' Drop targets
     set dtBank = new cvpmdroptarget
     With dtBank
         .InitDrop Array(sw41, sw42, sw43), Array(41, 42, 43)
         .Initsnd "droptarget", "resetdrop"
     End With

     ' Right Vuk
     Set bsRVuk = New cvpmBallStack
     With bsRVuk
         .InitSw 0, 52, 0, 0, 0, 0, 0, 0
'         .InitKick sw52, 165, 10
         .InitKick sw52, 165, 75
     .KickZ=150
         .InitExitSnd "popper_ball", "solenoid"
     End With

     Set bsVuk = New cvpmBallStack
     With bsVuk
         .InitSw 0, 38, 0, 0, 0, 0, 0, 0
         .InitKick sw38, 75, 35
     .KickZ=100
         .InitExitSnd "popper_ball", "solenoid"
     End With

     ' Power Scoop
     Set bsPScoop = New cvpmBallStack
     With bsPScoop
         .InitSw 0, 55, 0, 0, 0, 0, 0, 0
         .InitKick sw55, 167, 40
         .KickZ = 55
         .KickForceVar = 2
         .InitExitSnd "popper_ball", "solenoid"
         .Balls = 0
         .KickBalls = 1
     End With

     ' RIP Tombstone
     Set mTombStone = new cvpmMech
     With mTombStone
         .Mtype = vpmMechOneSol + vpmMechReverse
         .Sol1 = 15
         .Length = 220
         .Steps = 16
         .Acc = 30
         .Ret = 0
         .AddSw 33, 0, 0
         .AddSw 36, 15, 15
         .Callback = GetRef("RipUpdate")
         .Start
     End With

     ' Captive Ball
     ' Init Captive Ball
     Set cbCaptive = New cvpmCaptiveBall
     With cbCaptive
         .InitCaptive CapTrigger, CapWall, CapKicker, 348
         .Start
         .ForceTrans = .99
         .MinForce = 3.8
         .CreateEvents "cbCaptive"
     End With

     ' Impulse Plunger
     Const IMPowerSetting = 55 ' Plunger Power
     Const IMTime = 0.6        ' Time in seconds for Full Plunge
     Set plungerIM = New cvpmImpulseP
     With plungerIM
         .InitImpulseP swplunger, IMPowerSetting, IMTime
         .Random 0.3
         .InitExitSnd "plunger2", "plunger"
         .CreateEvents "plungerIM"
     End With

     ' Init GI, Flashers & Tomb Stones
     SolGi 1
     For Each x in RipStones:x.IsDropped = 1:Next
     Rip.IsDropped = 0


     KickBack.PullBack
     Diverter.IsDropped = 1
     ShakeDir = 0

     ' Main Timer init
     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1
'     table1.MaxBallSpeed = 60
'     StartShake

  dim bulb
  if NightDay <= 25 then
    for each bulb in Collection1
      bulb.intensity = bulb.intensity
    next
  end if
  if NightDay <= 50 and NightDay >=26 then
    for each bulb in Collection1
      bulb.intensity = bulb.intensity / 1.5
    next
  end if
  if NightDay <= 75 and NightDay >=51 then
    for each bulb in Collection1
      bulb.intensity = bulb.intensity / 2
    next
  end if
  if NightDay <= 100 and NightDay >=76 then
    for each bulb in Collection1
      bulb.intensity = bulb.intensity / 2.5
    next
  end if

 End Sub

 Sub table1_Paused:Controller.Pause = 1:End Sub
 Sub table1_unPaused:Controller.Pause = 0:End Sub

 '**********
 ' Keys
 '**********

 Sub Table1_KeyDown(ByVal Keycode)
     If keycode = keyFront Then vpmTimer.pulsesw 8 'Buy-in Button - 2 key
     If keycode = PlungerKey Then vpmTimer.PulseSw 62
     If keycode = KeyRules Then Rules
'      If keycode = LeftTiltKey Then LeftNudge 80, 1.2, 20: PlaySound "nudge_left"
'      If keycode = RightTiltKey Then RightNudge 280, 1.2, 20: PlaySound "nudge_right"
'      If keycode = CenterTiltKey Then CenterNudge 0, 1.6, 25: PlaySound "nudge_forward"
'      If keycode = 45 Then LeftNudge 80, 2, 30: PlaySound "nudge_left"
'      If keycode = 52 Then RightNudge 280, 2, 30: PlaySound "nudge_right"
     If vpmKeyDown(keycode) Then Exit Sub
 End Sub

 Sub Table1_KeyUp(ByVal Keycode)
     If vpmKeyUp(keycode) Then Exit Sub
 End Sub

 '**********************
 'Options Menu Variables
 '**********************

 Const cRegistryName = "JPTFTC"
 Const cOptRom = &H100

 Const RomSet1 = "tftc_303"
 Const RomSet2 = "tftc_300"
 Const RomSet3 = "tftc_200"
 Const RomSet4 = "tftc_104"

 Dim JPTCOptions, Fade

 Private Sub TFTCShowDips
     If Not IsObject(vpmDips) Then ' First time
         Set vpmDips = New cvpmDips
         With vpmDips
             .AddForm 100, 350, "TFTC Options"
             .AddFrameExtra 0, 0, 220, "Select Rom Set", 5 * cOptRom, Array("3.03", 0 * cOptRom, "3.00", 1 * cOptRom, "2.00", 2 * cOptRom, "1.04 Spanish", 3 * cOptRom)
         End With
     End If
     JPTCOptions = vpmDips.ViewDipsExtra(JPTCOptions)
     SaveValue cRegistryName, "Options", JPTCOptions
 End Sub

 Set vpmShowDips = GetRef("TFTCShowDips")


 '*********
 ' Switches
 '*********

 ' Slings & div switches

 Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 26
  PlaySoundAtVol SoundFX("left_slingshot"), ActiveBall, 1
  Me.TimerEnabled = 1
  End Sub

 Sub LeftSlingShot_Timer

 End Sub

 Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw 27
  PlaySoundAtVol SoundFX("right_slingshot"), ActiveBall, 1
  Me.TimerEnabled = 1
  End Sub

 Sub RightSlingShot_Timer

 End Sub

 ' Bumpers

      Sub Bumper1_Hit
      vpmTimer.PulseSw 49
      PlaySoundAtVol SoundFX("fx_bumper1"), ActiveBall, 1
      End Sub


      Sub Bumper2_Hit
      vpmTimer.PulseSw 50
      PlaySoundAtVol SoundFX("fx_bumper1"), ActiveBall, 1
       End Sub

      Sub Bumper3_Hit
      vpmTimer.PulseSw 51
      PlaySoundAtVol SoundFX("fx_bumper1"), ActiveBall, 1
       End Sub

 ' Drain holes
 Sub Drain_Hit:PlaysoundAtVol "drain",drain,1:bsTrough.AddBall Me:End Sub

 ' Rollovers
 Sub sw17_Hit:Controller.Switch(17) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
 Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

 Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
 Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

 Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
 Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

 Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
 Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

 Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
 Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

 Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
 Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

 Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
 Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

 Sub sw32_Hit:Controller.Switch(32) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
 Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

 Sub sw16_Hit():Controller.Switch(16) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
 Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

 ' Ramp Switches
 Sub sw44_Hit():Controller.Switch(44) = 1:PlaySoundAtVol "gate",ActiveBall, VolGates:End Sub
 Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

 Sub sw45_Hit():Controller.Switch(45) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
 Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

 Sub sw46_Hit():Controller.Switch(46) = 1:PlaySoundAtVol "gate",ActiveBall, VolGates:End Sub
 Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

 Sub sw47_Hit():Controller.Switch(47) = 1:LeftCount = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
 Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

 Sub sw57_Hit():Controller.Switch(57) = 1:RightCount = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
 Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub Rhelp3_Hit:RightCount2 = 1:End Sub

 ' Droptargets
 Sub sw41_Hit:dtBank.hit 1:End Sub
 Sub sw42_Hit:dtBank.hit 2:End Sub
 Sub sw43_Hit:dtBank.hit 3:End Sub

 ' Targets
 Sub sw20_Hit:vpmTimer.PulseSw 20:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
 'Sub sw20_Timer:sw20.IsDropped = 0:sw20a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
 Sub sw21_Hit:vpmTimer.PulseSw 21:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
 'Sub sw21_Timer:sw21.IsDropped = 0:sw21a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
 Sub sw22_Hit:vpmTimer.PulseSw 22:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
 'Sub sw22_Timer:sw22.IsDropped = 0:sw22a.IsDropped = 1:Me.TimerEnabled = 0:End Sub

 Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
 'Sub sw28_Timer:sw28.IsDropped = 0:sw28a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
 Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
 'Sub sw29_Timer:sw29.IsDropped = 0:sw29a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
 Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
 'Sub sw30_Timer:sw30.IsDropped = 0:sw30a.IsDropped = 1:Me.TimerEnabled = 0:End Sub

 Sub sw39_Hit:vpmTimer.PulseSw 39:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
 'Sub sw39_Timer:sw39.IsDropped = 0:sw39a.IsDropped = 1:Me.TimerEnabled = 0:End Sub

 Sub rip_Hit:vpmTimer.pulsesw 37:PlaySoundAtVol "wallhit",ActiveBall, 1:End Sub

 ' Spinners
 Sub sw40_Spin:vpmTimer.pulsesw 40:PlaySoundAtVol "spinner",sw40,VolSpin:End Sub
 Sub sw48_Spin:vpmTimer.pulsesw 48:PlaySoundAtVol "spinner",sw48,VolSpin:End Sub
 Sub sw56_Spin:vpmTimer.pulsesw 56:PlaySoundAtVol "spinner",sw56,VolSpin:End Sub

 ' Center Vuk
 Sub sw38a_Hit
     PlaySoundAtVol "kicker_enter_Center", ActiveBall, 1
    bsVuk.AddBall 0
 '    Controller.Switch(38) = 1
 '    BallInVuk = 1
     sw38a.Enabled = 0
  Me.DestroyBall
 End Sub

'Sub sw38_unhit:sw38.timerenabled = 1:end sub
'Sub sw38_Timer:sw38.enabled = 1:sw38.timerenabled = 0:End sub

' Sub SolPopper(Enabled)
'     If Enabled And BallInVuk = 1 Then
'         Vuk1.TimerEnabled = 1
'     End if
' End sub

 ' Vuk animation using kickersvuk

' Sub Vuk1_Timer():Me.TimerEnabled = 0:BallInVuk = 0:Playsound "popper":sw38.DestroyBall:Controller.Switch(38) = 0:Vuk1.createball:Vuk2.TimerEnabled = 1:End sub
' Sub Vuk2_Timer():Me.TimerEnabled = 0:Vuk1.DestroyBall:Vuk2.createball:Vuk3.TimerEnabled = 1:End sub
' Sub Vuk3_Timer():Me.TimerEnabled = 0:Vuk2.DestroyBall:Vuk3.createball:Vuk4.TimerEnabled = 1:End sub
' Sub Vuk4_Timer():Me.TimerEnabled = 0:Vuk3.DestroyBall:Vuk4.createball:Vuk5.TimerEnabled = 1:End sub
' Sub Vuk5_Timer():Me.TimerEnabled = 0:Vuk4.DestroyBall:Vuk5.createball:Vuk6.TimerEnabled = 1:End sub
' Sub Vuk6_Timer():Me.TimerEnabled = 0:Vuk5.DestroyBall:Vuk6.createball:sw38.Enabled = 1:Vuk6.kick 30, 8:End sub ' Saucers
 Sub sw53_Hit:PlaySoundAtVol "kicker_enter",sw53,VolKick:sw53.DestroyBall:vpmTimer.PulseSwitch 53, 100, "HandleTrough":sw38a.enabled = 1:End Sub
 Sub sw54_Hit:PlaySoundAtVol "kicker_enter",sw54,VolKick:sw54.DestroyBall:vpmTimer.PulseSwitch 54, 100, "HandleTrough":End Sub

 Sub HandleTrough(swNo):bsRVuk.AddBall 0:End Sub

 ' Power Scoop using the gobble effect by Bendigo
Dim aBall, aZpos

Sub sw55a_Hit
  Set aBall = ActiveBall
  aZpos = 50
  Me.TimerInterval = 2
  Me.TimerEnabled = 1
End Sub

Sub sw55a_Timer
  aBall.Z = aZpos
  aZpos = aZpos-2
  If aZpos <40 Then
    Me.TimerEnabled = 0
    Me.DestroyBall
    bspScoop.AddBall Me
    sw55a.Enabled = 0
  End If
End Sub

Sub k3trig_hit
  sw55a.enabled = 1
end sub

 ' Tombstone update
 Sub RipUpdate(newpos, speed, lastpos)
     RipStones(lastpos).IsDropped = True
     RipStones(newpos).IsDropped = False
 End Sub

 ' Ramp Helpers

 Sub TopRamp_Hit ' helps to make the loop
     ActiveBall.VelY = 13
     ActiveBall.VelX = -6
 End Sub


 ' gate hits - only sound
 Sub Gate1_Hit:PlaySoundAtVol "gate",gate1,VolGates:End Sub
 Sub Gate2_Hit:PlaySoundAtVol "gate",gate2,VolGates:End Sub
 Sub Gate3_Hit:PlaySoundAtVol "gate",gate3,VolGates:End Sub

 '********************
 'Solenoids & Flashers
 '********************

 SolCallback(1) = "SolRelease"
 SolCallback(2) = "bsBallRelease.SolOut"
 SolCallback(3) = "Auto_Plunger"
 SolCallback(4) = "dtBank.SolDropUp"
 SolCallback(5) = "bsPScoop.SolOut"
' SolCallback(6) = "SolPopper"
 SolCallback(6) = "bsVuk.SolOut"
 SolCallback(7) = "bsRVuk.SolOut"
 SolCallback(9) = "SolDiv"
 SolCallback(8) = "vpmSolSound ""Knocker"","
 SolCallback(11) = "SolGi"
 'SolCallback(16) = "SolShaker"
 SolCallback(22) = "vpmSolAutoPlunger KickBack,10,"
 SolCallback(25) = "SetLamp 65,"
 SolCallBack(26) = "SetLamp 66,"
 SolCallBack(27) = "SetLamp 67,"
 SolCallBack(28) = "SetLamp 68,"
 SolCallBack(29) = "SetLamp 69,"
 SolCallBack(30) = "SetLamp 70,"
 SolCallBack(31) = "SetLamp 71,"
 SolCallback(32) = "SetLamp 72,"

 Sub SolRelease(Enabled)
     If Enabled Then
         bsTrough.ExitSol_On
         BallRelease.CreateBall
         bsBallRelease.AddBall 0
     End If
 End Sub

 Sub Auto_Plunger(Enabled)
     If Enabled Then
         PlungerIM.AutoFire
     End If
 End Sub

 Sub SolDiv(Enabled):PlaySound "diverter":Diverter.IsDropped = NOT Enabled:End Sub

 Sub SolShaker(Enabled):ShakerMotor.Enabled = Abs(Enabled):End Sub

 Sub ShakerMotor_Timer
     ShakeDir = ABS(ShakeDir -1)
     PlaySound "quake"
     If ShakeDir=0 Then
         Nudge 270, 1
     Else
         Nudge 90, 1
     End If
 End Sub

   '**************
 ' Flipper Subs
 '**************

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
     PlaySoundAtVol SoundFX("FlipperUpLeft"), LeftFlipper, VolFlip
     LeftFlipper.RotateToEnd
     Else
     PlaySoundAtVol SoundFX("FlipperDown"), LeftFlipper, VoLFlip
     LeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
     PlaySoundAtVol SoundFX("FlipperUpRight"), RightFlipper,VolFlip
     RightFlipper.RotateToEnd:RightFlipper2.RotateToEnd
     Else
     PlaySoundAtVol SoundFX("FlipperDown"),RightFlipper,VolFlip
     RightFlipper.RotateToStart:RightFlipper2.RotateToStart
     End If
 End Sub

 ' GI
 Sub SolGi(Enabled)
  If Enabled then
    GIOff
  Else
    GiOn
  End If
 End Sub



 Dim LampState(200), FadingLevel(200), FadingState(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown
'Dim x

AllLampsOff()
LampTimer.Interval = 40 'lamp fading speed
LampTimer.Enabled = 1
'
FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
      FlashState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
        Next
    End If

    UpdateLamps
End Sub

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 50   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

 Sub UpdateLamps
NFadeL 1, l1a
NFadeL 2, l2a
NFadeL 3, l3a
NFadeL 4, l4a
NFadeL 5, l5a
NFadeL 6, l6a
NFadeL 7, l7a
NFadeL 8, l8a
NFadeL 9, l9a
NFadeL 10, l10a
NFadeL 11, l11a
'NFadeL 12, l12
NFadeL 13, l13
'NFadeL 14, l14a
NFadeL 15, l15
'NFadeL 16, l16a
'NFadeL 17, l17a
NFadeL 18, l18a
NFadeL 19, l19a
NFadeL 20, l20a
NFadeL 21, l21a
NFadeL 22, l22a
NFadeL 23, l23a
NFadeL 24, l24a
NFadeL 25, l25a
'NFadeL 26, l26a
NFadeL 27, l27
NFadeL 28, l28a
NFadeL 29, l29a
NFadeL 30, l30a
NFadeL 31, l31a
NFadeL 32, l32a
NFadeL 33, l33a
NFadeL 34, l34a
NFadeL 35, l35a
NFadeL 36, l36a
NFadeL 37, l37a
NFadeL 38, l38a
NFadeL 39, l39a
NFadeL 40, l40a
NFadeL 41, l41a
NFadeL 42, l42a
NFadeL 43, l43a
NFadeL 44, l44a
NFadeL 45, l45a
NFadeL 46, l46a
NFadeL 47, l47a
NFadeL 48, l48a
NFadeL 49, l49a
NFadeL 50, l50a
NFadeL 51, l51a
NFadeL 52, l52a
NFadeL 53, l53a
NFadeL 54, l54a
NFadeL 55, l55a
NFadeL 56, l56a
NFadeL 57, l57
NFadeL 58, l58
NFadeL 59, l59
NFadeL 60, l60a
NFadeL 61, l61a
NFadeL 62, l62
NFadeL 63, l63
NFadeL 64, l64

NFadeLm 70, f70a
NFadeL 70, f70b
'NFadeLm 126, f126a
'NFadeL 126, f126b
'NFadeLm 132, f132a
NFadeLm 68, f68a
NFadeLm 68, f68b
NFadeL 68, f68c
NFadeLm 69, f69a
NFadeLm 69, f69b
NFadeL 69, f69c
'NFadeL 131, f131a
'NFadeL 132, f132a
'NFadeLm 131, f131a
'NFadeLm 131, f131b
'NFadeLm 131, f131c
'NFadeL 131, f131d

' FlashAR 60, f160, "bf_on", "bf_a", "bf_b", ARRefresh
' FlashAR 61, f161, "bf_on", "bf_a", "bf_b", ARRefresh
' FlashAR 62, f162, "bf_on", "bf_a", "bf_b", ARRefresh
' NFadeL 60, bumper2b
' NFadeL 61, bumper1b
' NFadeL 62, bumper3b
' FadeL 63, l63, l63a
'
' FadeLm 120, l90, l90a
' FlashAR 120, f120, "wf_on", "wf_a", "wf_b", ARRefresh
' FadeLm 122, l89, l89a
' FlashAR 122, f189, "rf_on", "rf_a", "rf_b", ARRefresh
' FlashAR 121, f121, "yf_on", "yf_a", "yf_b", ARRefresh
' FlashARm 123, f125, "bf_on", "bf_a", "bf_b", ARRefresh
' FlashAR 123, f125a, "bf_on", "bf_a", "bf_b", ARRefresh
' FlashAR 126, f126, "yf_on", "yf_a", "yf_b", ARRefresh
' FlashAR 127, f127, "wf_on", "wf_a", "wf_b", ARRefresh
' FlashAR 129, f129, "wf_on", "wf_a", "wf_b", ARRefresh
' FlashARm 131, f131a, "rf_on", "rf_a", "rf_b", ARRefresh
' FlashAR 131, f131, "rf_on", "rf_a", "rf_b", ARRefresh
' FlashAR 132, f132, "rf_on", "rf_a", "rf_b", ARRefresh
   End Sub

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0
        Case 3:pri.image = c:FadingLevel(nr) = 1
        Case 4:pri.image = b:FadingLevel(nr) = 2
        Case 5:pri.image = a:FadingLevel(nr) = 3
    End Select
End Sub

''Lights

Sub NFadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0
        Case 5:a.State = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0
        Case 5:a.State = 1
    End Select
End Sub

' Flasher objects
' Uses own faster timer

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 1000 Then
                FlashLevel(nr) = 1000
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:FadingLevel(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub


Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
    LampState(nr) = abs(value) + 4
FadingLevel(nr ) = abs(value) + 4: FadingState(nr ) = abs(value) + 4
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub

Sub FlasherTimer_Timer()
'Flash 3, fire
'
'Flash 80, f80
'Flash 119, f119
'Flash 120, f120 'right ramp flash
'Flash 129, f29 'left loop / spinner flash
'Flash 131, f31 'vengeance flash
 End Sub


 '***************
 ' Rules
 '***************
 Dim RuleWindow
 RuleWindow = 0
 Sub Rules
     If RuleWindow = 0 Then
         Dim objShell:Set objShell = CreateObject("Wscript.Shell")
         objShell.Run "TFTC_JP_VP9.hta"
         RuleWindow = 1
         Controller.Pause = 1
     Else
         RuleWindow = 0
         Controller.Pause = 0
     End If
 End Sub




'******************************
'****END DONOR TABLE SCRIPT****
'******************************

Sub UpdateFlipperLogo_Timer
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
    RFlogo1.RotY = RightFlipper2.CurrentAngle
    LLogo.ObjRotZ = LeftFlipper.CurrentAngle -90
    Rlogo.ObjRotZ = RightFlipper.CurrentAngle -90
    Rlogo1.ObjRotZ = RightFlipper2.CurrentAngle -90
End Sub

Sub GIOn
  dim bulb
  for each bulb in Collection1
  bulb.state = 1
  next
End Sub

Sub GIOff
  dim bulb
  for each bulb in Collection1
  bulb.state = 0
  next
End Sub

 'Sub RightSlingShot_Timer:Me.TimerEnabled = 0:End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


'Sub LRRail_Hit:PlaySound "fx_metalrolling", 0, 150, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub
'
'Sub RLRail_Hit:PlaySound "fx_metalrolling", 0, 150, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub


Dim LeftCount:LeftCount = 0

Sub leftdrop_hit
  If LeftCount = 1 then
    playsoundAtVol "BallDrop", ActiveBall, 1
  End If
  LeftCount = 0
End Sub

Dim RightCount:RightCount = 0

Sub rightdrop_hit
  If RightCount = 1 then
    playsoundAtVol "BallDrop", ActiveBall, 1
  End If
  RightCount = 0
End Sub

Dim RightCount2:RightCount2 = 0
Sub RHelp1_hit
  If RightCount2 = 1 then
    playsoundAtVol "BallDrop", ActiveBall, 1
  End If
  RightCount2 = 0
End Sub

Sub RLS_Timer()

              RampGate4.RotZ = -(Gate2.currentangle)
              RampGate3.RotZ = -(Gate1.currentangle)
              RampGate2.RotZ = -(Spinner4.currentangle)
              RampGate1.RotZ = -(Spinner2.currentangle)
              SpinnerT1.RotZ = -(sw40.currentangle)
              SpinnerT2.RotZ = -(sw56.currentangle)
              SpinnerT4.RotZ = -(sw48.currentangle)
End Sub

Sub PrimT_Timer
  If sw41.isdropped = true then sw41f.rotatetoend:end if
  if sw41.isdropped = false then sw41f.rotatetostart:end if
  sw41p.transy = sw41f.currentangle
  If sw42.isdropped = true then sw42f.rotatetoend:end if
  if sw42.isdropped = false then sw42f.rotatetostart:end if
  sw42p.transy = sw42f.currentangle
  If sw43.isdropped = true then sw43f.rotatetoend:end if
  if sw43.isdropped = false then sw43f.rotatetostart:end if
  sw43p.transy = sw43f.currentangle
  If rip16.isdropped = false then tombstone.z = 0
  If rip15.isdropped = false then tombstone.z = 12
  If rip14.isdropped = false then tombstone.z = 24
  If rip13.isdropped = false then tombstone.z = 36
  If rip12.isdropped = false then tombstone.z = 48
  If rip11.isdropped = false then tombstone.z = 60
  If rip10.isdropped = false then tombstone.z = 72
  If rip9.isdropped = false then tombstone.z = 84
  If rip8.isdropped = false then tombstone.z = 96
  If rip7.isdropped = false then tombstone.z = 108
  If rip6.isdropped = false then tombstone.z = 120
  If rip5.isdropped = false then tombstone.z = 132
  If rip4.isdropped = false then tombstone.z = 144
  If rip3.isdropped = false then tombstone.z = 156
  If rip2.isdropped = false then tombstone.z = 168
  If rip1.isdropped = false then tombstone.z = 180
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************
' Getting often owerflow on Csng on this table - why ?

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
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

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 9 ' total number of balls
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

