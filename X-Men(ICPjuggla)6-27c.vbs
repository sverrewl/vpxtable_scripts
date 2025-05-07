  Option Explicit
    Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Added InitVpmFFlipsSAM
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

Const VolBump   = 2    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


Dim DesktopMode: DesktopMode = Table.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive58.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive58.visible=0
End if


  LoadVPM "01560000", "sam.VBS", 3.10

 '******************* Options *********************
' DMD/Backglass Controller Setting
Const cController = 0   '0=Use value defined in cController.txt, 1=VPinMAME, 2=UVP server, 3=B2S server, 4=B2S with DOF (disable VP mech sounds)
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







'********************
'Standard definitions
'********************

  Const cGameName = "xmn_151"

     Const UseSolenoids = 1
     Const UseLamps = 0
     Const UseSync = 1
     Const HandleMech = 0

     'Standard Sounds
     Const SSolenoidOn = "Solenoid"
     Const SSolenoidOff = ""
     Const SCoin = "CoinIn"


   'Variables
    Dim xx
    Dim bsL,mTLMag,mTRMag,mDMag,tbTrough,Bump1, Bump2, Bump3, bsRHole, MsLHole, bsLHole, bsTrough, bsSaucer, DTBank5, DTBank4, DTBank3, mDiverter
  Dim PlungerIM


     'Table Init
  Sub Table_Init
  vpmInit Me
  With Controller
        .GameName = cGameName
        .SplashInfoLine = "XMEN LE, Stern 2012"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 1
    .Hidden = 0
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
  End With

    On Error Goto 0


       '**Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 21, 20, 19, 18, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 8
    bsTrough.InitExitSnd "ballrelease", "Solenoid"
    bsTrough.Balls = 4

  '***Left Hole bsLHole
     Set bsLHole = New cvpmBallStack
     With bsLHole
         .InitSw 0, 4, 0, 0, 0, 0, 0, 0
         .InitKick sw4, 168, 12
         .KickZ = 0.4
         .InitExitSnd "popper_ball", "Solenoid"
         .KickForceVar = 2
     End With



  Set mTLMag= New cvpmMagnet
  With mTLMag
    .InitMagnet Magnet1, 16
    .GrabCenter = False
    .solenoid=32
    .CreateEvents "mTLMag"
  End With

  Set mDMag= New cvpmMagnet
  With mDMag
    .InitMagnet Magnet3, 16
    .GrabCenter = False
    .solenoid=4
    .CreateEvents "mDMag"
  End With


  'Wolvie2.isdropped=1
'wr2.alpha = 0

'************************************************************************************************************************************

  'DropTargets

      '**Main Timer init
           PinMAMETimer.Enabled = 1

'   Set bsSaucer=New cvpmBallStack
' bsSaucer.InitSaucer S37,37,283,15
' bsSaucer.InitExitSnd"Popper","SolOn"

'Nudging
    vpmNudge.TiltSwitch=-7
    vpmNudge.Sensitivity=3
    vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot)

  'Magneto Lock
lockPin1.Isdropped=1:lockPin2.Isdropped=1
  'switches
'    sw1b.IsDropped = 1
'    sw2b.IsDropped = 1
'    sw7b.IsDropped = 1
'    sw8b.IsDropped = 1
'    sw41b.IsDropped = 1
'    sw42b.IsDropped = 1
InitVpmFFlipsSAM
  End Sub

   Sub Table_Paused:Controller.Pause = 1:End Sub
   Sub Table_unPaused:Controller.Pause = 0:End Sub
   Sub Table_exit()

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
'   Sub Table_exit()
'
' Controller.Stop
'
'End sub


'*****Keys
 Sub Table_KeyDown(ByVal keycode)

  If Keycode = LeftFlipperKey then
    'SolLFlipper true
  End If
  If Keycode = RightFlipperKey then
'   SolRFlipper true
  End If
    If keycode = PlungerKey Then Plunger.Pullback
'   If keycode = LeftTiltKey Then LeftNudge 80, 1, 20
'    If keycode = RightTiltKey Then RightNudge 280, 1, 20
'    If keycode = CenterTiltKey Then CenterNudge 0, 1, 25
    If vpmKeyDown(keycode) Then Exit Sub

End Sub

Sub Table_KeyUp(ByVal keycode)
  If vpmKeyUp(keycode) Then Exit Sub
  If Keycode = LeftFlipperKey then
    'SolLFlipper false
  End If
  If Keycode = RightFlipperKey then
    'SolRFlipper False
  End If
  If Keycode = StartGameKey Then Controller.Switch(16) = 0
    If keycode = PlungerKey Then
    Plunger.Fire
  End If
End Sub

   'Solenoids
SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
SolCallback(3) = "bsLHole.SolOut"
SolCallback(4) = "mDMag.MagnetOn="
SolCallback(5) = "bsL.SolOut"
SolCallback(6) = "CLockUp"
SolCallback(7) = "CLockLatch"
SolCallback(9)  = "SetLamp 141,"
SolCallback(10) = "SetLamp 142,"
SolCallback(11) = "SetLamp 143,"
'SolCallback(12) = "SolURFlipper"
SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"
SolCallback(26) = "vpmSolDiverter RampDiverter,""Diverter"","



'Flashers

SolCallback(17) = "setlamp 117,"
SolCallback(18) = "setlamp 118,"
SolCallback(21) = "setlamp 121,"
SolCallback(32) = "setlamp 132,"
SolCallback(22) = "setlamp 122,"
SolCallback(25) = "setlamp 125,"
SolCallback(27) = "setlamp 127,"
SolCallback(28) = "setlamp 128,"
SolCallback(29) = "setlamp 129,"
SolCallback(31) = "setlamp 131,"


Sub solTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 22
  End If
 End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
  End If
 End Sub

Sub CLockUp(Enabled)
  If Enabled Then
lockPin1.Isdropped=0:lockPin2.Isdropped=0
  End If
 End Sub

Sub CLockLatch(Enabled)
  If Enabled Then
lockPin1.Isdropped=1:lockPin2.Isdropped=1
  End If
 End Sub

'Scoop
 Dim aBall, aZpos
 Dim bBall, bZpos

 Sub sw4_Hit
     Set bBall = ActiveBall
     PlaySoundAtVol "kicker_enter_center", ActiveBall, 1
     bZpos = 35
   ClearBallID
     Me.TimerInterval = 2
     Me.TimerEnabled = 1
 End Sub

 Sub sw4_Timer
     bBall.Z = bZpos
     bZpos = bZpos-4
     If bZpos <-30 Then
         Me.TimerEnabled = 0
         Me.DestroyBall
         bsLHole.AddBall Me
     End If
 End Sub

Sub SolDiverter(enabled)
  If enabled Then
    RampDiverter.rotatetoend : PlaySoundAtVol "Diverter", ActiveBall, 1
  Else
    RampDiverter.rotatetostart : PlaySoundAtVol "Diverter", ActiveBall, 1
  End If
End Sub

    Set bsL = New cvpmBallStack
    With bsL
        .InitSw 0, 55, 0, 0, 0, 0, 0, 0
        .InitKick sw55, 180, 2
        .InitExitSnd "popper_ball", "Solenoid"
        .KickForceVar = 3
    End With

   Sub Drain_Hit():PlaySoundAtVol "Drain", Drain, 1
  'ClearBallID
  BallCount = BallCount - 1
  bsTrough.AddBall Me
  If BallCount = 0 then GIOff
   End Sub

Sub sw1_Hit:Me.TimerEnabled = 1:sw1p.TransX = -2:vpmTimer.PulseSw 1:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
Sub sw1_Timer:Me.TimerEnabled = 0:sw1p.TransX = 0:End Sub
Sub sw2_Hit:Me.TimerEnabled = 1:sw2p.TransX = -2:vpmTimer.PulseSw 2:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
Sub sw2_Timer:Me.TimerEnabled = 0:sw2p.TransX = 0:End Sub
Sub sw7_Hit:Me.TimerEnabled = 1:sw7p.TransX = -2:vpmTimer.PulseSw 7:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
Sub sw7_Timer:Me.TimerEnabled = 0:sw7p.TransX = 0:End Sub
Sub sw8_Hit:Me.TimerEnabled = 1:sw8p.TransX = -2:vpmTimer.PulseSw 8:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
Sub sw8_Timer:Me.TimerEnabled = 0:sw8p.TransX = 0:End Sub
Sub sw11_Hit:Controller.Switch(11) = 1:PlaySoundAtVol "Gate",ActiveBall, VolGates:LeftCount = LeftCount + 1:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub
Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAtVol "Gate",ActiveBall, VolGates:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub
Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAtVol "fx_sensor",ActiveBall, 1:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub
Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "fx_sensor",ActiveBall, 1:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAtVol "fx_sensor",ActiveBall, 1:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAtVol "fx_sensor",ActiveBall, 1:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAtVol "fx_sensor",ActiveBall, 1:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAtVol "fx_sensor",ActiveBall, 1:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAtVol "fx_sensor",ActiveBall, 1:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw36_Hit:VengeanceHit 2:vpmTimer.PulseSw 36:PlaySoundAtVol "target",ActiveBall, 1:End Sub
'Sub sw36_Timer:Wolvie1.IsDropped = 0:Wolvie2.IsDropped = 1:wr1.alpha = 1:wr1.triggersingleupdate:wr2.alpha = 0:wr2.triggersingleupdate:Me.TimerEnabled = 0:End Sub
Sub sw38_Hit:PlaysoundAtVol "rollover",ActiveBall, VolRol:Controller.Switch(38)=1:End Sub   'Lock 2
Sub sw38_unHit:Controller.Switch(38)=0:End Sub
Sub sw39_Hit:PlaysoundAtVol "rollover",ActiveBall, VolRol:Controller.Switch(39)=1:End Sub   'Lock 3
Sub sw39_unHit:Controller.Switch(39)=0:End Sub
Sub sw40_Hit:PlaysoundAtVol "rollover",ActiveBall, VolRol:Controller.Switch(40)=1:End Sub   'Lock 4
Sub sw40_unHit:Controller.Switch(40)=0:End Sub
Sub sw41_Hit:Me.TimerEnabled = 1:sw41p.TransX = -2:vpmTimer.PulseSw 41:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
Sub sw41_Timer:Me.TimerEnabled = 0:sw41p.TransX = 0:End Sub
Sub sw42_Hit:Me.TimerEnabled = 1:sw42p.TransX = -2:vpmTimer.PulseSw 42:PlaySoundAtVol "target",ActiveBall, VolTarg:End Sub
Sub sw42_Timer:Me.TimerEnabled = 0:sw42p.TransX = 0:End Sub
Sub sw47_Spin:vpmTimer.PulseSw 47::playsoundAtVol"spinner",sw47, VolSpin:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAtVol "Gate",ActiveBall, VolGates:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub
Sub sw49_Hit:Controller.Switch(49) = 1:PlaySoundAtVol "fx_sensor",ActiveBall, 1:End Sub
Sub sw49_UnHit:Controller.Switch(49) = 0:End Sub
Sub sw52_Hit:Controller.Switch(52) = 1:PlaySoundAtVol "Gate",ActiveBall, VolGates:RightCount = RightCount + 1:End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub
Sub sw54_Hit:Controller.Switch(54) = 1:PlaySoundAtVol "fx_sensor",ActiveBall, 1:End Sub
Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub
Sub sw55a_Hit:ClearBallId:PlaySoundAtVol "kicker_enter",ActiveBall, VolKick:bsL.AddBall Me:End Sub


Sub SolLFlipper(Enabled)
     If Enabled Then
     PlaySoundAtVol "flipperupleft", LeftFlipper, VolFlip
     LeftFlipper.RotateToEnd
     Else
     PlaySoundAtVol "flipperdown", LeftFlipper, VolFLip
    LeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
     PlaySoundAtVol "flipperupright", RightFlipper, VolFlip
     RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
     PlaySoundAtVol "flipperdown", RightFlipper, VolFlip
    RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
 End Sub


Dim BallCount:BallCount = 0
   Sub BallRelease_UnHit()
  'NewBallID
    BallCount = BallCount + 1
    GIOn
  End Sub

 Dim RStep, LStep

 Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 26
  PlaySoundAtVol SoundFX("left_slingshot"), sling2, 1
  LSling.Visible = 0
  LSling1.Visible = 1
  sling2.TransZ = -20
  LStep = 0
  Me.TimerEnabled = 1
  End Sub

 Sub LeftSlingShot_Timer
  Select Case LStep
    Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
    Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
  End Select
  LStep = LStep + 1
 End Sub

 Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw 27
  PlaySoundAtVol SoundFX("right_slingshot"), sling1, 1
  RSling.Visible = 0
  RSling1.Visible = 1
  sling1.TransZ = -20
  RStep = 0
  Me.TimerEnabled = 1
  End Sub

 Sub RightSlingShot_Timer
  Select Case RStep
    Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
    Case 4 RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
  End Select
  RStep = RStep + 1
End Sub



    Const IMPowerSetting = 60
    Const IMTime = 0.6
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0
        .InitExitSnd "plunger2", "plunger"
        .CreateEvents "plungerIM"
    End With


   'Bumpers
      Sub Bumper1b_Hit
    vpmTimer.PulseSw 31
    PlaySoundAtVol ("fx_bumper1"), Bumper1b, VolBump
    End Sub


      Sub Bumper2b_Hit
    vpmTimer.PulseSw 30
    PlaySoundAtVol ("fx_bumper1"), Bumper2b, VolBump
    End Sub


      Sub Bumper3b_Hit
    vpmTimer.PulseSw 32
    PlaySoundAtVol ("fx_bumper1"), Bumper3b, VolBump
    End Sub







'***********************************************
'***********************************************
          ' Lamps
'***********************************************
'***********************************************

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

 Sub UpdateLamps()
NFadeL 3, l3
NFadeL 4, l4
NFadeL 5, l5
NFadeL 6, l6
NFadeL 7, l7
NFadeL 8, l8
NFadeL 9, l9
NFadeL 10, l10
NFadeL 11, l11
NFadeL 12, l12
NFadeL 13, l13
NFadeL 14, l14
NFadeL 15, l15
NFadeL 16, l16
NFadeL 17, l17
NFadeL 18, l18
NFadeL 19, l19
NFadeL 20, l20
NFadeL 21, l21
NFadeL 22, l22
NFadeL 23, l23
NFadeL 24, l24
NFadeL 25, l25
NFadeL 28, l28
NFadeL 29, l29
NFadeL 30, l30
NFadeL 31, l31
NFadeL 33, l33
NFadeL 34, l34
NFadeL 35, l35
NFadeL 36, l36
NFadeL 37, l37
NFadeL 38, l38
NFadeL 43, l43
NFadeL 44, l44
NFadeL 45, l45
NFadeL 46, l46
NFadeL 47, l47
NFadeL 48, l48
NFadeL 49, l49
NFadeL 50, l50
NFadeL 51, l51
NFadeL 52, l52
NFadeL 53, l53
NFadeL 54, l54
NFadeL 55, l55
NFadeL 56, l56
NFadeL 57, l57
NFadeL 60, l60 'top left bumper
NFadeL 61, l61 'right bumper
NFadeL 62, l62 'lower left bumper
NFadeL 65, l65
NFadeL 66, l66
NFadeL 67, l67
NFadeL 68, l68
NFadeL 69, l69
NFadeL 70, l70
NFadeL 71, l71

NFadeLm 117, f117a
NFadeLm 117, f117b
NFadeLm 117, f117c
NFadeLm 117, f117d
NFadeLm 117, f117e
NFadeLm 117, f117f
NFadeLm 118, f118a
NFadeLm 118, f118b
NFadeLm 118, f118c
NFadeLm 118, f118d
NFadeLm 118, f118e
NFadeLm 118, f118f

NFadeLm 122, f122a
NFadeL 122, f122b

'NFadeL 121, f121
NFadeL 125, f25
NFadeL 132, f132
NFadeL 59, f59

'NFadeL 131, f131

NfadeL 127, f127

NFadeLm 128, l28a
NFadeLm 128, l28b
NFadeL 128, l28c
NFadeLm 129, l29a
NFadeLm 129, l29b
NFadeL 129, l29c

NFadeL 141, f9
NFadeL 142, f10
NFadeL 143, f11

FadePrim 121, wolverine, "wolverine2_30", "wolverine2_20", "wolverine2_16","wolverine2"
FadePrim 131, Primitive5, "magnetotxt3_30", "magnetotxt3_20", "magnetotxt3_16","magnetotxt3"
End Sub

'Sub FlasherTimer_Timer()
' Flashm 117, f117b
' Flash 117, f117
' Flashm 118, f118b
' Flash 118, f118
' Flash 121, f121
' Flash 122, f122
' Flash 125, f125
' Flash 127, f127
' Flashm 128, f128a
' Flashm 128, f128b
' Flash 128, f128
' Flashm 129, f129a
' Flash 129, f129
' Flash 131, f131
'
' FlashAR 59, f59b, "wf_on", "wf_a", "wf_b"
'
' End Sub

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

Dim x
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



Sub swPlunger_Hit:BallinPlunger = 1:End Sub

Sub swPlunger_UnHit:BallinPlunger = 0:End Sub


' Create ball in kicker and assign a Ball ID used mostly in non-vpm tables
Sub CreateBallID(Kickername)
    For cnt = 1 to ubound(ballStatus)
        If ballStatus(cnt) = 0 Then
            Set currentball(cnt) = Kickername.createball
            currentball(cnt).uservalue = cnt
            ballStatus(cnt) = 1
            ballStatus(0) = ballStatus(0) + 1
            If coff = False Then
                If ballStatus(0)> 1 and XYdata.enabled = False Then XYdata.enabled = True
            End If
            Exit For
        End If
    Next
End Sub

Sub ClearBallID
    On Error Resume Next
    iball = ActiveBall.uservalue
    currentball(iball).UserValue = 0
    If Err Then Msgbox Err.description & vbCrLf & iball
    ballStatus(iBall) = 0
    ballStatus(0) = ballStatus(0) -1
    On Error Goto 0
End Sub


Sub Table_exit()
  Controller.Pause = False
  Controller.Stop
End Sub







'******************************
'****END DONOR TABLE SCRIPT****
'******************************

Sub UpdateFlipperLogo_Timer
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
    RFlogo1.RotY = RightFlipper1.CurrentAngle
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
Sub leftrampcounter_hit:LeftCount = 1:end sub
Sub leftdrop_hit
  If LeftCount = 1 then
    playsound "BallDrop"
  End If
  LeftCount = 0
End Sub

Dim RightCount:RightCount = 0
Sub rightrampcounter_hit:RightCount = 1:end sub
Sub rightdrop_hit
  If RightCount = 1 then
    playsound "BallDrop"
  End If
  RightCount = 0
End Sub

Sub RLS_Timer()
    RampGate3.RotZ = -(Spinner2.currentangle)
    RampGate1.RotZ = -(Spinner1.currentangle)
    RampGate4.RotZ = -(Spinner4.currentangle)
    RampGate2.RotZ = -(Spinner3.currentangle)
'    SpinnerT1.RotZ = -(sw9.currentangle)
    SpinnerT2.RotZ = -(sw47.currentangle)
' ampprim.objrotx = ampf.currentangle
End Sub

'Sub PrimT_Timer
' If Wall455.isdropped = true then ampf.rotatetoend:end if
' if Wall455.isdropped = false then ampf.rotatetostart:end if
' ampprim.objrotx = ampf.currentangle
'End Sub

Sub uppostt_Timer
  If lockpin1.isdropped = true then uppostf.rotatetoend
  if lockpin1.isdropped = false then uppostf.rotatetostart
  uppostp.transy = uppostf.currentangle
End Sub

dim vengeanceBall,zz
set vengeanceBall = kicker1.createball
kicker1.kick 0,0
Dim Mag1
  Set mag1= New cvpmMagnet
  With mag1
    .InitMagnet vengeanceTrigger, 25
    .GrabCenter = False
    .magnetOn = True
  End With

Sub VengeanceHit (Vmove)
  dim yy
  vengeanceBall.vely = Vmove / 2
  yy = (rnd(1) - 0.5) * Vmove
  vengeanceBall.velx = yy
  VengeanceTimer.Enabled = 1:zz=0
  'Playsound "solenoidon"
End Sub


Sub VengeanceTimer_Timer
  wolverine.roty = (vengeanceBall.x - vengeanceTrigger.x) * 2
  'wolverine.transY = (vengeanceBall.y - vengeanceTrigger.y) '* 2
  zz = zz + 1
  if zz = 100 then VengeanceTimer.Enabled = 0
End Sub

Sub PrimT_Timer
  if f9.State = 1 then f9a.visible = 1 else f9a.visible = 0
  if f10.State = 1 then f10a.visible = 1 else f10a.visible = 0
  if f11.State = 1 then f11a.visible = 1 else f11a.visible = 0
    if f25.State = 1 then f25a.visible = 1 else f25a.visible = 0
  if l28a.State = 1 then f28a.visible = 1 else f28a.visible = 0
  if l28a.State = 1 then f28b.visible = 1 else f28b.visible = 0
  if l28a.State = 1 then f28c.visible = 1 else f28c.visible = 0
  if l29a.State = 1 then f29a.visible = 1 else f29a.visible = 0
  if l29a.State = 1 then f29b.visible = 1 else f29b.visible = 0
  if l29a.State = 1 then f29c.visible = 1 else f29c.visible = 0
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Table" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Table.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Table" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Table.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Table" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Table.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
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

