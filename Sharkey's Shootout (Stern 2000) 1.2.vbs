'UnderTrough=43 needs to be added <?>
'UnderTrough switch drops a target when hit - no idea where switch goes
'Lamp 24=Bottom Up Post (UK)

Option Explicit

Const BallSize = 55
Const BallMass = 1.1

Const GIOnDuringAttractMode   = 0         '1 - GI on during attract, 0 - GI off during attract

'******************* Options *********************
' DMD/Backglass Controller Setting (by gtxjoe)
Const cController = 1   '0=Use value defined in cController.txt, 1=VPinMAME, 2=UVP server, 3=B2S server, 4=B2S with DOF (disable VP mech sounds)
'*************************************************

Dim cNewController
LoadVPM "01500000", "SEGA.VBS", 3.10
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
      Set Controller = CreateObject("B2S.Server")
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

'LoadVPM "01500000", "SEGA.VBS", 3.10
'Sub LoadVPM(VPMver, VBSfile, VBSver)
' On Error Resume Next
'   If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
'   ExecuteGlobal GetTextFile(VBSfile)
'   If Err Then MsgBox "Unable to open " & VBSfile & _
'   ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description : Err.Clear
'   Set Controller = CreateObject("VPinMAME.Controller")
'   If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
'   If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & _
'   " required." : Err.Clear
'   If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
' On Error Goto 0
'End Sub

Const UseSolenoids  = 1
Const UseLamps    = 0
Const UseSync   = True
Const UseGI     = False     'Only WPC games have special GI circuit.

Const SSolenoidOn = "SolOn"       'Solenoid activates
Const SSolenoidOff  = "SolOff"      'Solenoid deactivates
Const SFlipperOn  = "FlipperUp"   'Flipper activated
Const SFlipperOff = "FlipperDown" 'Flipper deactivated
Const SCoin     = "Quarter"     'Coin inserted

Sub table1_KeyDown(ByVal keycode)
  If KeyCode=LeftMagnaSave Then Controller.Switch(1)=1 'Aux Solenoid 1 left post save
  If KeyCode=RightMagnaSave Then Controller.Switch(8)=1 'Aux Solenoid 3 right post save
  If KeyDownHandler(keycode) Then Exit Sub
  'If keycode = PlungerKey Then Plunger.Pullback
    If keycode = PlungerKey Then Plunger.Pullback

End Sub

Sub table1_KeyUp(ByVal keycode)
  If KeyCode=LeftMagnaSave Then Controller.Switch(1)=0
  If KeyCode=RightMagnaSave Then Controller.Switch(8)=0
  If KeyUpHandler(keycode) Then Exit Sub
  'If keycode = PlungerKey Then Plunger.Fire:PlaySound"Plunger"
    If keycode = PlungerKey Then
    Plunger.Fire
  End If
End Sub

SolCallBack(1)      = "SolTrough"
SolCallBack(2)    = "SolAutofire"
SolCallback(3)    = "bsVUK.SolOut"
SolCallback(4)    = "dtSingle.SolHit 1,"
SolCallback(5)    = "dtDrop.SolDropUp"
SolCallback(6)    = "dtSingle.SolDropUp"
'SolCallback(7)   = "SetFlash 107," '107 108 121 122 123 132
SolCallback(7)      = "Multi107"
SolCallback(8)    = "SolFlasher"
'SolCallback(9)   = "vpmSolSound ""jet3"","
'SolCallback(10)    = "vpmSolSound ""jet3"","
'SolCallback(11)    = "vpmSolSound ""jet3"","
'SolCallBack(12)  = 'Top Orbit Magnet
'SolCallBack(13)  = 'Mystery Magnet
SolCallBack(17)   = "vpmSolSound ""sling"","
SolCallBack(18)   = "vpmSolSound ""sling"","
SolCallBack(19)   = "LoopPost.IsDropped=Not"
'SolCallBack(20)  = 'Motor Driver Relay Board
'SolCallBack(21)  = "SetFlash 121,"
SolCallback(21) = "Multi121"
SolCallBack(22) = "SetLamp 122," 'pf insert
SolCallBack(23) = "SolFlasherB"
'SolCallBack(23)  = "sol23flasher"
'SolCallBack(24)  = 'Optional Coin Meter
SolCallBack(25)   = "dtDrop.SolHit 1,"
SolCallBack(26)   = "dtDrop.SolHit 2,"
SolCallBack(27)   = "dtDrop.SolHit 3,"
SolCallBack(28)   = "dtDrop.SolHit 4,"
SolCallBack(29)   = "dtDrop.SolHit 5,"
SolCallBack(30)   = "dtDrop.SolHit 6,"
SolCallBack(31)   = "dtDrop.SolHit 7,"
SolCallback(32) = "SolFlasherC" 'pf insert
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"
'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Flipper1,"
SolCallback(sLRFlipper)="SolRFlipper"
SolCallback(sLLFlipper)="SolLFlipper"
SolCallback(33)="LeftPost.IsDropped=Not"
SolCallback(34)="Post.IsDropped=Not"
SolCallback(35)="RightPost.IsDropped=Not"



Sub SolTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 15
  End If
End Sub

Sub Multi107(Enabled)
  If Enabled Then
    SetFlash 107, 1
    l107a.State = 1
    l107b.state = 1
  Else
    SetFlash 107, 0
    l107a.State = 0
    l107b.state = 0
  End If
End Sub

Sub Multi121(Enabled)
  If Enabled Then
    SetFlash 121, 1
    l121.State = 1
  Else
    SetFlash 121, 0
    l121.State = 0
  End If
End Sub

Sub SolFlasher(Enabled)
  If Enabled Then
    SetLamp 108, 1
    SetLamp 118, 1
  Else
    SetLamp 108, 0
    SetLamp 118, 0
  End If
End Sub

Sub SolFlasherB(Enabled)
  If Enabled Then
    mball1.state = 1
    mball2.state = 1
  Else
    mball1.state = 0
    mball2.state = 0
  End If
End Sub

Sub SolFlasherC(Enabled)
  If Enabled Then
    SetLamp 132, 1
  Else
    SetLamp 132, 0
  End If
End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
  End If
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub

'***Slings and rubbers
 ' Slings
 'Dim LStep, RStep

 Sub LeftSlingShot_Slingshot
  Leftsling = True
  Controller.Switch(26) = 1
  PlaySound "slingshot":LeftSlingshot.TimerEnabled = 1
  End Sub

Dim Leftsling:Leftsling = False

Sub LS_Timer()
  If Leftsling = True and Left1.ObjRotZ < -7 then Left1.ObjRotZ = Left1.ObjRotZ + 2
  If Leftsling = False and Left1.ObjRotZ > -20 then Left1.ObjRotZ = Left1.ObjRotZ - 2
  If Left1.ObjRotZ >= -7 then Leftsling = False
  If Leftsling = True and Left2.ObjRotZ > -212.5 then Left2.ObjRotZ = Left2.ObjRotZ - 2
  If Leftsling = False and Left2.ObjRotZ < -199.5 then Left2.ObjRotZ = Left2.ObjRotZ + 2
  If Left2.ObjRotZ <= -212.5 then Leftsling = False
  If Leftsling = True and Left3.TransZ > -23 then Left3.TransZ = Left3.TransZ - 4
  If Leftsling = False and Left3.TransZ < -0 then Left3.TransZ = Left3.TransZ + 4
  If Left3.TransZ <= -23 then Leftsling = False
End Sub

 Sub LeftSlingShot_Timer:Me.TimerEnabled = 0:Controller.Switch(26) = 0:End Sub

 Sub RightSlingShot_Slingshot
  Rightsling = True
  Controller.Switch(27) = 1
  PlaySound "slingshot":RightSlingshot.TimerEnabled = 1
  End Sub

 Dim Rightsling:Rightsling = False

Sub RS_Timer()
  If Rightsling = True and Right1.ObjRotZ > 7 then Right1.ObjRotZ = Right1.ObjRotZ - 2
  If Rightsling = False and Right1.ObjRotZ < 20 then Right1.ObjRotZ = Right1.ObjRotZ + 2
  If Right1.ObjRotZ <= 7 then Rightsling = False
  If Rightsling = True and Right2.ObjRotZ < 212.5 then Right2.ObjRotZ = Right2.ObjRotZ + 2
  If Rightsling = False and Right2.ObjRotZ > 199.5 then Right2.ObjRotZ = Right2.ObjRotZ - 2
  If Right2.ObjRotZ >= 212.5 then Rightsling = False
  If Rightsling = True and Right3.TransZ > -23 then Right3.TransZ = Right3.TransZ - 4
  If Rightsling = False and Right3.TransZ < -0 then Right3.TransZ = Right3.TransZ + 4
  If Right3.TransZ <= -23 then Rightsling = False
End Sub

 Sub RightSlingShot_Timer:Me.TimerEnabled = 0:Controller.Switch(27) = 0:End Sub



'''bumper rings init

Dim Bump1, Bump2, Bump3
' Ring1a.IsDropped = 1:Ring1b.IsDropped = 1
'   Ring2a.IsDropped = 1:Ring2b.IsDropped = 1
'   Ring3a.IsDropped = 1:Ring3b.IsDropped = 1



''''''''''''bumper animation
      ''''''''''''bumper animation
      Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySound "bumper1":bump1 = 1:Me.TimerEnabled = 1:End Sub

      Sub Bumper2_Hit:vpmTimer.PulseSw 49:PlaySound "bumper1":bump2 = 1:Me.TimerEnabled = 1:End Sub

      Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySound "bumper1":bump3 = 1:Me.TimerEnabled = 1:End Sub


''''flipper primitive
Sub UpdateFlipperLogos
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
    UFlogo.RotY = Flipper1.CurrentAngle
End Sub



Sub SolLFlipper(Enabled)
     If Enabled Then
    PlaySound "flipperup"
    LeftFlipper.RotateToEnd
    Flipper1.RotateToEnd
     Else
    PlaySound "flipperdown"
    LeftFlipper.RotateToStart
    Flipper1.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
     PlaySound "flipperup"
     RightFlipper.RotateToEnd
     Else
     PlaySound "flipperdown"
     RightFlipper.RotateToStart
    End If
 End Sub


dim MotorCallback
Set MotorCallback = GetRef("GameTimer")
Sub GameTimer
    UpdateFlipperLogos
End Sub


    ' Impulse Plunger
    Const IMPowerSetting = 40
    Const IMTime = 0.6
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Random 0.3
        .InitExitSnd "plunger2", "plunger"
        .CreateEvents "plungerIM"
    End With

Dim xx
Dim bsTrough,bsVUK,dtDrop,dtSingle,Mag,OMag,mMystery,PlungerIM
'Dim LStep, Rstep

Sub table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

LeftPost.IsDropped=1
RightPost.IsDropped=1
Post.IsDropped=1
LoopPost.IsDropped=1
  table1.Yieldtime=1
  Controller.GameName="shrkysht"
  Controller.SplashInfoLine="Sharkey's Shootout - Stern 2000" & vbnewline & "Table by freneticamnesic"
  Controller.ShowTitle=0
  Controller.ShowDMDOnly=1
  Controller.ShowFrame=0
  Controller.HandleMechanics=0
  Controller.HandleKeyboard=0

     'Controller.Games("shrkysht")
  On Error Resume Next
    Controller.Run
    If Err Then MsgBox Err.Description
  On Error Goto 0
  PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch=56
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)
    'vpmMapLights AllLights


  Set Mag=New cvpmMagnet
  Mag.InitMagnet MyMag,40
  Mag.CreateEvents "Mag"
  Mag.GrabCenter=True
  Mag.Solenoid=13

  Set OMag=New cvpmMagnet
  OMag.InitMagnet OrbitMag,25
  OMag.CreateEvents "OMag"
  OMag.GrabCenter=False
  OMag.Solenoid=12

    Set bsTrough = New cvpmBallStack
  bsTrough.InitSw 0,14,13,12,11,10,0,0
  bsTrough.InitKick BallRelease,90,3
  bsTrough.InitEntrySnd "Solenoid", "Solenoid"
  bsTrough.InitExitSnd "BallRel", "Solenoid"
  bsTrough.Balls=4

' Set bsVUK=New cvpmBallStack
' bsVUK.InitSaucer Kicker3,44,271,12
' bsVUK.InitExitSnd "popper", "popper"
' bsVUK.KickForceVar=2

    Set bsVUK = New cvpmBallStack
    With bsVUK
        .InitSw 0, 44, 0, 0, 0, 0, 0, 0
        .InitKick Kicker3, 271, 12
        .KickZ = 0.4
        .InitExitSnd "scoopexit", "scoopexit"
    .InitAddSnd "scoopenter"
        '.KickForceVar = 2
        .KickAngleVar = 2
        .KickBalls = 1
    End With

  set dtDrop=new cvpmDropTarget
  dtDrop.InitDrop Array(B1,B2,B3,B4,B5,B6,B7),Array(17,18,19,20,21,22,23)
  dtDrop.InitSnd "flapclos","flapopen"

  set dtSingle=new cvpmDropTarget
  dtSingle.InitDrop W8,24
  dtSingle.InitSnd "flapclos","flapopen"

  set mMystery=new cvpmMech
  mMystery.MType=vpmMechOneSol+vpmMechCircle+vpmMechLinear
  mMystery.Sol1=20
  mMystery.Length=120
  mMystery.Steps=8
  mMystery.Callback=GetRef("UpdateMystery")
  mMystery.Start

'Plunger1.PullBack
Controller.Switch(9)=1 'playfield glass
Controller.Switch(10)=1 'lockdown

  If GIOnDuringAttractMode = 1 Then GI_AllOn

End Sub

Sub Kicker3_Hit
    PlaySound "scoopenter"
    'ClearBallId
  Kicker3.DestroyBall
    bsVUK.AddBall 0
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

Sub UpdateMystery(aNewPos,aSpeed,aLastPos)
Select Case aNewPos
  Case 0:Controller.Switch(33)=1:Controller.Switch(34)=0:Controller.Switch(35)=0:Controller.Switch(36)=0'Hurry Up
  Case 1:Controller.Switch(33)=1:Controller.Switch(34)=1:Controller.Switch(35)=0:Controller.Switch(36)=0'Myst Mil's
  Case 2:Controller.Switch(33)=0:Controller.Switch(34)=1:Controller.Switch(35)=0:Controller.Switch(36)=0'Prize Money
  Case 3:Controller.Switch(33)=0:Controller.Switch(34)=1:Controller.Switch(35)=1:Controller.Switch(36)=0'Jackpot
  Case 4:Controller.Switch(33)=0:Controller.Switch(34)=0:Controller.Switch(35)=1:Controller.Switch(36)=0'Mystery
  Case 5:Controller.Switch(33)=0:Controller.Switch(34)=0:Controller.Switch(35)=1:Controller.Switch(36)=1'Multiball
  Case 6:Controller.Switch(33)=0:Controller.Switch(34)=0:Controller.Switch(35)=0:Controller.Switch(36)=1'Extra Ball
  Case 7:Controller.Switch(33)=1:Controller.Switch(34)=0:Controller.Switch(35)=0:Controller.Switch(36)=1'Post Save
End Select
End Sub

Sub PrimT_Timer
  If B1.isdropped = true then sw17f.rotatetoend:dtw1a.isdropped = true:end if
  if B1.isdropped = false then sw17f.rotatetostart:dtw1a.isdropped = false:end if
  If B2.isdropped = true then sw18f.rotatetoend:dtw2a.isdropped = true:dtw2b.isdropped = true:end if
  if B2.isdropped = false then sw18f.rotatetostart:dtw2a.isdropped = false:dtw2b.isdropped = false:end if
  If B3.isdropped = true then sw19f.rotatetoend:dtw3a.isdropped = true:dtw2b.isdropped = true:end if
  if B3.isdropped = false then sw19f.rotatetostart:dtw3a.isdropped = false:dtw2b.isdropped = false:end if
  If B4.isdropped = true then sw20f.rotatetoend:dtw4a.isdropped = true:dtw2b.isdropped = true:end if
  if B4.isdropped = false then sw20f.rotatetostart:dtw4a.isdropped = false:dtw2b.isdropped = false:end if
  If B5.isdropped = true then sw21f.rotatetoend:dtw5a.isdropped = true:dtw2b.isdropped = true:end if
  if B5.isdropped = false then sw21f.rotatetostart:dtw5a.isdropped = false:dtw2b.isdropped = false:end if
  If B6.isdropped = true then sw22f.rotatetoend:dtw6a.isdropped = true:dtw2b.isdropped = true:end if
  if B6.isdropped = false then sw22f.rotatetostart:dtw6a.isdropped = false:dtw2b.isdropped = false:end if
  If B7.isdropped = true then sw23f.rotatetoend:dtw7a.isdropped = true:end if
  if B7.isdropped = false then sw23f.rotatetostart:dtw7a.isdropped = false:end if
' If b1.isdropped = 1 then dtw1a.isdropped = 1 else dtw1a.isdropped = 0:end if
' If b2.isdropped = 1 then dtw2a.isdropped = 1:dtw2b.isdropped = 1 else dtw2a.isdropped = 0:dtw2b.isdropped = 0:end if
' If b3.isdropped = 1 then dtw3a.isdropped = 1:dtw3b.isdropped = 1 else dtw3a.isdropped = 0:dtw3b.isdropped = 0:end if
' If b4.isdropped = 1 then dtw4a.isdropped = 1:dtw4b.isdropped = 1 else dtw4a.isdropped = 0:dtw4b.isdropped = 0:end if
' If b5.isdropped = 1 then dtw5a.isdropped = 1:dtw5b.isdropped = 1 else dtw5a.isdropped = 0:dtw5b.isdropped = 0:end if
' If b6.isdropped = 1 then dtw6a.isdropped = 1 else dtw6a.isdropped = 0:end if
' If b7.isdropped = 1 then dtw7a.isdropped = 1 else dtw7a.isdropped = 0:end if
  If W8.isdropped = true then w8f.rotatetoend
  if W8.isdropped = false then w8f.rotatetostart
  sw17p.transy = sw17f.currentangle
  sw18p.transy = sw18f.currentangle
  sw19p.transy = sw19f.currentangle
  sw20p.transy = sw20f.currentangle
  sw21p.transy = sw21f.currentangle
  sw22p.transy = sw22f.currentangle
  sw23p.transy = sw23f.currentangle
  w8p.transy = w8f.currentangle
End Sub


Sub sw43_Hit
    sw43.DestroyBall
  bsVUK.AddBall Me
End Sub

'solcallback(23) = sol23flasher
'sub sol23flasher(Enabled)
'    If enabled then
'       SetLamp 123,1
'    else
'       SetLamp 123,0
'    end if
'End Sub

'Sub Drain_Hit:bsTrough.AddBall Me:End Sub            '11-15
'Sub swPlunger_Hit:Controller.Switch(16)=1:End Sub        '16
'Sub swPlunger_unHit:Controller.Switch(16)=0:End Sub
Sub sw16_Hit:Controller.Switch(16)=1:End Sub        '16
Sub sw16_unHit:Controller.Switch(16)=0:End Sub
Sub B1_Hit:dtDrop.Hit 1:End Sub                 '17
Sub B2_Hit:dtDrop.Hit 2:End Sub                 '18
Sub B3_Hit:dtDrop.Hit 3:End Sub                 '19
Sub B4_Hit:dtDrop.Hit 4:End Sub                 '20
Sub B5_Hit:dtDrop.Hit 5:End Sub                 '21
Sub B6_Hit:dtDrop.Hit 6:End Sub                 '22
Sub B7_Hit:dtDrop.Hit 7:End Sub                 '23
Sub W8_Hit:dtSingle.Hit 1:End Sub               '24
Sub W1_Hit:vpmTimer.PulseSw 25:End Sub              '25
Sub W2_Hit:vpmTimer.PulseSw 26:End Sub              '26
Sub W3_Hit:vpmTimer.PulseSw 27:End Sub              '27
Sub W4_Hit:vpmTimer.PulseSw 28:End Sub              '28
Sub W5_Hit:vpmTimer.PulseSw 29:End Sub              '29
Sub W6_Hit:vpmTimer.PulseSw 30:End Sub              '30
Sub LL_Hit:Controller.Switch(31)=1:End Sub            '31
Sub LL_unHit:Controller.Switch(31)=0:End Sub
Sub Trigger2_Hit:Controller.Switch(32)=1:End Sub        '32
Sub Trigger2_unHit:Controller.Switch(32)=0:End Sub
Sub LLoop_Spin:vpmTimer.PulseSw 38:End Sub          '38
Sub Trigger7_Hit:Controller.Switch(39)=1:End Sub        '39
Sub Trigger7_unHit:Controller.Switch(39)=0:End Sub
Sub Trigger3_Hit:Controller.Switch(40)=1:End Sub        '40
Sub Trigger3_unHit:Controller.Switch(40)=0:End Sub
Sub Trigger4_Hit:Controller.Switch(41)=1:End Sub        '41
Sub Trigger4_unHit:Controller.Switch(41)=0:End Sub
Sub Trigger5_Hit:Controller.Switch(42)=1:End Sub        '42
Sub Trigger5_unHit:Controller.Switch(42)=0:End Sub
Sub Trigger8_Hit:Controller.Switch(43)=1:End Sub        '43
Sub Trigger8_unHit:Controller.Switch(43)=0:End Sub
'Sub Kicker3_Hit:bsVUK.AddBall 0:End Sub              '44 SuperVUK
Sub Trigger1_Hit:Controller.Switch(45)=1:End Sub        '45
Sub Trigger1_unHit:Controller.Switch(45)=0:End Sub
Sub trigLaneA_Hit:Controller.Switch(46)=1:End Sub       '46
Sub trigLaneA_unHit:Controller.Switch(46)=0:End Sub
Sub trigLaneB_Hit:Controller.Switch(47)=1:End Sub       '47
Sub trigLaneB_unHit:Controller.Switch(47)=0:End Sub
Sub TrigLaneC_Hit:Controller.Switch(48)=1:End Sub       '48
Sub TrigLaneC_unHit:Controller.Switch(48)=0:End Sub
'Sub Bumper1_Hit:vpmTimer.PulseSw 49:Playsound "bumper1":End Sub            '49
'Sub Bumper2_Hit:vpmTimer.PulseSw 50:Playsound "bumper1":End Sub            '50
'Sub Bumper3_Hit:vpmTimer.PulseSw 51:Playsound "bumper1":End Sub            '51
Sub CTARGET_Hit:vpmTimer.PulseSw 53:End Sub           '53
'Sub LeftSlingshot_Slingshot:vpmTimer.PulseSw 59:End Sub      '59
Sub LeftOutlane_Hit:Controller.Switch(57)=1:End Sub       '57
Sub LeftOutlane_unHit:Controller.Switch(57)=0:End Sub
Sub LeftInlane_Hit:Controller.Switch(58)=1:End Sub        '58
Sub LeftInlane_unHit:Controller.Switch(58)=0:End Sub
Sub RightOutlane_Hit:Controller.Switch(60)=1:End Sub      '60
Sub RightOutlane_unHit:Controller.Switch(60)=0:End Sub
Sub RightInlane_Hit:Controller.Switch(61)=1:End Sub       '61
Sub RightInlane_unHit:Controller.Switch(61)=0:End Sub
'Sub RightSlingshot_Slingshot:vpmTimer.PulseSw 62:End Sub   '62


Sub Drain_Hit
  PlaySound "Drain"
  bsTrough.AddBall Me
  Drain.TimerInterval = 200
  Drain.TimerEnabled = 1
End Sub


Sub Drain_Timer
  'Debug.print GI_TroughCheck & " " & Lampstate(4) & " " & Ballsavelight
  If GI_TroughCheck = 4 And LampState(23) = 0 Then
    GI_AllOff 0
  End If
  Drain.TimerEnabled = False
End Sub

Sub BallRelease_UnHit(): GI_AllOn:End Sub





 '****************************************************************************************************************************************
 '  JP's Fading Lamps 3.4 VP9 Fading only
 '      Based on PD's Fading Lights
 ' SetLamp 0 is Off
 ' SetLamp 1 is On
 ' LampState(x) current state
 '****************************************
 
 Dim LampState(200), FadingLevel(200), FadingState(200)
 Dim FlashState(200), FlashLevel(1000)
 Dim FlashSpeedUp, FlashSpeedDown
 Dim x
 AllLampsOff()
 LampTimer.Interval = 30
 LampTimer.Enabled = 1
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

' Sub LampTimer_Timer()
'     Dim chgLamp, num, chg, ii
'     chgLamp = Controller.ChangedLamps
'     If Not IsEmpty(chgLamp) Then
'         For ii = 0 To UBound(chgLamp)
'             LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
'      FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
'      FadingState(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
'         Next
'     End If
' 
'     UpdateLamps
'   'UpdateLeds
' End Sub

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
    NFadeLm 1, l1b
    NFadeL 1, l1a
    NFadeLm 2, l2b
    NFadeL 2, l2a
    NFadeLm 3, l3b
    NFadeL 3, l3a
    NFadeLm 4, l4b
    NFadeL 4, l4a
    NFadeLm 5, l5b
    NFadeL 5, l5a
    NFadeLm 6, l6b
    NFadeL 6, l6a
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
'   NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    NFadeL 27, l27
    NFadeL 28, l28
    NFadeL 29, l29
    NFadeL 30, l30
    NFadeL 31, l31
    NFadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    NFadeL 42, l42
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
    NFadeL 58, l58
    NFadeL 59, l59
    NFadeL 60, l60
    NFadeL 61, l61
    NFadeL 62, l62
    NFadeL 63, l63
    NFadeL 64, l64
    NFadeL 65, l65
    NFadeL 66, l66
    NFadeL 67, l67
'   NFadeL 68, l68
'   NFadeL 69, l69
'   NFadeL 70, l70
'   NFadeL 71, l71
'   NFadeL 72, l72
'   NFadeL 73, l73
'   NFadeL 74, l74
'   NFadeL 75, l75
'   NFadeL 76, l76
'   NFadeL 77, l77
'   NFadeL 78, l78
'   NFadeL 79, l79
'   NFadeL 80, l80
''l1.State = LampState(1)
''l2.State = LampState(2)
''l3.State = LampState(3)
''l4.State = LampState(4)
''l5.State = LampState(5)
''l6.State = LampState(6)
'l7.State = LampState(7)
'l8.State = LampState(8)
'l9.State = LampState(9)
'l10.State = LampState(10)
'l11.State = LampState(11)
'l12.State = LampState(12)
'l13.State = LampState(13)
'l14.State = LampState(14)
'l15.State = LampState(15)
'l16.State = LampState(16)
'l17.State = LampState(17)
'l18.State = LampState(18)
'l19.State = LampState(19)
'l20.State = LampState(20)
'l21.State = LampState(21)
'l22.State = LampState(22)
'l23.State = LampState(23)
''l24.State = LampState(24)
'l25.State = LampState(25)
'l26.State = LampState(26)
'l27.State = LampState(27)
'l28.State = LampState(28)
'l29.State = LampState(29)
'l30.State = LampState(30)
'l31.State = LampState(31)
'l32.State = LampState(32)
'l33.State = LampState(33)
'l34.State = LampState(34)
'l35.State = LampState(35)
'l36.State = LampState(36)
'l37.State = LampState(37)
'l38.State = LampState(38)
'l39.State = LampState(39)
'l40.State = LampState(40)
'l41.State = LampState(41)
'l42.State = LampState(42)
'l43.State = LampState(43)
'l44.State = LampState(44)
'l45.State = LampState(45)
'l46.State = LampState(46)
'l47.State = LampState(47)
'l48.State = LampState(48)
'l49.State = LampState(49)
'l50.State = LampState(50)
'l51.State = LampState(51)
'l52.State = LampState(52)
'l53.State = LampState(53)
'l54.State = LampState(54)
''l55.State = LampState(55)
''l56.State = LampState(56)
'l57.State = LampState(57)
'l58.State = LampState(58)
'l59.State = LampState(59)
'l60.State = LampState(60)
'l61.State = LampState(61)
'l62.State = LampState(62)
'l63.State = LampState(63)
''l64.State = LampState(64)
'l65.State = LampState(65)
'l66.State = LampState(66)
'l67.State = LampState(67)
''l68.State = LampState(68)
''l69.State = LampState(69)
''l70.State = LampState(70)
''l71.State = LampState(71)
''l72.State = LampState(72)
''l73.State = LampState(73)
''l74.State = LampState(74)
''l75.State = LampState(75)
''l76.State = LampState(76)
''l77.State = LampState(77)
''l78.State = LampState(78)
''l79.State = LampState(79)
''l80.State = LampState(80)

' add your light inserts here
'the light named l9a needs textures off and step a
'the light named l9 needs textures b and full
'l9a is light state pff (off) and light state a (on)
'l9 is light state b (off) and light state full on (on)

'FlashARm 1, l1, "bf_on", "bf_a", "bf_b", Refresh 'blue
'FlashAR 1, l1b, "bf_on", "bf_a", "bf_b", Refresh 'blue
'FlashARm 2, l2, "bf_on", "bf_a", "bf_b", Refresh 'blue
'FlashAR 2, l2b, "bf_on", "bf_a", "bf_b", Refresh 'blue
'FlashARm 3, l3, "bf_on", "bf_a", "bf_b", Refresh 'blue
'FlashAR 3, l3b, "bf_on", "bf_a", "bf_b", Refresh 'blue
'FlashARm 4, l4, "bf_on", "bf_a", "bf_b", Refresh 'blue
'FlashAR 4, l4b, "bf_on", "bf_a", "bf_b", Refresh 'blue
'FlashARm 5, l5, "bf_on", "bf_a", "bf_b", Refresh 'blue
'FlashAR 5, l5b, "bf_on", "bf_a", "bf_b", Refresh 'blue
'FlashARm 6, l6, "bf_on", "bf_a", "bf_b", Refresh 'blue
'FlashAR 6, l6b, "bf_on", "bf_a", "bf_b", Refresh 'blue
'
'
'FadeL 7, l7a, l7
'FadeL 8, l8a, l8
'FadeL 9, l9a, l9
'FadeL 10, l10a, l10
'FadeL 11, l11a, l11
'FadeL 12, l12a, l12
'FadeL 13, l13a, l13
'FadeL 14, l14a, l14
'FadeL 15, l15a, l15
'FadeL 16, l16a, l16
'FadeL 17, l17a, l17
'FadeL 18, l18a, l18
'FadeL 19, l19a, l19
'FadeL 20, l20a, l20
'FadeL 21, l21a, l21
'FadeL 22, l22a, l22
'FadeL 23, l23a, l23
''FadeL 24, l24a, l24
'FadeL 25, l25a, l25
'FadeL 26, l26a, l26
'FadeL 27, l27a, l27
'FadeL 28, l28a, l28
'FadeL 29, l29a, l29
'FadeL 30, l30a, l30
'FadeL 31, l31a, l31
'FadeL 32, l32a, l32
'FadeL 33, l33a, l33
'FadeL 34, l34a, l34
'FadeL 35, l35a, l35
'FadeL 36, l36a, l36
'FadeL 37, l37a, l37
'FadeL 38, l38a, l38
'FadeL 39, l39a, l39
'FadeL 40, l40a, l40
'FadeL 41, l41a, l41
'FadeL 42, l42a, l42
'FadeL 43, l43a, l43
'FadeL 44, l44a, l44
'FadeL 45, l45a, l45
'FadeL 46, l46a, l46
'FadeL 47, l47a, l47
'FadeL 48, l48a, l48
'FadeL 49, l49a, l49
'FadeL 50, l50a, l50
'FadeL 51, l51a, l51
'FadeL 52, l52a, l52
'FadeL 53, l53a, l53
'FadeL 54, l54a, l54
''FlashAR 55, l55, "gf_on", "gf_a", "gf_b", Refresh 'green
''FlashAR 56, l56, "bf_on", "bf_a", "bf_b", Refresh 'blue
'FadeL 57, l57a, l57
'FadeL 58, l58a, l58
'FadeL 59, l59a, l59
'FadeL 60, l60a, l60
'FlashAR 61, f61, "bf_on", "bf_a", "bf_b", Refresh
'FlashAR 62, f62, "bf_on", "bf_a", "bf_b", Refresh
'FlashAR 63, f63, "bf_on", "bf_a", "bf_b", Refresh
''FlashAR 64, l64, "rf_on", "rf_a", "rf_b", Refresh 'red
'FadeL 65, l65a, l65
'FadeL 66, l66a, l66
'FadeL 67, l67a, l67

'FLASHERS '107 121 123
'PF Inserts 108 122 132
'FlashARm 107, f107, "bf_on", "bf_a", "bf_b", Refresh
'FlashAR 107, f107a, "bf_on", "bf_a", "bf_b", Refresh
'
'FlashARm 108, f108, "rf_on", "rf_a", "rf_b", Refresh 'red insert
'FlashAR 108, f108a, "rf_on", "rf_a", "rf_b", Refresh 'red insert
'FlashAR 118, f108b, "rf_on", "rf_a", "rf_b", Refresh 'red insert
'
'FlashARm 121, f121, "bf_on", "bf_a", "bf_b", Refresh
'FlashAR 121, f121a, "bf_on", "bf_a", "bf_b", Refresh
'
'FlashAR 122, f122, "rf_on", "rf_a", "rf_b", Refresh 'red insert
'
'FlashARm 123, f123, "mb_on", "mb_a", "mb_b", Refresh
'FlashAR 123, f123a, "mb_on", "mb_a", "mb_b", Refresh
'FlashAR 124, f123b, "mb_on", "mb_a", "mb_b", Refresh
'
'FlashAR 132, f132, "rf_on", "rf_a", "rf_b", Refresh 'red insert





End Sub

Sub Reflections_Timer()
if l57.state = 1 then r57.opacity = 25 else r57.opacity = 0
if l58.state = 1 then r58.opacity = 25 else r58.opacity = 0
if l59.state = 1 then r59.opacity = 25 else r59.opacity = 0
if l60.state = 1 then r60.opacity = 25 else r60.opacity = 0
'r60.opacity = l60.intensity
End Sub

Sub FlasherTimer_Timer()

Flashm 1, f1b
Flash 1, f1a
Flashm 2, f2b
Flash 2, f2a
Flashm 3, f3b
Flash 3, f3a
Flashm 4, f4b
Flash 4, f4a
Flashm 5, f5b
Flash 5, f5a
Flashm 6, f6b
Flash 6, f6a

'Flash 7, f7
'Flash 8, f8
'Flash 9, f9
'Flash 10, f10
'Flash 11, f11
'Flash 12, f12
'Flash 13, f13
'Flash 14, f14
'Flash 15, f15
'Flash 16, f16
'Flash 17, f17
'Flash 18, f18
'Flash 19, f19
'Flash 20, f20
Flash 121, f121
Flashm 107, f107a
Flash 107, f107b
'Flashm 107, f107b
'
'Flash 25, f25
'Flash 26, f26
'Flash 27, f27
'Flash 28, f28
'Flash 29, f29
'Flash 30, f30
'Flash 31, f31
'Flash 32, f32
'Flash 33, f33
'Flash 34, f34
'Flash 35, f35
'Flash 36, f36
'Flash 37, f37
'Flash 38, f38
'Flash 39, f39
'Flash 40, f40
'Flash 41, f41
'Flash 42, f42
'Flash 43, f43
'Flash 44, f44
'Flash 45, f45
'Flash 46, f46
'Flash 47, f47
'Flash 48, f48
'Flash 49, f49
'Flash 50, f50
'Flash 51, f51
'Flash 52, f52
'Flash 53, f53
Flash 55, f55
Flash 56, f56
'Flash 57, f57
Flash 61, f61
Flash 62, f62
Flash 63, f63
Flash 64, f64
'Flash 65, f65
'Flash 66, f66
'Flash 67, f67



 End Sub
 
 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:FadingLevel(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub
 
' Sub SetLamp(nr, value):LampState(nr) = abs(value) + 4:End Sub

Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
    LampState(nr) = abs(value) + 4
FadingLevel(nr ) = abs(value) + 4: FadingState(nr ) = abs(value) + 4
End Sub

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

Sub FadeL(nr, a, b)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:FadingLevel(nr) = 0
        Case 3:b.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:FadingLevel(nr) = 3
        Case 5:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

'Sub FadeL(nr, a, b)
'     Select Case LampState(nr)
'         Case 2:b.state = 0:LampState(nr) = 0
'         Case 3:b.state = 1:LampState(nr) = 2
'         Case 4:a.state = 0:LampState(nr) = 3
'         Case 5:b.state = 1:LampState(nr) = 6
'         Case 6:a.state = 1:LampState(nr) = 1
'     End Select
' End Sub


Sub FlashAR(nr, ramp, a, b, c, r)                                                          'used for reflections when there is no off ramp
    Select Case FadingState(nr)
        Case 2:ramp.opacity = 0:r.State = ABS(r.state -1):FadingState(nr) = 0                'Off
        Case 3:ramp.image = c:r.State = ABS(r.state -1):FadingState(nr) = 2                'fading...
        Case 4:ramp.image = b:r.State = ABS(r.state -1):FadingState(nr) = 3                'fading...
        Case 5:ramp.image = a:ramp.alpha = 1:r.State = ABS(r.state -1):FadingState(nr) = 1 'ON
    End Select
End Sub


Sub FlashARm(nr, ramp, a, b, c, r)
    Select Case FadingState(nr)
        Case 2:ramp.opacity = 0:r.State = ABS(r.state -1)
        Case 3:ramp.image = c:r.State = ABS(r.state -1)
        Case 4:ramp.image = b:r.State = ABS(r.state -1)
        Case 5:ramp.image = a:ramp.alpha = 1:r.State = ABS(r.state -1)
    End Select
End Sub


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
            If FlashLevel(nr) > 255 Then
                FlashLevel(nr) = 255
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
            Object.opacity = FlashLevel(nr)
        Case 1         ' on
            Object.opacity = FlashLevel(nr)
    End Select
End Sub




'PinWizard Plunger script
'Turn on the Plunger timer if analog plunger is detected

Dim AnPlNewPos, AnPlOldPos
AnPlNewPos = 0:AnPlOldPos = 0
If Plunger.MotionDevice> 0 Then
    PBWPlunger.Enabled = 1
End If

Sub PBWPlunger_Timer()
    If Plunger.MotionDevice> 0 Then
    'If Plunger.Position>24 then exit sub
        'Plungers(AnPlOldPos).WidthBottom = 0:Plungers(AnPlOldPos).WidthTop = 0
        AnPlNewPos = Plunger.Position \ 2
        'Plungers(AnPlNewPos).WidthBottom = 45:Plungers(AnPlNewPos).WidthTop = 45 'change to the width of the ramps
        'PRefresh.State = ABS(PRefresh.state -1)
        AnPlOldPos = AnPlNewPos
    End If
End Sub

Sub Table1_Exit
    Controller.Stop
End Sub

Sub DTLightTimer_Timer()
' If b1.isdropped = 1 then dtw1a.isdropped = 1 else dtw1a.isdropped = 0:end if
' If b2.isdropped = 1 then dtw2a.isdropped = 1:dtw2b.isdropped = 1 else dtw2a.isdropped = 0:dtw2b.isdropped = 0:end if
' If b3.isdropped = 1 then dtw3a.isdropped = 1:dtw3b.isdropped = 1 else dtw3a.isdropped = 0:dtw3b.isdropped = 0:end if
' If b4.isdropped = 1 then dtw4a.isdropped = 1:dtw4b.isdropped = 1 else dtw4a.isdropped = 0:dtw4b.isdropped = 0:end if
' If b5.isdropped = 1 then dtw5a.isdropped = 1:dtw5b.isdropped = 1 else dtw5a.isdropped = 0:dtw5b.isdropped = 0:end if
' If b6.isdropped = 1 then dtw6a.isdropped = 1 else dtw6a.isdropped = 0:end if
' If b7.isdropped = 1 then dtw7a.isdropped = 1 else dtw7a.isdropped = 0:end if
End Sub

''***** GI routines
'
Dim ballsavelight

Dim MultiballFlag


Sub GI_AllOff (time) 'Turn GI Off
  'debug.print "GI OFF " & time
  'UpdateGI 0,0
  'RampsOff
  GILightsOff
  If time > 0 Then
    GI_AllOnT.Interval = time
    GI_AllOnT.Enabled = 0
    GI_AllOnT.Enabled = 1
  End If
End Sub

Sub GI_AllOn 'Turn GI On
  'UpdateGI 0,8
  'RampsOn
  'FlippersOn
  GILightsOn
  'OptoOn
End Sub

Sub GI_AllOnT_Timer 'Turn GI On timer
  'UpdateGI 0,8
  'RampsOn
  'FlippersOn
  GILightsOn
  GI_AllOnT.Enabled = 0
End Sub

Function GI_TroughCheck
  Dim Ballcount:  Ballcount = 0
  If Controller.Switch(11) = TRUE then Ballcount = Ballcount + 1
    If Controller.Switch(12) = TRUE then Ballcount = Ballcount + 1
    If Controller.Switch(13) = TRUE then Ballcount = Ballcount + 1
    If Controller.Switch(14) = TRUE then Ballcount = Ballcount + 1
  If Ballcount < 3 Then 'Keep track of multiball mode
    MultiballFlag = 1
  Else
    MultiballFlag = 0
  End If

  GI_TroughCheck = Ballcount

  'debug.print "Troughcheck " & ballcount & " Multiball " & MultiballFlag

  If ballcount = 4 then
    GameOverTimerCheck.Enabled = 1 'no ball in play
    'debug.print timer & "Game Over?"
  Else
    GameOverTimerCheck.Enabled = 0 'ball in play
    'debug.print timer & "Game Not Over"
  End If

End Function

GameOverTimerCheck.Interval = 30000

Sub GameOverTimerCheck_Timer
  debug.print timer & "Game Over!"
  If GIOnDuringAttractMode = 1 Then GI_AllOn
  GameOverTimerCheck.Enabled = 0
End Sub

Sub GILightsOn
  UFLogo.image = "flipper-l2"
  LFLogo.image = "flipper-l2"
  RFLogo.image = "flipper-r2"
  Primitive66.image = "gold"
  Primitive20.image = "[chrome-gold]"
  Primitive21.image = "[chrome-gold]"
  Primitive22.image = "[chrome-gold]"
  Primitive91.image = "[chrome-gold]"
  Primitive90.image = "[chrome-gold]"
  Primitive127.image = "[chrome-gold]"
  Primitive128.image = "[chrome-gold]"
  Primitive129.image = "[chrome-gold]"
  Primitive130.image = "[chrome-gold]"
  Primitive132.image = "[chrome-gold]"
  Primitive133.image = "[chrome-gold]"
  Primitive147.image = "[chrome-gold]"
  Primitive63.image = "[chrome-gold]"
  Primitive121.image = "[chrome-gold]"
  Primitive124.image = "[chrome-gold]"
  Primitive148.image = "[chrome-gold]"
  Primitive67.image = "gold"
  Primitive151.image = "gold"
  Primitive153.image = "gold"

  'GI_Light.State = 1
  newlighttest1.state = 1
  newlighttest2.state = 1
  newlighttest3.state = 1
  newlighttest4.state = 1
  newlighttest5.state = 1
  newlighttest6.state = 1
  newlighttest7.state = 1
  newlighttest11.state = 1
  Light1.state = 1
  Light2.state = 1
  Light3.state = 1
  Light4.state = 1
  Light5.state = 1
  Light6.state = 1
  Light7.state = 1
  Light8.state = 1
  Light9.state = 1
  Light10.state = 1
  Light11.state = 1
  Light12.state = 1
  Light13.state = 1
  Light14.state = 1
  Light15.state = 1
  Light16.state = 1
  Light17.state = 1
  Light18.state = 1
  Light19.state = 1
  Light20.state = 1
  Light21.state = 1
  Light22.state = 1
  Light23.state = 1
  Light26.state = 1
  Light27.state = 1
  Light28.state = 1
  Light29.state = 1
  Light30.state = 1
  Light31.state = 1
  Light32.state = 1
  Light33.state = 1
  Light34.state = 1
  Light35.state = 1
  Light36.state = 1
  Light37.state = 1
  Light38.state = 1
  Light39.state = 1
  Light40.state = 1
  Light41.state = 1
  Light42.state = 1
  Light43.state = 1
  Light44.state = 1
  Light45.state = 1
  Light46.state = 1
  Light47.state = 1
  Light48.state = 1
  Light49.state = 1
  Light51.state = 1
  Light53.state = 1
  Light56.state = 1
  Light57.state = 1
  GILight1.state = 1
  GILight2.state = 1
  GILight3.state = 1
  GILight4.state = 1
  Light24.state = 1
  Light25.state = 1
  Light50.state = 1
  Light52.state = 1
End Sub

Sub GILightsOff
  UFLogo.image = "flipper-l2off"
  LFLogo.image = "flipper-l2off"
  RFLogo.image = "flipper-r2off"
  Primitive66.image = "gold-off"
  Primitive20.image = "[chrome-gold-off]"
  Primitive21.image = "[chrome-gold-off]"
  Primitive22.image = "[chrome-gold-off]"
  Primitive91.image = "[chrome-gold-off]"
  Primitive90.image = "[chrome-gold-off]"
  Primitive127.image = "[chrome-gold-off]"
  Primitive128.image = "[chrome-gold-off]"
  Primitive129.image = "[chrome-gold-off]"
  Primitive130.image = "[chrome-gold-off]"
  Primitive132.image = "[chrome-gold-off]"
  Primitive133.image = "[chrome-gold-off]"
  Primitive147.image = "[chrome-gold-off]"
  Primitive63.image = "[chrome-gold-off]"
  Primitive121.image = "[chrome-gold-off]"
  Primitive124.image = "[chrome-gold-off]"
  Primitive148.image = "[chrome-gold-off]"
  Primitive67.image = "gold-off"
  Primitive151.image = "gold-off"
  Primitive153.image = "gold-off"

  'GI_Light.State = 0
  newlighttest1.state = 0
  newlighttest2.state = 0
  newlighttest3.state = 0
  newlighttest4.state = 0
  newlighttest5.state = 0
  newlighttest6.state = 0
  newlighttest7.state = 0
  newlighttest11.state = 0
  Light1.state = 0
  Light2.state = 0
  Light3.state = 0
  Light4.state = 0
  Light5.state = 0
  Light6.state = 0
  Light7.state = 0
  Light8.state = 0
  Light9.state = 0
  Light10.state = 0
  Light11.state = 0
  Light12.state = 0
  Light13.state = 0
  Light14.state = 0
  Light15.state = 0
  Light16.state = 0
  Light17.state = 0
  Light18.state = 0
  Light19.state = 0
  Light20.state = 0
  Light21.state = 0
  Light22.state = 0
  Light23.state = 0
  Light26.state = 0
  Light27.state = 0
  Light28.state = 0
  Light29.state = 0
  Light30.state = 0
  Light31.state = 0
  Light32.state = 0
  Light33.state = 0
  Light34.state = 0
  Light35.state = 0
  Light36.state = 0
  Light37.state = 0
  Light38.state = 0
  Light39.state = 0
  Light40.state = 0
  Light41.state = 0
  Light42.state = 0
  Light43.state = 0
  Light44.state = 0
  Light45.state = 0
  Light46.state = 0
  Light47.state = 0
  Light48.state = 0
  Light49.state = 0
  Light51.state = 0
  Light53.state = 0
  Light56.state = 0
  Light57.state = 0
  GILight1.state = 0
  GILight2.state = 0
  GILight3.state = 0
  GILight4.state = 0
  Light24.state = 0
  Light25.state = 0
  Light50.state = 0
  Light52.state = 0
End Sub

Sub RLS_Timer()
              Spinner1p.RotX = -(Spinner1.currentangle) +90
              Spinner2p.RotX = -(Spinner2.currentangle) +90
'              sw43p.RotX = -(sw43s.currentangle) +90
'              sw44p.RotX = -(sw44s.currentangle) +90
'              sw50p.RotX = -(sw50s.currentangle) +90
'              sw53p.RotX = -(sw53s.currentangle) +90
'              sw51p.RotX = -(sw51s.currentangle) +90
'              leftrampp.RotX = -(leftramps.currentangle) +90
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Vol2(ball1, ball2) ' Calculates the Volume of the sound based on the speed of two balls
    Vol2 = (Vol(ball1) + Vol(ball2) ) / 2
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'    JP's VP10 Collision & Rolling Sounds
'*****************************************

Const tnob = 8 ' total number of balls
ReDim rolling(tnob)
ReDim collision(tnob)
Initcollision

Sub Initcollision
    Dim i
    For i = 0 to tnob
        collision(i) = -1
        rolling(i) = False
    Next
End Sub

Sub CollisionTimer_Timer()
    Dim BOT, B, B1, B2, dx, dy, dz, distance, radii
    BOT = GetBalls

    ' rolling

  For B = UBound(BOT) +1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
  Next

    If UBound(BOT) = -1 Then Exit Sub

    For B = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next

    'collision

    If UBound(BOT) < 1 Then Exit Sub

    For B1 = 0 to UBound(BOT)
        For B2 = B1 + 1 to UBound(BOT)
            dz = INT(ABS((BOT(b1).z - BOT(b2).z) ) )
            radii = BOT(b1).radius + BOT(b2).radius
      If dz <= radii Then

            dx = INT(ABS((BOT(b1).x - BOT(b2).x) ) )
            dy = INT(ABS((BOT(b1).y - BOT(b2).y) ) )
            distance = INT(SQR(dx ^2 + dy ^2) )

            If distance <= radii AND (collision(b1) = -1 OR collision(b2) = -1) Then
                collision(b1) = b2
                collision(b2) = b1
                PlaySound("fx_collide"), 0, Vol2(BOT(b1), BOT(b2)), Pan(BOT(b1)), 0, Pitch(BOT(b1)), 0, 0
            Else
                If distance > (radii + 10)  Then
                    If collision(b1) = b2 Then collision(b1) = -1
                    If collision(b2) = b1 Then collision(b2) = -1
                End If
            End If
      End If
        Next
    Next
End Sub
