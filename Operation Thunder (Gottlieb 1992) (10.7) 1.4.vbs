'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Operation Thunder                                                  ########
'#######          (Gottlieb 1992)                                                    ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.0 FS mfuegemann 2014
'
' Thanks to:
' Akiles5000 for all the resource images of the playfield and plastics
' Batch for creating the Desktop Backdrop image
' Zaphod, Destruk and Scapino for the Vari Target solution from their VP8 table
' Fuzzel for providing the images for the playfield lights
' Kodiac for Flipper Primitive routine
' JPSalas for the Flasher fading code
' Arngrim for the SoundFX code
' Thalamus, table should definately have a hole drop sounds implemented, else ssf seems ok

option Explicit
Dim VRPosterR
Dim VRPosterL
Dim VRLogo
Dim Glass
Dim GlassScratch

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'********************* VR AND FSS OPTIONS ******************************************
'***********************************************************************************
'VR Logo - Set VRLogo = 0 to turn off VR Logo.
VRLogo = 1

'VR Poster-Right - Set VRPosterR = 0 to turn off VR Poster-Right.
VRPosterR = 1

'VR Poster-Left - Set VRPosterL = 0 to turn off VR Poster-Left.
VRPosterL = 1

'VR table Glass - Set Glass = 0 to turn off VR playfield Glass.
Glass = 1

'VR Playfield glass Scratches - Set to 0 if you want to turn them off.
GlassScratch = 1
'***********************************************************************************

Dim VarRol,VarHidden, loopvar
If OperationThunder.ShowDT = true then
  VarRol=0
  VarHidden=1
  For Each loopvar in dtdisplay
    loopvar.Visible = True
  Next
  DisplayTimer.Enabled=True
Else
  VarRol=1
  VarHidden=0
  For Each loopvar in dtdisplay
    loopvar.Visible = False
  Next
  DisplayTimer.Enabled=False
  PinCab_Metal_L_R.Visible = 0
End If

If RenderingMode = 2 or OperationThunder.ShowFSS = -1 then
  VarRol=0
  VarHidden=1
  For Each loopvar in VRFSSsegments
    loopvar.Y = 43
  Next
  For Each loopvar in dtdisplay
    loopvar.Visible = False
  Next
  DisplayTimer.Enabled = False
  UpdateLedsF.Enabled = True
  If RenderingMode = 2 Then
    if VRPosterR = 1 then VR_Poster1R.visible = true:VR_Poster2R.visible = true Else VR_Poster1R.visible = false:VR_Poster2R.visible = false
    if VRPosterL = 1 then VR_Poster3L.visible = true:VR_Poster4L.visible = true Else VR_Poster3L.visible = false:VR_Poster4L.visible = false
    if VRLogo = 1 then VR_Logo.visible = true Else VR_Logo.visible = false
    if GlassScratch = 1 then GlassImpurities.visible = true:GlassImpurities1.visible = true Else GlassImpurities.visible = false:GlassImpurities1.visible = false
    if Glass = 1 then WindowGlass.visible = true Else WindowGlass.visible = false
  End If
  PinCab_Metal_L_R.Visible = 1
End If

If RenderingMode <> 2 Then
  For Each loopvar in Room
    loopvar.Visible = 0
  Next
End If

LoadVPM "01560000","gts3.vbs",3.2

'-----------------------------------
'------  Global Cofigurations ------
'-----------------------------------
Const cgamename = "opthund"   'name of ROM file to be used
Const DimGI= -20        'set between -100 and 105 to dim or brighten GI lights (minus is darker)
Const DimFlashers = -75     'set between -200 and 0 to dim or brighten Flasher lights (minus is darker)
Const FreePlay = 1        'set to 1 to add a coin on StartGameKey
Const UseBackboxGIRelay = 1   'set to 1 to use the backbox GI relay for additional playfield light effects

If B2SOn = true Then
  VarHidden=1
  For Each loopvar in dtdisplay
    loopvar.Visible = False
  Next
  DisplayTimer.Enabled = False
  If RenderingMode <> 2 and OperationThunder.ShowFSS<>-1 then UpdateLedsF.Enabled = False
End If


'******************* Options *********************
' DMD/Backglass Controller Setting
'Const cController = 1    '0=Use value defined in cController.txt, 1=VPinMAME, 2=UVP server, 3=B2S server, 4=B2S with DOF (disable VP mech sounds)
'*************************************************

'Dim cNewController
'Sub LoadVPM(VPMver, VBSfile, VBSver)
' Dim FileObj, ControllerFile, TextStr
'
' On Error Resume Next
' If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
' ExecuteGlobal GetTextFile(VBSfile)
' If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
'
' cNewController = 1
' If cController = 0 then
'   Set FileObj=CreateObject("Scripting.FileSystemObject")
'   If Not FileObj.FolderExists(UserDirectory) then
'     Msgbox "Visual Pinball\User directory does not exist. Defaulting to vPinMame"
'   ElseIf Not FileObj.FileExists(UserDirectory & "cController.txt") then
'     Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
'     ControllerFile.WriteLine 1: ControllerFile.Close
'   Else
'     Set ControllerFile=FileObj.GetFile(UserDirectory & "cController.txt")
'     Set TextStr=ControllerFile.OpenAsTextStream(1,0)
'     If (TextStr.AtEndOfStream=True) then
'       Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
'       ControllerFile.WriteLine 1: ControllerFile.Close
'     Else
'       cNewController=Textstr.ReadLine: TextStr.Close
'     End If
'   End If
' Else
'   cNewController = cController
' End If
'
' Select Case cNewController
'   Case 1
'     Set Controller = CreateObject("VPinMAME.Controller")
'     If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
'     If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
'     If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
'   Case 2
'     Set Controller = CreateObject("UltraVP.BackglassServ")
'   Case 3,4
'     Set Controller = CreateObject("B2S.Server")
' End Select
' On Error Goto 0
'End Sub

'*************************************************************
'Toggle DOF sounds on/off based on cController value
'*************************************************************
'Dim ToggleMechSounds
'Function SoundFX (sound)
'    If cNewController= 4 and ToggleMechSounds = 0 Then
'        SoundFX = ""
'    Else
'        SoundFX = sound
'    End If
'End Function

'Sub DOF(dofevent, dofstate)
' If cController>2 Then
'   If dofstate = 2 Then
'     Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
'   Else
'     Controller.B2SSetData dofevent, dofstate
'   End If
' End If
'End Sub

Const UseSolenoids=1,UseLamps=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="",SFlipperOff="",SCoin="coin3"

'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------

'Sol1 Left Bumper
'Sol2 Right Bumper
'Sol3 Left Slingshot
'Sol4 Right Slingshot
SolCallback(5) = "bsRightKicker.Solout"     'Hole
SolCallback(6) = "DropTarget3Bank.SolDropUp"  'Left Drop Target 3 Bank
SolCallback(7) = "DropTarget3Bank.SolDropDown"  '#1 Trip, #2 Trip, #3 Trip
SolCallback(8) = "DropTarget5Bank.SolDropUp"  'Right Drop Target 5 Bank
SolCallback(9) = "LeaveLeftVUK"         'Left VUK
SolCallback(10) = "LeaveCenterVUK"        'Center VUK
SolCallback(11) = "SolRightVUK"         'Right VUK
SolCallback(12) = "SolLeftRampGate"       'Left Plunger Gate
SolCallback(13) = "SolRightRampGate"      'Right Plunger Gate
SolCallback(14) = "SolRightLaneFlasher"     'Right Lane Flasher
SolCallback(15) = "SolLeftLaneFlasher"      'Left Lane Flasher
SolCallback(16) = "SolLeftRampFlasher"      'Left Ramp Flasher
SolCallback(17) = "SolTopRightFlasher"      'Top Right Flasher
SolCallback(18) = "SolLeftMountainFlasherTop" 'Left Mountain Flasher
SolCallback(19) = "SolLeftMountainFlasherMid" 'Left Mountain Flasher
SolCallback(20) = "SolLeftMountainFlasherBot" 'Left Mountain Flasher
SolCallback(21) = "SolTopMountainFlasher"   'Top Mountain Flasher
SolCallback(22) = "SolRightMountainFlasherTop"  'Right Mountain Flasher
SolCallback(23) = "SolRightMountainFlasherBot"  'Right Mountain Flasher
SolCallback(24) = "SolVariTargetReset"      'Vari Target Reset
SolCallback(25) = "SolSpinningDisk"       'Spinning Disk
SolCallback(26) = "SolGIBackbox"        'Lightbox insert
'Sol27 Coin Lock
SolCallback(28) = "bsTrough.SolOut"       'BallRelease
SolCallback(29) = "bsTrough.SolIn"        'Outhole
SolCallback(30)  = "vpmSolSound SoundFX(""Knocker"",DOFKnocker)," 'Knocker
'Sol31 Tilt
SolCallback(32)= "SolGameOver"          'Game Over
SolCallback(38)= "SolLeftBumperLight"     'Left Bumper Light
SolCallback(39)= "SolRightBumperLight"      'Right Bumper Light

SolCallback(sllflipper) = "vpmsolflipper leftflipper,nothing,"
SolCallback(slrflipper) = "vpmsolflipper rightflipper,urightflipper,"

Sub SolRightVUK(enabled)
  If Enabled Then
    If bsRightVUK.Balls Then
      RightVUK.Timerenabled = True
      bsRightVUK.ExitSol_On
    End If
  End If
End Sub

Sub SolLeftRampGate(enabled)
  LeftRampLock.isdropped = enabled
  PlaySoundAtVol "solon", LeftRampLock, 1
End Sub

Sub SolRightRampGate(enabled)
  RightRampLock.isdropped = enabled
  PlaySoundAtVol "solon", RightRampLock, 1
End Sub

Sub SolGIBackbox(enabled)
  if (UseBackboxGIRelay = 1) and (not GIStarttimer.enabled) then
    For each obj in GIFlashers
      obj.visible = not enabled
    Next
    vrbggi.Visible = not enabled
  end if
End Sub

Sub SolGameOver(enabled)
  vpmNudge.SolGameOn enabled
  Flipperactive = enabled
end sub

Sub SolTopRightFlasher(enabled)
  if enabled then
    P_Flasher17.image = "Flasher_white"
    setflash 0, True
  else
    P_Flasher17.image = "Flasher_white_dark"
    setflash 0, False
  end if
end Sub
Sub SolLeftMountainFlasherTop(enabled)
  if enabled then
    P_Flasher18.image = "Flasher_white"
    setflash 1, True
  else
    P_Flasher18.image = "Flasher_white_dark"
    setflash 1, False
  end if
end Sub
Sub SolLeftMountainFlasherMid(enabled)
  if enabled then
    P_Flasher19.image = "Flasher_white"
    setflash 2, True
  else
    P_Flasher19.image = "Flasher_white_dark"
    setflash 2, False
  end if
end Sub
Sub SolLeftMountainFlasherBot(enabled)
  if enabled then
    P_Flasher20.image = "Flasher_white"
    setflash 3, True
  else
    P_Flasher20.image = "Flasher_white_dark"
    setflash 3, False
  end if
end Sub
Sub SolTopMountainFlasher(enabled)
  if enabled then
    P_Flasher21.image = "Flasher_white"
    setflash 4, True
  else
    P_Flasher21.image = "Flasher_white_dark"
    setflash 4, False
  end if
end Sub
Sub SolRightMountainFlasherTop(enabled)
  if enabled then
    P_Flasher22.image = "Flasher_white"
    setflash 5, True
  else
    P_Flasher22.image = "Flasher_white_dark"
    setflash 5, False
  end if
end Sub
Sub SolRightMountainFlasherBot(enabled)
  if enabled then
    P_Flasher23.image = "Flasher_white"
    setflash 6, True
  else
    P_Flasher23.image = "Flasher_white_dark"
    setflash 6, False
  end if
end Sub
Sub SolRightLaneFlasher(enabled)
  if enabled then
    setflash 7, True
    Flasher14Bulb.state = lightstateOn
  else
    setflash 7, False
    Flasher14Bulb.state = lightstateOff
  end if
end Sub
Sub SolLeftLaneFlasher(enabled)
  if enabled then
    setflash 8, True
    Flasher15Bulb.state = lightstateOn
  else
    setflash 8, False
    Flasher15Bulb.state = lightstateOff
  end if
end Sub
Sub SolLeftRampFlasher(enabled)
  if enabled then
    setflash 9, True
  else
    setflash 9, False
  end if
end Sub


'--------------------------
'------  Table Init  ------
'--------------------------
Dim bsTrough,obj,bsRightKicker,DropTarget3Bank,DropTarget5Bank,Flipperactive,TTable,bsLeftVUK,bsCenterVUK,bsRightVUK
Const swStartButton=4   'thanks to Destruk/Zaphod

Sub OperationThunder_Init
  vpminit me

  Flipperactive = False

    Controller.GameName=cGameName
    Controller.SplashInfoLine="Operation Thunder" & vbNewLine & "created by mfuegemann" & vbNewLine & "VPX conversion by Rascal"
    Controller.HandleKeyboard=False
    Controller.ShowTitle=0
  Controller.Hidden = VarHidden
  Controller.Games(cGameName).Settings.Value("rol")=VarRol
    Controller.ShowFrame=0
    Controller.ShowDMDOnly=1
  'Controller.Hidden = 1      'enable to hide DMD if You use a B2S backglass

    'DMD position for 3 Monitor Setup
    'Controller.Games(cGameName).Settings.Value("dmd_pos_x")=3850   'set this to 0 if You cannot find the DMD
    'Controller.Games(cGameName).Settings.Value("dmd_pos_y")=300    'set this to 0 if You cannot find the DMD
    'Controller.Games(cGameName).Settings.Value("dmd_width")=505
    'Controller.Games(cGameName).Settings.Value("dmd_height")=155
    'Controller.Games(cGameName).Settings.Value("rol")=0

  'Controller.Games(cGameName).Settings.Value("ddraw") = 0             'set to 0 if You have problems with DMD showing or table stutter

    Controller.HandleMechanics=0
    Controller.Run
    If Err Then MsgBox Err.Description
    On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled = true

    vpmNudge.TiltSwitch=151
    vpmNudge.Sensitivity=5
  vpmNudge.TiltObj = Array(LeftFlipper,RightFlipper,URightFlipper,Bumper10,Bumper11)

    vpmMapLights AllLights

    Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 14,0,0,24,0,0,0,0
        bsTrough.InitKick BallRelease,90,5
        bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors),SoundFX("Solenoid",DOFContactors)
        bsTrough.Balls=2

  bsTrough.AddBall 0

    ' Thalamus - more randomness to kickers pls
    Set bsRightKicker=New cvpmBallStack
        bsRightKicker.KickForceVar = 3
        bsRightKicker.KickAngleVar = 3
        bsRightKicker.InitSaucer RightKicker,34,180,5
        bsRightKicker.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("solon",DOFContactors)

  Set bsRightVUK=New cvpmBallStack
    bsRightVUK.InitSw 0,40,0,0,0,0,0,0
    bsRightVUK.InitKick RightVUK,245,8
    bsRightVUK.KickForceVar = 3
    bsRightVUK.KickAngleVar = 3
    bsRightVUK.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("Solenoid",DOFContactors)

  set DropTarget3Bank = new cvpmDropTarget
    DropTarget3Bank.InitDrop Array(Target22,Target32,Target42), Array(22,32,42)
    DropTarget3Bank.InitSnd SoundFX("Targetdrop1",DOFDropTargets),SoundFX("TargetBankreset1",DOFDropTargets)
    DropTarget3Bank.CreateEvents "DropTarget3Bank"

  set DropTarget5Bank = new cvpmDropTarget
    DropTarget5Bank.InitDrop Array(Target5,Target15,Target25,Target35,Target45), Array(5,15,25,35,45)
    DropTarget5Bank.InitSnd SoundFX("Targetdrop1",DOFDropTargets),SoundFX("TargetBankreset1",DOFDropTargets)
    DropTarget5Bank.CreateEvents "DropTarget5Bank"

  Set TTable=New cvpmTurnTable
    TTable.InitTurnTable TTableTrigger,30
    TTable.SpinUp = 15
    TTable.SpinDown = 15
    TTable.spinCW = False
    TTable.CreateEvents "TTable"

  SolVariTargetReset(True)

  if (DimGI >= -100) and (DimGI <= 105) then
    For each obj in GIFlashers
      obj.Opacity = obj.Opacity + DimGI
    Next
  end if

  GIStartTimer.enabled = True

End Sub

Sub OperationThunder_Exit()
  Controller.Pause = False
  Controller.Stop
  SaveLUT 'add this line in the sub OperationThunder_exit
End Sub


'------------------------------
'------  Trough Handler  ------
'------------------------------
Sub Drain_Hit()
  StopAllSounds
  SoundTimer.enabled = False
  bsTrough.AddBall Me
  PlaySoundAtVol "Drain5", Drain, 1
End Sub

'-------------------------------
'------  Keybord Handler  ------
'-------------------------------

Sub OperationThunder_KeyDown(ByVal keycode)
  if (keycode = StartGamekey) and (FreePlay = 1) then
    vpmtimer.pulsesw 2    'Free Play: add Coin on Game start
    VR_Cab_StartButton.y = VR_Cab_StartButton.y -5
  End If
  If keycode = PlungerKey Then
    Plunger.PullBack
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If
  if Flipperactive then
    If keycode = LeftFlipperKey Then
      VR_CabFlipperLeft.X = VR_CabFlipperLeft.X +10
      Controller.Switch(6)=1
      PlaySoundAtVol SoundFX("FlipperUp_Akiles",DOFFlippers), LeftFlipper, 1
    End If
    If keycode = RightFlipperKey Then
      VR_CabFlipperRight.X = VR_CabFlipperRight.X - 10
      Controller.Switch(7)=1
      PlaySoundAtVol SoundFX("FlipperUp_Akiles",DOFFlippers), RightFlipper, 1
    End If
  end if
  'LUT-Changer
  If Keycode = LeftMagnaSave Then
        LUTSet = LUTSet  + 1
    if LutSet > 15 then LUTSet = 0
        lutsetsounddir = 1
      If LutToggleSound then
        If lutsetsounddir = 1 And LutSet <> 15 Then
          Playsound "click", 0, 1, 0.1, 0.25
        End If
      If lutsetsounddir = -1 And LutSet <> 15 Then
          Playsound "click", 0, 1, 0.1, 0.25
        End If
      If LutSet = 15 Then
          Playsound "gun", 0, 1, 0.1, 0.25
        End If
        LutSlctr.enabled = true
      SetLUT
      ShowLUT
    End If
  End If

  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub OperationThunder_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
    VR_Primary_plunger.Y = 2335
  End If
  if Flipperactive then
    If keycode = LeftFlipperKey Then
      VR_CabFlipperLeft.X = VR_CabFlipperLeft.X -10
      Controller.Switch(6)=0
      PlaySoundAtVol SoundFX("FlipperDown_Akiles",DOFFlippers), LeftFlipper, 1
    End If
    If keycode = RightFlipperKey Then
      VR_CabFlipperRight.X = VR_CabFlipperRight.X +10
      Controller.Switch(7)=0
      PlaySoundAtVol SoundFX("FlipperDown_Akiles",DOFFlippers), RightFlipper, 1
    End If
  end if
  If Keycode = StartGameKey Then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y +5
  End If

  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub


'--------------------------
'------  Turn Table  ------
'--------------------------

Dim CurrSpeed
Sub TurnTableTimer_Timer
  CurrSpeed = TTable.speed * 15 / TTable.maxspeed

  P_TurnTable.objRotZ = P_TurnTable.objRotZ + CurrSpeed

  if P_TurnTable.objRotZ >= 360 then
    P_TurnTable.objRotZ = 0
  end if
  if P_TurnTable.objRotZ < 0 then
    P_TurnTable.objRotZ = 360 + P_TurnTable.objRotZ
  end if
End Sub

Sub SolSpinningDisk(enabled)
  TTable.MotorOn = enabled
End Sub


'---------------------------
'------  Vari Target  ------
'---------------------------

Dim Resistance
Resistance=6

Sub VT33_1_Hit
  If ActiveBall.VelY<0 Then
    VT33_1_Wall.IsDropped = True
    VT33_2_Wall.IsDropped = False
    Controller.Switch(33) = 1
    ActiveBall.VelY = ActiveBall.VelY + Resistance
  End if
End Sub

Sub VT33_2_Hit
  If ActiveBall.VelY<0 Then
    VT33_2_Wall.IsDropped = True
    VT33_3_Wall.IsDropped = False
    ActiveBall.VelY = ActiveBall.VelY + Resistance
  End If
End Sub

Sub VT33_3_Hit
  If ActiveBall.VelY<0 Then
    VT33_3_Wall.IsDropped = True
    VT43_1_Wall.IsDropped = False
    ActiveBall.VelY = ActiveBall.VelY + Resistance
  End If
End Sub

Sub VT43_1_Hit
  If ActiveBall.VelY<0 Then
    VT43_1_Wall.IsDropped = True
    VT43_2_Wall.IsDropped = False
    Controller.Switch(43) = 1
    ActiveBall.VelY = ActiveBall.VelY + Resistance
  End if
End Sub

Sub VT43_2_Hit
  If ActiveBall.VelY<0 Then
    VT43_2_Wall.IsDropped = True
    VT43_3_Wall.IsDropped = False
    ActiveBall.VelY = ActiveBall.VelY + Resistance
  End If
End Sub

Sub SolVariTargetReset(Enabled)
  If Enabled Then
    Controller.Switch(33)=0
    Controller.Switch(43)=0
    VT33_1_Wall.IsDropped = False
    VT33_2_Wall.IsDropped = True
    VT33_3_Wall.IsDropped = True
    VT43_1_Wall.IsDropped = True
    VT43_2_Wall.IsDropped = True
    VT43_3_Wall.IsDropped = True
  End If
End Sub

'-----------------------------
'------  VUK animation  ------
'-----------------------------

Sub LeftVUK_Hit
  Controller.Switch(20) = True
  StopAllSounds
  SoundTimer.enabled = False
End Sub

Sub LeaveLeftVUK(Enabled)
  If Enabled Then
      If Controller.Switch(20) = True Then
    PlaySoundAtVol "Popper", LeftVUK, 1
    Controller.Switch(20) = False
    LeftVUK.destroyball
    LeftVUK1.CreateBall
    vpmTimer.AddTimer 70,"LeftVUKLevel1"
    end if
  end if
end sub

Sub LeftVUKLevel1(swNo)
  LeftVUK1.DestroyBall
  LeftVUK2.CreateBall
  vpmTimer.AddTimer 70,"LeftVUKLevel2"
End Sub

Sub LeftVUKLevel2(swNo)
  LeftVUK2.DestroyBall
  LeftVUK3.CreateBall
  vpmTimer.AddTimer 70,"LeftVUKLevel3"
End Sub

Sub LeftVUKLevel3(swNo)
  LeftVUK3.DestroyBall
  LeftVUKTop.CreateBall
  LeftVUKTop.Kick 170,3
End Sub

Sub CenterVUK_Hit
  Controller.Switch(30) = True
  StopAllSounds
  SoundTimer.enabled = False
End Sub

Sub LeaveCenterVUK(Enabled)
  If Enabled Then
      If Controller.Switch(30) = True Then
    PlaySoundAtVol "Popper", CenterVUK, 1
    Controller.Switch(30) = False
    CenterVUK.destroyball
    CenterVUK1.CreateBall
    vpmTimer.AddTimer 70,"CenterVUKLevel1"
    end if
  end if
end sub

Sub CenterVUKLevel1(swNo)
  CenterVUK1.DestroyBall
  CenterVUK2.CreateBall
  vpmTimer.AddTimer 70,"CenterVUKLevel2"
End Sub

Sub CenterVUKLevel2(swNo)
  CenterVUK2.DestroyBall
  CenterVUK3.CreateBall
  vpmTimer.AddTimer 70,"CenterVUKLevel3"
End Sub

Sub CenterVUKLevel3(swNo)
  CenterVUK3.DestroyBall
  CenterVUKTop.CreateBall
  CenterVUKTop.Kick 170,3
End Sub


'------------------------------
'------  Switch Handler  ------
'------------------------------

Sub LeftSlingshot_Slingshot:PlaySoundAtVol SoundFX("LSling", DOFContactors), ActiveBall, 1:vpmTimer.PulseSw 12:End Sub
Sub RightSlingshot_Slingshot:PlaySoundAtVol SoundFX("RSling", DOFContactors), ActiveBall, 1:vpmTimer.PulseSw 13:End Sub

Sub Spinner44_Spin:vpmTimer.PulseSwitch 44,0,0:PlaySoundAtVol "soloff", ActiveBall, 1:End Sub

Sub Bumper10_hit
  vpmTimer.PulseSw 10
  PlaySoundAtVol SoundFX("bumper",DOFContactors), ActiveBall, 1
End Sub
Sub Bumper11_hit
  vpmTimer.PulseSw 11
  PlaySoundAtVol SoundFX("bumper",DOFContactors), ActiveBall, 1
End Sub

Sub RightKicker_Hit
  StopAllSounds
  SoundTimer.enabled = False
  bsRightKicker.addball me
end Sub

Sub RightVUK_Hit
  StopAllSounds
  SoundTimer.enabled = False
  bsRightVUK.addball me
end Sub

Sub UndergroundHole_Hit
  StopAllSounds
  SoundTimer.enabled = False
  vpmTimer.PulseSwitch 66,0,0
  UndergroundHole.destroyball
  'ActiveBall.Z = ActiveBall.Z - 60
  UndergroundHole.Timerenabled = True
  PlaySoundAtVol "Drain6", UndergroundHole, 1
end Sub
Sub UndergroundHole_Timer
  'UndergroundHole.destroyball
  UndergroundHole.Timerenabled = False
  bsRightVUK.addball 0
End Sub


sub Trigger46_hit:Controller.Switch(46)=1:P_T46.TransZ=-8:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub Trigger46_unhit:Controller.Switch(46)=0:Trigger46.timerenabled=True:End Sub

sub Trigger65_hit:Controller.Switch(65)=1:P_T65.TransZ=-8:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub Trigger65_unhit:Controller.Switch(65)=0:Trigger65.timerenabled=True:End Sub

sub Trigger70_hit:Controller.Switch(70)=1::PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub Trigger70_unhit:Controller.Switch(70)=0:End Sub

sub Trigger60_hit:Controller.Switch(60)=1:P_T60.TransZ=-8:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub Trigger60_unhit:Controller.Switch(60)=0:Trigger60.timerenabled=True:End Sub
sub Trigger62_hit:Controller.Switch(62)=1:P_T62.TransZ=-8:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub Trigger62_unhit:Controller.Switch(62)=0:Trigger62.timerenabled=True:End Sub
sub Trigger63_hit:Controller.Switch(63)=1:P_T63.TransZ=-8:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub Trigger63_unhit:Controller.Switch(63)=0:Trigger63.timerenabled=True:End Sub
sub Trigger61_hit:Controller.Switch(61)=1:P_T61.TransZ=-8:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub Trigger61_unhit:Controller.Switch(61)=0:Trigger61.timerenabled=True:End Sub

sub Trigger67_hit:Controller.Switch(67)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub Trigger67_unhit:Controller.Switch(67)=0:End Sub

Sub Target22_Hit:DropTarget3Bank.Hit 1:End Sub
Sub Target32_Hit:DropTarget3Bank.Hit 2:End Sub
Sub Target42_Hit:DropTarget3Bank.Hit 3:End Sub

Sub Target5_Hit:DropTarget5Bank.Hit 1:End Sub
Sub Target15_Hit:DropTarget5Bank.Hit 2:End Sub
Sub Target25_Hit:DropTarget5Bank.Hit 3:End Sub
Sub Target35_Hit:DropTarget5Bank.Hit 4:End Sub
Sub Target45_Hit:DropTarget5Bank.Hit 5:End Sub

Sub Target21_Hit:vpmTimer.PulseSw 21:MoveTarget21:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
Sub Target31_Hit:vpmTimer.PulseSw 31:MoveTarget31:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
Sub Target41_Hit:vpmTimer.PulseSw 41:MoveTarget41:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub

Sub Target16_Hit:vpmTimer.PulseSw 16:MoveTarget16:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
Sub Target26_Hit:vpmTimer.PulseSw 26:MoveTarget26:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
Sub Target36_Hit:vpmTimer.PulseSw 36:MoveTarget36:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
Sub Target17_Hit:vpmTimer.PulseSw 17:MoveTarget17:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
Sub Target27_Hit:vpmTimer.PulseSw 27:MoveTarget27:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
Sub Target37_Hit:vpmTimer.PulseSw 37:MoveTarget37:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub


'--------------------------------
'------  Helper Functions  ------
'--------------------------------

Sub Gate2_Hit:PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub

'Rollover Switches
sub Trigger46_Timer
  Trigger46.timerenabled = False
  P_T46.TransZ = 0
End Sub
sub Trigger65_Timer
  Trigger65.timerenabled = False
  P_T65.TransZ = 0
End Sub

sub Trigger60_Timer
  Trigger60.timerenabled = False
  P_T60.TransZ = 0
End Sub
sub Trigger61_Timer
  Trigger61.timerenabled = False
  P_T61.TransZ = 0
End Sub
sub Trigger62_Timer
  Trigger62.timerenabled = False
  P_T62.TransZ = 0
End Sub
sub Trigger63_Timer
  Trigger63.timerenabled = False
  P_T63.TransZ = 0
End Sub

'Round Targets
Sub MoveTarget21
  P_Target21.Transy = 5
  Target21.Timerenabled = False
  Target21.Timerenabled = True
End Sub
Sub Target21_Timer
  Target21.Timerenabled = False
  P_Target21.Transy = 0
End Sub

Sub MoveTarget31
  P_Target31.Transy = 5
  Target31.Timerenabled = False
  Target31.Timerenabled = True
End Sub
Sub Target31_Timer
  Target31.Timerenabled = False
  P_Target31.Transy = 0
End Sub

Sub MoveTarget41
  P_Target41.Transy = 5
  Target41.Timerenabled = False
  Target41.Timerenabled = True
End Sub
Sub Target41_Timer
  Target41.Timerenabled = False
  P_Target41.Transy = 0
End Sub

Sub MoveTarget16
  P_Target16.Transy = 5
  Target16.Timerenabled = False
  Target16.Timerenabled = True
End Sub
Sub Target16_Timer
  Target16.Timerenabled = False
  P_Target16.Transy = 0
End Sub

Sub MoveTarget26
  P_Target26.Transy = 5
  Target26.Timerenabled = False
  Target26.Timerenabled = True
End Sub
Sub Target26_Timer
  Target26.Timerenabled = False
  P_Target26.Transy = 0
End Sub

Sub MoveTarget36
  P_Target36.Transy = 5
  Target36.Timerenabled = False
  Target36.Timerenabled = True
End Sub
Sub Target36_Timer
  Target36.Timerenabled = False
  P_Target36.Transy = 0
End Sub

Sub MoveTarget17
  P_Target17.Transy = 5
  Target17.Timerenabled = False
  Target17.Timerenabled = True
End Sub
Sub Target17_Timer
  Target17.Timerenabled = False
  P_Target17.Transy = 0
End Sub

Sub MoveTarget27
  P_Target27.Transy = 5
  Target27.Timerenabled = False
  Target27.Timerenabled = True
End Sub
Sub Target27_Timer
  Target27.Timerenabled = False
  P_Target27.Transy = 0
End Sub

Sub MoveTarget37
  P_Target37.Transy = 5
  Target37.Timerenabled = False
  Target37.Timerenabled = True
End Sub
Sub Target37_Timer
  Target37.Timerenabled = False
  P_Target37.Transy = 0
End Sub

'additional Lamp Callback
Sub LampTimer_Timer
  if Controller.lamp(47) then
    P_Bulb47.image = "BulbRed_lit"
    P_Bulb47.disablelighting = True
    B47Flasher.visible = True
  else
    P_Bulb47.image = "BulbRed"
    P_Bulb47.disablelighting = False
    B47Flasher.visible = False
  end if
  if Controller.lamp(56) then
    P_Bulb56.image = "BulbRed_lit"
    P_Bulb56.disablelighting = True
    B56Flasher.visible = True
  else
    P_Bulb56.image = "BulbRed"
    P_Bulb56.disablelighting = False
    B56Flasher.visible = False
  end if
  if Controller.lamp(57) then
    P_Bulb57.image = "BulbBlue_lit"
    P_Bulb57.disablelighting = True
    B57Flasher.visible = True
  else
    P_Bulb57.image = "BulbBlue"
    P_Bulb57.disablelighting = False
    B57Flasher.visible = False
  end if
  if Controller.lamp(66) then
    P_Bulb66.image = "BulbRed_lit"
    P_Bulb66.disablelighting = True
    B66Flasher.visible = True
  else
    P_Bulb66.image = "BulbRed"
    P_Bulb66.disablelighting = False
    B66Flasher.visible = False
  end if
  if Controller.lamp(67) then
    P_Bulb67.image = "BulbGreen_lit"
    P_Bulb67.disablelighting = True
    B67Flasher.visible = True
  else
    P_Bulb67.image = "BulbGreen"
    P_Bulb67.disablelighting = False
    B67Flasher.visible = False
  end if
  if Controller.lamp(87) then
    P_Bulb87.image = "BulbRed_lit"
    P_Bulb87.disablelighting = True
    B87Flasher.visible = True
  else
    P_Bulb87.image = "BulbRed"
    P_Bulb87.disablelighting = False
    B87Flasher.visible = False
  end if

  if Controller.lamp(112) then
    P_BulbLB2.image = "BulbRed_lit"
    P_BulbLB2.disablelighting = True
    LB2Flasher.visible = True
  else
    P_BulbLB2.image = "BulbRed"
    P_BulbLB2.disablelighting = False
    LB2Flasher.visible = False
  end if
  if Controller.lamp(113) then
    P_BulbLB3.image = "BulbRed_lit"
    P_BulbLB3.disablelighting = True
    LB3Flasher.visible = True
  else
    P_BulbLB3.image = "BulbRed"
    P_BulbLB3.disablelighting = False
    LB3Flasher.visible = False
  end if
  if Controller.lamp(114) then
    P_BulbLB4.image = "BulbRed_lit"
    P_BulbLB4.disablelighting = True
    LB4Flasher.visible = True
  else
    P_BulbLB4.image = "BulbRed"
    P_BulbLB4.disablelighting = False
    LB4Flasher.visible = False
  end if
  if Controller.lamp(115) then
    P_BulbLB5.image = "BulbRed_lit"
    P_BulbLB5.disablelighting = True
    LB5Flasher.visible = True
  else
    P_BulbLB5.image = "BulbRed"
    P_BulbLB5.disablelighting = False
    LB5Flasher.visible = False
  end if
  if Controller.lamp(116) then
    P_BulbLB6.image = "BulbRed_lit"
    P_BulbLB6.disablelighting = True
    LB6Flasher.visible = True
  else
    P_BulbLB6.image = "BulbRed"
    P_BulbLB6.disablelighting = False
    LB6Flasher.visible = False
  end if
  if Controller.lamp(117) then
    P_BulbLB7.image = "BulbRed_lit"
    P_BulbLB7.disablelighting = True
    LB7Flasher.visible = True
  else
    P_BulbLB7.image = "BulbRed"
    P_BulbLB7.disablelighting = False
    LB7Flasher.visible = False
  end if
  if Controller.lamp(40) then
    Light40Flasher.visible = True
    Light40Flasher2.visible = True
  else
    Light40Flasher.visible = False
    Light40Flasher2.visible = False
  end if
  if Controller.lamp(41) then
    Light41Flasher.visible = True
    Light41Flasher2.visible = True
  else
    Light41Flasher.visible = False
    Light41Flasher2.visible = False
  end if
  if Controller.lamp(42) then
    Light42Flasher.visible = True
    Light42Flasher2.visible = True
  else
    Light42Flasher.visible = False
    Light42Flasher2.visible = False
  end if
  if Controller.lamp(43) then
    Light43Flasher.visible = True
    Light43Flasher2.visible = True
  else
    Light43Flasher.visible = False
    Light43Flasher2.visible = False
  end if
  if Controller.lamp(44) then
    Light44Flasher.visible = True
    Light44Flasher2.visible = True
  else
    Light44Flasher.visible = False
    Light44Flasher2.visible = False
  end if
  if Controller.lamp(45) then
    Light45Flasher.visible = True
    Light45Flasher2.visible = True
  else
    Light45Flasher.visible = False
    Light45Flasher2.visible = False
  end if

End sub

'Large Hole support
Sub LeftVUKHelper_Hit
  ActiveBall.x = LeftVUK.x
  ActiveBall.y = LeftVUK.y
End Sub

Sub CenterVUKHelper_Hit
  ActiveBall.x = CenterVUK.x
  ActiveBall.y = CenterVUK.y
End Sub

Sub UndergroundHoleHelper_Hit
  ActiveBall.x = UndergroundHole.x
  ActiveBall.y = UndergroundHole.y
End Sub

Sub RightVUKHelper_Hit
  if not RightVUK.Timerenabled then
    ActiveBall.x = RightVUK.x
    ActiveBall.y = RightVUK.y
  end if
End Sub

Sub RightVUK_Timer
  RightVUK.Timerenabled = False
End Sub

'Start GI
Sub GIStartTimer_Timer
  GIStartTimer.enabled = False
  For each obj in GIFlashers
    obj.visible = True
  Next
End Sub

'Flipper Primitives
sub FlipperTimer_Timer()
  LFPrim.objrotz=leftFlipper.CurrentAngle-90
  RFPrim.objrotz=rightFlipper.CurrentAngle-90
  URFPrim.objrotz=UrightFlipper.CurrentAngle-90
end sub


'-----------------------------
'------  Sound Handler  ------
'-----------------------------

Dim SoundActive,BallVel,Flying,SoundBall,OnRamp,WireRamp,Panning,BallXPos

Const TableWidth = 960                                    'used for sound panning, enter the value specified in table options "Table Width"
Const PanningFactor = 0.7                 'panning factor 0..1, 1 = left speaker off, if ball is rightmost and vice versa, 0 = no panning

Const SoundVel = 10                     'sound is played if ball velocity is above this threshold

Sub SoundTimer_Timer
  BallXPos = SoundBall.X
  if BallXPos < 0 then BallXPos = 0         'boundary check
  if BallXPos > TableWidth then BallXPos = Tablewidth
  Panning = ((2 * BallXPos / TableWidth) - 1) * PanningFactor   'scale the X position of the ball to the interval [-1,1]

  if SoundBall.Z > 30 then                    'Ball radius is 24, so the ball is definitely above the playfield - must be adjusted, if playfield level is not 0
    if Flying = 0 then StopAllSounds
    Flying = 1
    if OnRamp = 0 then exit sub             'no sound if "in the air"
  else
    if Flying = 1 then                  'Ball is returning to the playfield
    if OnRamp = 0 then                  'no sound if the ball returns down a ramp
      StopAllSounds
      PlaySound "BallCollision2",1,1,Panning,0
    end if
    Flying = 0
    OnRamp = 0
    WireRamp = 0
    end if
  end if

  BallVel = SQR((SoundBall.velx^2) + (SoundBall.vely^2))  'calculating the speed vector within the X/Y plane only once per timer call

  if SoundActive = 0 then                 'play new sounds only after the current sound has ended
    if BallVel > SoundVel then              'if velocity exceeds the threshold a sound will be played
    SoundActive = 1
    if OnRamp = 1 then
      if WireRamp = 1 then                'Ball is on a wire ramp
      PlaySound "WireRamp1",1,1,Panning,0.25
      else                        'Ball is on a normal ramp
      PlaySound "Ballroll3",1,1,Panning,0.25
      end if
    else                        'Ball is on the playfield, sound will be played as a loop
      PlaySound "roll1",-1,1,Panning,0.35
    end if
    end if
    Exit Sub
  else
    if BallVel < SoundVel then              'Sound is active and ball velocity drops below the threshold --> stop any sound active
    StopAllSounds                   'sound stopped so a new sound can be started
    end if
  end if
End Sub

Sub StopAllSounds()
  SoundActive = 0
  StopSound "WireRamp1"
  StopSound "Ballroll3"
  StopSound "roll1"
End Sub

Sub ShooterLaneLaunch_Hit
  if ActiveBall.vely < -6 then playsound "Launch",0,1,0.25,0.25
  Set SoundBall = Activeball  'Ball-assignment
  StopAllSounds
  SoundTimer.enabled = True
End Sub

Sub LeftVUKTopTrigger_Hit
  Set SoundBall = Activeball  'Ball-assignment
  StopAllSounds
  SoundTimer.enabled = True
  OnRamp = 1
End Sub

Sub RightVUKTopTrigger_Hit
  Set SoundBall = Activeball  'Ball-assignment
  StopAllSounds
  SoundTimer.enabled = True
  OnRamp = 1
  WireRamp = 1
End Sub

Sub LRampStart_Hit()
  StopAllSounds
  OnRamp = 1
End Sub

Sub LRampWireStart_Hit()
  StopAllSounds
  WireRamp = 1
End Sub

Sub LRampWireEnd_Hit()
  StopAllSounds
  WireRamp = 0
End Sub

Sub LRampMade_Hit()
  OnRamp = 0
End Sub

Sub RRampWireEnd_Hit()
  StopAllSounds
  WireRamp = 0
End Sub

Sub RRampMade_Hit()
  OnRamp = 0
End Sub

Sub RRampMade2_Hit()
  OnRamp = 0
End Sub


'---------------------------------------
'------  JP's Flasher Fading Sub  ------
'---------------------------------------

Dim Flashers
Flashers = Array(Flasher17a,Flasher18a,Flasher19a,Flasher20a,Flasher21a,Flasher22a,Flasher23a,Flasher14a,Flasher15a,Flasher16a,Flasher17b,Flasher18b,Flasher19b,Flasher20b,Flasher21b,Flasher22b,Flasher23b,Flasher14b,Flasher15b,Flasher16b,Flasher16c)

Dim FlashMaxOpacity
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown

FlashInit()
FlasherTimer.Interval = 5
FlasherTimer.Enabled = 1

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next
    FlashSpeedUp = 50    'fast speed when turning on the flasher
    FlashSpeedDown = 10  'slow speed when turning off the flasher, gives a smooth fading
    ' you could also change the default images for each flasher or leave it as in the editor
    ' for example
    ' Flasher1.Image = "fr"

  'added by mfuegemann to apply Dim settings
  FlashMaxOpacity = 255
  if DimFlashers < 0 then
    FlashMaxOpacity = FlashMaxOpacity + DimFlashers
    if FlashMaxOpacity < 0 then
      FlashMaxOpacity = 0
    end if
  end if

    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

dim x
Sub FlasherTimer_Timer()
  flashm 0, Flasher17b
  flash 0, Flasher17a
  flashm 1, Flasher18b
  flash 1, Flasher18a
  flashm 2, Flasher19b
  flash 2, Flasher19a
  flashm 3, Flasher20b
  flash 3, Flasher20a
  flashm 4, Flasher21b
  flash 4, Flasher21a
  flashm 5, Flasher22b
  flash 5, Flasher22a
  flashm 6, Flasher23b
  flash 6, Flasher23a
  flashm 7, Flasher14b
  flash 7, Flasher14a
  flashm 8, Flasher15b
  flash 8, Flasher15a
  flashm 9, Flasher16c
  flashm 9, Flasher16b
  flash 9, Flasher16a
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.Opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > FlashMaxOpacity Then        '255 original JP code
                FlashLevel(nr) = FlashMaxOpacity          '255 original JP code
                FlashState(nr) = -2 'completely on
            End if
            Object.Opacity = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
            Object.Opacity = FlashLevel(nr)
        Case 1         ' on
            Object.Opacity = FlashLevel(nr)
    End Select
End Sub

'LED taken from Victory Table (Gottlieb1987) by Sinbad
'https://vpinball.com/VPBdownloads/victory-gottlieb-1987-2-0-1/

Dim Digits(40)
Digits(0) = Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
Digits(1) = Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
Digits(2) = Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
Digits(3) = Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
Digits(4) = Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
Digits(5) = Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
Digits(6) = Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
Digits(7) = Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
Digits(8) = Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
Digits(9) = Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
Digits(10) = Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
Digits(11) = Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
Digits(12) = Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
Digits(13) = Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
Digits(14) = Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
Digits(15) = Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)
Digits(16) = Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
Digits(17) = Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
Digits(18) = Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
Digits(19) = Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
Digits(20) = Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
Digits(21) = Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
Digits(22) = Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
Digits(23) = Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
Digits(24) = Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
Digits(25) = Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
Digits(26) = Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
Digits(27) = Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
Digits(28) = Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
Digits(29) = Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
Digits(30) = Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
Digits(31) = Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)
Digits(32) = Array(c00, c05, c0c, c0d, c08, c01, c06, c0f, c02, c03, c04, c07, c0b, c0a, c09, c0e)
Digits(33) = Array(c10, c15, c1c, c1d, c18, c11, c16, c1f, c12, c13, c14, c17, c1b, c1a, c19, c1e)
Digits(34) = Array(c20, c25, c2c, c2d, c28, c21, c26, c2f, c22, c23, c24, c27, c2b, c2a, c29, c2e)
Digits(35) = Array(c30, c35, c3c, c3d, c38, c31, c36, c3f, c32, c33, c34, c37, c3b, c3a, c39, c3e)
Digits(36) = Array(c40, c45, c4c, c4d, c48, c41, c46, c4f, c42, c43, c44, c47, c4b, c4a, c49, c4e)
Digits(37) = Array(c50, c55, c5c, c5d, c58, c51, c56, c5f, c52, c53, c54, c57, c5b, c5a, c59, c5e)
Digits(38) = Array(c60, c65, c6c, c6d, c68, c61, c66, c6f, c62, c63, c64, c67, c6b, c6a, c69, c6e)
Digits(39) = Array(c70, c75, c7c, c7d, c78, c71, c76, c7f, c72, c73, c74, c77, c7b, c7a, c79, c7e)

Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 40) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
      Else
             end if
        Next
    End If
 End Sub

'VR and FSS Segment Display - Thanks to Rajo Joey
'*********************************************************************

Dim DigitsF(40)

DigitsF(0) = Array(fa00, fa05, fa0c, fa0d, fa08, fa01, fa06, fa0f, fa02, fa03, fa04, fa07, fa0b, fa0a, fa09, fa0e)
DigitsF(1) = Array(fa10, fa15, fa1c, fa1d, fa18, fa11, fa16, fa1f, fa12, fa13, fa14, fa17, fa1b, fa1a, fa19, fa1e)
DigitsF(2) = Array(fa20, fa25, fa2c, fa2d, fa28, fa21, fa26, fa2f, fa22, fa23, fa24, fa27, fa2b, fa2a, fa29, fa2e)
DigitsF(3) = Array(fa30, fa35, fa3c, fa3d, fa38, fa31, fa36, fa3f, fa32, fa33, fa34, fa37, fa3b, fa3a, fa39, fa3e)
DigitsF(4) = Array(fa40, fa45, fa4c, fa4d, fa48, fa41, fa46, fa4f, fa42, fa43, fa44, fa47, fa4b, fa4a, fa49, fa4e)
DigitsF(5) = Array(fa50, fa55, fa5c, fa5d, fa58, fa51, fa56, fa5f, fa52, fa53, fa54, fa57, fa5b, fa5a, fa59, fa5e)
DigitsF(6) = Array(fa60, fa65, fa6c, fa6d, fa68, fa61, fa66, fa6f, fa62, fa63, fa64, fa67, fa6b, fa6a, fa69, fa6e)
DigitsF(7) = Array(fa70, fa75, fa7c, fa7d, fa78, fa71, fa76, fa7f, fa72, fa73, fa74, fa77, fa7b, fa7a, fa79, fa7e)
DigitsF(8) = Array(fa80, fa85, fa8c, fa8d, fa88, fa81, fa86, fa8f, fa82, fa83, fa84, fa87, fa8b, fa8a, fa89, fa8e)
DigitsF(9) = Array(fa90, fa95, fa9c, fa9d, fa98, fa91, fa96, fa9f, fa92, fa93, fa94, fa97, fa9b, fa9a, fa99, fa9e)
DigitsF(10) = Array(faa0, faa5, faac, faad, faa8, faa1, faa6, faaf, faa2, faa3, faa4, faa7, faab, faaa, faa9, faae)
DigitsF(11) = Array(fab0, fab5, fabc, fabd, fab8, fab1, fab6, fabf, fab2, fab3, fab4, fab7, fabb, faba, fab9, fabe)
DigitsF(12) = Array(fac0, fac5, facc, facd, fac8, fac1, fac6, facf, fac2, fac3, fac4, fac7, facb, faca, fac9, face)
DigitsF(13) = Array(fad0, fad5, fadc, fadd, fad8, fad1, fad6, fadf, fad2, fad3, fad4, fad7, fadb, fada, fad9, fade)
DigitsF(14) = Array(fae0, fae5, faec, faed, fae8, fae1, fae6, faef, fae2, fae3, fae4, fae7, faeb, faea, fae9, faee)
DigitsF(15) = Array(faf0, faf5, fafc, fafd, faf8, faf1, faf6, faff, faf2, faf3, faf4, faf7, fafb, fafa, faf9, fafe)

DigitsF(16) = Array(fb00, fb05, fb0c, fb0d, fb08, fb01, fb06, fb0f, fb02, fb03, fb04, fb07, fb0b, fb0a, fb09, fb0e)
DigitsF(17) = Array(fb10, fb15, fb1c, fb1d, fb18, fb11, fb16, fb1f, fb12, fb13, fb14, fb17, fb1b, fb1a, fb19, fb1e)
DigitsF(18) = Array(fb20, fb25, fb2c, fb2d, fb28, fb21, fb26, fb2f, fb22, fb23, fb24, fb27, fb2b, fb2a, fb29, fb2e)
DigitsF(19) = Array(fb30, fb35, fb3c, fb3d, fb38, fb31, fb36, fb3f, fb32, fb33, fb34, fb37, fb3b, fb3a, fb39, fb3e)
DigitsF(20) = Array(fb40, fb45, fb4c, fb4d, fb48, fb41, fb46, fb4f, fb42, fb43, fb44, fb47, fb4b, fb4a, fb49, fb4e)
DigitsF(21) = Array(fb50, fb55, fb5c, fb5d, fb58, fb51, fb56, fb5f, fb52, fb53, fb54, fb57, fb5b, fb5a, fb59, fb5e)
DigitsF(22) = Array(fb60, fb65, fb6c, fb6d, fb68, fb61, fb66, fb6f, fb62, fb63, fb64, fb67, fb6b, fb6a, fb69, fb6e)
DigitsF(23) = Array(fb70, fb75, fb7c, fb7d, fb78, fb71, fb76, fb7f, fb72, fb73, fb74, fb77, fb7b, fb7a, fb79, fb7e)
DigitsF(24) = Array(fb80, fb85, fb8c, fb8d, fb88, fb81, fb86, fb8f, fb82, fb83, fb84, fb87, fb8b, fb8a, fb89, fb8e)
DigitsF(25) = Array(fb90, fb95, fb9c, fb9d, fb98, fb91, fb96, fb9f, fb92, fb93, fb94, fb97, fb9b, fb9a, fb99, fb9e)
DigitsF(26) = Array(fba0, fba5, fbac, fbad, fba8, fba1, fba6, fbaf, fba2, fba3, fba4, fba7, fbab, fbaa, fba9, fbae)
DigitsF(27) = Array(fbb0, fbb5, fbbc, fbbd, fbb8, fbb1, fbb6, fbbf, fbb2, fbb3, fbb4, fbb7, fbbb, fbba, fbb9, fbbe)
DigitsF(28) = Array(fbc0, fbc5, fbcc, fbcd, fbc8, fbc1, fbc6, fbcf, fbc2, fbc3, fbc4, fbc7, fbcb, fbca, fbc9, fbce)
DigitsF(29) = Array(fbd0, fbd5, fbdc, fbdd, fbd8, fbd1, fbd6, fbdf, fbd2, fbd3, fbd4, fbd7, fbdb, fbda, fbd9, fbde)
DigitsF(30) = Array(fbe0, fbe5, fbec, fbed, fbe8, fbe1, fbe6, fbef, fbe2, fbe3, fbe4, fbe7, fbeb, fbea, fbe9, fbee)
DigitsF(31) = Array(fbf0, fbf5, fbfc, fbfd, fbf8, fbf1, fbf6, fbff, fbf2, fbf3, fbf4, fbf7, fbfb, fbfa, fbf9, fbfe)

DigitsF(32) = Array(fc00, fc05, fc0c, fc0d, fc08, fc01, fc06, fc0f, fc02, fc03, fc04, fc07, fc0b, fc0a, fc09, fc0e)
DigitsF(33) = Array(fc10, fc15, fc1c, fc1d, fc18, fc11, fc16, fc1f, fc12, fc13, fc14, fc17, fc1b, fc1a, fc19, fc1e)
DigitsF(34) = Array(fc20, fc25, fc2c, fc2d, fc28, fc21, fc26, fc2f, fc22, fc23, fc24, fc27, fc2b, fc2a, fc29, fc2e)
DigitsF(35) = Array(fc30, fc35, fc3c, fc3d, fc38, fc31, fc36, fc3f, fc32, fc33, fc34, fc37, fc3b, fc3a, fc39, fc3e)
DigitsF(36) = Array(fc40, fc45, fc4c, fc4d, fc48, fc41, fc46, fc4f, fc42, fc43, fc44, fc47, fc4b, fc4a, fc49, fc4e)
DigitsF(37) = Array(fc50, fc55, fc5c, fc5d, fc58, fc51, fc56, fc5f, fc52, fc53, fc54, fc57, fc5b, fc5a, fc59, fc5e)
DigitsF(38) = Array(fc60, fc65, fc6c, fc6d, fc68, fc61, fc66, fc6f, fc62, fc63, fc64, fc67, fc6b, fc6a, fc69, fc6e)
DigitsF(39) = Array(fc70, fc75, fc7c, fc7d, fc78, fc71, fc76, fc7f, fc72, fc73, fc74, fc77, fc7b, fc7a, fc79, fc7e)

Sub UpdateLedsF_Timer
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For Each obj In DigitsF(num)
                If chg And 1 Then obj.Visible = stat And 1
                chg = chg \ 2:stat = stat \ 2
            Next
        Next
    End If
End Sub

Sub vrbgtimer_Timer()
    If bg50a.State = 1 then vrbg50.Visible = 1 Else vrbg50.Visible = 0
    If bg70a.State = 1 then vrbg70.Visible = 1 Else vrbg70.Visible = 0
    If bg75a.State = 1 then vrbg75.Visible = 1 Else vrbg75.Visible = 0
    If bg83a.State = 1 then vrbg83.Visible = 1 Else vrbg83.Visible = 0
End Sub

'**************************************************************
'LUT (Colour Look Up Table)

'0 = Fleep Natural Dark 1
'1 = Fleep Natural Dark 2
'2 = Fleep Warm Dark
'3 = Fleep Warm Bright
'4 = Fleep Warm Vivid Soft
'5 = Fleep Warm Vivid Hard
'6 = Skitso Natural and Balanced
'7 = Skitso Natural High Contrast
'8 = 3rdaxis Referenced THX Standard
'9 = CalleV Punchy Brightness and Contrast
'10 = HauntFreaks Desaturated
'11 = Tomate Washed Out
'12 = VPW Original 1 to 1
'13 = Bassgeige
'14 = Blacklight
'15 = B&W Comic Book

Dim LUTset, LutToggleSound
LutToggleSound = True
LoadLUT
SetLUT

'LUT selector timer

dim lutsetsounddir
sub LutSlctr_timer
  If lutsetsounddir = 1 And LutSet <> 15 Then
    Playsound "click", 0, 1, 0.1, 0.25
  End If
  If lutsetsounddir = -1 And LutSet <> 15 Then
    Playsound "click", 0, 1, 0.1, 0.25
  End If
  If LutSet = 15 Then
    Playsound "gun", 0, 1, 0.1, 0.25
  End If
  LutSlctr.enabled = False
end sub

'LUT Subs

Sub SetLUT
  OperationThunder.ColorGradeImage = "LUT" & LUTset
end sub

Sub LUTBox_Timer
  LUTBox.TimerEnabled = 0
  LUTBox.Visible = 0
  LUTBack.visible = 0
  VRLutdesc.visible = 0
End Sub

Sub ShowLUT
  LUTBox.visible = 1
  LUTBack.visible = 1
  VRLutdesc.visible = 1

  Select Case LUTSet
        Case 0: LUTBox.text = "Fleep Natural Dark 1": VRLUTdesc.imageA = "LUTcase0"
    Case 1: LUTBox.text = "Fleep Natural Dark 2": VRLUTdesc.imageA = "LUTcase1"
    Case 2: LUTBox.text = "Fleep Warm Dark": VRLUTdesc.imageA = "LUTcase2"
    Case 3: LUTBox.text = "Fleep Warm Bright": VRLUTdesc.imageA = "LUTcase3"
    Case 4: LUTBox.text = "Fleep Warm Vivid Soft": VRLUTdesc.imageA = "LUTcase4"
    Case 5: LUTBox.text = "Fleep Warm Vivid Hard": VRLUTdesc.imageA = "LUTcase5"
    Case 6: LUTBox.text = "Skitso Natural and Balanced": VRLUTdesc.imageA = "LUTcase6"
    Case 7: LUTBox.text = "Skitso Natural High Contrast": VRLUTdesc.imageA = "LUTcase7"
    Case 8: LUTBox.text = "3rdaxis Referenced THX Standard": VRLUTdesc.imageA = "LUTcase8"
    Case 9: LUTBox.text = "CalleV Punchy Brightness and Contrast": VRLUTdesc.imageA = "LUTcase9"
    Case 10: LUTBox.text = "HauntFreaks Desaturated" : VRLUTdesc.imageA = "LUTcase10"
      Case 11: LUTBox.text = "Tomate washed out": VRLUTdesc.imageA = "LUTcase11"
    Case 12: LUTBox.text = "VPW original 1on1": VRLUTdesc.imageA = "LUTcase12"
    Case 13: LUTBox.text = "bassgeige": VRLUTdesc.imageA = "LUTcase13"
    Case 14: LUTBox.text = "blacklight": VRLUTdesc.imageA = "LUTcase14"
    Case 15: LUTBox.text = "B&W Comic Book": VRLUTdesc.imageA = "LUTcase15"
  End Select

  LUTBox.TimerEnabled = 1

End Sub

Sub SaveLUT

  Dim FileObj
  Dim ScoreFile

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if

  if LUTset = "" then LUTset = 12 'failsafe

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "TableLUT.txt",True) 'Rename the tableLUT
  ScoreFile.WriteLine LUTset
  Set ScoreFile=Nothing
  Set FileObj=Nothing

End Sub

Sub LoadLUT

  Dim FileObj, ScoreFile, TextStr
  dim rLine

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    LUTset=12
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "TableLUT.txt") then  'Rename the tableLUT
    LUTset=12
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "TableLUT.txt")  'Rename the tableLUT
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    rLine = TextStr.ReadLine
    If rLine = "" then
      LUTset=12
      Exit Sub
    End if
    LUTset = int (rLine)
    Set ScoreFile = Nothing
      Set FileObj = Nothing
End Sub

'*****************************************************************************************************
' VR PLUNGER ANIMATION
'
' Code needed to animate the plunger. If you pull the plunger it will move in VR.
' IMPORTANT: there are two numeric values in the code that define the postion of the plunger and the
' range in which it can move. The first numeric value is the actual y position of the plunger primitive
' and the second is the actual y position + 135 to determine the range in which it can move.
'
' You need to to select the Primary_plunger primitive you copied from the
' template and copy the value of the Y position
' (e.g. 2114) into the code. The value that determines the range of the plunger is always the y
' position + 135 (e.g. 2249).
'
'*****************************************************************************************************

Sub TimerPlunger_Timer
  If VR_Primary_plunger.Y < 2435 then
      VR_Primary_plunger.Y = VR_Primary_plunger.Y + 3
  End If
End Sub

Sub TimerPlunger2_Timer
 'debug.print plunger.position
  VR_Primary_plunger.Y = 2335 + (5* Plunger.Position) -24
End Sub

' ***************** CODE BELOW IS FOR THE VR CLOCK.  This was originally taken from Rascal VP9 clock table *******************************
'*****************************************************************************************************************************************
Dim CurrentMinute ' for VR clock

Sub ClockTimer_Timer()

    'ClockHands Below *********************************************************
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
      Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())
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
  tmp = tableobj.y * 2 / OperationThunder.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / OperationThunder.width-1
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
  Vol = Csng(BallVel1(ball) ^2 / 2000)
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
  Pitch = BallVel1(ball) * 20
End Function

'Function BallVel1(ball) 'Calculates the ball speed
'  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
'End Function

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
        If BallVel1(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b))+5, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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

Function BallVel1(ball) 'Calculates the ball speed
    BallVel1 = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function
