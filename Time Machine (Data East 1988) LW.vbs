
 'Time Machine
 'Based on a JPSalas VP9's version
 'JPJ, Arngrim, Team PP VPX 1.0

    Option Explicit
    Randomize


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

 Const BallSize = 50.5
' Const BallMass = 2.1

'*********************************************************************************************************
'*** to show dmd in desktop Mod - Taken from ACDC Ninuzzu (THX) And Thanks To Rob Ross for Helping *******
Dim UseVPMDMD, DesktopMode
DesktopMode = Time_Machine.ShowDT
If NOT DesktopMode Then
  UseVPMDMD = False   'hides the internal VPMDMD when using the color ROM or when table is in Full Screen and color ROM is not in use
  metalsides.visible = False
  leftside.visible = True'put True or False if you want to see or not the wood side in fullscreen mod
  rightside.visible = True 'put True or False if you want to see or not the wood side in fullscreen mod
end if
If DesktopMode Then UseVPMDMD = True              'shows the internal VPMDMD when in desktop mode
'*********************************************************************************************************

 LoadVPM "00990300", "DE.VBS", 3.10

 Const cGameName="tmac_a24"
 Const UseSolenoids=2
 Const UseLamps=0
 Const UseSync=1
 Const sGI=11'Not Used


dim sw1
sw1=1

Dim x, bsTrough, bsVUK, vLock, bump1, bump2, bump3, PlungerIM, BallinPlunger, led, animA
Dim FastFlips
Dim GlobalSoundLevel
Dim TRA, TRAA, TRB, TRBB, TRC, TRCC
Dim vla, vlaa, vlb, vlbb, vlc, vlcc, vl
vla=0:vlb=0:vlc=0
vlaa=0:vlbb=0:vlcc=0

 ' Standard Sounds
 Const SSolenoidOn="Solenoid"
 Const SSolenoidOff=""
 Const SFlipperOn="FlipperUp"
 Const SFlipperOff="FlipperDown"
 Const SCoin="Coin"
 Const bladeArt = 1     '1=On (Art), 2=Sideblades Off.



'********    FLippers DE Anim and shadows     ******
  shfl.enabled = 1'             ****Activation****
'***************************************************

Dim FlashLevel1, FlashLevel2, FlashLevel3, FlashLevel4
FlasherLight1.IntensityScale = 0
Flasherlight2.IntensityScale = 0
Flasherlight3.IntensityScale = 0
Flasherlight4.IntensityScale = 0

set Lights(108) = L108State
set Lights(109) = L109State


 '******************************************************
 ' Table ini
 '******************************************************

 Sub Time_Machine_Init

' Thalamus : Was missing 'vpminit me'
vpminit me

TRA = 0
TRAA = 0
TRB = 0
TRBB = 0
TRC = 0
TRCC = 0
vl=0
Primitive6.image = "bumpcircle0"
Primitive9.image = "bumpcircle0"
Primitive7.image = "bumpcircle0"
GlobalSoundLevel = 1.2
Primitive47.image = "shadows"
Primitive47.visible = 1
'Primitive48.visible = 1
BallinPlunger = 0
AnimA = 0

RainbowTimer.enabled = 1
'*********** Led Mod ********************************************
Led = 1' 1 for Rainbow's mod 0 for normal and bonus rainbow's mod
'****************************************************************
  dim z


 '******************************************************
 ' Automatic screw rotation ;) by JPJ
 '******************************************************
    For Each x in AutomaticRot
  z = Int(Rnd*360)+1
    x.RotZ = z
  Next
    For Each x in AutomaticRot2
  z = Int(Rnd*360)+1
    x.RotY = z
  Next
'******************************************************

'**** Fastflips
    Set FastFlips = new cFastFlips
    with FastFlips
       .CallBackL = "SolLflipper"  'Point these to flipper subs
       .CallBackR = "SolRflipper"  '...
'       .CallBackUL = "SolULflipper"'...(upper flippers, if needed)
'       .CallBackUR = "SolURflipper"'...
       .TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
 '      .DebugOn = True        'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
    end with



' Autofire Plunger
    Const IMPowerSetting = 45
    Const IMTime = 0.4
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .InitExitSnd "plunger", "solenoid"
        .CreateEvents "plungerIM"
    End With

    With Controller
       .GameName=cGameName
       If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
       .SplashInfoLine="Time Machine, Data East 1988" & vbNewLine & "VP91x table by JPSalas v.1.0"
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
      End With
      Controller.Hidden = 0
      On Error Resume Next

      Controller.Run




    vpmNudge.TiltSwitch=1
    vpmNudge.Sensitivity=5
    vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)



      set bsTrough=new cvpmBallStack
    bsTrough.InitSw 0, 13, 12, 11, 10, 0, 0, 0
    bsTrough.InitKick BallRelease, 130, 2
    bsTrough.InitEntrySnd "Solenoid", "Solenoid"
    bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls=3


    ' Thalamus - more randomness to kickers pls
    set bsVUK=new cvpmBallStack
    bsVUK.KickForceVar = 3
    bsVUK.KickAngleVar = 3
    bsVUK.InitSw 0, 45, 0, 0, 0, 0, 0, 0
    bsVUK.InitKick Switch45, 130, 2
    bsVUK.InitEntrySnd "Solenoid", "Solenoidoff"
    bsVUK.InitExitSnd SoundFX("exitsubway",DOFContactors), SoundFX("popper",DOFContactors)



    ' Visible Lock - implements post ball lock
    Set vLock=New cvpmVLock
    vLock.InitVLock Array(Switch36, Switch37, Switch38), Array(Switch36Kick, Switch37Kick, Switch38Kick), Array(36, 37, 38)
    vLock.InitSnd "Solenoid", "Solenoidoff"
    vLock.CreateEvents "vLock"

    kickback.Pullback

    PinMAMETimer.Interval=PinMAMEInterval
    PinMAMETimer.Enabled=1


 End Sub

 Sub Time_Machine_Paused : Controller.Pause=True : End Sub
 Sub Time_Machine_unPaused : Controller.Pause=False : End Sub

 '**********
 ' Keys
 '**********

 Sub Time_Machine_KeyDown(ByVal keycode)
    If keycode = LeftTiltKey Then LeftNudge 80, 2, 20:PlaySound SoundFX("nudge_left",0)
    If keycode = RightTiltKey Then RightNudge 280, 2, 20:PlaySound SoundFX("nudge_right",0)
    If keycode = CenterTiltKey Then CenterNudge 0, 2, 25:PlaySound SoundFX("nudge_forward",0)
    If KeyDownHandler(keycode)Then Exit Sub
    If keycode = PlungerKey Then Plunger.Pullback
    If keycode=KeyRules Then Rules
    If KeyCode = LeftFlipperKey then FastFlips.FlipL True ':  FastFlips.FlipUL True
    If KeyCode = RightFlipperKey then FastFlips.FlipR True ':FastFlips.FlipUR True
 End Sub

 Sub Time_Machine_KeyUp(byval Keycode)
    If KeyUpHandler(keycode)Then Exit Sub
    If KeyCode = LeftFlipperKey then FastFlips.FlipL False ':  FastFlips.FlipUL True
    If KeyCode = RightFlipperKey then FastFlips.FlipR False' :FastFlips.FlipUR True
    If keycode = PlungerKey Then
        Plunger.Fire
        If(BallinPlunger = 1) then 'the ball is in the plunger lane
            PlaySoundAtVol "Plunger2", Plunger, 1
        else
            PlaySoundAtVol "Plunger", Plunger, 1
        end if
    End If
 End Sub

Sub ballinPL_UnHit:
  BallinPlunger = 1
End Sub
Sub Gate1_Hit:BallinPlunger = 0:End Sub
Sub swPlunger_Hit:Controller.Switch(14) = 1:End Sub 'in this sub you may add a switch, for example Controller.Switch(14) = 1

Sub swPlunger_UnHit:Controller.Switch(14) = 0:End Sub 'in this sub you may add a switch, for example Controller.Switch(14) = 0


 '*************************************
 '          Nudge System
 ' based on Noah's nudgetest table
 '*************************************

 Dim LeftNudgeEffect, RightNudgeEffect, NudgeEffect

 Sub LeftNudge(angle, strength, delay)
     vpmNudge.DoNudge angle, (strength * (delay-LeftNudgeEffect) / delay) + RightNudgeEffect / delay
     LeftNudgeEffect = delay
     RightNudgeEffect = 0
     RightNudgeTimer.Enabled = 0
     LeftNudgeTimer.Interval = delay
     LeftNudgeTimer.Enabled = 1
 End Sub

 Sub RightNudge(angle, strength, delay)
     vpmNudge.DoNudge angle, (strength * (delay-RightNudgeEffect) / delay) + LeftNudgeEffect / delay
     RightNudgeEffect = delay
     LeftNudgeEffect = 0
     LeftNudgeTimer.Enabled = 0
     RightNudgeTimer.Interval = delay
     RightNudgeTimer.Enabled = 1
 End Sub

 Sub CenterNudge(angle, strength, delay)
     vpmNudge.DoNudge angle, strength * (delay-NudgeEffect) / delay
     NudgeEffect = delay
     NudgeTimer.Interval = delay
     NudgeTimer.Enabled = 1
 End Sub

 Sub LeftNudgeTimer_Timer()
     LeftNudgeEffect = LeftNudgeEffect-1
     If LeftNudgeEffect = 0 then LeftNudgeTimer.Enabled = False
 End Sub

 Sub RightNudgeTimer_Timer()
     RightNudgeEffect = RightNudgeEffect-1
     If RightNudgeEffect = 0 then RightNudgeTimer.Enabled = False
 End Sub

 Sub NudgeTimer_Timer()
     NudgeEffect = NudgeEffect-1
     If NudgeEffect = 0 then NudgeTimer.Enabled = False
 End Sub


 '*******
 'Outhole
 '*******

 Sub OutHole_Hit()
    PlaySoundAtVol "drain", Drain, 1
    ClearBallID
    vpmTimer.PulseSwitch 10, 0, 0
    bsTrough.AddBall Me
 End Sub

 '*******
 'VUK
 '*******
 Dim numberofballsinhole:numberofballsinhole = 0

 Sub Switch44_Hit()
  numberofballsinhole = numberofballsinhole + 1
    vpmTimer.PulseSw 44
    ClearBallID
    Switch44.Destroyball
  TRC=0:TRCC=0:stopsound "Plasticramp3"
  PlaySoundAtVol "underballrolling", Switch44, 1
  SW45Wait.enabled = 1
'    bsVUK.AddBall 0

 End Sub


sub SW45Wait_timer()
  numberofballsinhole = numberofballsinhole - 1
    bsVUK.AddBall 0
    SW45Wait.enabled = 0
end sub
 '**************
 ' Flipper Subs
 '**************

 Sub LeftFlipper_Collide(parm)
     PlaySoundAtVol "rubber_flipper", ActiveBall, 1
 End Sub

 Sub RightFlipper_Collide(parm)
     PlaySoundAtVol "rubber_flipper", ActiveBall, 1
 End Sub

 '**********
 ' Solenoids
 '**********

 SolCallback(1)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
 SolCallback(2)="vpmSolSound SoundFX(""Chimer1"",DOFChimes),"
 SolCallback(3)="vpmSolSound SoundFX(""Chimer2"",DOFChimes),"
 SolCallback(4)="vpmSolSound SoundFX(""Chimer3"",DOFChimes),"
 SolCallback(5)="SetLamp 105,"
 SolCallBack(6)="SetLamp 106,"
 SolCallBack(7)="SetLamp 107,"
 SolCallBack(8)="SetLamp 108,"
 SolCallBack(9)="SetLamp 109,"
 SolCallback(11) = "GIUpdate"
 SolCallBack(12)="SetLamp 112,"
 SolCallback(13)="SetLamp 113,"
 SolCallback(14)="SetLamp 114,"
 SolCallback(15)="SetLamp 115,"
 SolCallback(16) = "SolKickBack" 'SolCallback(16)="vpmSolAutoPlunger KickBack,30,"
 SolCallback(23)  ="FastFlips.TiltSol" 'Fastflips
 SolCallback(25) = "SolTroughIn"     ' Outhole
 SolCallback(26) = "bsTrough.SolOut"     ' Ball Release
 SolCallback(27) = "SolVUKOut"     ' VUK
 SolCallback(28) = "SolLock"     ' Lock Ball
 SolCallback(sLRFlipper) = "SolRFlipper" 'Fastflips
 SolCallback(sLLFlipper) = "SolLFlipper" 'Fastflips

Sub SolLFlipper(Enabled)
     If Enabled Then
       PlaySoundAtVol SoundFX("LFlipperUp",DOFFlippers), LeftFlipper, 1
      LeftFlipper.RotateToEnd
      Controller.Switch(15)=True
  else
      PlaySoundAtVol SoundFX("LFlipperRel",DOFFlippers), LeftFlipper, 1
      LeftFlipper.RotateToStart
      Controller.Switch(15)=False
     End If
    End Sub

  Sub SolRFlipper(Enabled)
     If Enabled Then
       PlaySoundAtVol SoundFX("RFlipperUp",DOFFlippers), RightFlipper, 1
      RightFlipper.RotateToEnd
      Controller.Switch(16)=True
     Else
       PlaySoundAtVol SoundFX("RFlipperRel",DOFFlippers), RightFlipper, 1
      RightFlipper.RotateToStart
      Controller.Switch(16)=False
     End If
  End Sub

Sub SolKickBack(enabled)
    If enabled Then
       Kickback.Fire
       PlaySoundAtVol SoundFX ("KickBack",DOFContactors), KickBack, 1
    Else
       Kickback.PullBack
    End If
End Sub

 Sub SolTrough(Enabled)
   If Enabled Then
  If (ABS(bsTrough.Balls)+ABS(VLock.Balls)=3) And Not Timer2.Enabled Then bsTrough.ExitSol_On:Solanim.enabled = 1
   End If
  End Sub

  Sub SolTroughIn(Enabled)
   If Enabled Then
    bsTrough.EntrySol_On
    Timer2.Enabled=0
    Timer2.Enabled=1
   End If
  End Sub

  Sub SolVUKOut(Enabled)
   If Enabled Then
    Timer2.Enabled=0
    Timer2.Enabled=1
    bsVUK.ExitSol_On
   End If
  End Sub

   Sub SolLock(Enabled)
   If Enabled Then
    Timer2.Enabled=0
    Timer2.Enabled=1
    vLock.SolExit True
  Solanim.enabled=1
  vl=vl-1
   Else
    Timer2.Enabled=0
    Timer2.Enabled=1
    vLock.SolExit False
  vl=vl+1
   End If
  End Sub

  Sub Timer2_Timer
  debug.print vl
   Timer2.Enabled=0
  End Sub

Dim xx
Sub GIUpdate(enabled)
  If enabled Then
'Primitive47.visible = 0
Primitive47.image = "shadowsn"
Leftside.image = "leftsidetext3"
Rightside.image = "Rightsidetext3"
    For each xx in GI:xx.State = 0:Next
  else
'Primitive47.visible = 1
Primitive47.image = "shadows"
Leftside.image = "leftsidetext4"
Rightside.image = "Rightsidetext4"
    For each xx in GI:xx.State = 1:Next
  End If
End Sub

' Choose Side Blades
  if bladeArt = 1 then
    PinCab_Blades.Image = "Sidewalls TM"
    PinCab_Blades.visible = 1
  elseif bladeArt = 2 then
    PinCab_Blades.visible = 0
  End if

 '*********
 ' Switches
 '*********

 'Slings
 Sub LeftSlingshot_SlingShot() : PlaySoundAtVol SoundFX("LSlingshot",DOFContactors), ActiveBall, 1:vpmTimer.PulseSw 21 : End Sub
 Sub RightSlingshot_SlingShot() : PlaySoundAtVol SoundFX("RSlingshot",DOFContactors), ActiveBall, 1:vpmTimer.PulseSw 22 : End Sub

 '********************* Bumpers ****************************

 Sub Switch46_Hit() : PlaySoundAtVol SoundFX("bumperyellow",DOFContactors) , ActiveBall, 1: vpmTimer.PulseSw 46 : : bump1=1 : Me.TimerEnabled=1 : End Sub
 Sub Switch46_Timer(): Me.Timerenabled = 0 : End Sub

 Sub Switch48_Hit() : PlaySoundAtVol SoundFX("bumperblue",DOFContactors) , ActiveBall, 1: vpmTimer.PulseSw 48 : bump2=1 : Me.TimerEnabled=1 : End Sub
 Sub Switch48_Timer(): Me.Timerenabled = 0 : End Sub

 Sub Switch47_Hit() : PlaySoundAtVol SoundFX("bumperred",DOFContactors) , ActiveBall, 1: vpmTimer.PulseSw 47 : Me.TimerEnabled=1 : End Sub
 Sub Switch47_Timer():  Me.Timerenabled = 0 : End Sub

 '***********************************************************

 'Shooter
 Sub Switch14_Hit() : Controller.Switch(14)=1 : End Sub
 Sub Switch14_UnHit() : Controller.Switch(14)=0 : End Sub

 'Out Lanes
 Sub Switch17_Hit() : PlaySoundAtVol "sensor" , ActiveBall, 1: Controller.Switch(17)=1 : End Sub
 Sub Switch17_UnHit() : Controller.Switch(17)=0 : End Sub

 Sub Switch19_Hit() : PlaySoundAtVol "sensor" , ActiveBall, 1: Controller.Switch(19)=1 : End Sub
 Sub Switch19_UnHit() : Controller.Switch(19)=0 : End Sub

 'In Lanes
 Sub Switch18_Hit() : PlaySoundAtVol "sensor" , ActiveBall, 1: Controller.Switch(18)=1 : End Sub
 Sub Switch18_UnHit() : Controller.Switch(18)=0 : End Sub

 Sub Switch20_Hit() : PlaySoundAtVol "sensor" , ActiveBall, 1: Controller.Switch(20)=1 : End Sub
 Sub Switch20_UnHit() : Controller.Switch(20)=0 : End Sub

 'TopLane
 Sub Switch25_Hit() : PlaySoundAtVol "sensor" , ActiveBall, 1: Controller.Switch(25)=1 : End Sub
 Sub Switch25_UnHit() : Controller.Switch(25)=0 : End Sub
 Sub Switch26_Hit() : PlaySoundAtVol "sensor" , ActiveBall, 1: Controller.Switch(26)=1 : End Sub
 Sub Switch26_UnHit() : Controller.Switch(26)=0 : End Sub
 Sub Switch27_Hit() : PlaySoundAtVol "sensor" , ActiveBall, 1: Controller.Switch(27)=1 : End Sub
 Sub Switch27_UnHit() : Controller.Switch(27)=0 : End Sub

 'Ramps
 Sub Switch28_Hit() : Controller.Switch(28)=1 : End Sub
 Sub Switch28_UnHit() : Controller.Switch(28)=0 : End Sub
 Sub Switch29_Hit() : Controller.Switch(29)=1 : End Sub
 Sub Switch29_UnHit() : Controller.Switch(29)=0 : End Sub
 Sub Switch30_Hit() : Controller.Switch(30)=1: End Sub
 Sub Switch30_UnHit() : Controller.Switch(30)=0 : End Sub

 'RollOvers
 Sub Switch31_Hit() : Controller.Switch(31)=1 : Switch31L.State=1 : Me.TimerEnabled=1 : End Sub
 Sub Switch31_UnHit() : Controller.Switch(31)=0 : End Sub
 Sub Switch31_Timer : Me.TimerEnabled=0 : Switch31L.State=0 : End Sub
 Sub Switch32_Hit() : Controller.Switch(32)=1 : Switch32L.State=1 : Me.TimerEnabled=1 : End Sub
 Sub Switch32_UnHit() : Controller.Switch(32)=0 : End Sub
 Sub Switch32_Timer : Me.TimerEnabled=0 : Switch32L.State=0 : End Sub

 'Left Targets
Sub SWsqr5_hit():sqr5.transz = 10:sqr6.transz = 10:Me.TimerEnabled = 1:vpmTimer.PulseSw 35:PlaySoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub SWsqr5_Timer():sqr5.transz = 5:sqr6.transz = 5:SWsqr5a:Me.TimerEnabled = 0:End Sub
Sub SWsqr5a:sqr5.transZ = 0:sqr6.transz = 0:End Sub

Sub SWror5_hit():ror5.transz = 10:ror6.transz = 10:Me.TimerEnabled = 1:vpmTimer.PulseSw 34:PlaySoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub SWror5_Timer():ror5.transz = 5:ror6.transz = 5:SWror5a:Me.TimerEnabled = 0:End Sub
Sub SWror5a:ror5.transZ = 0:ror6.transz = 0:End Sub

Sub swtrr5_hit():trr5.transz = 10:trr6.transz = 10:Me.TimerEnabled = 1:vpmTimer.PulseSw 33:PlaySoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub swtrr5_Timer():trr5.transz = 5:trr6.transz = 5:SWtrr5a:Me.TimerEnabled = 0:End Sub
Sub swtrr5a:trr5.transZ = 0:trr6.transz = 0:End Sub

 'Center Targets
Sub SWsqr3_hit():sqr3.transz = 10:sqr4.transz = 10:Me.TimerEnabled = 1:vpmTimer.PulseSw 43:PlaySoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub SWsqr3_Timer():sqr3.transz = 5:sqr4.transz = 5:SWsqr3a:Me.TimerEnabled = 0:End Sub
Sub SWsqr3a:sqr3.transz = 0:sqr4.transz = 0:End Sub

Sub SWror3_hit():ror3.transz = 10:ror4.transz = 10:Me.TimerEnabled = 1:vpmTimer.PulseSw 42:PlaySoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub SWror3_Timer():ror3.transz = 5:ror4.transz = 5:SWror3a:Me.TimerEnabled = 0:End Sub
Sub SWror3a:ror3.transZ = 0:ror4.transz = 0:End Sub

Sub swtrr3_hit():trr3.transz = 10:trr4.transz = 10:Me.TimerEnabled = 1:vpmTimer.PulseSw 41:PlaySoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub swtrr3_Timer():trr3.transz = 5:trr4.transz = 5:SWtrr3a:Me.TimerEnabled = 0:End Sub
Sub swtrr3a:trr3.transZ = 0:trr4.transz = 0:End Sub

 'Right Targets
Sub SWsqr1_hit():sqr1.transz = 10:sqr2.transz = 10:Me.TimerEnabled = 1:vpmTimer.PulseSw 51:PlaySoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub SWsqr1_Timer():sqr1.transz = 5:sqr2.transz = 5:SWsqr1a:Me.TimerEnabled = 0:End Sub
Sub SWsqr1a:sqr1.transZ = 0:sqr2.transz = 0:End Sub

Sub SWror1_hit():ror1.transz = 10:ror2.transz = 10:Me.TimerEnabled = 1:vpmTimer.PulseSw 50:PlaySoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub SWror1_Timer():ror1.transz = 5:ror2.transz = 5:SWror1a:Me.TimerEnabled = 0:End Sub
Sub SWror1a:ror1.transZ = 0:ror2.transz = 0:End Sub

Sub swtrr1_hit():trr1.transz = 10:trr2.transz = 10:Me.TimerEnabled = 1:vpmTimer.PulseSw 49:PlaySoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub swtrr1_Timer():trr1.transz = 5:trr2.transz = 5:SWtrr1a:Me.TimerEnabled = 0:End Sub
Sub swtrr1a:trr1.transZ = 0:trr2.transz = 0:End Sub


Sub SW34_hit():sw34p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 34:End Sub
Sub SW34_Timer():sw34p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW35_hit():sw35p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 35:End Sub
Sub SW35_Timer():sw35p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW20_hit():sw20p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 20:End Sub
Sub SW20_Timer():sw20p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW21_hit():sw21p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 21:End Sub
Sub SW21_Timer():sw21p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW22_hit():sw22p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 22:End Sub
Sub SW22_Timer():sw22p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW38_hit():sw38p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 38:End Sub
Sub SW38_Timer():sw38p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW39_hit():sw39p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 39:End Sub
Sub SW39_Timer():sw39p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW40_hit():sw40p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 40:End Sub
Sub SW40_Timer():sw40p.transx = 0:Me.TimerEnabled = 0:End Sub



 '*************
 'Lock Switches
 '*************

 Sub SolLockRelease(enabled)
  solanim.enabled = 1
    vLock.SolExit enabled
 End Sub



'****************************************
'  JP's Fading Lamps v5.1 VP9 Fading only
'      Based on PD's Fading Lights
' SetLamp 0 is Off
' SetLamp 1 is On
' LampState(x) current state
'****************************************

Dim LampState(200)

AllLampsOff()
LampTimer.Interval = 35
LampTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
    If l52h.state = 1 then BumpCircle.enabled = 1
    If l53h.state = 1 then BumpCircle1.enabled = 1
    If l54h.state = 1 then BumpCircle2.enabled = 1
        Next
    End If
    If l52h.state = 0 then BumpCircle.enabled = 0:Primitive6.image = "bumpCircle0":circle = 0:end if
    If l53h.state = 0 then BumpCircle1.enabled = 0:Primitive7.image = "bumpCircle0":circle1=0:end if
    If l54h.state = 0 then BumpCircle2.enabled = 0:Primitive9.image = "bumpCircle0":circle2=0:end if
    UpdateLamps
End Sub

Sub UpdateLamps
    NFadeLm 105, f1a
    FadeLm 105, f1b, f1b1 '
    FadeLm 105, f1c, f1c1 '
    MFadeL 105, f1d, f1d1,48 '
    NFadeLm 106, f2d
    FadeLm 106, f2a, f2a1 '
    FadeLm 106, f2b, f2b1 '
    MFadeL 106, f2c, f2c1,48 '
    FadeLm 107, f3a, f3a1 '
    FadeLm 107, f3b, f3b1 '
    FadeLm 107, f3c, f3c1 '
    FadeL 107, f3d, f3d1 '
  NFadeL 108, L108State
    if L108State.State = 1 Then
    FlashLevel3 = 1 : FlasherFlash3_Timer
    FlashLevel4 = 1 : FlasherFlash4_Timer
      flasha5.Amount = 100
      flasha5.IntensityScale = 5
      flasha6.Amount = 60
      flasha6.IntensityScale = 3
      flasha7.Amount = 100
      flasha7.IntensityScale = 5
      flasha8.Amount = 60
      flasha8.IntensityScale = 3
  end If
    if L108State.State = 0 Then
      flasha5.Amount = 0
      flasha5.IntensityScale = 0
      flasha6.Amount = 0
      flasha6.IntensityScale = 0
      flasha7.Amount = 0
      flasha7.IntensityScale = 0
      flasha8.Amount = 0
      flasha8.IntensityScale = 0
  end If
  NFadeL 109, L109State
    if L109State.State = 1 Then
      FlashLevel1 = 1 : FlasherFlash1_Timer
      FlashLevel2 = 1 : FlasherFlash2_Timer
      flasha1.Amount = 100
      flasha1.IntensityScale = 5
      flasha2.Amount = 60
      flasha2.IntensityScale = 3
      flasha3.Amount = 100
      flasha3.IntensityScale = 5
      flasha4.Amount = 60
      flasha4.IntensityScale = 3
  end If
    if L109State.State = 0 Then
      flasha1.Amount = 0
      flasha1.IntensityScale = 0
      flasha2.Amount = 0
      flasha2.IntensityScale = 0
      flasha3.Amount = 0
      flasha3.IntensityScale = 0
      flasha4.Amount = 0
      flasha4.IntensityScale = 0
  end If
  MNFadeLm 112, f12a,39
    MNFadeL 112, f12b, 63
    MNFadeLm 113, f13a,40
    MNFadeL 113, f13b, 64
    MNFadeLm 114, f14a,32
    MNFadeL 114, f14b, 56
    MNFadeLm 115, f15a,31
    MNFadeL 115, f15b, 55


    if Not Desktopmode then
    FadeOldL 1, l1, l1a, l1b
    FadeOldL 2, l2, l2a, l2b
    FadeOldL 3, l3, l3a, l3b
    FadeOldL 4, l4, l4a, l4b
    FadeOldL 5, l5, l5a, l5b
    FadeOldL 6, l6, l6a, l6b
    FadeOldL 7, l7, l7a, l7b
    FadeOldL 8, l8, l8a, l8b
  else
    FadeOldLm 1, l1, l1a, l1b
    NFadeL 1, l1z
    FadeOldLm 2, l2, l2a, l2b
    NFadeL 2, l2z
    FadeOldLm 3, l3, l3a, l3b
    NFadeL 3, l3z
    FadeOldLm 4, l4, l4a, l4b
    NFadeL 4, l4z
    FadeOldLm 5, l5, l5a, l5b
    NFadeL 5, l5z
    FadeOldLm 6, l6, l6a, l6b
    NFadeL 6, l6z
    FadeOldLm 7, l7, l7a, l7b
    NFadeL 7, l7z
    FadeOldLm 8, l8, l8a, l8b
    NFadeL 8, l8z
  end if

    FadeL 9, l9, l9a
    FadeL 10, l10, l10a
    FadeL 11, l11, l11a
    FadeL 12, l12, l12a
    FadeL 13, l13, l13a
    FadeL 14, l14, l14a
    FadeL 15, l15, l15a
    FadeL 16, l16, l16a
    FadeL 17, l17, l17a
    FadeL 18, l18, l18a
    FadeL 19, l19, l19a
    FadeL 20, l20, l20a
    FadeL 21, l21, l21a
    FadeL 22, l22, l22a
    FadeL 23, l23, l23a
    FadeL 24, l24, l24a
  NFadeL 25, l25c
  NFadeL 26, l26c
  NFadeL 27, l27c
  NFadeL 28, l28c
  NFadeL 29, l29c
  NFadeL 30, l30c
if l25c.state=1 then f25c.visible=1 else if l25c.state=0 then f25c.visible=0
if l26c.state=1 then f26c.visible=1 else if l26c.state=0 then f26c.visible=0
if l27c.state=1 then f27c.visible=1 else if l27c.state=0 then f27c.visible=0
if l28c.state=1 then f28c.visible=1 else if l28c.state=0 then f28c.visible=0
if l29c.state=1 then f29c.visible=1 else if l29c.state=0 then f29c.visible=0
if l30c.state=1 then f30c.visible=1 else if l30c.state=0 then f30c.visible=0
    FadeL 31, l31, l31a
    FadeL 32, l32, l32a
    FadeL 33, l33, l33a
    FadeL 34, l34, l34a
    FadeL 35, l35, l35a
    FadeL 36, l36, l36a
    FadeL 37, l37, l37a
    FadeL 38, l38, l38a
    FadeL 39, l39, l39a
    FadeL 40, l40, l40a
    FadeL 41, l41, l41a
    FadeL 42, l42, l42a
    FadeL 43, l43, l43a
    FadeL 44, l44, l44a
    FadeL 45, l45, l45a
    FadeL 46, l46, l46a
    FadeL 47, l47, l47a
    FadeL 48, l48, l48a
    FadeL 49, l49, l49a
    FadeL 50, l50, l50a
    FadeL 51, l51, l51a
    NFadeLm 52, l52h
    NFadeL 52, l52
    NFadeLm 53, l53h
    NFadeL 53, l53
    NFadeLm 54, l54h
    NFadeL 54, l54

    FadeL 55, l55, l55a
    FadeL 56, l56, l56a
    FadeL 57, l57, l57a
    FadeL 58, l58, l58a
    FadeL 59, l59, l59a
    FadeL 60, l60, l60a
    FadeL 61, l61, l61a
 '   62 engine blackglass
    FadeL 63, l63, l63a
    FadeL 64, l64, l64a
End Sub

Sub AllLampsOff()
    Dim x
    For x = 0 to 200
        LampState(x) = 4
    Next
UpdateLamps:UpdateLamps:Updatelamps
End Sub

Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr)=0 Then Exit Sub
    If value = 1 AND LampState(nr)=1 Then Exit Sub
    LampState(nr) = abs(value) + 4
End Sub

'force the state of a lamp
Sub SetFLamp(nr, value)
    LampState(nr) = abs(value) + 4
End Sub

'Walls

Sub FadeW(nr, a, b, c)
    Select Case LampState(nr)
        Case 2:c.IsDropped = 1:LampState(nr) = 0                 'Off
        Case 3:b.IsDropped = 1:c.IsDropped = 0:LampState(nr) = 2 'fading...
        Case 4:a.IsDropped = 1:b.IsDropped = 0:LampState(nr) = 3 'fading...
        Case 5:c.IsDropped = 1:b.IsDropped = 0:LampState(nr) = 6 'ON
        Case 6:b.IsDropped = 1:a.IsDropped = 0:LampState(nr) = 1 'ON
    End Select
End Sub


'Lights

Sub FadeL(nr, a, b)
    Select Case LampState(nr)
        Case 4:a.state = 0:b.state = 0:LampState(nr) = 3
        Case 5:a.state = 1:b.state = 1:LampState(nr) = 6
    End Select
End Sub

Sub FadeLm(nr, a, b)
    Select Case LampState(nr)
        Case 4:a.state = 0:b.state = 0
        Case 5:a.state = 1:b.state = 1
    End Select
End Sub

Sub NFadeL(nr, a)
    Select Case LampState(nr)
        Case 4:a.state = 0:LampState(nr) = 0
        Case 5:a.State = 1:LampState(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, a)
    Select Case LampState(nr)
        Case 4:a.state = 0
        Case 5:a.State = 1
    End Select
End Sub

Sub FadeOldL(nr, a, b, c)
    Select Case LampState(nr)
        Case 2:c.state = 0:LampState(nr) = 0
        Case 3:b.state = 0:c.state = 1:LampState(nr) = 2
        Case 4:a.state = 0:b.state = 1:LampState(nr) = 3
        Case 5:a.state = 0:c.state = 0:b.state = 1:LampState(nr) = 6
        Case 6:b.state = 0:a.state = 1:LampState(nr) = 1
    End Select
End Sub

Sub FadeOldLm(nr, a, b, c)
    Select Case LampState(nr)
        Case 2:c.state = 0
        Case 3:b.state = 0:c.state = 1
        Case 4:a.state = 0:b.state = 1
        Case 5:a.state = 0:c.state = 0:b.state = 1
        Case 6:b.state = 0:a.state = 1
    End Select
End Sub

' Light acting as a flash. C is the light number to be restored

Sub MFadeL(nr, a, b, c)
    Select Case LampState(nr)
        Case 2:b.state = 0:LampState(nr) = 0:SetLamp c, LampState(c)
        Case 3:b.state = 1:LampState(nr) = 2
        Case 4:a.state = 0:LampState(nr) = 3
        Case 5:a.state = 1:LampState(nr) = 1
    End Select
End Sub

Sub MNFadeL(nr, a, c)
    Select Case LampState(nr)
        Case 4:a.state = 0:LampState(nr) = 0:SetFLamp c, LampState(c)
        Case 5:a.State = 1:LampState(nr) = 1
    End Select
End Sub

Sub MNFadeLm(nr, a, c)
    Select Case LampState(nr)
        Case 4:a.state = 0:SetFLamp c, LampState(c)
        Case 5:a.State = 1
    End Select
End Sub


 ' Rules
 Dim Msg(24)
 Sub Rules()
    Msg(0)="Time Machine - Data East 1988" &Chr(10)&Chr(10)
    Msg(1)=""
    Msg(2)="STARWARP - Ramps & completing top lanes spot letter in STARWARP to lite"
    Msg(3)=" 1 million point center ramp shot."
    Msg(4)=""
    Msg(5)="TIME TRAVEL TO MULTIBALL -  Triangles for the 70's, circles for the 60's, squares"
    Msg(6)=" for the 50's & shoot left and right ramp to go back in time. Play multiball"
    Msg(7)=" twice in one game to lite extra ball on outlanes."
    Msg(8)=""
    Msg(9)="JACKPOT - In multiball shoot targets to advanced value and center ramp scores SPECIAL."
    Msg(10)=""
    Msg(11)="MINI JACKPOT - Shoot STARWARP 3 times for mini JACKPOT 100K 1st ball,"
    Msg(12)=" 200K 2nd ball, 300K 3rd ball."
    Msg(13)=""
    Msg(14)="TOP LANES - score back panel value, spots STARWARP & 6X lites EXTRA BALL."
    Msg(15)=""
    Msg(16)="LASER LICK - flashing left or right ramp re-lights laserkick."
    Msg(17)=""
    Msg(18)="Pops add energy value, return lane to opposite ramp to score"
    Msg(19)=" Einstein shot energy value."
    Msg(20)=""
    Msg(21)=""
    Msg(22)=""
    Msg(23)=""
    For X=1 To 24
       Msg(0)=Msg(0)+Msg(X)&Chr(13)
    Next
    MsgBox Msg(0), , " Instructions and Rule Card"
 End Sub

'******************************************
' Use the motor callback to call div subs
'******************************************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates()
  RollingSound
End Sub

' kickers used in creating balls

Sub BallRelease_Unhit: NewBallID: End Sub
Sub Switch45_Unhit: newBallID: End Sub



'****************************************
 Dim tnopb, nosf

 tnopb = 3  ' <<<<< SET to the "Total Number Of Possible Balls" in play at any one time
 nosf = 7 ' <<<<< SET to the "Number Of Sound Files" used / B2B collision volume levels

 Dim currentball(3), ballStatus(3)
 Dim iball, cnt, coff, errMessage

 XYdata.interval = 1      ' Timer interval starts at 1 for the highest ball data sample rate
 coff = False       ' Collision off set to false

 For cnt = 0 to ubound(ballStatus)  ' Initialize/clear all ball stats, 1 = active, 0 = non-existant
  ballStatus(cnt) = 0
 Next

 '======================================================
 ' <<<<<<<<<<<<<< Ball Identification >>>>>>>>>>>>>>
 '======================================================
 ' Call this sub from every kicker(or plunger) that creates a ball.
 Sub NewBallID            ' Assign new ball object and give it ID for tracking
  For cnt = 1 to ubound(ballStatus)   ' Loop through all possible ball IDs
      If ballStatus(cnt) = 0 Then     ' If ball ID is available...
      Set currentball(cnt) = ActiveBall     ' Set ball object with the first available ID
      currentball(cnt).uservalue = cnt      ' Assign the ball's uservalue to it's new ID
      ballStatus(cnt) = 1       ' Mark this ball status active
      ballStatus(0) = ballStatus(0)+1     ' Increment ballStatus(0), the number of active balls
  If coff = False Then        ' If collision off, overrides auto-turn on collision detection
              ' If more than one ball active, start collision detection process
  If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
  End If
  Exit For          ' New ball ID assigned, exit loop
      End If
      Next
    'Debugger           ' For demo only, display stats
 End Sub

 ' Call this sub from every kicker that destroys a ball, before the ball is destroyed.
 Sub ClearBallID
    On Error Resume Next        ' Error handling for debugging purposes
      iball = ActiveBall.uservalue      ' Get the ball ID to be cleared
      currentball(iball).UserValue = 0      ' Clear the ball ID
      If Err Then Msgbox Err.description & vbCrLf & iball
      ballStatus(iBall) = 0         ' Clear the ball status
      ballStatus(0) = ballStatus(0)-1     ' Subtract 1 ball from the # of balls in play
      On Error Goto 0
 End Sub

 '=====================================================
 ' <<<<<<<<<<<<<<<<< XYdata_Timer >>>>>>>>>>>>>>>>>
 '=====================================================
 ' Ball data collection and B2B Collision detection.
 ReDim baX(tnopb,4), baY(tnopb,4), bVx(tnopb,4), bVy(tnopb,4), TotalVel(tnopb,4)
 Dim cForce, bDistance, xyTime, cFactor, id, id2, id3, B1, B2

 Sub XYdata_Timer()
  ' xyTime... Timers will not loop or start over 'til it's code is finished executing. To maximize
  ' performance, at the end of this timer, if the timer's interval is shorter than the individual
  ' computer can handle this timer's interval will increment by 1 millisecond.
     xyTime = Timer+(XYdata.interval*.001)  ' xyTime is the system timer plus the current interval time
  ' Ball Data... When a collision occurs a ball's velocity is often less than it's velocity before the
  ' collision, if not zero. So the ball data is sampled and saved for four timer cycles.
      If id2 >= 4 Then id2 = 0            ' Loop four times and start over
      id2 = id2+1               ' Increment the ball sampler ID
      For id = 1 to ubound(ballStatus)          ' Loop once for each possible ball
    If ballStatus(id) = 1 Then            ' If ball is active...
        baX(id,id2) = round(currentball(id).x,2)        ' Sample x-coord
        baY(id,id2) = round(currentball(id).y,2)        ' Sample y-coord
        bVx(id,id2) = round(currentball(id).velx,2)       ' Sample x-velocity
        bVy(id,id2) = round(currentball(id).vely,2)       ' Sample y-velocity
        TotalVel(id,id2) = (bVx(id,id2)^2+bVy(id,id2)^2)    ' Calculate total velocity
      If TotalVel(id,id2) > TotalVel(0,0) Then TotalVel(0,0) = int(TotalVel(id,id2))
      End If
      Next
  ' Collision Detection Loop - check all possible ball combinations for a collision.
  ' bDistance automatically sets the distance between two colliding balls. Zero milimeters between
  ' balls would be perfect, but because of timing issues with ball velocity, fast-traveling balls
  ' prevent a low setting from always working, so bDistance becomes more of a sensitivity setting,
  ' which is automated with calculations using the balls' velocities.
  ' Ball x/y-coords plus the bDistance determines B2B proximity and triggers a collision.
  id3 = id2 : B2 = 2 : B1 = 1           ' Set up the counters for looping
  Do
  If ballStatus(B1) = 1 and ballStatus(B2) = 1 Then     ' If both balls are active...
    bDistance = int((TotalVel(B1,id3)+TotalVel(B2,id3))^1.04)
    If ((baX(B1,id3)-baX(B2,id3))^2+(baY(B1,id3)-baY(B2,id3))^2)<2800+bDistance Then collide B1,B2 : Exit Sub
    End If
    B1 = B1+1             ' Increment ball1
    If B1 >= ballStatus(0) Then Exit Do       ' Exit loop if all ball combinations checked
    If B1 >= B2 then B1 = 1:B2 = B2+1       ' If ball1 >= reset ball1 and increment ball2
  Loop

    If ballStatus(0) <= 1 Then XYdata.enabled = False       ' Turn off timer if one ball or less

  If XYdata.interval >= 40 Then coff = True : XYdata.enabled = False  ' Auto-shut off
  If Timer > xyTime * 3 Then coff = True : XYdata.enabled = False   ' Auto-shut off
      If Timer > xyTime Then XYdata.interval = XYdata.interval+1    ' Increment interval if needed
 End Sub

 ' Ball data collection and B2B Collision detection. jpsalas: added height check
ReDim baX(tnopb, 4), baY(tnopb, 4), baZ(tnopb, 4), bVx(tnopb, 4), bVy(tnopb, 4), TotalVel(tnopb, 4)

Sub XYdata_Timer()
    xyTime = Timer + (XYdata.interval * .001)
    If id2 >= 4 Then id2 = 0
    id2 = id2 + 1
    For id = 1 to ubound(ballStatus)
        If ballStatus(id) = 1 Then
            baX(id, id2) = round(currentball(id).x, 2)
            baY(id, id2) = round(currentball(id).y, 2)
            baZ(id, id2) = round(currentball(id).z, 2)
            bVx(id, id2) = round(currentball(id).velx, 2)
            bVy(id, id2) = round(currentball(id).vely, 2)
            TotalVel(id, id2) = (bVx(id, id2) ^2 + bVy(id, id2) ^2)
            If TotalVel(id, id2) > TotalVel(0, 0) Then TotalVel(0, 0) = int(TotalVel(id, id2) )
        End If
    Next

    id3 = id2:B2 = 2:B1 = 1
    Do
        If ballStatus(B1) = 1 and ballStatus(B2) = 1 Then
            bDistance = int((TotalVel(B1, id3) + TotalVel(B2, id3) ) ^1.04)
      If ABS(baZ(B1, id3) - baZ(B2, id3)) < 50 Then
        If((baX(B1, id3) - baX(B2, id3) ) ^2 + (baY(B1, id3) - baY(B2, id3) ) ^2) < 2800 + bDistance Then collide B1, B2:Exit Sub
      End If
        End If
        B1 = B1 + 1
        If B1 >= ballStatus(0) Then Exit Do
        If B1 >= B2 then B1 = 1:B2 = B2 + 1
    Loop

    If ballStatus(0) <= 1 Then XYdata.enabled = False

    If XYdata.interval >= 40 Then coff = True:XYdata.enabled = False
    If Timer > xyTime * 3 Then coff = True:XYdata.enabled = False
    If Timer > xyTime Then XYdata.interval = XYdata.interval + 1
End Sub

'Calculate the collision force and play sound
Dim cTime, cb1, cb2, avgBallx, cAngle, bAngle1, bAngle2

Sub Collide(cb1, cb2)
    If TotalVel(0, 0) / 1.8 > cFactor Then cFactor = int(TotalVel(0, 0) / 1.8)
    avgBallx = (bvX(cb2, 1) + bvX(cb2, 2) + bvX(cb2, 3) + bvX(cb2, 4) ) / 4
    If avgBallx < bvX(cb2, id2) + .1 and avgBallx > bvX(cb2, id2) -.1 Then
        If ABS(TotalVel(cb1, id2) - TotalVel(cb2, id2) ) < .000005 Then Exit Sub
    End If
    If Timer < cTime Then Exit Sub
    cTime = Timer + .1
    GetAngle baX(cb1, id3) - baX(cb2, id3), baY(cb1, id3) - baY(cb2, id3), cAngle
    id3 = id3 - 1:If id3 = 0 Then id3 = 4
    GetAngle bVx(cb1, id3), bVy(cb1, id3), bAngle1
    GetAngle bVx(cb2, id3), bVy(cb2, id3), bAngle2
    cForce = Cint((abs(TotalVel(cb1, id3) * Cos(cAngle-bAngle1) ) + abs(TotalVel(cb2, id3) * Cos(cAngle-bAngle2) ) ) )
    If cForce < 4 Then Exit Sub
    cForce = Cint((cForce) / (cFactor / nosf) )
    If cForce > nosf-1 Then cForce = nosf-1
    PlaySound("collide" & cForce)
End Sub

' Get angle
Dim Xin, Yin, rAngle, Radit, wAngle, Pi
Pi = Round(4 * Atn(1), 6) '3.1415926535897932384626433832795

Sub GetAngle(Xin, Yin, wAngle)
    If Sgn(Xin) = 0 Then
        If Sgn(Yin) = 1 Then rAngle = 3 * Pi / 2 Else rAngle = Pi / 2
        If Sgn(Yin) = 0 Then rAngle = 0
        Else
            rAngle = atn(- Yin / Xin)
    End If
    If sgn(Xin) = -1 Then Radit = Pi Else Radit = 0
    If sgn(Xin) = 1 and sgn(Yin) = 1 Then Radit = 2 * Pi
    wAngle = round((Radit + rAngle), 4)
End Sub

 '=================================================
 ' <<<<<<<< GetAngle(X, Y, Anglename) >>>>>>>>
 '=================================================
 ' A repeated function which takes any set of coordinates or velocities and calculates an angle in radians.
 Pi = Round(4*Atn(1),6)         '3.1415926535897932384626433832795

 Sub GetAngle(Xin, Yin, wAngle)
    If Sgn(Xin) = 0 Then
      If Sgn(Yin) = 1 Then rAngle = 3 * Pi/2 Else rAngle = Pi/2
      If Sgn(Yin) = 0 Then rAngle = 0
    Else
      rAngle = atn(-Yin/Xin)      ' Calculates angle in radians before quadrant data
    End If
    If sgn(Xin) = -1 Then Radit = Pi Else Radit = 0
    If sgn(Xin) = 1 and sgn(Yin) = 1 Then Radit = 2 * Pi
    wAngle = round((Radit + rAngle),4)    ' Calculates angle in radians with quadrant data
  '"wAngle = round((180/Pi) * (Radit + rAngle),4)" ' Will convert radian measurements to degrees - to be used in future
 End Sub


 '*****************
 '**   sounds    **
 '*****************

  Sub Rhelp_Hit() : PlaySoundAtVol "ballhit", ActiveBall, 1 : End Sub
  Sub Lhelp_Hit() : PlaySoundAtVol "ballhit", ActiveBall, 1 : End Sub
' Sub Targets_Hit(idx):PlaySound"SwitchTarget":End Sub

  Sub Gates_Hit(idx):PlaySoundAtVol"Gate", ActiveBall, 1:End Sub

  Sub Metal_Hit (idx):RandomSoundMetal():End Sub
  Sub RandomSoundMetal()
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySound "metal_1", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
      Case 2 : PlaySound "metal_2", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
      Case 3 : PlaySound "metal_3", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
      Case 4 : PlaySound "metal_4", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
    End Select
  End Sub
  Sub Rubbers_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
      PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*20*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 1 AND finalspeed <= 20 then
      RandomSoundRubber()
    End If
  End Sub
  Sub RandomSoundRubber()
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*13*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
      Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*11*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
      Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*12*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
    End Select
  End Sub

Sub TriggerRampA_Hit:
  if TRA=0 and TRAA=0 then playsound "Plasticramp", 2, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):TRA=1:exit sub
  if (TRA=1 and TRAA=1) or (TRA=1 and TRAA=0) then stopsound "Plasticramp":TRA=0:TRAA=0
End Sub
Sub TriggerRampAA_Unhit:
  TRAA=1
End Sub
Sub TriggerRampA1_Hit:
  stopsound "Plasticramp":TRA=0:TRAA=0
End Sub
Sub TRAAS_Unhit:
  TRA=0:TRAA=0:stopsound "Plasticramp"
  playsound "Wirerampright", 2, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub TriggerRampB_Hit:
  if TRB=0 and TRBB=0 then playsound "Plasticramp2", 2, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):TRB=1:exit sub
  if (TRB=1 and TRBB=1) or (TRB=1 and TRBB=0) then stopsound "Plasticramp2":TRB=0:TRBB=0
End Sub
Sub TriggerRampBB_Unhit:
  TRBB=1
End Sub
Sub TriggerRampB1_Hit:
  stopsound "Plasticramp2":TRB=0:TRBB=0
End Sub
Sub TRBBS_hit:
  TRB=0:TRBB=0:stopsound "Plasticramp2"
  playsound "Wirerampright", 2, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub
Sub TRBBS1_hit:
  debug.print "metal2 ok"
  playsound "Wireramp_Stop", 0, Vol(ActiveBall)+5, Pan(ActiveBall), 0, Pitch(ActiveBall)+10, 1, 0, AudioFade(ActiveBall)
End Sub

Sub TriggerRampC_Hit:
  if TRC=0 and TRCC=0 then playsound "Plasticramp3", 2, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):TRC=1:exit sub
  if (TRC=1 and TRCC=1) or (TRC=1 and TRCC=0) then stopsound "Plasticramp3":TRC=0:TRCC=0
End Sub
Sub TriggerRampCC_Unhit:
  TRCC=1
End Sub
Sub TriggerRampA1_Hit:
  stopsound "Plasticramp3":TRC=0:TRCC=0
End Sub
Sub TRCCS_Unhit:
  TRC=0:TRCC=0:stopsound "Plasticramp3"
  playsound "Wirerampright", 2, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub MetTrig01_Hit()
PlaySound "metal_1", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub
Sub MetTrig02_Hit()
PlaySound "metal_2", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub
Sub MetTrig03_Hit()
PlaySound "metal_3", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub
Sub MetTrig04_Hit()
PlaySound "metal_4", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub
Sub MetTrig05_Hit()
PlaySound "metal_2", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub
Sub MetTrig06_Hit()
PlaySound "metal_4", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub WoodTrig01_Hit()
PlaySound "woodhit", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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

' Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Time_Machine" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / Time_Machine.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Time_Machine" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / Time_Machine.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Time_Machine" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / Time_Machine.width-1
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

Const tnob = 7 ' total number of balls
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
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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


'********************************************************
'********    Flupper's subs for his Flashers    *********
'********************************************************

'*** white flasher ***
Sub FlasherFlash1_Timer()
  dim flashx3, matdim
  If not FlasherFlash1.TimerEnabled Then
    FlasherFlash1.TimerEnabled = True
    FlasherFlash1.visible = 1
    FlasherLit1.visible = 1
Yellowflash.state = 1
  End If
  flashx3 = FlashLevel1 * FlashLevel1 * FlashLevel1
  Flasherflash1.opacity = 1000 * flashx3
  FlasherLit1.BlendDisableLighting = 10 * flashx3
  Flasherbase1.BlendDisableLighting =  flashx3
  FlasherLight1.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel1)
  FlasherLit1.material = "domelit" & matdim
  FlashLevel1 = FlashLevel1 * 0.85 - 0.01
  If FlashLevel1 < 0.15 Then
    FlasherLit1.visible = 0
  Else
    FlasherLit1.visible = 1
  end If
  If FlashLevel1 < 0 Then
    FlasherFlash1.TimerEnabled = False
    FlasherFlash1.visible = 0
Yellowflash.state = 0
  End If
End Sub

'*** Red flasher ***
Sub FlasherFlash2_Timer()
  dim flashx3, matdim
  If not Flasherflash2.TimerEnabled Then
    Flasherflash2.TimerEnabled = True
    Flasherflash2.visible = 1
    Flasherlit2.visible = 1
Blueflash.state = 1
  End If
  flashx3 = FlashLevel2 * FlashLevel2 * FlashLevel2
  Flasherflash2.opacity = 1500 * flashx3
  Flasherlit2.BlendDisableLighting = 10 * flashx3
  Flasherbase2.BlendDisableLighting =  flashx3
  Flasherlight2.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel2)
  Flasherlit2.material = "domelit" & matdim
  FlashLevel2 = FlashLevel2 * 0.9 - 0.01
  If FlashLevel2 < 0.15 Then
    Flasherlit2.visible = 0
  Else
    Flasherlit2.visible = 1
  end If
  If FlashLevel2 < 0 Then
    Flasherflash2.TimerEnabled = False
    Flasherflash2.visible = 0
Blueflash.state = 0
  End If
End Sub

'*** blue flasher ***
Sub FlasherFlash3_Timer()
  dim flashx3, matdim
  If not Flasherflash3.TimerEnabled Then
    Flasherflash3.TimerEnabled = True
    Flasherflash3.visible = 1
    Flasherlit3.visible = 1
Whiteflash.state = 1
  End If
  flashx3 = FlashLevel3 * FlashLevel3 * FlashLevel3
  Flasherflash3.opacity = 8000 * flashx3
  Flasherlit3.BlendDisableLighting = 10 * flashx3
  Flasherbase3.BlendDisableLighting =  flashx3
  Flasherlight3.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel3)
  Flasherlit3.material = "domelit" & matdim
  FlashLevel3 = FlashLevel3 * 0.85 - 0.01
  If FlashLevel3 < 0.15 Then
    Flasherlit3.visible = 0
  Else
    Flasherlit3.visible = 1
  end If
  If FlashLevel3 < 0 Then
    Flasherflash3.TimerEnabled = False
    Flasherflash3.visible = 0
Whiteflash.state = 0
  End If
End Sub

'*** blue flasher vertical (script is the same as for blue Flasher) ***
Sub FlasherFlash4_Timer()
  dim flashx3, matdim
  If not Flasherflash4.TimerEnabled Then
    Flasherflash4.TimerEnabled = True
    Flasherflash4.visible = 1
    Flasherlit4.visible = 1
Greenflash.state = 1
  End If
  flashx3 = FlashLevel4 * FlashLevel4 * FlashLevel4
  Flasherflash4.opacity = 8000 * flashx3
  Flasherlit4.BlendDisableLighting = 10 * flashx3
  Flasherbase4.BlendDisableLighting =  flashx3
  Flasherlight4.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel4)
  Flasherlit4.material = "domelit" & matdim
  FlashLevel4 = FlashLevel4 * 0.85 - 0.01
  If FlashLevel4 < 0.15 Then
    Flasherlit4.visible = 0
  Else
    Flasherlit4.visible = 1
  end If
  If FlashLevel4 < 0 Then
    Flasherflash4.TimerEnabled = False
    Flasherflash4.visible = 0
Greenflash.state = 0
  End If
End Sub


'************************************************************************

'***************************************
'**     Data-east Flippers Ninuzzu    **
'***************************************

Sub shfl_Timer()
  LeftFlipperP.RotZ = LeftFlipper.CurrentAngle
  LeftFlipperSh.RotZ = LeftFlipper.currentangle
  RightFlipperP.RotZ = RightFlipper.CurrentAngle
  RightFlipperSh.RotZ = RightFlipper.currentangle

  if ABS(Vlock.balls) = 1 and vla = 0 then
      playsound "Wirestop", 1, GlobalSoundLevel, -0,2, 0, , 1, 0
    vla=1:vlb=0:vlc=0:vlaa=0:vlbb=0:vlcc=0
  end if

  if ABS(Vlock.balls) = 2 and vla = 1 and vlb = 0 then
    playsound "Wirestopball", 1, GlobalSoundLevel, -0,2, 0, , 1, 0
    vla=1:vlb=1:vlc=0:vlaa=0:vlbb=0:vlcc=0
  end if
  if ABS(Vlock.balls) = 3 and vla = 1 and vlb = 1 and vlc = 0 then
    playsound "Wirestopball", 1, GlobalSoundLevel, -0,2, 0, , 1, 0
    vla=1:vlb=1:vlc=1:vlaa=0:vlbb=0:vlcc=1
  end if

  if ABS(Vlock.balls) = 1 and vla = 1 and vlb = 1 and vlc=0 and vlbb=1 then
      playsound "Wirestop", 1, GlobalSoundLevel, -0,2, 0, , 1, 0
    vla=1:vlb=0:vlc=0:vlcc=0:vlbb=0:vlaa=1
  end if

  if ABS(Vlock.balls) = 2 and vla = 1 and vlb = 1 and vlc=1 and vlcc=1 then
    playsound "Wirestopball", 1, GlobalSoundLevel, -0,2, 0, , 1, 0
    vla=1:vlb=1:vlc=0:vlcc=0:vlbb=1
  end if
  if ABS(Vlock.balls) = 0 then
    vla=0:vlb=0:vlc=0:vlcc=0:vlbb=0:vlaa=0
  end if

 if numberofballsinhole >0 then SW45Wait.enabled = 1:end if
End Sub


'***************************************
'**      JPJ's Animation Bumpers      **
'***************************************
dim circle, circle1, circle2, z0, z1, z2
circle=0
circle1=0
circle2=0

Sub BumpCircle_Timer()
Select Case circle
    Case 0
      z0=Int(Rnd*360)+1
      primitive6.roty=z0
      Primitive6.image = "bumpcircle1"
    Case 1
      z0=Int(Rnd*360)+1
      primitive6.roty=z0
      Primitive6.image = "bumpcircle2"
    Case 2
      z0=Int(Rnd*360)+1
      primitive6.roty=z0
      Primitive6.image = "bumpcircle1"
    Case 3
      z0=Int(Rnd*360)+1
      primitive6.roty=z0
      Primitive6.image = "bumpcircle2"
    Case 4
      z0=Int(Rnd*360)+1
      primitive6.roty=z0
      Primitive6.image = "bumpcircle2"
    Case 5
      z0=Int(Rnd*360)+1
      primitive6.roty=z0
      Primitive6.image = "bumpcircle2"
    Case 6
      z0=Int(Rnd*360)+1
      primitive6.roty=z0
      Primitive6.image = "bumpcircle1"
    Case 7
      z0=Int(Rnd*360)+1
      primitive6.roty=z0
      Primitive6.image = "bumpcircle0"
    Case 8
      z0=Int(Rnd*360)+1
      primitive6.roty=z0
      Primitive6.image = "bumpcircle2"
    Case 9
      z0=Int(Rnd*360)+1
      primitive6.roty=z0
      Primitive6.image = "bumpcircle0"
  End Select
      circle=circle+1
  if circle = 10 then circle = 0:me.enabled = 0:end If
End Sub

Sub BumpCircle1_Timer()
Select Case circle1
    Case 0
      z1=Int(Rnd*360)+1
      primitive7.roty=z1
      Primitive7.image = "bumpcircle1"
    Case 1
      z1=Int(Rnd*360)+1
      primitive7.roty=z1
      Primitive7.image = "bumpcircle2"
    Case 2
      z1=Int(Rnd*360)+1
      primitive7.roty=z1
      Primitive7.image = "bumpcircle1"
    Case 3
      z1=Int(Rnd*360)+1
      primitive7.roty=z1
      Primitive7.image = "bumpcircle1"
    Case 4
      z1=Int(Rnd*360)+1
      primitive7.roty=z1
      Primitive7.image = "bumpcircle1"
    Case 5
      z1=Int(Rnd*360)+1
      primitive7.roty=z1
      Primitive7.image = "bumpcircle2"
    Case 6
      z1=Int(Rnd*360)+1
      primitive7.roty=z1
      Primitive7.image = "bumpcircle2"
    Case 7
      z1=Int(Rnd*360)+1
      primitive7.roty=z1
      Primitive7.image = "bumpcircle1"
    Case 8
      z1=Int(Rnd*360)+1
      primitive7.roty=z1
      Primitive7.image = "bumpcircle2"
    Case 9
      z1=Int(Rnd*360)+1
      primitive7.roty=z1
      Primitive7.image = "bumpcircle0"
  End Select
      circle1=circle1+1
  if circle1 = 10 then circle1 = 0:me.enabled = 0:end If
End Sub

Sub BumpCircle2_Timer()
Select Case circle2
    Case 0
      z2=Int(Rnd*360)+1
      primitive9.roty=z2
      Primitive9.image = "bumpcircle1"
    Case 1
      z2=Int(Rnd*360)+1
      primitive9.roty=z2
      Primitive9.image = "bumpcircle0"
    Case 2
      z2=Int(Rnd*360)+1
      primitive9.roty=z2
      Primitive9.image = "bumpcircle1"
    Case 3
      z2=Int(Rnd*360)+1
      primitive9.roty=z2
      Primitive9.image = "bumpcircle2"
    Case 4
      z2=Int(Rnd*360)+1
      primitive9.roty=z2
      Primitive9.image = "bumpcircle2"
    Case 5
      z2=Int(Rnd*360)+1
      primitive9.roty=z2
      Primitive9.image = "bumpcircle2"
    Case 6
      z2=Int(Rnd*360)+1
      primitive9.roty=z2
      Primitive9.image = "bumpcircle1"
    Case 7
      z2=Int(Rnd*360)+1
      primitive9.roty=z2
      Primitive9.image = "bumpcircle1"
    Case 8
      z2=Int(Rnd*360)+1
      primitive9.roty=z2
      Primitive9.image = "bumpcircle2"
    Case 9
      z2=Int(Rnd*360)+1
      primitive9.roty=z2
      Primitive9.image = "bumpcircle0"
  End Select
      circle2=circle2+1
  if circle2 = 10 then circle2 = 0:me.enabled = 0:end If
End Sub


'********************************************
'*  Led part, adapted from JP Salas script  *
'********************************************

Dim obj, rGreen, rRed, rBlue, RGBFactor, RGBStep

Sub RainbowTimer_Timer 'rainbow led light color changing
  RGBFactor =20
            Select Case RGBStep
              Case 0 'Green
                rGreen = rGreen + RGBFactor
                If rGreen > 255 then
                  rGreen = 255
                  RGBStep = 1
                End If
              Case 1 'Red
                rRed = rRed - RGBFactor
                If rRed < 0 then
                  rRed = 0
                  RGBStep = 2
                End If
              Case 2 'Blue
                rBlue = rBlue + RGBFactor
                If rBlue > 255 then
                  rBlue = 255
                  RGBStep = 3
                End If
              Case 3 'Green
                rGreen = rGreen - RGBFactor
                If rGreen < 0 then
                  rGreen = 0
                  RGBStep = 4
                End If
              Case 4 'Red
                rRed = rRed + RGBFactor
                If rRed > 255 then
                  rRed = 255
                  RGBStep = 5
                End If
              Case 5 'Blue
                rBlue = rBlue - RGBFactor
                If rBlue < 0 then
                  rBlue = 0
                  RGBStep = 0
                End If
            End Select


    if BallinPlunger=1 then
              For each obj in Led1
                obj.color = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
                obj.colorfull = RGB(rRed, rGreen, rBlue)
              Next
          PRefresh.State = 1
        Else
          PRefresh.state = 0
    end if
  If led = 1 Then
        For each obj in RainbowLights
          obj.color = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
          obj.colorfull = RGB(rRed, rGreen, rBlue)
        Next
        exit Sub
  end if

  if led = 0 and f3a.state = 1 or f1b.state = 1 or f2a.state = 1 then
          For each obj in RainbowLights
            obj.color = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
            obj.colorfull = RGB(rRed, rGreen, rBlue)
          Next
    else
      if led = 0 and f3a.state = 0 and f1b.state = 0 and f2a.state = 0 and l18.state = 0 then
          For each obj in RainbowLights
            obj.color = RGB(210, 251, 255)
            obj.colorfull = RGB(255, 255, 255)
          Next
      end if
    end if

      if led = 0 and l18.state=1 and bstrough.balls<3 then
          For each obj in RainbowLights
            obj.color = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
            obj.colorfull = RGB(rRed, rGreen, rBlue)
          Next
    else
      if led = 0 and l18.state = 0 and f3a.state = 0 and f1b.state = 0 and f2a.state = 0 then
          For each obj in RainbowLights
            obj.color = RGB(210, 251, 255)
            obj.colorfull = RGB(255, 255, 255)
          Next
      end if
    end if
End Sub


Sub solanim_timer()
  Select Case animA
    Case 0
      Primitive49.rotx=80
      SolMetal.TransZ = 5
      SolSpring.Size_Z = 90
    Case 1
      Primitive49.rotx=85
      SolMetal.TransZ = 10
      SolSpring.Size_Z = 79
    Case 2
      Primitive49.rotx=90
      SolMetal.TransZ = 15
      SolSpring.Size_Z = 69
    Case 3
      Primitive49.rotx=95
      SolMetal.TransZ = 20
      SolSpring.Size_Z = 58
    Case 4
      Primitive49.rotx=100
      SolMetal.TransZ = 25
      SolSpring.Size_Z = 48
    Case 5
      Primitive49.rotx=105
      SolMetal.TransZ = 30
      SolSpring.Size_Z = 37
    Case 6
      Primitive49.rotx=110
      SolMetal.TransZ = 35
      SolSpring.Size_Z = 27
    Case 7
      Primitive49.rotx=115
      SolMetal.TransZ = 35
      SolSpring.Size_Z = 25
    Case 8
      Primitive49.rotx=107
      SolMetal.TransZ = 30
      SolSpring.Size_Z = 37
    Case 9
      Primitive49.rotx=100
      SolMetal.TransZ = 25
      SolSpring.Size_Z = 48
    Case 10
      Primitive49.rotx=95
      SolMetal.TransZ = 20
      SolSpring.Size_Z = 58
    Case 11
      Primitive49.rotx=90
      SolMetal.TransZ = 15
      SolSpring.Size_Z = 69
    Case 12
      Primitive49.rotx=85
      SolMetal.TransZ = 10
      SolSpring.Size_Z = 79
    Case 13
      Primitive49.rotx=80
      SolMetal.TransZ = 5
      SolSpring.Size_Z = 90
    Case 14
      Primitive49.rotx=75
      SolMetal.TransZ = 0
      SolSpring.Size_Z = 100
  End Select
      animA=animA+1
  if animA = 15 then animA = 0:me.enabled = 0:end If
end Sub


'cFastFlips by nFozzy
'Bypasses pinmame callback for faster and more responsive flippers
'Version 1.1 beta2 (More proper behaviour, extra safety against script errors)
'*************************************************
Function NullFunction(aEnabled):End Function    '1 argument null function placeholder
Class cFastFlips
    Public TiltObjects, DebugOn, hi
    Private SubL, SubUL, SubR, SubUR, FlippersEnabled, Delay, LagCompensation, Name, FlipState(3)

    Private Sub Class_Initialize()
        Delay = 0 : FlippersEnabled = False : DebugOn = False : LagCompensation = False
        Set SubL = GetRef("NullFunction"): Set SubR = GetRef("NullFunction") : Set SubUL = GetRef("NullFunction"): Set SubUR = GetRef("NullFunction")
    End Sub

    'set callbacks
    Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : Decouple sLLFlipper, aInput: End Property
    Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : End Property
    Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : Decouple sLRFlipper, aInput:  End Property
    Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : End Property
    Public Sub InitDelay(aName, aDelay) : Name = aName : delay = aDelay : End Sub   'Create Delay
    'Automatically decouple flipper solcallback script lines (only if both are pointing to the same sub) thanks gtxjoe
    Private Sub Decouple(aSolType, aInput)  : If StrComp(SolCallback(aSolType),aInput,1) = 0 then SolCallback(aSolType) = Empty End If : End Sub

    'call callbacks
    Public Sub FlipL(aEnabled)
        FlipState(0) = aEnabled 'track flipper button states: the game-on sol flips immediately if the button is held down (1.1)
        If not FlippersEnabled and not DebugOn then Exit Sub
        subL aEnabled
    End Sub

    Public Sub FlipR(aEnabled)
        FlipState(1) = aEnabled
        If not FlippersEnabled and not DebugOn then Exit Sub
        subR aEnabled
    End Sub

    Public Sub FlipUL(aEnabled)
        FlipState(2) = aEnabled
        If not FlippersEnabled and not DebugOn then Exit Sub
        subUL aEnabled
    End Sub

    Public Sub FlipUR(aEnabled)
        FlipState(3) = aEnabled
        If not FlippersEnabled and not DebugOn then Exit Sub
        subUR aEnabled
    End Sub

    Public Sub TiltSol(aEnabled)    'Handle solenoid / Delay (if delayinit)
        If delay > 0 and not aEnabled then  'handle delay
            vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
            LagCompensation = True
        else
            If Delay > 0 then LagCompensation = False
            EnableFlippers(aEnabled)
        end If
    End Sub

    Sub FireDelay() : If LagCompensation then EnableFlippers False End If : End Sub

    Private Sub EnableFlippers(aEnabled)
        If aEnabled then SubL FlipState(0) : SubR FlipState(1) : subUL FlipState(2) : subUR FlipState(3)
        FlippersEnabled = aEnabled
        If TiltObjects then vpmnudge.solgameon aEnabled
        If Not aEnabled then
            subL False
            subR False
            If not IsEmpty(subUL) then subUL False
            If not IsEmpty(subUR) then subUR False
        End If
    End Sub


    End Class

' Thalamus : Exit in a clean and proper way
Sub Time_Machine_exit
  Controller.Pause = False
  Controller.Stop
End Sub

' Thalamus, this script is a mess - it has the callback for ball collitions but also ancient code for the same

