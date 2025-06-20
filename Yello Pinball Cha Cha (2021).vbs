' *********************************************************************
' *********************************************************************
' **              -----------------------------------                **
' **            /                                     \              **
' **           |       ******             ******       |             **
' **           |        ******           ******        |             **
' **           |         ******         ******         |             **
' **           |          ******       ******          |             **
' **           |           ******     ******           |             **
' **           |            ******   ******            |             **
' **           |             ****** ******             |             **
' **           |              ***********              |             **
' **           |                *******                |             **
' **           |                *******                |             **
' **           |                *******                |             **
' **           |                *******                |             **
' **           |                *******                |             **
' **           |                *******                |             **
' **           |                *******                |             **
' **            \                                     /              **
' **              -----------------------------------                **
' **                    YELLO - PINBALL CHA CHA                      **
' **    >>DEDICATED TO FELLOW YELLO MUSIC LOVERS EVERYWHERE<<        **
' **                      Table Script v1.0                          **
' **       (C) 2021  Art and Visuals  Mark Paulik Marcade MODs       **
' **                shoeboxtheater.tumblr.com                        **
' **                                                                 **
' **                       Many Special thanks:                      **
' **              32ASSASSIN (DEFENDER WILLIAMS 1982) VPX            **
' **                  GTXJOE (DEFENDER WILLIAMS 1982) VP9.           **
' **                  DESTRUK (DEFENDER 1.4.0) VP8.                  **
' **              IVANTBA (TRANSFORMERS THE MOVIE 1986)              **
' **          and anyone else who contributed not listed here!       **
' **                                                                 **
' *********************************************************************
' *********************************************************************


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'**********************Options********************************

dim FlipperReturnLow, FlipperReturnHigh, FlipperReturnOverride, SoundLevelMult, DesktopMode : DesktopMode = Table1.ShowDT

SoundLevelMult = 0 'Increase table SFX (may cause some normalization)

Const HardFlips = 1
Const FastFlips = 1
const FlipperReturnMod = 1 'Extra flipper return curve
FlipperReturnLow = 81 '% adjust flipper return min/max values
FlipperReturnHigh = 110 '%
FlipperReturnOverride = 0 'This can affect the rigidity of the flipper when it's at rest. 0 = Disabled

const SingleScreenFS = 1 '1 = VPX display 2 = Vpinmame display rotated
const VPXdisplay = 0 'Enable/Disable VPX display. Disable for greater performance. (Default: 1)

'*************************************************************

Ramp15.visible = Table1.ShowDT
Ramp16.visible = Table1.ShowDT


LoadVPM "", "S7.VBS", 2.0

Const cGameName="dfndr_l4"

dim VPMversion
Set MotorCallback = GetRef("UpdateSolenoids")

Sub UpdateSolenoids
  Dim Changed, Count, funcName, ii, sol11, solNo
  Changed = Controller.ChangedSolenoids
  If Not IsEmpty(Changed) Then
    sol11 = Controller.Solenoid(11)
    Count = UBound(Changed, 1)
    For ii = 0 To Count
      solNo = Changed(ii, CHGNO)
      ' multiplex solenoid #11 fixed in VPM 1.52beta and newer
      if VPMversion < "01519901" then
        If SolNo < 11 And sol11 Then solNo = solNo + 32
      else
        ' no need to evaluate sol 11 anymore, VPM does it
        if SolNo > 50 then solNo = solNo - 18
      end if
      funcName = SolCallback(solNo)
      If funcName <> "" Then Execute funcName & " CBool(" & Changed(ii, CHGSTATE) &")"
    Next
  End If
 BRG.ObjRotZ = BallReturnGate.CurrentAngle -5
End Sub

dim BallMass : BallMass = 1
Dim gameRun
Dim yy:For each yy in GiLights:yy.intensity=7  : Next   'Global Light Collection intensity. (collection name after "in")

Const SCoin="fx_coin"

Const UseSolenoids = 0
Const UseLamps = 0
Const UseGi = 0
Const UseSync = 0
Const HandleMech = 0

'***********************************************************************************
'****             Solenoid                            ****
'***********************************************************************************

Const sDTLeft1    = 1 ' Drop Target 5-bank, Left Lander Drop Targets
Const sDTLeft2    = 2 ' Drop Target 5-bank, Left Lander Drop Targets
Const sDTLeft3    = 3 ' Drop Target 5-bank, Left Lander Drop Targets
Const sDTLeft4    = 4 ' Drop Target 5-bank, Left Lander Drop Targets
Const sDTLeft5    = 5 ' Drop Target 5-bank, Left Lander Drop Targets
Const sDTLeftRel  = 6 ' Drop Target 5-bank, Left Lander Drop Targets
Const sDTRight1   = 33  ' Drop Target 5-bank, Right Lander Drop Targets
Const sDTRight2   = 34  ' Drop Target 5-bank, Right Lander Drop Targets
Const sDTRight3   = 35  ' Drop Target 5-bank, Right Lander Drop Targets
Const sDTRight4   = 36  ' Drop Target 5-bank, Right Lander Drop Targets
Const sDTRight5   = 37  ' Drop Target 5-bank, Right Lander Drop Targets
Const sDTRightRel = 38  ' Drop Target 5-bank, Right Lander Drop Targets

Const sDTPodBot   = 7 ' First Drop Down Left Upper Lane
Const sDTPodTop   = 39  ' Second Drop Down Left Upper Lane
Const sLockupRel  = 40  ' Lock Right Lane

Const sDTBaitBot  = 9 ' Playfield Drop Target Left
Const sDTBaitTop  = 10  ' Drop Down Target Top, located at the start of the Right Lane
Const sDTBaitMid  = 41  ' Playfield Drop Target Right
Const sDTBaitRel  = 42  ' Playfield Drop Target Reset

Const BallRelease = 8 ' Ball Release
Const Drain       = 12  ' Drain

Const sKickBack   = 13  ' Left Auto Plunger
Const sGI       = 14  ' Generall Illumination
Const sBell       = 15  ' Bell Sound
Const sLJet       = 17  ' Left Jet Bumper
Const sRJet       = 18  ' Right Jet Bumper

Const sGate       = 22
Const sEnable     = 25

SolCallback(sDTLeft1)     = "dtLBank.SolUnHit 1," ' Drop Target 5-bank, Left Lander Drop Targets
SolCallback(sDTLeft2)     = "dtLBank.SolUnHit 2," ' Drop Target 5-bank, Left Lander Drop Targets
SolCallback(sDTLeft3)     = "dtLBank.SolUnHit 3," ' Drop Target 5-bank, Left Lander Drop Targets
SolCallback(sDTLeft4)     = "dtLBank.SolUnHit 4," ' Drop Target 5-bank, Left Lander Drop Targets
SolCallback(sDTLeft5)     = "dtLBank.SolUnHit 5," ' Drop Target 5-bank, Left Lander Drop Targets
SolCallback(sDTLeftRel)   = "dtLBank.SolDropDown" ' Drop Target 5-bank, Left Lander Drop Targets

SolCallback(sDTRight1)    = "dtRBank.SolUnHit 1," ' Drop Target 5-bank, Right Lander Drop Targets
SolCallback(sDTRight2)    = "dtRBank.SolUnHit 2," ' Drop Target 5-bank, Right Lander Drop Targets
SolCallback(sDTRight3)    = "dtRBank.SolUnHit 3," ' Drop Target 5-bank, Right Lander Drop Targets
SolCallback(sDTRight4)    = "dtRBank.SolUnHit 4," ' Drop Target 5-bank, Right Lander Drop Targets
SolCallback(sDTRight5)    = "dtRBank.SolUnHit 5," ' Drop Target 5-bank, Right Lander Drop Targets
SolCallback(sDTRightRel)  = "dtRBank.SolDropDown" ' Drop Target 5-bank, Right Lander Drop Targets

SolCallback(sDTBaitBot)   = "dtBait.SolUnhit 1,"  ' Playfield Drop Target Left
SolCallback(sDTBaitMid)   = "dtBait.SolUnhit 2,"
SolCallback(sDTBaitTop)   = "dtBait.SolUnhit 3,"
SolCallback(sDTBaitRel)   = "dtBait.SolDropDown"  ' Playfield Drop Target Reset

SolCallback(sDTPodBot)    = "dtBPod.SolDropUp"    ' First Drop Down Left Upper Lane
SolCallback(sDTPodTop)    = "dtTPod.SolDropUp"    ' Second Drop Down Left Upper Lane

SolCallback(sLockupRel)   = "bsLockup.SolOut"   ' Lock Right Lane

SolCallback(sKickBack)    = "SolKickBack"     ' Left Auto Plunger
SolCallback(sGI)          = "SolGI"

SolCallback(sBell)      = "vpmSolSound ""Bell"","
SolCallback(sLJet)        = "vpmSolSound ""LeftBumper_Hit"","
SolCallback(sRJet)        = "vpmSolSound ""RightBumper_Hit"","
SolCallback(sGate)        = "vpmSolDiverter BallReturnGate,True," ' Ball Return Gate to Plunger Lane
SolCallback(sEnable)    = "SolRun"


Dim bsTrough, bsLeftBallPopper, bsRightBallPopper, bsRightBallPopper42, bsRightBallPopper43, bsRightBallPopper44, bsRightLock, bsLeftLock, dtdrop, dt3bank
Dim bsLockup, dtLBank, dtRBank, dtBait, dtBaitBot, dtBPod, dtTPod

'********************************************************************************************
' Only for VPX 10.2 and higher.
' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Return to previous State
'********************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first version

    If TypeName(MyLight) = "Light" Then

        If FinalState = 2 Then
            FinalState = MyLight.State 'Keep the current light state
        End If
        MyLight.BlinkInterval = BlinkPeriod
        MyLight.Duration 2, TotalPeriod, FinalState
    ElseIf TypeName(MyLight) = "Flasher" Then

        Dim steps

        ' Store all blink information
        steps = Int(TotalPeriod / BlinkPeriod + .5) 'Number of ON/OFF steps to perform
        If FinalState = 2 Then                      'Keep the current flasher state
            FinalState = ABS(MyLight.Visible)
        End If
        MyLight.UserValue = steps * 10 + FinalState 'Store # of blinks, and final state

        ' Start blink timer and create timer subroutine
        MyLight.TimerInterval = BlinkPeriod
        MyLight.TimerEnabled = 0
        MyLight.TimerEnabled = 1
        ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
    End If
End Sub

Sub Table1_Init
  vpmInit Me
  With Controller
    .GameName = cGameName
      NVOffset(1)
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Yello Pinball Cha Cha (2021) by Marcade MODs"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Hidden = 0
    .Games(cGameName).Settings.Value("sound") = 0 '1= rotated display, 0= normal
'   .SetDisplayPosition 0,0,GetPlayerHWnd 'if you can't see the DMD then uncomment this line
    On Error Resume Next
    .Run GetPlayerHwnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With

  VPMversion=Controller.version

  If SingleScreenFS = 2 and not Table1.ShowDT then
    Controller.Games(cGameName).Settings.Value("rol") = 1
  Else
    Controller.Games(cGameName).Settings.Value("rol") = 0
    Controller.Hidden = 1
  End If
  If VPXdisplay = 0 then
    Controller.Hidden = 0
  End If

PlaySound "Intro",-0

  On Error Goto 0

  ' Nudging
  vpmNudge.TiltSwitch = 1
  vpmNudge.Sensitivity = 2
  vpmNudge.TiltObj = Array(SlingL, SlingR, bumper1, bumper2)

  'Drop Target 5-bank, Left Lander Drop Targets
  Set dtLBank = New cvpmDropTarget
  with dtLBank
    .InitDrop Array(sw13,sw14,sw15,sw16,sw17),Array(13,14,15,16,17)
    .InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
'   .CreateEvents"dtLBank"
    End With

  'Drop Target 5-bank, Right Lander Drop Targets
  Set dtRBank = New cvpmDropTarget
  with dtRBank
    .InitDrop Array(sw23,sw24,sw25,sw26,sw27),Array(23,24,25,26,27)
    .InitSnd SoundFX("DTDrop2",DOFDropTargets),SoundFX("DTReset",DOFContactors)
'   .CreateEvents"dtRBank"
    End With

  'Drop Target Left Upper Lane
  Set dtBPod = New cvpmDropTarget
  with dtBPod
    .InitDrop sw39,39
    .InitSnd SoundFX("DTDrop3",DOFContactors),SoundFX("DTReset",DOFContactors)
    End With

  Set dtTPod = New cvpmDropTarget
  with dtTPod
    .InitDrop sw40,40
    .InitSnd SoundFX("DTDrop4",DOFContactors),SoundFX("DTReset",DOFContactors)
    End With

  Set bsLockup = New cvpmBallStack
    bsLockup.InitSw 0, 42,43,44,0,0,0,0
    bsLockup.InitKick sw42,180,1
    bsLockup.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  Set dtBait = New cvpmDropTarget
  with dtBait
    .InitDrop Array(Array(sw33,sw33a),Array(sw34,sw34a),sw35),Array(33,34,35)
    .InitSnd SoundFX("DTDrop5",DOFContactors),SoundFX("DTReset",DOFContactors)
    End With

     ' Main Timer init
     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1
   LampTimer.Enabled = 1

  ' Init Kickback
    KickBack.Pullback

  'start trough
  sw49.CreateSizedBallWithMass 25, BallMass
  Sw48.CreateSizedBallWithMass 25, BallMass
  sw47.CreateSizedBallWithMass 25, BallMass

  BallSearch  'init switches

  gameRun = False
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_exit():Controller.Pause = False:Controller.Stop:End Sub
Sub Destroyer_Hit:Me.destroyball : End Sub

' Start the random music
    MusicOn

Sub MusicOn
    Dim x
    x = INT(39 * RND(1) )
    Select Case x
       Case 0:PlayMusic "Yello_01.mp3"
       Case 1:PlayMusic "Yello_02.mp3"
       Case 2:PlayMusic "Yello_03.mp3"
       Case 3:PlayMusic "Yello_04.mp3"
       Case 4:PlayMusic "Yello_05.mp3"
       Case 5:PlayMusic "Yello_06.mp3"
     Case 6:PlayMusic "Yello_07.mp3"
     Case 7:PlayMusic "Yello_08.mp3"
     Case 8:PlayMusic "Yello_09.mp3"
     Case 9:PlayMusic "Yello_10.mp3"
     Case 10:PlayMusic "Yello_11.mp3"
     Case 11:PlayMusic "Yello_12.mp3"
     Case 12:PlayMusic "Yello_13.mp3"
     Case 13:PlayMusic "Yello_14.mp3"
     Case 14:PlayMusic "Yello_15.mp3"
     Case 15:PlayMusic "Yello_16.mp3"
     Case 16:PlayMusic "Yello_17.mp3"
     Case 17:PlayMusic "Yello_18.mp3"
     Case 18:PlayMusic "Yello_19.mp3"
     Case 19:PlayMusic "Yello_20.mp3"
     Case 20:PlayMusic "Yello_21.mp3"
     Case 21:PlayMusic "Yello_22.mp3"
     Case 22:PlayMusic "Yello_23.mp3"
     Case 23:PlayMusic "Yello_24.mp3"
     Case 24:PlayMusic "Yello_25.mp3"
     Case 25:PlayMusic "Yello_26.mp3"
     Case 26:PlayMusic "Yello_27.mp3"
     Case 27:PlayMusic "Yello_28.mp3"
     Case 28:PlayMusic "Yello_29.mp3"
     Case 29:PlayMusic "Yello_30.mp3"
     Case 30:PlayMusic "Yello_31.mp3"
     Case 31:PlayMusic "Yello_32.mp3"
     Case 32:PlayMusic "Yello_33.mp3"
     Case 33:PlayMusic "Yello_34.mp3"
       Case 34:PlayMusic "Yello_35.mp3"
       Case 35:PlayMusic "Yello_36.mp3"
       Case 36:PlayMusic "Yello_37.mp3"
     Case 37:PlayMusic "Yello_38.mp3"
     Case 38:PlayMusic "Yello_39.mp3"
     Case 39:PlayMusic "Yello_40.mp3"

     End Select
 End Sub

 Sub Table1_MusicDone()
    MusicOn
 End Sub

'***********************************************************************************
'****               Plunger & Flipper                 ****
'***********************************************************************************

Sub Table1_KeyDown(ByVal keycode)
  If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAtVol "plungerpull", Plunger, 1:Plunger.Pullback
  If keycode = LeftFlipperKey Then flipnf 0, 1: exit sub : End If
  If keycode = RightFlipperKey Then flipnf 1, 1 : controller.Switch(59) = 1 : exit sub : End If
  if Keycode = KeyRules then
    If DesktopMode Then
      P_InstructionsDT.PlayAnim 0, 3*CGT/500
    Else
      P_InstructionsFS.PlayAnim 0, 3*CGT/500
    End If
    Exit Sub
  End If
  If keycode = RightMagnaSave Then controller.Switch(60) = 1
  If keycode = LeftMagnaSave Then controller.Switch(61) = 1
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then
    Plunger.Fire
    if BallInPlunger then
      If keycode = PlungerKey Then PlaySoundAtVol "plunger3", Plunger, 1:Plunger.Fire
      FlashForMs yBoxFlasher, 2000, 40, 0
    Else
      If keycode = PlungerKey Then PlaySoundAtVol "plunger3", Plunger, 1:Plunger.Fire
      FlashForMs yBoxFlasher, 2000, 40, 0
    end if
  End If
  if Keycode = KeyRules then
    If DesktopMode Then
      P_InstructionsDT.ShowFrame 0
    Else
      P_InstructionsFs.ShowFrame 0
    End If
    Exit Sub
  End If

    If keycode = LeftFlipperKey Then : flipnf 0, 0: exit sub : End If
  If keycode = RightFlipperKey Then : flipnf 1, 0: controller.Switch(59) = 0 : exit sub:End If

  If keycode = RightMagnaSave Then controller.Switch(60) = 0
  If keycode = LeftMagnaSave Then controller.Switch(61) = 0
End Sub


'***********************************************************************************
'****               Trough Handling                 ****
'***********************************************************************************

SolCallback(BallRelease)   = "TroughOut"      ' Ball Release
SolCallback(Drain)         = "TroughIn"       ' Drain

Sub TroughIn(enabled)
  if Enabled then
    sw46.Kick 58, 16 :
    Tgate.Open = True
    If sw46.BallCntOver > 0 Then PlaySoundAtVol SoundFX("Trough1",DOFcontactors), ActiveBall, .5: End If
  Else
    Tgate.Open = False
    BallSearch
  End If
end Sub

Sub TroughOut(enabled)
  if Enabled then
    sw47.Kick 58, 8 :PlaySoundAtVol SoundFX("BallRelease",DOFcontactors), ActiveBall, .4
  End If
end Sub

'***********************************************************************************
'****           Ball Ramp Trough Switches                 ****
'***********************************************************************************

Sub Sw46_hit():PlaySoundAtVol "fx_drain", Drain, 1:controller.Switch(46) = 1 : End Sub 'Drain
'Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
Sub TroughSFX_Hit(): PlaySoundAtVol "Trough2", ActiveBall, .2 : End Sub
Sub Sw49_hit():controller.Switch(49) = 1 : UpdateTrough : End Sub
Sub Sw48_hit():controller.Switch(48) = 1 : UpdateTrough : End Sub
Sub Sw47_hit():controller.Switch(47) = 1 : UpdateTrough : End Sub

Sub Sw46_UnHit():controller.Switch(46) = 0 : UpdateTrough : End Sub
Sub Sw49_UnHit():controller.Switch(49) = 0 : UpdateTrough : End Sub
Sub Sw48_UnHit():controller.Switch(48) = 0 : UpdateTrough : End Sub
Sub Sw47_UnHit():controller.Switch(47) = 0 : UpdateTrough : End Sub

Sub UpdateTrough: TroughTimer.enabled = 1 : TroughTimer.Interval = 200: end sub

Sub TroughTimer_Timer()
  If sw47.BallCntOver = 0 then sw48.kick 58, 12
  If sw48.BallCntOver = 0 then sw49.kick 58, 12
End Sub

Sub BallSearch()  'In case of hard pinmame reset. Called by PF solenoids firing empty.
  if Sw46.BallCntOver > 0 then controller.Switch(46) = 1 else controller.Switch(46) = 0
  if Sw49.BallCntOver > 0 then controller.Switch(49) = 1 else controller.Switch(49) = 0
  if Sw48.BallCntOver > 0 then controller.Switch(48) = 1 else controller.Switch(48) = 0
  if Sw47.BallCntOver > 0 then controller.Switch(47) = 1 else controller.Switch(47) = 0
End Sub

'***********************************************************************************
'****               Rollover Switches                 ****
'***********************************************************************************

'Primitives from Victory (Gottlieb 1987)

Sub sw9_Hit():PlaySoundAtVol "Sensor", ActiveBall, 1:Controller.Switch(9) = 1:sw9p.Visible = 0:End Sub
Sub sw9_UnHit():Controller.Switch(9) = 0:sw9p.Visible = 1:End Sub
Sub sw10_Hit():PlaySoundAtVol "Sensor2", ActiveBall, 1:Controller.Switch(10) = 1:sw10p.Visible = 0:End Sub
Sub sw10_UnHit():Controller.Switch(10) = 0:sw10p.Visible = 1:End Sub
Sub sw11_Hit():PlaySoundAtVol "Sensor3", ActiveBall, 1:Controller.Switch(11) = 1:sw11p.Visible = 0:End Sub
Sub sw11_UnHit():Controller.Switch(11) = 0:sw11p.Visible = 1:End Sub
Sub sw12_Hit():PlaySoundAtVol "Sensor4", ActiveBall, 1:Controller.Switch(12) = 1:sw12p.Visible = 0:End Sub
Sub sw12_UnHit():Controller.Switch(12) = 0:sw12p.Visible = 1:End Sub

Sub sw45_Hit():PlaySoundAtVol "Sensor5", ActiveBall, 1:Controller.Switch(45) = 1:sw45p.Visible = 0:End Sub
Sub sw45_UnHit():Controller.Switch(45) = 0:sw45p.Visible = 1:End Sub
Sub sw51_Hit():PlaySoundAtVol "Sensor6", ActiveBall, 1:Controller.Switch(51) = 1:sw51p.Visible = 0:End Sub
Sub sw51_UnHit():Controller.Switch(51) = 0:sw51p.Visible = 1:End Sub
Sub sw52_Hit():PlaySoundAtVol "Sensor7", ActiveBall, 1:Controller.Switch(52) = 1:sw52p.Visible = 0:End Sub
Sub sw52_UnHit():Controller.Switch(52) = 0:sw52p.Visible = 1:End Sub
Sub sw53_Hit():PlaySoundAtVol "Sensor8", ActiveBall, 1:Controller.Switch(53) = 1:sw53p.Visible = 0:End Sub
Sub sw53_UnHit():Controller.Switch(53) = 0:sw53p.Visible = 1:End Sub
Sub sw54_Hit():PlaySoundAtVol "Sensor9", ActiveBall, 1:Controller.Switch(54) = 1:sw54p.Visible = 0:End Sub
Sub sw64_Hit:Controller.Switch(64) = 1:PlaySound "guitarRiff",0,1,-0.15,0.25:End Sub
Sub sw64_Unhit:Controller.Switch(64) = 0:ActiveBall.VelY= ActiveBall.VelY/3:End Sub


'***********************************************************************************
'****               Stand Up Targets                  ****
'***********************************************************************************

Sub sw18_hit:vpmTimer.pulseSw 18:PlaySoundAtVol "target", ActiveBall,  .1:End Sub
Sub sw19_hit:vpmTimer.pulseSw 19:PlaySoundAtVol "target2", ActiveBall, .1:End Sub
Sub sw20_hit:vpmTimer.pulseSw 20:PlaySoundAtVol "target3", ActiveBall, .1:End Sub
Sub sw21_hit:vpmTimer.pulseSw 21:PlaySoundAtVol "target4", ActiveBall, .1:End Sub
Sub sw22_hit:vpmTimer.pulseSw 22:PlaySoundAtVol "target5", ActiveBall, .1:End Sub

Sub sw28_hit:vpmTimer.pulseSw 28:PlaySoundAtVol "target6", ActiveBall, .1:End Sub
Sub sw29_hit:vpmTimer.pulseSw 29:PlaySoundAtVol "target7", ActiveBall, .1:End Sub
Sub sw30_hit:vpmTimer.pulseSw 30:PlaySoundAtVol "target8", ActiveBall, .1:End Sub
Sub sw31_hit:vpmTimer.pulseSw 31:PlaySoundAtVol "target9", ActiveBall, .1:End Sub
Sub sw32_hit:vpmTimer.pulseSw 32:PlaySoundAtVol "target10", ActiveBall, .1:End Sub

Sub sw36_hit:vpmTimer.pulseSw 36:PlaySoundAtVol "target11", ActiveBall, .1:End Sub
Sub sw37_hit:vpmTimer.pulseSw 37:PlaySoundAtVol "target12", ActiveBall, .1:End Sub

Sub sw41_hit:vpmTimer.pulseSw 41:PlaySoundAtVol "target13", ActiveBall, .1:End Sub


'***********************************************************************************
'****               Drop Targets                    ****
'***********************************************************************************

'Left Lander Targets
Sub Sw13_Dropped:dtLBank.Hit 1 : End Sub
Sub Sw14_Dropped:dtLBank.Hit 2 : End Sub
Sub Sw15_Dropped:dtLBank.Hit 3 : End Sub
Sub Sw16_Dropped:dtLBank.Hit 4 : End Sub
Sub Sw17_Dropped:dtLBank.Hit 5 : End Sub

'Right Lander Targets
Sub sw23_Dropped:dtRBank.Hit 1 : End Sub
Sub Sw24_Dropped:dtRBank.Hit 2 : End Sub
Sub Sw25_Dropped:dtRBank.Hit 3 : End Sub
Sub Sw26_Dropped:dtRBank.Hit 4 : End Sub
Sub Sw27_Dropped:dtRBank.Hit 5 : End Sub

'Baiter Targets
Sub sw33_Dropped : dtBait.Hit 1  : End Sub
Sub sw34_Dropped : dtBait.Hit 2  : End Sub
Sub sw35_Dropped : dtBait.Hit 3  : End Sub

'Pod Targets
Sub sw39_Dropped:dtBPod.Hit 1 : End Sub
Sub sw40_Dropped:dtTPod.Hit 1 : End Sub

Sub Sw33a_Hit() : PlaySoundAtVol "soloff", ActiveBall, .2 : End Sub 'dont drop target but play sound
Sub Sw34a_Hit() : PlaySoundAtVol "soloff2", ActiveBall, .2 : End Sub 'dont drop target but play sound

'***********************************************************************************
'****               Switches                      ****
'***********************************************************************************

Sub sw38_hit:vpmTimer.pulseSw 38 : PlaySoundAtVol "soloff", ActiveBall, .1:End Sub

'***********************************************************************************
'****               Bumper                          ****
'***********************************************************************************

Const swJet1 = 55 ' Left Jet Bumper
Const swJet2 = 56 ' Right Jet Bumper

Sub Bumper1_Hit():vpmTimer.PulseSwitch (swJet1), 0, "" : PlaySoundAtVol SoundFX("LeftBumper_Hit",DOFContactors), ActiveBall, 1
  FlashForMs yBoxFlasher, 100, 2, 0  : End Sub
Sub Bumper2_Hit():vpmTimer.PulseSwitch (swJet2), 0, "" : PlaySoundAtVol SoundFX("RightBumper_Hit",DOFContactors), ActiveBall, 1
  FlashForMs yBoxFlasher, 100, 2, 0 : End Sub

'***********************************************************************************
'****               Drain hole & Kicker               ****
'***********************************************************************************

Sub sw42_Hit:bsLockup.AddBall me : PlaySoundAtVol "popper_ball", sw42, 1: End Sub

'***********************************************************************************
'****               Knocker                           ****
'***********************************************************************************

Sub KnockerSol(enabled)
  If enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

'***********************************************************************************
'****               KickBack                          ****
'***********************************************************************************

Sub SolKickBack(enabled)
    If enabled Then
       Kickback.Fire
       PlaySoundAtVol SoundFX("DiverterLeft",DOFContactors), Kickback, .5
    Else
       KickBack.PullBack
    End If
End Sub

'***********************************************************************************
'****               Shooter Lane                      ****
'***********************************************************************************

Sub sw50_Hit():PlaySoundAtVol "Sensor", ActiveBall, 1:Controller.Switch(50) = 1:sw50p.Visible = 0:End Sub
Sub sw50_UnHit():Controller.Switch(50) = 0:sw50p.Visible = 1:End Sub

dim BallInPlunger :BallInPlunger = False
sub PlungerLane_hit():ballinplunger = True: End Sub
Sub PlungerLane_unhit():BallInPlunger = False: End Sub

'===================
' NF Flipper Return Mod 2017
'===================
dim returnspeed, lfstep, rfstep, LFstep1
returnspeed = cInt(leftflipper.return*100)
LFstep = 1
LFstep1 = 1
RFstep = 1

sub LeftFlipper_timer()
  select case LFstep
    Case 1: LeftFlipper.return = ReturnCurve(LFstep) :LFstep = LFstep + 1
    Case 2: LeftFlipper.return = ReturnCurve(LFstep) :LFstep = LFstep + 1
    Case 3: LeftFlipper.return = ReturnCurve(LFstep) :LFstep = LFstep + 1
    Case 4: LeftFlipper.return = ReturnCurve(LFstep) :LFstep = LFstep + 1
    Case 5: LeftFlipper.return = ReturnCurve(LFstep) :LFstep = LFstep + 1
    Case 6: LeftFlipper.return = ReturnCurve(LFstep) :LFstep = LFstep + 1
    Case 7: LeftFlipper.timerenabled = 0 : LFstep = 1
  end select
' tb2.text = LeftFlipper.Return
end sub

sub LeftFlipper1_timer()
  select case LFstep1
    Case 1: LeftFlipper1.return = ReturnCurve(LFstep1) :LFstep1 = LFstep1 + 1
    Case 2: LeftFlipper1.return = ReturnCurve(LFstep1) :LFstep1 = LFstep1 + 1
    Case 3: LeftFlipper1.return = ReturnCurve(LFstep1) :LFstep1 = LFstep1 + 1
    Case 4: LeftFlipper1.return = ReturnCurve(LFstep1) :LFstep1 = LFstep1 + 1
    Case 5: LeftFlipper1.return = ReturnCurve(LFstep1) :LFstep1 = LFstep1 + 1
    Case 6: LeftFlipper1.return = ReturnCurve(LFstep1) :LFstep1 = LFstep1 + 1
    Case 7: LeftFlipper1.timerenabled = 0 : LFstep1 = 1
  end select
' tb2.text = LeftFlipper.Return
end sub

sub RightFlipper_timer()
  select case RFstep
    Case 1: RightFlipper.return = ReturnCurve(RFstep) :RFstep = RFstep + 1
    Case 2: RightFlipper.return = ReturnCurve(RFstep) :RFstep = RFstep + 1
    Case 3: RightFlipper.return = ReturnCurve(RFstep) :RFstep = RFstep + 1
    Case 4: RightFlipper.return = ReturnCurve(RFstep) :RFstep = RFstep + 1
    Case 5: RightFlipper.return = ReturnCurve(RFstep) :RFstep = RFstep + 1
    Case 6: RightFlipper.return = ReturnCurve(RFstep) :RFstep = RFstep + 1
    Case 7: RightFlipper.timerenabled = 0 : RFstep = 1
  end select
end sub

FlipperReturnLow = cInt(FlipperReturnLow) '% adjust flipper return min/max values

Function ReturnCurve(step)  'Adjust flipper curve
  dim x, s
  s = FlipperReturnHigh - FlipperReturnLow
  Select Case step
    case 0 'initial curve
      x = FlipperReturnLow
    case 1 '16ms
      x = FlipperReturnLow + ((s/6) * 1)
    case 2 '32ms
      x = FlipperReturnLow + ((s/6) * 2)
    case 3 'etc
      x = FlipperReturnLow + ((s/6) * 3)
    case 4
      x = FlipperReturnLow + ((s/6) * 4)
    case 5
      x = FlipperReturnLow + ((s/6) * 5)
    case 6
      if FlipperReturnOverride > 0 then
        ReturnCurve = FlipperReturnOverride
        Exit Function
      else
        x = FlipperReturnHigh
      End If
  End Select
  ReturnCurve = (returnspeed * x) / 10000
' tb.text = x & vbnewline & ReturnCurve
End Function

'====================
'   HARD
'       FLIPS
'====================
'just switches EOStorque when hit
dim defaultEOS, hardEOS
defaulteos = leftflipper.eostorque
hardeos = 2200 / leftflipper.strength 'eos equivalent to 2200 strength
hardeos = 2200 / leftflipper1.strength  'eos equivalent to 2200 strength
if hardflips = 0 then TriggerLF.enabled = 0 :TriggerLF1.enabled = 0 : TriggerRF.enabled = 0 end if

Sub TriggerLF_Hit()
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 0 then
    leftflipper.eostorque = hardeos
    End if
end sub

Sub TriggerLF1_Hit()
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 0 then
    LeftFlipper1.eostorque = hardeos
    End if
end sub

Sub TriggerRF_Hit()
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 0 then
    rightflipper.eostorque = hardeos
    End if
end sub

'***********************************************************************************
'****   nFozzy's FastFlips NF 'Pre-solid state flippers lag reduction   ****
'***********************************************************************************

dim FlippersEnabled

Sub SolRun(enabled)
  FlippersEnabled = Enabled
  if enabled then
    gameRun = True
'   PFGI.State = 1
    SLingL.Disabled = 0
    SLingR.Disabled = 0
  Elseif not Enabled Then
    gameRun = false
'   PFGI.State = 0
    if leftflipper.startangle > leftflipper.endangle Then
      if leftflipper.currentangle < leftflipper.startangle then leftflipper.rotatetostart : leftflippersound 0 : end if
    elseif leftflipper.startangle < leftflipper.endangle Then
      if leftflipper.currentangle > leftflipper.startangle then leftflipper.rotatetostart : leftflippersound 0 : end If
    end If
    if leftflipper1.startangle > leftflipper1.endangle Then
      if leftflipper1.currentangle < leftflipper1.startangle then leftflipper1.rotatetostart : end if
    elseif leftflipper1.startangle < leftflipper1.endangle Then
      if leftflipper1.currentangle > leftflipper1.startangle then leftflipper1.rotatetostart : end If
    end If
    if rightflipper.startangle > rightflipper.endangle Then
      if rightflipper.currentangle < rightflipper.startangle then rightflipper.rotatetostart : rightflippersound 0 : end if
    elseif rightflipper.startangle < rightflipper.endangle Then
      if rightflipper.currentangle > rightflipper.startangle then rightflipper.rotatetostart : rightflippersound 0 : end If
    end If
    SLingL.Disabled = 1
    SLingR.Disabled = 1
  end if
  vpmNudge.SolGameOn enabled
End Sub

sub flipnf(LR, DU)  'object, left-right, downup
  if FastFlips = 0 then
    if LR = 0 then
      controller.switch(LLFlip) = DU
    elseif LR = 1 then
      controller.switch(LRFlip) = DU
    End If
    exit Sub
  end if
  if LR = 0 Then    'left flipper
    if DU = 1 then
      If FlippersEnabled = True then
        leftflipper.rotatetoend
        leftflipper1.rotatetoend
        LeftFlipperSound 1
      end if
      controller.Switch(swLLFlip) = True
    Elseif DU = 0 then
      If FlippersEnabled = True then
        leftflipper.rotatetoStart
        leftflipper1.rotatetoStart
        LeftFlipperSound 0
      end if
      controller.Switch(swLLFlip) = False
    end if
  elseif LR = 1 then    ''right flipper
    if DU = 1 then
      If FlippersEnabled = True then
        RightFlipper.rotatetoend
        RightFlipperSound 1
      end if
      controller.Switch(swLRFlip) = True
    Elseif DU = 0 then
      If FlippersEnabled = True then
        RightFlipper.rotatetoStart
        RightFlipperSound 0
      end if
      controller.Switch(swLRFlip) = False
    end if
  end if
end sub

sub LeftFlipperSound(updown)'called along with the flipper, so feel free to add stuff, EOStorque tweaks, animation updates, upper flippers, whatever.
  if updown = 1 Then
    PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), LeftFlipper, 1
  Else
    PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), LeftFlipper, 1
    LeftFlipper.eostorque = defaultEOS
    LeftFlipper1.eostorque = defaultEOS
      if FlipperReturnMod = 1 then
        LeftFlipper.TimerEnabled = 1
        LeftFlipper1.TimerEnabled = 1
        LeftFlipper.TimerInterval = 16  '400 test
        LeftFlipper1.TimerInterval = 16 '400 test
        LeftFlipper.return = FlipperReturnLow / 10000
        LeftFlipper1.return = FlipperReturnLow / 10000
        LFstep = 1
        LFstep1 = 1
      end if
  end if
end sub
sub RightFlipperSound(updown)'called along with the flipper, so feel free to add stuff, EOStorque tweaks, animation updates, upper flippers, whatever.
  if updown = 1 Then
    PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), RightFlipper, 1
  Else
    PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), RightFlipper, 1
    RightFlipper.eostorque = defaultEOS
    LeftFlipper1.eostorque = defaultEOS
      if FlipperReturnMod = 1 then
        RightFlipper.TimerEnabled = 1
        LeftFlipper1.TimerEnabled = 1
        RightFlipper.TimerInterval = 16 '400 test
        LeftFlipper1.TimerInterval = 16 '400 test
        RightFlipper.return = FlipperReturnLow / 10000
        LeftFlipper1.return = FlipperReturnLow / 10000
        RFstep = 1
        LFstep1 = 1
      end if
  end if
end sub

'***********************************************************************************
'****       Special JP Flippers 'Legacy Flippers using callbacks    ****
'***********************************************************************************

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  if FastFlips = 0 then
    If Enabled Then
      PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), LeftFlipper, 1
      LeftFlipper.RotateToEnd
      LeftFlipper1.RotateToEnd
    Else
      PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), LeftFlipper, 1
      LeftFlipper.RotateToStart
      LeftFlipper1.RotateToStart
      LeftFlipper.eostorque = defaulteos
      LeftFlipper1.eostorque = defaulteos
      if FlipperReturnMod = 1 then
        LeftFlipper.TimerEnabled = 1
        LeftFlipper1.TimerEnabled = 1
        LeftFlipper.TimerInterval = 16
        LeftFlipper1.TimerInterval = 16
        LeftFlipper.return = FlipperReturnLow / 10000
        LeftFlipper1.return = FlipperReturnLow / 10000
        LFstep = 1
        LFstep1 = 1
      end if
    End If
  End If
End Sub

Sub SolRFlipper(Enabled)
  if FastFlips = 0 then
    If Enabled Then
      PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), RightFlipper, 1
      RightFlipper.RotateToEnd
    Else
      PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), RightFlipper, 1
      RightFlipper.RotateToStart
      RightFlipper.eostorque = defaulteos
      if FlipperReturnMod = 1 then
        rightflipper.TimerEnabled = 1
        rightflipper.TimerInterval = 16
        rightflipper.return = FlipperReturnLow / 10000
        RFstep = 1
      end if
    End If
  End If
End Sub

'***********************************************************************************
'****             Slingshot                           ****
'***********************************************************************************
Dim SLPos,SRPos

Sub SlingL_Slingshot()
' If FlipperEnabled = False Then Exit Sub
  LSling1.Visible = 0:LSling2.Visible = 0:LSling3.Visible = 0: LSling4.Visible = 1: LSling.TransZ = -27
  vpmTimer.PulseSw 57
  PlaySoundAtVol SoundFX ("left_slingshot",DOFContactors), ActiveBall,1 : DOF dLeftSlingshot, 2
  SLPos = 0: Me.TimerEnabled = 1
' LightshowChangeSide
End Sub

Sub SlingL_Timer
    Select Case SLPos
        Case 2: LSling1.Visible = 0:LSling2.Visible = 0:LSling3.Visible = 1: LSling4.Visible = 0: LSling.TransZ = -17
    Case 3: LSling1.Visible = 0:LSling2.Visible = 1:LSling3.Visible = 0: LSling4.Visible = 0: LSling.TransZ = -8
    Case 4: LSling1.Visible = 1:LSling2.Visible = 0:LSling3.Visible = 0: LSling4.Visible = 0: LSling.TransZ = 0 :Me.TimerEnabled = 0
    End Select
    SLPos = SLPos + 1
End Sub

Sub SlingR_Slingshot()
' If FlipperEnabled = False Then Exit Sub
  RSling1.Visible = 0:RSling2.Visible = 0:RSling3.Visible = 0: RSling4.Visible = 1: RSling.TransZ = -27
  vpmTimer.PulseSw 58
  PlaySoundAtVol SoundFX ("right_slingshot",DOFContactors), Activeball, 1 : DOF dRightSlingshot, 2
  SRPos = 0: Me.TimerEnabled = 1
' LightshowChangeSide
End Sub

Sub SlingR_Timer
    Select Case SRPos
        Case 2: RSling1.Visible = 0:RSling2.Visible = 0:RSling3.Visible = 1: RSling4.Visible = 0: RSling.TransZ = -17
    Case 3: RSling1.Visible = 0:RSling2.Visible = 1:RSling3.Visible = 0: RSling4.Visible = 0: RSling.TransZ = -8
    Case 4: RSling1.Visible = 1:RSling2.Visible = 0:RSling3.Visible = 0: RSling4.Visible = 0: RSling.TransZ = 0 :Me.TimerEnabled = 0
    End Select
    SRPos = SRPos + 1
End Sub

Sub Slingshots_Init
  LSling1.Visible = 1:LSling2.Visible = 0:LSling3.Visible = 0: LSling4.Visible = 0: LSling.TransZ = 0
  RSling1.Visible = 1:RSling2.Visible = 0:RSling3.Visible = 0: RSling4.Visible = 0: RSling.TransZ = 0
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

'Dim bulb
Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
  UpdateLamps
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps
'    On Error Resume Next
  nFadeLm 9, Light9
  nFadeLm 10, Light10
  nFadeLm 11, Light11
  nFadeLm 12, Light12
  nFadeLm 13, Light13
  nFadeLm 14, Light14
  nFadeLm 15, Light15
  nFadeLm 16, Light16
  nFadeLm 17, Light17
  nFadeLm 18, Light18
  nFadeLm 19, Light19
  nFadeLm 20, Light20
  nFadeLm 21, Light21
  nFadeLm 22, Light22
  nFadeLm 23, Light23
  nFadeLm 24, Light24
  nFadeLm 25, Light25
  nFadeLm 26, Light26
  nFadeLm 27, Light27
  nFadeLm 28, Light28
  nFadeLm 29, Light29
  nFadeLm 30, Light30
  nFadeLm 31, Light31
  nFadeLm 32, Light32
  nFadeLm 33, Light33
  nFadeLm 34, Light34
  nFadeLm 35, Light35
  NFadeLm 36, Light36
  NFadeLm 37, Light37
  NFadeLm 38, Light38
  NFadeLm 39, Light39
  FlashC 39, Flash39b
  NFadeLm 40, Light40
  FlashC 40, Flash40b
  NFadeLm 41, Light41
  FlashC 41, Flash41b
  NFadeLm 42, Light42
  FlashC 42, Flash42b
  nFadeLm 43, Light43
  nFadeLm 44, Light44
  nFadeLm 45, Light45
  NFadeLm 46, Light46
  NFadeLm 47, Light47
  NFadeLm 48, Light48
  NFadeLm 49, Light49
  NFadeLm 50, Light50
  NFadeLm 51, Light51
  NFadeLm 52, Light52
  NFadeLm 53, Light53
  NFadeLm 54, Light54
  nFadeLm 55, Light55
  nFadeLm 56, Light56
  nFadeLm 57, Light57
  nFadeLm 58, Light58
  nFadeLm 59, Light59
  nFadeLm 60, Light60
  nFadeLm 61, Light61
  nFadeLm 62, Light62


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

Sub Flashc(nr, object)
    Select Case FadingLevel(nr)
    Case 3
            FadingLevel(nr) = 0 'completely off
            Object.IntensityScale = FlashLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * CGT)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
               FadingLevel(nr) = 3 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * CGT)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 6 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 6 ' on
      FadingLevel(nr) = 1 'completely on
      Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

dim DisplayColor, DisplayColorG
InitDisplay
Sub InitDisplay()
  dim x
  for each x in Display
    x.Color = RGB(1,1,1)
  next
  Displaytimer.enabled = 1
  DisplayColor = dx.Color
  DisplayColorG = dxG.Color
  If DesktopMode then
    for each x in Display
      x.height = 380
      x.x = x.x - 956
      x.y = x.y -37
      x.rotx = -42
      x.visible = 1
    next
    for each x in Display2
      x.x = x.x +109
      x.y = x.y + 1
    next
    for each x in Display3
      x.height = 325
      x.x = x.x +32
      x.y = x.y +30
    next
    for each x in Display4
      x.height = 325
      x.x = x.x +125
      x.y = x.y +30
    next
  Else
    Select Case SingleScreenFS
      Case 1
        Displaytimer.enabled = 1
        for each x in Display
          x.visible = 1
        next
      case 0, 2
        Displaytimer.enabled = 0
        for each x in Display
          x.visible = 0
        next
    End Select
  End If

  if VPXdisplay = 0 Then
    Displaytimer.enabled = 0
    for each x in Display
      x.visible = 0
    next
  End If
End Sub

Sub Displaytimer_Timer
  On Error Resume Next
  Dim ii, jj, obj, b, x
  Dim ChgLED,num, chg, stat
  ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        For Each obj In Digits(num)
          If chg And 1 Then FadeDisplay obj, stat And 1
          chg=chg\2 : stat=stat\2
        Next
      Next
    End If
End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
  Else
    Object.Color = RGB(1,1,1)
  End If
End Sub

Dim Digits(32)
Digits(14)=Array(D211,D212,D213,D214,D215,D216,D217,D218)
Digits(15)=Array(D219,D220,D221,D222,D223,D224,D225,D226)
Digits(16)=Array(D227,D228,D229,D230,D231,D232,D233,D234)
Digits(17)=Array(D235,D236,D237,D238,D239,D240,D241,D242)
Digits(18)=Array(D243,D244,D245,D246,D247,D248,D249,D250)
Digits(19)=Array(D251,D252,D253,D254,D255,D256,D257,D258)
Digits(20)=Array(D259,D260,D261,D262,D263,D264,D265,D266)
Digits(21)=Array(D267,D268,D269,D270,D271,D272,D273,D274)
Digits(22)=Array(D275,D276,D277,D278,D279,D280,D281,D282)
Digits(23)=Array(D283,D284,D285,D286,D287,D288,D289,D290)
Digits(24)=Array(D291,D292,D293,D294,D295,D296,D297,D298)
Digits(25)=Array(D299,D300,D301,D302,D303,D304,D305,D306)
Digits(26)=Array(D307,D308,D309,D310,D311,D312,D313,D314)
Digits(27)=Array(D315,D316,D317,D318,D319,D320,D321,D322)
Digits(28)=Array(D323,D324,D325,D326,D327,D328,D329,D330)
Digits(29)=Array(D331,D332,D333,D334,D335,D336,D337,D338)
Digits(30)=Array(D339,D340,D341,D342,D343,D344,D345,D346)
Digits(31)=Array(D347,D348,D349,D350,D351,D352,D353,D354)


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







' Walls

Sub FadeWS(nr, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 1:d.IsDropped = 0:FadingLevel(nr) = 0 'Off
        Case 3:a.IsDropped = 1:b.IsDropped = 0:c.IsDropped = 1:d.IsDropped = 1:FadingLevel(nr) = 2 'fading...
        Case 4:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 0:d.IsDropped = 1:FadingLevel(nr) = 3 'fading...
        Case 5:a.IsDropped = 0:b.IsDropped = 1:c.IsDropped = 1:d.IsDropped = 1:FadingLevel(nr) = 1 'ON
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
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
        Case 4:object.image = b:FadingLevel(nr) = 0
        Case 5:object.image = a:FadingLevel(nr) = 1
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

Sub Flashc(nr, object)
    Select Case FadingLevel(nr)
    Case 3
            FadingLevel(nr) = 0 'completely off
            Object.IntensityScale = FlashLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr))
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
               FadingLevel(nr) = 3 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr))
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 6 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 6 ' on
      FadingLevel(nr) = 1 'completely on
      Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

' RGB Leds

Sub RGBLED (object,red,green,blue)
object.color = RGB(0,0,0)
object.colorfull = RGB(2.5*red,2.5*green,2.5*blue)
object.state=1
End Sub

' Modulated Flasher and Lights objects

Sub SetLampMod(nr, value)
    If value > 0 Then
    LampState(nr) = 1
  Else
    LampState(nr) = 0
  End If
  FadingLevel(nr) = value
End Sub

Sub FlashMod(nr, object)
  Object.IntensityScale = FadingLevel(nr)/255
End Sub

Sub LampMod(nr, object)
Object.IntensityScale = FadingLevel(nr)/255
Object.State = LampState(nr)
End Sub




Sub SolGi(enabled)
  If enabled Then
    LightCenter1.State = 1 'Red Center Light On
    LightCenter2.State = 1 'Red Center Light On
    LightCenter3.State = 1 'Red Center Light On
    LightCenter4.State = 1 'Red Center Light On
    LightCenter5.State = 1 'Red Center Light On
     GiOFF
   GiFlasher1.Visible = 0
   GiFlasher2.Visible = 0
   GiFlasher3.Visible = 0
   GiFlasher4.Visible = 0

  Playsound "fx_relay_off"
  Table1.ColorGradeImage = "ColorGrade_1"
   Else
    LightCenter1.State = 0 'Red Center Light Off
    LightCenter2.State = 0 'Red Center Light Off
    LightCenter3.State = 0 'Red Center Light Off
    LightCenter4.State = 0 'Red Center Light Off
    LightCenter5.State = 0 'Red Center Light Off
     GiON
   GiFlasher1.Visible = 1
   GiFlasher2.Visible = 1
   GiFlasher3.Visible = 1
   GiFlasher4.Visible = 1

  Playsound "fx_relay_on"
  Table1.ColorGradeImage = "ColorGrade_8"
    GITimer.enabled = 1
 End If
End Sub

Sub GITimer_Timer
    LightCenter1.State = 0 'Red Center Light Off
    LightCenter2.State = 0 'Red Center Light Off
    LightCenter3.State = 0 'Red Center Light Off
    LightCenter4.State = 0 'Red Center Light Off
    LightCenter5.State = 0 'Red Center Light Off
' PFGI.State = 1
' plasticsOn.WidthTop=1024
' plasticsOn.WidthBottom=1024
  GITimer.enabled = 0
End Sub

Sub GiON
  dim xx
    For each xx in GiLights
        xx.State = 1
    Next
End Sub

Sub GiOFF
  dim xx
    For each xx in GiLights
        xx.State = 0
    Next
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

Function LVL(input)
  LVL = Input * SoundLevelMult
End Function

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


Sub zApron_Hit (idx)
  PlaySound "woodhitaluminium", 0, LVL((Vol(ActiveBall)^2.5)*10), Pan(ActiveBall)/4, 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub zInlanes_Hit (idx)
  PlaySound "MetalHit2", 0, LVL(Vol(ActiveBall)*5), Pan(ActiveBall)*55, 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub RampEntry_Hit()
  If activeball.vely < -10 then
    PlaySound "ramp_hit2", 0, LVL(Vol(ActiveBall)/5), Pan(ActiveBall), 0, Pitch(ActiveBall)*10, 1, 0, AudioFade(ActiveBall)
  Elseif activeball.vely > 3 then
    PlaySound "PlayfieldHit", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End If
End Sub

Sub RampHit2_Hit()
  PlaySound "ramp_hit3", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub RampEntry2_Hit()
  Playsound "WireRamp", 0, LVL(0.3)
End Sub

Sub RampSound0_Hit()
  If Activeball.velx > 0 then
    PlaySoundAtVol "pmaReriW", ActiveBall, 1
  Else
    PlaySoundAtVol "WireRamp1", ActiveBall, 1
  End If
End Sub

Sub Metals_Hit (idx)
  PlaySound "metalhit_medium", 0, LVL(Vol(ActiveBall)*30 ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound SoundFX("target",DOFTargets), 0, LVL(0.2), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  PlaySound SoundFX("targethit",0), 0, LVL(Vol(ActiveBall)*18 ), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  FlashForMs yBoxFlasher, 1000, 40, 0
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 15 then
    PlaySound "fx_rubber2", 0, LVL(Vol(ActiveBall)*30 ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if

  If finalspeed >= 6 AND finalspeed <= 15 then
    RandomSoundRubber()
  else
    PlaySound "rubber_hit_3", 0, LVL(Vol(ActiveBall)*30 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End If
End Sub

Sub RubberPosts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, LVL(Vol(ActiveBall)*80 ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 15 then
    RandomSoundRubber()
  else
    PlaySound "rubber_hit_3", 0, LVL(Vol(ActiveBall)*1 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, LVL(Vol(ActiveBall)*50 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, LVL(Vol(ActiveBall)*50 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, LVL(Vol(ActiveBall)*50 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

'Ramp drops using collision events
Sub col_UPF_Fall_hit():if FallSFX1.Enabled = 0 then playsound "drop_mono", 0, LVL(0.5), 0 :FallSFX1.Enabled = 1 :end if :end sub'name,loopcount,volume,pan,randompitch
Sub FallSFX1_Timer():me.Enabled = 0:end sub

Sub Col_Rramp_Fall_Hit():if FallSFX2.Enabled = 0 then playsound "drop_mono2", 0, LVL(0.5), 0.05 :FallSFX2.Enabled = 1 :end if :end sub'name,loopcount,volume,pan,randompitch
Sub FallSFX2_Timer():me.Enabled = 0:end sub

Sub col_Lramp_Fall_Hit():if FallSFX3.Enabled = 0 then playsound "drop_mono3", 0, LVL(0.5), -0.05 :FallSFX3.Enabled = 1  :end if :end sub'name,loopcount,volume,pan,randompitch
Sub FallSFX3_Timer():me.Enabled = 0:end sub

'***********************************************************************************
'****                 DOF reference                     ****
'***********************************************************************************
const dShooterLane          = 201 ' Shooterlane
const dLeftSlingshot        = 211 ' Left Slingshot
const dRightSlingshot         = 212 ' Right Slingshot

Sub Table1_Exit()
  Controller.Games(cGameName).Settings.Value("sound") = 1
  Controller.Stop
End Sub

