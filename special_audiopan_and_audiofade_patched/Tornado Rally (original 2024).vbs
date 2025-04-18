Option Explicit

'*****************************************************************************************************
'*****************************************************************************************************

'        &&&&&&       &&&&&&&     &&&&    &&&&&&&&&&&&&  &&&&&&&&&&&      &&&&&&&&
'     &&&&&&&&&&&&    &&&&&&&   &&&&&&    &&&&&&&&&&&&   &&&&&&&&&&&    &&&&&&&&&&&
'    &&&&&&&&&&&&&&&    &&&&&  &&&&&&         &&&&       &&&&    &&&    &&&&&   &&&
'   &&&&&&     &&&&&&   &&&&&&&&&&&           &&&&       &&&&&&&&&       &&&&&&&
'  &&&&&&      &&&&&&   &&&&&&&&&             &&&&       &&&&&&&&&&        &&&&&&&
'  &&&&&&&     &&&&&&   &&&&&&&&&&&           &&&&       &&&&           &&&&  &&&&&&
'   &&&&&&&&&&&&&&&&      &&&& &&&&&          &&&&       & &&&   &&&     &&&&  &&&&&
'    &&&&&&&&&&&&&&  &&&&& &&   &&&&&&  &&  &&&&&&&&&&&  &&& &&&&&&&&&&&&  &&&&&&&&& &&&&
'      &&&&&&&&&&&   &&&&& &      &&&&  &&  &&&&&&&&&&&  &&& &&&&&&& &&&& &&&&&&&&& &&&&&

'*****************************************************************************************************
'*****************************************************************************************************


'*****************************************************************************************************
' CREDITS
' Table designed and created by Flying Rabbit Studios
' Backglass lighting by Hauntedfreaks
' DOF by MauiPunter
' Big THANK YOU to Apophis and the VPW Community
'
' Hay bales scan for VR Room by Abby Crawford via Sketchfab CC Attribution
'
' Initial table created by fuzzel, jimmyfingers, jpsalas, toxie & unclewilly (in alphabetical order)
' Flipper primitives by zany
' Ball rolling sound script by jpsalas
' Ball shadow by ninuzzu
' Ball control & ball dropping sound by rothbauerw
' Positional sound helper functions by djrobx
' Plus a lot of input from the whole community (sorry if we forgot you :/)
'*****************************************************************************************************


'*****************************************************************************************************
'*****OPTIONS****************************************************************************************
'*****************************************************************************************************


Const BackglassOn = True

Dim SideRailsVisible
SideRailsVisible = (Table1.Option("Show Side Rails", 0, 1, 1, 0, 0, Array("True", "False")) = 0)
Sub Table1_OptionEvent(ByVal eventId)
    SideRailsVisible = (Table1.Option("Show Side Rails", 0, 1, 1, 0, 0, Array("True", "False")) = 0)
  VR_Cab_Siderails.visible = SideRailsVisible
End Sub

Dim MixedRealityPassthru
MixedRealityPassthru = (Table1.Option("MixedReality", 0, 1, 1, 1, 0, Array("False", "True")) = 1)
Sub Table1_OptionEvent(ByVal eventId)
    MixedRealityPassthru = (Table1.Option("MixedReality", 0, 1, 1, 1, 0, Array("False", "True")) = 1)
  If VRRoom = 1 Then
    If MixedRealityPassthru = True Then
      MR_Cube.visible = True
      For each Reels in DesktopStuff : Reels.Visible = False : Next
      For each VRcab in VRstuff : VRcab.Visible = True : Next
      For each VRrm in VRenv : VRrm.visible = False : Next
    Else
      For each VRrm in VRenv : VRrm.visible = True : Next
      For each VRcab in VRstuff : VRcab.Visible = True : Next
    End If
  End If

End Sub


'*****************************************************************************************************
'*****************************************************************************************************


Const BallSize = 50        'Ball diameter in VPX units; must be 50
Const BallMass = 1          'Ball mass must be 1

'First, try to load the Controller.vbs (DOF), which helps controlling additional hardware like lights, gears, knockers, bells and chimes (to increase realism)
'This table uses DOF via the 'SoundFX' calls that are inserted in some of the PlaySound commands, which will then fire an additional event, instead of just playing a sample/sound effect
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

'If using Visual PinMAME (VPM), place the ROM/game name in the constant below,
'both for VPM, and DOF to load the right DOF config from the Configtool, whether it's a VPM or an Original table

Const cGameName = "OKIES_TornadoRally"
LoadEM


Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed
Dim B2SOn   'True/False if want backglass
If BackglassOn = False Then B2SOn = False

Dim BIP

Dim Credit

Dim CurrScore

Dim TgtScore

Dim HiScore


loadhs

Dim PumpJackPlay

Dim AdvBonus

Dim Bonus2x

Dim BumperBonus

Dim RedSpecial

Dim WhaleBallLocked

Dim WhaleBIP: WhaleBIP = False

Const HSFileName="OKIES.txt"


Dim Tilt: Tilt = 0
Const TiltSensitivity = 6
Dim Tilted: Tilted = False



'"viewmodes"
Dim VRRoom, cab_mode, DesktopMode: DesktopMode = Table1.ShowDT
DesktopMode = Table1.ShowDT
If RenderingMode = 2 Then VRRoom=1 Else VRRoom=0      'VRRoom set based on RenderingMode starting in version 10.72
if Not DesktopMode and VRRoom=0 Then cab_mode=1 Else cab_mode=0

dim VRcab
dim VRrm
dim Reels

If VRRoom = 0 and cab_mode = 1 Then
  For each VRcab in VRstuff : VRcab.Visible = False : Next
  For each VRrm in VRenv : VRrm.visible = False : Next
  For each Reels in DesktopStuff : Reels.Visible = False : Next
  If SideRailsVisible = True Then
    VR_Cab_Siderails.visible = True
  End If
Else
  For each VRcab in VRstuff : VRcab.Visible = False : Next
  For each VRrm in VRenv : VRrm.visible = False : Next
  If SideRailsVisible = True Then
    VR_Cab_Siderails.visible = True
  End If
End If
If VRRoom = 1 Then
  If MixedRealityPassthru = True Then
    MR_Cube.visible = True
    For each Reels in DesktopStuff : Reels.Visible = False : Next
    For each VRcab in VRstuff : VRcab.Visible = True : Next
    For each VRrm in VRenv : VRrm.visible = False : Next
  Else
    For each VRrm in VRenv : VRrm.visible = True : Next
    For each VRcab in VRstuff : VRcab.Visible = True : Next
  End If
End If





'"Volume Options"

Const VolumeDial = 0.8

Const musicvol = 0.9

Const chimevol = 0.7

Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1

Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height



Dim EnableRetractPlunger
EnableRetractPlunger = false 'Change to true to enable retracting the plunger at a linear speed; wait for 1 second at the maximum position; move back towards the resting position; nice for button/key plungers






Dim xx, bsTrough, initContacts, dtRDrop, dtLDrop, bsLSaucer, bsMSaucer, bsRSaucer


Sub OKIES_Paused:Controller.Pause = 1:End Sub

Sub OKIESSpiderMan_unPaused:Controller.Pause = 0:End Sub

Sub OKIES_Exit
    If b2son then controller.stop
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
  If VR_Primary_plunger.Y < -135 then
      VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
 'debug.print plunger.position
  VR_Primary_plunger.Y = 0 + (5* Plunger.Position) -20
End Sub

sub startGame_timer
    dim xx
    playsound "poweron"
    LampTimer.enabled=1
    For each xx in GI:xx.State = 1: Next        '*****GI Lights On
    me.enabled=false
    VR_Backbox_Backglass.blenddisablelighting = 3
end sub







Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

'Const BallSize = 25  'Ball radius


Sub Table1_Init()
  loadhs
  If HiScore = 0 Then
    HiScore = 500
  End If

  ResetBonuses
  ResetScore
  TargetRESET
  TargetXtraRESET
  BIP = 0
  EMReel1.SetValue Credit
  AttractMode
  If Credit>0 then DOF 137, DOFOn
  DOF 109, DOFPulse
End Sub


Sub LSattract_PlayDone()
  AttractMode
End Sub

Sub GameStart()
  Credit = Credit - 1
  if Credit < 1 then DOF 137, DOFOff
  ResetBonuses
  ResetScore
  TargetRESET
  TargetXtraBall.isDropped = True
  BIP = 3
  DisplayBIP
  BallRelease.CreateBall
  BallRelease.Kick 90, 7
  RandomSoundBallRelease BallRelease
End Sub




Dim BIPL: BIPL = 0

Sub Table1_KeyDown(ByVal keycode)
  If Keycode = StartGameKey Then VR_Cab_StartButton.y = VR_Cab_StartButton.Y - 4

  If Keycode = StartGameKey AND BIP = 0 AND Credit > 0 Then
      If B2SOn = True Then Controller.B2SSetData 10,1
      If B2SOn = True Then Controller.B2SSetData 11,1
      Credit = Credit - 1
        If Credit < 1 Then
          Credit = 0
          If B2SOn = True Then Controller.B2SSetData 9,0
        End If
      EMReel1.SetValue Credit
      soundStartButton()
      GameStart
      LSattract.StopPlay
  End If

  If keycode = AddCreditKey or keycode = 4 Then
    If B2SOn = True Then Controller.B2SSetData 9,1
    Credit = Credit + 1
      If Credit = 0 Then
        Credit = 1
      End If
    savehs
    EMReel1.SetValue Credit
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25

    End Select
    'OKIESwinLS
    If Credit>0 then DOF 137, DOFOn
  End If

  If keycode = PlungerKey Then
        If EnableRetractPlunger Then
            Plunger.PullBackandRetract
        Else
        Plunger.PullBack
        End If
    Plunger.Pullback:SoundPlungerPull()
  End If



  If keycode = LeftFlipperKey And Tilted = False And BIP > 0 Then
'   LeftFlipper.RotateToEnd
    SolLFlipper True  'This would be called by the solenoid callbacks if using a ROM
    PlaySound SoundFXDOF("fx_flipperup",101,DOFOn,DOFFlippers), 0, .67, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
    VR_CabFlipperLeft.X = VR_CabFlipperLeft.X + 8
  End If

  If keycode = RightFlipperKey And Tilted = False  And BIP > 0 Then
'   RightFlipper.RotateToEnd
    SolRFlipper True  'This would be called by the solenoid callbacks if using a ROM
    PlaySound SoundFXDOF("fx_flipperup",102,DOFOn,DOFFlippers), 0, .67, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
    VR_CabFlipperRight.X = VR_CabFlipperRight.X - 8
  End If

  If keycode = LeftTiltKey Then Nudge 90, 5:SoundNudgeLeft(): CheckTilt

  If keycode = RightTiltKey Then Nudge 270, 5:SoundNudgeRight(): CheckTilt

  If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter(): CheckTilt

  If keycode = 20 Then CheckTilt

End Sub

Sub Table1_KeyUp(ByVal keycode)
  If Keycode = StartGameKey Then VR_Cab_StartButton.y = VR_Cab_StartButton.y + 4

  If keycode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
  End If

  If keycode = LeftFlipperKey And Tilted = False Then
'   LeftFlipper.RotateToStart
    SolLFlipper False   'This would be called by the solenoid callbacks if using a ROM
    PlaySound SoundFXDOF("fx_flipperdown",101,DOFOff,DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
    VR_CabFlipperLeft.X = 0
  End If

  If keycode = RightFlipperKey And Tilted = False Then
'   RightFlipper.RotateToStart
    SolRFlipper False   'This would be called by the solenoid callbacks if using a ROM
    PlaySound SoundFXDOF("fx_flipperdown",102,DOFOff,DOFFlippers), 0, 1, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
    VR_CabFlipperRight.X = 0
  End If

    'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If
End Sub


Sub Drain_Hit()
  If Tilted = False Then EOBBonusCollect
  RandomSoundDrain drain
  PlaySound "DrainSynth", 0, musicvol + 0.5, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
  Drain.DestroyBall
  If TimerShootAgain.enabled = False Then
    BIP = BIP - 1
  End If

  If WhaleBIP = True Then
    WhaleBIP = False
      If BIP > 0 Then
        DisplayBIP
        Exit Sub
      End If
    Else
      If Tilted = True Then
        TiltEnd
      End if
  End If

  If BIP = 0 Then
    If WhaleBallLocked = True then
      LSwhale.UpdateInterval = 22
      LSwhale.Play SeqBlinking, ,3,85
      whaletimer.Enabled = True
      BIP = 0
      Exit Sub
    Else
      Select Case Int(Rnd*2)+1
        Case 1 : PlaySound "TakeMeBack2Tulsa", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
        Case 2 : Playsound "ThisLand", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
      End Select
      TimerOKIESreset.enabled=True
      AttractMode
      savehs
    End If

  Else
    BallRelease.CreateBall
    BallRelease.Kick 90, 7
    RandomSoundBallRelease BallRelease

  End If
  DisplayBIP

  ResetBonuses
End Sub



Sub DisplayBIP
Dim gol
  Select Case BIP
    Case 0 : '*** GAME OVER
        If cab_mode = 0 Then
          For each gol in GAMEOVER_light : gol.State = 1 : Next
        End if
        If B2SOn Then
        Controller.B2SSetData 21,0
        Controller.B2SSetData 22,0
        Controller.B2SSetData 23,0
        Controller.B2SSetData 24,0
        Controller.B2SSetData 25,1
        Controller.B2SSetData 10,0
        Controller.B2SSetData 11,0

        End If
    Case 1 : For each gol in GAMEOVER_light : gol.State = 0 : Next
        If B2SOn Then
        Controller.B2SSetData 21,1
        Controller.B2SSetData 22,0
        Controller.B2SSetData 23,0
        Controller.B2SSetData 24,0
        Controller.B2SSetData 25,0

        End If
    Case 2  : For each gol in GAMEOVER_light : gol.State = 0 : Next
        If B2SOn Then
        Controller.B2SSetData 21,1
        Controller.B2SSetData 22,1
        Controller.B2SSetData 23,0
        Controller.B2SSetData 24,0
        Controller.B2SSetData 25,0

        End If
    Case 3 : For each gol in GAMEOVER_light : gol.State = 0 : Next
        If B2SOn Then
        Controller.B2SSetData 21,1
        Controller.B2SSetData 22,1
        Controller.B2SSetData 23,1
        Controller.B2SSetData 24,0
        Controller.B2SSetData 25,0

        End If
    Case 4 :For each gol in GAMEOVER_light : gol.State = 0 : Next
        If B2SOn Then
        Controller.B2SSetData 21,1
        Controller.B2SSetData 22,1
        Controller.B2SSetData 23,1
        Controller.B2SSetData 24,1
        Controller.B2SSetData 25,0

        End If
  End Select
  EMReel_Ball.SetValue BIP
End Sub



'*********
' TILT
'*********
Dim tlt

Sub CheckTilt
  Tilt = Tilt + TiltSensitivity
  TiltDecreaseTimer.Enabled = True

  If Tilt > 15 Then 'Tilt the Table1
    StopAllSounds
    PlaySound "buzz", 0, musicvol + 0.5, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
    Tilted = True
    If B2SOn Then Controller.B2SSetData 33,1
    If cab_mode = 0 Then
      For each tlt in TILT_light : tlt.State = 1 : Next
    End if
    DisableTable True
  End If

end Sub

Sub TiltDecreaseTimer_Timer
  ' DecreaseTilt
  If Tilt> 0 Then
    Tilt = Tilt - 0.1
  Else
    TiltDecreaseTimer.Enabled = False
  End If
End Sub

Sub DisableTable(Enabled)
  If Enabled Then
    lsOFF.play SeqAllOff
    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart
    LeftSlingShot.Disabled = 1
    RightSlingShot.Disabled = 1
    Bumper1.Force = 0
    Bumper2.Force = 0
    Bumper3.Force = 0
    RedSpecialEnd
  Else
    LeftSlingShot.Disabled = 0
    RightSlingShot.Disabled = 0
    Bumper1.Force = 10
    Bumper2.Force = 10
    Bumper3.Force = 10
  End If
end Sub

Sub TiltEnd
  Tilt = 0
  Tilted = False
  lsOFF.StopPlay
  lsOFF.UpdateInterval = 12
  lsOFF.Play SeqUpOn, 30, 1
  If B2SOn Then Controller.B2SSetData 33,0
  For each tlt in TILT_light : tlt.State = 0 : Next
  DisableTable False
end Sub

Sub StopAllSounds
  StopSound "OKLAHOMAspell"
  StopSound "YeeHaw01wav"
  StopSound "WhereTheWindComesSweeping"
  StopSound "boomersooner"
  StopSound "IveBeenToOklahoma"
  StopSound "OkieFromMuskogee"
  StopSound "GetYourKicks"
  StopSound "Tornado"
end Sub



'***************
' end of TILT
'***************


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Anemometer: BP_Anemometer=Array(BM_Anemometer, LM_GI_Bumper1Light_Anemometer, LM_GI_Bumper2Light_Anemometer, LM_GI_GI_TopApron_L_Anemometer, LM_GI_LightWeatherman_Anemomete, LM_IL_LightToplane1_Anemometer)
Dim BP_BRing1: BP_BRing1=Array(BM_BRing1, LM_GI_Bumper1Light_BRing1, LM_GI_Bumper3Light_BRing1, LM_GI_LightsGI_Pops_BRing1, LM_IL_LightToplane1_BRing1, LM_IL_LightToplane2_BRing1)
Dim BP_BRing2: BP_BRing2=Array(BM_BRing2, LM_GI_Bumper1Light_BRing2, LM_GI_Bumper2Light_BRing2, LM_GI_Bumper3Light_BRing2, LM_GI_LightsGI_Pops_BRing2, LM_IL_LightEOB4_BRing2, LM_IL_LightDLR03_BRing2, LM_IL_LightXtraBall_Arrow_BRing)
Dim BP_BRing3: BP_BRing3=Array(BM_BRing3, LM_GI_Bumper2Light_BRing3, LM_GI_Bumper3Light_BRing3, LM_GI_LightsGI_TARamp_BRing3, LM_IL_LightToplane1_BRing3, LM_IL_LightToplane2_BRing3)
Dim BP_Bluey_d2: BP_Bluey_d2=Array(BM_Bluey_d2, LM_GI_Bumper1Light_Bluey_d2, LM_GI_Bumper2Light_Bluey_d2, LM_GI_Bumper3Light_Bluey_d2, LM_GI_LightsGI_Pops_Bluey_d2, LM_GI_GI_PumpJack_Bluey_d2, LM_GI_LightsGI_TARamp_Bluey_d2, LM_IL_LightBLU_Bluey_d2, LM_IL_LightTA_Arrow_Bluey_d2, LM_IL_LightXtraBall_Arrow_Bluey)
Dim BP_Bsocket1: BP_Bsocket1=Array(BM_Bsocket1, LM_GI_Bumper1Light_Bsocket1, LM_GI_Bumper2Light_Bsocket1, LM_GI_Bumper3Light_Bsocket1, LM_GI_LightsGI_Pops_Bsocket1, LM_IL_LightToplane1_Bsocket1, LM_IL_LightToplane2_Bsocket1)
Dim BP_Bsocket2: BP_Bsocket2=Array(BM_Bsocket2, LM_GI_Bumper1Light_Bsocket2, LM_GI_Bumper2Light_Bsocket2, LM_GI_Bumper3Light_Bsocket2, LM_GI_LightsGI_Pops_Bsocket2, LM_IL_LightEOB3_Bsocket2, LM_IL_LightEOB4_Bsocket2, LM_IL_LightDLR03_Bsocket2, LM_IL_LightToplane2_Bsocket2)
Dim BP_Bsocket3: BP_Bsocket3=Array(BM_Bsocket3, LM_GI_Bumper1Light_Bsocket3, LM_GI_Bumper2Light_Bsocket3, LM_GI_Bumper3Light_Bsocket3, LM_GI_LightsGI_TARamp_Bsocket3, LM_IL_LightToplane1_Bsocket3, LM_IL_LightToplane2_Bsocket3, LM_IL_LightToplane3_Bsocket3)
Dim BP_DT_66: BP_DT_66=Array(BM_DT_66, LM_GI_GI_PumpJack_DT_66, LM_GI_RedFlasherLight_DT_66, LM_IL_LightEOB1_DT_66, LM_IL_LightDLR01_DT_66, LM_IL_LightInlane_L_DT_66, LM_IL_LightDLR_Arrow_DT_66, LM_IL_LightRT66_Arrow_DT_66, LM_IL_LightBonus_DT_66, LM_GI_gi4_DT_66)
Dim BP_DT_BLU: BP_DT_BLU=Array(BM_DT_BLU, LM_GI_GI_PumpJack_DT_BLU, LM_IL_LightBLU_DT_BLU, LM_IL_LightEOB1_DT_BLU, LM_IL_LightEOB2_DT_BLU, LM_IL_LightEOB3_DT_BLU, LM_IL_LightTA_Arrow_DT_BLU, LM_IL_LightPJ_DT_BLU)
Dim BP_DT_TA: BP_DT_TA=Array(BM_DT_TA, LM_GI_Bumper3Light_DT_TA, LM_GI_GI_PumpJack_DT_TA, LM_IL_LightEOB2_DT_TA, LM_IL_LightEOB3_DT_TA, LM_IL_LightEOB4_DT_TA, LM_IL_LightTA_Arrow_DT_TA, LM_IL_LightXtraBall_Arrow_DT_TA)
Dim BP_DT_XB: BP_DT_XB=Array(BM_DT_XB, LM_GI_Bumper3Light_DT_XB, LM_GI_GI_PumpJack_DT_XB, LM_IL_LightEOB3_DT_XB, LM_IL_LightEOB4_DT_XB, LM_IL_LightDLR02_DT_XB, LM_IL_LightDLR03_DT_XB, LM_IL_LightXtraBall_Arrow_DT_XB)
Dim BP_Driller_UltraLoPoly: BP_Driller_UltraLoPoly=Array(BM_Driller_UltraLoPoly, LM_GI_Bumper1Light_Driller_Ultr, LM_GI_Bumper2Light_Driller_Ultr, LM_GI_Bumper3Light_Driller_Ultr, LM_GI_LightsGI_Pops_Driller_Ult, LM_GI_RedFlasherLight_Driller_U, LM_GI_GI_TopApron_L_Driller_Ult, LM_GI_LightsGI_TARamp_Driller_U, LM_GI_LightWeatherman_Driller_U, LM_IL_Light_Alley01_Driller_Ult, LM_IL_Light_Alley02_Driller_Ult, LM_IL_Light_Alley03_Driller_Ult, LM_IL_Light_Alley04_Driller_Ult, LM_IL_Light_Alley05_Driller_Ult, LM_IL_LightBLU_Driller_UltraLoP, LM_IL_LightEOB2_Driller_UltraLo, LM_IL_LightEOB3_Driller_UltraLo, LM_IL_LightEOB4_Driller_UltraLo, LM_IL_LightTA_Arrow_Driller_Ult, LM_IL_LightDLR01_Driller_UltraL, LM_IL_LightDLR02_Driller_UltraL, LM_IL_LightDLR03_Driller_UltraL, LM_IL_LightToplane1_Driller_Ult, LM_IL_LightToplane2_Driller_Ult, LM_IL_LightXtraBall_Arrow_Drill)
Dim BP_Funnel: BP_Funnel=Array(BM_Funnel, LM_GI_Bumper1Light_Funnel, LM_GI_Bumper2Light_Funnel, LM_GI_Bumper3Light_Funnel, LM_GI_LightsGI_Pops_Funnel, LM_GI_GI_TopApron_L_Funnel, LM_GI_LightsGI_TARamp_Funnel, LM_GI_LightWeatherman_Funnel, LM_IL_Light_Alley05_Funnel, LM_IL_LightToplane1_Funnel, LM_IL_LightToplane2_Funnel, LM_IL_LightToplane3_Funnel)
Dim BP_GateRT66_Wire: BP_GateRT66_Wire=Array(BM_GateRT66_Wire, LM_GI_GI_LeftSpecial_GateRT66_W, LM_GI_GI_PumpJack_GateRT66_Wire, LM_GI_RedFlasherLight_GateRT66_, LM_IL_LightInlane_L_GateRT66_Wi, LM_IL_LightOutlane_L_GateRT66_W, LM_IL_LightBonus_GateRT66_Wire, LM_GI_gi3_GateRT66_Wire, LM_GI_gi4_GateRT66_Wire)
Dim BP_Gate_Alley_Wire: BP_Gate_Alley_Wire=Array(BM_Gate_Alley_Wire, LM_GI_LightsGI_Pops_Gate_Alley_, LM_GI_LightWeatherman_Gate_Alle)
Dim BP_Gate_Flap: BP_Gate_Flap=Array(BM_Gate_Flap)
Dim BP_Gate_Flap_002: BP_Gate_Flap_002=Array(BM_Gate_Flap_002, LM_GI_GI_PumpJack_Gate_Flap_002, LM_IL_LightPJ_Gate_Flap_002)
Dim BP_LFlipper: BP_LFlipper=Array(BM_LFlipper, LM_IL_LightShootAgain_LFlipper, LM_GI_gi3_LFlipper)
Dim BP_Metal: BP_Metal=Array(BM_Metal, LM_GI_Bumper1Light_Metal, LM_GI_Bumper2Light_Metal, LM_GI_Bumper3Light_Metal, LM_GI_GI_LeftSpecial_Metal, LM_GI_LightsGI_Pops_Metal, LM_GI_GI_PumpJack_Metal, LM_GI_RedFlasherLight_Metal, LM_GI_GI_TopApron_L_Metal, LM_GI_GI_TopApron_R_Metal, LM_GI_LightsGI_TARamp_Metal, LM_GI_LightWeatherman_Metal, LM_IL_Light_Alley02_Metal, LM_IL_Light_Alley03_Metal, LM_IL_LightBLU_Metal, LM_IL_LightE_Metal, LM_IL_LightEOB1_Metal, LM_IL_LightEOB2_Metal, LM_IL_LightEOB3_Metal, LM_IL_LightEOB4_Metal, LM_IL_LightTA_Arrow_Metal, LM_IL_LightI_Metal, LM_IL_LightDLR01_Metal, LM_IL_LightDLR02_Metal, LM_IL_LightDLR03_Metal, LM_IL_LightInlane_L_Metal, LM_IL_LightInlane_R_Metal, LM_IL_LightK_Metal, LM_IL_LightDLR_Arrow_Metal, LM_IL_LightO_Metal, LM_IL_LightRT66_Arrow_Metal, LM_IL_LightOutlane_L_Metal, LM_IL_LightOutlane_R_Metal, LM_IL_LightPJ_Metal, LM_IL_LightBonus_Metal, LM_IL_LightS_Metal, LM_IL_LightBLU_Arrow_Metal, LM_IL_LightShootAgain_Metal, LM_IL_LightToplane1_Metal, _
  LM_IL_LightToplane2_Metal, LM_IL_LightXtraBall_Arrow_Metal, LM_IL_LightToplane3_Metal, LM_GI_gi1_Metal, LM_GI_gi2_Metal, LM_GI_gi3_Metal, LM_GI_gi4_Metal)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_GI_Bumper1Light_Parts, LM_GI_Bumper2Light_Parts, LM_GI_Bumper3Light_Parts, LM_GI_GI_LeftSpecial_Parts, LM_GI_LightsGI_Pops_Parts, LM_GI_GI_PumpJack_Parts, LM_GI_RedFlasherLight_Parts, LM_GI_GI_TopApron_L_Parts, LM_GI_GI_TopApron_R_Parts, LM_GI_LightsGI_TARamp_Parts, LM_GI_LightWeatherman_Parts, LM_IL_Light_Alley01_Parts, LM_IL_Light_Alley02_Parts, LM_IL_Light_Alley03_Parts, LM_IL_Light_Alley04_Parts, LM_IL_Light_Alley05_Parts, LM_IL_LightBLU_Parts, LM_IL_LightE_Parts, LM_IL_LightEOB1_Parts, LM_IL_LightEOB2_Parts, LM_IL_LightEOB3_Parts, LM_IL_LightEOB4_Parts, LM_IL_LightTA_Arrow_Parts, LM_IL_LightI_Parts, LM_IL_LightDLR01_Parts, LM_IL_LightDLR02_Parts, LM_IL_LightDLR03_Parts, LM_IL_LightInlane_L_Parts, LM_IL_LightInlane_R_Parts, LM_IL_LightK_Parts, LM_IL_LightDLR_Arrow_Parts, LM_IL_LightO_Parts, LM_IL_LightRT66_Arrow_Parts, LM_IL_LightOutlane_L_Parts, LM_IL_LightOutlane_R_Parts, LM_IL_LightPJ_Parts, LM_IL_LightBonus_Parts, LM_IL_LightS_Parts, _
  LM_IL_LightBLU_Arrow_Parts, LM_IL_LightShootAgain_Parts, LM_IL_LightToplane1_Parts, LM_IL_LightToplane2_Parts, LM_IL_LightXtraBall_Arrow_Parts, LM_IL_LightToplane3_Parts, LM_GI_gi1_Parts, LM_GI_gi2_Parts, LM_GI_gi3_Parts, LM_GI_gi4_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GI_Bumper1Light_Playfield, LM_GI_Bumper2Light_Playfield, LM_GI_Bumper3Light_Playfield, LM_GI_LightsGI_Pops_Playfield, LM_GI_GI_PumpJack_Playfield, LM_GI_RedFlasherLight_Playfield, LM_GI_GI_TopApron_L_Playfield, LM_GI_GI_TopApron_R_Playfield, LM_GI_LightsGI_TARamp_Playfield, LM_GI_LightWeatherman_Playfield, LM_IL_Light_Alley01_Playfield, LM_IL_Light_Alley02_Playfield, LM_IL_Light_Alley03_Playfield, LM_IL_Light_Alley04_Playfield, LM_IL_Light_Alley05_Playfield, LM_IL_LightBLU_Playfield, LM_IL_LightE_Playfield, LM_IL_LightEOB1_Playfield, LM_IL_LightEOB2_Playfield, LM_IL_LightEOB3_Playfield, LM_IL_LightEOB4_Playfield, LM_IL_LightTA_Arrow_Playfield, LM_IL_LightI_Playfield, LM_IL_LightDLR01_Playfield, LM_IL_LightDLR02_Playfield, LM_IL_LightDLR03_Playfield, LM_IL_LightInlane_L_Playfield, LM_IL_LightInlane_R_Playfield, LM_IL_LightK_Playfield, LM_IL_LightDLR_Arrow_Playfield, LM_IL_LightO_Playfield, LM_IL_LightRT66_Arrow_Playfield, LM_IL_LightOutlane_L_Playfield, _
  LM_IL_LightOutlane_R_Playfield, LM_IL_LightPJ_Playfield, LM_IL_LightBonus_Playfield, LM_IL_LightS_Playfield, LM_IL_LightBLU_Arrow_Playfield, LM_IL_LightShootAgain_Playfield, LM_IL_LightToplane1_Playfield, LM_IL_LightToplane2_Playfield, LM_IL_LightXtraBall_Arrow_Playf, LM_IL_LightToplane3_Playfield)
Dim BP_RFlipper: BP_RFlipper=Array(BM_RFlipper, LM_IL_LightShootAgain_RFlipper, LM_GI_gi1_RFlipper)
Dim BP_RoundTgt01: BP_RoundTgt01=Array(BM_RoundTgt01, LM_GI_Bumper3Light_RoundTgt01, LM_GI_LightsGI_Pops_RoundTgt01, LM_GI_RedFlasherLight_RoundTgt0, LM_IL_LightDLR01_RoundTgt01, LM_IL_LightDLR02_RoundTgt01, LM_IL_LightDLR03_RoundTgt01)
Dim BP_RoundTgt02: BP_RoundTgt02=Array(BM_RoundTgt02, LM_GI_Bumper3Light_RoundTgt02, LM_GI_LightsGI_Pops_RoundTgt02, LM_GI_RedFlasherLight_RoundTgt0, LM_IL_LightEOB4_RoundTgt02, LM_IL_LightDLR01_RoundTgt02, LM_IL_LightDLR02_RoundTgt02, LM_IL_LightDLR03_RoundTgt02, LM_IL_LightXtraBall_Arrow_Round)
Dim BP_RoundTgt03: BP_RoundTgt03=Array(BM_RoundTgt03, LM_GI_Bumper3Light_RoundTgt03, LM_IL_LightEOB4_RoundTgt03, LM_IL_LightDLR01_RoundTgt03, LM_IL_LightDLR02_RoundTgt03, LM_IL_LightDLR03_RoundTgt03, LM_IL_LightXtraBall_Arrow_Round)
Dim BP_SW1: BP_SW1=Array(BM_SW1, LM_IL_LightToplane1_SW1)
Dim BP_SW2: BP_SW2=Array(BM_SW2, LM_IL_LightToplane1_SW2, LM_IL_LightToplane2_SW2)
Dim BP_SW3: BP_SW3=Array(BM_SW3, LM_IL_LightToplane1_SW3)
Dim BP_SW4: BP_SW4=Array(BM_SW4, LM_IL_LightInlane_R_SW4, LM_IL_LightOutlane_R_SW4, LM_GI_gi2_SW4)
Dim BP_SW5: BP_SW5=Array(BM_SW5, LM_IL_LightOutlane_R_SW5)
Dim BP_SW6: BP_SW6=Array(BM_SW6, LM_IL_LightInlane_L_SW6, LM_IL_LightOutlane_L_SW6, LM_GI_gi4_SW6)
Dim BP_SW7: BP_SW7=Array(BM_SW7, LM_IL_LightOutlane_L_SW7)
Dim BP_Spinner_Alley_Wire: BP_Spinner_Alley_Wire=Array(BM_Spinner_Alley_Wire, LM_GI_GI_LeftSpecial_Spinner_Al, LM_GI_RedFlasherLight_Spinner_A, LM_IL_Light_Alley01_Spinner_All, LM_IL_LightInlane_L_Spinner_All, LM_IL_LightRT66_Arrow_Spinner_A, LM_IL_LightBonus_Spinner_Alley_)
' Arrays per lighting scenario
Dim BL_GI_Bumper1Light: BL_GI_Bumper1Light=Array(LM_GI_Bumper1Light_Anemometer, LM_GI_Bumper1Light_BRing1, LM_GI_Bumper1Light_BRing2, LM_GI_Bumper1Light_Bluey_d2, LM_GI_Bumper1Light_Bsocket1, LM_GI_Bumper1Light_Bsocket2, LM_GI_Bumper1Light_Bsocket3, LM_GI_Bumper1Light_Driller_Ultr, LM_GI_Bumper1Light_Funnel, LM_GI_Bumper1Light_Metal, LM_GI_Bumper1Light_Parts, LM_GI_Bumper1Light_Playfield)
Dim BL_GI_Bumper2Light: BL_GI_Bumper2Light=Array(LM_GI_Bumper2Light_Anemometer, LM_GI_Bumper2Light_BRing2, LM_GI_Bumper2Light_BRing3, LM_GI_Bumper2Light_Bluey_d2, LM_GI_Bumper2Light_Bsocket1, LM_GI_Bumper2Light_Bsocket2, LM_GI_Bumper2Light_Bsocket3, LM_GI_Bumper2Light_Driller_Ultr, LM_GI_Bumper2Light_Funnel, LM_GI_Bumper2Light_Metal, LM_GI_Bumper2Light_Parts, LM_GI_Bumper2Light_Playfield)
Dim BL_GI_Bumper3Light: BL_GI_Bumper3Light=Array(LM_GI_Bumper3Light_BRing1, LM_GI_Bumper3Light_BRing2, LM_GI_Bumper3Light_BRing3, LM_GI_Bumper3Light_Bluey_d2, LM_GI_Bumper3Light_Bsocket1, LM_GI_Bumper3Light_Bsocket2, LM_GI_Bumper3Light_Bsocket3, LM_GI_Bumper3Light_DT_TA, LM_GI_Bumper3Light_DT_XB, LM_GI_Bumper3Light_Driller_Ultr, LM_GI_Bumper3Light_Funnel, LM_GI_Bumper3Light_Metal, LM_GI_Bumper3Light_Parts, LM_GI_Bumper3Light_Playfield, LM_GI_Bumper3Light_RoundTgt01, LM_GI_Bumper3Light_RoundTgt02, LM_GI_Bumper3Light_RoundTgt03)
Dim BL_GI_GI_LeftSpecial: BL_GI_GI_LeftSpecial=Array(LM_GI_GI_LeftSpecial_GateRT66_W, LM_GI_GI_LeftSpecial_Metal, LM_GI_GI_LeftSpecial_Parts, LM_GI_GI_LeftSpecial_Spinner_Al)
Dim BL_GI_GI_PumpJack: BL_GI_GI_PumpJack=Array(LM_GI_GI_PumpJack_Bluey_d2, LM_GI_GI_PumpJack_DT_66, LM_GI_GI_PumpJack_DT_BLU, LM_GI_GI_PumpJack_DT_TA, LM_GI_GI_PumpJack_DT_XB, LM_GI_GI_PumpJack_Gate_Flap_002, LM_GI_GI_PumpJack_GateRT66_Wire, LM_GI_GI_PumpJack_Metal, LM_GI_GI_PumpJack_Parts, LM_GI_GI_PumpJack_Playfield)
Dim BL_GI_GI_TopApron_L: BL_GI_GI_TopApron_L=Array(LM_GI_GI_TopApron_L_Anemometer, LM_GI_GI_TopApron_L_Driller_Ult, LM_GI_GI_TopApron_L_Funnel, LM_GI_GI_TopApron_L_Metal, LM_GI_GI_TopApron_L_Parts, LM_GI_GI_TopApron_L_Playfield)
Dim BL_GI_GI_TopApron_R: BL_GI_GI_TopApron_R=Array(LM_GI_GI_TopApron_R_Metal, LM_GI_GI_TopApron_R_Parts, LM_GI_GI_TopApron_R_Playfield)
Dim BL_GI_LightWeatherman: BL_GI_LightWeatherman=Array(LM_GI_LightWeatherman_Anemomete, LM_GI_LightWeatherman_Driller_U, LM_GI_LightWeatherman_Funnel, LM_GI_LightWeatherman_Gate_Alle, LM_GI_LightWeatherman_Metal, LM_GI_LightWeatherman_Parts, LM_GI_LightWeatherman_Playfield)
Dim BL_GI_LightsGI_Pops: BL_GI_LightsGI_Pops=Array(LM_GI_LightsGI_Pops_BRing1, LM_GI_LightsGI_Pops_BRing2, LM_GI_LightsGI_Pops_Bluey_d2, LM_GI_LightsGI_Pops_Bsocket1, LM_GI_LightsGI_Pops_Bsocket2, LM_GI_LightsGI_Pops_Driller_Ult, LM_GI_LightsGI_Pops_Funnel, LM_GI_LightsGI_Pops_Gate_Alley_, LM_GI_LightsGI_Pops_Metal, LM_GI_LightsGI_Pops_Parts, LM_GI_LightsGI_Pops_Playfield, LM_GI_LightsGI_Pops_RoundTgt01, LM_GI_LightsGI_Pops_RoundTgt02)
Dim BL_GI_LightsGI_TARamp: BL_GI_LightsGI_TARamp=Array(LM_GI_LightsGI_TARamp_BRing3, LM_GI_LightsGI_TARamp_Bluey_d2, LM_GI_LightsGI_TARamp_Bsocket3, LM_GI_LightsGI_TARamp_Driller_U, LM_GI_LightsGI_TARamp_Funnel, LM_GI_LightsGI_TARamp_Metal, LM_GI_LightsGI_TARamp_Parts, LM_GI_LightsGI_TARamp_Playfield)
Dim BL_GI_RedFlasherLight: BL_GI_RedFlasherLight=Array(LM_GI_RedFlasherLight_DT_66, LM_GI_RedFlasherLight_Driller_U, LM_GI_RedFlasherLight_GateRT66_, LM_GI_RedFlasherLight_Metal, LM_GI_RedFlasherLight_Parts, LM_GI_RedFlasherLight_Playfield, LM_GI_RedFlasherLight_RoundTgt0, LM_GI_RedFlasherLight_RoundTgt0, LM_GI_RedFlasherLight_Spinner_A)
Dim BL_GI_gi1: BL_GI_gi1=Array(LM_GI_gi1_Metal, LM_GI_gi1_Parts, LM_GI_gi1_RFlipper)
Dim BL_GI_gi2: BL_GI_gi2=Array(LM_GI_gi2_Metal, LM_GI_gi2_Parts, LM_GI_gi2_SW4)
Dim BL_GI_gi3: BL_GI_gi3=Array(LM_GI_gi3_GateRT66_Wire, LM_GI_gi3_LFlipper, LM_GI_gi3_Metal, LM_GI_gi3_Parts)
Dim BL_GI_gi4: BL_GI_gi4=Array(LM_GI_gi4_DT_66, LM_GI_gi4_GateRT66_Wire, LM_GI_gi4_Metal, LM_GI_gi4_Parts, LM_GI_gi4_SW6)
Dim BL_IL_LightBLU: BL_IL_LightBLU=Array(LM_IL_LightBLU_Bluey_d2, LM_IL_LightBLU_DT_BLU, LM_IL_LightBLU_Driller_UltraLoP, LM_IL_LightBLU_Metal, LM_IL_LightBLU_Parts, LM_IL_LightBLU_Playfield)
Dim BL_IL_LightBLU_Arrow: BL_IL_LightBLU_Arrow=Array(LM_IL_LightBLU_Arrow_Metal, LM_IL_LightBLU_Arrow_Parts, LM_IL_LightBLU_Arrow_Playfield)
Dim BL_IL_LightBonus: BL_IL_LightBonus=Array(LM_IL_LightBonus_DT_66, LM_IL_LightBonus_GateRT66_Wire, LM_IL_LightBonus_Metal, LM_IL_LightBonus_Parts, LM_IL_LightBonus_Playfield, LM_IL_LightBonus_Spinner_Alley_)
Dim BL_IL_LightDLR01: BL_IL_LightDLR01=Array(LM_IL_LightDLR01_DT_66, LM_IL_LightDLR01_Driller_UltraL, LM_IL_LightDLR01_Metal, LM_IL_LightDLR01_Parts, LM_IL_LightDLR01_Playfield, LM_IL_LightDLR01_RoundTgt01, LM_IL_LightDLR01_RoundTgt02, LM_IL_LightDLR01_RoundTgt03)
Dim BL_IL_LightDLR02: BL_IL_LightDLR02=Array(LM_IL_LightDLR02_DT_XB, LM_IL_LightDLR02_Driller_UltraL, LM_IL_LightDLR02_Metal, LM_IL_LightDLR02_Parts, LM_IL_LightDLR02_Playfield, LM_IL_LightDLR02_RoundTgt01, LM_IL_LightDLR02_RoundTgt02, LM_IL_LightDLR02_RoundTgt03)
Dim BL_IL_LightDLR03: BL_IL_LightDLR03=Array(LM_IL_LightDLR03_BRing2, LM_IL_LightDLR03_Bsocket2, LM_IL_LightDLR03_DT_XB, LM_IL_LightDLR03_Driller_UltraL, LM_IL_LightDLR03_Metal, LM_IL_LightDLR03_Parts, LM_IL_LightDLR03_Playfield, LM_IL_LightDLR03_RoundTgt01, LM_IL_LightDLR03_RoundTgt02, LM_IL_LightDLR03_RoundTgt03)
Dim BL_IL_LightDLR_Arrow: BL_IL_LightDLR_Arrow=Array(LM_IL_LightDLR_Arrow_DT_66, LM_IL_LightDLR_Arrow_Metal, LM_IL_LightDLR_Arrow_Parts, LM_IL_LightDLR_Arrow_Playfield)
Dim BL_IL_LightE: BL_IL_LightE=Array(LM_IL_LightE_Metal, LM_IL_LightE_Parts, LM_IL_LightE_Playfield)
Dim BL_IL_LightEOB1: BL_IL_LightEOB1=Array(LM_IL_LightEOB1_DT_66, LM_IL_LightEOB1_DT_BLU, LM_IL_LightEOB1_Metal, LM_IL_LightEOB1_Parts, LM_IL_LightEOB1_Playfield)
Dim BL_IL_LightEOB2: BL_IL_LightEOB2=Array(LM_IL_LightEOB2_DT_BLU, LM_IL_LightEOB2_DT_TA, LM_IL_LightEOB2_Driller_UltraLo, LM_IL_LightEOB2_Metal, LM_IL_LightEOB2_Parts, LM_IL_LightEOB2_Playfield)
Dim BL_IL_LightEOB3: BL_IL_LightEOB3=Array(LM_IL_LightEOB3_Bsocket2, LM_IL_LightEOB3_DT_BLU, LM_IL_LightEOB3_DT_TA, LM_IL_LightEOB3_DT_XB, LM_IL_LightEOB3_Driller_UltraLo, LM_IL_LightEOB3_Metal, LM_IL_LightEOB3_Parts, LM_IL_LightEOB3_Playfield)
Dim BL_IL_LightEOB4: BL_IL_LightEOB4=Array(LM_IL_LightEOB4_BRing2, LM_IL_LightEOB4_Bsocket2, LM_IL_LightEOB4_DT_TA, LM_IL_LightEOB4_DT_XB, LM_IL_LightEOB4_Driller_UltraLo, LM_IL_LightEOB4_Metal, LM_IL_LightEOB4_Parts, LM_IL_LightEOB4_Playfield, LM_IL_LightEOB4_RoundTgt02, LM_IL_LightEOB4_RoundTgt03)
Dim BL_IL_LightI: BL_IL_LightI=Array(LM_IL_LightI_Metal, LM_IL_LightI_Parts, LM_IL_LightI_Playfield)
Dim BL_IL_LightInlane_L: BL_IL_LightInlane_L=Array(LM_IL_LightInlane_L_DT_66, LM_IL_LightInlane_L_GateRT66_Wi, LM_IL_LightInlane_L_Metal, LM_IL_LightInlane_L_Parts, LM_IL_LightInlane_L_Playfield, LM_IL_LightInlane_L_SW6, LM_IL_LightInlane_L_Spinner_All)
Dim BL_IL_LightInlane_R: BL_IL_LightInlane_R=Array(LM_IL_LightInlane_R_Metal, LM_IL_LightInlane_R_Parts, LM_IL_LightInlane_R_Playfield, LM_IL_LightInlane_R_SW4)
Dim BL_IL_LightK: BL_IL_LightK=Array(LM_IL_LightK_Metal, LM_IL_LightK_Parts, LM_IL_LightK_Playfield)
Dim BL_IL_LightO: BL_IL_LightO=Array(LM_IL_LightO_Metal, LM_IL_LightO_Parts, LM_IL_LightO_Playfield)
Dim BL_IL_LightOutlane_L: BL_IL_LightOutlane_L=Array(LM_IL_LightOutlane_L_GateRT66_W, LM_IL_LightOutlane_L_Metal, LM_IL_LightOutlane_L_Parts, LM_IL_LightOutlane_L_Playfield, LM_IL_LightOutlane_L_SW6, LM_IL_LightOutlane_L_SW7)
Dim BL_IL_LightOutlane_R: BL_IL_LightOutlane_R=Array(LM_IL_LightOutlane_R_Metal, LM_IL_LightOutlane_R_Parts, LM_IL_LightOutlane_R_Playfield, LM_IL_LightOutlane_R_SW4, LM_IL_LightOutlane_R_SW5)
Dim BL_IL_LightPJ: BL_IL_LightPJ=Array(LM_IL_LightPJ_DT_BLU, LM_IL_LightPJ_Gate_Flap_002, LM_IL_LightPJ_Metal, LM_IL_LightPJ_Parts, LM_IL_LightPJ_Playfield)
Dim BL_IL_LightRT66_Arrow: BL_IL_LightRT66_Arrow=Array(LM_IL_LightRT66_Arrow_DT_66, LM_IL_LightRT66_Arrow_Metal, LM_IL_LightRT66_Arrow_Parts, LM_IL_LightRT66_Arrow_Playfield, LM_IL_LightRT66_Arrow_Spinner_A)
Dim BL_IL_LightS: BL_IL_LightS=Array(LM_IL_LightS_Metal, LM_IL_LightS_Parts, LM_IL_LightS_Playfield)
Dim BL_IL_LightShootAgain: BL_IL_LightShootAgain=Array(LM_IL_LightShootAgain_LFlipper, LM_IL_LightShootAgain_Metal, LM_IL_LightShootAgain_Parts, LM_IL_LightShootAgain_Playfield, LM_IL_LightShootAgain_RFlipper)
Dim BL_IL_LightTA_Arrow: BL_IL_LightTA_Arrow=Array(LM_IL_LightTA_Arrow_Bluey_d2, LM_IL_LightTA_Arrow_DT_BLU, LM_IL_LightTA_Arrow_DT_TA, LM_IL_LightTA_Arrow_Driller_Ult, LM_IL_LightTA_Arrow_Metal, LM_IL_LightTA_Arrow_Parts, LM_IL_LightTA_Arrow_Playfield)
Dim BL_IL_LightToplane1: BL_IL_LightToplane1=Array(LM_IL_LightToplane1_Anemometer, LM_IL_LightToplane1_BRing1, LM_IL_LightToplane1_BRing3, LM_IL_LightToplane1_Bsocket1, LM_IL_LightToplane1_Bsocket3, LM_IL_LightToplane1_Driller_Ult, LM_IL_LightToplane1_Funnel, LM_IL_LightToplane1_Metal, LM_IL_LightToplane1_Parts, LM_IL_LightToplane1_Playfield, LM_IL_LightToplane1_SW1, LM_IL_LightToplane1_SW2, LM_IL_LightToplane1_SW3)
Dim BL_IL_LightToplane2: BL_IL_LightToplane2=Array(LM_IL_LightToplane2_BRing1, LM_IL_LightToplane2_BRing3, LM_IL_LightToplane2_Bsocket1, LM_IL_LightToplane2_Bsocket2, LM_IL_LightToplane2_Bsocket3, LM_IL_LightToplane2_Driller_Ult, LM_IL_LightToplane2_Funnel, LM_IL_LightToplane2_Metal, LM_IL_LightToplane2_Parts, LM_IL_LightToplane2_Playfield, LM_IL_LightToplane2_SW2)
Dim BL_IL_LightToplane3: BL_IL_LightToplane3=Array(LM_IL_LightToplane3_Bsocket3, LM_IL_LightToplane3_Funnel, LM_IL_LightToplane3_Metal, LM_IL_LightToplane3_Parts, LM_IL_LightToplane3_Playfield)
Dim BL_IL_LightXtraBall_Arrow: BL_IL_LightXtraBall_Arrow=Array(LM_IL_LightXtraBall_Arrow_BRing, LM_IL_LightXtraBall_Arrow_Bluey, LM_IL_LightXtraBall_Arrow_DT_TA, LM_IL_LightXtraBall_Arrow_DT_XB, LM_IL_LightXtraBall_Arrow_Drill, LM_IL_LightXtraBall_Arrow_Metal, LM_IL_LightXtraBall_Arrow_Parts, LM_IL_LightXtraBall_Arrow_Playf, LM_IL_LightXtraBall_Arrow_Round, LM_IL_LightXtraBall_Arrow_Round)
Dim BL_IL_Light_Alley01: BL_IL_Light_Alley01=Array(LM_IL_Light_Alley01_Driller_Ult, LM_IL_Light_Alley01_Parts, LM_IL_Light_Alley01_Playfield, LM_IL_Light_Alley01_Spinner_All)
Dim BL_IL_Light_Alley02: BL_IL_Light_Alley02=Array(LM_IL_Light_Alley02_Driller_Ult, LM_IL_Light_Alley02_Metal, LM_IL_Light_Alley02_Parts, LM_IL_Light_Alley02_Playfield)
Dim BL_IL_Light_Alley03: BL_IL_Light_Alley03=Array(LM_IL_Light_Alley03_Driller_Ult, LM_IL_Light_Alley03_Metal, LM_IL_Light_Alley03_Parts, LM_IL_Light_Alley03_Playfield)
Dim BL_IL_Light_Alley04: BL_IL_Light_Alley04=Array(LM_IL_Light_Alley04_Driller_Ult, LM_IL_Light_Alley04_Parts, LM_IL_Light_Alley04_Playfield)
Dim BL_IL_Light_Alley05: BL_IL_Light_Alley05=Array(LM_IL_Light_Alley05_Driller_Ult, LM_IL_Light_Alley05_Funnel, LM_IL_Light_Alley05_Parts, LM_IL_Light_Alley05_Playfield)
Dim BL_World: BL_World=Array(BM_Anemometer, BM_BRing1, BM_BRing2, BM_BRing3, BM_Bluey_d2, BM_Bsocket1, BM_Bsocket2, BM_Bsocket3, BM_DT_66, BM_DT_BLU, BM_DT_TA, BM_DT_XB, BM_Driller_UltraLoPoly, BM_Funnel, BM_Gate_Flap, BM_Gate_Flap_002, BM_GateRT66_Wire, BM_Gate_Alley_Wire, BM_LFlipper, BM_Metal, BM_Parts, BM_Playfield, BM_RFlipper, BM_RoundTgt01, BM_RoundTgt02, BM_RoundTgt03, BM_SW1, BM_SW2, BM_SW3, BM_SW4, BM_SW5, BM_SW6, BM_SW7, BM_Spinner_Alley_Wire)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Anemometer, BM_BRing1, BM_BRing2, BM_BRing3, BM_Bluey_d2, BM_Bsocket1, BM_Bsocket2, BM_Bsocket3, BM_DT_66, BM_DT_BLU, BM_DT_TA, BM_DT_XB, BM_Driller_UltraLoPoly, BM_Funnel, BM_Gate_Flap, BM_Gate_Flap_002, BM_GateRT66_Wire, BM_Gate_Alley_Wire, BM_LFlipper, BM_Metal, BM_Parts, BM_Playfield, BM_RFlipper, BM_RoundTgt01, BM_RoundTgt02, BM_RoundTgt03, BM_SW1, BM_SW2, BM_SW3, BM_SW4, BM_SW5, BM_SW6, BM_SW7, BM_Spinner_Alley_Wire)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_Bumper1Light_Anemometer, LM_GI_Bumper1Light_BRing1, LM_GI_Bumper1Light_BRing2, LM_GI_Bumper1Light_Bluey_d2, LM_GI_Bumper1Light_Bsocket1, LM_GI_Bumper1Light_Bsocket2, LM_GI_Bumper1Light_Bsocket3, LM_GI_Bumper1Light_Driller_Ultr, LM_GI_Bumper1Light_Funnel, LM_GI_Bumper1Light_Metal, LM_GI_Bumper1Light_Parts, LM_GI_Bumper1Light_Playfield, LM_GI_Bumper2Light_Anemometer, LM_GI_Bumper2Light_BRing2, LM_GI_Bumper2Light_BRing3, LM_GI_Bumper2Light_Bluey_d2, LM_GI_Bumper2Light_Bsocket1, LM_GI_Bumper2Light_Bsocket2, LM_GI_Bumper2Light_Bsocket3, LM_GI_Bumper2Light_Driller_Ultr, LM_GI_Bumper2Light_Funnel, LM_GI_Bumper2Light_Metal, LM_GI_Bumper2Light_Parts, LM_GI_Bumper2Light_Playfield, LM_GI_Bumper3Light_BRing1, LM_GI_Bumper3Light_BRing2, LM_GI_Bumper3Light_BRing3, LM_GI_Bumper3Light_Bluey_d2, LM_GI_Bumper3Light_Bsocket1, LM_GI_Bumper3Light_Bsocket2, LM_GI_Bumper3Light_Bsocket3, LM_GI_Bumper3Light_DT_TA, LM_GI_Bumper3Light_DT_XB, LM_GI_Bumper3Light_Driller_Ultr, _
  LM_GI_Bumper3Light_Funnel, LM_GI_Bumper3Light_Metal, LM_GI_Bumper3Light_Parts, LM_GI_Bumper3Light_Playfield, LM_GI_Bumper3Light_RoundTgt01, LM_GI_Bumper3Light_RoundTgt02, LM_GI_Bumper3Light_RoundTgt03, LM_GI_GI_LeftSpecial_GateRT66_W, LM_GI_GI_LeftSpecial_Metal, LM_GI_GI_LeftSpecial_Parts, LM_GI_GI_LeftSpecial_Spinner_Al, LM_GI_GI_PumpJack_Bluey_d2, LM_GI_GI_PumpJack_DT_66, LM_GI_GI_PumpJack_DT_BLU, LM_GI_GI_PumpJack_DT_TA, LM_GI_GI_PumpJack_DT_XB, LM_GI_GI_PumpJack_Gate_Flap_002, LM_GI_GI_PumpJack_GateRT66_Wire, LM_GI_GI_PumpJack_Metal, LM_GI_GI_PumpJack_Parts, LM_GI_GI_PumpJack_Playfield, LM_GI_GI_TopApron_L_Anemometer, LM_GI_GI_TopApron_L_Driller_Ult, LM_GI_GI_TopApron_L_Funnel, LM_GI_GI_TopApron_L_Metal, LM_GI_GI_TopApron_L_Parts, LM_GI_GI_TopApron_L_Playfield, LM_GI_GI_TopApron_R_Metal, LM_GI_GI_TopApron_R_Parts, LM_GI_GI_TopApron_R_Playfield, LM_GI_LightWeatherman_Anemomete, LM_GI_LightWeatherman_Driller_U, LM_GI_LightWeatherman_Funnel, LM_GI_LightWeatherman_Gate_Alle, LM_GI_LightWeatherman_Metal, _
  LM_GI_LightWeatherman_Parts, LM_GI_LightWeatherman_Playfield, LM_GI_LightsGI_Pops_BRing1, LM_GI_LightsGI_Pops_BRing2, LM_GI_LightsGI_Pops_Bluey_d2, LM_GI_LightsGI_Pops_Bsocket1, LM_GI_LightsGI_Pops_Bsocket2, LM_GI_LightsGI_Pops_Driller_Ult, LM_GI_LightsGI_Pops_Funnel, LM_GI_LightsGI_Pops_Gate_Alley_, LM_GI_LightsGI_Pops_Metal, LM_GI_LightsGI_Pops_Parts, LM_GI_LightsGI_Pops_Playfield, LM_GI_LightsGI_Pops_RoundTgt01, LM_GI_LightsGI_Pops_RoundTgt02, LM_GI_LightsGI_TARamp_BRing3, LM_GI_LightsGI_TARamp_Bluey_d2, LM_GI_LightsGI_TARamp_Bsocket3, LM_GI_LightsGI_TARamp_Driller_U, LM_GI_LightsGI_TARamp_Funnel, LM_GI_LightsGI_TARamp_Metal, LM_GI_LightsGI_TARamp_Parts, LM_GI_LightsGI_TARamp_Playfield, LM_GI_RedFlasherLight_DT_66, LM_GI_RedFlasherLight_Driller_U, LM_GI_RedFlasherLight_GateRT66_, LM_GI_RedFlasherLight_Metal, LM_GI_RedFlasherLight_Parts, LM_GI_RedFlasherLight_Playfield, LM_GI_RedFlasherLight_RoundTgt0, LM_GI_RedFlasherLight_RoundTgt0, LM_GI_RedFlasherLight_Spinner_A, LM_GI_gi1_Metal, LM_GI_gi1_Parts, _
  LM_GI_gi1_RFlipper, LM_GI_gi2_Metal, LM_GI_gi2_Parts, LM_GI_gi2_SW4, LM_GI_gi3_GateRT66_Wire, LM_GI_gi3_LFlipper, LM_GI_gi3_Metal, LM_GI_gi3_Parts, LM_GI_gi4_DT_66, LM_GI_gi4_GateRT66_Wire, LM_GI_gi4_Metal, LM_GI_gi4_Parts, LM_GI_gi4_SW6, LM_IL_LightBLU_Bluey_d2, LM_IL_LightBLU_DT_BLU, LM_IL_LightBLU_Driller_UltraLoP, LM_IL_LightBLU_Metal, LM_IL_LightBLU_Parts, LM_IL_LightBLU_Playfield, LM_IL_LightBLU_Arrow_Metal, LM_IL_LightBLU_Arrow_Parts, LM_IL_LightBLU_Arrow_Playfield, LM_IL_LightBonus_DT_66, LM_IL_LightBonus_GateRT66_Wire, LM_IL_LightBonus_Metal, LM_IL_LightBonus_Parts, LM_IL_LightBonus_Playfield, LM_IL_LightBonus_Spinner_Alley_, LM_IL_LightDLR01_DT_66, LM_IL_LightDLR01_Driller_UltraL, LM_IL_LightDLR01_Metal, LM_IL_LightDLR01_Parts, LM_IL_LightDLR01_Playfield, LM_IL_LightDLR01_RoundTgt01, LM_IL_LightDLR01_RoundTgt02, LM_IL_LightDLR01_RoundTgt03, LM_IL_LightDLR02_DT_XB, LM_IL_LightDLR02_Driller_UltraL, LM_IL_LightDLR02_Metal, LM_IL_LightDLR02_Parts, LM_IL_LightDLR02_Playfield, LM_IL_LightDLR02_RoundTgt01, _
  LM_IL_LightDLR02_RoundTgt02, LM_IL_LightDLR02_RoundTgt03, LM_IL_LightDLR03_BRing2, LM_IL_LightDLR03_Bsocket2, LM_IL_LightDLR03_DT_XB, LM_IL_LightDLR03_Driller_UltraL, LM_IL_LightDLR03_Metal, LM_IL_LightDLR03_Parts, LM_IL_LightDLR03_Playfield, LM_IL_LightDLR03_RoundTgt01, LM_IL_LightDLR03_RoundTgt02, LM_IL_LightDLR03_RoundTgt03, LM_IL_LightDLR_Arrow_DT_66, LM_IL_LightDLR_Arrow_Metal, LM_IL_LightDLR_Arrow_Parts, LM_IL_LightDLR_Arrow_Playfield, LM_IL_LightE_Metal, LM_IL_LightE_Parts, LM_IL_LightE_Playfield, LM_IL_LightEOB1_DT_66, LM_IL_LightEOB1_DT_BLU, LM_IL_LightEOB1_Metal, LM_IL_LightEOB1_Parts, LM_IL_LightEOB1_Playfield, LM_IL_LightEOB2_DT_BLU, LM_IL_LightEOB2_DT_TA, LM_IL_LightEOB2_Driller_UltraLo, LM_IL_LightEOB2_Metal, LM_IL_LightEOB2_Parts, LM_IL_LightEOB2_Playfield, LM_IL_LightEOB3_Bsocket2, LM_IL_LightEOB3_DT_BLU, LM_IL_LightEOB3_DT_TA, LM_IL_LightEOB3_DT_XB, LM_IL_LightEOB3_Driller_UltraLo, LM_IL_LightEOB3_Metal, LM_IL_LightEOB3_Parts, LM_IL_LightEOB3_Playfield, LM_IL_LightEOB4_BRing2, _
  LM_IL_LightEOB4_Bsocket2, LM_IL_LightEOB4_DT_TA, LM_IL_LightEOB4_DT_XB, LM_IL_LightEOB4_Driller_UltraLo, LM_IL_LightEOB4_Metal, LM_IL_LightEOB4_Parts, LM_IL_LightEOB4_Playfield, LM_IL_LightEOB4_RoundTgt02, LM_IL_LightEOB4_RoundTgt03, LM_IL_LightI_Metal, LM_IL_LightI_Parts, LM_IL_LightI_Playfield, LM_IL_LightInlane_L_DT_66, LM_IL_LightInlane_L_GateRT66_Wi, LM_IL_LightInlane_L_Metal, LM_IL_LightInlane_L_Parts, LM_IL_LightInlane_L_Playfield, LM_IL_LightInlane_L_SW6, LM_IL_LightInlane_L_Spinner_All, LM_IL_LightInlane_R_Metal, LM_IL_LightInlane_R_Parts, LM_IL_LightInlane_R_Playfield, LM_IL_LightInlane_R_SW4, LM_IL_LightK_Metal, LM_IL_LightK_Parts, LM_IL_LightK_Playfield, LM_IL_LightO_Metal, LM_IL_LightO_Parts, LM_IL_LightO_Playfield, LM_IL_LightOutlane_L_GateRT66_W, LM_IL_LightOutlane_L_Metal, LM_IL_LightOutlane_L_Parts, LM_IL_LightOutlane_L_Playfield, LM_IL_LightOutlane_L_SW6, LM_IL_LightOutlane_L_SW7, LM_IL_LightOutlane_R_Metal, LM_IL_LightOutlane_R_Parts, LM_IL_LightOutlane_R_Playfield, _
  LM_IL_LightOutlane_R_SW4, LM_IL_LightOutlane_R_SW5, LM_IL_LightPJ_DT_BLU, LM_IL_LightPJ_Gate_Flap_002, LM_IL_LightPJ_Metal, LM_IL_LightPJ_Parts, LM_IL_LightPJ_Playfield, LM_IL_LightRT66_Arrow_DT_66, LM_IL_LightRT66_Arrow_Metal, LM_IL_LightRT66_Arrow_Parts, LM_IL_LightRT66_Arrow_Playfield, LM_IL_LightRT66_Arrow_Spinner_A, LM_IL_LightS_Metal, LM_IL_LightS_Parts, LM_IL_LightS_Playfield, LM_IL_LightShootAgain_LFlipper, LM_IL_LightShootAgain_Metal, LM_IL_LightShootAgain_Parts, LM_IL_LightShootAgain_Playfield, LM_IL_LightShootAgain_RFlipper, LM_IL_LightTA_Arrow_Bluey_d2, LM_IL_LightTA_Arrow_DT_BLU, LM_IL_LightTA_Arrow_DT_TA, LM_IL_LightTA_Arrow_Driller_Ult, LM_IL_LightTA_Arrow_Metal, LM_IL_LightTA_Arrow_Parts, LM_IL_LightTA_Arrow_Playfield, LM_IL_LightToplane1_Anemometer, LM_IL_LightToplane1_BRing1, LM_IL_LightToplane1_BRing3, LM_IL_LightToplane1_Bsocket1, LM_IL_LightToplane1_Bsocket3, LM_IL_LightToplane1_Driller_Ult, LM_IL_LightToplane1_Funnel, LM_IL_LightToplane1_Metal, LM_IL_LightToplane1_Parts, _
  LM_IL_LightToplane1_Playfield, LM_IL_LightToplane1_SW1, LM_IL_LightToplane1_SW2, LM_IL_LightToplane1_SW3, LM_IL_LightToplane2_BRing1, LM_IL_LightToplane2_BRing3, LM_IL_LightToplane2_Bsocket1, LM_IL_LightToplane2_Bsocket2, LM_IL_LightToplane2_Bsocket3, LM_IL_LightToplane2_Driller_Ult, LM_IL_LightToplane2_Funnel, LM_IL_LightToplane2_Metal, LM_IL_LightToplane2_Parts, LM_IL_LightToplane2_Playfield, LM_IL_LightToplane2_SW2, LM_IL_LightToplane3_Bsocket3, LM_IL_LightToplane3_Funnel, LM_IL_LightToplane3_Metal, LM_IL_LightToplane3_Parts, LM_IL_LightToplane3_Playfield, LM_IL_LightXtraBall_Arrow_BRing, LM_IL_LightXtraBall_Arrow_Bluey, LM_IL_LightXtraBall_Arrow_DT_TA, LM_IL_LightXtraBall_Arrow_DT_XB, LM_IL_LightXtraBall_Arrow_Drill, LM_IL_LightXtraBall_Arrow_Metal, LM_IL_LightXtraBall_Arrow_Parts, LM_IL_LightXtraBall_Arrow_Playf, LM_IL_LightXtraBall_Arrow_Round, LM_IL_LightXtraBall_Arrow_Round, LM_IL_Light_Alley01_Driller_Ult, LM_IL_Light_Alley01_Parts, LM_IL_Light_Alley01_Playfield, LM_IL_Light_Alley01_Spinner_All, _
  LM_IL_Light_Alley02_Driller_Ult, LM_IL_Light_Alley02_Metal, LM_IL_Light_Alley02_Parts, LM_IL_Light_Alley02_Playfield, LM_IL_Light_Alley03_Driller_Ult, LM_IL_Light_Alley03_Metal, LM_IL_Light_Alley03_Parts, LM_IL_Light_Alley03_Playfield, LM_IL_Light_Alley04_Driller_Ult, LM_IL_Light_Alley04_Parts, LM_IL_Light_Alley04_Playfield, LM_IL_Light_Alley05_Driller_Ult, LM_IL_Light_Alley05_Funnel, LM_IL_Light_Alley05_Parts, LM_IL_Light_Alley05_Playfield)
Dim BG_All: BG_All=Array(BM_Anemometer, BM_BRing1, BM_BRing2, BM_BRing3, BM_Bluey_d2, BM_Bsocket1, BM_Bsocket2, BM_Bsocket3, BM_DT_66, BM_DT_BLU, BM_DT_TA, BM_DT_XB, BM_Driller_UltraLoPoly, BM_Funnel, BM_Gate_Flap, BM_Gate_Flap_002, BM_GateRT66_Wire, BM_Gate_Alley_Wire, BM_LFlipper, BM_Metal, BM_Parts, BM_Playfield, BM_RFlipper, BM_RoundTgt01, BM_RoundTgt02, BM_RoundTgt03, BM_SW1, BM_SW2, BM_SW3, BM_SW4, BM_SW5, BM_SW6, BM_SW7, BM_Spinner_Alley_Wire, LM_GI_Bumper1Light_Anemometer, LM_GI_Bumper1Light_BRing1, LM_GI_Bumper1Light_BRing2, LM_GI_Bumper1Light_Bluey_d2, LM_GI_Bumper1Light_Bsocket1, LM_GI_Bumper1Light_Bsocket2, LM_GI_Bumper1Light_Bsocket3, LM_GI_Bumper1Light_Driller_Ultr, LM_GI_Bumper1Light_Funnel, LM_GI_Bumper1Light_Metal, LM_GI_Bumper1Light_Parts, LM_GI_Bumper1Light_Playfield, LM_GI_Bumper2Light_Anemometer, LM_GI_Bumper2Light_BRing2, LM_GI_Bumper2Light_BRing3, LM_GI_Bumper2Light_Bluey_d2, LM_GI_Bumper2Light_Bsocket1, LM_GI_Bumper2Light_Bsocket2, LM_GI_Bumper2Light_Bsocket3, _
  LM_GI_Bumper2Light_Driller_Ultr, LM_GI_Bumper2Light_Funnel, LM_GI_Bumper2Light_Metal, LM_GI_Bumper2Light_Parts, LM_GI_Bumper2Light_Playfield, LM_GI_Bumper3Light_BRing1, LM_GI_Bumper3Light_BRing2, LM_GI_Bumper3Light_BRing3, LM_GI_Bumper3Light_Bluey_d2, LM_GI_Bumper3Light_Bsocket1, LM_GI_Bumper3Light_Bsocket2, LM_GI_Bumper3Light_Bsocket3, LM_GI_Bumper3Light_DT_TA, LM_GI_Bumper3Light_DT_XB, LM_GI_Bumper3Light_Driller_Ultr, LM_GI_Bumper3Light_Funnel, LM_GI_Bumper3Light_Metal, LM_GI_Bumper3Light_Parts, LM_GI_Bumper3Light_Playfield, LM_GI_Bumper3Light_RoundTgt01, LM_GI_Bumper3Light_RoundTgt02, LM_GI_Bumper3Light_RoundTgt03, LM_GI_GI_LeftSpecial_GateRT66_W, LM_GI_GI_LeftSpecial_Metal, LM_GI_GI_LeftSpecial_Parts, LM_GI_GI_LeftSpecial_Spinner_Al, LM_GI_GI_PumpJack_Bluey_d2, LM_GI_GI_PumpJack_DT_66, LM_GI_GI_PumpJack_DT_BLU, LM_GI_GI_PumpJack_DT_TA, LM_GI_GI_PumpJack_DT_XB, LM_GI_GI_PumpJack_Gate_Flap_002, LM_GI_GI_PumpJack_GateRT66_Wire, LM_GI_GI_PumpJack_Metal, LM_GI_GI_PumpJack_Parts, LM_GI_GI_PumpJack_Playfield, _
  LM_GI_GI_TopApron_L_Anemometer, LM_GI_GI_TopApron_L_Driller_Ult, LM_GI_GI_TopApron_L_Funnel, LM_GI_GI_TopApron_L_Metal, LM_GI_GI_TopApron_L_Parts, LM_GI_GI_TopApron_L_Playfield, LM_GI_GI_TopApron_R_Metal, LM_GI_GI_TopApron_R_Parts, LM_GI_GI_TopApron_R_Playfield, LM_GI_LightWeatherman_Anemomete, LM_GI_LightWeatherman_Driller_U, LM_GI_LightWeatherman_Funnel, LM_GI_LightWeatherman_Gate_Alle, LM_GI_LightWeatherman_Metal, LM_GI_LightWeatherman_Parts, LM_GI_LightWeatherman_Playfield, LM_GI_LightsGI_Pops_BRing1, LM_GI_LightsGI_Pops_BRing2, LM_GI_LightsGI_Pops_Bluey_d2, LM_GI_LightsGI_Pops_Bsocket1, LM_GI_LightsGI_Pops_Bsocket2, LM_GI_LightsGI_Pops_Driller_Ult, LM_GI_LightsGI_Pops_Funnel, LM_GI_LightsGI_Pops_Gate_Alley_, LM_GI_LightsGI_Pops_Metal, LM_GI_LightsGI_Pops_Parts, LM_GI_LightsGI_Pops_Playfield, LM_GI_LightsGI_Pops_RoundTgt01, LM_GI_LightsGI_Pops_RoundTgt02, LM_GI_LightsGI_TARamp_BRing3, LM_GI_LightsGI_TARamp_Bluey_d2, LM_GI_LightsGI_TARamp_Bsocket3, LM_GI_LightsGI_TARamp_Driller_U, _
  LM_GI_LightsGI_TARamp_Funnel, LM_GI_LightsGI_TARamp_Metal, LM_GI_LightsGI_TARamp_Parts, LM_GI_LightsGI_TARamp_Playfield, LM_GI_RedFlasherLight_DT_66, LM_GI_RedFlasherLight_Driller_U, LM_GI_RedFlasherLight_GateRT66_, LM_GI_RedFlasherLight_Metal, LM_GI_RedFlasherLight_Parts, LM_GI_RedFlasherLight_Playfield, LM_GI_RedFlasherLight_RoundTgt0, LM_GI_RedFlasherLight_RoundTgt0, LM_GI_RedFlasherLight_Spinner_A, LM_GI_gi1_Metal, LM_GI_gi1_Parts, LM_GI_gi1_RFlipper, LM_GI_gi2_Metal, LM_GI_gi2_Parts, LM_GI_gi2_SW4, LM_GI_gi3_GateRT66_Wire, LM_GI_gi3_LFlipper, LM_GI_gi3_Metal, LM_GI_gi3_Parts, LM_GI_gi4_DT_66, LM_GI_gi4_GateRT66_Wire, LM_GI_gi4_Metal, LM_GI_gi4_Parts, LM_GI_gi4_SW6, LM_IL_LightBLU_Bluey_d2, LM_IL_LightBLU_DT_BLU, LM_IL_LightBLU_Driller_UltraLoP, LM_IL_LightBLU_Metal, LM_IL_LightBLU_Parts, LM_IL_LightBLU_Playfield, LM_IL_LightBLU_Arrow_Metal, LM_IL_LightBLU_Arrow_Parts, LM_IL_LightBLU_Arrow_Playfield, LM_IL_LightBonus_DT_66, LM_IL_LightBonus_GateRT66_Wire, LM_IL_LightBonus_Metal, LM_IL_LightBonus_Parts, _
  LM_IL_LightBonus_Playfield, LM_IL_LightBonus_Spinner_Alley_, LM_IL_LightDLR01_DT_66, LM_IL_LightDLR01_Driller_UltraL, LM_IL_LightDLR01_Metal, LM_IL_LightDLR01_Parts, LM_IL_LightDLR01_Playfield, LM_IL_LightDLR01_RoundTgt01, LM_IL_LightDLR01_RoundTgt02, LM_IL_LightDLR01_RoundTgt03, LM_IL_LightDLR02_DT_XB, LM_IL_LightDLR02_Driller_UltraL, LM_IL_LightDLR02_Metal, LM_IL_LightDLR02_Parts, LM_IL_LightDLR02_Playfield, LM_IL_LightDLR02_RoundTgt01, LM_IL_LightDLR02_RoundTgt02, LM_IL_LightDLR02_RoundTgt03, LM_IL_LightDLR03_BRing2, LM_IL_LightDLR03_Bsocket2, LM_IL_LightDLR03_DT_XB, LM_IL_LightDLR03_Driller_UltraL, LM_IL_LightDLR03_Metal, LM_IL_LightDLR03_Parts, LM_IL_LightDLR03_Playfield, LM_IL_LightDLR03_RoundTgt01, LM_IL_LightDLR03_RoundTgt02, LM_IL_LightDLR03_RoundTgt03, LM_IL_LightDLR_Arrow_DT_66, LM_IL_LightDLR_Arrow_Metal, LM_IL_LightDLR_Arrow_Parts, LM_IL_LightDLR_Arrow_Playfield, LM_IL_LightE_Metal, LM_IL_LightE_Parts, LM_IL_LightE_Playfield, LM_IL_LightEOB1_DT_66, LM_IL_LightEOB1_DT_BLU, LM_IL_LightEOB1_Metal, _
  LM_IL_LightEOB1_Parts, LM_IL_LightEOB1_Playfield, LM_IL_LightEOB2_DT_BLU, LM_IL_LightEOB2_DT_TA, LM_IL_LightEOB2_Driller_UltraLo, LM_IL_LightEOB2_Metal, LM_IL_LightEOB2_Parts, LM_IL_LightEOB2_Playfield, LM_IL_LightEOB3_Bsocket2, LM_IL_LightEOB3_DT_BLU, LM_IL_LightEOB3_DT_TA, LM_IL_LightEOB3_DT_XB, LM_IL_LightEOB3_Driller_UltraLo, LM_IL_LightEOB3_Metal, LM_IL_LightEOB3_Parts, LM_IL_LightEOB3_Playfield, LM_IL_LightEOB4_BRing2, LM_IL_LightEOB4_Bsocket2, LM_IL_LightEOB4_DT_TA, LM_IL_LightEOB4_DT_XB, LM_IL_LightEOB4_Driller_UltraLo, LM_IL_LightEOB4_Metal, LM_IL_LightEOB4_Parts, LM_IL_LightEOB4_Playfield, LM_IL_LightEOB4_RoundTgt02, LM_IL_LightEOB4_RoundTgt03, LM_IL_LightI_Metal, LM_IL_LightI_Parts, LM_IL_LightI_Playfield, LM_IL_LightInlane_L_DT_66, LM_IL_LightInlane_L_GateRT66_Wi, LM_IL_LightInlane_L_Metal, LM_IL_LightInlane_L_Parts, LM_IL_LightInlane_L_Playfield, LM_IL_LightInlane_L_SW6, LM_IL_LightInlane_L_Spinner_All, LM_IL_LightInlane_R_Metal, LM_IL_LightInlane_R_Parts, LM_IL_LightInlane_R_Playfield, _
  LM_IL_LightInlane_R_SW4, LM_IL_LightK_Metal, LM_IL_LightK_Parts, LM_IL_LightK_Playfield, LM_IL_LightO_Metal, LM_IL_LightO_Parts, LM_IL_LightO_Playfield, LM_IL_LightOutlane_L_GateRT66_W, LM_IL_LightOutlane_L_Metal, LM_IL_LightOutlane_L_Parts, LM_IL_LightOutlane_L_Playfield, LM_IL_LightOutlane_L_SW6, LM_IL_LightOutlane_L_SW7, LM_IL_LightOutlane_R_Metal, LM_IL_LightOutlane_R_Parts, LM_IL_LightOutlane_R_Playfield, LM_IL_LightOutlane_R_SW4, LM_IL_LightOutlane_R_SW5, LM_IL_LightPJ_DT_BLU, LM_IL_LightPJ_Gate_Flap_002, LM_IL_LightPJ_Metal, LM_IL_LightPJ_Parts, LM_IL_LightPJ_Playfield, LM_IL_LightRT66_Arrow_DT_66, LM_IL_LightRT66_Arrow_Metal, LM_IL_LightRT66_Arrow_Parts, LM_IL_LightRT66_Arrow_Playfield, LM_IL_LightRT66_Arrow_Spinner_A, LM_IL_LightS_Metal, LM_IL_LightS_Parts, LM_IL_LightS_Playfield, LM_IL_LightShootAgain_LFlipper, LM_IL_LightShootAgain_Metal, LM_IL_LightShootAgain_Parts, LM_IL_LightShootAgain_Playfield, LM_IL_LightShootAgain_RFlipper, LM_IL_LightTA_Arrow_Bluey_d2, LM_IL_LightTA_Arrow_DT_BLU, _
  LM_IL_LightTA_Arrow_DT_TA, LM_IL_LightTA_Arrow_Driller_Ult, LM_IL_LightTA_Arrow_Metal, LM_IL_LightTA_Arrow_Parts, LM_IL_LightTA_Arrow_Playfield, LM_IL_LightToplane1_Anemometer, LM_IL_LightToplane1_BRing1, LM_IL_LightToplane1_BRing3, LM_IL_LightToplane1_Bsocket1, LM_IL_LightToplane1_Bsocket3, LM_IL_LightToplane1_Driller_Ult, LM_IL_LightToplane1_Funnel, LM_IL_LightToplane1_Metal, LM_IL_LightToplane1_Parts, LM_IL_LightToplane1_Playfield, LM_IL_LightToplane1_SW1, LM_IL_LightToplane1_SW2, LM_IL_LightToplane1_SW3, LM_IL_LightToplane2_BRing1, LM_IL_LightToplane2_BRing3, LM_IL_LightToplane2_Bsocket1, LM_IL_LightToplane2_Bsocket2, LM_IL_LightToplane2_Bsocket3, LM_IL_LightToplane2_Driller_Ult, LM_IL_LightToplane2_Funnel, LM_IL_LightToplane2_Metal, LM_IL_LightToplane2_Parts, LM_IL_LightToplane2_Playfield, LM_IL_LightToplane2_SW2, LM_IL_LightToplane3_Bsocket3, LM_IL_LightToplane3_Funnel, LM_IL_LightToplane3_Metal, LM_IL_LightToplane3_Parts, LM_IL_LightToplane3_Playfield, LM_IL_LightXtraBall_Arrow_BRing, _
  LM_IL_LightXtraBall_Arrow_Bluey, LM_IL_LightXtraBall_Arrow_DT_TA, LM_IL_LightXtraBall_Arrow_DT_XB, LM_IL_LightXtraBall_Arrow_Drill, LM_IL_LightXtraBall_Arrow_Metal, LM_IL_LightXtraBall_Arrow_Parts, LM_IL_LightXtraBall_Arrow_Playf, LM_IL_LightXtraBall_Arrow_Round, LM_IL_LightXtraBall_Arrow_Round, LM_IL_Light_Alley01_Driller_Ult, LM_IL_Light_Alley01_Parts, LM_IL_Light_Alley01_Playfield, LM_IL_Light_Alley01_Spinner_All, LM_IL_Light_Alley02_Driller_Ult, LM_IL_Light_Alley02_Metal, LM_IL_Light_Alley02_Parts, LM_IL_Light_Alley02_Playfield, LM_IL_Light_Alley03_Driller_Ult, LM_IL_Light_Alley03_Metal, LM_IL_Light_Alley03_Parts, LM_IL_Light_Alley03_Playfield, LM_IL_Light_Alley04_Driller_Ult, LM_IL_Light_Alley04_Parts, LM_IL_Light_Alley04_Playfield, LM_IL_Light_Alley05_Driller_Ult, LM_IL_Light_Alley05_Funnel, LM_IL_Light_Alley05_Parts, LM_IL_Light_Alley05_Playfield)
' VLM  Arrays - End



'******** CUSTOM STUFF>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'******** VARIABLES

TimerInit.Enabled = True

dim il

  For each il in InsertLights : il.State = 1 : Next



Sub TimerInit_timer

  dim il

  For each il in InsertLights : il.State = 0 : Next

  RedFlasherLight.State = 0

  TimerInit.Enabled = False

end Sub

Sub ResetBonuses
  EOBBonusReset
  RedSpecialEnd
  BumperMultReset
  Bonus2x = 1
end Sub



'******** BUMPERS!!
Dim B1Step, B2Step, B3Step


Sub Bumper1_Animate
  Dim z, BP
  z = Bumper1.CurrentRingOffset
  For Each BP in BP_BRing1 : BP.transz = z: Next
  B1Step = 0 : Bumper1Light.State = 0 : Bumper1.TimerEnabled = 1
End Sub

Sub Bumper2_Animate
  Dim z, BP
  z = Bumper2.CurrentRingOffset
  For Each BP in BP_BRing2 : BP.transz = z: Next
  B2Step = 0 : Bumper3Light.State = 0 : Bumper2.TimerEnabled = 1
End Sub

Sub Bumper3_Animate
  Dim z, BP
  z = Bumper3.CurrentRingOffset
  For Each BP in BP_BRing3 : BP.transz = z: Next
  B3Step = 0 : Bumper2Light.State = 0 : Bumper3.TimerEnabled = 1
End Sub

Sub Bumper1_Timer
  Select Case B1Step
    Case 2: Bumper1.TimerEnabled = 0: Bumper1Light.State = 1
  End Select
  B1Step = B1Step + 1
End Sub

Sub Bumper2_Timer
  Select Case B2Step
    Case 2: Bumper2.TimerEnabled = 0: Bumper3Light.State = 1
  End Select
  B2Step = B2Step + 1
End Sub

Sub Bumper3_Timer
  Select Case B3Step
    Case 2: Bumper3.TimerEnabled = 0: Bumper2Light.State = 1
  End Select
  B3Step = B3Step + 1
End Sub

Sub Bumper1_Hit
  If Tilted = False Then
    TgtScore = TgtScore + (10 * BumperBonus)
    RandomSoundBumperTop Bumper1
  End If

end Sub

Sub Bumper2_Hit
  If Tilted = False Then
    TgtScore = TgtScore + (10 * BumperBonus)
    RandomSoundBumperMiddle Bumper2
  End if

end Sub

Sub Bumper3_Hit
  If Tilted = False Then
    TgtScore = TgtScore + (10 * BumperBonus)
    RandomSoundBumperBottom Bumper3
  End if

end Sub

'******** BUMPERS MULTIPLIERS!!
Sub Trigger001_Hit
  BumperBonus = 3
  LightToplane1.State = 1
  LightToplane2.State = 0
  LightToplane3.State = 0
end Sub

Sub Trigger002_Hit
  BumperBonus = 5
  LightToplane1.State = 0
  LightToplane2.State = 1
  LightToplane3.State = 0
end Sub

Sub Trigger003_Hit
  BumperBonus = 2
  LightToplane1.State = 0
  LightToplane2.State = 0
  LightToplane3.State = 1
end Sub

Sub BumperMultReset
  BumperBonus = 1
  LightToplane1.State = 0
  LightToplane2.State = 0
  LightToplane3.State = 0
end Sub


'******** Targets
Function DLRtgtCheck()
  if LightDLR01.State = 1 And LightDLR02.State = 1 And LightDLR03.State = 1 Then
  LightI.State = 1
  OKIEScheck()
  TimerDLRreset.enabled=True
  End If
End Function

Function OKIEScheck()
  if LightO.State = 1 And LightK.State = 1 And LightI.State = 1 And LightE.State = 1 And LightS.State = 1 Then
  TgtScore = TgtScore + 155
  OKIESwinLS
  end If
End Function

Sub OKIESwinEnd()
  LSwhale.UpdateInterval = 22
  LSwhale.Play SeqBlinking, ,3,85
  whaletimer.enabled=true
  TimerOKIESreset.enabled=True
End Sub

Function TargetRESET()
  TargetBLU.isDropped = False
  TargetRT66.isDropped = False
  TargetTA.isDropped = False
End Function

Sub TargetXtraRESET
  TargetXtraBall.isDropped = False
End Sub

Sub TargetBLU_Hit
  TgtScore = TgtScore + 150
  DropTargetSound()
  LightBLU_Arrow.State = 2
End Sub

Sub TargetTA_Hit
  TgtScore = TgtScore + 150
  DropTargetSound()
  LightTA_Arrow.State = 2
End Sub

Sub TargetRT66_Hit
  TgtScore = TgtScore + 150
  DropTargetSound()
  LightDLR_Arrow.State = 2
End Sub

Sub TargetXtraBall_Hit
  PlaySound "ricochet", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
  TgtScore = TgtScore + 150
  DropTargetSound()
  BIP = BIP + 1
  DisplayBIP
  RedSpecialEnd
end Sub

Sub Kicker_Whale_hit
  PlaySound "gulp", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
  SoundSaucerLock
  TgtScore = TgtScore + 150
  Kicker_Whale.enabled=false
  WhaleBallLocked = True
  whaleEndballTimer.enabled=true
  LightBLU.State = 1
  LightBLU_Arrow.State = 0
  LightS.State = 1

  LSwhale.UpdateInterval = 22
  LSwhale.Play SeqBlinking, ,3,85

  OKIEScheck()
end sub

Sub whaleEndballTimer_timer
  BallRelease.CreateBall
  BallRelease.Kick 90, 7
  RandomSoundBallRelease BallRelease
  whaleEndballTimer.enabled=false
end sub


Sub GateShootAgain_Hit
  if PumpJackPlay = 0 Then
    PumpJack.PlayAnimEndless(2)
    PumpJackPlay = 1
  Else
    PumpJack.ContinueAnim(2)
  End If
  PlaySound "PumpJackFoley", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
  LSpumpjack.UpdateInterval = 22
  LSpumpjack.Play SeqBlinking, ,6,45
  PumpJackTimer.Enabled = True
  LightPJ.State = 1
  ResetBonuses
End Sub



'******** ALLEY LIGHTS
Sub Spinner_Alley_Spin
  If Tilted = False Then
    TgtScore = TgtScore + 1

    if Light_Alley01.State = 1 Then
      If B2SOn Then
        Controller.B2SSetData 15,0
        Controller.B2SSetData 16,0
      End If
      Light_Alley01.State = 0
    Else
      Light_Alley01.State = 1
      If B2SOn Then
        Controller.B2SSetData 15,1
        Controller.B2SSetData 16,1
      End If

    end if
    if Light_Alley02.State = 1 Then
      Light_Alley02.State = 0
    Else
      Light_Alley02.State = 1

    end if
    if Light_Alley03.State = 1 Then
      Light_Alley03.State = 0
    Else
      Light_Alley03.State = 1

    end if
    if Light_Alley04.State = 1 Then
      Light_Alley04.State = 0
    Else
      Light_Alley04.State = 1

    end if
    if Light_Alley05.State = 1 Then
      Light_Alley05.State = 0
    Else
      Light_Alley05.State = 1
    end if
    AlleyLightTimer.Enabled=True
  End if

end Sub


Sub SpinnerTA_Spin
  If Tilted = False Then
    BackglassFlash.enabled = True
    StormTimer.enabled = True
    TgtScore = TgtScore + 3
    LightE.State = 1
    LightTA_Arrow.State = 0
    OKIEScheck()
    LightWeatherman.State = 2
    LightsGI_TARamp.State = 2
    LightsGI_TopApron_L.State = 2
    LightsGI_TopApron_R.State = 2
    Bumper1Light.State = 2
    Bumper2Light.State = 2
    Bumper3Light.State = 2
    TimerTASpin.Enabled = True
    TornadoSounds
  End if
end Sub

Sub Gate_Alley_Hit
  PlaySound "racecar", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
  TgtScore = TgtScore + 150
  LightO.State = 1
  LightRT66_Arrow.State = 0

  OKIEScheck()
end Sub

Sub GateEntry_Hit
  StartBallSound
  LightShootAgain.State = 1
  if LightO.State = 0 Then
    LightRT66_Arrow.State = 2
  End If
  TimerShootAgain.enabled=True
end Sub

Sub GatePops_Hit
  PlaySound "GetYourKicks", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
  TgtScore = TgtScore + 150
  LightsGI_POPS.State = 2
  PopsLightTimer.enabled = True
end Sub


'******** RED SPECIAL!!

Sub TargetBonus_hit
  If RedSpecial = False Then
    RedSpecial = True
    LightBonus.State=1
    RedSpecialLit
    RedSpecialSound
    RedSpecialTimer.enabled = True
  End If
  RandomSoundBumper()

end Sub

Sub RedSpecialLit
  RedFlasherLight.State = 2
  LightXtraBall_Arrow.State = 2
  LightInlane_L.State = 1
  LightInlane_R.State = 1
  LightOutlane_L.State = 1
  LightOutlane_R.State = 1
  TargetXtraRESET
end Sub

Sub RedSpecialEnd
  StopSound "Stampede"
  RedSpecial = False
  Bonus2x = 1
  RedFlasherLight.State = 0
  LightXtraBall_Arrow.State = 0
  LightBonus.State = 0
  LightInlane_L.State = 0
  LightInlane_R.State = 0
  LightOutlane_L.State = 0
  LightOutlane_R.State = 0
  TargetXtraBall.isDropped = True
end Sub

Sub LeftInlane_Hit
  If Tilted = False Then
    If RedSpecial = True Then
      TgtScore = TgtScore + 75
    End If
    RedSpecialEnd
    EOBBonusAdvance
  End if
end Sub

Sub RightInlane_Hit
  If Tilted = False Then
    If RedSpecial = True Then
      TgtScore = TgtScore + 75
    End If
    RedSpecialEnd
    EOBBonusAdvance
  End if

end Sub

Sub LeftOutlane_Hit
  If Tilted = False Then
    PlaySound "moo01", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
    If RedSpecial = True Then
      Bonus2x = 2
    End If
    RedSpecialEnd
  End if
end Sub

Sub RightOutlane_Hit
  If Tilted = False Then
    PlaySound "moo02", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
    If RedSpecial = True Then
      Bonus2x = 2
    End If
    RedSpecialEnd
  End if
end Sub


'******** TIMERS

Sub AnemometerSpinTimer_timer
  Dim az : az = SpinnerTA.CurrentAngle
  Anemometer.rotZ = az
  Dim BP: For each BP in BP_Anemometer: BP.rotz = az: Next
end Sub


Sub AlleyLightTimer_timer
  Light_Alley01.State = 0
  Light_Alley02.State = 0
  Light_Alley03.State = 0
  Light_Alley04.State = 0
  Light_Alley05.State = 0
  AlleyLightTimer.enabled=false
end Sub

Sub TimerDLRreset_timer
  LightDLR01.State = 0
  LightDLR02.State = 0
  LightDLR03.State = 0
  TimerDLRreset.enabled=False
end Sub

Sub TimerOKIESreset_timer
  LightO.State=0
  LightK.State=0
  LightI.State=0
  LightE.State=0
  LightS.State=0
  LightRT66_Arrow.State = 0
  LightBLU_Arrow.State = 0
  LightDLR_Arrow.State = 0
  LightTA_Arrow.State = 0
  TimerTargetReset.enabled = True
  TimerOKIESreset.enabled=False
end Sub

Sub whaletimer_timer
  Kicker_Whale.kick 180,10
  SoundSaucerKick 1, Kicker_Whale
  LightBLU.State = 0
  Kicker_Whale.enabled=true
  WhaleBallLocked = False
  WhaleBIP = True
  BIP = BIP + 1
  DisplayBIP
  whaletimer.enabled=false
end sub

Sub TimerTargetReset_timer
  TargetRESET()
  TimerTargetReset.enabled = False
end Sub

Sub TimerTASpin_timer
  LightWeatherman.State = 1
  LightsGI_TARamp.State = 1
  LightsGI_TopApron_L.State = 1
  LightsGI_TopApron_R.State = 1
  Bumper1Light.State = 1
  Bumper2Light.State = 1
  Bumper3Light.State = 1
  If B2SOn = True Then Controller.B2SSetData 11,1
  TgtScore = TgtScore + 150
  TimerTASpin.enabled = False
end Sub

Sub TimerShootAgain_timer
  LightShootAgain.State=0
  TimerShootAgain.enabled=False
end Sub

Sub PumpJackTimer_timer
  PumpJack.StopAnim()
  PumpJackPlay = 0
  LightPJ.State = 0
  PumpJackTimer.enabled=False
end Sub

Sub PopsLightTimer_timer
  LightsGI_POPS.State = 0
  PopsLightTimer.enabled = False
end Sub

Sub RedSpecialTimer_timer
  RedSpecialEnd
  RedSpecialTimer.enabled = False
end Sub

Dim TAglass

Sub BackglassFlash_timer
  If TAglass = 0 Then
    TAglass = 1
    If B2SOn = True Then Controller.B2SSetData 11,1
  Else
    TAglass = 0
    If B2SOn = True Then Controller.B2SSetData 11,0
  End If
end Sub



Sub StormTimer_timer
  StormTimer.enabled= False
  BackglassFlash.enabled = False
end Sub


'******** DRILLER Targets

Sub TargetDLR01_Hit
  If Tilted = False Then
    LightDLR01.State = 1
    TgtScore = TgtScore + 50
    DLRtgtCheck()
  End if
end Sub

Sub TargetDLR02_Hit
  If Tilted = False Then
    LightDLR02.State = 1
    TgtScore = TgtScore + 50
    DLRtgtCheck()
  End if
end Sub

Sub TargetDLR03_Hit
  If Tilted = False Then
    LightDLR03.State = 1
    TgtScore = TgtScore + 50
    DLRtgtCheck()
  End If
end Sub


Sub TargetDLR01_Animate
  Dim x : x = TargetDLR01.CurrentAnimOffset
  Dim BP : For Each BP in BP_RoundTgt01 : BP.rotx = x : Next
end Sub

Sub TargetDLR02_Animate
  Dim x : x = TargetDLR02.CurrentAnimOffset
  Dim BP : For Each BP in BP_RoundTgt02 : BP.rotx = x : Next
end Sub

Sub TargetDLR03_Animate
  Dim x : x = TargetDLR03.CurrentAnimOffset
  Dim BP : For Each BP in BP_RoundTgt03 : BP.rotx = x : Next
end Sub



Sub GateRT66_Hit
  LightK.State = 1
  LightDLR_Arrow.State = 0
  OKIEScheck()
end Sub




'******** DROP Targets
Sub TargetRT66_Animate
  Dim z : z = TargetRT66.CurrentAnimOffset
  Dim BP : For Each BP in BP_DT_66 : BP.transz = z: Next
End Sub

Sub TargetBLU_Animate
  Dim z : z = TargetBLU.CurrentAnimOffset
  Dim BP : For Each BP in BP_DT_BLU : BP.transz = z: Next
End Sub

Sub TargetTA_Animate
  Dim z : z = TargetTA.CurrentAnimOffset
  Dim BP : For Each BP in BP_DT_TA : BP.transz = z: Next
End Sub

Sub TargetXtraBall_Animate
  Dim z : z = TargetXtraBall.CurrentAnimOffset
  Dim BP : For Each BP in BP_DT_XB : BP.transz = z: Next
End Sub


Sub GateShootAgain_Animate
  Dim a : a = (GateShootAgain.CurrentAngle * -1) + 25
  Dim BP : For Each BP in BP_Gate_Flap_002 : BP.rotx = a: Next
End Sub


Sub Spinner_Alley_Animate
  Dim a : a = Spinner_Alley.CurrentAngle
  Dim BP : For Each BP in BP_Spinner_Alley_Wire: BP.rotx = a: Next

End Sub

'************ Switches

Sub Trigger001_Animate
  Dim z : z = Trigger001.CurrentAnimOffset
  Dim BP : For Each BP in BP_SW1 : BP.transx = -z: Next
End Sub

Sub Trigger002_Animate
  Dim z : z = Trigger002.CurrentAnimOffset
  Dim BP : For Each BP in BP_SW2 : BP.transx = -z: Next
End Sub

Sub Trigger003_Animate
  Dim z : z = Trigger003.CurrentAnimOffset
  Dim BP : For Each BP in BP_SW3 : BP.transx = -z: Next
End Sub

Sub LeftInlane_Animate
  Dim z : z = LeftInlane.CurrentAnimOffset
  Dim BP : For Each BP in BP_SW6 : BP.transx = -z: Next
End Sub

Sub LeftOutlane_Animate
  Dim z : z = LeftOutlane.CurrentAnimOffset
  Dim BP : For Each BP in BP_SW7 : BP.transx = -z: Next
End Sub

Sub RightInlane_Animate
  Dim z : z = RightInlane.CurrentAnimOffset
  Dim BP : For Each BP in BP_SW4 : BP.transx = -z: Next
End Sub

Sub RightOutlane_Animate
  Dim z : z = RightOutlane.CurrentAnimOffset
  Dim BP : For Each BP in BP_SW5 : BP.transx = -z: Next
End Sub


'****************************
'******** SCORING!!
'****************************

Sub ScoreTimer_timer
  If CurrScore + 100 <= TgtScore Then
      ScoreChime100(Spinner_Alley)
      CurrScore = CurrScore + 100
  Elseif CurrScore + 10 <= TgtScore Then
      ScoreChime100(Spinner_Alley)
      CurrScore = CurrScore + 10
  Elseif CurrScore + 1 <= TgtScore Then
      ScoreChime10(Spinner_Alley)
      CurrScore = CurrScore + 1
  end if

  If B2SOn Then
    Controller.B2SSetScorePlayer 1, CurrScore
  End If

  If CurrScore >= 10000 Then  '*Adding stars**
    If B2SOn Then Controller.B2SSetData 31,1
    If cab_mode = 0 Then lt_Star1.state = 1
    Else
    lt_Star1.state = 0
  End If
  If CurrScore >= 20000 Then  '*Adding stars**
    If B2SOn Then Controller.B2SSetData 32,1
    If cab_mode = 0 Then lt_Star2.state = 1
    Else
    lt_Star2.state = 0
  End If
  If CurrScore >= 30000 Then  '*Adding stars**
    If B2SOn Then Controller.B2SSetData 34,1
    If cab_mode = 0 Then lt_Star3.state = 1
    Else
    lt_Star3.state = 0
  End If
    If CurrScore >= 40000 Then  '*Adding stars**
    If B2SOn Then Controller.B2SSetData 35,1
    If cab_mode = 0 Then lt_Star4.state = 1
    Else
    lt_Star4.state = 0
  End If

  EMReel_Score.SetValue CurrScore

  If CurrScore > HiScore Then
      HiScore = CurrScore
  End If
  EMReel_HiScore.SetValue HiScore
  If B2SOn Then
    Controller.B2SSetScorePlayer 2, HiScore
  End If

  If HiScore >= 10000 Then  '*Adding stars**
    If B2SOn Then Controller.B2SSetData 41,1
    If cab_mode = 0 Then lt_StarHi1.state = 1
    Else
    lt_StarHi1.state = 0
  End If
  If HiScore >= 20000 Then  '*Adding stars**
    If B2SOn Then Controller.B2SSetData 42,1
    If cab_mode = 0 Then lt_StarHi2.state = 1
    Else
    lt_StarHi2.state = 0
  End If
  If HiScore >= 30000 Then  '*Adding stars**
    If B2SOn Then Controller.B2SSetData 43,1
    If cab_mode = 0 Then lt_StarHi3.state = 1
    Else
    lt_StarHi3.state = 0
  End If
    If HiScore >= 40000 Then  '*Adding stars**
    If B2SOn Then Controller.B2SSetData 44,1
    If cab_mode = 0 Then lt_StarHi4.state = 1
    Else
    lt_StarHi4.state = 0
  End If

end Sub

Sub ResetScore
  TgtScore = 0
  CurrScore = 0
  If B2SOn Then
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScorePlayer 1, 0
    Controller.B2SSetData 31,0
    Controller.B2SSetData 32,0
    Controller.B2SSetData 33,0
  End If
End Sub


Sub EOBBonusCollect
  Select Case AdvBonus
    Case 1: TgtScore = TgtScore + (100 * Bonus2x)
    Case 2: TgtScore = TgtScore + (200 * Bonus2x)
    Case 3: TgtScore = TgtScore + (500 * Bonus2x)
    Case 4: TgtScore = TgtScore + (1000 * Bonus2x)
  End Select
End Sub

Sub EOBBonusAdvance
  AdvBonus = AdvBonus + 1

  Select Case AdvBonus
    Case 1: LightEOB1.State = 1
    Case 2: LightEOB2.State = 1
    Case 3: LightEOB3.State = 1
    Case 4: LightEOB4.State = 1
    Case 5: AdvBonus = 4
  end Select
end sub

Sub EOBBonusReset
  AdvBonus = 0
  LightEOB1.State = 0
  LightEOB2.State = 0
  LightEOB3.State = 0
  LightEOB4.State = 0
end Sub

'*****************
'*** LOAD/SAVE HIGHSCORE
'***********************

sub savehs
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HSFileName,True)
    ScoreFile.WriteLine 0
    ScoreFile.WriteLine HiScore
    ScoreFile.WriteLine Credit
    ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub

sub loadhs
  Dim FileObj
  Dim ScoreFile
    dim temp1
    dim temp2
  dim temp3
  Dim TextStr

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & HSFileName) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & HSFileName)
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    temp1=TextStr.ReadLine
    temp2=textstr.readline
    temp3=textstr.readline

    TextStr.Close
      HiScore = cdbl(temp2)
    Credit = cdbl(temp3)

    Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub


'********************
'********* LIGHT SEQUENCERS

Sub AttractMode()
    LSattract.UpdateInterval = 12
    LSattract.Play SeqUpOn, 30, 1
    LSattract.UpdateInterval = 12
    LSattract.Play SeqDownOn, 30, 1
    LSattract.UpdateInterval = 12
    LSattract.Play SeqRightOn, 30, 1
    LSattract.UpdateInterval = 12
    LSattract.Play SeqLeftOn, 30, 1
    LSattract.UpdateInterval = 12
    LSattract.Play SeqCircleOutOn, 30, 1
    LSattract.UpdateInterval = 12
    LSattract.Play SeqCircleInOn, 30, 1
    LSattract.UpdateInterval = 12
    LSattract.Play SeqHatch1VertOn, 30, 1
    LSattract.UpdateInterval = 12
    LSattract.Play SeqHatch2VertOn, 30, 1
End Sub



Sub OKIESwinLS()
  PlaySound "YourDoinFineOklahoma", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
  TgtScore = TgtScore + 1234
  LS_OKIES.Play SeqRandom, 1,,4000
  LS_OKIES.UpdateInterval = 12
  LS_OKIES.Play SeqBlinking, ,6,85
  LS_Arrows.UpdateInterval = 24
  LS_Arrows.Play SeqBlinking, ,6,45
  LS_Arrows.UpdateInterval = 42
  LS_Arrows.Play SeqBlinking, ,6,45
  LS_Arrows.UpdateInterval = 8
  LS_Arrows.Play SeqCircleOutOn, 30, 1
  LS_Arrows.UpdateInterval = 8
  LS_Arrows.Play SeqCircleInOn, 30, 1
  LS_Arrows.UpdateInterval = 8
  LS_Arrows.Play SeqCircleOutOn, 30, 1
  LS_Arrows.UpdateInterval = 8
  LS_Arrows.Play SeqCircleInOn, 30, 1

End Sub

Sub LS_OKIES_PlayDone()
  OKIESwinEnd
End Sub

'*****GI Lights On
dim lxx

For each lxx in GI:lxx.State = 1: Next


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
'    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
    RandomSoundSlingshotRight Sling1
  RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0:gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
'    PlaySound SoundFX("left_slingshot",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
  RandomSoundSlingshotLeft Sling2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  gi3.State = 0:Gi4.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0:gi3.State = 1:Gi4.State = 1
    End Select
    LStep = LStep + 1
End Sub



'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

'' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
'Sub PlaySoundAt(soundname, tableobj)
'    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
'End Sub

'Sub PlaySoundAtBall(soundname)
'    PlaySoundAt soundname, ActiveBall
'End Sub




'******************************************************
'   BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

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
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    '***Ball Drop Sounds***
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz > -7 Then
          RandomSoundBallBouncePlayfieldSoft BOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard BOT(b)
        End If
      End If
    End If
    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If
  Next
End Sub




'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'     * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'     * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'     * Create a Timer called RampRoll, that is enabled, with a interval of 100
'     * Set RampBalls and RampType variable to Total Number of Balls
' Usage:
'     * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'     * To stop tracking ball
'        * call WireRampOff
'        * Otherwise, the ball will auto remove if it's below 30 vp units
'

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
' Setup:  Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RampBalls(6,2)
Dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(6)

Sub WireRampOn(input)
  Waddball ActiveBall, input
  RampRollUpdate
End Sub

Sub WireRampOff()
  WRemoveBall ActiveBall.ID
End Sub

' WaddBall (Active Ball, Boolean)
Sub Waddball(input, RampInput) 'This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  Dim x
  For x = 1 To UBound(RampBalls)  'Check, don't add balls twice
    If RampBalls(x, 1) = input.id Then
      If Not IsEmpty(RampBalls(x,1) ) Then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 To UBound(RampBalls)
    If IsEmpty(RampBalls(x, 1)) Then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      Exit Sub
    End If
    If x = UBound(RampBalls) Then  'debug
      Debug.print "WireRampOn error, ball queue Is full: " & vbNewLine & _
      RampBalls(0, 0) & vbNewLine & _
      TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbNewLine & _
      TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbNewLine & _
      TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbNewLine & _
      TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbNewLine & _
      TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbNewLine & _
      " "
    End If
  Next
End Sub

' WRemoveBall (BallId)
Sub WRemoveBall(ID) 'This subroutine is called from the RampRollUpdate subroutine and is used to remove and stop the ball rolling sounds
  '   Debug.Print "In WRemoveBall() + Remove ball from loop array"
  Dim ballcount
  ballcount = 0
  Dim x
  For x = 1 To UBound(RampBalls)
    If ID = RampBalls(x, 1) Then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    If Not IsEmpty(Rampballs(x,1)) Then ballcount = ballcount + 1
  Next
  If BallCount = 0 Then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer()
  RampRollUpdate
End Sub

Sub RampRollUpdate()  'Timer update
  Dim x
  For x = 1 To UBound(RampBalls)
    If Not IsEmpty(RampBalls(x,1) ) Then
      If BallVel(RampBalls(x,0) ) > 1 Then ' if ball is moving, play rolling sound
        If RampType(x) Then
          PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      End If
      If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
  Next
  If Not RampBalls(0,0) Then RampRoll.enabled = 0
End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  Me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbNewLine & _
  "1 " & TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbNewLine & _
  "2 " & TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbNewLine & _
  "3 " & TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbNewLine & _
  "4 " & TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbNewLine & _
  "5 " & TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbNewLine & _
  "6 " & TypeName(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbNewLine & _
  " "
End Sub

Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
  BallPitch = pSlope(BallVel(ball), 1, - 1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, - 4000, 60, 7000)
End Function

Sub PopsRampStart_Hit
  WireRampOn false
End Sub

Sub PopsWireRampStart_Hit
  WireRampOn true
End Sub

Sub PopsWireRampStop_Hit
  WireRampOn false
End Sub

Sub PopsRampStop_Hit
    WireRampOff
End Sub

Sub TARampStart_Hit
  WireRampOn false
End Sub

Sub TAWireRampStart_Hit
  WireRampOn true
End Sub
Sub TAWireRampStop_Hit
  WireRampOn false
End Sub

Sub TARampStop_Hit
    WireRampOff
End Sub


Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************






'*****************************************
' ninuzzu's FLIPPER SHADOWS v3 (VPX 10.8)
'*****************************************

Sub LeftFlipper_Animate()
    FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  Dim BP: For each BP in BP_LFlipper: BP.RotZ = LeftFlipper.CurrentAngle - 90: Next
End Sub

Sub RightFlipper_Animate()
    FlipperRSh.RotZ = RightFlipper.CurrentAngle
  Dim BP: For each BP in BP_RFlipper: BP.RotZ = RightFlipper.CurrentAngle -90: Next
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        End If
    BallShadow(b).X = BOT(b).X + (BOT(b).X - (Table1.Width/2)) * 1.25 / BallSize
        BallShadow(b).Y = BOT(b).Y + 12
    BallShadow(b).Size_X = 5
    BallShadow(b).Size_Y = 5

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




Sub LeftFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub


Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandomSoundBumper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx_bumper1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "fx_bumper2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "fx_bumper3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


'**************************************  CUSTOM SOUNDS
Sub RandSoundChime1k()
  Select Case Int(Rnd*2)+1
    Case 1:
        StopSound "SJ_Chime_1000a"
        PlaySound "SJ_Chime_1000a", 0, chimevol, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2:
        StopSound "SJ_Chime_1000b"
        PlaySound "SJ_Chime_1000b", 0, chimevol, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandSoundChime100()
  Select Case Int(Rnd*2)+1
    Case 1:
        StopSound "SJ_Chime_100a"
        PlaySound "SJ_Chime_100a", 0, chimevol, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2:
        StopSound "SJ_Chime_100b"
        PlaySound "SJ_Chime_100b", 0, chimevol, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub DropTargetSound()
  PlaySound "droptargetdropped", 0, 0.05, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
end Sub

Sub ScoreChime10(obj)
  PlaySound "SJ_Chime_10a" ', 0, chimevol, 0, 0, 0, 1, 0, 0
end Sub

Sub ScoreChime100(obj)
  PlaySound "SJ_Chime_100a" ', 0, chimevol, 0, 0, 0, 1, 0, 0
end Sub

Sub ScoreChime1000(obj)
  PlaySound "SJ_Chime_1000a" ', 0, chimevol, 0, 0, 0, 1, 0, 0
end Sub

Dim StartSoundPlaying

Sub StartBallSound
  If StartSoundPlaying = False Then
    StartSoundPlaying = True
    StartSoundTimer.enabled = True
    Select Case Int(Rnd*7)+1
      Case 1 : PlaySound "OklahomaOklahomaOklahoma", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
      Case 2 : Playsound "OKLAHOMAspell", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
      Case 3 : Playsound "YeeHaw01wav", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
      Case 4 : Playsound "WhereTheWindComesSweeping", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
      Case 5 : Playsound "boomersooner", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
      Case 6 : Playsound "IveBeenToOklahoma", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
      Case 7 : Playsound "OkieFromMuskogee", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
    End Select
  End If
end Sub

Sub StartSoundTimer_timer
  StartSoundPlaying = False
  StartSoundTimer.enabled = False
End Sub

Sub TornadoSounds
    PlaySound "Tornado", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
end Sub

Sub RedSpecialSound
  PlaySound "Stampede", 0, musicvol, AudioPan(speaker), 0, 0, 1, 0, AudioFade(speaker)
end Sub


'**********************************
'   ZMAT: General Math Functions
'**********************************
' These get used throughout the script.

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
  dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
  dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    End If
  ElseIf dx = 0 Then
    If dy = 0 Then
      Atn2 = 0
    Else
      Atn2 = Sgn(dy) * pi / 2
    End If
  End If
End Function

Function ArcCos(x)
  If x = 1 Then
    ArcCos = 0/180*PI
  ElseIf x = -1 Then
    ArcCos = 180/180*PI
  Else
    ArcCos = Atn(-x/Sqr(-x * x + 1)) + 2 * Atn(1)
  End If
End Function

Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

Function min(a,b)
  If a > b Then
    min = b
  Else
    min = a
  End If
End Function

' Used for drop targets
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy) 'Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
  Dim AB, BC, CD, DA
  AB = (bx * py) - (by * px) - (ax * py) + (ay * px) + (ax * by) - (ay * bx)
  BC = (cx * py) - (cy * px) - (bx * py) + (by * px) + (bx * cy) - (by * cx)
  CD = (dx * py) - (dy * px) - (cx * py) + (cy * px) + (cx * dy) - (cy * dx)
  DA = (ax * py) - (ay * px) - (dx * py) + (dy * px) + (dx * ay) - (dy * ax)

  If (AB <= 0 And BC <= 0 And CD <= 0 And DA <= 0) Or (AB >= 0 And BC >= 0 And CD >= 0 And DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
  Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
  Dim rotxy
  rotxy = RotPoint(ax,ay,angle)
  rax = rotxy(0) + px
  ray = rotxy(1) + py
  rotxy = RotPoint(bx,by,angle)
  rbx = rotxy(0) + px
  rby = rotxy(1) + py
  rotxy = RotPoint(cx,cy,angle)
  rcx = rotxy(0) + px
  rcy = rotxy(1) + py
  rotxy = RotPoint(dx,dy,angle)
  rdx = rotxy(0) + px
  rdy = rotxy(1) + py

  InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
  Dim rx, ry
  rx = x * dCos(angle) - y * dSin(angle)
  ry = x * dSin(angle) + y * dCos(angle)
  RotPoint = Array(rx,ry)
End Function



'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. and, special scripting
'
' TriggerLF and RF should now be 27 vp units from the flippers. In addition, 3 degrees should be added to the end angle
' when creating these triggers.
'
' RF.ReProcessBalls Activeball and LF.ReProcessBalls Activeball must be added the flipper_collide subs.
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.4            | 0.4                   | 0.375                  | 0.375              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger


        x.AddPt "Polarity", 0, 0, 0
        x.AddPt "Polarity", 1, 0.05, - 2.7
        x.AddPt "Polarity", 2, 0.16, - 2.7
        x.AddPt "Polarity", 3, 0.22, - 0
        x.AddPt "Polarity", 4, 0.25, - 0
        x.AddPt "Polarity", 5, 0.3, - 1
        x.AddPt "Polarity", 6, 0.4, - 2
        x.AddPt "Polarity", 7, 0.5, - 2.7
        x.AddPt "Polarity", 8, 0.65, - 1.8
        x.AddPt "Polarity", 9, 0.75, - 0.5
        x.AddPt "Polarity", 10, 0.81, - 0.5
        x.AddPt "Polarity", 11, 0.88, 0
        x.AddPt "Polarity", 12, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.15, 0.85
    x.AddPt "Velocity", 2, 0.2, 0.9
    x.AddPt "Velocity", 3, 0.23, 0.95
    x.AddPt "Velocity", 4, 0.41, 0.95
    x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
    x.AddPt "Velocity", 6, 0.62, 1.0
    x.AddPt "Velocity", 7, 0.702, 0.968
    x.AddPt "Velocity", 8, 0.95,  0.968
    x.AddPt "Velocity", 9, 1.03,  0.945
    x.AddPt "Velocity", 10, 1.5,  0.945

  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
    LF.SetObjects "LF", LeftFlipper, TriggerLF
    RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub

'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 80
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 3.7
'   x.AddPt "Polarity", 2, 0.16, - 3.7
'   x.AddPt "Polarity", 3, 0.22, - 0
'   x.AddPt "Polarity", 4, 0.25, - 0
'   x.AddPt "Polarity", 5, 0.3, - 2
'   x.AddPt "Polarity", 6, 0.4, - 3
'   x.AddPt "Polarity", 7, 0.5, - 3.7
'   x.AddPt "Polarity", 8, 0.65, - 2.3
'   x.AddPt "Polarity", 9, 0.75, - 1.5
'   x.AddPt "Polarity", 10, 0.81, - 1
'   x.AddPt "Polarity", 11, 0.88, 0
'   x.AddPt "Polarity", 12, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
''*******************************************
''  Late 80's early 90's
'
'Sub InitPolarity()
' dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 5
'   x.AddPt "Polarity", 2, 0.16, - 5
'   x.AddPt "Polarity", 3, 0.22, - 0
'   x.AddPt "Polarity", 4, 0.25, - 0
'   x.AddPt "Polarity", 5, 0.3, - 2
'   x.AddPt "Polarity", 6, 0.4, - 3
'   x.AddPt "Polarity", 7, 0.5, - 4.0
'   x.AddPt "Polarity", 8, 0.7, - 3.5
'   x.AddPt "Polarity", 9, 0.75, - 3.0
'   x.AddPt "Polarity", 10, 0.8, - 2.5
'   x.AddPt "Polarity", 11, 0.85, - 2.0
'   x.AddPt "Polarity", 12, 0.9, - 1.5
'   x.AddPt "Polarity", 13, 0.95, - 1.0
'   x.AddPt "Polarity", 14, 1, - 0.5
'   x.AddPt "Polarity", 15, 1.1, 0
'   x.AddPt "Polarity", 16, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945

' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

'*******************************************
' Early 90's and after

'Sub InitPolarity()
' Dim x, a
' a = Array(LF, RF)
' For Each x In a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 5.5
'   x.AddPt "Polarity", 2, 0.16, - 5.5
'   x.AddPt "Polarity", 3, 0.20, - 0.75
'   x.AddPt "Polarity", 4, 0.25, - 1.25
'   x.AddPt "Polarity", 5, 0.3, - 1.75
'   x.AddPt "Polarity", 6, 0.4, - 3.5
'   x.AddPt "Polarity", 7, 0.5, - 5.25
'   x.AddPt "Polarity", 8, 0.7, - 4.0
'   x.AddPt "Polarity", 9, 0.75, - 3.5
'   x.AddPt "Polarity", 10, 0.8, - 3.0
'   x.AddPt "Polarity", 11, 0.85, - 2.5
'   x.AddPt "Polarity", 12, 0.9, - 2.0
'   x.AddPt "Polarity", 13, 0.95, - 1.5
'   x.AddPt "Polarity", 14, 1, - 1.0
'   x.AddPt "Polarity", 15, 1.05, -0.5
'   x.AddPt "Polarity", 16, 1.1, 0
'   x.AddPt "Polarity", 17, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.23, 0.85
'   x.AddPt "Velocity", 2, 0.27, 1
'   x.AddPt "Velocity", 3, 0.3, 1
'   x.AddPt "Velocity", 4, 0.35, 1
'   x.AddPt "Velocity", 5, 0.6, 1 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before\
' modified 2024 by rothbauerw
' Added Reprocessballs for flipper collisions (LF.Reprocessballs Activeball and RF.Reprocessballs Activeball must be added to the flipper collide subs
' Improved handling to remove correction for backhand shots when the flipper is raised

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef, FlipStartAngle
  Private Balls(20), balldata(20)
  Private Name

  Dim PolarityIn, PolarityOut
  Dim VelocityIn, VelocityOut
  Dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    ReDim PolarityIn(0)
    ReDim PolarityOut(0)
    ReDim VelocityIn(0)
    ReDim VelocityOut(0)
    ReDim YcoefIn(0)
    ReDim YcoefOut(0)
    Enabled = True
    TimeDelay = 50
    LR = 1
    Dim x
    For x = 0 To UBound(balls)
      balls(x) = Empty
      Set Balldata(x) = new SpoofBall
    Next
  End Sub

  Public Sub SetObjects(aName, aFlipper, aTrigger)

    If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: .SetObjects error: first argument must be a String (And name of Object). Found:" & TypeName(aName) End If
    If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: .SetObjects error: Second argument must be a flipper. Found:" & TypeName(aFlipper) End If
    If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & TypeName(aTrigger) End If
    If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
    Name = aName
    Set Flipper = aFlipper
    FlipperStart = aFlipper.x
    FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
    FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y

    Dim str
    str = "Sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
    ExecuteGlobal(str)
    str = "Sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
    ExecuteGlobal(str)

  End Sub

  ' Legacy: just no op
  Public Property Let EndPoint(aInput)

  End Property

  Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      Case "Polarity"
        ShuffleArrays PolarityIn, PolarityOut, 1
        PolarityIn(aIDX) = aX
        PolarityOut(aIDX) = aY
        ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity"
        ShuffleArrays VelocityIn, VelocityOut, 1
        VelocityIn(aIDX) = aX
        VelocityOut(aIDX) = aY
        ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef"
        ShuffleArrays YcoefIn, YcoefOut, 1
        YcoefIn(aIDX) = aX
        YcoefOut(aIDX) = aY
        ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
  End Sub

  Public Sub AddBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If IsEmpty(balls(x)) Then
        Set balls(x) = aBall
        Exit Sub
      End If
    Next
  End Sub

  Private Sub RemoveBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If TypeName(balls(x) ) = "IBall" Then
        If aBall.ID = Balls(x).ID Then
          balls(x) = Empty
          Balldata(x).Reset
        End If
      End If
    Next
  End Sub

  Public Sub Fire()
    Flipper.RotateToEnd
    processballs
  End Sub

  Public Property Get Pos 'returns % position a ball. For debug stuff.
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x)) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x)) Then
        balldata(x).Data = balls(x)
      End If
    Next
    FlipStartAngle = Flipper.currentangle
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub

  Public Sub ReProcessBalls(aBall) 'save data of balls in flipper range
    If FlipperOn() Then
      Dim x
      For x = 0 To UBound(balls)
        If Not IsEmpty(balls(x)) Then
          if balls(x).ID = aBall.ID Then
            If isempty(balldata(x).ID) Then
              balldata(x).Data = balls(x)
            End If
          End If
        End If
      Next
    End If
  End Sub

  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef, BalltoFlip, BalltoBase, NoCorrection, checkHit
      Ycoef = 1

      'y safety Exit
      If aBall.VelY > -8 Then 'ball going down
        RemoveBall aBall
        Exit Sub
      End If

      'Find balldata. BallPos = % on Flipper
      For x = 0 To UBound(Balls)
        If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          BalltoFlip = DistanceFromFlipperAngle(BallData(x).x, BallData(x).y, Flipper, FlipStartAngle)
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
        NoCorrection = 1
      Else
        checkHit = 50 + (20 * BallPos)

        If BalltoFlip > checkHit or (PartialFlipCoef < 0.5 and BallPos > 0.22) Then
          NoCorrection = 1
        Else
          NoCorrection = 0
        End If
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        'If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled and NoCorrection = 0 Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef*VelCoef)
      End If
      If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  Dim x, aCount
  aCount = 0
  ReDim a(UBound(aArray) )
  For x = 0 To UBound(aArray)   'Shuffle objects in a temp array
    If Not IsEmpty(aArray(x) ) Then
      If IsObject(aArray(x)) Then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  If offset < 0 Then offset = 0
  ReDim aArray(aCount-1+offset)   'Resize original array
  For x = 0 To aCount-1       'set objects back into original array
    If IsObject(a(x)) Then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
  BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)    'Set up line via two points, no clamping. Input X, output Y
  Dim x, y, b, m
  x = input
  m = (Y2 - Y1) / (X2 - X1)
  b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x
      y = .y
      z = .z
      velx = .velx
      vely = .vely
      velz = .velz
      id = .ID
      mass = .mass
      radius = .radius
    End With
  End Property
  Public Sub Reset()
    x = Empty
    y = Empty
    z = Empty
    velx = Empty
    vely = Empty
    velz = Empty
    id = Empty
    mass = Empty
    radius = Empty
  End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  Dim y 'Y output
  Dim L 'Line
  'find active line
  Dim ii
  For ii = 1 To UBound(xKeyFrame)
    If xInput <= xKeyFrame(ii) Then
      L = ii
      Exit For
    End If
  Next
  If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)    'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )     'Clamp lower
  If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )    'Clamp upper

  LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************
' To add the flipper tricks you must
'  - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'  - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines
'  - Include FlipperActivate and FlipperDeactivate in the Flipper solenoid subs

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
     Dim gBOT
     gBOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
          gBOT(b).velx = gBOT(b).velx / 1.3
          gBOT(b).vely = gBOT(b).vely - 0.5
        End If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
  End If
End Sub


Dim FCCDamping: FCCDamping = 0.4

Sub FlipperCradleCollision(ball1, ball2, velocity)
  if velocity < 0.7 then exit sub   'filter out gentle collisions
    Dim DoDamping, coef
    DoDamping = false
    'Check left flipper
    If LeftFlipper.currentangle = LFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, LeftFlipper) OR FlipperTrigger(ball2.x, ball2.y, LeftFlipper) Then DoDamping = true
    End If
    'Check right flipper
    If RightFlipper.currentangle = RFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, RightFlipper) OR FlipperTrigger(ball2.x, ball2.y, RightFlipper) Then DoDamping = true
    End If
    If DoDamping Then
    coef = FCCDamping
        ball1.velx = ball1.velx * coef: ball1.vely = ball1.vely * coef: ball1.velz = ball1.velz * coef
        ball2.velx = ball2.velx * coef: ball2.vely = ball2.vely * coef: ball2.velz = ball2.velz * coef
    End If
End Sub





'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
  DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function DistanceFromFlipperAngle(ballx, bally, Flipper, Angle)
  DistanceFromFlipperAngle = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Angle + 90)) + Flipper.x, Sin(Radians(angle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function

'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
'Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0
    SOSRampup = 2.5
  Case 1
    SOSRampup = 6
  Case 2
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'   Const EOSReturn = 0.055  'EM's
  Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
'   Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
  FlipperPress = 0
  Flipper.eostorqueangle = EOSA
  Flipper.eostorque = EOST * EOSReturn / FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b, gBOT
        gBOT = GetBalls

    For b = 0 To UBound(gBOT)
      If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3 * Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
    If FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If
  End If
End Sub

Const LiveDistanceMin = 5  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)
Const BaseDampen = 0.55

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
    Dim Dir, LiveDist
    Dir = Flipper.startangle / Abs(Flipper.startangle)    '-1 for Right Flipper
    Dim LiveCatchBounce   'If live catch is not perfect, it won't freeze ball totally
    Dim CatchTime
    CatchTime = GameTime - FCount
    LiveDist = Abs(Flipper.x - ball.x)

    If CatchTime <= LiveCatch And parm > 3 And LiveDist > LiveDistanceMin And LiveDist < LiveDistanceMax Then
        If CatchTime <= LiveCatch * 0.5 Then   'Perfect catch only when catch time happens in the beginning of the window
            LiveCatchBounce = 0
        Else
            LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)  'Partial catch when catch happens a bit late
        End If

        If LiveCatchBounce = 0 And ball.velx * Dir > 0 And LiveDist > 30 Then ball.velx = 0

        If ball.velx * Dir > 0 And LiveDist < 30 Then
            ball.velx = BaseDampen * ball.velx
            ball.vely = BaseDampen * ball.vely
            ball.angmomx = BaseDampen * ball.angmomx
            ball.angmomy = BaseDampen * ball.angmomy
            ball.angmomz = BaseDampen * ball.angmomz
        Elseif LiveDist > 30 Then
            ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
            ball.angmomx = 0
            ball.angmomy = 0
            ball.angmomz = 0
        End If
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
    End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************





'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen ActiveBall
  TargetBouncer ActiveBall, 0.7
End Sub

Dim RubbersD        'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False    'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1    'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64    'there's clamping so interpolate up to 56 at least

Dim SleevesD  'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False    'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn   'tbpOut.text
  Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
  End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
  End Sub

  Public Sub Dampen(aBall)
    If threshold Then
      If BallSpeed(aBall) < threshold Then Exit Sub
    End If
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
    "actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
    If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)

    aBall.velx = aBall.velx * coef
    aBall.vely = aBall.vely * coef
    aBall.velz = aBall.velz * coef
    If debugOn Then TBPout.text = str
  End Sub

  Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
      aBall.velx = aBall.velx * coef
      aBall.vely = aBall.vely * coef
      aBall.velz = aBall.velz * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    Dim x
    For x = 0 To UBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
    Next
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub
End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
  Public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize
    ReDim ballvel(0)
    ReDim ballvelx(0)
    ReDim ballvely(0)
  End Sub

  Public Sub Update() 'tracks in-ball-velocity
    Dim str, b, AllBalls, highestID
    allBalls = GetBalls

    For Each b In allballs
      If b.id >= HighestID Then highestID = b.id
    Next

    If UBound(ballvel) < highestID Then ReDim ballvel(highestID)  'set bounds
    If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)  'set bounds
    If UBound(ballvely) < highestID Then ReDim ballvely(highestID)  'set bounds

    For Each b In allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
Sub RDampen_Timer
  Cor.Update
End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************


'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9  'Level of bounces. Recommmended value of 0.7-1

Sub TargetBouncer(aBall,defvalue)
  Dim zMultiplier, vel, vratio
  If TargetBouncerEnabled = 1 And aball.z < 30 Then
    '   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    vel = BallSpeed(aBall)
    If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
    Select Case Int(Rnd * 6) + 1
      Case 1
        zMultiplier = 0.2 * defvalue
      Case 2
        zMultiplier = 0.25 * defvalue
      Case 3
        zMultiplier = 0.3 * defvalue
      Case 4
        zMultiplier = 0.4 * defvalue
      Case 5
        zMultiplier = 0.45 * defvalue
      Case 6
        zMultiplier = 0.5 * defvalue
    End Select
    aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
    aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
    aBall.vely = aBall.velx * vratio
    '   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    '   debug.print "conservation check: " & BallSpeed(aBall)/vel
  End If
End Sub

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer ActiveBall, 1
End Sub



'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS
Set LS = New SlingshotCorrection
Dim RS
Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00, - 4
  AddSlingsPt 1, 0.45, - 7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4
End Sub

Sub AddSlingsPt(idx, aX, aY)    'debugger wrapper for adjusting flipper script In-game
  Dim a
  a = Array(LS, RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function
'
'Function RotPoint(x,y,angle)
' dim rx, ry
' rx = x*dCos(angle) - y*dSin(angle)
' ry = x*dSin(angle) + y*dCos(angle)
' RotPoint = Array(rx,ry)
'End Function

Class SlingshotCorrection
  Public DebugOn, Enabled
  Private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut

  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
    Enabled = True
  End Sub

  Public Property Let Object(aInput)
    Set Slingshot = aInput
  End Property

  Public Property Let EndPoint1(aInput)
    SlingX1 = aInput.x
    SlingY1 = aInput.y
  End Property

  Public Property Let EndPoint2(aInput)
    SlingX2 = aInput.x
    SlingY2 = aInput.y
  End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    Dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1
      YL = SlingY1
      XR = SlingX2
      YR = SlingY2
    Else
      XL = SlingX2
      YL = SlingY2
      XR = SlingX1
      YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If Abs(XR - XL) > Abs(YR - YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If Not IsEmpty(ModIn(0) ) Then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      '   debug.print " BallPos=" & BallPos &" Angle=" & Angle
      '   debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled Then aBall.Velx = RotVxVy(0)
      If Enabled Then aBall.Vely = RotVxVy(1)
      '   debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      '   debug.print " "
    End If
  End Sub
End Class


'*******************************************
' ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled) 'Left flipper solenoid callback
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire  'leftflipper.rotatetoend

    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
        RandomSoundReflipUpLeft LeftFlipper
      Else
        SoundFlipperUpAttackLeft LeftFlipper
        RandomSoundFlipperUpLeft LeftFlipper
      End If

  Else
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart

    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel

  End If
End Sub

Sub SolRFlipper(Enabled) 'Right flipper solenoid callback
  If Enabled Then
    FlipperActivate RightFlipper, RFPress
    RF.Fire 'rightflipper.rotatetoend

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If

  Else
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart

    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel

  End If
End Sub



'******************************************************
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'  Metals (all metal objects, metal walls, metal posts, metal wire guides)
'  Apron (the apron walls and plunger wall)
'  Walls (all wood or plastic walls)
'  Rollovers (wire rollover triggers, star triggers, or button triggers)
'  Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'  Gates (plate gates)
'  GatesWire (wire gates)
'  Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
' - Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
' - Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
' - Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
' - Ball rolling sounds sub
'
' Tutorial videos by Apophis
' Audio : Adding Fleep Part 1         https://youtu.be/rG35JVHxtx4?si=zdN9W4cZWEyXbOz_
' Audio : Adding Fleep Part 2         https://youtu.be/dk110pWMxGo?si=2iGMImXXZ0SFKVCh
' Audio : Adding Fleep Part 3         https://youtu.be/ESXWGJZY_EI?si=6D20E2nUM-xAw7xy


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1        'volume level; range [0, 1]
NudgeRightSoundLevel = 1        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1        'volume level; range [0, 1]
StartButtonSoundLevel = 0.1      'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1        'volume level; range [0, 1]
RollingSoundFactor = 1.1 / 5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010    'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635    'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0            'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45          'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel    'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel    'sound helper; not configurable
SlingshotSoundLevel = 0.95            'volume level; range [0, 1]
BumperSoundFactor = 4.25            'volume multiplier; must not be zero
KnockerSoundLevel = 1              'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2      'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5      'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075 / 5      'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025      'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025      'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8    'volume level; range [0, 1]
WallImpactSoundFactor = 0.075          'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075 / 3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5      'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10  'volume multiplier; must not be zero
DTSoundLevel = 0.25        'volume multiplier; must not be zero
RolloverSoundLevel = 0.25      'volume level; range [0, 1]
SpinnerSoundLevel = 0.5      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8          'volume level; range [0, 1]
BallReleaseSoundLevel = 1        'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015  'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025 / 5      'volume multiplier; must not be zero

'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / tableheight - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

' Thalamus, AudioFade - Patched
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 5) 'was 10
	Else
		AudioFade = CSng( - (( - tmp) ^ 5) ) 'was 10
	End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / tablewidth - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

' Thalamus, AudioPan - Patched
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 5) 'was 10
	Else
		AudioPan = CSng( - (( - tmp) ^ 5) ) 'was 10
	End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = CSng(BallVel(ball) ^ 2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = CSng((ball.velz) ^ 2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
End Function

Function RndInt(min, max) ' Sets a random number integer between min and max
  RndInt = Int(Rnd() * (max - min + 1) + min)
End Function

Function RndNum(min, max) ' Sets a random number between min and max
  RndNum = Rnd() * (max - min) + min
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////

Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * VolumeDial, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub

Sub SoundPlungerPull()
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////

Sub KnockerSolenoid()
  PlaySoundAtLevelStatic SoundFXDOF("Knocker_1",138,DOFPulse,DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFXDOF("BallRelease" & Int(Rnd * 7) + 1,135,DOFPulse,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFXDOF("Sling_L" & Int(Rnd * 10) + 1,131,DOFPulse,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFXDOF("Sling_R" & Int(Rnd * 8) + 1,135,DOFPulse,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFXDOF("Bumpers_Top_" & Int(Rnd * 5) + 1,111,DOFPulse,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFXDOF("Bumpers_Middle_" & Int(Rnd * 5) + 1,109,DOFPulse,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFXDOF("Bumpers_Bottom_" & Int(Rnd * 5) + 1,115,DOFPulse,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////

Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////

Sub RandomSoundFlipperUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm / 10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm / 10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd * 4) + 1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////

Sub Rubbers_Hit(idx)
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 5 Then
    RandomSoundRubberStrong 1
  End If
  If finalspeed <= 5 Then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////

Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd * 10) + 1
    Case 1
      PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 2
      PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 3
      PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 4
      PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 5
      PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 6
      PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 7
      PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 8
      PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 9
      PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 10
      PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd * 9) + 1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
  RandomSoundWall()
End Sub

Sub RandomSoundWall()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 5) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 4) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 3) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

Sub RandomSoundMetal()
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
  RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
  RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////

Sub RandomSoundBottomArchBallGuide()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Bounce_" & Int(Rnd * 2) + 1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd * 3) + 1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(ActiveBall.id) < 4) And cor.ballvely(ActiveBall.id) > 7 Then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd * 3) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd * 7) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

Sub RandomSoundTargetHitStrong()
  PlaySoundAtLevelActiveBall SoundFXDOF("Target_Hit_" & Int(Rnd * 4) + 5,113,DOFPulse,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFXDOF("Target_Hit_" & Int(Rnd * 4) + 1,113,DOFPulse,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 10 Then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft ActiveBall
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////

Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd * 9) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd * 7) + 1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////

Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd * 5) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, ActiveBall
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub Arch1_hit()
  If ActiveBall.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If ActiveBall.velx <  - 8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If ActiveBall.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If ActiveBall.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0
      PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1
      PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////

Sub OnBallBallCollision(ball1, ball2, velocity)

  FlipperCradleCollision ball1, ball2, velocity

  Dim snd
  Select Case Int(Rnd * 7) + 1
    Case 1
      snd = "Ball_Collide_1"
    Case 2
      snd = "Ball_Collide_2"
    Case 3
      snd = "Ball_Collide_3"
    Case 4
      snd = "Ball_Collide_4"
    Case 5
      snd = "Ball_Collide_5"
    Case 6
      snd = "Ball_Collide_6"
    Case 7
      snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05    'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025 * RelayGISoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025 * RelayGISoundLevel, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025 * RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025 * RelayFlashSoundLevel, obj
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************


