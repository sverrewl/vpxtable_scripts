'*****************************************************************************************************
'*      William's Wonderland (1955)                                *
'*      VPX Table Primary Build/Script by JINO0372                             *
'*      Artwork by JINO0372                                    *
'*      Backglass & dB2S by JINO0372                               *
'*                                                   *
'*                                                   *                                                   *
'*      ++Special thanks to BALATER++                              *
'*                                                   *
'*      I used the script for 4 Roses as a guide to script this table.               *
'*                                                   *
'*      ++Special thanks to ITCHIGO AND PBECKER++                        *
'*                                                   *
'*      Their VP9 table helped understand the scoring and gameplay.                        *
'*                                                   *
'*          ++Additional thanks to the following++                           *
'*                                                   *
'*      JPSALAS - Rolling sounds script, LUT selection, Material Physics             *
'*      JLOULOULOU - Flipper Tricks Script                             *
'*      ROTHBAUERW -  Manual Ball Control and Ball Dropping Sound                *
'*      NINUZZU - Flipper shadows                                      *
'*                                                   *
'*      ++Also thanks to everyone in the VP community!!!++                     *
'*                                                   *
'*                                                   *
'*****************************************************************************************************


Option Explicit
Randomize


On Error Resume Next
ExecuteGlobal GetTextFile("core.vbs")
If Err Then MsgBox "You need the core.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

Const cGameName = "wonderland"
Const B2STableName="wonderland"

Dim B2SOn
Dim Score
Dim BallstoPlay
Dim Credits
Dim GameinProgress
Dim GameStart
Dim BallsPerGame
Dim Highscore
Dim Tilt
Dim TiltSensor
Dim Replay
Dim BallinLane
Dim Playedballs
Dim aXpos
Dim aXadd
Dim aYpos
Dim aYadd
Dim aZpos
Dim aBall
Dim GotReplay
Dim cball
Dim X
Dim SpotNum
Dim gball
Dim TopLiteUp
Dim BottomLiteUp
Dim SkillHoleBonus
Dim TopBumperLights

Dim EnableRetractPlunger
EnableRetractPlunger = false 'Change to true to enable retracting the plunger at a linear speed; wait for 1 second at the maximum position; move back towards the resting position; nice for button/key plungers

Dim LifterOption

LifterOption=2 '1 = manual lifter...2 = auto lifter

Dim DesktopMode: DesktopMode = Wonderland.ShowDT

If DesktopMode = True then

  For Each X in DT
    X.Visible=True
    Next

  Else

  For Each X in DT
    X.Visible=False
    Next


End If

  Kicker2.CreateBall
  Kicker2.Kick 180,1
    Kicker2.Enabled=False
  Kicker003.CreateBall
  Kicker003.Kick 180,1
  Kicker003.Enabled=False
  Kicker004.CreateBall
  Kicker004.Kick 180,1
  Kicker004.Enabled=False
  Kicker005.CreateBall
  Kicker005.Kick 180,1
  Kicker005.Enabled=False
  Kicker006.CreateBall
  Kicker006.Kick 180,1
  Kicker006.Enabled=False


Sub Wonderland_Init()
LoadEM
loadhs
loadcreds
LoadLUT 'JP's LUT Control
UpdatePostIt
ballinlane=1

If B2SOn then
  Controller.B2SSetGameOver 1
  Controller.B2SSetScorePlayer1 Score
  Controller.B2SSetTilt 0
  Controller.B2SSetCredits Credits
  Controller.B2SSetdata 61,0
  Controller.B2SSetdata 62,0
  Controller.B2SSetdata 63,0
  Controller.B2SSetdata 64,0
  Controller.B2SSetdata 65,0
  Controller.B2SSetdata 66,0
  Controller.B2SSetdata 67,0
  Controller.B2SSetdata 68,0
  Controller.B2SSetdata 69,0
  Controller.B2SSetdata 71,0
  Controller.B2SSetdata 72,0
  Controller.B2SSetdata 73,0
  Controller.B2SSetdata 74,0
  Controller.B2SSetdata 75,0
  Controller.B2SSetdata 76,0
  Controller.B2SSetdata 77,0
  Controller.B2SSetdata 78,0
  Controller.B2SSetdata 79,0
  Controller.B2SSetdata 81,0
  Controller.B2SSetdata 82,0
  Controller.B2SSetdata 83,0
  Controller.B2SSetdata 84,0
  Controller.B2SSetdata 85,0
  Controller.B2SSetdata 86,0
  Controller.B2SSetdata 87,0
  Controller.B2SSetdata 88,0
  Controller.B2SSetdata 89,0
  Controller.B2SSetdata 90,0
End If

Replay=0
TiltSensor=0
GameinProgress=0
GameStart=0
Score=0
BallsPerGame=5
BallstoPlay=0
ScoreText.Text="0"
SpotNum=0
TopLiteUp=0
BottomLiteUp=0
SkillHoleBonus=0
TopBumperLights=0

If Credits > 0 Then
  Credittext.Text=Credits
  CreditsReel.SetValue(Credits)
  If B2SOn Then Controller.B2SSetCredits Credits
  DOF 119, DOFOn
Else
  Credittext.Text= "0"
  DOF 119, DOFOff
end If

If BallstoPlay > 0 Then
BallsText.Text=BallstoPlay
TopLiteUpText.Text=TopLiteUp
BottomLightUpText.Text=BottomLiteUp
TopBumperText.Text=TopBumperLights
else
BallsText.text = "0"
TopLiteUpText.text = "0"
BottomLightUpText.text = "0"
TopBumperText.text = "0"
end If

if Credits > 0 Then DOF 121, DOFOn

If DesktopMode=True Then
GameOver.Visible=True
GameOver.SetValue 1
Else
GameOver.Visible=False
GameOver.SetValue 1
End If

gball=0
SkillLt1.State=0
SkillLt2.State=0
SkillLt3.State=0
SkillLt4.State=0
SkillLt5.State=0
TopSpecial.State=0

For Each X in BTopLights
  X.State=0
  Next

For Each X in Millions  'backglass million values
  X.State=0
  Next

For Each X in HundredThousand 'backglass million values
  X.State=0
  Next

For Each X in TenThousand  'backglass million values
  X.State=0
  Next


For Each X in SpotLights 'flower spot lights
  X.State=0
    Next

For Each X in BottomSpecialLights 'bottom specials
  X.State=0
  Next

For each X in PFLights '1-10 lane lights
  X.State = 0
  Next

End Sub

Sub EndofGame()
  credittext.text=credits
  CreditsReel.SetValue(Credits)
  BallsText.Text=BallstoPlay
  TopLiteUpText.Text=TopLiteUp
  BottomLightUpText.Text=BottomLiteUp
  TopBumperText.Text=TopBumperLights
  GameinProgress=0
    StopSound "buzz"
    StopSound "buzzl"

  If DesktopMode = True Then
  GameOver.Visible=True
  GameOver.SetValue 1
  Else
  GameOver.Visible=False
  GameOver.SetValue 1
  End If

  If gball=5 then
    Credits=Credits+1
  ElseIf gball=5 and BottomLiteUp=10 Then
    Credits=Credits+5
  End If

  If B2SOn Then Controller.B2SSetCredits Credits

  For Each X in SpotLights 'flower spot lights
    X.State=0
    Next

End Sub



Sub Wonderland_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
        If EnableRetractPlunger Then
            Plunger.PullBackandRetract
        Else
        Plunger.PullBack
        End If
    PlaySound "plungerpull",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  If HSEnterMode Then HighScoreProcessKey(keycode)

  If keycode=LeftFlipperKey Then
    If Tilt=0 Then
      If GameinProgress=1 Then
        LeftFlipper.TimerEnabled = True 'This line is only for ninuzzu's flipper shadows!
        SolLFlipper 1
      End If
    End If
  End If

  If keycode=RightFlipperKey Then
    If Tilt=0 Then
      If GameinProgress=1 Then
        RightFlipper.TimerEnabled = True 'This line is only for ninuzzu's flipper shadows!
        SolRFlipper 1
      End If
    End If
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    TiltSensor=TiltSensor+25
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    TiltSensor=TiltSensor+25
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 9:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    TiltSensor=TiltSensor+25
  End If

  If GameinProgress=0 Then ' JP's LUT Control
    If keycode = LeftMagnaSave Then bLutActive = True: Lutbox.text = "level of darkness " & LUTImage + 1
    If keycode = RightMagnaSave AND bLutActive Then NextLUT:End If
  End If

  If LifterOption=1 and keycode = RightMagnasave and BallstoPlay>0 and BallinLane=0 then
    playsound "ballup"
    BallLifter.Enabled=True
  End If


  If Keycode = StartGameKey And Not HSEnterMode=true then
    If GameinProgress=0 and credits>0 then
      cball=0
      DrainLock.IsDropped=1
      TrayOpen.TimerInterval=3    'Set to ms
            TrayOpen.TimerEnabled = True
      ballinlane=0
      playedballs=0
      Credits=Credits-1
      If Credits < 1 Then DOF 121, DOFOff
      CreditText.text=Credits
      CreditsReel.SetValue(Credits)
      BallsText.Text=BallstoPlay
      TopLiteUpText.Text=TopLiteUp
      BottomLightUpText.Text=BottomLiteUp
      TopBumperText.Text=TopBumperLights
      If B2SOn then
        Controller.B2SSetData 35, 0 ' GameOver off
        Controller.B2SSetPlayerUp 1
        Controller.B2SSetTilt 0
        Controller.B2SSetScorePlayer1 0
        Controller.B2SSetCredits Credits
      End If
      GameStart=1
'     playsound "DrainShorter"
      PlaySound "GameStart"
      GameinProgress=1
      GameOver.SetValue 0
      Start_Game()
    End If
  End IF
  If Keycode= 6 then
    PlaySoundAtVol "coin3", Drain, 1
    If credits<47 then Credits=Credits+1:DOF 121, DOFOn:end if
    Credittext.text=Credits
    CreditsReel.SetValue(Credits)
    BallsText.Text=BallstoPlay
    TopLiteUpText.Text=TopLiteUp
    BottomLightUpText.Text=BottomLiteUp
    TopBumperText.Text=TopBumperLights
    If B2SOn then
      Controller.B2SSetCredits Credits
    End If
  End if

End Sub

Sub Wonderland_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySound "plunger",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  If Tilt=0 and GameinProgress=1 Then
    If keycode = LeftFlipperKey Then SolLFlipper 0
    If keycode = RightFlipperKey Then SolRFlipper 0
  End If

  If keycode=24 and GameStart=0 and GameinProgress=0 then
    call Toggleroutine()
    playsound "metal4"
  End If

  If GameinProgress=0 Then 'JP's LUT Control
    If keycode = LeftMagnaSave Then bLutActive = False: LutBox.text = ""
  End If

End Sub

'*******************
'  Flipper Subs
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), LeftFlipper, 1
    PlayLoopSoundAtVol "buzzl", LeftFlipper, 1
    LeftFlipper.RotateToEnd
  Else
    PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper, 1
    StopSound "buzzl"
    LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), RightFlipper, 1
    PlayLoopSoundAtVol "buzz", RightFlipper, 1
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), RightFlipper, 1
    StopSound "buzz"
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber2", 0, parm / 60, Audiopan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber2", 0, parm / 60, Audiopan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'**********************************************
'    Flipper adjustments - enable tricks
'             by JLouLouLou
'**********************************************

Dim FlipperPower
Dim FlipperElasticity
Dim EOSTorque, EOSAngle
Dim SOSTorque, SOSAngle
Dim FullStrokeEOS_Torque

Dim LLiveCatchTimer
Dim RLiveCatchTimer
Dim LiveCatchSensivity

FlipperPower = 5000
FlipperElasticity = 0.25
EOSTorque = 0.2
EOSAngle = 6
SOSTorque = 0.1
SOSAngle = 6
FullStrokeEOS_Torque = 0.5
LiveCatchSensivity = 25 'adjust as you prefer
LLiveCatchTimer = 0
RLiveCatchTimer = 0

Sub RealTimeFast_Timer 'flipper's tricks timer
    'Flipper Stroke Routine
    If LeftFlipper.CurrentAngle => LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque:End If                                                       'Start of Stroke for Tap pass and Tap shoot
    If LeftFlipper.CurrentAngle < LeftFlipper.StartAngle - SOSAngle and LeftFlipper.CurrentAngle > LeftFlipper.EndAngle + EOSAngle Then LeftFlipper.Strength = FlipperPower:End If      'Full Stroke
    If LeftFlipper.CurrentAngle < LeftFlipper.EndAngle + EOSAngle and LeftFlipper.CurrentAngle > LeftFlipper.EndAngle Then LeftFlipper.Strength = FlipperPower * EOSTorque:End If       'EOS Stroke
    If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle Then LeftFlipper.Strength = FlipperPower * FullStrokeEOS_Torque:End If                                                           'Bunny Bump Remover

    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque:End If                                                     'Start of Stroke for Tap pass and Tap shoot
    If RightFlipper.CurrentAngle > RightFlipper.StartAngle + SOSAngle and RightFlipper.CurrentAngle < RightFlipper.EndAngle - EOSAngle Then RightFlipper.Strength = FlipperPower:End If 'Full Stroke
    If RightFlipper.CurrentAngle > RightFlipper.EndAngle - EOSAngle and RightFlipper.CurrentAngle < RightFlipper.EndAngle Then RightFlipper.Strength = FlipperPower * EOSTorque:End If  'EOS Stroke
    If RightFlipper.CurrentAngle = RightFlipper.EndAngle Then RightFlipper.Strength = FlipperPower * FullStrokeEOS_Torque:End If                                                        'Bunny Bump Remover

    'Live Catch Routine
    If LeftFlipper.CurrentAngle <= LeftFlipper.EndAngle + EOSAngle and LeftFlipper.CurrentAngle => LeftFlipper.EndAngle Then
        LLiveCatchTimer = LLiveCatchTimer + 1
        If LLiveCatchTimer < LiveCatchSensivity Then
            LeftFlipper.Elasticity = 0.2
        Else
            LeftFlipper.Elasticity = FlipperElasticity
            LLiveCatchTimer = LiveCatchSensivity
        End If
    Else
        LLiveCatchTimer = 0
    End If
    If RightFlipper.CurrentAngle => RightFlipper.EndAngle - EOSAngle and RightFlipper.CurrentAngle <= RightFlipper.EndAngle Then
        RLiveCatchTimer = RLiveCatchTimer + 1
        If RLiveCatchTimer < LiveCatchSensivity Then
            RightFlipper.Elasticity = 0.2
        Else
            RightFlipper.Elasticity = FlipperElasticity
            RLiveCatchTimer = LiveCatchSensivity
        End If
    Else
        RLiveCatchTimer = 0
    End If
End Sub

Sub Toggleroutine()
  if ballspergame=3 then
    ballspergame=5
  else ballspergame=3
  end if
end sub


Sub BallLifter_Timer
  BallLifter.enabled=false
  Kicker1.CreateBall
  Kicker1.kick 90,2
  ballinlane=1
  BallLifter2.enabled=true
End Sub

Sub BallLifter2_Timer
  BallLifter2.enabled=false
End Sub

Sub TiltTimer_Timer
  If TiltSensor>1 then TiltSensor=TiltSensor-1
  If TiltSensor>99 then
    TiltSensor=0
    If Tilt=0 then playsound "buzzer" End If
    Tilt=1
    TiltLight.Visible=True
    TiltLight.SetValue 1
    If B2SOn then Controller.B2SSetTilt 1
  End If
End Sub


Sub Start_Game()
  If GameStart=1 and GameinProgress=1 then

    Score=0

    If B2SOn then
      Controller.B2SSetdata 61,0
      Controller.B2SSetdata 62,0
      Controller.B2SSetdata 63,0
      Controller.B2SSetdata 64,0
      Controller.B2SSetdata 65,0
      Controller.B2SSetdata 66,0
      Controller.B2SSetdata 67,0
      Controller.B2SSetdata 68,0
      Controller.B2SSetdata 69,0
      Controller.B2SSetdata 71,0
      Controller.B2SSetdata 72,0
      Controller.B2SSetdata 73,0
      Controller.B2SSetdata 74,0
      Controller.B2SSetdata 75,0
      Controller.B2SSetdata 76,0
      Controller.B2SSetdata 77,0
      Controller.B2SSetdata 78,0
      Controller.B2SSetdata 79,0
      Controller.B2SSetdata 81,0
      Controller.B2SSetdata 82,0
      Controller.B2SSetdata 83,0
      Controller.B2SSetdata 84,0
      Controller.B2SSetdata 85,0
      Controller.B2SSetdata 86,0
      Controller.B2SSetdata 87,0
      Controller.B2SSetdata 88,0
      Controller.B2SSetdata 89,0
      Controller.B2SSetdata 90,1
    End If

    Ballstoplay=5
    BallstoPlay=BallsPerGame
    New_Ball
    TiltLight.Visible=False
    TiltLight.SetValue 0
    Tilt=0
        BallsText.Text=BallstoPlay
    gball=0
    SkillLt1.State=0
    SkillLt2.State=0
    SkillLt3.State=0
    SkillLt4.State=0
    SkillLt5.State=0
    SpotLt1.State=1
    SpotNum=1
    TopSpecial.State=0
    TopLiteUp=0
    TopBumperLights=0
    BottomLiteUp=0
    TopLiteUpText.Text=TopLiteUp
    BottomLightUpText.Text=BottomLiteUp
    TopBumperText.Text=TopBumperLights

    GameOver.Visible=False
    GameOver.SetValue 0

    '*****GI Lights On
    dim xx

    For each xx in GI:xx.State = 1: Next

    For Each X in BTopLights
      X.State=0
      Next

    For Each X in Millions  'backglass million values
      X.State=0
      Next

    For Each X in HundredThousand 'backglass million values
      X.State=0
      Next

    For Each X in TenThousand  'backglass million values
      X.State=0
      Next

    For Each X in BottomSpecialLights 'outlane specials
      X.State=0
      Next

    For each X in PFLights '1-10 lane lights
      X.State = 1
      Next
  End If
End Sub

Sub New_Ball
  ballinlane=0
End Sub


Sub Drain1_Hit()
  PlaySound "drain",0,1,AudioPan(Drain1),0.25,0,0,1,AudioFade(Drain1)
  Drain1.DestroyBall
  cball=cball+1

  If LifterOption = 2 and cball = 5 Then
    If BallinLane=0 and BallstoPlay>0 Then
      PlaySound "ballup"
      BallLifter.Enabled=True
    End If
  End If


End Sub

Sub TrayOpen_Timer()
  TrayOpen.visible=0
  If cball = 5 Then
  TrayOpen.visible=1
  End If
End Sub

Sub Gate001_hit()
  DrainLock.IsDropped=0
  PlaySound "Gate5",0,1,AudioPan(Gate001),0.25,0,0,1,AudioFade(Gate001)
End Sub

Sub RubberWall1_hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Drain_Hit()
  DOF 117, DOFPulse
    TiltLight.Visible=False
  TiltLight.SetValue 0
  Tilt=0
  If B2SOn then Controller.B2SSetTilt 0
  BallstoPlay=BallstoPlay-1
  If BallstoPlay=0 then
    EndofGame()
    GameOver.SetValue 1
    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart
    If B2SOn then Controller.B2SSetGameOver 1
    playsound "motorleer"
    If score>HighScore Then
      HighScore=score
      HighScoreEntryInit()
      UpdatePostIt
      savehs
    End if
  End If

  If LifterOption=1 AND BallstoPlay>0 then
    Credittext.Text=Credits
    CreditsReel.SetValue(Credits)
    BallsText.Text=BallstoPlay
    TopLiteUpText.Text=TopLiteUp
    BottomLightUpText.Text=BottomLiteUp
    TopBumperText.Text=TopBumperLights
      If B2SOn then
        Controller.B2SSetCredits Credits
      End If
    New_Ball
  End If


  If LifterOption=2 AND BallstoPlay>0 then
    Credittext.Text=Credits
    CreditsReel.SetValue(Credits)
    BallsText.Text=BallstoPlay
    TopLiteUpText.Text=TopLiteUp
    BottomLightUpText.Text=BottomLiteUp
    TopBumperText.Text=TopBumperLights
      If B2SOn then
        Controller.B2SSetCredits Credits
      End If
    New_Ball
      If LifterOption=2 and Ballinlane=0 and ballstoplay>0 then
        playsound "ballup"
        BallLifter.Enabled=True
      End If
  End If

End Sub


'*****Skill Hole

Sub gobblerr_Hit()
  AddScore 500000
  PlaySound "50"
  gobblerR.TimerEnabled=True
  gobblerr.TimerInterval=2000
  Set aBall=ActiveBall      ' aBall = the ball Being Gobbled
  aXpos=aBall.X
  aXadd=(gobblerR.X-aXpos)/15     ' 15 ticks to roll ball into hole
  aYpos=aBall.Y
  aYadd=(gobblerR.Y-aYpos)/15
  aBall.X=aXpos         ' Put Ball Back
  aBall.Y=aYpos
  aZpos=50            ' Including Starting Height
  aBall.Z=aZPos
  aBall.VelX=0          ' Stop It Rolling
  aBall.VelY=0
  gobblerR.TimerInterval=3      ' Set timer to 3ms
  gobblerR.TimerEnabled=True      ' And Turn It On
End Sub

Sub gobblerr_Timer()
  aZpos = aZpos - 1       'Subtract 1 from ball Z position and repeat the line above
  If aZpos > 25 Then
    aYpos = aYpos + aYadd   ' Move X & X towards Hole Center
    aXpos = aXpos + aXadd
    wall016.visible=0
  Else
    aYpos=gobblerR.Y          'Ball Now in the hole
    aXpos=gobblerR.X
  End If
  aBall.Z=aZpos         ' Move the ball Z position to the value of variable aZpos
  aBall.X=aXpos         ' Roll to Center of Kicker
  aBall.Y=aYpos
  aBall.VelX=0          ' Stop It Really Rolling
  aBall.VelY=0
  If aZpos > -40 Then       'If the ball Z position is above -30 cycle
    Exit Sub
  End If
  Wall016.visible=1

' Recreate drain to try and fix issues

    PlaySoundAtVol "drain", Drain, 1
    DOF 117, DOFPulse
    TiltLight.SetValue 0
    Tilt=0
  gobblerR.TimerEnabled=False     ' Stop this timer
  gobblerR.DestroyBall    'Moved to above       ' Delete Ball From Table
    If B2SOn then Controller.B2SSetTilt 0
  Kicker2.createball
  Kicker2.kick 90, 2
  BallstoPlay=BallstoPlay-1
  gball=gball+1
  If BallstoPlay=0 then
    EndofGame()
    GameOver.SetValue 1
    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart
    If B2SOn then Controller.B2SSetGameOver 1
    playsound "motorleer"
    If score>HighScore Then
      HighScore=score
      HighScoreEntryInit()
      UpdatePostIt
      savehs
    End if
  End If

  If LifterOption=1 AND BallstoPlay>0 then
    Credittext.Text=Credits
    CreditsReel.SetValue(Credits)
    BallsText.Text=BallstoPlay
    TopLiteUpText.Text=TopLiteUp
    BottomLightUpText.Text=BottomLiteUp
    TopBumperText.Text=TopBumperLights
      If B2SOn then
        Controller.B2SSetCredits Credits
      End If
    New_Ball
  End If

  If LifterOption=2 AND BallstoPlay>0 then
    Credittext.Text=Credits
    CreditsReel.SetValue(Credits)
    BallsText.Text=BallstoPlay
    TopLiteUpText.Text=TopLiteUp
    BottomLightUpText.Text=BottomLiteUp
    TopBumperText.Text=TopBumperLights
      If B2SOn then
        Controller.B2SSetCredits Credits
      End If
    New_Ball
      If LifterOption=2 and Ballinlane=0 and ballstoplay>0 then
        playsound "ballup"
        BallLifter.Enabled=True
      End If
  End If


  Gob_End()
End Sub

Sub Gob_End()
  playsound"GobbleFF"
  gobend.Interval=1000      ' Set timer to 3ms
  gobend.Enabled=True
End Sub

sub gobend_timer()
  gobend.Enabled=0
  playsound "ballup"
    If Tilt= 0 then GobbleLights()
end Sub

Sub GobbleLights()
  If gball=1 then SkillLt1.State=1
  If gball=2 then SkillLt2.State=1
  If gball=3 then SkillLt3.State=1
  If gball=4 then SkillLt4.State=1
  If gball=5 then SkillLt5.State=1
End Sub


'*****Top Kicker

Sub Top_Kicker_Hit()            'routine name

    Top_Kicker.TimerInterval=450          '  .450 second timer
    Top_Kicker.TimerEnabled = True         '  enables timer

    If Tilt = 0 then
    Addscore 500000              'add a score of some type here
    PlaySound "50" 'play any sound you put in when ball first hits kicker


  If SpotLt1.State=1 and Lt1.State=1 then
    Lt1.State=0
    TopLiteUp=TopLiteUp+1
    BottomLiteUp=BottomLiteUp+1
  End If
  If SpotLt2.State=1 and Lt2.State=1 then
    Lt2.State=0
    TopLiteUp=TopLiteUp+1
    BottomLiteUp=BottomLiteUp+1
  End If
  If SpotLt3.State=1 and Lt3.State=1 then
    Lt3.State=0
    TopLiteUp=TopLiteUp+1
    BottomLiteUp=BottomLiteUp+1
  End If
  If SpotLt4.State=1 and Lt4.State=1 then
    Lt4.State=0
    TopLiteUp=TopLiteUp+1
    BottomLiteUp=BottomLiteUp+1
  End If
  If SpotLt5.State=1 and Lt5.State=1 then
    Lt5.State=0
    BottomLiteUp=BottomLiteUp+1
  End If
  If SpotLt6.State=1 and Lt6.State=1 then
    Lt6.State=0
    BottomLiteUp=BottomLiteUp+1
  End If
  If SpotLt7.State=1 and Lt7.State=1 then
    Lt7.State=0
    BottomLiteUp=BottomLiteUp+1
  End If
  If SpotLt8.State=1 and Lt8.State=1 then
    Lt8.State=0
    BottomLiteUp=BottomLiteUp+1
  End If
  If SpotLt9.State=1 and Lt9.State=1 then
    Lt9.State=0
    TopBumperLights=TopBumperLights+1
    BottomLiteUp=BottomLiteUp+1
  End If
  If SpotLt10.State=1 then
    Lt10.State=0
    TopBumperLights=TopBumperLights+1
    BottomLiteUp=BottomLiteUp+1
  End If
End If

  If Tilt = 0 then Spots()

End Sub

Sub Spots()
  SpotNum=SpotNum+1
  If SpotNum=1 then SpotLt1.State=1 Else SpotLt1.State=0
  If SpotNum=2 then SpotLt2.State=1 Else SpotLt2.State=0
  If SpotNum=3 then SpotLt3.State=1 Else SpotLt3.State=0
  If SpotNum=4 then SpotLt4.State=1 Else SpotLt4.State=0
  If SpotNum=5 then SpotLt5.State=1 Else SpotLt5.State=0
  If SpotNum=6 then SpotLt6.State=1 Else SpotLt6.State=0
  If SpotNum=7 then SpotLt7.State=1 Else SpotLt7.State=0
  If SpotNum=8 then SpotLt8.State=1 Else SpotLt8.State=0
  If SpotNum=9 then SpotLt9.State=1 Else SpotLt9.State=0
  If SpotNum=10 then SpotLt10.State=1 Else SpotLt10.State=0
  If SpotNum>10 then SpotLt1.State=1
  If SpotNum>10 then SpotNum=1
End Sub


Sub Top_Kicker_Timer()
    Top_Kicker.kick 187-INT(RND()*5), 4  'When timer runs out we will do this. Kick out at 184 +/- 2 degrees
    If Tilt=0 then PlaySound "kicker",0,1,AudioPan(Top_Kicker),0.25,0,0,1,AudioFade(Top_Kicker)  'playsound when ball leaves kicker
    Top_Kicker.Timerenabled = False    ' stops timer
End Sub


'*****Bumpers

Sub Bumper1_Hit
  If Tilt=0 then
    AddScore (10000)
    PlaySound "1"
    PlaySound SoundFX("fx_bumper4",DOFContactors), 0,1,AudioPan(Bumper1),0,0,0,1,AudioFade(Bumper1)
    Spots()
  End If
End Sub

Sub Bumper1_Timer
' Add timer events if needed
End Sub

Sub Bumper2_Hit
  If Tilt=0 then
    AddScore (10000)
    PlaySound "1"
    PlaySound SoundFX("fx_bumper4",DOFContactors), 0,1,AudioPan(Bumper2),0,0,0,1,AudioFade(Bumper2)
    Spots()
  End If
End Sub

Sub Bumper2_Timer
' Add timer events if needed
End Sub

Sub Bumper3_Hit
  If Tilt=0 then
    AddScore (10000)
    PlaySound "1"
    PlaySound SoundFX("fx_bumper4",DOFContactors), 0,1,AudioPan(Bumper3),0,0,0,1,AudioFade(Bumper3)
    Spots()
  End If
End Sub

Sub Bumper3_Timer
' Add timer events if needed
End Sub

Sub Bumper_Top_1_Hit
  If BTopLight1.State=1 and Tilt=0 Then
  AddScore (100000)
  PlaySound "10"
  PlaySound SoundFX("fx_bumper4",DOFContactors), 0,1,AudioPan(Bumper_Top_1),0,0,0,1,AudioFade(Bumper_Top_1)
  End If
End Sub

Sub Bumper_Top_1_Timer
' Add timer events if needed
End Sub

Sub Bumper_Top_2_Hit
  If BTopLight2.State=1 and Tilt=0 Then
  AddScore (100000)
  PlaySound "10"
  PlaySound SoundFX("fx_bumper4",DOFContactors), 0,1,AudioPan(Bumper_Top_2),0,0,0,1,AudioFade(Bumper_Top_2)
  End If
End Sub

Sub Bumper_Top_2_Timer
' Add timer events if needed
End Sub


'*****Slingshots

Sub RightSlingShot_Slingshot
  If Tilt=0 then
    PlaySoundAtVol SoundFXDOF("right_slingshot",205,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 206,2
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    Addscore (10000)
    PlaySound "1"
  End If
End Sub

dim lstep, rstep

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling0.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  If Tilt=0 Then
    PlaySoundAtVol SoundFXDOF("left_slingshot",203,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 204,2
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    Addscore (10000)
    PlaySound "1"
  End If
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing0.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'*****Leaf Switches

Sub Sling3_Slingshot()
  If Tilt=0 Then
    PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*400, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    AddScore(10000)
    PlaySound "1"
  End If
End Sub

Sub Sling4_Slingshot()
  If Tilt=0 Then
    PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*400, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    AddScore(10000)
    PlaySound "1"
  End If
End Sub

Sub Sling5_Slingshot()
  If Tilt=0 Then
    PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*400, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    AddScore(10000)
    PlaySound "1"
  End If
End Sub

Sub Sling6_Slingshot()
  If Tilt=0 Then
    PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*400, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    AddScore(10000)
    PlaySound "1"
  End If
End Sub

Sub Sling7_Slingshot()
  If Tilt=0 Then
    PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*400, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    AddScore(10000)
    PlaySound "1"
  End If
End Sub

Sub Sling8_Slingshot()
  If Tilt=0 Then
    PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*400, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    AddScore(10000)
    PlaySound "1"
  End If
End Sub


'*****Triggers

'10000 pts

Sub No_1_Trigger_Hit()
    If Lt1.State=1 and Tilt=0 Then
      TopLiteUp=TopLiteUp+1
      BottomLiteUp=BottomLiteUp+1
      PlaySoundAtVol "1", ActiveBall, 1
      AddScore(10000)
      Lt1.State=0
    ElseIf Tilt=0 Then
      PlaySoundAtVol "1", ActiveBall, 1
      AddScore(10000)
    Else
      PlaySoundAtVol "pinhit_low", ActiveBall, 1
    End If
End Sub

Sub No_4_Trigger_Hit()
    If Lt4.State=1 and Tilt=0 then
      TopLiteUp=TopLiteUp+1
      BottomLiteUp=BottomLiteUp+1
      PlaySoundAtVol "1", ActiveBall, 1
      AddScore(10000)
      Lt4.State=0
    ElseIf Tilt=0 Then
      PlaySoundAtVol "1", ActiveBall, 1
      AddScore(10000)
    Else
      PlaySoundAtVol "pinhit_low", ActiveBall, 1
    End If
End Sub


'100000 pts

Sub No_2_Trigger_Hit()
    If Lt2.State=1 and Tilt=0 then
      TopLiteUp=TopLiteUp+1
      BottomLiteUp=BottomLiteUp+1
      PlaySound "10"
      AddScore(100000)
      Lt2.State=0
    ElseIf Tilt=0 Then
      PlaySound "10"
      AddScore(100000)
    Else
      PlaySoundAtVol "pinhit_low", ActiveBall, 1
    End If
End Sub

Sub No_3_Trigger_Hit()
    If Lt3.State=1 and Tilt=0 then
      TopLiteUp=TopLiteUp+1
      BottomLiteUp=BottomLiteUp+1
      PlaySound "10"
      AddScore(100000)
      Lt3.State=0
    ElseIf Tilt=0 Then
      PlaySound "10"
      AddScore(100000)
    Else
      PlaySoundAtVol "pinhit_low", ActiveBall, 1
    End If
End Sub

Sub No_5_Trigger_Hit()
    If Lt5.State=1 and Tilt=0 then
      BottomLiteUp=BottomLiteUp+1
      PlaySound "10"
      AddScore(100000)
      Lt5.State=0
    ElseIf Tilt=0 Then
      PlaySound "10"
      AddScore(100000)
    Else
      PlaySoundAtVol "pinhit_low", ActiveBall, 1
    End If
End Sub

Sub No_6_Trigger_Hit()
    If Lt6.State=1 and Tilt=0 then
      BottomLiteUp=BottomLiteUp+1
      PlaySound "10"
      AddScore(100000)
      Lt6.State=0
    ElseIf Tilt=0 Then
      PlaySound "10"
      AddScore(100000)
    Else
      PlaySoundAtVol "pinhit_low", ActiveBall, 1
    End If
End Sub

'100000 or 500000 when special is lit

Sub No_9_Trigger_Hit()
    If Lt9.State=1 and Tilt=0 then
      BottomLiteUp=BottomLiteUp+1
      TopBumperLights=TopBumperLights+1
      PlaySound "10"
      AddScore(100000)
      Lt9.State=0
    Elseif LeftSpecial.State=1 and Tilt=0 Then
      PlaySound "50"
      AddScore(500000)
    ElseIf Tilt=0 Then
      PlaySound "10"
      AddScore(100000)
    Else
      PlaySoundAtVol "pinhit_low", ActiveBall, 1
    End If
End Sub

Sub No_10_Trigger_Hit()
    If Lt10.State=1 and Tilt=0 then
      BottomLiteUp=BottomLiteUp+1
      TopBumperLights=TopBumperLights+1
      PlaySound "10"
      AddScore(100000)
      Lt10.State=0
    Elseif RightSpecial.State=1 and Tilt=0 Then
      PlaySound "50"
      AddScore(500000)
    ElseIf Tilt=0 Then
      PlaySound "10"
      AddScore(100000)
    Else
      PlaySoundAtVol "pinhit_low", ActiveBall, 1
    End If
End Sub

'Special Triggers

Sub SpecialTrigger_Hit()
  If BotSpecial.state=1 and Tilt=0 then
    AddScore(500000)
    PlaySound "50"
    PlaySound "Knocke"
    Credits=Credits+5
  ElseIf Tilt=0 Then
    AddScore(100000)
    PlaySound "10"
  Else
    PlaySoundAtVol "pinhit_low", ActiveBall, 1
  end if
End Sub

Sub SpecialTrigger2_Hit()
  If TopSpecial.state=1 and Tilt=0 then
    AddScore(100000)
    PlaySound "10"
    PlaySound "Knocke"
    Credits=Credits+1
  ElseIf Tilt=0 Then
    AddScore(100000)
    PlaySound "10"
  Else
    PlaySoundAtVol "pinhit_low", ActiveBall, 1
  end if
End Sub


'*****Targets

Sub No_7_Target_Hit()
    If Lt7.State=1 and Tilt=0 then
      BottomLiteUp=BottomLiteUp+1
      PlaySound "10"
      AddScore(100000)
      Lt7.State=0
    ElseIf Tilt=0 Then
      PlaySound "10"
      AddScore(100000)
    Else
      PlaySoundAtVol "pinhit_low", ActiveBall, 1
    End If
End Sub

Sub No_8_Target_Hit()
    If Lt8.State=1 and Tilt=0 then
      BottomLiteUp=BottomLiteUp+1
      PlaySound "10"
      AddScore(100000)
      Lt8.State=0
    ElseIf Tilt=0 Then
      PlaySound "10"
      AddScore(100000)
    Else
      PlaySoundAtVol "pinhit_low", ActiveBall, 1
    End If
End Sub


'*****Scoring


Sub AddScore(f)
  If Tilt=0 then
    Score=Score+f
      If B2SOn Then
        Controller.B2SSetScorePlayer1 Score
        B2SScoreLights()
      End If
    End If

  ScoreText.Text=FormatNumber(Score,0,-1,0,-1)
  TopLiteUpText.Text=TopLiteUp
  BottomLightUpText.Text=BottomLiteUp
  TopBumperText.Text=TopBumperLights
    ScoreLights()
  BotSpecialLights()
  TopSpecialLights()
  BumperTopLights()
End Sub

Sub ScoreLights()

Dim SLScore
Dim R5
Dim R6
Dim R7

SLScore=(Cstr(Score)) 'convert score to string
R5=(Right(SLScore,5))   'numbers right by 5 or TenThousand
R6=(Right(SLScore,6))   'numbers right by 6 or HundredThousand
R7=(Right(SLScore,7))   'numbers right by 7 or Millions

If R5 >= 10000 and R5 <=19999 then TT1.State=1 Else TT1.State = 0
If R5 >= 20000 and R5 <=29999 then TT2.State=1 Else TT2.State = 0
If R5 >= 30000 and R5 <=39999 then TT3.State=1 Else TT3.State = 0
If R5 >= 40000 and R5 <=49999 then TT4.State=1 Else TT4.State = 0
If R5 >= 50000 and R5 <=59999 then TT5.State=1 Else TT5.State = 0
If R5 >= 60000 and R5 <=69999 then TT6.State=1 Else TT6.State = 0
If R5 >= 70000 and R5 <=79999 then TT7.State=1 Else TT7.State = 0
If R5 >= 80000 and R5 <=89999 then TT8.State=1 Else TT8.State = 0
If R5 >= 90000 and R5 <=99999 then TT9.State=1 Else TT9.State = 0

If R6 >= 100000 and R6 <=199999 then HT1.State=1 Else HT1.State = 0
If R6 >= 200000 and R6 <=299999 then HT2.State=1 Else HT2.State = 0
If R6 >= 300000 and R6 <=399999 then HT3.State=1 Else HT3.State = 0
If R6 >= 400000 and R6 <=499999 then HT4.State=1 Else HT4.State = 0
If R6 >= 500000 and R6 <=599999 then HT5.State=1 Else HT5.State = 0
If R6 >= 600000 and R6 <=699999 then HT6.State=1 Else HT6.State = 0
If R6 >= 700000 and R6 <=799999 then HT7.State=1 Else HT7.State = 0
If R6 >= 800000 and R6 <=899999 then HT8.State=1 Else HT8.State = 0
If R6 >= 900000 and R6 <=999999 then HT9.State=1 Else HT9.State = 0

If R7 >= 1000000 and R7 <=1999999 then M1.State=1 Else M1.State = 0
If R7 >= 2000000 and R7 <=2999999 then M2.State=1 Else M2.State = 0
If R7 >= 3000000 and R7 <=3999999 then M3.State=1 Else M3.State = 0
If R7 >= 4000000 and R7 <=4999999 then M4.State=1 Else M4.State = 0
If R7 >= 5000000 and R7 <=5999999 then M5.State=1 Else M5.State = 0
If R7 >= 6000000 and R7 <=6999999 then M6.State=1 Else M6.State = 0
If R7 >= 7000000 and R7 <=7999999 then M7.State=1 Else M7.State = 0
If R7 >= 8000000 and R7 <=8999999 then M8.State=1 Else M8.State = 0
If R7 >= 9000000 and R7 <=9999999 then M9.State=1 Else M9.State = 0

End Sub

Sub B2SScoreLights()

Dim SLScore
Dim R5
Dim R6
Dim R7

SLScore=(Cstr(Score)) 'convert score to string
R5=(Right(SLScore,5))   'numbers right by 5 or TenThousand
R6=(Right(SLScore,6))   'numbers right by 6 or HundredThousand
R7=(Right(SLScore,7))   'numbers right by 7 or Millions

If R5 >= 10000 and R5 <=19999 then Controller.B2SSetData 61, 1 Else Controller.B2SSetData 61, 0
If R5 >= 20000 and R5 <=29999 then Controller.B2SSetData 62, 1 Else Controller.B2SSetData 62, 0
If R5 >= 30000 and R5 <=39999 then Controller.B2SSetData 63, 1 Else Controller.B2SSetData 63, 0
If R5 >= 40000 and R5 <=49999 then Controller.B2SSetData 64, 1 Else Controller.B2SSetData 64, 0
If R5 >= 50000 and R5 <=59999 then Controller.B2SSetData 65, 1 Else Controller.B2SSetData 65, 0
If R5 >= 60000 and R5 <=69999 then Controller.B2SSetData 66, 1 Else Controller.B2SSetData 66, 0
If R5 >= 70000 and R5 <=79999 then Controller.B2SSetData 67, 1 Else Controller.B2SSetData 67, 0
If R5 >= 80000 and R5 <=89999 then Controller.B2SSetData 68, 1 Else Controller.B2SSetData 68, 0
If R5 >= 90000 and R5 <=99999 then Controller.B2SSetData 69, 1 Else Controller.B2SSetData 69, 0

If R6 >= 100000 and R6 <=199999 then Controller.B2SSetData 71, 1 Else Controller.B2SSetData 71, 0
If R6 >= 200000 and R6 <=299999 then Controller.B2SSetData 72, 1 Else Controller.B2SSetData 72, 0
If R6 >= 300000 and R6 <=399999 then Controller.B2SSetData 73, 1 Else Controller.B2SSetData 73, 0
If R6 >= 400000 and R6 <=499999 then Controller.B2SSetData 74, 1 Else Controller.B2SSetData 74, 0
If R6 >= 500000 and R6 <=599999 then Controller.B2SSetData 75, 1 Else Controller.B2SSetData 75, 0
If R6 >= 500000 and R6 <=699999 then Controller.B2SSetData 76, 1 Else Controller.B2SSetData 76, 0
If R6 >= 700000 and R6 <=799999 then Controller.B2SSetData 77, 1 Else Controller.B2SSetData 77, 0
If R6 >= 800000 and R6 <=899999 then Controller.B2SSetData 78, 1 Else Controller.B2SSetData 78, 0
If R6 >= 900000 and R6 <=999999 then Controller.B2SSetData 79, 1 Else Controller.B2SSetData 79, 0

If R7 >= 1000000 and R7 <=1999999 then Controller.B2SSetData 81, 1 Else Controller.B2SSetData 81, 0
If R7 >= 2000000 and R7 <=2999999 then Controller.B2SSetData 82, 1 Else Controller.B2SSetData 82, 0
If R7 >= 3000000 and R7 <=3999999 then Controller.B2SSetData 83, 1 Else Controller.B2SSetData 83, 0
If R7 >= 4000000 and R7 <=4999999 then Controller.B2SSetData 84, 1 Else Controller.B2SSetData 84, 0
If R7 >= 5000000 and R7 <=5999999 then Controller.B2SSetData 85, 1 Else Controller.B2SSetData 85, 0
If R7 >= 5000000 and R7 <=6999999 then Controller.B2SSetData 86, 1 Else Controller.B2SSetData 86, 0
If R7 >= 7000000 and R7 <=7999999 then Controller.B2SSetData 87, 1 Else Controller.B2SSetData 87, 0
If R7 >= 8000000 and R7 <=8999999 then Controller.B2SSetData 88, 1 Else Controller.B2SSetData 88, 0
If R7 >= 9000000 and R7 <=9999999 then Controller.B2SSetData 89, 1 Else Controller.B2SSetData 89, 0

End Sub

Sub BotSpecialLights()
  If BottomLiteUp >= 10 Then
    LeftSpecial.State=1
    RightSpecial.State=1
    BotSpecial.State=1
  Else
    LeftSpecial.State=0
    RightSpecial.State=0
    BotSpecial.State=0
  End If
End Sub

Sub TopSpecialLights()
  If TopLiteUp >= 4 Then
    TopSpecial.State=1
  Else
    TopSpecial.State=0
  End If
End Sub

Sub BumperTopLights()
  If TopBumperLights >= 2 Then
    BTopLight1.State=1
    BTopLight2.State=1
  Else
    BTopLight1.State=0
    BTopLight2.State=0
  End If
End Sub


Sub Wonderland_Exit()
    savecreds
  If B2SOn Then Controller.Stop
End Sub


'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD: And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
'TO SHOW THE SCORE ON POST-IT ADD LINE AT RELEVENT LOCATION THAT HAS:  UpdatePostIt
'TO INITIATE ADDING INITIALS ADD LINE AT RELEVENT LOCATION THAT HAS:  HighScoreEntryInit()
'ADD THE FOLLOWING LINES TO TABLE_INIT TO SETUP POSTIT
' if HSA1="" then HSA1=25
' if HSA2="" then HSA2=25
' if HSA3="" then HSA3=25
' UpdatePostIt
'ADD HSA1, HSA2 AND HSA3 TO SAVE AND LOAD VALUES FOR TABLE
'ADD A TIMER NAMED HighScoreFlashTimer WITH INTERVAL 100 TO TABLE
'SET HSSSCOREX BELOW TO WHATEVER VARIABLE YOU USE FOR HIGH SCORE.
'ADD OBJECTS TO PLAYFIELD (EASIEST TO JUST COPY FROM THIS TABLE)
'IMPORT POST-IT IMAGES

dim brarray
brarray = array("ball go", "ball 1","ball 2","ball 3","ball 4","ball 5","ball")
sub ballaff(nbball)
dim nbll
  nbll=nbball
  if nbll>5 then nbll =6
  ballin.image = brarray(nbll)
end sub


Dim HSA1, HSA2, HSA3
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim HSArray
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex 'Define 6 different score values for each reel to use
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")
Const hsFlashDelay = 4

' ***********************************************************
'  HiScore DISPLAY
' ***********************************************************

Sub UpdatePostIt
  dim tempscore
  HSScorex = highscore
  TempScore = HSScorex
  HSScore1 = 0
  HSScore10 = 0
  HSScore100 = 0
  HSScoreK = 0
  HSScore10k = 0
  HSScore100k = 0
  HSScoreM = 0
  if len(TempScore) > 0 Then
    HSScore1 = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore10 = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore100 = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScoreK = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore10k = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore100k = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScoreM = cint(right(Tempscore,1))
  end If
  Pscore6.image = HSArray(HSScoreM):If HSScorex<1000000 Then PScore6.image = HSArray(10)
  Pscore5.image = HSArray(HSScore100K):If HSScorex<100000 Then PScore5.image = HSArray(10)
  PScore4.image = HSArray(HSScore10K):If HSScorex<10000 Then PScore4.image = HSArray(10)
  PScore3.image = HSArray(HSScoreK):If HSScorex<1000 Then PScore3.image = HSArray(10)
  PScore2.image = HSArray(HSScore100):If HSScorex<100 Then PScore2.image = HSArray(10)
  PScore1.image = HSArray(HSScore10):If HSScorex<10 Then PScore1.image = HSArray(10)
  PScore0.image = HSArray(HSScore1):If HSScorex<1 Then PScore0.image = HSArray(10)
  if HSScorex<1000 then
    PComma.image = HSArray(10)
  else
    PComma.image = HSArray(11)
  end if
  if HSScorex<1000000 then
    PComma2.image = HSArray(10)
  else
    PComma2.image = HSArray(11)
  end if
' if showhisc=1 and showhiscnames=1 then
'   for each object in hiscname:object.visible=1:next
    HSName1.image = ImgFromCode(HSA1, 1)
    HSName2.image = ImgFromCode(HSA2, 2)
    HSName3.image = ImgFromCode(HSA3, 3)
'   else
'   for each object in hiscname:object.visible=0:next
' end if
End Sub

Function ImgFromCode(code, digit)
  Dim Image
  if (HighScoreFlashTimer.Enabled = True and hsLetterFlash = 1 and digit = hsCurrentLetter) then
    Image = "postitBL"
  elseif (code + ASC("A") - 1) >= ASC("A") and (code + ASC("A") - 1) <= ASC("Z") then
    Image = "postit" & chr(code + ASC("A") - 1)
  elseif code = 27 Then
    Image = "PostitLT"
    elseif code = 0 Then
    image = "PostitSP"
    Else
      msgbox("Unknown display code: " & code)
  end if
  ImgFromCode = Image
End Function

Sub HighScoreEntryInit()
  HSA1=0:HSA2=0:HSA3=0
  HSEnterMode = True
  hsCurrentDigit = 0
  hsCurrentLetter = 1:HSA1=1
  HighScoreFlashTimer.Interval = 250
  HighScoreFlashTimer.Enabled = True
  hsLetterFlash = hsFlashDelay
End Sub

Sub HighScoreFlashTimer_Timer()
  hsLetterFlash = hsLetterFlash-1
  UpdatePostIt
  If hsLetterFlash=0 then 'switch back
    hsLetterFlash = hsFlashDelay
  end if
End Sub


' ***********************************************************
'  HiScore ENTER INITIALS
' ***********************************************************

Sub HighScoreProcessKey(keycode)
    If keycode = LeftFlipperKey Then
    hsLetterFlash = hsFlashDelay
    Select Case hsCurrentLetter
      Case 1:
        HSA1=HSA1-1:If HSA1=-1 Then HSA1=26 'no backspace on 1st digit
        UpdatePostIt
      Case 2:
        HSA2=HSA2-1:If HSA2=-1 Then HSA2=27
        UpdatePostIt
      Case 3:
        HSA3=HSA3-1:If HSA3=-1 Then HSA3=27
        UpdatePostIt
     End Select
    End If

  If keycode = RightFlipperKey Then
    hsLetterFlash = hsFlashDelay
    Select Case hsCurrentLetter
      Case 1:
        HSA1=HSA1+1:If HSA1>26 Then HSA1=0
        UpdatePostIt
      Case 2:
        HSA2=HSA2+1:If HSA2>27 Then HSA2=0
        UpdatePostIt
      Case 3:
        HSA3=HSA3+1:If HSA3>27 Then HSA3=0
        UpdatePostIt
     End Select
  End If

    If keycode = PlungerKey Then
    Select Case hsCurrentLetter
      Case 1:
        hsCurrentLetter=2 'ok to advance
        HSA2=HSA1 'start at same alphabet spot
'       EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
      Case 2:
        If HSA2=27 Then 'bksp
          HSA2=0
          hsCurrentLetter=1
        Else
          hsCurrentLetter=3 'enter it
          HSA3=HSA2 'start at same alphabet spot
        End If
      Case 3:
        If HSA3=27 Then 'bksp
          HSA3=0
          hsCurrentLetter=2
        Else
          savehs 'enter it
          HighScoreFlashTimer.Enabled = False
          HSEnterMode = False
        End If
    End Select
    UpdatePostIt
    End If
End Sub
sub savehs
    savevalue "Wonderland", "hiscore", Highscore
    savevalue "Wonderland", "score1", score
  savevalue "Wonderland", "hsa1", HSA1
  savevalue "Wonderland", "hsa2", HSA2
  savevalue "Wonderland", "hsa3", HSA3
end sub

sub loadhs
    dim temp
    temp = LoadValue("Wonderland", "hiscore")
    If (temp <> "") then Highscore = CDbl(temp)
    temp = LoadValue("Wonderland", "score1")
    If (temp <> "") then score = CDbl(temp)
  temp = LoadValue("Wonderland", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("Wonderland", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("Wonderland", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
end sub

sub savecreds
    savevalue "Wonderland1", "credit", Credits
end sub

sub loadcreds
    dim temp
  temp = LoadValue("Wonderland1", "credit")
    If (temp <> "") then Credits = CDbl(temp)
end sub

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


'*************
'   JP'S LUT
'*************

Dim bLutActive, LUTImage

Sub LoadLUT
    Dim x
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "")Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage + 1)MOD 22:UpdateLUT:SaveLUT:Lutbox.text = "level of darkness " & LUTImage + 1:End Sub

Sub UpdateLUT
    Select Case LutImage
        Case 0:Wonderland.ColorGradeImage = "LUT0"
        Case 1:Wonderland.ColorGradeImage = "LUT1"
        Case 2:Wonderland.ColorGradeImage = "LUT2"
        Case 3:Wonderland.ColorGradeImage = "LUT3"
        Case 4:Wonderland.ColorGradeImage = "LUT4"
        Case 5:Wonderland.ColorGradeImage = "LUT5"
        Case 6:Wonderland.ColorGradeImage = "LUT6"
        Case 7:Wonderland.ColorGradeImage = "LUT7"
        Case 8:Wonderland.ColorGradeImage = "LUT8"
        Case 9:Wonderland.ColorGradeImage = "LUT9"
        Case 10:Wonderland.ColorGradeImage = "LUT10"
        Case 11:Wonderland.ColorGradeImage = "LUT Warm 0"
        Case 12:Wonderland.ColorGradeImage = "LUT Warm 1"
        Case 13:Wonderland.ColorGradeImage = "LUT Warm 2"
        Case 14:Wonderland.ColorGradeImage = "LUT Warm 3"
        Case 15:Wonderland.ColorGradeImage = "LUT Warm 4"
        Case 16:Wonderland.ColorGradeImage = "LUT Warm 5"
        Case 17:Wonderland.ColorGradeImage = "LUT Warm 6"
        Case 18:Wonderland.ColorGradeImage = "LUT Warm 7"
        Case 19:Wonderland.ColorGradeImage = "LUT Warm 8"
        Case 20:Wonderland.ColorGradeImage = "LUT Warm 9"
        Case 21:Wonderland.ColorGradeImage = "LUT Warm 10"
    End Select
End Sub


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
' ninuzzu's FLIPPER SHADOWS v2
'*****************************************

'Add TimerEnabled=True to Wonderland_KeyDown procedure
' Example :
'Sub Wonderland_KeyDown(ByVal keycode)
'    If keycode = LeftFlipperKey Then
'        LeftFlipper.TimerEnabled = True 'Add this
'        LeftFlipper.RotateToEnd
'    End If
'    If keycode = RightFlipperKey Then
'        RightFlipper.TimerEnabled = True 'And add this
'        RightFlipper.RotateToEnd
'    End If
'End Sub

Sub LeftFlipper_Init()
    LeftFlipper.TimerInterval = 10
End Sub

Sub RightFlipper_Init()
    RightFlipper.TimerInterval = 10
End Sub

Sub LeftFlipper_Timer()
  LFlip.RotZ = LeftFlipper.CurrentAngle
  LFlip1.RotZ = LeftFlipper.CurrentAngle
    FlipperLSh.RotZ = LeftFlipper.CurrentAngle
    If LeftFlipper.CurrentAngle = LeftFlipper.StartAngle Then
        LeftFlipper.TimerEnabled = False
    End If
End Sub

Sub RightFlipper_Timer()
  RFlip.RotZ = RightFlipper.CurrentAngle
  RFlip1.RotZ = RightFlipper.CurrentAngle
    FlipperRSh.RotZ = RightFlipper.CurrentAngle
    If RightFlipper.CurrentAngle = RightFlipper.StartAngle Then
        RightFlipper.TimerEnabled = False
    End If
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

'Sub LeftFlipper_Collide(parm)
'   RandomSoundFlipper()
'End Sub

'Sub RightFlipper_Collide(parm)
'   RandomSoundFlipper()
'End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub
