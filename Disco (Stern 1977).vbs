Option Explicit

'*****************************************************************************************************
' CREDITS
' VPX conversion by JamieP
'
'
'
'*****************************************************************************************************
' DOF config - OUTHERE and a little by leeoneil
'
' Option for more lights effects with DOF (Undercab effects on bumpers and slingshots)
' "True" to activate (False by default)
Const Epileptikdof = False
'
' Flippers L/R - 101/102
' Slingshot L/R - 103/104
' Bumpers L/R - 105/106
' Drop Targets L/R - 107/108/109
' Ballrelease - 110
' Left Upper SlingShot - 111
' Kicker - 105
'
' LED backboard
' Flasher Outside Left - 205/207/211/215/216/219/221/240/250/260
' Flasher left - 206/210/212/215/217/241/250/260
' Flasher center - 204/206/209/212/214/215/242/250/260
' Flasher right - 204/210/214/215/218/243/250/260
' Flasher Outside - 203/208/213/215/220/222/244/250/260
' MX Flashers F/R - 204-206
' Start Button - 200
' Undercab - 201/202
' Strobe - 230
' Knocker - 300
' Chimes - 301/302/303



'First, try to load the Controller.vbs (DOF), which helps controlling additional hardware like lights, gears, knockers, bells and chimes (to increase realism)
'This table uses DOF via the 'SoundFX' calls that are inserted in some of the PlaySound commands, which will then fire an additional event, instead of just playing a sample/sound effect
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

'If using Visual PinMAME (VPM), place the ROM/game name in the constant below,
'both for VPM, and DOF to load the right DOF config from the Configtool, whether it's a VPM or an Original table

Const cGameName = "disco_Stern_1977"


Dim EnableRetractPlunger
EnableRetractPlunger = false 'Change to true to enable retracting the plunger at a linear speed; wait for 1 second at the maximum position; move back towards the resting position; nice for button/key plungers


If Table1.ShowDT = false then
    For each xx in DesktopStuff:xx.Visible = 0: Next
  For each xx in MatchDT:xx.Visible = 0: Next
End If


'########################VARS################################################

Dim a,x,xx,BIP,Credits,Players,PlayerUp,BellEnabled,maxBalls,GameStarted,TableTilted,MotorRun,MotorMult,PlayerScore(2)
Dim MatchNumber,Bonus,BonusHold,TempBonus,MadeDISC,MadeSCO,MadeDISCO
Dim Replay1,Replay2,Replay1Paid(2),Replay2Paid(2),Freeplay, Attract

Dim spinLightStep,changeRelay

Const BallSize = 25
Const BallMass = 1.7
Freeplay = 1
Attract = 0

Replay1 = 750000
Replay2 = 900000

If Freeplay = 1 Then Apron.Image = "plasticsFREEPLAY"

'########################KEYS################################################
'##########################################################################################################
dim rms,lms
Sub Table1_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAtVol"plungerpull", Plunger, 1


  If Keycode = StartGameKey and (Credits>0 or Freeplay = 1) and GameStarted=false and Players=0 and Players<2 and BIP<2 And Not HSEnterMode=true And not rms = true Then
    Credits=Credits-1
    Players = Players +1
    CreditsReel.SetValue(Credits)
    If B2SOn Then Controller.B2SSetCanPlay 31,1
    If B2SOn Then Controller.B2SSetCredits Credits
    can1Play.State = 1
    playsound "feat"
    StartGame
    if Credits = 0 then DOF 200, DOFOff end if
  ElseIf Keycode = StartGameKey and (Credits>0 or Freeplay = 1) and GameStarted=true and Players>0 and Players<2 and BIP<2 And Not HSEnterMode=true And not rms = true Then
    Players = Players +1
    Credits=Credits-1
    CreditsReel.SetValue(Credits)
    If B2SOn Then Controller.B2SSetCanPlay 31,2
    If B2SOn Then Controller.B2SSetCredits Credits
    can2Play.State=1
    playsound "feat"
    If B2SOn Then Controller.B2SSetCredits Credits
  Elseif keycode = StartGameKey And Not HSEnterMode=true And rms = true Then
    deletehs
  End If



  If keycode = LeftMagnaSave Then lms = True

  If keycode = RightMagnaSave And lms = True Then rms = True

  If HSEnterMode Then
    HighScoreProcessKey(keycode)
  End If

  If keycode = LeftFlipperKey and GameStarted = True and TableTilted = False Then
          LeftFlipper.TimerEnabled = True 'This line is only for ninuzzu's flipper shadows!
    LeftFlipper.EOSTorque = 0.75: LeftFlipper.RotateToEnd
    PlaySound SoundFXDOF ("fx_flipperup",101, DOFOn, DOFFlippers), 0, .67, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
  End If

  If keycode = RightFlipperKey and GameStarted = True and TableTilted = False Then
          RightFlipper.TimerEnabled = True 'This line is only for ninuzzu's flipper shadows!
    RightFlipper.EOSTorque = 0.75: RightFlipper.RotateToEnd
    PlaySound SoundFXDOF ("fx_flipperup",102, DOFOn, DOFFlippers), 0, .67, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
  End If
  If keycode = MechanicalTilt Then
    Tiltit
  End If
  If keycode = AddCreditKey or keycode = 4 then
    PlaySoundAtVol "coin3", Drain, 1
    AddCoin
    DOf 200, DOFOn
  end if

End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
  If keycode = RightMagnaSave Then rms = False
  If keycode = LeftMagnaSave Then lms = False

  If keycode = LeftFlipperKey Then
    LeftFlipper.EOSTorque = 0.15: LeftFlipper.RotateToStart
    PlaySound SoundFXDOF ("fx_flipperdown",101, DOFOff, DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.EOSTorque = 0.15: RightFlipper.RotateToStart
    PlaySound SoundFXDOF ("fx_flipperdown",102, DOFOff, DOFFlippers), 0, 1, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
  End If

End Sub

'########################################################################################
'########################################################################################
Sub Table1_Init()
  LoadEM
  DOF 201, DOFOn
  For each xx in GI:xx.State = 1: Next
  BIP = 0
  PlayerScore(0)=0
  PlayerScore(1)=0
  maxBalls=5
  Matchnumber = 100
  spinLightStep=0
  loadhs
  if HSA1="" then HSA1=25
  if HSA2="" then HSA2=25
  if HSA3="" then HSA3=25
  UpdatePostIt
  If Table1.ShowDT = True then
    LightGameover.State=1
    For each xx in MatchDT:xx.Visible = 0: Next
    TextBox001.Visible = False
  End If
  If B2SOn Then
    Controller.B2SSetData 1,1
    Controller.B2SSetGameOver 35,1
  End If
  AttractMode
  'For each xx in AllLights:xx.State = 1: Next
End Sub

Sub AttractMode()
  If Attract = 0 Then Exit Sub
  Select Case a
      Case 0 : AttractSeq.Play SeqUpOn,25
        Case 1 : AttractSeq.Play SeqDownOn,25
      Case 2 : AttractSeq.Play SeqRightOn,25
      Case 3 : AttractSeq.Play SeqLeftOn,25
      Case 4 : AttractSeq.Play SeqArcBottomLeftUpOn,90
      Case 5 : AttractSeq.Play SeqArcBottomLeftDownOn,90
      Case 6 : AttractSeq.Play SeqArcBottomRightUpOn,90
      Case 7 : AttractSeq.Play SeqArcBottomRightDownOn,90
      Case 8 : AttractSeq.Play SeqArcTopLeftUpOn,90
      Case 9 : AttractSeq.Play SeqArcTopLeftDownOn,90
      Case 10 : AttractSeq.Play SeqArcTopRightUpOn,90
      Case 11 : AttractSeq.Play SeqArcTopRightDownOn,90

  End Select
End Sub

Sub AttractSeq_PlayDone
  a=a+1
  If a >11 Then a = 0
  AttractMode
End Sub

Sub SpinLightAdvance(spinner)
  aSpinnerLights(spinLightStep).State = 0
  spinLightStep = spinLightStep +1
  If spinLightStep>8 Then
    spinLightStep=0
    If (spinner = 1 and LightSpinnerLeft.State = 1) or (spinner = 2 and LightSpinnerRight.State = 1) Then AdvanceBonus
  End If
  If spinLightStep = 8 Then
    If changeRelay = 1 then LightSpinnerLeft.State = 1 else LightSpinnerRight.State = 1
  Else
    LightSpinnerLeft.State = 0
    LightSpinnerRight.State = 0
  End If
  aSpinnerLights(spinLightStep).State = 1

End Sub

Sub ResetLights()
  For each xx in aDISCOhitLights:xx.State = 1: Next
  For each xx in aDISCOtriggerLights:xx.State = 1: Next
  For each xx in aDISCOlights:xx.State = 0: Next
  For each xx in aLeftLights:xx.State = 1: Next
  MadeDISC = 0
  MadeDISCO = 0
  MadeSCO = 0
  DoubleBonus.State = 0
  TripleBonus.State = 0
  LightDoubleBonus.State = 1
  LightTripleBonus.State = 0
  LightSpecialLeft.State = 0
  LightSpecialRight.State = 0
  LightLeft10k.State = 0
  If Bonus <> 40000 or Bonus <> 90000 Then
    LightExtraBallRight.State = 0
    LightExtraBallLeft.State = 0
  End If
  Change(1)
End Sub


Sub AddCoin()
  Credits=Credits+1
  DOF 200, DOFOn
  if Credits>25 then Credits=25
  If B2SOn Then Controller.B2SSetCredits Credits
  CreditsReel.SetValue(Credits)
End Sub


Sub Change(newball)
  If newball = 0 then If changeRelay = 0 then changeRelay = 1 Else changeRelay = 0

  If changeRelay = 1 Then
    LightBumperLeft.State=1
    LightBumperRight.State=1
    LightButton003.State = 1
    If MadeDISCO = 1 Then
      LightSpecialRight.State = 0
      LightSpecialLeft.State = 1
    End If

  Else
    LightBumperLeft.State=0
    LightBumperRight.State=0
    LightButton003.State = 0
    If MadeDISCO = 1 Then
      LightSpecialRight.State = 1
      LightSpecialLeft.State = 0
    End If
  End If

End Sub

Sub StartGame
    if not GameStarted Then
      ShutdownTimer.Enabled = 0
      AttractSeq.StopPlay
      If Table1.ShowDT = True then
        LightGameover.State=0
      End If
      PlayerScore(0)=0
      PlayerScore(1)=0
      For each xx in GI:xx.State = 1: Next
      If B2SOn Then
        Controller.B2SStopAnimation "Attract" 'turn off Attract
        Controller.B2SSetTilt 33,0
        Controller.B2SSetGameOver 35,0 'turn off Gameover light
        Controller.B2SSetScorePlayer 1, PlayerScore(0)
        Controller.B2SSetScorePlayer 2, PlayerScore(1)
        Controller.B2SSetPlayerUp 30, 1
        Controller.B2ssetmatch 34, 0
      End If
      BIP = 0
      PlayerUp=0
      EMReel001.setvalue(0)
      EMReel002.setvalue(0)
      For each xx in MatchDT:xx.Visible = 0: Next
      GameStarted = True
      PlaySound "Motor2"
      PlaySound "initialize"
      StartTimer.enabled = True
      TableTilted = False
      For each xx in BonusLadder:xx.State = 0: Next
      Bonus = 0
      BonusHold = False
      ShootAgain.State = 0
      Replay1Paid(0) = False
      Replay2Paid(0) = False
      Replay1Paid(1) = False
      Replay2Paid(1) = False
      DOF 201, DOFOn
      DOF 260, DOFPulse
  end If
End Sub


Sub StartTimer_Timer
  NextBall
  me.Enabled = False
End Sub

Sub NextBall()

  If ShootAgain.State = 2 Then
      BallRelease.CreateSizedBallWithMass BallSize,BallMass
      BallRelease.Kick 90, 7
      PlaySound SoundFXDOF ("ballrelease",110, DOFPulse, DOFContactors), 0,1,AudioPan(BallRelease),0.25,0,0,1,AudioFade(BallRelease)
            If Epileptikdof = True Then DOF 203, DOFPulse End If
      ResetLights
      Exit Sub
  End If
  If B2SOn Then Controller.B2SSetTilt 33,0
  If B2SOn Then Controller.B2SSetPlayerUp 30,0
  PlayerUp = PlayerUp +1
  If PlayerUp > Players Then ' Back to player 1 Up
    PlayerUp =1
    if BIP < maxBalls Then
      BIP = BIP + 1
      Player1UpLight.State=1
      Player2UpLight.State=0
      For each xx in BIPlights:xx.State = 0: Next
      BIPlights(BIP-1).State = 1
      If B2SOn Then Controller.B2SSetBallInPlay 32,BIP
      If B2SOn Then Controller.B2SSetPlayerUp 30,1
      BallRelease.CreateSizedBallWithMass BallSize,BallMass
      BallRelease.Kick 90, 7
      PlaySound SoundFXDOF ("ballrelease", 110, DOfPulse, DOFContactors), 0,1,AudioPan(BallRelease),0.25,0,0,1,AudioFade(BallRelease)
            If Epileptikdof = True Then DOF 203, DOFPulse End If
    Else
      GameOver
      For each xx in BIPlights:xx.State = 0: Next
      Exit Sub
    End If
  ElseIf PlayerUp = 1 Then  'First ball
      BIP = BIP + 1
      Player1UpLight.State=1
      Player2UpLight.State=0
      For each xx in BIPlights:xx.State = 0: Next
      BIPlights(BIP-1).State = 1
      If B2SOn Then Controller.B2SSetBallInPlay 32,BIP
      If B2SOn Then Controller.B2SSetPlayerUp 30,1
      BallRelease.CreateSizedBallWithMass BallSize,BallMass
      BallRelease.Kick 90, 7
      PlaySound SoundFXDOF ("ballrelease", 110, DOFPulse, DOFContactors), 0,1,AudioPan(BallRelease),0.25,0,0,1,AudioFade(BallRelease)
            If Epileptikdof = True Then DOF 203, DOFPulse End If

  Else ' Player 2 Up
      Player1UpLight.State=0
      Player2UpLight.State=1
      For each xx in BIPlights:xx.State = 0: Next
      BIPlights(BIP-1).State = 1
      If B2SOn Then Controller.B2SSetBallInPlay 32,BIP
      If B2SOn Then Controller.B2SSetPlayerUp 30,2
      BallRelease.CreateSizedBallWithMass BallSize,BallMass

      BallRelease.Kick 90, 7
      PlaySound SoundFXDOF ("ballrelease",110, DOFPulse, DOFContactors), 0,1,AudioPan(BallRelease),0.25,0,0,1,AudioFade(BallRelease)
            If Epileptikdof = True Then DOF 203, DOFPulse End If
  End If
  ResetLights
End Sub

Sub BallReleaseTimer_Timer()
  Me.Enabled=False
  NextBall()
End Sub

Sub CheckReplay()
  If PlayerScore(PlayerUp-1)=>Replay1 and Replay1Paid(PlayerUp-1)=false then
    Replay1Paid(PlayerUp-1)=True
    If Freeplay = 1 Then :AddShootAgain:Else AddReplay :End If
  End If
  If PlayerScore(PlayerUp-1)=>Replay2 and Replay2Paid(PlayerUp-1)=false then
    Replay2Paid(PlayerUp-1)=True
    If Freeplay = 1 Then :AddShootAgain:Else AddReplay :End If
  End If
End Sub

Sub AddShootAgain
  If ShootAgain.State=1 Then Exit Sub
  ShootAgain.State=1
  If B2SOn Then Controller.B2SSetShootAgain 36, 1
  PlaySound SoundFXDOF ("knocker", 300, DOFPulse, DOFKnocker)
  DOF 230, DOFPulse
End Sub

Sub DisplayScores()
  If B2SOn Then Controller.B2SSetScorePlayer PlayerUp, PlayerScore(PlayerUp-1)
  If Table1.ShowDT = True then
    EMReel001.setvalue(PlayerScore(0))
    EMReel002.setvalue(PlayerScore(1))
  End If
End Sub

Sub Addpoints(ScorePar)
  If TableTilted = True Then Exit Sub
  Select Case ScorePar
    Case 100
      PlayerScore(PlayerUp-1) = PlayerScore(PlayerUp-1) + ScorePar
      DisplayScores
      PlaySound SoundFXDOF ("Bell100", 301, DOFPulse, DOFChimes)
      CheckReplay
    Case 1000
      PlayerScore(PlayerUp-1) = PlayerScore(PlayerUp-1) + ScorePar
      DisplayScores
      PlaySound SoundFXDOF ("Bell1000", 302, DOFPulse, DOFChimes)
      CheckReplay
    Case 10000
      PlayerScore(PlayerUp-1) = PlayerScore(PlayerUp-1) + ScorePar
      DisplayScores
      PlaySound SoundFXDOF ("Bell10000", 303, DOFPulse, DOFChimes)
      CheckReplay
    Case Else
      PlaySound ("scoremtr")
      MotorRun = ScorePar\100 : MotorMult = 100
      If ScorePar>1000 Then MotorRun = ScorePar\1000 : MotorMult = 1000
      If ScorePar>10000 Then MotorRun = ScorePar\10000 : MotorMult = 10000
      ScoreMotor.Enabled = True
  End Select
            MatchNumber=MatchNumber+100
            If MatchNumber=1000 Then MatchNumber=0
End Sub

Sub ScoreMotor_Timer()
  If MotorRun>0 Then
      PlayerScore(PlayerUp-1) = PlayerScore(PlayerUp-1) + MotorMult
      DisplayScores
      PlaySound "Bell"&MotorMult
      MotorRun = MotorRun -1
      MatchNumber=MatchNumber+100
      If MatchNumber=1000 Then MatchNumber=0
      CheckReplay
  Else
    If TempBonus = 1 Then BonusTimer.Enabled = True
    Me.Enabled = False
  End If
End Sub

Sub AdvanceBonus()
  If Bonus > 0 Then BonusLadder((Bonus\10000)-1).State = 0
  Bonus = Bonus +10000
  If Bonus >100000 then Bonus = 100000
  BonusLadder((Bonus\10000)-1).State = 1
  If Bonus = 40000 or Bonus = 90000 and ShootAgain.State = 0 Then
    If changeRelay = 1 then LightExtraBallRight.State = 1 Else LightExtraBallLeft.State = 1
  Else
    LightExtraBallRight.State = 0
    LightExtraBallLeft.State = 0
  End If
End Sub

Sub BonusTimer_Timer()
  If TableTilted = True Then Exit Sub
    TempBonus = 1
    If Bonus>0 Then
      If ShootAgain.State = 1 Then ShootAgain.State = 2
      BonusLadder((Bonus\10000)-1).State = 0
      If Bonus > 10000 Then BonusLadder((Bonus\10000)-2).State = 1
      Bonus = Bonus - 10000
      If TripleBonus.State = 1 Then
        Addpoints 30000
        Me.Enabled = False
      ElseIf DoubleBonus.State = 1 Then
        Addpoints 20000
        Me.Enabled = False
      Else
        Addpoints 10000
      End If

    Else
      BallReleaseTimer.Enabled=True
      PlaySound "scoremtr"
      TempBonus = 0
      Me.Enabled = False
    End If
End Sub



Sub CheckDisco()
  'check for DISC
  If aDISCOlights(0).State=1 and aDISCOlights(1).State=1 and aDISCOlights(2).State=1 and aDISCOlights(3).State=1 Then
    MadeDISC = 1
    If DoubleBonus.State = 1 then LightTripleBonus.State = 1
  End If
  'check for SCO
  If aDISCOlights(2).State=1 and aDISCOlights(3).State=1 and aDISCOlights(4).State=1 Then
    MadeSCO = 1
    LightLeft10k.State = 1
  End If
  'check for DISCO
  If aDISCOlights(0).State=1 and aDISCOlights(1).State=1 and aDISCOlights(2).State=1 and aDISCOlights(3).State=1 and aDISCOlights(4).State=1 Then
    If changeRelay = 1 Then LightSpecialLeft.State = 1 Else LightSpecialRight.State = 1
    MadeDISCO = 1
  Else
    LightSpecialLeft.State = 0
    LightSpecialRight.State = 0
  End If
End Sub

Sub AddReplay()
  AddCoin
  PlaySound SoundFXDOF ("knocker", 300, DOFPulse, DOFKnocker)
    DOF 230, DOFPulse
End Sub

Sub Match()
    If B2SOn Then
    If MatchNumber = 0 Then
      Controller.B2ssetmatch 34, 10
    Else Controller.B2ssetmatch 34, MatchNumber/100
    End If
  End If
  If Table1.ShowDT = True then MatchDT(MatchNumber/100).Visible = 1
    If MatchNumber= INT(PlayerScore(0) Mod 1000) Then AddReplay
    If Players = 2 Then If MatchNumber= INT(PlayerScore(1) Mod 1000) Then AddReplay
End Sub

Sub Tiltit()
  If BIP>0 Then
    PlaySound "Tilt"
    DOF 101, DOFOff
        DOF 102, DOFOff
    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart
    For each xx in GI:xx.State = 0: Next
    For each xx in BonusLadder:xx.State = 0: Next
    For each xx in Bumpers:xx.HasHitEvent = 0: Next
    For each xx in Slingshots:xx.HasHitEvent = 0: Next
    ResetLights
    ShootAgain.State = 0
    If B2SOn Then
      Controller.B2SSetPlayerUp 30,0
      Controller.B2SSetTilt 33,1
      Controller.B2SSetData 1,0
      Controller.B2SSetShootAgain 36, 0
    End If
    TableTilted = True
    If Table1.ShowDT = True Then TextBox001.Visible = True
  End If
End Sub

Sub TiltTimer_Timer
    TiltTimer.Enabled = 0
    TableTilted = False
    For each xx in GI:xx.State = 1: Next
    For each xx in Bumpers:xx.HasHitEvent = 1: Next
    For each xx in Slingshots:xx.HasHitEvent = 1: Next
    For each xx in BonusLadder:xx.State = 0: Next
    Bonus = 0
    If Table1.ShowDT = True Then TextBox001.Visible = False
    If B2SOn Then Controller.B2SSetData 1,1
    BallReleaseTimer.Enabled=True
    PlaySound "scoremtr"
End Sub

Sub GameOver
  GameStarted = False
    If Freeplay = 0 Then Match
  If B2SOn Then Controller.B2SSetGameOver 35,1 'turn on Gameover light
  If Table1.ShowDT = True then LightGameover.State=1
  If PlayerScore(1) > PlayerScore(0) Then
    CheckNewHighScorePostIt(PlayerScore(1))
      Player1UpLight.State=0
      Player2UpLight.State=1
      If B2SOn Then Controller.B2SSetPlayerUp 30,2
  Else
    CheckNewHighScorePostIt(PlayerScore(0))
      Player1UpLight.State=1
      Player2UpLight.State=0
      If B2SOn Then Controller.B2SSetPlayerUp 30,1
  End If
  Players=0
  BIP = 0
    If B2SOn Then Controller.B2SSetBallInPlay 32,0
  can1Play.state=0
  can2Play.state=0
  If B2SOn Then Controller.B2SSetCanPlay 31,0
  DOF 260, DOfPulse
End Sub

Sub PowerDown
  PlaySound "endgame"
  ShutdownTimer.Enabled = 1
  'For each xx in AllLights:xx.State = 1: Next
  If Attract = 1 Then AttractSeq.Play SeqDownOn,25
End Sub

Sub ShutdownTimer_timer
  ResetLights
  AttractMode
  If B2SOn Then Controller.B2SStartAnimation "Attract" 'turn on Attract
  If B2SOn Then Controller.B2SSetPlayerUp 30, 0
  DOF 201, DOFOff
  ShutdownTimer.Enabled = 0
End Sub


'******************************
'     DOF lights ball entrance
'******************************
'
Sub Trigger009_Hit
    DOF 250, DOFPulse
End sub



'#########################################################################################
Sub Trigger008_Hit()
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  ShootAgain.State = 0
  If B2SOn Then Controller.B2SSetShootAgain 36, 0
End Sub

Sub Trigger012_Hit()
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  DOF 213, DOFPulse
  IF LightButton003.State = 1 Then Addpoints(1000) Else Addpoints(100)
End Sub

Sub Drain_Hit()
  PlaySound "drain",0,1,AudioPan(Drain),0.25,0,0,1,AudioFade(Drain)
    DOF 215, DOFPulse
  Drain.DestroyBall
  If TableTilted = True Then
    TiltTimer.Enabled = 1
  Else
    BonusHold = False
    If ShootAgain.State = 1 Then ShootAgain.State = 2
    BonusTimer.Enabled = True
  End If
End Sub

Sub aDISCOtargets_Hit(idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    DOF 240+idx, DOFPulse
  aDISCOhitLights(idx).State = 0
  If idx <4 Then aDISCOtriggerLights(idx).State = 0
  aDISCOlights(idx).State = 1
  Addpoints 1000
  CheckDisco
End Sub

Sub aDISCOtriggers_Hit(idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    DOF 270+idx, DOFPulse
  aDISCOtriggerLights(idx).State = 0
  aDISCOhitLights(idx).State = 0
  aDISCOlights(idx).State = 1
  If idx = 2 Then
    Addpoints 5000
    If LightExtraBallLeft.State = 1 then
      ShootAgain.State = 1
      PlaySound SoundFXDOF ("knocker", 300, DOFPulse, DOFKnocker)
    End If
  Elseif idx = 3 Then
    Addpoints 5000
    If LightExtraBallRight.State = 1 then
      ShootAgain.State = 1
      PlaySound SoundFXDOF ("knocker", 300, DOFPulse, DOFKnocker)
    End If
  Else Addpoints 1000
  End If
  CheckDisco
End Sub

Sub aLeftTriggers_Hit(idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  If LightLeft10k.State=1 Then Addpoints 10000 Else Addpoints 1000
    DOF 280+idx, DOFPulse
  aLeftLights(idx).State = 0: aLeftTriggers(idx).TimerEnabled = 1
End Sub

Sub aLeftTriggers_Timer(idx)
  aLeftLights(idx).State = 1
  aLeftTriggers(idx).TimerEnabled = 0
End Sub

Sub aAdvanceTriggers_Hit(idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  AdvanceBonus
  Change(0)
  PlaySound SoundFXDOF ("Bell1000", 303, DOFPulse, DOFChimes)
  aLightButtons(idx).State = 0 : aAdvanceTriggers(idx).TimerEnabled = 1

End Sub

Sub aAdvanceTriggers_Timer(idx)
    aLightButtons(idx).State = 1
  aAdvanceTriggers(idx).TimerEnabled = 0
End Sub

Sub TargetSTAR_Hit()
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  Addpoints 1000
  If LightDoubleBonus.State = 1 Then
    DoubleBonus.State = 1
    LightDoubleBonus.State = 0
  Elseif LightTripleBonus.State = 1 Then
    TripleBonus.State = 1
    LightTripleBonus.State = 0
    DoubleBonus.State = 0
  End If
  CheckDisco
End Sub

Sub Spinner001_Spin()
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner001), 0.25, 0, 0, 1, AudioFade(Spinner001)
    DOF 216, DOFPulse
  SpinLightAdvance(1)
  Addpoints 100
End Sub

Sub Spinner002_Spin()
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner002), 0.25, 0, 0, 1, AudioFade(Spinner002)
    DOF 217, DOFPulse
  SpinLightAdvance(2)
  Addpoints 100
End Sub


Sub LeftOutlane_Hit()
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  DOF 219, DOfPulse
  If LightSpecialLeft.State = 1 Then
    If Freeplay = 1 Then ShootAgain.State = 1 Else AddReplay
    PlaySound SoundFXDOF ("knocker", 300, DOFPulse, DOFKnocker)
  Else
    Addpoints 10000
  End If
End Sub

Sub RightOutlane_Hit()
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  DOF 220, DOfPulse
  If LightSpecialRight.State = 1 Then
    If Freeplay = 1 Then ShootAgain.State = 1 Else AddReplay
    PlaySound SoundFXDOF ("knocker", 300, DOFPulse, DOFKnocker)
  Else
    Addpoints 10000
  End If
End Sub

Sub ScoringRubbers_Hit(idx)
  RandomSoundRubber()
  Addpoints 100
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************

Dim RStep, Lstep, LUstep

Sub RightSlingShot_Slingshot
  If TableTilted = True Then Exit Sub
    PlaySoundAtVol SoundFXDOF ("right_slingshot",104,DOFPulse,DOFcontactors), ActiveBall, 1
    DOF 203, DOFPulse
    If Epileptikdof = True Then DOF 206, DOFPulse End If
    If Epileptikdof = True Then DOF 202, DOFPulse End If
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  Addpoints 100
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  If TableTilted = True Then Exit Sub
    PlaySoundAtVol SoundFXDOF ("left_slingshot",103,DOFPulse,DOFcontactors), ActiveBall, 1
    DOF 205, DOFPulse
    If Epileptikdof = True Then DOF 204, DOFPulse End If
    If Epileptikdof = True Then DOF 202, DOFPulse End If
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  Addpoints 100
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub LeftUpperSlingShot_Slingshot
  If TableTilted = True Then Exit Sub
    PlaySoundAtVol SoundFXDOF ("left_slingshot",111,DOFPulse,DOFcontactors), ActiveBall, 1
    DOF 205, DOFPulse
    If Epileptikdof = True Then DOF 204, DOFPulse End If
    If Epileptikdof = True Then DOF 202, DOFPulse End If
    LUSling.Visible = 0
    LUSling1.Visible = 1
    sling3.rotx = 20
    LUStep = 0
    LeftUpperSlingShot.TimerEnabled = 1
  Addpoints 100
End Sub

Sub LeftUpperSlingShot_Timer
    Select Case LUStep
        Case 3:LUSLing1.Visible = 0:LUSLing2.Visible = 1:sling3.rotx = 10
        Case 4:LUSLing2.Visible = 0:LUSLing.Visible = 1:sling3.rotx = 0:LeftUpperSlingShot.TimerEnabled = 0
    End Select
    LUStep = LUStep + 1
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


Dim TableWidth, TableHeight
TableWidth = Table1.width
TableHeight = Table1.height

'********************************************************************
'      JP's VP10 Rolling Sounds (+ JamieP Ballshadows +rothbauerw's Dropping Sounds)
'********************************************************************

Const tnob = 4 ' total number of balls
ReDim rolling(tnob)
Dim shadowm,shadowc
shadowc=Ballsize*0.75 'distance of shadow from ball at the furthest point
shadowm= ((TableWidth/2)+shadowc)/(TableWidth/2) 'slope of the linear change around the centerpoint of the table
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b,speedfactorx,speedfactory
  Const maxvel = 35 'max ball velocity
   BOT = GetBalls
    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    aBallShadow(b).visible = 0
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

    For b = 0 to UBound(BOT)
        aBallShadow(b).X = shadowm*BOT(b).X - shadowc'y=mx+c
    aBallShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            aBallShadow(b).visible = 1
        Else
            aBallShadow(b).visible = 0
        End If
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

End Sub

'*****************************************************************************************

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

Sub Rails_Hit (idx)
  PlaySound "woodhit", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

Sub Bumpers_Hit(idx)
  If idx = 0 Then
    If LightBumperLeft.State=1 Then Addpoints 1000 else Addpoints 100
    DOF 105, DOFPulse : DOF 207, DOFPulse
        If Epileptikdof = True Then DOF 202, DOFPulse End If
  End If
  If idx = 1 Then
    If LightBumperRight.State=1 Then Addpoints 1000 else Addpoints 100
    DOF 106, DOFPulse : DOF 208, DOFPulse
        If Epileptikdof = True Then DOF 202, DOFPulse End If
  End If
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySound "fx_bumper1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "fx_bumper2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "fx_bumper3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 4 : PlaySound "fx_bumper4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
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


Sub CheckNewHighScorePostIt(newScore)
    If CLng(newScore) > CLng(HSAHighScore) Then
      HSAHighScore=newScore
      HighScoreEntryInit()
    Else
      PowerDown
    End If
End Sub


Dim HSA1, HSA2, HSA3,HSAHighScore
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim HSArray
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex 'Define 6 different score values for each reel to use
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")
Const hsFlashDelay = 4
Const HighScoreFilename = "Disco.txt"
Dim FileObj,ScoreFile
sub savehs
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HighScoreFilename,True)
    ScoreFile.WriteLine HSAHighScore
    ScoreFile.WriteLine HSA1
    ScoreFile.WriteLine HSA2
    ScoreFile.WriteLine HSA3
    ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub

sub loadhs
  Dim FileObj
  Dim ScoreFile
  Dim TextStr
    Dim SavedDataTempScore, SavedDataTempHSA1, SavedDataTempHSA2, SavedDataTempHSA3 'HighScore
    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & HighScoreFilename) then
    SetDefaultHSTD:UpdatePostIt:savehs
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & HighScoreFilename)
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      SetDefaultHSTD:UpdatePostIt:savehs
      Exit Sub
    End if
    SavedDataTempScore=Textstr.ReadLine ' HighScore
    HSA1=Textstr.ReadLine ' HighScore
    HSA2=Textstr.ReadLine ' HighScore
    HSA3=Textstr.ReadLine ' HighScore
    TextStr.Close
    HSAHighScore=SavedDataTempScore
    UpdatePostIt
      Set ScoreFile = Nothing
      Set FileObj = Nothing
end sub

sub deletehs
  debug.print "delete HS file"
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If FileObj.FileExists(UserDirectory & HighScoreFilename) then
    SetDefaultHSTD:UpdatePostIt:savehs
    Exit Sub
  End if

end Sub
Sub SetDefaultHSTD  'bad data or missing file - reset and resave
  HSAHighScore = 10000
  HSA1 = 1
  HSA2 = 1
  HSA3 = 1
End Sub


' ***********************************************************
'  HiScore DISPLAY
' ***********************************************************

Sub UpdatePostIt
  dim tempscore
  HSScorex = HSAHighScore
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
  PScore6.image = HSArray(HSScoreM):If HSScorex<1000000 Then PScore6.image = HSArray(10)
  PScore5.image = HSArray(HSScore100K):If HSScorex<100000 Then PScore5.image = HSArray(10)
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
'   for each objekt in hiscname:objekt.visible=1:next
    HSName1.image = ImgFromCode(HSA1, 1)
    HSName2.image = ImgFromCode(HSA2, 2)
    HSName3.image = ImgFromCode(HSA3, 3)
'   else
'   for each objekt in hiscname:objekt.visible=0:next
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
  PlaySound SoundFXDOF ("knocker", 300, DOFPulse, DOFKnocker)
    DOF 230, DOFPulse
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
    PlaySound SoundFX("solon",DOFContactors)
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
    PlaySound SoundFX("soloff",DOFContactors)
  End If

    If (keycode = StartGameKey or keycode = PlungerKey) Then
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
          PowerDown
        End If
    End Select
    UpdatePostIt
    PlaySound SoundFX("Bell10",DOFContactors)
    End If
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub
