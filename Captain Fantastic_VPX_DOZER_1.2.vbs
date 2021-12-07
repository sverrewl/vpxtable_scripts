' **********************
' 2016: directb2s / DOF
' by STAT, stefanaustria
' by arngrim
' **********************

'---------------------
'-  Script by ROSVE  - Score System by STAT
'---------------------

'Additional Code by Dozer.
' Thalamus: added Chimes not included in table - see :
' https://www.vpforums.org/index.php?showtopic=40334&p=410744
' Thanks Carny
Randomize

Dim Ballsize,BallMass
BallSize = 50
BallMass = (Ballsize^3)/125000

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "captfantastic"

'//Game Options/////////////////

Const GI_Dim = 1 'Dim GI with flipper current draw.
Const Lamp_Flicker = 1 'Flicker lamp under right slingshot plastic.
Const Loose_Plastic = 1 'Simulate loose left hand slingshot plastic.
Const BallsToPlay = 3 'Set to 3 or 5

'///////////////////////////////
'Drop Target Special / Lane Adjustment plugs.
'See manual "Feature Operation & Scoring" for
'overview of these settings (Page 7) or reference
'included image in archive.

'1 = Conservative / 0 = Liberal

sadjust = 1 'Special Plug Adjust
ladjust = 1 'Lane Plug Adjust

'///////////////////////////////

 Const NudgeOn = 1
 Dim robject, sadjust, ladjust
 Dim Renudge,RenudgeForce
 Dim roll
 Dim ball,ball1
 Dim GameSeq
 Dim MaxPlayers
 Dim ActivePlayer
 Dim Round
 Dim BallInPlay
 Dim Coins
 Dim Tilt
 Dim FreePlayNum
 Dim DoubleBonusFlag
 Dim CollectL
 Dim HiScore
 Dim TopLights,Targets,DoodleBug,InPlunger
 Dim mhole
 Dim Mem(30)
 Dim Alt10p
 Dim bump1,bump2,bump3,bump4
 Dim NOB
 Dim ft
 Dim mhole1,mhole2
 Dim PRpull,PRdir
 Dim SReels(4)

 Dim STATScores(4)

 Sub captfant_Init()
  LoadEM
    HiScore=""
    On Error Resume Next
  HiScore=Cdbl(LoadValue("captfant","DBHiScore"))
  If HiScore="" Then HiScore=0
    Coins=0
    GameSeq=0  'Game Over
  If B2SOn Then Controller.B2SSetGameover 1
    GameOverBG.State = 1
    ResetBumperRings
    For i=1 to 3
       SetBumperOff i
    Next
    MaxPlayers=0
    CloseGate
    InPlunger=0
    FLFlash 1, FLMaxLight, 0, 1, 1
    ResetMemory 1
    AltLights
    NOB=0
    Set SReels(1) = ScoreReel1
    Set SReels(2) = ScoreReel2
    Set SReels(3) = ScoreReel3
    Set SReels(4) = ScoreReel4
    DOF 165,1
    anidown = 1
   End Sub


 Sub InitGame()
    UpdateBIP
    For i=1 to 3
       SetBumperOff i
    Next
    DT1.IsDropped=False
    DT2.IsDropped=False
    DT3.IsDropped=False
    DT4.IsDropped=False
    DT5.IsDropped=False
    PlaySoundat SoundFXDOF("TargetBankReset",135,DOFPulse,DOFContactors),BulbFil13

  If B2SOn Then Controller.B2SSetGameover 0
    GameOverBG.State = 0
    CloseGate
    FLSetOff 1,FLMaxLight
    LDT15.State = 1
    LDT16.State = 1
    ResetMemory 0
    'Turn on Shoot Again light if mem(2)=1
    If mem(2)=1 Then FLSetOn 17,17
    AltLights
    AddBonus 1
    FLSetOn 19,20
 End Sub

Dim object

    If ShowDT = True Then
    For each object in DT_Stuff
    Object.visible = 1
    Next
        For each object in Cab_Stuff
        Object.visible = 0
        Next
        CabinetRailLeft.visible = 1:CabinetRailRight.visible = 1
        Card1.Y = 1856:Card2.Y = 1856
  End If

  If ShowDt = False Then
    For each object in DT_Stuff
    Object.visible = 0
    Next
        For each object in Cab_Stuff
        Object.visible = 1
        Next
        CabinetRailLeft.visible = 0:CabinetRailRight.visible = 0
        'LDB.visible = 0
  End If

    If BallsToPlay = 3 Then
    Card2.image = "ic3"
    Else
    Card2.image = "ic5"
    End If

    If Lamp_Flicker = 1 Then
    GI_Flicker.enabled = 1
    Else
    GI_Flicker.enabled = 0
    End If

Dim BOOT1,BOOT2,BOOT3,BOOT4, REPA(10), REPB(10), REPC(10)

Sub AddScoresToPlayer(s,p)    ' * Score System by STAT
  If s = 10 Then
    PlaySound SoundFXDOF("10a",141,DOFPulse,DOFChimes)
  ElseIf s = 100 Then
    PlaySound SoundFXDOF("100a",142,DOFPulse,DOFChimes)
  ElseIf s = 200 Then
    PlaySound SoundFXDOF("200a",142,DOFPulse,DOFChimes)
  ElseIf s = 300 Then
    PlaySound SoundFXDOF("300a",142,DOFPulse,DOFChimes)
  ElseIf s = 500 Then
    PlaySound SoundFXDOF("500a",142,DOFPulse,DOFChimes)
  ElseIf s = 1000 Then
    PlaySound SoundFXDOF("1000a",143,DOFPulse,DOFChimes)
  ElseIf s = 3000 Then
    PlaySound SoundFXDOF("3000a",143,DOFPulse,DOFChimes)
  End If
  STATScores(p) = STATScores(p) + s
    If B2SOn Then
  Controller.B2SSetScorePlayer p, STATScores(p)
    End If
     'Handle Score based replays.
     If STATScores(p) > 70000 AND repa(p) = 0 Then
        ReplayAction
        repa(p) = 1
     End If

     If STATScores(p) > 99000 AND repb(p) = 0 Then
        ReplayAction
        repb(p) = 1
     End If

     If STATScores(p) > 131000 AND repc(p) = 0 Then
        ReplayAction
        repc(p) = 1
     End If

  If STATScores(p) > 99999 Then
        If B2SOn Then
    Controller.B2SSetData 149+p,1
        End If
        If p=1 AND boot1 = 0 Then
        PlaySound "Buzzer"
        OTT1.State = 1
        boot1 = 1
        End If
        If p = 2 AND boot2 = 0 Then
        PlaySound "Buzzer"
        OTT2.State = 1
        boot2 = 1
        End If
        If p=3 AND boot3 = 0 Then
        PlaySound "Buzzer"
        OTT3.State = 1
        boot3 = 1
        End If
        If p=4 AND boot4 = 0 Then
        PlaySound "Buzzer"
        OTT4.State = 1
        boot4 = 1
        End If

    End if
   If GameSeq = 1 Then
    SReels(p).addvalue(s)
   End If
End Sub


 Sub captfant_KeyDown(ByVal keycode)
      If keycode = AddCreditKey Then
         If Coins < 9 Then
       Coins=Coins+1
    If B2SOn Then Controller.B2SSetCredits Coins
    End If
        PlaySound "Coin"
  End If
      If keycode = 5 Then
         If Coins < 9 Then
       Coins=Coins+1
    If B2SOn Then Controller.B2SSetCredits Coins
    End If
        PlaySound "Coin"
  End If
     If keycode = StartGameKey Then
         If MaxPlayers < 4 And Coins>0 Then
       MaxPlayers=MaxPlayers+1
            AddPlayer
        End If
  End If

    If KeyCode = LeftMagnaSave Then
    If anidown = 1 Then
    anidir = 1:IC_Ani.enabled = 1:PlaySound "TargetDrop"
    Else
    anidir = 2:IC_Ani.enabled = 1:PlaySound "TargetDrop"
    End If
    End If

  If keycode = PlungerKey Then
    Plunger.PullBack
         PlaySound "SSPlungePull"
         PRdir=0
         PRLight.TimerInterval=60
         PRLight.TimerEnabled=true
  End If

  If keycode = LeftFlipperKey Then
         If Tilt<3 And Round>0 Then
       Flipper2.RotateToEnd
           Flipper4.RotateToEnd
           PlaySoundat SoundFXDOF("SSFlipper1",101,DOFOn,DOFFlippers), flipper2
         End If
        If Gi_Dim = 1 AND GameSeq=1 AND Tilt<3 Then
        For Each gxx in GI_Lights
        gxx.IntensityScale = 0.9
        Next
        DOF 166,1:DOF 165,0
        gi_bright.enabled = 1
        End If
  End If

  If keycode = RightFlipperKey Then
         If Tilt<3 And Round>0 Then
       Flipper3.RotateToEnd
       Flipper5.RotateToEnd
           PlaySoundat SoundFXDOF("SSFlipper1",102,DOFOn,DOFFlippers), flipper3
         End If
         If Gi_Dim = 1 AND GameSeq=1 AND Tilt<3 Then
        For Each gxx in GI_Lights
        gxx.IntensityScale = 0.9
        Next
        DOF 166,1:DOF 165,0
        gi_bright.enabled = 1
        End If
  End If

    If keycode = MechanicalTilt Then
    Tilt = 3
    TiltOn
    PlaySound "buzzer"
    End If

  If keycode = LeftTiltKey Then
      PlaySound SoundFX("nudge_left",0)
      If BallInPlay=1 Then
       If Tilt<3 Then
        'If NudgeOn=1 Then
     '  Nudge 90, 1.2
          ' RenudgeForce=1.0
       ' Else
        '   Nudge 90,0.7
         '  RenudgeForce=0.4
        'End If
        'Renudge=-90

         Nudge 90, 2

          Tilt2Timer.Enabled=1
         If Round>0 Then
           Tilt=Tilt+0.6
           If Tilt>2.9 Then
              TiltOn
           Else
             TiltTimer.Enabled=1
           End If
         End If
       End If
      End If
  End If

  If keycode = RightTiltKey Then
      PlaySound SoundFX("nudge_right",0)
      If BallInPlay=1 Then
       If Tilt<3 Then
        'If NudgeOn=1 Then
     '  Nudge 270, 1.2
          ' RenudgeForce=1.0
        'Else
         '  Nudge 270,0.7
          ' RenudgeForce=0.4
        'End If
        'Renudge=90

    Nudge 270, 2

          Tilt2Timer.Enabled=1
         If Round>0 Then
           Tilt=Tilt+0.6
           If Tilt>2.9 Then
              TiltOn
           Else
             TiltTimer.Enabled=1
           End If
         End If
       End If
      End If
  End If

  If keycode = CenterTiltKey Then
      PlaySound SoundFX("nudge_forward",0)
      If BallInPlay=1 Then
       If Tilt<3 Then
    'Renudge=160+(Rnd*40)
          Tilt2Timer.Enabled=1
        'If NudgeOn=1 Then
     '  Nudge -Renudge, 1.3
          ' RenudgeForce=1.0
        'Else
         '  Nudge -Renudge,0.9
          ' RenudgeForce=0.4
        'End If

         Nudge 0, 2

         If Round>0 Then
           Tilt=Tilt+0.6
           If Tilt>2.9 Then
              TiltOn
           Else
             TiltTimer.Enabled=1
           End If
         End If
       End If
      End If
  End If

End Sub

Sub captfant_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
        PlaySound "SSPlungeRel"
        PRLight.TimerInterval=10
        PRdir=1
        If InPlunger=1 Then
           PlaySound "Shoot"
        End If
  End If

  If keycode = LeftFlipperKey Then
     Flipper2.RotateToStart
       Flipper4.RotateToStart
       If Tilt<3 And Round>0 Then
        PlaySoundat SoundFXDOF("SSFlipper2",101,DOFOff,DOFFlippers), flipper2
       End If
  End If

  If keycode = RightFlipperKey Then
     Flipper3.RotateToStart
     Flipper5.RotateToStart
       If Tilt<3 And Round>0 Then
        PlaySoundat SoundFXDOF("SSFlipper2",102,DOFOff,DOFFlippers), flipper3
       End If
  End If

End Sub

Sub Drain_Hit()
     PlaySoundat "SSDrain", Drain
   DOF 133, DOFPulse
     BallInPlay=0
   Drain.DestroyBall
     If Tilt<3 Then
     GameSeq=3 'Collect Bonus then start new ball
     Else
     GameSeq=2
     End If
     SeqTimer.Enabled=1
     SeqTimer.Interval=250
End Sub

Sub ShooterLane_Hit()
  DOF 134, DOFPulse
End Sub

 '----------------------------------------------------------
 '     TILT
 '-----------------------------------------------------------

 Sub TiltTimer_Timer()
    TiltTimer.Enabled=0
    If Tilt<3 Then
      Tilt=0
    End If
 End Sub

  Sub Tilt2Timer_Timer()
     Tilt2Timer.Enabled=0
     Nudge renudge,RenudgeForce
  End Sub

 Sub TiltOn()

  If B2SOn Then Controller.B2SSetTilt 1
    PlaySound SoundFX("solon",0)
    For i=1 to 3
       SetBumperOff i
    Next
    CloseGate
    Bumper1.force = 0:Bumper2.Force = 0:Bumper3.Force = 0
    LeftSlingShot.SlingshotThreshold = 100:RightSlingShot.SlingshotThreshold = 100
    For each Object in Bonus_Lights
    object.state = 0
    next
    If B2Son Then
    PlaySoundat SoundFXDOF("SSFlipper2",101,DOFOff,DOFFlippers), flipper2
    PlaySoundat SoundFXDOF("SSFlipper2",102,DOFOff,DOFFlippers), flipper3
     Flipper2.RotateToStart
     Flipper4.RotateToStart
     Flipper3.RotateToStart
     Flipper5.RotateToStart
    Else
    PlaySoundat SoundFX("SSFlipper2",DOFFlippers), flipper2
    PlaySoundat SoundFX("SSFlipper2",DOFFlippers), flipper3
     Flipper2.RotateToStart
     Flipper4.RotateToStart
     Flipper3.RotateToStart
     Flipper5.RotateToStart
    'DOF 101,0
    'DOF 102,0
     End If
 End Sub

 '----------------------------------------
  'Add Players and start game
 '----------------------------------------
 Sub AddPlayer()
   Dim sr
  If Round=0 Or Round=3 Or Round=5 Then
    If Coins>0 Then
     Coins=Coins-1
  If B2SOn Then
    Controller.B2SSetData 28,Coins
    Controller.B2SSetCredits Coins
    Controller.B2SSetCanPlay MaxPlayers
    Controller.B2SSetMatch 0
    Controller.B2SSetGameover 0
    STATScores(1) = 0
    STATScores(2) = 0
    STATScores(3) = 0
    STATScores(4) = 0
    Controller.B2SSetScorePlayer1 0
    Controller.B2SSetScorePlayer2 0
    Controller.B2SSetScorePlayer3 0
    Controller.B2SSetScorePlayer4 0
    Controller.B2SSetData 150,0:BOOT1=0:OTT1.STATE = 0
    Controller.B2SSetData 151,0:BOOT2=0:OTT2.STATE = 0
    Controller.B2SSetData 152,0:BOOT3=0:OTT3.STATE = 0
    Controller.B2SSetData 153,0:BOOT4=0:OTT4.STATE = 0
  End If
        BOOT1=0:OTT1.STATE = 0
        BOOT2=0:OTT2.STATE = 0
    BOOT3=0:OTT3.STATE = 0
        BOOT4=0:OTT4.STATE = 0
        SReels(1).setvalue(0)
        SReels(2).setvalue(0)
        SReels(3).setvalue(0)
        SReels(4).setvalue(0)
        STATScores(1) = 0
    STATScores(2) = 0
    STATScores(3) = 0
    STATScores(4) = 0
        GameOverBG.State = 0
        MatchBG.State = 0
        TextBox2.text = ""
    if MaxPlayers = 1 Then
         Round=BallsToPlay
         ActivePlayer=0
         GameSeq=2
         SeqTimer.Interval=800
        SeqTimer.Enabled=1
     End If
   PlaySound "bally-addplayer"
     End If
   End If
 End Sub


 '------------------------------------------------------------------
 '    Game Sequence
 '------------------------------------------------------------------
 Sub SeqTimer_Timer()
    Select Case GameSeq
       Case 0 'Game Over
          ' Attract Mode
       Case 1 'New Ball
              SeqTimer.Enabled=0
              If LightShootAgain.State=0 Then InitGame
              LightShootAgain.State=0
        If B2SOn Then Controller.B2SSetShootAgain 0
              Set ball = Kicker2.CreateBall
              BallInPlay=1
              Kicker2.Kick 45,14
              PlaySoundat SoundFXDOF("ballrelease",109,DOFPulse,DOFContactors), kicker2
       Case 2 'Drain
              'Prepare for next ball
        If B2SOn Then Controller.B2SSetTilt 0
              Bumper1.force = 7:Bumper2.Force = 7:Bumper3.Force = 7
              LeftSlingShot.SlingshotThreshold = 2:RightSlingShot.SlingshotThreshold = 2
              Tilt=0
              IF mem(2)=0 Then
                 If ActivePlayer=MaxPlayers Then

                    If Light1_X17.State = 1 Then
                    Round = Round
                    Else
                    Round=Round-1
                    End If

                    If Round<1 Then
                       ActivePlayer=0
                    Else
                       ActivePlayer=1
                    End If
                 Else
                    If Light1_X17.State = 1 Then
                    ActivePlayer = ActivePlayer
                    Else
                    ActivePlayer=ActivePlayer+1
                    End If
                 End If

        If B2SOn Then
          Controller.B2SSetScoreRolloverPlayer1 0
          Controller.B2SSetScoreRolloverPlayer2 0
          Controller.B2SSetScoreRolloverPlayer3 0
          Controller.B2SSetScoreRolloverPlayer4 0
          Controller.B2SSetScoreRollover 24+ActivePlayer, 2
        End if

              End If
              IF Round <1 Then
                 'End of game
           FreePlayNum = Int(rnd*9)
          If B2SOn Then
            Controller.B2SSetCanPlay 0
            Controller.B2SSetMatch FreePlayNum+1
            Controller.B2SSetGameover 1
            Controller.B2SSetBallinPlay 0
            Controller.B2SSetScoreRolloverPlayer1 0
            Controller.B2SSetScoreRolloverPlayer2 0
            Controller.B2SSetScoreRolloverPlayer3 0
            Controller.B2SSetScoreRolloverPlayer4 0
          End if
                    GameOverBG.State = 1
                    REPA(1) = 0:REPA(2) = 0:REPA(3) = 0:REPA(4) = 0
          REPB(1) = 0:REPB(2) = 0:REPB(3) = 0:REPB(4) = 0
          REPC(1) = 0:REPC(2) = 0:REPC(3) = 0:REPC(4) = 0
                    MatchBG.State = 1
                    TextBox2.text = FreePlayNum * 10
                   ' Check Match Wheel ------
                   For i=1 to MaxPlayers
                     If FreePlayNum = (STATScores(i) mod 10) Then
                        ReplayAction
                     End If
                  Next
                   '--------------------------
                  PlaySound "MekSound"
                  SeqTimer.Enabled=0
                  For i=1 to 3
                     SetBumperOff i
                  Next
                  'Hiscore
                  For i=1 To MaxPlayers
                     HStemp=STATScores(i)
                     If HStemp>HiScore Then
                        HiScore=HStemp
                        SaveValue  "captfant","DBHiScore",HiScore
                     End If
                  Next
                  MaxPlayers=0
                  If B2Son Then
          PlaySoundat SoundFXDOF("SSFlipper2",101,DOFOff,DOFFlippers), flipper2
          PlaySoundat SoundFXDOF("SSFlipper2",102,DOFOff,DOFFlippers), flipper3
          Flipper2.RotateToStart
          Flipper4.RotateToStart
          Flipper3.RotateToStart
          Flipper5.RotateToStart
          Else
          PlaySoundat SoundFX("SSFlipper2",DOFFlippers), flipper2
          PlaySoundat SoundFX("SSFlipper2",DOFFlippers), flipper3
          Flipper2.RotateToStart
          Flipper4.RotateToStart
          Flipper3.RotateToStart
          Flipper5.RotateToStart
          'DOF 101,0
          'DOF 102,0
          End If
              Else
                 'Start new ball
                 GameSeq=1
                 SeqTimer.Enabled=1
                 SeqTimer.Interval=300
                 If ft=1 Then
                    ft=0
                    PlaySound "SSResetGame"
                    SeqTimer.Interval=1200
                 Else
                    PlaySound "SSNewBall"
                 End If
              End If
       Case 3 'Collect Bonus
                  BCC=0
                  SeqTimer.Enabled=False
                  CollectBonusTimer.Enabled = True
       End Select
 End Sub

 '--------------------------------------------------------------------
 '------- TARGETS
 '--------------------------------------------------------------------
Sub T11_Hit()
   If Tilt>2 Then Exit Sub
   PlaySoundat SoundFXDOF("TargetSound",125,DOFPulse,DOFContactors),T11
  AddScoresToPlayer 100,ActivePlayer
End Sub
Sub T11_Timer()
   T11.TimerEnabled=False
   T12.IsDropped=True
   T11.IsDropped=False
End Sub

Sub T21_Hit()
   If Tilt>2 Then Exit Sub
   PlaySoundat SoundFXDOF("TargetSound",126,DOFPulse,DOFContactors),T21
  AddScoresToPlayer 100,ActivePlayer
End Sub
Sub T21_Timer()
   T21.TimerEnabled=False
   T22.IsDropped=True
   T21.IsDropped=False
End Sub

Sub DT1_Hit()
   If Tilt>2 Then Exit Sub
   DTL1.state = 0:DTGI.Intensity = DTGI.Intensity - 2
   PlaySoundat SoundFXDOF("TargetDrop",127,DOFPulse,DOFContactors),DT1
   DT1.IsDropped=True
   mem(15)=mem(15)+1
  AddScoresToPlayer 500,ActivePlayer
   DTargs.enabled = 1
End Sub
Sub DT2_Hit()
   If Tilt>2 Then Exit Sub
   DTL2.state = 0:DTGI.Intensity = DTGI.Intensity - 2
   PlaySoundat SoundFXDOF("TargetDrop",127,DOFPulse,DOFContactors),DT2
   DT2.IsDropped=True
   mem(15)=mem(15)+1
  AddScoresToPlayer 500,ActivePlayer
   DTargs.enabled = 1
End Sub
Sub DT3_Hit()
   If Tilt>2 Then Exit Sub
   DTL3.state = 0:DTGI.Intensity = DTGI.Intensity - 2
   PlaySoundat SoundFXDOF("TargetDrop",127,DOFPulse,DOFContactors),DT3
   DT3.IsDropped=True
   mem(15)=mem(15)+1
  AddScoresToPlayer 500,ActivePlayer
   DTargs.enabled = 1
End Sub
Sub DT4_Hit()
   If Tilt>2 Then Exit Sub
   DTL4.state = 0:DTGI.Intensity = DTGI.Intensity - 2
   PlaySoundat SoundFXDOF("TargetDrop",127,DOFPulse,DOFContactors),DT4
   DT4.IsDropped=True
   mem(15)=mem(15)+1
  AddScoresToPlayer 500,ActivePlayer
   DTargs.enabled = 1
End Sub
Sub DT5_Hit()
   If Tilt>2 Then Exit Sub
   DTL5.state = 0:DTGI.Intensity = DTGI.Intensity - 2
   PlaySoundat SoundFXDOF("TargetDrop",127,DOFPulse,DOFContactors),DT5
   DT5.IsDropped=True
   mem(15)=mem(15)+1
  AddScoresToPlayer 500,ActivePlayer
   DTargs.enabled = 1
End Sub

Dim gxx

Sub GI_Bright_Timer()
For Each gxx in GI_Lights
    gxx.Intensityscale = 1.1
    Next
    DOF 166,0:DOF 165,1
    me.enabled = 0
End Sub

 '----------------------------------------------------------------------
 '     BUMPER
 '----------------------------------------------------------------------
Sub Bumper1_Hit()
    If Tilt>2 Then Exit Sub
       If mem(10)=0 Then
      AddScoresToPlayer 10,ActivePlayer
          AltLights
       Else
      AddScoresToPlayer 100,ActivePlayer
      End If
      PlaySoundAtVol SoundFXDOF("Bumper",103,DOFPulse,DOFContactors), Primitive3, 0.05
End Sub

Sub Bumper2_Hit()
    If Tilt>2 Then Exit Sub
       If mem(10)=0 Then
      AddScoresToPlayer 10,ActivePlayer
          AltLights
       Else
      AddScoresToPlayer 100,ActivePlayer
      End If
       PlaySoundAtVol SoundFXDOF("Bumper",104,DOFPulse,DOFContactors), Primitive1, 0.05
End Sub

Sub Bumper3_Hit()
    If Tilt>2 Then Exit Sub
       If mem(11)=0 Then
      AddScoresToPlayer 10,ActivePlayer
          AltLights
       Else
      AddScoresToPlayer 100,ActivePlayer
      End If
           PlaySoundAtVol SoundFXDOF("Bumper",105,DOFPulse,DOFContactors), Primitive2, 0.05
End Sub

 Sub SetBumperOn(ByVal bnr)
   Dim a,i
   a=(bnr-1)*3
   For i=a To a+2
   Next
   a=(bnr-1)*2
   For i=a To a+1
     Bumpers(i).State=1
   Next
End Sub

Sub SetBumperOff(ByVal bnr)
   Dim a,i
   a=(bnr-1)*3
   For i=a To a+2
   Next
   a=(bnr-1)*2
   For i=a To a+1
     Bumpers(i).State=0
   Next
End Sub

Sub ResetBumperRings()
End Sub

 '------------------------------------------------------------------------
 '           TRIGGERS
 '------------------------------------------------------------------------
Sub Trigger1_Hit()
    PlaySoundAtBall "sensor"
  DOF 111, DOFPulse
    Trigger1.TimerEnabled=True
    If Tilt>2 Then Exit Sub
    AddScoresToPlayer 100,ActivePlayer

    If mem(16)>1 And mem(18)=0 AND Light1_X14.state = 1 Then
       FLSetOff 14,14
       mem(18)=1
       ReplayAction
    End If

   If mem(16)>0 And mem(18)=0 AND Light1_X14.state = 1 AND sadjust = 0 Then
       FLSetOff 14,14
       mem(18)=1
       ReplayAction
    End If

End Sub
Sub Trigger1_Timer()
    Trigger1.TimerEnabled=False
End Sub

Sub Trigger2_Hit()
    PlaySoundAtBall "sensor"
  DOF 112, DOFPulse
     Trigger2.TimerEnabled=True
    If Tilt>2 Then Exit Sub
    AddScoresToPlayer 300,ActivePlayer
    AddBonus 3
    If mem(16)>0 And mem(17)=0 AND Light1_X15.State = 1 Then
       FLSetOff 15,15
       FLSetOn 17,17
       mem(17)=1
       If B2SOn Then Controller.B2SSetShootAgain 1
    End If
End Sub
Sub Trigger2_Timer()
    Trigger2.TimerEnabled=False
End Sub

Sub Trigger3_Hit()
    PlaySoundAtBall "sensor"
  DOF 113, DOFPulse
    Trigger3.TimerEnabled=True
    If Tilt>2 Then Exit Sub
    AddScoresToPlayer 100,ActivePlayer
    If mem(6) =1 Then
       AddBonus 3
    AddScoresToPlayer 200,ActivePlayer
       Flipper1.TimerEnabled = True
    End If
End Sub
Sub Trigger3_Timer()
    Trigger3.TimerEnabled=False
End Sub
Sub Flipper1_Timer()
   Flipper1.TimerEnabled = False
   CloseGate
End Sub

Sub Trigger4_Hit()
    PlaySoundAtBall "sensor"
  DOF 114, DOFPulse
    Trigger4.TimerEnabled=True
    If Tilt>2 Then Exit Sub
    If mem(13)=0 Then
      mem(13)=1
      FLSetOff 19,19
      FLSetOn 12,12
      If mem(14)=1 Then
         FLSetOn 11,11
         mem(12)=1
      End If
    End If
    If mem(1)=1 Then
       AddBonus 3
       AddScoresToPlayer 300,ActivePlayer
       Else
       AddScoresToPlayer 100,ActivePlayer
    End If
End Sub
Sub Trigger4_Timer()
    Trigger4.TimerEnabled=False
End Sub

Sub Trigger5_Hit()
    PlaySoundAtBall "sensor"
  DOF 115, DOFPulse
    Trigger5.TimerEnabled=True
    If Tilt>2 Then Exit Sub
    AddScoresToPlayer 100,ActivePlayer
    If mem(1)=0 Then
       OpenGate
    End If
End Sub
Sub Trigger5_Timer()
    Trigger5.TimerEnabled=False
End Sub

Sub Trigger6_Hit()
    PlaySoundAtBall "sensor"
  DOF 116, DOFPulse
    Trigger6.TimerEnabled=True
    If Tilt>2 Then Exit Sub
    If mem(14)=0 Then
      mem(14)=1
      FLSetOff 20,20
      FLSetOn 13,13
      If mem(13)=1 Then
         FLSetOn 11,11
         mem(12)=1
      End If
    End If
    If mem(1)=1 Then
       AddBonus 3
       AddScoresToPlayer 300,ActivePlayer
       Else
       AddScoresToPlayer 100,ActivePlayer
    End If
End Sub
Sub Trigger6_Timer()
    Trigger6.TimerEnabled=False
End Sub


Sub Btriggs_Hit(idx)
PlaySoundatBall "sensor"
If Tilt>2 Then Exit Sub
    If mem(1)=1 Then
       AddBonus 1
    End If
If mem(1)=1 Then
    AddBonus 1
    End If
If BallsToPlay=3 Then
AddScoresToPlayer 1000,ActivePlayer
Else
AddScoresToPlayer 100,ActivePlayer
End If
End Sub


Sub Trigger13_Hit()
    PlaySoundAtBall "sensor"
    'ROlights 7,1
    Trigger13.TimerEnabled=True
    If Tilt>2 Then Exit Sub
     AltLights
     AddScoresToPlayer 10, ActivePlayer
End Sub
Sub Trigger13_Timer()
    Trigger13.TimerEnabled=False
End Sub

Sub Trigger14_Hit()
    PlaySoundAtBall "sensor"
    Trigger14.TimerEnabled=True
    If Tilt>2 Then Exit Sub
     AltLights
     AddScoresToPlayer 10, ActivePlayer
End Sub
Sub Trigger14_Timer()
    Trigger14.TimerEnabled=False
    ROlights 8,0
End Sub


Sub ROlights(Byval nr, Byval st)
   Select Case nr
      Case 1,2,3,4,5,6
         If mem(1)=0 Then

         Else

         End If
      Case 7,8
         If mem(1)=1 Then

         Else

         End If
      End Select
End Sub

'-----
Sub TriggerPlunger_Hit()
    InPlunger=1
End Sub

Sub TriggerPlunger_Unhit()
    InPlunger=0
    ' Shoot Again Off
    mem(2)=0
    FLSetOff 30,30
End Sub

 Sub TriggerNudge1_Hit()
    NudgeOn=0
  End Sub
 Sub TriggerNudge1_UnHit()
    NudgeOn=1
  End Sub
 Sub TriggerNudge2_Hit()
    NudgeOn=0
  End Sub
 Sub TriggerNudge2_UnHit()
    NudgeOn=1
  End Sub
'-----------------------------------------------------------
'     SLINGS and WALLS
'-----------------------------------------------------------
Sub R14_Hit()
    If Tilt>2 Then Exit Sub
    AddScoresToPlayer 10,ActivePlayer
    AltLights
End Sub
Sub R30_Hit()
    If Tilt>2 Then Exit Sub
    AddScoresToPlayer 10,ActivePlayer
    AltLights
End Sub
Sub R1_Hit()
    If Tilt>2 Then Exit Sub
    AddScoresToPlayer 10,ActivePlayer
End Sub
'-------------------------------------------------------------------
'      Alternators
'-------------------------------------------------------------------

Sub AltLights()
   mem(1)=Abs(mem(1)-1)
   Select Case mem(1)
      Case 0
         SetBumperOn 1
         SetBumperOn 2
         SetBumperOff 3
         For each robject in BRolls
     robject.state = 0
     Next
         mem(10)=1
         mem(11)=0
         FLSetOff 21,21
         FLSetOn 22,22
         FLSetOff 23,23

         If mem(16)>0  And mem(17)=0 AND sadjust = 1  AND ladjust = 1 Then
            FLSetOff 15,15 'Extra ball
         End If
         If mem(16)>1 And mem(18)=0 AND sadjust = 1  AND ladjust = 1 Then
            FLSetOn 14,14 'Special
         End If

         If mem(16)>0  And mem(17)=0 AND sadjust = 1 AND ladjust = 0 Then
            'FLSetOff 15,15 'Extra ball
         End If
         If mem(16)>1 And mem(18)=0 AND sadjust = 1 AND ladjust = 0 Then
            'FLSetOn 14,14 'Special
         End If



     If mem(16)>0  And mem(17)=0 AND sadjust = 0  AND ladjust = 1 Then
            FLSetOff 15,15 'Extra ball
         End If
         If mem(16)>0 And mem(18)=0 AND sadjust = 0  AND ladjust = 1 Then
            FLSetOn 14,14 'Special
         End If

         If mem(16)>0  And mem(17)=0 AND sadjust = 0 AND ladjust = 0 Then
            'FLSetOff 15,15 'Extra ball
         End If
         If mem(16)>0 And mem(18)=0 AND sadjust = 0 AND ladjust = 0 Then
            'FLSetOn 14,14 'Special
         End If




      Case 1
         SetBumperOff 1
         SetBumperOff 2
         SetBumperOn 3
         For each robject in BRolls
     robject.state = 1
     Next
         mem(10)=0
         mem(11)=1
         FLSetOn 21,21
         FLSetOff 22,22
         FLSetOn 23,23

         If mem(16)>0 And mem(17)=0 AND sadjust = 1 AND ladjust = 1 Then
            FLSetOn 15,15 'Extra ball
         End If
         If mem(16)>1 And mem(18)=0 AND sadjust = 1 AND ladjust = 1 Then
            FLSetOff 14,14 'Special
         End If

         If mem(16)>0 And mem(17)=0 AND sadjust = 1 AND ladjust = 0 Then
            'FLSetOn 15,15 'Extra ball
         End If
         If mem(16)>1 And mem(18)=0 AND sadjust = 1 AND ladjust = 0 Then
            'FLSetOff 14,14 'Special
         End If



         If mem(16)>0 And mem(17)=0 AND sadjust = 0 AND ladjust = 1 Then
            FLSetOn 15,15 'Extra ball
         End If
         If mem(16)>0 And mem(18)=0 AND sadjust = 0 AND ladjust = 1 Then
            FLSetOff 14,14 'Special
         End If

         If mem(16)>0 And mem(17)=0 AND sadjust = 0 AND ladjust = 0 Then
            'FLSetOn 15,15 'Extra ball
         End If
         If mem(16)>0 And mem(18)=0 AND sadjust = 0 AND ladjust = 0 Then
            'FLSetOff 14,14 'Special
         End If

   End Select
End Sub

Sub DTargs_Timer()
   me.enabled = 0
   If mem(15)=5 Then
      PlaySoundat SoundFXDOF("TargetBankReset",135,DOFPulse,DOFContactors),BulbFil13
      DT1.IsDropped=False:DT2.IsDropped=False:DT3.IsDropped=False:DT4.IsDropped=False:DT5.IsDropped=False
      DTL1.state = 1:DTL2.state = 1:DTL3.state = 1:DTL4.state = 1:DTL5.state = 1:DTGI.Intensity = 12
      mem(15)=0
    AddScoresToPlayer 3000,ActivePlayer

      If mem(16)<3 Then mem(16)=mem(16)+1

      Select case mem(16)
         Case 1 ' X-Ball light on
            If sadjust = 1 Then
            FLSetOn 15,15
            Else
            FLSetOn 15,15
            FLSetOn 14,14
            FLSetOn 18,18
            End If

         Case 2 ' Special light on
            If sadjust = 1 Then
            FLSetOn 18,18
            FLSetOn 14,14 ' Lane Special
            Else
            ReplayAction
            End If
         Case 3 ' Reward Special
            ReplayAction
      End Select
   End If
End Sub

 '------------------------------------------------------------------------
 '             KICKERS
 '------------------------------------------------------------------------


 '------------------------------------------------------------------------
 '             BONUS
 '------------------------------------------------------------------------
Sub AddBonus(ByVal ab)
   If Tilt>2 Then Exit Sub
   mem(8)=mem(8)+ab
End Sub

Sub AddBonusTimer_Timer()
   If Tilt>2 Then mem(8)=0
   If mem(8)>0 Then
      If mem(9)<15 Then
         mem(9)=mem(9)+1
         PlaySound "SSResetBonusNoScore"
         SetBonusLights
      End If
      mem(8)=mem(8)-1
   End If
End Sub

Dim BCC
Sub CollectBonusTimer_Timer()
   If mem(9)>0 Then
      If mem(12)=1 then
      BCC=BCC+1
      Select Case BCC
         Case 1
      AddScoresToPlayer 1000,ActivePlayer
            SReels(ActivePlayer).addvalue(1000)
            Exit Sub
         Case 2
            Exit Sub
         Case 3
            BCC=0
      End Select
   End If
   AddScoresToPlayer 1000,ActivePlayer
   SReels(ActivePlayer).addvalue(1000)
   PlaySound "SSResetBonusNoScore"
   mem(9)=mem(9)-1
   SetBonusLights
   Else
      CollectBonusTimer.Enabled = False
      SeqTimer.Interval=250
      SeqTimer.Enabled=True
      GameSeq=2
   End If
End Sub

Sub SetBonusLights()
   FLSetOff 1,10
   If mem(9)>0 Then
      If mem(9)<11 Then
         FLSetOn 1,mem(9)
      Else
         FLSetOn 10,10
         FLSetOn 1,mem(9)-10
      End If
   End If
End Sub

'------------------------------------------------------------------------
'             TARGETS, GATE, POST AND BUG
'------------------------------------------------------------------------

Sub UpPost()
   If Post.IsDropped=True Then
      Post.IsDropped=false
      PlaySound "solon"
   End If
End Sub

Sub DownPost()
   If Post.IsDropped=false Then
      Post.IsDropped=true
      PlaySound "solon"
   End If
End Sub

Sub OpenGate()
   If mem(6)=0 Then
      Flipper1.RotateToEnd
      PlaySoundat "OpenGate", Flipper1
      mem(6)=1
      FLSetOn 16,16
      DOF 150,2
   End If
End Sub

Sub CloseGate()
   If mem(6)=1 Then
      Flipper1.RotateToStart
      PlaySoundat "OpenGate", Flipper1
      mem(6)=0
      FLSetOff 16,16
      DOF 150,2
   End If
End Sub

Sub OpenTopGate()
   If mem(5)=0 Then
      Gate2.Open = True
      PlaySound "OpenGate"
      mem(5)=1
      FLSetOn 32,32
    End If
End Sub

Sub CloseTopGate()
   If mem(5)=1 Then
      Gate2.Open = False
      PlaySound "OpenGate"
      mem(5)=0
      FLSetOff 32,32
   End if
End Sub


 '------------------------------------------------------------------------
 '             SOUNDS
 '------------------------------------------------------------------------

 Sub Gate1_Hit()
    PlaySoundat "Gate", Primitive4
 End Sub

 Sub Gate2_Hit()
    PlaySoundat "Gate", Primitive53
 End Sub

Dim RStep, Lstep, Tstep, Blstep, TLstep, BRstep, TRstep
Dim dtest
Sub LeftSlingShot_Slingshot
    If Loose_Plastic = 1 Then plasdir = Int(Rnd*12)+1:dtest = Int(Rnd*12)+1 : Plas.enabled = 1
   If Tilt>2 Then Exit Sub
    PlaySoundat SoundFXDOF("slingshot",106,DOFPulse,DOFContactors), SLING3
    LSling.Visible = 0
    LSling1.Visible = 1
    sling3.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval  = 10
    AddScoresToPlayer 10,ActivePlayer
    AltLights
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling3.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling3.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
   If Tilt>2 Then Exit Sub
    PlaySoundat SoundFXDOF("slingshot",107,DOFPulse,DOFContactors), SLING4
    RSling.Visible = 0
    RSling1.Visible = 1
    SLING4.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval  = 10
    AddScoresToPlayer 10,ActivePlayer
    AltLights
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling4.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling4.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub RightSlingShot_Slingshot
   If Tilt>2 Then Exit Sub
    PlaySoundat SoundFXDOF("slingshot",107,DOFPulse,DOFContactors), SLING4
    RSling.Visible = 0
    RSling1.Visible = 1
    SLING4.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval  = 10
    AddScoresToPlayer 10,ActivePlayer
    AltLights
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling4.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling4.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub TopSlingShot_Hit
   If Tilt>2 Then Exit Sub
    'PlaySoundat SoundFXDOF("slingshot",107,DOFPulse,DOFContactors), SLING4
    TSling.Visible = 0
    TSling1.Visible = 1
    'SLING4.TransZ = -20
    TStep = 0
    TopSlingShot.TimerEnabled = 1
  TopSlingShot.TimerInterval  = 10
    AddScoresToPlayer 10,ActivePlayer
    AltLights
End Sub

Sub TopSlingShot_Timer
    Select Case TStep
        Case 3:TSLing1.Visible = 0:TSLing2.Visible = 1
        Case 4:TSLing2.Visible = 0:TSLing.Visible = 1:TopSlingShot.TimerEnabled = 0
    End Select
    TStep = TStep + 1
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************
Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / captfant.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / captfant.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / captfant.width-1
    If tmp > 0 Then
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

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, vol)
    PlaySound soundname, 1, (vol), AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

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

Sub RollingSoundUpdate()
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
End Sub


'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'*********Sound Effects**************************************************************************************************
                                       'Use these for your sound effects like ball rolling, etc.

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub flipper2_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub flipper3_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub flipper4_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub flipper5_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub RandomSoundHole()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySound "fx_Hole1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "fx_Hole2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "fx_Hole3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 4 : PlaySound "fx_Hole4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub Metals_Hit(idx)
RandomSoundMetal
End Sub

Sub RandomSoundMetal()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx_metal_hit_1"
    Case 2 : PlaySound "fx_metal_hit_2"
    Case 3 : PlaySound "fx_metal_hit_3"
  End Select
End Sub

Sub Woods_Hit(idx)
RandomSoundWood
End Sub

Sub RandomSoundWood()
  Select Case Int(Rnd*2)+1
    Case 1 : PlaySound "fx_woodhit"
    Case 2 : PlaySound "fx_woodhit1"
    Case 3 : PlaySound "fx_woodhit2"
  End Select
End Sub



'---------------------------------------------------------------------
'       Handle Replay and Hiscore
'---------------------------------------------------------------------

 Sub ReplayAction()
   If Coins<9 Then Coins=Coins+1
  If B2SOn Then Controller.B2SSetCredits Coins
   PlaySound SoundFXDOF("knocker",108,DOFPulse,DOFKnocker)
   DOF 132, DOFPulse
 End Sub

Function GetScore(ByVal pl)
   Dim temphi
   pl=(pl-1)*6
   temphi=ReelPulses(pl+6,2)
   temphi=temphi+(ReelPulses(pl+5,2)*10)
   temphi=temphi+(ReelPulses(pl+4,2)*100)
   temphi=temphi+(ReelPulses(pl+3,2)*1000)
   temphi=temphi+(ReelPulses(pl+2,2)*10000)
   temphi=temphi+(ReelPulses(pl+1,2)*100000)
   GetScore=temphi
End Function

 '----------------------------------------------------
 '   RULES
 '----------------------------------------------------

 sub DisplayRuleSheet()
  dim rules

  msgBox rules, vbOKonly, "       RULESHEET --- WILLIAMS/DOODLE BUG"
End Sub



'********************************************************************************************
Sub UpdateBIP()
  If BallsToPlay=3 Then
    If B2SOn Then
    Controller.B2SSetBallinPlay 0
    Controller.B2SSetBallinPlay 4-Round
        End If
        If Round = 3 Then
        B1BG.state = 1
        Else
        B1BG.state = 0
        End If
        If Round = 2 Then
        B2BG.state = 1
        Else
        B2BG.state = 0
        End If
        If Round = 1 Then
        B3BG.state = 1
        Else
        B3BG.state = 0
        End If
    'End If
  Else
    If B2SOn Then
    Controller.B2SSetBallinPlay 0
    Controller.B2SSetBallinPlay 6-Round
        End If
        If Round = 5 Then
        B1BG.state = 1
        Else
        B1BG.state = 0
        End If
        If Round = 4 Then
        B2BG.state = 1
        Else
        B2BG.state = 0
        End If
        If Round = 3 Then
        B3BG.state = 1
        Else
        B3BG.state = 0
        End If
        If Round = 2 Then
        B4BG.state = 1
        Else
        B4BG.state = 0
        End If
        If Round = 1 Then
        B5BG.state = 1
        Else
        B5BG.state = 0
        'End If
    End If
  End If
End Sub

Sub ResetMemory(ByVal All)
    If All=1 Then
       For i=0 to 30
          Mem(i)=0
       Next
    Else
       For i=5 to 30
          Mem(i)=0
       Next
    End If
End Sub

Sub PRreset()
   For  i=1 To 10
      PLU(i).WidthBottom = 0:PLU(i).WidthTop = 0
   Next
   PLU(i).WidthBottom = 40:PLU(i).WidthTop = 30
   PRLight.State=Abs(PRLight.State-1)
End Sub

Sub PRLight_Timer()
   If PRdir=0 Then
      If PRpull<10 Then
         PLU(PRpull).WidthBottom = 0:PLU(PRpull).WidthTop = 0
         PRpull=PRpull+1
         PLU(PRpull).WidthBottom = 30:PLU(PRpull).WidthTop = 30
         PRLight.State=Abs(PRLight.State-1)
      End If
   Else
      If PRpull>0 Then
         PLU(PRpull).WidthBottom = 0:PLU(PRpull).WidthTop = 0
         PRpull=PRpull-3
         If PRpull<0 Then PRpull=0
         PLU(PRpull).WidthBottom = 30:PLU(PRpull).WidthTop = 30
         PRLight.State=Abs(PRLight.State-1)
      Else
         PRLight.TimerEnabled=false
      End If
   End If
End Sub


'*********************************************************************************************
'*****************               FADING LIGHTS CODE                ***************************
'*****************                     ROSVE                       ***************************
'*********************************************************************************************
Dim FLData(100,7)
Dim FLMaxLight
FLMaxLight=23
' 1=command --> 0=set off   1=set on   2=flash   99=no change
' 2=state
' 3=repeats   999=continous
' 4=on duration
' 5=off duration
' 6=on dur mem
' 7=off dur mem

Sub FLSetOn(ByVal FLlight1, ByVal FLlight2)
   For i=FLlight1 to  FLlight2
      FLData(i,1)=1
      FLData(i,2)=0
      FLData(i,3)=0
      FLData(i,4)=0
      FLData(i,5)=0
   Next
End Sub

Sub FLSetOff(ByVal FLlight1, ByVal FLlight2)
   For i=FLlight1 to  FLlight2
      FLData(i,1)=0
      FLData(i,2)=4
      FLData(i,3)=0
      FLData(i,4)=0
      FLData(i,5)=0
   Next
End Sub

Sub FLFlash(ByVal FLlight1, ByVal FLlight2, ByVal FLrepeat, ByVal FLondur, ByVal FLoffdur)
   For i=FLlight1 to  FLlight2
      FLData(i,1)=2
      FLData(i,2)=0
      FLData(i,3)=FLrepeat
      FLData(i,4)=FLondur
      FLData(i,5)=FLoffdur
      FLData(i,6)=FLondur
      FLData(i,7)=FLoffdur
   Next
End Sub

Sub FLTimer_Timer()
   For i=1 To FLMaxLight
      Select Case FLData(i,1)
      Case 0
         FLData(i,2)=FLData(i,2)+1
         Select Case FLData(i,2)
         Case 5
            'L2(i).State=0
         Case 7
            L1(i).State=0
            FLData(i,1)=99
         End Select
      Case 1
         FLData(i,2)=FLData(i,2)+1
         Select Case FLData(i,2)
         Case 1
            L1(i).State=1
         Case 3
            'L2(i).State=1
            FLData(i,1)=99
         End Select

      Case 2
         FLData(i,2)=FLData(i,2)+1
         Select Case FLData(i,2)
         Case 1
            L1(i).State=1
         Case 3
            'L2(i).State=1
         Case 4 'on Duration
            IF FLData(i,4)>0 Then
               FLData(i,4)=FLData(i,4)-1
               FLData(i,2)=FLData(i,2)-1
            Else
               FLData(i,4)=FLData(i,6)
            End If
         Case 5
            'L2(i).State=0
         Case 7
            L1(i).State=0
            IF FLData(i,3)>0 Then
               IF FLData(i,5)>0 Then
                  FLData(i,5)=FLData(i,5)-1
                  FLData(i,2)=FLData(i,2)-1
                Else
                      FLData(i,5)=FLData(i,7)
                      FLData(i,3)=FLData(i,3)-1
                      If FLData(i,3)=998 Then FLData(i,3)=999
                      FLData(i,2)=0
                End If
            Else
               FLData(i,1)=99
            End If
         End Select
      End Select
   Next
End Sub


Dim a(20),b(20)

Sub Update_Stuff_Timer()
DiverterP.ObjRotZ = Flipper1.CurrentAngle + 90

If DT5.isdropped = 1 Then
DT5L.state = 1
'DTL5.State = 0
If DT5L.X => 163 Then
DT5L.X = DT5L.X - 6
DT5P.X = DT5P.X - 6
End If
Else
If DT5L.X <= 208 Then
DT5L.X = DT5L.X + 6
DT5P.X = DT5P.X + 6
End If
If DT5L.X => 208 Then
DT5L.intensity = 5
End If
End If

If DT4.isdropped = 1 Then
DT4L.state = 1
'DTL4.state = 0
If DT4L.X => 155 Then
DT4L.X = DT4L.X - 6
DT4P.X = DT4P.X - 6
End If
Else
If DT4L.X <= 199 Then
DT4L.X = DT4L.X + 6
DT4P.X = DT4P.X + 6
End If
If DT4L.X => 199 Then
DT4L.intensity = 5
End If
End If

If DT3.isdropped = 1 Then
DT3L.state = 1
'DTL3.state = 0
If DT3L.X => 142 Then
DT3L.X = DT3L.X - 6
DT3P.X = DT3P.X - 6
End If
Else
If DT3L.X <= 188 Then
DT3L.X = DT3L.X + 6
DT3P.X = DT3P.X + 6
End If
If DT3L.X => 188 Then
DT3L.intensity = 5
End If
End If

If DT2.isdropped = 1 Then
DT2L.state = 1
'DTL2.state = 0
If DT2L.X => 131 Then
DT2L.X = DT2L.X - 6
DT2P.X = DT2P.X - 6
End If
Else
If DT2L.X <= 180 Then
DT2L.X = DT2L.X + 6
DT2P.X = DT2P.X + 6
End If
If DT2L.X => 180 Then
DT2L.intensity = 5
End If
End If

If DT1.isdropped = 1 Then
DT1L.state = 1
'DTL1.state = 0
If DT1L.X => 120 Then
DT1L.X = DT1L.X - 6
DT1P.X = DT1P.X - 6
End If
Else
If DT1L.X <= 172 Then
DT1L.X = DT1L.X + 6
DT1P.X = DT1P.X + 6
End If
If DT1L.X => 172 Then
DT1L.intensity = 5
End If
End If

ScoreReel12.setvalue(HiScore)

BallShadowUpdate
RollingSoundUpdate

If ActivePlayer = 1 Then
P1BG.State = 1
Else
P1BG.State = 0
End If

If ActivePlayer = 2 Then
P2BG.State = 1
Else
P2BG.State = 0
End If

If ActivePlayer = 3 Then
P3BG.State = 1
Else
P3BG.State = 0
End If

If ActivePlayer = 4 Then
P4BG.State = 1
Else
P4BG.State = 0
End If

ScoreReel5.setvalue(Coins)


If MaxPlayers = 1 Then
MP1BG.state = 1
Else
MP1BG.State = 0
End If

If MaxPlayers = 2 Then
MP2BG.state = 1
Else
MP2BG.State = 0
End If

If MaxPlayers = 3 Then
MP3BG.state = 1
Else
MP3BG.State = 0
End If

If MaxPlayers = 4 Then
MP4BG.state = 1
Else
MP4BG.State = 0
End If

If Tilt > 2.9 Then
TILTBG.state = 1
Else
TILTBG.state = 0
End If

If b3Bumper.State = 1 Then
B3S.visible = 0
Else
B3S.visible = 1
End If

If b2Bumper.State = 1 Then
B2S.visible = 0
Else
B2S.visible = 1
End If

If b1Bumper.State = 1 Then
B1S.visible = 0
Else
B1S.visible = 1
End If

FlipperLSh.RotZ = Flipper2.currentangle
FlipperLSh1.RotZ = Flipper4.currentangle
FlipperRSh.RotZ = Flipper3.currentangle
FlipperRSh1.RotZ = Flipper5.currentangle

DT5P7.X = Gate1.CurrentAngle + 152
DT5P8.X = Gate2.CurrentAngle + 733

'Handle Cabinet View HighScore.

'Split HiScore intvar into string and
'assign to separate reels.

Dim str:str = CStr(right("000000" & HiScore,6))
Dim I
for I = 0 to len(str)-1
    a(I)=Cint(Mid(str,I+1,1))
next

ScoreReel6.setvalue(a(0))
ScoreReel7.setvalue(a(1))
ScoreReel8.setvalue(a(2))
ScoreReel9.setvalue(a(3))
ScoreReel10.setvalue(a(4))
ScoreReel11.setvalue(a(5))

If Coins > 0 Then
Clight.State = 1
Else
Clight.State = 0
End If

End Sub

Dim anidir, aniup, anidown

Sub IC_ani_Timer
Select Case anidir
Case 1:
Card3.visible = 1:Card4.visible = 1:Light1.State = 1:Light2.State = 1
If Card3.X => 486 Then
aniup = 1:anidown = 0
me.enabled = 0
bulb_move.enabled = 1
End If
Card3.size_X = Card3.size_X + 0.012
Card3.size_Y = Card3.size_Y + 0.012
Card3.Y = Card3.Y - 4
Card3.X = Card3.X + 2

Card4.size_X = Card4.size_X + 0.012
Card4.size_Y = Card4.size_Y + 0.012
Card4.Y = Card4.Y - 4
Card4.X = Card4.X + 2

Case 2:
If Card3.X <= 156.4871 Then
Card3.visible = 0:Card4.visible = 0
aniup = 0:anidown = 1
Light1.X = 115
Light2.X = 865
me.enabled = 0
End If
Card3.size_X = Card3.size_X - 0.012
Card3.size_Y = Card3.size_Y - 0.012
Card3.Y = Card3.Y + 4
Card3.X = Card3.X - 2

Card4.size_X = Card4.size_X - 0.012
Card4.size_Y = Card4.size_Y - 0.012
Card4.Y = Card4.Y + 4
Card4.X = Card4.X - 2
End Select
End Sub

Sub Bulb_Move_Timer
If Light1.X => 840 Then
Light1.State = 0:light2.State = 0
me.enabled = 0
End If
Light1.X = Light1.X + 3
Light2.X = Light2.X - 3
End Sub

Dim plasdir

plasdir = 1

Sub plas_timer()
Select Case plasdir
Case 1:
Glass.image = "plasticsr1":plasdir = 2
Case 2:
Glass.image = "plasticsr2":plasdir = 3
Case 3:
Glass.image = "plasticsr3":plasdir = 4
Case 4:
Glass.image = "plasticsr2":plasdir = 5
Case 5:
Glass.image = "plasticsr1":plasdir = 6
Case 6:
Glass.image = "plastics2":me.enabled = 0
Case 7:
Glass.image = "plastics_sr1":plasdir = 8
Case 8:
Glass.image = "plastics_sr2":plasdir = 9
Case 9:
Glass.image = "plastics_sr3":plasdir = 10
Case 10:
Glass.image = "plastics_sr2":plasdir = 11
Case 11:
Glass.image = "plastics_sr1":plasdir = 12
Case 12:
Glass.image = "plastics2":me.enabled = 0:
End Select
End Sub

Dim GIF
Sub GI_Flicker_Timer
GIF = Int(Rnd*6)+1
Flick_Gi GIF
End Sub

Sub Flick_GI(gif)
Select Case gif
Case 1:GI_whiteupper14.state = 1:BLGI.state = 1
Case 2:GI_whiteupper14.state = 0:BLGI.state = 0
Case 3:GI_whiteupper14.state = 2:BLGI.state = 2
Case 4:GI_whiteupper14.state = 1:BLGI.state = 1
Case 5:GI_whiteupper14.state = 0:BLGI.state = 0
Case 6:GI_whiteupper14.state = 2:BLGI.state = 2
End Select
End Sub

'*****************************************
' Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5)


Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls

  ' render the shadow for each ball
    For b = 0 to UBound(BOT)
    If BOT(b).X < captfant.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (captfant.Width/2))/7)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (captfant.Width/2))/7)) - 10
    End If

      BallShadow(b).Y = BOT(b).Y + 10
      BallShadow(b).Z = 1
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

Sub captfant_Exit
  If B2SOn Then Controller.stop
End Sub


'////////////////////////////////////////////////////////////////
