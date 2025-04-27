'---------------------------------------------------------------------------
'IMPACTO
'Original VP9 table by luvthatapex&lucasbuck
'Ported and updated For VPX by tgx during Covid 2021
'B2S Script by Loserman
'Table is FS only
'Designed for use with Editoy backglass.
'---------------------------------------------------------------------------
'
Option explicit
 Randomize

ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

BallSize = 50 ' 50 is the normal size
BallMass = 1.7  ' 1 is the normal ball mass.

Const cGameName = "Impacto"

'' Load the core.vbs for supporting Subs and functions
'    On Error Resume Next
'        ExecuteGlobal GetTextFile("core.vbs")
'    If Err Then MsgBox "Can't open core.vbs"
    ExecuteGlobal GetTextFile("b2s.vbs")
  If Err Then MsgBox "Cannot find b2s.vbs!"
    On Error Goto 0

'Dim Controller
'Set Controller = CreateObject("B2S.Server")
'Controller.Run


Dim B2SOn
B2SOn = true
If B2SOn Then Set Controller = CreateObject("B2S.Server"): Controller.Run


Dim bsTrough, dt5t
Dim bump1, bump2
Dim x
Dim BallBox(5)
Dim MatchBox(10)

Set BallBox(1) = Ball1
Set BallBox(2) = Ball2
Set BallBox(3) = Ball3
Set BallBox(4) = Ball4
Set BallBox(5) = Ball5

Set MatchBox(1) = Match00
Set MatchBox(2) = Match10
Set MatchBox(3) = Match20
Set MatchBox(4) = Match30
Set MatchBox(5) = Match40
Set MatchBox(6) = Match50
Set MatchBox(7) = Match60
Set MatchBox(8) = Match70
Set MatchBox(9) = Match80
Set MatchBox(10) = Match90

Dim Digit(4)
Dim Steps(4)
Dim Match
Dim BonusPause
Dim Score
Dim GotReplay
Dim HolePause
Dim HoleExit
Dim StartPause
Dim Credits
Dim HighScore
Dim BallsLeft
Dim BallActive
Dim CreditPause
Dim SPReady
Dim TiltSensor
Dim MachineTilt
Dim P10
Dim P100
Dim P1000
Dim f, g, h, i
Dim RedDone
Dim LitLane

'Table Init
  LoadEM
    TiltB1.IsDropped=1
  TiltB2.IsDropped=1

 ' Bumpers
 Sub Bumper1_Hit
  PlaySoundAtVol "Rpop", ActiveBall, 1
  If BumpRight.State=0 Then
    P100 = P100 + 1
  Else
    P1000 = P1000 + 1
    End If
 End Sub

 Sub Bumper2_Hit
  PlaySoundAtVol "Lpop", ActiveBall, 1
  If BumpLeft.State=0 Then
    P100 = P100 + 1
  Else
    P1000 = P1000 + 1
   End If
  End Sub

Sub BumpersOn()
  BumpLeft.state=1
  BumpLeftb.state=1
  BumpRight.state=0
  BumpRightb.state=0
End Sub

Sub BumpersOff()
  BumpLeft.state=0
  BumpLeftb.state=0
  BumpRight.state=1
  BumpRightb.state=1
End Sub

Sub ScoreTimer_Timer()
    h = 0
    If MachineTilt = 1 then
        P10 = 0
        P100 = 0
        P1000 = 0
    End If
    If P10> 0 or P100> 0 or P1000> 0 then
        StopSound "Bell10"
        StopSound "Bell100"
        StopSound "Bell1000"
        PlaySound "Cluper"
    End If
    If P1000> 0 and h = 0 then
        P1000 = P1000-1
        Digit(2) = Digit(2) + 1
        Digit2.AddValue(1)
        If Digit(2) = 10 then
            Digit(2) = 0
            Digit(1) = Digit(1) + 1
            Digit1.AddValue(1)
    B2SData(1)=Chr(Digit(1)+1)
        End If
    B2SData(2)=Chr(Digit(2)+1)
        If Digit(1) = 10 then Digit(1) = 0
    B2SData(1)=Chr(Digit(1)+1)
        AddScore(1000)
        If MachineTilt = 0 then PlaySound "Bell1000"
        h = 1
    End If
    If P100> 0 and h = 0 then
        P100 = P100-1
        Digit(3) = Digit(3) + 1
        Digit3.AddValue(1)
        If Digit(3) = 10 then
            Digit(3) = 0
            Digit(2) = Digit(2) + 1
            Digit2.AddValue(1)
    B2SData(2)=Chr(Digit(2)+1)
        End If
    B2SData(3)=Chr(Digit(3)+1)
        If Digit(2) = 10 then
            Digit(2) = 0
    B2SData(2)=Chr(Digit(2)+1)
            Digit(1) = Digit(1) + 1
            Digit1.AddValue(1)
    B2SData(1)=Chr(Digit(1)+1)
        End If
        If Digit(1) = 10 then Digit(1) = 0
        AddScore(100)
        If MachineTilt = 0 then PlaySound "Bell100"
        h = 1
    End If
    If P10> 0 and h = 0 then
        P10 = P10-1
        Digit(4) = Digit(4) + 1
        Digit4.AddValue(1)
        If Digit(4) = 10 then
            Digit(4) = 0
            Digit(3) = Digit(3) + 1
            Digit3.AddValue(1)
    B2SData(3)=Chr(Digit(3)+1)
        End If
    B2SData(4)=Chr(Digit(4)+1)
        If Digit(3) = 10 then
            Digit(3) = 0
    B2SData(3)=Chr(Digit(3)+1)
            Digit(2) = Digit(2) + 1
            Digit2.AddValue(1)
    B2SData(2)=Chr(Digit(2)+1)
        End If
        If Digit(2) = 10 then
            Digit(2) = 0
    B2SData(2)=Chr(Digit(2)+1)
            Digit(1) = Digit(1) + 1
            Digit1.AddValue(1)
    B2SData(1)=Chr(Digit(1)+1)
        End If
        If Digit(1) = 10 then Digit(1) = 0: B2SData(1)=Chr(Digit(1)+1)
        AddScore(10)
        If MachineTilt = 0 then PlaySound "Bell10"
    End If
'***************************************

    If B2Stimer.Enabled=false then SetB2SData 6,0
'***************************************
    If BallActive = 2 then
        BonusPause = BonusPause-1
        If BonusPause <1 then
            If BallsLeft> 1 then
                NextBall
            Else
                If BonusPause = 0 then PlaySound "GameOver"
                If BonusPause = -6 then
                    Match = INT(RND * 10)
                    MatchBox(Match + 1).Text = Match& "0"
                    If Match>9 Then Match 9
                If Match = 0 then
          Controller.B2SSetData 34,10
                Else
          Controller.B2SSetData 34,Match
                End If
                    BallBox(5).Text = ""
                    Controller.B2SSetData 32,0
                    If HighScore <Score then HighScore = Score
                    HighBox.Text = "High Score: " &FormatNumber(HighScore, 0, -1, 0, -1)
                    g = Score MOD 100
                    If g = Match * 10 then
                        Credits = Credits + 1
                        PlaySound "Knocker"
                        CreditPause = 1
                    End If
                    BallActive = 0
                    GoBox.Text = "GAME OVER"
          'Lights Out
            LeftOut.state=0
            LeftIn.state=0
            RightOut.state=0
            RightIn.state=0
            For Each i in GI_Lights
              i.state=0
            Next
                    Controller.B2SSetData 35,1
                    TiltBox.Text = ""
                    Controller.B2SSetData 33,0
                    StopSound "Buzz"
                    SaveData
                End If
            End If
        End If
    End If
End Sub

Sub GameTimer_Timer()
    If HolePause> 0 then
        If P10 = 0 and P100 = 0 and P1000 = 0 then HolePause = HolePause-1
        If HolePause = 0 then
            PlaySound "popper_ball"
        End If
    End If

    If StartPause> 0 then
        StartPause = StartPause-1
        If StartPause = 185 then PlaySound "Initialize"
        If StartPause = 178 then
            For f = 1 to 10
                MatchBox(f).Text = ""
            Next
            Controller.B2SSetData 34,0
            For f = 1 to 5
                BallBox(f).Text = ""
            Next
            Controller.B2SSetData 32,0

            HighBox.Text = ""
            GOBox.Text = ""
            Controller.B2SSetData 35,0
            TiltBox.Text = ""
            Controller.B2SSetData 33,0
            For f = 1 to 4
                Steps(f) = 10-Digit(f)
                If Digit(f) = 5 then Steps(f) = 15
                If Digit(f) = 0 then Steps(f) = 20

                Digit(f) = 0
            Next
        End If
        If StartPause = 157 or StartPause = 151 or StartPause = 145 or StartPause = 139 or StartPause = 133 then DoReset
        If StartPause = 118 or StartPause = 112 or StartPause = 106 or StartPause = 100 or StartPause = 94 then DoReset
        If StartPause = 79 or StartPause = 73 or StartPause = 67 or StartPause = 61 or StartPause = 55 then DoReset
        If StartPause = 40 or StartPause = 34 or StartPause = 28 or StartPause = 22 or StartPause = 16 then DoReset
        If StartPause = 8 then PlaySound "ballrel"
        If StartPause = 0 then
            CheckLights
            TiltBox.Text = ""
            controller.B2SSetData 33,0
            For f = 1 to 5
                BallBox(f).Text = ""
            Next
            BallBox(6-BallsLeft).Text = (6-BallsLeft)
            Controller.B2SSetData 32,6-BallsLeft
            BallExit.CreateBall
            BallExit.Kick 90, 8
        End If
    End If

    If CreditPause> 0 then
        CreditPause = CreditPause-1
        If CreditPause = 0 then
            Credits = Credits + 1
            If Credits> 9 then Credits = 9
            CreditBox.SetValue(Credits)
            Controller.B2SSetCredits Credits
        End If
    End If

    If TiltSensor> 99 and MachineTilt = 0 and BallActive = 1 then
        PlaySound "solenoidon"
        MachineTilt = 1
        TiltBox.Text = "TILT"
        TiltB1.IsDropped=0
      TiltB2.IsDropped=0
        Controller.B2SSetData 33,1
        P10 = 0
        P100 = 0
        P1000 = 0
    End If
    If TiltSensor> 0 then
        TiltSensor = TiltSensor-1
    End If
End Sub

  ' Table init
  Sub Impacto_Init
' LoadController
  AllLightsOut
  ' init 100k
    Controller.B2SSetData 26,0
    Controller.B2SSetData 27,0
    ' Init Bumper Rings,rollovers,Roto and Standup targets
    TiltB1.IsDropped=1
  TiltB2.IsDropped=1
    Score = 0
    Credits = 0
    Match = 0
    HighScore = 50000
    SPReady = 0
    RedDone = 0
    LitLane = -1
    LoadData
    CheckLights
    DoDigits
    PlaySound "Motor2"
    GOBox.Text = "GAME OVER"
    Controller.B2SSetData 35,1
    MatchBox(Match + 1).Text = Match& "0"
    If Match>9 Then Match 9
    If Match = 0 then
    Controller.B2SSetData 34,10
    Else
    Controller.B2SSetData 34,Match
    End If
    CreditBox.SetValue(Credits)
    Controller.B2SSetCredits Credits
    HighBox.Text = "High Score: " &FormatNumber(HighScore, 0, -1, 0, -1)
  End Sub

Sub AddScore(f)
    If MachineTilt = 0 then
        Score = Score + f
    Controller.B2SSetScorePlayer1 Score
        If Score> 99990 then
           Score100K.Text = "100,000"
           Controller.B2SSetData 26,1
        End If
        If GotReplay = 0 then
            If Score> (82000) -1 then
                GotReplay = 1
                PlaySound "Knocker"
                CreditPause = 1
            End If
        End If
        If GotReplay = 1 then
            If Score> (99000) -1 then
                GotReplay = 2
                PlaySound "Knocker"
                CreditPause = 1
            End If
        End If
    End If
End Sub

Sub DoDigits

    If Score> 99990 then
        Score100K.Text = "100,000"
        Controller.B2SSetData 26,1
        Score = Score -(Int(Score / 100000) * 100000)
    End If
    Digit(1) = Int(Score / 10000)
    Score = Score -(Int(Score / 10000) * 10000)
    Digit(2) = Int(Score / 1000)
    Score = Score -(Int(Score / 1000) * 1000)
    Digit(3) = Int(Score / 100)
    Score = Score -(Int(Score / 100) * 100)
    Digit(4) = Score / 10
    Digit1.SetValue(Digit(1) )
    Digit2.SetValue(Digit(2) )
    Digit3.SetValue(Digit(3) )
    Digit4.SetValue(Digit(4) )
    Controller.B2SSetScorePlayer 1,Score
    Score = 0
End Sub

Sub DoReset
    If Steps(1)> 0 then
        Steps(1) = Steps(1) -1
        Digit1.AddValue(1)
    End If
    If Steps(2)> 0 then
        Steps(2) = Steps(2) -1
        Digit2.AddValue(1)
    End If
    If Steps(3)> 0 then
        Steps(3) = Steps(3) -1
        Digit3.AddValue(1)
    End If
    If Steps(4)> 0 then
        Steps(4) = Steps(4) -1
        Digit4.AddValue(1)
    End If
End Sub

Sub CheckLights
End Sub

Sub SaveData
    SaveValue "HTD", "HighScore", HighScore
    SaveValue "HTD", "Credits", Credits
    SaveValue "HTD", "Score", Score
    SaveValue "HTD", "Match", Match
End Sub

Sub LoadData
    Dim value
    Value = LoadValue("HTD", "HighScore")
    If(Value <> "") then HighScore = CDbl(Value) End If

    Value = LoadValue("HTD", "Credits")
    If(Value <> "") then Credits = CDbl(Value) End If

    Value = LoadValue("HTD", "Score")
    If(Value <> "") then Score = CDbl(Value) End If

    Value = LoadValue("HTD", "Match")
    If(Value <> "") then Match = CDbl(Value) End If

End Sub

Sub FirstBall
    StopSound "GameOver"
    StartPause = 186
    Score = 0
  Controller.B2SSetScorePlayer1 Score
    GotReplay = 0
    BallsLeft = 6
  ResetCenter
    NextBall
End Sub

Sub NextBall
    P10 = 0
    P100 = 0
    P1000 = 0
    TiltSensor = 0
    MachineTilt = 0
    BallsLeft = BallsLeft-1
    If StartPause <1 then StartPause = 10
    BallActive = 1
  ResetCenter
End Sub

Sub Impacto_KeyDown(ByVal keycode)

    If keycode = PlungerKey Then
        Plunger.PullBack
    End If

    If keycode = LeftFlipperKey and MachineTilt = 0 and BallActive> 0 Then
        LeftFlipper.RotateToEnd
        PlaySoundAtVol "FlipperUp", LeftFlipper, 1
        PlayLoopSoundAtVol "Buzz", LeftFlipper, 1
    End If

    If keycode = RightFlipperKey and MachineTilt = 0 and BallActive> 0 Then
        RightFlipper.RotateToEnd
        PlaySoundAtVol "FlipperUp", RightFlipper, 1
        PlayLoopSoundAtVol "Buzz", RightFlipper, 1
    End If

    If keycode = LeftTiltKey and MachineTilt = 0 and BallActive = 1 Then
        Nudge 15, 1.4 :PlaySound "nudge_left"
        TiltSensor = TiltSensor + 50
    End If

    If keycode = RightTiltKey and MachineTilt = 0 and BallActive = 1 Then
        Nudge 345, 1.4:PlaySound "nudge_right"
        TiltSensor = TiltSensor + 50
    End If

    If keycode = CenterTiltKey and MachineTilt = 0 and BallActive = 1 Then
        Nudge 0, 1.4:PlaySound "nudge_forward"
        TiltSensor = TiltSensor + 50
    End If

    If keycode = 6 and CreditPause = 0 then
        PlaySound "Coin"
        CreditPause = 60
    Playsound "reel"
    End If

    If keycode = 2 and BallActive = 0 and StartPause = 0 then
        If Credits> 0 then
            Credits = Credits-1
            CreditBox.SetValue(Credits)
            Controller.B2SSetCredits Credits
      'Init Lamps
        LeftOut.state=1
        RightIn.state=1
        For Each i in GI_Lights
          i.state=1
        Next
        FirstBall
        End If
    End If
End Sub

Sub Impacto_KeyUp(ByVal keycode)

    If keycode = PlungerKey Then
        Plunger.Fire
    playsoundAtVol "plunger", Plunger, 1
    End If

    If keycode = LeftFlipperKey Then
        LeftFlipper.RotateToStart
        If MachineTilt = 0 and BallActive> 0 then
            PlaySoundAtVol "FlipperDown", LeftFlipper, 1
            StopSound "Buzz"
        End If
    End If

    If keycode = RightFlipperKey Then
        RightFlipper.RotateToStart
        If MachineTilt = 0 and BallActive> 0 then
            PlaySoundAtVol "FlipperDown", RightFlipper, 1
            StopSound "Buzz"
        End If
    End If
End Sub

Sub Drain_Hit()
    Drain.DestroyBall
    PlaySoundAtVol "Drain", Drain, 1
    BallActive = 2
    BonusPause = 10
  playsound "drop_reset_center"
  sw1.IsDropped = 0:sw2.IsDropped = 0:sw3.IsDropped = 0:sw8.IsDropped = 0
  sw4.IsDropped = 0:sw5.IsDropped = 0:sw6.IsDropped = 0:sw7.IsDropped = 0
  sw1a.IsDropped = 0:sw2a.IsDropped = 0:sw3a.IsDropped = 0:sw8a.IsDropped = 0
  sw4a.IsDropped = 0:sw5a.IsDropped = 0:sw6a.IsDropped = 0:sw7a.IsDropped = 0
    LeftLoop.state=0: RightLoop.state=0
  TopRight.state=0:TopLeft.state=0
  BumpersOff
End Sub

' Slingshots
'alternate inlane and outlane lamps.
'starts with left outlane/right inlane lit

Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySoundAtVol "slingshot", sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
  P10 = P10 + 1
  Alternate
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol "slingshot", sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
  P10 = P10 + 1
  Alternate
End Sub

 ' Rollovers
  Sub sw12_Hit
   CheckCenter
   P100 = P100 + 5
   PlaySoundAtVol "sensor", ActiveBall, 1
  End Sub

  Sub sw13_Hit
  P100 = P100 + 5
  PlaySoundAtVol "sensor", ActiveBall, 1
  End Sub

  Sub sw14_Hit
   CheckCenter
   P100 = P100 + 5
   PlaySoundAtVol "sensor", ActiveBall, 1
  End Sub

  Sub sw10_Hit
  P100 = P100 + 5
  PlaySoundAtVol "sensor", ActiveBall, 1
  'Special
    If RightLoop.state=1 then
     Credits = Credits + 1
     PlaySound "Knocker"
     CreditPause = 1
     RightLoop.state=0
  End If
  End Sub

  Sub sw9_Hit
   P100 = P100 + 5
   PlaySound "sensor"
  'Special
    If LeftLoop.state=1 then
     Credits = Credits + 1
     PlaySound "Knocker"
     CreditPause = 1
  End If
  End Sub

 'rollovers double
  Sub sw32b_Hit
  If LeftOut.state=1 then
    P100 = P100 + 5
  Else
    P10 = P10 + 5
  End If
    PlaySoundAtVol "sensor", ActiveBall, 1
  End Sub

  Sub sw41_Hit
  If LeftIn.state=1 then
    P100 = P100 + 5
  Else
    P10 = P10 + 5
  End If
  PlaySoundAtVol "sensor", ActiveBall, 1
  End Sub

  Sub sw41b_Hit
  If RightIn.state=1 then
    P100 = P100 + 5
  Else
    P10 = P10 + 5
  End If
  PlaySoundAtVol "sensor", ActiveBall, 1
  End Sub

  Sub sw34b_Hit
  If RightOut.state=1 then
    P100 = P100 + 5
  Else
    P10 = P10 + 5
  End If
   PlaySoundAtVol "sensor", ActiveBall, 1
  End Sub

  'SlingShots sws
  'Drop Targets
  'left bank
  Sub sw1_Hit
sw1.IsDropped = 1:sw1a.IsDropped = 1
P100 = P100 + 5
PlaySoundAtVol "drop_down_left", ActiveBall, 1
  If sw1.isdropped = 1 and sw2.isdropped = 1 and sw3.isdropped = 1 and sw8.isdropped=1 then
  LeftLoop.state=1
  End If
End Sub
  Sub sw2_Hit
 sw2.IsDropped = 1:sw2a.IsDropped = 1
 P100 = P100 + 5
 PlaySoundAtVol "drop_down_left", ActiveBall, 1
  If sw1.isdropped = 1 and sw2.isdropped = 1 and sw3.isdropped = 1 and sw8.isdropped=1 then
  LeftLoop.state=1
  End If
End Sub
  Sub sw3_Hit
   sw3.IsDropped = 1:sw3a.IsDropped = 1
   P100 = P100 + 5
   PlaySoundAtVol "drop_down_left", ActiveBall, 1
  If sw1.isdropped = 1 and sw2.isdropped = 1 and sw3.isdropped = 1 and sw8.isdropped=1 then
  LeftLoop.state=1
  End If
  End Sub
  Sub sw8_Hit
   sw8.IsDropped = 1:sw8a.IsDropped = 1
   P100 = P100 + 5
   PlaySoundAtVol "drop_down_left", ActiveBall, 1
  If sw1.isdropped = 1 and sw2.isdropped = 1 and sw3.isdropped = 1 and sw8.isdropped=1 then
  LeftLoop.state=1
  End If
  End Sub
  'right bank
  Sub sw4_Hit
   sw4.IsDropped = 1:sw4a.IsDropped = 1
   P100 = P100 + 5
   PlaySoundAtVol "drop_down_right", ActiveBall, 1
  If sw4.isdropped = 1 and sw5.isdropped = 1 and sw6.isdropped = 1 and sw7.isdropped=1 then
  RightLoop.state=1
  End If
  End Sub
  Sub sw5_Hit
   sw5.IsDropped = 1:sw5a.IsDropped = 1
   P100 = P100 + 5
   PlaySoundAtVol "drop_down_right", ActiveBall, 1
  If sw4.isdropped = 1 and sw5.isdropped = 1 and sw6.isdropped = 1 and sw7.isdropped=1 then
  RightLoop.state=1
  End If
  End Sub
  Sub sw6_Hit
   sw6.IsDropped = 1:sw6a.IsDropped = 1
   P100 = P100 + 5
   PlaySoundAtVol "drop_down_right", ActiveBall, 1
  If sw4.isdropped = 1 and sw5.isdropped = 1 and sw6.isdropped = 1 and sw7.isdropped=1 then
  RightLoop.state=1
  End If
  End Sub
  Sub sw7_Hit
  sw7.IsDropped = 1:sw7a.IsDropped = 1
  P100 = P100 + 5
  PlaySoundAtVol "drop_down_right", ActiveBall, 1
  If sw4.isdropped = 1 and sw5.isdropped = 1 and sw6.isdropped = 1 and sw7.isdropped=1 then
  RightLoop.state=1
  End If
  End Sub

  ' Gate
  Sub Gate_Hit():PlaySoundAtVol "gate", ActiveBall, 1:End Sub
  Sub Gate2_Hit():PlaySoundAtVol "gate", ActiveBall, 1:End Sub
  Sub Wall471_hit():playsoundAtVol "metalhit", ActiveBall, 1:end sub

Sub Kicker1_Hit
  CheckSpecial
End Sub

Sub Kicker1_Timer
  Kicker1.Kick 190,5              'Kick the ball out after a delay
  PlaySoundAtVol"Kicker", Kicker1, 1
  Me.TimerEnabled=False
  ResetBonus
End Sub

Sub AllLightsOut()
  TopLeft.state=0
  TopRight.state=0
  LeftLoop.state=0
  RightLoop.state=0
  Special.state=0
  TwentyFiveK.state=0
  TwentyK.state=0
  FifteenK.state=0
  TenK.state=0
  FiveK.state=0
  LeftOut.state=0
  LeftIn.state=0
  RightOut.state=0
  RightIn.state=0
  For Each i in GI_Lights
    i.state=0
  Next
End Sub

Sub ResetBonus()
  FiveK.state=0
  TenK.state=0
  FifteenK.state=0
  TwentyK.state=0
  TwentyfiveK.state=0
  Special.state=0
End Sub

Sub Alternate()
'alternate inlane/outlane lighting
If LeftOut.state =1 then
  LeftIn.state=1
  RightOut.state=1
  RightIn.state=0
    LeftOut.state=0
    BumpersOn
Else
  LeftIn.state=0
  RightOut.state=0
  RightIn.state=1
    LeftOut.state=1
  BumpersOff
End If
End Sub

Sub CheckCenter()
' If Special.IsDropped = False Then Special.IsDropped = False: Exit Sub
  If TwentyFiveK.state=1 Then TwentyFiveK.state=0:Special.state=1: Exit Sub
  If TwentyK.state =1 Then TwentyK.state=0:TwentyFiveK.state=1: Exit Sub
  If FifteenK.state=1 Then FifteenK.state=0: TwentyK.state=1: Exit Sub
  If TenK.state=1 Then TenK.state=0: FifteenK.state=1
  If FiveK.state=1 Then FiveK.state=0: TenK.state = 1
End Sub

Sub ResetCenter()
  FiveK.state = 1
  TenK.state=0
  FifteenK.state=0
  TwentyK.state=0
  TwentyfiveK.state=0
  Special.state=0
End Sub

Sub CheckSpecial()
  If FiveK.state=1 then P1000=P1000+5
  If TenK.state=1 then  P1000=P1000+10
  If FifteenK.state=1 then P1000=P1000+15
  If TwentyK.state=1 then P1000=P1000+20
  If TwentyfiveK.state=1 then P1000=P1000+25
  If Special.state=1 then
      Credits = Credits + 1
      PlaySound "Knocker"
      CreditPause = 1
  End If
  Kicker1.TimerEnabled=True
End Sub

 'Random Rubber Strikes

Sub Rubber001_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber002_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber003_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber004_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber005_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber006_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber007_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber008_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber009_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber010_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber011_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber012_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub RubberBand11_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber013_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber014_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber015_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber016_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber017_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber018_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber019_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber020_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber021_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber026_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Rubber027_Hit()
  PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
End Sub

Sub Trigger001_Hit()
  PlaySoundAtVol "gate_tgx", ActiveBall, 1
End Sub

Sub Wall001_Hit()
  PlaySoundAtVol "fx_collide", ActiveBall, 1
End Sub

Sub Wall002_Hit()
  PlaySoundAtVol "fx_collide", ActiveBall, 1
End Sub

Sub Wall005_Hit()
  PlaySoundAtVol "fx_collide", ActiveBall, 1
End Sub

Sub Wall006_Hit()
  PlaySoundAtVol "fx_collide", ActiveBall, 1
End Sub

Sub Wall010_Hit()
  PlaySoundAtVol "fx_collide", ActiveBall, 1
End Sub

Sub Wall0110_Hit()
  PlaySoundAtVol "fx_collide", ActiveBall, 1
End Sub

Sub Wall148_Hit()
  PlaySoundAtVol "fx_collide", ActiveBall, 1
End Sub

Sub Ramp003_Hit()
  PlaySoundAtVol "fx_collide", ActiveBall, 1
End Sub

Sub Ramp005_Hit()
  PlaySoundAtVol "fx_collide", ActiveBall, 1
End Sub

Sub BallPlungerStop_Hit()
  PlaySoundAtVol "rubber_hit_3", Plunger, 1
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


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 4 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b, speedfactorx, speedfactory
  Const maxvel=28
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
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

      ' jps ball speed control
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
    Next
End Sub
'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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
        If BOT(b).X < Impacto.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Impacto.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Impacto.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub
