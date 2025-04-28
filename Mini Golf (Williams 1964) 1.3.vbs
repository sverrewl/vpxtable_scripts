'                   |X|
'                   |X|
'                   |||
'                   |||
'                   |^|
'                   | |          .-::":-.
'     .----=--.-':'-; <        .'''..''..'.
'    /=====  /'.'.'.'\ |      /..''..''..''\
'   |====== |.'.'.'.'.||       ;'..''..''..''.;
'    \=====  \'.'.'.'/ /       ;'..''..''..'..;
'     '--=-=-='-:.:-'-`           \..''..''..''/
'                  '.''..''...'
'                  '-..::-'
'
'*        Mini Golf (Williams 1964)
'*        CactusDude - Primary author
'*      JPO - Backglass code and useful tweaks/additons
'*      Thalamus - Golfer primitive

option explicit
Randomize
ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "MiniGolf1964"

Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed

Const B2STableName="MiniGolf1964"

Dim B2SOn 'False = no backglass, True = yes backglass


'VARIABLE***
Const TwoPlayersAlternate = False 'If True, players alternate holes - If False, Player 1 shoots all holes/balls before Player 2 goes


'* this value adjusts score motor behavior - 0 allows you to continue scoring while the score motor is running - 1 sets score motor to behave more like a real EM
Const ScoreMotorAdjustment=1

'* this is a debug setting to use an older scoring routine vs a newer score routine - don't change this value
Const ScoreAdditionAdjustment=1

dim ScoreChecker
dim CheckAllScores,CheckAllPoints
dim sortscores(2)
dim sortpoints(2)
dim sortplayerpoints(2)
dim sortplayers(2)
dim ChimesOn
Dim TextStr,TextStr2
Dim i,xx, LStep, RStep
Dim obj
Dim bgpos
Dim kgpos
Dim dooralreadyopen
Dim kgdooralreadyopen
Dim TargetSpecialLit
Dim Points210counter
Dim Points500counter
Dim Points1000counter
Dim Points2000counter
Dim BallsPerGame
Dim InProgress
Dim BallInPlay
Dim CreditsPerCoin
Dim Score100K(4)
Dim Score(4)
Dim ScoreDisplay(4)
Dim HighScorePaid(4)
Dim HighScore
Dim HighScoreReward
Dim BonusMultiplier
Dim Credits
Dim Match
Dim Replay1
Dim Replay2
Dim Replay3
Dim Replay4
Dim Replay1Paid(4)
Dim Replay2Paid(4)
Dim Replay3Paid(4)
Dim Replay4Paid(4)
Dim BallReplay1
Dim BallReplay2
Dim BallReplay3
Dim ReplayBalls1Paid(4)
Dim ReplayBalls2Paid(4)
Dim ReplayBalls3Paid(4)
Dim TableTilted
Dim TiltCount
Dim debugscore

Dim OperatorMenu, StartGameState

Dim Player
Dim Players
Dim tagback
Dim TempPlayerUp, rst

Dim Moving

'May not actually need all those variables. Might also need more, which I can add here


Sub Table1_Init()
  If Table1.ShowDT = False then
    For each obj in DesktopItems
      obj.visible=False
    next
  End If
  Player1Lights False
  Player2Lights False

  LoadEM
  OperatorMenu=0
  StartGameState=0
  ChimesOn=0
  BallsPerGame=27
  HolesAtStart
  LaunchAngle = 0
  StrokePower = 30
  Strokes = 0
  Strokes2 = 0
  HolesScored = 0
  HolesScored2 = 0
  EMReelStrokes1.ResetToZero
  EMReelHoles1.ResetToZero
  EMReelStrokes2.ResetToZero
  EMReelHoles2.ResetToZero
  Credits=0
  Players = 0

  Forfeit = False
  Forfeit2 = False

  TableTilted=false
  'TiltReel.SetValue(1)

  If Table1.ShowDT = False then
    Controller.B2SSetData 70,1
    Controller.B2SSetData 71,1
    Controller.B2SSetData 72,1
  End if
End Sub

Sub Table1_exit()
  If B2SOn Then Controller.Stop
end sub

Sub Table1_KeyDown(ByVal keycode)


  If keycode = PlungerKey and InProgress=true Then
    GolfSwing.enabled = True'TeeOff
  End If

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    LaunchAngleLeft
    AutoLeft.enabled = True
    Moving = True
  End If

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
    LaunchAngleRight
    AutoRight.enabled = True
    Moving = True
  End If

  If keycode = LeftMagnaSave and InProgress=true and TableTilted=false Then
    StrokePower = 30
    GolfSwing.enabled = True'TeeOff
  End If

  If keycode = RightMagnaSave  and InProgress=true and TableTilted=false Then
    StrokePower = 15
    GolfSwing.enabled = True'TeeOff
  End If

' If keycode = LeftTiltKey Then
'   Nudge 90, 2
'   TiltIt
' End If
'
' If keycode = RightTiltKey Then
'   Nudge 270, 2
'   TiltIt
' End If
'
' If keycode = CenterTiltKey Then
'   Nudge 0, 2
'   TiltIt
' End If
'
' If keycode = MechanicalTilt Then
'   TiltCount=2
'   TiltIt
' End If

  If keycode = AddCreditKey or keycode = 4 then
    If B2SOn Then
      'Controller.B2SSetScorePlayer6 HighScore

    End If
    playsound "CoinIn"
    Credits = Credits + 1
    EMReelCredits.SetValue(Credits)
    If Table1.ShowDT = False then
      Controller.B2SSetCredits Credits
    end if
  end if


   if keycode = StartGameKey and Credits>0 and InProgress=true and Players>0 and Players<2 and BallInPlay<2 then
    Credits=Credits-1
    Controller.B2SSetCredits Credits
    If Credits < 1 Then DOF 232, 0
    EMReelCredits.SetValue(Credits)
    Players = Players + 1
    EMReelPlayer.SetValue(Players)
    If Table1.ShowDT = False then
      Controller.B2SSetCanPlay 2
    end if
'   CanPlayLight1.state=0
'   CanPlayLight2.state=1
    playsound "click"
    HoleLightUpdate.enabled = True
    end if

  if keycode=StartGameKey and Credits>0 and InProgress=false and Players=0 then

    Credits=Credits-1
    Controller.B2SSetCredits Credits

    If Credits < 1 Then DOF 232, 0
    EMReelCredits.SetValue(Credits)
    Players = Players + 1
    EMReelPlayer.SetValue(Players)
'   CanPlayLight1.state=1
    If Table1.ShowDT = False then
      Controller.B2SSetCanPlay 1
      Controller.B2SSetPlayerUp 1
      Controller.B2SSetGameOver 0
    end if
    Player1Lights True
    Player2Lights False
    PlaySound "Mini Golf Start"
    'MatchReel.SetValue(0)
    Player=1
    playsound "startup_norm"
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=true
    rst=0
    BallInPlay=1
    InProgress=true
'   GameOverReel.setvalue(0)

    ResetGame

    TeeLoad.enabled = True

    HoleLightUpdate.enabled = True

  end if



End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = LeftFlipperKey  and InProgress=true and TableTilted=false Then
    AutoLeft.enabled = False
    Moving = False
  End If

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
    AutoRight.enabled = False
    Moving = False
  End If


End Sub

'Start Tilting

'Sub TiltTimer_Timer()
' if TiltCount > 0 then TiltCount = TiltCount - 1
' if TiltCount = 0 then
'   TiltTimer.Enabled = False
' end if
'end sub

'Sub TiltIt()
'   TiltCount = TiltCount + 1
'   if TiltCount = 3 then
'     TableTilted=True
'     'PlasticsOff
'     'BumpersOff
'     LeftFlipper.RotateToStart
'     RightFlipper.RotateToStart
'     TiltReel.SetValue(1)
'     If Table1.ShowDT = False then
'       Controller.B2SSetTilt 1
'     end if
'     BallInPlay=BallsPerGame
'   else
'     TiltTimer.Interval = 500
'     TiltTimer.Enabled = True
'   end if
'
'end sub

'End Tilting

'Start Ball Shooter / KickerTee

Dim LaunchAngle, HoldFire, Strokes, Strokes2
Dim GolferLeft, GolferRight
Dim StrokePower
Dim GolferStanceRot, GolferStanceTransX, GolferStanceTransZ

'********The exact numbers for the GolferRight/Left equations will need tweaked when I put the real prim in, probably
Sub ResetGolfer
  If Forfeit = True OR Forfeit2 = True then Exit Sub end if
  GolferLeft =2
  GolferRight =2
  LaunchAngle =2
    GolferRotation.enabled = True
end sub


Sub LaunchAngleLeft
  LaunchAngle = LaunchAngle - 1
  If LaunchAngle < -60 Then LaunchAngle = -60 End If 'may need adjusted
  GolferLeft = GolferLeft + 0.5
  GolferRight = GolferRight - 0.5
  If GolferLeft > 30 Then GolferLeft = 30 end if      '-Limits the golfers' movement to the kickers' angles
  If GolferRight < -30 Then GolferRight = -30 end if
  GolferRotation.enabled = True
End Sub

Sub LaunchAngleRight
  LaunchAngle = LaunchAngle + 1
  If LaunchAngle >60 Then LaunchAngle = 60 End If 'may need adjusted
  GolferRight = GolferRight + 0.5
  GolferLeft = GolferLeft - 0.5
  If GolferLeft < -30 Then GolferLeft = -30 end if      '-Limits the golfers' movement to the kickers' angles
  If GolferRight > 30 Then GolferRight = 30 end if
  GolferRotation.enabled = True
End Sub

Sub GolferRotation_Timer()
  GolferRotation.enabled = False
  GolferStandIn.RotY = (90 - (GolferLeft) + (GolferRight))
  GolferTorso.RotY = (90 - (GolferLeft) + (GolferRight))
  GolferPlatform.RotY = (90 - (GolferLeft) + (GolferRight))
  End Sub


Sub TeeUp
  If Strokes = BallsPerGame And Players = 1 then InProgress = false : GameOver.enabled = true :Exit Sub end if '1 Player, out of balls
  If Strokes = BallsPerGame And Strokes2 = BallsPerGame then InProgress = false : GameOver.enabled = true :Exit Sub end if '2 Players, both out of balls
  If Strokes = BallsPerGame And Players = 2 then Player = 2 end if '2 Player, Player 1 out of balls
  If Strokes2 = BallsPerGame And Players = 2 then Player = 1 end if '2 Player, Player 2 out of balls
  KickerTee.CreateBall.radius = 15
End Sub

Sub TeeOff
  If HoldFire = True or Moving = True then Exit Sub end if
  PlaySound"TeeOffHit"
  KickerTee.Kick LaunchAngle,StrokePower
  'TeeReset.enabled = True
  TeeReload.enabled = True
  If Player = 1 Then Strokes = Strokes + 1 end if
  If Player = 2 Then Strokes2 = Strokes2 + 1 end if
  UpdateStrokesReels
  CheckStrokes
  HoldFire = True
End Sub

Sub TeeLoad_timer()
  TeeLoad.enabled = False
  TeeReset.enabled = True
End Sub

Sub TeeReload_timer()
  TeeReload.enabled = False
  PlaySound"BallReload"
  TeeReset.enabled = True
End Sub

Sub TeeReset_timer()
  TeeReset.enabled = False
  HoldFire = False
  TeeUp
End Sub

Sub GolfSwing_timer()
  GolfSwing.enabled = False
  If HoldFire = True or Moving = True then Exit Sub end if
  TeeOff
  GolferTorso.ObjRotX = (-15)
  GolferTorso.TransX = (-30)
  GolferTorso.TransZ = (GolferLeft*(2/3))
  GolfSwingReset.enabled = True
End Sub

Sub GolfSwingReset_timer()
  GolfSwingReset.enabled = False
  GolferTorso.ObjRotX = 0
  GolferTorso.TransX = 0
  GolferTorso.TransZ = 0
End Sub


Dim Forfeit, Forfeit2

Sub CheckStrokes
  If Players = 1 Then Exit Sub end If
  If Strokes < BallsPerGame AND Strokes2 < BallsPerGame Then Exit Sub end If
  If Forfeit = True Then Exit Sub end If
  If Forfeit2 = True Then Exit Sub end If


  If Strokes = BallsPerGame Then
    PlaySound"knocker"
    DefaultToTwo
    Forfeit = True
  end if

  If Strokes2 = BallsPerGame Then
    PlaySound"knocker"
    DefaultToOne
    Forfeit2 = True
  end if

  'Could also have the B2S and/or Desktop update who the player is when that happenes here?
End Sub

Sub DefaultToOne
  DefaultTo1.enabled = Ture
End Sub

Sub DefaultTo1_timer()
  DefaultTo1.enabled = False
  Player = 1
  Player1Lights True
  Player2Lights False
  If Table1.ShowDT = False then
    Controller.B2SSetPlayerUp 1
  end if
  If TwoPlayersAlternate = True then ResumeSinglePlayerGame.enabled = True end if
End Sub

Sub DefaultToTwo
  DefaultTo2.enabled = True
End Sub

Sub DefaultTo2_timer()
  DefaultTo2.enabled = False
  Player = 2
  Player1Lights False
  Player2Lights True
  If Table1.ShowDT = False then
    Controller.B2SSetPlayerUp 2
  end if
  If TwoPlayersAlternate = True then Exit Sub end if
    Hole1Up = True
    Hole2Up = False
    Hole3Up = False
    Hole4Up = False
    Hole5Up = False
    Hole6Up = False
    Hole7Up = False
    Hole8Up = False
    Hole9Up = False
    HoleLightUpdate.enabled = True
End Sub


Sub AutoLeft_Timer()
  LaunchAngleLeft
End Sub

Sub AutoRight_Timer()
  LaunchAngleRight
End Sub



'End Ball Shooter / KickerTee

Sub GameOver_timer()
  GameOver.enabled = false
  AutoLeft.enabled = False
  AutoRight.enabled = False
  InProgress = False
  HoldFire = True
  KickerTee.DestroyBall
  PlaySound"Chime100"
  Players = 0
  Light001.state = False
  EMReelPlayer.ResetToZero
  If Table1.ShowDT = False then
    Controller.B2SSetPlayerUp 0
    Controller.B2SSetCanPlay 0
    Controller.B2SSetGameOver 1
  end if
  Player2Lights False
End Sub

Sub ResetGame
  Strokes = 0
  Strokes2 = 0
  HolesScored = 0
  HolesScored2 = 0
  Forfeit = False
  Forfeit2 = False
  HoldFire = False
  HolesAtStart
  TableTilted = False
  EMReelStrokes1.ResetToZero
  EMReelHoles1.ResetToZero
  EMReelStrokes2.ResetToZero
  EMReelHoles2.ResetToZero
  KickerTee.DestroyBall
End Sub


'Start Desktop reels code

Sub UpdateStrokesReels
  EMReelStrokes1.SetValue Strokes
  EMReelStrokes2.SetValue Strokes2
  If Table1.ShowDT = False then
    Controller.B2SSetScore 1, Strokes
    Controller.B2SSetScore 2, Strokes2
  end if
End Sub

Sub UpdateHolesReels
  EMReelHoles1.SetValue HolesScored
  EMReelHoles2.SetValue HolesScored2
  If Table1.ShowDT = False then
    Controller.B2SSetScore 3, HolesScored
    Controller.B2SSetScore 4, HolesScored2
  end if
End Sub

'Start Hole Targeting

Dim HolesScored, HolesScored2

Dim Hole1Up, Hole2Up, Hole3Up, Hole4Up, Hole5Up, Hole6Up, Hole7Up, Hole8Up, Hole9Up

Sub HolesAtStart
  Hole1Up = True
  Hole2Up = False
  Hole3Up = False
  Hole4Up = False
  Hole5Up = False
  Hole6Up = False
  Hole7Up = False
  Hole8Up = False
  Hole9Up = False
End Sub

Sub DoHoles(ThisPlayer)
  PlaySound"HitItIn"
  If Player = 1 Then HolesScored = HolesScored + 1 end if
  If Player = 2 Then HolesScored2 = HolesScored2 + 1 end if
  UpdateHolesReels
  tagback = false

    if Players = 2 then
      if TwoPlayersAlternate = True AND ThisPlayer = 1 then
          Player = 2
          Player1Lights False
          Player2Lights True
          tagback = true
          If Table1.ShowDT = False then
            Controller.B2SSetPlayerUp 2
          end if
      end if
      if TwoPlayersAlternate = True AND ThisPlayer = 2 AND tagback = false then
          Player = 1
          Player1Lights True
          Player2Lights False
          If Table1.ShowDT = False then
            Controller.B2SSetPlayerUp 1
          end if
      end if
      if TwoPlayersAlternate = False and ThisPlayer = 2 then
          Player = 2
          Player1Lights False
          Player2Lights True
          If Table1.ShowDT = False then
            Controller.B2SSetPlayerUp 2
          end if
      end if

    if Players = 1 then
        Player = 1
        Player1Lights True
        Player2Lights False
        If Table1.ShowDT = False then
          Controller.B2SSetPlayerUp 1
        end if
    end if
'   end if


      If Forfeit2 = True OR Strokes2 = BallsPerGame then
        Player = 1
        Player1Lights True
        Player2Lights False
        If Table1.ShowDT = False then
          Controller.B2SSetPlayerUp 1
        ResumeSinglePlayerGame.enabled = True
        end if
      end if

      If Forfeit = True OR Strokes = BallsPerGame then
        Player = 2
        Player1Lights False
        Player2Lights True
        If Table1.ShowDT = False then
          Controller.B2SSetPlayerUp 2
        end if
      end if
    end if

  If Players = 1 OR TwoPlayersAlternate = False then Exit Sub end if
  ResetGolfer
End Sub


Sub Kicker001_Hit
  Kicker001.DestroyBall
  If Hole1Up = False then PlaySoundAtVol"fx_ball_drop4", ActiveBall, 1:Exit Sub end if
  if Players = 1 then
    Hole1Up = False
    Hole2Up = True
    HoleLightUpdate.enabled = True
  end if
  if Players = 2 AND Player = 2 then
    Hole1Up = False
    Hole2Up = True
    HoleLightUpdate.enabled = True
  end If
  if Players = 2 AND Player = 1 AND TwoPlayersAlternate = False then
    Hole1Up = False
    Hole2Up = True
    HoleLightUpdate.enabled = True
  end if
  DoHoles Player
End Sub

Sub Kicker002_Hit
  Kicker002.DestroyBall
  If Hole2Up = False then PlaySoundAtVol"fx_ball_drop4", ActiveBall, 1:Exit Sub end if
  if Players = 1 or (Players = 2 and Player = 2) or (Players = 2 and Player = 1 and TwoPlayersAlternate = False) then
    Hole2Up = False
    Hole3Up = True
    HoleLightUpdate.enabled = True
  end if
  DoHoles Player
End Sub

Sub Kicker003_Hit
  Kicker003.DestroyBall
  If Hole3Up = False then PlaySoundAtVol"fx_ball_drop4", ActiveBall, 1:Exit Sub end if
  if Players = 1 or (Players = 2 and Player = 2) or (Players = 2 and Player = 1 and TwoPlayersAlternate = False) then
    Hole3Up = False
    Hole4Up = True
    HoleLightUpdate.enabled = True
  end if
  DoHoles Player
End Sub

Sub Kicker004_Hit
  Kicker004.DestroyBall
  If Hole4Up = False then PlaySoundAtVol"fx_ball_drop4", ActiveBall, 1:Exit Sub end if
  if Players = 1 or (Players = 2 and Player = 2) or (Players = 2 and Player = 1 and TwoPlayersAlternate = False) then
    Hole4Up = False
    Hole5Up = True
    HoleLightUpdate.enabled = True
  end if
  DoHoles Player
End Sub

Sub Kicker005_Hit
  Kicker005.DestroyBall
  If Hole5Up = False then PlaySoundAtVol"fx_ball_drop4", ActiveBall, 1:Exit Sub end if
  if Players = 1 or (Players = 2 and Player = 2) or (Players = 2 and Player = 1 and TwoPlayersAlternate = False) then
    Hole5Up = False
    Hole6Up = True
    HoleLightUpdate.enabled = True
  end if
  DoHoles Player
End Sub

Sub Kicker006_Hit
  Kicker006.DestroyBall
  If Hole6Up = False then PlaySoundAtVol"fx_ball_drop4", ActiveBall, 1:Exit Sub end if
  if Players = 1 or (Players = 2 and Player = 2) or (Players = 2 and Player = 1 and TwoPlayersAlternate = False) then
    Hole6Up = False
    Hole7Up = True
    HoleLightUpdate.enabled = True
  end if
  DoHoles Player
End Sub

Sub Kicker007_Hit
  Kicker007.DestroyBall
  If Hole7Up = False then PlaySoundAtVol"fx_ball_drop4", ActiveBall, 1:Exit Sub end if
  if Players = 1 or (Players = 2 and Player = 2) or (Players = 2 and Player = 1 and TwoPlayersAlternate = False) then
    Hole7Up = False
    Hole8Up = True
    HoleLightUpdate.enabled = True
  end if
  DoHoles Player
End Sub

Sub Kicker008_Hit
  Kicker008.DestroyBall
  If Hole8Up = False then PlaySoundAtVol"fx_ball_drop4", ActiveBall, 1:Exit Sub end if
  if Players = 1 or (Players = 2 and Player = 2) or (Players = 2 and Player = 1 and TwoPlayersAlternate = False) then
    Hole8Up = False
    Hole9Up = True
    HoleLightUpdate.enabled = True
  end if
  DoHoles Player
End Sub

Sub Kicker009_Hit
  Kicker009.DestroyBall
  If Hole9Up = False then PlaySoundAtVol"fx_ball_drop4", ActiveBall, 1:Exit Sub end if
  PlaySound"HitItIn"
  Player1Lights False
  If Player = 1 Then HolesScored = HolesScored + 1 end if
  If Player = 2 Then HolesScored2 = HolesScored2 + 1 end if
  UpdateHolesReels

  if Players = 1 or (Players = 2 and Player = 2) then
    GameOver.enabled = true
  end if

  if Player = 1 AND TwoPlayersAlternate = False then
    Player = 2
    Hole1Up = True
    Hole9Up = False
    HoleLightUpdate.enabled = True
    Player1Lights False
    Player2Lights True
    If Table1.ShowDT = False then
      Controller.B2SSetPlayerUp 2
    end if
  Else
    Player2Lights true
    If Table1.ShowDT = False then
      Controller.B2SSetPlayerUp 2
    end if
    Player = 2
    ResetGolfer
  end if
End Sub

Sub HoleLightUpdate_timer()
  HoleLightUpdate.enabled = False
  If  Hole1Up = True Then Light001.state = True:Light002.state = False:Light003.state = False:Light004.state = False:Light005.state = False:Light006.state = False:Light007.state = False:Light008.state = False:Light009.state = False end if
  If  Hole2Up = True Then Light001.state = False:Light002.state = True:Light003.state = False:Light004.state = False:Light005.state = False:Light006.state = False:Light007.state = False:Light008.state = False:Light009.state = False end if
  If  Hole3Up = True Then Light001.state = False:Light002.state = False:Light003.state = True:Light004.state = False:Light005.state = False:Light006.state = False:Light007.state = False:Light008.state = False:Light009.state = False end if
  If  Hole4Up = True Then Light001.state = False:Light002.state = False:Light003.state = False:Light004.state = True:Light005.state = False:Light006.state = False:Light007.state = False:Light008.state = False:Light009.state = False end if
  If  Hole5Up = True Then Light001.state = False:Light002.state = False:Light003.state = False:Light004.state = False:Light005.state = True:Light006.state = False:Light007.state = False:Light008.state = False:Light009.state = False end if
  If  Hole6Up = True Then Light001.state = False:Light002.state = False:Light003.state = False:Light004.state = False:Light005.state = False:Light006.state = True:Light007.state = False:Light008.state = False:Light009.state = False end if
  If  Hole7Up = True Then Light001.state = False:Light002.state = False:Light003.state = False:Light004.state = False:Light005.state = False:Light006.state = False:Light007.state = True:Light008.state = False:Light009.state = False end if
  If  Hole8Up = True Then Light001.state = False:Light002.state = False:Light003.state = False:Light004.state = False:Light005.state = False:Light006.state = False:Light007.state = False:Light008.state = True:Light009.state = False end if
  If  Hole9Up = True Then Light001.state = False:Light002.state = False:Light003.state = False:Light004.state = False:Light005.state = False:Light006.state = False:Light007.state = False:Light008.state = False:Light009.state = True end if
    If Table1.ShowDT = False then
    Controller.B2SSetData 81,0
    Controller.B2SSetData 82,0
    Controller.B2SSetData 83,0
    Controller.B2SSetData 84,0
    Controller.B2SSetData 85,0
    Controller.B2SSetData 86,0
    Controller.B2SSetData 87,0
    Controller.B2SSetData 88,0
    Controller.B2SSetData 89,0
    If  Hole1Up = True Then Controller.B2SSetData 81,1 end if
    If  Hole2Up = True Then Controller.B2SSetData 82,1 end if
    If  Hole3Up = True Then Controller.B2SSetData 83,1 end if
    If  Hole4Up = True Then Controller.B2SSetData 84,1 end if
    If  Hole5Up = True Then Controller.B2SSetData 85,1 end if
    If  Hole6Up = True Then Controller.B2SSetData 86,1 end if
    If  Hole7Up = True Then Controller.B2SSetData 87,1 end if
    If  Hole8Up = True Then Controller.B2SSetData 88,1 end if
    If  Hole9Up = True Then Controller.B2SSetData 89,1 end if
  end if
End Sub

Sub ResumeSinglePlayerGame_timer()
  ResumeSinglePlayerGame.enabled = False
  If HolesScored = 0 Then Hole1Up = True:Hole2Up = False:Hole3Up = False:Hole4Up = False:Hole5Up = False:Hole6Up = False:Hole7Up = False:Hole8Up = False:Hole9Up = False end if
  If HolesScored = 1 Then Hole1Up = False:Hole2Up = True:Hole3Up = False:Hole4Up = False:Hole5Up = False:Hole6Up = False:Hole7Up = False:Hole8Up = False:Hole9Up = False end if
  If HolesScored = 2 Then Hole1Up = False:Hole2Up = False:Hole3Up = True:Hole4Up = False:Hole5Up = False:Hole6Up = False:Hole7Up = False:Hole8Up = False:Hole9Up = False end if
  If HolesScored = 3 Then Hole1Up = False:Hole2Up = False:Hole3Up = False:Hole4Up = True:Hole5Up = False:Hole6Up = False:Hole7Up = False:Hole8Up = False:Hole9Up = False end if
  If HolesScored = 4 Then Hole1Up = False:Hole2Up = False:Hole3Up = False:Hole4Up = False:Hole5Up = True:Hole6Up = False:Hole7Up = False:Hole8Up = False:Hole9Up = False end if
  If HolesScored = 5 Then Hole1Up = False:Hole2Up = False:Hole3Up = False:Hole4Up = False:Hole5Up = False:Hole6Up = True:Hole7Up = False:Hole8Up = False:Hole9Up = False end if
  If HolesScored = 6 Then Hole1Up = False:Hole2Up = False:Hole3Up = False:Hole4Up = False:Hole5Up = False:Hole6Up = False:Hole7Up = True:Hole8Up = False:Hole9Up = False end if
  If HolesScored = 7 Then Hole1Up = False:Hole2Up = False:Hole3Up = False:Hole4Up = False:Hole5Up = False:Hole6Up = False:Hole7Up = False:Hole8Up = True:Hole9Up = False end if
  If HolesScored = 8 Then Hole1Up = False:Hole2Up = False:Hole3Up = False:Hole4Up = False:Hole5Up = False:Hole6Up = False:Hole7Up = False:Hole8Up = False:Hole9Up = True end if
  HoleLightUpdate.enabled = True
  Player = 1
End Sub


Sub prt(stuff)
   Dim WriteFileObject, WriteFileName
   Set WriteFileObject = CreateObject("Scripting.FileSystemObject")
   Set WriteFileName = WriteFileObject.OpenTextFile("c:\ESD\VPOut.txt", 8, True )
   WriteFileName.WriteLine(stuff)
   WriteFileName.close
End Sub

sub Player1Lights(TorF)
  For each obj in Player1
    obj.State = TorF
  next
end sub

sub Player2Lights(TorF)
  For each obj in Player2
    obj.State = TorF
  next
end sub

'End Hole Targeting

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

