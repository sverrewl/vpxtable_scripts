'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~                                       ~~~
'~~~                Big Game               ~~~
'~~~            (Rock-Ola 1935)            ~~~
'~~~                                       ~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Made by CactusDude
'Released September 2023
'Version 0.22
' Thalamus - increased the ball rollling volume, could be improved by a solenoid call from kickout in the bottom

Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

Const cGameName = "BigGameRO"


'~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~  Configuration  ~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~
Const BallsperGame = 8
Const FreePlay = True
Const RollingSoundFactor = 2.5    'set sound level factor here for Ball Rolling Sound, 1=default level

Const VolumeMultiplier = 3        'adjusts table sound volume

Const ResetHighscore = False    'enable to manually reset the Highscores

Const BallSize = 20         'Ball radius

Const BackglassUsed = True          'True to use the backglass, False to turn it off

Dim i,HighScore,Credits,GameStarted,BallinPlay,BallReleaseLocked,BallsonBoard
Dim LaneClear

'--------------------------
'------  Table Init  ------
'--------------------------

Sub BigGameRO_Init
  FinalBall = False
  LoadHighScore
  If ResetHighScore Then
    SetDefaultHSTD
  end If
  UpdatePostIt

  Randomize

  if FreePlay then
    Credits = 5
    DOF 126, DOFOn
  end If

' TiltKicker.CreateBall
  TiltCount = 0
  BallinPlay = 0
  GameStarted = False
'If BallsonBoard=8 then BallReleaseLocked=True End if
  UpdateScoreReel.enabled = True
  LaneClear = True
  If BackglassUsed = True then
    BackglassOff
  End if
End Sub


'~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~  Game Features  ~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Sub MainTimer_Timer
' if (BallsonBoard = BallsperGame) and (not BallReleaseLocked) Then
'   If BallVel(LBall) = 0 Then
'     GameEndTimer.enabled = True
'   Else
'     if GameEndTimer.enabled Then
'       GameEndTimer.enabled = False
'     end If
'   end If
' end If
'End Sub

Sub StartGame
  if Credits > 0 Then
    if not GameStarted then
      GameEndTimer.interval = 2000
      Occupancy = 0
      BallReleaseLocked = False
      Playerscore = 0
      Oldscore = 0
      Val10 = 0
      Val100 = 0
      Val1000 = 0
      Val10000 = 0
      Val100000 = 0
      LeftTrail = False
      RightTrail = False
      LeftReturn = False
      RightReturn = False
      UpdateScoreReel.enabled = True
      GameStarted = True
      BallinPlay = 0
'     P_BallinPlay.image = CStr(0)
'     P_BallinPlay1.image = CStr(0)
      Credits = Credits - 1
      If Credits < 1 Then DOF 126, DOFOff
      if FreePlay and (Credits = 0) then
        Credits = 5
        DOF 126, DOFOn
      end If
    end If
  end If
End Sub

Sub NextBall
  If LaneClear = False then Exit Sub end if
  if BallsonBoard < BallsperGame Then 'was <=, which seemed wrong
    BallReleaseLocked = False
    PlaySoundAtVol "BallLift", BallRelease, 1
'   BallinPlay = BallinPlay + 1
    end If
    BallRelease.Timerenabled = True
' End If
End Sub

Sub BallRelease_Timer
  BallRelease.Timerenabled = False
  BallRelease.CreateSizedBall(25)
  BallRelease.Kick 270, 7
  DOF 122, DOFPulse
End Sub

Sub CheckFinish
  If Occupancy = 8 AND BallsonBoard >= 8 Then
    BallOccupancy.Enabled = False
    LastBall.Enabled = True
  End if
End Sub

Sub LastBall_Timer
  LastBall.Enabled = False
  EndGame
End Sub

Sub EndGame
  CheckNewHighScorePostIt(Playerscore)
  PlaySound "Chime"
  GameStarted = False
  BallinPlay = 0
End Sub

'--------------------------------
'------  Helper Functions  ------
'--------------------------------
Sub PostWall_Hit
  P_Left1.RotY = 185
  P_Left2.RotY = 185
  P_Left3.RotY = 185
  PostWall.timerenabled = True
End Sub

Sub PostWall_Timer
  PostWall.timerenabled = False
  P_Left1.RotY = 180
  P_Left2.RotY = 180
  P_Left3.RotY = 180
End Sub

Sub Drain_Hit()
' PlaySound "drain3",0,0.2,0,0.25
' DOF 121, DOFPulse
  Drain.DestroyBall
End Sub

Dim BOT2, LBall, FinalBall, BallIsGoingOutFromShooterLane
Dim Occupancy

Sub EnterPF_Hit
        BallIsGoingOutFromShooterLane = True
End Sub

Sub EnterPF001_Hit
  If BallIsGoingOutFromShooterLane = True Then
    BallIsGoingOutFromShooterLane = False
    BallsonBoard = BallsonBoard + 1
' If BallsonBoard >= 8 Then
'   If BallsonBoard = 8 then BallReleaseLocked = True
'   FinalBall = True
  Else
    BallReleaseLocked = False
  End If
' CheckFinish
  Set LBall = ActiveBall
' End If
End Sub

Dim BoardDir
Sub MoveBoardTimer_Timer

  P_Board.Transy = P_Board.Transy + BoardDir
  if P_Board.Transy < -40 Then
    KickerHoleWall.isdropped = True
  Else
    KickerHoleWall.isdropped = False
  End If

  if P_Board.Transy < -57 Then
    BoardDir = -Boarddir
    HoldBoardTimer.enabled = True
    MoveBoardTimer.enabled = False
  end If

  if P_Board.Transy >= 0 Then
    P_Board.Transy = 0
    MoveBoardTimer.enabled = False
    StartGame
  end If
End Sub

Sub HoldBoardTimer_Timer
  HoldBoardTimer.enabled = False
  MoveBoardTimer.enabled = True
' PlaySound "BallIn",0,0.75,0,0.25,0,0,1,0
end Sub

Sub GateLeft_Hit
  FlapLeft.ObjRotZ = -15
' FlapRight.ObjRotZ = 15
  FlapReset.enabled = True
End Sub


Sub GateRight_Hit
' FlapLeft.ObjRotZ = -15
  FlapRight.ObjRotZ = 15
  FlapReset.enabled = True
End Sub

Sub FlapReset_timer()
  FlapReset.enabled = False
  FlapLeft.ObjRotZ = 45
  FlapRight.ObjRotZ = -45
End Sub

Dim Val10,Val100,Val1000,Val10000,Val100000,Score10,Score100,Score1000,Score10000,Score100000,TempScore,Oldscore
Dim PlayerScore,Special1,Special2,Special3

'-------------------------------
'------  Keybord Handler  ------
'-------------------------------
Sub BigGameRO_KeyDown(ByVal keycode)
  If keycode = StartGameKey Then
      if BallsonBoard>7 or BallsonBoard=0 then
      If LastBall.Enabled = False then
      BallsonBoard = 0
      Ballinplay = 0
      BoardDir = -1
      MoveBoardTimer.enabled = True
      PlaySoundAtVol "CoinStart", Drain, 1
      PlayerScore=0
      ClearKickers
    end if
    end If
  End if

  If keycode = PlungerKey Then
    Plunger.PullBack
'   PlaySound "plungerpull",0,1,Pan(Plunger),0.25,0,0,1,0
  End If

  If keycode = LeftFlipperKey Then

  End If

If keycode = RightFlipperKey or keycode = RightMagnaSave and LaneClear = True then
    if GameStarted and BallReleaseLocked = False Then
    NextBall
'   PlaySound "ballrelease"
  end If
End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
'   TiltIt
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
'   TiltIt
  End If

  If keycode = CenterTiltKey OR keycode = 57 Then
    Nudge 0, 2
    StuckBallSweep
'   TiltIt
  End If

End Sub

Sub BigGameRO_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySoundAtVol "PlungerHit", Plunger, 1
  End If

  If keycode = LeftFlipperKey Then

  End If

  If keycode = RightFlipperKey Then

  End If

End Sub

If BigGameRO.ShowDT = false then
'    Scoretext.Visible = false
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                  Bumper Spring Movement - REPLACE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Sub SpringWall_Hit 'Not doing anything - oh well
' Spring.transX = +10000
' SpringReset.enabled = True
'End Sub

'Sub SpringReset_Timer()
' SpringReset.enabled = False
' Spring.transX = -10000
'End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                  Lane Check
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub ShooterLaneLaunch_Hit
  LaneClear = False
End Sub

Sub Gate001_Hit
  LaneClear = True
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                Balls on Board Check
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub BallOccupancy_Timer
  If Occupancy = 8 then
    BallReleaseLocked = True
    CheckFinish
  end if

  If BallsonBoard = 8 OR LaneClear = False then
    BallReleaseLocked = True
  end if

  If BallsonBoard < 8 AND LaneClear = True then
    BallReleaseLocked = False
  end if
end Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                  Game/Backglass Reseting
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Sub StuckBallSweep
  If BallsonBoard = 0 then ClearKickers
  PlayerScore = 0
end Sub

Sub ClearKickers
      KTriggerHole.DestroyBall
      K2500A.DestroyBall 'Jaguar
      K2500B.DestroyBall
      K2500C.DestroyBall
      K2000A.DestroyBall 'Elephant
      K2000B.DestroyBall
      K2000C.DestroyBall
      K2000C.DestroyBall
      K1500A.DestroyBall 'Tiger
      K1500B.DestroyBall
      K1500C.DestroyBall
      K1000A.DestroyBall 'Ape
      K1000B.DestroyBall
      K500TwinL.DestroyBall
      K500TwinR.DestroyBall
      K500TRailL.DestroyBall
      K500TRailR.DestroyBall
      K500A.DestroyBall
      K500B.DestroyBall
      K500C.DestroyBall
      K500D.DestroyBall
      K500E.DestroyBall
      K500F.DestroyBall
      BackglassOff
      BallsonBoard = 0
      BallReleaseLocked = False
      ResetWins
      UpdateScoreReel.enabled = True
end Sub

Sub BackglassOff
'   Controller.B2SSetScorePlayer1 (0) '**I'll need to update with backglass stuff
'   Controller.B2SSetScorePlayer2 (0) '**Also, should consider making a variable so B2S is optional for desktop users (and add in desktop scoring somehow?)
'   Controller.B2SSetScorePlayer3 (0)
'   Controller.B2SSetScorePlayer4 (0)
    Controller.B2SSetData 101,0
    Controller.B2SSetData 102,0
    Controller.B2SSetData 103,0
    Controller.B2SSetData 104,0
    Controller.B2SSetData 105,0
    Controller.B2SSetData 106,0
    Controller.B2SSetData 107,0
    Controller.B2SSetData 108,0
    Controller.B2SSetData 109,0
    Controller.B2SSetData 110,0
    Controller.B2SSetData 111,0
    Controller.B2SSetData 112,0
End Sub

Sub ResetWins '**Win subs need renamed and updated for this table
  Triggered = False
  SkillShotWin = False
  PartridgeWin = False
  PhesantWin = False
  RabbitWin = False
  SquirrelWin = False
  DuckWin = False
'And Hits
  PartridgeHitA = False '**Need redone
  PartridgeHitB = False
  PartridgeHitC = False
  PhesantHitA = False
  PhesantHitB = False
  PhesantHitC = False
  RabbitHitA = False
  RabbitHitB = False
  RabbitHitC = False
  SquirrelHitA = False
  SquirrelHitB = False
  DuckHitA = False
  DuckHitB = False
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                  Backglass/Scoring
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim Controller

Set Controller = CreateObject("B2S.Server")
Controller.B2SName = ""
Controller.Run

'Skill Shot
Dim SkillShotWin
Sub KSkillShot_Hit:Controller.B2SSetData 108,1:SkillShotWin = True
  Addpoints(2000)
  SkillShotCheck
End sub

Sub KSkillShot_unHit:Addpoints(-2000):SkillShotWin = False
End Sub

Sub SkillShotCheck
  If PartridgeHitA=True and PartridgeHitB=True and PartridgeHitC=True and SkillShotWin = True then PartridgeWin=True end If
  If PartridgeWin = True then Controller.B2SSetScorePlayer1(8):Controller.B2SSetData 118,1 end if
  If PhesantHitA=True and PhesantHitB=True and PhesantHitC=True and SkillShotWin = True then PhesantWin=True end If
  If PhesantWin = True then Controller.B2SSetScorePlayer2(6):Controller.B2SSetData 117,1 end if
  If RabbitHitA=True and RabbitHitB=True and RabbitHitC=True and SkillShotWin = True then RabbitWin=True end If
  If RabbitWin = True then Controller.B2SSetScorePlayer3(4):Controller.B2SSetData 116,1 end if
  If RabbitHitA=True and RabbitHitB=True and RabbitHitC=True and SkillShotWin = True then RabbitWin=True end If
  If RabbitWin = True then Controller.B2SSetScorePlayer3(4):Controller.B2SSetData 116,1 end if
  If SquirrelHitA=True and SquirrelHitB=True and SkillShotWin = True then SquirrelWin=True end If
  If SquirrelWin = True then Controller.B2SSetScorePlayer4(2):Controller.B2SSetData 115,1 end if
  If DuckHitA=True and DuckHitB=True and SkillShotWin = True then DuckWin=True end If
  If DuckWin = True then Controller.B2SSetScorePlayer4(2):Controller.B2SSetData 115,1 end if
End Sub

'Partridge
Dim PartridgeHitA,PartridgeHitB,PartridgeHitC,PartridgeWin
Sub KPartridgeA_Hit:Controller.B2SSetData 112,1:PartridgeHitA=True
  Addpoints(1500)
  If PartridgeHitA=True and PartridgeHitB=True and PartridgeHitC=True and SkillShotWin = True then PartridgeWin=True end If
  If PartridgeWin = True then Controller.B2SSetScorePlayer1(8):Controller.B2SSetData 118,1 end if end sub
Sub KPartridgeB_Hit:Controller.B2SSetData 113,1:PartridgeHitB=True
  Addpoints (1500)
  If PartridgeHitA=True and PartridgeHitB=True and PartridgeHitC=True and SkillShotWin = True then PartridgeWin=True end If
  If PartridgeWin = True then Controller.B2SSetScorePlayer1(8):Controller.B2SSetData 118,1 end if end sub
Sub KPartridgeC_Hit:Controller.B2SSetData 114,1:PartridgeHitC=True
  Addpoints(1500)
  If PartridgeHitA=True and PartridgeHitB=True and PartridgeHitC=True and SkillShotWin = True then PartridgeWin=True end If
  If PartridgeWin = True then Controller.B2SSetScorePlayer1(8):Controller.B2SSetData 118,1 end if end sub

Sub KPartridgeA_unHit
  Addpoints(-1500)
  PartridgeHitA = False
  PartridgeWin = False
End Sub

Sub KPartridgeB_unHit
  Addpoints(-1500)
  PartridgeHitB = False
  PartridgeWin = False
End Sub

Sub KPartridgeC_unHit
  Addpoints(-1500)
  PartridgeHitC = False
  PartridgeWin = False
End Sub

'Phesant
Dim PhesantHitA,PhesantHitB,PhesantHitC,PhesantWin
Sub KPhesantA_Hit:Controller.B2SSetData 109,1:PhesantHitA=True
  Addpoints(1000)
  If PhesantHitA=True and PhesantHitB=True and PhesantHitC=True and SkillShotWin = True then PhesantWin=True end If
  If PhesantWin = True then Controller.B2SSetScorePlayer2(6):Controller.B2SSetData 117,1 end if end sub
Sub KPhesantB_Hit:Controller.B2SSetData 110,1:PhesantHitB=True
  Addpoints(1000)
  If PhesantHitA=True and PhesantHitB=True and PhesantHitC=True and SkillShotWin = True then PhesantWin=True end If
  If PhesantWin = True then Controller.B2SSetScorePlayer2(6):Controller.B2SSetData 117,1 end if end sub
Sub KPhesantC_Hit:Controller.B2SSetData 111,1:PhesantHitC=True
  Addpoints(1000)
  If PhesantHitA=True and PhesantHitB=True and PhesantHitC=True and SkillShotWin = True then PhesantWin=True end If
  If PhesantWin = True then Controller.B2SSetScorePlayer2(6):Controller.B2SSetData 117,1 end if end sub

Sub KPhesantA_unHit
  Addpoints(-1000)
  PhesantHitA = False
  PhesantWin = False
End Sub

Sub KPhesantB_unHit
  Addpoints(-1000)
  PhesantHitB = False
  PhesantWin = False
End Sub

Sub KPhesantC_unHit
  Addpoints(-1000)
  PhesantHitC = False
  PhesantWin = False
End Sub

'Rabbit
Dim RabbitHitA,RabbitHitB,RabbitHitC,RabbitWin
Sub KRabbitA_Hit:Controller.B2SSetData 105,1:RabbitHitA=True
  Addpoints(800)
  If RabbitHitA=True and RabbitHitB=True and RabbitHitC=True and SkillShotWin = True then RabbitWin=True end If
  If RabbitWin = True then Controller.B2SSetScorePlayer3(4):Controller.B2SSetData 116,1 end if end sub
Sub KRabbitB_Hit:Controller.B2SSetData 106,1:RabbitHitB=True
  Addpoints(800)
  If RabbitHitA=True and RabbitHitB=True and RabbitHitC=True and SkillShotWin = True then RabbitWin=True end If
  If RabbitWin = True then Controller.B2SSetScorePlayer3(4):Controller.B2SSetData 116,1 end if end sub
Sub KRabbitC_Hit:Controller.B2SSetData 107,1:RabbitHitC=True
  Addpoints(800)
  If RabbitHitA=True and RabbitHitB=True and RabbitHitC=True and SkillShotWin = True then RabbitWin=True end If
  If RabbitWin = True then Controller.B2SSetScorePlayer3(4):Controller.B2SSetData 116,1 end if end sub

Sub KRabbitA_unHit
  Addpoints(-800)
  RabbitHitA = False
  RabbitWin = False
End Sub

Sub KRabbitB_unHit
  Addpoints(-800)
  RabbitHitB = False
  RabbitWin = False
End Sub

Sub KRabbitC_unHit
  Addpoints(-800)
  RabbitHitC = False
  RabbitWin = False
End Sub

'Squirrel
Dim SquirrelHitA,SquirrelHitB,SquirrelWin
Sub KSquirrelA_Hit:Controller.B2SSetData 103,1:SquirrelHitA=True
  Addpoints(600)
  If SquirrelHitA=True and SquirrelHitB=True and SkillShotWin = True then SquirrelWin=True end If
  If SquirrelWin = True then Controller.B2SSetScorePlayer4(2):Controller.B2SSetData 115,1 end if end sub
Sub KSquirrelB_Hit:Controller.B2SSetData 104,1:SquirrelHitB=True
  Addpoints(600)
  If SquirrelHitA=True and SquirrelHitB=True and SkillShotWin = True then SquirrelWin=True end If
  If SquirrelWin = True then Controller.B2SSetScorePlayer4(2):Controller.B2SSetData 115,1 end if end sub

Sub KSquirrelA_unHit
  Addpoints(-600)
  SquirrelHitA = False
  SquirrelWin = False
End Sub

Sub KSquirrelB_unHit
  Addpoints(-600)
  SquirrelHitB = False
  SquirrelWin = False
End Sub

'Duck
Dim DuckHitA,DuckHitB,DuckWin

Sub KDuckA_Hit:Controller.B2SSetData 101,1:DuckHitA=True
  Addpoints(500)
  If DuckHitA=True and DuckHitB=True and SkillShotWin = True then DuckWin=True end If
  If DuckWin = True then Controller.B2SSetScorePlayer4(2):Controller.B2SSetData 115,1 end if end sub
Sub KDuckB_Hit:Controller.B2SSetData 102,1:DuckHitB=True
    Addpoints(500)
  If DuckHitA=True and DuckHitB=True and SkillShotWin = True then DuckWin=True end If
  If DuckWin = True then Controller.B2SSetScorePlayer4(2):Controller.B2SSetData 115,1 end if end sub

Sub KDuckA_unHit
  Addpoints(-500)
  DuckHitA = False
  DuckWin = False
End Sub

Sub KDuckB_unHit
  Addpoints(-500)
  DuckHitB = False
  DuckWin = False
End Sub

'Animal occupancy test - still need to properly code wins
Dim Triggered
'TriggerHole
Sub KTriggerHole_Hit
  Triggered = True
  Occupancy = Occupancy + 1
  If BackglassUsed = True Then
    Controller.B2SSetData 101,1
  End if
  PlaySoundAtVol "Trigger gunshot", ActiveBall, 1
End Sub

'Jaguar
Sub K2500A_Hit
  Addpoints (2500)
  Occupancy = Occupancy + 1
  If BackglassUsed = True Then
    Controller.B2SSetData 102,1
  End if
End Sub

Sub K2500B_Hit
  Addpoints (2500)
  Occupancy = Occupancy + 1
  If BackglassUsed = True Then
    Controller.B2SSetData 103,1
  End if
End Sub

Sub K2500C_Hit
  Addpoints (2500)
  Occupancy = Occupancy + 1
  If BackglassUsed = True Then
    Controller.B2SSetData 104,1
  End if
End Sub

'Elephant
Sub K2000A_Hit
  Addpoints (2000)
  Occupancy = Occupancy + 1
  If BackglassUsed = True Then
    Controller.B2SSetData 105,1
  End if
End Sub

Sub K2000B_Hit
  Addpoints (2000)
  Occupancy = Occupancy + 1
  If BackglassUsed = True Then
    Controller.B2SSetData 106,1
  End if
End Sub

Sub K2000C_Hit
  Addpoints (2000)
  Occupancy = Occupancy + 1
  If BackglassUsed = True Then
    Controller.B2SSetData 107,1
  End if
End Sub

'Tiger
Sub K1500A_Hit
  Addpoints (1500)
  Occupancy = Occupancy + 1
  If BackglassUsed = True Then
    Controller.B2SSetData 108,1
  End if
End Sub

Sub K1500B_Hit
  Addpoints (1500)
  Occupancy = Occupancy + 1
  If BackglassUsed = True Then
    Controller.B2SSetData 109,1
  End if
End Sub

Sub K1500C_Hit
  Addpoints (1500)
  Occupancy = Occupancy + 1
  If BackglassUsed = True Then
    Controller.B2SSetData 110,1
  End if
End Sub

'Ape
Sub K1000A_Hit
  Addpoints (1000)
  Occupancy = Occupancy + 1
  If BackglassUsed = True Then
    Controller.B2SSetData 111,1
  End if
End Sub

Sub K1000B_Hit
  Addpoints (1000)
  Occupancy = Occupancy + 1
  If BackglassUsed = True Then
    Controller.B2SSetData 112,1
  End if
End Sub

'Non-lethal/non-return points (500)

Sub K500A_Hit
  Addpoints (500)
  Occupancy = Occupancy + 1
End Sub

Sub K500B_Hit
  Addpoints (500)
  Occupancy = Occupancy + 1
End Sub

Sub K500C_Hit
  Addpoints (500)
  Occupancy = Occupancy + 1
End Sub

Sub K500D_Hit
  Addpoints (500)
  Occupancy = Occupancy + 1
End Sub

Sub K500E_Hit
  Addpoints (500)
  Occupancy = Occupancy + 1
End Sub

Sub K500F_Hit
  Addpoints (500)
  Occupancy = Occupancy + 1
End Sub

'Twin Return
Dim LeftReturn, RightReturn

Sub K500TwinL_hit()
  LeftReturn = True
  Occupancy = Occupancy + 1
  Addpoints(500)
  CheckReturn
End Sub

Sub K500TwinR_hit()
  RightReturn = True
  Occupancy = Occupancy + 1
  Addpoints(500)
  CheckReturn
End Sub

Sub CheckReturn
  If LeftReturn = True AND RightReturn = True then
    K500TwinL.DestroyBall
    K500TwinR.DestroyBall
    LeftReturn = False
    RightReturn = False
    Addpoints(-1000)
    Occupancy = Occupancy - 2
    BallsonBoard = BallsonBoard - 2
    PlaySoundAtVol "Ball Return", K500TwinL, 1
  Else
    Exit Sub
  End If
End Sub

'Lost Trail
Dim LeftTrail, RightTrail

Sub K500TrailL_hit()
  LeftTrail = True
  Occupancy = Occupancy + 1
  Addpoints(500)
  CheckTrail
End Sub

Sub K500TrailR_hit()
  RightTrail = True
  Occupancy = Occupancy + 1
  Addpoints(500)
  CheckTrail
End Sub

Sub CheckTrail
  If LeftTrail = True AND RightTrail = True then
    K500TrailL.DestroyBall
    K500TrailR.DestroyBall
    LeftTrail = False
    RightTrail = False
    Addpoints(-1000)
    Occupancy = Occupancy - 2
    BallsonBoard = BallsonBoard - 2
    PlaySoundAtVol "Ball Return", K500TrailR, 1
  Else
    Exit Sub
  End If
End Sub

'***

'Free Play Ball Returns

Sub FreePlayLeft_hit()
  FreePlayLeft.DestroyBall
  BallsonBoard = BallsonBoard - 1
End sub

Sub FreePlayRight_hit()
  FreePlayRight.DestroyBall
  BallsonBoard = BallsonBoard - 1
End sub

'~~~~~~~~~~~~~~~~~~~~
'        Tilt (currently not working)
'~~~~~~~~~~~~~~~~~~~~
Dim TableTilted
Dim TiltCount
'Dim CheckTilt

'Sub TiltIt()
'   TiltCount = TiltCount + 1
'   CheckTilt
'End sub

'Sub CheckTilt
' If TiltCount > 4 then TiltKicker.Kick
'End sub

'Haven't figure out how to get the ball to kick out when you tilt it

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
  Vol = Csng(BallVel(ball) ^2 / 100)
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

Const tnob = 11 ' total number of balls
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


'Sub Gates_Hit (idx)
' PlaySound "Gate Hit", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'End Sub

Sub Bumper_Hit '(idx)
  PlaySoundAtVol "BumperHit", ActiveBall, 1
End Sub

Sub Pegs_Hit(idx)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "Peg3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "Peg4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub Guides_Hit(idx)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "Peg3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

'---------------------------
'----- High Score Code -----
'---------------------------
Const HighScoreFilename = "RockOlaBigGame.txt"

Dim HSArray,HSAHighScore, HSA1, HSA2, HSA3
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex 'Define 6 different score values for each reel to use
HSArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma","Tape")
Const DefaultHighScore = 0

Sub LoadHighScore
  Dim FileObj
  Dim ScoreFile
  Dim TextStr
    Dim SavedDataTemp3 'HighScore
    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & HighScoreFilename) then
    SetDefaultHSTD:UpdatePostIt:SaveHighScore
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & HighScoreFilename)
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      SetDefaultHSTD:UpdatePostIt:SaveHighScore
      Exit Sub
    End if
    SavedDataTemp3=Textstr.ReadLine ' HighScore
    TextStr.Close
    HSAHighScore=SavedDataTemp3
    UpdatePostIt
      Set ScoreFile = Nothing
      Set FileObj = Nothing
End Sub

Sub SetDefaultHSTD  'bad data or missing file - reset and resave
  HSAHighScore = DefaultHighScore
  SaveHighScore
End Sub

Sub UpdatePostIt
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
  End If
End Sub

Dim FileObj,ScoreFile
Sub SaveHighScore
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HighScoreFilename,True)
    ScoreFile.WriteLine HSAHighScore
    ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub

Sub CheckNewHighScorePostIt(newScore)
    If CLng(newScore) > CLng(HSAHighScore) Then
'     AddSpecial
      HSAHighScore=newScore
      SaveHighScore
      UpdatePostIt
    End If
End Sub

Sub Addpoints(ScorePar)
    if GameStarted Then
      Playerscore = Playerscore + ScorePar
    end If

'   if (Playerscore >= Special1Score) and not Special1 Then
'     Special1 = True
'     AddSpecial
'   end If
'   if (Playerscore >= Special2Score) and not Special2 Then
'     Special2 = True
'     AddSpecial
'   end If
'   if (Playerscore >= Special3Score) and not Special3 Then
'     Special3 = True
'     AddSpecial
'   end If
    UpdateScoreReel.enabled = True
End Sub

'xxx

Sub UpdateScoreReel_Timer
  TempScore = Playerscore
  Score10 = 0
  Score100 = 0
  Score1000 = 0
  Score10000 = 0
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score10 = cstr(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score100 = cstr(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score1000 = cstr(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score10000 = cstr(right(Tempscore,1))
  end If

  Val10 = ReelValue(Val10,Score10,1)
  Val100 = ReelValue(Val100,Score100,2)
  Val1000 = ReelValue(Val1000,Score1000,3)
  Val10000 = ReelValue(Val10000,Score10000,0)
  Tempscore = Val10 * 10 + Val100 * 100 + Val1000 * 1000 + Val10000 * 10000
  if Oldscore <> TempScore Then
'   playsound "solon",0,0.3,0.1,0.25
    Oldscore = TempScore
    P_Reel1.image = cstr(Val10)
    P_Reel2.image = cstr(Val100)
    P_Reel3.image = cstr(Val1000)
    P_Reel4.image = cstr(Val10000)
  Else
    UpdateScoreReel.enabled = False
  end If
End Sub

Function ReelValue(ValPar,ScorPar,ChimePar)
  ReelValue = cint(ValPar)
  if ReelValue <> cint(ScorPar) Then
    ReelValue = ReelValue + 1
    if ReelValue > 9 Then
      ReelValue = 0
    end If
  end If
End Function

Sub SetScoreReel
  TempScore = Playerscore
  Score10 = 0
  Score100 = 0
  Score1000 = 0
  Score10000 = 0
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score10 = cstr(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score100 = cstr(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score1000 = cstr(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score10000 = cstr(right(Tempscore,1))
  end If
  P_Reel1.image = cstr(Score10)
  P_Reel2.image = cstr(Score100)
  P_Reel3.image = cstr(Score1000)
  P_Reel4.image = cstr(Score10000)
End Sub


Sub LeafKicker_hit()
  LeafKicker.kick 0, 25 'need to make the angle more variable? Maybe the velocity as well
End sub

' ,\/~~~\_                            _/~~~~\
' |  ---, `\_    ___,-------~~\__  /~' ,,''  |
' | `~`, ',,\`-~~--_____    ---  - /, ,--/ '/'
'  `\_|\ _\`    ______,---~~~\  ,_   '\_/' /'
'    \,_|   , '~,/'~   /~\ ,_  `\_\ \_  \_\'
'    ,/   /' ,/' _,-'~~  `\  ~~\_ ,_  `\  `\
'  /@@ _/  /' ./',-                 \       `@,
'  @@ '   |  ___/  /'  /  \  \ '\__ _`~|, `, @@
'/@@ /  | | ',___  |  |    `  | ,,---,  |  | `@@,
'@@@ \  | | \ \*_`\ |        / / *_/' | \  \  @@@
'@@@ |  | `| '   ~ / ,          ~     /  |    @@@
'`@@ |   \ `\     ` |         | |  _/'  /'  | @@'
' @@ |    ~\ /--'~  |       , |  \__   |    | |@@
' @@, \     | ,,|   |       ,,|   | `\     /',@@
' `@@, ~\   \ '     |       / /    `' '   / ,@@
'  @@@,    \    ~~\ `\/~---'~/' _ /'~~~~~~~~--,_
'   `@@@_,---::::::=  `-,| ,~  _=:::::''''''    `
'   ,/~~_---'_,-___     _-__  ' -~~~\_```---
'     ~`   ~~_/'// _,--~\_/ '~--, |\_
'          /' /'| `@@@@@,,,,,@@@@  | \
'               `     `@@@@@@'

' Thalamus : Exit in a clean and proper way
Sub BigGameRO_exit
  Controller.Pause = False
  Controller.Stop
End Sub
