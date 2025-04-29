'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Treff - Fingerschlagautomat                                        ########
'#######          (Fa. Walter Steiner 1949)                                          ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
'
' Version 1.0
'
' Thanks to Tobias Ballmann for providing a detailed construction manual and details on the special parts
'
' There is no ROM needed to run the game. Start a game with the usual StartGameKey. Then You will get a
' set of playing coins to start Your game. Enter a coin by pressing either of the Flipper keys. Then use
' the plunger key to adjust Your hit force. After releasing of the plunger, the coin is thrown onto the
' playfield. Try to hit one of the pockets. The coin slides below te playfield and re-enters above the
' release mechanics, returning the indicated number of coins to the player.
' The number "5" pocket also awards the Reserve (situated at the rightmost lane).
' All coins, falling into the upper pocket next to the Reserve, are lost and remain in the cash box.
' The game is over once You have no coins left or if You can beat the machine by clearing the playfield of
' all coins. The coins in Your possession will be shown as the score.
'
' Version 1.1
'
' Added coin faces and rotation
'
' Version 1.2
'
' Added B2S backglass
'
' Version 1.3
'
' Added sounds



Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

'############################################
'######  Settings                      ######
'############################################

Const TotalPlayerCoins = 20   'number of coins per game
Const Volume = 0.25       'volume adjustment from 0-1


Const cGameName = "Treff"

Const BallSize = 29  'Ball radius for coin sized balls

Dim KickForce, i, Obj, CoinsOnTable, FirstRun
Randomize

Sub Treff_init
  LoadEM
  if B2SOn then
    Controller.B2SSetData 100,0 'CoinReady
    Controller.B2SSetData 34,0 'GameWon
    Controller.B2SSetData 32,0 'GameOn
    Controller.B2SSetData 35,1 'GameOver
  end if


  FirstRun = True
  InitTable


  if not Treff.ShowDT then
    ScoreText.visible = False
    ScoreText2.visible = False
  End If
End Sub

sub Treff_Keydown(ByVal keycode)
  if keycode = startgamekey Then
    If GameTimer.enabled = False Then
      if FirstRun = False Then
        InitTable
      Else
        FirstRun = False
      End if
      StartGame
    End if
  end If

  If keycode = LeftFlipperkey or keycode = RightFlipperkey Then
    NextCoin
  end If

  if keycode = plungerkey Then
    Plunger.Pullback
  end If
end Sub

sub Treff_Keyup(ByVal keycode)
  if keycode = plungerkey Then
    Playsound "Treff_Shoot", 1, 1*Volume, 0.7, 0.1
    Plunger.Fire
    HitRing.TransX = 42
    ResetRingTimer.enabled = True
  end If
end Sub


'############################################
'######  Game Code                     ######
'############################################

Dim CoinsLeft, PlungerOccupied
CoinsLeft = 0
PlungerOccupied = False
ScoreText.Text = "Game Over"
ScoreText2.Text = "Start New Game"

Sub StartGame
  If GameTimer.enabled = False Then
    CoinsLeft = TotalPlayerCoins
    GameTimer.enabled = True
    if B2SOn then
      Controller.B2SSetData 35,0 'GameOver
      Controller.B2SSetData 34,0 'GameWon
      Controller.B2SSetData 32,1 'GameOn

      Controller.B2SSetData 200,1 'GI
      Controller.B2SSetData 201,1 'GI
      Controller.B2SSetData 202,1 'GI
      Controller.B2SSetData 203,1 'GI

      B2SShowCoins(CoinsLeft)
    end if
  End If
End Sub

Dim x
Sub B2SShowCoins(ValuePar)
  'Clear old values
  Controller.B2SSetData 10,0
  Controller.B2SSetData 20,0
  Controller.B2SSetData 30,0
  Controller.B2SSetData 40,0
  Controller.B2SSetData 50,0
  Controller.B2SSetData 60,0
  Controller.B2SSetData 1,0
  Controller.B2SSetData 2,0
  Controller.B2SSetData 3,0
  Controller.B2SSetData 4,0
  Controller.B2SSetData 5,0
  Controller.B2SSetData 6,0
  Controller.B2SSetData 7,0
  Controller.B2SSetData 8,0
  Controller.B2SSetData 9,0
  Controller.B2SSetData 99,0  '0

  if ValuePar > 900 Then
    Exit Sub
  End If

  if ValuePar > 9 Then
    Controller.B2SSetData ValuePar-(ValuePar mod 10),1
  End If

  if ValuePar mod 10 = 0 Then
    Controller.B2SSetData 99,1
  Else
    Controller.B2SSetData (ValuePar mod 10),1
  End If
End Sub

Sub NextCoin
  if GameTimer.enabled Then
    if PlungerOccupied = False and CoinsLeft > 0 Then
      Playsound "Treff_InsertCoin", 1, 0.5*Volume, 0.4, 0.5
      CoinsLeft = CoinsLeft - 1
      CoinsOnTable = CoinsOnTable + 1
      PlungerOccupied = True
      BallInsert.createsizedball BallSize
      ballinsert.Kick 180, 5
      if B2SOn then
        Controller.B2SSetData 100,1 'CoinReady
      End if
    End if
  End If
End Sub

'Cash Return
Sub Drain_hit
  Playsound "Treff_Payout", 1, 1*Volume, 0.4, 0.5
  Drain.destroyball
  CoinsOnTable = CoinsOnTable - 1
  CoinsLeft = CoinsLeft + 1
End Sub

'Cash Return
Sub Drain2_hit
  Playsound "Treff_Payout", 1, 1*Volume, 0.4, 0.5
  Drain2.destroyball
  CoinsOnTable = CoinsOnTable - 1
  CoinsLeft = CoinsLeft + 1
End Sub


'Cash Box
Sub Drain3_hit
  Playsound "Treff_CoinLost", 1, 1*Volume, 0.4, 0.5
  Drain3.destroyball
  CoinsOnTable = CoinsOnTable - 1
End Sub


Sub GameTimer_Timer
  ScoreText.Text = "Coins         " & CoinsLeft
  ScoreText2.Text = ""
  if B2SOn then
    B2SShowCoins(CoinsLeft)
  End If
  if GameOverTimer.enabled and CoinsLeft > 0 Then
    GameOverTimer.enabled = False
  End if

  If CoinsOnTable <= 0 Then
    GameOverTimer.enabled = False
    GameTimer.enabled = False
    ScoreText.Text = "Game Won"
    ScoreText2.Text = "Score "& CoinsLeft
    if B2SOn then
      B2SShowCoins(CoinsLeft)
      Controller.B2SSetData 32,0 'GameOn
      Controller.B2SSetData 35,0 'GameOver
      Controller.B2SSetData 34,1 'GameWon

      Controller.B2SSetData 200,0 'GI
      Controller.B2SSetData 201,0 'GI
      Controller.B2SSetData 202,0 'GI
      Controller.B2SSetData 203,0 'GI
      Controller.B2SSetData 204,1 'GameWon_GI
    End If
  End If
End Sub

Sub GameOverTimer_Timer
  GameTimer.enabled = False
  ScoreText.Text = "Game Over"
  ScoreText2.Text = "Start New Game"

  if B2SOn then
    B2SShowCoins(999) 'clear coin display
    Controller.B2SSetData 32,0 'GameOn
    Controller.B2SSetData 34,0 'GameWon
    Controller.B2SSetData 35,1 'GameOver

    Controller.B2SSetData 200,0 'GI
    Controller.B2SSetData 201,0 'GI
    Controller.B2SSetData 202,0 'GI
    Controller.B2SSetData 203,0 'GI
    Controller.B2SSetData 204,0 'GameWon_GI
  End If
End Sub


Sub InitTable
  CoinTimer.enabled = False

  For each Obj in AllLaneHoles
    if Obj.Ballcntover <> 0 Then
      Obj.destroyball
    end if
  Next

  'Minimal Coin distribution

  R005.createsizedball BallSize
  R004.createsizedball BallSize
  Lane5_010.createsizedball BallSize
  Lane5_009.createsizedball BallSize
  Lane4_006.createsizedball BallSize
  Lane4_005.createsizedball BallSize
  Lane4_004.createsizedball BallSize
  Lane3_008.createsizedball BallSize
  Lane3_007.createsizedball BallSize
  Lane3_006.createsizedball BallSize
  Lane3_005.createsizedball BallSize
  Lane3_004.createsizedball BallSize
  Lane2_010.createsizedball BallSize
  Lane2_009.createsizedball BallSize
  Lane2_008.createsizedball BallSize
  Lane1_009.createsizedball BallSize
  Lane1_008.createsizedball BallSize

  ' -> 17
  i = round(rnd * 5)

  Select Case i
'Max    3=8=3=3=7=7
'     3-5-1-1-2-4 -> 16
'case 0: '3-5-1-1-2-4
'case 1: '2-6-2-2-3-1
'case 2: '2-4-3-1-3-3
'case 3: '3-2-1-3-6-1
'case 4: '2-3-2-2-2-5
'case 5: '1-6-2-0-5-2

    case 0: '3-5-1-1-2-4
      R003.createsizedball BallSize
      R002.createsizedball BallSize
      R001.createsizedball BallSize
      Lane5_008.createsizedball BallSize
      Lane5_007.createsizedball BallSize
      Lane5_006.createsizedball BallSize
      Lane5_005.createsizedball BallSize
      Lane5_004.createsizedball BallSize
      Lane4_003.createsizedball BallSize
      Lane3_003.createsizedball BallSize
      Lane2_007.createsizedball BallSize
      Lane2_006.createsizedball BallSize
      Lane1_007.createsizedball BallSize
      Lane1_006.createsizedball BallSize
      Lane1_005.createsizedball BallSize
      Lane1_004.createsizedball BallSize

    case 1: '2-6-2-2-3-1
      R003.createsizedball BallSize
      R002.createsizedball BallSize
      Lane5_008.createsizedball BallSize
      Lane5_007.createsizedball BallSize
      Lane5_006.createsizedball BallSize
      Lane5_005.createsizedball BallSize
      Lane5_004.createsizedball BallSize
      Lane5_003.createsizedball BallSize
      Lane4_003.createsizedball BallSize
      Lane4_002.createsizedball BallSize
      Lane3_003.createsizedball BallSize
      Lane3_002.createsizedball BallSize
      Lane2_007.createsizedball BallSize
      Lane2_006.createsizedball BallSize
      Lane2_005.createsizedball BallSize
      Lane1_007.createsizedball BallSize

    case 2: '2-4-3-1-3-3
      R003.createsizedball BallSize
      R002.createsizedball BallSize
      Lane5_008.createsizedball BallSize
      Lane5_007.createsizedball BallSize
      Lane5_006.createsizedball BallSize
      Lane5_005.createsizedball BallSize
      Lane4_003.createsizedball BallSize
      Lane4_002.createsizedball BallSize
      Lane4_001.createsizedball BallSize
      Lane3_003.createsizedball BallSize
      Lane2_007.createsizedball BallSize
      Lane2_006.createsizedball BallSize
      Lane2_005.createsizedball BallSize
      Lane1_007.createsizedball BallSize
      Lane1_006.createsizedball BallSize
      Lane1_005.createsizedball BallSize

    case 3: '3-2-1-3-6-1
      R003.createsizedball BallSize
      R002.createsizedball BallSize
      R001.createsizedball BallSize
      Lane5_008.createsizedball BallSize
      Lane5_007.createsizedball BallSize
      Lane4_003.createsizedball BallSize
      Lane3_003.createsizedball BallSize
      Lane3_002.createsizedball BallSize
      Lane3_001.createsizedball BallSize
      Lane2_007.createsizedball BallSize
      Lane2_006.createsizedball BallSize
      Lane2_005.createsizedball BallSize
      Lane2_004.createsizedball BallSize
      Lane2_003.createsizedball BallSize
      Lane2_002.createsizedball BallSize
      Lane1_007.createsizedball BallSize

    case 4: '2-3-2-2-2-5
      R003.createsizedball BallSize
      R002.createsizedball BallSize
      Lane5_008.createsizedball BallSize
      Lane5_007.createsizedball BallSize
      Lane5_006.createsizedball BallSize
      Lane4_003.createsizedball BallSize
      Lane4_002.createsizedball BallSize
      Lane3_003.createsizedball BallSize
      Lane3_002.createsizedball BallSize
      Lane2_007.createsizedball BallSize
      Lane2_006.createsizedball BallSize
      Lane1_007.createsizedball BallSize
      Lane1_006.createsizedball BallSize
      Lane1_005.createsizedball BallSize
      Lane1_004.createsizedball BallSize
      Lane1_003.createsizedball BallSize

    case 5: '1-6-2-0-5-2
      R003.createsizedball BallSize
      Lane5_008.createsizedball BallSize
      Lane5_007.createsizedball BallSize
      Lane5_006.createsizedball BallSize
      Lane5_005.createsizedball BallSize
      Lane5_004.createsizedball BallSize
      Lane5_003.createsizedball BallSize
      Lane4_003.createsizedball BallSize
      Lane4_002.createsizedball BallSize
      Lane2_007.createsizedball BallSize
      Lane2_006.createsizedball BallSize
      Lane2_005.createsizedball BallSize
      Lane2_004.createsizedball BallSize
      Lane2_003.createsizedball BallSize
      Lane1_007.createsizedball BallSize
      Lane1_006.createsizedball BallSize
  End Select

  CoinsOnTable = 33
  CoinsLeft = 0
  CoinTimer.enabled = True
End Sub


'############################################
'######  Mechanics                     ######
'############################################

Sub ResetRingTimer_Timer
  HitRing.TransX = HitRing.TransX - 1
  if HitRing.TransX <= 0 Then
    ResetRingTimer.enabled = False
    HitRing.TransX = 0
  End If
End Sub

Sub ExitPlunger_Hit
  PlungerOccupied = False
  if B2SOn then
    Controller.B2SSetData 100,0 'CoinReady
  End If
  if CoinsLeft = 0 Then
    GameOverTimer.enabled = True
  End If
End Sub

Sub Pocket2_hit
  Playsound "Treff_Pocket", 1, 1*Volume, 0.2, 0.5
  Pocket2.destroyball
  Exit2.createsizedball BallSize
  Exit2.Kick 180, 5
End Sub

Sub Pocket3_hit
  Playsound "Treff_Pocket", 1, 1*Volume, 0, 0.5
  Pocket3.destroyball
  Exit3.createsizedball BallSize
  Exit3.Kick 180, 5
End Sub

Sub Pocket4_hit
  Playsound "Treff_Pocket", 1, 1*Volume, -0.2, 0.5
  Pocket4.destroyball
  Exit4.createsizedball BallSize
  Exit4.Kick 180, 5
End Sub

Sub Pocket5_hit
  Playsound "Treff_Pocket", 1, 1*Volume, -0.4, 0.5
  Pocket5.destroyball
  Exit5.createsizedball BallSize
  Exit5.Kick 180, 5
End Sub


Dim CoinModel, BOT, b, CoordOverride, OverrideRotZ
CoinModel = Array (Disc001,Disc002,Disc003,Disc004,Disc005,Disc006,Disc007,Disc008,Disc009,Disc010,Disc011,Disc012,Disc013,Disc014,Disc015,Disc016,Disc017,Disc018,Disc019,Disc020,Disc021,Disc022,Disc023,Disc024,Disc025,Disc026,Disc027,Disc028,Disc029,Disc030,Disc031,Disc032,Disc033,Disc034,Disc035,Disc036,Disc037,Disc038,Disc039,Disc040,Disc041,Disc042,Disc043,Disc044,Disc045,Disc046,Disc047,Disc048,Disc049,Disc050,Disc051,Disc052,Disc053,Disc054,Disc055,Disc056,Disc057,Disc058,Disc059,Disc060)
Const tnob = 60


Sub CoinTimer_Timer

    BOT = GetBalls

    ' hide coins for deleted balls
    If UBound(BOT)<(tnob-1) Then
        For b = (UBound(BOT) + 1) to (tnob-1)
            CoinModel(b).visible = 0
        Next
    End If

    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

    ' render the coin for each ball
  For b = 0 to UBound(BOT)
    BOT(b).visible = False

    ' Initialize coin face and rotation
    if BOT(b).PlayfieldReflectionScale = 1 Then
      if rnd < 0.5 Then
        BOT(b).image = "Coin_front"
      Else
        BOT(b).image = "Coin_back"
      end if

      'add 1000 to rotation to recognize an initialized rotation of 1 degree (init value)
      BOT(b).PlayfieldReflectionScale = 1000 + round(rnd * 360)

      'store ball position for movement check in some non-vital variables
      BOT(b).UserValue = BOT(b).X
      BOT(b).BulbIntensityScale = BOT(b).Y
    End if

    if (ABS(BOT(b).UserValue - BOT(b).X) > 0.1) Then
      BOT(b).PlayfieldReflectionScale = BOT(b).PlayfieldReflectionScale - ((BOT(b).UserValue - BOT(b).X) * 0.5) + (rnd * 2)
      If BOT(b).PlayfieldReflectionScale < 1000 or BOT(b).PlayfieldReflectionScale > 1360 Then
        BOT(b).PlayfieldReflectionScale = 1000
      End if

      'save new position
      BOT(b).UserValue = BOT(b).X
      BOT(b).BulbIntensityScale = BOT(b).Y
    End if

    'update coin
    CoinModel(b).X = BOT(b).X
    CoinModel(b).Y = BOT(b).Y
    CoinModel(b).Image = BOT(b).Image
    CoinModel(b).ObjRotZ = BOT(b).PlayfieldReflectionScale - 1000
    CoinModel(b).visible = 1
  Next
End Sub

Dim RPos,i_r
RPos = Array (R005,R004,R003,R002,R001)
Sub LaneRTimer_Timer
  If Release_R Then
    if R005.Ballcntover Then
      Playsound "Treff_CoinRelease", 1, 1*Volume, 0.5, 0.5
      ReleaseRTimer.enabled = False
      R005.Kick 180, 2
    End if
    ReleaseRTimer.enabled = True
  End if

  for i_r = 0 to 3
    if RPos(i_r).Ballcntover = 0 Then
      if RPos(i_r+1).Ballcntover <> 0 Then
        RPos(i_r+1).Kick 180, 2
      End If
    End If
  Next
  R_Wall.isdropped = (RPos(4).Ballcntover = 0)
End Sub

Sub ReleaseRTimer_Timer
  if R001.Ballcntover = 0 and R002.Ballcntover = 0 and R003.Ballcntover = 0 and R004.Ballcntover = 0 and R005.Ballcntover = 0 Then
    ReleaseRTimer.enabled = False
    Release_R = False
    TriggerR.timerenabled = True
  End if
End Sub


Dim L1Pos,i_1
L1Pos = Array (Lane1_009,Lane1_008,Lane1_007,Lane1_006,Lane1_005,Lane1_004,Lane1_003,Lane1_002,Lane1_001)
Sub Lane1Timer_Timer
  for i_1 = 0 to 7
    if L1Pos(i_1).Ballcntover = 0 Then
      if L1Pos(i_1+1).Ballcntover <> 0 Then
        L1Pos(i_1+1).Kick 180, 2
      End If
    End If
  Next
  Wall_L1.isdropped = (L1Pos(8).Ballcntover = 0)
End Sub

Dim L2Pos,i_2
L2Pos = Array (Lane2_010,Lane2_009,Lane2_008,Lane2_007,Lane2_006,Lane2_005,Lane2_004,Lane2_003,Lane2_002,Lane2_001)
Sub Lane2Timer_Timer
  for i_2 = 0 to 8
    if L2Pos(i_2).Ballcntover = 0 Then
      if L2Pos(i_2+1).Ballcntover <> 0 Then
        L2Pos(i_2+1).Kick 180, 2
      End If
    End If
  Next
  Wall_L2.isdropped = (L2Pos(9).Ballcntover = 0)
End Sub

Dim L3Pos,i_3
L3Pos = Array (Lane3_008,Lane3_007,Lane3_006,Lane3_005,Lane3_004,Lane3_003,Lane3_002,Lane3_001)
Sub Lane3Timer_Timer
  for i_3 = 0 to 6
    if L3Pos(i_3).Ballcntover = 0 Then
      if L3Pos(i_3+1).Ballcntover <> 0 Then
        L3Pos(i_3+1).Kick 180, 2
      End If
    End If
  Next
  Wall_L3.isdropped = (L3Pos(7).Ballcntover = 0)
End Sub

Dim L4Pos,i_4
L4Pos = Array (Lane4_006,Lane4_005,Lane4_004,Lane4_003,Lane4_002,Lane4_001)
Sub Lane4Timer_Timer
  for i_4 = 0 to 4
    if L4Pos(i_4).Ballcntover = 0 Then
      if L4Pos(i_4+1).Ballcntover <> 0 Then
        L4Pos(i_4+1).Kick 180, 2
      End If
    End If
  Next
  Wall_L4.isdropped = (L4Pos(5).Ballcntover = 0)
End Sub

Dim L5Pos,i_5
L5Pos = Array (Lane5_010,Lane5_009,Lane5_008,Lane5_007,Lane5_006,Lane5_005,Lane5_004,Lane5_003,Lane5_002,Lane5_001)
Sub Lane5Timer_Timer
  for i_5 = 0 to 8
    if L5Pos(i_5).Ballcntover = 0 Then
      if L5Pos(i_5+1).Ballcntover <> 0 Then
        L5Pos(i_5+1).Kick 180, 2
      End If
    End If
  Next
  Wall_L5.isdropped = (L5Pos(9).Ballcntover = 0)
End Sub

Sub Trigger2R_Hit
  Playsound "Treff_Hit", 1, 1*Volume, 0.2, 0.5
  L1_Klinke.ObjRotZ = -45
  Trigger2R.timerenabled = True
  If Lane1_009.Ballcntover Then
    Playsound "Treff_CoinRelease", 1, 1*Volume, 0.3, 0.5
    Lane1_009.kick 180,2
  End if
End Sub

Sub Trigger2R_Timer
  L1_Klinke.ObjRotZ = L1_Klinke.ObjRotZ + 5
  If L1_Klinke.ObjRotZ >= 0 Then
    Trigger2R.timerenabled = False
    L1_Klinke.ObjRotZ = 0
  End If
End Sub

Sub Trigger2L_Hit
  Playsound "Treff_Hit", 1, 1*Volume, 0, 0.5
  L2_Klinke.ObjRotZ = -45
  Trigger2L.timerenabled = True
  If Lane2_010.Ballcntover Then
    Playsound "Treff_CoinRelease", 1, 1*Volume, 0.1, 0.5
    Lane2_010.kick 180,2
  End if
End Sub

Sub Trigger2L_Timer
  L2_Klinke.ObjRotZ = L2_Klinke.ObjRotZ + 5
  If L2_Klinke.ObjRotZ >= 0 Then
    Trigger2L.timerenabled = False
    L2_Klinke.ObjRotZ = 0
  End If
End Sub

Sub Trigger3_Hit
  Playsound "Treff_Hit", 1, 1*Volume, -0.2, 0.5
  L3_Klinke.ObjRotZ = -45
  Trigger3.timerenabled = True
  If Lane3_008.Ballcntover Then
    Playsound "Treff_CoinRelease", 1, 1*Volume, -0.1, 0.5
    Lane3_008.kick 180,2
  End if
End Sub

Sub Trigger3_Timer
  L3_Klinke.ObjRotZ = L3_Klinke.ObjRotZ + 5
  If L3_Klinke.ObjRotZ >= 0 Then
    Trigger3.timerenabled = False
    L3_Klinke.ObjRotZ = 0
  End If
End Sub

Sub Trigger5A_Hit
  Playsound "Treff_Hit", 1, 1*Volume, -0.4, 0.5
  L4_Klinke.ObjRotZ = -45
  Trigger5A.timerenabled = True
  If Lane4_006.Ballcntover Then
    Playsound "Treff_CoinRelease", 1, 1*Volume, -0.3, 0.5
    Lane4_006.kick 180,2
  End if
End Sub

Sub Trigger5A_Timer
  L4_Klinke.ObjRotZ = L4_Klinke.ObjRotZ + 5
  If L4_Klinke.ObjRotZ >= 0 Then
    Trigger5A.timerenabled = False
    L4_Klinke.ObjRotZ = 0
  End If
End Sub

Sub Trigger5B_Hit
  Playsound "Treff_Hit", 1, 1*Volume, -0.4, 0.5
  L5_Klinke.ObjRotZ = 45
  Trigger5B.timerenabled = True
  If Lane5_010.Ballcntover Then
    Playsound "Treff_CoinRelease", 1, 1*Volume, -0.5, 0.5
    Lane5_010.kick 180,2
  End if
End Sub

Sub Trigger5B_Timer
  L5_Klinke.ObjRotZ = L5_Klinke.ObjRotZ - 5
  If L5_Klinke.ObjRotZ <= 0 Then
    Trigger5B.timerenabled = False
    L5_Klinke.ObjRotZ = 0
  End If
End Sub

Dim Release_R
Release_R = False
Sub TriggerR_Hit
  R_Klinke1.ObjRotZ = 21
  R_Klinke2.ObjRotZ = -28
  Release_R = True
End Sub

Sub TriggerR_Timer
  R_Klinke1.ObjRotZ = R_Klinke1.ObjRotZ - 2
  R_Klinke2.ObjRotZ = R_Klinke2.ObjRotZ + 2.7

  If R_Klinke1.ObjRotZ <= 0 Then
    TriggerR.timerenabled = False
    R_Klinke1.ObjRotZ = 0
    R_Klinke2.ObjRotZ = 0
  End If
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

