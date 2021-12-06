Option Explicit
Randomize
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

' Thalamus 2018-11-01 : Improved directional sounds
' Thalamus : tables seems to be lacking tilt routine
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


Dim Attract
Dim AwardLR,AwardMR,AwardHR
Dim AwardLY,AwardMY,AwardHY
Dim AwardLG,AwardMG,AwardHG
Dim BIP
Dim BIL
Dim Bypass
Dim BWait
Dim credit,oldcredit,C
Dim Ballz
Dim ballinplay
Dim Coin
Dim LastB
Dim InProgress
Dim MX,MS,TT,xx,xxx,kk,ll,ff,ss,bb,fa,ex,sms,rl
Dim MaxCredit
Dim Odds
Dim Objekt
Dim Rollers
Dim ScreenLock
Dim RNDM,RNDMR,RNDMY,RNDMG,RNDMOK
Dim Red,Yellow,Green
Dim RArray,YArray,GArray,FArray,a,b,x
Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys
Const BallSize = 50  'Ball radius

'******************************************************************************************
'***************  Game Settings ***********************************************************
' Coin Choice:    1: Nickel (1 Credit)
'         2: Quarter (5 Credits
'         3: Half Dollar (10 Credits)
'         4: Dollar (20 Credits)

Coin = 4
'
'   Max Credits:    Set Maximium allowable buy-in credits to start (100 or less recommended, 999 Maximum)
'
MaxCredit = 100
'
'
'******************************************************************************************
'******************************************************************************************

RArray = Array(R1,R2,R3,R4,R5,R6,R7,R8)
YArray = Array(Y1,Y2,Y3,Y4,Y5,Y6,Y7,Y8)
GArray = Array(G1,G2,G3,G4,G5,G6,G7,G8)
FArray = Array(F1,F2,F3,F4,F5)

If Table1.ShowDT = True then
  For each objekt in backdropstuff : objekt.visible = 1 : next
    Else
  For each objekt in backdropstuff : objekt.visible = 0 : next
End If

sub Table1_init
  LoadEM
  For each xx in GI:xx.State = 1: Next
  MX=2
  MS=2
  TT=1
  BIP = 0
  Ballz = 6
  RNDMR=1
  RNDMY=1
  RNDMG=1
  RNDMOK=1
  xxx = 1
  LastB=0
  Red = 0
  Yellow =0
  Green = 0
  Attract = 1
  InProgress = 0
  ScreenLock = 0
  gamov.text="GAME OVER"
  ballz2play.text = ""
  credittxt.setvalue(credit)
  MagicScreen.setvalue(MX)
  oldcredit=0
  If B2SOn then
    With Controller
    .B2SSetCredits 1,MX
    .B2SSetCredits 9,9
    .B2SSetCredits 2,1
    .B2SSetCredits 3,2
    .B2SSetCredits 11,11
    .B2SSetCredits 15,15
    .B2SSetCredits 6,4
    .B2SSetCredits 7,19
    .B2SSetCredits 8,7
    .B2SSetCredits 4,22
    .B2SSetCredits 10,18
    .B2SSetCredits 26,25
    .B2SSetCredits 12,24
    .B2SSetCredits 13,16
    .B2SSetCredits 14,13
    .B2SSetCredits 5,17
    .B2SSetCredits 16,6
    .B2SSetCredits 17,23
    .B2SSetCredits 18,5
    .B2SSetCredits 19,21
    .B2SSetCredits 20,20
    .B2SSetCredits 21,12
    .B2SSetCredits 22,8
    .B2SSetCredits 23,14
    .B2SSetCredits 24,3
    .B2SSetCredits 25,10
    End With
  End If
  Playsound "IdleMachine",-1
End Sub
'-----------------------------------------------------------------------------------------
Sub coindelay_timer()
  addcredit
    coindelay.enabled=false
End Sub

Sub addcredit()
  if credit > Maxcredit then Exit Sub
  Select Case Coin
    Case 1: credit = credit + 1
    Case 2: credit = credit + 5
    Case 3: credit = credit + 10
    Case 4: credit = credit + 20
  End Select
  If credit > MaxCredit then credit = MaxCredit
  UpdateCredits
End sub

Sub UpdateCredits()
  C = Credit - OldCredit
  credittxt.AddValue(C)
    If B2SOn then
        Controller.B2SSetCredits 29,Credit \ 10
        Controller.B2SSetCredits 30,Credit Mod 10
    Controller.B2SSetCredits 31,Credit \ 100
  End If
  If Credit > 998 then Credit = 999
  If Credit = 999 then playsound "DoubleBells"
  oldcredit=credit
End Sub
'-----------------------------------------------------------------------------------------
Sub Table1_KeyDown(ByVal keycode)
  if keycode = 19 Then
    if InProgress=0 then exit Sub
    RollerCard.visible=True
    RollerCard.image = "RollerCard"&rollers
  end if

  if keycode=AddCreditKey then
    playsoundAtVol "coininsert", drain, 1
    coindelay.enabled=true
    end if

  If keycode = PlungerKey Then
    Plunger.PullBack
    PlaySound "plungerpull",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  If keycode = 2 Then                   '*****Start Game #1
  If Attract = 0 AND InProgress = 0 Then
    NewGame
    AttractOff
    InProgress = 1
    OldCredit = credit
   End If
    End If

  If keycode = 3 Then                   '*****Blue Button (Scores) #2
  If Credit > 0 AND InProgress = 0 Then
    AttractOff
    Credit = Credit - 1
    For Each xx in Anticipation_Scores:xx.State = 2: Next
    Anticipation_Scores_Timer.enabled = 1
    If B2Son then Controller.B2SStartAnimation "scrs":End If
    Scores
  If Credit < 1 then Credit = 0
    RandomSoundCycle
    UpdateCredits
   End If
    End If

  If keycode = 4 Then                                   '*****Green Button (Features) #3
    If Credit > 0 AND InProgress = 0 Then
    AttractOff
    Credit = Credit - 1
    For Each xx in Anticipation:xx.State = 2: Next
    Anticipation_Timer.enabled = 1
    If B2Son then Controller.B2SStartAnimation "ftrs":End If
    Features
  If Credit < 1 then Credit = 0
    RandomSoundCycle
    UpdateCredits
   End If
  End If

  If keycode = 5 Then                 '*****Red Button (Both) #4
    If Credit > 0 AND InProgress = 0 Then
    AttractOff
    Credit = Credit - 1
    For Each xx in Anticipation:xx.State = 2: Next
    For Each xx in Anticipation_Scores:xx.State = 2: Next
    Anticipation_Timer.enabled = 1
    Anticipation_Scores_Timer.enabled = 1
    If B2Son then
      Controller.B2SStartAnimation "scrs"
      Controller.B2SStartAnimation "ftrs"
    End If
    Select Case Int(Rnd*2)+1
      Case 1: Scores
      Case 2: Features
    End Select
    If Credit < 1 then Credit = 0
    RandomSoundCycle
    UpdateCredits
   End If
  End If
'-----------------------------------------------------------------------------------------
'******************Right/Left Arrows or Flipper Buttons = Scroll Magic Screen *******************************
' If keycode = 205 Then    'Right Arrow Key
  If keycode = RightFlipperKey Then
  If ScreenLock=1 then exit Sub
    If InProgress=0 then exit Sub
    IF LK.state = 1 then MS = 1 : End If
    IF LO.state = 1 then MS = 2 : End If

    IF LA.state = 1 then MS = 3 : End If
    IF LB.state = 1 then MS = 4 : End If
    IF LC.state = 1 then MS = 5 : End If
    IF LD.state = 1 then MS = 6 : End If
    IF LE.state = 1 then MS = 7 : End If
    IF LF.state = 1 then MS = 8 : End If
    IF LG.state = 1 then MS = 9 : End If

    MX = MX + 1
    IF MX > MS then MX = MS:End If
      MagicScreen.setvalue(MX)
    IF MX = MS then playsound "buzz":End If
    IF MX <> MS then playsound "rotate":End If
    If B2SOn then Controller.B2SSetCredits 1,MX
  End If

' If keycode = 203 Then  'LeftArrow Key
  If keycode = LeftFlipperKey Then
  If Screenlock=1 then exit Sub
    If InProgress=0 then exit Sub
    IF MX = 0 then playsound "buzz": Exit Sub: End If
    IF LK.state=0 and MX=2 then playsound "buzz": Exit Sub: End If
    IF LO.state=0 and MX=1 then playsound "buzz": Exit Sub: End If
    MX = MX - 1
    IF MS = MX then MX = MS:End If
    MagicScreen.setvalue(MX)
    IF MX = MS then playsound "buzz":End If
    IF MX <> MS then playsound "rotate":End If
    If B2SOn then Controller.B2SSetCredits 1,MX
  End If
'-----------------------------------------------------------------------------------------
'***************Nudge Keys **************************************************************
  If keycode = LeftTiltKey Then
    Nudge 90, 2
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
  End If

 If keycode = MechanicalTilt
  ' PlaySound "fx_nudge",0,1,1,0,25
  ' CheckTilt
 End If
'-----------------------------------------------------------------------------------------
 '************************************* Manual Ball Control*******************************
  If keycode = 46 Then          ' C Key
    If EnableBallControl = 1 Then
      EnableBallControl = 0
    Else
      EnableBallControl = 1
    End If
  End If
    If EnableBallControl = 1 Then
    If keycode = 48 Then        ' B Key
      If BCboost = 1 Then
        BCboost = BCboostmulti
      Else
        BCboost = 1
      End If
    End If
    If keycode = 203 Then BCleft = 1  ' Left Arrow
    If keycode = 200 Then BCup = 1    ' Up Arrow
    If keycode = 208 Then BCdown = 1  ' Down Arrow
    If keycode = 205 Then BCright = 1 ' Right Arrow
  End If
End Sub
'-----------------------------------------------------------------------------------------
Sub Table1_KeyUp(ByVal keycode)
  if keycode = 19 Then                   'R Key to show current "ROLLER" Card
    if InProgress=0 then exit Sub
    RollerCard.visible=false
  end If

  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySound "plunger",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  If keycode = LeftFlipperKey Then

  End If

  If keycode = RightFlipperKey Then

  End If
'************************************* Manual Ball Control*******************************
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If
End Sub
'-----------------------------------------------------------------------------------------

Sub Drain_Hit()
  Drain.DestroyBall
  PlaySound "drain",0,1,AudioPan(Drain),0.25,0,0,1,AudioFade(Drain)
  If BIL = 1 then BWait = 1:Exit Sub
  Bypass=1
  BallRelease.CreateBall
  BallRelease.Kick 90, 7
  PlaySound SoundFX("Ball_Lifter",DOFContactors), 0,1,AudioPan(BallRelease),0.25,0,0,1,AudioFade(BallRelease)
End Sub

Sub balltrigger_unHit()
  If Bypass=1 then Bypass =0:Exit Sub
  If BWait = 1 Then
    BallRelease.CreateBall
    BallRelease.Kick 90, 7
    PlaySound SoundFX("Ball_Lifter",DOFContactors), 0,1,AudioPan(BallRelease),0.25,0,0,1,AudioFade(BallRelease)
    BWait = 0
      Else
    If LastB=1 then Exit Sub
    If InProgress=1 Then NextBall
  End If
End Sub

Sub NextBall()
  F21.State=0
  BIP = BIP + 1
  If BIP = Ballz Then LastB = 1: Exit Sub :End If
    BallCountTimer.enabled=1
  ScreenLock_Timer.enabled=1
  BallRelease.CreateBall
  BallRelease.Kick 90, 7
  PlaySound SoundFX("Ball_Lifter",DOFContactors), 0,1,AudioPan(BallRelease),0.25,0,0,1,AudioFade(BallRelease)
End Sub

Sub BallCountTimer_Timer
  Select Case xxx
    Case 1: ballinplay1.text = 1:BallinPlay2.text = ""
    Case 2: ballinplay2.text = 2:BallinPlay1.text = ""
    Case 3: BallinPlay3.text = 3:BallinPlay2.text = ""
    Case 4: BallinPlay4.text = 4:BallinPlay3.text = "":BallinPlay5.text = ""
    Case 5: BallinPlay5.text = 5:BallinPlay4.text = "":BallinPlay6.text = ""
    Case 6: BallinPlay6.text = 6:BallinPlay5.text = "":BallinPlay7.text = ""
    Case 7: BallinPlay7.text = 7:BallinPlay6.text = "":Ballinplay8.text = ""
    Case 8: Ballinplay8.text = 8:BallinPlay7.text = "":BallinPlay6.text = ""
  End Select
  xxx = xxx + 1
  BallCountTimer.enabled=0
End Sub

Sub Trigger1_Hit():   BIL = 1:End Sub
Sub Trigger1_unHit(): BIL = 0:End Sub

Sub LastBall()
  EndGame
End Sub

Sub ScreenLock_Timer_Timer
  If BIP = 3 And F6.state=0 Then F21.State=2:End If
  If BIP = 4 And F6.state=0 Then ScreenLock=1:End If
  If BIP = 4 And F6.state=1 And F9.state=0 And F10.State=0 Then F21.State=2:End If
  If BIP = 5 And F9.state=0 Then ScreenLock=1:End If
  If BIP = 5 and F9.state=1 Then F21.State=2:End If
  If BIP = 6 And F10.state=0 Then Screenlock=1:End If
  If BIP = 6 And F10.state=1 Then F21.State=2:End If
  If BIP > 6 Then ScreenLock=1:End If
  ScreenLock_Timer.enabled=0
End Sub
'-----------------------------------------------------------------------------------------

Sub NewGame()
  gamov.text=""
  balls2play.text="BALL IN PLAY"
  ballz2playText.text = "BALLS PER GAME"
  ballz2play.text = (Ballz-1)
  For each kk in Kickers:kk.DestroyBall:Next
  Playsound "Balldrop"
  For each ll in Num_Lights:ll.State = 0: Next
  ScreenLock=0
  LastB=0
  BIP = 0
  SS = 0
  BB = 0
  ex=1
  sms = 1
  xxx = 1
  MX=2
  MS=2
  MagicScreen.setvalue(MX)
  Select Case Int(Rnd*6)+1
    Case 1: RedR.state=1  :Rollers=1
    Case 2: RedO.state=1  :Rollers=2
    Case 3: RedL.state=1  :Rollers=3
    Case 4: RedL1.state=1   :Rollers=4
    Case 5: RedE.state=1  :Rollers=5
    Case 6: RedR1.state=1 :Rollers=6
  End Select
  RollerCard.image = "RollerCard"&rollers
  If B2SOn then
    Controller.B2SSetCredits 1,MX
    for x = 101 to 125
    Controller.B2SSetData (x),0
    Next
  End If
  Playsound "StartGame"
  AwardFeatures
  NextBall
End Sub

Sub EndGame()
  GameOverTimer.enabled=1
  InProgress=0
  gamov.text="GAME OVER"
  For Each xx in BallsinPlayTxt:xx.text = "": Next
  ballz2play.text = ""
End Sub

Sub GameOverTimer_Timer
' For Each xx in Score_Lights:xx.State = 2: Next
' For Each ff in Feature_lights:ff.State = 2: Next
  Award
  BIP = 0
  Ballz = 6
  xxx = 1
  Attract = 1
  RNDMR=1
  RNDMY=1
  RNDMG=1
  RNDMOK=1
  TT=1
  F21.state=0
  UpdateCredits
  GameOverTimer.enabled=0
End Sub

Sub AttractOff()
  If Attract = 0 then exit sub
  For Each xx in Score_Lights:xx.State = 0: Next
  For Each ff in Feature_lights:ff.State = 0: Next
  For Each xx in RollerLights:xx.State = 0: Next
  YellowLight.state=0
  RedLight.state=0
  Attract = 0
End Sub

Sub Anticipation_Timer_Timer
  Anticipation_Timer.enabled=0
  Stoplights.enabled=1
End Sub

Sub Anticipation_Scores_Timer_Timer
  Anticipation_Scores_Timer.enabled=0
  Stoplights.enabled=1
End Sub

Sub StopLights_Timer
  For Each xx in Anticipation:xx.State = 0: Next
  For Each xx in Anticipation_Scores:xx.State = 0: Next
  Stoplights.enabled=0
End Sub
'-----------------------------------------------------------------------------------------

Sub BackglassTimer_Timer
If B2Son Then
  If R1.state=1 then controller.B2SSetData 1,1 Else controller.B2SSetData 1,0:End If
  If R2.state=1 then controller.B2SSetData 18,1 Else controller.B2SSetData 18,0:End If
  If R3.state=1 then controller.B2SSetData 19,1 Else controller.B2SSetData 19,0:End If
  If R4.state=1 then controller.B2SSetData 20,1 Else controller.B2SSetData 20,0:End If
  If R5.state=1 then controller.B2SSetData 21,1 Else controller.B2SSetData 21,0:End If
  If R6.state=1 then controller.B2SSetData 22,1 Else controller.B2SSetData 22,0:End If
  If R7.state=1 then controller.B2SSetData 23,1 Else controller.B2SSetData 23,0:End If
  If R8.state=1 then controller.B2SSetData 24,1 Else controller.B2SSetData 24,0:End If
  If Y1.state=1 then controller.B2SSetData 10,1 Else controller.B2SSetData 10,0:End If
  If Y2.state=1 then controller.B2SSetData 11,1 Else controller.B2SSetData 11,0:End If
  If Y3.state=1 then controller.B2SSetData 12,1 Else controller.B2SSetData 12,0:End If
  If Y4.state=1 then controller.B2SSetData 13,1 Else controller.B2SSetData 13,0:End If
  If Y5.state=1 then controller.B2SSetData 14,1 Else controller.B2SSetData 14,0:End If
  If Y6.state=1 then controller.B2SSetData 15,1 Else controller.B2SSetData 15,0:End If
  If Y7.state=1 then controller.B2SSetData 16,1 Else controller.B2SSetData 16,0:End If
  If Y8.state=1 then controller.B2SSetData 17,1 Else controller.B2SSetData 17,0:End If
  If G1.state=1 then controller.B2SSetData 25,1 Else controller.B2SSetData 25,0:End If
  If G2.state=1 then controller.B2SSetData 26,1 Else controller.B2SSetData 26,0:End If
  If G3.state=1 then controller.B2SSetData 27,1 Else controller.B2SSetData 27,0:End If
  If G4.state=1 then controller.B2SSetData 28,1 Else controller.B2SSetData 28,0:End If
  If G5.state=1 then controller.B2SSetData 29,1 Else controller.B2SSetData 29,0:End If
  If G6.state=1 then controller.B2SSetData 30,1 Else controller.B2SSetData 30,0:End If
  If G7.state=1 then controller.B2SSetData 31,1 Else controller.B2SSetData 31,0:End If
  If G8.state=1 then controller.B2SSetData 32,1 Else controller.B2SSetData 32,0:End If
  If F1.state=1 then controller.B2SSetData 38,1 Else controller.B2SSetData 38,0:End If
  If F2.state=1 then controller.B2SSetData 39,1 Else controller.B2SSetData 39,0:End If
  If F3.state=1 then controller.B2SSetData 40,1 Else controller.B2SSetData 40,0:End If
  If F4.state=1 then controller.B2SSetData 41,1 Else controller.B2SSetData 41,0:End If
  If F5.state=1 then controller.B2SSetData 42,1 Else controller.B2SSetData 42,0:End If
  If F6.state=1 then controller.B2SSetData 43,1 Else controller.B2SSetData 43,0:End If
  If F7.state=1 then controller.B2SSetData 44,1 Else controller.B2SSetData 44,0:End If
  If F8.state=1 then controller.B2SSetData 45,1 Else controller.B2SSetData 45,0:End If
  If F9.state=1 then controller.B2SSetData 46,1 Else controller.B2SSetData 46,0:End If
  If F10.state=1 then controller.B2SSetData 47,1 Else controller.B2SSetData 47,0:End If
  If F21.state=2 then controller.B2SSetData 49,1 Else controller.B2SSetData 49,0:End If
  If EB.state=1 then controller.B2SSetData 50,1 Else controller.B2SSetData 50,0:End If
  If EB1.state=1 then controller.B2SSetData 51,1 Else controller.B2SSetData 51,0:End If
  If EB2.state=1 then controller.B2SSetData 52,1 Else controller.B2SSetData 52,0:End If
  If LO.state=1 then controller.B2SSetData 53,1 Else controller.B2SSetData 53,0:End If
  If LK.state=1 then controller.B2SSetData 54,1 Else controller.B2SSetData 54,0:End If
  If LA.state=1 then controller.B2SSetData 55,1 Else controller.B2SSetData 55,0:End If
  If LB.state=1 then controller.B2SSetData 56,1 Else controller.B2SSetData 56,0:End If
  If LC.state=1 then controller.B2SSetData 57,1 Else controller.B2SSetData 57,0:End If
  If LD.state=1 then controller.B2SSetData 58,1 Else controller.B2SSetData 58,0:End If
  If LE.state=1 then controller.B2SSetData 59,1 Else controller.B2SSetData 59,0:End If
  If LF.state=1 then controller.B2SSetData 60,1 Else controller.B2SSetData 60,0:End If
  If LG.state=1 then controller.B2SSetData 61,1 Else controller.B2SSetData 61,0:End If
  If RedR.state=1 then controller.B2SSetData 62,1 Else controller.B2SSetData 62,0:End If
  If RedO.state=1 then controller.B2SSetData 63,1 Else controller.B2SSetData 63,0:End If
  If RedL.state=1 then controller.B2SSetData 64,1 Else controller.B2SSetData 64,0:End If
  If RedL1.state=1 then controller.B2SSetData 65,1 Else controller.B2SSetData 65,0:End If
  If RedE.state=1 then controller.B2SSetData 66,1 Else controller.B2SSetData 66,0:End If
  If RedR1.state=1 then controller.B2SSetData 67,1 Else controller.B2SSetData 67,0:End If
End If
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
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

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function


'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1       'Do Not Change - default setting
BCvel = 4       'Controls the speed of the ball movement
BCyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
  Set ControlActiveBall = ActiveBall
  ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
  ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
  If EnableBallControl and ControlBallInPlay then
    If BCright = 1 Then
      ControlActiveBall.velx =  BCvel*BCboost
    ElseIf BCleft = 1 Then
      ControlActiveBall.velx = -BCvel*BCboost
    Else
      ControlActiveBall.velx = 0
    End If

    If BCup = 1 Then
      ControlActiveBall.vely = -BCvel*BCboost
    ElseIf BCdown = 1 Then
      ControlActiveBall.vely =  BCvel*BCboost
    Else
      ControlActiveBall.vely = bcyveloffset
    End If
  End If
End Sub


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 10 ' total number of balls
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9,BallShadow10)

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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub
'*****************************************
Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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

Sub Rubber_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 2 AND finalspeed <= 10 then
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

Sub RandomSoundCycle()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "cycle1"
    Case 2 : PlaySound "cycle2"
    Case 3 : PlaySound "cycle3"
  End Select
End Sub

Sub Kicker1_Hit()
  L1.state=1
  If B2Son Then controller.B2SSetData 101,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker2_Hit()
  L2.state=1
  If B2Son Then controller.B2SSetData 102,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker3_Hit()
  L3.state=1
  If B2Son Then controller.B2SSetData 103,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker4_Hit()
  L4.state=1
  If B2Son Then controller.B2SSetData 104,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker5_Hit()
  L5.state=1
  If B2Son Then controller.B2SSetData 105,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker6_Hit()
  L6.state=1
  If B2Son Then controller.B2SSetData 106,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker7_Hit()
  L7.state=1
  If B2Son Then controller.B2SSetData 107,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker8_Hit()
  L8.state=1
  If B2Son Then controller.B2SSetData 108,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker9_Hit()
  L9.state=1
  If B2Son Then controller.B2SSetData 109,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall:End If
End Sub

Sub Kicker10_Hit()
  L10.state=1
  If B2Son Then controller.B2SSetData 110,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker11_Hit()
  L11.state=1
  If B2Son Then controller.B2SSetData 111,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker12_Hit()
  L12.state=1
  If B2Son Then controller.B2SSetData 112,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker13_Hit()
  L13.state=1
  If B2Son Then controller.B2SSetData 113,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker14_Hit()
  L14.state=1
  If B2Son Then controller.B2SSetData 114,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker15_Hit()
  L15.state=1
  If B2Son Then controller.B2SSetData 115,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker16_Hit()
  L16.state=1
  If B2Son Then controller.B2SSetData 116,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker17_Hit()
  L17.state=1
  If B2Son Then controller.B2SSetData 117,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker18_Hit()
  L18.state=1
  If B2Son Then controller.B2SSetData 118,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker19_Hit()
  L19.state=1
  If B2Son Then controller.B2SSetData 119,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker20_Hit()
  L20.state=1
  If B2Son Then controller.B2SSetData 120,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker21_Hit()
  L21.state=1
  If B2Son Then controller.B2SSetData 121,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker22_Hit()
  L22.state=1
  If B2Son Then controller.B2SSetData 122,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker23_Hit()
  L23.state=1
  If B2Son Then controller.B2SSetData 123,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker24_Hit()
  L24.state=1
  If B2Son Then controller.B2SSetData 124,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Kicker25_Hit()
  L25.state=1
  If B2Son Then controller.B2SSetData 125,1 :End If
  Playsound "WoodHit" : If LastB=1 Then LastBall
End Sub

Sub Trigger2_Hit()
  If YellowLight.state=1 then Playsound "Cycle5"
  F9.state=1
End Sub

Sub Trigger3_Hit()
  If RedLight.state=1 then Playsound "Cycle5"
  F10.state=1
End Sub

Sub Scores()
  RNDM = 4
Select Case Int(Rnd*(RNDM))+1
  Case 1:
  Select Case RNDMR
    Case 1: RArray(0).State=1
    Case 2: RArray(1).State=1
    Case 3: RArray(2).State=1
    Case 4: RArray(3).State=1
    Case 5: RArray(4).State=1
    Case 6: RArray(5).State=1
    Case 7: RArray(6).State=1
    Case 8: RArray(7).State=1
  End Select
    RNDMR = RNDMR+1
    IF RNDMR >7 then RNDMR = 8: End If
  Case 2:
  Select Case RNDMY
    Case 1: YArray(0).State=1
    Case 2: YArray(1).State=1
    Case 3: YArray(2).State=1
    Case 4: YArray(3).State=1
    Case 5: YArray(4).State=1
    Case 6: YArray(5).State=1
    Case 7: YArray(6).State=1
    Case 8: YArray(7).State=1
  End Select
    RNDMY = RNDMY+1
    IF RNDMY >7 then RNDMY = 8: End If
  Case 3:
  Select Case RNDMG
    Case 1: GArray(0).State=1
    Case 2: GArray(1).State=1
    Case 3: GArray(2).State=1
    Case 4: GArray(3).State=1
    Case 5: GArray(4).State=1
    Case 6: GArray(5).State=1
    Case 7: GArray(6).State=1
    Case 8: GArray(7).State=1
  End Select
    RNDMG = RNDMG+1
    IF RNDMG >7 then RNDMG = 8: End If
  Case 4:
  Select Case Int(Rnd*2)+1
    Case 1:
    Case 2:
  End Select
End Select
  IF R8.state=1 then for a = 0 to 6: RArray(a).State=0:Next:Red=8:End If
  IF R7.state=1 then for a = 0 to 5: RArray(a).State=0:Next:Red=7:End If
  IF R6.state=1 then for a = 0 to 4: RArray(a).State=0:Next:Red=6:End If
  IF R5.state=1 then for a = 0 to 3: RArray(a).State=0:Next:Red=5:End If
  IF R4.state=1 then for a = 0 to 2: RArray(a).State=0:Next:Red=4:End If
  IF R3.state=1 then for a = 0 to 1: RArray(a).State=0:Next:Red=3:End If
  IF R2.state=1 then for a = 0 to 0: RArray(a).State=0:Next:Red=2:End If
  IF R1.state=1 then Red = 1:End If
  IF Y8.state=1 then for a = 0 to 6: YArray(a).State=0:Next:Yellow=8:End If
  IF Y7.state=1 then for a = 0 to 5: YArray(a).State=0:Next:Yellow=7:End If
  IF Y6.state=1 then for a = 0 to 4: YArray(a).State=0:Next:Yellow=6:End If
  IF Y5.state=1 then for a = 0 to 3: YArray(a).State=0:Next:Yellow=5:End If
  IF Y4.state=1 then for a = 0 to 2: YArray(a).State=0:Next:Yellow=4:End If
  IF Y3.state=1 then for a = 0 to 1: YArray(a).State=0:Next:Yellow=3:End If
  IF Y2.state=1 then for a = 0 to 0: YArray(a).State=0:Next:Yellow=2:End If
    If Y1.state=1 then Yellow = 1:End If
  IF G8.state=1 then for a = 0 to 6: GArray(a).State=0:Next:Green=8:End If
  IF G7.state=1 then for a = 0 to 5: GArray(a).State=0:Next:Green=7:End If
  IF G6.state=1 then for a = 0 to 4: GArray(a).State=0:Next:Green=6:End If
  IF G5.state=1 then for a = 0 to 3: GArray(a).State=0:Next:Green=5:End If
  IF G4.state=1 then for a = 0 to 2: GArray(a).State=0:Next:Green=4:End If
  IF G3.state=1 then for a = 0 to 1: GArray(a).State=0:Next:Green=3:End If
  IF G2.state=1 then for a = 0 to 0: GArray(a).State=0:Next:Green=2:End If
  IF G1.state=1 then Green = 1:End If
  Odds = Red + Yellow + Green
End Sub

Sub Features()
  If Odds <= 16 Then Odds = 10 :End If
  If Odds > 16 AND Odds <= 20 Then Odds = 15 :End If
  If Odds > 20 then Odds = 20 :End If
  FA=Int(Rnd*6)-1
  If FA = -1 then FA = 0
  RNDM = 4
Select Case Int(Rnd*(RNDM))+1
  Case 1:
  Select Case Int(Rnd*(Odds))+1
    Case 1: For b = fa to fa: FArray(b).State=1:Next
    Case 2: For b = fa to fa: FArray(b).State=1:Next
    Case 3: For b = fa to fa: FArray(b).State=1:Next
    Case 4: For b = fa to fa: FArray(b).State=1:Next
    Case 5: For b = fa to fa: FArray(b).State=1:Next
    Case 6: For b = fa to fa: FArray(b).State=1:Next
    Case 7: For b = fa to fa: FArray(b).State=1:Next
    Case 8: For b = fa to fa: FArray(b).State=1:Next
  End Select

  Case 2:
  Select Case Int(Rnd*(Odds))+1
    Case 1: OK_Game_Timer.enabled=1
    Case 2: SwitchMagicScreen.enabled=1
    Case 3: OK_Game_Timer.enabled=1
    Case 4: SwitchMagicScreen.enabled=1
    Case 5: OK_Game_Timer.enabled=1
    Case 6: SwitchMagicScreen.enabled=1
    Case 7: OK_Game_Timer.enabled=1
    Case 8: SwitchMagicScreen.enabled=1
    Case 9: OK_Game_Timer.enabled=1
    Case 10: SwitchMagicScreen.enabled=1
  End Select

  Case 3:
  Select Case Int(Rnd*(Odds))+1
    Case 1: Tree_Timer.enabled=1
    Case 2: Tree_Timer.enabled=1
    Case 3: ExtraBalls.enabled=1
    Case 4: Tree_Timer.enabled=1
  End Select

  Case 4:
  Select Case Int(Rnd*(Odds-5))+1
    Case 1: SuperMagicScreen.enabled=1
    Case 2:
  End Select
End Select
End Sub

Sub Tree_Timer_Timer()
  Select Case TT
    Case 1: F6.state=1
    Case 2: F7.state=1
    Case 3: F8.state=1
    Case 4: F9.state=1
    Case 5: F10.state=1
  End Select
  TT = TT + 1
  Tree_Timer.enabled=0
End Sub

Sub SwitchMagicScreen_timer()
  Select Case sms
    Case 1: LA.state=1
    Case 2: LB.state=1
    Case 3: LC.state=1
    Case 4: LD.state=1
    Case 5: LE.state=1
    Case 6: LF.state=1
    Case 7: LG.state=1
  End Select
  sms = sms +1
    SwitchMagicScreen.enabled=0
End Sub

Sub SuperMagicScreen_timer()
  Select Case Int(Rnd*3)+1
    Case 1:LA.state=1:LB.state=1:LC.state=1:PlaySound "Cycle4"
    Case 2:LA.state=1:LB.state=1:LC.state=1:LD.state=1:LE.state=1:PlaySound "Cycle4"
    Case 3:LA.state=1:LB.state=1:LC.state=1:LD.state=1:LE.state=1:LF.state=1:LG.state=1:PlaySound "Cycle4"
  End Select
  SuperMagicScreen.enabled=0
End Sub

Sub OK_Game_Timer_timer()
  Select Case Int(Rnd*(Odds))+1
    Case 5:
    Select Case RNDMOK
      Case 1: LK.state=1:RNDMOK = RNDMOK+1
      Case 2: LO.state=1:IF RNDMOK > 2 then RNDMOK = 2
    End Select
  End Select
  OK_Game_Timer.enabled=0
End Sub

Sub ExtraBalls_timer()
  Select Case ex
    Case 1: EB.state=1: Ballz=7
    Case 2: EB1.state=1:Ballz=8
    Case 3: EB2.state=1:Ballz=9
  End Select
  ex = ex + 1
  If ex > 3 then ex = 3
' xxx = Ballz-Ballz+1
  ExtraBalls.enabled=0
End Sub

Sub AwardFeatures()
  If F8.state=1 then RedLight.state=1
  If F7.state=1 then YellowLight.state=1
End Sub

Sub Award()
  Dim R,RR,RRR,RRRR,Y,YY,YYY,YYYY,G,GG,GGG,GGGG,BBB
      If Red = 8 Then AwardLR = 192 : AwardMR = 480 : AwardHR = 600 : End If
      If Red = 7 Then AwardLR = 120 : AwardMR = 240 : AwardHR = 450 : End If
      If Red = 6 Then AwardLR = 64  : AwardMR = 144 : AwardHR = 300 : End If
      If Red = 5 Then AwardLR = 32  : AwardMR = 96  : AwardHR = 200 : End If
      If Red = 4 Then AwardLR = 16  : AwardMR = 50  : AwardHR = 96  : End If
      If Red = 3 Then AwardLR = 8   : AwardMR = 24  : AwardHR = 96  : End If
      If Red = 2 Then AwardLR = 6   : AwardMR = 20  : AwardHR = 75  : End If
      If Red = 1 Then AwardLR = 4   : AwardMR = 16  : AwardHR = 75  : End If

      If Yellow = 8 Then AwardLY = 192 : AwardMY = 480 : AwardHY = 600 : End If
      If Yellow = 7 Then AwardLY = 120 : AwardMY = 240 : AwardHY = 450 : End If
      If Yellow = 6 Then AwardLY = 64  : AwardMY = 144 : AwardHY = 300 : End If
      If Yellow = 5 Then AwardLY = 32  : AwardMY = 96  : AwardHY = 200 : End If
      If Yellow = 4 Then AwardLY = 16  : AwardMY = 50  : AwardHY = 96  : End If
      If Yellow = 3 Then AwardLY = 8   : AwardMY = 24  : AwardHY = 96  : End If
      If Yellow = 2 Then AwardLY = 6   : AwardMY = 20  : AwardHY = 75  : End If
      If Yellow = 1 Then AwardLY = 4   : AwardMY = 16  : AwardHY = 75  : End If

      If Green = 8 Then AwardLG = 192 : AwardMG = 480 : AwardHG = 600 : End If
      If Green = 7 Then AwardLG = 120 : AwardMG = 240 : AwardHG = 450 : End If
      If Green = 6 Then AwardLG = 64  : AwardMG = 144 : AwardHG = 300 : End If
      If Green = 5 Then AwardLG = 32  : AwardMG = 96  : AwardHG = 200 : End If
      If Green = 4 Then AwardLG = 16  : AwardMG = 50  : AwardHG = 96  : End If
      If Green = 3 Then AwardLG = 8   : AwardMG = 24  : AwardHG = 96  : End If
      If Green = 2 Then AwardLG = 6   : AwardMG = 20  : AwardHG = 75  : End If
      If Green = 1 Then AwardLG = 4   : AwardMG = 16  : AwardHG = 75  : End If

  If MX=0 Then
      '3 only rows (Yellow)
      If L9.State=1 And L4.State=1  And L25.State=1 then Credit = Credit + AwardLY: End If
      If L7.State=1 And L22.State=1 And L18.State=1 then Credit = Credit + AwardLY: End If
            '(Green)
      If L2.State=1  And L22.State=1  And L17.State=1 then Credit = Credit + AwardLG: End If
      If L16.State=1 And L13.State=1  And L17.State=1 then Credit = Credit + AwardLG: End If
      If L14.State=1 And L21.State=1  And L17.State=1 then Credit = Credit + AwardLG: End If
      '(Red)
      If L5.State=1 And L21.State=1  And L20.State=1 then Credit = Credit + AwardLR: End If
      If L2.State=1 And L11.State=1  And L15.State=1 then Credit = Credit + AwardLR: End If

      'Red Row
      R=0
      If L11.State=1 And L22.State=1 And L13.State=1 And L21.state=1 And L3.State=1 then R=1:Credit=Credit + AwardHR:End If
      If R=0 And L11.State=1 And L22.State=1 And L13.State=1 And L21.State=1 then R=1:Credit=Credit + AwardMR:End If
      If R=0 And L22.State=1 And L13.State=1 And L21.State=1 And L3.state=1  then R=1:Credit=Credit + AwardMR:End If
      If R=0 And L11.State=1 And L22.State=1 And L13.State=1 then Credit=Credit + AwardLR: End If
      If R=0 And L22.State=1 And L13.State=1 And L21.State=1 then Credit=Credit + AwardLR: End If
      If R=0 And L13.State=1 And L21.State=1 And L3.State=1  then Credit=Credit + AwardLR: End If

      'Yellow Row
      Y=0
      If L2.State=1  And L7.State=1  And L16.State=1 And L5.state=1  And L14.State=1 then Y=1:Credit=Credit + AwardHY:End If
      If Y=0 And L2.State=1  And L7.State=1  And L16.State=1 And L5.State=1  then Y=1:Credit=Credit + AwardMY:End If
      If Y=0 And L7.State=1  And L16.State=1 And L5.State=1  And L14.state=1 then Y=1:Credit=Credit + AwardMY:End If
      If Y=0 And L2.State=1  And L7.State=1  And L16.State=1 then Credit=Credit + AwardLY: End If
      If Y=0 And L7.State=1  And L16.State=1 And L5.State=1  then Credit=Credit + AwardLY: End If
      If Y=0 And L16.State=1 And L5.State=1  And L14.State=1 then Credit=Credit + AwardLY: End If

            'Yellow Row #2
      YY=0
      If L12.State=1 And L8.State=1  And L14.State=1 And L3.state=1  And L10.State=1 then YY=1:Credit=Credit + AwardHY:End If
      If YY=0 And L12.State=1 And L8.State=1  And L14.State=1 And L3.State=1  then YY=1:Credit=Credit + AwardMY:End If
      If YY=0 And L8.State=1  And L14.State=1 And L3.State=1  And L10.state=1 then YY=1:Credit=Credit + AwardMY:End If
      If YY=0 And L12.State=1 And L8.State=1  And L14.State=1 then Credit=Credit + AwardLY: End If
      If YY=0 And L8.State=1  And L14.State=1 And L3.State=1  then Credit=Credit + AwardLY: End If
      If YY=0 And L14.State=1 And L3.State=1  And L10.State=1 then Credit=Credit + AwardLY: End If

            'Green Row
      G=0
      If L15.State=1 And L18.State=1 And L17.State=1 And L20.state=1 And L10.State=1 then G=1:Credit=Credit + AwardHG:End If
      If G=0 And L15.State=1 And L18.State=1 And L17.State=1 And L20.State=1 then G=1:Credit=Credit + AwardMG:End If
      If G=0 And L18.State=1 And L17.State=1 And L20.State=1 And L10.state=1 then G=1:Credit=Credit + AwardMG:End If
      If G=0 And L15.State=1 And L18.State=1 And L17.State=1 then Credit=Credit + AwardLG: End If
      If G=0 And L18.State=1 And L17.State=1 And L20.State=1 then Credit=Credit + AwardLG: End If
      If G=0 And L17.State=1 And L20.State=1 And L10.State=1 then Credit=Credit + AwardLG: End If

      If F5.state=1 Then 'Orange as Green
        GG=0
        If GG=0 And L6.State=1  And L23.State=1 And L24.State=1 And L19.State=1 And L1.State=1 then GG=1:Credit = Credit + AwardHG:OKGAME:End If
        If GG=0 And L6.State=1  And L23.State=1 And L24.State=1 And L19.State=1 then GG=1:Credit = Credit + AwardMG:OKGAME:End If
        If GG=0 And L6.State=1  And L23.State=1 And L24.State=1 And L1.State=1  then GG=1:Credit = Credit + AwardMG:OKGAME:End If
        If GG=0 And L6.State=1  And L23.State=1 And L19.State=1 And L1.State=1  then GG=1:Credit = Credit + AwardMG:OKGAME:End If
        If GG=0 And L6.State=1  And L24.State=1 And L19.State=1 And L1.State=1  then GG=1:Credit = Credit + AwardMG:OKGAME:End If
        If GG=0 And L23.State=1 And L24.State=1 And L19.State=1 And L1.State=1  then GG=1:Credit = Credit + AwardMG:OKGAME:End If
        If GG=0 And L6.State=1  And L23.State=1 And L24.State=1 then GG=1:Credit = Credit + AwardLG:OKGAME:End If
        If GG=0 And L6.State=1  And L23.State=1 And L19.State=1 then GG=1:Credit = Credit + AwardLG:OKGAME:End If
        If GG=0 And L6.State=1  And L23.State=1 And L1.State=1  then GG=1:Credit = Credit + AwardLG:OKGAME:End If
        If GG=0 And L6.State=1  And L24.State=1 And L19.State=1 then GG=1:Credit = Credit + AwardLG:OKGAME:End If
        If GG=0 And L6.State=1  And L24.State=1 And L1.State=1  then GG=1:Credit = Credit + AwardLG:OKGAME:End If
        If GG=0 And L6.State=1  And L19.State=1 And L1.State=1  then GG=1:Credit = Credit + AwardLG:OKGAME:End If
        If GG=0 And L23.State=1 And L24.State=1 And L19.State=1 then GG=1:Credit = Credit + AwardLG:OKGAME:End If
        If GG=0 And L23.State=1 And L24.State=1 And L1.State=1  then GG=1:Credit = Credit + AwardLG:OKGAME:End If
        If GG=0 And L23.State=1 And L19.State=1 And L1.State=1  then GG=1:Credit = Credit + AwardLG:OKGAME:End If
        If GG=0 And L24.State=1 And L19.State=1 And L1.State=1  then GG=1:Credit = Credit + AwardLG:OKGAME:End If
      End If
          '2 in Orange Ok GAME
      If GG=0 And L6.State=1  And L23.State=1 then OKGAME : End If
      If GG=0 And L6.State=1  And L24.State=1 then OKGAME : End If
      If GG=0 And L6.State=1  And L19.State=1 then OKGAME : End If
      If GG=0 And L6.State=1  And L1.State=1  then OKGAME : End If
      If GG=0 And L23.State=1 And L24.State=1 then OKGAME : End If
      If GG=0 And L23.State=1 And L19.State=1 then OKGAME : End If
      If GG=0 And L23.State=1 And L1.State=1  then OKGAME : End If
      If GG=0 And L24.State=1 And L19.State=1 then OKGAME : End If
      If GG=0 And L24.State=1 And L1.State=1  then OKGAME : End If
      If GG=0 And L19.State=1 And L1.State=1  then OKGAME : End If
  End If


  If MX=1 Then
      R=0
      If L1.State=1 And L2.State=1  And L11.State=1  And L15.state=1 then R=1:Credit=Credit + AwardMR:End If
      If R=0 And L1.State=1 And L2.State=1  And L11.State=1 then Credit = Credit + AwardLR: End If
      If R=0 And L2.State=1 And L11.State=1 And L15.State=1 then Credit = Credit + AwardLR: End If

      RR=0
      If L23.State=1 And L5.State=1  And L21.State=1  And L20.state=1 then RR=1:Credit=Credit + AwardMR:End If
      If RR=0 And L23.State=1 And L5.State=1  And L21.State=1 then Credit = Credit + AwardLR: End If
      If RR=0 And L5.State=1  And L21.State=1 And L20.State=1 then Credit = Credit + AwardLR: End If

      RRR=0
      If L2.State=1  And L7.State=1  And L16.State=1 And L5.state=1 And L14.State=1 then RRR=1:Credit=Credit + AwardHR:End If
      If RRR=0 And L2.State=1  And L7.State=1  And L16.State=1 And L5.State=1  then RRR=1:Credit=Credit + AwardMR:End If
      If RRR=0 And L7.State=1  And L16.State=1 And L5.State=1  And L14.state=1 then RRR=1: Credit=Credit + AwardMR:End If
      If RRR=0 And L2.State=1  And L7.State=1  And L16.State=1 then Credit=Credit + AwardLR: End If
      If RRR=0 And L7.State=1  And L16.State=1 And L5.State=1  then Credit=Credit + AwardLR: End If
      If RRR=0 And L16.State=1 And L5.State=1  And L14.State=1 then Credit=Credit + AwardLR: End If

      Y=0
      If L1.State=1  And L19.State=1 And L24.State=1 And L23.state=1 And L8.State=1 then Y=1:Credit=Credit + AwardHY:End If
      If Y=0 And L1.State=1  And L19.State=1 And L24.State=1 And L23.State=1 then Y=1:Credit=Credit + AwardMY:End If
      If Y=0 And L19.State=1 And L24.State=1 And L23.State=1 And L8.state=1  then Y=1:Credit=Credit + AwardMY:End If
      If Y=0 And L1.State=1  And L19.State=1 And L24.State=1 then Credit=Credit + AwardLY: End If
      If Y=0 And L19.State=1 And L24.State=1 And L23.State=1 then Credit=Credit + AwardLY: End If
      If Y=0 And L24.State=1 And L23.State=1 And L8.State=1  then Credit=Credit + AwardLY: End If

      YY=0
      If L15.State=1 And L18.State=1 And L17.State=1 And L20.state=1 And L10.State=1 then YY=1:Credit=Credit + AwardHY:End If
      If YY=0 And L15.State=1 And L18.State=1 And L17.State=1 And L20.State=1 then YY=1:Credit=Credit + AwardMY:End If
      If YY=0 And L18.State=1 And L17.State=1 And L20.State=1 And L10.state=1 then YY=1:Credit=Credit + AwardMY:End If
      If YY=0 And L15.State=1 And L18.State=1 And L17.State=1 then Credit=Credit + AwardLY: End If
      If YY=0 And L18.State=1 And L17.State=1 And L20.State=1 then Credit=Credit + AwardLY: End If
      If YY=0 And L17.State=1 And L20.State=1 And L10.State=1 then Credit=Credit + AwardLY: End If

      YYY=0
      If L12.State=1 And L8.State=1  And L14.State=1 And L3.state=1  And L10.State=1 then YYY=1:Credit=Credit + AwardHY:End If
      If YYY=0 And L12.State=1 And L8.State=1  And L14.State=1 And L3.State=1  then YYY=1:Credit=Credit + AwardMY:End If
      If YYY=0 And L8.State=1  And L14.State=1 And L3.State=1  And L10.state=1 then YYY=1:Credit=Credit + AwardMY:End If
      If YYY=0 And L12.State=1 And L8.State=1  And L14.State=1 then Credit=Credit + AwardLY: End If
      If YYY=0 And L8.State=1  And L14.State=1 And L3.State=1  then Credit=Credit + AwardLY: End If
      If YYY=0 And L14.State=1 And L3.State=1  And L10.State=1 then Credit=Credit + AwardLY: End If

      YYYY=0
      If L19.State=1 And L7.State=1  And L22.State=1 And L18.state=1 then YYYY=1:Credit=Credit + AwardMY:End If
      If YYYY=0 And L19.State=1 And L7.State=1  And L22.State=1 then Credit = Credit + AwardLY: End If
      If YYYY=0 And  L7.State=1  And L22.State=1 And L18.State=1 then Credit=Credit + AwardLY: End If

      G=0
      If L11.State=1 And L22.State=1 And L13.State=1 And L21.state=1 And L3.State=1 then G=1:Credit=Credit + AwardHG:End If
      If G=0 And L11.State=1 And L22.State=1 And L13.State=1 And L21.State=1 then G=1:Credit=Credit + AwardMG:End If
      If G=0 And L22.State=1 And L13.State=1 And L21.State=1 And L3.state=1  then G=1:Credit=Credit + AwardMG:End If
      If G=0 And L11.State=1 And L22.State=1 And L13.State=1 then Credit=Credit + AwardLG: End If
      If G=0 And L22.State=1 And L13.State=1 And L21.State=1 then Credit=Credit + AwardLG: End If
      If G=0 And L13.State=1 And L21.State=1 And L3.State=1  then Credit=Credit + AwardLG: End If

      GG=0
      If L1.State=1  And L7.State=1  And L13.State=1  And L20.state=1 then GG=1:Credit=Credit + AwardMG:End If
      If GG=0 And L1.State=1  And L7.State=1  And L13.State=1 then Credit = Credit + AwardLG: End If
      If GG=0 And L7.State=1  And L13.State=1 And L20.State=1 then Credit=Credit + AwardLG: End If

      GGG=0
      If L24.State=1  And L16.State=1 And L13.State=1 And L17.state=1 then GGG=1:Credit = Credit + AwardMG:End If
      If GGG=0 And L24.State=1  And L16.State=1 And L13.State=1 then Credit = Credit + AwardLG: End If
      If GGG=0 And L16.State=1  And L13.State=1 And L17.State=1 then Credit = Credit + AwardLG: End If

      GGGG=0
      If L8.State=1  And L5.State=1  And L13.State=1 And L18.state=1 then GGGG=1:Credit=Credit + AwardMG:End If
      If GGGG=0 And L8.State=1  And L5.State=1  And L13.State=1 then Credit = Credit + AwardLG: End If
      If GGGG=0 And L5.State=1  And L13.State=1 And L18.State=1 then Credit = Credit + AwardLG: End If

      If F5.state=1 Then 'Orange as Green
      BBB=0
        If BBB=0 And L9.State=1 And L4.State=1  And L25.State=1 And L6.State=1 then BBB=1:Credit=Credit + AwardMG::OKGAME:End If
        If BBB=0 And L9.State=1 And L4.State=1  And L25.State=1 then BBB=1:Credit=Credit + AwardLG:OKGAME:End If
        If BBB=0 And L9.State=1 And L4.State=1  And L6.State=1  then BBB=1:Credit=Credit + AwardLG:OKGAME:End If
        If BBB=0 And L9.State=1 And L25.State=1 And L6.State=1  then BBB=1:Credit=Credit + AwardLG:OKGAME:End If
        If BBB=0 And L4.State=1 And L25.State=1 And L6.State=1  then BBB=1:Credit=Credit + AwardLG:OKGAME:End If
      End If

      '2 in OK Game
      If BBB=0 And L9.State=1  And L4.State=1  then OKGAME :End If
      If BBB=0 And L9.State=1  And L25.State=1 then OKGAME :End If
      If BBB=0 And L9.State=1  And L6.State=1  then OKGAME :End If
      If BBB=0 And L4.State=1  And L25.State=1 then OKGAME :End If
      If BBB=0 And L4.State=1  And L6.State=1  then OKGAME :End If
      If BBB=0 And L25.State=1 And L6.State=1  then OKGAME :End If
  End If


  If MX=2 Then 'All Stripe Card
      R=0
      If L9.State=1 And L1.State=1  And L2.State=1  And L11.state=1 And L15.State=1 then R=1:Credit=Credit + AwardHR:End If
      If R=0 And L9.State=1 And L1.State=1  And L2.State=1  And L11.state=1 then R=1:Credit=Credit + AwardMR:End If
      If R=0 And L1.State=1 And L2.State=1  And L11.State=1 And L15.state=1 then R=1:Credit=Credit + AwardMR:End If
      If R=0 And L9.State=1 And L1.State=1  And L2.State=1  then Credit=Credit + AwardLR: End If
      If R=0 And L1.State=1 And L2.State=1  And L11.State=1 then Credit=Credit + AwardLR: End If
      If R=0 And L2.State=1 And L11.State=1 And L15.State=1 then Credit=Credit + AwardLR: End If

            RR=0
      If L1.State=1  And L19.State=1 And L24.State=1 And L23.state=1 And L8.State=1 then RR=1:Credit=Credit + AwardHR:End If
      If RR=0 And L1.State=1  And L19.State=1 And L24.State=1 And L23.State=1 then RR=1:Credit=Credit + AwardMR:End If
      If RR=0 And L19.State=1 And L24.State=1 And L23.State=1 And L8.state=1  then RR=1:Credit=Credit + AwardMR:End If
      If RR=0 And L1.State=1  And L19.State=1 And L24.State=1 then Credit=Credit + AwardLR: End If
      If RR=0 And L19.State=1 And L24.State=1 And L23.State=1 then Credit=Credit + AwardLR: End If
      If RR=0 And L24.State=1 And L23.State=1 And L8.State=1  then Credit=Credit + AwardLR: End If

            RRR=0
        If L15.State=1 And L18.State=1 And L17.State=1 And L20.state=1 And L10.State=1 then RRR=1:Credit=Credit + AwardHR:End If
      If RRR=0 And L15.State=1 And L18.State=1 And L17.State=1  And L20.State=1 then RRR=1:Credit=Credit + AwardMR:End If
      If RRR=0 And  L18.State=1 And L17.State=1 And L20.State=1 And L10.state=1 then RRR=1:Credit=Credit + AwardMR:End If
      If RRR=0 And  L15.State=1 And L18.State=1 And L17.State=1 then Credit=Credit + AwardLR: End If
      If RRR=0 And  L18.State=1 And L17.State=1 And L20.State=1 then Credit=Credit + AwardLR: End If
      If RRR=0 And  L17.State=1 And L20.State=1 And L10.State=1 then Credit=Credit + AwardLR: End If

      RRRR=0
      If L6.State=1 And L23.State=1 And L5.State=1  And L21.state=1 And L20.State=1 then RRRR=1:Credit=Credit + AwardHR:End If
      If RRRR=0 And L6.State=1 And L23.State=1 And L5.State=1  And L21.State=1 then RRRR=1:Credit=Credit + AwardMR:End If
      If RRRR=0 And L23.State=1 And L5.State=1 And L21.State=1 And L20.state=1 then RRRR=1:Credit=Credit + AwardMR:End If
      If RRRR=0 And L6.State=1 And L23.State=1 And L5.State=1  then Credit=Credit + AwardLR: End If
      If RRRR=0 And L23.State=1 And L5.State=1 And L21.State=1 then Credit=Credit + AwardLR: End If
      If RRRR=0 And L5.State=1 And L21.State=1 And L20.State=1 then Credit=Credit + AwardLR: End If

      Y=0
      If L9.State=1 And L4.State=1 And L25.State=1 And L6.state=1 And L12.State=1 then Y=1:Credit=Credit + AwardHY:End If
      If Y=0 And L9.State=1  And L4.State=1  And L25.State=1 And L6.state=1  then Y=1:Credit=Credit + AwardMY:End If
      If Y=0 And L4.State=1  And L25.State=1 And L6.State=1  And L12.state=1 then Y=1:Credit=Credit + AwardMY:End If
      If Y=0 And L9.State=1  And L4.State=1  And L25.State=1 then Credit=Credit+AwardLY: End If
      If Y=0 And L4.State=1  And L25.State=1 And L6.State=1  then Credit=Credit+AwardLY: End If
      If Y=0 And L25.State=1 And L6.State=1  And L12.State=1 then Credit=Credit+AwardLY: End If

      YY=0
      If L4.State=1 And L19.State=1 And L7.State=1 And L22.state=1 And L18.State=1 then YY=1:Credit=Credit + AwardHY:End If
      If YY=0 And L4.State=1  And L19.State=1 And L7.State=1  And L22.State=1 then YY=1:Credit=Credit + AwardMY:End If
      If YY=0 And L19.State=1 And L7.State=1  And L22.State=1 And L18.state=1 then YY=1:Credit=Credit + AwardMY:End If
      If YY=0 And L4.State=1  And L19.State=1 And L7.State=1  then Credit=Credit + AwardLY: End If
      If YY=0 And L19.State=1 And L7.State=1  And L22.State=1 then Credit=Credit + AwardLY: End If
      If YY=0 And L7.State=1  And L22.State=1 And L18.State=1 then Credit=Credit + AwardLY: End If

      YYY=0
      If L12.State=1 And L8.State=1 And L14.State=1 And L3.state=1 And L10.State=1 then YYY=1:Credit=Credit + AwardHY:End If
      If YYY=0 And L12.State=1 And L8.State=1  And L14.State=1 And L3.State=1  then YYY=1:Credit=Credit + AwardMY:End If
      If YYY=0 And L8.State=1  And L14.State=1 And L3.State=1  And L10.state=1 then YYY=1:Credit=Credit + AwardMY:End If
      If YYY=0 And L12.State=1 And L8.State=1  And L14.State=1 then Credit=Credit + AwardLY: End If
      If YYY=0 And L8.State=1  And L14.State=1 And L3.State=1  then Credit=Credit + AwardLY: End If
      If YYY=0 And L14.State=1 And L3.State=1  And L10.State=1 then Credit=Credit + AwardLY: End If

      YYYY=0
      If L11.State=1 And L22.State=1 And L13.State=1 And L21.state=1 And L3.State=1 then YYYY=1:Credit=Credit + AwardHY:End If
      If YYYY=0 And L11.State=1 And L22.State=1 And L13.State=1 And L21.State=1 then YYYY=1:Credit=Credit + AwardMY:End If
      If YYYY=0 And L22.State=1 And L13.State=1 And L21.State=1 And L3.state=1  then YYYY=1:Credit=Credit + AwardMY:End If
      If YYYY=0 And L11.State=1 And L22.State=1 And L13.State=1 then Credit=Credit + AwardLY: End If
      If YYYY=0 And L22.State=1 And L13.State=1 And L21.State=1 then Credit=Credit + AwardLY: End If
      If YYYY=0 And L13.State=1 And L21.State=1 And L3.State=1  then Credit=Credit + AwardLY: End If

      G=0
      If L9.State=1  And L19.State=1 And L16.State=1 And L21.state=1 And L10.State=1 then G=1:Credit=Credit + AwardHG:End If
      If G=0 And L9.State=1  And L19.State=1 And L16.State=1 And L21.state=1 then G=1:Credit=Credit + AwardMG:End If
      If G=0 And L19.State=1 And L16.State=1 And L21.State=1 And L10.state=1 then G=1:Credit=Credit + AwardMG:End If
      If G=0 And L9.State=1  And L19.State=1 And L16.State=1 then Credit=Credit + AwardLG: End If
      If G=0 And L19.State=1 And L16.State=1 And L21.State=1 then Credit=Credit + AwardLG: End If
      If G=0 And L16.State=1 And L21.State=1 And L10.State=1 then Credit=Credit +AwardLG: End If

      GG=0
      If L25.State=1 And L24.State=1 And L16.State=1 And L13.state=1 And L17.State=1 then GG=1:Credit=Credit + AwardHG:End If
      If GG=0 And L25.State=1 And L24.State=1 And L16.State=1 And L13.State=1 then GG=1:Credit=Credit + AwardMG:End If
      If GG=0 And L24.State=1 And L16.State=1 And L13.State=1 And L17.state=1 then GG=1:Credit=Credit + AwardMG:End If
      If GG=0 And L25.State=1 And L24.State=1 And L16.State=1 then Credit=Credit + AwardLG: End If
      If GG=0 And L24.State=1 And L16.State=1 And L13.State=1 then Credit=Credit + AwardLG: End If
      If GG=0 And L16.State=1 And L13.State=1 And L17.State=1 then Credit=Credit + AwardLG: End If

      GGG=0
      If L12.State=1 And L23.State=1 And L16.State=1 And L22.state=1 And L15.State=1 then GGG=1:Credit=Credit + AwardHG:End If
      If GGG=0 And L12.State=1 And L23.State=1 And L16.State=1 And L22.State=1 then GGG=1:Credit=Credit + AwardMG:End If
      If GGG=0 And L23.State=1 And L16.State=1 And L22.State=1 And L15.state=1 then GGG=1:Credit=Credit + AwardMG:End If
      If GGG=0 And L12.State=1 And L23.State=1 And L16.State=1 then Credit=Credit + AwardLG: End If
      If GGG=0 And L23.State=1 And L16.State=1 And L22.State=1 then Credit=Credit + AwardLG: End If
      If GGG=0 And L16.State=1 And L22.State=1 And L15.State=1 then Credit=Credit + AwardLG: End If

      GGGG=0
      If L2.State=1  And L7.State=1  And L16.State=1 And L5.state=1  And L14.State=1 then GGGG=1:Credit=Credit + AwardHG:End If
      If GGGG=0 And L2.State=1  And L7.State=1  And L16.State=1 And L5.State=1  then GGGG=1:Credit=Credit + AwardMG:End If
      If GGGG=0 And L7.State=1  And L16.State=1 And L5.State=1  And L14.state=1 then GGGG=1:Credit=Credit + AwardMG:End If
      If GGGG=0 And L2.State=1  And L7.State=1  And L16.State=1 then Credit=Credit + AwardLG: End If
      If GGGG=0 And L7.State=1  And L16.State=1 And L5.State=1  then Credit=Credit + AwardLG: End If
      If GGGG=0 And L16.State=1 And L5.State=1  And L14.State=1 then Credit=Credit + AwardLG: End If
  End If


  If MX=3 Then 'Card 1
      R=0
      If L11.State=1 And L22.State=1 And L13.State=1 And L21.state=1 And L3.State=1 then R=1:Credit=Credit + AwardHR:End If
      If R=0 And L11.State=1 And L22.State=1 And L13.State=1 And L21.State=1 then R=1:Credit=Credit + AwardMR:End If
      If R=0 And L22.State=1 And L13.State=1 And L21.State=1 And L3.state=1  then R=1:Credit=Credit + AwardMR:End If
      If R=0 And L11.State=1 And L22.State=1 And L13.State=1 then Credit=Credit + AwardLR: End If
      If R=0 And L22.State=1 And L13.State=1 And L21.State=1 then Credit=Credit + AwardLR: End If
      If R=0 And L13.State=1 And L21.State=1 And L3.State=1  then Credit=Credit + AwardLR: End If

      RR=0
      If L9.State=1 And L1.State=1  And L2.State=1  And L11.state=1 then RR=1:Credit=Credit + AwardMR:End If
      If RR=0 And L9.State=1 And L1.State=1  And L2.State=1  then Credit=Credit + AwardLR: End If
      If RR=0 And L1.State=1 And L2.State=1  And L11.State=1 then Credit=Credit + AwardLR: End If

      RRR=0
      If L9.State=1  And L4.State=1  And L25.State=1 And L6.State=1 then RRR=1:Credit=Credit + AwardMR:End If
      If RRR=0 And L9.State=1 And L4.State=1  And L25.State=1 then Credit=Credit + AwardLR: End If
      If RRR=0 And L4.State=1 And L25.State=1 And L6.State=1  then Credit=Credit + AwardLR: End If

      RRRR=0
      If L6.State=1  And L23.State=1 And L5.State=1  And L21.State=1 then RRRR=1:Credit=Credit + AwardMR:End If
      If RRRR=0 And L6.State=1  And L23.State=1 And L5.State=1  then Credit=Credit + AwardLR: End If
      If RRRR=0 And L23.State=1 And L5.State=1  And L21.State=1 then Credit=Credit + AwardLR: End If

      Y=0
      If L2.State=1  And L7.State=1  And L16.State=1 And L5.state=1 And L14.State=1 then Y=1:Credit=Credit + AwardHY:End If
      If Y=0 And L2.State=1  And L7.State=1  And L16.State=1 And L5.state=1  then Y=1:Credit=Credit + AwardMY:End If
      If Y=0 And L7.State=1  And L16.State=1 And L5.State=1  And L14.state=1 then Y=1:Credit=Credit + AwardMY:End If
      If Y=0 And L2.State=1  And L7.State=1  And L16.State=1 then Credit=Credit+ AwardLY: End If
      If Y=0 And L7.State=1  And L16.State=1 And L5.State=1  then Credit=Credit+ AwardLY: End If
      If Y=0 And L16.State=1 And L5.State=1  And L14.State=1 then Credit=Credit+AwardLY: End If

      YY=0
      If L4.State=1  And L19.State=1 And L7.State=1  And L22.State=1 then YY=1:Credit=Credit + AwardMY:End If
      If YY=0 And L4.State=1  And L19.State=1 And L7.State=1  then Credit=Credit + AwardLY: End If
      If YY=0 And L19.State=1 And L7.State=1  And L22.State=1 then Credit=Credit + AwardLY: End If

      YYY=0
      If L12.State=1 And L8.State=1 And L14.State=1 And L3.State=1 then YYY=1:Credit=Credit + AwardMY:End If
      If YYY=0 And L12.State=1 And L8.State=1 And L14.State=1 then Credit=Credit + AwardLY: End If
      If YYY=0 And L8.State=1 And L14.State=1 And L3.State=1  then Credit=Credit + AwardLY: End If

      'Yellow Super Section
      If F4.State=1 Then
        YYYY=0
        If L15.State=1 And L18.State=1 And L17.State=1 then YYYY=1:Credit=Credit + AwardMY: End If
        If YYYY=0 And L15.State=1 And L18.State=1 then YYYY=1:Credit=Credit + AwardLY: End If
        If YYYY=0 And L15.State=1 And L17.State=1 then YYYY=1:Credit=Credit + AwardLY: End If
        If YYYY=0 And L17.State=1 And L18.State=1 then YYYY=1:Credit=Credit + AwardLY: End If
      End If
      'Yellow Stripes 3
      If YYYY=0 And L15.State=1 And L18.State=1 And L17.State=1 then Credit=Credit + AwardLY: End If

      G=0
      If L1.State=1  And L19.State=1 And L24.State=1 And L23.state=1  And L8.State=1 then G=1:Credit=Credit + AwardHG:End If
      If G=0 And L1.State=1  And L19.State=1 And L24.State=1 And L23.State=1 then G=1:Credit=Credit + AwardMG:End If
      If G=0 And L19.State=1 And L24.State=1 And L23.State=1 And L8.state=1  then G=1:Credit=Credit + AwardMG:End If
      If G=0 And L1.State=1  And L19.State=1 And L24.State=1 then Credit=Credit + AwardLG: End If
      If G=0 And L19.State=1 And L24.State=1 And L23.State=1 then Credit=Credit + AwardLG: End If
      If G=0 And L24.State=1 And L23.State=1 And L8.State=1  then Credit=Credit + AwardLG: End If

      GG=0
      If L4.State=1  And L24.State=1 And L5.State=1 And L3.state=1 then GG=1:Credit=Credit + AwardMG:End If
      If GG=0 And L4.State=1  And L24.State=1 And L5.State=1 then Credit=Credit + AwardLG: End If
      If GG=0 And L24.State=1 And L5.State=1  And L3.State=1 then Credit=Credit + AwardLG: End If

      GGG=0
      If L25.State=1 And L24.State=1 And L16.State=1 And L13.State=1 then GGG=1:Credit=Credit + AwardMG:End If
      If GGG=0 And L25.State=1 And L24.State=1 And L16.State=1 then Credit=Credit + AwardLG: End If
      If GGG=0 And L24.State=1 And L16.State=1 And L13.State=1 then Credit=Credit + AwardLG: End If

      GGGG=0
      If L6.State=1  And L24.State=1 And L7.State=1  And L11.State=1 then GGGG=1:Credit=Credit + AwardMG:End If
      If GGGG=0 And L6.State=1  And L24.State=1 And L7.State=1  then Credit=Credit + AwardLG: End If
      If GGGG=0 And L24.State=1 And L7.State=1  And L11.State=1 then Credit=Credit + AwardLG: End If
  End If


  If MX=4 Then 'Card 2
      R=0
      If L2.State=1  And L7.State=1  And L16.State=1 And L5.state=1 And L14.State=1 then R=1:Credit=Credit + AwardHR:End If
      If R=0 And L2.State=1  And L7.State=1  And L16.State=1 And L5.state=1  then R=1:Credit=Credit + AwardMR:End If
      If R=0 And L7.State=1  And L16.State=1 And L5.State=1  And L14.state=1 then R=1:Credit=Credit + AwardMR:End If
      If R=0 And L2.State=1  And L7.State=1  And L16.State=1 then Credit=Credit+AwardLR: End If
      If R=0 And L7.State=1  And L16.State=1 And L5.State=1  then Credit=Credit+AwardLR: End If
      If R=0 And L16.State=1 And L5.State=1  And L14.State=1 then Credit=Credit+AwardLR: End If

            'Red 3 Choice only Rows
      If L9.State=1  And L1.State=1  And L2.State=1  then Credit=Credit + AwardLR: End If
      If L6.State=1  And L23.State=1 And L5.State=1  then Credit=Credit + AwardLR: End If


      'Red Super Section 2 as 3, 3 as 4
      If F3.State=1 Then
        RR=0
        If L21.State=1 And L3.State=1  And L10.State=1 then RR=1:Credit=Credit + AwardMR: End If
        If RR=0 And L21.State=1 And L3.State=1  then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L21.State=1 And L10.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L3.State=1  And L10.State=1 then RR=1:Credit=Credit + AwardLR: End If
      End If
      If F3.State=0 Then
        If L21.State=1 And L3.State=1  And L10.State=1 then Credit=Credit + AwardLR: End If
      End If


      'Yellow 3 Choice Rows
      If L4.State=1 And L19.State=1 And L7.State=1  then Credit=Credit + AwardLY: End If
      If L12.State=1 And L8.State=1 And L14.State=1 then Credit=Credit + AwardLY: End If

      Y=0
      If L1.State=1  And L19.State=1 And L24.State=1 And L23.state=1 And L8.State=1 then Y=1:Credit=Credit + AwardHY:End If
      If Y=0 And L1.State=1  And L19.State=1 And L24.State=1 And L23.State=1 then Y=1:Credit=Credit + AwardMY:End If
      If Y=0 And L19.State=1 And L24.State=1 And L23.State=1 And L8.state=1  then Y=1:Credit=Credit + AwardMY:End If
      If Y=0 And L1.State=1  And L19.State=1 And L24.State=1 then Credit=Credit + AwardLY: End If
      If Y=0 And L19.State=1 And L24.State=1 And L23.State=1 then Credit=Credit + AwardLY: End If
      If Y=0 And L24.State=1 And L23.State=1 And L8.State=1  then Credit=Credit + AwardLY: End If

      'yellow super section feature 2 as 3, 3 as 4, 4 as 5
      If F4.State=1 Then
        YY=0
        If L11.State=1 And L15.State=1 And L22.State=1 And L13.State=1 then YY=1:Credit=Credit + AwardHY:End If
        If YY=0 And L11.State=1 And L15.State=1 And L22.State=1 then YY=1:Credit=Credit + AwardMY: End If
        If YY=0 And L11.State=1 And L15.State=1 And L13.State=1 then YY=1:Credit=Credit + AwardMY: End If
        If YY=0 And L11.State=1 And L22.State=1 And L13.State=1 then YY=1:Credit=Credit + AwardMY: End If
        If YY=0 And L15.State=1 And L22.State=1 And L13.State=1 then YY=1:Credit=Credit + AwardMY: End If
        If YY=0 And L11.State=1 And L15.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L11.State=1 And L22.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L11.State=1 And L13.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L15.State=1 And L22.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L15.State=1 And L13.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L22.State=1 And L13.State=1 then YY=1:Credit=Credit + AwardLY: End If
      End If
      If F4.State=0 Then
      YYY=0
      If L11.State=1 And L15.State=1 And L22.State=1 And L13.State=1 then YYY=1:Credit=Credit + AwardMY:End If
      If YYY=0 And L11.State=1 And L15.State=1 And L22.State=1 then Credit=Credit + AwardLY: End If
      If YYY=0 And L11.State=1 And L15.State=1 And L13.State=1 then Credit=Credit + AwardLY: End If
      If YYY=0 And L11.State=1 And L22.State=1 And L13.State=1 then Credit=Credit + AwardLY: End If
      If YYY=0 And L15.State=1 And L22.State=1 And L13.State=1 then Credit=Credit + AwardLY: End If
      End If

      G=0
      If L9.State=1  And L4.State=1  And L25.State=1 And L6.state=1  And L12.State=1 then G=1:Credit=Credit + AwardHG:End If
      If G=0 And L9.State=1  And L4.State=1  And L25.State=1 And L6.state=1  then G=1:Credit=Credit + AwardMG:End If
      If G=0 And L4.State=1  And L25.State=1 And L6.State=1  And L12.state=1 then G=1:Credit=Credit + AwardMG:End If
      If G=0 And L9.State=1  And L4.State=1  And L25.State=1 then Credit=Credit+AwardLG: End If
      If G=0 And L4.State=1  And L25.State=1 And L6.State=1  then Credit=Credit+AwardLG: End If
      If G=0 And L25.State=1 And L6.State=1  And L12.State=1 then Credit=Credit+AwardLG: End If

            'Green 3 Choice Rows
      If L25.State=1 And L19.State=1 And L2.State=1  then Credit=Credit + AwardLG: End If
      If L25.State=1 And L24.State=1 And L16.State=1 then Credit=Credit + AwardLG: End If
      If L25.State=1 And L23.State=1 And L14.State=1 then Credit=Credit + AwardLG: End If
      If L18.State=1 And L17.State=1 And L20.State=1 then Credit=Credit + AwardLG: End If
  End If

  If MX=5 Then

      R=0
      If L1.State=1  And L19.State=1 And L24.State=1 And L23.state=1  And L8.State=1 then R=1:Credit=Credit + AwardHR:End If
      If R=0 And L1.State=1  And L19.State=1 And L24.State=1 And L23.State=1 then R=1:Credit=Credit + AwardMR:End If
      If R=0 And L19.State=1 And L24.State=1 And L23.State=1 And L8.State=1  then R=1:Credit=Credit + AwardMR:End If
      If R=0 And L1.State=1  And L19.State=1 And L24.State=1 then Credit=Credit + AwardLR: End If
      If R=0 And L19.State=1 And L24.State=1 And L23.State=1 then Credit=Credit + AwardLR: End If
      If R=0 And L24.State=1 And L23.State=1 And L8.State=1  then Credit=Credit + AwardLR: End If

      'red super section feature 2 as 3, 3 as 4, 4 as 5
      If F3.State=1 Then
        RR=0
        If L5.State=1  And L14.State=1 And L3.State=1  And L10.State=1 then RR=1:Credit=Credit + AwardHR:End If
        If RR=0 And L5.State=1  And L14.State=1 And L3.State=1  And L20.State=1 then RR=1:Credit=Credit + AwardHR:End If
        If RR=0 And L5.State=1  And L3.State=1  And L10.State=1 And L20.State=1 then RR=1:Credit=Credit + AwardHR:End If
        If RR=0 And L5.State=1  And L14.State=1 And L10.State=1 And L20.State=1 then RR=1:Credit=Credit + AwardHR:End If
        If RR=0 And L14.State=1 And L3.State=1  And L10.State=1 And L20.State=1 then Credit=Credit + AwardHR:End If
        If RR=0 And L5.State=1  And L14.State=1 And L3.State=1  then RR=1:Credit=Credit + AwardMR: End If
        If RR=0 And L5.State=1  And L3.State=1  And L10.State=1 then RR=1:Credit=Credit + AwardMR: End If
        If RR=0 And L5.State=1  And L10.State=1 And L20.State=1 then RR=1:Credit=Credit + AwardMR: End If
        If RR=0 And L14.State=1 And L3.State=1  And L10.State=1 then RR=1:Credit=Credit + AwardMR: End If
        If RR=0 And L14.State=1 And L10.State=1 And L20.State=1 then RR=1:Credit=Credit + AwardMR: End If
        If RR=0 And L3.State=1  And L10.State=1 And L20.State=1 then RR=1:Credit=Credit + AwardMR: End If
        If RR=0 And L5.State=1  And L14.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L5.State=1  And L3.State=1  then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L5.State=1  And L10.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L5.State=1  And L20.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L14.State=1 And L3.State=1  then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L14.State=1 And L10.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L14.State=1 And L20.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L3.State=1  And L10.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L3.State=1  And L20.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L10.State=1 And L20.State=1 then RR=1:Credit=Credit + AwardLR: End If
      End If
      'Red Stripes 5 choices:
      If F3.State=0 Then
        RRR=0
        If L5.State=1  And L14.State=1 And L3.State=1  And L10.State=1 And L20.State=1 then RRR=1:Credit=Credit + AwardHR:End If
        If RRR=0 And L5.State=1  And L14.State=1 And L3.State=1  And L10.State=1 then RRR=1:Credit=Credit + AwardMR:End If
        If RRR=0 And L5.State=1  And L14.State=1 And L3.State=1  And L20.State=1 then RRR=1:Credit=Credit + AwardMR:End If
        If RRR=0 And L5.State=1  And L3.State=1  And L10.State=1 And L20.State=1 then RRR=1:Credit=Credit + AwardMR:End If
        If RRR=0 And L5.State=1  And L14.State=1 And L10.State=1 And L20.State=1 then RRR=1:Credit=Credit + AwardMR:End If
        If RRR=0 And L14.State=1 And L3.State=1  And L10.State=1 And L20.State=1 then RRR=1:Credit=Credit + AwardMR:End If
        If RRR=0 And L5.State=1  And L14.State=1 And L3.State=1  then RRR=1:Credit=Credit + AwardLR: End If
        If RRR=0 And L5.State=1  And L14.State=1 And L10.State=1 then RRR=1:Credit=Credit + AwardLR: End If
        If RRR=0 And L5.State=1  And L14.State=1 And L20.State=1 then RRR=1:Credit=Credit + AwardLR: End If
        If RRR=0 And L5.State=1  And L3.State=1  And L10.State=1 then RRR=1:Credit=Credit + AwardLR: End If
        If RRR=0 And L5.State=1  And L3.State=1  And L20.State=1 then RRR=1:Credit=Credit + AwardLR: End If
        If RRR=0 And L5.State=1  And L10.State=1 And L20.State=1 then RRR=1:Credit=Credit + AwardLR: End If
        If RRR=0 And L14.State=1 And L3.State=1  And L10.State=1 then RRR=1:Credit=Credit + AwardLR: End If
        If RRR=0 And L14.State=1 And L3.State=1  And L20.State=1 then RRR=1:Credit=Credit + AwardLR: End If
        If RRR=0 And L14.State=1 And L10.State=1 And L20.State=1 then RRR=1:Credit=Credit + AwardLR: End If
        If RRR=0 And L3.State=1  And L10.State=1 And L20.State=1 then RRR=1:Credit=Credit + AwardLR: End If
      End If

      'yellow super section feature 2 as 3, 3 as 4, 4 as 5
      If F4.State=1  Then
        Y=0
        If L16.State=1 And L7.State=1  And  L2.State=1  And L11.State=1 then Y=1:Credit=Credit  + AwardHY:End If
        If Y=0 And L16.State=1 And L7.State=1  And  L2.State=1  And L15.State=1 then Y=1:Credit=Credit  + AwardHY:End If
        If Y=0 And L16.State=1 And L2.State=1  And  L11.State=1 And L15.State=1 then Y=1:Credit=Credit  + AwardHY:End If
        If Y=0 And L16.State=1 And L7.State=1  And  L11.State=1 And L15.State=1 then Y=1:Credit=Credit  + AwardHY:End If
        If Y=0 And L7.State=1  And L2.State=1  And  L11.State=1 And L15.State=1 then Y=1:Credit=Credit  + AwardHY:End If
        If Y=0 And L16.State=1 And L7.State=1  And L2.State=1  then Y=1:Credit=Credit + AwardMY: End If
        If Y=0 And L16.State=1 And L7.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardMY: End If
        If Y=0 And L16.State=1 And L7.State=1  And L15.State=1 then Y=1:Credit=Credit + AwardMY: End If
        If Y=0 And L16.State=1 And L2.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardMY: End If
        If Y=0 And L16.State=1 And L2.State=1  And L15.State=1 then Y=1:Credit=Credit + AwardMY: End If
        If Y=0 And L16.State=1 And L11.State=1 And L15.State=1 then Y=1:Credit=Credit + AwardMY: End If
        If Y=0 And L7.State=1  And L2.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardMY: End If
        If Y=0 And L7.State=1  And L2.State=1  And L15.State=1 then Y=1:Credit=Credit + AwardMY: End If
        If Y=0 And L7.State=1  And L11.State=1 And L15.State=1 then Y=1:Credit=Credit + AwardMY: End If
        If Y=0 And L2.State=1  And L11.State=1 And L15.State=1 then Y=1:Credit=Credit + AwardMY: End If
        If Y=0 And L16.State=1 And L7.State=1  then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L16.State=1 And L2.State=1  then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L16.State=1 And L11.State=1 then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L16.State=1 And L15.State=1 then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L7.State=1  And L2.State=1  then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L7.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L7.State=1  And L15.State=1 then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L2.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L2.State=1  And L15.State=1 then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L11.State=1 And L15.State=1 then Y=1:Credit=Credit + AwardLY:End If
      End If
      'yellow stripes 5 choices:
      If F4.State=0 Then
        YY=0
        If L16.State=1 And L7.State=1  And L2.State=1  And L11.State=1 And L15.State=1 then YY=1:Credit=Credit + AwardHY:End If
        If YY=0 And L16.State=1 And L7.State=1  And L2.State=1  And L11.State=1 then YY=1:Credit=Credit + AwardMY:End If
        If YY=0 And L16.State=1 And L7.State=1  And L2.State=1  And L15.State=1 then YY=1:Credit=Credit + AwardMY:End If
        If YY=0 And L16.State=1 And L2.State=1  And L11.State=1 And L15.State=1 then YY=1:Credit=Credit + AwardMY:End If
        If YY=0 And L16.State=1 And L7.State=1  And L11.State=1 And L15.State=1 then YY=1:Credit=Credit + AwardMY:End If
        If YY=0 And L7.State=1  And L2.State=1  And L11.State=1 And L15.State=1 then YY=1:Credit=Credit + AwardMY:End If
        If YY=0 And L16.State=1 And L7.State=1  And L2.State=1  then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L16.State=1 And L7.State=1  And L11.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L16.State=1 And L7.State=1  And L15.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L16.State=1 And L2.State=1  And L11.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L16.State=1 And L2.State=1  And L15.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L16.State=1 And L11.State=1 And L15.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L7.State=1  And L2.State=1  And L11.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L7.State=1  And L2.State=1  And L15.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L7.State=1  And L11.State=1 And L15.State=1 then YY=1:Credit=Credit + AwardLY: End If
        If YY=0 And L2.State=1  And L11.State=1 And L15.State=1 then YY=1:Credit=Credit + AwardLY: End If
      End If
      YYY=0
      If L9.State=1 And L4.State=1 And L25.State=1 And L6.state=1 And L12.State=1 then YYY=1:Credit=Credit + AwardHY:End If
      If YYY=0 And L9.State=1  And L4.State=1  And L25.State=1 And L6.state=1  then YYY=1:Credit=Credit + AwardMY:End If
      If YYY=0 And L4.State=1  And L25.State=1 And L6.State=1  And L12.state=1 then YYY=1:Credit=Credit + AwardMY:End If
      If YYY=0 And L9.State=1  And L4.State=1  And L25.State=1 then YYY=1:Credit=Credit + AwardLY: End If
      If YYY=0 And L4.State=1  And L25.State=1 And L6.State=1  then YYY=1:Credit=Credit + AwardLY: End If
      If YYY=0 And L25.State=1 And L6.State=1  And L12.State=1 then YYY=1:Credit=Credit + AwardLY: End If

      G=0
      If L22.State=1 And L13.State=1 And L21.State=1 And L17.state=1 then G=1:Credit=Credit + AwardMG: End If
      If G=0 And L22.State=1 And L13.State=1 And L21.State=1 then Credit=Credit + AwardLG: End If
      If G=0 And L22.State=1 And L13.State=1 And L17.State=1 then Credit=Credit + AwardLG: End If
      If G=0 And L13.State=1 And L17.State=1 And L21.State=1 then Credit=Credit + AwardLG: End If
  End If

  If MX=6 Then 'Card 4
      'red super section feature 2 as 3, 3 as 4, 4 as 5
      If F3.State=1 Then
        R=0
        If R=0 And  L23.State=1 And L8.State=1  And L14.State=1 And L3.State=1  then R=1:Credit=Credit + AwardHR:End If
        If R=0 And  L23.State=1 And L8.State=1  And L14.State=1 And L21.State=1 then R=1:Credit=Credit + AwardHR:End If
        If R=0 And  L23.State=1 And L14.State=1 And L3.State=1  And L21.State=1 then R=1:Credit=Credit + AwardHR:End If
        If R=0 And  L23.State=1 And L8.State=1  And L3.State=1  And L21.State=1 then R=1:Credit=Credit + AwardHR:End If
        If R=0 And  L8.State=1  And L14.State=1 And L3.State=1  And L21.State=1 then R=1:Credit=Credit + AwardHR:End If
        If R=0 And L23.State=1 And L8.State=1 And L14.State=1 then R=1:Credit=Credit + AwardMR: End If
        If R=0 And L23.State=1 And L8.State=1 And L3.State=1  then R=1:Credit=Credit + AwardMR: End If
        If R=0 And L23.State=1 And L8.State=1 And L21.State=1 then R=1:Credit=Credit + AwardMR: End If
        If R=0 And L23.State=1 And L14.State=1  And L3.State=1  then R=1:Credit=Credit + AwardMR: End If
        If R=0 And L23.State=1 And L14.State=1  And L21.State=1 then R=1:Credit=Credit + AwardMR: End If
        If R=0 And L23.State=1 And L3.State=1 And L21.State=1 then R=1:Credit=Credit + AwardMR: End If
        If R=0 And L8.State=1  And L14.State=1  And L3.State=1  then R=1:Credit=Credit + AwardMR: End If
        If R=0 And L8.State=1  And L14.State=1  And L21.State=1 then R=1:Credit=Credit + AwardMR: End If
        If R=0 And L8.State=1  And L3.State=1 And L21.State=1 then R=1:Credit=Credit + AwardMR: End If
        If R=0 And L14.State=1 And L3.State=1 And L21.State=1 then R=1:Credit=Credit + AwardMR: End If
        If R=0 And  L23.State=1 And L8.State=1  then R=1:Credit=Credit+AwardLR:End If
        If R=0 And  L23.State=1 And L14.State=1 then R=1:Credit=Credit+AwardLR:End If
        If R=0 And  L23.State=1 And L3.State=1  then R=1:Credit=Credit+AwardLR:End If
        If R=0 And  L23.State=1 And L21.State=1 then R=1:Credit=Credit+AwardLR:End If
        If R=0 And  L8.State=1  And L14.State=1 then R=1:Credit=Credit+AwardLR:End If
        If R=0 And  L8.State=1  And L3.State=1  then R=1:Credit=Credit+AwardLR:End If
        If R=0 And  L8.State=1  And L21.State=1 then R=1:Credit=Credit+AwardLR:End If
        If R=0 And  L14.State=1 And L3.State=1  then R=1:Credit=Credit+AwardLR:End If
        If R=0 And  L14.State=1 And L21.State=1 then R=1:Credit=Credit+AwardLR:End If
        If R=0 And  L3.State=1  And L21.State=1 then R=1:Credit=Credit+AwardLR:End If
      End If

      'Red Stripes 5 choices:
      If F3.state=0 Then
        RR=0
        If RR=0 And L23.State=1 And L8.State=1  And L14.State=1 And L3.State=1  And L21.State=1 then RR=1:Credit=Credit + AwardHR:End If
        If RR=0 And L23.State=1 And L8.State=1  And L14.State=1 And L3.State=1  then RR=1:Credit=Credit + AwardMR:End If
        If RR=0 And L23.State=1 And L8.State=1  And L14.State=1 And L21.State=1 then RR=1:Credit=Credit + AwardMR:End If
        If RR=0 And L23.State=1 And L8.State=1  And L3.State=1  And L21.State=1 then RR=1:Credit=Credit + AwardMR:End If
        If RR=0 And L23.State=1 And L14.State=1 And L3.State=1  And L21.State=1 then RR=1:Credit=Credit + AwardMR:End If
        If RR=0 And L8.State=1  And L14.State=1 And L3.State=1  And L21.State=1 then RR=1:Credit=Credit + AwardMR:End If
        If RR=0 And L23.State=1 And L8.State=1  And L14.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L23.State=1 And L8.State=1  And L3.State=1  then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L23.State=1 And L8.State=1  And L21.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L23.State=1 And L14.State=1 And L3.State=1  then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L23.State=1 And L14.State=1 And L21.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L23.State=1 And L3.State=1  And L21.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L8.State=1  And L14.State=1 And L3.State=1  then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L8.State=1  And L14.State=1 And L21.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L8.State=1  And L3.State=1  And L21.State=1 then RR=1:Credit=Credit + AwardLR: End If
        If RR=0 And L14.State=1 And L3.State=1  And L21.State=1 then RR=1:Credit=Credit + AwardLR: End If
      End If

            'Red Row
      RRR=0
      If L9.State=1  And L4.State=1  And L25.State=1 And L6.state=1  And L12.State=1 then RRR=1:Credit=Credit + AwardHR:End If
      If RRR=0 And L9.State=1  And L4.State=1  And L25.State=1 And L6.state=1  then RRR=1:Credit=Credit + AwardMR:End If
      If RRR=0 And L4.State=1  And L25.State=1 And L6.State=1  And L12.state=1 then RRR=1:Credit=Credit + AwardMR:End If
        If RRR=0 And L9.State=1  And L4.State=1  And L25.State=1 then RRR=1:Credit=Credit+ AwardLR: End If
      If RRR=0 And L4.State=1  And L25.State=1 And L6.State=1  then RRR=1:Credit=Credit+ AwardLR: End If
      If RRR=0 And L25.State=1 And L6.State=1  And L12.State=1 then RRR=1:Credit=Credit+ AwardLR: End If
        If RRR=0 And L22.State=1 And L18.State=1 And L15.State=1 then RRR=1:Credit=Credit+ AwardLR: End If

      'yellow super section feature 2 as 3, 3 as 4, 4 as 5
      If F4.State=1  Then
        Y=0
        If Y=0 And L24.State=1  And L19.State=1 And L1.State=1  And L2.State=1  then Y=1:Credit=Credit + AwardHY:End If
        If Y=0 And L24.State=1  And L19.State=1 And L1.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardHY:End If
        If Y=0 And L24.State=1  And L1.State=1  And L2.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardHY:End If
        If Y=0 And L24.State=1  And L19.State=1 And L2.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardHY:End If
        If Y=0 And L19.State=1  And L1.State=1  And L2.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardHY:End If
        If Y=0 And L24.State=1  And L19.State=1 And L1.State=1  then Y=1:Credit = Credit +AwardMY: End If
        If Y=0 And L24.State=1  And L19.State=1 And L2.State=1  then Y=1:Credit = Credit +AwardMY: End If
        If Y=0 And L24.State=1  And L19.State=1 And L11.State=1 then Y=1:Credit = Credit +AwardMY: End If
        If Y=0 And L24.State=1  And L1.State=1  And L2.State=1  then Y=1:Credit = Credit +AwardMY: End If
        If Y=0 And L24.State=1  And L1.State=1  And L11.State=1 then Y=1:Credit = Credit +AwardMY: End If
        If Y=0 And L24.State=1  And L2.State=1  And L11.State=1 then Y=1:Credit = Credit +AwardMY: End If
        If Y=0 And L19.State=1  And L1.State=1  And L2.State=1  then Y=1:Credit = Credit +AwardMY: End If
        If Y=0 And L19.State=1  And L1.State=1  And L11.State=1 then Y=1:Credit = Credit +AwardMY: End If
        If Y=0 And L19.State=1  And L2.State=1  And L11.State=1 then Y=1:Credit = Credit +AwardMY: End If
        If Y=0 And L1.State=1 And L2.State=1  And L11.State=1 then Y=1:Credit = Credit +AwardMY: End If
        If Y=0 And L24.State=1  And L19.State=1 then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L24.State=1  And L1.State=1  then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L24.State=1  And L2.State=1  then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L24.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L19.State=1  And L1.State=1  then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L19.State=1  And L2.State=1  then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L19.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L1.State=1 And L2.State=1  then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L1.State=1 And L11.State=1 then Y=1:Credit=Credit + AwardLY:End If
        If Y=0 And L2.State=1 And L11.State=1 then Y=1:Credit=Credit + AwardLY:End If
      End If
        '5 block yellow stripe
      If F4.State=0 Then
        Y=0
        If Y=0 And L24.State=1  And L19.State=1 And L1.State=1  And L2.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardHY:End If
        If Y=0 And L24.State=1  And L19.State=1 And L1.State=1  And L2.State=1  then Y=1:Credit=Credit + AwardMY:End If
        If Y=0 And L24.State=1  And L19.State=1 And L1.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardMY:End If
        If Y=0 And L24.State=1  And L19.State=1 And L2.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardMY:End If
        If Y=0 And L24.State=1  And L1.State=1  And L2.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardMY:End If
        If Y=0 And L19.State=1  And L1.State=1  And L2.State=1  And L11.State=1 then Y=1:Credit=Credit + AwardMY:End If
        If Y=0 And L24.State=1  And L19.State=1 And L1.State=1  then Y=1:Credit = Credit +AwardLY: End If
        If Y=0 And L24.State=1  And L19.State=1 And L2.State=1  then Y=1:Credit = Credit +AwardLY: End If
        If Y=0 And L24.State=1  And L19.State=1 And L11.State=1 then Y=1:Credit = Credit +AwardLY: End If
        If Y=0 And L24.State=1  And L1.State=1  And L2.State=1  then Y=1:Credit = Credit +AwardLY: End If
        If Y=0 And L24.State=1  And L1.State=1  And L11.State=1 then Y=1:Credit = Credit +AwardLY: End If
        If Y=0 And L24.State=1  And L2.State=1  And L11.State=1 then Y=1:Credit = Credit +AwardLY: End If
        If Y=0 And L19.State=1  And L1.State=1  And L2.State=1  then Y=1:Credit = Credit +AwardLY: End If
        If Y=0 And L19.State=1  And L1.State=1  And L11.State=1 then Y=1:Credit = Credit +AwardLY: End If
        If Y=0 And L19.State=1  And L2.State=1  And L11.State=1 then Y=1:Credit = Credit +AwardLY: End If
        If Y=0 And L1.State=1 And L2.State=1  And L11.State=1 then Y=1:Credit = Credit +AwardLY: End If
      End If

      '6 green blocks
      G=0
      If G=0 And L7.State=1 And L16.State=1 And L5.State=1  And L13.State=1 And L17.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L7.State=1 And L5.State=1  And L13.State=1 And L17.State=1 And L20.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L7.State=1 And L16.State=1 And L5.State=1  And L17.State=1 And L20.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L7.State=1 And L16.State=1 And L5.State=1  And L13.State=1 And L20.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L16.State=1  And L5.State=1  And L13.State=1 And L17.State=1 And L20.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L7.State=1 And L16.State=1 And L5.State=1  And L13.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L7.State=1 And L16.State=1 And L5.State=1  And L17.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L7.State=1 And L16.State=1 And L5.State=1  And L20.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L7.State=1 And L16.State=1 And L13.State=1 And L17.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L7.State=1 And L16.State=1 And L13.State=1 And L20.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L7.State=1 And L5.State=1  And L13.State=1 And L17.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L7.State=1 And L5.State=1  And L13.State=1 And L20.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L7.State=1 And L5.State=1  And L17.State=1 And L20.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L7.State=1 And L13.State=1 And L17.State=1 And L20.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L16.State=1  And L5.State=1  And L13.State=1 And L17.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L16.State=1  And L5.State=1  And L13.State=1 And L20.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L16.State=1  And L13.State=1 And L17.State=1 And L20.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L5.State=1 And L13.State=1 And L17.State=1 And L20.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L7.State=1 And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L7.State=1 And L16.State=1 And L13.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L7.State=1 And L16.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L7.State=1 And L16.State=1 And L20.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L7.State=1 And L5.State=1  And L13.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L7.State=1 And L5.State=1  And L17.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L7.State=1 And L5.State=1  And L20.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L7.State=1 And L13.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L7.State=1 And L13.State=1 And L20.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L7.State=1 And L17.State=1 And L20.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L16.State=1  And L5.State=1  And L13.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L16.State=1  And L5.State=1  And L17.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L16.State=1  And L5.State=1  And L20.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L16.State=1  And L13.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L16.State=1  And L13.State=1 And L20.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L16.State=1  And L17.State=1 And L20.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L5.State=1 And L13.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L5.State=1 And L13.State=1 And L20.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L5.State=1 And L17.State=1 And L20.State=1 then G=1:Credit = Credit + AwardLG: End If
  End If

  If MX=7 Then 'Card 5
      'Red Stripe super section feature 2 as 3, 3 as 4, 4 as 5
      If F3.State=1 Then
        R=0
        If R=0 And L6.State=1 And L12.State=1 And L8.State=1  And L14.State=1 then R=1:Credit = Credit + AwardHR:End  If
        If R=0 And L6.State=1 And L12.State=1 And L8.State=1  And L5.State=1  then R=1:Credit = Credit + AwardHR:End  If
        If R=0 And L6.State=1 And L8.State=1  And L14.State=1 And L5.State=1  then R=1:Credit = Credit + AwardHR:End  If
        If R=0 And L6.State=1 And L12.State=1 And L14.State=1 And L5.State=1  then R=1:Credit = Credit + AwardHR:End  If
        If R=0 And L12.State=1  And L8.State=1  And L14.State=1 And L5.State=1  then R=1:Credit = Credit + AwardHR:End  If
        If R=0 And L6.State=1 And L12.State=1 And L8.State=1  then R=1:Credit = Credit + AwardMR: End If
        If R=0 And L6.State=1 And L12.State=1 And L14.State=1 then R=1:Credit = Credit + AwardMR: End If
        If R=0 And L6.State=1 And L12.State=1 And L5.State=1  then R=1:Credit = Credit + AwardMR: End If
        If R=0 And L6.State=1 And L8.State=1  And L14.State=1 then R=1:Credit = Credit + AwardMR: End If
        If R=0 And L6.State=1 And L8.State=1  And L5.State=1  then R=1:Credit = Credit + AwardMR: End If
        If R=0 And L6.State=1 And L14.State=1 And L5.State=1  then R=1:Credit = Credit + AwardMR: End If
        If R=0 And L12.State=1  And L8.State=1  And L14.State=1 then R=1:Credit = Credit + AwardMR: End If
        If R=0 And L12.State=1  And L8.State=1  And L5.State=1  then R=1:Credit = Credit + AwardMR: End If
        If R=0 And L12.State=1  And L14.State=1 And L5.State=1  then R=1:Credit = Credit + AwardMR: End If
        If R=0 And L8.State=1 And L14.State=1 And L5.State=1  then R=1:Credit = Credit + AwardMR: End If
        If R=0 And L6.State=1 And L12.State=1 then  R=1:Credit = Credit + AwardLR:End If
        If R=0 And L6.State=1 And L8.State=1  then  R=1:Credit = Credit + AwardLR:End If
        If R=0 And L6.State=1 And L14.State=1 then  R=1:Credit = Credit + AwardLR:End If
        If R=0 And L6.State=1 And L5.State=1  then  R=1:Credit = Credit + AwardLR:End If
        If R=0 And L12.State=1  And L8.State=1  then  R=1:Credit = Credit + AwardLR:End If
        If R=0 And L12.State=1  And L14.State=1 then  R=1:Credit = Credit + AwardLR:End If
        If R=0 And L12.State=1  And L5.State=1  then  R=1:Credit = Credit + AwardLR:End If
        If R=0 And L8.State=1 And L14.State=1 then  R=1:Credit = Credit + AwardLR:End If
        If R=0 And L8.State=1 And L5.State=1  then  R=1:Credit = Credit + AwardLR:End If
        If R=0 And L14.State=1  And L5.State=1  then  R=1:Credit = Credit + AwardLR:End If
      End If
             'Red Stripes 5 blocks
      If F3.State=0 Then
        R=0
        If R=0 And L6.State=1 And L12.State=1 And L8.State=1  And L14.State=1 And L5.State=1 then R=1:Credit = Credit + AwardHR:End If
        If R=0 And L6.State=1 And L12.State=1 And L8.State=1  And L14.State=1 then R=1:Credit = Credit + AwardMR::End If
        If R=0 And L6.State=1 And L12.State=1 And L8.State=1  And L5.State=1  then R=1:Credit = Credit + AwardMR::End If
        If R=0 And L6.State=1 And L12.State=1 And L14.State=1 And L5.State=1  then R=1:Credit = Credit + AwardMR::End If
        If R=0 And L6.State=1 And L8.State=1  And L14.State=1 And L5.State=1  then R=1:Credit = Credit + AwardMR::End If
        If R=0 And L12.State=1  And L8.State=1  And L14.State=1 And L5.State=1  then R=1:Credit = Credit + AwardMR::End If
        If R=0 And L6.State=1 And L12.State=1 And L8.State=1  then R=1:Credit = Credit + AwardLR: End If
        If R=0 And L6.State=1 And L12.State=1 And L14.State=1 then R=1:Credit = Credit + AwardLR: End If
        If R=0 And L6.State=1 And L12.State=1 And L5.State=1  then R=1:Credit = Credit + AwardLR: End If
        If R=0 And L6.State=1 And L8.State=1  And L14.State=1 then R=1:Credit = Credit + AwardLR: End If
        If R=0 And L6.State=1 And L8.State=1  And L5.State=1  then R=1:Credit = Credit + AwardLR: End If
        If R=0 And L6.State=1 And L14.State=1 And L5.State=1  then R=1:Credit = Credit + AwardLR: End If
        If R=0 And L12.State=1  And L8.State=1  And L14.State=1 then R=1:Credit = Credit + AwardLR: End If
        If R=0 And L12.State=1  And L8.State=1  And L5.State=1  then R=1:Credit = Credit + AwardLR: End If
        If R=0 And L12.State=1  And L14.State=1 And L5.State=1  then R=1:Credit = Credit + AwardLR: End If
        If R=0 And L8.State=1 And L14.State=1 And L5.State=1  then R=1:Credit = Credit + AwardLR: End If
      End If

      'Red 5 blocks
      RR=0
      If RR=0 And L7.State=1  And L11.State=1 And L22.State=1 And L15.State=1 And L18.State=1 then RR=1:Credit = Credit + AwardHR:End If
      If RR=0 And L7.State=1  And L11.State=1 And L22.State=1 And L15.State=1 then RR=1:Credit = Credit + AwardMR:End If
      If RR=0 And L7.State=1  And L11.State=1 And L22.State=1 And L18.State=1 then RR=1:Credit = Credit + AwardMR:End If
      If RR=0 And L7.State=1  And L11.State=1 And L15.State=1 And L18.State=1 then RR=1:Credit = Credit + AwardMR:End If
      If RR=0 And L7.State=1  And L22.State=1 And L15.State=1 And L18.State=1 then RR=1:Credit = Credit + AwardMR:End If
      If RR=0 And L11.State=1 And L22.State=1 And L15.State=1 And L18.State=1 then RR=1:Credit = Credit + AwardMR:End If
      If RR=0 And L7.State=1  And L11.State=1 And L22.State=1 then RR=1:Credit = Credit + AwardLR: End If
      If RR=0 And L7.State=1  And L11.State=1 And L15.State=1 then RR=1:Credit = Credit + AwardLR: End If
      If RR=0 And L7.State=1  And L11.State=1 And L18.State=1 then RR=1:Credit = Credit + AwardLR: End If
      If RR=0 And L7.State=1  And L22.State=1 And L15.State=1 then RR=1:Credit = Credit + AwardLR: End If
      If RR=0 And L7.State=1  And L22.State=1 And L18.State=1 then RR=1:Credit = Credit + AwardLR: End If
      If RR=0 And L7.State=1  And L15.State=1 And L18.State=1 then RR=1:Credit = Credit + AwardLR: End If
      If RR=0 And L11.State=1 And L22.State=1 And L15.State=1 then RR=1:Credit = Credit + AwardLR: End If
      If RR=0 And L11.State=1 And L22.State=1 And L18.State=1 then RR=1:Credit = Credit + AwardLR: End If
      If RR=0 And L11.State=1 And L15.State=1 And L18.State=1 then RR=1:Credit = Credit + AwardLR: End If
      If RR=0 And L22.State=1 And L15.State=1 And L18.State=1 then RR=1:Credit = Credit + AwardLR: End If

            'yellow super section feature 2 as 3, 3 as 4, 4 as 5
      If F4.State=1  Then
        Y=0
        If Y=0 And L25.State=1  And L4.State=1  And L9.State=1  And L1.State=1  then  Y=1:Credit = Credit + AwardHY:End If
        If Y=0 And L25.State=1  And L4.State=1  And L9.State=1  And L2.State=1  then  Y=1:Credit = Credit + AwardHY:End If
        If Y=0 And L25.State=1  And L9.State=1  And L1.State=1  And L2.State=1  then  Y=1:Credit = Credit + AwardHY:End If
        If Y=0 And L25.State=1  And L4.State=1  And L1.State=1  And L2.State=1  then  Y=1:Credit = Credit + AwardHY:End If
        If Y=0 And L4.State=1 And L9.State=1  And L1.State=1  And L2.State=1  then  Y=1:Credit = Credit + AwardHY:End If
        If Y=0 And L25.State=1  And L4.State=1  And L9.State=1  then Y=1:Credit = Credit + AwardMY: End If
        If Y=0 And L25.State=1  And L4.State=1  And L1.State=1  then Y=1:Credit = Credit + AwardMY: End If
        If Y=0 And L25.State=1  And L4.State=1  And L2.State=1  then Y=1:Credit = Credit + AwardMY: End If
        If Y=0 And L25.State=1  And L9.State=1  And L1.State=1  then Y=1:Credit = Credit + AwardMY: End If
        If Y=0 And L25.State=1  And L9.State=1  And L2.State=1  then Y=1:Credit = Credit + AwardMY: End If
        If Y=0 And L25.State=1  And L1.State=1  And L2.State=1  then Y=1:Credit = Credit + AwardMY: End If
        If Y=0 And L4.State=1 And L9.State=1  And L1.State=1  then Y=1:Credit = Credit + AwardMY: End If
        If Y=0 And L4.State=1 And L9.State=1  And L2.State=1  then Y=1:Credit = Credit + AwardMY: End If
        If Y=0 And L4.State=1 And L1.State=1  And L2.State=1  then Y=1:Credit = Credit + AwardMY: End If
        If Y=0 And L9.State=1 And L1.State=1  And L2.State=1  then Y=1:Credit = Credit + AwardMY: End If
        If Y=0 And L25.State=1  And L4.State=1  then  Y=1:Credit = Credit + AwardLY:End If
        If Y=0 And L25.State=1  And L9.State=1  then  Y=1:Credit = Credit + AwardLY:End If
        If Y=0 And L25.State=1  And L1.State=1  then  Y=1:Credit = Credit + AwardLY:End If
        If Y=0 And L25.State=1  And L2.State=1  then  Y=1:Credit = Credit + AwardLY:End If
        If Y=0 And L4.State=1 And L9.State=1  then  Y=1:Credit = Credit + AwardLY:End If
        If Y=0 And L4.State=1 And L1.State=1  then  Y=1:Credit = Credit + AwardLY:End If
        If Y=0 And L4.State=1 And L2.State=1  then  Y=1:Credit = Credit + AwardLY:End If
        If Y=0 And L9.State=1 And L1.State=1  then  Y=1:Credit = Credit + AwardLY:End If
        If Y=0 And L9.State=1 And L2.State=1  then  Y=1:Credit = Credit + AwardLY:End If
        If Y=0 And L1.State=1 And L2.State=1  then  Y=1:Credit = Credit + AwardLY:End If
      End If
      'yellow stripes 5 blocks
      If F4.State=0 Then
        Y=0
        If Y=0 And L25.State=1  And L4.State=1  And L9.State=1  And L1.State=1 And  L2.State=1 then Y=1:Credit = Credit + AwardHY:End If
        If Y=0 And L25.State=1  And L4.State=1  And L9.State=1  And L1.State=1 then Y=1:Credit = Credit + AwardMY:End If
        If Y=0 And L25.State=1  And L4.State=1  And L9.State=1  And L2.State=1 then Y=1:Credit = Credit + AwardMY:End If
        If Y=0 And L25.State=1  And L4.State=1  And L1.State=1  And L2.State=1 then Y=1:Credit = Credit + AwardMY:End If
        If Y=0 And L25.State=1  And L9.State=1  And L1.State=1  And L2.State=1 then Y=1:Credit = Credit + AwardMY:End If
        If Y=0 And L4.State=1 And L9.State=1  And L1.State=1  And L2.State=1 then Y=1:Credit = Credit + AwardMY:End If
        If Y=0 And L25.State=1  And L4.State=1  And L9.State=1  then Y=1:Credit = Credit + AwardLY: End If
        If Y=0 And L25.State=1  And L4.State=1  And L1.State=1  then Y=1:Credit = Credit + AwardLY: End If
        If Y=0 And L25.State=1  And L4.State=1  And L2.State=1  then Y=1:Credit = Credit + AwardLY: End If
        If Y=0 And L25.State=1  And L9.State=1  And L1.State=1  then Y=1:Credit = Credit + AwardLY: End If
        If Y=0 And L25.State=1  And L9.State=1  And L2.State=1  then Y=1:Credit = Credit + AwardLY: End If
        If Y=0 And L25.State=1  And L1.State=1  And L2.State=1  then Y=1:Credit = Credit + AwardLY: End If
        If Y=0 And L4.State=1 And L9.State=1  And L1.State=1  then Y=1:Credit = Credit + AwardLY: End If
        If Y=0 And L4.State=1 And L9.State=1  And L2.State=1  then Y=1:Credit = Credit + AwardLY: End If
        If Y=0 And L4.State=1 And L1.State=1  And L2.State=1  then Y=1:Credit = Credit + AwardLY: End If
        If Y=0 And L9.State=1 And L1.State=1  And L2.State=1  then Y=1:Credit = Credit + AwardLY: End If
      End If
            'Yellow 3 blocks
      If L3.State=1 And L10.State=1 And L20.State=1then:Credit = Credit + AwardLR: End If
            'Green 6 blocks
      G=0
      If G=0 And L19.State=1  And L24.State=1 And L23.State=1 And L16.State=1 And L13.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L19.State=1  And L23.State=1 And L16.State=1 And L13.State=1 And L21.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L19.State=1  And L24.State=1 And L23.State=1 And L13.State=1 And L21.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L19.State=1  And L24.State=1 And L23.State=1 And L16.State=1 And L21.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L24.State=1  And L23.State=1 And L16.State=1 And L13.State=1 And L21.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L19.State=1  And L24.State=1 And L23.State=1 And L16.State=1 then G=1:Credit = Credit + AwardMG:End  If
      If G=0 And L19.State=1  And L24.State=1 And L23.State=1 And L13.State=1 then G=1:Credit = Credit + AwardMG:End  If
      If G=0 And L19.State=1  And L24.State=1 And L23.State=1 And L21.State=1 then G=1:Credit = Credit + AwardMG:End  If
      If G=0 And L19.State=1  And L24.State=1 And L16.State=1 And L13.State=1 then G=1:Credit = Credit + AwardMG:End  If
      If G=0 And L19.State=1  And L24.State=1 And L16.State=1 And L21.State=1 then G=1:Credit = Credit + AwardMG:End  If
      If G=0 And L19.State=1  And L23.State=1 And L16.State=1 And L13.State=1 then G=1:Credit = Credit + AwardMG:End  If
      If G=0 And L19.State=1  And L23.State=1 And L16.State=1 And L21.State=1 then G=1:Credit = Credit + AwardMG:End  If
      If G=0 And L19.State=1  And L23.State=1 And L13.State=1 And L21.State=1 then G=1:Credit = Credit + AwardMG:End  If
      If G=0 And L19.State=1  And L16.State=1 And L13.State=1 And L21.State=1 then G=1:Credit = Credit + AwardMG:End  If
      If G=0 And L24.State=1  And L23.State=1 And L16.State=1 And L13.State=1 then G=1:Credit = Credit + AwardMG:End  If
      If G=0 And L24.State=1  And L23.State=1 And L16.State=1 And L21.State=1 then G=1:Credit = Credit + AwardMG:End  If
      If G=0 And L24.State=1  And L16.State=1 And L13.State=1 And L21.State=1 then G=1:Credit = Credit + AwardMG:End  If
      If G=0 And L23.State=1  And L16.State=1 And L13.State=1 And L21.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L19.State=1  And L24.State=1 And L23.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L19.State=1  And L24.State=1 And L16.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L19.State=1  And L24.State=1 And L13.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L19.State=1  And L24.State=1 And L21.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L19.State=1  And L23.State=1 And L16.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L19.State=1  And L23.State=1 And L13.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L19.State=1  And L23.State=1 And L21.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L19.State=1  And L16.State=1 And L13.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L19.State=1  And L16.State=1 And L21.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L19.State=1  And L13.State=1 And L21.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L24.State=1  And L23.State=1 And L16.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L24.State=1  And L23.State=1 And L13.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L24.State=1  And L23.State=1 And L21.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L24.State=1  And L16.State=1 And L13.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L24.State=1  And L16.State=1 And L21.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L24.State=1  And L13.State=1 And L21.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L23.State=1  And L16.State=1 And L13.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L23.State=1  And L16.State=1 And L21.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L23.State=1  And L13.State=1 And L21.State=1 then G=1:Credit = Credit + AwardLG: End If
  End If

  If MX=8 Then 'Card 6
      'Red 5 Blocks
      R=0
      If R=0 And L2.State=1 And L11.State=1 And L19.State=1 And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardHR:End If
      If R=0 And L2.State=1 And L19.State=1 And L7.State=1  And L22.State=1 And L18.State=1 then R=1:Credit = Credit + AwardHR:End If
      If R=0 And L2.State=1 And L11.State=1 And L19.State=1 And L22.State=1 And L18.State=1 then R=1:Credit = Credit + AwardHR:End If
      If R=0 And L2.State=1 And L11.State=1 And L19.State=1 And L7.State=1  And L18.State=1 then R=1:Credit = Credit + AwardHR:End If
      If R=0 And L11.State=1  And L19.State=1 And L7.State=1  And L22.State=1 And L18.State=1 then R=1:Credit = Credit + AwardHR:End If
      If R=0 And L2.State=1 And L11.State=1 And L19.State=1 And L7.State=1  then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L2.State=1 And L11.State=1 And L19.State=1 And L22.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L2.State=1 And L11.State=1 And L19.State=1 And L18.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L2.State=1 And L11.State=1 And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L2.State=1 And L11.State=1 And L7.State=1  And L18.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L2.State=1 And L19.State=1 And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L2.State=1 And L19.State=1 And L7.State=1  And L18.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L2.State=1 And L19.State=1 And L22.State=1 And L18.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L2.State=1 And L7.State=1  And L22.State=1 And L18.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L11.State=1  And L19.State=1 And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L11.State=1  And L19.State=1 And L7.State=1  And L18.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L11.State=1  And L7.State=1  And L22.State=1 And L18.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L19.State=1  And L7.State=1  And L22.State=1 And L18.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L2.State=1 And L11.State=1 And L19.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L2.State=1 And L11.State=1 And L7.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L2.State=1 And L11.State=1 And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L2.State=1 And L11.State=1 And L18.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L2.State=1 And L19.State=1 And L7.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L2.State=1 And L19.State=1 And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L2.State=1 And L19.State=1 And L18.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L2.State=1 And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L2.State=1 And L7.State=1  And L18.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L2.State=1 And L22.State=1 And L18.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L11.State=1  And L19.State=1 And L7.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L11.State=1  And L19.State=1 And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L11.State=1  And L19.State=1 And L18.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L11.State=1  And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L11.State=1  And L7.State=1  And L18.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L11.State=1  And L22.State=1 And L18.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L19.State=1  And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L19.State=1  And L7.State=1  And L18.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L19.State=1  And L22.State=1 And L18.State=1 then R=1:Credit = Credit + AwardLR: End If

      'red super section feature 2 as 3, 3 as 4
      RR=0
      If F3.State=1 Then
        If  L12.State=1 And L23.State=1 And L8.State=1  then RR=1:Credit=Credit + AwardMR:End If
        If RR=0 And L12.State=1 And L23.State=1 then Credit=Credit + AwardLR:End If
        If RR=0 And L12.State=1 And L8.State=1  then Credit=Credit + AwardLR:End If
        If RR=0 And L23.State=1 And L8.State=1  then Credit=Credit + AwardLR:End If
      End If

      'Red Stripes 3 Blocks
      If F3.State=0 Then
        If  L12.State=1 And L8.State=1  And L23.State=1 then Credit=Credit + AwardLR: End If
      End If

            'Yellow superstate 2 as 3
      If F4.State=1  Then
        If  L9.State=1  And L1.State=1  then Credit=Credit  + AwardLY:End If
      End If

      'Yellow 4 Blocks
      Y=0
      If  L14.State=1 And L3.State=1  And L21.State=1 And L10.State=1 then Y=1:Credit=Credit + AwardMY:End If
      If Y=0 And L14.State=1 And  L3.State=1  And L21.State=1 then Y=1:Credit=Credit + AwardLY: End If
      If Y=0 And L14.State=1 And  L3.State=1  And L10.State=1 then Y=1:Credit=Credit + AwardLY: End If
      If Y=0 And L14.State=1 And  L21.State=1 And L10.State=1 then Y=1:Credit=Credit + AwardLY: End If
      If Y=0 And L3.State=1  And  L21.State=1 And L10.State=1 then Y=1:Credit=Credit + AwardLY: End If

      'Blue feature F1: 3 Blue = 5 Green   F2: 2 Blue = 5 Green
      BBB=0
      If F1.State=1  Then
        If  L13.State=1 And L17.State=1 And L20.State=1 then BBB=1:Credit=Credit + AwardHG: End If
      End If

      If F2.State=1 Then
        If BBB=0 And L13.State=1 And L17.State=1 then Credit=Credit + AwardHG:End If
        If BBB=0 And L13.State=1 And L20.State=1 then Credit=Credit + AwardHG:End If
        If BBB=0 And L17.State=1 And L20.State=1 then Credit=Credit + AwardHG:End If
      End If

      'Green 7block
      G=0
      If G=0 And L4.State=1 And L25.State=1 And L6.State=1  And L24.State=1 And L16.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L25.State=1 And L6.State=1  And L24.State=1 And L5.State=1  then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L25.State=1 And L6.State=1  And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L25.State=1 And L6.State=1  And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L25.State=1 And L6.State=1  And L16.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L25.State=1 And L16.State=1 And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L6.State=1  And L24.State=1 And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L6.State=1  And L24.State=1 And L16.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L6.State=1  And L24.State=1 And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L25.State=1 And L24.State=1 And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L25.State=1 And L24.State=1 And L16.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L25.State=1 And L24.State=1 And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L6.State=1  And L16.State=1 And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L25.State=1  And L6.State=1  And L24.State=1 And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L25.State=1  And L6.State=1  And L24.State=1 And L16.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L25.State=1  And L6.State=1  And L24.State=1 And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L25.State=1  And L24.State=1 And L16.State=1 And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L25.State=1  And L6.State=1  And L16.State=1 And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L6.State=1 And L24.State=1 And L16.State=1 And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L4.State=1 And L25.State=1 And L6.State=1  And L24.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L25.State=1 And L6.State=1  And L16.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L25.State=1 And L6.State=1  And L5.State=1  then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L25.State=1 And L6.State=1  And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L25.State=1 And L24.State=1 And L16.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L25.State=1 And L24.State=1 And L5.State=1  then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L25.State=1 And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L6.State=1  And L24.State=1 And L16.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L6.State=1  And L24.State=1 And L5.State=1  then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L6.State=1  And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L6.State=1  And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L6.State=1  And L16.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L24.State=1 And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L24.State=1 And L16.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L25.State=1  And L6.State=1  And L24.State=1 And L16.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L25.State=1  And L6.State=1  And L24.State=1 And L5.State=1  then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L25.State=1  And L6.State=1  And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L25.State=1  And L24.State=1 And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L25.State=1  And L24.State=1 And L16.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L6.State=1 And L24.State=1 And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L6.State=1 And L24.State=1 And L16.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L6.State=1 And L16.State=1 And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L4.State=1 And L25.State=1 And L6.State=1  then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L25.State=1 And L24.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L25.State=1 And L16.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L25.State=1 And L5.State=1  then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L25.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L6.State=1  And L24.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L6.State=1  And L16.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L6.State=1  And L5.State=1  then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L6.State=1  And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L24.State=1 And L16.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L24.State=1 And L5.State=1  then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L16.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L4.State=1 And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L25.State=1  And L6.State=1  And L24.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L25.State=1  And L6.State=1  And L16.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L25.State=1  And L6.State=1  And L5.State=1  then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L25.State=1  And L6.State=1  And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L25.State=1  And L24.State=1 And L16.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L25.State=1  And L24.State=1 And L5.State=1  then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L25.State=1  And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L25.State=1  And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L25.State=1  And L16.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L25.State=1  And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L6.State=1 And L24.State=1 And L16.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L6.State=1 And L24.State=1 And L5.State=1  then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L6.State=1 And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L6.State=1 And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L6.State=1 And L16.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L6.State=1 And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L24.State=1  And L16.State=1 And L5.State=1  then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L24.State=1  And L16.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L24.State=1  And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
      If G=0 And L16.State=1  And L5.State=1  And L15.State=1 then G=1:Credit = Credit + AwardLG: End If
  End If

  If MX=9 Then 'Final Card
      'Red 6 Blocks
      R=0
      If R=0 And L4.State=1 And L19.State=1 And L1.State=1  And L2.State=1  And L7.State=1  then R=1:Credit = Credit + AwardHR:End If
      If R=0 And L4.State=1 And L1.State=1  And L2.State=1  And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardHR:End If
      If R=0 And L4.State=1 And L19.State=1 And L1.State=1  And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardHR:End If
      If R=0 And L4.State=1 And L19.State=1 And L1.State=1  And L2.State=1  And L22.State=1 then R=1:Credit = Credit + AwardHR:End If
      If R=0 And L19.State=1  And L1.State=1  And L2.State=1  And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardHR:End If
      If R=0 And L4.State=1 And L19.State=1 And L1.State=1  And L2.State=1  then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L4.State=1 And L19.State=1 And L1.State=1  And L7.State=1  then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L4.State=1 And L19.State=1 And L1.State=1  And L22.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L4.State=1 And L19.State=1 And L2.State=1  And L7.State=1  then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L4.State=1 And L19.State=1 And L2.State=1  And L22.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L4.State=1 And L1.State=1  And L2.State=1  And L7.State=1  then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L4.State=1 And L1.State=1  And L2.State=1  And L22.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L4.State=1 And L1.State=1  And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L4.State=1 And L2.State=1  And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L19.State=1  And L1.State=1  And L2.State=1  And L7.State=1  then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L19.State=1  And L1.State=1  And L2.State=1  And L22.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L19.State=1  And L2.State=1  And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L1.State=1 And L2.State=1  And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardMR:End If
      If R=0 And L4.State=1 And L19.State=1 And L1.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L4.State=1 And L19.State=1 And L2.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L4.State=1 And L19.State=1 And L7.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L4.State=1 And L19.State=1 And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L4.State=1 And L1.State=1  And L2.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L4.State=1 And L1.State=1  And L7.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L4.State=1 And L1.State=1  And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L4.State=1 And L2.State=1  And L7.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L4.State=1 And L2.State=1  And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L4.State=1 And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L19.State=1  And L1.State=1  And L2.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L19.State=1  And L1.State=1  And L7.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L19.State=1  And L1.State=1  And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L19.State=1  And L2.State=1  And L7.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L19.State=1  And L2.State=1  And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L19.State=1  And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L1.State=1 And L2.State=1  And L7.State=1  then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L1.State=1 And L2.State=1  And L22.State=1 then R=1:Credit = Credit + AwardLR: End If
      If R=0 And L1.State=1 And L7.State=1  And L22.State=1 then R=1:Credit = Credit + AwardLR: End If

      'red super section feature 2 as 3
      If F3.State=1 Then
        If  L6.State=1  And L12.State=1 then Credit = Credit + AwardLR:End  If
      End If

      'Yellow 6 Blocks
      Y=0
      If Y=0 And L8.State=1 And L14.State=1 And L5.State=1  And L3.State=1  And L10.State=1 then Y=1:Credit = Credit + AwardHY:End If
      If Y=0 And L8.State=1 And L5.State=1  And L3.State=1  And L10.State=1 And L20.State=1 then Y=1:Credit = Credit + AwardHY:End If
      If Y=0 And L8.State=1 And L14.State=1 And L5.State=1  And L10.State=1 And L20.State=1 then Y=1:Credit = Credit + AwardHY:End If
      If Y=0 And L8.State=1 And L14.State=1 And L5.State=1  And L3.State=1  And L20.State=1 then Y=1:Credit = Credit + AwardHY:End If
      If Y=0 And L14.State=1  And L5.State=1  And L3.State=1  And L10.State=1 And L20.State=1 then Y=1:Credit = Credit + AwardHY:End If
      If Y=0 And L8.State=1 And L14.State=1 And L5.State=1  And L3.State=1  then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L8.State=1 And L14.State=1 And L5.State=1  And L10.State=1 then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L8.State=1 And L14.State=1 And L5.State=1  And L20.State=1 then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L8.State=1 And L14.State=1 And L3.State=1  And L10.State=1 then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L8.State=1 And L14.State=1 And L3.State=1  And L20.State=1 then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L8.State=1 And L5.State=1  And L3.State=1  And L10.State=1 then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L8.State=1 And L5.State=1  And L3.State=1  And L20.State=1 then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L8.State=1 And L5.State=1  And L10.State=1 And L20.State=1 then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L8.State=1 And L3.State=1  And L10.State=1 And L20.State=1 then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L14.State=1  And L5.State=1  And L3.State=1  And L10.State=1 then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L14.State=1  And L5.State=1  And L3.State=1  And L20.State=1 then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L14.State=1  And L3.State=1  And L10.State=1 And L20.State=1 then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L5.State=1 And L3.State=1  And L10.State=1 And L20.State=1 then Y=1:Credit = Credit + AwardMY:End If
      If Y=0 And L8.State=1 And L14.State=1 And L5.State=1  then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L8.State=1 And L14.State=1 And L3.State=1  then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L8.State=1 And L14.State=1 And L10.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L8.State=1 And L14.State=1 And L20.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L8.State=1 And L5.State=1  And L3.State=1  then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L8.State=1 And L5.State=1  And L10.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L8.State=1 And L5.State=1  And L20.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L8.State=1 And L3.State=1  And L10.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L8.State=1 And L3.State=1  And L20.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L8.State=1 And L10.State=1 And L20.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L14.State=1  And L5.State=1  And L3.State=1  then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L14.State=1  And L5.State=1  And L10.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L14.State=1  And L5.State=1  And L20.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L14.State=1  And L3.State=1  And L10.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L14.State=1  And L3.State=1  And L20.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L14.State=1  And L10.State=1 And L20.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L5.State=1 And L3.State=1  And L10.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L5.State=1 And L3.State=1  And L20.State=1 then Y=1:Credit = Credit + AwardLY: End If
      If Y=0 And L5.State=1 And L10.State=1 And L20.State=1 then Y=1:Credit = Credit + AwardLY: End If

      'Blue feature F1: 3 Blue = 5 Green   F2: 2 Blue = 5 Green
      BBB=0
      If F1.State=1  Then
        If  L16.State=1 And L13.State=1 And L21.State=1 then BBB=1:Credit=Credit + AwardHG: End If
      End If

      If F2.State=1 Then
        If BBB=0 And L16.State=1 And L13.State=1 then Credit=Credit + AwardHG:End If
        If BBB=0 And L16.State=1 And L21.State=1 then Credit=Credit + AwardHG:End If
        If BBB=0 And L13.State=1 And L21.State=1 then Credit=Credit + AwardHG:End If
      End If

      'Green 7 Blocks
      G=0
      If G=0 And L23.State=1  And L25.State=1 And L11.State=1 And L24.State=1 And L18.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L23.State=1  And L25.State=1 And L11.State=1 And L24.State=1 And L17.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L23.State=1  And L25.State=1 And L11.State=1 And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L23.State=1  And L25.State=1 And L11.State=1 And L18.State=1 And L17.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L23.State=1  And L25.State=1 And L11.State=1 And L18.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L23.State=1  And L25.State=1 And L18.State=1 And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L23.State=1  And L11.State=1 And L24.State=1 And L18.State=1 And L17.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L23.State=1  And L11.State=1 And L24.State=1 And L18.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L23.State=1  And L11.State=1 And L24.State=1 And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L23.State=1  And L25.State=1 And L24.State=1 And L18.State=1 And L17.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L23.State=1  And L25.State=1 And L24.State=1 And L18.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L23.State=1  And L25.State=1 And L24.State=1 And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L23.State=1  And L11.State=1 And L18.State=1 And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L25.State=1  And L11.State=1 And L24.State=1 And L18.State=1 And L17.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L25.State=1  And L11.State=1 And L24.State=1 And L18.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L25.State=1  And L11.State=1 And L24.State=1 And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L25.State=1  And L24.State=1 And L18.State=1 And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L25.State=1  And L11.State=1 And L18.State=1 And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If
      If G=0 And L11.State=1  And L24.State=1 And L18.State=1 And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardHG:End If

      If G=0 And L23.State=1  And L25.State=1 And L11.State=1 And L24.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L25.State=1 And L11.State=1 And L18.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L25.State=1 And L11.State=1 And L17.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L25.State=1 And L11.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L25.State=1 And L24.State=1 And L18.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L25.State=1 And L24.State=1 And L17.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L25.State=1 And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L11.State=1 And L24.State=1 And L18.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L11.State=1 And L24.State=1 And L17.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L11.State=1 And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L11.State=1 And L18.State=1 And L17.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L11.State=1 And L18.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L24.State=1 And L18.State=1 And L17.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L23.State=1  And L24.State=1 And L18.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L25.State=1  And L11.State=1 And L24.State=1 And L18.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L25.State=1  And L11.State=1 And L24.State=1 And L17.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L25.State=1  And L11.State=1 And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L25.State=1  And L24.State=1 And L18.State=1 And L17.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L25.State=1  And L24.State=1 And L18.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L11.State=1  And L24.State=1 And L18.State=1 And L17.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L11.State=1  And L24.State=1 And L18.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If
      If G=0 And L11.State=1  And L18.State=1 And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardMG:End If

      If G=0 And L23.State=1  And L25.State=1 And L11.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L25.State=1 And L24.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L25.State=1 And L18.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L25.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L25.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L11.State=1 And L24.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L11.State=1 And L18.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L11.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L11.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L24.State=1 And L18.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L24.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L18.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L18.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L23.State=1  And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L25.State=1  And L11.State=1 And L24.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L25.State=1  And L11.State=1 And L18.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L25.State=1  And L11.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L25.State=1  And L11.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L25.State=1  And L24.State=1 And L18.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L25.State=1  And L24.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L25.State=1  And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L25.State=1  And L18.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L25.State=1  And L18.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L25.State=1  And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L11.State=1  And L24.State=1 And L18.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L11.State=1  And L24.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L11.State=1  And L24.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L11.State=1  And L18.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L11.State=1  And L18.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L11.State=1  And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L24.State=1  And L18.State=1 And L17.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L24.State=1  And L18.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L24.State=1  And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
      If G=0 And L18.State=1  And L17.State=1 And L15.State=1 then G=1:Credit = Credit + AwardLG:End If
  End If
End Sub

Sub OKGAME()
  Playsound "ScoreBells"
  UpdateCredits
  For Each xx in Score_Lights:xx.State = 0: Next
  For Each ff in Feature_lights:ff.State = 0: Next
  YellowLight.state=0
  RedLight.state=0
  OKGAME_TIMER.enabled=1
  If RedR.state=1 Then
    R6.state=1:Red=6
    Y4.state=1:Yellow=6
    Select Case Int(Rnd*3)+1
      Case 1: G3.state=1:Green=3
      Case 2: G2.state=1:Green=2
      Case 3: G1.state=1:Green=1
    End Select
    LA.state=1
    LB.state=1
    LC.state=1
    LD.state=1
    LO.state=1
    LK.state=1
    RedR.state=0
  End If

  If RedO.state=1 Then
    R5.state=1:Red=5
    Y6.state=1:Yellow=6
    G4.state=1:Green=4
    LA.state=1
    LB.state=1
    LC.state=1
    LD.state=1
    LO.state=1
    LK.state=1
    F3.state=1
    F8.state=1
    RedO.state=0
  End If

  If RedL.state=1 Then
    R6.state=1:Red=6
    Y6.state=1:Yellow=6
    G5.state=1:Green=5
    LA.state=1
    LB.state=1
    LC.state=1
    LD.state=1
    LO.state=1
    LK.state=1
    F4.state=1
    F7.state=1
    RedL.state=0
  End If

  If RedL1.state=1 Then
    R5.state=1:Red=5
    Y7.state=1:Yellow=7
    G6.state=1:Green=6
    LA.state=1
    LB.state=1
    LC.state=1
    LD.state=1
    LE.state=1
    LO.state=1
    LK.state=1
    F4.state=1
    RedL1.state=0
  End If

  If RedE.state=1 Then
    R7.state=1:Red=7
    Y6.state=1:Yellow=6
    G7.state=1:Green=7
    LA.state=1
    LB.state=1
    LC.state=1
    LD.state=1
    LE.state=1
    LO.state=1
    LK.state=1
    F3.state=1
    F6.state=1
    RedE.state=0
  End If

  If RedR1.state=1 Then
    R8.state=1:Red=8
    Y8.state=1:Yellow=8
    G8.state=1:Green=8
    LA.state=1
    LB.state=1
    LC.state=1
    LD.state=1
    LE.state=1
    LF.state=1
    LG.state=1
    LO.state=1
    LK.state=1
    F3.state=1
    F9.state=1
    RedR1.state=0
  End If
End Sub

Sub OKGAME_TIMER_Timer()
  InProgress=1
  Ballz=6
  ballz2playText.text = "BALLS PER GAME"
  ballz2play.text = (Ballz-1)
  xxx=1
  NewGame
  OKGAME_TIMER.enabled=0
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  On Error Resume Next
  Controller.Pause = False
  Controller.Stop
End Sub
