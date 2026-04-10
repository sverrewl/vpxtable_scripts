'*************************************
'Star-Jet (Bally 1963) - IPD No. 2347
'************************************

'*****************************************************************************************************
' CREDITS
' Art/Blender bord
' Script rothbauerw
' backglass image used by permission from popotte
'*****************************************************************************************************

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

Const BallSize = 25
Const BallMass = 1

Const cGameName = "star_jet"

'******************************************************
'             OPTIONS
'******************************************************

Const BallsPerGame = 5        '3 or 5
Const VolumeDial = 2        'Change volume of hit events
Const RollingSoundFactor = 5    'set sound level factor here for Ball Rolling Sound, 1=default level
Const LiftKey = 1         '0 for Plunger key, 1 for Right Manga Save
Const VROn = False          'set to True if playing in VR

'******************************************************
'             VARIABLES
'******************************************************

Dim i
Dim obj
Dim InProgress:InProgress = False
Dim BallInPlay
Dim MaxPlayers:MaxPlayers = 2
Dim Players, Player, HighPlayer
Dim PrevScore(4)
Dim Score(4), isScoring: isScoring = False
Dim HighScore
Dim Initials
Dim Credits
Dim Match
Dim Replay1:Replay1 = 1300
Dim Replay2:Replay2 = 1500
Dim Replay3:Replay3 = 1600
Dim Replay4:Replay4 = 1700
Dim TableTilted, Tilted(4)
Dim TiltCount
Dim GameOn:GameOn = False
Dim ReelP1Value1, ReelP1Value10, ReelP1Value100, P1K
Dim ReelP2Value1, ReelP2Value10, ReelP2Value100, P2K
Dim GameOver, ShootAgain, ShootAgainBlink, ShootAgainTime: GameOver = False: ShootAgain = 0
Dim TableWidth, TableHeight
Dim lsstate, rsstate
Dim st1state, st2state, st3state, st4state, st5state, spstate
Dim bump1state, bump2state, bump3state
Dim togglestate, togglestar, togglek, togglebip, togglepup, togglesa, toggleplayers, toggletilt

'******************************************************
'             TABLE INIT
'******************************************************

Dim DesktopMode, FSSMode

Sub Table1_Init
  LoadEM

  If B2SOn or Table1.ShowDT = False or Table1.ShowFSS = True or VROn then
    for each i in DT : i.visible = 0 : next
    DesktopMode = False
    FSSMode = True
  Else
    DesktopMode = True
    FSSMode = False
  End If

  loadhs
  if HSA1="" then HSA1=23
  if HSA2="" then HSA2=10
  if HSA3="" then HSA3=18
  UpdatePostIt

  TableTilted = 0

  InProgress = false

  If Credits > 0 Then DOF 125, 1
End Sub

Sub Table1_Exit()
  savehs
  DOF 101, 0
  DOF 102, 0
  DOF 125, 0
  DOF 80, 0
  DOF 25, 0
  DOF 26, 0
  DOF 30, 0
  DOF 31, 0
  DOF 32, 0
  DOF 33, 0
  DOF 34, 0
  DOF 35, 0
  DOF 36, 0

  If B2SOn Then Controller.Stop
End Sub

TableWidth = Table1.width
TableHeight = Table1.height

If Table1.ShowDT = false and not VROn then
  lockdown.visible=0
  layerrails.visible=0
End If

If Not VRon Then
  starLight1a.visible = False
  starLight2a.visible = False
  starLight3a.visible = False
  starLight4a.visible = False
  starLight5a.visible = False
End If

Dim BootCount:BootCount = 0

Sub BootTable_Timer()
  dim xx
  If BootCount = 0 Then
    BootCount = 1
    me.interval = 100
    PlaySoundAt "poweron", Plunger
  Else
    ScoreReel1.setvalue score(1)
    ScoreReel2.setvalue score(2)
    CreditReel.setvalue Credits

    '*****GI Lights On
    If bcuCount <> 0 then
      bcuCount = bcuCount - 1
    Else
      bcuCount = BumperMax - 1
    End If
    ToggleStates

    playfield_meshoff.visible=0
    For each obj in layer1all: obj.image = "layer1": Next
    For each obj in layer2all: obj.image = "layer2": Next
    For each obj in layer3all: obj.image = "layer3": Next
    For each obj in layer4all: obj.image = "layer4": Next
    For each obj in bumpcapsall: obj.image = "bumpcaps": Next
    For each obj in bumpskirtsall: obj.image = "bumpskirts": Next
    For each obj in bumpbasesall: obj.image = "bumpbases":Next
    backbox.image = "backbox"
    backglassGI.image="gion scaled"
    lflip.image="lflipGION"
    rflip.image="rflipGION"
    gate1.image="gate1"
    For each obj in GI: obj.visible = true: Next
    for each xx in GIlights: xx.state = 1: Next
    SetReels

    GameOn = True
    me.enabled = False
  End If
End Sub

Sub BootB2S_Timer()
  If B2SOn Then
    Controller.B2SSetCredits Credits
    Controller.B2SSetScorePlayer 1,Score(1)
    Controller.B2SSetScorePlayer 2,Score(2)
    If Score(1) > 999 Then  Controller.B2SSetData 25, P1K
    If Score(2) > 999 Then  Controller.B2SSetData 26, P2K
  End If

  me.enabled = False
End Sub

'******************************************************
'             KEYS
'******************************************************

Dim InSL
InSL = 0

Sub Table1_KeyDown(ByVal keycode)

  If keycode = PlungerKey Then
    If Liftkey = 0 Then
      If InSL = 1 Then
        Plunger.PullBack
        PlaySoundAt "plungerpull", Plunger
      End If
    Else
      Plunger.PullBack
      PlaySoundAt "plungerpull", Plunger
    End If

    If LiftKey = 0 and BallRelease.ballcntover <> 0 Then
      PlaySoundAt "ballout", BallRelease
    End If
  End If

  If keycode = RightMagnaSave Then
    If LiftKey = 1 and BallRelease.ballcntover <> 0 Then
      PlaySoundAt "ballout", BallRelease
    End If
  End If

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=0 Then
    LeftFlipper.RotateToEnd
    PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers),LeftFlipper, 0.67
    DOF 101, 1
    PlaySoundAtVolLoops "buzzL",LeftFlipper,0.05,-1
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=0 Then
    RightFlipper.RotateToEnd
    PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers),RightFlipper, 0.67
    DOF 102, 1
    PlaySoundAtVolLoops "buzz",LeftFlipper,0.05,-1
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
    checktilt
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
    checktilt
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
    checktilt
  End If

  If keycode = MechanicalTilt Then
    gametilted
  End If

  If keycode = AddCreditKey And GameOn then
    playsoundAt "coinin", drain
    coindelay.enabled=true
  end if

  if keycode=StartGameKey and Credits>0 and InProgress=false And GameOn And Not HSEnterMode=true  and PulseCreditReel.enabled = 0 then
    AddCredit(-1)
    playsoundAt "click", drain
    Players = 1
    Player = 1
    BallInPlay = 1
    toggleplayers = 1
    StartNewGame.enabled = True
  elseif keycode=StartGameKey and Credits>0 and InProgress=true And GameOn And Not HSEnterMode=true and BallInPlay = 1 and Players = 1 and PulseCreditReel.enabled = 0 then
    AddCredit(-1)
    playsoundAt "click", drain
    Players = 2
    toggleplayers = 1
  end if

  If HSEnterMode Then HighScoreProcessKey(keycode)

End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    If Liftkey =0 Then
      If InSL = 1 Then
        Plunger.Fire
        PlaySoundAt "plunger", Plunger
      End If
    Else
        Plunger.Fire
        PlaySoundAt "plunger", Plunger
    End If

    If LiftKey = 0 and BallRelease.ballcntover <> 0 Then
      ReleaseBall.visible = True
      BallRelease.Kick 90, 7
    End If

  End If

  If keycode = RightMagnaSave Then
    If LiftKey = 1 and BallRelease.ballcntover <> 0 Then
      ReleaseBall.visible = True
      BallRelease.Kick 90, 7
    End If
  End If

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=0 Then
    LeftFlipper.RotateToStart
    PlaySoundAt SoundFX("FlipperDown",DOFFlippers), LeftFlipper
    DOF 101, 0
    StopSound "buzzL"
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=0 Then
    RightFlipper.RotateToStart
    PlaySoundAt SoundFX("FlipperDown",DOFFlippers), RightFlipper
    DOF 102, 0
    StopSound "buzz"
  End If
End Sub

sub coindelay_timer
  AddCredit(1)
  coindelay.enabled=false
end sub

Sub shooterlane_hit()
  InSL = 1
End Sub

Sub shooterlane_unhit()
  InSL = 0
End Sub


'******************************************************
'           GAME TIMER
'     Ball Rolling and Shadows, Primitive Updates,
'     B2S, Desktop, and FSS Updates
'           Lights
'******************************************************

Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)
ShootAgainBlink = 500
ShootAgainTime = 0

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to (tnob - 1)
        rolling(i) = False
    Next
End Sub

Sub GameTimer_Timer()

  '****** Flipper shadows and primitives ******
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  LFlip.RotZ = LeftFlipper.CurrentAngle
  RFlip.RotZ = RightFlipper.CurrentAngle
  If GameOn Then
    If bumper1L.state = 1 then
      bumpskirt2.image="bumpskirtslit"
      bumpbase2.image="bumpbaseslit"
      bumpcap2.image="bumpcapslit"
      Else
      bumpskirt2.image="bumpskirts"
      bumpbase2.image="bumpbases"
      bumpcap2.image="bumpcaps"
    End If
    If bumper2L.state = 1 then
      bumpskirt1.image="bumpskirtslit"
      bumpbase1.image="bumpbaseslit"
      bumpcap1.image="bumpcapslit"
      Else
      bumpskirt1.image="bumpskirts"
      bumpbase1.image="bumpbases"
      bumpcap1.image="bumpcaps"
    End If
    If bumper3L.state = 1 then
      bumpskirt3.image="bumpskirtslit"
      bumpbase3.image="bumpbaseslit"
      bumpcap3.image="bumpcapslit"
      Else
      bumpskirt3.image="bumpskirts"
      bumpbase3.image="bumpbases"
      bumpcap3.image="bumpcaps"
    End If
  End If

  '****** Insert and Bumper Lights ******
  If togglestate then
    rslingl.state = rsstate
    rslinglb.state = rsstate

    lslingl.state = lsstate
    lslinglb.state = lsstate

    Bumper1L.state = bump1state
    Bumper1La.state = bump1state

    Bumper2L.state = bump2state
    Bumper2La.state = bump2state

    Bumper3L.state = bump3state
    Bumper3La.state = bump3state

    togglestate = 0
  End If

  If togglestar Then
    starLight1.state = st1state
    starLight1a.state = st1state
    starLight1b.state = st1state

    starLight2.state = st2state
    starLight2a.state = st2state
    starLight2b.state = st2state

    starLight3.state = st3state
    starLight3a.state = st3state
    starLight3b.state = st3state

    starLight4.state = st4state
    starLight4a.state = st4state
    starLight4b.state = st4state

    starLight5.state = st5state
    starLight5a.state = st5state
    starLight5b.state = st5state

    spLight1.state = spstate
    spLight1b.state = spstate
    spLight2.state = spstate
    spLight2b.state = spstate

    togglestar = 0
  End If


  'B2S Updates ***************************************************************************
  If B2Son and GameOn Then
    If InProgress = False and GameOver = False Then
      If Match = 0 Then
        Controller.B2SSetMatch 100
      Else
        Controller.B2SSetMatch Match*10
      End If
      Controller.B2SSetGameOver 1
      Controller.B2SSetData 80,1
      Controller.B2SSetBallInPlay 0
      Controller.B2SSetData 52,0
      Controller.B2SSetData 53,0
    End If

    If InProgress = True and GameOver = True Then
      Controller.B2SSetMatch 0
      Controller.B2SSetScorePlayer 1,0
      Controller.B2SSetScorePlayer 2,0
      Controller.B2SSetGameOver 0
    End If

    If toggleplayers = 1 Then
      If Players = 1 or Players = 2 Then Controller.B2SSetData 54,1 else  Controller.B2SSetData 54,0
      If Players = 2 Then Controller.B2SSetData 55,1 else  Controller.B2SSetData 55,0
    End If

    If togglepup = 1 Then
      If Player = 1 Then Controller.B2SSetData 52,1 else  Controller.B2SSetData 52,0
      If Player = 2 Then Controller.B2SSetData 53,1 else  Controller.B2SSetData 53,0
    End If

    If togglebip = 1 Then
      Controller.B2SSetBallInPlay BallInPlay
    End If

    If togglek = 1 Then
      Controller.B2SSetData 25, P1K
      Controller.B2SSetData 26, P2K
    End If

    If togglesa = 1 then
      If ShootAgain = 1 Then
        Controller.B2SSetShootAgain 1
      Else
        Controller.B2SSetShootAgain 0
      End If
    End If

    If Toggletilt = 1 Then
      If Tilted(1) = 1 Then Controller.B2SSetData 50,1 Else Controller.B2SSetData 50,0
      If Tilted(2) = 1 Then Controller.B2SSetData 51,1 Else Controller.B2SSetData 51,0
    End If

  End If

  'Desktop Updates ***********************************************************************
  If DesktopMode and GameOn Then

    If InProgress = False and GameOver = False Then
      If Match = 0 Then
        Matchtxt.text = "Match 0"
      Else
        Matchtxt.text = "Match " & Match
      End If
      gamov.text = "GAME OVER"
      biptext.text= " "
    End If

    If InProgress = True and GameOver = True Then
      Matchtxt.text = ""
      ScoreReel1.resettozero
      ScoreReel2.resettozero
      gamov.text = ""
    End If

    If togglepup = 1 Then
      If Player = 1 Then
        pup1txt.text = "1 Up"
        pup2txt.text = ""
      Elseif Player = 2 Then
        pup1txt.text = ""
        pup2txt.text = "2 Up"
      Else
        pup1txt.text = ""
        pup2txt.text = ""
      End If
    End If

    If togglebip = 1 Then
      If BallInPlay = 0 Then
        biptext.text = ""
      Else
        biptext.text = "Ball In Play " & BallInPlay
      End If
    End If

    If togglesa = 1 then
      If ShootAgain = 1 Then
        ShootAgaintxt.text = "Shoot Again"
      Else
        ShootAgaintxt.text = " "
      End If
    End If

    If Toggletilt = 1 Then
      If Tilted(1) = 1 Then Tilttxt.text = "TILT - 1"
      If Tilted(2) = 1 Then Tilttxt.text = "TILT - 2"
      If Tilted(1) = 0 and Tilted(2) = 0 Then Tilttxt.text = ""
    End If
  End If

  'FSS Updates ***************************************************************************
  If FSSMode and GameOn Then
    If InProgress = False and GameOver = False Then
      If Match = 0 Then lampmatch0.visible = 1 else lampmatch0.visible = 0
      If Match = 1 Then lampmatch1.visible = 1 else lampmatch1.visible = 0
      If Match = 2 Then lampmatch2.visible = 1 else lampmatch2.visible = 0
      If Match = 3 Then lampmatch3.visible = 1 else lampmatch3.visible = 0
      If Match = 4 Then lampmatch4.visible = 1 else lampmatch4.visible = 0
      If Match = 5 Then lampmatch5.visible = 1 else lampmatch5.visible = 0
      If Match = 6 Then lampmatch6.visible = 1 else lampmatch6.visible = 0
      If Match = 7 Then lampmatch7.visible = 1 else lampmatch7.visible = 0
      If Match = 8 Then lampmatch8.visible = 1 else lampmatch8.visible = 0
      If Match = 9 Then lampmatch9.visible = 1 else lampmatch9.visible = 0

      lampgameover.visible=1 'Turn on Game Over

      lampbip1.visible = 0
      lampbip2.visible = 0
      lampbip3.visible = 0
      lampbip4.visible = 0
      lampbip5.visible = 0

      lampp1up.visible=0
      lampp2up.visible=0
    End If

    If InProgress = True and GameOver = True Then
      lampmatch0.visible = 0
      lampmatch1.visible = 0
      lampmatch2.visible = 0
      lampmatch3.visible = 0
      lampmatch4.visible = 0
      lampmatch5.visible = 0
      lampmatch6.visible = 0
      lampmatch7.visible = 0
      lampmatch8.visible = 0
      lampmatch9.visible = 0

      lampgameover.visible=0
    End If

    If toggleplayers = 1 Then
      if players = 1 Then
        lampp1canplay.visible=1
        lampp2canplay.visible=0
      end if
      if players = 2 Then
        lampp1canplay.visible=1
        lampp2canplay.visible=1
      End if
    End If

    If togglepup = 1 Then   'Set Player Up using variable 'Player'
      if player = 1 Then
        lampp1up.visible=1
        lampp2up.visible=0
      End if
      if player = 2 Then
        lampp1up.visible=0
        lampp2up.visible=1
      End if
    End If

    If togglebip = 1 Then
      If BallInPlay = 1 Then lampbip1.visible = 1 else lampbip1.visible = 0
      If BallInPlay = 2 Then lampbip2.visible = 1 else lampbip2.visible = 0
      If BallInPlay = 3 Then lampbip3.visible = 1 else lampbip3.visible = 0
      If BallInPlay = 4 Then lampbip4.visible = 1 else lampbip4.visible = 0
      If BallInPlay = 5 Then lampbip5.visible = 1 else lampbip5.visible = 0
    End If

    If togglek = 1 Then
      If P1K = 1 then lampp11k.visible=1 else lampp11k.visible=0
      If P2K = 1 then lampp21k.visible=1 else lampp21k.visible=0
    End If

    If togglesa = 1 then
      If ShootAgain = 1 Then
        lampshootagain.visible=1
      Else
        lampshootagain.visible=0
      End If
    End If

    If Toggletilt = 1 Then  'Set tilt using variable 'TableTilted'
      If Tilted(1) = 1 Then lampp1tilt.visible=1 Else lampp1tilt.visible=0
      If Tilted(2) = 1 Then lampp2tilt.visible=1 Else lampp2tilt.visible=0
    End If
  End If

  'Updates Toggles ***********************************************************************

  If InProgress = False and GameOver = False and GameOn Then
    GameOver = True
  End If

  If InProgress = True and GameOver = True and GameOn Then
    GameOver = False
  End If

  If togglesa = 1 Then
    togglesa = 0
    if ShootAgain = 1 or ShootAgain = 2 Then
      ShootAgainTime = GameTime
    End If
  end if

  If toggleplayers = 1 Then
    toggleplayers = 0
  End If

  If togglepup = 1 Then
    togglepup = 0
  End If

  If togglebip = 1 Then
    togglebip = 0
  End If

  If togglek = 1 Then
    togglek = 0
  End If

  If ShootAgain > 0 Then
    If GameTime - ShootAgainTime > ShootAgainBlink Then
      togglesa = 1
      If ShootAgain = 1 Then
        ShootAgain = 2
      Else
        ShootAgain = 1
      End If
    End If
  End If

  If toggletilt= 1 Then
    toggletilt = 0
  End If

  If layer2dbumpl1.visible = True Then
    layer2dbumpl1.visible = False
    layer2dbumpl2.visible = True
  ElseIf layer2dbumpl2.visible = True Then
    layer2dbumpl2.visible = False
    layer2dbumpl3.visible = True
  ElseIf layer2dbumpl3.visible = True Then
    layer2dbumpl3.visible = False
    layer2dbumpl.visible = True
  End If

  If layer2dbumpr1.visible = True Then
    layer2dbumpr1.visible = False
    layer2dbumpr2.visible = True
  ElseIf layer2dbumpr2.visible = True Then
    layer2dbumpr2.visible = False
    layer2dbumpr3.visible = True
  ElseIf layer2dbumpr3.visible = True Then
    layer2dbumpr3.visible = False
    layer2dbumpr.visible = True
  End If


  '****** Ball rolling sounds, ball shadows, and ball drop sounds ******
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls and hide shadows
  For b = UBound(BOT) + 1 to (tnob - 1)
    rolling(b) = False
    StopSound("fx_ballrolling" & b)
    BallShadow(b).visible = 0
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

    BallShadow(b).X = BOT(b).X - (TableWidth/2 - BOT(b).X)/20
    ballShadow(b).Y = BOT(b).Y + 10

    If BOT(b).Z > 22 and BOT(b).Z < 35 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If

    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySound "fx_ball_drop" & b, 0, (ABS(BOT(b).velz)/17)^2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    End If

    If InRect(BOT(b).x, BOT(b).y, 877,1557,927,1557,927,1735,877,1735) Then
      InSL = 1
    End If
  Next
End Sub

'******************************************************
'           RUBBERS/SLINGS
'******************************************************

Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("right_slingshot",DOFContactors), Sling1
  DOF 104, 2

  if rsstate = 1 then
    addscore 10, 0
  else
    addscore 1, 0
  end if

    RSling.Visible = 0
    RSling1.Visible = 1
    sling2.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  PlaySoundAt SoundFX("left_slingshot",DOFContactors), Sling2
  DOF 103, 2

  if lsstate = 1 then
    addscore 10, 0
  else
    addscore 1, 0
  end if

  LSling.Visible = 0
  LSling1.Visible = 1
  sling1.rotx = 20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'******************************************************
'           Toggle States
'******************************************************

Dim BumperLights(32)  ' Bumper Control Unit
Dim bcuCount: bcuCount = 0
Dim BumperMax

BumperMax = 32
BumperLights(0)  = Array(1,1,1)
BumperLights(1)  = Array(0,1,1)
BumperLights(2)  = Array(0,0,1)
BumperLights(3)  = Array(0,0,0)
BumperLights(4)  = Array(1,0,0)
BumperLights(5)  = Array(1,1,0)
BumperLights(6)  = Array(1,1,1)
BumperLights(7)  = Array(0,0,0)
BumperLights(8)  = Array(0,1,1)
BumperLights(9)  = Array(1,0,0)
BumperLights(10) = Array(1,1,0)
BumperLights(11) = Array(0,0,0)
BumperLights(12) = Array(1,1,1)
BumperLights(13) = Array(1,1,1)
BumperLights(14) = Array(0,0,0)
BumperLights(15) = Array(0,1,1)
BumperLights(16) = Array(0,0,0)
BumperLights(17) = Array(0,1,1)
BumperLights(18) = Array(0,1,1)
BumperLights(19) = Array(0,0,1)
BumperLights(20) = Array(1,1,1)
BumperLights(21) = Array(1,0,0)
BumperLights(22) = Array(1,1,0)
BumperLights(23) = Array(0,0,0)
BumperLights(24) = Array(1,1,1)
BumperLights(25) = Array(0,0,1)
BumperLights(26) = Array(0,0,0)
BumperLights(27) = Array(1,0,0)
BumperLights(28) = Array(1,1,0)
BumperLights(29) = Array(0,0,1)
BumperLights(30) = Array(1,0,0)
BumperLights(31) = Array(0,0,1)

Sub ToggleStates()
  PlaySoundAt "cluper", MotorSounds

  bump1state = BumperLights(bcuCount)(0)
  bump2state = BumperLights(bcuCount)(1)
  bump3state = BumperLights(bcuCount)(2)

  rsstate = BumperLights(bcuCount)(0)
  lsstate = BumperLights(bcuCount)(2)

  togglestate = 1
  bcuCount = bcuCount + 1
  if bcuCount >= BumperMax then bcuCount = 0
End Sub


'******************************************************
'             SWITCHES
'******************************************************

Sub startrig1_hit()
  layer2wire001.transz = -25
  PlaySoundAtBall "sensor"
  AddScore 10, 0
  st1state = 0
  togglestar = 1
  checkStarLights
End Sub

Sub startrig1_unhit()
  layer2wire001.transz = 0
End Sub

Sub startrig2_hit()
  layer2wire002.transz = -25
  PlaySoundAtBall "sensor"
  AddScore 10, 0
  st2state = 0
  togglestar = 1
  checkStarLights
End Sub

Sub startrig2_unhit()
  layer2wire002.transz = 0
End Sub

Sub startrig3_hit()
  layer2wire003.transz = -25

  PlaySoundAtBall "sensor"
  AddScore 10, 0
  st3state = 0
  togglestar = 1
  checkStarLights
End Sub

Sub startrig3_unhit()
  layer2wire003.transz = 0
End Sub

Sub startrig4_hit()
  layer2wire004.transz = -25

  PlaySoundAtBall "sensor"
  AddScore 10, 0
  st4state = 0
  togglestar = 1
  checkStarLights
End Sub

Sub startrig4_unhit()
  layer2wire004.transz = 0
End Sub

Sub startrig5_hit()
  layer2wire005.transz = -25
  PlaySoundAtBall "sensor"
  AddScore 10, 0
  st5state = 0
  togglestar = 1
  checkStarLights
End Sub

Sub startrig5_unhit()
  layer2wire005.transz = 0
End Sub

Sub leftoutlane_hit()
  layer2wire006.transz = -25
  PlaySoundAtBall "sensor"
  If spstate = 1 then
    FreeGame
    spstate = 0
    togglestar = 1
  Else
    MultiScore 5, 10
  End If
End Sub

Sub leftoutlane_unhit()
    layer2wire006.transz = 0
End Sub

Sub rightoutlane_hit()
  layer2wire007.transz = -25
  PlaySoundAtBall "sensor"
  If spstate = 1 then
    FreeGame
    spstate = 0
    togglestar = 1
  Else
    MultiScore 5, 10
  End If
End Sub

Sub rightoutlane_unhit()
    layer2wire007.transz = 0
End Sub

Sub CheckStarLights()
  If st1state = 0 and st2state = 0 and st3state = 0 and st4state = 0 and st1state = 0 Then
    spstate = 1
    togglestar = 1
  End If
End Sub

Sub Switch1ptA_hit()
  AddScore 1, 0
End Sub

'Sub Switch1ptB_hit()
' AddScore 1
'End Sub
'
'Sub Switch1ptC_hit()
' AddScore 1
'End Sub
'
'Sub Switch1ptD_hit()
' AddScore 1
'End Sub

Sub Bumper4_hit()
  AddScore 1, 0

  deadbumpskirtleft.transz = skirtTrans*3/4
  bumper4t.enabled=1
End Sub

sub bumper4t_timer
  deadbumpskirtleft.transz = 0
  me.enabled=0
end sub


Sub Bumper5_hit()
  AddScore 1, 0

  deadbumpskirtright.transz = skirtTrans*3/4
  bumper5t.enabled=1
End Sub

sub bumper5t_timer
  deadbumpskirtright.transz = 0
  me.enabled=0
end sub

Sub phys_rubbersdbumpl_hit
  layer2dbumpl1.visible = True
  layer2dbumpl.visible = False
End Sub

Sub phys_rubbersdbumpr_hit
  layer2dbumpr1.visible = True
  layer2dbumpr.visible = False
End Sub


'******************************************************
'             BUMPERS
'******************************************************

Dim SkirtTrans: SkirtTrans = -8

Sub Bumper1_hit()
  If bump1state = 1 Then
    AddScore 10, 0
  Else
    AddScore 1, 0
  End If
  PlaySoundAt SoundFXDOF("fx_bumper1",107,DOFPulse,DOFContactors), Bumper1

  BumpSkirt2.transz = skirtTrans
  me.timerenabled=1
End Sub

sub bumper1_timer
  BumpSkirt2.transz = 0
  me.timerenabled=0
end sub


Sub Bumper2_hit()
  If bump2state = 1 Then
    AddScore 10, 0
  Else
    AddScore 1, 0
  End If
  PlaySoundAt SoundFXDOF("fx_bumper2",108,DOFPulse,DOFContactors), Bumper2

  BumpSkirt1.transz = skirtTrans
  me.timerenabled=1
End Sub

sub bumper2_timer
  BumpSkirt1.transz = 0
  me.timerenabled=0
end sub

Sub Bumper3_hit()
  If bump3state = 1 Then
    AddScore 10, 0
  Else
    AddScore 1, 0
  End If
  PlaySoundAt SoundFXDOF("fx_bumper3",109,DOFPulse,DOFContactors), Bumper3

  BumpSkirt3.transz = skirtTrans
  me.timerenabled=1
End Sub

sub bumper3_timer
  BumpSkirt3.transz = 0
  me.timerenabled=0
end sub


'******************************************************
'             TARGETS
'******************************************************

Dim TargetTrans:TargetTrans = -8

Sub target1_hit()
  layer2target001.transy = TargetTrans
  target1.timerenabled = 1
  CheckKick 2
End Sub

Sub Target1_timer()
  layer2target001.transy = layer2target001.transy - TargetTrans/2
  if layer2target001.transy >= 0 then layer2target001.transy = 0:me.timerenabled=false
End Sub

Sub target2_hit()
  layer2target002.transy = TargetTrans
  target2.timerenabled = 1
  CheckKick 1
End Sub

Sub Target2_timer()
  layer2target002.transy = layer2target002.transy - TargetTrans/2
  if layer2target002.transy >= 0 then layer2target002.transy = 0:me.timerenabled=false
End Sub


Sub target3_hit()
  layer2target003.transy = TargetTrans
  target3.timerenabled = 1
  CheckKick 2
End Sub

Sub Target3_timer()
  layer2target003.transy = layer2target003.transy - TargetTrans/2
  if layer2target003.transy >= 0 then layer2target003.transy = 0:me.timerenabled=false
End Sub

Sub target4_hit()
  layer2target004.transy = TargetTrans
  target4.timerenabled = 1
  CheckKick 3
End Sub

Sub Target4_timer()
  layer2target004.transy = layer2target004.transy - TargetTrans/2
  if layer2target004.transy >= 0 then layer2target004.transy = 0:me.timerenabled=false
End Sub

Sub target5_hit()
  layer2target005.transy = TargetTrans
  target5.timerenabled = 1
  CheckKick 1
End Sub

Sub Target5_timer()
  layer2target005.transy = layer2target005.transy - TargetTrans/2
  if layer2target005.transy >= 0 then layer2target005.transy = 0:me.timerenabled=false
End Sub

'******************************************************
'             KICKERS
'******************************************************

Dim KickBalls
Dim KickerBall1, KickerBall2

Sub leftkick_Hit
  PlaySoundAt "kicker_enter_center", leftkick
  set KickerBall1 = activeball

  BIA = BIP - leftkick.ballcntover - rightkick.ballcntover

  If BIA = 0 Then
    ShootAgain = 1
    togglesa = 1
    Drain.timerenabled = 1
  End If
  kickstepl1 = 0
End Sub

dim kickstep1, kickstep2

Sub leftkick_timer
'   Select Case kickstep1
'        Case 3:pkickarm.rotz=15
'        Case 4:pkickarm.rotz=15
'        Case 5:pkickarm.rotz=15
'        Case 6:pkickarm.rotz=15
'        Case 7:pkickarm.rotz=8
'        Case 8:pkickarm.rotz=3
'   Case 9:pkickarm.rotz=0:leftkick.TimerEnabled = 0:
'    End Select
'   kickstep1 = kickstep1 + 1
End Sub

Sub rightkick_Hit
  PlaySoundAt "kicker_enter_center", rightkick
  set KickerBall2 = activeball

  BIA = BIP - leftkick.ballcntover - rightkick.ballcntover

  If BIA = 0 Then
    ShootAgain = 1
    togglesa = 1
    Drain.timerenabled = 1
  End If
  kickstepr2 = 0
End Sub

dim kickstepl1, kickstepr2

Sub rightkick_timer
'   Select Case kickstep1
'        Case 3:pkickarm.rotz=15
'        Case 4:pkickarm.rotz=15
'        Case 5:pkickarm.rotz=15
'        Case 6:pkickarm.rotz=15
'        Case 7:pkickarm.rotz=8
'        Case 8:pkickarm.rotz=3
'   Case 9:pkickarm.rotz=0:sw23b.TimerEnabled = 0: if sw23b.ballcntover > 0 then SolLSaucer -1
'    End Select
'   kickstep1 = kickstep1 + 1
End Sub


Sub CheckKick(x)
  if x = 1 and leftkick.ballcntover > 0 Then
    KickBalls = 1
    CheckKickTimer.enabled = True
  elseif x = 2 and rightkick.ballcntover > 0 Then
    KickBalls = 2
    CheckKickTimer.enabled = True
  elseif x = 3 Then
    If leftkick.ballcntover > 0 and rightkick.ballcntover > 0 Then
      KickBalls = 3
      CheckKickTimer.enabled = True
    elseif leftkick.ballcntover > 0 Then
      KickBalls = 1
      CheckKickTimer.enabled = True
    elseif rightkick.ballcntover > 0 Then
      KickBalls = 2
      CheckKickTimer.enabled = True
    Else
      AddScore 10, 0
    End If
  else
    AddScore 10, 0
  end if
End Sub

Sub CheckKickTimer_timer()
  If KickBalls = 1 Then
    MultiScore 5, 10

    PlaySoundAt "popper_ball", leftkick
    KickBall KickerBall1, 165, 8, 5, 30
    DOF 110, 2

    leftkick.timerenabled = 1
  ElseIf KickBalls = 2 Then
    MultiScore 5, 10

    PlaySoundAt "popper_ball", rightkick
    KickBall KickerBall2, -165, 8, 5, 30
    DOF 111, 2

    rightkick.timerenabled = 1
  ElseIf KickBalls = 3 Then
    MultiScore 10, 10

    PlaySoundAt "popper_ball", leftkick
    KickBall KickerBall1, 165, 8, 5, 30
    DOF 110, 2

    PlaySoundAt "popper_ball", rightkick
    KickBall KickerBall2, -165, 8, 5, 30
    DOF 111, 2

    rightkick.timerenabled = 1
    leftkick.timerenabled = 1
  End If
  KickBalls = 0
  me.enabled = false
End Sub

'*** PI returns the value for PI
Function PI()
  PI = 4*Atn(1)
End Function

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

'******************************************************
'           SCORING/CREDITS
'******************************************************

Sub AddScore(x,y)

  ShootAgain = 0
  togglesa = 1

  If TableTilted = 0 and (isScoring = False or y = 1) Then
    PrevScore(Player) = Score(Player)
    If x = 1 Then
      Play1Bell
      CheckForRoll 1, Player
      Score(Player) = Score(Player) + 1
      AdvMatch
      ToggleStates
      isScoring = True
    ElseIf x = 10 Then
      Play10Bell
      CheckForRoll 10, Player
      Score(Player) = Score(Player) + 10
      AdvMatch
      isScoring = True
    End If
    CheckFreeGame
  End If

  EVAL("ScoreReel"&player).setvalue score(player)
  If B2SOn Then
    Controller.B2SSetScorePlayer Player,Score(Player)
  End If
  If Score(1) >  999 Then P1K = 1: Togglek = 1
  If Score(2) >  999 Then P2K = 1: Togglek = 1
End Sub

Sub Play1Bell()
  PlaySoundAt SoundFX("10_Point_Bell_Loud",DOFBell), MotorSounds
  DOF 141, 2
End Sub

Sub Play10Bell()
  PlaySoundAt SoundFX("10_Point_Bell_Loud",DOFBell), MotorSounds
  DOF 141, 2
End Sub

Sub CheckForRoll(x, pnum)
  Select Case (x)
    Case 1:
      If pnum = 1 then
        PulseReelP1_1.enabled = 1
        ReelP1Value1 = (ReelP1Value1 + 1) mod 10
        PlaySoundAtVol "reelclick", emreel1_p1, ReelClickVol
        if emreel1_p1.rotx = 261 then CheckForRoll 10, 1
      Else
        PulseReelP2_1.enabled = 1
        ReelP2Value1 = (ReelP2Value1 + 1) mod 10
        PlaySoundAtVol "reelclick", emreel1_p2, ReelClickVol
        if emreel1_p2.rotx = 261 then CheckForRoll 10, 2
      End If
    Case 10:
      If pnum = 1 then
        PulseReelP1_10.enabled = 1
        ReelP1Value10 = (ReelP1Value10 + 1) mod 10
        PlaySoundAtVol "reelclick", emreel10_p1, ReelClickVol
        if emreel10_p1.rotx = 261 then CheckForRoll 100, 1
      Else
        PulseReelP2_10.enabled = 1
        ReelP2Value10 = (ReelP2Value10 + 1) mod 10
        PlaySoundAtVol "reelclick", emreel10_p2, ReelClickVol
        if emreel10_p2.rotx = 261 then CheckForRoll 100, 2
      End If
    Case 100:
      If pnum = 1 then
        PulseReelP1_100.enabled = 1
        ReelP1Value100 = (ReelP1Value100 + 1) mod 10
        PlaySoundAtVol "reelclick", emreel100_p1, ReelClickVol
        if emreel100_p1.rotx = 261 then CheckForRoll 1000, 1
      Else
        PulseReelP2_100.enabled = 1
        ReelP2Value100 = (ReelP2Value100 + 1) mod 10
        PlaySoundAtVol "reelclick", emreel100_p2, ReelClickVol
        if emreel100_p2.rotx = 261 then CheckForRoll 1000, 2
      End If
    Case 1000:
      If pnum = 1 then
        PlaySoundAtVol "reelclick", emreel100_p1, ReelClickVol
      Else
        PlaySoundAtVol "reelclick", emreel100_p2, ReelClickVol
        End If
  End Select
End Sub

Sub AddCredit(direction)
  If PulseCreditReel.enabled = 0 Then
    if direction > 0 and credits >= 25 Then
      'do Nothing
    Else
      if direction = 1 Then
        'playsound SoundFXDOF("knocker",124,DOFPulse,DOFKnocker)
      end if
      credits = credits + direction

      CreditDir = direction
      PulseCreditReel.enabled = 1

      PlaySoundAtVolLoops "reelclick", MotorSounds, 0.5, 0
    End If

    If Credits > 0 Then
      DOF 125, 1
    Else
      DOF 125, 0
    End If

    CreditReel.setvalue Credits
    If B2SOn Then Controller.B2SSetCredits Credits
  End If
End Sub

Sub CheckFreeGame()
  If PrevScore(Player) < Replay1 And Score(Player) >= Replay1 Then FreeGame
  If PrevScore(Player) < Replay2 And Score(Player) >= Replay2 Then FreeGame
  If PrevScore(Player) < Replay3 And Score(Player) >= Replay3 Then FreeGame
  If PrevScore(Player) < Replay4 And Score(Player) >= Replay4 Then FreeGame
End Sub

Sub FreeGame()
  playsoundat SoundFXDOF("knocker",124,DOFPulse,DOFKnocker), MotorSounds
  AddCredit(1)
End Sub

Dim mCountDown, mScore

Sub MultiScore(cycle, Score)
  If isScoring = false Then
    mScore = Score
    mCountDown = cycle - 1
    AddScore mscore, 1
    MultiScoreTimer.enabled = true
  End If
End Sub

Sub MultiScoreTimer_Timer
  mCountDown = mCountDown - 1
  AddScore mscore, 1
  If mCountDown < 1 Then
    me.enabled = False
  End If
End Sub

'******************************************************
'             EMREELS
'******************************************************

Dim ReelStep:ReelStep = 6
Dim ReelClickVol:ReelClickVol=0.5

Sub PulseReelP1_1_Timer
  dim done:done=0
  emreel1_p1.rotx = emreel1_p1.rotx + ReelStep
  if emreel1_p1.rotx > 360 then emreel1_p1.rotx = emreel1_p1.rotx - 360

  Select Case (emreel1_p1.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
    if MultiScoreTimer.enabled = False Then isScoring = False
  End If
End Sub

Sub PulseReelP1_10_Timer
  dim done:done=0
  emreel10_p1.rotx = emreel10_p1.rotx + ReelStep
  if emreel10_p1.rotx > 360 then emreel10_p1.rotx = emreel10_p1.rotx - 360

  Select Case (emreel10_p1.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
    if MultiScoreTimer.enabled = False Then isScoring = False
  End If
End Sub

Sub PulseReelP1_100_Timer
  dim done:done=0
  emreel100_p1.rotx = emreel100_p1.rotx + ReelStep
  if emreel100_p1.rotx > 360 then emreel100_p1.rotx = emreel100_p1.rotx - 360

  Select Case (emreel100_p1.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
    if MultiScoreTimer.enabled = False Then isScoring = False
  End If
End Sub

Sub PulseReelP2_1_Timer
  dim done:done=0
  emreel1_p2.rotx = emreel1_p2.rotx + ReelStep
  if emreel1_p2.rotx > 360 then emreel1_p2.rotx = emreel1_p2.rotx - 360

  Select Case (emreel1_p2.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
    if MultiScoreTimer.enabled = False Then isScoring = False
  End If
End Sub

Sub PulseReelP2_10_Timer
  dim done:done=0
  emreel10_p2.rotx = emreel10_p2.rotx + ReelStep
  if emreel10_p2.rotx > 360 then emreel10_p2.rotx = emreel10_p2.rotx - 360

  Select Case (emreel10_p2.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
    if MultiScoreTimer.enabled = False Then isScoring = False
  End If
End Sub

Sub PulseReelP2_100_Timer
  dim done:done=0
  emreel100_p2.rotx = emreel100_p2.rotx + ReelStep
  if emreel100_p2.rotx > 360 then emreel100_p2.rotx = emreel100_p2.rotx - 360

  Select Case (emreel100_p2.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
    if MultiScoreTimer.enabled = False Then isScoring = False
  End If
End Sub

Dim CreditDir

Sub PulseCreditReel_Timer
  dim done:done=0

  emcredits1.rotx = emcredits1.rotx - ReelStep*CreditDir
  if emcredits1.rotx > 360 then emcredits1.rotx = emcredits1.rotx - 360
  if emcredits1.rotx < 0 then emcredits1.rotx = emcredits1.rotx + 360

  If CreditDir = 1 and emcredits1.rotx < 333 and emcredits1.rotx >= 297 then
    emcredits10.rotx = emcredits10.rotx - ReelStep*CreditDir
  Elseif CreditDir = -1 and emcredits1.rotx <= 333 and emcredits1.rotx > 297 then
    emcredits10.rotx = emcredits10.rotx - ReelStep*CreditDir
  End If

  Select Case (emcredits1.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
  End If
End Sub

Sub ResetReels()
  If emreel1_p1.rotx <> 297 then PulseReelP1_1.enabled = true:PlaySoundAtVol "reelclick", emreel1_p1, ReelClickVol
  If emreel10_p1.rotx <> 297 then PulseReelP1_10.enabled = true:PlaySoundAtVol "reelclick", emreel10_p1, ReelClickVol
  If emreel100_p1.rotx <> 297 then PulseReelP1_100.enabled = true:PlaySoundAtVol "reelclick", emreel100_p1, ReelClickVol

  If emreel1_p2.rotx <> 297 then PulseReelP2_1.enabled = true:PlaySoundAtVol "reelclick", emreel1_p2, ReelClickVol
  If emreel10_p2.rotx <> 297 then PulseReelP2_10.enabled = true:PlaySoundAtVol "reelclick", emreel10_p2, ReelClickVol
  If emreel100_p2.rotx <> 297 then PulseReelP2_100.enabled = true:PlaySoundAtVol "reelclick", emreel100_p2, ReelClickVol
End Sub

Sub SetReels()
  dim digangles(10)
  digangles(0) = 297
  digangles(1) = 333
  digangles(2) = 9
  digangles(3) = 45
  digangles(4) = 81
  digangles(5) = 117
  digangles(6) = 153
  digangles(7) = 189
  digangles(8) = 225
  digangles(9) = 261

  If Score(1) > 0 Then emreel1_p1.rotx = digangles(Int(Mid(StrReverse(score(1)),1,1))):ReelP1Value1=Mid(StrReverse(score(1)),1,1)
  If Score(1) > 9 Then emreel10_p1.rotx = digangles(Int(Mid(StrReverse(score(1)),2,1)))::ReelP1Value10=Mid(StrReverse(score(1)),2,1)
  If Score(1) > 99 Then emreel100_p1.rotx = digangles(Int(Mid(StrReverse(score(1)),3,1))):ReelP1Value100=Mid(StrReverse(score(1)),3,1)
  If Score(1) > 999 Then P1K = 1: togglek= 1

  If Score(2) > 0 Then emreel1_p2.rotx = digangles(Int(Mid(StrReverse(score(2)),1,1))):ReelP2Value1=Mid(StrReverse(score(2)),1,1)
  If Score(2) > 9 Then emreel10_p2.rotx = digangles(Int(Mid(StrReverse(score(2)),2,1)))::ReelP2Value10=Mid(StrReverse(score(2)),2,1)
  If Score(2) > 99 Then emreel100_p2.rotx = digangles(Int(Mid(StrReverse(score(2)),3,1))):ReelP2Value100=Mid(StrReverse(score(2)),3,1)
  If Score(2) > 999 Then P2K = 1: togglek= 1

  digangles(0) = 297
  digangles(9) = 333
  digangles(8) = 9
  digangles(7) = 45
  digangles(6) = 81
  digangles(5) = 117
  digangles(4) = 153
  digangles(3) = 189
  digangles(2) = 225
  digangles(1) = 261

  If Credits > 9 Then emcredits10.rotx = digangles(Int(Mid(StrReverse(Credits),2,1)))
  emcredits1.rotx = digangles(Int(Mid(StrReverse(Credits),1,1)))

End Sub

'******************************************************
'             DRAIN
'******************************************************


Dim BIP, BIA
BIP = 0
BIA = 0

Sub Drain_Hit()
  PlaySoundat "drain", Drain
  Drain.DestroyBall
  BIP = BIP - 1

  BIA = BIP - leftkick.ballcntover - rightkick.ballcntover

  If TableTilted = 1 Then
    ShootAgain = 0
    ResetTilt
  End If

  If BIA = 0 and ShootAgain < 1 then
    PlaySoundAtVolLoops "End_Of_Game", MotorSounds, 0.25, 0
    NextBall
  Elseif BIA = 0 and ShootAgain > 0 Then
    Drain.timerenabled = True
  End If
End Sub

Drain.timerinterval = 900
dim ReleaseBall

Sub Drain_Timer()
  If ShootAgain < 1 Then
    ResetTable
  End If

  Set ReleaseBall = BallRelease.CreateSizedballWithMass(Ballsize,Ballmass)
  ReleaseBall.visible = False
  BIP = BIP + 1

  DOF 105, DOFPulse
  PlaySoundAt SoundFX("solenoid",DOFContactors), BallRelease
  me.timerenabled = False
End Sub

Sub NextBall()
  If Player = Players Then
    BallInPlay = BallInPlay + 1
    togglebip = 1
    Player = 1
    togglepup = 1
  Else
    Player = Player + 1
    togglepup = 1
  End If
  if ballinplay > BallsPerGame then
    InProgress = False
    BallInPlay = 0
    togglebip = 1
    Player = 0
    togglepup = 1

    CheckMatch

    LeftFlipper.RotateToStart
    StopSound "buzzL"
    DOF 101, 0

    RightFlipper.RotateToStart
    StopSound "buzz"
    DOF 102, 0

    If Score(2) > Score(1) Then
      HighPlayer = 2
    Else
      HighPlayer = 1
    End If

    If Score(HighPlayer) >= HighScore Then
      FreeGame
      HighScore = Score(HighPlayer)
      HighScoreEntryInit()
    End If
  Else
    If Tilted(Player) = 1 Then
      NextBall
    Else
      Drain.timerenabled = 1
    End If
  End If
  PlayQuietClick
End Sub

Sub ResetTable
  st1state = 1
  st2state = 1
  st3state = 1
  st4state = 1
  st5state = 1
  spstate = 0
  togglestar = 1
End Sub


'******************************************************
'           START NEW GAME
'******************************************************

Dim NGCount

Sub StartNewGame_Timer()
  NGCount = NGCount + 1

  Select Case (NGCount)
    Case 1:
      InProgress=true
      P1K = 0
      P2K = 0
      togglek = 1
      togglebip = 1
      togglepup = 1

      Tilted(1) = 0
      Tilted(2) = 0
      ResetTilt

      PlaySoundAtVol "MotorRunning", MotorSounds, 0.25
      PlayLoudClick
      ResetReels
    Case 2:
      PlayLoudClick
      ResetReels
    Case 3:
      PlayLoudClick
      ResetReels
    Case 4:
      PlayLoudClick
      ResetReels
    Case 5:
      PlayLoudClick
      ResetReels
    Case 7:
      PlayQuietClick
      ResetReels
    Case 8:
      PlayQuietClick
      ResetReels
    Case 9:
      PlayQuietClick
      ResetReels
    Case 10:
      PlayQuietClick
      ResetReels
    Case 11:
      PlayQuietClick
      ResetReels
    Case 13:
      Drain.timerenabled = true
      Score(1) = 0
      Score(2) = 0
      SetReels

      me.enabled = false
      NGCount = 0
  End Select
End Sub

Sub PlayLoudClick()
  PlaySoundAtVol "Motor_Click_Loud_Long", MotorSounds, 0.05
End Sub

Sub PlayQuietClick()
  PlaySoundAtVol "Motor_Click_Quiet_Long2", MotorSounds, 0.05
End Sub

'******************************************************
'               MATCH
'******************************************************

Dim MatchStep

Sub CheckMatch()
  MatchStep=0
  MatchTimer.enabled = 1
End Sub

Sub MatchTimer_Timer()
  If MatchStep < Players Then
    MatchStep = MatchStep + 1
    if Match=(Score(MatchStep) mod 10) then
      if Tilted(MatchStep) = 0 Then FreeGame
    end if
  Else
    me.enabled = false
  End If
End Sub

Sub AdvMatch()
  Match = Match - 1
  if Match < 0 then Match = 9
End Sub


'******************************************************
'               TILT
'******************************************************

Dim TiltSens

Sub CheckTilt
  If Tilttimer.Enabled = True Then
   TiltSens = TiltSens + 1
   if TiltSens = 3 Then GameTilted
  Else
   TiltSens = 1
   Tilttimer.Enabled = True
  End If
End Sub

Sub Tilttimer_Timer()
  If TiltSens > 0 Then TiltSens = TiltSens - 1
  If TiltSens = 0 Then Tilttimer.Enabled = False
End Sub

Sub GameTilted
  TableTilted = 1
  Tilted(Player) = 1
  toggletilt = 1

  Bumper1.threshold = 100
  Bumper2.threshold = 100
  Bumper3.threshold = 100

  LeftSlingShot.SlingShotThreshold = 100
  RightSlingShot.SlingShotThreshold = 100

  PlayLoudClick

  LeftFlipper.RotateToStart
  StopSound "buzzL"
  DOF 101, 0

  RightFlipper.RotateToStart
  StopSound "buzz"
  DOF 102, 0
End Sub

Sub ResetTilt
  TableTilted = 0
  toggletilt = 1

  Bumper1.threshold = 1.6
  Bumper2.threshold = 1.6
  Bumper3.threshold = 1.6

  LeftSlingShot.SlingShotThreshold = 1.4
  RightSlingShot.SlingShotThreshold = 1.4
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

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(sound, tableobj, Vol)
    PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

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

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

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

Sub RDampen_Timer()
  Cor.Update
End Sub

'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub


dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 0.01, 1.25
RubbersD.addpoint 2, 3.77, 0.96
RubbersD.addpoint 3, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 4, 15.84, 0.874
RubbersD.addpoint 5, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False  'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    if gametime > 100 then Report
  End Sub

  public sub Dampen(aBall)
    if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

' Thalamus - patched : ' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()   'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


End Class

'Tracks ball velocity for judging bounce calculations & angle
'apologies to JimmyFingers is this is what his script does. I know his tracks ball velocity too but idk how it works in particular
dim cor : set cor = New CoRTracker
cor.debugOn = False
'cor.update() - put this on a low interval timer
Class CoRTracker
  public DebugOn 'tbpIn.text
  public ballvel

  Private Sub Class_Initialize : redim ballvel(0) : End Sub
  'TODO this would be better if it didn't do the sorting every ms, but instead every time it's pulled for COR stuff
  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs
    if uBound(allballs) < 0 then if DebugOn then str = "no balls" : TBPin.text = str : exit Sub else exit sub end if: end if
    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      if DebugOn then
        dim s, bs 'debug spacer, ballspeed
        bs = round(BallSpeed(b),1)
        if bs < 10 then s = " " else s = "" end if
        str = str & b.id & ": " & s & bs & vbnewline
        'str = str & b.id & ": " & s & bs & "z:" & b.z & vbnewline
      end if
    Next
    if DebugOn then str = "ubound ballvels: " & ubound(ballvel) & vbnewline & str : if TBPin.text <> str then TBPin.text = str : end if
  End Sub
End Class

'******************************************************
'   FLIPPER CORRECTION SUPPORTING FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

'Methods:
'.TimeDelay - Delay before trigger shuts off automatically. Default = 80 (ms)
'.AddPoint - "Polarity", "Velocity", "Ycoef" coordinate points. Use one of these 3 strings, keep coordinates sequential. x = %position on the flipper, y = output
'.Object - set to flipper reference. Optional.
'.StartPoint - set start point coord. Unnecessary, if .object is used.

'Called with flipper -
'ProcessBalls - catches ball data.
' - OR -
'.Fire - fires flipper.rotatetoend automatically + processballs. Requires .Object to be set to the flipper.

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart, FlipperEnd, LR, PartialFlipCoef
  Private Balls(20), balldata(20)

  dim PolarityIn, PolarityOut
  dim VelocityIn, VelocityOut
  dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
    Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
  End Sub

  Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
  Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
  Public Property Get StartPoint : StartPoint = FlipperStart : End Property
  Public Property Let EndPoint(aInput) : if IsObject(aInput) then FlipperEnd = aInput.x else FlipperEnd = aInput : end if : End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property

  Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
    if gametime > 100 then Report aChooseArray
  End Sub

  Public Sub Report(aChooseArray)   'debug, reports all coords in tbPL.text
    if not DebugOn then exit sub
    dim a1, a2 : Select Case aChooseArray
      case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
      Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
      Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut
      case else :tbpl.text = "wrong string" : exit sub
    End Select
    dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    tbpl.text = str
  End Sub

  Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

  Private Sub RemoveBall(aBall)
    dim x : for x = 0 to uBound(balls)
      if TypeName(balls(x) ) = "IBall" then
        if aBall.ID = Balls(x).ID Then
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
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        balldata(x).Data = balls(x)
        if DebugOn then StickL.visible = True : StickL.x = balldata(x).x    'debug TODO
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
    if abs(Flipper.currentAngle - Flipper.EndAngle) < 30 Then
      PartialFlipCoef = 0
    End If
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1
      dim teststr : teststr = "Cutoff"
      tmp = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
      if tmp < 0.1 then 'if real ball position is behind flipper, exit Sub to prevent stucks  'Disabled 1.03, I think it's the Mesh that's causing stucks, not this
        if DebugOn then TestStr = "real pos < 0.1 ( " & round(tmp,2) & ")" : tbpl.text = Teststr
        'RemoveBall aBall
        'Exit Sub
      end if

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        if DebugOn then teststr = "y velocity: " & round(aBall.vely, 3) & "exit sub" : tbpl.text = teststr
        RemoveBall aBall
        exit Sub
      end if
      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          'TB.TEXT = balldata(x).id & " " & BALLDATA(X).X & VBNEWLINE & FLIPPERSTART & " " & FLIPPEREND
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        if DebugOn then set tmp = new spoofball : tmp.data = aBall : End If
        if IsEmpty(BallData(idx).id) and aBall.VelY < -12 then 'if tip hit with no collected data, do vel correction anyway
          if PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1) > 1.1 then 'adjust plz
            VelCoef = LinearEnvelope(5, VelocityIn, VelocityOut)
            if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
            if Enabled then aBall.Velx = aBall.Velx*VelCoef'VelCoef
            if Enabled then aBall.Vely = aBall.Vely*VelCoef'VelCoef
            if DebugOn then teststr = "tip protection" & vbnewline & "velcoef: " & round(velcoef,3) & vbnewline & round(PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1),3) & vbnewline
            'debug.print teststr
          end if
        Else
     :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
          if Enabled then aBall.Velx = aBall.Velx*VelCoef
          if Enabled then aBall.Vely = aBall.Vely*VelCoef
        end if
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
      'debug
      if DebugOn then
        TestStr = teststr & "%pos:" & round(BallPos,2)
        if IsEmpty(PolarityOut(0) ) then
          teststr = teststr & vbnewline & "(Polarity Disabled)" & vbnewline
        else
          teststr = teststr & "+" & round(1 *(AddX*ycoef*PartialFlipcoef),3)
          if BallPos >= PolarityOut(uBound(PolarityOut) ) then teststr = teststr & "(MAX)" & vbnewline else teststr = teststr & vbnewline end if
          if Ycoef < 1 then teststr = teststr &  "ycoef: " & ycoef & vbnewline
          if PartialFlipcoef < 1 then teststr = teststr & "PartialFlipcoef: " & round(PartialFlipcoef,4) & vbnewline
        end if

        teststr = teststr & vbnewline & "Vel: " & round(BallSpeed(tmp),2) & " -> " & round(ballspeed(aBall),2) & vbnewline
        teststr = teststr & "%" & round(ballspeed(aBall) / BallSpeed(tmp),2)
        tbpl.text = TestSTR
      end if
    Else
      'if DebugOn then tbpl.text = "td" & timedelay
    End If
    RemoveBall aBall
  End Sub
End Class

'================================
'Helper Functions


Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
    if not IsEmpty(aArray(x) ) Then
      if IsObject(aArray(x)) then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  if offset < 0 then offset = 0
  redim aArray(aCount-1+offset) 'Resize original array
  for x = 0 to aCount-1   'set objects back into original array
    if IsObject(a(x)) then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub


Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

Function NullFunctionZ(aEnabled):End Function '1 argument null function placeholder  TODO move me or replac eme

Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
      id = .ID : mass = .mass : radius = .radius
    end with
  End Property
  Public Sub Reset()
    x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty
    id = Empty : mass = Empty : radius = Empty
  End Sub
End Class


Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  'Clamp if on the boundry lines
  'if L=1 and Y < yLvl(LBound(yLvl) ) then Y = yLvl(lBound(yLvl) )
  'if L=uBound(xKeyFrame) and Y > yLvl(uBound(yLvl) ) then Y = yLvl(uBound(yLvl) )
  'clamp 2.0
  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function


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
  HSScorex = HighScore
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


'******************************************************
'           LOAD & SAVE TABLE
'******************************************************
sub savehs
  savevalue "StarJet_VPX", "Credits", Credits
  savevalue "StarJet_VPX", "Match", Match
  savevalue "StarJet_VPX", "BumperState", bcuCount
  savevalue "StarJet_VPX", "HighScore", Highscore
  savevalue "StarJet_VPX", "HSA1", HSA1
  savevalue "StarJet_VPX", "HSA2", HSA2
  savevalue "StarJet_VPX", "HSA3", HSA3
  savevalue "StarJet_VPX", "Score", Score(1)
  savevalue "StarJet_VPX", "Score2", Score(2)
end sub

sub loadhs
  HighScore=0
  Credits=0
  Match=0

  dim temp
  temp = LoadValue("StarJet_VPX", "Credits")

  temp = LoadValue("StarJet_VPX", "Credits")
  If (temp <> "") then Credits = CDbl(temp)

  temp = LoadValue("StarJet_VPX", "Match")
  If (temp <> "") then Match = CDbl(temp)

  temp = LoadValue("StarJet_VPX", "BumperState")
  If (temp <> "") then bcuCount = CDbl(temp)

  temp = LoadValue("StarJet_VPX", "HighScore")
  If (temp <> "") then HighScore = CDbl(temp)

  temp = LoadValue("StarJet_VPX", "hsa1")
  If (temp <> "") then HSA1 = CDbl(temp)

  temp = LoadValue("StarJet_VPX", "hsa2")
  If (temp <> "") then HSA2 = CDbl(temp)

  temp = LoadValue("StarJet_VPX", "hsa3")
  If (temp <> "") then HSA3 = CDbl(temp)

  temp = LoadValue("StarJet_VPX", "Score")
  If (temp <> "") then Score(1) = CDbl(temp)

  temp = LoadValue("StarJet_VPX", "Score2")
  If (temp <> "") then Score(2) = CDbl(temp)

  if HighScore=0 then HighScore=500

  SetReels

  if bcuCount > BumperMax Then bcuCount = 0
  If isNull(bcuCount) then bcuCount = 0
end sub

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

    If keycode = StartGameKey Then
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


'******************************************************
'           FUNCTIONS
'******************************************************

'*** PI returns the value for PI
Function PI()
  PI = 4*Atn(1)
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


