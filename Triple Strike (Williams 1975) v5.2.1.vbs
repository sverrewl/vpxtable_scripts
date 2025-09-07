
' Triple Strike 5.2 (Williams 1975) - Aug 2025 - Rawd - Hybrid release

'************************************************************************************************************************************************************
'************************************************************************************************************************************************************
' I started work on this table in 2010...  It's been a long road..  :)
' I received help on the original VP9 table from: Greywolf (original artwork), Rascal, Noah Fentz, UncleWilly, Wizards_Hat, Bob5453, JP Salas & Sabbat Moon.


' Thanks to ** CKPin ** for his help with the update from VP9 to VPX in 2017 along with his own upgrades.(Options menu, 3/5 ball game, new flipper etc).
' I couldn't have done the original VPX conversion without his help, and a lot of his work remains in this release.
' Other contributers at this time included Arngrim, rosve, BorgDog, jesperpark, & Haunts.


' Thanks to Rascal and Steely for all of their help with the VR Room.
' Thanks to Basti for the original Minimal room models and textures from his Safe Cracker VR Release
' Thanks to VPW team for their Example table (Dynamic shadows and ball rolling sounds used)
' Thanks to Leon Alexanian for the EM reel and Ballcount reset code.
' Thanks to Scampa123 and Leon Alexanian for testing and finding bugs

' Lots of people have had their hand in this table now, and it shows.  Thanks everyone!

' More 5.2.0 Updates include:  VRroom cabinet fixes to metals, legs and backbox, Positioned all the table sound effects, added suacer hit sound effect
'*************************************************************************************************************************************************************
'*************************************************************************************************************************************************************

' *** Big Thanks Apophis for the work done on the 5.2 Update below....

' apophis updates
' - organized layers
' - reworked all the physical rubbers
' - added playfield mesh
' - added VPW physics stuff: physical trough, flipper corrections, flipper tricks, rubber dampeners, materials, global physics settings
' - tweaked physical parameters to match video
' - added a subset of Fleep sounds
' - baked GI and bumper cap
' - added drop target shadows

' Thanks to Tomate for the updated bumper and bake!

'**************************************************************************************************************************************************************


Option Explicit
Randomize

Const BallMass = 1
Const BallSize = 50

' Load the controller.vbs file for DOF
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs."
On Error Goto 0

ExecuteGlobal GetTextFile("core.vbs")
If Err Then MsgBox "Unable to open core.vbs."
On Error Goto 0




' ******* OPTIONS ************************************************************************************************************************
' ****************************************************************************************************************************************
' ****************************************************************************************************************************************

GlassImpurities.visible = 1             '0 = Glass scratches off      1 = Glass scratches on - This will only work in VR mode.  It will default to OFF in Desktop or FS
Dim FlipperImage: FlipperImage = 3      ' Use 1 through 5.  Can be hard set here, and swapped on the fly with left Magnasave
Dim BallRollVolume:  BallRollVolume = 0.35  ' Level of ball rolling volume. Value between 0 and 1
Dim KnockerSoundOn: KnockerSoundOn = 1      ' Set this to 0 if you don't want knocker sound on coin drop

'***** END OPTIONS ***********************************************************************************************************************
'*****************************************************************************************************************************************
'*****************************************************************************************************************************************






Dim VRRoom  '0 = Desktop or cabinet,1 = VR
If RenderingMode = 2 Then
  VRRoom = 1
Else
  VRRoom = 0
End If


' Constants
' BackGlass ROM IDs
Const BIP=32
Const TILT_ID=33
Const MATCH_ID=34
Const GAMEOVER_ID=35
Const SAMESHOOTER=36
Const HISCORE=50
'*****************************************************************************************************
'*****************************************************************************************************
Dim Lspeed(7), Lbob(7), blobAng(7), blobRad(7), blobSiz(7), Blob, Bcnt   ' for VR Lavalamp
Dim theball    'name of the ball when regarding speedx and speedy
Dim gBOT
Dim mhole
Dim score
Dim ball
Dim Points
Dim T, obj
Dim Loops, EndLoop
Dim BonusCount, BonusCountMax, Total, BonusScore, Pinbonus(10),a,b
Dim hold1, hold2, hold3
Dim credit, credit2
Dim highscore, postitscore
Dim postitval
Dim tilted
Dim tilt : tilt=false
Dim tilts
Dim gameon
Dim knockered
Dim match
Dim position
Dim PosY
Dim speedx
Dim speedy
Dim finalspeed
Dim ballinplay
Dim Bump1
Dim targeta
Dim targetb
Dim targetc
Dim targetd
Dim cGameName, B2SName
Dim DesktopMode: DesktopMode = Table1.ShowDT
Dim CurrentMinute ' for VR clock
Dim VRDKCounter: VRDKCounter = 1  'VR TV
Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Const VolumeDial = 1        ' Recommended values should be no greater than 1.

cGameName = "triplestrike"   'assumes:  romname entry in directouputconfig*.ini
B2SName = "triplestrike"     'assumes:  B2STableSettings.xml has startup error messages set to 0
ballinplay=false



' Setup VRRoom *******
If VRRoom = 1 then
  for each Object in VRStuff: object.visible = 1: next
  for each object in backdropstuff : object.visible = 0 : next  ' Turns off Backdrop stuff for VR Users.
  BeerTimer.enabled = true
End If

If VRRoom = 0 then
  for each Object in VRStuff: object.visible = 0: next
  GlassImpurities.visible = 0  ' Make sure the scratches are never seen in desktop/FS mode.
  ClockTimer.enabled = False
  LavaTimer.enabled = False
  TimerAnimateCard1.enabled = False
  BeerTimer.enabled = False
End If
'*********************


Sub Table1_Init()

  LoadEM
  loadStoredDataTimer.Enabled=True
  gameon=False
  ballinplay=false
  LoadHighScore
  SetCredits
  OptionMenu_Init
  set theball=kicker2.createball
  gBOT = Array(theball)

  ' Dont run B2S if in VR mode...
  If VRRoom = 1 then
    if B2sOn = true then
      Controller.LaunchBackglass=0
      B2sOn=False
    end if
  End If


  If Desktopmode = false then
    'Cabinet Mode
    for each object in backdropstuff : object.visible = 0 : next
  Else
    'Desktop Mode
    for each object in backdropstuff : object.visible = 1 : next
  End If

  targeta = false
  targetb = false
  targetc = false
  targetd = false

  Reset
  For each obj in AllLights 'Turn all the lights off at table initialisation
    obj.State=0
  Next
  Loops=0           'Used to loop sounds
  'Init Bumper rings
  ringa.IsDropped=1:ringb.IsDropped=1:ringc.IsDropped=1
  postitval=0
  Tilts = 0
  SetBackGlass 45, 1  ' Start BG light animation
  BG.Image="2025BGOFF"
  ' reset Postit here...
  Postit.ImageA="PostItBlank"
  if B2SOn then Controller.B2SSetData 99,0  ' NewB2S call to set 100,000 light off at table load

End Sub


Sub TimerAnimateCard1_Timer()
  VRDKTube.Image = "gil " & VRDKCounter
  VRDKCounter = VRDKCounter + 1
  If VRDKCounter > 17 Then
    VRDKCounter = 1
  End If
End Sub


Sub Table1_Exit
  savehs
  If B2SOn Then Controller.Stop
End Sub


Sub savehs
  On Error Resume Next
  'HighScore = 0: Postitscore = 0  ' FOR TESTING: Resets the saved value
  SaveValue "TripleStrike", "Option1", HighScore
  SaveValue "TripleStrike", "Option2", PostItScore
  SaveValue "TripleStrike", "Option3", credit
  SaveValue "TripleStrike", "Option4", Balls
  SaveValue "TripleStrike", "Option5", liberal
  SaveValue "TripleStrike", "Option6", replays
End Sub


' Uses the VPReg.stg file in the User directory
Sub LoadHighScore
  HighScore=LoadValue("TripleStrike", "Option1")
  If HighScore="" Then
    HighScore=0:PostItScore=0
    SaveValue "TripleStrike", "Option1", HighScore
    SaveValue "TripleStrike", "Option2", PostItScore
  Else
    PostItScore=LoadValue("TripleStrike", "Option2")
  End If
  ' credit = 0
  credit=LoadValue("TripleStrike", "Option3")
  If credit="" Then
    credit=2
    SaveValue "TripleStrike", "Option3", credit
  End If
  Balls=LoadValue("TripleStrike", "Option4")
  If Balls="" Then
    Balls=3   ' first time run will pick 3
    SaveValue "TripleStrike", "Option4", balls
  End If
  liberal=LoadValue("TripleStrike", "Option5")
  If liberal="" Then
    liberal=0
    SaveValue "TripleStrike", "Option5", liberal
  End If
  replays=LoadValue("TripleStrike", "Option6")
  If replays="" Then
    replays=2
    SaveValue "TripleStrike", "Option6", replays
  End If
  backglassTimer.interval = 1000
  backglassTimer.enabled = True
End Sub


' Turns on GI and Backglass in VR on a 1500ms timer
Sub loadStoredDataTimer_Timer()
  loadStoredDataTimer.Enabled=False
  loadStoredDataTimer.Interval=1500
  'Turn Plastics lights On..
  dim g: for each g in GILights: g.state = 1: next  'turn on GI lights
  'Turn VR BG On
    BG.Image="NewTSBGFinalOff"
  VRGameOver.visible = true
End Sub


Sub gi001_animate
    dim a: a = gi001.GetInPlayIntensity/gi001.Intensity
    UpdateMaterial "Plastic with an image4", 0, 0, 0, 0, 0, 0, a, RGB(255,255,255), RGB(100,100,100), RGB(0,0,0), 0, 1, 0, 0, 0, 0
  BumperBase.blenddisablelighting = 0.5*a
End Sub


' BackGlass Handlers
Sub backglassTimer_Timer()
  backglassTimer.Enabled=False
  backglassTimer.interval = 1000

  ' **********  Below is code to update the VR score reels  THANKS for help RASCAL!  ************

  score = highscore    ' settings score to Highscore so I can read it for VR reels.   This is now displayed on the desk beside the wall.

  If score = 0 then
    HFR1.ImageA="Trans"
  Else
    HFR1.ImageA="Paper"&Right(score, 1)
  end If

  If score >9 then
    HFR10.ImageA="Paper"&Left(Right(score,2),1)
  Else
    HFR10.ImageA="Trans"
  end If

  If score >99 then
    HFR100.ImageA="Paper"&Left(Right(score,3),1)
  Else
    HFR100.ImageA="Trans"
  End If

  If score >999 then
    HFR1000.ImageA="Paper"&Left(Right(score,4),1)
  Else
    HFR1000.ImageA="Trans"
  End If

  If score >9999 then
    HFR10000.ImageA="Paper"&Left(Right(score,5),1)
  Else
    HFR10000.ImageA="Trans"
  end If

  If score >99999 then
    HFR100000.ImageA="Paper"&Left(Right(score,6),1)
  Else
    HFR100000.ImageA="Trans"
  End If
End Sub


Sub SetBackGlass(id, value)
  If B2SOn Then
    Controller.B2SSetData id, value
  End If
End Sub


Sub SetCredits()
  Dim credit10, credit1
  If B2SOn Then
    credit10 = int(credit/10)
    credit1 = int(credit - credit10*10)
    Controller.B2SSetCredits 28, credit10
    Controller.B2SSetCredits 29, credit1
  End If
  CreditReel.SetValue(credit)

  '***********************  This is where I need to add VR credit reel code ************
  '************************ set credits on primitive reels
  vrSetCredit credit

  If credit > 0 then
    creditlight.state=1:DOF 120, DOFOn
  Else
    creditlight.state=0:DOF 120, DOFOff
  End If
End Sub


Sub LeftFlipper_animate
  dim a: a = LeftFlipper.CurrentAngle
  LFLogo.RotY = a
  FlipperLsh.rotz = a
End Sub

Sub RightFlipper_animate
  dim a: a = RightFlipper.CurrentAngle
  RFlogo.RotY = a
  FlipperRsh.rotz = a
End Sub


Sub TiltCheck
  If Tilted = FALSE then
    If Tilts = 0 then
      Tilts = Tilts + int(rnd*200)
      TiltTimer.Enabled = TRUE
    Else
      Tilts = Tilts + int(rnd*200)
    End If
    If Tilts >= 400 and Tilted = FALSE then
      Tilted = TRUE
      PlaySound "tilt"
      SetBackGlass TILT_ID, 1
      tiltreel.setvalue(1)
      VRTiltReel.visible = true
      ' turn off bumpers
      bumper1.force=0
      target153.disabled=true
      target169.disabled=true
      gameover()
      PlaySound "Evil2"
      ' Update Highscore here.. even if you tilt, it should count if it's a high score
      topscore
    End If
  End if
End Sub


Sub TiltTimer_Timer()
  If Tilts > 0 then
    Tilts = Tilts - 2
  Else
    TiltTimer.Enabled = FALSE
  End If
End Sub


Dim FirstGame
FirstGame = True

Sub Reset()   'Called at table initialisation and game start
  If B2SOn Then
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScorePlayer1 0
    Controller.B2SSetData 99,0  ' NewB2S call to reset 100,000 light off
  End If

  Randomize
  match= int(10*rnd)*10
  '  Score = 0
  ball = 0
  knockered=0
  tilted = false

  For each obj in AllLights 'Set all lights off
    obj.State=0
  Next
  For each obj in PinLights 'Set all pin lights on
    obj.State=1
  Next

  For each obj in PinLights2  'Set all pin lights on
    obj.State=1
  Next


  ' Lol. below is so dumb.. I dont know why it wouldnt work otherwise...
  If FirstGame = True then
    For each obj in PinLights2  'Set all pin lights off
      obj.State=0
    Next
    FirstGame = false
  End if

  For each obj in DTargLights 'Set all Drop Target Lights On
    obj.State=1
  Next
  For each obj in DTargs    'Raise all Drop Targets
    obj.IsDropped=false
  Next
  For each obj in DTShadows   'Raise all Drop Targets
    obj.visible=true
  Next

  bumper1.force = 13
  target153.disabled=false
  target169.disabled=false
  SetBackGlass TILT_ID, 0
  SetBackGlass GAMEOVER_ID, 1
  SetBackGlass BIP, 0
  SetBackGlass SAMESHOOTER, 0
  SetBackGlass MATCH_ID, 0

  '***************************vr prim score reset to 0
  vrResetScores
  BG.Image="NewTSBGFinalOff"  ' resets the VR bG to have the 100,000 point light OFF

  ' reset Postit here...
  Postit.ImageA="PostItBlank"
  tiltreel.resettozero
  VRTiltReel.visible = false
  gameoverreel.setvalue(1)
  ' ****  Below sets Gameovereeel on VR backglass ****
  VRGameOver.visible = false
  ball1reel.resettozero
  ball2reel.resettozero
  ball3reel.resettozero
  ball4reel.resettozero
  ball5reel.resettozero
  postitreel.setvalue(0)  ' This is the desktop Post-it .. play with this later..

  '*******  SET ALL VR BALL IN PLAY REELS TO OFF HERE  *****

  VRball1reel.visible = false
  VRball2reel.visible = false
  VRball3reel.visible = false
  VRball4reel.visible = false
  VRball5reel.visible = false

  emreel100.setvalue(0)
  emreel110.setvalue(0)
  emreel120.setvalue(0)
  emreel130.setvalue(0)
  emreel140.setvalue(0)
  emreel150.setvalue(0)
  emreel160.setvalue(0)
  emreel170.setvalue(0)
  emreel180.setvalue(0)
  emreel190.setvalue(0)

  '*********** Set VR Match Reels here below..... **********

  VRMatch00.visible = False
  VRMatch10.visible = False
  VRMatch20.visible = False
  VRMatch30.visible = False
  VRMatch40.visible = False
  VRMatch50.visible = False
  VRMatch60.visible = False
  VRMatch70.visible = False
  VRMatch80.visible = False
  VRMatch90.visible = False
End Sub


Sub ballcountersub

  Select Case ballcount

    Case 1
      ball1reel.setValue(1)
      ball2reel.resettozero
      ball3reel.resettozero
      ball4reel.resettozero
      ball5reel.resettozero
      ' **************  Add VRBallinPlay Reel code Here **************

      VRball1reel.visible = true
      VRball2reel.visible = false
      VRball3reel.visible = false
      VRball4reel.visible = false
      VRball5reel.visible = false

    Case 2
      ball1reel.resettozero
      ball2reel.setValue(1)
      ball3reel.resettozero
      ball4reel.resettozero
      ball5reel.resettozero
      ' **************  Add VRBallinPlay Reel code Here **************

      VRball1reel.visible = false
      VRball2reel.visible = true
      VRball3reel.visible = false
      VRball4reel.visible = false
      VRball5reel.visible = false

    Case 3
      ball1reel.resettozero
      ball2reel.resettozero
      ball3reel.setValue(1)
      ball4reel.resettozero
      ball5reel.resettozero
      ' **************  Add VRBallinPlay Reel code Here **************

      VRball1reel.visible = false
      VRball2reel.visible = false
      VRball3reel.visible = true
      VRball4reel.visible = false
      VRball5reel.visible = false

    Case 4
      ball1reel.resettozero
      ball2reel.resettozero
      ball3reel.resettozero
      ball4reel.setValue(1)
      ball5reel.resettozero
      ' **************  Add VRBallinPlay Reel code Here **************

      VRball1reel.visible = false
      VRball2reel.visible = false
      VRball3reel.visible = false
      VRball4reel.visible = true
      VRball5reel.visible = false

    Case 5
      ball1reel.resettozero
      ball2reel.resettozero
      ball3reel.resettozero
      ball4reel.resettozero
      ball5reel.setValue(1)
      ' **************  Add VRBallinPlay Reel code Here **************

      VRball1reel.visible = false
      VRball2reel.visible = false
      VRball3reel.visible = false
      VRball4reel.visible = false
      VRball5reel.visible = true

  End Select

  if ballcount=balltotal then
    ballcountertimer.enabled=0
    exit sub
  end if
  ballcount = ballcount + 1

end sub

Sub ballcountertimer_timer()
  ballcountersub()
end sub

Dim ballcount
Dim balltotal

Sub StartGame()

  for each Object in ScorePaper: object.visible = 0: next  ' Hide HighScore Paper
  backglassTimer.Enabled=False
  Reset
  ball = Balls  'Option Menu Link - Number of Balls
  ballkickertimer2.enabled=true
  credit = credit -1
  creditreel.addvalue(-1)
  SetCredits
  SetBackGlass GAMEOVER_ID, 0
  SetBackGlass BIP, ball
  SetBackGlass SAMESHOOTER, 0
  SetBackGlass MATCH_ID, 0
  SetBackGlass HISCORE, 0
  SetBackGlass 45, 0
  SetBackGlass 41, 0:SetBackGlass 42, 0:SetBackGlass 43, 0:SetBackGlass 44, 0
  emreel1.resettozero
  gameoverreel.resettozero
  VRGameOver.visible = false

  '**************  Reseting Score reel to 0 here.... ****************
  vrResetScores
  ballcount=1
  if balls=3 then balltotal=3
  if balls=5 then balltotal=5
  ballcountertimer.enabled=1
  Playsound "startup3"
  PlaySound "lightreset"  ' Sound for light reset
End Sub


' General function to add score, updates the score reel
Sub AddScore(points)
  Score=Score+points
  '*****************************add score to the primitive reels
  addscoreVR points

  ' **********  Below is code to update the VR score reels  THANKS RASCAL! - REMOVED USING PRIMS NOW - SAVE FOR DESKTOP REINTEGRATION ************

  If score >99999 then BG.Image="NewTSBGFinalOn" :postitreel.setvalue(1)   'This is where we turn the 100,000 point light on.  We must reset this at game start.... leave  line right above, ill make it invisible...
  If score >199999 then Postit.ImageA="PostIt2" :postitreel.setvalue(2)   'turning on VR postit notes here...
  If score >299999 then Postit.ImageA="PostIt3" :postitreel.setvalue(3)
  If score >399999 then Postit.ImageA="PostIt4" :postitreel.setvalue(4)
  If score >499999 then Postit.ImageA="PostIt5" :postitreel.setvalue(5)

  If B2SOn Then Controller.B2SSetScorePlayer1 Score

  If B2SOn Then
    If score >99999 then Controller.B2SSetData 99,1
    If score >199999 then Controller.B2SSetData 99,2
    If score >299999 then Controller.B2SSetData 99,3
    If score >399999 then Controller.B2SSetData 99,4
    If score >499999 then Controller.B2SSetData 99,5
  end If


  EMReel1.addvalue(points)
  'Option Menu Link - Replay Scores
  For i= 1 to 4
    If score => replayscore(i) and knockered<i Then
      knockered=knockered+1
      Replay_Award
    End If
  Next
End Sub


Sub Replay_Award

  If liberal=1 and extraballlight.state=lightstateoff or knockered=1 and extraballlight.state=lightstateoff Then  ' added and statement here..

    extraballlight.state=lightstateon
    ball = ball+1
    SetBackGlass SAMESHOOTER, 1
  Exit sub ' we want to exit the sub here cause we dont want knocker to sound for extra ball, just for replay.
  End If

  PlaySoundAtLevelStatic SoundFXDOF("fpknocker",108,DOFPulse,DOFKnocker), 1, kicker1

  If knockered<>1 Then
    if credit < 9 then credit = credit + 1:creditreel.addvalue(1):SetCredits
  End If
End Sub

'Used by Match
Sub Add_credit
  PlaySoundAtLevelStatic SoundFXDOF("fpknocker",108,DOFPulse,DOFKnocker), 1, kicker1
  if credit < 9 then credit = credit + 1:creditreel.addvalue(1):SetCredits
End Sub


Sub bumper1_hit() 'Top bumper
  If Tilted = true then exit sub
  PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
  PlaySoundAtLevelStatic SoundFXDOF("bumper1",101,DOFPulse,DOFContactors), 1, Bumper1
  AddScore 100
  Bump1=1:ringTimer.enabled=1
End Sub

' Bumper animation
Sub ringTimer_Timer()
  Select Case bump1
    Case 1:Ringa.IsDropped = 0:bump1 = 2
    Case 2:ringb.IsDropped = 0:ringa.IsDropped = 1:bump1 = 3
    Case 3:ringc.IsDropped = 0:ringb.IsDropped = 1:bump1 = 4
    Case 4:ringc.IsDropped = 1: ringTimer.Enabled = 0
  End Select
End Sub


' Left Slingshot
Sub Target169_slingshot()
  If Tilted = true then exit sub
  LS.VelocityCorrect(ActiveBall)
  PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
  PlaySoundAtLevelStatic SoundFXDOF("slingshotl",105,DOFPulse,DOFContactors), 1, SlingMetal5
  AddScore 10
End Sub


' Right Slingshot
Sub Target153_slingshot()
  If Tilted = true then exit sub
  RS.VelocityCorrect(ActiveBall)
  PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
  PlaySoundAtLevelStatic SoundFXDOF("slingshotl",102,DOFPulse,DOFContactors), 1, SlingMetal6
  AddScore 10
End Sub


Sub PointShots_Hit(T)
  If Tilted = true then exit sub
  PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
  AddScore 10
End Sub


'Ball @ Plunger
Sub trigger11_hit()
  DOF 119, DOFOn
  DOF 122, DOFPulse
End sub

Dim SoundReady : SoundReady = true

Sub SoundTrigger_Hit()
  SoundReady = true
End Sub

Sub trigger11_unhit()
  If SoundReady = True then PlaySoundAtLevelStatic "plunger-full-t2", 1, TSPlunger
  SoundReady = false
End sub


' triggers and lights for pins
' Turn off the light of the pin that was hit
' Adds 10 points regardless if light is on or off.
Sub PinTriggers_Hit(T)
  If Tilted = true then exit sub
  PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
  PinLights(T).state=0
  PinLights2(T).state=0

  If liberal=2 Then                    ' Option Menu Link
    Pin_Opposite = OppositeMap(T)    ' T comes in as the pin number - 1
    PinLights(Pin_Opposite).state=0  ' PinLights requires pin number - 1
    PinLights2(Pin_Opposite).state=0
  End If
  checkpins
  AddScore 10
End Sub

' Check to see if all the pins are down,
' if they are then reset lights & add bonus light
Sub checkpins()
  For each obj in PinLights
    If obj.state=1 then Exit Sub
  Next
  For each obj in PinLights
    obj.state=1
  Next

  For each obj in PinLights2
    If obj.state=1 then Exit Sub
  Next
  For each obj in PinLights2
    obj.state=1
  Next

  PlaySound "lightreset"
  DOF 123, DOFPulse
  If lightbonus1.state=1 then
    If lightbonus2.state=1 then
      lightbonus3.state=1
      PlaySound "bonus-advance"
    Else
      lightbonus2.state=1
      PlaySound "bonus-advance"
    End If
  Else
    lightbonus1.state=1
    PlaySound "bonus-advance"
  End If
End Sub

Sub DTargA_hit():DOF 115, DOFPulse:SetBackGlass 41, 1:targeta = true:End Sub
Sub DTargB_hit():DOF 116, DOFPulse:SetBackGlass 42, 1:targetb = true:End Sub
Sub DTargC_hit():DOF 117, DOFPulse:SetBackGlass 43, 1:targetc = true:End Sub
Sub DTargD_hit():DOF 118, DOFPulse:SetBackGlass 44, 1:targetd = true:End Sub


Sub DTargs_hit(T) 'drop targets  hit
  PlaySoundAtLevelStatic SoundFX("targetdrop",DOFDropTargets), 1, DTargLights(T)
  DTargs(T).isdropped=true
  DTargLights(T).state=lightstateoff
  DTShadows(T).visible=false
  checkextraball
  checkabcd
  checktargets
  If Tilted = false then
    PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
    AddScore 1000
  End If
End Sub


Sub checkextraball()  ' check extra ball
  For each obj in DTargLights
    If obj.state=1 Then Exit Sub
  Next

  If leftlight.state=lightstateon and extraballlight.state=lightstateoff Then
    DOF 108, DOFPulse
    ball = ball+1
    extraballlight.state=lightstateon
  End If
End Sub


Sub checkabcd() ' Check abcd lights
  For each obj in DTargLights
    If obj.state=1 Then Exit Sub
  Next
  For each obj in DTargLights
    obj.state=1
  Next
  starlight1.state=lightstateon
  starlight2.state=lightstateon
  starlight3.state=lightstateon
  tenxlight.state=lightstateon
  leftlight.state=lightstateon
  rightlight.state=lightstateon
end sub


Sub checktargets()  ' check targets
  For each obj in DTargs
    If obj.IsDropped=False Then Exit Sub
  Next
  dtargettimer.enabled=true
end sub


' Added timer for abcd target reset
Sub dtargettimer_timer
  For each obj in DTargs    'Raise all Drop Targets
    obj.IsDropped=false
  Next
  For each obj in DTShadows   'Raise all Drop Targets
    obj.visible=true
  Next
  SetBackGlass 41, 0:SetBackGlass 42, 0:SetBackGlass 43, 0:SetBackGlass 44, 0
  PlaySoundAtLevelStatic SoundFX("dropreset",DOFDropTargets), 1, Light5
  if targeta = true then DOF 115, DOFPulse
  if targetb = true then DOF 116, DOFPulse
  if targetc = true then DOF 117, DOFPulse
  if targetd = true then DOF 118, DOFPulse
  targeta = false
  targetb = false
  targetc = false
  targetd = false
  dtargettimer.enabled=false
End Sub


Sub kicker1_hit()'Kicker at the top of the playfield

  SoundSaucerLock
  dim speedx
  dim speedy
  dim finalspeed
  finalspeed=BallSpeed(theball)
  If Tilted = true then
    kicker1.timerenabled=true
    exit sub
  end if

    kicker1.timerenabled=true
    if tenxlight.state=lightstateoff then

      LoopSound
    end if
    if tenxlight.state=lightstateon then
      AddBonus1Kb 5 ' OK I know it's not a bonus, but it uses the exact same code
    end if
    checktoplights  ' HAD to make a new addbonus sub because it conflicted with ball release.
    toplight1.state=lightstateon
End Sub


Sub checktoplights()
  if toplight4.state=lightstateon then
    toplight5.state=lightstateon
    speciallight.state=lightstateon
  end if

  if toplight3.state=lightstateon then toplight4.state=lightstateon
  if toplight2.state=lightstateon then
    toplight3.state=lightstateon
    bonuslight.state=lightstateon
  end if

  if toplight2.state=lightstateon and ball=1 and extraballlight.state=lightstateoff Then   ' added this to take out addaball that CKPin added.
    DOF 108, DOFPulse
    ball = ball+1
    extraballlight.state=lightstateon
  end if
  if toplight1.state=lightstateon then toplight2.state=lightstateon
End Sub


Sub LoopSound()   'Used to loop the bell sound 5 times when the ball lands in the kicker
  Loops=0
  LoopSoundTimer.Interval=400
  LoopSoundTimer.Enabled=true
End Sub


Sub LoopSoundTimer_Timer()  'Used to loop the bell sound 5 times when the ball lands in the kicker
  PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
  PlaySoundAtLevelStatic "bonus-scan", 1, Bumper1

  LoopSoundTimer.Interval=155  '' This is what times the space between bells.
  If loops=4 then LoopSoundTimer.Enabled=false
  'Gave a pause to EM reel for 500 points in top kicker
  Loops=Loops+1
  AddScore (100)
End Sub

' knocker timer for delayed knocker on coin-in  also delayed credit reel animation
Sub knockertimer_timer
  'PlaySound SoundFXDOF("knocker-t1", 108, DOFPulse, DOFKnocker)
  if KnockerSoundOn then PlaySoundAtLevelStatic SoundFXDOF("FPKnocker",108,DOFPulse,DOFKnocker), 1, kicker1
  if credit < 9 then credit = credit + 1:creditreel.addvalue(1):SetCredits
  knockertimer.enabled=false
End Sub


Sub kicker1_timer
    kicker1.kick 189+INT(RND*9),20+INT(RND*4)
    PlaySoundAtLevelStatic SoundFXDOF("eject",107,DOFPulse,DOFContactors), 1, kicker1
    kicker1.timerenabled=false
End Sub


Sub StarTriggers_Hit(T)
  if tilted = true then exit sub         ' Star triggers
  DOF 130 + T, DOFPulse
  If StarLights(T).State=0 then
    AddScore 100
    PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
  Else
    AddScore 1000
    PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
  End If
End Sub


' lane triggers
sub lefttrigger1_hit()
  DOF 111, DOFPulse
  if tilted = true then exit sub
  PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
  if speciallight.state=0 then
    AddScore 1000
    PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
  end if
  if speciallight.state=1 then
    AddScore 10000
  end if
end sub

sub lefttrigger2_hit()
  DOF 112, DOFPulse
  if tilted = true then exit sub
  if leftlight.state=1 then
    AddScore 1000
    PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
  end if
  if leftlight.state=0 then
    AddScore 100
    PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
  end if
end sub

sub righttrigger1_hit():DOF 113, DOFPulse
  if tilted = true then exit sub
  if leftlight.state=1 then
    AddScore 1000
    PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
  end if
  if leftlight.state=0 then
    AddScore 100
    PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
  end if
end sub

sub righttrigger2_hit():DOF 114, DOFPulse
  if tilted = true then exit sub
  PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
  AddScore 1000
end sub


' Keys
Sub Table1_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.PullBack
    TimerVRPlunger.Enabled = True   'VR Plunger
    TimerVRPlunger2.Enabled = False   'VR Plunger
  End If

  ' Below code will START the HighScoreResetTimer
  If keycode = RightMagnaSave Then
    HighScoreResetTimer.enabled = True
  End If

  If keycode = LeftMagnaSave Then
    FlipperImage = FlipperImage + 1
    If FlipperImage = 6 then FlipperImage =1
    SetFlipperImage
  End If

  If keycode = StartGameKey Then
    FrontButton.y = FrontButton.y -5              ' VR Start Button Animation here
  End If

  If keycode = AddcreditKey and credit < 9 Then
    PlaySoundAtLevelStatic "CoinIn",1, Coinslot2
    knockertimer.enabled=True
  End If


  ' Option Menu Link - Flipper Control
  If gameon=false Then
    OptionMenuKeyDownCheck(keycode)
  Else

    If keycode = LeftFlipperKey Then
      If GameOn = TRUE then
        If Tilted = FALSE Then
          SolLFlipper 1
          PlaySoundAtLevelStatic SoundFXDOF("Flipperup2",103,DOFOn, DOFFlippers), 1, LFLogo
          PlaySoundAtLevelStatic "Flipperbuzz2",1, FlipperButtonLeft
          FlipperButtonLeft.X = 2112.9 + 9   ' VR Flipper animation
        end if
      end if
    end if

    If keycode = RightFlipperKey Then
      If GameOn = TRUE then
        If Tilted = FALSE then
          SolRFlipper 1
          PlaySoundAtLevelStatic SoundFXDOF("Flipperup2",104,DOFOn, DOFFlippers), 1, RFLogo
          PlaySoundAtLevelStatic "Flipperbuzz1",1, FlipperButtonRight
          FlipperButtonRight.X = 2104 - 9   ' VR Flipper animation
        End If
      end if
    end if

    ' Below was added for Pinball cabinets that use VP's fake tilt.  (no real tilt bob)

    If keycode = CenterTiltKey and tilted = false and ballinplay=true Then
      Nudge 0, 3
      tiltcheck
    End If

    If keycode = LeftTiltKey and tilted = false and ballinplay=true Then      'and theball.x >500 ????
      Nudge 90, 3
      tiltcheck
    end if
    If keycode = RightTiltKey and tilted = false and ballinplay=true Then
      Nudge 270, 3
      tiltcheck
    End If

    'added below for cab users with real physical tilt mechs...
    If keycode = MechanicalTilt Then
      Tilts = 800
      tiltcheck
    End If

  End If
End Sub


Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    TimerVRPlunger.Enabled = False  'VR Plunger
    TimerVRPlunger2.Enabled = True   ' VR Plunger
    TSPlunger.Y = 975   ' VR Plunger
    PlaySoundAtLevelStatic "Plunger", 1, TSPlunger
  End if

  'Below code will stop the HighScoreResetTimer
  If keycode = RightMagnaSave Then
    HighScoreResetTimer.enabled = False
  End If

  If keycode = StartGameKey Then
    FrontButton.y = FrontButton.y + 5                ' VR Start Button Animation here  This is KeyUp - Reset
  End If

  If keycode = StartGameKey and gameon=false and ball <= 0 and credit >0 and OperatorMenu=0  Then      ' AND MENU ISNT OPEN....
    gameon=true
    StartGame
  End If

  If keycode = LeftFlipperKey Then
    ' Option Menu Link - Timer Turn Off
    OperatorMenuTimer.Enabled = false

    If GameOn = TRUE then
      If Tilted = FALSE then
        SolLFlipper 0
        FlipperButtonLeft.x = 2112.9   'Resets VR flipper button
        PlaySoundAtLevelStatic SoundFXDOF("FlipperDown2",103,DOFOff, DOFFlippers), 1, LFLogo
        stopSound "Flipperbuzz2"
      end if
    end if
  end if
  If keycode = RightFlipperKey Then
    If GameOn = TRUE then
      If Tilted = FALSE then
        SolRFlipper 0
        FlipperButtonRight.x = 2104   'Resets VR flipper button
        PlaySoundAtLevelStatic SoundFXDOF("Flipperdown2",104,DOFOff, DOFFlippers), 1, RFLogo
        stopSound "Flipperbuzz1"
      end if
    end if
  end if
End Sub


Sub matchplay
  if match = 0 then match = 100
  SetBackGlass MATCH_ID, match
  Select Case match
    Case 100: emreel100.setvalue(1): VRMatch00.visible = True
    Case 10:  emreel110.setvalue(1): VRMatch10.visible = True
    Case 20:  emreel120.setvalue(1): VRMatch20.visible = True
    Case 30:  emreel130.setvalue(1): VRMatch30.visible = True
    Case 40:  emreel140.setvalue(1): VRMatch40.visible = True
    Case 50:  emreel150.setvalue(1): VRMatch50.visible = True
    Case 60:  emreel160.setvalue(1): VRMatch60.visible = True
    Case 70:  emreel170.setvalue(1): VRMatch70.visible = True
    Case 80:  emreel180.setvalue(1): VRMatch80.visible = True
    Case 90:  emreel190.setvalue(1): VRMatch90.visible = True
  End Select
  if match = 100 and int(right(score,2))= 0 then
    Add_credit
  Else
    if match = int(right(score,2)) then Add_credit
  End If
End sub

Sub Drain_Hit()
  Drain.Timerenabled = True
End Sub

Sub Drain_Timer()
  Drain.Timerenabled = False
  Drain.kick 57, 20
  DOF 121, DOFPulse
  Tilts = 0
  if tilted = true then
    ball = 0
    gameover()
    SetBackGlass GAMEOVER_ID, 1
    gameoverreel.setvalue(1)
    VRGameOver.visible = true
  else
    singlepinbonus
    checkstrikelights
  end if
  ballinplay=false
  PlaySoundAtLevelStatic "drain-t1", 1, Drain

  SetBackGlass SAMESHOOTER, 0
  ''Fixed below.. with above....  extra balls should not stack.  Going to turn the playfiled extra ball light and BG EB light off here too...  It will turn back on if you get an extra ball by score during bonus count.
  ball = ball - 1
End Sub

' VR Plunger stuff below..........
Sub TimerVRPlunger_Timer
  If TSPlunger.Y < 1100 then    'guessing on number..  trial and error..
    TSPlunger.Y = TSPlunger.Y + 5
  End If
End Sub


Sub TimerVRPlunger2_Timer
  TSPlunger.Y = 975 + (5* Plunger.Position) -20
End Sub


Sub gameover()
  for each Object in ScorePaper: object.visible = 1: next  ' Show High Score Paper

  gameon=false
  SolLFlipper 0
  SolRFlipper 0
  stopSound "Flipperbuzz2"      ' in case flippers are up when the game ends
  stopSound "Flipperbuzz1"
  SetBackGlass BIP, 0
  ball1reel.resettozero
  ball2reel.resettozero
  ball3reel.resettozero
  ball4reel.resettozero
  ball5reel.resettozero
  ' *******  VR BallinPlay Reset below.********
  VRball1reel.visible = false
  VRball2reel.visible = false
  VRball3reel.visible = false
  VRball4reel.visible = false
  VRball5reel.visible = false

  If tilted=false then
    matchplay
  end if
  SetBackGlass 41, 0:SetBackGlass 42, 0:SetBackGlass 43, 0:SetBackGlass 44, 0
  SetBackGlass 45, 1
End sub


'BONUS SYSTEM
'Added a singlepin bonus system to come before the bonus system
sub singlepinbonus
  a=0
  For each obj in PinLights
    a=a+1
    If obj.State=0 Then:PinBonus(a)=1:Else PinBonus(a)=0:End If
  next
  a=0
  b=10
  pinBonusTimer.Enabled=True
end sub

Sub pinbonusTimer_Timer()
  a=a+1
  PlaySoundAtLevelStatic "motorstep", 1, Bumper1
  PlaySoundAtLevelStatic "bonus-scan", 1, Bumper1
  If Pinbonus(a)=1 Then PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes):AddScore 1000
  If a=10 Then
    pinbonusTimer.Enabled=False
    a=0
    Bonuspausetimer.enabled=true
  End If
End Sub

sub Bonuspausetimer_timer()         ' Adds a pause between single pin bonus and strike bonus
  bonus()
  Bonuspausetimer.enabled=false
end sub

sub bonus()
  BonusScore=0
  if lightbonus1.state=0 then afterBonusTimer.Enabled=true           'for kicking out the ball when no strikes are lit
  If lightbonus3.state=1 then   'Check bonus lights & add 10K, 20K or 30K bonus as appropriate
    BonusScore=BonusScore+30
  Else
    if lightbonus2.state=1 then
      BonusScore=BonusScore+20
    Else
      if lightbonus1.state=1 then BonusScore=BonusScore+10
    End If
  End If
  AddBonus1K BonusScore     'Run add bonus routine (plays sounds & updates score reel)
end sub


Sub afterbonustimer_timer
  bonusholdcheck
  playsound "lightreset"
  toplight1.state=0       'If not game over then reset lights
  toplight2.state=0
  toplight3.state=0
  toplight4.state=0
  toplight5.state=0
  starlight1.state=0
  starlight2.state=0
  starlight3.state=0
  leftlight.state=0
  rightlight.state=0
  bonuslight.state=0
  speciallight.state=0
  tenxlight.state=0

  If Ball <= 0 then
    gameover()
    stopsound "flipperbuzz2"
    SetBackGlass GAMEOVER_ID, 1
    gameoverreel.setvalue(1)
    VRGameOver.visible = true
    Topscore
  End If
  If Ball > 0 then ballkickertimer.enabled=true
  afterbonustimer.enabled=false
End Sub


sub topscore
  if clng (score) > clng (highscore) then
    highscore=score:postitscore=postitval
    savehs
  End if
  backglassTimer.Interval=100  '' I upped this time a little bit.   This is after your game, and the high score comes back up on the EM reel.  ***********   This has now been changed to the VR desk, so I made the time low again to update right after game.
  backglassTimer.Enabled=True
end sub


Sub bonusholdcheck()        'Only reset pin lights if bonus hold is not lit
  If bonuslight.state=0 and Ball > 0 then  'only if bonus light is off, pinlights and strikes turn off
    For each obj in PinLights
      obj.state=1
    Next

    For each obj in PinLights2
      obj.state=1
    Next

    PlaySound "lightreset"  ' lightreset sound, comes before ball eject
    lightbonus1.state=0
    lightbonus2.state=0
    lightbonus3.state=0
  End If
  ' This checks the value of the hold light and turns the strike light back on if holds bonus was on
  If bonuslight.state=1 then
    if hold1=1 then lightbonus1.state=1
    if hold2=1 then lightbonus2.state=1
    if hold3=1 then lightbonus3.state=1
  end if
  hold1=0
  hold2=0
  hold3=0
End Sub

Sub BulbInt_Hit()
  theball.BulbIntensityScale = 1
  DOF 119, DOFOff  ' turns off DOF lights in plunger lane.
end sub


Sub ballkickertimer2_timer
  if ball =>1 then
    ballinplay=true                       ' longer starup for first ball
    kicker2.kick 30,17
    ballkickertimer2.enabled=false
    PlaySoundAtLevelStatic SoundFXDOF("loader-ballroll",109,DOFPulse, DOFContactors), 1, Plunger
  End If
end sub


sub ballkickertimer_timer

  if ball =>1 then
    ballinplay=true ' named theball for speed equations later
    kicker2.kick 30,17
    PlaySoundAtLevelStatic "eject", 1, Drain
    PlaySoundAtLevelStatic SoundFXDOF("loader-ballroll",109,DOFPulse, DOFContactors), 1, Plunger
    extraballlight.state=lightstateoff          '' I think we need to reset the playfieled extra ball light here.  -  Had to do this if you gain an extra ball during bonus..  Youll get your extra ball. but need to turn off the light.
  End If
  SetBackGlass BIP, ball
  Select Case ball

    Case 5:
      ball1reel.resettozero
      ball2reel.resettozero
      ball3reel.resettozero
      ball4reel.resettozero
      ball5reel.setValue(1)
      '******* VRBallReel Case.... *****
      VRball1reel.visible = false
      VRball2reel.visible = false
      VRball3reel.visible = false
      VRball4reel.visible = false
      VRball5reel.visible = true

    Case 4:
      ball1reel.resettozero
      ball2reel.resettozero
      ball3reel.resettozero
      ball4reel.setValue(1)
      ball5reel.resettozero
      '******* VRBallReel Case.... *****
      VRball1reel.visible = false
      VRball2reel.visible = false
      VRball3reel.visible = false
      VRball4reel.visible = true
      VRball5reel.visible = false

    Case 3:
      ball1reel.resettozero
      ball2reel.resettozero
      ball3reel.setValue(1)
      ball4reel.resettozero
      ball5reel.resettozero
      '******* VRBallReel Case.... *****
      VRball1reel.visible = false
      VRball2reel.visible = false
      VRball3reel.visible = true
      VRball4reel.visible = false
      VRball5reel.visible = false

    Case 2:
      ball1reel.resettozero
      ball2reel.setValue(1)
      ball3reel.resettozero
      ball4reel.resettozero
      ball5reel.resettozero
      '******* VRBallReel Case.... *****
      VRball1reel.visible = false
      VRball2reel.visible = true
      VRball3reel.visible = false
      VRball4reel.visible = false
      VRball5reel.visible = false

    Case 1:
      ball1reel.setValue(1)
      ball2reel.resettozero
      ball3reel.resettozero
      ball4reel.resettozero
      ball5reel.resettozero
      '******* VRBallReel Case.... *****
      VRball1reel.visible = true
      VRball2reel.visible = false
      VRball3reel.visible = false
      VRball4reel.visible = false
      VRball5reel.visible = false
  End Select

  ballkickertimer.enabled=false
end sub


Sub AddBonus1Kb(Total)        'Add 1000's in bonus, play sound & update score reel
  BonusCount=0
  BonusCountMax=Total
  If Total=0 then exit sub          'afterbonustimer.enabled=true : ' added this to spit out ball
  BonusTimerb.Interval=400      'Slow start to bonus timer
  BonusTimerb.Enabled=true
End Sub


Sub AddBonus1K(Total)       'Add 1000's in bonus, play sound & update score reel
  BonusCount=0
  BonusCountMax=Total
  If Total=0 then exit sub          'afterbonustimer.enabled=true : ' added this to spit out ball
  BonusTimer.Interval=400     'Slow start to bonus timer
  BonusTimer.Enabled=true
End Sub


' This checks to see if the strike lights are on and adds a value for the bonus hold check
Sub checkstrikelights
  hold1=0
  hold2=0
  hold3=0
  if lightbonus1.state=1 then hold1=1
  if lightbonus2.state=1 then hold2=1
  if lightbonus3.state=1 then hold3=1
End Sub

Sub BonusTimerb_timer()
  PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
  PlaySoundAtLevelStatic "bonus-scan", 1, Bumper1
  AddScore 1000
  BonusCount=BonusCount+1
  'add a space in between 5 counts
  if bonuscount mod 5 =0 then
    bonustimerb.interval=280
  else bonustimerb.interval=155
  end if
  If BonusCount=BonusCountMax then BonusTimerb.Enabled=false  'end the timer when all the bonus is counted
End Sub


Sub BonusTimer_timer()
  PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
  PlaySoundAtLevelStatic "motorstep", 1, Bumper1
  PlaySoundAtLevelStatic "bonus-scan", 1, Bumper1
  AddScore 1000
  BonusCount=BonusCount+1
  'add a space in between 5 counts
  if bonuscount mod 5 =0 then
    bonustimer.interval=280
  else
    bonustimer.interval=155
  end if

  if bonuscount =10 and lightbonus1.state=1 and lightbonus2.state=0 and lightbonus3.state=0 then lightbonus1.state=0
  if bonuscount =10 and lightbonus1.state=1 and lightbonus2.state=1 and lightbonus3.state=0 then lightbonus2.state=0
  if bonuscount =10 and lightbonus1.state=1 and lightbonus2.state=1 and lightbonus3.state=1 then lightbonus3.state=0
  if bonuscount =20 and lightbonus1.state=1 and lightbonus2.state=0 then lightbonus1.state=0
  if bonuscount =20 and lightbonus1.state=1 and lightbonus2.state=1 then lightbonus2.state=0
  if bonuscount =30 and lightbonus1.state=1 then lightbonus1.state=0
  If BonusCount=BonusCountMax then BonusTimer.Enabled=false 'end the timer when all the bonus is counted
  If BonusCount=BonusCountMax then afterbonustimer.enabled=true       ' Timer to shutoff bonus lights and spit out ball
End Sub


Sub Timer2_Timer()
  If credit<>credit2 Then
    credit2=credit
    SetCredits
  End If
End Sub


'************************************************************************************
'START Option Menu Support
Dim operatormenu, options, game_state
Dim balls
Dim replays, liberal, liblight, object
Dim replayscore(4), i
Dim Pin_Opposite
Dim OppositeMap(10)


Sub OptionMenu_Init
  'Check for legal option parameter values
  if balls="" or  (balls<>3 and balls<>5) then balls=3  ' Legal Values 3 or 5    ' first time run will pick 3
  if replays="" or replays<1 or replays>2 then replays=2  ' Legal Values 1 or 2
  if liberal="" or liberal<0 or liberal>2 then liberal=0  ' Legal Values 0 or 1 or 2

  'Replay Score Thresholds
  SetReplayScores
  SetOppositeMap

  RepCard.image = "VRReplayCard"&replays
  OptionBalls.image="OptionsBalls"&balls
  OptionReplays.image="OptionsReplays"&replays
  OptionLiberal.image="OptionsLiberal"&liberal
  if balls=3 then
    InstCard.image="VRInstCard3balls"
  else
    InstCard.image="VRInstCard5balls"
  end if
  game_state=gameon 'Link to State of Play - Menu Not Active During Gameplay
End Sub


Sub OptionMenuKeyDownCheck(keycode)
  game_state=gameon 'Link to State of Play - Menu Not Active During Gameplay
  If keycode=LeftFlipperKey and game_state = false and OperatorMenu=0 then
    OperatorMenuTimer.Enabled = true
  end if

  If keycode=LeftFlipperKey and game_state = false and OperatorMenu=1 then
    Options=Options+1
    If Options=5 then Options=1
    Select Case (Options)
      Case 1:
        Option1.visible=true
        Option4.visible=False
      Case 2:
        Option2.visible=true
        Option1.visible=False
      Case 3:
        Option3.visible=true
        Option2.visible=False
      Case 4:
        Option4.visible=true
        Option3.visible=False
    End Select
  end if
  If keycode=RightFlipperKey and game_state = false and OperatorMenu=1 then
    Select Case (Options)
      Case 1:
        if balls=3 then
          balls=5
          InstCard.image="VRInstCard5balls"
        else
          balls=3
          InstCard.image="VRInstCard3balls"
        end if
        OptionBalls.image = "OptionsBalls"&balls
      Case 2:
        Liberal = Liberal + 1
        If Liberal>2 Then Liberal = 0
        OptionLiberal.image= "OptionsLiberal"&Liberal
      Case 3:
        replays=replays+1
        if replays>2 then
          replays=1
        end if
        SetReplayScores
        OptionReplays.image = "OptionsReplays"&replays
        repcard.image = "VRReplayCard"&replays
      Case 4:
        OperatorMenu=0
        savehs    ' Save Configured Options
        HideOptions
    End Select
  End If
end Sub


Sub OperatorMenuTimer_Timer
  OperatorMenu=1
  Displayoptions
  Options=1
End Sub

Sub DisplayOptions
  OptionsBack.visible = true
  Option1.visible = True
  OptionBalls.visible = True
  OptionReplays.visible = True
  OptionLiberal.visible = True
End Sub

Sub HideOptions
  for each object In OptionMenu
    object.visible = false
  next
End Sub

Sub SetReplayScores
  If replays=2 Then
    replayscore(1) = 80000:replayscore(2) = 126000:replayscore(3) = 157000:replayscore(4) = 200000
  Else
    replayscore(1) = 60000:replayscore(2) = 100000:replayscore(3) = 127000:replayscore(4) = 148000
  End If
End Sub

Sub SetOppositeMap
  OppositeMap(0) = 0
  OppositeMap(1) = 2
  OppositeMap(2) = 1
  OppositeMap(3) = 5
  OppositeMap(4) = 4
  OppositeMap(5) = 3
  OppositeMap(6) = 9
  OppositeMap(7) = 8
  OppositeMap(8) = 7
  OppositeMap(9) = 6
End Sub

'END Option Menu Support
'************************************************************************************

Sub HighScoreResetTimer_Timer()
  ResetHighScore
  HighScoreResetTimer.enabled = false
End Sub


Sub ResetHighScore()
  HighScore = 0
  SaveValue "TripleStrike", "Option1", HighScore
  score = HighScore    ' settings score to Highscore so I can read it for VR reels.   This is now displayed on the desk beside the wall  Restting to 0 below..

  HFR1.ImageA="Trans"
  HFR10.ImageA="Trans"
  HFR100.ImageA="Trans"
  HFR1000.ImageA="Trans"
  HFR10000.ImageA="Trans"
  HFR100000.ImageA="Trans"
End Sub


' VR Clock code below....
Sub ClockTimer_Timer()
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
  Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())
End Sub


' Lavalamp code below.  Thank you STEELY!
Bcnt = 0
For Each Blob in Lspeed
  Lbob(Bcnt) = .2
  Lspeed(Bcnt) = Int((5 * Rnd) + 2) * .05
  Bcnt = Bcnt + 1
Next


Sub LavaTimer_Timer()
  Bcnt = 0
  For Each Blob in Lava
    If Blob.TransZ <= LavaBase.Size_Z * 1.5 Then  'Change blob direction to up
      Lspeed(Bcnt) = Int((5 * Rnd) + 2) * .05   'travel speed
      blobAng(Bcnt) = Int((359 * Rnd) + 1)      'blob location/angle from center
      blobRad(Bcnt) = Int((40 * Rnd) + 10)      'blob distance from center, radius
      blobSiz(Bcnt) = Int((150 * Rnd) + 100)    'blob size
      Blob.Size_x = blobSiz(Bcnt):Blob.Size_y = blobSiz(Bcnt):Blob.Size_z = blobSiz(Bcnt)
      Blob.X = Round(Cos(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + LavaBase.X  'place blob
      Blob.Y = Round(Sin(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + LavaBase.Y
    End If

    If Blob.TransZ => LavaBase.Size_Z*5 Then    'Change blob direction to down
      blobAng(Bcnt) = Int((359 * Rnd) + 1)      'blob location/angle from center
      blobRad(Bcnt) = Int((40 * Rnd) + 10)      'blob distance from center,radius
      Blob.X = Round(Cos(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + LavaBase.X  'place blob
      Blob.Y = Round(Sin(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + LavaBase.Y
      Lspeed(Bcnt) = Int((8 * Rnd) + 5) * .05:Lspeed(Bcnt) = Lspeed(Bcnt) * -1        'travel speed
    End If

    'Make blob wobble
    If Blob.Size_x > blobSiz(Bcnt) + blobSiz(Bcnt)*.15 or Blob.Size_x < blobSiz(Bcnt) - blobSiz(Bcnt)*.15  Then Lbob(Bcnt) = Lbob(Bcnt) * -1
    Blob.Size_x = Blob.Size_x + Lbob(Bcnt)
    Blob.Size_y = Blob.Size_y + Lbob(Bcnt)
    Blob.Size_z = Blob.Size_Z - Lbob(Bcnt) * .66
    Blob.TransZ = Blob.TransZ + Lspeed(Bcnt)    'Move blob
    Bcnt = Bcnt + 1
  Next

End Sub


'********************************VR DRUM SCORING ROUTINES
'*********************************
' ***************************************************************************
'          BASIC FSS(EM) 1-4 player 5x drums, 1 credit drum CORE CODE
' ****************************************************************************

Dim ix, nxp,npp, reels(5, 7), scores(6,2)
'********************for one player games uncomment the next line
Dim player: player = 1
'reset scores to defaults
for nxp =0 to 5
  scores(nxp,0 ) = 0
  scores(nxp,1 ) = 0
Next

'reset EM drums to defaults
For nxp =0 to 3
  For  npp =0 to 6
    reels(nxp, npp) =0 ' default to zero
  Next
Next


Sub SetScore(player, ndx , val)

  Dim ncnt
  if player = 5 or player = 6 then
    if val > 0 then
      If(ndx = 0)Then ncnt = val * 10
      If(ndx = 1)Then ncnt = val

      scores(player, 0) = scores(player, 0) + ncnt
    end if
  else
    if val > 0 then

      If(ndx = -1)then ncnt = val * 100000
      If(ndx = 0)then ncnt = val * 10000
      If(ndx = 1)then ncnt = val * 1000
      If(ndx = 2)Then ncnt = val * 100
      If(ndx = 3)Then ncnt = val * 10
      If(ndx = 4)Then ncnt = val

      scores(player, 0) = scores(player, 0) + ncnt
    end if
  end if
End Sub


Sub SetDrum(player, drum , val)

  If val = 0 then
    Select case player
      case -1: ' the credit drum
        emp4r6.ObjrotX = 0 ' 285
      Case 0:
        Select Case drum
          Case 1: emp1r1.ObjrotX = 0 ' 283
          Case 2: emp1r2.ObjrotX= 0
          Case 3: emp1r3.ObjrotX=0
          Case 4: emp1r4.ObjrotX=0
          Case 5: emp1r5.ObjrotX=0
        End Select
      Case 1:
        Select Case drum
          Case 1: emp2r1.ObjrotX= 0
          Case 2: emp2r2.ObjrotX= 0
          Case 3: emp2r3.ObjrotX= 0
          Case 4: emp2r4.ObjrotX=0
          Case 5: emp2r5.ObjrotX= 0
        End Select
      Case 2:
        Select Case drum
          Case 1: emp3r1.ObjrotX=0
          Case 2: emp3r2.ObjrotX=0
          Case 3: emp3r3.ObjrotX=0
          Case 4: emp3r4.ObjrotX=0
          Case 5: emp3r5.ObjrotX=0
        End Select
      Case 3:
        Select Case drum
          Case 1: emp4r1.ObjrotX=0
          Case 2: emp4r2.ObjrotX=0
          Case 3: emp4r3.ObjrotX=0
          Case 4: emp4r4.ObjrotX=0
          Case 5: emp4r5.ObjrotX=0
        End Select
    End Select

  else
    Select case player

      Case -1: ' the credit drum
        emp4r6.ObjrotX = emp4r6.ObjrotX + val
      Case 0:
        Select Case drum
          Case 1: emp1r1.ObjrotX= emp1r1.ObjrotX + val
          Case 2: emp1r2.ObjrotX= emp1r2.ObjrotX + val
          Case 3: emp1r3.ObjrotX= emp1r3.ObjrotX + val
          Case 4: emp1r4.ObjrotX= emp1r4.ObjrotX + val
          Case 5: emp1r5.ObjrotX= emp1r5.ObjrotX + val
        End Select
      Case 1:
        Select Case drum
          Case 1: emp2r1.ObjrotX= emp2r1.ObjrotX+val
          Case 2: emp2r2.ObjrotX= emp2r2.ObjrotX+val
          Case 3: emp2r3.ObjrotX= emp2r3.ObjrotX+val
          Case 4: emp2r4.ObjrotX= emp2r4.ObjrotX+val
          Case 5: emp2r5.ObjrotX= emp2r5.ObjrotX+val
        End Select
      Case 2:
        Select Case drum
          Case 1: emp3r1.ObjrotX=emp3r1.ObjrotX+ val
          Case 2: emp3r2.ObjrotX=emp3r2.ObjrotX+ val
          Case 3: emp3r3.ObjrotX=emp3r3.ObjrotX+ val
          Case 4: emp3r4.ObjrotX=emp3r4.ObjrotX+ val
          Case 5: emp3r5.ObjrotX=emp3r5.ObjrotX+ val
        End Select
      Case 3:
        Select Case drum
          Case 1: emp4r1.ObjrotX=emp4r1.ObjrotX+ val
          Case 2: emp4r2.ObjrotX=emp4r2.ObjrotX+ val
          Case 3: emp4r3.ObjrotX=emp4r3.ObjrotX+ val
          Case 4: emp4r4.ObjrotX=emp4r4.ObjrotX+ val
          Case 5: emp4r5.ObjrotX= emp4r5.ObjrotX+ val
        End Select

    End Select
  end if
End Sub


Sub SetReel(player, drum, val)

  Dim  inc , cur, dif, fix, fval
  inc = 33.5
  fval = -5 ' graphic seam between 5 & 6 fix value, easier to fix here than photoshop

  If  (player <= 3) or (drum = -1) then
    If drum = -1 then drum = 0
    cur =reels(player, drum)
    If val <> cur then ' something has changed
      Select Case drum

        Case 0: ' credits drum
          if val > cur then
            dif =val - cur
            fix =0
            If cur < 5 and cur+dif > 5 then
              fix = fix- fval
            end if
            dif = dif * inc
            dif = dif-fix
            SetDrum -1,0,  -dif
          Else
            if val = 0 Then
              SetDrum -1,0,  0' reset the drum to abs. zero
            Else
              dif = 11 - cur
              dif = dif + val
              dif = dif * inc
              dif = dif-fval
              SetDrum -1,0,   -dif
            end if
          end if
        Case 1:

          if val > cur then
            dif =val - cur
            fix =0
            If cur < 5 and cur+dif > 5 then
              fix = fix- fval
            end if
            dif = dif * inc
            dif = dif-fix
            SetDrum player,drum,  -dif
          Else
            if val = 0 Then
              SetDrum player,drum,  0' reset the drum to abs. zero
            Else
              dif = 11 - cur
              dif = dif + val
              dif = dif * inc
              dif = dif-fval
              SetDrum player,drum,   -dif
            end if
          end if
          reels(player, drum) = val

        Case 2:

          if val > cur then
            dif =val - cur
            fix =0
            If cur < 5 and cur+dif > 5 then
              fix = fix- fval
            end if
            dif = dif * inc
            dif = dif-fix
            SetDrum player,drum,  -dif
          Else
            if val = 0 Then
              SetDrum player,drum,  0 ' reset the drum to abs. zero
            Else
              dif = 11 - cur
              dif = dif + val
              dif = dif * inc
              dif = dif-fval
              SetDrum player,drum,  -dif
            end if
          end if
          reels(player, drum) = val

        Case 3:

          if val > cur then
            dif =val - cur
            fix =0
            If cur < 5 and cur+dif > 5 then
              fix = fix- fval
            end if
            dif = dif * inc
            dif = dif-fix

            SetDrum player,drum,  -dif
          Else
            if val = 0 Then
              SetDrum player,drum,  0 ' reset the drum to abs. zero
            Else
              dif = 11 - cur
              dif = dif + val
              dif = dif * inc
              dif = dif-fval
              SetDrum player,drum,  -dif
            end if

          end if
          reels(player, drum) = val

        Case 4:

          if val > cur then
            dif =val - cur
            fix =0
            If cur < 5 and cur+dif > 5 then
              fix = fix- fval
            end if
            dif = dif * inc
            dif = dif-fix
            SetDrum player,drum,  -dif
          Else
            if val = 0 Then
              SetDrum player,drum,  0 ' reset the drum to abs. zero
            Else
              dif = 11 - cur
              dif = dif + val
              dif = dif * inc
              dif = dif-fval
              SetDrum player,drum,  -dif
            end if

          end if
          reels(player, drum) = val

        Case 5:

          if val > cur then
            dif =val - cur
            fix =0
            If cur < 5 and cur+dif > 5 then
              fix = fix- fval
            end if
            dif = dif * inc
            dif = dif-fix
            SetDrum player,drum,  -dif
          Else
            if val = 0 Then
              SetDrum player,drum,  0 ' reset the drum to abs. zero
            Else
              dif = 11 - cur
              dif = dif + val
              dif = dif * inc
              dif = dif-fval
              SetDrum player,drum,  -dif
            end if

          end if
          reels(player, drum) = val
      End Select

    end if
  end if
End Sub

Dim Score1, Score10,Score100,Score1000,Score10000,Score100000, Scoreimage
Dim ActivePLayer, nplayer,playr,value, stemp, cred

'**********  Scoring

sub addscoreVR(points)
  if points=10 or points=100 or points=1000 or points=10000 then   ' HAVING ISSUE HERE WITH 10,000 POINT SPECIAL Lane light....  fixed.. added or points=10000
    addpointsVR Points

    ''************************** I am not sure if below is doing anything?  We dont have these timeres set as variables? eg. AddScore10Timer   *****************************************************
  else
    If Points < 100 and AddScore10Timer.enabled = false Then
      Add10 = Points \ 10
      vrAddScore10Timer.Enabled = TRUE
    ElseIf Points < 1000 and AddScore100Timer.enabled = false Then
      Add100 = Points \ 100
      vrAddScore100Timer.Enabled = TRUE
    ElseIf AddScore1000Timer.enabled = false Then
      Add1000 = Points \ 1000
      vrAddScore1000Timer.Enabled = TRUE
    End If
  End If
End Sub

Sub vrAddScore10Timer_Timer()
  if Add10 > 0 then
    AddPointsVR 10
    Add10 = Add10 - 1
  Else
    Me.Enabled = FALSE
  End If
End Sub

Sub vrAddScore100Timer_Timer()
  if Add100 > 0 then
    AddPointsVR 100
    Add100 = Add100 - 1
  Else
    Me.Enabled = FALSE
  End If
End Sub

Sub vrAddScore1000Timer_Timer()
  if Add1000 > 0 then
    AddPointsVR 1000
    Add1000 = Add1000 - 1
  Else
    Me.Enabled = FALSE
  End If

  '******************************* DO WE NEED TO ADD 10,000 here above also?  For the Special light in left lane? ********************************

End Sub

Sub AddPointsVR(Points)       ' Sounds: there are 3 sounds: tens, hundreds and thousands

  ActivePLayer = 0

  Score100000 =0
  Score10000 =0
  Score1000 =0
  Score100 =0
  Score10 =0
  Score1 =0

  value = Score

  ' do 100 thousands thousands
  if(value >= 500000)  then:  Score100000 = 5 : value = value - 500000: end if
  if(value >= 400000)  then:  Score100000 = 4 : value = value - 400000: end if
  if(value >= 300000)  then:  Score100000 = 3 : value = value - 300000: end if
  if(value >= 200000)  then:  Score100000 = 2 : value = value - 200000: end if
  if(value >= 100000)  then:  Score100000 = 1 : value = value - 100000: end if


  ' do ten thousands
  if(value >= 90000)  then:  Score10000 =9 : value = value - 90000: end if
  if(value >= 80000)  then:  Score10000 =8 : value = value - 80000: end if
  if(value >= 70000)  then:  Score10000 =7 : value = value - 70000: end if
  if(value >= 60000)  then:  Score10000 =6 : value = value - 60000: end if
  if(value >= 50000)  then:  Score10000 =5 : value = value - 50000: end if
  if(value >= 40000)  then:  Score10000 =4 : value = value - 40000: end if
  if(value >= 30000)  then:  Score10000 =3 : value = value - 30000: end if
  if(value >= 20000)  then:  Score10000 =2 : value = value - 20000: end if
  if(value >= 10000)  then:  Score10000 =1 : value = value - 10000: end if

  ' do thousands
  if(value >= 9000)  then:  Score1000 =9 : value = value - 9000: end if
  if(value >= 8000)  then:  Score1000 =8 : value = value - 8000: end if
  if(value >= 7000)  then:  Score1000 =7 : value = value - 7000: end if
  if(value >= 6000)  then:  Score1000 =6 : value = value - 6000: end if
  if(value >= 5000)  then:  Score1000 =5 : value = value - 5000: end if
  if(value >= 4000)  then:  Score1000 =4 : value = value - 4000: end if
  if(value >= 3000)  then:  Score1000 =3 : value = value - 3000: end if
  if(value >= 2000)  then:  Score1000 =2 : value = value - 2000: end if
  if(value >= 1000)  then:  Score1000 =1 : value = value - 1000: end if

  ' do hundreds
  if(value >= 900)  then:  Score100 =9 : value = value - 900: end if
  if(value >= 800)  then:  Score100 =8 : value = value - 800: end if
  if(value >= 700)  then:  Score100 =7 : value = value - 700: end if
  if(value >= 600)  then:  Score100 =6 : value = value - 600: end if
  if(value >= 500)  then:  Score100 =5 : value = value - 500: end if
  if(value >= 400)  then:  Score100 =4 : value = value - 400: end if
  if(value >= 300)  then:  Score100 =3 : value = value - 300: end if
  if(value >= 200)  then:  Score100 =2 : value = value - 200: end if
  if(value >= 100)  then:  Score100 =1 : value = value - 100: end if

  ' do tens
  if(value >= 90)  then:  Score10 =9 : value = value - 90: end if
  if(value >= 80)  then:  Score10 =8 : value = value - 80: end if
  if(value >= 70)  then:  Score10 =7 : value = value - 70: end if
  if(value >= 60)  then:  Score10 =6 : value = value - 60: end if
  if(value >= 50)  then:  Score10 =5 : value = value - 50: end if
  if(value >= 40)  then:  Score10 =4 : value = value - 40: end if
  if(value >= 30)  then:  Score10 =3 : value = value - 30: end if
  if(value >= 20)  then:  Score10 =2 : value = value - 20: end if
  if(value >= 10)  then:  Score10 =1 : value = value - 10: end if
  ' do ones
  if(value >= 9)  then:  Score1 =9 : value = value - 9: end if
  if(value >= 8)  then:  Score1 =8 : value = value - 8: end if
  if(value >= 7)  then:  Score1 =7 : value = value - 7: end if
  if(value >= 6)  then:  Score1 =6 : value = value - 6: end if
  if(value >= 5)  then:  Score1 =5 : value = value - 5: end if
  if(value >= 4)  then:  Score1 =4 : value = value - 4: end if
  if(value >= 3)  then:  Score1 =3 : value = value - 3: end if
  if(value >= 2)  then:  Score1 =2 : value = value - 2: end if
  if(value >= 1)  then:  Score1 =1 : value = value - 1: end if

  scores(0,1) = scores(0,0)
  scores(0,0) = 0
  scores(1,1) = scores(1,0)
  scores(1,0) = 0
  scores(2,1) = scores(2,0)
  scores(2,0) = 0
  scores(3,1) = scores(3,0)
  scores(3,0) = 0

  For  ix =0 to 6
    reels(0, ix) =0
    reels(1, ix) =0
    reels(2, ix) =0
    reels(3, ix) =0
  Next

  For  ix =0 to 4'4

    SetDrum ix, 1 , 0
    SetDrum ix, 2 , 0
    SetDrum ix, 3 , 0
    SetDrum ix, 4 , 0
    SetDrum ix, 5 , 0
  Next

  For playr =0 to 3'3

    If (ActivePLayer) = playr Then
      nplayer = playr

      SetScore nplayer,-1, Score100000 ' store 100K+ score no reel
      SetReel nplayer, 1 , Score10000 : SetScore nplayer,0,Score10000
      SetReel nplayer, 2 , Score1000 : SetScore nplayer,1,Score1000
      SetReel nplayer, 3 , Score100 : SetScore nplayer,2,Score100
      SetReel nplayer, 4 , Score10 : SetScore nplayer,3,Score10
      SetReel nplayer, 5 , 0 : SetScore nplayer,4,0 ' assumes ones position is always zero

    else
      nplayer = playr
      value =scores(nplayer, 1)

      ' do hundred thousands

      SetScore nplayer,-1, stemp ' store 100K+ score no reel

      ' do ten thousands
      if(value >= 90000)  then:  SetReel nplayer, 1 , 9 : SetScore nplayer,0,9 : value = value - 90000: end if
      if(value >= 80000)  then:  SetReel nplayer, 1 , 8 : SetScore nplayer,0,8 : value = value - 80000: end if
      if(value >= 70000)  then:  SetReel nplayer, 1 , 7 : SetScore nplayer,0,7 : value = value - 70000: end if
      if(value >= 60000)  then:  SetReel nplayer, 1 , 6 : SetScore nplayer,0,6 : value = value - 60000: end if
      if(value >= 50000)  then:  SetReel nplayer, 1 , 5 : SetScore nplayer,0,5 : value = value - 50000: end if
      if(value >= 40000)  then:  SetReel nplayer, 1 , 4 : SetScore nplayer,0,4 : value = value - 40000: end if
      if(value >= 30000)  then:  SetReel nplayer, 1 , 3 : SetScore nplayer,0,3 : value = value - 30000: end if
      if(value >= 20000)  then:  SetReel nplayer, 1 , 2 : SetScore nplayer,0,2 : value = value - 20000: end if
      if(value >= 10000)  then:  SetReel nplayer, 1 , 1 : SetScore nplayer,0,1 : value = value - 10000: end if

      ' do thousands
      if(value >= 9000)  then:  SetReel nplayer, 2 , 9 : SetScore nplayer,1,9 : value = value - 9000: end if
      if(value >= 8000)  then:  SetReel nplayer, 2 , 8 : SetScore nplayer,1,8 : value = value - 8000: end if
      if(value >= 7000)  then:  SetReel nplayer, 2 , 7 : SetScore nplayer,1,7 : value = value - 7000: end if
      if(value >= 6000)  then:  SetReel nplayer, 2 , 6 : SetScore nplayer,1,6 : value = value - 6000: end if
      if(value >= 5000)  then:  SetReel nplayer, 2 , 5 : SetScore nplayer,1,5 : value = value - 5000: end if
      if(value >= 4000)  then:  SetReel nplayer, 2 , 4 : SetScore nplayer,1,4 : value = value - 4000: end if
      if(value >= 3000)  then:  SetReel nplayer, 2 , 3 : SetScore nplayer,1,3 : value = value - 3000: end if
      if(value >= 2000)  then:  SetReel nplayer, 2 , 2 : SetScore nplayer,1,2 : value = value - 2000: end if
      if(value >= 1000)  then:  SetReel nplayer, 2 , 1 : SetScore nplayer,1,1 : value = value - 1000: end if

      'do hundreds
      if(value >= 900) then: SetReel nplayer, 3 , 9 : SetScore nplayer,2,9 : value = value - 900: end if
      if(value >= 800) then: SetReel nplayer, 3 , 8 : SetScore nplayer,2,8 : value = value - 800: end if
      if(value >= 700) then: SetReel nplayer, 3 , 7 : SetScore nplayer,2,7 : value = value - 700: end if
      if(value >= 600) then: SetReel nplayer, 3 , 6 : SetScore nplayer,2,6 : value = value - 600: end if
      if(value >= 500) then: SetReel nplayer, 3 , 5 : SetScore nplayer,2,5 : value = value - 500: end if
      if(value >= 400) then: SetReel nplayer, 3 , 4 : SetScore nplayer,2,4 : value = value - 400: end if
      if(value >= 300) then: SetReel nplayer, 3 , 3 : SetScore nplayer,2,3 : value = value - 300: end if
      if(value >= 200) then: SetReel nplayer, 3 , 2 : SetScore nplayer,2,2 : value = value - 200: end if
      if(value >= 100) then: SetReel nplayer, 3 , 1 : SetScore nplayer,2,1 : value = value - 100: end if

      'do tens
      if(value >= 90) then: SetReel nplayer, 4 , 9 : SetScore nplayer,3,9 : value = value - 90: end if
      if(value >= 80) then: SetReel nplayer, 4 , 8 : SetScore nplayer,3,8 : value = value - 80: end if
      if(value >= 70) then: SetReel nplayer, 4 , 7 : SetScore nplayer,3,7 : value = value - 70: end if
      if(value >= 60) then: SetReel nplayer, 4 , 6 : SetScore nplayer,3,6 : value = value - 60: end if
      if(value >= 50) then: SetReel nplayer, 4 , 5 : SetScore nplayer,3,5 : value = value - 50: end if
      if(value >= 40) then: SetReel nplayer, 4 , 4 : SetScore nplayer,3,4 : value = value - 40: end if
      if(value >= 30) then: SetReel nplayer, 4 , 3 : SetScore nplayer,3,3 : value = value - 30: end if
      if(value >= 20) then: SetReel nplayer, 4 , 2 : SetScore nplayer,3,2 : value = value - 20: end if
      if(value >= 10) then: SetReel nplayer, 4 , 1 : SetScore nplayer,3,1 : value = value - 10: end if

      'do ones
      if(value >= 9) then: SetReel nplayer, 5 , 9 : SetScore nplayer,4,9 : value = value - 9: end if
      if(value >= 8) then: SetReel nplayer, 5 , 8 : SetScore nplayer,4,8 : value = value - 8: end if
      if(value >= 7) then: SetReel nplayer, 5 , 7 : SetScore nplayer,4,7 : value = value - 7: end if
      if(value >= 6) then: SetReel nplayer, 5 , 6 : SetScore nplayer,4,6 : value = value - 6: end if
      if(value >= 5) then: SetReel nplayer, 5 , 5 : SetScore nplayer,4,5 : value = value - 5: end if
      if(value >= 4) then: SetReel nplayer, 5 , 4 : SetScore nplayer,4,4 : value = value - 4: end if
      if(value >= 3) then: SetReel nplayer, 5 , 3 : SetScore nplayer,4,3 : value = value - 3: end if
      if(value >= 2) then: SetReel nplayer, 5 , 2 : SetScore nplayer,4,2 : value = value - 2: end if
      if(value >= 1) then: SetReel nplayer, 5 , 1 : SetScore nplayer,4,1 : value = value - 1: end if
    end if
  Next
end sub


'**** reset score by counting down
'***** reset score

Dim tenk, onek, hun, tens, rst:

sub vrResetScores()

  dim i
  'reset scores to defaults
  for nxp =0 to 5
    scores(nxp,0 ) = 0
    scores(nxp,1 ) = 0
  Next

  rst=0
  resettimer.enabled=true

end sub

sub resettimer_timer


  if score => 100000 Then
    do until score < 100000
      score = score - 100000
    Loop
  end If

  rst=rst+1

  if score>0 then

    tenk = Int(score/10000)
    onek = Int((score-tenk*10000)/1000)
    hun = Int((score-tenk*10000-onek*1000)/100)
    tens = Int((score-tenk*10000-onek*1000-hun*100)/10)

  end if

  if tenk>0 then tenk=tenk-1 else tenk=0: end if
  if onek>0 then onek=onek-1 else onek=0: end if
  if hun>0 then hun=hun-1 else hun=0: end if
  if tens>0 then tens=tens-1 else tens=0: end if

  score=tenk*10000+onek*1000+hun*100+tens*10

  Setreel 0, 1 , tenk
  Setreel 0, 2 , onek
  Setreel 0, 3 , hun
  Setreel 0, 4 , tens
  Setreel 0, 5 , 0

  if rst=18 then
    resettimer.enabled=false
  end if

end sub


Sub vrScoreLight(nr, st)
  dim q, r
  for r = 1 to 5   'if less then 4 players comment the appropriate lines
    EVAL("emp1r"&r).Image = "ScoreReel"
    EVAL("emp2r"&r).Image = "ScoreReel"
    EVAL("emp3r"&r).Image = "ScoreReel"
    EVAL("emp4r"&r).Image = "ScoreReel"
  Next
  if st = 1 then
    Select Case nr
        case 1 :  for r = 1 to 5
          EVAL("emp1r"&r).Image = "ScoreReelOn"
        Next
        case 2 :  for r = 1 to 5
          EVAL("emp2r"&r).Image = "ScoreReelOn"
        Next
        case 3 :  for r = 1 to 5
          EVAL("emp3r"&r).Image = "ScoreReelOn"
        Next
        case 4 :  for r = 1 to 5
          EVAL("emp4r"&r).Image = "ScoreReelOn"
        Next
    end select
  end if
end Sub


Sub vrSetCredit(credit)
  if credit > 9 then credit = 9
  cred =reels(4, 0)
  reels(4, 0) = 0
  SetDrum -1,0,  0

  SetReel 0,-1,  credit
  reels(4, 0) = credit
End Sub


Sub VRcredLight(state)
  Select Case state
    case 0 :  emp4r6.Image = "ScoreReel"
    case 1 :  emp4r6.Image = "ScoreReelOn"
  end Select
End sub


'***************************<<***end vr drums code
' ****************************************************************************


'************************* NEW BALL SHADOW CODE - Thanks VPW Team! ***************

Sub FrameTimer_Timer()
  BSUpdate
End Sub


'***************************************************************
' ZSHA: Ambient ball shadows
'***************************************************************

' For dynamic ball shadows, Check the "Raytraced ball shadows" box for the specific light.
' Also make sure the light's z position is around 25 (mid ball)
Const tnob = 2
Const lob = 0

'Ambient (Room light source)
Const AmbientBSFactor = 0.9    '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(7)

'Initialization
BSInit

Sub BSInit()
  Dim iii
  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 3 + iii / 1000
    objBallShadow(iii).visible = 0
  Next
End Sub


Sub BSUpdate
  Dim s: For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow

    'Primitive shadow on playfield, flasher shadow in ramps
    '** If on main and upper pf
    If gBOT(s).Z > 20 And gBOT(s).Z < 30 Then
      objBallShadow(s).visible = 1
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
    '** No shadow if ball is off the main playfield (this may need to be adjusted per table)
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub


'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************

'******************************************************
'****  BALL ROLLING AND DROP SOUNDS
'******************************************************

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


Sub RollingBall_timer()

  Dim b

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 To tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(gBOT)
    If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ <  - 1 And gBOT(b).z < 55 And gBOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz >  - 7 Then
          RandomSoundBallBouncePlayfieldSoft gBOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard gBOT(b)
        End If
      End If
    End If

    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If
  Next
End Sub


' I had to add this trigger due to weird reflection anomolies on the kicker area.
Sub RefTrigger_hit()
  Table1.BallReflection = 0
End Sub

Sub RefTrigger_unhit()
  Table1.BallReflection = 1
End Sub


' make bubbles random Speed
Sub BeerTimer_Timer()

  Randomize(21)
  BeerBubble1.z = BeerBubble1.z + Rnd(1)*0.5
  if BeerBubble1.z > -771 then BeerBubble1.z = -955

  BeerBubble2.z = BeerBubble2.z + Rnd(1)*1
  if BeerBubble2.z > -768 then BeerBubble2.z = -955

  BeerBubble3.z = BeerBubble3.z + Rnd(1)*1
  if BeerBubble3.z > -768 then BeerBubble3.z = -955

  BeerBubble4.z = BeerBubble4.z + Rnd(1)*0.75
  if BeerBubble4.z > -774 then BeerBubble4.z = -955

  BeerBubble5.z = BeerBubble5.z + Rnd(1)*1
  if BeerBubble5.z > -771 then BeerBubble5.z = -955

  BeerBubble6.z = BeerBubble6.z + Rnd(1)*1
  if BeerBubble6.z > -774 then BeerBubble6.z = -955

  BeerBubble7.z = BeerBubble7.z + Rnd(1)*0.8
  if BeerBubble7.z > -768 then BeerBubble7.z = -955

  BeerBubble8.z = BeerBubble8.z + Rnd(1)*1
  if BeerBubble8.z > -771 then BeerBubble8.z = -955
End Sub


Sub SetFlipperimage()

  if FlipperImage = 1 then
    LFLogo.image = "Flipper-r1"
    RFLogo.image = "Flipper-r1"
  End If

  if FlipperImage = 2 then
    LFLogo.image = "Flipper-r2"
    RFLogo.image = "Flipper-r2"
  End If

  if FlipperImage = 3 then
    LFLogo.image = "Flipper-r3"
    RFLogo.image = "Flipper-r3"
  End If

  if FlipperImage = 4 then
    LFLogo.image = "Flipper-r4"
    RFLogo.image = "Flipper-r4"
  End If

  if FlipperImage = 5 then
    LFLogo.image = "Flipper-r5"
    RFLogo.image = "Flipper-r5"
  End If
End Sub


' tiered Flipper rubber hit sounds based on ball speed...
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall

  finalspeed=BallSpeed(theball)
  if finalspeed > 4 and finalspeed < 8 then PlaySoundAtLevelStatic "FlipperrubberSoft", 1, LFLogo
  if finalspeed > 8 then PlaySoundAtLevelStatic "FlipperrubberHard", 1, LFLogo
End Sub


Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall

  finalspeed=BallSpeed(theball)
  if finalspeed > 4 and finalspeed < 8 then PlaySoundAtLevelStatic "FlipperrubberSoft", 1, RFLogo
  if finalspeed > 8 then PlaySoundAtLevelStatic "FlipperrubberHard", 1, RFLogo
End Sub


Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled) 'Left flipper solenoid callback
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire  'leftflipper.rotatetoend

  Else
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled) 'Right flipper solenoid callback
  If Enabled Then
    FlipperActivate RightFlipper, RFPress
    RF.Fire 'rightflipper.rotatetoend

  Else
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
  End If
End Sub



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


'*******************************************
' Late 70's to early 80's

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
Const EOSReturn = 0.045  'late 70's to mid 80's

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
    Dim b', BOT


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
  LS.Object = target169
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = target153
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00, - 2
  AddSlingsPt 1, 0.45, - 5
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  2
  AddSlingsPt 5, 1.00,  5
End Sub

Sub AddSlingsPt(idx, aX, aY)    'debugger wrapper for adjusting flipper script In-game
  Dim a
  a = Array(LS, RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub

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

  If tmp > 0 Then
    AudioFade = CSng(tmp ^ 10)
  Else
    AudioFade = CSng( - (( - tmp) ^ 10) )
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

  If tmp > 0 Then
    AudioPan = CSng(tmp ^ 10)
  Else
    AudioPan = CSng( - (( - tmp) ^ 10) )
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
  PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 8) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
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
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
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
  PlaySoundAtLevelStatic "SaucerEnter2", SaucerLockSoundLevel, ActiveBall
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


' *****  END !!!!! ....omg  *********************************








