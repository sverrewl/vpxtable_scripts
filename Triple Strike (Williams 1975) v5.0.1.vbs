
' Triple Strike 5.0 (Williams 1975) - Nov 2021 - Rawd - Hybrid release

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


' This is a complete overhaul of the table. Hybrid version, New plastics, new lights and Dynamic shadowing, playfield meshes, new sounds, physics tweaks, new VR room, code completely cleaned and restructured


' Lots of people have had their hand in this table now, and it shows.  Thanks everyone!
'*************************************************************************************************************************************************************
'*************************************************************************************************************************************************************


Option Explicit
Randomize

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

 Dim   VRRoom: VRRoom = 0           '0 = Desktop or cabinet,      1 = VR
 Const DynamicBallShadowsOn = 1   '0 = no dynamic ball shadows, 1 = enable dynamic ball shadow
 Const AmbientShadowOn = 1        '0 = no regular ball shadow   1 = regular ball shadow on
 GlassImpurities.visible = 1        '0 = Glass scratches off      1 = Glass scratches on - This will only work in VR mode.  It will default to OFF in Desktop or FS
 Dim FlipperImage: FlipperImage = 3 ' Use 1 through 5.  Can be hard set here, and swapped on the fly with left Magnasave

'***** END OPTIONS ***********************************************************************************************************************
'*****************************************************************************************************************************************
'*****************************************************************************************************************************************




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
Dim ballspeed    'name of the ball when regarding speedx and speedy
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
SetVRRoomSign.visible = true ' Instruction sign in VR mode to set VR Room = 1
for each Object in VRStuff: object.visible = 0: next
for each Object in FakeBulbs: object.state = 0: next
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

    'Below is doing a check for people running desktop, and turning B2S OFF, Some have B2S enabled for other tables that require it on desktop (PUP)
      if B2sOn = true then
    Controller.LaunchBackglass=0
    B2sOn=False
      end if
  End If

  Set mHole = New cvpmMagnet  'Upper Hole (Using low powered Magnet to simulate drop in
  With mHole                    'playfield surface around the saucer)/Thanks JPSalas!
      .InitMagnet Umagnet, 2.85
      .GrabCenter = 0
      .MagnetOn = 1
      .CreateEvents "mHole"
  End With

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
  BG.Image="NewTSBGFinalOff"  ' resets the VR bG to have the 100,000 point light OFF
  ' reset Postit here...
  Postit.ImageA="PostItBlank"
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


' Loads the saved values, sets the backglass credits and initializes the Option Menu
Sub loadStoredDataTimer_Timer()
  loadStoredDataTimer.Enabled=False
  loadStoredDataTimer.Interval=1000
  LoadHighScore
  SetCredits
  OptionMenu_Init
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


Sub UpdateFlipperLogo_Timer
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
End Sub


Sub TiltCheck
  If Tilted = FALSE then
    If Tilts = 0 then
      Tilts = Tilts + int(rnd*200)
      TiltTimer.Enabled = TRUE
    Else
      Tilts = Tilts + int(rnd*200)
    End If
    If Tilts >= 300 and Tilted = FALSE then
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
       Tilts = Tilts - 1
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

 For each obj in PinLights2 'Set all pin lights on
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

  bumper1.force = 8
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
  'emreel1.resettozero
  tiltreel.resettozero
  VRTiltReel.visible = false
  gameoverreel.setvalue(1)
  ' ****  Below sets Gameovereeel on VR backglass ****
  VRGameOver.visible = true
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
  'balltext.text = formatnumber(ball,0,-1,0,-1)
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
  PlaySound SoundFXDOF("fpknocker", 108, DOFPulse, DOFKnocker)
  If liberal=1 and extraballlight.state=lightstateoff or knockered=1 and extraballlight.state=lightstateoff Then  ' added and statement here..

     extraballlight.state=lightstateon
     ball = ball+1
      SetBackGlass SAMESHOOTER, 1
  End If
  If knockered<>1 Then
      credit = credit + 1:creditreel.addvalue(1):SetCredits
  End If
End Sub

'Used by Match
Sub Add_credit
  PlaySound SoundFXDOF("fpknocker", 108, DOFPulse, DOFKnocker)
  credit = credit + 1:creditreel.addvalue(1):SetCredits
End Sub


Sub bumper1_hit() 'Top bumper
  If Tilted = true then exit sub
  PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
  PlaySoundAtVol SoundFXDOF("bumper1", 101, DOFPulse, DOFContactors), ActiveBall, 1
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
  PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
  PlaySoundAtVol SoundFXDOF("slingshotl", 105, DOFPulse, DOFContactors), ActiveBall, 1
  AddScore 10
End Sub


' Right Slingshot
Sub Target153_slingshot()
  If Tilted = true then exit sub
  PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
  PlaySoundAtVol SoundFXDOF("slingshotl", 102, DOFPulse, DOFContactors), ActiveBall, 1
  AddScore 10
End Sub


Sub PointShots_Hit(T)
  If Tilted = true then exit sub
  PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
  AddScore 10
End Sub


'Ball @ Plunger
Sub trigger11_hit()
  BallSpeed.BulbIntensityScale = 0 '- allows to dampen/scale the intensity of (bulb-)light reflections on each ball (e.g. to simulate shadowing in the ball lane, etc)
  DOF 119, DOFOn
  DOF 122, DOFPulse
End sub

Dim SoundReady : SoundReady = true

Sub SoundTrigger_Hit()
SoundReady = true
End Sub

Sub trigger11_unhit()
If SoundReady = True then Playsound "plunger-full-t2"
SoundReady = false
End sub


' Rail hit sounds
Sub wall159_hit
     speedx=ballspeed.velx: speedy=ballspeed.vely
    finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)
    if finalspeed > 2 and finalspeed < 6 then PlaySoundAtVol "metalhitsoft", ActiveBall, 1
    if finalspeed > 6 then PlaySoundAtVol "metalhit", ActiveBall, 1
end sub


Sub wall133_hit
     speedx=ballspeed.velx: speedy=ballspeed.vely
    finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)
    if finalspeed > 2 and finalspeed < 6 then PlaySoundAtVol "metalhitsoft", ActiveBall, 1
    if finalspeed > 6 then PlaySoundAtVol "metalhit", ActiveBall, 1
end sub


'Drain Walls sounds..
Sub wall121_hit
     speedx=ballspeed.velx: speedy=ballspeed.vely
    finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)
    if finalspeed > 2 and finalspeed <6 then PlaySoundAtVol "metalhitsoft", ActiveBall, 1
    if finalspeed > 6 then PlaySoundAtVol "metalhit", ActiveBall, 1
end sub

Sub wall123_hit
     speedx=ballspeed.velx: speedy=ballspeed.vely
    finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)
    if finalspeed > 2 and finalspeed <6 then PlaySoundAtVol "metalhitsoft", ActiveBall, 1
    if finalspeed > 6 then PlaySoundAtVol "metalhit", ActiveBall, 1
end sub

'Side Walls sounds..
Sub LeftWood_hit
     speedx=ballspeed.velx: speedy=ballspeed.vely
    finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)
    if finalspeed > 3 and finalspeed <9 then PlaySoundAtVol "WoodSoft", ActiveBall, 1
    if finalspeed > 9 then PlaySoundAtVol "WoodLoud", ActiveBall, 1
end sub

Sub RightWood_hit
     speedx=ballspeed.velx: speedy=ballspeed.vely
    finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)
    if finalspeed > 3 and finalspeed <9 then PlaySoundAtVol "WoodSoft", ActiveBall, 1
    if finalspeed > 9 then PlaySoundAtVol "WoodLoud", ActiveBall, 1
end sub


' gate sounds
Sub gate2_hit:PlaySoundAtVol "gate", ActiveBall, 1:end sub
Sub gate3_hit:PlaySoundAtVol "gate", ActiveBall, 1:end sub

' rubber sounds
Sub WallHit(finalspeedcheck)
  speedx=ballspeed.velx: speedy=ballspeed.vely
    finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)
    if finalspeed > finalspeedcheck then PlaySoundAtVol "rubber", ActiveBall, 1 else PlaySoundAtVol "rubberS", ActiveBall, 1:end if
End Sub

Sub wall140_hit:WallHit(11):End Sub
Sub wall1_hit:WallHit(11):End Sub
Sub wall2_hit:WallHit(11):End Sub
Sub wall5_hit:WallHit(11):End Sub
Sub wall146_hit:WallHit(12):End Sub
Sub wall39_hit:WallHit(12):End Sub
Sub wall51_hit:WallHit(12):End Sub
Sub wall45_hit:WallHit(10):End Sub
Sub wall59_hit:WallHit(10):End Sub
Sub wall168_hit:WallHit(12):End Sub
Sub wall169_hit:WallHit(12):End Sub
Sub wall176_hit:WallHit(12):End Sub
Sub wall172_hit:WallHit(12):End Sub
Sub wall177_hit:WallHit(12):End Sub
Sub wall175_hit:WallHit(12):End Sub
Sub wall170_hit:WallHit(12):End Sub
Sub wall174_hit:WallHit(12):End Sub
Sub wall20_hit:WallHit(12):End Sub
Sub wall178_hit:WallHit(12):End Sub
Sub wall179_hit:WallHit(12):End Sub
Sub wall100_hit:WallHit(12):End Sub
Sub wall157_hit:WallHit(12):End Sub
Sub wall180_hit:WallHit(12):End Sub
Sub wall181_hit:WallHit(12):End Sub

' triggers and lights for pins
' Turn off the light of the pin that was hit
' Adds 10 points regardless if light is on or off.
Sub PinTriggers_Hit(T)
  If Tilted = true then exit sub
  PlaySoundAtVol SoundFXDOF("bell10",141,DOFPulse,DOFChimes), ActiveBall, 1
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
  PlaySoundAtVol SoundFX("targetdrop",DOFDropTargets), ActiveBall, 1
  DTargs(T).isdropped=true
  DTargLights(T).state=lightstateoff
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
    SetBackGlass 41, 0:SetBackGlass 42, 0:SetBackGlass 43, 0:SetBackGlass 44, 0
    PlaySoundAtVol SoundFX("dropreset", DOFDropTargets), targetc, 1
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
    dim speedx
    dim speedy
    dim finalspeed
    speedx=ballspeed.velx: speedy=ballspeed.vely
    finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)
  If Tilted = true then
    kicker1.timerenabled=true
    exit sub
    end if
    if finalspeed > 11 then
      Kicker1.kick 0,0
      ballspeed.vely=speedy   ' If the ball is going faster than '10',
      ballspeed.velx=speedx   ' then continue on its path and bump it up in the air speed 9, and hit glass
    ballspeed.velz=10
    playsound "collide6"
    end if
    if finalspeed > 8 then
      Kicker1.kick 0,0
      ballspeed.vely=speedy  ' If the ball is going faster than '8', then continue on its path and bump
      ballspeed.velx=speedx  ' it up in the air speed 4
    ballspeed.velz=5
    else
      kicker1.timerenabled=true
    if tenxlight.state=lightstateoff then

        LoopSound
    end if
    if tenxlight.state=lightstateon then
        AddBonus1Kb 5 ' OK I know it's not a bonus, but it uses the exact same code
      end if
    checktoplights  ' HAD to make a new addbonus sub because it conflicted with ball release.
    toplight1.state=lightstateon
    end if
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
  Playsound "bonus-scan"

    LoopSoundTimer.Interval=155  '' This is what times the space between bells.
    If loops=4 then LoopSoundTimer.Enabled=false
    'if loops=0 then AddScore (500):Score=Score+500:scoretext.text = formatnumber(score,0,-1,0,-1): end if
    'Gave a pause to EM reel for 500 points in top kicker
    Loops=Loops+1
    AddScore (100)
End Sub


' knocker timer for delayed knocker on coin-in  also delayed credit reel animation
Sub knockertimer_timer
  PlaySound SoundFXDOF("knocker-t1", 108, DOFPulse, DOFKnocker)
  credit = credit + 1:creditreel.addvalue(1):SetCredits
  knockertimer.enabled=false
End Sub


Sub kicker1_timer
  kicker1.kick 189+INT(RND*9),16+INT(RND*4)
    Playsound SoundFXDOF("eject",107,DOFPulse,DOFContactors)
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

''added below for cab users with real physical tilt mechs...
If keycode = MechanicalTilt Then
Tilts = 800
tiltcheck
End If

If keycode = StartGameKey Then
    FrontButton.y = FrontButton.y -5              ' VR Start Button Animation here
    End If

    If keycode = AddcreditKey Then PlaySoundAtVol "CoinIn", Drain, 1: knockertimer.enabled=true: End If
    ' Option Menu Link - Flipper Control
    If gameon=false Then
        OptionMenuKeyDownCheck(keycode)
    Else

      If keycode = LeftFlipperKey Then
    If GameOn = TRUE then
            If Tilted = FALSE Then
                LeftFlipper.RotateToEnd
                PlaySoundAtVol SoundFXDOF("Flipperup2", 103, DOFOn, DOFFlippers), LeftFlipper, 1
                PlayLoopSoundAtVol "Flipperbuzz2", LeftFlipper, 1
                FlipperButtonLeft.X = 2112.9 + 9   ' VR Flipper animation
            end if
        end if
      end if

    If keycode = RightFlipperKey Then
        If GameOn = TRUE then
            If Tilted = FALSE then
                RightFlipper.RotateToEnd
                PlaySoundAtVol SoundFXDOF("Flipperup2", 104, DOFOn, DOFFlippers), RightFlipper, 1
                PlayLoopSoundAtVol "Flipperbuzz1", RightFlipper, 1
                FlipperButtonRight.X = 2104 - 9   ' VR Flipper animation
            End If
        end if
      end if

' Below was added for Pinball cabinets that use VP's fake tilt.  (no real tilt bob)

 If keycode = CenterTiltKey and tilted = false and ballinplay=true Then

      Tilts = 800
            tiltcheck
        End If

      If keycode = LeftTiltKey and tilted = false and ballinplay=true Then      'and ballspeed.x >500 ????
    if ballspeed.y <500 then
            nudge 90,.45
            tiltcheck
    else
            Nudge 90, 1
            tiltcheck
        End If
    end if
      If keycode = RightTiltKey and tilted = false and ballinplay=true Then
    if ballspeed.y <500 then
            Nudge 270, .45
            tiltcheck
        else
            Nudge 270, 1
            tiltcheck
        End If
    End If
    End If
End Sub


Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
Plunger.Fire
TimerVRPlunger.Enabled = False  'VR Plunger
TimerVRPlunger2.Enabled = True   ' VR Plunger
TSPlunger.Y = 975   ' VR Plunger
PlaySoundAtVol "Plunger", Plunger, 1
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
        LeftFlipper.RotateToStart
            FlipperButtonLeft.x = 2112.9   'Resets VR flipper button
            PlaySoundAtVol SoundFXDOF("Flipperdown2", 103, DOFOff, DOFFlippers), LeftFlipper, 1
        stopSound "Flipperbuzz2"
        end if
        end if
    end if
  If keycode = RightFlipperKey Then
       If GameOn = TRUE then
    If Tilted = FALSE then
      RightFlipper.RotateToStart
          FlipperButtonRight.x = 2104   'Resets VR flipper button
          PlaySoundAtVol SoundFXDOF("Flipperdown2", 104, DOFOff, DOFFlippers), RIghtFlipper, 1
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
  Drain.DestroyBall
  PlaySoundAtVol "drain-t1", Drain, 1
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
  LeftFlipper.RotateToStart
  RightFlipper.RotateToStart
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
    PlaySound"motorstep"
    Playsound "bonus-scan"
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
BallSpeed.BulbIntensityScale = 1
DOF 119, DOFOff  ' turns off DOF lights in plunger lane.
end sub


Sub ballkickertimer2_timer
   if ball =>1 then
      set ballspeed=kicker2.createball
    ballinplay=true                       ' longer starup for first ball
    kicker2.kick 30,17
      ballkickertimer2.enabled=false
        PlaySoundAtVol SoundFXDOF("loader-ballroll", 109, DOFPulse, DOFContactors), kicker2, 1
  End If
end sub


sub ballkickertimer_timer
    if ball =>1 then
        set ballspeed=kicker2.createball
        ballinplay=true' named ballspeed for speed equations later
        kicker2.kick 30,17
        PlaySoundAtVol "eject", kicker2, 1
        PlaySound SoundFXDOF("loader-ballroll", 109, DOFPulse, DOFContactors)
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
  Playsound "bonus-scan"
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
  Playsound "motorstep"
    Playsound "bonus-scan"
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

'Sub ResetHighScore()
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

    ' Ive split the 2 different shadows into seperate subs below so they can be run independantly..
  If DynamicBallShadowsOn=1 Then DynamicBSUpdate 'update Dynamic ball shadows
    if AmbientShadowOn=1 Then AmbientBSUpdate ' run regular shadow
    FlipperLsh.rotz= LeftFlipper.currentangle
    FlipperRsh.rotz= RightFlipper.currentangle
End Sub


Const tnob = 1 ' total number of balls
Const lob = 0 'locked balls on start; might need some fiddling depending on how your locked balls are done

' *** Shadow Options ***
Const fovY          = -2  'Offset y position under ball to account for layback or inclination (more pronounced need further back, -2 seems best for alignment at slings)
Const DynamicBSFactor     = 0.95  '0 to 1, higher is darker, 1 will always be maxed even with 2 sources
Const AmbientBSFactor     = 0.7 '0 to 1, higher is darker
Const Wideness        = 20  'Sets how wide the shadows can get (20 +5 thinness should be most realistic)
Const Thinness        = 5   'Sets minimum as ball moves away from source
' ***        ***

Dim sourcenames, currentShadowCount
sourcenames = Array ("","","","","","","","","","","","")
currentShadowCount = Array (0,0,0,0,0,0,0,0,0,0,0,0)
dim objrtx1(20), objrtx2(20)
dim objBallShadow(20)
DynamicBSInit

sub DynamicBSInit()
  Dim iii

  for iii = 0 to tnob                 'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = iii/1000 + 0.01
    objrtx1(iii).visible = 0
    'objrtx1(iii).uservalue=0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = (iii)/1000 + 0.02
    objrtx2(iii).visible = 0
    'objrtx2(iii).uservalue=0
    currentShadowCount(iii) = 0
    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    objBallShadow(iii).Z = iii/1000 + 0.04
  Next
end sub


Sub AmbientBSUpdate()

  Dim ShadowOpacity, ShadowOpacity2
  Dim s, Source, LSd, b, currentMat, AnotherSource, BOT
  BOT = GetBalls
  If UBound(BOT) = lob - 1 Then Exit Sub    'No balls in play, exit

'The Magic happens here
  For s = lob to UBound(BOT)

' *** Normal "ambient light" ball shadow

      If BOT(s).X < tablewidth/2 Then
        objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/13)) + 5
      Else
        objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/13)) - 5
      End If
      objBallShadow(s).Y = BOT(s).Y + fovY

  If BOT(s).Z < 30 Then 'or BOT(s).Z > 105 Then   'Defining when (height-wise) you want ambient shadows
        objBallShadow(s).visible = 1
      Else
        objBallShadow(s).visible = 0
      end if
next
End Sub


Sub DynamicBSUpdate
  Dim falloff:  falloff = 150     'Max distance to light sources, can be changed if you have a reason
  Dim ShadowOpacity, ShadowOpacity2
  Dim s, Source, LSd, b, currentMat, AnotherSource, BOT
  BOT = GetBalls

  If UBound(BOT) = lob - 1 Then Exit Sub    'No balls in play, exit

'The Magic happens here
  For s = lob to UBound(BOT)

' *** Dynamic shadows
    For Each Source in DynamicSources
      LSd=DistanceFast((BOT(s).x-Source.x),(BOT(s).y-Source.y)) 'Calculating the Linear distance to the Source
      If BOT(s).Z < 30 Then 'Or BOT(s).Z > 105 Then       'Defining when (height-wise) you want dynamic shadows
        If LSd < falloff and Source.state=1 Then          'If the ball is within the falloff range of a light and light is on
          currentShadowCount(s) = currentShadowCount(s) + 1 'Within range of 1 or 2
          if currentShadowCount(s) = 1 Then         '1 dynamic shadow source
            sourcenames(s) = source.name
            currentMat = objrtx1(s).material
            objrtx2(s).visible = 0 : objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
            objrtx1(s).rotz = AnglePP(Source.x, Source.y, BOT(s).X, BOT(s).Y) + 90
            ShadowOpacity = (falloff-LSd)/falloff                 'Sets opacity/darkness of shadow by distance to light
            objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness           'Scales shape of shadow with distance/opacity
            UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^2,RGB(0,0,0),0,0,False,True,0,0,0,0
            currentMat = objBallShadow(s).material
            UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-ShadowOpacity),RGB(0,0,0),0,0,False,True,0,0,0,0

          Elseif currentShadowCount(s) = 2 Then
                                'Same logic as 1 shadow, but twice
            currentMat = objrtx1(s).material
            set AnotherSource = Eval(sourcenames(s))
            objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
            objrtx1(s).rotz = AnglePP(AnotherSource.x, AnotherSource.y, BOT(s).X, BOT(s).Y) + 90
            ShadowOpacity = (falloff-(((BOT(s).x-AnotherSource.x)^2+(BOT(s).y-AnotherSource.y)^2)^0.5))/falloff
            objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
            UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0

            currentMat = objrtx2(s).material
            objrtx2(s).visible = 1 : objrtx2(s).X = BOT(s).X : objrtx2(s).Y = BOT(s).Y + fovY
            objrtx2(s).rotz = AnglePP(Source.x, Source.y, BOT(s).X, BOT(s).Y) + 90
            ShadowOpacity2 = (falloff-LSd)/falloff
            objrtx2(s).size_y = Wideness*ShadowOpacity2+Thinness
            UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity2*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
            currentMat = objBallShadow(s).material
            UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
          end if
        Else
          currentShadowCount(s) = 0
        End If
      Else                  'Hide dynamic shadows everywhere else
        objrtx2(s).visible = 0 : objrtx1(s).visible = 0
      End If
    Next
  Next
End Sub


Function DistanceFast(x, y)
  dim ratio, ax, ay
  'Get absolute value of each vector
  ax = abs(x)
  ay = abs(y)
  'Create a ratio
  ratio = 1 / max(ax, ay)
  ratio = ratio * (1.29289 - (ax + ay) * ratio * 0.29289)
  if ratio > 0 then
    DistanceFast = 1/ratio
  Else
    DistanceFast = 0
  End if
end Function

Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function
              'Enable these functions if they are not already present elswhere in your table
Dim PI: PI = 4*Atn(1)

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    end if
  ElseIf dx = 0 Then
    if dy = 0 Then
      Atn2 = 0
    else
      Atn2 = Sgn(dy) * pi / 2
    end if
  End If
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

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

Dim RollingSoundFactor
RollingSoundFactor = 1.1/5


Sub RollingBall_timer()
'textbox1.text = balls
'TextBox001.text=ballcount
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If
  Next
End Sub


'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / tableheight-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = Csng((ball.velz) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  PitchPlayfieldRoll = BallVel(ball) ^2 * 15
End Function

Function RndInt(min, max)
  RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
  RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'************************************ End Fleep Sound Functions *******************************************************


' I had to add this trigger due to weird reflection anomolies on the kicker area.
Sub RefTrigger_hit()
Table1.BallReflection = 0
BallShadow0.z = BallShadow0.z -20
End Sub

Sub RefTrigger_unhit()
Table1.BallReflection = 1
BallShadow0.z = BallShadow0.z +20
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
    speedx=ballspeed.velx: speedy=ballspeed.vely
    finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)

    if finalspeed > 4 and finalspeed < 8 then PlaySoundAtVol "FlipperrubberSoft", ActiveBall, 1
    if finalspeed > 8 then PlaySoundAtVol "FlipperrubberHard", LeftFlipper, 1
End Sub


Sub RightFlipper_Collide(parm)
    speedx=ballspeed.velx: speedy=ballspeed.vely
    finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)

    if finalspeed > 4 and finalspeed < 8 then PlaySoundAtVol "FlipperrubberSoft", ActiveBall, 1
    if finalspeed > 8 then PlaySoundAtVol "FlipperrubberHard", RightFlipper, 1
End Sub


' *****  END !!!!! ....omg  *********************************

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

' Const Pi = 3.1415927

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

