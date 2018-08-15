' Triple Strike (Williams 1975)
' - VPX Conversion

Option Explicit
Randomize

' Thalamus 2018-08-15 : Improved directional sounds
' You need to import standard fx_ballrolling sounds and create a timer named RollingTimer

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 200

' Load the controller.vbs file for DOF
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs."
On Error Goto 0

ExecuteGlobal GetTextFile("core.vbs")
If Err Then MsgBox "Unable to open core.vbs."
On Error Goto 0

' Constants
' BackGlass ROM IDs
Const BIP=32
Const TILT_ID=33
Const MATCH_ID=34
Const GAMEOVER_ID=35
Const SAMESHOOTER=36
Const HISCORE=50

' Variables
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
Dim addball
Dim cGameName, B2SName
Dim DesktopMode: DesktopMode = Table1.ShowDT

cGameName = "triplestrike"   'assumes:  romname entry in directouputconfig*.ini
B2SName = "triplestrike"     'assumes:  B2STableSettings.xml has startup error messages set to 0
ballinplay=false


Sub Table1_Init()
  LoadEM
  loadStoredDataTimer.Enabled=True
  gameon=False
  addball=0
  ballinplay=false

  If Table1.ShowDT = false then
	'Cabinet Mode
	for each object in backdropstuff : object.visible = 0 : next
  Else
	'Desktop Mode
    for each object in backdropstuff : object.visible = 1 : next
  End If

  Set mHole = New cvpmMagnet	'Upper Hole (Using low powered Magnet to simulate drop in
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
  For each obj in AllLights	'Turn all the lights off at table initialisation
     obj.State=0
  Next
  Loops=0						'Used to loop sounds

 'Init Bumper rings
  ringa.IsDropped=1:ringb.IsDropped=1:ringc.IsDropped=1

  postitval=0
  Tilts = 0

  SetBackGlass 45, 1	' Start BG light animation
End Sub


Sub Table1_Exit
  savehs
  If B2SOn Then Controller.Stop
End Sub

Sub savehs
  'Highscore = 0: Postitscore = 0  ' FOR TESTING: Resets the saved value
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
  credit=LoadValue("TripleStrike", "Option3")
  If credit="" Then
    credit=2
    SaveValue "TripleStrike", "Option3", credit
  End If
  Balls=LoadValue("TripleStrike", "Option4")
  If Balls="" Then
    Balls=5
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
  backglassTimer.interval = 500
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
  backglassTimer.interval = 5000
  If B2SOn Then
    Controller.B2SSetScorePlayer1 0
    Controller.B2SSetScorePlayer1 HighScore
    SetBackGlass HISCORE, 1
  End If
  EMReel1.resettozero
  EMReel1.addvalue(highscore)
  hiscorereel.setvalue(1)
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
    If credit > 0 then
        creditlight.state=1:DOF 120, DOFOn
    Else
        creditlight.state=0:DOF 120, DOFOff
    End If
End Sub


Sub UpdateFlipperLogo_Timer
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
	'LFLogoUP.RotY = LeftFlipper1.CurrentAngle
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
      ' turn off bumpers
      bumper1.force=0
      target153.disabled=true
      target169.disabled=true
 	  gameover()
      End If
    End If
End Sub

Sub MechCheckTilt
      Tilted = TRUE
      PlaySound "tilt"
      SetBackGlass TILT_ID, 1
      tiltreel.setvalue(1)
      ' turn off bumpers
      bumper1.force=0
      target153.disabled=true
      target169.disabled=true
 	  gameover()
End Sub

Sub TiltTimer_Timer()
    If Tilts > 0 then
       Tilts = Tilts - 1
    Else
       TiltTimer.Enabled = FALSE
    End If
End Sub


Sub Reset()		'Called at table initialisation and game start
  If B2SOn Then
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScorePlayer1 0
  End If

  Randomize
  match= int(10*rnd)*10
  matchtext.text = match

  Score = 0
  ball = 0

  knockered=0
  tilted = false

  scoretext.text = formatnumber(score,0,-1,0,-1)
  balltext.text = formatnumber(ball,0,-1,0,-1)

  For each obj in AllLights	'Set all lights off
  	obj.State=0
  Next
  For each obj in PinLights	'Set all pin lights on
  	obj.State=1
  Next
  For each obj in DTargLights	'Set all Drop Target Lights On
  	obj.State=1
  Next
  For each obj in DTargs		'Raise all Drop Targets
  	obj.IsDropped=false
  Next

  bumper1.force = 8
  target153.disabled=false
  target169.disabled=false

' StepB2SData 0, 5, 0,2 , 50, ""
  SetBackGlass TILT_ID, 0
  SetBackGlass GAMEOVER_ID, 1
  SetBackGlass BIP, 0
  SetBackGlass SAMESHOOTER, 0
  SetBackGlass MATCH_ID, 0

  emreel1.resettozero
  tiltreel.resettozero
  gameoverreel.setvalue(1)
  hiscorereel.setvalue(1)
  sameplayerreel.resettozero
  ball1reel.resettozero
  ball2reel.resettozero
  ball3reel.resettozero
  ball4reel.resettozero
  ball5reel.resettozero
  postitreel.setvalue(0)

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

End Sub


Sub StartGame()
  backglassTimer.Enabled=False
  Reset
  ball = Balls	'Option Menu Link - Number of Balls
  balltext.text = formatnumber(ball,0,-1,0,-1)
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
  hiscorereel.resettozero
  sameplayerreel.resettozero
  Select Case ball
      Case 5:
          ball1reel.resettozero
          ball2reel.resettozero
          ball3reel.resettozero
          ball4reel.resettozero
          ball5reel.setValue(1)
      Case 4:
          ball1reel.resettozero
          ball2reel.resettozero
          ball3reel.resettozero
          ball4reel.setValue(1)
          ball5reel.resettozero
      Case 3:
          ball1reel.resettozero
          ball2reel.resettozero
          ball3reel.setValue(1)
          ball4reel.resettozero
          ball5reel.resettozero
      Case 2:
          ball1reel.resettozero
          ball2reel.setValue(1)
          ball3reel.resettozero
          ball4reel.resettozero
          ball5reel.resettozero
      Case 1:
          ball1reel.setValue(1)
          ball2reel.resettozero
          ball3reel.resettozero
          ball4reel.resettozero
          ball5reel.resettozero
  End Select
' ball3reel.setvalue(1)

  Playsound "startup3"
  PlaySound "lightreset"  ' Sound for light reset
End Sub

' General function to add score, updates the score reel
Sub AddScore(points)
  Score=Score+points
  scoretext.text = formatnumber(score,0,-1,0,-1)
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
  If liberal=1 or knockered=1 Then
      addball=addball+1:extraballlight.state=lightstateon
      If ball<5 Then ball = ball+1
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
  PlaySound SoundFXDOF("bumper1", 101, DOFPulse, DOFContactors)
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
  PlaySound SoundFXDOF("slingshotl", 105, DOFPulse, DOFContactors)
  AddScore 10
End Sub

' Right Slingshot
Sub Target153_slingshot()
  If Tilted = true then exit sub
  PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
  PlaySound SoundFXDOF("slingshotl", 102, DOFPulse, DOFContactors)
  AddScore 10
End Sub

Sub PointShots_Hit(T)
  If Tilted = true then exit sub
  PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
  AddScore 10
End Sub

'Ball @ Plunger
Sub trigger11_hit()
  If addball>0 then
     extraballlight.state=lightstateon
     SetBackGlass SAMESHOOTER, 1
     sameplayerreel.setvalue(1)
  End If
  DOF 119, DOFOn
  DOF 122, DOFPulse
End sub

Sub trigger11_unhit()
  Playsound "plunger-full-t2"
End sub

Sub trigger14_hit()
  DOF 119, DOFOff
End sub

' Sub BallRollSounds_hit(T) 'ball rolling sounds
'   PlaySound "ballrolling"
' End Sub


' Rail hit sounds
Sub wall159_hit
     speedx=ballspeed.velx: speedy=ballspeed.vely
  	finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)
  	if finalspeed > 2 then PlaySound "metalhit"
end sub

Sub wall133_hit
     speedx=ballspeed.velx: speedy=ballspeed.vely
  	finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)
  	if finalspeed > 2 then PlaySound "metalhit"
end sub


' gate sounds
Sub gate2_hit:PlaySoundAt "gate",gate2:end sub
Sub gate3_hit:PlaySoundAt "gate",gate3:end sub

' rubber sounds
Sub WallHit(finalspeedcheck)
 	speedx=ballspeed.velx: speedy=ballspeed.vely
  	finalspeed=SQR(ballspeed.velx * ballspeed.velx + ballspeed.vely * ballspeed.vely)
  	if finalspeed > finalspeedcheck then PlaySound "rubber" else PlaySound "rubberS":end if
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
  PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
  PinLights(T).state=0
  If liberal=2 Then                    ' Option Menu Link
      Pin_Opposite = OppositeMap(T)    ' T comes in as the pin number - 1
      PinLights(Pin_Opposite).state=0  ' PinLights requires pin number - 1
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
  PlaySound SoundFX("targetdrop",DOFDropTargets)
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


Sub checkextraball()	' check extra ball
 	For each obj in DTargLights
  		If obj.state=1 Then Exit Sub
  	Next
  	If leftlight.state=lightstateon then
       DOF 108, DOFPulse:addball = addball + 1
       If ball<5 Then ball = ball+1
       extraballlight.state=lightstateon
       SetBackGlass SAMESHOOTER, 1
       sameplayerreel.setvalue(1)
    End If
End Sub

Sub checkabcd()	' Check abcd lights
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


Sub checktargets()	' check targets
  	For each obj in DTargs
  		If obj.IsDropped=False Then Exit Sub
  	Next
  	dtargettimer.enabled=true
end sub


' Added timer for abcd target reset
Sub dtargettimer_timer
    For each obj in DTargs		'Raise all Drop Targets
  	    obj.IsDropped=false
    Next
    SetBackGlass 41, 0:SetBackGlass 42, 0:SetBackGlass 43, 0:SetBackGlass 44, 0
    PlaySound SoundFX("dropreset", DOFDropTargets)
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
        AddBonus1Kb 5	' OK I know it's not a bonus, but it uses the exact same code
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
    if toplight2.state=lightstateon and ball=1 then
      DOF 108, DOFPulse:addball = addball + 1
      If ball<5 Then ball = ball+1
      extraballlight.state=lightstateon
      SetBackGlass SAMESHOOTER, 1
      sameplayerreel.setvalue(1)
    end if
    if toplight1.state=lightstateon then toplight2.state=lightstateon
End Sub

Sub LoopSound()		'Used to loop the bell sound 5 times when the ball lands in the kicker
  	Loops=0
  	LoopSoundTimer.Interval=400
  	LoopSoundTimer.Enabled=true
End Sub

Sub LoopSoundTimer_Timer()	'Used to loop the bell sound 5 times when the ball lands in the kicker
    PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
 	Playsound "bonus-scan"

  	LoopSoundTimer.Interval=80
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
 	If keycode = PlungerKey Then Plunger.PullBack: End If

    If keycode = AddcreditKey Then PlaySound "CoinIn": knockertimer.enabled=true: End If

    ' Option Menu Link - Flipper Control
    If gameon=false Then
        OptionMenuKeyDownCheck(keycode)
    Else

  	  If keycode = LeftFlipperKey Then
 		If GameOn = TRUE then
            If Tilted = FALSE Then
                LeftFlipper.RotateToEnd
                PlaySound SoundFXDOF("Flipperup2", 103, DOFOn, DOFFlippers)
                playSound "Flipperbuzz2"
            end if
        end if
      end if

 	  If keycode = RightFlipperKey Then
        If GameOn = TRUE then
            If Tilted = FALSE then
                RightFlipper.RotateToEnd
                PlaySound SoundFXDOF("Flipperup2", 104, DOFOn, DOFFlippers)
                playSound "Flipperbuzz1"
            End If
        end if
      end if

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
        If keycode = CenterTiltKey and tilted = false Then
            Nudge 0, 1.08
            tiltcheck
        End If
 	  End If

	If keycode = MechanicalTilt Then
		mechchecktilt
	End If

    End If
End Sub


Sub Table1_KeyUp(ByVal keycode)
 	If keycode = PlungerKey Then Plunger.Fire

  	If keycode = StartGameKey and gameon=false and ball <= 0 and credit >0  Then
        gameon=true
        StartGame
  	End If

 	If keycode = LeftFlipperKey Then

        ' Option Menu Link - Timer Turn Off
        OperatorMenuTimer.Enabled = false

        If GameOn = TRUE then
		  If Tilted = FALSE then
 		    LeftFlipper.RotateToStart
            PlaySound SoundFXDOF("Flipperdown2", 103, DOFOff, DOFFlippers)
 		    stopSound "Flipperbuzz2"
 	      end if
        end if
    end if
 	If keycode = RightFlipperKey Then
       If GameOn = TRUE then
		If Tilted = FALSE then
 		  RightFlipper.RotateToStart
          PlaySound SoundFXDOF("Flipperdown2", 104, DOFOff, DOFFlippers)
 		  stopSound "Flipperbuzz1"
 	    end if
      end if
    end if

End Sub


Sub matchplay
  if match = 0 then match = 100
  SetBackGlass MATCH_ID, match
  Select Case match
      Case 100: emreel100.setvalue(1)
      Case 10:  emreel110.setvalue(1)
      Case 20:  emreel120.setvalue(1)
      Case 30:  emreel130.setvalue(1)
      Case 40:  emreel140.setvalue(1)
      Case 50:  emreel150.setvalue(1)
      Case 60:  emreel160.setvalue(1)
      Case 70:  emreel170.setvalue(1)
      Case 80:  emreel180.setvalue(1)
      Case 90:  emreel190.setvalue(1)
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
  else
 	singlepinbonus
    checkstrikelights
  end if
  ballinplay=false
  Drain.DestroyBall
  PlaySound "drain-t1"
  if addball>0 then addball=addball-1
  ball = ball - 1
  balltext.text = formatnumber(ball,0,-1,0,-1)
End Sub


Sub gameover()
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
 	If lightbonus3.state=1 then		'Check bonus lights & add 10K, 20K or 30K bonus as appropriate
  		 BonusScore=BonusScore+30
  	Else
  		if lightbonus2.state=1 then
  			BonusScore=BonusScore+20
 		Else
  			if lightbonus1.state=1 then BonusScore=BonusScore+10
  		End If
  	End If
 	AddBonus1K BonusScore			'Run add bonus routine (plays sounds & updates score reel)
end sub


Sub afterbonustimer_timer
  	bonusholdcheck
    playsound "lightreset"

 	toplight1.state=0				'If not game over then reset lights
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
    If addball=0 Then
        SetBackGlass SAMESHOOTER, 0
        sameplayerreel.resettozero
    End If
  	If Ball <= 0 then
    	gameover()
 	    stopsound "flipperbuzz2"
        SetBackGlass GAMEOVER_ID, 1
        gameoverreel.setvalue(1)
        Topscore
    End If
    If Ball > 0 then ballkickertimer.enabled=true
    afterbonustimer.enabled=false
End Sub


sub topscore
  if clng (score) > clng (highscore) then
  	highscore=score:postitscore=postitval
    savehs
 	textbox1.text=highscore
  End if
  backglassTimer.Interval=5000
  backglassTimer.Enabled=True
end sub


Sub bonusholdcheck()				'Only reset pin lights if bonus hold is not lit
  If bonuslight.state=0 and Ball > 0 then  'only if bonus light is off, pinlights and strikes turn off
  		For each obj in PinLights
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


Sub ballkickertimer2_timer
   if ball =>1 then
  	 	set ballspeed=kicker2.createball
 		ballinplay=true                       ' longer starup for first ball
 		kicker2.kick 30,17
 	    ballkickertimer2.enabled=false
        PlaySound SoundFXDOF("loader-ballroll", 109, DOFPulse, DOFContactors)
 	End If
end sub


sub ballkickertimer_timer
    if ball =>1 then
        set ballspeed=kicker2.createball
        ballinplay=true' named ballspeed for speed equations later
        kicker2.kick 30,17
        Playsound "eject"
        PlaySound SoundFXDOF("loader-ballroll", 109, DOFPulse, DOFContactors)
    End If
    SetBackGlass BIP, ball
    Select Case ball
      Case 5:
          ball1reel.resettozero
          ball2reel.resettozero
          ball3reel.resettozero
          ball4reel.resettozero
          ball5reel.setValue(1)
      Case 4:
          ball1reel.resettozero
          ball2reel.resettozero
          ball3reel.resettozero
          ball4reel.setValue(1)
          ball5reel.resettozero
      Case 3:
          ball1reel.resettozero
          ball2reel.resettozero
          ball3reel.setValue(1)
          ball4reel.resettozero
          ball5reel.resettozero
      Case 2:
          ball1reel.resettozero
          ball2reel.setValue(1)
          ball3reel.resettozero
          ball4reel.resettozero
          ball5reel.resettozero
      Case 1:
          ball1reel.setValue(1)
          ball2reel.resettozero
          ball3reel.resettozero
          ball4reel.resettozero
          ball5reel.resettozero
    End Select
    ballkickertimer.enabled=false
end sub


Sub AddBonus1Kb(Total)				'Add 1000's in bonus, play sound & update score reel
  	BonusCount=0
  	BonusCountMax=Total
  	'if total=0 then BonusTimer.Enabled=true
  	If Total=0 then exit sub          'afterbonustimer.enabled=true : ' added this to spit out ball
  	BonusTimerb.Interval=400			'Slow start to bonus timer
  	BonusTimerb.Enabled=true
End Sub


Sub AddBonus1K(Total)				'Add 1000's in bonus, play sound & update score reel
  	BonusCount=0
  	BonusCountMax=Total
  	'if total=0 then BonusTimer.Enabled=true
  	If Total=0 then exit sub          'afterbonustimer.enabled=true : ' added this to spit out ball
  	BonusTimer.Interval=400			'Slow start to bonus timer
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
        bonustimerb.interval=220
    else bonustimerb.interval=80
    end if
  	If BonusCount=BonusCountMax then BonusTimerb.Enabled=false	'end the timer when all the bonus is counted
End Sub


Sub BonusTimer_timer()
  	PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
 	Playsound "motorstep"
  	Playsound "bonus-scan"
 	AddScore 1000
 	BonusCount=BonusCount+1

    'add a space in between 5 counts
    if bonuscount mod 5 =0 then
        bonustimer.interval=180
    else
        bonustimer.interval=80
    end if

    if bonuscount =10 and lightbonus1.state=1 and lightbonus2.state=0 and lightbonus3.state=0 then lightbonus1.state=0
    if bonuscount =10 and lightbonus1.state=1 and lightbonus2.state=1 and lightbonus3.state=0 then lightbonus2.state=0
    if bonuscount =10 and lightbonus1.state=1 and lightbonus2.state=1 and lightbonus3.state=1 then lightbonus3.state=0

    if bonuscount =20 and lightbonus1.state=1 and lightbonus2.state=0 then lightbonus1.state=0
    if bonuscount =20 and lightbonus1.state=1 and lightbonus2.state=1 then lightbonus2.state=0

    if bonuscount =30 and lightbonus1.state=1 then lightbonus1.state=0

    If BonusCount=BonusCountMax then BonusTimer.Enabled=false	'end the timer when all the bonus is counted
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
	if balls="" or  (balls<>3 and balls<>5) then balls=5	' Legal Values 3 or 5
	if replays="" or replays<1 or replays>2 then replays=2	' Legal Values 1 or 2
	if liberal="" or liberal<0 or liberal>2 then liberal=0	' Legal Values 0 or 1 or 2

   'Replay Score Thresholds
    SetReplayScores
    SetOppositeMap

	RepCard.image = "ReplayCard"&replays
	OptionBalls.image="OptionsBalls"&balls
	OptionReplays.image="OptionsReplays"&replays
	OptionLiberal.image="OptionsLiberal"&liberal
	if balls=3 then
		InstCard.image="InstCard3balls"
	  else
		InstCard.image="InstCard5balls"
	end if
    game_state=gameon	'Link to State of Play - Menu Not Active During Gameplay
End Sub


Sub OptionMenuKeyDownCheck(keycode)
    game_state=gameon	'Link to State of Play - Menu Not Active During Gameplay
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
				InstCard.image="InstCard5balls"
			  else
				balls=3
				InstCard.image="InstCard3balls"
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
			repcard.image = "ReplayCard"&replays
		Case 4:
			OperatorMenu=0
			savehs		' Save Configured Options
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
       replayscore(1) = 60000:replayscore(2) = 106000:replayscore(3) = 127000:replayscore(4) = 148000
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


' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

Sub PlaySoundAtVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

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

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
  Vol = Vol + 3
  debug.print vol
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function


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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

