Option Explicit
Randomize


Const cGameName = "King Tut"
Const B2STableName = "King Tut"
Dim HSFileName
HSFileName="King Tut.txt"

' Basic DOF config
'   Note. There is two King Tut tables.
'   These are in this script so best arrange a candidate config or liase with admins to get King Tut 69 listed
'
'   101 Left Flipper, 102 Right Flipper
'   103 Left Sling, 104 Right Sling
'   105 Bumper1, 106 Bumper2, 107 Bumper3, 108 Bumper 4
'   111 kicker
'   125 Ball Release
'   128 Knocker
'   131 gate
'   142 Chime 10 & 100
'   170 - 171 Mushrooms




Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    On Error Resume Next
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
End Sub

LoadCoreFiles

Dim Controller    ' B2S
Dim B2SScore    ' B2S Score Displayed
Dim B2SOn
B2Son = true

'********************  Flasher position  *******************

Sub SetBackglass()

Dim object

For Each object In Backglass_flashers

object.x = object.x - 0

object.height = - object.y - 140

object.y = -20   'adjusts the distance from the backglass towards the user

Next
  KI_flash.y = KI_flash.y + 5
  NG_flash.y = NG_flash.y + 5
  TUT_flash.y = TUT_flash.y + 5
  faro_flash.y = faro_flash.y + 5
  creditblock.y = creditblock.y - 5
End Sub

'----- General Sound Options -----
Const VolumeDial = 0.8        'Recommended values should be no greater than 1.
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1
Const volmaster = 1         'Recommended values should be no greater than 1.
Dim volchimes, volbumpers, volmotor
volchimes = 0.1           'Recommended values should be no greater than 1.
volbumpers = 0.1          'Recommended values should be no greater than 1.
volmotor = 0.1            'Recommended values should be no greater than 1.

Dim RollingSoundFactor

RollingSoundFactor = 1.1/5

Dim sounddamp: sounddamp = 0.3
dim holdrelay

holdrelay = true    ' set to false to turn off the audio

'----- Phsyics Mods -----
' Enhance micro bounces on flippers
Const FlipperCoilRampupMode = 1     '0 = fast, 1 = medium, 2 = slow (tap passes should work)

Dim smartflippers, authentic
smartflippers = true    ' set to false if you don't want to use flipper corrections

authentic = true      ' set to false if you want the game to power on automatically

'**********  Table options    ********************

Dim VRroom
dim desktopmode, cabinetmode, FSSmode
Dim forcollections

desktopmode = Table1.ShowDT
FSSmode =  Table1.ShowFSS

If RenderingMode=2 or FSSmode Then
  VRroom = true
  FSSmode = true
  desktopmode = false
  cabinetmode = false
  SetBackglass
  Set Controller = CreateObject("b2s.server")
  controller.launchbackglass= false
  b2son=false
  Set Controller = CreateObject("VPinMAME.Controller")
  For Each forcollections in HSrampsFC: forcollections.visible = False: Next
end If

If RenderingMode<>2 and not FSSmode Then
  For Each forcollections in VRstuff: forcollections.visible = False: Next
  For Each forcollections in HSrampsothers: forcollections.visible = False: Next
  Set Controller = CreateObject("b2s.server")
  controller.launchbackglass=1
  b2son=true

  if desktopmode then
    cabinetmode = false
    CabRailleft.visible = 1
    CabRailright.visible = 1
  Else
    cabinetmode = true
    CabRailleft.visible = 0
    CabRailright.visible = 0
  end If
end if

Dim scoretotal, score1, score2, score3, score4, score5, hsscore
Dim match, credits
Dim ballinplay
dim cycle
dim replay1, replay2, replay3, replay4
dim Replay1Paid, Replay2Paid, Replay3Paid, Replay4Paid
dim InProgress: InProgress = false
dim matchposition
dim liberal
dim ballstarted: ballstarted = 0
dim tiltgame: tiltgame = false        ' change to true and change instruction card if game tilts
dim specialslit


liberal = 1 ' if conservative, set to 0

ballinplay = 0
cycle = 0

replay1 = 3800
replay2 = 4200
replay3 = 4600
replay4 = 5000

Replay1Paid=False
Replay2Paid=False
Replay3Paid=False
Replay4Paid=False


Dim bgpos
Dim dooralreadyopen




'******* Table Initialisation*******

Sub Table1_Init

Dim Obj

LoadEM

loadhs
  for each obj in kingtut: obj.state=0: next
  specialslit=0
  holdr.enabled = holdrelay
  if creditreel.objrotx="" then creditreel.objrotx= 57.24
  if Credits = "" then credits = 0
  if match  = "" then match = 10
  if score1 =  "" then score1 = 0
  if score2 =  "" then score2 = 0
  if score3 =  "" then score3 = 0
  if score4 =  "" then score4 = 0
  if score5 =  "" then score5 = 0
  if HSScore =  "" then HSScore = 0
  if matchposition = "" then matchposition = 0

  rscore1.objrotx =  108 + score1 * 36
  rscore2.objrotx =  108 + score2 * 36
  rscore3.objrotx =  108 + score3 * 36
  rscore4.objrotx =  108 + score4 * 36

If B2SOn then
  Controller.B2SSetreel 1, score1
  Controller.B2SSetreel 2, score2
  Controller.B2SSetreel 3, score3
  Controller.B2SSetreel 4, score4
  Controller.B2SSetGameOver 0
  Controller.B2SSetTilt 0

  Controller.B2SSetMatch 0
  Controller.B2SSetData 97, 0   'Tut
  Controller.B2SSetData 98, 0   'ng
  Controller.B2SSetData 99, 0   'Ki
  Controller.B2SSetData 90, 0   'Bally
  Controller.B2SSetData 91, 0   'Faro
  Controller.B2SSetBallInPlay ballinplay

  if credits <10 then Controller.B2SSetCredits credits else Controller.B2SSetCredits 9 end if

end if

  SetHSLine 1, "HIGH SCORE"
  SetHSLine 2, HSScore

  for each obj in bottgate
    obj.isdropped=true
    next
  dooralreadyopen = 0
    bgpos=6
    bottgate(bgpos).isdropped=false
  delaystart.enabled = 1
End Sub

sub delaystart_timer
  Turnlightson
  delaystart.enabled = 0
end sub

Dim tiltcount, tabletilted: tiltcount = 0: TableTilted=false


TimerVRPlunger2.enabled = true   '  This sits outside of a sub, and tells the timer2 to be enabled at table load

Sub TimerVRPlunger_Timer
if VRPlunger.Y < 2402 then VRPlunger.Y = VRPlunger.y +5  'If the plunger is not fully extend it, then extend it by 5 coordinates in the Y,
End Sub

Sub TimerVRPlunger2_Timer
VRPlunger.Y = 2332 + (5* Plunger.Position) -20  ' This follows our dummy plunger position for analog plunger hardware users.
end sub

'******************
' Table stop/pause
'******************

Sub Table1_Paused
End Sub

Sub Table1_unPaused
End Sub

Sub Table1_Exit
  savehs
  Controller.Stop
End Sub


Sub Table1_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
        Plunger.PullBack
          PlaySound "plungerpull", 0, volmaster
      TimerVRPlunger.enabled = true  ' We enable the plunger timer below..  look for the TimerVPlunger Sub.
      TimerVRPlunger2.enabled = False' We disaable the plunger2 timer below..  look for the TimerVPlunger2 Sub.
  End If

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
      if smartflippers then
        FlipperDeActivate LeftFlipper, LFPress
        SolLFlipper True
      else
        LeftFlipper.RotateToEnd
      end if

      playFieldSound "FlipUpL", 0, leftFlipper, 1
      playFieldSound "FlipBuzzL", -1, leftFlipper, 1
    dof 101,1

  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
      if smartflippers then
        FlipperDeActivate RightFlipper, RFPress
        SolRFlipper True
      else
        RightFlipper.RotateToEnd
      end if
      playFieldSound "FlipUpR", 0, RightFlipper,1
      playFieldSound "FlipBuzzR", -1, RightFlipper,1
    dof 102,1


  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
    if InProgress=true then TiltIt

  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
    if InProgress=true then TiltIt

  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
    if InProgress=true then TiltIt
  End If

  If keycode = MechanicalTilt Then
    TiltCount=2
    if InProgress=true then TiltIt
  End If

  If (keycode = AddCreditKey or keycode = 4) then
    Credits=Credits+1
    if credits>30 then
      credits = 30
    else
      advancercredit(1)
    end if
    PlaySound "coin insert", 0, volmaster * volmotor
    PlaySound "knocker", 0, volmaster
    DOF 128, DOFPulse
    If B2SOn Then
      if credits <10 then Controller.B2SSetCredits credits else Controller.B2SSetCredits 9 end if
    end if
  end if

  if keycode=StartGameKey and Credits>0 and InProgress=false then
    startsequence.enabled=true
  end if

  If keycode = LeftFlipperKey Then
    Primary_flipper_button_left.X = 2097.845 + 8
  End If

  If keycode = RightFlipperKey Then
    Primary_flipper_button_right.X = 2107.33 - 8
  End If

  If keycode = LeftMagnaSave Then bLutActive = True
  If keycode = RightMagnaSave Then
    If bLutActive Then NextLUT: End If
  End If


End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySound "plungerspring", 0, volmaster
    TimerVRPlunger.enabled = false  'Disabling the plungertimer.. this is used to animate the plunger
    TimerVRPlunger2.enabled = true  'enabling the plungertimer.. this is used to animate the plunger
    VRPlunger.Y = 2332 ' Putting the plunger back to start position when we let go of it with a button. change number to the plunger position
  End If

  If keycode = LeftFlipperKey  and InProgress=true and TableTilted=false Then
      if smartflippers then
        FlipperDeActivate LeftFlipper, LFPress
        SolLFlipper false
      else
        LeftFlipper.RotateToStart
      end if
      playFieldSound "FlipDownL", 0, leftFlipper, 1
      stopSound "FlipBuzzLA"
      stopSound "FlipBuzzLB"
      stopSound "FlipBuzzLC"
      stopSound "FlipBuzzLD"
      dof 101,0
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
      if smartflippers then
        FlipperDeActivate RightFlipper, RFPress
        SolRFlipper false
      Else
        RightFlipper.RotateToStart
      end if
      playFieldSound "FlipDownR", 0, RightFlipper, 1
      stopSound "FlipBuzzRA"
      stopSound "FlipBuzzRB"
      stopSound "FlipBuzzRC"
      stopSound "FlipBuzzRD"
      dof 102,0
  End If

  If keycode = LeftFlipperKey Then
    Primary_flipper_button_left.X = 2097.845
  End If

  If keycode = RightFlipperKey Then
    Primary_flipper_button_right.X = 2107.33
  End If

  If keycode = LeftMagnaSave Then bLutActive = False

End Sub


Sub Drain_Hit()

  dim obj

    For Each obj in Bumpers: obj.hasHitEvent = true: Next
    leftslingshot.disabled = 0
    rightslingshot.disabled = 0
    PostDown
    if dooralreadyopen=1 then closeg.enabled=true end if
    kicker.enabled = 0: kickon.state = 0
    If TableTilted= true and tiltgame then
      inprogress = False
      tiltflash.visible = 1
      bipflash.visible = 0
      If B2Son then
        Controller.B2SSetTilt 1
        Controller.B2SSetBallInPlay 0
      end if
    else
      ballinplay = ballinplay + 1
      TableTilted=false
      tiltflash.visible = 0
      If B2Son then Controller.B2SSetTilt 0 end if

      if ballinplay < 6 then
        PlaySound "ball return", 0, volmaster*volmotor
        drain.timerenabled = 1
        If B2Son then Controller.B2SSetBallInPlay ballinplay
        bipflash.imagea = "B"&ballinplay
      Else
        PlaySound "gameover", 0, volmaster
        inprogress = False
        checkmatch.enabled = 1
        gameoverflash.visible = 1
        If match <> 10 then matchflash.imagea = match else matchflash.imagea = 0 end if
        matchflash.visible = 1
        bipflash.visible = 0
        SetHSLine 2, HSScore
        If B2Son then
          Controller.B2SSetGameOver 1
          Controller.B2SSetMatch match  '0 to 10 - 0 no display
          Controller.B2SSetBallInPlay 0
        end if
      end if
    end if

End Sub

Sub drain_timer()
  ballstarted = 0
  Drain.Kick 135, 50
  DOF 125, 2
  drain.timerenabled = 0
end sub

sub checkmatch_timer
  if score1 = match then addreplay
  checkmatch.enabled = 0
end sub


Sub Plunger_Init()
  Drain.CreateBall
End Sub

Sub TiltIt()
  dim obj
  if InProgress = true then
    TiltCount = TiltCount + 1
    if TiltCount = 3 then
      PostDown
      if dooralreadyopen=1 then closeg.enabled=true end if
      kicker.enabled = 0: kickon.state = 0
      TableTilted=True
      For Each obj in Bumpers: obj.hasHitEvent = False: Next
      For each obj in tilt: obj.state = false: next
      leftslingshot.disabled = 1
      rightslingshot.disabled = 1
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      PlaySoundAtVol "tilt", Bumperredbottom, 1
      If B2Son then
        Controller.B2SSetTilt 1
        tiltflash.visible = 1
        If tiltgame then Controller.B2SSetBallInPlay 0:bipflash.visible = 0: end if
      end if
      TiltCount = 0
    else
      TiltTimer.Interval = 500
      TiltTimer.Enabled = True
    end if
  end if
end sub

Sub Flipperanimation_Timer()
    FlipperLShadow.RotZ = LeftFlipper.CurrentAngle
    FlipperRShadow.RotZ = RightFlipper.CurrentAngle
    Pgate.rotx = -90 - Gate2.currentAngle * .2

    if incr1 > 0 then incr1 = incr1+1
    if incr2 > 0 then incr2 = incr2+1
    if incr3 > 0 then incr3 = incr3+1
    if incr4 > 0 then incr4 = incr4+1

  if incr1 = 2 then
    rscore1.objrotx = rscore1.objrotx + 18
    if rscore1.objrotx > 360 then rscore1.objrotx = rscore1.objrotx - 360
    incr1 = 0
  end if
  if incr2 = 2 then
    rscore2.objrotx = rscore2.objrotx + 18
    if rscore2.objrotx > 360 then rscore2.objrotx = rscore2.objrotx - 360
    incr2 = 0
  end if
  if incr3 = 2 then
    rscore3.objrotx = rscore3.objrotx + 18
    if rscore3.objrotx > 360 then rscore3.objrotx = rscore3.objrotx - 360
    incr3 = 0
  end if
  if incr4 = 2 then
    rscore4.objrotx = rscore4.objrotx + 18
    if rscore4.objrotx > 360 then rscore4.objrotx = rscore4.objrotx - 360
    incr4 = 0
  end if
End Sub

Sub Turnlightson

  Dim obj

  playsound "relay", 0, 1

    for each obj in gi: obj.state=1:next    'turn on GI lights
'   For each obj In dlights: obj.state = 1:Next 'optional GI lights

'**** bulb items
    For Each obj In lightson: obj.state = 1:Next
    For each obj in filaments: obj.blenddisablelighting = 120: Next
    For each obj in bulbs: obj.blenddisablelighting = 12: Next


    Backglass.imagea = "backglass kt bright"
    KI_flash.visible = 1  'KI
    NG_flash.visible = 1  'NG
    TUT_flash.visible = 1 'TuT
    bally_flash.visible = 1 'Bally
    faro_flash.visible = 1  'Faro

    gameoverflash.visible = 1
    If match <> 10 then matchflash.imagea = match else matchflash.imagea = 0 end if
    matchflash.visible = 1

If B2SOn then
  Controller.B2SSetGameOver 1
  Controller.B2SSetTilt 0
  if credits <10 then Controller.B2SSetCredits credits else Controller.B2SSetCredits 9 end if
  Controller.B2SSetMatch match
  Controller.B2SSetData 97, 1   'Tut
  Controller.B2SSetData 98, 1   'ng
  Controller.B2SSetData 99, 1   'Ki
  Controller.B2SSetData 90, 1   'Bally
  Controller.B2SSetData 91, 1   'Faro
  Controller.B2SSetBallInPlay ballinplay

end if

  lightflash.enabled = 1

end sub

sub holdr_timer
  playsound "hold relay"
end sub


' *********************************************
' **** Handle Slingshot and Contact Switch Hits
' *********************************************
'
'  -----------------------
'  Ball hit Left Slingshot
'  -----------------------
'
Dim Lstep, Rstep    'counters for sling animation

Sub LeftSlingShot_SlingShot()
  if TableTilted=false then
  If B2SOn Then DOF 103, DOFPulse
  PlaySound "slingshot", 0, volmaster*volbumpers
    AddScore (1)          ' 1 points
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
  end if
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 2:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 3:sling2.TransZ = 0:LSLing2.Visible = 0:LSLing0.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'  ------------------------
'  Ball hit Right Slingshot
'  ------------------------
'
Sub RightSlingShot_SlingShot()
  if TableTilted=false then
  If B2SOn Then DOF 104, DOFPulse
  PlaySound "slingshot", 0, volmaster*volbumpers
  AddScore (1)          ' 1 points
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 1
    RightSlingShot.TimerEnabled = 1
  end if
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 2:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 3:sling1.TransZ = 0:RSLing2.Visible = 0:RSLing0.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub


'********************************** bumpers

Sub Bumpergreenleft_Hit
  if TableTilted=false then
' Bumpergreenright.playhit
  if  T3.state=1 then AddScore(10) else AddScore(1) end if
  stopsound "bumper"
  playsound "bumper", 0, volmaster*volbumpers
  dof 105,2
  end if
End sub

Sub Bumpergreenright_Hit
  if TableTilted=false then
' Bumpergreenleft.playhit
  if  T3.state=1 then AddScore(10) else AddScore(1) end if
  stopsound "bumper"
  playsound "bumper", 0, volmaster*volbumpers
  dof 106,2
  end if
End sub

Sub Bumperredtop_Hit
  if TableTilted=false then
' Bumperredbottom.playhit
  if  T1.state=1 then AddScore(10) else AddScore(1) end if
  stopsound "bumper"
  playsound "bumper", 0, volmaster*volbumpers
  dof 107,2
  end if
End sub


Sub Bumperredbottom_Hit
  if TableTilted=false then
' Bumperredtop.playhit
  if  T1.state=1 then AddScore(10) else AddScore(1) end if
  stopsound "bumper"
  playsound "bumper", 0, volmaster*volbumpers
  dof 108,2
  end if
End sub


'************************* Triggers


Sub leftTlane_Hit
  PlaySoundAtVol "sensor", leftTlane, 1
  if TableTilted=false then
    T1.state=1
    T2.state=1
    BulbB2.state=1
    BulbB4.state=1
    AddScore(100)
    if specialslit=0 then checkspecials
  end if
End Sub


Sub Ulane_Hit
  PlaySoundAtVol "sensor", Ulane, 1
  if TableTilted=false then
  U1.state=1
  U2.state=1
  PostUp
  AddScore(100)
  if specialslit=0 then checkspecials
  end if
End Sub


Sub righttlane_Hit
  PlaySoundAtVol "sensor", righttlane, 1
  if TableTilted=false then
  T3.state=1
  T4.state=1
  BulbB1.state=1
  BulbB3.state=1
  AddScore(100)
  if specialslit=0 then checkspecials
  end if
End Sub



Sub leftspeciallane_Hit
  PlaySoundAtVol "sensor", leftspeciallane, 1
  if TableTilted=false then
  AddScore(100)
  if specialslit = 1 and specialleft.state = 1 then addreplay
  end if
End Sub


Sub rightspeciallane_Hit
  PlaySoundAtVol "sensor", rightspeciallane, 1
  if TableTilted=false then
  AddScore(100)
  if specialslit = 1 and specialright.state = 1 then addreplay
  end if
End Sub


Sub leftoutlane_Hit
  PlaySoundAtVol "sensor", leftoutlane, 1
  if TableTilted=false then
  if kickon.state = 0 then AddScore(100)
  end if
End Sub


Sub Rightoutlane_Hit
  PlaySoundAtVol "sensor", Rightoutlane, 1
  if TableTilted=false then
  AddScore(100)
  end if
End Sub


Sub CloseGateTrigger_Hit

End Sub


Sub ballplayed_Hit
  if dooralreadyopen=1 then closeg.enabled=true end if
End Sub



'************************* Targets


Sub TargetCentre_Hit
  if TableTilted=false then
  PostUp
  AddScore(100)
  if specialslit = 1 and specialctr.state = 1 then addreplay
  end if
End Sub


Sub TargetK_Hit
  if TableTilted=false then
  K1.state=1
  K2.state=1
  AddScore(100)
  if specialslit=0 then checkspecials
  end if
End Sub

Sub TargetI_Hit
  if TableTilted=false then
  I1.state=1
  I2.state=1
  AddScore(100)
  if specialslit=0 then checkspecials
  end if
End Sub

Sub TargetN_Hit
  if TableTilted=false then
  kicker.enabled = 0
  kickon.state = 0
  N1.state=1
  N2.state=1
  AddScore(100)
  if specialslit=0 then checkspecials
  end if
End Sub

Sub TargetG_Hit
  if TableTilted=false then
  if dooralreadyopen=1 then closeg.enabled=true end if
  G1.state=1
  G2.state=1
  AddScore(100)
  if specialslit=0 then checkspecials
  end if
End Sub

'**************************   check that specials is lit

Sub checkspecials

  if k1.state = 1 and I1.state = 1 and N1.state = 1 and G1.state = 1 and T1.state = 1 and U1.state = 1 and T3.state = 1 then specialslit=1

  if specialslit=1 Then
    if match = 10 then specialctr.state=1:specialleft.state=0:specialright.state=0 end if
    if match = 4 then specialctr.state=0: specialleft.state=1:specialright.state=0 end if
    if match = 9 then specialctr.state=0: specialleft.state=0:specialright.state=1 end if
    if match = 5 then specialctr.state=0: specialleft.state=1:specialright.state=0 end if
    if match = 2 then specialctr.state=1: specialleft.state=0:specialright.state=0 end if
    if match = 8 then specialctr.state=0: specialleft.state=1:specialright.state=0 end if
    if match = 3 then specialctr.state=0: specialleft.state=0:specialright.state=1 end if
    if match = 7 then specialctr.state=1: specialleft.state=0:specialright.state=0 end if
    if match = 1 then specialctr.state=0: specialleft.state=1:specialright.state=0 end if
    if match = 6 then specialctr.state=0: specialleft.state=0:specialright.state=1 end if
  end if
end sub


'**************************  1 point rubbers

Sub rubberwall1_Hit
  if TableTilted=false then
  AddScore(1)
  end if
End Sub

Sub rubberwall2_Hit
  if TableTilted=false then
  AddScore(1)
  end if
End Sub

Sub rubberwall3_Hit
  if TableTilted=false then
  AddScore(1)
  end if
End Sub

Sub rubberwall4_Hit
  if TableTilted=false then
  AddScore(1)
  end if
End Sub

Sub rubberwall5_Hit
  if TableTilted=false then
  AddScore(1)
  end if
End Sub

Sub rubberwall6_Hit
  if TableTilted=false then
  AddScore(1)
  end if
End Sub

Sub rubberwall7_Hit
  if TableTilted=false then
  AddScore(1)
  end if
End Sub





'********************************** Post up and down

Sub walBallSaver_Init()
  walBallSaver.IsDropped = true
End Sub

Sub PostUp
  if walBallSaver.IsDropped = True Then walBallSaver.IsDropped = False else exit sub
  PlaySoundAtVol "fx_solenoidon2", priHole18, 1
  PostPrim.image = "popup_post"
' PostPrim.TransZ = 30
  postupAnimation.enabled = 1
End Sub

Sub PostDown
  if walBallSaver.IsDropped = false Then walBallSaver.IsDropped = true else exit sub
  PlaySoundAtVol "fx_solenoidoff2", priHole18, 1
  PostPrim.image = "popup_post_dark"
' PostPrim.TransZ = 0
  postupAnimation.enabled = 1
End Sub

'************** Post up  Animation

Dim postupcase:postupcase=0

Sub postupAnimation_Timer()
  postupcase = postupcase + 1
  if postlight.State = 0 then
    Select Case postupcase
      Case 1: PostPrim.transz = 7
      Case 2: PostPrim.transz = 14
      Case 3: PostPrim.transz = 21
      Case 4: PostPrim.transz = 30
      if postupcase = 4 then
          postupcase = 0
          postupAnimation.enabled = 0
          postlight.State = 1
          postlight2.State = 1
      end if
    End Select
  Else
    Select Case postupcase
      Case 1: PostPrim.transz = 21
      Case 2: PostPrim.transz = 14
      Case 3: PostPrim.transz = 7
      Case 4: PostPrim.transz = 0
      if postupcase = 4 then
          postupcase = 0
          postupAnimation.enabled = 0
          postlight.State = 0
          postlight2.State = 0
      end if
    End Select
  end if
End Sub

'********************************** Gate open and close



 sub openg_timer
    bottgate(bgpos).isdropped=true
    bgpos=bgpos-1
    bottgate(bgpos).isdropped=false
  primgate.RotY=30+(bgpos*10)
     if bgpos=0 then
    PlaySoundAt "fx_gate_outlane", primgate
    gateon.state=1
    openg.enabled=false
    dooralreadyopen=1
  end if

 end sub

sub closeg_timer
    bottgate(bgpos).isdropped=true
    bgpos=bgpos+1
    bottgate(bgpos).isdropped=false
    primgate.RotY=30+(bgpos*10)
  if bgpos=6 then
    PlaySoundAt "fx_gate_outlane", primgate
    gateon.state=0
    closeg.enabled=false
    DOF 131,2
    dooralreadyopen=0
  end if
end sub

'********************************** rollovers



Dim buttonnumber

Sub Downpostleft_Hit()
  PlaySound "fx_trigger"
  if TableTilted=false then
  PostDown
  AddScore(1)
  buttonnumber = 1
  rollOverAnimation.enabled = 1
  end if
End Sub

Sub Downpostright_Hit()
  PlaySound "fx_trigger"
  if TableTilted=false then
  PostDown
  AddScore(1)
  buttonnumber = 2
  rollOverAnimation.enabled = 1
  end if
End Sub
'************** Button Animation

Dim buttoncase:buttoncase=0

Sub rollOverAnimation_Timer()
  buttoncase = buttoncase + 1
  Select Case buttoncase
    Case 1: EVAL("DPprim" & buttonNumber).transz = 2
    Case 2: EVAL("DPprim" & buttonNumber).transz = 1
    Case 3: EVAL("DPprim" & buttonNumber).transz = -1
    Case 4: EVAL("DPprim" & buttonNumber).transz = -2
        buttoncase = 0
        rollOverAnimation.enabled = 0
  End Select
End Sub

'************** Mushrooms

Sub mushrubber1_Hit
  if TableTilted=false then
  Dof 170,2
  Cap = 1
  MushroomAnimation.enabled = 1
  playFieldSound "MRCollision", 0, mushrubber1, 1
  AddScore(10)
  if dooralreadyopen=0 then openg.enabled=true end if
  end if
end sub

Sub mushrubber2_Hit
  if TableTilted=false then
  Dof 171, 2
  Cap = 2
  MushroomAnimation.enabled = 1
  playFieldSound "MRCollision", 0, mushrubber2, 1
  AddScore(10)
  kicker.enabled = 1
  kickon.state = 1
  end if
end sub



Dim Mush, cap:Mush=0
Sub MushroomAnimation_Timer
  Mush = Mush + 1
  Select Case Mush
    Case 1: EVAL("mushroomcap" & Cap).transz = 3
    Case 2: EVAL("mushroomcap" & Cap).transz = 5
    Case 3: EVAL("mushroomcap" & Cap).transz = 3
    Case 4: EVAL("mushroomcap" & Cap).transz = 0
        Mush = 0
        MushroomAnimation.enabled = 0
  End Select
End Sub

'************** Kick back
Dim KickIP, dir, strong

Sub Kicker_Hit
  Kicker.timerenabled=1
End Sub

Sub Kicker_Timer()
  DOF 111, DOFPulse
  StrongDir.enabled=1
  Kicker.Kick Dir, Strong
  Stantuffo.enabled=1
  Kicker.timerenabled=0
End Sub

Sub Stantuffo_timer()
  Primitive83.TransY=-70
  StantuffoR.enabled=1
  Stantuffo.enabled=0
End Sub

Sub StantuffoR_timer()
  Primitive83.TransY=Primitive83.TransY+5
  If Primitive83.TransY=0 Then
  StantuffoR.enabled=0
  End If
End Sub

Sub Kicker_UnHit
  PlaySound "popper_ball",0,.75,0,0.25
  kicker.enabled = 0
  kickon.state = 0
  AddScore (100)
End Sub

Sub StrongDir_timer()
KickIP = Int(Rnd*6)+1
Select Case KickIP
    Case 1 : Dir = 0 : Strong = 30
    Case 2 : Dir = 0 : Strong = 27
  Case 3 : Dir = 0 : Strong = 24
  Case 4 : Dir = 0 : Strong = 21
  Case 5 : Dir = 5 : Strong = 30
  Case 6 : Dir = 5 : Strong = 27
End Select
StrongDir.enabled=0
End Sub


'
' ******************************
' **** Scoring and credits *****
' ******************************
'
'  AddScore is main scoring routine
'
'  ScoreToAdd May be:
'      1, 10, 100
'

Sub AddScore(ScoreToAdd)
  if TableTilted=true then exit sub

if scoretoadd = 1 then
  stopSound "1 pt bell"
  PlaySound "1 pt bell", 0, volchimes*volmaster, 1
  advancerscore1
  if score1 <> 9 then
    score1 = score1 + 1
  else
    if score1 = 9 and score2 <> 9 then score1 = 0: score2 = score2 + 1: advancerscore2: end if
    if score1 = 9 and score2 = 9 and score3 <> 9 then score1 = 0: score2 = 0 : score3 = score3 + 1: advancerscore2: advancerscore3: end if
    if score1 = 9 and score2 = 9 and score3 = 9 and score4 <> 9 then score1 = 0: score2 = 0 :score3 = 0 : score4 = score4 + 1: advancerscore2: advancerscore3: advancerscore4: end if
    if score1 = 9 and score2 = 9 and score3 = 9 and score4 = 9 then score1 = 0: score2 = 0 :score3 = 0 : score4 = 0: advancerscore2: advancerscore3: advancerscore4:score5 = score5 + 1: end if
  end if
end if

if scoretoadd = 10 then

  StopSound "10-100pt bell"
  PlaySound "10-100pt bell", 0, volchimes*volmaster, 1
  dof 142, 2
  advancerscore2
  if score2 <> 9 then
    score2 = score2 + 1
  else
    if score2 = 9 and score3 <> 9 then score2 = 0 : score3 = score3 + 1:advancerscore3: end if
    if score2 = 9 and score3 = 9 and score4 <> 9 then score2 = 0 :score3 = 0 : score4 = score4 + 1:advancerscore3:advancerscore4: end if
    if score2 = 9 and score3 = 9 and score4 = 9 then score2 = 0 :score3 = 0 : score4 = 0: advancerscore3: advancerscore4:score5 = score5 + 1: end if
  end if
end if

if scoretoadd = 100 then

  StopSound "10-100pt bell"
  PlaySound "10-100Pt Bell", 0, volchimes*volmaster, 1
  dof 142, 2
  advancerscore3
  if score3 <> 9 then
    score3 = score3 + 1
  else
    if score3 = 9 and score4 <> 9 then score3 = 0 : score4 = score4 + 1:advancerscore4: end if
    if score3 = 9 and score4 = 9 then score3 = 0 :score4 = 0:advancerscore4:score5 = score5 + 1:  end if
  end if

end if

scoretotal = score1+10*score2+100*score3+1000*score4+10000*score5             'Total of current score

if scoretotal > HSscore then HSscore = scoretotal

' score1 = 1's: score2 = 10'2: score3 = 100's: score4 = 1000's: score5 = 10000's

If B2SOn then
Controller.B2SSetreel 4, score4
Controller.B2SSetreel 3, score3
Controller.B2SSetreel 2, score2
Controller.B2SSetreel 1, score1
end if

'
'  Check for replays Earned on Scoring
'
  If Scoretotal=>Replay1 And _
       Replay1Paid=False Then             'If the first replay score is reached,
       AddReplay()                    'then give a replay and
     Replay1Paid=True                 'mark the replay as paid
  End If
  If scoretotal=>Replay2 And _
       Replay2Paid=False Then             'Same as above,except for 2nd replay
     AddReplay()
     Replay2Paid=True
  End If
  If scoretotal=>Replay3 And _
       RePlay3Paid=False Then             'Same for 3rd replay
     AddReplay()
     Replay3Paid=True
    End If
  If scoretotal=>Replay4 And _
       RePlay4Paid=False Then             'Same for 4th replay
     AddReplay()
     Replay4Paid=True
  End If

'*********** calculate match, specials alternation

If scoretoadd = 1 then

matchposition = matchposition + 1

if matchposition = 10 then matchposition = 0 end if



select Case matchposition
  case 0: match = 10:if specialslit=1 then specialctr.state=1:specialleft.state=0:specialright.state=0 end if
  case 1: match = 4:if specialslit=1 then specialctr.state=0: specialleft.state=1:specialright.state=0 end if
  case 2: match = 9:if specialslit=1 then specialctr.state=0: specialleft.state=0:specialright.state=1 end if
  case 3: match = 5:if specialslit=1 then specialctr.state=0: specialleft.state=1:specialright.state=0 end if
  case 4: match = 2:if specialslit=1 then specialctr.state=1: specialleft.state=0:specialright.state=0 end if
  case 5: match = 8:if specialslit=1 then specialctr.state=0: specialleft.state=1:specialright.state=0 end if
  case 6: match = 3:if specialslit=1 then specialctr.state=0: specialleft.state=0:specialright.state=1 end if
  case 7: match = 7:if specialslit=1 then specialctr.state=1: specialleft.state=0:specialright.state=0 end if
  case 8: match = 1:if specialslit=1 then specialctr.state=0: specialleft.state=1:specialright.state=0 end if
  case 9: match = 6:if specialslit=1 then specialctr.state=0: specialleft.state=0:specialright.state=1 end if
end select

end if

End Sub

Sub Addreplay

  credits = credits + 1
  if credits>30 then
    credits = 30
  else
    advancercredit(1)
  end if
  PlaySound "knocker", 0, volmaster
  dof 128,2
If B2SOn then
    if credits <10 then Controller.B2SSetCredits credits else Controller.B2SSetCredits 9 end if
end if

end sub


' *********  score / credit reel movement

' score reel "0" = 306 deg. 36 deg per increment
' credit reel "0" = 21.2 deg . 10.6 deg per increment
Dim incr1, incr2, incr3, incr4

Sub advancerscore1
  playsound "reel", 0, volmaster
  rscore1.objrotx = rscore1.objrotx + 18
  incr1 = 1
End Sub

Sub advancerscore2
  playsound "reel", 0, volmaster
  rscore2.objrotx = rscore2.objrotx + 18
  incr2 = 1
End Sub

Sub advancerscore3
  playsound "reel", 0, volmaster
  rscore3.objrotx = rscore3.objrotx + 18
  incr3 = 1
End Sub

Sub advancerscore4
  playsound "reel", 0, volmaster
  rscore4.objrotx = rscore4.objrotx + 18
  incr4 = 1
End Sub

Sub advancercredit(sign)      'function to be used AFTER adding credit elsewhere in script
  playsound "reel", 0, volmaster
  creditreel.objrotx = creditreel.objrotx - sign*9.54
End Sub



'*********
'   LUT
'*********

Dim bLutActive, LUTImage
Sub LoadLUT
  bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
  UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT: LUTImage = (LUTImage +1 ) MOD 9: UpdateLUT: SaveLUT: End Sub

Sub UpdateLUT
Select Case LutImage
Case 0: table1.ColorGradeImage = "LUT0"
Case 1: table1.ColorGradeImage = "LUT1"
Case 2: table1.ColorGradeImage = "LUT2"
Case 3: table1.ColorGradeImage = "LUT3"
Case 4: table1.ColorGradeImage = "LUT4"
Case 5: table1.ColorGradeImage = "LUT5"
Case 6: table1.ColorGradeImage = "LUT6"
Case 7: table1.ColorGradeImage = "LUT7"
Case 8: table1.ColorGradeImage = "LUT8"
End Select
End Sub

'**** flashing lights - backglass

DIM min
Dim max
Dim lampe1
Dim lampe2
Dim lampe3
Dim zero


Sub lightflash_Timer()
min =0
max =6
lampe1=(Int((max-min+1)*Rnd+min))
min =0
max =6
lampe2=(Int((max-min+1)*Rnd+min))
min =0
max =6
lampe3=(Int((max-min+1)*Rnd+min))


if b2son then
  If lampe1 >0 then Controller.B2SSetData 99, 1 'KI
  If lampe1 =0 then Controller.B2SSetData 99, 0 'KI
  If lampe2 >0 then Controller.B2SSetData 98, 1 'NG
  If lampe2 =0 then Controller.B2SSetData 98, 0 'NG
  If lampe3 >0 then Controller.B2SSetData 97, 1 'TuT
  If lampe3 =0 then Controller.B2SSetData 97, 0 'TuT
end if

if VRroom or FSSmode then
  If lampe1 >0 then KI_flash.visible = 1  'KI
  If lampe1 =0 then KI_flash.visible = 0  'KI
  If lampe2 >0 then NG_flash.visible = 1  'NG
  If lampe2 =0 then NG_flash.visible = 0  'NG
  If lampe3 >0 then TUT_flash.visible = 1 'TuT
  If lampe3 =0 then TUT_flash.visible = 0 'TuT
end if

min =250
max =500
zero=(Int((max-min+1)*Rnd+min))
lightflash.Interval = zero
End Sub

'******************************* save / load scores

Dim TextStr

sub savehs

  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HSFileName,True)
    ScoreFile.WriteLine creditreel.objrotx
    ScoreFile.WriteLine Credits
    scorefile.writeline Match
    ScoreFile.WriteLine score1
    ScoreFile.WriteLine score2
    ScoreFile.WriteLine score3
    ScoreFile.WriteLine score4
    ScoreFile.WriteLine score5
    scorefile.writeline HSScore
    scorefile.writeline matchposition
    ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub

sub loadhs

  Dim FileObj
  Dim ScoreFile
    dim temp1, temp2, temp3, temp4, temp5, temp6, temp7, temp8, temp9, temp10, temp11, temp12, temp13, temp14, temp15, temp16, temp17, temp18

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & HSFileName) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & HSFileName)
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    temp1=TextStr.ReadLine
    temp2=textstr.readline
    temp3=textstr.readline
    temp4=textstr.readline
    temp5=textstr.readline
    temp6=textstr.readline
    temp7=textstr.readline
    temp8=textstr.readline
    temp9=textstr.readline
    temp10=textstr.readline
    TextStr.Close

    creditreel.objrotx = cdbl(temp1)
    Credits=cdbl(temp2)
    match = cdbl(temp3)
    score1 =cdbl(temp4)
    score2 =cdbl(temp5)
    score3 =cdbl(temp6)
    score4 =cdbl(temp7)
    score5 = cdbl(temp8)
    HSScore = cdbl(temp9)
    matchposition = cdbl(temp10)
    Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub

'********  Post it note for High score1

Function GetHSChar(String, Index)
  dim ThisChar
  dim FileName
  ThisChar = Mid(String, Index, 1)
  FileName = "PostIt"
  if ThisChar = " " or ThisChar = "" then
    FileName = FileName & "BL"
  elseif ThisChar = "<" then
    FileName = FileName & "LT"
  elseif ThisChar = "_" then
    FileName = FileName & "SP"
  else
    FileName = FileName & ThisChar
  End If
  GetHSChar = FileName
End Function

Sub SetHsLine(LineNo, String)
  dim Letter
  dim ThisDigit
  dim ThisChar
  dim StrLen
  dim LetterLine
  dim Index
  dim StartHSArray
  dim EndHSArray
  dim LetterName
  dim xfor
  StartHSArray=array(0,1,15)
  EndHSArray=array(0,10,19)
' StartHSArray=array(0,1,12)    - orig
' EndHSArray=array(0,10,21)    - orig
  StrLen = len(string)
  Index = 1

  for xfor = StartHSArray(LineNo) to EndHSArray(LineNo)
    Eval("HS"&xfor).image = GetHSChar(String, Index)
    Eval("HS0"&xfor).image = GetHSChar(String, Index)
    Index = Index + 1
  next

End Sub



' *************************************
' ****  Game Start Reset Sequence  ****
' *************************************
'
' When started, Plunger Timer gets here every 200ms until we are done.
' Entire startup process is staged to take 2.6 seconds to
' emulate an EM machine Game Start Reset sequence
'

Dim Phase: Phase = 0


Sub startsequence_Timer()
  Dim Obj
  Phase=Phase+1
  Select Case Phase

    Case 1:
    PlaySound"reset", 0, volmotor*volmaster           'Start of Game sound
    Credits=Credits-1
    advancercredit(-1)
    InProgress=true
    ballinplay = 0
    Replay1Paid=False           ' Reset High Scores Hit flags
    Replay2Paid=False
    Replay3Paid=False
        Replay4Paid=False
    ballstarted = 0
    TableTilted=False
    score5 = 0

      If B2SOn then
        if credits <10 then Controller.B2SSetCredits credits else Controller.B2SSetCredits 9 end if
        Controller.B2SSetGameOver 0
        Controller.B2SSetTilt 0
        Controller.B2SSetMatch 0  '0 to 10 - 0 no display
        Controller.B2SSetBallInPlay ballinplay '1 to 5 - 0 no display
      End If

      tiltflash.visible = 0
      gameoverflash.visible = 0
      matchflash.visible = 0

      bulbb1.state = 0
      bulbb2.state = 0
      bulbb3.state = 0
      bulbb4.state = 0
      for each obj in kingtut: obj.state=0: next
      specialctr.state=0:specialleft.state=0:specialright.state=0
      specialslit=0

    Case 1265:
    stopSound"reset"
    If score1<>0 or score2<>0 or score3<>0 or score4<>0 then playsound "1265", 0, volmotor*volmaster else playsound "2871", 0, volmotor*volmaster
If score1<>0 then
  score1 = score1 + 1
  if score1 = 10 then score1 = 0
  If B2SOn then Controller.B2SSetreel 1, score1
  advancerscore1
end if

If score2<>0 then
  score2 = score2 + 1
  if score2 = 10 then score2 = 0
  If B2SOn then Controller.B2SSetreel 2, score2
  advancerscore2
end if

If score3<>0 then
  score3 = score3 + 1
  if score3 = 10 then score3 = 0
  If B2SOn then Controller.B2SSetreel 3, score3
  advancerscore3
end if

If score4<>0 then
  score4 = score4 + 1
  if score4 = 10 then score4 = 0
  If B2SOn then Controller.B2SSetreel 4, score4
  advancerscore4
end if

    bipflash.visible = 1
    bipflash.imagea = "b5"
    If B2SOn then Controller.B2SSetBallInPlay 5

    Case 1453:
    If score1<>0 or score2<>0 or score3<>0 or score4<>0 then playsound "1453", 0, volmotor*volmaster else playsound "2871", 0, volmotor*volmaster
If score1<>0 then
  score1 = score1 + 1
  if score1 = 10 then score1 = 0
  If B2SOn then Controller.B2SSetreel 1, score1
  advancerscore1
end if

If score2<>0 then
  score2 = score2 + 1
  if score2 = 10 then score2 = 0
  If B2SOn then Controller.B2SSetreel 2, score2
  advancerscore2
end if

If score3<>0 then
  score3 = score3 + 1
  if score3 = 10 then score3 = 0
  If B2SOn then Controller.B2SSetreel 3, score3
  advancerscore3
end if

If score4<>0 then
  score4 = score4 + 1
  if score4 = 10 then score4 = 0
  If B2SOn then Controller.B2SSetreel 4, score4
  advancerscore4
end if

    bipflash.imagea = "b4"
    If B2SOn then Controller.B2SSetBallInPlay 4

        Case 1632:
    If score1<>0 or score2<>0 or score3<>0 or score4<>0 then playsound "1632", 0, volmotor*volmaster else playsound "2871", 0, volmotor*volmaster

If score1<>0 then
  score1 = score1 + 1
  if score1 = 10 then score1 = 0
  If B2SOn then Controller.B2SSetreel 1, score1
  advancerscore1
end if

If score2<>0 then
  score2 = score2 + 1
  if score2 = 10 then score2 = 0
  If B2SOn then Controller.B2SSetreel 2, score2
  advancerscore2
end if

If score3<>0 then
  score3 = score3 + 1
  if score3 = 10 then score3 = 0
  If B2SOn then Controller.B2SSetreel 3, score3
  advancerscore3
end if

If score4<>0 then
  score4 = score4 + 1
  if score4 = 10 then score4 = 0
  If B2SOn then Controller.B2SSetreel 4, score4
  advancerscore4
end if

    bipflash.imagea = "b3"
    If B2SOn then Controller.B2SSetBallInPlay 3

    Case 1811:
    If score1<>0 or score2<>0 or score3<>0 or score4<>0 then playsound "1811", 0, volmotor*volmaster else playsound "2871", 0, volmotor*volmaster
If score1<>0 then
  score1 = score1 + 1
  if score1 = 10 then score1 = 0
  If B2SOn then Controller.B2SSetreel 1, score1
  advancerscore1
end if

If score2<>0 then
  score2 = score2 + 1
  if score2 = 10 then score2 = 0
  If B2SOn then Controller.B2SSetreel 2, score2
  advancerscore2
end if

If score3<>0 then
  score3 = score3 + 1
  if score3 = 10 then score3 = 0
  If B2SOn then Controller.B2SSetreel 3, score3
  advancerscore3
end if

If score4<>0 then
  score4 = score4 + 1
  if score4 = 10 then score4 = 0
  If B2SOn then Controller.B2SSetreel 4, score4
  advancerscore4
end if

    bipflash.imagea = "b2"
    If B2SOn then Controller.B2SSetBallInPlay 2

    Case 1944:
    playsound "1944", 0, volmotor*volmaster
If score1<>0 then
  score1 = score1 + 1
  if score1 = 10 then score1 = 0
  If B2SOn then Controller.B2SSetreel 1, score1
  advancerscore1
end if

If score2<>0 then
  score2 = score2 + 1
  if score2 = 10 then score2 = 0
  If B2SOn then Controller.B2SSetreel 2, score2
  advancerscore2
end if

If score3<>0 then
  score3 = score3 + 1
  if score3 = 10 then score3 = 0
  If B2SOn then Controller.B2SSetreel 3, score3
  advancerscore3
end if

If score4<>0 then
  score4 = score4 + 1
  if score4 = 10 then score4 = 0
  If B2SOn then Controller.B2SSetreel 4, score4
  advancerscore4
end if

    bipflash.imagea = "b1"
    If B2SOn then Controller.B2SSetBallInPlay 1

    Case 2349:
    If score1<>0 or score2<>0 or score3<>0 or score4<>0 then playsound "2349", 0, volmotor*volmaster else playsound "2871", 0, volmotor*volmaster
If score1<>0 then
  score1 = score1 + 1
  if score1 = 10 then score1 = 0
  If B2SOn then Controller.B2SSetreel 1, score1
  advancerscore1
end if

If score2<>0 then
  score2 = score2 + 1
  if score2 = 10 then score2 = 0
  If B2SOn then Controller.B2SSetreel 2, score2
  advancerscore2
end if

If score3<>0 then
  score3 = score3 + 1
  if score3 = 10 then score3 = 0
  If B2SOn then Controller.B2SSetreel 3, score3
  advancerscore3
end if

If score4<>0 then
  score4 = score4 + 1
  if score4 = 10 then score4 = 0
  If B2SOn then Controller.B2SSetreel 4, score4
  advancerscore4
end if

    bipflash.visible = 0
    If B2SOn then Controller.B2SSetBallInPlay 0

    Case 2525:
    If score1<>0 or score2<>0 or score3<>0 or score4<>0 then playsound "2525", 0, volmotor*volmaster else playsound "2871", 0, volmotor*volmaster
If score1<>0 then
  score1 = score1 + 1
  if score1 = 10 then score1 = 0
  If B2SOn then Controller.B2SSetreel 1, score1
  advancerscore1
end if

If score2<>0 then
  score2 = score2 + 1
  if score2 = 10 then score2 = 0
  If B2SOn then Controller.B2SSetreel 2, score2
  advancerscore2
end if

If score3<>0 then
  score3 = score3 + 1
  if score3 = 10 then score3 = 0
  If B2SOn then Controller.B2SSetreel 3, score3
  advancerscore3
end if

If score4<>0 then
  score4 = score4 + 1
  if score4 = 10 then score4 = 0
  If B2SOn then Controller.B2SSetreel 4, score4
  advancerscore4
end if

    Case 2713:
    If score1<>0 or score2<>0 or score3<>0 or score4<>0 then playsound "1811", 0, volmotor*volmaster else playsound "2713", 0, volmotor*volmaster
    playsound "2713", 0, volmotor*volmaster
If score1<>0 then
  score1 = score1 + 1
  if score1 = 10 then score1 = 0
  If B2SOn then Controller.B2SSetreel 1, score1
  advancerscore1
end if

If score2<>0 then
  score2 = score2 + 1
  if score2 = 10 then score2 = 0
  If B2SOn then Controller.B2SSetreel 2, score2
  advancerscore2
end if

If score3<>0 then
  score3 = score3 + 1
  if score3 = 10 then score3 = 0
  If B2SOn then Controller.B2SSetreel 3, score3
  advancerscore3
end if

If score4<>0 then
  score4 = score4 + 1
  if score4 = 10 then score4 = 0
  If B2SOn then Controller.B2SSetreel 4, score4
  advancerscore4
end if

    Case 2871:
    playsound "2871", 0, volmotor*volmaster
If score1<>0 then
  score1 = score1 + 1
  if score1 = 10 then score1 = 0
  If B2SOn then Controller.B2SSetreel 1, score1
  advancerscore1
end if

If score2<>0 then
  score2 = score2 + 1
  if score2 = 10 then score2 = 0
  If B2SOn then Controller.B2SSetreel 2, score2
  advancerscore2
end if

If score3<>0 then
  score3 = score3 + 1
  if score3 = 10 then score3 = 0
  If B2SOn then Controller.B2SSetreel 3, score3
  advancerscore3
end if

If score4<>0 then
  score4 = score4 + 1
  if score4 = 10 then score4 = 0
  If B2SOn then Controller.B2SSetreel 4, score4
  advancerscore4
end if

    Case 3047:
    playsound "3047", 0, volmotor*volmaster
If score1<>0 then
  score1 = score1 + 1
  if score1 = 10 then score1 = 0
  If B2SOn then Controller.B2SSetreel 1, score1
  advancerscore1
end if

If score2<>0 then
  score2 = score2 + 1
  if score2 = 10 then score2 = 0
  If B2SOn then Controller.B2SSetreel 2, score2
  advancerscore2
end if

If score3<>0 then
  score3 = score3 + 1
  if score3 = 10 then score3 = 0
  If B2SOn then Controller.B2SSetreel 3, score3
  advancerscore3
end if

If score4<>0 then
  score4 = score4 + 1
  if score4 = 10 then score4 = 0
  If B2SOn then Controller.B2SSetreel 4, score4
  advancerscore4
end if



  Case 3357:
    ballinplay = ballinplay + 1
    If B2SOn then Controller.B2SSetBallInPlay ballinplay '1 to 5 - 0 no display
    Drain.Kick 135, 45
    bipflash.visible = 1
    bipflash.imagea = "b"&ballinplay



    Case 4026:

           Phase=0

         startsequence.Enabled=False        'Disable this timer
  End Select
End Sub



'*********************************************************************
'                 End Game script - start all supporting functions
'*********************************************************************

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

Sub PlaySoundAtVol(soundname, tableobj, Vol) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, Vol, Pan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

sub GameTimer_Timer
  RollingUpdate
  cor.Update
end Sub

'****************************************************** (from big brave)
'****  BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 1 ' total number of balls
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

Sub RollingUpdate()
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows

    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz > -7 Then
'         RandomSoundBallBouncePlayfieldSoft BOT(b)
        Else
'         RandomSoundBallBouncePlayfieldHard BOT(b)
        End If
      End If
    End If
    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If


  Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, sounddamp*Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit(idx)
  PlaySound "gate4", 0, 2*Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, sounddamp*Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
  RubbersD.dampen Activeball
  TargetBouncer Activeball, 1
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, sounddamp*Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, sounddamp*Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, sounddamp*Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, sounddamp*Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, sounddamp*Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, sounddamp*Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, sounddamp*Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim BIPL : BIPL = False       'Ball in plunger lane


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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

'**********************
' Ball Collision Sound
'**********************

'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine.

'New algorithm for OnBallBallCollision
'For stereo, xGain is a Playsound volume multiplier that provides a Constant Power pan.
'For quad, multiple PlaySound commands are launched together that are panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'XVol and YVol are Playsound volume multipliers that provide a Constant Power "pan" and "fade".
'Subtracting XVol or YVol from 1 yeilds an inverse response.

Dim PFOption: PFOption = 1  ' 1,2 or 3

'Sub OnBallBallCollision(ball1, ball2, velocity)
' If PFOption = 1 or PFOption = 2 Then
'   PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) * xGain(ball1), AudioPan(ball1), 0, Pitch(ball1), 0, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
' End If
' If PFOption = 3 Then
'   PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  *    YVol(ball1),  -1, 0, Pitch(ball1), 0, 0, -1 'Top Left Playfield Speaker
'   PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) *    YVol(ball1),   1, 0, Pitch(ball1), 0, 0, -1 'Top Right Playfield Speaker
'   PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  * (1-YVol(ball1)), -1, 0, Pitch(ball1), 0, 0,  1 'Bottom Left Playfield Speaker
'   PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) * (1-YVol(ball1)),  1, 0, Pitch(ball1), 0, 0,  1 'Bottom Right Playfield Speaker
' End If
'End Sub

Sub PlayFieldSound (SoundName, Looper, TableObject, VolMult)
'Plays the sound of a table object at the table object's coordinates.
'For stereo, xGain is a Playsound volume multiplier that provides a Constant Power pan.
'For quad, multiple PlaySound commands are launched together that are panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'XVol and YVol are Playsound volume multipliers that provide a Constant Power "pan" and "fade".
'Subtracting XVol or YVol from 1 yeilds an inverse response.

  If PFOption = 1 Or PFOption = 2 Then
    If Looper = -1 Then
      PlaySound SoundName&"A", Looper, VolMult * xGain(TableObject), AudioPan(TableObject), 0, 0, 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
    If Looper = 0 Then
      PlaySound SoundName, Looper, VolMult * xGain(TableObject), AudioPan(TableObject), 0, 0, 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
  End If
  If PFOption = 3 Then
    If Looper = -1 Then
      PlaySound SoundName&"A", Looper, VolMult *    XVol(TableObject)  *    YVol(TableObject),  -1, 0, 0, 0, 0, -1  'Top Left PF Speaker
      PlaySound SoundName&"B", Looper, VolMult * (1-XVol(TableObject)) *    YVol(TableObject),   1, 0, 0, 0, 0, -1  'Top Right PF Speaker
      PlaySound SoundName&"C", Looper, VolMult *    XVol(TableObject)  * (1-YVol(TableObject)), -1, 0, 0, 0, 0,  1  'Bottom Left PF Speaker
      PlaySound SoundName&"D", Looper, VolMult * (1-XVol(TableObject)) * (1-YVol(TableObject)),  1, 0, 0, 0, 0,  1  'Bottom Right PF Speaker
    End If
    If Looper = 0 Then
      PlaySound SoundName, Looper, VolMult *    XVol(TableObject)  *    YVol(TableObject),  -1, 0, 0, 0, 0, -1  'Top Left PF Speaker
      PlaySound SoundName, Looper, VolMult * (1-XVol(TableObject)) *    YVol(TableObject),   1, 0, 0, 0, 0, -1  'Top Right PF Speaker
      PlaySound SoundName, Looper, VolMult *    XVol(TableObject)  * (1-YVol(TableObject)), -1, 0, 0, 0, 0,  1  'Bottom Left PF Speaker
      PlaySound SoundName, Looper, VolMult * (1-XVol(TableObject)) * (1-YVol(TableObject)),  1, 0, 0, 0, 0,  1  'Bottom Right PF Speaker
    End If
  End If
End Sub

Sub PlayFieldSoundAB (SoundName, Looper, VolMult)
'Plays the sound of a table object at the Active Ball's location.
'For stereo, xGain is a Playsound volume multiplier that provides a Constant Power pan.
'For quad, multiple PlaySound commands are launched together that are panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'XVol and YVol are Playsound volume multipliers that provide a Constant Power "pan" and "fade".
'Subtracting XVol or YVol from 1 yeilds an inverse response.

  If PFOption = 1 Or PFOption = 2 Then
    PlaySound SoundName, Looper, VolMult * Vol(ActiveBall) * xGain(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
  End If
  If PFOption = 3 Then
    PlaySound SoundName, Looper, VolMult * Vol(ActiveBall) *    XVol(ActiveBall)  *    YVol(ActiveBall),  -1, 0, Pitch(ActiveBall), 0, 0, -1  'Top Left PF Speaker
    PlaySound SoundName, Looper, VolMult * Vol(ActiveBall) * (1-XVol(ActiveBall)) *    YVol(ActiveBall),   1, 0, Pitch(ActiveBall), 0, 0, -1  'Top Right PF Speaker
    PlaySound SoundName, Looper, VolMult * Vol(ActiveBall) *    XVol(ActiveBall)  * (1-YVol(ActiveBall)), -1, 0, Pitch(ActiveBall), 0, 0,  1  'Bottom Left PF Speaker
    PlaySound SoundName, Looper, VolMult * Vol(ActiveBall) * (1-XVol(ActiveBall)) * (1-YVol(ActiveBall)),  1, 0, Pitch(ActiveBall), 0, 0,  1  'Bottom Right PF Speaker
  End If
End Sub

Sub PlayReelSound (SoundName, Pan)
Dim ReelVolAdj
ReelVolAdj = 0.2
'Provides a Constant Power Pan for the backglass reel sound volume to match the playfield's Constant Power Pan response
  If showDT = False Then  '-3dB for desktop mode
    If Pan = "Lpan" Then PlaySound SoundName, 0, ReelVolAdj * 1.00, -0.12, 0, 0, 0, 0, 0  'Panned 3/4 Left at 0dB * ReelVolAdj
    If Pan = "Mpan" Then PlaySound SoundName, 0, ReelVolAdj * 0.33,  0.00, 0, 0, 0, 0, 0  'Panned Middle at -3dB * ReelVolAdj
    If Pan = "Rpan" Then PlaySound SoundName, 0, ReelVolAdj * 1.00,  0.12, 0, 0, 0, 0, 0  'Panned 3/4 Right at 0dB * ReelVolAdj
  Else
    If Pan = "Lpan" Then PlaySound SoundName, 0, ReelVolAdj * 0.33, -0.12, 0, 0, 0, 0, 0  'Panned 3/4 Left at -3dB * ReelVolAdj
    If Pan = "Mpan" Then PlaySound SoundName, 0, ReelVolAdj * 0.11,  0.00, 0, 0, 0, 0, 0  'Panned Middle at -6dB * ReelVolAdj
    If Pan = "Rpan" Then PlaySound SoundName, 0, ReelVolAdj * 0.33,  0.12, 0, 0, 0, 0, 0  'Panned 3/4 Right at -3dB * ReelVolAdj
  End If
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************


Function xGain(TableObj)
'xGain algorithm calculates a PlaySound Volume parameter multiplier to provide a Constant Power "pan".
'PFOption=1:  xGain = 1 at PF Left, xGain = 0.32931 (-3dB for PlaySound's volume parameter) at PF Center and xGain = 1 at PF Right.  Used for Left & Right stereo PF Speakers.
'PFOption=2:  xGain = 1 at PF Top, xGain = 0.32931 (-3dB for PlaySound's volume parameter) at PF Center and xGain = 1 at PF Bottom.  Used for Top & Bottom stereo PF Speakers.
  Dim tmp
    If PFOption=1 Then tmp = TableObj.x * 2 / table1.width-1
  If PFOption=2 Then tmp = TableObj.y * 2 / table1.height-1
  If tmp < 0 Then
  xGain = 0.3293074856*EXP(-0.9652695455*tmp^3 - 2.452909811*tmp^2 - 2.597701999*tmp)
  Else
  xGain = 0.3293074856*EXP(-0.9652695455*-tmp^3 - 2.452909811*-tmp^2 - 2.597701999*-tmp)
  End If
End Function

Function XVol(tableobj)
'XVol algorithm calculates a PlaySound Volume parameter multiplier for a tableobj based on its X table position to provide a Constant Power "pan".
'XVol = 1 at PF Left, XVol = 0.32931 (-3dB for PlaySound's volume parameter) at PF Center and XVol = 0 at PF Right
Dim tmpx
  If PFOption = 3 Then
    tmpx = tableobj.x * 2 / table1.width-1
    XVol = 0.3293074856*EXP(-0.9652695455*tmpx^3 - 2.452909811*tmpx^2 - 2.597701999*tmpx)
  End If
End Function

Function YVol(tableobj)
'YVol algorithm calculates a PlaySound Volume parameter multiplier for a tableobj based on its Y table position to provide a Constant Power "fade".
'YVol = 1 at PF Top, YVol = 0.32931 (-3dB for PlaySound's volume parameter) at PF Center and YVol = 0 at PF Bottom
Dim tmpy
  If PFOption = 3 Then
    tmpy = tableobj.y * 2 / table1.height-1
    YVol = 0.3293074856*EXP(-0.9652695455*tmpy^3 - 2.452909811*tmpy^2 - 2.597701999*tmpy)
  End If
' TB2.text = "yVol=" & Round(yVol,4)
End Function



'**********************************



'******************************************************
'                FLIPPER POLARITY AND RUBBER DAMPENER
'                        SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
        dim x, aCount : aCount = 0
        redim a(uBound(aArray) )
        for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
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
        redim aArray(aCount-1+offset)        'Resize original array
        for x = 0 to aCount-1                'set objects back into original array
                if IsObject(a(x)) then
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
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
        dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
        Y = M*x+b
        PSlope = Y
End Function

' Used for flipper correction
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

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
        dim y 'Y output
        dim L 'Line
        dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
                if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
        Next
        if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
        Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

        if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
        if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

        LinearEnvelope = Y
End Function



'*****************
' Maths
'*****************
Dim PI: PI = 4*Atn(1)

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

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

'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function


'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1 'EM's to late 80's
'Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0:
    SOSRampup = 2.5
  Case 1:
    SOSRampup = 6
  Case 2:
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
'Const EOSReturn = 0.025  'mid 90's and later

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
  Flipper.eostorque = EOST*EOSReturn/FReturn


  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b, BOT
    BOT = GetBalls

    For b = 0 to UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3*Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
    if FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If

  End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
    Dim Dir
    Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
    Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
    Dim CatchTime : CatchTime = GameTime - FCount

    if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
            if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
                    LiveCatchBounce = 0
            else
                    LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
            end If

            If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
            ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
            ball.angmomx= 0
            ball.angmomy= 0
            ball.angmomz= 0
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
    End If
End Sub



'******************************************************
'****  PHYSICS DAMPENERS
'******************************************************
'
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR



Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
  TargetBouncer Activeball, 0.7
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

dim FlippersD : Set FlippersD = new Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

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

' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  public sub Dampenf(aBall, parm) 'Rubberizer is handle here
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
' Thalamus - patched :       aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    End If
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



'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets, 2 = orig TargetBouncer
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7 when TargetBouncerEnabled=1, and 1.1 when TargetBouncerEnabled=2

sub TargetBouncer(aBall,defvalue)  'use defvalue to manipulate the bounce of the ball/target interaction
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled = 1 and aball.z < 30 then
        'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        vel = BallSpeed(aBall)
        if aBall.velx = 0 then vratio = 1 else vratio = aBall.vely/aBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1: zMultiplier = 0.2*defvalue
      Case 2: zMultiplier = 0.25*defvalue
            Case 3: zMultiplier = 0.3*defvalue
      Case 4: zMultiplier = 0.4*defvalue
            Case 5: zMultiplier = 0.45*defvalue
            Case 6: zMultiplier = 0.5*defvalue
        End Select
        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
        aBall.vely = aBall.velx * vratio
        'debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        'debug.print "conservation check: " & BallSpeed(aBall)/vel
    elseif TargetBouncerEnabled = 2 and aball.z < 30 then
    'debug.print "velz: " & activeball.velz
    Select Case Int(Rnd * 4) + 1
      Case 1: zMultiplier = defvalue+1.1
      Case 2: zMultiplier = defvalue+1.05
      Case 3: zMultiplier = defvalue+0.7
      Case 4: zMultiplier = defvalue+0.3
    End Select
    aBall.velz = aBall.velz * zMultiplier * TargetBouncerFactor
    'debug.print "----> velz: " & activeball.velz
    'debug.print "conservation check: " & BallSpeed(aBall)/vel
  end if
end sub

Sub TargetBounce_Hit
  TargetBouncer activeball, 1
End Sub

Sub ShadowHide_Hit()
Shadowblock.visible = false
End Sub

Sub ShadowShow_Hit()
Shadowblock.visible = true
End Sub


'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
    public ballvel, ballvelx, ballvely, ballvelz, ballangmomx, ballangmomy, ballangmomz

    Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : redim ballvelz(0) : redim ballangmomx(0) : redim ballangmomy(0): redim ballangmomz(0): End Sub

    Public Sub Update()    'tracks in-ball-velocity
        dim str, b, AllBalls, highestID : allBalls = getballs

        for each b in allballs
            if b.id >= HighestID then highestID = b.id
        Next

        if uBound(ballvel) < highestID then redim ballvel(highestID)    'set bounds
        if uBound(ballvelx) < highestID then redim ballvelx(highestID)    'set bounds
        if uBound(ballvely) < highestID then redim ballvely(highestID)    'set bounds
        if uBound(ballvelz) < highestID then redim ballvelz(highestID)    'set bounds
        if uBound(ballangmomx) < highestID then redim ballangmomx(highestID)    'set bounds
        if uBound(ballangmomy) < highestID then redim ballangmomy(highestID)    'set bounds
        if uBound(ballangmomz) < highestID then redim ballangmomz(highestID)    'set bounds

        for each b in allballs
            ballvel(b.id) = BallSpeed(b)
            ballvelx(b.id) = b.velx
            ballvely(b.id) = b.vely
            ballvelz(b.id) = b.velz
            ballangmomx(b.id) = b.angmomx
            ballangmomy(b.id) = b.angmomy
            ballangmomz(b.id) = b.angmomz
        Next
    End Sub
End Class


'******************************************************
'****  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Flippers:   https://www.youtube.com/watch?v=FWvM9_CdVHw
' Dampeners:  https://www.youtube.com/watch?v=tqsxx48C6Pg
' Physics:    https://www.youtube.com/watch?v=UcRMG-2svvE
'
'
' Note: BallMass must be set to 1. BallSize should be set to 50 (in other words the ball radius is 25)
'
' Recommended Table Physics Settings
' | Gravity Constant             | 0.97      |
' | Playfield Friction           | 0.15-0.25 |
' | Playfield Elasticity         | 0.25      |
' | Playfield Elasticity Falloff | 0         |
' | Playfield Scatter            | 0         |
' | Default Element Scatter      | 2         |
'
' Bumpers
' | Force         | 9.5-10.5 |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 4-5  |
' | Slingshot Theshold | 2-3  |
' | Elasticity         | 0.85 |
' | Friction           | 0.8  |
' | Scatter Angle      | 1    |



'******************************************************
'****  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. an object or point to tell the script where the tip of the flipper is at rest (EndPointLp, EndPointRp)
' 4. and, special scripting
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
' | EOS Torque         | 0.3            | 0.3                   | 0.275                  | 0.275              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'


'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity


InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
Sub InitPolarity()
        dim x, a : a = Array(LF, RF)', LF1, RF1)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
                x.enabled = True
                x.TimeDelay = 80  '*****Important, this variable is an offset for the speed that the ball travels down the table to determine if the flippers have been fired
              'This is needed because the corrections to ball trajectory should only applied if the flippers have been fired and the ball is in the trigger zones.
              'FlipAT is set to GameTime when the ball enters the flipper trigger zones and if GameTime is less than FlipAT + this time delay then changes to velocity
              'and trajectory are applied.  If the flipper is fired before the ball enters the trigger zone then with this delay added to FlipAT the changes
              'to tragectory and velocity will not be applied.  Also if the flipper is in the final 20 degrees changes to ball values will also not be applied.
              '"Faster" tables will need a smaller value while "slower" tables will need a larger value to give the ball more time to get to the flipper.
              'If this value is not set high enough the Flipper Velocity and Polarity corrections will NEVER be applied.

        Next

        AddPt "Polarity", 0, 0, 0
        AddPt "Polarity", 1, 0.05, -2.7
        AddPt "Polarity", 2, 0.33, -2.7
        AddPt "Polarity", 3, 0.37, -2.7
        AddPt "Polarity", 4, 0.41, -2.7
        AddPt "Polarity", 5, 0.45, -2.7
        AddPt "Polarity", 6, 0.576,-2.7
        AddPt "Polarity", 7, 0.66, -1.8
        AddPt "Polarity", 8, 0.743, -0.5
        AddPt "Polarity", 9, 0.81, -0.5
        AddPt "Polarity", 10, 0.88, 0

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        LF.Object = LeftFlipper
'   LF1.Object = LeftFlipper1
        LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
        RF.Object = RightFlipper
'   RF1.Object = RightFlipper1
        RF.EndPoint = EndPointRp
End Sub


Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub


'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)', LF1, RF1)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
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
  Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property
  Public Property Get EndPointY: EndPointY = FlipperEndY : End Property

  Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
    if gametime > 100 then Report aChooseArray
  End Sub

  Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
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

'********Triggered by a ball hitting the flipper trigger area
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

'*********Used to rotate flipper since this is removed from the key down for the flippers
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
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))   '% of flipper swing
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub

'***********gameTime is a global variable of how long the game has progressed in ms
'***********This function lets the table know if the flipper has been fired
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

'***********This is turned on when a ball leaves the flipper trigger area
  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        RemoveBall aBall
        exit Sub
      end if

      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
        Dim BOT, b

        If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
                EOSNudge1 = 1
                'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
                If Flipper2.currentangle = EndAngle2 Then
                        BOT = GetBalls
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
                                        'Debug.Print "ball in flip1. exit"
                                         exit Sub
                                end If
                        Next
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
                                        BOT(b).velx = BOT(b).velx / 1.3
                                        BOT(b).vely = BOT(b).vely - 0.5
                                end If
                        Next
                End If
        Else
      If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 then
        EOSNudge1 = 0
      end if
        End If
End Sub

'*******************************************
'  Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
  If Enabled Then
    'LF.Fire
    leftflipper.rotatetoend

    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
'     RandomSoundReflipUpLeft LeftFlipper
    Else
'     SoundFlipperUpAttackLeft LeftFlipper
'     RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
'     RandomSoundFlipperDownLeft LeftFlipper
    End If
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    'RF.Fire
    rightflipper.rotatetoend

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
'     RandomSoundReflipUpRight RightFlipper
    Else
'     SoundFlipperUpAttackRight RightFlipper
'     RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
'     RandomSoundFlipperDownRight RightFlipper
    End If
  End If
End Sub

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RandomSoundFlipper()
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1)

Sub BallShadowUpdate_timer()
    Dim BOT, b
  Dim midpoint

  midpoint = (Table1.Width - 100)/2

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
'     If BOT(b).X < midpoint Then
      if BOT(b).X < 2*midpoint then
        BallShadow(b).X = BOT(b).X + 15*((BOT(b).X - midpoint)/midpoint)
      end if

'       BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (midpoint))/15)) + 6
'     Else
'       BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (midpoint))/15)) - 6
'     End If

      if BOT(b).X => (Table1.Width - 100) then BallShadow(b).X = BOT(b).X  end if

        ballShadow(b).Y = BOT(b).Y + 12

        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub
