'
' MM     MM    AAAAA    RRRRRRRR   YY      YY       SSSSSSS   HH     HH  EEEEEEEE   LL        LL        EEEEEEEE  YY      YY   ''   SSSSSSS
' MMM   MMM   AA   AA   RR     RR   YY    YY       SS     SS  HH     HH  EE         LL        LL        EE         YY    YY    ''  SS     SS
' MMMM MMMM  AA     AA  RR     RR    YY  YY        SS         HH     HH  EE         LL        LL        EE          YY  YY    ''   SS
' MM MMM MM  AA     AA  RR     RR     YYYY         SS         HH     HH  EE         LL        LL        EE           YYYY          SS
' MM  M  MM  AAAAAAAAA  RRRRRRRR       YY           SSSSSSS   HHHHHHHHH  EEEEEEE    LL        LL        EEEEEEE       YY            SSSSSSS
' MM     MM  AA     AA  RR  RR         YY                 SS  HH     HH  EE         LL        LL        EE            YY                  SS
' MM     MM  AA     AA  RR   RR        YY                 SS  HH     HH  EE         LL        LL        EE            YY                  SS
' MM     MM  AA     AA  RR    RR       YY          SS     SS  HH     HH  EE         LL        LL        EE            YY           SS     SS
' MM     MM  AA     AA  RR     RR      YY           SSSSSSS   HH     HH  EEEEEEEEE  LLLLLLLL  LLLLLLLL  EEEEEEEEE     YY            SSSSSSS
'
'
' FFFFFFFFF  RRRRRRRR     AAAAA    NN     NN  KK     KK  EEEEEEEE  NN     NN   SSSSSSS   TTTTTTTT  EEEEEEEE  II  NN     NN
' FF         RR     RR   AA   AA   NNN    NN  KK    KK   EE        NNN    NN  SS     SS     TT     EE        II  NNN    NN
' FF         RR     RR  AA     AA  NNNN   NN  KK   KK    EE        NNNN   NN  SS            TT     EE        II  NNNN   NN
' FF         RR     RR  AA     AA  NN NN  NN  KK  KK     EE        NN NN  NN  SS            TT     EE        II  NN NN  NN
' FFFFF      RRRRRRRR   AAAAAAAAA  NN  NN NN  KK KK      EEEEEEE   NN  NN NN   SSSSSSS      TT     EEEEEEE   II  NN  NN NN
' FF         RR  RR     AA     AA  NN   NNNN  KKKK       EE        NN   NNNN         SS     TT     EE        II  NN   NNNN
' FF         RR   RR    AA     AA  NN    NNN  KK  KK     EE        NN    NNN         SS     TT     EE        II  NN    NNN
' FF         RR    RR   AA     AA  NN     NN  KK    KK   EE        NN     NN  SS     SS     TT     EE        II  NN     NN
' FF         RR     RR  AA     AA  NN     NN  KK     KK  EEEEEEEEE NN     NN   SSSSSSS      TT     EEEEEEEEE II  NN     NN     by Sega, 1995
'
'
' VPX version by Schreibi34 and Herweh
'
' Version 1.0: Initial release
' Version 1.0.1: Correct LUT initialization
'        Added a wall for Frank
'        Added a playfield wall left to the Geneva scoop
' Version 1.1:   Some physical improvements to match the real table, i.e. at the Sarcophagus scoop, the kickback and the plastic ramp
'        Less reflection on the transparent plastic over the bumper area
'        Correct GI-Off initialisation at table startup
'
' Schreibi34:
' Visuals and gameplay
' "It's simply amazing what you do in Blender. So thrilling, so inspiring, so extraordinary. I'm overwhelmed every time when I see what you have created.
'  You are the man. And a very relaxed, cool guy too. And for sure an awesome drummer ;-).
'  Once again it was a pleasure and an honor to work with you" (this was written by Herweh)
'
' Herweh:
' Scripting, VP development and gameplay
' "Everything you see below in this script! Without him Mr. Basic v2.0 Schreibi34 would have been totally lost! What he has done on this table goes beyond
'  my imagination. From animating Frank to implementing RothbauerW's physics to the texture swaps to all sorts of extras. Please check out all the stuff
'  that the script Voodoo Master has provided for you below!
'  Thanks for being such a cool guy! It was a pleasure!" (this was written by Schreibi34)
'
' --------------------------------------------------------------------------------------------------------------------------------------------------------
'
' Special thanks to:
'
' - Dark: We all know what he did! Frank is just jawdropping!!
' - Dids666: For starting all of this and his awesome meshes. Thanks for letting us finish this!!
' - Sheltemke: For stitching together the PF image and playtesting
' - Mlager8: For PS help on the sling plastics
' - RothbauerW: For his awesome physics' guide at vpinball.com and some more help on the physics
' - Mark70, Bord, Thalamus: For beta testing and some good hints
' - Batch: For the DT background image (Sorry for making it a bit darker! Batch's original is in the image manager for those who like more pop)
' - Schotty VH: For gameplay hints


Option Explicit
Randomize


Dim tablewidth: tablewidth = Frankenstein.width
Dim tableheight: tableheight = Frankenstein.height

Dim EnableGI, EnableFlasher, EnableReflectionsAtBall, EnableLUTChangeWithRightMagnaSave, SelectSidewalls, ShowBallShadow, LetTheBallJump, AnimateFranksHead, RailsVisible

' ****************************************************
' OPTIONS
' ****************************************************

' ENABLE/DISABLE GI (general illumination)
' 0 = GI is off
' 1 = GI is on
EnableGI = 1 '(do not change here, use F12 menu in Game)

' ENABLE/DISABLE flasher
' 0 = Flashers are off
' 1 = Flashers are on
EnableFlasher = 1 '(do not change here, use F12 menu in Game)

' ENABLE/DISABLE insert reflections at the ball
' 0 = reflections are off
' 1 = reflections are on (value can be between 0 and 1 and controls the reflections' intensity)
EnableReflectionsAtBall = 1 '(do not change here, use F12 menu in Game)

' ENABLED/DISABLE THE ABILITY TO CHANGE THE LUT FILE WITH THE RIGHT MAGNASAVE BUTTON
' 0 = disabled
' 1 = enabled (default)
EnableLUTChangeWithRightMagnaSave = 1 '(do not change here, use F12 menu in Game)

' SELECT YOUR PREFERRED SIDEWALLS
' 0 = autoselect (default)
' 1 = use sidewalls rendered for fullscreen view
' 2 = use sidewalls rendered for a medium view
' 3 = use sidewalls rendered for desktop view
' 4 = use sidewalls rendered for general illumination is off
SelectSidewalls = 0 '(do not change here, use F12 menu in Game)

' ANIMATE FRANK'S HEAD MOVING AROUND
' 0 = animation is off or motor is broken
' 1 = Frank's head is randonly moving around
' 2 = Frank's head is following the nearest ball
' 3 = Frank is watching the action (looking to switches or firing solenoids of the table)
' 4 = Frank's head is moving around like in some youtube videos
AnimateFranksHead = 4 '(do not change here, use F12 menu in Game)

' SHOW BALL SHADOWS
' 0 = no ball shadows
' 1 = ball shadows are visible
ShowBallShadow = 1 '(do not change here, use F12 menu in Game)

' LET THE BALL JUMP A BIT
' 0 = off
' 1 to 6 = ball jump intensity (3 = default)
LetTheBallJump = 3 '(do not change here, use F12 menu in Game)

' SIDE RAILS AND LOCKBAR VISIBILITY
'   0 = hide side rails
'   1 = show side rails
'   2 = side rails visible just in desktop mode
RailsVisible = 2 '(do not change here, use F12 menu in Game)

' volume for rolling sound and head motor sound
'Const RollingSoundFactor = 0.15
Const HeadMotorSoundFactor = 0.6

' LUT names and inital LUT
Dim luts, lutpos
luts = Array("ColorGradeLUT256x16_1to1", "LUT_MSF_1", "LUT_MSF_2", "LUT_MSF_3", "LUT_MSF_4", "LUT_MSF_5", "LUT_MSF_6")
lutpos = 0


'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************


' ****************************************************
' standard definitions
' ****************************************************

Const UseSolenoids  = 2
Const UseLamps    = 0
Const UseSync     = 0
Const HandleMech  = 0
Const UseGI     = 1

'Standard Sounds
Const SSolenoidOn   = "fx_Solenoid"
Const SSolenoidOff  = ""
Const SFlipperOn  = ""
Const SFlipperOff   = ""
Const SCoin     = ""


'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************

Const cDMDRotation  = -1      '-1 for No change, 0 - DMD rotation of 0?, 1 - DMD rotation of 90?
Const cGameName   = "frankst"   'ROM name
Const ballsize    = 50
Const ballmass    = 1

If Version < 10600 Then
  MsgBox "This table requires Visual Pinball 10.6 or newer!" & vbNewLine & "Your version: " & Replace(Version/1000,",","."), , "Laser War VPX"
End If

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package."
On Error Goto 0

LoadVPM "01550000", "DE.VBS", 3.26

Dim DesktopMode: DesktopMode = Frankenstein.ShowDT
Dim i, bsTrough

If LetTheBallJump > 6 Then LetTheBallJump = 6 : If LetTheBallJump < 0 Then LetTheBallJump = 0
If EnableReflectionsAtBall < 0 Or EnableReflectionsAtBall > 1 Then EnableReflectionsAtBall = 0

LUTBox.Visible = False



' ****************************************************
' table init
' ****************************************************
Sub Frankenstein_Init()
  vpmInit Me
  With Controller
        .GameName       = cGameName
        .SplashInfoLine   = "Frankenstein (Sega 1995)"
    .HandleMechanics  = False
    .HandleKeyboard   = False
    .ShowDMDOnly    = True
    .ShowFrame      = False
    .ShowTitle      = False
    .Hidden       = False
    If cDMDRotation >= 0 Then .Games(cGameName).Settings.Value("rol") = cDMDRotation
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With
  ' initialize some table settings
  InitTable
End Sub

Sub Frankenstein_Paused()   : Controller.Pause = True : End Sub
Sub Frankenstein_UnPaused() : Controller.Pause = False : End Sub
Sub Frankenstein_Exit()   : Controller.Stop : End Sub

Sub InitTable()
  ' tilt
  vpmNudge.TiltSwitch   = 1
  vpmNudge.Sensitivity  = 1
  vpmNudge.TiltObj    = Array(Bumper13, Bumper14, Bumper15, Bumper16, LeftSlingshot, RightSlingshot)

  ' ball trough
  Set bsTrough = New cvpmBallStack
  With bsTrough
    .InitSw 0, 14, 13, 12, 11, 10, 9, 0
    .InitKick BallLockout, 80, 10
    .Balls = 6
  End With

  ' init lights, flippers, bumpers, ...
  InitLights InsertLights

  ' init timers
  PinMAMETimer.Interval         = PinMAMEInterval
    PinMAMETimer.Enabled          = True
  LampTimer.Enabled           = True
  GraphicsTimer.Interval        = 10
  GraphicsTimer.Enabled       = True
  'RollingSoundTimer.Interval     = 10
  'RollingSoundTimer.Enabled      = True
  GITimer.Interval          = 15
  FrankHeadTimer.Interval       = 17
  FrankHeadTimer.Enabled        = (AnimateFranksHead = 1 Or AnimateFranksHead = 2 Or AnimateFranksHead = 3)

  pSiderailsLockbar.Visible       = RailsVisible = 1 Or (RailsVisible = 2 And DesktopMode)
  pSiderailsLockbarGIOff.Visible    = pSiderailsLockbar.Visible

  Frankenstein.ColorGradeImage    = luts(lutpos)
End Sub


' ****************************************************
' keys
' ****************************************************
Sub Frankenstein_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
       Controller.Switch(62) = True
    VRshifter.objrotX = VRShifter.ObjrotX -20
    end If
  If keycode = AddCreditKey Then
    Select Case Int(Rnd * 3)
      Case 0: PlaySound "Coin_In_1", 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound "Coin_In_2", 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound "Coin_In_3", 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    Primary_FlipperButtonLeft.x = Primary_FlipperButtonLeft.x + 8
end If
  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
' : Controller.Switch(64) = True
    Primary_FlipperButtonRight.x = Primary_FlipperButtonRight.x - 8
end If
  If keycode = LeftTiltKey Then Nudge 90, 1 : SoundNudgeLeft      ' Sets the nudge angle and power
  If keycode = RightTiltKey Then Nudge 270, 1 : SoundNudgeRight   ' ^
  If keycode = CenterTiltKey Then Nudge 0, 1 : SoundNudgeCenter
If keycode = RightMagnaSave And EnableLUTChangeWithRightMagnaSave = 1 Then
    LUTBox.Visible = True
    lutpos = lutpos + 1 : If lutpos > UBound(luts) Then lutpos = 0 : End If
    Frankenstein.ColorGradeImage = luts(lutpos)
    Dim tekst : tekst = "LUT id: " & lutpos & " " & luts(lutpos)
    LUTBox.Text = tekst
    vpmTimer.AddTimer 3000, "If LUTBox.Text =" + chr(34) + tekst + chr(34) + " Then LUTBox.Visible = False'"
  End If
  If keycode = StartGameKey then Primary_StartButton.y = Primary_StartButton.y - 2 : Primary_StartButton2.y = Primary_StartButton2.y - 2
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Frankenstein_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Controller.Switch(62) = False
    VRshifter.objrotX = VRShifter.ObjrotX +20
  End If
  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
       ' Controller.Switch(63) = False
    Primary_FlipperButtonLeft.x = Primary_FlipperButtonLeft.x - 8
  End If
  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
      ' Controller.Switch(64) = False
    Primary_FlipperButtonRight.x = Primary_FlipperButtonRight.x + 8
  End If
  If keycode = StartGameKey then Primary_StartButton.y = Primary_StartButton.y + 2 : Primary_StartButton2.y = Primary_StartButton2.y + 2
  If vpmKeyUp(keycode) Then Exit Sub
End Sub


' ****************************************************
' *** solenoids
' ****************************************************
' tools
SolCallback(1)          = "SolLockout"
SolCallback(2)      = "SolBallRelease"
SolCallback(3)      = "SolAutoLaunch"
SolCallback(4)      = "SolLowerScoopEject"
SolCallback(5)      = "SolTopScoopEject"
SolCallback(6)      = "SolIceCaveEject"
SolCallback(7)      = "SolVUKEject"
SolCallback(8)      = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(9)      = "SolTrapLeftHand"
SolCallback(11)     = "SolGI"
SolCallback(12)     = "SolTrapRightHand"
SolCallback(13)     = "SolFranksArms"
SolCallback(14)     = "SolOrbitDiverter"
SolCallback(16)     = "SolKickback"

' flasher
SolCallback(15)     = "SolFlasher 0, "  ' around VUK
SolCallback(25)     = "SolFlasher 1, "  ' above ANK
SolCallback(26)     = "SolFlasher 2, "  ' above FR
SolCallback(27)     = "SolFlasher 3, "  ' above IN
SolCallback(28)     = "SolFlasher 4, "  ' at pop bumpers
SolCallback(29)     = "SolFlasher 5, "  ' at spinner
SolCallback(30)     = "SolFlasher 6, "  ' upper left corner
SolCallback(31)     = "SolFlasher 7, "  ' around VUK

' flipper
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sURFlipper) = "SolURFlipper"


' ******************************************************
' outhole, drain and ball release
' ******************************************************
Sub Drain_Hit()
  BallSearch
  'PlaySoundAtVol "drain", Drain, 0.01
    RandomSoundDrain Drain
    FrankIsWatching ActiveBall
  bsTrough.AddBall Me
  If bsTrough.Balls = 6 Then StopFranksHeadMode4
End Sub

Sub SolLockout(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX(SSolenoidOn, DOFContactors), BallLockout
    If bsTrough.Balls = 6 Then StartFranksHeadMode4
    If bsTrough.Balls > 0 Then bsTrough.SolOut True
  End If
End Sub
Sub BallRelease_Hit()   : Controller.Switch(15) = True  : End Sub
Sub BallRelease_Unhit() : Controller.Switch(15) = False : End Sub
Sub SolBallRelease(Enabled)
  If Enabled Then
    If Controller.Switch(15) Then
             RandomSoundBallRelease BallRelease
    BallRelease.Kick 85, 7
    End If
  End If
End Sub

' find balls that have fallen off the table
Sub BallSearch()
  Dim b
  For Each b In GetBalls
    If b.Y > 2200 Then
      b.X = 155 : b.Y = 550 : b.VelX = 0 : b.VelY = 0
    End If
  Next
End Sub


' ****************************************************
' flipper subs
' ****************************************************

Const Angle = 20
Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.Fire
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
'            RightUpperFlipper.RotateToEnd
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
'            RightUpperFlipper.RotateToEnd
    End If
  Else
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
'            RightUpperFlipper.RotateToStart
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolURFlipper(Enabled)
  If Enabled Then
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
            RightUpperFlipper.RotateToEnd
    Else
            RightUpperFlipper.RotateToEnd
    End If
  Else
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
            RightUpperFlipper.RotateToStart
    End If
  End If
End Sub

' flipper hit sounds

Sub RightUpperFlipper_Collide(parm)
   ' PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3), 1
End Sub


' ****************************************************
' sling shots and animations
' ****************************************************
Dim LeftStep, RightStep

Sub LeftSlingShot_Slingshot()
    LS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw 47
    RandomSoundSlingshotLeft pLeftSlingHammer
FrankIsWatching pLeftSlingHammer
    LeftStep = 0
  LeftSlingShot.TimerInterval = 5
  LeftSlingShot_Timer
End Sub
Sub LeftSlingShot_Timer()
    Select Case LeftStep
    Case 0: pRubberLeftSling1.Visible = False : pRubberLeftSling2.Visible = True : pLeftSlingHammer.TransZ = -12 : LeftSlingShot.TimerEnabled = True
        Case 1: pRubberLeftSling2.Visible = False : pRubberLeftSling3.Visible = True : pLeftSlingHammer.TransZ = -24 : LeftSlingShot.TimerInterval = 50
        Case 2: pRubberLeftSling3.Visible = False : pRubberLeftSling2.Visible = True : pLeftSlingHammer.TransZ = -12
        Case 3: pRubberLeftSling2.Visible = False : pRubberLeftSling1.Visible = True : pLeftSlingHammer.TransZ = 0 : LeftSlingShot.TimerEnabled = False
    End Select
    LeftStep = LeftStep + 1
End Sub

Sub RightSlingShot_Slingshot()
    LS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw 48
    RandomSoundSlingshotLeft pLeftSlingHammer
    FrankIsWatching pRightSlingHammer
    RightStep = 0
  RightSlingShot.TimerInterval = 5
  RightSlingShot_Timer
End Sub
Sub RightSlingShot_Timer()
    Select Case RightStep
    Case 0: pRubberRightSling1.Visible = False : pRubberRightSling2.Visible = True : pRightSlingHammer.TransZ = -12 : RightSlingShot.TimerEnabled = True
    Case 1: pRubberRightSling2.Visible = False : pRubberRightSling3.Visible = True : pRightSlingHammer.TransZ = -24 : RightSlingShot.TimerInterval = 50
        Case 2: pRubberRightSling3.Visible = False : pRubberRightSling2.Visible = True : pRightSlingHammer.TransZ = -12
        Case 3: pRubberRightSling2.Visible = False : pRubberRightSling1.Visible = True : pRightSlingHammer.TransZ = 0 : RightSlingShot.TimerEnabled = False
    End Select
    RightStep = RightStep + 1
End Sub


' ****************************************************
' bumpers
' ****************************************************
Sub Bumper13_Hit() : vpmTimer.PulseSw 41 : RandomSoundBumperTop Bumper13 : FrankIsWatching ActiveBall : End Sub
Sub Bumper14_Hit() : vpmTimer.PulseSw 42 : RandomSoundBumperLeft Bumper14 : FrankIsWatching ActiveBall : End Sub
'Sub Bumper15_Hit() : vpmTimer.PulseSw 43 : RandomSoundBumperTop Bumper15 : FrankIsWatching ActiveBall : End Sub
'Sub Bumper16_Hit() : vpmTimer.PulseSw 44 : RandomSoundBumperTop Bumper16 : FrankIsWatching ActiveBall : End Sub
Sub Bumper15_Hit() : vpmTimer.PulseSw 43 : RandomSoundBumperMiddle Bumper15 : FrankIsWatching ActiveBall : End Sub
Sub Bumper16_Hit() : vpmTimer.PulseSw 44 : RandomSoundBumperBottom Bumper16 : FrankIsWatching ActiveBall : End Sub


' ****************************************************
' switches
' ****************************************************
' orbit
Sub sw33_Hit()   : Controller.Switch(33) = True  : RandomSoundRollover : FrankIsWatching ActiveBall : End Sub
Sub sw33_Unhit() : Controller.Switch(33) = False : End Sub
Sub sw34_Hit()   : Controller.Switch(34) = True  : RandomSoundRollover : FrankIsWatching ActiveBall : End Sub
Sub sw34_Unhit() : Controller.Switch(34) = False : End Sub
Sub sw35_Hit()   : Controller.Switch(35) = True  : RandomSoundRollover : FrankIsWatching ActiveBall : End Sub
Sub sw35_Unhit() : Controller.Switch(35) = False : End Sub
Sub sw36_Hit()   : Controller.Switch(36) = True  : RandomSoundRollover : FrankIsWatching ActiveBall : End Sub
Sub sw36_Unhit() : Controller.Switch(36) = False : End Sub
Sub sw56_Hit()   : Controller.Switch(56) = True  : RandomSoundRollover : FrankIsWatching ActiveBall : End Sub
Sub sw56_Unhit() : Controller.Switch(56) = False : End Sub

' kickback at left outlane
Sub sw37_Hit()   : Controller.Switch(37) = True  : RandomSoundRollover : FrankIsWatching ActiveBall : End Sub
Sub sw37_Unhit() : Controller.Switch(37) = False : End Sub

' inlanes and outlanes
Sub sw38_Hit()   : Controller.Switch(38) = True  : RandomSoundRollover : FrankIsWatching ActiveBall : End Sub
Sub sw38_Unhit() : Controller.Switch(38) = False : End Sub
Sub sw39_Hit()   : Controller.Switch(39) = True  : RandomSoundRollover : FrankIsWatching ActiveBall : End Sub
Sub sw39_Unhit() : Controller.Switch(39) = False : End Sub
Sub sw40_Hit()   : Controller.Switch(40) = True  : RandomSoundRollover : FrankIsWatching ActiveBall : End Sub
Sub sw40_Unhit() : Controller.Switch(40) = False : End Sub

' shooter lane
Sub sw16_Hit()   : Controller.Switch(16) = True  : RandomSoundRollover : FrankIsWatching ActiveBall : End Sub
Sub sw16_UnHit() : Controller.Switch(16) = False : End Sub

' ramp trigger
Sub sw45_Hit()   : Controller.Switch(45) = True  : FrankIsWatching ActiveBall : End Sub
Sub sw45_Unhit() : Controller.Switch(45) = False : End Sub
Sub sw46_Hit()   : Controller.Switch(46) = True  : FrankIsWatching ActiveBall : End Sub
Sub sw46_Unhit() : Controller.Switch(46) = False : End Sub

' spinner
'Sub sw54_Spin()  : vpmTimer.PulseSw 54 : PlaySoundAt "fx_spinner", sw54 : End Sub SoundSpinner
Sub sw54_Spin()  : vpmTimer.PulseSw 54 : SoundSpinner sw54 : End Sub


' ****************************************************
' stand-up targets
' ****************************************************
Dim sw17Dir, sw18Dir, sw19Dir, sw20Dir, sw21Dir, sw22Dir, sw23Dir, sw24Dir, sw25Dir, sw26Dir, sw27Dir, sw28Dir, sw29Dir
Sub sw17_Hit()   : MoveStandUpTarget 17, sw17, sw17P, sw17PGIOff, sw17Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw17_Timer() : MoveStandUpTarget 17, sw17, sw17P, sw17PGIOff, sw17Dir, False : End Sub
Sub sw18_Hit()   : MoveStandUpTarget 18, sw18, sw18P, sw18PGIOff, sw18Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw18_Timer() : MoveStandUpTarget 18, sw18, sw18P, sw18PGIOff, sw18Dir, False : End Sub
Sub sw19_Hit()   : MoveStandUpTarget 19, sw19, sw19P, sw19PGIOff, sw19Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw19_Timer() : MoveStandUpTarget 19, sw19, sw19P, sw19PGIOff, sw19Dir, False : End Sub
Sub sw20_Hit()   : MoveStandUpTarget 20, sw20, sw20P, sw20PGIOff, sw20Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw20_Timer() : MoveStandUpTarget 20, sw20, sw20P, sw20PGIOff, sw20Dir, False : End Sub
Sub sw21_Hit()   : MoveStandUpTarget 21, sw21, sw21P, sw21PGIOff, sw21Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw21_Timer() : MoveStandUpTarget 21, sw21, sw21P, sw21PGIOff, sw21Dir, False : End Sub
Sub sw22_Hit()   : MoveStandUpTarget 22, sw22, sw22P, sw22PGIOff, sw22Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw22_Timer() : MoveStandUpTarget 22, sw22, sw22P, sw22PGIOff, sw22Dir, False : End Sub
Sub sw23_Hit()   : MoveStandUpTarget 23, sw23, sw23P, sw23PGIOff, sw23Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw23_Timer() : MoveStandUpTarget 23, sw23, sw23P, sw23PGIOff, sw23Dir, False : End Sub
Sub sw24_Hit()   : MoveStandUpTarget 24, sw24, sw24P, sw24PGIOff, sw24Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw24_Timer() : MoveStandUpTarget 24, sw24, sw24P, sw24PGIOff, sw24Dir, False : End Sub
Sub sw25_Hit()   : MoveStandUpTarget 25, sw25, sw25P, sw25PGIOff, sw25Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw25_Timer() : MoveStandUpTarget 25, sw25, sw25P, sw25PGIOff, sw25Dir, False : End Sub
Sub sw26_Hit()   : MoveStandUpTarget 26, sw26, sw26P, sw26PGIOff, sw26Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw26_Timer() : MoveStandUpTarget 26, sw26, sw26P, sw26PGIOff, sw26Dir, False : End Sub
Sub sw27_Hit()   : MoveStandUpTarget 27, sw27, sw27P, sw27PGIOff, sw27Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw27_Timer() : MoveStandUpTarget 27, sw27, sw27P, sw27PGIOff, sw27Dir, False : End Sub
Sub sw28_Hit()   : MoveStandUpTarget 28, sw28, sw28P, sw28PGIOff, sw28Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw28_Timer() : MoveStandUpTarget 28, sw28, sw28P, sw28PGIOff, sw28Dir, False : End Sub
Sub sw29_Hit()   : MoveStandUpTarget 29, sw29, sw29P, sw29PGIOff, sw29Dir, True : FrankIsWatching ActiveBall : End Sub
Sub sw29_Timer() : MoveStandUpTarget 29, sw29, sw29P, sw29PGIOff, sw29Dir, False : End Sub

Sub MoveStandUpTarget(id, sw, prim, primGIOff, counter, init)
  If init And prim.RotX < 10 Then
    OnStandupTargetHit
    vpmTimer.PulseSw id
    sw.TimerInterval = 5
    sw.TimerEnabled  = True
    counter = 5
  End If
  prim.TransY = prim.TransY - counter / 4
  prim.RotX = prim.RotX + counter
  primGIOff.TransY = prim.TransY
  primGIOff.RotX = prim.RotX
  If prim.RotX >= 10 Then counter = -0.5
  If prim.RotX = 0 Then sw.TimerEnabled = False
End Sub


' ****************************************************
' rotating spinner
' ****************************************************
Sub SpinnerTimer_Timer()
  pSpinner54.RotX     = 360 - sw54.CurrentAngle
  pSpinner54GIOff.RotX  = pSpinner54.RotX
    pSpinner54Rod.TransZ  = -Sin(sw54.CurrentAngle * 2 * 3.14 / 360) * 5
    pSpinner54Rod.TransX  = Sin((sw54.CurrentAngle - 90) * 2 * 3.14 / 360) * -5
End Sub


' ****************************************************
' saucer
' ****************************************************
Dim sw53Step : sw53Step = 0
Dim sw55Step : sw55Step = 0
Dim sw31Step : sw31Step = 0
Dim sw32Step : sw32Step = 0
Dim sw53Ball : Set sw53Ball = Nothing
Dim sw55Ball : Set sw55Ball = Nothing
Dim sw31Ball : Set sw31Ball = Nothing
Dim sw32Ball : Set sw32Ball = Nothing

Sub sw53_Hit()   : SoundSaucerLock : End Sub
Sub sw53_Unhit() : Controller.Switch(53) = False : End Sub
Sub sw55_Hit()   : SoundSaucerLock : End Sub
Sub sw55_Unhit() : Controller.Switch(55) = False : End Sub
'Sub sw31s_Hit()  : PlaySoundAt "fx_balldrop" & Int(Rnd*3), sw31s : End Sub
Sub sw31s_Hit()  : RandomSoundDelayedBallDropOnPlayfield sw31s: End Sub
Sub sw31_Hit()   : PlaySoundAt "fx_loosemetalplate", sw31 : End Sub
Sub sw31_Unhit() : Controller.Switch(31) = False : sw31e.Enabled = True : End Sub
Sub sw31e_Hit()  : KickScoop ActiveBall, 107, 25 : sw31e.Enabled = False : End Sub
'Sub sw32s_Hit()  : PlaySoundAt "fx_balldrop" & Int(Rnd*3), sw32s : End Sub
Sub sw32s_Hit()  : RandomSoundDelayedBallDropOnPlayfield sw32s: End Sub
Sub sw32_Hit()   : PlaySoundAt "fx_loosemetalplate", sw32 : End Sub
Sub sw32_Unhit() : Controller.Switch(32) = False : sw32e.Enabled = True : End Sub
'Sub sw32e_Hit()  : KickScoop ActiveBall, 169, 16 : sw32e.Enabled = False : End Sub
Sub sw32e_Hit()  : KickScoop ActiveBall, 165, 20 : sw32e.Enabled = False : End Sub

Sub SolIceCaveEject(Enabled)
  If Enabled Then
    MoveHammer pSaucer55Hammer, 0
    sw55Step      = 0
    sw55.TimerInterval  = 11
    sw55.TimerEnabled   = True
  End If
End Sub
Sub SolVUKEject(Enabled)
  If Enabled Then
    sw53Step      = 0
    sw53.TimerInterval  = 11
    sw53.TimerEnabled   = True
  End If
End Sub
Sub SolTopScoopEject(Enabled)
  If Enabled Then
    sw31Step      = 0
    sw31.TimerInterval  = 11
    sw31.TimerEnabled   = True
  End If
End Sub
Sub SolLowerScoopEject(Enabled)
  If Enabled Then
    sw32Step      = 0
    sw32.TimerInterval  = 11
    sw32.TimerEnabled   = True
  End If
End Sub

Sub sw55_Timer()
  ' ice cave
  SaucerAction sw55Step, pSaucer55Hammer, sw55Ball, 55, sw55, 200, 10
  sw55Step = sw55Step + 1
End Sub
Sub sw53_Timer()
  ' VUK
  SaucerAction sw53Step, Nothing, sw53Ball, 53, sw53, -1, 50+Rnd()*10
  sw53Step = sw53Step + 1
End Sub
Sub sw31_Timer()
  ' top scoop
  SaucerAction sw31Step, Nothing, sw31Ball, 31, sw31, 290, 150+Rnd()*20
  sw31Step = sw31Step + 1
End Sub
Sub sw32_Timer()
  ' bottom scoop
  SaucerAction sw32Step, Nothing, sw32Ball, 32, sw32, 345, 110+Rnd()*15
  sw32Step = sw32Step + 1
End Sub

Sub SaucerAction(step, pHammer, kBall, id, switch, kickangle, kickpower)
  Select Case step
    Case 0    : MoveHammer pHammer, -1 : PlaySoundAt SoundFX("fx_vuk_exit", DOFContactors), switch
    Case 1    : MoveHammer pHammer, 25 : If Controller.Switch(id) And Not (kBall Is Nothing) Then KickBall kBall, kickangle, IIF(kickpower>=100,1,10), kickpower, IIF(kickpower>=100,0,5), IIF(kickpower>=100,0,30)
    Case 13   : Controller.Switch(id) = False
    Case 23,24,25,26,27,28,29 : MoveHammer pHammer, -3
    Case 30   : MoveHammer pHammer,  0 : switch.TimerEnabled = False
    Case Else   : ' nothing to do
  End Select
End Sub

Sub KickBall(kBall, kAngle, kAngleVar, kVel, kVelZ, kLiftZ)
  If kBall Is Nothing Then Exit Sub
  Dim rAngle
  rAngle  = 3.14159265 * (kAngle + (kAngleVar / 2 - kAngleVar * Rnd()) - 90) / 180
  kVel  = kVel + (kVel/20 - kVel/10 * Rnd())
  kVelZ = kVelZ + (kVelZ/20 - kVelZ/10 * Rnd())
  With kBall
    .Z = .Z + kLiftZ
    If kAngle >= 0 Then
      .VelZ = kVelZ
      .VelX = cos(rAngle) * kVel
      .VelY = sin(rAngle) * kVel
    Else
      .VelZ = kVel
      If kAngle = -2 Then
        .VelX = Rnd() * 2
        .VelY = Rnd() * 2
      End If
    End If
  End With
  Set kBall = Nothing
End Sub

Sub MoveHammer(pHammer, rotZ)
  If Not (pHammer Is Nothing) Then
    If rotZ = 0 Then
      pHammer.RotZ = 0
    Else
      pHammer.RotZ = pHammer.RotZ + rotZ
    End If
  End If
End Sub

Sub KickScoop(actBall, angle, power)
  With actBall
    .VelX = 0 : .VelY = 0 : .VelZ = 0
    KickBall actBall, angle, 5, power, 1, 20
  End With
End Sub


' ****************************************************
' upper ramp traps
' ****************************************************
rUpperRampTrap1.Collidable  = False
rUpperRampTrap2.Collidable  = False
rUpperRamp002.Collidable  = True

Sub SolTrapLeftHand(Enabled)
  If Enabled Then
    OpenTrapDoor trUpperRampTrap1
  End If
End Sub
Sub SolTrapRightHand(Enabled)
  If Enabled Then
    OpenTrapDoor trUpperRampTrap2
  End If
End Sub

Sub trUpperRampTrap1_Hit()
  Set ballFrankLeftHand     = ActiveBall
  'isBallOnWireRamp = False
  rUpperRampTrap1.Collidable    = True
  rUpperRamp002.Collidable    = False
  PlaySoundAt "ferris_ball_drop", trUpperRampTrap1
End Sub
Sub kFrankLeftHand_Hit()
  rUpperRampTrap1.Collidable  = False
  rUpperRamp002.Collidable  = True
  PlaySoundAt "ferris_hit" & Int(Rnd()*2+1), kFrankLeftHand
' vpmTimer.AddTimer 15000, "SolFranksArms True'"
End Sub

Sub trUpperRampTrap2_Hit()
  Set ballFrankRightHand      = ActiveBall
  'isBallOnWireRamp = False
  rUpperRampTrap2.Collidable    = True
  rUpperRamp002.Collidable    = False
  PlaySoundAt "ferris_ball_drop", trUpperRampTrap2
End Sub
Sub kFrankRightHand_Hit()
  rUpperRampTrap2.Collidable  = False
  rUpperRamp002.Collidable  = True
  PlaySoundAt "ferris_hit" & Int(Rnd()*2+1), kFrankRightHand
' vpmTimer.AddTimer 15000, "SolFranksArms True'"
End Sub

Sub OpenTrapDoor(sw)
  rUpperRamp002.Collidable = True
  sw.Enabled = True
  PlaySoundAt SoundFX(SSolenoidOn, DOFContactors), sw
  sw.TimerInterval = 9 : sw.TimerEnabled = True
End Sub
Sub CloseTrapDoor(sw)
  rUpperRamp002.Collidable = True
  sw.Enabled = False
  PlaySoundAt SoundFX(SSolenoidOn, DOFContactors), sw
  sw.TimerInterval = 11 : sw.TimerEnabled = True
End Sub

Sub trUpperRampTrap1_Timer()
  If trUpperRampTrap1.TimerInterval = 9 Then
    If pTrap1.ObjRotY <= -30 Then
      trUpperRampTrap1.TimerEnabled = False : PlaySoundAt "fx_metalhit0", trUpperRampTrap1
    Else
      pTrap1.ObjRotY = pTrap1.ObjRotY - 10
    End If
  Else
    If pTrap1.ObjRotY >= 0 Then
      trUpperRampTrap1.TimerEnabled = False : PlaySoundAt "fx_sensor", trUpperRampTrap1
    Else
      pTrap1.ObjRotY = pTrap1.ObjRotY + 2
    End If
  End If
End Sub
Sub trUpperRampTrap2_Timer()
  If trUpperRampTrap2.TimerInterval = 9 Then
    If pTrap2.ObjRotY <= -30 Then
      trUpperRampTrap2.TimerEnabled = False : PlaySoundAt "fx_metalhit0", trUpperRampTrap2
    Else
      pTrap2.ObjRotY = pTrap2.ObjRotY - 10
    End If
  Else
    If pTrap2.ObjRotY >= 0 Then
      trUpperRampTrap2.TimerEnabled = False : PlaySoundAt "fx_sensor", trUpperRampTrap2
    Else
      pTrap2.ObjRotY = pTrap2.ObjRotY + 2
    End If
  End If
End Sub


' ****************************************************
' Frank
' ****************************************************
Dim ballFrankLeftHand  : Set ballFrankLeftHand = Nothing
Dim ballFrankRightHand : Set ballFrankRightHand = Nothing

Sub SolFranksArms(Enabled)
  If Enabled Then
    MoveArms
    rUpperRamp002.Collidable  = True
    rUpperRampTrap1.Collidable  = False
    rUpperRampTrap2.Collidable  = False
  End If
End Sub

Sub MoveArms()
  FrankArmsTimer.Interval = 3
  FrankArmsTimer.Enabled  = True
  PlaySoundAt SoundFX("ferris_exit", DOFContactors), pFrankHead
End Sub
Sub FrankArmsTimer_Timer()
  If FrankArmsTimer.Interval = 7 Then
    ' move the arms back
    pFrankLeftArm.ObjRotX = pFrankLeftArm.ObjRotX + 0.5
    pFrankLeftArmGIOff.ObjRotX = pFrankLeftArm.ObjRotX
    pFrankLeftArmFOnGIOff.ObjRotX = pFrankLeftArm.ObjRotX
    pFrankRightArm.ObjRotX = pFrankRightArm.ObjRotX + 0.5
    pFrankRightArmGIOff.ObjRotX = pFrankRightArm.ObjRotX
    pFrankRightArmFOnGIOff.ObjRotX = pFrankRightArm.ObjRotX
    If pFrankLeftArm.ObjRotX >= 0 Then FrankArmsTimer.Enabled = False
  ElseIf FrankArmsTimer.Interval = 5 Then
    ' throw the ball(s)
    ThrowBall ballFrankLeftHand, kFrankLeftHand, trUpperRampTrap1
    ThrowBall ballFrankRightHand, kFrankRightHand, trUpperRampTrap2
    FrankArmsTimer.Interval = 7
  Else
    ' move the arms to the front
    pFrankLeftArm.ObjRotX = pFrankLeftArm.ObjRotX - 5
    pFrankLeftArmGIOff.ObjRotX = pFrankLeftArm.ObjRotX
    pFrankLeftArmFOnGIOff.ObjRotX = pFrankLeftArm.ObjRotX
    MoveBall ballFrankLeftHand
    pFrankRightArm.ObjRotX = pFrankRightArm.ObjRotX - 5
    pFrankRightArmGIOff.ObjRotX = pFrankRightArm.ObjRotX
    pFrankRightArmFOnGIOff.ObjRotX = pFrankRightArm.ObjRotX
    MoveBall ballFrankRightHand
    If pFrankLeftArm.ObjRotX <= -25 Then FrankArmsTimer.Interval = 5
  End If
End Sub

Sub MoveBall(ballInFranksHand)
  If Not (ballInFranksHand Is Nothing) Then
    With ballInFranksHand
      .Y = .Y + 8 : .Z = .Z + 9
    End With
  End If
End Sub
Sub ThrowBall(ballInFranksHand, ballKicker, trigger)
  ballKicker.Kick 180, 20
  CloseTrapDoor trigger
  If Not (ballInFranksHand Is Nothing) Then
    'isBallOnWireRamp = False
    Set ballInFranksHand = Nothing
  End If
End Sub

Dim newAngleF, directionF, speedF, breakF, headStateF, headStateTimerF
newAngleF = 0
directionF = 0
Sub FrankHeadTimer_Timer()
  If AnimateFranksHead = 1 Then
    FrankHeadTimer.Interval = 55
    Const maxAngle1 = 45
    If directionF = 0 Or (directionF = -1 And newAngleF >= pFrankHead.ObjRotZ) Or (directionF = 1 And newAngleF <= pFrankHead.ObjRotZ) Then
      If breakF > 0 Then
        breakF = breakF - 1
      Else
        speedF = Rnd()*2.5 + 0.5
        breakF = Int(Rnd()*25)
        newAngleF = Int(Rnd()*(2*maxAngle1+1)) - maxAngle1
        If newAngleF < pFrankHead.ObjRotZ Then
          directionF = -1
        Else
          directionF = 1
        End If
      End If
      StopMotorSound
    Else
      ' move Frank's head one step
      pFrankHead.ObjRotZ = pFrankHead.ObjRotZ + speedF * directionF
      StartMotorSound
    End If
  ElseIf AnimateFranksHead = 2 Then
    Const maxAngle2 = 45 ' 75
    Const rotStep = 3
    ' identify the nearest ball
    Dim balls, b, distance, nearestBall, nearestDistance
    nearestBall   = -1
    nearestDistance = 0
    balls       = GetBalls
    For b = 0 to UBound(balls)
      distance = SQR((pFrankHead.X - balls(b).X) ^ 2 + (pFrankHead.Y - balls(b).Y) ^ 2)
      If nearestBall = -1 Or distance < nearestDistance Then
        nearestBall   = b
        nearestDistance = distance
      End If
    Next
    ' follow that ball with Frank's head
    If nearestBall > -1 Then
      Dim angle, newAngle
      angle = GetFranksHeadAngle(balls(nearestBall))
      newAngle = pFrankHead.ObjRotZ
      If pFrankHead.ObjRotZ < angle Then
        newAngle = pFrankHead.ObjRotZ + rotStep
        If newAngle > angle Then newAngle = angle
        If newAngle > maxAngle2 Then newAngle = maxAngle2
      ElseIf pFrankHead.ObjRotZ > angle Then
        newAngle = pFrankHead.ObjRotZ - rotStep
        If newAngle < angle Then newAngle = angle
        If newAngle < -maxAngle2 Then newAngle = -maxAngle2
      End If
      pFrankHead.ObjRotZ = newAngle : StartMotorSound
    Else
      pFrankHead.ObjRotZ = 0 : StopMotorSound
    End If
  ElseIf AnimateFranksHead = 3 Then
    FrankHeadTimer.Interval = 25
    Const maxAngle3 = 45
    If Abs(newAngleF) > maxAngle3 Then
      If newAngleF < 0 Then newAngleF = -maxAngle3 Else newAngleF = maxAngle3
    End If
    If (directionF = -1 And newAngleF >= pFrankHead.ObjRotZ) Or (directionF = 1 And newAngleF <= pFrankHead.ObjRotZ) Then
      directionF = 0
      newAngleF = pFrankHead.ObjRotZ
    End If
    If directionF = 0 Then
      If newAngleF <> pFrankHead.ObjRotZ Then
        speedF = Rnd()*3.5 + 2.5
        directionF = 99 ' any value but not 1 or -1
      End If
      StopMotorSound
    Else
      If newAngleF < pFrankHead.ObjRotZ Then
        directionF = -1
      Else
        directionF = 1
      End If
      ' move Frank's head one step
      pFrankHead.ObjRotZ = pFrankHead.ObjRotZ + speedF * directionF
      StartMotorSound
    End If
  ElseIf AnimateFranksHead = 4 Then
    FrankHeadTimer.Interval = 11
    speedF = 1
    Select Case headStateF
    Case -1
      If pFrankHead.ObjRotZ > 0 Then pFrankHead.ObjRotZ = pFrankHead.ObjRotZ - 1 : StartMotorSound
      If pFrankHead.ObjRotZ < 0 Then pFrankHead.ObjRotZ = pFrankHead.ObjRotZ + 1 : StartMotorSound
      If pFrankHead.ObjRotZ = 0 Then FrankHeadTimer.Enabled = False : StopMotorSound
    Case 1
      If pFrankHead.ObjRotz > -17 Then
        pFrankHead.ObjRotz = pFrankHead.ObjRotz - speedF : StartMotorSound
      Else
        StopMotorSound
      End If
      headStateTimerF = headStateTimerF + 1
      If headStateTimerF >= 200 Then headStateF = 2
    Case 2
      If pFrankHead.ObjRotz > -35 Then
        pFrankHead.ObjRotz = pFrankHead.ObjRotz - speedF : StartMotorSound
      Else
        StopMotorSound
      End If
      If pFrankHead.ObjRotz = -35 Then headStateF = 3
    Case 3
      If pFrankHead.ObjRotz < 0 Then
        pFrankHead.ObjRotz = pFrankHead.ObjRotz + speedF : StartMotorSound
      Else
        StopMotorSound
      End If
      headStateTimerF = headStateTimerF + 1
      If headStateTimerF >= 300 Then headStateF = 4
    Case 4
      If pFrankHead.ObjRotz > -35 Then
        pFrankHead.ObjRotz = pFrankHead.ObjRotz - speedF : StartMotorSound
      Else
        StopMotorSound
      End If
      headStateTimerF = headStateTimerF + 1
      If headStateTimerF >= 700 Then headStateF = 5
    Case 5
      If pFrankHead.ObjRotz < 0 Then
        pFrankHead.ObjRotz = pFrankHead.ObjRotz + speedF : StartMotorSound
      Else
        StopMotorSound
      End If
      headStateTimerF = headStateTimerF + 1
      If headStateTimerF >= 800 Then headStateF = 6
    Case 6
      If pFrankHead.ObjRotz < 17 Then
        pFrankHead.ObjRotz = pFrankHead.ObjRotz + speedF : StartMotorSound
      Else
        StopMotorSound
      End If
      headStateTimerF = headStateTimerF + 1
      If headStateTimerF >= 850 Then headStateF = 7
    Case 7
      If pFrankHead.ObjRotz < 35 Then
        pFrankHead.ObjRotz = pFrankHead.ObjRotz + speedF : StartMotorSound
      Else
        StopMotorSound
      End If
      headStateTimerF = headStateTimerF + 1
      If headStateTimerF >= 900 Then headStateF = 8
    Case 8
      If pFrankHead.ObjRotz > 0 Then
        pFrankHead.ObjRotz = pFrankHead.ObjRotz - speedF : StartMotorSound
      Else
        StopMotorSound
      End If
      headStateTimerF = headStateTimerF + 1
      If headStateTimerF >= 1100 Then headStateF = 1 : headStateTimerF = 1
    End Select
  End If
  pFrankHeadGIOff.ObjRotZ = pFrankHead.ObjRotZ
  pFrankHeadFOnGIOff.ObjRotZ = pFrankHead.ObjRotZ
End Sub

Sub FrankIsWatching(aObject)
  newAngleF = GetFranksHeadAngle(aObject)
End Sub

Function GetFranksHeadAngle(aObject)
  If aObject Is Nothing Then Exit Function
  Dim aa, bb, cc, xx
  aa = pFrankHead.X - aObject.X
  bb = pFrankHead.Y - aObject.Y
  cc = SQR(aa^2 + bb^2)
  xx = aa / cc
  GetFranksHeadAngle = Atn(xx / Sqr(-xx * xx + 1)) * 180 / CSng(4*Atn(1))
End Function

Sub StartFranksHeadMode4()
  If AnimateFranksHead <> 4 Then Exit Sub
  headStateF = 1 : headStateTimerF = 1
  FrankHeadTimer.Enabled = True
End Sub
Sub StopFranksHeadMode4()
  If AnimateFranksHead <> 4 Then Exit Sub
  headStateF = -1 : headStateTimerF = 1
End Sub

Dim isMotorSOundOn : isMotorSoundOn = False
Sub StartMotorSound()
  If isMotorSoundOn Then Exit Sub
  isMotorSoundOn = True
  PlayLoopSoundAtVol "fx_motor", pFrankBody, HeadMotorSoundFactor
End Sub
Sub StopMotorSound()
  isMotorSoundOn = False
  StopSound "fx_motor"
End Sub


' ****************************************************
' diverter
' ****************************************************
rDiverter.Collidable = False

Sub SolOrbitDiverter(Enabled)
  PlaySoundAt "fx_sensor", sw35
  FrankIsWatching sw35
  rDiverter.Collidable = Enabled
End Sub


' ****************************************************
' auto launch
' ****************************************************
AutoPlunger.PullBack

Sub SolAutoLaunch(Enabled)
  vpmSolAutoPlungeS AutoPlunger, SoundFX(SSolenoidOn, DOFContactors), 5, Enabled
End Sub


' ****************************************************
' kick back
' ****************************************************
Sub SolKickback(Enabled)
    Kickback.Enabled = Enabled
End Sub

Dim kickbackBallVel : kickbackBallVel = 1
Dim kickbackBall    : Set kickbackBall = Nothing
Sub Kickback_Hit()
  Kickback.TimerEnabled  = False
  kickbackBallVel  = BallVel(ActiveBall)
  Set kickbackBall = ActiveBall
  PlaySoundAt SoundFX(SSolenoidOn, DOFContactors), Kickback
  Kickback.TimerInterval = 40
    Kickback.TimerEnabled  = True
End Sub
Sub Kickback_Timer()
  If Kickback.TimerInterval = 40 Then
    Kickback.Kick 0, Int(60 + kickbackBallVel/2 + Rnd()*10)
    kickbackBall.VelX = 0
    'PlaySoundAt SoundFX("plunger", DOFContactors), Kickback
        SoundPlungerReleaseBall2
  pKickback.TransY = 0
    Kickback.TimerInterval = 20
  ElseIf Kickback.TimerInterval = 20 Then
    pKickback.TransY = pKickback.TransY + 10
    If pKickback.TransY = 50 Then Kickback.TimerInterval = 500
  ElseIf Kickback.TimerInterval = 500 Then
    Kickback.TimerInterval = 50
  ElseIf Kickback.TimerInterval = 50 Then
    pKickback.TransY = pKickback.TransY - 5
    If pKickback.TransY = 0 Then Kickback.TimerInterval = 1000
  ElseIf Kickback.TimerInterval = 1000 Then
    Kickback.TimerEnabled = False
  End If
End Sub


' *********************************************************************
' lamps and illumination
' *********************************************************************
' inserts
Dim PFLights(200,3), PFLightsCount(200), isGIOn, isGIChanged
isGIOn = False
isGIChanged = False

Sub InitLights(aColl)
  ' init inserts
  Dim obj, idx
  For Each obj In aColl
    idx = obj.TimerInterval
    Set PFLights(idx, PFLightsCount(idx)) = obj
    PFLightsCount(idx) = PFLightsCount(idx) + 1
    If Right(obj.Name,1) = "a" Then obj.IntensityScale = EnableReflectionsAtBall
  Next
  InitGI
End Sub
Sub LampTimer_Timer()
  Dim chgLamp, num, chg, ii, nr, xxx, obj
  xxx = -1
    chgLamp = Controller.ChangedLamps
  If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
      Select Case chglamp(ii,0)
      ' bumper lights
      Case 13,14,15,16
        If chgLamp(ii,1) = 1 Then aBumperValue(chglamp(ii,0)-13) = -99 Else aBumperValue(chglamp(ii,0)-13) = 3
      ' tiny plastic lamps
      Case 29
        fPlasticBulb29.Visible  = (chglamp(ii,1) = 1)
        lPlasticBulb29.State  = IIF(chglamp(ii,1) = 1, LightStateOn, LightStateOff)
      Case 30
        fPlasticBulb30.Visible  = (chglamp(ii,1) = 1)
        lPlasticBulb30.State  = IIF(fPlasticBulb30.Visible Or fPlasticBulb31.Visible Or fPlasticBulb32.Visible, LightStateOn, LightStateOff)
      Case 31
        fPlasticBulb31.Visible  = (chglamp(ii,1) = 1)
        lPlasticBulb30.State  = IIF(fPlasticBulb30.Visible Or fPlasticBulb31.Visible Or fPlasticBulb32.Visible, LightStateOn, LightStateOff)
      Case 32
        fPlasticBulb32.Visible  = (chglamp(ii,1) = 1)
        lPlasticBulb30.State  = IIF(fPlasticBulb30.Visible Or fPlasticBulb31.Visible Or fPlasticBulb32.Visible, LightStateOn, LightStateOff)
      Case 48
        fPlasticBulb48.Visible  = (chglamp(ii,1) = 1)
        lPlasticBulb48.State  = IIF(chglamp(ii,1) = 1, LightStateOn, LightStateOff)
      Case 56
        fPlasticBulb56.Visible  = (chglamp(ii,1) = 1)
        lPlasticBulb56.State  = IIF(chglamp(ii,1) = 1, LightStateOn, LightStateOff)
      Case Else
        For nr = 1 to PFLightsCount(chgLamp(ii,0)) : PFLights(chgLamp(ii,0),nr - 1).State = chgLamp(ii,1) : Next
      End Select
        Next
  End If
End Sub

' GI
Dim GILevel : GILevel = 0


Sub SolGI(IsOff)
    SetGI Not IsOff
End Sub


Sub InitGI()
  Dim obj
  pFlasher2Aa.Visible = False
  pFlasher6Aa.Visible = False
  pFlasher7Aa.Visible = False
  pFlasher15Aa.Visible = False
Dim k
For Each k In GiElementsDyn.Keys
    Set obj = GiElementsDyn(k)
    obj.Visible = False
Next

  UseGIMaterial GIWireTrigger, "Metal0.8", 0
    UseGIMaterial sGates, "Metal Chrome S34", 0
    UseGIMaterial Array(kIceCave,kVUK,pSpinner54Rod), "Metal Chrome S34", 0
End Sub
Sub SetGI(IsOn)
    If EnableGI = 0 And Not isGIOn Then Exit Sub
    If isGIOn <> IsOn Then
        isGIOn = IsOn
        isGIChanged = True
        If isGIOn Then
            ' GI goes on
            PlaySoundAtVol "fx_relay_on", pFrankHead, 1
            GILevel = 1 ' IIF(FastMode,4,1)
            DOF 101, DOFOn
        Else
            ' GI goes off
            PlaySoundAtVol "fx_relay_off", pFrankHead, 1
            GILevel = -4
            DOF 101, DOFOff
        End If
        InvalidateFrank
    End If
End Sub



Sub SolFlasherb(id, enabled)
  If enabled Then aFlasherValue(id) = -99 Else aFlasherValue(id) = 3
End Sub

Dim aBumper     : aBumper     = Array(pBumperCap13Off, pBumperCap14Off, pBumperCap15Off, pBumperCap16Off)
Dim aBumperOn     : aBumperOn   = Array(pBumperCap13, pBumperCap14, pBumperCap15, pBumperCap16)
Dim aBumperLight  : aBumperLight  = Array(lBumper13, lBumper14, lBumper15, lBumper16)
Dim aBumperValue  : aBumperValue  = Array(-99, -99, -99, -99)
Dim aFlasher    : aFlasher    = Array(pFlasher15AaOff, Nothing, pFlasher2AaOff, Nothing, Nothing, Nothing, pFlasher6AaOff, pFlasher7AaOff)
Dim aFlasherOn    : aFlasherOn  = Array(pFlasher15Aa, Nothing, pFlasher2Aa, Nothing, Nothing, Nothing, pFlasher6Aa, pFlasher7Aa)
Dim aFlasherLight   : aFlasherLight = Array(Array(lFlasher15Ab,lFlasher15Ac,lFlasher15Ad,lFlasher15Ae,lFlasher15Af,lFlasher15Ag,lFlasher15Az,lFlasher15Ba,lFlasher15Ca,lFlasher15Da,lFlasher15Db), _
                      Array(lFlasher1Aa), _
                      Array(lFlasher2Ab,lFlasher2Ac,lFlasher2Ad,lFlasher2Ae,lFlasher2Af,lFlasher2Ag,lFlasher2Ah,lFlasher2Az,lFlasher2Ba), _
                      Array(lFlasher3Aa), _
                      Array(lFlasher4Aa,lFlasher4Ba,lFlasher4Bb,lFlasher4Ca,lFlasher4Cb,lFlasher4Da,lFlasher4Db), _
                      Array(lFlasher5Aa,lFlasher5Ba,lFlasher5Bb), _
                      Array(lFlasher6Ab,lFlasher6Ac,lFlasher6Ad,lFlasher6Ae,lFlasher6Az,lFlasher6Ba,lFlasher6Ca), _
                      Array(lFlasher7Ab,lFlasher7Ac,lFlasher7Ad,lFlasher7Ae,lFlasher7Af,lFlasher7Ag,lFlasher7Az,lFlasher7Ba,lFlasher7Ca))
Dim aFlasherValue   : aFlasherValue = Array(0, 0, 0, 0, 0, 0, 0, 0)
Sub LightsTimer_Timer()
Dim k
  Dim idx, obj, coll
  ' set bumper and flasher off images
  If isGIChanged Then
    For idx = 0 To UBound(aBumper)
      If Not aBumper(idx) Is Nothing Then
        aBumper(idx).Image = IIF(isGIOn, "BumperCap Light-Off GI-On", "BumperCap Light-Off GI-Off")
      End If
    Next
    For idx = 0 To UBound(aFlasher)
      If Not aFlasher(idx) Is Nothing Then
        aFlasher(idx).Image = IIF(isGIOn, "Flasher Light-Off GI-On", "Flasher Light-Off GI-Off")
      End If
    Next
    For Each obj In Array(pRubberLeftSling1,pRubberLeftSling2,pRubberLeftSling3,pRubberRightSling1,pRubberRightSling2,pRubberRightSling3)
      obj.Image = IIF(isGIOn, "SP-Wire-GuidesTexComb", "SP-Wire-GuidesTexComb GI-Off")
    Next
    isGIChanged = False
  End If
  ' bumper lights
  For idx = 0 To UBound(aBumper)
    If aBumperValue(idx) > -1 Or aBumperValue(idx) = -99 Then
      If aBumperValue(idx) = -99 Then
        aBumperOn(idx).Visible  = True
        aBumper(idx).Material   = "Prims platt4"
        aBumper(idx).Visible  = False
        If IsArray(aBumperLight(idx)) Then
          For Each obj In aBumperLight(idx) : obj.State = LightStateOn : obj.IntensityScale = 1 : Next
        Else
          If Not aBumperLight(idx) Is Nothing Then aBumperLight(idx).State = LightStateOn : aBumperLight(idx).IntensityScale = 1 : End If
        End If
      ElseIf aBumperValue(idx) = 0 Then
        aBumper(idx).Visible  = True
        aBumper(idx).Material   = "Prims platt"
        aBumperOn(idx).Visible  = False
        If IsArray(aBumperLight(idx)) Then
          For Each obj In aBumperLight(idx) : obj.State = LightStateOn : obj.IntensityScale = 0 : Next
        Else
          If Not aBumperLight(idx) Is Nothing Then aBumperLight(idx).State = LightStateOff : aBumperLight(idx).IntensityScale = 0 : End If
        End If
      ElseIf aBumperValue(idx) <= 8 Then
        aBumperOn(idx).Visible  = True
        aBumper(idx).Visible  = True
        aBumper(idx).Material   = "Prims platt" & aBumperValue(idx)
        If IsArray(aBumperLight(idx)) Then
          For Each obj In aBumperLight(idx) : obj.IntensityScale = aBumperValue(idx) / 4 : Next
        Else
          If Not aBumperLight(idx) Is Nothing Then aBumperLight(idx).IntensityScale = aBumperValue(idx) / 4 : End If
        End If
      End If
      aBumperValue(idx) = aBumperValue(idx) - 1
    End If
  Next
  ' flasher
  For idx = 0 To UBound(aFlasher)
    If aFlasherValue(idx) > -1 Or aFlasherValue(idx) = -99 Then
      If aFlasherValue(idx) = -99 Then
        ' flasher is on
        If Not aFlasher(idx) Is Nothing Then
          aFlasherOn(idx).Visible = True
          aFlasher(idx).Material  = "Prims platt4"
          aFlasher(idx).Visible   = False
          IlluminateFrank idx, aFlasherValue(idx)
        End If
        If IsArray(aFlasherLight(idx)) Then
          For Each obj In aFlasherLight(idx) : obj.State = LightStateOn : obj.IntensityScale = 1 : Next
        Else
          If Not aFlasherLight(idx) Is Nothing Then aFlasherLight(idx).State = LightStateOn : aFlasherLight(idx).IntensityScale = 1
        End If
      ElseIf aFlasherValue(idx) = 0 Then
        ' flasher is off
        If Not aFlasher(idx) Is Nothing Then
          aFlasher(idx).Visible   = True
          aFlasher(idx).Material  = "Prims platt"
          aFlasherOn(idx).Visible = False
          IlluminateFrank idx, aFlasherValue(idx)
        End If
        If IsArray(aFlasherLight(idx)) Then
          For Each obj In aFlasherLight(idx) : obj.State = LightStateOff : obj.IntensityScale = 0 : Next
        Else
          If Not aFlasherLight(idx) Is Nothing Then aFlasherLight(idx).State = LightStateOff : aFlasherLight(idx).IntensityScale = 0
        End If
      ElseIf aFlasherValue(idx) <= 4 Then
        ' flasher goes off
        If Not aFlasher(idx) Is Nothing Then
          aFlasher(idx).Visible   = True
          aFlasher(idx).Material  = "Prims platt" & aFlasherValue(idx)
          aFlasherOn(idx).Visible = False
          IlluminateFrank idx, aFlasherValue(idx)
        End If
        If IsArray(aFlasherLight(idx)) Then
          For Each obj In aFlasherLight(idx) : obj.IntensityScale = aFlasherValue(idx) / 4 : Next
        Else
          If Not aFlasherLight(idx) Is Nothing Then aFlasherLight(idx).IntensityScale = aFlasherValue(idx) / 4
        End If
      End If
      aFlasherValue(idx) = aFlasherValue(idx) - 1
    End If
  Next
' ' GI
  If GILevel > 0 And GILevel < 5 Then
    ' GI goes on, from 1 to 4
    For Each obj In GIBulbs
      obj.IntensityScale = GILevel / 4
      If GILevel > 0 Then obj.State = LightStateOn
    Next
    If GILevel = 4 Then
      For Each obj In GIFlasher : obj.Visible = True : Next
    End If
    ' playfield and playfield overlay
    wPlayfieldGIOff.TopMaterial   = "Playfield" & GILevel
    wPlayfieldGIOff.Visible     = (wPlayfieldGIOff.TopMaterial <> "Playfield4")
    If DesktopMode Then
      wTextlayerGIOff.Visible   = False
      wTextlayer.Image      = "Playfield Textlayer"
    Else
      wTextlayerGIOff.TopMaterial = "PlayfieldText" & GILevel
      wTextlayerGIOff.Visible   = (wTextlayerGIOff.TopMaterial <> "PlayfieldText4")
      wTextlayer.TopMaterial    = "PlayfieldText" & IIF((4-GILevel)=0,"",4-GILevel)
      wTextlayer.Visible      = (wTextlayer.TopMaterial <> "PlayfieldText4")
    End If
    ' Frank
    pFrankHeadGIOff.Material    = "Prims platt" & GILevel
    pFrankBodyGIOff.Material    = "Prims platt" & GILevel
    pFrankLeftArmGIOff.Material   = "Prims platt" & GILevel
    pFrankRightArmGIOff.Material  = "Prims platt" & GILevel
    pFrankHeadGIOff.Visible     = (pFrankHeadGIOff.Material <> "Prims platt4")
    pFrankBodyGIOff.Visible     = pFrankHeadGIOff.Visible
    pFrankLeftArmGIOff.Visible    = pFrankHeadGIOff.Visible
    pFrankRightArmGIOff.Visible   = pFrankHeadGIOff.Visible
    ' playfield elements
'Dim k, obj
'For Each k In GiElementsGiOffDyn.Keys
'    Set obj = GiElementsGiOffDyn(k)
'    obj.Material = "Prims platt" & GILevel
'    obj.Visible  = pFrankHeadGIOff.Visible
'Next

'Dim k, obj

For Each k In GiElementsGiOffDyn.Keys
    Set obj = GiElementsGiOffDyn(k)

    ' --- Default material ---
    Dim newMat : newMat = "Prims platt" & GILevel

    ' --- Special cases by NAME (safe even if VR-only objects don't exist) ---
    Select Case obj.Name
        Case "pPlasticRamp_VR_GiOn", "pPlasticRamp_VR_GiOff"
            newMat = "Prims platt ramptrans"

        Case "WireRampVR_GI_ON", "WireRampVR_GI_OFF"
            newMat = "Metal S34 GoldON"
    End Select

    ' --- Apply material and visibility ---
    obj.Material = newMat
    obj.Visible  = pFrankHeadGIOff.Visible
Next

    pSiderailsLockbarGIOff.Material = "Prims platt" & GILevel
    pSiderailsLockbarGIOff.Visible  = (pFrankHeadGIOff.Visible And (RailsVisible = 1 Or (RailsVisible = 2 And DesktopMode)))
    If GILevel = 1 Then
'Dim k, obj
For Each k In GiElementsDyn.Keys
    Set obj = GiElementsDyn(k)
    obj.Visible = True
Next

      SetSidewallsVisibility True
      pSiderailsLockbar.Visible = (RailsVisible = 1 Or (RailsVisible = 2 And DesktopMode))
    End If
    UseGIMaterial GIWireTrigger, "Metal0.8", GILevel
    UseGIMaterial sGates, "Metal Chrome S34", GILevel
    UseGIMaterial Array(kIceCave,kVUK,pSpinner54Rod), "Metal Chrome S34", GILevel
    If GILevel = 4 Then AdjustLights
    GILevel = GILevel + 1
  ElseIf GILevel < 0 And GILevel > -5 Then
    ' GI goes off, from -4 to -1
    For Each obj In GIBulbs
      obj.IntensityScale = (Abs(GILevel)-1) / 4
      If GILevel = -1 Then obj.State = LightStateOff
    Next
    If GILevel = -1 Then
      For Each obj In GIFlasher : obj.Visible = False : Next
    End If
    ' playfield and playfield overlay
    wPlayfieldGIOff.TopMaterial   = "Playfield" & IIF(GILevel=-1,"",Abs(GILevel+1))
    wPlayfieldGIOff.Visible     = (wPlayfieldGIOff.TopMaterial <> "Playfield4")
    If DesktopMode Then
      wTextlayerGIOff.Visible   = False
      wTextlayer.Image      = "Playfield Textlayer GI-Off"
    Else
      wTextlayerGIOff.TopMaterial = "PlayfieldText" & IIF(GILevel=-1,"",Abs(GILevel+1))
      wTextlayerGIOff.Visible   = (wTextlayerGIOff.TopMaterial <> "PlayfieldText4")
      wTextlayer.TopMaterial    = "PlayfieldText" & (5+GILevel)
      wTextlayer.Visible      = (wTextlayer.TopMaterial <> "PlayfieldText4")
    End If
    ' Frank
    pFrankHeadGIOff.Material    = "Prims platt" & IIF(GILevel=-1," Frank",Abs(GILevel+1))
    pFrankBodyGIOff.Material    = "Prims platt" & IIF(GILevel=-1," Frank",Abs(GILevel+1))
    pFrankLeftArmGIOff.Material   = "Prims platt" & IIF(GILevel=-1," Frank",Abs(GILevel+1))
    pFrankRightArmGIOff.Material  = "Prims platt" & IIF(GILevel=-1," Frank",Abs(GILevel+1))
    pFrankHeadGIOff.Visible     = (pFrankHeadGIOff.Material <> "Prims platt4")
    pFrankBodyGIOff.Visible     = pFrankHeadGIOff.Visible
    pFrankLeftArmGIOff.Visible    = pFrankHeadGIOff.Visible
    pFrankRightArmGIOff.Visible   = pFrankHeadGIOff.Visible
    ' playfield elements
'Dim k, obj
'For Each k In GiElementsGiOffDyn.Keys
'    Set obj = GiElementsGiOffDyn(k)
'    obj.Material = "Prims platt" & IIF(GILevel = -1, IIF(Left(obj.Name, 2) = "sw", " Frank", ""), Abs(GILevel + 1))
'    obj.Visible  = pFrankHeadGIOff.Visible
'Next



For Each k In GiElementsGiOffDyn.Keys
    Set obj = GiElementsGiOffDyn(k)

    ' --- Default material logic (preserved exactly) ---
    newMat = "Prims platt" & IIf( _
                GILevel = -1, _
                IIf(Left(obj.Name, 2) = "sw", " Frank", ""), _
                Abs(GILevel + 1) _
             )

    ' --- Replace for VR exceptions ---
    Select Case obj.Name
        Case "pPlasticRamp_VR_GiOn", "pPlasticRamp_VR_GiOff"
            newMat = "Prims platt ramptrans"

        Case "WireRampVR_GI_ON", "WireRampVR_GI_OFF"
            newMat = "Metal S34 GoldON"
    End Select

    obj.Material = newMat
    obj.Visible  = pFrankHeadGIOff.Visible
Next


    pSiderailsLockbarGIOff.Material = "Prims platt" & IIF(GILevel=-1,"",Abs(GILevel+1))
    pSiderailsLockbarGIOff.Visible  = (pFrankHeadGIOff.Visible And (RailsVisible = 1 Or (RailsVisible = 2 And DesktopMode)))
If GILevel = -1 Then
   ' Dim k, obj
    For Each k In GiElementsDyn.Keys
        Set obj = GiElementsDyn(k)
        obj.Visible = False
    Next

    SetSidewallsVisibility False
    pSiderailsLockbar.Visible = False
End If

    UseGIMaterial GIWireTrigger, "Metal0.8", Abs(GILevel+1)
    UseGIMaterial sGates, "Metal Chrome S34", Abs(GILevel+1)
    UseGIMaterial Array(kIceCave,kVUK,pSpinner54Rod), "Metal Chrome S34", Abs(GILevel+1)
    If GILevel = -4 Then AdjustLights
    GILevel = GILevel + 1
  End If
End Sub

Sub InvalidateFrank()
  If aFlasherValue(0) <> -1 Then IlluminateFrank 0, aFlasherValue(0)
  If aFlasherValue(2) <> -1 Then IlluminateFrank 2, aFlasherValue(2)
End Sub
Sub IlluminateFrank(flasherID, flasherVal)
  If flasherID <> 0 And flasherID <> 2 Then Exit Sub
  If flasherVal <= -99 Then
    If flasherID = 0 Then
      pFrankHeadFOnGIOff.Image    = "FrankHead GI-Off RFlasher-On"
      pFrankBodyFOnGIOff.Image    = "FrankBody GI-Off RFlasher-On"
      pFrankLeftArmFOnGIOff.Image   = "FrankLArm GI-Off RFlasher-On"
      pFrankRightArmFOnGIOff.Image  = "FrankRArm GI-Off RFlasher-On"
    ElseIf flasherID = 2 Then
      pFrankHeadFOnGIOff.Image    = "FrankHead GI-Off LFlasher-On"
      pFrankBodyFOnGIOff.Image    = "FrankBody GI-Off LFlasher-On"
      pFrankLeftArmFOnGIOff.Image   = "FrankLArm GI-Off LFlasher-On"
      pFrankRightArmFOnGIOff.Image  = "FrankRArm GI-Off LFlasher-On"
    End If
    pFrankHeadFOnGIOff.Material   = "Prims platt" & IIF(isGIOn, "4", " Frank")
  ElseIf flasherVal > 0 Then
    pFrankHeadFOnGIOff.Material   = "Prims platt" & IIF(isGIOn, "4", 4-flasherVal)
  Else
    pFrankHeadFOnGIOff.Material   = "Prims platt4"
  End If
  pFrankBodyFOnGIOff.Material   = pFrankHeadFOnGIOff.Material
  pFrankLeftArmFOnGIOff.Material  = pFrankHeadFOnGIOff.Material
  pFrankRightArmFOnGIOff.Material = pFrankHeadFOnGIOff.Material
  pFrankHeadFOnGIOff.Visible    = (pFrankHeadFOnGIOff.Material <> "Prims platt4")
  pFrankBodyFOnGIOff.Visible    = pFrankHeadFOnGIOff.Visible
  pFrankLeftArmFOnGIOff.Visible   = pFrankHeadFOnGIOff.Visible
  pFrankRightArmFOnGIOff.Visible  = pFrankHeadFOnGIOff.Visible
End Sub

Sub AdjustLights()
  AdjustInserts
  AdjustWhiteFlasher
End Sub
Dim areInsertsHigh : areInsertsHigh = False
Const HighInsertsFactor = 2.7
Dim areWhiteFlasherHigh : areWhiteFlasherHigh = False
Const HighWhiteFlasherFactor = 3
Sub AdjustInserts()
  If (isGIOn And Not areInsertsHigh) Or (Not isGIOn And areInsertsHigh) Then Exit Sub
  Dim obj
  For Each obj In InsertLights
    If areInsertsHigh Then
      obj.IntensityScale = obj.IntensityScale / HighInsertsFactor
    Else
      obj.IntensityScale = obj.IntensityScale * HighInsertsFactor
    End If
  Next
  areInsertsHigh = Not areInsertsHigh
End Sub
Sub AdjustWhiteFlasher()
  If (isGIOn And Not areWhiteFlasherHigh) Or (Not isGIOn And areWhiteFlasherHigh) Then Exit Sub
  Dim obj
  For Each obj In WhiteFlasher
    If areWhiteFlasherHigh Then
      obj.Intensity = obj.Intensity / HighWhiteFlasherFactor
    Else
      obj.Intensity = obj.Intensity * HighWhiteFlasherFactor
    End If
  Next
  areWhiteFlasherHigh = Not areWhiteFlasherHigh
End Sub

Sub UseGIMaterial(coll, matName, currentGILevel)
  Dim obj, mat
  mat = CreateMaterialName(matName, currentGILevel)
  For Each obj In coll : obj.Material = mat : Next
End Sub
Function CreateMaterialName(matName, currentGILevel)
  CreateMaterialName = matName & IIF(currentGILevel=0, " Dark", IIF(currentGILevel<4, " Dark" & currentGILevel,""))
End Function

Dim setImage : setImage = False
Sub SetSidewallsVisibility(mode)
  Select Case SelectSidewalls
  Case 0
    If DesktopMode Then pSidewallsDT.Visible = mode Else pSidewallsFS.Visible = mode
  Case 1
    pSidewallsFS.Visible = mode
  Case 2
    pSidewalls.Visible = mode
  Case 3
    pSidewallsDT.Visible = mode
  Case 4
    ' nothing to do
    If Not setImage Then setImage = True : pSidewalls.Image = "SidewPlasTexCombined GI-Off"
    pSidewalls.Visible = mode
  End Select
End Sub


' *********************************************************************
' sound stuff and a bit ball jumping
' *********************************************************************
'Sub sGates_Hit(idx) : PlaySoundAtVolPitch SoundFX("fx_gate",DOFContactors), ActiveBall, 0.02, .25 : End Sub
'Sub sMetalWalls_Hit(idx) : PlaySoundAtBallAbsVol "fx_metalhit" & Int(Rnd*3), Minimum(Vol(ActiveBall),0.5) : End Sub
'Sub sPlastics_Hit(idx) : PlaySoundAtBallAbsVol "fx_ball_hitting_plastic", Minimum(Vol(ActiveBall),0.5) : End Sub
Sub sRubberPosts_Hit(idx) : OnRubberPostHit : End Sub
Sub OnDropTargetHit() : OnTargetHit : End Sub
Sub OnStandupTargetHit() : OnTargetHit : End Sub
Sub OnTargetHit() : ActiveBall.VelZ = ActiveBall.VelZ * (0.5 + (Rnd()*LetTheBallJump + Rnd()*LetTheBallJump + 1) / 6) : End Sub
Sub OnRubberWallHit() : ActiveBall.VelZ = ActiveBall.VelZ * (0.8 + (Rnd()*(LetTheBallJump-1) + 1) / 6) : End Sub
Sub OnRubberPostHit() : ActiveBall.VelZ = ActiveBall.VelZ * (0.9 + (Rnd()*(LetTheBallJump-1) + Rnd()*(LetTheBallJump-1) + 1) / 6) : End Sub

Sub RollOverSound() : PlaySoundAtVolPitch SoundFX("fx_rollover",DOFContactors), ActiveBall, 2, .25 : End Sub
Sub SwitchSound(switch) : PlaySoundAt SoundFX("fx_sensor",DOFContactors), switch : End Sub

Sub trUpperRampStart_Hit() : SlowDownBall ActiveBall : End Sub
'Sub trUpperRampWireStart_Hit() : isBallOnWireRamp = True : End Sub
'Sub trUpperRampWireEnd_Hit() : isBallOnWireRamp = False : End Sub

'Sub trLowerRampBeforeStart_Hit() : isBallOnRamp = False : End Sub
'Sub trLowerRampStart_Hit() : isBallOnRamp = True : SpeedUpRamp ActiveBall : End Sub
Sub trLowerRampStart2_Hit() : SpeedUpRamp ActiveBall : End Sub
'Sub trLowerRampEnd_Hit() : isBallOnRamp = False : End Sub

' ball collision sound
Sub SlowDownBall(actBall)
  With actBall
    If .VelY < 0 Then .VelY = .VelY / 2 : .VelX = .VelX / 2
  End With
End Sub

Sub SpeedUpRamp(actBall)
  With actBall
    If .VelY < 0 Then .VelY = .VelY * 1
  End With
End Sub


' *********************************************************************
' supporting Surround Sound Feedback (SSF) functions
' *********************************************************************
' set position as table object (Use object or light but NOT wall) and Vol to 1
'Sub PlaySoundAt(sound, tableobj)
' PlaySound sound, 1, 1, AudioPan(tableobj), 0, 0, 0, 1, AudioFade(tableobj)
'End Sub

' set position as table object and Vol manually.
'Sub PlaySoundAtVol(sound, tableobj, Vol)
' PlaySound sound, 1, Vol, AudioPan(tableobj), 0, 0, 0, 1, AudioFade(tableobj)
'End Sub
' set position as table object and Vol + RndPitch manually
Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
End Sub

' set all as per ball position & speed.
'Sub PlaySoundAtBall(sound)
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
'End Sub
' set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3
'Sub PlaySoundAtBallVol(sound, VolMult)
' PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
'End Sub
Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

' requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound
Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0, 0, 1, 1, AudioFade(tableobj)
End Sub




Function BallVelZ(ball) 'Calculates the ball speed in the -Z
    BallVelZ = INT((ball.VelZ) * -1 )
End Function


' *************************************************************
' ball shadow and ramp look
' *************************************************************
Const tnob = 6 ' total number of balls

Dim BallShadow : BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5, BallShadow6)

Sub GraphicsTimer_Timer()
  Dim ii

  ' maybe show ball shadows
  If ShowBallShadow <> 0 Then
    Dim BOT
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT) < tnob - 1 Then
      For ii = UBound(BOT) + 1 To tnob - 1
        If BallShadow(ii).Visible Then BallShadow(ii).Visible = False
      Next
    End If
    ' render the shadow for each ball
    For ii = 0 to UBound(BOT)
      If BOT(ii).X < Frankenstein.Width/2 Then
        BallShadow(ii).X = ((BOT(ii).X) - (Ballsize/6) + ((BOT(ii).X - (Frankenstein.Width/2))/7)) + 6
      Else
        BallShadow(ii).X = ((BOT(ii).X) + (Ballsize/6) + ((BOT(ii).X - (Frankenstein.Width/2))/7)) - 6
      End If
      BallShadow(ii).Y = BOT(ii).Y + 12
      BallShadow(ii).Visible = (BOT(ii).Z > 20)
    Next
  End If

  ' maybe move primitive flippers
  pLeftFlipper.ObjRotZ  = LeftFlipper.CurrentAngle +180
  pLeftFlipperGIOff.ObjRotZ  = LeftFlipper.CurrentAngle +180
  pRightFlipper.ObjRotZ = RightFlipper.CurrentAngle +180
  pRightFlipperGIOff.ObjRotZ = RightFlipper.CurrentAngle +180
  pRightUpperFlipper.ObjRotZ = RightUpperFlipper.CurrentAngle +180
  pRightUpperFlipperGIOff.ObjRotZ = RightUpperFlipper.CurrentAngle +180
  pLeftFlipperShadow.ObjRotZ  = LeftFlipper.CurrentAngle + 1
  pRightFlipperShadow.ObjRotZ = RightFlipper.CurrentAngle + 1
  pRightUpperFlipperShadow.ObjRotZ = RightUpperFlipper.CurrentAngle + 1
End Sub


' *************************************************************
'      rolling and realtime sounds
' *************************************************************
'ReDim rolling(tnob)
'For i = 0 to tnob : rolling(i) = False : Next
'Dim isBallOnRamp : isBallOnRamp = False
'Dim isBallOnWireRamp : isBallOnWireRamp = False

Sub Switches_Timer() 'was previously rollingtimer, but that has been replaced. kept the logic that checks ROM switched and renamed timer
    Dim BOT, b, ub
    BOT = GetBalls

    ' Guard: ensure we have an array of balls and a valid upper bound
    If Not IsArray(BOT) Then Exit Sub
    On Error Resume Next
    ub = UBound(BOT)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    If ub < 0 Then Exit Sub

    For b = 0 To ub
        ' sw53 (VUK)
        If InCircle(BOT(b).X, BOT(b).Y, sw53.X, sw53.Y, 20) And BOT(b).Z < 0 Then
            If Abs(BOT(b).VelX) < 3 And Abs(BOT(b).VelY) < 3 Then
                BOT(b).VelX = BOT(b).VelX / 3
                BOT(b).VelY = BOT(b).VelY / 3
                If Not Controller.Switch(53) Then
                    Controller.Switch(53) = True
                    Set sw53Ball = BOT(b)
                    FrankIsWatching sw53Ball
                End If
            End If
        End If

        ' sw55 (Ice Cave)
        If InCircle(BOT(b).X, BOT(b).Y, sw55.X, sw55.Y, 20) And BOT(b).Z < 0 Then
            If Abs(BOT(b).VelX) < 3 And Abs(BOT(b).VelY) < 3 Then
                BOT(b).VelX = BOT(b).VelX / 3
                BOT(b).VelY = BOT(b).VelY / 3
                If Not Controller.Switch(55) Then
                    Controller.Switch(55) = True
                    Set sw55Ball = BOT(b)
                    FrankIsWatching sw55Ball
                End If
            End If
        End If

        ' sw31 (Top scoop)
        If InCircle(BOT(b).X, BOT(b).Y, sw31.X, sw31.Y, 20) And BOT(b).Z < 0 Then
            If Abs(BOT(b).VelX) < 3 And Abs(BOT(b).VelY) < 3 Then
                BOT(b).VelX = BOT(b).VelX / 3
                BOT(b).VelY = BOT(b).VelY / 3
                If Not Controller.Switch(31) Then
                    Controller.Switch(31) = True
                    Set sw31Ball = BOT(b)
                    FrankIsWatching sw31Ball
                End If
            End If
        End If

        ' sw32 (Lower scoop)
        If InCircle(BOT(b).X, BOT(b).Y, sw32.X, sw32.Y, 20) And BOT(b).Z < 0 Then
            If Abs(BOT(b).VelX) < 3 And Abs(BOT(b).VelY) < 3 Then
                BOT(b).VelX = BOT(b).VelX / 3
                BOT(b).VelY = BOT(b).VelY / 3
                If Not Controller.Switch(32) Then
                    Controller.Switch(32) = True
                    Set sw32Ball = BOT(b)
                    FrankIsWatching sw32Ball
                End If
            End If
        End If
    Next
End Sub




' *********************************************************************
' some more general methods
' *********************************************************************


Function InCircle(pX, pY, centerX, centerY, radius)
  Dim route
  route = Sqr((pX - centerX) ^ 2 + (pY - centerY) ^ 2)
  InCircle = (route < radius)
End Function

Function Minimum(val1, val2)
  Minimum = IIF(val1<val2,val2,val1)
End Function

Const fileName = "C:\Games\Visual Pinball\User\Log.txt"
Sub WriteLine(text)
  Dim fso, fil
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set fil = fso.OpenTextFile(fileName, 8, true)
  fil.WriteLine Date & " " & Time & ": " & text
  fil.Close
  Set fso = Nothing
End Sub


Function Maximum(obj1, obj2)
  If obj1 > obj2 Then
    Maximum = obj1
  Else
    Maximum = obj2
  End If
End Function

Function GetBallID(actBall)
  Dim b, BOT, ret
  ret = -1
  BOT = GetBalls
  For b = 0 to UBound(BOT)
    If actBall Is BOT(b) Then
      ret = b : Exit For
    End If
  Next
  GetBallID = ret
End Function

'**************************************************
'        Flipper Collision Subs
'NOTE: COpy and overwrite collision sound from original collision subs over
'RandomSoundFlipper()' below
'**************************************************'

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub



'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

' Put all the Post and Pin objects in dPosts collection. Make sure dPosts fires hit events.
Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

' This collection contains the bottom sling posts. They are not in the dPosts collection so that the TargetBouncer is not applied to them, but they should still have dampening applied
' If you experience airballs with posts or targets, consider adding them to this collection
Sub NoTargetBouncer_Hit
    RubbersD.dampen ActiveBall
End Sub

' Put all the Sleeve objects in dSleeves collection. Make sure dSleeves fires hit events.
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
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
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
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7-1.0

sub TargetBouncer(aBall,defvalue)
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
  end if
end sub

'******************************************************
'                TRACK ALL BALL VELOCITIES
'                 FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
        public ballvel, ballvelx, ballvely

        Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

        Public Sub Update()        'tracks in-ball-velocity
                dim str, b, AllBalls, highestID : allBalls = getballs

                for each b in allballs
                        if b.id >= HighestID then highestID = b.id
                Next

                if uBound(ballvel) < highestID then redim ballvel(highestID)        'set bounds
                if uBound(ballvelx) < highestID then redim ballvelx(highestID)        'set bounds
                if uBound(ballvely) < highestID then redim ballvely(highestID)        'set bounds

                for each b in allballs
                        ballvel(b.id) = BallSpeed(b)
                        ballvelx(b.id) = b.velx
                        ballvely(b.id) = b.vely
                Next
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
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, AutoPlunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall2()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, kickback
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

Sub RandomSoundBumperLeft(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Left_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
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

Sub Posts_Hit(idx)
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
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
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

'/////////////////////////////  GI AND FLASHER   ////////////////////////////

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
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************

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
'****************************************************************
' VR Stuff
'****************************************************************
'///////////////////////---- VR Room ----////////////////////////
Dim VRRoomChoice : VRRoomChoice = 2         '1 - Minimal Room, 2 - Hallway Room
Dim VRTest : VRTest = False

'****************************************************************
'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim Romvolume
Dim RampSetChoice          ' 0 = original ramps, 1 = alternate ramps

Dim GiElementsDyn
Dim GiElementsGiOffDyn

Set GiElementsDyn = CreateObject("Scripting.Dictionary")
Set GiElementsGiOffDyn = CreateObject("Scripting.Dictionary")


' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False

Sub Frankenstein_OptionEvent(ByVal eventId)

    ' one-time: disable static pre-rendering once user touches options
    If eventId = 1 And Not dspTriggered Then
        dspTriggered = True
        DisableStaticPreRendering = True
    End If

    ' ==========================
    '  F12 USER OPTIONS
    ' ==========================

    'ROMVolume = Frankenstein.Option("ROM Volume (-32 off; 0 max) - Requires Restart; altsound set to 0", -32, 0, 1, 0, 0)
If RenderingMode = 2 or VRTest = True Then
    RampSetChoice = Frankenstein.Option( _
        "Ramps Visual Style", _
        0, 1, 1, 1, 0, _
        Array("Standard", "Optimized for VR") _
    )
Else
    RampSetChoice = Frankenstein.Option( _
        "Ramps Visual Style", _
        0, 1, 1, 0, 0, _
        Array("Standard", "Optimized for VR") _
    )
End If

UpdateGiCollections

    ' GI on/off
EnableGI = Frankenstein.Option("General Illumination", 0, 1, 1, 1, 0, _
    Array("GI Off", "GI On"))


    ' Flashers on/off
    EnableFlasher = Frankenstein.Option("Flashers", 0, 1, 1, 1, 0, _
        Array("Flashers Off", "Flashers On"))

    ' Ball reflections (0 = no reflections, 1 = full)
    EnableReflectionsAtBall = Frankenstein.Option("Ball Reflections", 0, 1, 1, 1, 0, _
        Array("Reflections Off", "Reflections On"))

    ' Right MagnaSave LUT change
    EnableLUTChangeWithRightMagnaSave = Frankenstein.Option("Right MagnaSave Changes LUT", 0, 1, 1, 1, 0, _
        Array("Disabled", "Enabled"))

    ' Sidewalls: uses your SetSidewallsVisibility cases 0–4
    SelectSidewalls = Frankenstein.Option("Sidewall Mode - Requires Restart", 0, 4, 1, 0, 0, _
        Array("Auto", "for Fullscreen View", "for a Medium View", "for Desktop View", "For when Gi is off"))

    ' Ball shadow (your shadow logic will read this)
    ShowBallShadow = Frankenstein.Option("Ball Shadow", 0, 1, 1, 1, 0, _
        Array("Shadow Off", "Shadow On"))

  ' Ball jump intensity
  LetTheBallJump = Frankenstein.Option("Ball Jump", 0, 6, 1, 0, 0, _
    Array("Off", "Minimal", "Low", "Medium", "High", "Very High", "Maximum"))

  ' Frank head animation mode
  ' 0 = Off, 1 = Random, 2 = Follow Ball, 3 = Watch Table, 4 = Video Style
  AnimateFranksHead = Frankenstein.Option("Frank Head Mode", 0, 4, 1, 4, 0, _
    Array("Off", "Random", "Follow Ball", "Watch Table", "Video Style"))


    ' Rails visibility (used in your GI block and SetSidewallsVisibility)
    ' 0 = Hide, 1 = Show, 2 = Desktop Only
    RailsVisible = Frankenstein.Option("Side Rails Visibility", 0, 2, 1, 2, 0, _
        Array("Hide Rails", "Show Rails", "Desktop Only"))


If RenderingMode = 2 or VRTest = True Then
    ' VRRoom
  VRRoomChoice = Frankenstein.Option("VR Room", 1, 3, 1, 2, 0, Array("Minimal", "Hallway", "Cab Only"))
  LoadVRRoom
End If


    ' ==========================
    '  APPLY INSTANTLY
    ' ==========================
    If eventId = 1 Then

        ' ---- GI: force one transition to match the new option ----
        If EnableGI = 0 Then
            ' User turned GI Off: force GI off once
            If isGIOn Then SetGI False
        Else
            ' User turned GI On: force GI on once if currently off
            If Not isGIOn Then SetGI True
        End If

        ' Other instant apply stuff
        ApplyReflectionsOption
        ApplySideVisibilityOptions
        ApplyFlasherOption
    End If
End Sub

Sub UpdateGiCollections()
    Dim o

    ' Recreate dynamic dictionaries from scratch
    Set GiElementsDyn      = CreateObject("Scripting.Dictionary")
    Set GiElementsGiOffDyn = CreateObject("Scripting.Dictionary")

    ' Copy everything from static GI-ON collection except the parts we swap/manage here
    For Each o In GiElements
        If Not ( _
            (o Is pPlasticRamp) _
            Or (o Is pPlasticRamp_VR_GiOn) _
            Or (o Is WireRampDT_GI_ON) _
            Or (o Is WireRampVR_GI_ON) _
            Or (o Is pPlasticTransM) _
            Or (o Is pPlasticTransR) _
            Or (o Is pPlasticTransM_VR_GiOn) _
            Or (o Is pPlasticTransR_VR_GiOn) _
            Or (o Is pPlasticTransB) _
            Or (o Is pPlasticTransB_VR_GiOn) _
        ) Then
            GiElementsDyn.Add o.Name, o
        End If
    Next

    ' Copy everything from static GI-OFF collection except the parts we swap/manage here
    For Each o In GiElementsGiOff
        If Not ( _
            (o Is pPlasticRampGiOff) _
            Or (o Is pPlasticRamp_VR_GiOff) _
            Or (o Is WireRampDT_GI_OFF) _
            Or (o Is WireRampVR_GI_OFF) _
            Or (o Is pPlasticTransMGIOff) _
            Or (o Is pPlasticTransRGIOff) _
            Or (o Is pPlasticTransM_VR_GiOff) _
            Or (o Is pPlasticTransR_VR_GiOff) _
            Or (o Is pPlasticTransBGIOff) _
            Or (o Is pPlasticTransB_VR_GiOff) _
        ) Then
            GiElementsGiOffDyn.Add o.Name, o
        End If
    Next

    ' Baseline: hide all managed swap parts
    pPlasticRamp.Visible = False
    pPlasticRampGiOff.Visible = False
    pPlasticRamp_VR_GiOn.Visible = False
    pPlasticRamp_VR_GiOff.Visible = False

    WireRampDT_GI_ON.Visible = False
    WireRampDT_GI_OFF.Visible = False
    WireRampVR_GI_ON.Visible = False
    WireRampVR_GI_OFF.Visible = False

    pPlasticTransM.Visible = False
    pPlasticTransR.Visible = False
    pPlasticTransMGIOff.Visible = False
    pPlasticTransRGIOff.Visible = False

    pPlasticTransM_VR_GiOn.Visible = False
    pPlasticTransR_VR_GiOn.Visible = False
    pPlasticTransM_VR_GiOff.Visible = False
    pPlasticTransR_VR_GiOff.Visible = False

    pPlasticTransB.Visible = False
    pPlasticTransBGIOff.Visible = False
    pPlasticTransB_VR_GiOn.Visible = False
    pPlasticTransB_VR_GiOff.Visible = False

    ' Add the correct parts to the correct dynamic dictionaries and show the correct set
    If RampSetChoice = 0 Then
        ' Original set (DT / non-VR)

        GiElementsDyn.Add pPlasticRamp.Name, pPlasticRamp
        GiElementsDyn.Add WireRampDT_GI_ON.Name, WireRampDT_GI_ON
        GiElementsDyn.Add pPlasticTransM.Name, pPlasticTransM
        GiElementsDyn.Add pPlasticTransR.Name, pPlasticTransR
        GiElementsDyn.Add pPlasticTransB.Name, pPlasticTransB

        GiElementsGiOffDyn.Add pPlasticRampGiOff.Name, pPlasticRampGiOff
        GiElementsGiOffDyn.Add WireRampDT_GI_OFF.Name, WireRampDT_GI_OFF
        GiElementsGiOffDyn.Add pPlasticTransMGIOff.Name, pPlasticTransMGIOff
        GiElementsGiOffDyn.Add pPlasticTransRGIOff.Name, pPlasticTransRGIOff
        GiElementsGiOffDyn.Add pPlasticTransBGIOff.Name, pPlasticTransBGIOff

        If isGIOn <> False Then
            pPlasticRamp.Visible = True
            WireRampDT_GI_ON.Visible = True
            pPlasticTransM.Visible = True
            pPlasticTransR.Visible = True
            pPlasticTransB.Visible = True

            pPlasticRampGiOff.Visible = False
            WireRampDT_GI_OFF.Visible = False
            pPlasticTransMGIOff.Visible = False
            pPlasticTransRGIOff.Visible = False
            pPlasticTransBGIOff.Visible = False
        Else
            pPlasticRamp.Visible = False
            WireRampDT_GI_ON.Visible = False
            pPlasticTransM.Visible = False
            pPlasticTransR.Visible = False
            pPlasticTransB.Visible = False

            pPlasticRampGiOff.Visible = True
            WireRampDT_GI_OFF.Visible = True
            pPlasticTransMGIOff.Visible = True
            pPlasticTransRGIOff.Visible = True
            pPlasticTransBGIOff.Visible = True
        End If

    Else
        ' VR set

        GiElementsDyn.Add pPlasticRamp_VR_GiOn.Name, pPlasticRamp_VR_GiOn
        GiElementsDyn.Add WireRampVR_GI_ON.Name, WireRampVR_GI_ON
        GiElementsDyn.Add pPlasticTransM_VR_GiOn.Name, pPlasticTransM_VR_GiOn
        GiElementsDyn.Add pPlasticTransR_VR_GiOn.Name, pPlasticTransR_VR_GiOn
        GiElementsDyn.Add pPlasticTransB_VR_GiOn.Name, pPlasticTransB_VR_GiOn

        GiElementsGiOffDyn.Add pPlasticRamp_VR_GiOff.Name, pPlasticRamp_VR_GiOff
        GiElementsGiOffDyn.Add WireRampVR_GI_OFF.Name, WireRampVR_GI_OFF
        GiElementsGiOffDyn.Add pPlasticTransM_VR_GiOff.Name, pPlasticTransM_VR_GiOff
        GiElementsGiOffDyn.Add pPlasticTransR_VR_GiOff.Name, pPlasticTransR_VR_GiOff
        GiElementsGiOffDyn.Add pPlasticTransB_VR_GiOff.Name, pPlasticTransB_VR_GiOff

        If isGIOn <> False Then
            pPlasticRamp_VR_GiOn.Visible = True
            WireRampVR_GI_ON.Visible = True
            pPlasticTransM_VR_GiOn.Visible = True
            pPlasticTransR_VR_GiOn.Visible = True
            pPlasticTransB_VR_GiOn.Visible = True

            pPlasticRamp_VR_GiOff.Visible = False
            WireRampVR_GI_OFF.Visible = False
            pPlasticTransM_VR_GiOff.Visible = False
            pPlasticTransR_VR_GiOff.Visible = False
            pPlasticTransB_VR_GiOff.Visible = False
        Else
            pPlasticRamp_VR_GiOn.Visible = False
            WireRampVR_GI_ON.Visible = False
            pPlasticTransM_VR_GiOn.Visible = False
            pPlasticTransR_VR_GiOn.Visible = False
            pPlasticTransB_VR_GiOn.Visible = False

            pPlasticRamp_VR_GiOff.Visible = True
            WireRampVR_GI_OFF.Visible = True
            pPlasticTransM_VR_GiOff.Visible = True
            pPlasticTransR_VR_GiOff.Visible = True
            pPlasticTransB_VR_GiOff.Visible = True
        End If
    End If
End Sub





Sub ApplyFlasherOption()
    Dim i
    If EnableFlasher = 0 Then
        For i = 0 To UBound(aFlasherValue)
            aFlasherValue(i) = 0  ' will be treated as "off" and fade out
        Next
    End If
End Sub


Sub ApplyReflectionsOption()
    Dim i, j, obj
    For i = 0 To UBound(PFLightsCount)
        For j = 0 To PFLightsCount(i) - 1
            Set obj = PFLights(i, j)
            If Not obj Is Nothing Then
                If Right(obj.Name, 1) = "a" Then
                    obj.IntensityScale = EnableReflectionsAtBall
                End If
            End If
        Next
    Next
End Sub


Sub SolFlasher(id, enabled)
    ' If flashers are disabled via F12, force them toward "off" and ignore ROM calls
    If EnableFlasher = 0 Then
        aFlasherValue(id) = 3   ' will fade to off in LightsTimer
        Exit Sub
    End If

    If enabled Then
        aFlasherValue(id) = -99 ' flasher on
    Else
        aFlasherValue(id) = 3   ' flasher off / fade
    End If
End Sub


Sub ApplySideVisibilityOptions()
    Dim railsOn

    ' Re-evaluate sidewalls based on current GI state and new SelectSidewalls
    If GILevel > 0 Then
        SetSidewallsVisibility True
    Else
        SetSidewallsVisibility False
    End If

    ' Rails: mirror what your GI code does, but on demand
    railsOn = (RailsVisible = 1 Or (RailsVisible = 2 And DesktopMode))

    pSiderailsLockbar.Visible    = (GILevel > 0 And railsOn)
    pSiderailsLockbarGIOff.Visible = (GILevel <= 0 And railsOn)
End Sub


'******************************************************
' ZBRL: BALL ROLLING AND DROP SOUNDS
'******************************************************
'
' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

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
  Dim b', BOT
  gBOT = GetBalls


  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ < -1 and gBOT(b).z < 55 and gBOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz > -7 Then
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


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************

'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS: Set LS = New SlingshotCorrection
Dim RS: Set RS = New SlingshotCorrection
'Dim TS: Set TS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'TS.Object = TopSlingshot
  'TS.EndPoint1 = EndPoint1TS
  'TS.EndPoint2 = EndPoint2TS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00, - 4
  AddSlingsPt 1, 0.45, - 7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4
End Sub

Sub AddSlingsPt(idx, aX, aY)    'debugger wrapper for adjusting flipper script In-game
  Dim a
  a = Array(LS, RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub

' The following sub are needed, however they exist in the ZMAT maths section of the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function
'
'Function RotPoint(x,y,angle)
' dim rx, ry
' rx = x*dCos(angle) - y*dSin(angle)
' ry = x*dSin(angle) + y*dCos(angle)
' RotPoint = Array(rx,ry)
'End Function

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

Dim gBot
Sub GameTimer_Timer()
  'gBOT = GetBalls
  RollingUpdate
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub


'--- Add this near the top of your script ---
Function IIf(condition, truePart, falsePart)
    If condition Then
        IIf = truePart
    Else
        IIf = falsePart
    End If
End Function

'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'          * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'          * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'          * Create a Timer called RampRoll, that is enabled, with a interval of 100
'          * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'          * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'          * To stop tracking ball
'                 * call WireRampOff
'                 * Otherwise, the ball will auto remove if it's below 30 vp units
'

dim RampMinLoops : RampMinLoops = 4

' RampBalls
'      Setup:        Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RammBalls(6,2)
'      Description:
dim RampBalls(5,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(5)

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID : End Sub


' WaddBall (Active Ball, Boolean)
'     Description: This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
Sub Waddball(input, RampInput)  'Add ball
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  dim x : for x = 1 to uBound(RampBalls)  'Check, don't add balls twice
    if RampBalls(x, 1) = input.id then
      if Not IsEmpty(RampBalls(x,1) ) then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 to uBound(RampBalls)
    if IsEmpty(RampBalls(x, 1)) then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      exit Sub
    End If
    if x = uBound(RampBalls) then   'debug
      Debug.print "WireRampOn error, ball queue is full: " & vbnewline & _
      RampBalls(0, 0) & vbnewline & _
      Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbnewline & _
      Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbnewline & _
      Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbnewline & _
      Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbnewline & _
      Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbnewline & _
      " "
    End If
  next
End Sub

' WRemoveBall (BallId)
'    Description: This subroutine is called from the RampRollUpdate subroutine
'                 and is used to remove and stop the ball rolling sounds
Sub WRemoveBall(ID)   'Remove ball
  'Debug.Print "In WRemoveBall() + Remove ball from loop array"
  dim ballcount : ballcount = 0
  dim x : for x = 1 to Ubound(RampBalls)
    if ID = RampBalls(x, 1) then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    if not IsEmpty(Rampballs(x,1)) then ballcount = ballcount + 1
  next
  if BallCount = 0 then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer():RampRollUpdate:End Sub

Sub RampRollUpdate()    'Timer update
  dim x : for x = 1 to uBound(RampBalls)
    if Not IsEmpty(RampBalls(x,1) ) then
      if BallVel(RampBalls(x,0) ) > 1 then ' if ball is moving, play rolling sound
        If RampType(x) then
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end if
  next
  if not RampBalls(0,0) then RampRoll.enabled = 0

End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbnewline & _
  "1 " & Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbnewline & _
  "2 " & Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbnewline & _
  "3 " & Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbnewline & _
  "4 " & Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbnewline & _
  "5 " & Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbnewline & _
  "6 " & Typename(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbnewline & _
  " "
End Sub


Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(ball), 1, -1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function


'Ramp triggers


Sub trLowerRampStart_Hit()
    If ActiveBall.VelY < 0 Then   ' going up the table
        WireRampOn True           ' plastic
    Else
        WireRampOff               ' coming from above, treat as exit
    End If
End Sub

Sub trUpperRampWireEnd_Hit
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
End Sub

Sub trLowerrampend_Hit 'end of ramp
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
End Sub

Sub RHP001_Hit() 'start of plastic ramp
           WireRampOn True           ' plastic
End Sub

Sub trUpperRampWireStart_Hit
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
        WireRampOn False          ' wire
        'WireRampOn False          ' wire
End Sub

Sub RHP002_Hit
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
End Sub

Sub RHP00002_Hit
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
End Sub

Sub RampTrigger1_Hit
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTrigger2_Hit
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTrigger3_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 50
    activeball.AngMomZ = -abs(activeball.AngMomZ) * 3
    WireRampOff
End Sub

Sub RampTrigger4_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 50
    activeball.AngMomZ = abs(activeball.AngMomZ) * 3
    WireRampOff
End Sub

Sub RampTrigger5_Hit
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTrigger6_Hit
  WireRampOff
End Sub

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub




'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************

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
'Dim ULF : Set ULF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  Dim x, a
  a = Array(LF, RF)
  For Each x In a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 5.5
    x.AddPt "Polarity", 2, 0.16, - 5.5
    x.AddPt "Polarity", 3, 0.20, - 0.75
    x.AddPt "Polarity", 4, 0.25, - 1.25
    x.AddPt "Polarity", 5, 0.3, - 1.75
    x.AddPt "Polarity", 6, 0.4, - 3.5
    x.AddPt "Polarity", 7, 0.5, - 5.25
    x.AddPt "Polarity", 8, 0.7, - 4.0
    x.AddPt "Polarity", 9, 0.75, - 3.5
    x.AddPt "Polarity", 10, 0.8, - 3.0
    x.AddPt "Polarity", 11, 0.85, - 2.5
    x.AddPt "Polarity", 12, 0.9, - 2.0
    x.AddPt "Polarity", 13, 0.95, - 1.5
    x.AddPt "Polarity", 14, 1, - 1.0
    x.AddPt "Polarity", 15, 1.05, -0.5
    x.AddPt "Polarity", 16, 1.1, 0
    x.AddPt "Polarity", 17, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.23, 0.85
    x.AddPt "Velocity", 2, 0.27, 1
    x.AddPt "Velocity", 3, 0.3, 1
    x.AddPt "Velocity", 4, 0.35, 1
    x.AddPt "Velocity", 5, 0.6, 1 '0.982
    x.AddPt "Velocity", 6, 0.62, 1.0
    x.AddPt "Velocity", 7, 0.702, 0.968
    x.AddPt "Velocity", 8, 0.95,  0.968
    x.AddPt "Velocity", 9, 1.03,  0.945
    x.AddPt "Velocity", 10, 1.5,  0.945

  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
  LF.SetObjects "LF", LeftFlipper, TriggerLF
  RF.SetObjects "RF", RightFlipper, TriggerRF
    'ULF.SetObjects "ULF", LeftFlipper1, TriggerLF1
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
     Dim BOT
     BOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          gBOT(b).velx = BOT(b).velx / 1.3
          gBOT(b).vely = BOT(b).vely - 0.5
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

Dim LFPress, RFPress, ULFPress, LFCount, RFCount
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
'Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
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
'   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
'   Const EOSReturn = 0.035  'mid 80's to early 90's
    Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
    Dim BOT, b

    FlipperPress = 0
    Flipper.eostorqueangle = EOSA
    Flipper.eostorque = EOST * EOSReturn / FReturn

    If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
        BOT = GetBalls   ' fresh list of valid balls on the table

        If IsArray(BOT) Then
            For b = 0 To UBound(BOT)
                If IsObject(BOT(b)) Then
                    ' check for cradle near this flipper
                    If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then
                        ' clamp downward speed a bit
                        If BOT(b).VelY >= -0.4 Then BOT(b).VelY = -0.4
                    End If
                End If
            Next
        End If
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

'*************
'VR Stuff
'*************

DIM VRThings, VRRoom

Sub LoadVRRoom
  for each VRThings in VRCab:VRThings.visible = 1:Next
  for each VRThings in VRMin:VRThings.visible = 0:Next
  for each VRThings in VRMega:VRThings.visible = 0:Next
  If RenderingMode = 2 or VRTest Then
    VRRoom = VRRoomChoice
    RailsVisible = 0
    Primary_CabBack.visible = 1
    Primary_CabBack.sidevisible = 1
  Else
    VRRoom = 0
    Primary_CabBack.visible = 1
    Primary_CabBack.sidevisible = 1
  End If
  If VRRoom = 1 Then
    for each VRThings in VRCab:VRThings.visible = 1:Next
    for each VRThings in VRMin:VRThings.visible = 1:Next
    for each VRThings in VRMega:VRThings.visible = 0:Next
  End If
  If VRRoom = 2 Then
    for each VRThings in VRCab:VRThings.visible = 1:Next
    for each VRThings in VRMin:VRThings.visible = 0:Next
    for each VRThings in VRMega:VRThings.visible = 1:Next
  End If
  If VRRoom = 3 Then
    for each VRThings in VRCab:VRThings.visible = 1:Next
    for each VRThings in VRMin:VRThings.visible = 0:Next
    for each VRThings in VRMega:VRThings.visible = 0:Next
  End If
End Sub

Sub L64_animate
  If l64.state > .05 Then
    Primary_StartButton2.disablelighting = 1
  Else
    Primary_StartButton2.disablelighting = 0
  End if
End Sub

Sub L63_animate
  If l63.state > .05 Then
    Primary_BuyInButton2.disablelighting = 1
  Else
    Primary_BuyInButton2.disablelighting = 0
  End if
End Sub


Const GATE_MIN_VELY = 2   ' tweak: 1–5
Const GATE_MIN_SPEED = 6  ' tweak: 4–10

Sub GateBackStop_Hit()
    Dim vy, spd
    vy  = ActiveBall.VelY
    spd = Sqr(ActiveBall.VelX*ActiveBall.VelX + ActiveBall.VelY*ActiveBall.VelY)

    ' Only if the ball is moving DOWN the table (toward flippers)
    If vy > GATE_MIN_VELY And spd > GATE_MIN_SPEED Then
        PlaySoundAtVol "Gate_FastTrigger_1", gShooterLane, 0.6
    End If
End Sub
