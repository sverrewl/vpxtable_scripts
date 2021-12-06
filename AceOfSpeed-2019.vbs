Option Explicit
Randomize
'
'__   ______ AAAAAA   CCCCCC  EEEEEEEE                    SSSSS    PPPPPP  EEEEEEE EEEEEEE DDDDDD
'   _    __AAAAAAAA CCCCCCCC EEEEEEEE                   SSS SSS   PPPPPPP EEEEEEE EEEEEEE DDDDDDD           #
'         AA    AA CC       EE                  fff    SS        PP   PP EE      EE      DD   DD          ###
'___  __ AA    AA CC       EEEE               ff       SSS      PPPPPPP EEEE    EEEE    DD   DD         ######
'____   AAAAAAAA CC       EEEE        oooo   ff         SS     PPPPPP  EEEE    EEEE    DD   DD        #########
'      AA    AA CC       EE         oo  oo  ffff        SSS   PP      EE      EE      DD   DD        ##########
'  ___AA    AA CC       EE         oo  oo  ff     SS    SS   PP      EE      EE      DD   DD          #######
'_ __AA    AA CCCCCCCC EEEEEEEE   oo  oo  ff     SSS  SSS   PP      EEEEEEE EEEEEEE DDDDDDD             ##
'  _AA    AA  CCCCCC  EEEEEEEE    oooo   ff      SSSSSS    PP      EEEEEEE EEEEEEE DDDDDD             #####
'
' Version 1.0
'
' Mod of MOUSIN' AROUN (Bally/Midway, 1989) by Mussinger  _2019_
'
'
' Mousin' Around (Version 1.1) Great development by Herweh 2019 for Visual Pinball 10
'
' "If you don't have this great table it's a need to !!! Go for it ;)"
'
' First of all : Thanks to Herweh to let me modify this table. He made a awesome work. Also Thanks to everybody who worked on this table with him (and before under VP9)
'
' Ace of Speed WIP special thanks to :
' - This is THE "thank you" ... Without him this table wouldn't look like this ... BIG THANKS to David Vicente, the illustrator/artist of most of the Artworks elements he let me modify and adapt for this project.
' - David Maldonado for his daily support, advises , brain storming and for the wonderfull 3D stuff he's preparing for a future release. My friend your are the number one ! ;)
' - JPJ for his solutions to some issues and for  the 3D Charger RT model and integration help.
' - A big big thank to Scott Wickberg who helped a lot with codes and the 3D pistons! Scott, you are a Magician ;)
' - Thx to David Paiva for his patiency an help with the PupPack
' - Thanks to virtuals pinball facebook groups , specialy: the frenchy one "Monte Ton Cab" , Scott's "Orbital creators club" and "Pinup Popper Artists Association"
' - Thanks to Rik Laubach for his beta-testing time ;)
'... Sorry if I forget somebody  :)  Gonna have to thank my wife Vero also, for the tones of hours she let me pass in front of the computer!
'
' FOR YOUR INFO:
' If you encounter performance issues, try to set some of the objects to "Static rendering". Open layer 7, select all screws and/or locknuts and select "Static Rendering" in the options windows.
'
'

Dim GIColorMod, FlipperColorMod, LetTheBallJump, ActivateMouseHoleMultiball, EnableGI, EnableFlasher, ShowBallShadow, ShadowOpacityGIOff, ShadowOpacityGIOn, RailsVisible, DrainSoundFx


' ****************************************************
' OPTIONS
' ****************************************************

' Drain Sound
'   0 = Tires squealing (default)
'   1 = Sad Trombone
'   2 = None
DrainSoundFx = 0

' GI COLOR MOD (Change them with Left Magna-Save)
'   0 = White bulbs and GI (default)
'   1 = Yellow bulbs and GI
'   2 = Red bulbs and GI
'   3 = Blue bulbs and GI
GIColorMod = 0

' FLIPPER COLOR MOD (Change them with Right Magna-Save)
' 0 = Purpple neon / Williams
' 1 = White/red/Bally style (default)
'   2 = Yellow/white/Williams
'   3 = Red/white/Williams
' 4 = Blue/white/Williams
FlipperColorMod = 1

' LET THE BALL JUMP A BIT
' 0 = off
' 0 to 6 = ball jump intensity
LetTheBallJump = 3

' ACTIVATE HERWEH'S ADDITIONAL MOUSE HOLE MULTIBALL (May create PupPack bugs)
' 0 = Not active
' 1 = up to 5-ball mouse hole multiball with changing traps every 20 seconds
' 2 = 5-ball mouse hole multiball
ActivateMouseHoleMultiball = 0

' ENABLE/DISABLE GI (general illumination)
' 0 = GI is off
' 1 = GI is on (value is a multiplicator for GI intensity - decimal values like 0.7 or 1.33 are valid too)
EnableGI = 1

' ENABLE/DISABLE flasher
' 0 = Flashers are off
' 1 = Flashers are on
EnableFlasher = 1

' SHOW BALL SHADOWS
' 0 = no ball shadows
' 1 = ball shadows are visible
ShowBallShadow = 1

' PLAYFIELD SHADOW INTENSITY DURING GI OFF OR ON (adds additional visual depth)
' usable range is 0 (lighter) - 100 (darker)
ShadowOpacityGIOff = 90
ShadowOpacityGIOn  = 60

' SIDE RAILS VISIBILITY
'   0 = hide side rails
'   1 = show side rails
RailsVisible = 1

' Notes:
' If you are encounter performance issues and low frame rate. I would reccomend disabling "Reflect Elements On Playfield".
' This will increase performance for lesser powerful systems.

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
Const SCoin     = "fx_coin"


'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************

Const cDMDRotation  = -1      '-1 for No change, 0 - DMD rotation of 0?, 1 - DMD rotation of 90?
'Const cGameName  = "acesofspeed" 'ROM name
Const cGameName   = "mousn_l1"  'ROM name
Const ballsize    = 50
Const ballmass    = 1


If Version < 10500 Then
  MsgBox "This table requires Visual Pinball 10.5 or newer!" & vbNewLine & "Your version: " & Replace(Version/1000,",","."), , "Ace of Speed VPX"
End If

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package."
On Error Goto 0

LoadVPM "01560000", "S11.VBS", 3.26

Dim DesktopMode: DesktopMode = MousinAround.ShowDT
Dim i, bsTrough, bsMouseHoleMultiball

If LetTheBallJump > 6 Then LetTheBallJump = 6 : If LetTheBallJump < 0 Then LetTheBallJump = 0

' **** enables manual ball control with C key (enable/disable control) ********************
' **** and B key (speed boost) and arrow keys *********************************************
Dim IsDebugBallsModeOn : IsDebugBallsModeOn = False

' *** Dynamic flipper friction mod ***
Const DynamicFlipperFriction = True
Const DynamicFlipperFrictionResting = 0.4
Const DynamicFlipperFrictionActive = 1.0


' ****************************************************
' table init
' ****************************************************
 playsound "intro", -1

Sub MousinAround_Init()
  vpmInit Me
  With Controller
        .GameName       = cGameName
        .SplashInfoLine   = "Ace of Speed"  & vbNewLine & "Mod Table by Mussinger, created on Mousin Around VPX from Herweh "
    .Games(cGameName).Settings.Value("sound") = 0
    .HandleMechanics  = False
    .HandleKeyboard   = False
    .ShowDMDOnly    = True
    .ShowFrame      = False
    .ShowTitle      = False
    .Hidden       = DesktopMode
    If cDMDRotation >= 0 Then .Games(cGameName).Settings.Value("rol") = cDMDRotation
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With
  ' initialize some table settings
  InitTable
End Sub

Sub MousinAround_Paused()   : Controller.Pause = True : End Sub
Sub MousinAround_UnPaused() : Controller.Pause = False : End Sub
Sub MousinAround_Exit()   : Controller.Games(cGameName).Settings.Value("sound")=1:Controller.Stop : End Sub

Sub InitTable()
  ' tilt
  vpmNudge.TiltSwitch   = 1
  vpmNudge.Sensitivity  = 5
  vpmNudge.TiltObj    = Array(Bumper52, Bumper53, Bumper54, LeftSlingshot, RightSlingshot)

  ' ball trough
  Set bsTrough = New cvpmBallStack
  With bsTrough
    .InitSw 0, 11, 12, 13, 0, 0, 0, 0
    .InitKick BallRelease, 90, 10
    .Balls = 3
  End With

  ' mouse hole multiball
  Set bsMouseHoleMultiball = New cvpmBallStack
  With bsMouseHoleMultiball
    .InitSw 0, 0, 0, 0, 0, 0, 0, 0
    .InitKick MouseHoleMultiball, 135, 0.1
    .Balls = 0
  End With

  ' init lights, flippers, traps, mouse hole diverter and target motor bank
  InitLights InsertLights
  InitFlippers
  InitTraps
  InitMouseHoleDiverter
  InitMotorBank

  ' init timers
  PinMAMETimer.Interval         = PinMAMEInterval
    PinMAMETimer.Enabled          = True
  LampTimer.Enabled           = True
  GraphicsTimer.Interval        = 10 ' -1
  GraphicsTimer.Enabled       = True
  BallControlTimer.Interval     = 1
  BallControlTimer.Enabled      = False
  RollingSoundTimer.Interval      = 10
  RollingSoundTimer.Enabled     = True

  ' rails visibility
  Leftrail.Visible  = (RailsVisible <> 0)
  Rightrail.Visible   = (RailsVisible <> 0)

  ' backdrop objects
  r57.Visible = DesktopMode
  r58.Visible = DesktopMode
  r59.Visible = DesktopMode
  r60.Visible = DesktopMode
  r61.Visible = DesktopMode
  r62.Visible = DesktopMode
  r63.Visible = DesktopMode
  DisplayTimer.Enabled = DesktopMode

  ' maybe start the timer for the mouse hole multiball modes
  MouseHoleProcessTimer_Timer
  MouseHoleRandomTimer_Timer
End Sub

ChargerRT.blenddisablelighting=0.3 ' Better looking for the ChargerRT Car

' ****************************************************
' For multiball sound
' ****************************************************
Dim countballs
Dim InMultiball : InMultiball = 0
Dim WasMulti : WasMulti = 0
Dim LoosedBall : LoosedBall = 1
Sub LookForBalls_timer()
  countballs = bsTrough.Balls
  If countballs = 0 Then
    InMultiball = 3
    WasMulti = 3
  End If
End Sub


' *********************************************************************
' pistons
' *********************************************************************

  dim posit:posit = 0
  dim stopit:stopit = 0
  sub pistontm_timer()
    if piston1.z = 120 Then
      posit = 0
      stopit = stopit+1
    end If
    if stopit > 7 then pistontm.enabled = 0 : Exit Sub
    if piston1.z = 180 Then
      posit = 1
    end if
    if posit = 0 Then
      piston1.z = piston1.z + 20
      piston2.z = piston2.z - 20
    Else
      piston1.z = piston1.z - 20
      piston2.z = piston2.z + 20
    end If
  end Sub

  sub runpistons
  pistontm.enabled = 1
  end Sub

' *********************************************************************
' Car Jumper
' *********************************************************************
dim carmove : carmove = 0

' 1 jump center loop
  dim pposit:pposit= 0
  dim stoppit:stoppit = 0
  Sub CCarTm_timer()
    If ChargerRT.ObjRotY = 0 Then
      pposit = 0
      stoppit = stoppit+1
    end If
    If stoppit > 1 Then ChargerRT.ObjRotY = 0 : CCarTm.enabled = 0 : carmove = 0 : Exit Sub
    If ChargerRT.ObjRotY = 30 Then
      pposit = 1
    End If
    If pposit = 0 Then
      ChargerRT.ObjRotY = ChargerRT.ObjRotY + 5
    Else
      ChargerRT.ObjRotY= ChargerRT.ObjRotY - 10
    end If
  end Sub

  sub onejump
  CCarTm.enabled = 1
  end Sub

'big jumps right ramp
  dim positt:positt= 0
  dim stopitt:stopitt = 0
  Sub CarTm_timer()
    If ChargerRT.ObjRotY = 0 Then
      positt = 0
      stopitt = stopitt+1
    end If
    If stopitt > 4 Then CarTm.enabled = 0 : carmove = 0 : Exit Sub
    If ChargerRT.ObjRotY = 20 Then
      positt = 1
    End If
    If positt = 0 Then
      ChargerRT.ObjRotY = ChargerRT.ObjRotY + 5
    Else
      ChargerRT.ObjRotY= ChargerRT.ObjRotY - 5
    end If
  end Sub

  sub runcar
  CarTm.enabled = 1
  end Sub

'small jumps left ramp
  dim posittt:posittt= 0
  dim stopittt:stopittt = 0
  sub CarTTm_timer()
    if ChargerRT.ObjRotY = 0 Then
      posittt = 0
      stopittt = stopittt+1
    end If
    if stopittt > 8 Then CarTTm.enabled = 0 : carmove = 0 : Exit Sub
    if ChargerRT.ObjRotY = 4 Then
      posittt = 1
    end if
    if posittt = 0 Then
      ChargerRT.ObjRotY = ChargerRT.ObjRotY + 2
    Else
      ChargerRT.ObjRotY= ChargerRT.ObjRotY - 2
    end If
  end Sub

  sub rumblecar
  CarTTm.enabled = 1
  end Sub

Sub sw001_Hit
  if carmove < 1 then
  carmove = 1
  runcar
  stopitt = 0
  stopsound "Muscle-drift": playsound "Muscle-drift"
  end if
End Sub

Sub sw002_Hit
  if carmove < 1 then
  carmove = 1
  rumblecar
  stopittt = 0
  end if
End Sub

' ****************************************************
' keys
' ****************************************************
Sub MousinAround_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then Plunger.PullBack : PlaySoundAt "plungerpull", Plunger
  If keycode = LeftFlipperKey Then Controller.Switch(58) = True
    If keycode = RightFlipperKey Then Controller.Switch(57) = True
    If keycode = LeftTiltKey Then Nudge 90,2: Playsound SoundFX("fx_nudge",0)
    If keycode = RightTiltKey Then Nudge 270,2: Playsound SoundFX("fx_nudge",0)
    If keycode = CenterTiltKey Then Nudge 0,3: Playsound SoundFX("fx_nudge",0)
  If keycode = LeftMagnaSave Then GIColorMod = (GIColorMod + 1) MOD 4 : ResetGI
  If keycode = RightMagnaSave Then FlipperColorMod = (FlipperColorMod + 1) MOD 5 : ResetFlippers
  If vpmKeyDown(keycode) Then Exit Sub
  ' *** manual ball control ***
  If IsDebugBallsModeOn Then
    If keycode = 46 Then If contball = 1 Then contball = 0 : BallControlTimer.Enabled = False Else contball = 1 : BallControlTimer.Enabled = True: End If : End If ' C Key
    If keycode = 48 Then If bcboost = 1 Then bcboost = bcboostmulti Else bcboost = 1 : End If : End If 'B Key
    If keycode = 203 Then bcleft = 1        ' Left Arrow
    If keycode = 200 Then bcup = 1          ' Up Arrow
    If keycode = 208 Then bcdown = 1        ' Down Arrow
    If keycode = 205 Then bcright = 1       ' Right Arrow
  End If
End Sub

Sub MousinAround_KeyUp(ByVal keycode)
    If keycode = PlungerKey Then Plunger.Fire : PlaySoundAt "plunger", Plunger
  If keycode = LeftFlipperKey Then Controller.Switch(58) = False
    If keycode = RightFlipperKey Then Controller.Switch(57) = False
    If vpmKeyUp(keycode) Then Exit Sub
  ' *** manual ball control ***
  If IsDebugBallsModeOn Then
    If keycode = 203 Then bcleft = 0        ' Left Arrow
    If keycode = 200 Then bcup = 0          ' Up Arrow
    If keycode = 208 Then bcdown = 0        ' Down Arrow
    If keycode = 205 Then bcright = 0       ' Right Arrow
  End If
End Sub


' ****************************************************
' manual ball control
' ****************************************************
Sub StartControl_Hit() : Set ControlBall = ActiveBall : contballinplay = True : PlaySound "carstart": stopSound "people-cars" :End Sub
Sub StopControl_Hit() : contballinplay = false : End Sub

Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1         'Do Not Change - default setting
bcvel = 4           'Controls the speed of the ball movement
bcyveloffset = -0.01    'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3      'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControlTimer_Timer()
  GetABall
    If Contball and ContBallInPlay then
        If bcright = 1 Then ControlBall.velx = bcvel*bcboost Else If bcleft = 1 Then ControlBall.velx = - bcvel*bcboost Else ControlBall.velx=0 : End If : End If
        If bcup = 1 Then ControlBall.vely = -bcvel*bcboost Else If bcdown = 1 Then ControlBall.vely = bcvel*bcboost Else ControlBall.vely= bcyveloffset : End If : End If
    End If
End Sub
Sub GetABall()
  Dim BOT : BOT = GetBalls()
  If UBound(BOT) > -1 Then
    contballinplay = True
    Set ControlBall = BOT(UBound(BOT))
  End If
End Sub


' ****************************************************
' *** solenoids
' ****************************************************
SolCallback(1)          = "SolOuthole"
SolCallback(2)          = "SolBallRelease"
SolCallback(3)      = "SolRightTrapUp"
SolCallback(4)      = "SolLeftTrapUp"
SolCallback(5)      = "SolRightTrapDown"
SolCallback(8)      = "SolLeftTrapDown"
SolCallback(7)      = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(10)     = "SolGI"
SolCallback(11)     = "SolMotorBank"
SolCallback(13)     = "SolKickback"
SolCallback(14)     = "SolMouseHoleDiverter"
SolCallback(16)     = "SolMouseHoleRelease"
SolCallback(22)     = "SolOrbitGate"

SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"

SolCallback(15)     = "Flash 15,"   ' center flashers
SolCallback(25)     = "Flash 25,"   ' right flipper flasher
SolCallback(26)     = "Flash 26,"   ' left flipper flasher
SolCallback(27)     = "Flash 27,"   ' left side flasher
SolCallback(28)     = "Flash28"     ' backboard flasher
SolCallback(29)     = "Flash 29,"   ' top right flasher
SolCallback(30)     = "Flash 30,"   ' right ramp flasher
SolCallback(31)     = "Flash 31,"   ' left ramp flasher
SolCallback(32)     = "Flash 32,"   ' timer flasher


' ******************************************************
' outhole, drain and ball release
' ******************************************************
Sub SolOuthole(Enabled)
  If Enabled Then
    bsTrough.SolIn True
  End If
End Sub
Sub SolBallRelease(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX(IIF(bsTrough.Balls>0,"fx_ballrel","fx_solenoid"), DOFContactors), BallRelease
    bsTrough.SolOut True
  End If
' ******************************************************
' ******** multiball sound on/off    *******************
' ******************************************************

  If InMultiball = 0 and LoosedBall = 1 Then
    stopsound "intro" : stopsound "musclecar"
    stopsound "ROCK_N_ROLL" :   StopSound "Godsmack"
    playsound "ROCK_N_ROLL" ,-1 '
  End If
  If LoosedBall = 0 Then
    'nothing
  End If
  If InMultiball = 3 Then
    StopSound "ROCK_N_ROLL"
    PlaySound "Godsmack", -1
  End If
'********
End Sub

Sub Drain_Hit()
'********

  If WasMulti = 0 Then
    Stopsound "ROCK_N_ROLL"
    If DrainSoundFx = 0 then playsound "Tires"
    If DrainSoundFx = 1 then playsound "sad"
    If DrainSoundFx = 2 then playsound "fx_relay_off"
    playsound "people-cars"
    LoosedBall = 1
  End If
  If WasMulti = 3 Then
    playsound "lol" : playsound "people-cars"
    WasMulti = 2
    LoosedBall = 0
    InMultiball = 0
  Else If WasMulti = 2 Then
    playsound "lol" : playsound "people-cars"
    StopSound "Godsmack"
    Playsound "ROCK_N_ROLL"
    WasMulti = 0
    LoosedBall = 0
    InMultiball = 0
    End If
  End If
' ******************************************************
' ******** multiball sound on/off Code END      ********
' ******************************************************

  BallSearch
  PlaySoundAtVol "drain", Drain, .5
  If lockedMouseHoleBallsOnPF > 0 Then
    Drain.DestroyBall
    lockedMouseHoleBallsOnPF = lockedMouseHoleBallsOnPF - 1
  Else
    vpmTimer.PulseSw 9
    bsTrough.AddBall Me
    If currentMultiballIsRunning And bsTrough.Balls >= 2 Then
      wait3Seconds        = True
      currentMultiballIsRunning   = False
      trapLeftUnloaded      = False
      trapRightUnloaded     = False
      readyForMultiball     = False : Check4Multiball20
      lockedMouseHoleBalls    = 0
      lockedMouseHoleBallsOnPF  = 0
      isMouseHoleOpen       = False : MoveMouseHoleDiverter
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
Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), LeftFlipper, 2
    If DynamicFlipperFriction Then LeftFlipper.Friction = DynamicFlipperFrictionActive
    LeftFlipper.RotateToEnd
    Else
    PlaySoundAt SoundFX("fx_flipperdown",DOFContactors), LeftFlipper
    If DynamicFlipperFriction Then LeftFlipper.Friction = DynamicFlipperFrictionResting
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors),RightFlipper, 2
    If DynamicFlipperFriction Then RightFlipper.Friction = DynamicFlipperFrictionActive
    RightFlipper.RotateToEnd
    Else
    PlaySoundAt SoundFX("fx_flipperdown",DOFContactors),RightFlipper
    If DynamicFlipperFriction Then RightFlipper.Friction = DynamicFlipperFrictionResting
        RightFlipper.RotateToStart
    End If
End Sub

' flipper hit sounds
Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3)+1, 1
End Sub
Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3)+1, 1
End Sub


' ****************************************************
' sling shots and animations
' ****************************************************
Dim LeftStep, RightStep

Sub LeftSlingShot_Slingshot()
  vpmTimer.PulseSw 55
  PlaySoundAt SoundFX("fx_left_slingshot", DOFContactors), LeftSlingHammer
    LeftStep = 0
  LeftSlingShot.TimerInterval = 20
    LeftSlingShot.TimerEnabled  = True
  LeftSlingShot_Timer
End Sub
Sub LeftSlingShot_Timer()
    Select Case LeftStep
    Case 0: LeftSling1.Visible = False : LeftSling3.Visible = True : LeftSlingHammer.TransZ = -28
        Case 1: LeftSling3.Visible = False : LeftSling2.Visible = True : LeftSlingHammer.TransZ = -10
        Case 2: LeftSling2.Visible = False : LeftSling1.Visible = True : LeftSlingHammer.TransZ = 0 : LeftSlingShot.TimerEnabled = 0
    End Select
    LeftStep = LeftStep + 1
End Sub

Sub RightSlingShot_Slingshot()
  vpmTimer.PulseSw 56
  PlaySoundAt SoundFX("fx_right_slingshot", DOFContactors), RightSlingHammer
    RightStep = 0
  RightSlingShot.TimerInterval = 20
    RightSlingShot.TimerEnabled  = True
  RightSlingShot_Timer
End Sub
Sub RightSlingShot_Timer()
    Select Case RightStep
    Case 0: RightSling1.Visible = False : RightSling3.Visible = True : RightSlingHammer.TransZ = -28
        Case 1: RightSling3.Visible = False : RightSling2.Visible = True : RightSlingHammer.TransZ = -10
        Case 2: RightSling2.Visible = False : RightSling1.Visible = True : RightSlingHammer.TransZ = 0 : RightSlingShot.TimerEnabled = 0
    End Select
    RightStep = RightStep + 1
End Sub



' ****************************************************
' kick back
' ****************************************************
Sub SolKickback(Enabled)
    Kickback.Enabled = Enabled
End Sub

Dim KickbackBallVel : KickbackBallVel = 1
Sub Kickback_Hit()
Playsound "imback"
  Kickback.TimerEnabled  = False
  KickbackBallVel = BallVel(ActiveBall)
  Kickback.TimerInterval = 40
    Kickback.TimerEnabled  = True
End Sub
Sub Kickback_Timer()
  If Kickback.TimerInterval = 40 Then
    Kickback.Kick 0, Int(35 + KickbackBallVel*3/4 + Rnd()*20)
    PlaySoundAt SoundFX("plunger", DOFContactors), Kickback
    Kickback.TimerInterval = 20
  ElseIf Kickback.TimerInterval = 20 Then
    pKickback.TransY = pKickback.TransY + 10
    If pKickback.TransY = 30 Then Kickback.TimerInterval = 500
  ElseIf Kickback.TimerInterval = 500 Then
    Kickback.TimerInterval = 50
  ElseIf Kickback.TimerInterval = 50 Then
    pKickback.TransY = pKickback.TransY - 3
    If pKickback.TransY = 00 Then Kickback.TimerEnabled = False
  End If
End Sub


' ****************************************************
' orbit gate
' ****************************************************
Sub SolOrbitGate(Enabled)
  GateTopRight.Open = Enabled
End Sub



' ****************************************************
' additional mouse hole multiball stuff
' ****************************************************
Dim readyForMultiball     : readyForMultiball     = False
Dim isMouseHoleOpen       : isMouseHoleOpen     = False
Dim trapLeftUnloaded      : trapLeftUnloaded      = False
Dim trapRightUnloaded     : trapRightUnloaded     = False
Dim lockedMouseHoleBalls    : lockedMouseHoleBalls    = 0
Dim lockedMouseHoleBallsOnPF  : lockedMouseHoleBallsOnPF  = 0
Dim startMultiball        : startMultiball      = False
Dim currentMultiballIsRunning : currentMultiballIsRunning = False
Dim startExtraBalls       : startExtraBalls     = False
Dim wait3Seconds        : wait3Seconds        = False
Dim noStandardDisplay     : noStandardDisplay     = False
Dim openedMouseHoleTrap     : openedMouseHoleTrap   = 0

Sub Check4Multiball()
  If MouseHoleMultiball.TimerEnabled And MouseHoleMultiball.TimerInterval = 20000 Then Exit Sub
  ' wait for 3 seconds, then check whether multiball is available
  MouseHoleMultiball.TimerEnabled  = False
  MouseHoleMultiball.TimerInterval = 3000
  MouseHoleMultiball.TimerEnabled  = True
End Sub
Sub Check4Multiball20()
  If MouseHoleMultiball.TimerEnabled And MouseHoleMultiball.TimerInterval = 20000 Then Exit Sub
  ' wait for 20 seconds, then check whether multiball is available
  MouseHoleMultiball.TimerEnabled  = False
  MouseHoleMultiball.TimerInterval = 20000
  MouseHoleMultiball.TimerEnabled  = True
End Sub
Sub Check4MultiballSW14()
  If MouseHoleMultiball.TimerEnabled And MouseHoleMultiball.TimerInterval = 20000 Then Exit Sub
  If bsTrough.Balls = 0 Then Exit Sub
  ' immediatelly check whether multiball is available
  readyForMultiball = False
  MouseHoleMultiball.TimerEnabled  = False
  MouseHoleMultiball.TimerInterval = 3000
  MouseHoleMultiball_Timer
End Sub

Dim MouseHoleLightStep : MouseHoleLightStep = 0


Sub MouseHoleProcessTimer_Timer()
  If (ActivateMouseHoleMultiball = 1 Or ActivateMouseHoleMultiball = 2) Then
    If Not MouseHoleProcessTimer.Enabled Then MouseHoleProcessTimer.Enabled = True
    If readyForMultiball And Not currentMultiballIsRunning And Not wait3Seconds Then
      ' get thru the mouse hole multiball steps
      If startExtraBalls Then
        ' unlock the mouse hole multiballs
        startExtraBalls       = False
        readyForMultiball     = False
        currentMultiballIsRunning   = True
        If lockedMouseHoleBalls > 0 Then StartMouseHoleWaitTimer 750
      ElseIf (trapLeftUnloaded And trapRightUnloaded) Then 'Or (trapLeftUnloaded And ballInRightTrap Is Nothing) Or (trapRightUnloaded And ballInLeftTrap Is Nothing) Then
        ' maybe start the mouse hole balls
        startExtraBalls   = True
        ' reset all the other multiball settings
        isMouseHoleOpen   = False : MoveMouseHoleDiverter
        trapLeftUnloaded  = False
        trapRightUnloaded = False
        sw32.Enabled    = True
        sw61.Enabled    = True
        sw60.Enabled    = True
        MouseHole.Enabled   = True
        MouseHoleMultiball.Enabled  = False
      ElseIf MouseHoleMultiball.Enabled And lockedMouseHoleBalls >= 2 Then
        ' disable mouse hole multiball objects
        sw32.Enabled    = True
        sw61.Enabled    = True
        sw60.Enabled    = True
        MouseHole.Enabled   = True
        MouseHoleMultiball.Enabled  = False
      ElseIf (ActivateMouseHoleMultiball = 1 And openedMouseHoleTrap < 3 And isMouseHoleOpen) Or _
           (ActivateMouseHoleMultiball = 2 And isMouseHoleOpen And lockedMouseHoleBalls >= 2) Then
        ' close mouse hole as two balls are locked
        isMouseHoleOpen   = False : MoveMouseHoleDiverter
      ElseIf (ActivateMouseHoleMultiball = 1 And openedMouseHoleTrap >= 3 And Not isMouseHoleOpen And lockedMouseHoleBalls < 2) Or _
           (ActivateMouseHoleMultiball = 2 And Not isMouseHoleOpen And lockedMouseHoleBalls < 2) Then
        ' open the mouse hole for additional locking
        isMouseHoleOpen   = True : MoveMouseHoleDiverter
        sw32.Enabled    = False
        sw61.Enabled    = False
        sw60.Enabled    = False
        MouseHole.Enabled   = False
        MouseHoleMultiball.Enabled  = True
        ' start clip
        MouseHoleWaitTimer.Enabled  = False
        StartMouseHoleWaitTimer 1001
      End If
    End If
    ' maybe flash the mouse hole entry light
    If readyForMultiball And isMouseHoleOpen Then
      MouseHoleLightStep = MouseHoleLightStep + 1
      If MouseHoleLightStep = 10 Then
        fMouseHole.Visible = True
      ElseIf MouseHoleLightStep >= 20 Then
        MouseHoleLightStep = 0
        fMouseHole.Visible = False
      End If
    Else
      MouseHoleLightStep  = 0
      fMouseHole.Visible  = False
      ' maybe close mouse hole
      If isMouseHoleOpen Then
        isMouseHoleOpen = False : MoveMouseHoleDiverter
      End If
    End If
  End If
End Sub

Sub MouseHoleMultiball_Hit()
  ' lock a ball at the mouse hole
  bsMouseHoleMultiball.AddBall Me
  lockedMouseHoleBalls = lockedMouseHoleBalls + 1
  ' start clip
  MouseHoleWaitTimer.Enabled  = False
  StartMouseHoleWaitTimer 1000
End Sub
Sub MouseHoleMultiball_Timer()
  MouseHoleMultiball.TimerEnabled = False
  wait3Seconds = False
  If readyForMultiball Then Exit Sub
  Dim obj
  readyForMultiball = True
  For Each obj In Array(l17,l18,l19,l20,l21,l25,l26,l27,l28)
    If obj.State <> LightStateOn Then
      readyForMultiball = False
      Exit For
    End If
  Next
  If Not readyForMultiball Then
    lockedMouseHoleBalls    = 0
    lockedMouseHoleBallsOnPF  = 0
    startExtraBalls       = False
    currentMultiballIsRunning = False
  Else
    MouseHoleRandomTimer_Timer
  End If
End Sub

Dim MouseHoleWaitStep : MouseHoleWaitStep = 0
Sub StartMouseHoleWaitTimer(interval)
  If MouseHoleWaitTimer.Enabled And interval <> 750 Then Exit Sub
  MouseHoleWaitTimer.Enabled  = False
  MouseHoleWaitTimer.Interval = interval
  MouseHoleWaitTimer_Timer
End Sub
Sub MouseHoleWaitTimer_Timer()
  If Not MouseHoleWaitTimer.Enabled Then MouseHoleWaitTimer.Enabled = True
  MouseHoleWaitStep = MouseHoleWaitStep + 1
  If MouseHoleWaitTimer.Interval = 750 Then
    ' unlock locked mouse hole balls
    If bsMouseHoleMultiball.Balls <= 0 Or lockedMouseHoleBalls <= 0 Then
      MouseHoleWaitTimer.Enabled = False
      MouseHoleWaitStep = 0
    Else
      PlaySoundAt SoundFX("fx_solenoid", DOFContactors), MouseHoleMultiball
      bsMouseHoleMultiball.SolOut True
      lockedMouseHoleBalls    = lockedMouseHoleBalls - 1
      lockedMouseHoleBallsOnPF  = lockedMouseHoleBallsOnPF + 1
    End If
  ElseIf MouseHoleWaitTimer.Interval = 1000 Then
    ' ball is locked in mouse hole trap
    Select Case MouseHoleWaitStep
    Case 1
      If DesktopMode Then noStandardDisplay = True
      PlaySound "ma_Gotcha"
      PlaySound "v8"
      SolGI True
      SetMouseHoleLightsOff
    Case 2
      SetLEDs "     YOU     " & "   GOT IT!    ", 1
      StartMouseHoleFlasher
    Case 4
      SetLEDs "     GARAGE     " & " BALL " & lockedMouseHoleBalls & " LOCKED  ", 2
    Case 5
      StopMouseHoleFlasher
    Case 6
      SetLEDs "     GARAGE     " & " BALL " & lockedMouseHoleBalls & " LOCKED  ", 1
      SolGI False
      SetMouseHoleLightsOn
    Case 7
      If bsMouseHoleMultiball.Balls > 2 Then
        StartMouseHoleFlasherS
      Else
        ' stop and reset current clip
        MouseHoleWaitTimer.Enabled = False
        ClearAllLEDs : If DesktopMode Then noStandardDisplay = False
        MouseHoleWaitStep = 0
        ' create next ball at the ball release
        PlaySoundAt SoundFX("fx_ballrel", DOFContactors), BallRelease
        BallRelease.CreateSizedballWithMass Ballsize/2,Ballmass
        BallRelease.Kick 90,10
        SolOrbitGate False
      End If
    Case 8
      ' kick next ball out of the mouse hole exit
      MouseHoleWaitTimer.Enabled = False
      StopMouseHoleFlasher
      ClearAllLEDs : If DesktopMode Then noStandardDisplay = False
      MouseHoleWaitStep = 0
      PlaySoundAt SoundFX("fx_solenoid", DOFContactors), MouseHoleMultiball
      bsMouseHoleMultiball.SolOut True
    End Select
  ElseIf MouseHoleWaitTimer.Interval = 1001 Then
    ' mouse hole trap is open too
    Select Case MouseHoleWaitStep
    Case 1
      If DesktopMode Then noStandardDisplay = True
      SetLEDs "    GARAGE    " & "  TRAP IS OPEN  ", 3
    Case 3
      MouseHoleWaitTimer.Enabled = False
      ClearAllLEDs : If DesktopMode Then noStandardDisplay = False
      MouseHoleWaitStep = 0
    End Select
  End If
End Sub

' the timer to randomly pick one of up to three traps and open it for locking the ball
Dim MouseHoleRandomStep : MouseHoleRandomStep = 0
Sub MouseHoleRandomTimer_Timer()
  If ActivateMouseHoleMultiball = 1 Then
    If Not MouseHoleRandomTimer.Enabled Then MouseHoleRandomTimer.Enabled = True
    If readyForMultiball And Not currentMultiballIsRunning And Not wait3Seconds And bsTrough.Balls < 3 Then
      If MouseHoleRandomStep >= 20 Then MouseHoleRandomStep = 0
      MouseHoleRandomStep = MouseHoleRandomStep + 1
      Dim availableTraps : availableTraps = IIF(lockedMouseHoleBalls=0,2,0) + IIF(lockedMouseHoleBalls=1,1,0) + IIF(ballInLeftTrap Is Nothing,1,0) + IIF(ballInRightTrap Is Nothing,1,0)
      If MouseHoleRandomStep = 1 Then
        ClearAllLEDs : If DesktopMode Then noStandardDisplay = False
        ' maybe close current trap
        If openedMouseHoleTrap = 1 Then TrapDown 34, leftTrap42, pLeftTrap, ballInLeftTrap
        If openedMouseHoleTrap = 2 Then TrapDown 33, rightTrap41, pRightTrap, ballInRightTrap
        ' open new trap
        If availableTraps >= 3 Then
          openedMouseHoleTrap = 0
          Do While openedMouseHoleTrap = 0 Or (Not ballInLeftTrap Is Nothing And openedMouseHoleTrap = 1) Or (Not ballInRightTrap Is Nothing And openedMouseHoleTrap = 2)
            openedMouseHoleTrap = Int(Rnd()*4) + 1
          Loop
        ElseIf availableTraps = 2 Then
          If ballInLeftTrap Is Nothing And ballInRightTrap Is Nothing Then
            openedMouseHoleTrap = IIF(openedMouseHoleTrap=1,2,1)
          ElseIf ballInLeftTrap Is Nothing And lockedMouseHoleBalls < 2 Then
            openedMouseHoleTrap = IIF(openedMouseHoleTrap=1,3,1)
          ElseIf ballInRightTrap Is Nothing And lockedMouseHoleBalls < 2 Then
            openedMouseHoleTrap = IIF(openedMouseHoleTrap=2,3,2)
          End If
        ElseIf availableTraps = 1 Then
          openedMouseHoleTrap = 0
          If ballInLeftTrap Is Nothing Then
            openedMouseHoleTrap = 1
          ElseIf ballInRightTrap Is Nothing Then
            openedMouseHoleTrap = 2
          ElseIf lockedMouseHoleBalls < 2 Then
            openedMouseHoleTrap = 3
          End If
        ElseIf availableTraps = 0 Then
          openedMouseHoleTrap = 0
        End If
      ElseIf MouseHoleRandomStep >= 16 And MouseHoleRandomStep <= 20 Then
        ' countdown for next trap
        If DesktopMode Then noStandardDisplay = True
        SetLEDs " NEXT BALL TRAP " & " OPENS IN " & (21-MouseHoleRandomStep) & " SEC ", IIF(MouseHoleRandomStep >= 19,1,3)
      End If
    ElseIf MouseHoleRandomTimer.Enabled Then
      MouseHoleRandomTimer.Enabled = False
      openedMouseHoleTrap = 0
      MouseHoleRandomStep = 0
      ClearAllLEDs : If DesktopMode Then noStandardDisplay = False
    End If
  End If
End Sub

Dim MouseHoleLightsOn : MouseHoleLightsOn = ""
Sub SetMouseHoleLightsOff()
  Dim obj
  MouseHoleLightsOn = ","
  For Each obj In InsertLights
    If Left(obj.Name,1) = "l" Then
      If obj.State = LightStateOn Then
        obj.State = LightStateOff
        MouseHoleLightsOn = MouseHoleLightsOn & obj.Name & ","
      End If
    End If
  Next
End Sub
Sub SetMouseHoleLightsOn()
  Dim obj
  For Each obj In InsertLights
    If InStr(MouseHoleLightsOn, "," & obj.Name & ",") > 0 Then
      obj.State = LightStateOn
    End If
  Next
  MouseHoleLightsOn = ""
End Sub

Dim isFlasherOn   : isFlasherOn     = False
Dim justExitFlasher : justExitFlasher   = False
Sub StartMouseHoleFlasher()
  justExitFlasher = False
  MouseHoleFlasherTimer.Enabled = True
End Sub
Sub StartMouseHoleFlasherS()
  justExitFlasher = True
  MouseHoleFlasherTimer.Enabled = True
End Sub
Sub StopMouseHoleFlasher()
  MouseHoleFlasherTimer.Enabled = False
  isFlasherOn = True : justExitFlasher = False : MouseHoleFlasherTimer_Timer
End Sub
Sub MouseHoleFlasherTimer_Timer()
  isFlasherOn = Not isFlasherOn
  Flash28 isFlasherOn
  If Not justExitFlasher Then
    Flash 29, isFlasherOn
    Flash 30, isFlasherOn
    Flash 31, isFlasherOn
  End If
End Sub

Dim ledChar(25), ledNo(9), ledDot
ledChar(0)  = 1+2+4+16+32+64+2048
ledChar(1)  = 1+2+4+8+512+2048+8192
ledChar(2)  = 1+8+16+32
ledChar(3)  = 1+2+4+8+512+8192
ledChar(4)  = 1+8+16+32+64
ledChar(5)  = 1+16+32+64
ledChar(6)  = 1+4+8+16+32+2048
ledChar(7)  = 2+4+16+32+64+2048
ledChar(8)  = 1+8+512+8192
ledChar(9)  = 2+4+8+16
ledChar(10) = 16+32+64+1024+4096
ledChar(11) = 8+16+32
ledChar(12) = 2+4+16+32+256+1024
ledChar(13) = 2+4+16+32+256+4096
ledChar(14) = 1+2+4+8+16+32
ledChar(15) = 1+2+16+32+64+2048
ledChar(16) = 1+2+4+8+16+32+4096
ledChar(17) = 1+2+16+32+64+2048+4096
ledChar(18) = 1+4+8+32+64+2048
ledChar(19) = 1+512+8192
ledChar(20) = 2+4+8+16+32
ledChar(21) = 0
ledChar(22) = 2+4+16+32+16384
ledChar(23) = 256+1024+4096+16384
ledChar(24) = 256+1024+8192
ledChar(25) = 1+8+1024+16384
ledNo(0)  = 1+2+4+8+16+32+16384+1024
ledNo(1)  = 512+8192
ledNo(2)  = 1+2+8+16+64+2048
ledNo(3)  = 1+2+4+8+2048
ledNo(4)  = 2+4+32+64+2048
ledNo(5)  = 1+4+8+32+64+2048
ledDot    = 32768

Sub ClearAllLEDs()
  Dim ii
  For ii = 1 To 32 : SetLED ii, 0 : Next
  MouseHoleLEDTimer.Enabled = False
End Sub
Dim currentLEDText : currentLEDText = ""
Dim currentLEDMode : currentLEDMode = 0
Sub SetLEDs(text, mode)
  currentLEDText    = text
  currentLEDMode    = mode
  MouseHoleLEDStep  = 0
  MouseHoleLEDTimer_Timer
End Sub
Sub SetLED(digit, value)
  If DesktopMode Then
    If digit > 32 then Exit Sub
    Dim obj, digitValue
    digitValue = value
    For Each obj In Digits(digit-1)
      If digitValue And 1 Then
        obj.State = LightStateOn : digitValue = digitValue - 1
      Else
        obj.State = LightStateOff
      End If
      digitValue = digitValue / 2
    Next
  ElseIf B2SOn Then
    Controller.B2SSetLED digit, value
  End If
End Sub

Dim MouseHoleLEDStep : MouseHoleLEDStep = 0
Sub MouseHoleLEDTimer_Timer()
  If Not MouseHoleLEDTimer.Enabled Then MouseHoleLEDTimer.Enabled = True
  Dim ii, char
  If currentLEDMode = 2 Then
    If MouseHoleLEDStep > 40 Then MouseHoleLEDStep = 0 : For ii = 1 To 32 : SetLED ii, 0 : Next
    MouseHoleLEDStep = MouseHoleLEDStep + 1
  ElseIf currentLEDMode = 3 Then
    If MouseHoleLEDStep > 3 Then For ii = 1 To 32 : SetLED ii, 0 : Next
    If MouseHoleLEDStep > 6 Then MouseHoleLEDStep = 0
    MouseHoleLEDStep = MouseHoleLEDStep + 1
  End If
  If currentLEDMode <> 3 Or MouseHoleLEDStep <= 3 Then
    For ii = 1 To IIF(currentLEDMode=2 And MouseHoleLEDStep<Len(currentLEDText), MouseHoleLEDStep, Len(currentLEDText))
      char = UCase(Mid(currentLEDText,ii,1))
      If Asc(char) >= 65 And Asc(char) <= 90 Then
        SetLED ii, ledChar(Asc(char)-65)
      ElseIf Asc(char) >= 48 And Asc(char) <= 53 Then
        SetLED ii, ledNo(Asc(char)-48)
      ElseIf Asc(char) = 46 Then
        SetLED ii, ledDot
      Else
        SetLED ii, 0
      End If
    Next
  End If
End Sub


' ****************************************************
' switches
' ****************************************************
' stand-up targets
Sub sw17_Hit()   : StandUpTarget 17, sw17, pTarget17, pTargetPost17, True  : Check4Multiball :PlaySound "sexy1":  End Sub
Sub sw17_Timer() : StandUpTarget 17, sw17, pTarget17, pTargetPost17, False : End Sub
Sub sw18_Hit()   : StandUpTarget 18, sw18, pTarget18, pTargetPost18, True  : Check4Multiball :PlaySound "sexy2":  End Sub
Sub sw18_Timer() : StandUpTarget 18, sw18, pTarget18, pTargetPost18, False : End Sub
Sub sw19_Hit()   : StandUpTarget 19, sw19, pTarget19, pTargetPost19, True  : Check4Multiball :PlaySound "sexy3":  End Sub
Sub sw19_Timer() : StandUpTarget 19, sw19, pTarget19, pTargetPost19, False : End Sub
Sub sw20_Hit()   : StandUpTarget 20, sw20, pTarget20, pTargetPost20, True  : Check4Multiball :PlaySound "sexy1":  End Sub
Sub sw20_Timer() : StandUpTarget 20, sw20, pTarget20, pTargetPost20, False : End Sub
Sub sw21_Hit()   : StandUpTarget 21, sw21, pTarget21, pTargetPost21, True  : Check4Multiball :PlaySound "sexy2":  End Sub
Sub sw21_Timer() : StandUpTarget 21, sw21, pTarget21, pTargetPost21, False : End Sub
Sub sw25_Hit()   : StandUpTarget 25, sw25, pTarget25, pTargetPost25, True  : Check4Multiball :PlaySound "sexy3":  End Sub
Sub sw25_Timer() : StandUpTarget 25, sw25, pTarget25, pTargetPost25, False : End Sub
Sub sw26_Hit()   : StandUpTarget 26, sw26, pTarget26, pTargetPost26, True  : Check4Multiball :PlaySound "sexy1":  End Sub
Sub sw26_Timer() : StandUpTarget 26, sw26, pTarget26, pTargetPost26, False : End Sub
Sub sw27_Hit()   : StandUpTarget 27, sw27, pTarget27, pTargetPost27, True  : Check4Multiball :PlaySound "sexy2":  End Sub
Sub sw27_Timer() : StandUpTarget 27, sw27, pTarget27, pTargetPost27, False : End Sub
Sub sw28_Hit()   : StandUpTarget 28, sw28, pTarget28, pTargetPost28, True  : Check4Multiball :PlaySound "sexy3":  End Sub
Sub sw28_Timer() : StandUpTarget 28, sw28, pTarget28, pTargetPost28, False : End Sub
Sub sw36_Hit()   : StandUpTarget 36, sw36, pTarget36, pTargetPost36, True  : PlaySound "sexy-mother" : End Sub
Sub sw36_Timer() : StandUpTarget 36, sw36, pTarget36, pTargetPost36, False : End Sub

' standup-targets in front of the motor bank
Sub sw29_Hit()   : StandUpTarget 29, sw29, pTarget29, pTargetPost29, True  : PlaySound "yessir1": End Sub
Sub sw29_Timer() : StandUpTarget 29, sw29, pTarget29, pTargetPost29, False : End Sub
Sub sw30_Hit()   : StandUpTarget 30, sw30, pTarget30, pTargetPost30, True  : PlaySound "yessir2":End Sub
Sub sw30_Timer() : StandUpTarget 30, sw30, pTarget30, pTargetPost30, False : End Sub
Sub sw31_Hit()   : StandUpTarget 31, sw31, pTarget31, pTargetPost31, True  : PlaySound "yessir3":End Sub
Sub sw31_Timer() : StandUpTarget 31, sw31, pTarget31, pTargetPost31, False : End Sub

' top lanes
Sub sw22_Hit()   : Controller.Switch(22) = True  : RollOverSound : PlaySound "sexy1":  End Sub
Sub sw22_Unhit() : Controller.Switch(22) = False : SolOrbitGate True : End Sub
Sub sw23_Hit()   : Controller.Switch(23) = True  : RollOverSound : PlaySound "sexy2": End Sub
Sub sw23_Unhit() : Controller.Switch(23) = False : SolOrbitGate True : End Sub
Sub sw24_Hit()   : Controller.Switch(24) = True  : RollOverSound : PlaySound "sexy3": End Sub
Sub sw24_Unhit() : Controller.Switch(24) = False : SolOrbitGate True : End Sub

' orbit lanes
Sub sw15_Hit()   : Controller.Switch(15) = True  : RollOverSound : Playsound "goodboy" : End Sub
Sub sw15_Unhit() : Controller.Switch(15) = False : End Sub
Sub sw16_Hit()   : Controller.Switch(16) = True  : stopsound "mucleorbit": playsound "mucleorbit" :RollOverSound : End Sub
Sub sw16_Unhit() : Controller.Switch(16) = False : End Sub

' inlanes and outlanes
Sub sw38_Hit() :   Controller.Switch(38) = True  : RollOverSound : PlaySound "whistle" : Check4Multiball : End Sub
Sub sw38_Unhit() : Controller.Switch(38) = False : End Sub
Sub sw39_Hit() :   Controller.Switch(39) = True  : RollOverSound : PlaySound "heybaby" : Check4Multiball : End Sub
Sub sw39_Unhit() : Controller.Switch(39) = False : End Sub
Sub sw40_Hit() :   Controller.Switch(40) = True  : RollOverSound : PlaySound "this-is-ridiculous" : End Sub
Sub sw40_Unhit() : Controller.Switch(40) = False : End Sub

' kickback at left outlane
Sub sw51_Hit() :   Controller.Switch(51) = True  : RollOverSound : End Sub
Sub sw51_Unhit() : Controller.Switch(51) = False : End Sub

' plunger switch
Sub sw14_Hit()   : Controller.Switch(14) = True  : RollOverSound : Check4MultiballSW14 : positt= 0 : End Sub
Sub sw14_UnHit() : Controller.Switch(14) = False : End Sub

' mouse hole switches
Sub sw61_Hit()   : Controller.Switch(61) = True  : RollOverSound : End Sub
Sub sw61_Unhit() : Controller.Switch(61) = False : End Sub
Sub sw60_Hit()   : Controller.Switch(60) = True  : RollOverSound : End Sub
Sub sw60_Unhit() : Controller.Switch(60) = False : End Sub

' bumpers
Sub Bumper52_Hit() : vpmTimer.PulseSw 52 : PlaySoundAt SoundFX("fx_bumper1",DOFContactors), Bumper52 : End Sub
Sub Bumper53_Hit() : vpmTimer.PulseSw 53 : PlaySoundAt SoundFX("fx_bumper2",DOFContactors), Bumper53 : End Sub
Sub Bumper54_Hit() : vpmTimer.PulseSw 54 : PlaySoundAt SoundFX("fx_bumper3",DOFContactors), Bumper54 : End Sub

' ramp triggers
Sub sw47_Hit() : vpmTimer.PulseSw 47 : stopsound "Muscle-vroom" : playsound "Muscle-vroom"  End Sub
Sub sw46_Hit() : vpmTimer.PulseSw 46 : RollOverSound : End Sub
Sub sw37_Hit() : vpmTimer.PulseSw 37 : RollOverSound : End Sub
Sub sw35_Hit() : vpmTimer.PulseSw 35 : SlowDownBall ActiveBall,1 : End Sub
Sub sw44_Hit() : vpmTimer.PulseSw 44 : CheckGuideToTop : End Sub
Sub sw32_Hit() : vpmTimer.PulseSw 32 : RollOverSound : End Sub

Sub StandUpTarget(id, sw, prim, primPost, isHit)
  If isHit And sw.TimerEnabled Then Exit Sub
  If id = 29 Or id = 30 Or id = 31 Or id = 36 Then prim.TransX = prim.TransX + 1 Else prim.TransZ = prim.TransZ + 1
  primPost.TransZ = primPost.TransZ + 1
  If primPost.TransZ >= 5 Then
    If id = 29 Or id = 30 Or id = 31 Or id = 36 Then prim.TransX = 0 Else prim.TransZ = 0
    primPost.TransZ = 0
    sw.TimerEnabled = False
    Exit Sub
  End If
  If Not sw.TimerEnabled Then
    StandupTargetHit
    vpmTimer.PulseSw id
    sw.TimerInterval = 11
    sw.TimerEnabled = True
  End If
End Sub


' *********************************************************************
' cheese loop
' *********************************************************************
Sub trCheeseLoopEnter_Hit()
  If ActiveBall.VelY < -10 Then
    PlaySoundAt "fx_metalrolling0", trCheeseLoopEnter
    CheeseLoopSoundTimer.Enabled = True
    Stopsound "v8-forlooping" : Playsound "v8-forlooping"
    Stopit=0 : runpistons
    if carmove < 1 then
    carmove = 1
    Stoppit=0 : onejump
    End If
  End If
End Sub

Sub CheeseLoopSoundTimer_Timer()
  CheeseLoopSoundTimer.Enabled = False
  StopSound "fx_metalrolling1"
End Sub


' *********************************************************************
' mouse hole
' *********************************************************************
Sub MouseHole_Hit()
  PlaySoundAt "ferris_hit2", MouseHole
  Controller.Switch(59) = True
End Sub
Sub SolMouseHoleRelease(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("fx_solenoid", DOFContactors), MouseHole, 2
    MouseHole.Kick 135, 0.1
    Controller.Switch(59) = False
  End If
End Sub

Sub MouseHoleSlowDown_Hit()
  With ActiveBall
    PlaySound "fx_balldrop" & Int(Rnd()*3), 0, ABS(.VelZ)/17*3, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    .VelX = 0 : .VelY = 0
  End With
End Sub


' *********************************************************************
' ball traps up and down
' *********************************************************************
Dim ballInLeftTrap  : Set ballInLeftTrap  = Nothing
Dim ballInRightTrap : Set ballInRightTrap   = Nothing
Dim isLeftTrapUp  : isLeftTrapUp      = False

Sub InitTraps()
  SolLeftTrapDown True
  SolRightTrapDown True
End Sub

' traps going up or down
Sub SolLeftTrapDown(Enabled)
  If Enabled Then
    TrapDown 34, leftTrap42, pLeftTrap, ballInLeftTrap
  End If
End Sub
Sub SolLeftTrapUp(Enabled)
  If Enabled Then
    If Not ballInLeftTrap Is Nothing Then trapLeftUnloaded = True
    TrapUp 34, 42, leftTrap42, pLeftTrap, ballInLeftTrap
  End If
End Sub
Sub SolRightTrapDown(Enabled)
  If Enabled Then
    TrapDown 33, rightTrap41, pRightTrap, ballInRightTrap
  End If
End Sub
Sub SolRightTrapUp(Enabled)
  If Enabled Then
    If Not ballInRightTrap Is Nothing Then trapRightUnloaded = True
    TrapUp 33, 41, rightTrap41, pRightTrap, ballInRightTrap
  End If
End Sub

' trap is hit by a ball
Sub leftTrap42_Hit()
  BallInTrap 34, 42, leftTrap42, pLeftTrap, ballInLeftTrap, ActiveBall
  LoosedBall = 0
End Sub
Sub leftTrap42_Timer()
  MoveTrap 42, leftTrap42, pLeftTrap, ballInLeftTrap, 0
End Sub
Sub rightTrap41_Hit()
  BallInTrap 33, 41, rightTrap41, pRightTrap, ballInRightTrap, ActiveBall
  LoosedBall = 0
End Sub
Sub rightTrap41_Timer()
  MoveTrap 41, rightTrap41, pRightTrap, ballInRightTrap, 0
End Sub

Sub TrapDown(id, trap, prim, ball)
  trap.Enabled = False
  PlaySoundAt SoundFX("fx_solenoid", DOFContactors), trap
  PlaySound "ahooha_horn"
  Controller.Switch(id) = True
  ' move trap prim and ball down
  trap.TimerEnabled = False
  MoveTrap 0, trap, prim, ball, 15
End Sub
Sub TrapUp(id, idBallInTrap, trap, prim, ball)
  trap.Enabled = True
  PlaySoundAt SoundFX("fx_solenoid", DOFContactors), trap
  PlaySound "ahooha_horn"
  Playsound "V8"
  Playsound "locksopen"

  Controller.Switch(id) = False
  ' move trap prim and ball up
  trap.TimerEnabled = False
  MoveTrap 0, trap, prim, ball, 11
End Sub

Sub BallInTrap(id, idBallInTrap, trap, prim, ball, ballA)
  PlaySoundAt SoundFX("fx_solenoid", DOFContactors), trap
  Set ball = ballA : ball.VelX = 0 : ball.VelY = 0 : ball.VelZ = 0
  Controller.Switch(idBallInTrap) = True
  TrapDown id, trap, prim, ball
  PlaySoundAt "ferris_hit1", trap
  PlaySound "ahooha_horn"
End Sub

Sub MoveTrap(idBallInTrap, trap, prim, ball, interval)
  If Not trap.TimerEnabled Then
    trap.TimerInterval = interval
    trap.TimerEnabled  = True
  End If
  If trap.TimerInterval = 15 Then
    ' trap goes down
    prim.TransZ = prim.TransZ - 6
    If Not ball Is Nothing Then ball.Z = ball.Z - 6
    If prim.TransZ <= -60 Then
      trap.TimerEnabled = False
      If Not ball Is Nothing Then ball.Z = ball.Z - 25
    End If
  ElseIf trap.TimerInterval = 11 Then
    If ActivateMouseHoleMultiball = 1 Then
      trap.Enabled = currentMultiballIsRunning Or Not ((Left(trap.Name,4)="left" And openedMouseHoleTrap <> 1 And ballInLeftTrap Is Nothing) Or (Left(trap.Name,5)="right" And openedMouseHoleTrap <> 2 And ballInRightTrap Is Nothing))
      If Not trap.Enabled Then Exit Sub
    End If
    ' trap goes up
    If prim.TransZ <= -60 And Not ball Is Nothing Then ball.Z = ball.Z + 25
    prim.TransZ = prim.TransZ + 12
    If Not ball Is Nothing Then ball.Z = ball.Z + 12
    If prim.TransZ >= 0 Then
      trap.TimerInterval = 100
    End If
  Else
    trap.TimerEnabled = False
    ' kick the trapped ball
    If Not ball Is Nothing Then
      trap.Kick 185, 0.5
      Controller.Switch(idBallInTrap) = False
      Set ball = Nothing
    End If
  End If
End Sub


' *********************************************************************
' ball diverter on ramp
' *********************************************************************
Sub InitMouseHoleDiverter()
  Controller.Switch(49) = False
  SolMouseHoleDiverter True
End Sub

Sub SolMouseHoleDiverter(Enabled)
  If Enabled Then
    Controller.Switch(49) = Not Controller.Switch(49)
    MoveMouseHoleDiverter
  End If
End Sub

Dim DivStep : DivStep = 0
Sub MoveMouseHoleDiverter()
  MouseHoleDiverterTimer.Enabled  = False
  PlaySoundAt SoundFX("fx_popper", DOFContactors), pDiverterArm
  WallDiverter.IsDropped      = Not (Controller.Switch(49) And Not isMouseHoleOpen)
  CheckGuideToTop
  DivStep             = IIF(Controller.Switch(49) And Not isMouseHoleOpen, -3, 3)
  MouseHoleDiverterTimer.Interval = 11
  MouseHoleDiverterTimer_Timer
End Sub
Sub MouseHoleDiverterTimer_Timer()
  If Not MouseHoleDiverterTimer.Enabled Then MouseHoleDiverterTimer.Enabled = True
  If pDiverterArm.ObjRotZ > 23 Then
    pDiverterArm.ObjRotZ = 23
    MouseHoleDiverterTimer.Enabled  = False
  ElseIf pDiverterArm.ObjRotZ < 0 Then
    pDiverterArm.ObjRotZ = 0
    MouseHoleDiverterTimer.Enabled  = False
  End If
  If MouseHoleDiverterTimer.Enabled Then
    pDiverterArm.ObjRotZ = pDiverterArm.ObjRotZ + DivStep
  End If
End Sub

Sub CheckGuideToTop()
  WallGuideToTop.IsDropped = Not WallDiverter.IsDropped
End Sub
Sub TriggerGuideToTop_Hit()
  WallGuideToTop.IsDropped = True
End Sub


' *********************************************************************
' target motor bank
' *********************************************************************
Sub InitMotorBank()
  Controller.Switch(43)       = True
  Controller.Switch(50)       = False
  MotorBankTimer.Interval     = 39
  MotorBankTimer.Enabled      = False
  MotorBankSwitchTimer.Enabled  = False
End Sub

Dim isMotorbankUp : isMotorbankUp = True
Dim MBStep : MBStep = 0
Sub SolMotorBank(Enabled)
  MotorBankSwitchTimer.Enabled  = False
  If Enabled Then
    If Controller.Switch(43) And Not Controller.Switch(50) Then
      Controller.Switch(43) = False
      MotorBankSwitchTimer.Interval = 1999
      MotorBankSwitchTimer.Enabled  = True
      MBStep = -1 : MotorBankTimer_Timer 'MotorBankTimer.Enabled = True
    ElseIf Not Controller.Switch(43) And Controller.Switch(50) Then
      Controller.Switch(50) = False
      MotorBankSwitchTimer.Interval = 2001
      MotorBankSwitchTimer.Enabled  = True
      MBStep = 1 : MotorBankTimer_Timer 'MotorBankTimer.Enabled = True
    End If
  Else
    MotorBankSwitchTimer.Enabled = False
    If Controller.Switch(43) = Controller.Switch(50) Then
      MotorBankSwitchTimer.Interval = 50
      MotorBankSwitchTimer.Enabled  = True
      MBStep = MBStep * -1 : MotorBankTimer_Timer
    End If
  End If

End Sub
Sub MotorBankSwitchTimer_Timer()
  MotorBankSwitchTimer.Enabled = False
  If MotorBankSwitchTimer.Interval = 50 Then
    Controller.Switch(43) = Not Controller.Switch(43)
  ElseIf MotorBankSwitchTimer.Interval = 1999 Then
    Controller.Switch(50) = True
  ElseIf MotorBankSwitchTimer.Interval = 2001 Then
    Controller.Switch(43) = True
  End If
End Sub
Sub MotorBankTimer_Timer()
  If Not MotorBankTimer.Enabled Then MotorBankTimer.Enabled = True
  Dim obj, obj1
  Set obj1 = Nothing
  For Each obj In Array(pTarget29, pTarget30, pTarget31, pTargetPost29, pTargetPost30, pTargetPost31)
    obj.TransY = obj.TransY + MBStep
    If obj1 Is Nothing Then Set obj1 = obj
  Next
  pMotorbankWall.TransZ = pMotorbankWall.TransZ + MBStep
  ' switch off
  If obj1.TransY = -1 And MBStep = -1 Then
    PlayLoopSoundAtVol "fx_motor", sw30, 0.5
  ElseIf obj1.TransY = -52 And MBStep = 1 Then
    PlayLoopSoundAtVol "fx_motor", sw30, 0.5
    isMotorbankUp = True : UpdateGILightsUnderCheeseLoop
    For Each obj In Array(sw29, sw30, sw31, MotorbankWall) : obj.Collidable = True : Next
  End If
  ' switch on
  If obj1.TransY <= -53 Or obj1.TransY >= 0 Then
    MotorBankTimer.Enabled = False
    StopSound "fx_motor"
    If MBStep = -1 Then
      isMotorbankUp = False : UpdateGILightsUnderCheeseLoop
      For Each obj In Array(sw29, sw30, sw31, MotorbankWall) : obj.Collidable = False : Next
    End If
  End If
End Sub


' *********************************************************************
' lamps and illumination
' *********************************************************************
' inserts
Dim PFLights(100,3), PFLightsCount(100), currentGILevel, isGIOn

Sub InitLights(aColl)
  ' init inserts
  Dim obj, idx
  For Each obj In aColl
    idx = obj.TimerInterval
    Set PFLights(idx, PFLightsCount(idx)) = obj
    PFLightsCount(idx) = PFLightsCount(idx) + 1
  Next
  ' init flasher
  InitFlasher
  ' init GI
  InitGI True
End Sub
Sub LampTimer_Timer()
  Dim chgLamp, num, chg, ii, nr, xxx, obj
  xxx = -1
    chgLamp = Controller.ChangedLamps
  If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
      Select Case chglamp(ii,0)
      Case 8,49,50,51,52,53,54,55,56
        If chgLamp(ii,1) Then
          PFLights(chgLamp(ii,0),0).Opacity = IIF((chgLamp(ii,1) = 56),400,800)
          PFLights(chgLamp(ii,0),0).Visible = True
        Else
          PFLights(chgLamp(ii,0),0).Opacity = PFLights(chgLamp(ii,0),0).Opacity / 2
          BackboardLampTimer.Enabled = True
        End If
      Case 57,58,59,60,61,62,63
        PFLights(chgLamp(ii,0),0).SetValue IIF(chgLamp(ii,1),1,0)
      Case Else
        For nr = 1 to PFLightsCount(chgLamp(ii,0))
          If Len(MouseHoleLightsOn) = 0 Then PFLights(chgLamp(ii,0),nr - 1).State = chgLamp(ii,1)
        Next
      End Select
        Next
  End If
End Sub
Sub BackboardLampTimer_Timer()
  Dim obj, allZero : allZero = True
  For Each obj In Array(f8,f49,f50,f51,f52,f53,f54,f55,f56)
    If obj.Opacity > 50 Then
      allZero = False
      obj.Opacity = obj.Opacity / 2
    Else
      obj.Visible = False
    End If
  Next
  If allZero Then BackboardLampTimer.Enabled = False
End Sub

' flasher
Const minFlasherNo = 15
Const maxFlasherNo = 32
ReDim fValue(maxFlasherNo,14) : For i = minFlasherNo To maxFlasherNo : fValue(i,0) = 0 : Next

Sub InitFlasher()
  fValue(15,4)    = Array(fLight15,fBulb15)
  fValue(15,5)    = WhiteFlasher
  fValue(25,4)    = Array(fLight25,fPlasticsLight25,fBulb25)
  fValue(25,5)    = WhiteFlasher
  fValue(26,4)    = Array(fLight26,fPlasticsLight26,fBulb26)
  fValue(26,5)    = WhiteFlasher
  fValue(27,4)    = Array(fLight27,fPlasticsLight27,fBulb27,fLight27a,fPlasticsLight27a,fBulb27a,fLight27b)
  fValue(27,5)    = WhiteFlasher
  fValue(29,4)    = Array(fLight29,fPlasticsLight29,fBulb29)
  fValue(29,5)    = WhiteFlasher
  fValue(30,4)    = Array(fLight30,fPlasticsLight30,fBulb30,fLight30a,fPlasticsLight30a,fBulb30a,fLight30b,fPlasticsLight30b,fBulb30b)
  fValue(30,5)    = WhiteFlasher
  fValue(31,4)    = Array(fLight31,fPlasticsLight31,fBulb31,fLight31a,fPlasticsLight31a,fBulb31a)
  fValue(31,5)    = WhiteFlasher
  Set fValue(32,4)  = fLight32
  fValue(32,5)    = WhiteFlasher
  ' start flasher timer
  FlasherTimer.Interval = 25
  If Not FlasherTimer.Enabled Then FlasherTimer_Timer
End Sub

Sub Flash(flasherNo, flasherValue)
  If EnableFlasher = 0 Then Exit Sub
  ' set value
  fValue(flasherNo,0) = IIF(flasherValue,1,fDecrease(fValue(flasherNo,7)))
  ' start flasher timer
  If Not FlasherTimer.Enabled Then FlasherTimer_Timer
End Sub
Sub FlasherTimer_Timer()
  Dim ii, allZero, flashx3, matdim, obj
  allZero = True
  If Not FlasherTimer.Enabled Then
    FlasherTimer.Enabled = True
  End If
  For ii = minFlasherNo To maxFlasherNo
    If (IsObject(fValue(ii,1)) Or IsObject(fValue(ii,4)) Or IsArray(fValue(ii,4))) And fValue(ii,0) >= 0 Then
      allZero = False
      If IsObject(fValue(ii,1)) Then
        If Not fValue(ii,1).Visible Then
          fValue(ii,1).Visible = True : If IsObject(fValue(ii,2)) Then fValue(ii,2).Visible = True
        End If
      End If
      ' calc values
      flashx3 = fValue(ii,0) ^ 3
      matdim  = Round(10 * fValue(ii,0))
      ' set flasher object values
      If IsObject(fValue(ii,1)) Then fValue(ii,1).Opacity = fOpacity(fValue(ii,5)) * flashx3
      If IsObject(fValue(ii,2)) Then fValue(ii,2).BlendDisableLighting = 10 * flashx3 : fValue(ii,2).Material = "domelit" & matdim : fValue(ii,2).Visible = Not (fValue(ii,0) < 0.15)
      If IsObject(fValue(ii,3)) Then fValue(ii,3).BlendDisableLighting =  flashx3
      If IsObject(fValue(ii,4)) Then fValue(ii,4).IntensityScale = flashx3 : fValue(ii,4).State = IIF(fValue(ii,4).IntensityScale<=0,0,1)
      If IsArray(fValue(ii,4)) Then
        For Each obj In fValue(ii,4) : obj.IntensityScale = flashx3 : obj.State = IIF(obj.IntensityScale<=0,0,1) : Next
      End If
      ' decrease flasher value
      If fValue(ii,0) < 1 Then fValue(ii,0) = fValue(ii,0) * fDecrease(fValue(ii,5)) - 0.01
      ' some special handling for flasher
      If Not isGIOn Then
        If ii = 31 Then
          pLeftTrap.Material = "Plastic with an image" & fMaterial(fValue(ii,0))
        ElseIf ii = 30 Then
          pRightTrap.Material = "Plastic with an image" & fMaterial(fValue(ii,0))
        End If
      End If
    End If
  Next
  If allZero Then
    FlasherTimer.Enabled = False
  End If
End Sub
Function fOpacity(fColor)
  If fColor = RedFlasher Then
    fOpacity = RedFlasherOpacity
  ElseIf fColor = BlueFlasher Then
    fOpacity = BlueFlasherOpacity
  Else
    fOpacity = WhiteFlasherOpacity
  End If
End Function
Function fDecrease(fColor)
  If fColor = RedFlasher Then
    fDecrease = RedFlasherDecrease
  ElseIf fColor = BlueFlasher Then
    fDecrease = BlueFlasherDecrease
  Else
    fDecrease = WhiteFlasherDecrease
  End If
End Function
Function fMaterial(fValue)
  If fValue > 0.8 Then
    fMaterial = ""
  ElseIf fValue > 0.6 Then
    fMaterial = " 0.7"
  ElseIf fValue > 0.4 Then
    fMaterial = " 0.5"
  ElseIf fValue > 0.2 Then
    fMaterial = " 0.2"
  Else
    fMaterial = " Dark"
  End If
End Function

Dim WhiteFlasher, WhiteFlasherOpacity, WhiteFlasherDecrease
WhiteFlasher      = 1
WhiteFlasherOpacity   = 1000
WhiteFlasherDecrease  = 0.9

Dim RedFlasher, RedFlasherOpacity, RedFlasherDecrease
RedFlasher        = 2
RedFlasherOpacity   = 1500
RedFlasherDecrease    = 0.85

Dim BlueFlasher, BlueFlasherOpacity, BlueFlasherDecrease
BlueFlasher       = 3
BlueFlasherOpacity    = 8000
BlueFlasherDecrease   = 0.85

Dim YellowInsertFlasher
YellowInsertFlasher   = 4

Sub Flash28(Enabled)
  pFlasher28.Image  = "Dome" & IIF(Enabled,"On","Off")
  pFlasher28.Material = "Orange Flashers" & IIF(Enabled,"On","")
  f28.State       = IIF(Enabled,LightStateOn,LightStateOff)
  f28a.State      = f28.State
End Sub

' general illumination
Sub ResetGI()
  InitGI False
End Sub
Sub InitGI(startUp)
  If startUp Then
    isGIOn = False
    SolGI False
  End If
  Dim coll, obj
  ' init GI overhead
  With GIOverhead
    .IntensityScale   = IIF(isGIOn,1,0)
    .State        = LightStateOn
    If GIColorMod = 1 Then
      .Color      = YellowOverhead
      .ColorFull    = YellowOverheadFull
      .Intensity    = YellowOverheadI
    ElseIf GIColorMod = 2 Then
      .Color      = RedOverhead
      .ColorFull    = RedOverheadFull
      .Intensity    = RedOverheadI
    ElseIf GIColorMod = 3 Then
      .Color      = BlueOverhead
      .ColorFull    = BlueOverheadFull
      .Intensity    = BlueOverheadI
    Else
      .Color      = WhiteOverhead
      .ColorFull    = WhiteOverheadFull
      .Intensity    = WhiteOverheadI
    End If
  End With
  ' init GI bulbs
  For Each obj In GIBulbs
    obj.IntensityScale  = IIF(isGIOn,1,0)
    obj.State     = LightStateOn
    If GIColorMod = 1 Then
      obj.Color   = YellowBulbs
      obj.ColorFull = YellowBulbsFull
      obj.Intensity   = YellowBulbsI * EnableGI
    ElseIf GIColorMod = 2 Then
      obj.Color   = RedBulbs
      obj.ColorFull = RedBulbsFull
      obj.Intensity   = RedBulbsI * EnableGI
    ElseIf GIColorMod = 3 Then
      obj.Color   = BlueBulbs
      obj.ColorFull = BlueBulbsFull
      obj.Intensity   = BlueBulbsI * EnableGI
    Else
      obj.Color   = WhiteBulbs
      obj.ColorFull   = WhiteBulbsFull
      obj.Intensity   = WhiteBulbsI * EnableGI
    End If
  Next
  ' init GI lights
  For Each obj in GILeft
    obj.IntensityScale  = IIF(isGIOn,1,0)
    obj.State     = LightStateOn
    If GIColorMod = 1 Then
      obj.Color   = Yellow
      obj.ColorFull = YellowFull
      obj.Intensity   = YellowI * EnableGI
    ElseIf GIColorMod = 2 Then
      obj.Color   = Red
      obj.ColorFull = RedFull
      obj.Intensity   = RedI * EnableGI
    ElseIf GIColorMod = 3 Then
      obj.Color   = Blue
      obj.ColorFull = BlueFull
      obj.Intensity   = BlueI * EnableGI
    Else
      obj.Color   = White
      obj.ColorFull = WhiteFull
      obj.Intensity   = WhiteI * EnableGI
    End If
  Next
  For Each obj In GIPlasticsLeft
    obj.IntensityScale  = IIF(isGIOn,1,0)
    obj.State     = LightStateOn
    If GIColorMod = 1 Then
      obj.Color   = Yellow
      obj.ColorFull = YellowFull
      obj.Intensity   = YellowPlasticI * EnableGI
    ElseIf GIColorMod = 2 Then
      obj.Color   = Red
      obj.ColorFull = RedFull
      obj.Intensity   = RedPlasticI * EnableGI
    ElseIf GIColorMod = 3 Then
      obj.Color   = Blue
      obj.ColorFull = BlueFull
      obj.Intensity   = BluePlasticI * EnableGI
    Else
      obj.Color   = White
      obj.ColorFull = WhiteFull
      obj.Intensity   = WhitePlasticI * EnableGI
    End If
  Next
  For Each obj in GIRight
    obj.IntensityScale  = IIF(isGIOn,1,0)
    obj.State     = LightStateOn
    If GIColorMod = 1 Then
      obj.Color   = Yellow
      obj.ColorFull = YellowFull
      obj.Intensity   = YellowI * EnableGI
    ElseIf GIColorMod = 2 Then
      obj.Color   = Blue
      obj.ColorFull = BlueFull
      obj.Intensity   = BlueI * EnableGI
    ElseIf GIColorMod = 3 Then
      obj.Color   = Red
      obj.ColorFull = RedFull
      obj.Intensity   = RedI * EnableGI
    Else
      obj.Color   = White
      obj.ColorFull = WhiteFull
      obj.Intensity   = WhiteI * EnableGI
    End If
  Next
  For Each obj In GIPlasticsRight
    obj.IntensityScale  = IIF(isGIOn,1,0)
    obj.State     = LightStateOn
    If GIColorMod = 1 Then
      obj.Color   = Yellow
      obj.ColorFull = YellowFull
      obj.Intensity   = YellowPlasticI * EnableGI * IIF(obj.TimerInterval=-2,0.5,1)
    ElseIf GIColorMod = 2 Then
      obj.Color   = Blue
      obj.ColorFull = BlueFull
      obj.Intensity   = BluePlasticI * EnableGI * IIF(obj.TimerInterval=-2,0.5,1)
    ElseIf GIColorMod = 3 Then
      obj.Color   = Red
      obj.ColorFull = RedFull
      obj.Intensity   = RedPlasticI * EnableGI * IIF(obj.TimerInterval=-2,0.5,1)
    Else
      obj.Color   = White
      obj.ColorFull = WhiteFull
      obj.Intensity   = WhitePlasticI * EnableGI * IIF(obj.TimerInterval=-2,0.5,1)
    End If
  Next
  ' init flasher bulbs
  For Each obj In FlasherBulbs
    obj.IntensityScale  = IIF(isGIOn,1,0)
    obj.State     = LightStateOn
    If GIColorMod = 1 Then
      obj.Color   = YellowBulbs
      obj.ColorFull = YellowBulbsFull
      obj.Intensity   = YellowBulbsI * EnableGI
    ElseIf GIColorMod = 2 Then
      obj.Color   = RedBulbs
      obj.ColorFull = RedBulbsFull
      obj.Intensity   = RedBulbsI * EnableGI
    ElseIf GIColorMod = 3 Then
      obj.Color   = BlueBulbs
      obj.ColorFull = BlueBulbsFull
      obj.Intensity   = BlueBulbsI * EnableGI
    Else
      obj.Color   = WhiteBulbs
      obj.ColorFull   = WhiteBulbsFull
      obj.Intensity   = WhiteBulbsI * EnableGI
    End If
  Next
End Sub

Dim GIDir : GIDir = 0
Dim GIStep : GIStep = 0
Sub SolGI(IsOff)
  If EnableGI = 0 And Not isGIOn Then Exit Sub
  If isGIOn <> Not IsOff Then
    isGIOn = Not IsOff
    If isGIOn Then
      ' GI goes on
      PlaySoundAtVol "fx_relay_on", fLight32, 2
      GIDir = 1 : GITimer_Timer
      DOF 101, DOFOn
    Else
      ' GI goes off
      PlaySoundAtVol "fx_relay_off", fLight32, 2
      GIDir = -1 : GITimer_Timer
      DOF 101, DOFOff
    End If
  End If
End Sub
Sub GITimer_Timer()
  If Not GITimer.Enabled Then GITimer.Enabled = True
  GIStep      = GIStep + GIDir
  ' set opacity of the shadow overlays
  fGIOff.Opacity  = (ShadowOpacityGIOff / 4) * (4 - GIStep)
  fGIOn.Opacity   = (ShadowOpacityGIOn / 4) * GIStep
  ' set GI illumination
  Dim coll, obj
  For Each coll In Array(GILeft,GIPlasticsLeft,GIRight,GIPlasticsRight,GIBulbs)
    For Each obj In coll
      If obj.TimerInterval <> -1 Then obj.IntensityScale = GIStep/4
    Next
  Next
  UpdateGILightsUnderCheeseLoop
  GIOverhead.IntensityScale = GIStep/4
  For Each obj In GIBumperLights
    obj.IntensityScale  = GIStep/4 * 0.6
    obj.State       = IIF(GIStep <= 0, LightstateOff, LightStateOn)
  Next
  ' targets, posts, pegs etc
  If GIStep = 4 Then
    For Each obj In GIYellowTargets : obj.Material = "TargetYellow" : Next : sw36.Material = "Plastic Yellow1"
    For Each obj In GIWhiteTargets : obj.Material = "TargetWhite" : Next
    For Each obj In GIYellowPosts : obj.Material = "Plastic White" : Next
    For Each obj In GIWhitePosts : obj.Material = "Plastic White" : Next
    For Each obj In GIPlasticPegs : obj.Material = "TransparentPlasticWhite2" : Next
    For Each obj In GIRubbers : obj.Material = "Rubber White" : Next
    For Each obj In GITopLanes : obj.Material = "Plastic Blue" : Next
    For Each obj In GIWireTrigger: obj.Material = "Metal0.8" : Next
    For Each obj In GILocknuts: obj.Material = "Metal0.8" : Next
    For Each obj In GIScrews: obj.Material = "Metal0.8" : Next
    For Each obj In GIGates: obj.Material = "Metal Wires" : Next
    For Each obj In GIBumpers : obj.BaseMaterial = "Plastic Yellow" : obj.SkirtMaterial = "Plastic Yellow" : Next
    For Each obj In GIBumperTops : obj.Material = "BumperTopYellow" : Next
    For Each obj In GIRampDecals : obj.Material = "Plastics Light Light Dark" : Next
    For Each obj In Array(RampSideDecalLeft,RampSideDecalRight) : obj.SideMaterial = "Plastics Light Light Dark" : Next
    For Each obj In Array(RampDecalYow1,RampDecalYow2) : obj.Material = "Plastics Light Light Dark" : Next
    For Each obj In Array(RampStopper1,RampStopper2,RampStopper3) : obj.Material = "Plastics" : Next
    For Each obj In Array(pLeftTrap,pRightTrap,pCheeseLoopHut,pGateTopLeft,pGateTopRight,pLeftFlipperBally,pRightFlipperBally) : obj.Material = "Plastic with an image" : Next
    For Each obj In Array(pDiverterAxisA,pDiverterAxisB,rBBF1,rBBF2,rBBF3) : obj.Material = "Metal Chrome S34" : Next
    For Each obj In Array(rMetalWall1,rMetalWall2,rMetalWall3,rMetalWall4,rMetalWall5,rMetalWall6,rMetalWall7,rMetalWall8,rMetalWall9) : obj.Material = "Metal Chrome S34" : Next
    For Each obj In Array(pCheeseLoopRamp,pGuideLaneLeft,pGuideLaneRight,pRampTrigger37,pRampTrigger46) : obj.Material = "Metal S34" : Next
    For Each obj In Array(pMetalGate,pDiverterArm,pDiverterFixer) : obj.Material = "Metal Diverter" : Next
    For Each obj In Array(pCheeseLoopGate,pCheeseLoopFixer,pCheeseLoopNutC,pCheeseLoopNutD) : obj.Material = "Metal" : Next
    pLevelPlate.Material = "Metal0.8"
    If FlipperColorMod = 2 Then
      pLeftFlipperBat.Material = "flipperbatyellow" : pRightFlipperBat.Material = "flipperbatyellow"
    ElseIf FlipperColorMod = 3 Then
      pLeftFlipperBat.Material = "flipperbatred" : pRightFlipperBat.Material = "flipperbatred"
    ElseIf FlipperColorMod = 4 Then
      pLeftFlipperBat.Material = "flipperbatblue" : pRightFlipperBat.Material = "flipperbatblue"
    End If
    wBackboard.IsDropped = False : wBackboardDark.IsDropped = True
  ElseIf GIStep = 0 Then
    For Each obj In GIYellowTargets : obj.Material = "TargetYellow Dark" : Next : sw36.Material = "Plastic Yellow Dark"
    For Each obj In GIWhiteTargets : obj.Material = "TargetWhite Dark" : Next
    For Each obj In GIYellowPosts : obj.Material = "Plastic White Dark" : Next
    For Each obj In GIWhitePosts : obj.Material = "Plastic White Dark" : Next
    For Each obj In GIPlasticPegs : obj.Material = "TransparentPlasticWhite2 Dark" : Next
    For Each obj In GIRubbers : obj.Material = "Rubber White Dark" : Next
    For Each obj In GIWireTrigger: obj.Material = "Metal0.8 Dark" : Next
    For Each obj In GILocknuts: obj.Material = "Metal0.8 Dark" : Next
    For Each obj In GIScrews: obj.Material = "Metal0.8 Dark" : Next
    For Each obj In GITopLanes : obj.Material = "Plastic Blue Dark" : Next
    For Each obj In GIGates: obj.Material = "Metal Wires Dark" : Next
    For Each obj In GIBumpers : obj.BaseMaterial = "Plastic Yellow Dark" : obj.SkirtMaterial = "Plastic Yellow Dark" : Next
    For Each obj In GIBumperTops : obj.Material = "BumperTopYellow Dark" : Next
    For Each obj In GIRampDecals : obj.Material = "Plastics Light Dark" : Next
    For Each obj In Array(RampSideDecalLeft,RampSideDecalRight) : obj.SideMaterial = "Plastics Light Dark" : Next
    For Each obj In Array(RampDecalYow1,RampDecalYow2) : obj.Material = "Plastics Dark" : Next
    For Each obj In Array(RampStopper1,RampStopper2,RampStopper3) : obj.Material = "Plastics Light Dark" : Next
    For Each obj In Array(pLeftTrap,pRightTrap,pCheeseLoopHut,pGateTopLeft,pGateTopRight,pLeftFlipperBally,pRightFlipperBally) : obj.Material = "Plastic with an image Dark" : Next
    For Each obj In Array(pDiverterAxisA,pDiverterAxisB,rBBF1,rBBF2,rBBF3) : obj.Material = "Metal Chrome S34 Dark" : Next
    For Each obj In Array(rMetalWall1,rMetalWall2,rMetalWall3,rMetalWall4,rMetalWall5,rMetalWall6,rMetalWall7,rMetalWall8,rMetalWall9) : obj.Material = "Metal Chrome S34 Dark" : Next
    For Each obj In Array(pCheeseLoopRamp,pGuideLaneLeft,pGuideLaneRight,pRampTrigger37,pRampTrigger46) : obj.Material = "Metal S34 Dark" : Next
    For Each obj In Array(pMetalGate,pDiverterArm,pDiverterFixer) : obj.Material = "Metal Diverter Dark" : Next
    For Each obj In Array(pCheeseLoopGate,pCheeseLoopFixer,pCheeseLoopNutC,pCheeseLoopNutD) : obj.Material = "Metal Dark" : Next
    pLevelPlate.Material = "Metal0.8 Dark"
    If FlipperColorMod = 2 Then
      pLeftFlipperBat.Material = "flipperbatyellow Dark" : pRightFlipperBat.Material = "flipperbatyellow Dark"
    ElseIf FlipperColorMod = 3 Then
      pLeftFlipperBat.Material = "flipperbatred Dark" : pRightFlipperBat.Material = "flipperbatred Dark"
    ElseIf FlipperColorMod = 4 Then
      pLeftFlipperBat.Material = "flipperbatblue Dark" : pRightFlipperBat.Material = "flipperbatblue Dark"
    End If
    wBackboard.IsDropped = True : wBackboardDark.IsDropped = False
  End If
  ' ramps
  pRampLeft.Material = "rampsGI" & (GIStep * 2)
  pRampRight.Material = "rampsGI" & (GIStep * 2)
  pRampCenter.Material = "rampsGI" & (GIStep * 2)
  ' GI on/off goes in 4 steps so maybe stop timer
  If (GIDir = 1 And GIStep = 4) Or (GIDir = -1 And GIStep = 0) Then
    GITimer.Enabled = False
  End If
End Sub
Sub UpdateGILightsUnderCheeseLoop()
  GILightL8.IntensityScale  = IIF(isMotorbankUp, GIStep/4, 0)
  GILightL8a.IntensityScale   = IIF(isMotorbankUp, 0, GIStep/4)
  GILightR8.IntensityScale  = IIF(isMotorbankUp, GIStep/4, 0)
  GILightR8a.IntensityScale   = IIF(isMotorbankUp, 0, GIStep/4)
End Sub


' *********************************************************************
' colors
' *********************************************************************
Dim White, WhiteFull, WhiteI, WhiteP, WhitePlastic, WhitePlasticFull, WhitePlasticI, WhiteBumper, WhiteBumperFull, WhiteBumperI, WhiteBulbs, WhiteBulbsFull, WhiteBulbsI, WhiteOverheadFull, WhiteOverhead, WhiteOverheadI
WhiteFull = rgb(255,255,255)
White = rgb(255,255,180)
WhiteI = 15
WhitePlasticFull = rgb(255,255,180)
WhitePlastic = rgb(255,255,180)
WhitePlasticI = 25
WhiteBumperFull = rgb(255,255,180)
WhiteBumper = rgb(255,255,180)
WhiteBumperI = 1
WhiteBulbsFull = rgb(255,255,180)
WhiteBulbs = rgb(255,255,180)
WhiteBulbsI = 100 * ShadowOpacityGIOff
WhiteOverheadFull = rgb(255,255,180)
WhiteOverhead = rgb(255,255,180)
WhiteOverheadI = .25

Dim Yellow, YellowFull, YellowI, YellowPlastic, YellowPlasticFull, YellowPlasticI, YellowBumper, YellowBumperFull, YellowBumperI, YellowBulbs, YellowBulbsFull, YellowBulbsI, YellowOverheadFull, YellowOverhead, YellowOverheadI
YellowFull = rgb(255,255,0)
Yellow = rgb(255,255,0)
YellowI = 5
YellowPlasticFull = rgb(255,255,0)
YellowPlastic = rgb(255,255,0)
YellowPlasticI = 40
YellowBumperFull = rgb(255,255,0)
YellowBumper = rgb(255,255,0)
YellowBumperI = 1
YellowBulbsFull = rgb(255,255,0)
YellowBulbs = rgb(255,255,0)
YellowBulbsI = 250 * ShadowOpacityGIOff
YellowOverheadFull = rgb(255,255,10)
YellowOverhead = rgb(255,255,10)
YellowOverheadI = 0.5

Dim Red, RedFull, RedI, RedPlastic, RedPlasticFull, RedPlasticI, RedBumper, RedBumperFull, RedBumperI, RedBulbs, RedBulbsFull, RedBulbsI, RedOverheadFull, RedOverhead, RedOverheadI
RedFull = rgb(255,75,75)
Red = rgb(255,75,75)
RedI = 15
RedPlasticFull = rgb(255,75,75)
RedPlastic = rgb(255,75,75)
RedPlasticI = 20
RedBumperFull = rgb(255,0,0)
RedBumper = rgb(255,0,0)
RedBumperI = 2
RedBulbsFull = rgb(255,75,75)
RedBulbs = rgb(255,75,75)
RedBulbsI = 250 * ShadowOpacityGIOff
RedOverheadFull = rgb(255,10,10)
RedOverhead = rgb(255,10,10)
RedOverheadI = 1

Dim Blue, BlueFull, BlueI, BluePlastic, BluePlasticFull, BluePlasticI, BlueBumper, BlueBumperFull, BlueBumperI, BlueBulbs, BlueBulbsFull, BlueBulbsI,  BlueOverheadFull, BlueOverhead, BlueOverheadI
BlueFull = rgb(75,75,255)
Blue = rgb(75,75,255)
BlueI = 50
BluePlasticFull = rgb(75,75,255)
BluePlastic = rgb(75,75,255)
BluePlasticI = 20
BlueBumperFull = rgb(0,0,255)
BlueBumper = rgb(0,0,255)
BlueBumperI = 1
BlueBulbsFull = rgb(75,75,255)
BlueBulbs = rgb(75,75,255)
BlueBulbsI = 125 * ShadowOpacityGIOff
BlueOverheadFull = rgb(10,10,255)
BlueOverhead = rgb(10,10,255)
BlueOverheadI = .8


' *********************************************************************
' digital display
' *********************************************************************
Dim Digits(32)
Digits(0)=Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
Digits(1)=Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
Digits(2)=Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
Digits(3)=Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
Digits(4)=Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
Digits(5)=Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
Digits(6)=Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
Digits(7)=Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
Digits(8)=Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
Digits(9)=Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
Digits(10)=Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
Digits(11)=Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
Digits(12)=Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
Digits(13)=Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
Digits(14)=Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
Digits(15)=Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)
Digits(16)=Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
Digits(17)=Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
Digits(18)=Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
Digits(19)=Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
Digits(20)=Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
Digits(21)=Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
Digits(22)=Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
Digits(23)=Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
Digits(24)=Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
Digits(25)=Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
Digits(26)=Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
Digits(27)=Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
Digits(28)=Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
Digits(29)=Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
Digits(30)=Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
Digits(31)=Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)

Sub DisplayTimer_Timer()
    Dim chgLED, ii, num, chg, stat, obj
  chgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(chgLED) Then
    If DesktopMode And Not noStandardDisplay Then
      For ii = 0 To UBound(chgLED)
        num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
        if (num < 32) then
          For Each obj In Digits(num)
            If chg And 1 Then obj.State = stat And 1
            chg = chg\2 : stat = stat\2
          Next
        End If
      Next
    End If
    End If
End Sub


' *********************************************************************
' some special physics behaviour
' *********************************************************************
' target or rubber post is hit so let the ball jump a bit
Sub DropTargetHit()
  DropTargetSound
  TargetHit
End Sub

Sub StandupTargetHit()
  StandUpTargetSound
  TargetHit
End Sub

Sub TargetHit()
    ActiveBall.VelZ = ActiveBall.VelZ * (0.5 + (Rnd()*LetTheBallJump + Rnd()*LetTheBallJump + 1) / 6)
End Sub

Sub RubberPostHit()
  ActiveBall.VelZ = ActiveBall.VelZ * (0.9 + (Rnd()*(LetTheBallJump-1) + Rnd()*(LetTheBallJump-1) + 1) / 6)
End Sub

Sub RubberRingHit()
  ActiveBall.VelZ = ActiveBall.VelZ * (0.8 + (Rnd()*(LetTheBallJump-1) + 1) / 6)
End Sub

Sub trSlowdownOrbit_Hit()
  SlowDownBall ActiveBall,2
End Sub

Sub SlowDownBall(aBall, mode)
  If aBall.VelY > 0 Then
    If mode = 1 And BallVel(aBall) >= 8 Then
      With aBall : .VelX = .VelX / 2 : .VelY = .VelY / 2 : End With
    ElseIf mode = 2 Then
      With aBall : .VelX = .VelX / 4 : .VelY = .VelY / 4 : End With
    End If
  End If
End Sub


' *********************************************************************
' sound stuff
' *********************************************************************
Sub RollOverSound()
  PlaySoundAtVolPitch SoundFX("fx_rollover",DOFContactors), ActiveBall, 0.02, .25
End Sub
Sub DropTargetSound()
  PlaySoundAtVolPitch SoundFX("fx_droptarget",DOFTargets), ActiveBall, 2, .25
End Sub
Sub StandUpTargetSound()
  PlaySoundAtVolPitch SoundFX("fx_target",DOFTargets), ActiveBall, 2, .25
End Sub
Sub GateSound()
  PlaySoundAtVolPitch SoundFX("fx_gate",DOFContactors), ActiveBall, 0.02, .25
End Sub

Sub swRightRampStopper_Hit()
  PlaySoundAtBallVol "fx_balldrop" & Int(Rnd()*3), 0.2
End Sub


' *********************************************************************
' Supporting Surround Sound Feedback (SSF) functions
' *********************************************************************
' set position as table object (Use object or light but NOT wall) and Vol to 1
Sub PlaySoundAt(sound, tableobj)
  PlaySound sound, 1, 1, AudioPan(tableobj), 0, 0, 0, 1, AudioFade(tableobj)
End Sub
' set position as table object and Vol manually.
Sub PlaySoundAtVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0, 0, 0, 1, AudioFade(tableobj)
End Sub
' set position as table object and Vol + RndPitch manually
Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
End Sub

' set all as per ball position & speed.
Sub PlaySoundAtBall(sound)
  PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub
' set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3
Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub
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


' *********************************************************************
' Supporting Ball & Sound Functions
' *********************************************************************
Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 200)
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
    VolZ = Csng(BallVelZ(ball) ^2 / 2000)
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table.
    Dim tmp : tmp = tableobj.x * 2 / MousinAround.Width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function AudioFade(ball)
    Dim tmp : tmp = ball.y * 2 / MousinAround.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
    BallVelZ = INT((ball.VelZ) * -1 )
End Function


' *************************************************************
' ball shadow and flipper primitives outfit
' *************************************************************
Const tnob = 5 ' total number of balls

Dim BallShadow : BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5)

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
      If BOT(ii).X < MousinAround.Width/2 Then
        BallShadow(ii).X = ((BOT(ii).X) - (Ballsize/6) + ((BOT(ii).X - (MousinAround.Width/2))/7)) + 6
      Else
        BallShadow(ii).X = ((BOT(ii).X) + (Ballsize/6) + ((BOT(ii).X - (MousinAround.Width/2))/7)) - 6
      End If
      BallShadow(ii).Y = BOT(ii).Y + 12
      BallShadow(ii).Visible = (BOT(ii).Z > 20)
    Next
  End If

  ' maybe move primitive flippers
  If FlipperColorMod > 1 Then
    pLeftFlipperBat.ObjRotZ  = LeftFlipper.CurrentAngle + 1
    pRightFlipperBat.ObjRotZ = RightFlipper.CurrentAngle + 1
  ElseIf FlipperColorMod = 1 Then
    pLeftFlipperBally.ObjRotZ  = LeftFlipper.CurrentAngle - 90
    pRightFlipperBally.ObjRotZ = RightFlipper.CurrentAngle + 90
  End If
  pLeftFlipperShadow.ObjRotZ  = LeftFlipper.CurrentAngle + 1
  pRightFlipperShadow.ObjRotZ = RightFlipper.CurrentAngle + 1
End Sub
'
Sub ResetFlippers()
  InitFlippers
End Sub
Sub InitFlippers()
  ' initialize flipper colors
  pLeftFlipperBat.Visible   = (FlipperColorMod > 1)
  pRightFlipperBat.Visible  = pLeftFlipperBat.Visible
  pLeftFlipperBally.Visible   = (FlipperColorMod = 1)
  pRightFlipperBally.Visible  = pLeftFlipperBally.Visible
  LeftFlipper.Visible     = (Not pLeftFlipperBat.Visible And Not pLeftFlipperBally.Visible)
  RightFlipper.Visible    = LeftFlipper.Visible
  Select Case FlipperColorMod
  Case 2
    pLeftFlipperBat.Image     = "flipperbatyellow"
    pLeftFlipperBat.Material    = "flipperbatyellow"
    pRightFlipperBat.Image    = "flipperbatyellow"
    pRightFlipperBat.Material   = "flipperbatyellow"
  Case 3
    pLeftFlipperBat.Image     = "flipperbatred"
    pLeftFlipperBat.Material    = "flipperbatred"
    pRightFlipperBat.Image    = "flipperbatred"
    pRightFlipperBat.Material   = "flipperbatred"
  Case 4
    pLeftFlipperBat.Image     = "flipperbatblue"
    pLeftFlipperBat.Material  = "flipperbatblue"
    pRightFlipperBat.Image    = "flipperbatblue"
    pRightFlipperBat.Material = "flipperbatblue"
  End Select
End Sub


' *************************************************************
'      rolling and realtime sounds
' *************************************************************
ReDim rolling(tnob) : For i = 0 to tnob : rolling(i) = False : Next
Dim isBallOnWireRamp : isBallOnWireRamp = False

Sub RollingSoundTimer_Timer()
    Dim BOT, b
    BOT = GetBalls

  ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound "fx_ballrolling" & b
        StopSound "fx_plasticrolling" & b
        StopSound "fx_metalrolling" & b
    Next

  ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
    If BallVel(BOT(b) ) > 1 Then
            rolling(b) = True
      If isBallOnWireRamp Then
        ' ball on wire ramp
        StopSound "fx_ballrolling" & b
        StopSound "fx_plasticrolling" & b
        PlaySound "fx_metalrolling" & b, -1, Vol(BOT(b)) / 2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
      ElseIf BOT(b).Z > 30 Then
        ' ball on plastic ramp
        StopSound "fx_ballrolling" & b
        StopSound "fx_metalrolling" & b
        PlaySound "fx_plasticrolling" & b, -1, Vol(BOT(b)) / 2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
      Else
        ' ball on playfield
        StopSound "fx_plasticrolling" & b
        StopSound "fx_metalrolling" & b
        PlaySound "fx_ballrolling" & b, -1, Vol(BOT(b)) / 2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
      End If
    Else
      If rolling(b) Then
                StopSound "fx_ballrolling" & b
        StopSound "fx_plasticrolling" & b
        StopSound "fx_metalrolling" & b
                rolling(b) = False
            End If
    End If

    '   ball drop sounds matching the adjusted height params but not the way down the ramps
    If BOT(b).VelZ < -2 And BOT(b).Z < 55 And BOT(b).Z > 27 then 'And Not InRect(BOT(b).X, BOT(b).Y, 610,320, 740,320, 740,550, 610,550) And Not InRect(BOT(b).X, BOT(b).Y, 180,400, 230,400, 230, 550, 180,550) Then
      PlaySound "fx_balldrop" & Int(Rnd()*3), 0, ABS(BOT(b).VelZ)/17*5, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    End If

    Next
End Sub


' *********************************************************************
' more realtime sounds
' *********************************************************************
' ball collision sound
Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound "fx_collide", 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' rubber hit sounds
Sub RubberWalls_Hit(idx)
    PlaySoundAtBallVol "fx_rubber_hit_" & Int(Rnd*3)+1, 10
  RubberRingHit
End Sub
Sub RubberPosts_Hit(idx)
    PlaySoundAtBallVol "fx_rubber_hit_" & Int(Rnd*3)+1, 10
  RubberPostHit
End Sub

' metal hit sounds
Sub MetalWalls_Hit(idx)
  PlaySoundAtBallAbsVol "fx_metalhit" & Int(Rnd*3), Minimum(Vol(ActiveBall),0.5)
End Sub

' gates sound
Sub Gates_Hit(idx)
  GateSound
End Sub

' sound at ramp rubber at the diverter
Sub RampRubber_Hit()
  PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3)+1, 10
End Sub


' *********************************************************************
' some more general methods
' *********************************************************************
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

Function InCircle(pX, pY, centerX, centerY, radius)
  Dim route
  route = Sqr((pX - centerX) ^ 2 + (pY - centerY) ^ 2)
  InCircle = (route < radius)
End Function

Function Minimum(val1, val2)
  Minimum = IIF(val1<val2,val2,val1)
End Function

Const fileName = "C:\Visual Pinball\User\Log.txt"
Sub WriteLine(text)
  Dim fso, fil
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set fil = fso.OpenTextFile(fileName, 8, true)
  fil.WriteLine Date & " " & Time & ": " & text
  fil.Close
  Set fso = Nothing
End Sub

Function IIF(bool, obj1, obj2)
  If bool Then
    IIF = obj1
  Else
    IIF = obj2
  End If
End Function










