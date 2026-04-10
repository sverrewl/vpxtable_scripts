Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

Const cGameName="sshtl_l7",UseSolenoids=2,UseLamps=0,UseGI=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff"

LoadVPM "01560000", "S11.vbs", 3.36

'If you want to play with liberal operator settings you can change this (but some of the visuals will be wonky):
Tooeasy=0

'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.8

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
End if

Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

dim ballsize,ballmass,tooeasy

ballsize=1
ballmass=1

Dim NullFader : set NullFader = new NullFadingObject
Dim FadeLights : Set FadeLights = New LampFader

'**********************************************************************************************************
'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsLKick, bsRKick, DTBank, DTBankT

Sub Table1_Init
   SetLocale(1033)
   PFGI(False)
  Stopper.Collidable=0
  StopperDown.Collidable=1
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Space Shuttle"&chr(13)&"bord"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
          .hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

  vpmNudge.TiltSwitch=1
    vpmNudge.Sensitivity=3
    vpmNudge.TiltObj=Array(sw25,sw26,sw27,LeftSlingshot,RightSlingshot)

           Set bsTrough = New cvpmBallStack
               bsTrough.InitSw 9,12,11,10,0,0,0,0
             bsTrough.InitKick BallRelease, 160, 8
               RandomSoundBallRelease sw36
               bsTrough.Balls = 3

      Set bsLKick = New cvpmBallStack
          bsLKick.InitSaucer sw16,16, 185, 18
        bsLKick.KickForceVar = 2
        bsLKick.KickAngleVar = 2
        SoundSaucerKick 1, sw16

          Set bsRKick = New cvpmBallStack
              bsRKick.InitSaucer sw24,24, 172, 18
            bsRKick.KickForceVar = 2
            bsRKick.KickAngleVar = 2
              SoundSaucerKick 1, sw24
End Sub

Sub Table1_Exit(): Controller.Stop:End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:playsoundat"plungerpull", plunger
  if KeyCode = LeftTiltKey Then Nudge 90, 4:SoundNudgeLeft()
  if KeyCode = RightTiltKey Then Nudge 270, 4:SoundNudgeRight()
  if KeyCode = CenterTiltKey Then Nudge 0, 4:SoundNudgeCenter()
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
                Select Case Int(rnd*3)
                        Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
                        Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
                        Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
                End Select
     End If
     If KeyCode = PlungerKey Then Plunger.Pullback:SoundPlungerPull()
     if keycode=StartGameKey then soundStartButton()
    ' Manual Ball Control
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

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If KeyCode = PlungerKey Then
          Plunger.Fire
          If BIPL = 1 Then
              SoundPlungerReleaseBall()                        'Plunger release sound when there is a ball in shooter lane
          Else
              SoundPlungerReleaseNoBall()                        'Plunger release sound when there is no ball in shooter lane
          End If
  End If
    'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If
End Sub

dim BIPL

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1) = "bsTrough.SolIn"
SolCallback(2) = "bsTrough.SolOut"
SolCallback(3) = "bsLKick.solout"
SolCallback(4) = "bsRKick.SolOut"
SolCallback(5) = "dropupcenter"
SolCallback(6) = "dropupleft"
SolCallback(7) = "PostUp"
SolCallback(8) = "PostDown"
SolCallback(9) = "SetLamp 101,"
SolCallback(10) = "SetLamp 102,"
SolCallback(11) = "PFGI"
SolCallback(13) = "Diverter"
SolCallback(14) = "SetLamp 103,"
solcallback(15) = "SolKnocker" 'this is actually a bell

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'******************************************************
'           FLIPPERS
'******************************************************
Const ReflipAngle = 10

Sub SolLFlipper(Enabled)
  If Enabled Then
                LF.Fire  'leftflipper.rotatetoend
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
                RF.Fire 'rightflipper.rotatetoend
                If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
                        RandomSoundReflipUpRight RightFlipper
                Else
                        SoundFlipperUpAttackRight RightFlipper
                        RandomSoundFlipperUpRight RightFlipper
                End If
  Else
                RightFlipper.RotateToStart
                If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
                        RandomSoundFlipperDownRight RightFlipper
                End If
                FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub


' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
' Sub LeftFlipper_Collide(parm)
'         LeftFlipperCollide parm
' End Sub
'
' Sub RightFlipper_Collide(parm)
'         RightFlipperCollide parm
' End Sub
'******************************************************
'                        FLIPPER TRICKS
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
                                        BOT(b).velx = BOT(b).velx / 1.7
                                        BOT(b).vely = BOT(b).vely - 1
                                end If
                        Next
                End If
        Else
                If Flipper1.currentangle <> EndAngle1 then
                        EOSNudge1 = 0
                end if
        End If
End Sub

'*****************
' Maths
'*****************
Const PI = 3.1415927

Function dSin(degrees)
        dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
        dcos = cos(degrees * Pi/180)
End Function

'*************************************************
' Check ball distance from Flipper for Rem
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

' Used for drop targets and stand up targets
Function Atn2(dy, dx)
        dim pi
        pi = 4*Atn(1)

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
' End - Check ball distance from Flipper for Rem
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
Const EOSTnew = 1
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.045

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
                Dim BOT, b
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
        End If
End Sub


'**************************************************
'        Flipper Collision Subs
'NOTE: COpy and overwrite collision sound from original collision subs over
'RandomSoundFlipper()' below
'**************************************************'

Sub LeftFlipper_Collide(parm)
    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
'    RandomSoundFlipper() 'Remove this line if Fleep is integrated
    LeftFlipperCollide parm   'This is the Fleep code
End Sub

Sub RightFlipper_Collide(parm)
    CheckLiveCatch Activeball, RightFlipper, RFCount, parm
'    RandomSoundFlipper() 'Remove this line if Fleep is integrated
    RightFlipperCollide parm  'This is the Fleep code
End Sub

' Diverter

Sub Diverter(Enabled)
  If Enabled then
    bottomgate.rotatetoend
  Else
    bottomgate.rotatetostart
  end if
End Sub

Sub PostUp(Enabled)
    If Enabled Then
    Stopper.Collidable=1
    StopperDown.Collidable=0
    Playsoundat "dropup", centerpost_prim
        CenterPost_prim.TransZ = 24
    End If
End Sub

Sub PostDown(Enabled)
    If Enabled Then
    Stopper.Collidable=0
    StopperDown.Collidable=1
    Playsoundat "soloff", centerpost_prim
        CenterPost_prim.Transz = 1
    End If
End Sub

dim GiIsOff

'Playfield GI
Sub PFGI(Enabled)
  If Enabled Then  ' GI state inverted, enabled = OFF
    dim xx
    For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"
    GIIsOff=true
    outerleft.image="outerleftoff"
    outerright.image="outerrightoff"
    spaceman.image="outerleftoff"
    outerback.image="outerbackoff"
    outerbackmetal.image="outerbackoff"
    outerbackmetal1.image="outermetalsoff"
'   outerbackmetal2.image="backplasticsoff"
    outerbackplastic.image="backplasticsoff"
    bumpera1.image="bumps1off"
    bumperb1.image="bumps1off"
    bumperc1.image="bumps1off"
    bumpera2.image="bumps2off"
    bumperb2.image="bumps2off"
    bumperc2.image="bumps2off"
    bumpera3.image="bumps3off"
    bumperb3.image="bumps3off"
    bumperc3.image="bumps3off"
    rampplate.image="rampsoff"
    ramp.image="rampsoff"
    plastics1.image="plastics1off"
    plastics2.image="plastics3off"
    sw17.image="sw17off"
    sw18.image="sw18off"
    sw19.image="sw19off"
    psw20.image="sw20off"
    sw21.image="sw21off"
    sw22.image="sw22off"
    sw23.image="sw23off"
    psw33.image="sw33off"
    psw34.image="sw34off"
    psw35.image="sw35off"
    sw38.image="ramptargetoff"
    rampposts_physics.image="rubrampoff"
    posts_spinner.image="metals_underrampoff"
    metals_underramp.image="metals_underrampoff"
    leftkick_metal.image="kickleftoff"
    leftkickplastic.image="kickleftoff"
    rightkick_metal.image="kickrightoff"
    rightkickplastic.image="kickrightoff"
    rightlane_posts.image="lanerightoff"
    rightlane_posts_lib.image="lanerightoff"
    rightlane_metal.image="lanerightoff"
    leftlane_posts.image="laneleftoff"
    leftlane_posts_lib.image="laneleftoff"
    leftlane_metal.image="laneleftoff"
    shuttle.image="shuttleoff"
    rubbers1.image="bakeryoff"
    rubbers1_lib.image="bakeryoff"
    metalstravaganza.image="bakeryoff"
    acorns.image="bakeryoff"
    acorns_lib.image="bakeryoff"
    lsling.image="bakeryoff"
    rsling.image="bakeryoff"
    rubberposts.image="rubsoff"
    sw37.image="spin1off"
    rampflaps.blenddisablelighting=0.5
  Else
    For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"
    GIIsOff=false
    outerleft.image="outerleft"
    outerright.image="outerright"
    spaceman.image="outerleft"
    outerback.image="outerback"
    outerbackmetal.image="outerback"
    outerbackmetal1.image="outermetals"
'   outerbackmetal2.image="backplastics"
    outerbackplastic.image="backplastics"
    bumpera1.image="bumps1"
    bumperb1.image="bumps1"
    bumperc1.image="bumps1"
    bumpera2.image="bumps2"
    bumperb2.image="bumps2"
    bumperc2.image="bumps2"
    bumpera3.image="bumps3"
    bumperb3.image="bumps3"
    bumperc3.image="bumps3"
    rampplate.image="ramps"
    ramp.image="ramps"
    plastics1.image="plastics1"
    sw17.image="sw17"
    sw18.image="sw18"
    sw19.image="sw19"
    psw20.image="sw20"
    sw21.image="sw21"
    sw22.image="sw22"
    sw23.image="sw23"
    psw33.image="sw33"
    psw34.image="sw34"
    psw35.image="sw35"
    sw38.image="ramptarget"
    rampposts_physics.image="rubramp"
    posts_spinner.image="metals_underramp"
    metals_underramp.image="metals_underramp"
    leftkick_metal.image="kickleft"
    leftkickplastic.image="kickleft"
    rightkick_metal.image="kickright"
    rightkickplastic.image="kickright"
    rightlane_posts.image="laneright"
    rightlane_posts_lib.image="laneright"
    rightlane_metal.image="laneright"
    leftlane_posts.image="laneleft"
    leftlane_posts_lib.image="laneleft"
    leftlane_metal.image="laneleft"
    shuttle.image="shuttle"
    rubbers1.image="bakery"
    rubbers1_lib.image="bakery"
    metalstravaganza.image="bakery"
    acorns.image="bakery"
    acorns_lib.image="bakery"
    lsling.image="bakery"
    rsling.image="bakery"
    rubberposts.image="rubs"
    sw37.image="spin1"
    rampflaps.blenddisablelighting=1
  End If
  ' Update lamps that may swap textures when GI changes here
  SetLamp 110, Not Enabled
  SetLamp 111, Enabled
  UpdateLamp25
  UpdateLamp26
  UpdateLamp27
End Sub

  if tooeasy=1 Then
    acorns.visible=0
    acorns_lib.visible=1
    rubbers1.visible=0
    rubbers1_lib.visible=1
    libwall001.collidable=1
    libwall002.collidable=1
    libwall003.collidable=1
    leftlane_posts.visible=0
    leftlane_posts_lib.visible=1
    rightlane_posts.visible=0
    rightlane_posts_lib.visible=1
  Else
    acorns.visible=1
    acorns_lib.visible=0
    rubbers1.visible=1
    rubbers1_lib.visible=0
    libwall001.collidable=0
    libwall002.collidable=0
    libwall003.collidable=0
    leftlane_posts.visible=1
    leftlane_posts_lib.visible=0
    rightlane_posts.visible=1
    rightlane_posts_lib.visible=0
  End if

Sub Wall26_hit
  RandomSoundFlipperBallGuide
End Sub

Sub Wall18_hit
  RandomSoundFlipperBallGuide
End Sub

Sub bottomgate_hit
  RandomSoundFlipperBallGuide
End Sub

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : RandomSoundDrain drain : End Sub
Sub sw16_Hit():bsLKick.AddBall 0:Controller.Switch(16)=1:SoundSaucerLock:End Sub
Sub sw24_Hit():bsRKick.AddBall 0:Controller.Switch(24)=1:SoundSaucerLock:End Sub

'Drop Targets
Sub Sw20_Hit:DTHit 20:End Sub
Sub Sw33_Hit:DTHit 33:End Sub
Sub Sw34_Hit:DTHit 34:End Sub
Sub Sw35_Hit:DTHit 35:End Sub

Sub dropupleft(enabled)
  if enabled then
'   PlaySoundAt SoundFX(DTResetSound,DOFContactors), psw42
    DTRaise 33
    DTRaise 34
    DTRaise 35
  end if
End Sub

Sub dropupcenter(enabled)
  if enabled then
'   PlaySoundAt SoundFX(DTResetSound,DOFContactors), cutout_prim1
    DTRaise 20
  end if
End Sub

'Standup targets
Sub sw17_Hit:vpmTimer.PulseSw 17:End Sub
Sub sw18_Hit:vpmTimer.PulseSw 18:End Sub
Sub sw19_Hit:vpmTimer.PulseSw 19:End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:End Sub

'Wire Triggers
Sub sw13_Hit:Controller.Switch(13) = 1:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub
Sub sw14_Hit:Controller.Switch(14) = 1:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub
Sub sw15_Hit:Controller.Switch(15) = 1:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw30_Hit:Controller.Switch(30) = 1:End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub
Sub sw31_Hit:Controller.Switch(31) = 1:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub
Sub sw36_Hit:Controller.Switch(36) = 1:BIPL=1:End Sub
Sub sw36_UnHit:Controller.Switch(36) = 0:BIPL=0:End Sub

'Spinner
Sub sw37_Spin():VPMTimer.PulseSw 37:End Sub

'Ramp Switches
Sub sw32_Hit : vpmTimer.PulseSw(32): End Sub
Sub sw45_Hit : vpmTimer.PulseSw(45) : End Sub

'Bumpers

dim bump1step,bump2step,bump3step

Sub sw25_hit
  vpmTimer.PulseSw(25)
  RandomSoundBumperTop bumpera3
  bumpera3.RotY=5
  Bump1Step = 0
  sw25.timerenabled = 1
End Sub

Sub sw25_timer
  Select Case Bump1Step
    Case 3: bumpera3.RotY=3
    Case 4: bumpera3.RotY=-2
    Case 5: bumpera3.RotY=1
    Case 6: bumpera3.RotY=-1
    Case 7: bumpera3.RotY=0:sw25.timerenabled = 0
  End Select
  Bump1Step=Bump1Step + 1
End Sub

Sub sw26_hit
  vpmTimer.PulseSw(26)
  RandomSoundBumperMiddle bumperb3
  bumperb3.RotY=5
  Bump2Step = 0
  sw26.timerenabled = 1
End Sub

Sub sw26_timer
  Select Case Bump2Step
    Case 3: bumperb3.RotY=3
    Case 4: bumperb3.RotY=-2
    Case 5: bumperb3.RotY=1
    Case 6: bumperb3.RotY=-1
    Case 7: bumperb3.RotY=0:sw26.timerenabled = 0
  End Select
  Bump2Step=Bump2Step + 1
End Sub

Sub sw27_hit
  vpmTimer.PulseSw(27)
  RandomSoundBumperBottom bumperc3
  bumperc3.RotY=5
  Bump3Step = 0
  sw27.timerenabled = 1
End Sub

Sub sw27_timer
  Select Case Bump3Step
    Case 3: bumperc3.RotY=3
    Case 4: bumperc3.RotY=-2
    Case 5: bumperc3.RotY=1
    Case 6: bumperc3.RotY=-1
    Case 7: bumperc3.RotY=0:sw27.timerenabled = 0
  End Select
  Bump3Step=Bump3Step + 1
End Sub

Sub sw42_Hit
  vpmTimer.PulseSw 42
End Sub

Sub sw43_Hit
  vpmTimer.PulseSw 43
End Sub

Sub sw44_Hit
  vpmTimer.PulseSw 44
End Sub

Sub sw46_Hit
  vpmTimer.PulseSw 46
End Sub

Sub sw47_Hit
  vpmTimer.PulseSw 47
End Sub

Sub sw48_Hit
  vpmTimer.PulseSw 48
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 40
    RandomSoundSlingshotRight rsling
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 16
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 39
    RandomSoundSlingshotLeft lsling
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'***************************************************

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
FrameTimer.Interval = -1 'lamp fading speed
FrameTimer.Enabled = 1

' Called once per frame (16.6ms at 60hz)
Sub FrameTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
       SetLamp ChgLamp(ii,0), ChgLamp(ii,1)
        Next
    End If
  FadeLights.Update
  FlipperUpdate
  BallShadowUpdate
End Sub

' Called every 1ms.
Sub OneMsec_Timer()
  FadeLights.Update1
End Sub

' div lamp subs

Sub InitLamps()
  'FadeLights.Callback(120)="RampCallback "
  FadeLights.obj(7) = array(f7)
  FadeLights.obj(8) = array(f8)
  FadeLights.obj(9) = array(f9)
  FadeLights.obj(10) = array(f10)
  FadeLights.obj(11) = array(f11)
  FadeLights.obj(12) = array(f12)
  FadeLights.obj(13) = array(f13)
  FadeLights.obj(14) = array(f14)
  FadeLights.obj(15) = array(f15,f15a)
  FadeLights.obj(16) = array(f16)
  FadeLights.obj(17) = array(f17)
  FadeLights.obj(18) = array(f18)
  FadeLights.obj(19) = array(f19)
  FadeLights.obj(20) = array(f20)
  FadeLights.obj(21) = array(f21)
  FadeLights.obj(22) = array(f22)
  FadeLights.obj(23) = array(f23)
  FadeLights.obj(24) = array(f24)
  FadeLights.obj(28) = array(f28)
  FadeLights.obj(29) = array(f29)
  FadeLights.obj(30) = array(f30)
  FadeLights.obj(31) = array(f31)
  FadeLights.obj(32) = array(f32)
  FadeLights.obj(33) = array(f33)
  FadeLights.obj(34) = array(f34)
  FadeLights.obj(35) = array(f35)
  FadeLights.obj(36) = array(f36)
  FadeLights.obj(37) = array(f37)
  FadeLights.obj(41) = array(f41,f41a)
  FadeLights.obj(42) = array(f42)
  FadeLights.obj(43) = array(f43)
  FadeLights.obj(44) = array(f44)
  FadeLights.obj(45) = array(f45)
  FadeLights.obj(46) = array(f46)
  FadeLights.obj(47) = array(f47)
  FadeLights.obj(48) = array(f48)
  FadeLights.obj(49) = array(f49)
  FadeLights.obj(50) = array(f50)
  FadeLights.obj(51) = array(f51)
  FadeLights.obj(52) = array(f52)
  FadeLights.obj(53) = array(f53)
  FadeLights.obj(54) = array(f54)
  FadeLights.obj(55) = array(f55)
  FadeLights.obj(56) = array(f56)
  FadeLights.obj(57) = array(f57)
  FadeLights.obj(58) = array(f58)
  FadeLights.obj(59) = array(f59)
  FadeLights.obj(60) = array(f60)
  FadeLights.obj(61) = array(f61)
  FadeLights.obj(62) = array(f62)
  FadeLights.obj(63) = array(f63)
  FadeLights.obj(64) = array(f64)
  FadeLights.obj(101) = array(fspace,fspace1,fspace2)
  FadeLights.obj(102) = array(fshuttle,fshuttle1,fshuttle2)
  FadeLights.obj(110) = array(gion)
  FadeLights.obj(111) = array(GIoff)
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub UpdateLamp7
  if FadeLights.state(7) <> 0 Then
      centerpost_prim.image="pp_popupon"
    else
      centerpost_prim.image="pp_popupoff"
    end if
End Sub

Sub UpdateLamp25
  if FadeLights.state(25) <> 0 Then
    if FadeLights.state(110) <> 0 Then
      if FadeLights.state(26) <> 0 and FadeLights.state(27) <> 0 Then
        plastics2.image="plastics3lit"
      Else
        plastics2.image="plastics3bump1"
        if FadeLights.state(26) <> 0 Then
          plastics2.image="plastics3bump12"
        else
          if FadeLights.state(27) <> 0 Then
            plastics2.image="plastics3bump13"
          else
            plastics2.image="plastics3bump1"
          end if
        end if
      end if
      bumpera1.image="bumps1lit"
      bumpera2.image="bumps2lit"
      bumpera3.image="bumps3lit"
    Else
      if FadeLights.state(26) <> 0 and FadeLights.state(27) <> 0 Then
        plastics2.image="plastics3offlit"
      Else
        if FadeLights.state(26) <> 0 Then
          plastics2.image="plastics3offlit12"
        else
          if FadeLights.state(27) <> 0 Then
            plastics2.image="plastics3offlit13"
          else
            plastics2.image="plastics3offlit1"
          end if
        end if
      End if
      bumpera1.image="bumps1lit"
      bumpera2.image="bumps2off"
      bumpera3.image="bumps3offlit"
    End if
  Else
    if FadeLights.state(110) <> 0 Then
      bumpera1.image="bumps1"
      bumpera2.image="bumps2"
      bumpera3.image="bumps3"
      if FadeLights.state(26) <> 0 and FadeLights.state(27) <> 0 Then
        plastics2.image="plastics3bump23"
      Else
        if FadeLights.state(26) <> 0 Then
          plastics2.image="plastics3bump2"
        else
          if FadeLights.state(27) <> 0 Then
            plastics2.image="plastics3bump3"
          Else
            plastics2.image="plastics3"
          end if
        end if
      End if
    Else
      bumpera1.image="bumps1off"
      bumpera2.image="bumps2off"
      bumpera3.image="bumps3off"
      if FadeLights.state(26) <> 0 and FadeLights.state(27) <> 0 Then
        plastics2.image="plastics3offlit23"
      Else
        if FadeLights.state(26) <> 0 Then
          plastics2.image="plastics3offlit2"
        end if
        if FadeLights.state(27) <> 0 Then
        plastics2.image="plastics3offlit3"
        end if
      plastics2.image="plastics3off"
      End if
    End if
  End if
End Sub

Sub UpdateLamp26
  if FadeLights.state(26) <> 0 Then
    if FadeLights.state(110) <> 0 Then
      if FadeLights.state(25) <> 0 and FadeLights.state(27) <> 0 Then
        plastics2.image="plastics3lit"
      Else
        if FadeLights.state(25) <> 0 Then
          plastics2.image="plastics3bump12"
        else
          if FadeLights.state(27) <> 0 Then
            plastics2.image="plastics3bump23"
          else
            plastics2.image="plastics3bump2"
          end if
        end if
      End if
      bumperb1.image="bumps1lit"
      bumperb2.image="bumps2lit"
      bumperb3.image="bumps3lit"
    Else
      if FadeLights.state(25) <> 0 and FadeLights.state(27) <> 0 Then
        plastics2.image="plastics3offlit"
      Else
        if FadeLights.state(25) <> 0 Then
          plastics2.image="plastics3offlit12"
        else
          if FadeLights.state(27) <> 0 Then
            plastics2.image="plastics3offlit23"
          else
            plastics2.image="plastics3offlit2"
          end if
        End if
      bumperb1.image="bumps1"
      bumperb2.image="bumps2off"
      bumperb3.image="bumps3offlit"
      End if
    End if
  Else
    if FadeLights.state(110) <> 0 Then
      bumperb1.image="bumps1"
      bumperb2.image="bumps2"
      bumperb3.image="bumps3"
      if FadeLights.state(25) <> 0 and FadeLights.state(27) <> 0 Then
        plastics2.image="plastics3bump13"
      Else
        if FadeLights.state(25) <> 0 Then
          plastics2.image="plastics3bump1"
        else
          if FadeLights.state(27) <> 0 Then
            plastics2.image="plastics3bump3"
          else
            plastics2.image="plastics3"
          end If
        end if
      End if
    Else
      bumperb1.image="bumps1off"
      bumperb2.image="bumps2off"
      bumperb3.image="bumps3off"
      if FadeLights.state(25) <> 0 and FadeLights.state(27) <> 0 Then
        plastics2.image="plastics3offlit13"
      Else
        if FadeLights.state(25) <> 0 Then
          plastics2.image="plastics3offlit1"
        end if
        if FadeLights.state(27) <> 0 Then
        plastics2.image="plastics3offlit3"
        end if
      plastics2.image="plastics3off"
      End if
    End if
  End if
End Sub

Sub UpdateLamp27
  if FadeLights.state(27) <> 0 Then
    if FadeLights.state(110) <> 0 Then
      if FadeLights.state(25) <> 0 and FadeLights.state(26) <> 0 Then
        plastics2.image="plastics3lit"
      Else
        if FadeLights.state(25) <> 0 Then
          plastics2.image="plastics3bump13"
        else
          if FadeLights.state(26) <> 0 Then
            plastics2.image="plastics3bump23"
          Else
            plastics2.image="plastics3bump3"
          end if
        end if
      End if
      bumperc1.image="bumps1lit"
      bumperc2.image="bumps2lit"
      bumperc3.image="bumps3lit"
    Else
      if FadeLights.state(25) <> 0 and FadeLights.state(26) <> 0 Then
        plastics2.image="plastics3offlit"
      Else
        if FadeLights.state(25) <> 0 Then
          plastics2.image="plastics3offlit13"
        else
          if FadeLights.state(26) <> 0 Then
            plastics2.image="plastics3offlit23"
          else
            plastics2.image="plastics3offlit1"
          end if
        End if
      bumperc1.image="bumps1"
      bumperc2.image="bumps2off"
      bumperc3.image="bumps3offlit"
      End if
    End if
  Else
    if FadeLights.state(110) <> 0 Then
      bumperc1.image="bumps1"
      bumperc2.image="bumps2"
      bumperc3.image="bumps3"
      if FadeLights.state(25) <> 0 and FadeLights.state(26) <> 0 Then
        plastics2.image="plastics3bump12"
      Else
        if FadeLights.state(25) <> 0 Then
          plastics2.image="plastics3bump1"
        else
          if FadeLights.state(26) <> 0 Then
            plastics2.image="plastics3bump2"
          Else
            plastics2.image="plastics3"
          end if
        end if
      End if
    Else
      bumperc1.image="bumps1off"
      bumperc2.image="bumps2off"
      bumperc3.image="bumps3off"
      if FadeLights.state(25) <> 0 and FadeLights.state(26) <> 0 Then
        plastics2.image="plastics3offlit12"
      Else
        if FadeLights.state(25) <> 0 Then
          plastics2.image="plastics3offlit1"
        end if
        if FadeLights.state(26) <> 0 Then
        plastics2.image="plastics3offlit2"
        end if
      plastics2.image="plastics3off"
      End if
    End if
  End if
End Sub

Sub SetLamp(nr, value)
  ' If the lamp state is not changing, just exit.
  if FadeLights.state(nr) = value then exit sub

  FadeLights.state(nr) = value
  if nr = 7 then UpdateLamp7
  if nr = 25 then UpdateLamp25
  if nr = 26 then UpdateLamp26
  if nr = 27 then UpdateLamp27
End Sub



' *** NFozzy's lamp fade routines ***


Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class 'todo do better

Class LampFader
  Public FadeSpeedDown(140), FadeSpeedUp(140)
  Private Lock(140), Loaded(140), OnOff(140)
  Public UseFunction
  Private cFilter
  Private UseCallback(140), cCallback(140)
  Public Lvl(140), Obj(140)

  Sub Class_Initialize()
    dim x : for x = 0 to uBound(OnOff)  'Set up fade speeds
      if FadeSpeedDown(x) <= 0 then FadeSpeedDown(x) = 1/100  'fade speed down
      if FadeSpeedUp(x) <= 0 then FadeSpeedUp(x) = 1/80'Fade speed up
      UseFunction = False
      lvl(x) = 0
      OnOff(x) = False
      Lock(x) = True : Loaded(x) = False
    Next

    for x = 0 to uBound(OnOff)    'clear out empty obj
      if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
    Next
  End Sub

  Public Property Get Locked(idx) : Locked = Lock(idx) : End Property   'debug.print Lampz.Locked(100)  'debug
  Public Property Get state(idx) : state = OnOff(idx) : end Property
  Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
  Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property

  Public Property Let state(ByVal idx, input) 'Major update path
    input = cBool(input)
    if OnOff(idx) = Input then : Exit Property : End If 'discard redundant updates
    OnOff(idx) = input
    Lock(idx) = False
    Loaded(idx) = False
  End Property

  Public Sub TurnOnStates() 'If obj contains any light objects, set their states to 1 (Fading is our job!)
    dim debugstr
    dim idx : for idx = 0 to uBound(obj)
      if IsArray(obj(idx)) then
        dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
        for x = 0 to uBound(tmp)
          if typename(tmp(x)) = "Light" then DisableState tmp(x) : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline

        Next
      Else
        if typename(obj(idx)) = "Light" then DisableState obj(idx) : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline

      end if
    Next
    'debug.print debugstr
  End Sub
  Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub 'turn state to 1

  Public Sub Update1()   'Handle all boolean numeric fading. If done fading, Lock(True). Update on a '1' interval Timer!
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x)
          if Lvl(x) > 1 then Lvl(x) = 1 : Lock(x) = True
        elseif Not OnOff(x) then 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x)
          if Lvl(x) < 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
  End Sub


  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    dim x,xx : for x = 0 to uBound(OnOff)
      if not Loaded(x) then
        if IsArray(obj(x) ) Then  'if array
          If UseFunction then
            for each xx in obj(x) : xx.IntensityScale = cFilter(Lvl(x)) : Next
          Else
            for each xx in obj(x) : xx.IntensityScale = Lvl(x) : Next
          End If
        else            'if single lamp or flasher
          If UseFunction then
            obj(x).Intensityscale = cFilter(Lvl(x))
          Else
            obj(x).Intensityscale = Lvl(x)
          End If
        end if
        ' Sleazy hack for regional decimal point problem
        If UseCallBack(x) then execute cCallback(x) & " CSng(" & CInt(10000 * Lvl(x)) & " / 10000)" 'Callback
        If Lock(x) Then
          if Lvl(x) = 1 or Lvl(x) = 0 then Loaded(x) = True 'finished fading
        end if
      end if
    Next
  End Sub
End Class

'*****************************************
' FLIPPER SHADOWS
'*****************************************

sub FlipperUpdate()
  batleftshadow.rotz = LeftFlipper.CurrentAngle
  batrightshadow.rotz  = RightFlipper.CurrentAngle
  lflip.rotz = leftflipper.currentangle
  rflip.rotz = rightflipper.currentangle
  Diverter_prim.rotz = bottomgate.currentangle-55
End Sub

'*****************************************
' BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate()
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

'******************************************************
'   FLIPPER CORRECTION SUPPORTING FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF, RF1)
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

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity
dim RF1 : Set RF1 = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF, RF1)
  for each x in a
    'safety coefficient (diminishes polarity correction only)
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1

    x.enabled = True
    x.TimeDelay = 80
  Next

  '"Polarity" Profile
'"Polarity" Profile<br>
' AddPt "Polarity", 0, 0, 0
' AddPt "Polarity", 1, 0.1, 0
' AddPt "Polarity", 2, 0.14, -2.25
' AddPt "Polarity", 3, 0.2, -2.25
' AddPt "Polarity", 4, 0.28, -3.25
' AddPt "Polarity", 5, 0.31, -3.25
' AddPt "Polarity", 6, 0.34, -3.75
' AddPt "Polarity", 7, 0.37, -3.75
' AddPt "Polarity", 8, 0.4, -4.5
' AddPt "Polarity", 9, 0.45, -3.5
' AddPt "Polarity", 10, 0.48, -3.5
' AddPt "Polarity", 11, 0.51, -3.75
' AddPt "Polarity", 12, 0.55, -3.75
' AddPt "Polarity", 13, 0.58, -3
' AddPt "Polarity", 14, 0.6, -2.75
' AddPt "Polarity", 15, 0.62, -2.75
' AddPt "Polarity", 16, 0.65, -2.5
' AddPt "Polarity", 17, 0.8, -2
' AddPt "Polarity", 18, 0.85, -1.9
' AddPt "Polarity", 19, 1.0, -1
' AddPt "Polarity", 20, 1.2, 0

  'rf.report "Polarity"
  AddPt "Polarity", 0, 0, -2.7
  AddPt "Polarity", 1, 0.16, -2.7
  AddPt "Polarity", 2, 0.33, -2.7
  AddPt "Polarity", 3, 0.37, -2.7 '4.2
  AddPt "Polarity", 4, 0.41, -2.7
  AddPt "Polarity", 5, 0.45, -2.7 '4.2
  AddPt "Polarity", 6, 0.576,-2.7
  AddPt "Polarity", 7, 0.66, -1.8'-2.1896
  AddPt "Polarity", 8, 0.743, -0.5
  AddPt "Polarity", 9, 0.81, -0.5
  AddPt "Polarity", 10, 0.88, 0

  '"Velocity" Profile
  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,  1.05
  addpt "Velocity", 3, 0.53,  1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,  0.945

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub
Sub TriggerRF1_Hit() : RF1.Addball activeball : End Sub
Sub TriggerRF1_UnHit() : RF1.PolarityCorrect activeball : End Sub


' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
' Sub RDampen_Timer()
'   Cor.Update
' End Sub

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
RubbersD.addpoint 0, 0, 1.1 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

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

Sub RDampen_Timer()
        Cor.Update
End Sub


'******************************************************
'                DROP TARGETS INITIALIZATION
'******************************************************

'Set array with drop target objects

dim DT20, DT33, DT34, DT35
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'         primary:                         primary target wall to determine drop
'        secondary:                        wall used to simulate the ball striking a bent or offset target after the initial Hit
'        prim:                                primitive target used for visuals and animation
'                                                        IMPORTANT!!!
'                                                        rotz must be used for orientation
'                                                        rotx to bend the target back
'                                                        transz to move it up and down
'                                                        the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
'        switch:                                ROM switch number
'        animate:                        Arrary slot for handling the animation instrucitons, set to 0
'
'        Values for annimate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target

' Center Bank
DT20 = Array(sw20, sw20y, psw20, 20, 0)

' Left Bank
DT33 = Array(sw33, sw33y, psw33, 33, 0)
DT34 = Array(sw34, sw34y, psw34, 34, 0)
DT35 = Array(sw35, sw35y, psw35, 35, 0)

'Add all the Drop Target Arrays to Drop Target Animation Array
' DTAnimationArray = Array(DT1, DT2, ....)
Dim DTArray
DTArray = Array(DT20, DT33, DT34, DT35)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 110                         'in milliseconds
Const DTDropUpSpeed = 40                         'in milliseconds
Const DTDropUnits = 60                         'VP units primitive drops
Const DTDropUpUnits = 10                         'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8                                 'max degrees primitive rotates when hit
Const DTDropDelay = 20                         'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40                         'time in milliseconds before target drops back to normal up position after the solendoid fires to raise the target
Const DTBrickVel = 30                                'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0                        'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "targethit"        'Drop Target Hit sound
Const DTDropSound = "DTDrop"                'Drop Target Drop sound
Const DTResetSound = "DTReset"        'Drop Target reset sound

Const DTMass = 0.2                                'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'                                DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
        Dim i
        i = DTArrayID(switch)

        PlaySoundAtVol  DTHitSound, Activeball, Vol(Activeball)*22.5
        DTArray(i)(4) =  DTCheckBrick(Activeball,DTArray(i)(2))
    If DTArray(i)(4) = 1 or DTArray(i)(4) = 3 or DTArray(i)(4) = 4 Then
                DTBallPhysics Activeball, DTArray(i)(2).rotz, DTMass
        End If
        DoDTAnim
End Sub

Sub DTRaise(switch)
        Dim i
        i = DTArrayID(switch)

        DTArray(i)(4) = -1
        DoDTAnim
End Sub

Sub DTDrop(switch)
        Dim i
        i = DTArrayID(switch)

        DTArray(i)(4) = 1
        DoDTAnim
End Sub

Function DTArrayID(switch)
        Dim i
        For i = 0 to uBound(DTArray)
                If DTArray(i)(3) = switch Then DTArrayID = i:Exit Function
        Next
End Function


sub DTBallPhysics(aBall, angle, mass)
        dim rangle,bangle,calc1, calc2, calc3
        rangle = (angle - 90) * 3.1416 / 180
        bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

        calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
        calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
        calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

        aBall.velx = calc1 * cos(rangle) + calc2
        aBall.vely = calc1 * sin(rangle) + calc3
End Sub

'Add Timer name DTAnim to editor to handle drop target animations
DTAnim.interval = 10
DTAnim.enabled = True

Sub DTAnim_Timer()
        DoDTAnim
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - tan(bangle) * aball.x + tan(rangle2) * dtprim.x) / (tan(rangle2) - tan(bangle))
  Yintersect = tan(rangle2) * Xintersect + (dtprim.y - tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

  If perpvel > 0 and  perpvelafter <= 0 Then
    If DTEnableBrick = 1 and  perpvel > DTBrickVel and DTBrickVel <> 0 and cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
    DTCheckBrick = 4
  Else
    DTCheckBrick = 0
  End If
End Function


Sub DoDTAnim()
        Dim i
        For i=0 to Ubound(DTArray)
                DTArray(i)(4) = DTAnimate(DTArray(i)(0),DTArray(i)(1),DTArray(i)(2),DTArray(i)(3),DTArray(i)(4))
        Next
End Sub

Function DTAnimate(primary, secondary, prim, switch,  animate)
        dim transz
        Dim animtime, rangle

        rangle = prim.rotz * 3.1416 / 180

        DTAnimate = animate

        if animate = 0  Then
                primary.uservalue = 0
                DTAnimate = 0
                Exit Function
        Elseif primary.uservalue = 0 then
                primary.uservalue = gametime
        end if

        animtime = gametime - primary.uservalue

        If (animate = 1 or animate = 4) and animtime < DTDropDelay Then
      primary.collidable = 0
      If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
      prim.rotx = DTMaxBend * cos(rangle)
      prim.roty = DTMaxBend * sin(rangle)
      DTAnimate = animate
      Exit Function
    elseif (animate = 1 or animate = 4) and animtime > DTDropDelay Then
      primary.collidable = 0
      If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
      prim.rotx = DTMaxBend * cos(rangle)
      prim.roty = DTMaxBend * sin(rangle)
      animate = 2
      PlaySoundAt SoundFX(DTDropSound,DOFDropTargets),prim
    End If

        if animate = 2 Then
                transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
                if prim.transz > -DTDropUnits  Then
                        prim.transz = transz
                end if

                prim.rotx = DTMaxBend * cos(rangle)/2
                prim.roty = DTMaxBend * sin(rangle)/2

                if prim.transz <= -DTDropUnits Then
                        prim.transz = -DTDropUnits
                        secondary.collidable = 0
                        controller.Switch(Switch) = 1
                        primary.uservalue = 0
                        DTAnimate = 0
                        Exit Function
                Else
                        DTAnimate = 2
                        Exit Function
                end If
        End If

        If animate = 3 and animtime < DTDropDelay Then
                primary.collidable = 0
                secondary.collidable = 1
                prim.rotx = DTMaxBend * cos(rangle)
                prim.roty = DTMaxBend * sin(rangle)
        elseif animate = 3 and animtime > DTDropDelay Then
                primary.collidable = 1
                secondary.collidable = 0
                prim.rotx = 0
                prim.roty = 0
                primary.uservalue = 0
                DTAnimate = 0
                Exit Function
        End If

        if animate = -1 Then
                transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1

                If prim.transz = -DTDropUnits Then
                        Dim BOT, b
                        BOT = GetBalls

                        For b = 0 to UBound(BOT)
                                If InRect(BOT(b).x,BOT(b).y,prim.x-25,prim.y-10,prim.x+25, prim.y-10,prim.x+25,prim.y+25,prim.x -25,prim.y+25) Then
                                        BOT(b).velz = 20
                                End If
                        Next
                End If

                if prim.transz < 0 Then
                        prim.transz = transz
                elseif transz > 0 then
                        prim.transz = transz
                end if

                if prim.transz > DTDropUpUnits then
                        DTAnimate = -2
                        prim.rotx = 0
                        prim.roty = 0
                        primary.uservalue = gametime
                end if
                primary.collidable = 0
                secondary.collidable = 1
                controller.Switch(Switch) = 0

        End If

        if animate = -2 and animtime > DTRaiseDelay Then
                prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits
                if prim.transz < 0 then
                        prim.transz = 0
                        primary.uservalue = 0
                        DTAnimate = 0

                        primary.collidable = 1
                        secondary.collidable = 0
                end If
        End If
End Function

'******************************************************
'                DROP TARGET
'                SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for drop targets
Function Atn2(dy, dx)
        dim pi
        pi = 4*Atn(1)

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

' Used for drop targets
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

' Used for drop targets
Function Distance(ax,ay,bx,by)
        Distance = SQR((ax - bx)^2 + (ay - by)^2)
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

'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                         'volume level; range [0, 1]
NudgeRightSoundLevel = 1                        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                       'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                       'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                     'volume level; range [0, 1]
PlungerPullSoundLevel = 1                       'volume level; range [0, 1]
RollingSoundFactor = 1.1/5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                      'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                  'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel               'sound helper; not configurable
SlingshotSoundLevel = 0.95                        'volume level; range [0, 1]
BumperSoundFactor = 4.25                        'volume multiplier; must not be zero
KnockerSoundLevel = 1                           'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                  'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                     'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                     'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                    'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                 'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                 'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                     'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                          'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                     'volume multiplier; must not be zero
DTSoundLevel = 0.25                           'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                       'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                           'volume level; range [0, 1]
BallReleaseSoundLevel = 1                       'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                   'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                         'volume multiplier; must not be zero


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
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
    PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


' *********************************************************************
'                     Fleep  Supporting Ball & Sound Functions
' *********************************************************************

Dim tablewidth, tableheight : tablewidth = table1.width : tableheight = table1.height

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

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
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
  PlaySoundAtLevelStatic SoundFX("bell",DOFBell), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////
Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*7)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
    PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm/10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm/10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong 1
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong 1
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

Sub RandomSoundWall()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
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
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft Activeball
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
  If Activeball.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If activeball.velx < -8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If Activeball.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If activeball.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Activeball
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  Dim snd
  Select Case Int(Rnd*7)+1
    Case 1 : snd = "Ball_Collide_1"
    Case 2 : snd = "Ball_Collide_2"
    Case 3 : snd = "Ball_Collide_3"
    Case 4 : snd = "Ball_Collide_4"
    Case 5 : snd = "Ball_Collide_5"
    Case 6 : snd = "Ball_Collide_6"
    Case 7 : snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////


'******************************************************
'   BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 3 ' total number of balls
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

Sub RollingTimer_Timer()
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

    '***Ball Drop Sounds***
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz > -7 Then
          RandomSoundBallBouncePlayfieldSoft BOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard BOT(b)
        End If
      End If
    End If
    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If
  Next
End Sub


'******************************************************
'         KNOCKER
'******************************************************

Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid 'Add knocker position object
  End If
End Sub

'******************************************************
'         Saucer Kick
'******************************************************

Sub SolSaucer(Enabled)
  If Enabled Then
    If BallInSaucer = True Then
      SoundSaucerKick 1, Kicker
    Else
      SoundSaucerKick 0, Kicker
    End If
  End If
End Sub
