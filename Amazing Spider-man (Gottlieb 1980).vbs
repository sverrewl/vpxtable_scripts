'VPX Amazing Spider-Man
' BorgDog mod of gtxjoe table
' dof fixes by arngrim
'
'   gtxjoe v1.0
'
'       Thanks to all VP contributors present and past
'       Inspiration from Bob5453 and Gaston for their previous Spider-man versions
 
Option Explicit
Randomize
 
Const cGameName     = "spidermn"  
 
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0
 
LoadVPM "01000100", "sys80.vbs", 2.31
 
Const UseSolenoids  = 1
Const UseLamps      = 1
Const UseGI         = 0
 
' Standard Sounds
Const SSolenoidOn  = ""
Const SSolenoidOff = ""
Const SCoin        = "fx_coin"
 
Dim xx, bsTrough, initContacts, dtRDrop, dtLDrop, bsLSaucer, bsMSaucer, bsRSaucer
Sub AmazingSpiderMan_Init()
 
 
    On Error Resume Next
      With Controller
        .GameName                               = cGameName
        .SplashInfoLine                         = "Amazing Spider-Man, Gottlieb 1980"
        .HandleKeyboard                         = 0
        .ShowTitle                              = 0
        .ShowDMDOnly                            = 1
        .ShowFrame                              = 0
        .ShowTitle                              = 0
        .hidden                                 = 1
        .Games(cGameName).Settings.Value("rol") = 0
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
      End With
    On Error Goto 0
       
' basic pinmame timer
    PinMAMETimer.Interval   = PinMAMEInterval
    PinMAMETimer.Enabled    = 1
 
' Nudging
    vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1,Bumper2,LeftSlingshot,RightSlingShot)
 
   
'   ball stack for trough: Gottlieb System 80
    Set bsTrough = New cvpmBallStack
        bsTrough.InitSw 0, 67, 0, 0, 0, 0, 0, 0
        bsTrough.InitKick Drain, 60,25
        bsTrough.InitExitSnd Soundfx("BallRelease",DOFContactors), Soundfx("SolOn",DOFContactors)
        bsTrough.Balls = 1
 
 
    Set dtLDrop = New cvpmDropTarget
    dtLDrop.InitDrop Array(sw1,sw11,sw21),Array(1,11,21)
    dtLDrop.InitSnd SoundFX("DropTarget",DOFTargets), Soundfx("DropTargetreset",DOFContactors)
 
    Set dtRDrop = New cvpmDropTarget
    dtRDrop.InitDrop Array(sw0,sw10,sw20,sw30,sw40),Array(0,10,20,30,40)
    dtRDrop.InitSnd SoundFX("DropTarget",DOFTargets),Soundfx("DropTargetreset",DOFContactors)
 
    if b2son then: for each xx in backdropstuff: xx.visible=0: next
    startGame.enabled=true
    NTMBox.visible=0    'turn off number to match box
    FindDips        'find if match enabled, if so turn back on number to match box
 
End Sub
 
Sub AmazingSpiderMan_Paused:Controller.Pause = 1:End Sub
 
Sub AmazingSpiderMan_unPaused:Controller.Pause = 0:End Sub
 
Sub AmazingSpiderMan_Exit
    If b2son then controller.stop
End Sub
 
sub startGame_timer
    dim xx
    playsound "poweron"
    LampTimer.enabled=1
    For each xx in GI:xx.State = 1: Next        '*****GI Lights On
    me.enabled=false
end sub
 
Sub AmazingSpiderMan_KeyDown(ByVal keycode)
 
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then
        Plunger.PullBack
        PlaySound "plungerpull",0,1,0.25,0.25
    End If
    If keycode=AddCreditKey then playsound "coin": vpmTimer.pulseSW (swCoin1): end if
 
   
End Sub
 
Sub AmazingSpiderMan_KeyUp(ByVal keycode)
 
    If keycode = 61 then FindDips
 
    If keycode = PlungerKey Then
        Plunger.Fire
        PlaySound "plunger",0,1,0.25,0.25
    End If
   
    If vpmKeyUp(keycode) Then Exit Sub
 
End Sub
 
Sub Drain_Hit()
    PlaySound "drain",0,1,0,0.25
    bsTrough.AddBall Me
End Sub
 
 
'*****************************************
'Solenoids
'*****************************************
SolCallBack(1)  = "SolSaucerMid"
SolCallBack(2)  = "SolSaucerOuter"
SolCallback(5)  = "Lraised"     '"dtLDrop.SolDropUp"
SolCallback(6)  = "Rraised"     '"dtRDrop.SolDropUp"
SolCallback(8)  = "SolKnocker"
SolCallback(9)  = "SolTrough"
 
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
 
Sub Rraised(enabled)
    if enabled then Rreset.enabled=True
End Sub
 
Sub Rreset_timer
    dtRDrop.DropSol_On 
    For each light in DTRightLights: light.state=0: Next
    Rreset.enabled=false
End Sub
 
Sub Lraised(enabled)
    if enabled then Lreset.enabled=True
End Sub
 
Sub Lreset_timer
    dtLDrop.DropSol_On 
    For each light in DTLeftLights: light.state=0: Next
    Lreset.enabled=false
End Sub
 
 
Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
        GIlflipper.state=1
    Else
        PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
        GIlflipper.state=0
    End If
End Sub
 
Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
    End If
End Sub
 
Sub SolTrough(Enabled)
    If Enabled Then bsTrough.ExitSol_On
End Sub
 
Sub SolKnocker(Enabled)
    If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub
 
Sub SolSaucerMid(Enabled)
    If Enabled Then
        sw12.kick 173,6
        Controller.Switch(12) = False
        PlaySound SoundFX("popper_ball",DOFContactors)
        sw12.uservalue=1
        sw12.timerenabled=1
        PkickarmSW12.rotz=15
    end if
End Sub
 
Sub sw12_timer
    select case sw12.uservalue
      case 2:
        PkickarmSW12.rotz=0
        me.timerenabled=0
    end Select
    sw12.uservalue=sw12.uservalue+1
End Sub
 
Sub SolSaucerOuter(Enabled)
    If Enabled Then
        Sw2.kick 176,6
        Controller.Switch(2) = False
        PlaySound SoundFX("popper_ball",DOFContactors)
        sw22.kick 170,6
        Controller.Switch(22) = False
        sw2.uservalue=1
        sw2.timerenabled=1
        PkickarmSW2.rotz=15
        PkickarmSW22.rotz=15
    end if
End Sub
 
Sub sw2_timer
    select case sw2.uservalue
      case 2:
        PkickarmSW2.rotz=0
        me.timerenabled=0
    end Select
    sw2.uservalue=sw2.uservalue+1
End Sub
 
'*****************************************
'Switches
'*****************************************
 
Sub sw0_Hit  : dtRDrop.Hit 1 : GIsw0.state=1: End Sub
Sub sw10_Hit : dtRDrop.Hit 2 : GIsw10.state=1: End Sub
Sub sw20_Hit : dtRDrop.Hit 3 : GIsw20.state=1: End Sub
Sub sw30_Hit : dtRDrop.Hit 4 : GIsw30.state=1: End Sub
Sub sw40_Hit : dtRDrop.Hit 5 : GIsw40.state=1: End Sub
Sub sw1_Hit  : dtLDrop.Hit 1 : GIsw1.state=1: End Sub
Sub sw11_Hit : dtLDrop.Hit 2 : GIsw11.state=1: End Sub
Sub sw21_Hit : dtLDrop.Hit 3 : GIsw21.state=1: End Sub
 
Sub sw31_Hit  : debug.print "31":vpmTimer.PulseSw 31:DOF 111, DOFPulse:End Sub
Sub sw31a_Hit : vpmTimer.PulseSw 31:DOF 112, DOFPulse:End Sub
 
Sub sw32a_Hit : vpmTimer.PulseSw 32: End Sub
Sub sw32b_Hit : vpmTimer.PulseSw 32: End Sub
Sub sw32c_Hit : vpmTimer.PulseSw 32: End Sub
Sub sw32d_Hit : vpmTimer.PulseSw 32: End Sub
 
Sub sw41_Hit :  debug.print "41": vpmTimer.PulseSw 41:DOF 112, DOFPulse:End Sub
Sub sw41a_Hit : vpmTimer.PulseSw 41:DOF 109, DOFPulse:End Sub
 
    Sub DingwallA_hit()
        vpmTimer.PulseSw 42
        SlingA.visible=0
        SlingA1.visible=1
        me.uservalue=1
        Me.timerenabled=1
    End Sub
 
    sub dingwalla_timer                            
        select case dingwalla.uservalue
            Case 1: SlingA1.visible=0: SlingA.visible=1
            case 2: SlingA.visible=0: SlingA2.visible=1
            Case 3: SlingA2.visible=0: SlingA.visible=1: Me.timerenabled=0
        end Select
        me.uservalue=me.uservalue+1
    end sub
 
    Sub DingwallB_hit()
        vpmTimer.PulseSw 42
        SlingB.visible=0
        SlingB1.visible=1
        me.uservalue=1
        Me.timerenabled=1
    End Sub
 
    sub dingwallb_timer                                 'default 50 timer
        select case DingwallB.uservalue
            Case 1: Slingb1.visible=0: SlingB.visible=1
            case 2: SlingB.visible=0: Slingb2.visible=1
            Case 3: Slingb2.visible=0: SlingB.visible=1: Me.timerenabled=0
        end Select
        DingwallB.uservalue=DingwallB.uservalue+1
    end sub
 
Sub sw42a_Hit : vpmTimer.PulseSw 42: End Sub
Sub sw42b_Hit : vpmTimer.PulseSw 42: End Sub
Sub sw42c_Hit : vpmTimer.PulseSw 42: End Sub
Sub sw42d_Hit : vpmTimer.PulseSw 42: End Sub
 
 
Sub sw50_Hit : vpmTimer.PulseSw 50: End Sub
Sub sw51_Hit : vpmTimer.PulseSw 51:DOF 105, DOFPulse:End Sub
Sub sw51a_Hit : vpmTimer.PulseSw 51:DOF 107, DOFPulse:End Sub
Sub sw61_Hit : vpmTimer.PulseSw 61:DOF 106, DOFPulse:End Sub
Sub sw61a_Hit : vpmTimer.PulseSw 61:DOF 108, DOFPulse:End Sub
Sub sw70_Hit : vpmTimer.PulseSw 70:playsound SoundFX("target",DOFTargets): End Sub
Sub sw71_Hit : vpmTimer.PulseSw 71:playsound SoundFX("target",DOFTargets): End Sub
 
Sub sw60_spin
    vpmTimer.PulseSw 60
    PlaySound "fx_spinner",0,.25,0,0.25
'   playsound "fx_spinner"
End Sub
 
Sub sw2_Hit
    Controller.Switch(2) = True
    PlaySound "kicker_enter"
End Sub
Sub sw12_Hit
    Controller.Switch(12) = True
    PlaySound "kicker_enter"
End Sub
Sub sw22_Hit
    Controller.Switch(22) = True
    PlaySound "kicker_enter"
End Sub
 
 
Sub Bumper1_Hit
    vpmTimer.PulseSw 72
    PlaySound SoundFX("fx_bumper2",DOFContactors)
    DOF 104,DOFPulse
End Sub
 
Sub Bumper2_Hit
    vpmTimer.PulseSw 72
    PlaySound SoundFX("fx_bumper2",DOFContactors)
    DOF 103,DOFPulse
End Sub
 
 
 
'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep
 
Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw 42
    DOF 102,DOFPulse
    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.objroty = -15
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub
 
Sub RightSlingShot_Timer
    Select Case RStep
        Case 2:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.objroty = -7
        Case 3:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.objroty = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub
 
Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 42
    DOF 101,DOFPulse
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.objroty = 15
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub
 
Sub LeftSlingShot_Timer
    Select Case LStep
        Case 2:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.objroty = 7
        Case 3:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.objroty = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub
 
Sub FlipperTimer_Timer()
 
End Sub
 
 
 
 
 
 
' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************
 
Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 500)
End Function
 
Function Vol2(ball1, ball2) ' Calculates the Volume of the sound based on the speed of two balls
    Vol2 = (Vol(ball1) + Vol(ball2) ) / 2
End Function
 
Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / AmazingSpiderMan.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function
 
Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function
 
Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function
 
'*****************************************
'    JP's VP10 Collision & Rolling Sounds
'*****************************************
 
Const tnob = 16 ' total number of balls
ReDim rolling(tnob)
ReDim collision(tnob)
Initcollision
 
Sub Initcollision
    Dim i
    For i = 0 to tnob
        collision(i) = -1
        rolling(i) = False
    Next
End Sub
 
Sub CollisionTimer_Timer()
 
    'Added flipper and gate and spinner rotation here
    dim PI:PI=3.1415926
    LFLogo.RotY = LeftFlipper.CurrentAngle
    LFLogo1.RotY = LeftFlipper1.CurrentAngle
    RFLogo.RotY = RightFlipper.CurrentAngle
    RFLogo1.RotY = RightFlipper1.CurrentAngle
    PrimGate3.Rotz = Gate3.CurrentAngle * 70/90
    PrimGate1.Rotz = Gate1.CurrentAngle * 70/90
    PrimSw60.Rotz = sw60.CurrentAngle
    SpinnerRod.TransZ = sin( (sw60.CurrentAngle+180) * (2*PI/360)) * 5
    SpinnerRod.TransX = -1*(sin( (sw60.CurrentAngle- 90) * (2*PI/360)) * 5)
 
 
    Dim BOT, B, B1, B2, dx, dy, dz, distance, radii
    BOT = GetBalls
 
    ' rolling
   
    For B = UBound(BOT) +1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next
 
    If UBound(BOT) = -1 Then Exit Sub
 
    For B = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
 
    'collision
 
    If UBound(BOT) < 1 Then Exit Sub
 
    For B1 = 0 to UBound(BOT)
        For B2 = B1 + 1 to UBound(BOT)
            dz = INT(ABS((BOT(b1).z - BOT(b2).z) ) )
            radii = BOT(b1).radius + BOT(b2).radius
            If dz <= radii Then
 
            dx = INT(ABS((BOT(b1).x - BOT(b2).x) ) )
            dy = INT(ABS((BOT(b1).y - BOT(b2).y) ) )
            distance = INT(SQR(dx ^2 + dy ^2) )
 
            If distance <= radii AND (collision(b1) = -1 OR collision(b2) = -1) Then
                collision(b1) = b2
                collision(b2) = b1
                PlaySound("fx_collide"), 0, Vol2(BOT(b1), BOT(b2)), Pan(BOT(b1)), 0, Pitch(BOT(b1)), 0, 0
            Else
                If distance > (radii + 10)  Then
                    If collision(b1) = b2 Then collision(b1) = -1
                    If collision(b2) = b1 Then collision(b2) = -1
                End If
            End If
            End If
        Next
    Next
End Sub
 
 
'************************************
' What you need to add to your table
'************************************
 
' a timer called CollisionTimer. With a fast interval, like from 1 to 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc
 
 
'******************************************
' Explanation of the rolling sound routine
'******************************************
 
' sounds are played based on the ball speed and position
 
' the routine checks first for deleted balls and stops the rolling sound.
 
' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.
 
' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.
 
 
'**************************************
' Explanation of the collision routine
'**************************************
 
' The Double For loop: This is a double cycle used to check the collision between a ball and the other ones.
' If you look at the parameters of both cycles, you’ll notice they are designed to avoid checking
' collision twice. For example, I will never check collision between ball 2 and ball 1,
' because I already checked collision between ball 1 and 2. So, if we have 4 balls,
' the collision checks will be: ball 1 with 2, 1 with 3, 1 with 4, 2 with 3, 2 with 4 and 3 with 4.
 
' Sum first the radius of both balls, and if the height between them is higher then do not calculate anything more,
' since the balls are on different heights so they can't collide.
 
' The next 3 lines calculates distance between xth and yth balls with the Pytagorean theorem,
 
' The first "If": Checking if the distance between the two balls is less than the sum of the radius of both balls,
' and both balls are not already colliding.
 
' Why are we checking if balls are already in collision?
' Because we do not want the sound repeting when two balls are resting closed to each other.
 
' Set the collision property of both balls to True, and we assign the number of the ball colliding
 
' Play the collide sound of your choice using the VOL, PAN and PITCH functions
 
' Last lines: If the distance between 2 balls is more than the radius of a ball,
' then there is no collision and then set the collision property of the ball to False (-1).
 
Sub Rollovers_Hit (idx)
    'PlaySound "rollover", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    PlaySound "rollover"
End Sub
 
Sub Pins_Hit (idx)
'   PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    PlaySound "pinhit_low"
End Sub
 
Sub Targets_Hit (idx)
    'PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    PlaySound "target"
End Sub
 
Sub Metals_Thin_Hit (idx)
    'PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    PlaySound "metalhit_thin"
End Sub
 
Sub Metals_Medium_Hit (idx)
    'PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    PlaySound "metalhit_medium"
End Sub
 
Sub Metals2_Hit (idx)
    PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub
 
Sub Gates_Hit (idx)
    PlaySound "fx_gate", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub
 
Sub Spinner_Spin
    PlaySound "fx_spinner",0,.25,0,0.25
End Sub
 
Sub Rubbers_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub
 
Sub Posts_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
        RandomSoundRubber()
    End If
End Sub
 
Sub RandomSoundRubber()
    Select Case Int(Rnd*3)+1
        Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End Select
End Sub
 
Sub LeftFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub
Sub LeftFlipper1_Collide(parm)
    RandomSoundFlipper()
End Sub
 
Sub RightFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub
Sub RightFlipper1_Collide(parm)
    RandomSoundFlipper()
End Sub
 
Sub RandomSoundFlipper()
    Select Case Int(Rnd*3)+1
        Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End Select
End sub
 
'Finding an individual dip state based on scapino's Strikes and spares dip code - from unclewillys pinball pool, added another section to get high score award to set replay cards
Dim TheDips(32)
Sub FindDips
    Dim DipsNumber
    DipsNumber = Controller.Dip(1)
    TheDips(16) = Int(DipsNumber/128)
    If TheDips(16) = 1 then DipsNumber = DipsNumber - 128 end if
    TheDips(15) = Int(DipsNumber/64)
    If TheDips(15) = 1 then DipsNumber = DipsNumber - 64 end if
    TheDips(14) = Int(DipsNumber/32)
    If TheDips(14) = 1 then DipsNumber = DipsNumber - 32 end if
    TheDips(13) = Int(DipsNumber/16)
    If TheDips(13) = 1 then DipsNumber = DipsNumber - 16 end if
    TheDips(12) = Int(DipsNumber/8)
    If TheDips(12) = 1 then DipsNumber = DipsNumber - 8 end if
    TheDips(11) = Int(DipsNumber/4)
    If TheDips(11) = 1 then DipsNumber = DipsNumber - 4 end if
    TheDips(10) = Int(DipsNumber/2)
    If TheDips(10) = 1 then DipsNumber = DipsNumber - 2 end if
    TheDips(9) = Int(DipsNumber)
    DipsNumber = Controller.Dip(2)
    TheDips(24) = Int(DipsNumber/128)
    If TheDips(24) = 1 then DipsNumber = DipsNumber - 128 end if
    TheDips(23) = Int(DipsNumber/64)
    If TheDips(23) = 1 then DipsNumber = DipsNumber - 64 end if
    TheDips(22) = Int(DipsNumber/32)
    If TheDips(22) = 1 then DipsNumber = DipsNumber - 32 end if
    TheDips(21) = Int(DipsNumber/16)
    If TheDips(21) = 1 then DipsNumber = DipsNumber - 16 end if
    TheDips(20) = Int(DipsNumber/8)
    If TheDips(20) = 1 then DipsNumber = DipsNumber - 8 end if
    TheDips(19) = Int(DipsNumber/4)
    If TheDips(19) = 1 then DipsNumber = DipsNumber - 4 end if
    TheDips(18) = Int(DipsNumber/2)
    If TheDips(18) = 1 then DipsNumber = DipsNumber - 2 end if
    TheDips(17) = Int(DipsNumber)
    DipsTimer.Enabled=1
End Sub
 
 Sub DipsTimer_Timer()
    dim match
    match = TheDips(18)
    if match = 1 and NOT b2son then
        NTMBox.visible=1
      else
        NTMBox.visible=0
    end if
'   hsaward = TheDips(22)
'   BPG = TheDips(17)
'   If BPG = 1 then
'       instcard.image="InstCard3Balls"
'     Else
'       instcard.image="InstCard5Balls"
'   End if
'   replaycard.image="replaycard"&hsaward
    DipsTimer.enabled=0
 End Sub
 
Sub editDips
    Dim vpmDips : Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700,400,"Spiderman - DIP switches"
        .AddFrame 2,0,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
        .AddFrame 2,76,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
        .AddFrame 2,122,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
        .AddFrame 2,168,190,"Game mode",&H00100000,Array("replay",0,"extra ball",&H00100000)'dip 21
        .AddFrame 2,214,190,"3rd coin chute credits control",&H00001000,Array("no effect",0,"add 9",&H00001000)'dip 13
        .AddFrame 2,260,190,"Tilt penalty",&H10000000,Array("game over",0,"ball in play",&H10000000)'dip 29
        .AddChk 2,310,140,Array("Sound when scoring?",&H01000000)'dip 25
        .AddChk 2,325,140,Array("Replay button tune?",&H02000000)'dip 26
        .AddChk 2,340,140,Array("Coin switch tune?",&H04000000)'dip 27
        .AddFrame 205,0,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
        .AddFrameExtra 205,76,190,"Attract tune",&H0200,Array("no attract tune",0,"attract tune played every 6 minutes",&H0200)'S-board dip 2
        .AddFrame 205,122,190,"Balls per game",&H00010000,Array("5 balls",0,"3 balls",&H00010000)'dip 17
        .AddFrame 205,168,190,"Replay limit",&H00040000,Array("no limit",0,"one per game",&H00040000)'dip 19
        .AddFrame 205,214,190,"Novelty",&H00080000,Array("normal game mode",0,"50K for special/extra ball",&H0080000)'dip 20
        .AddFrameExtra 205,260,190,"Sound option",&H0100,Array("sound mode",0,"tone mode",&H0100)'S-board dip 1
        .AddChk 205,310,140,Array("Credits displayed?",&H08000000)'dip 28
        .AddChk 205,325,140,Array("Match feature",&H00020000)'dip 18
        .AddChk 205,340,140,Array("Attract features",&H20000000)'dip 30
        .AddLabel 50,360,300,20,"After hitting OK, press F3 to reset game with new settings."
    End With
    Dim extra
    extra = Controller.Dip(4) + Controller.Dip(5)*256
    extra = vpmDips.ViewDipsExtra(extra)
    Controller.Dip(4) = extra And 255
    Controller.Dip(5) = (extra And 65280)\256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")
 
 
 Dim N1,O1, Light
 N1=0:O1=0:
 Set LampCallback=GetRef("UpdateMultipleLamps")
 
Sub UpdateMultipleLamps
  if startGame.enabled=0 then
    N1=Controller.Lamp(0) 'Game Over triggers match and BIP
    If N1 Then
        BIPBox.text="BALL IN PLAY"
        NTMBox.text=""
        GOBox.text=""
      Else
        BIPBox.text=""
        NTMBox.text="NUMBER TO MATCH"
        GOBox.text="GAME OVER"
    End If
 
    N1=Controller.Lamp(1) 'Tilt
    If N1 Then
        TILTBox.text="TILT"
      Else
        TILTBox.text=""
    End If
 
    N1=Controller.Lamp(3) 'Shoot Again
    If N1 Then
        ShootAgainBox.text="SHOOT AGAIN"
      Else
        ShootAgainBox.text=""
    End If
 
    N1=Controller.Lamp(10) 'HIGH SCORE TO DATE
    If N1 Then
        HStoDateBox.text="HIGH SCORE TO DATE"
      Else
        HStoDateBox.text=""
    End If
  end if
End Sub
 
'**********************************************************************************************************
 
'Map lights to an array
'**********************************************************************************************************
'Set Lights(0) = l0  ' Game Over Relay
'Set Lights(1) = l1 ' Tilt relay
'Set Lights(2) = l2 ' Coin Lockout Relay
Set Lights(3) = l3 ' Shoot Again Playfield and backglass
Set Lights(4) = l4 ' first player
Set Lights(5) = l5 ' second player
Set Lights(6) = l6 'third player
Set Lights(7) = l7 ' fourth player
 
'Set Lights(10) = l10 ' high game to date - backbox
'Set Lights(11) = l11 ' game over - backbox
Set Lights(12) = l12
Set Lights(13) = l13
Set Lights(14) = l14
Set Lights(15) = l15
Set Lights(16) = l16
Set Lights(17) = l17
Set Lights(18) = l18
Set Lights(19) = l19
Set Lights(20) = l20
Set Lights(21) = l21
Set Lights(22) = l22
Set Lights(23) = l23
Set Lights(24) = l24
Set Lights(25) = l25
Set Lights(26) = l26
Set Lights(27) = l27
Set Lights(28) = l28
Set Lights(29) = l29
Set Lights(30) = l30
Set Lights(31) = l31
Set Lights(32) = l32
Set Lights(33) = l33
Set Lights(34) = l34
Set Lights(35) = L35
Set Lights(36) = l36
Set Lights(37) = l37
Set Lights(38) = l38
Set Lights(39) = l39
Set Lights(40) = l40
Set Lights(41) = l41
Set Lights(42) = l42
Set Lights(43) = l43
Set Lights(44) = l44
Lights(45) = array(l45,l45b)
Lights(46) = array(l46,l46b)
Lights(47) = array(l47,l47b)
 
Lights(49) = array(l49,l49b)
Lights(50) = array(l50,l50b)
Set Lights(51) = l51
 
 
 
Dim Digits(32)
 
'Score displays
 
Digits(0)=Array(a1,a2,a3,a4,a5,a6,a7,n,a8)
Digits(1)=Array(a9,a10,a11,a12,a13,a14,a15,n,a16)
Digits(2)=Array(a17,a18,a19,a20,a21,a22,a23,n,a24)
Digits(3)=Array(a25,a26,a27,a28,a29,a30,a31,n,a32)
Digits(4)=Array(a33,a34,a35,a36,a37,a38,a39,n,a40)
Digits(5)=Array(a41,a42,a43,a44,a45,a46,a47,n,a48)
Digits(6)=Array(a49,a50,a51,a52,a53,a54,a55,n,a56)
Digits(7)=Array(a57,a58,a59,a60,a61,a62,a63,n,a64)
Digits(8)=Array(a65,a66,a67,a68,a69,a70,a71,n,a72)
Digits(9)=Array(a73,a74,a75,a76,a77,a78,a79,n,a80)
Digits(10)=Array(a81,a82,a83,a84,a85,a86,a87,n,a88)
Digits(11)=Array(a89,a90,a91,a92,a93,a94,a95,n,a96)
Digits(12)=Array(a97,a98,a99,a100,a101,a102,a103,n,a104)
Digits(13)=Array(a105,a106,a107,a108,a109,a110,a111,n,a112)
Digits(14)=Array(a113,a114,a115,a116,a117,a118,a119,n,a120)
Digits(15)=Array(a121,a122,a123,a124,a125,a126,a127,n,a128)
Digits(16)=Array(a129,a130,a131,a132,a133,a134,a135,n,a136)
Digits(17)=Array(a137,a138,a139,a140,a141,a142,a143,n,a144)
Digits(18)=Array(a145,a146,a147,a148,a149,a150,a151,n,a152)
Digits(19)=Array(a153,a154,a155,a156,a157,a158,a159,n,a160)
Digits(20)=Array(a161,a162,a163,a164,a165,a166,a167,n,a168)
Digits(21)=Array(a169,a170,a171,a172,a173,a174,a175,n,a176)
Digits(22)=Array(a177,a178,a179,a180,a181,a182,a183,n,a184)
Digits(23)=Array(a185,a186,a187,a188,a189,a190,a191,n,a192)
 
'Ball in Play and Credit displays
 
Digits(26)=Array(e00,e01,e02,e03,e04,e05,e06,n)
Digits(27)=Array(e10,e11,e12,e13,e14,e15,e16,n)
Digits(24)=Array(f00,f01,f02,f03,f04,f05,f06,n)
Digits(25)=Array(f10,f11,f12,f13,f14,f15,f16,n)
 
 
Sub DisplayTimer_Timer
    Dim ChgLED,ii,num,chg,stat, obj
    ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
        If not b2son Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
            if (num < 28 ) then
                For Each obj In Digits(num)
                    If chg And 1 Then obj.State = stat And 1
                    chg = chg\2 : stat = stat\2
                Next
            else
 
            end if
        next
        end if
    end if
End Sub
 
 
 
Sub B2SCommand(nr, state)
    If cController = 3 Then
        Controller.B2SSetData nr, state
    End If
End Sub
 
Sub DOF(dofevent, dofstate)
    '*******Use DOF 1**, 1 to activate a ledwiz output*******************
    '*******Use DOF 1**, 0 to deactivate a ledwiz output*****************
    '*******Use DOF 1**, 2 to pulse a ledwiz output**********************
    If cController > 2 Then
        If dofstate = 2 Then
            Controller.B2SSetData dofevent, 1
            Controller.B2SSetData dofevent, 0
        Else
            Controller.B2SSetData dofevent, dofstate
        End If
    End If
End Sub