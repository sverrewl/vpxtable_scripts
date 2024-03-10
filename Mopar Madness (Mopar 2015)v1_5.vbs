'********Black Pyramid********
  '***********Bally 1984*********
  '********By unclewilly*********
 'light numbers based on script by nico193

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
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
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.



 Const movieSpeed= 190 'msecs

    'Option Explicit
    'Randomize

    LoadVPM "01560000", "Bally.VBS", 3.26

    Sub LoadVPM(VPMver, VBSfile, VBSver)
        On Error Resume Next
        If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
        ExecuteGlobal GetTextFile(VBSfile)
        If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
        Set Controller = CreateObject("b2s.server")
        If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
        If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
        If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
        On Error Goto 0
    End Sub

  'Variables
  Dim STPos, STDir, xx
  STPos=0:STDir=1
  Dim bsTrough, bsSaucer, dtBank, plungerIM
  Dim Bump1, Bump2
  Dim cGameName

Dim AttractVideo
AttractVideo = Array("attract ", 588)

  Const UseSolenoids = 1
  Const UseLamps = 0
  Const UseGI = 0
  Const UseSync = 0
  Const HandleMech = 0

  'Standard Sounds
   Const SSolenoidOn = "Solenoid"
   Const SSolenoidOff = ""
   Const SCoin = "CoinIn"


  'Table Init
    Sub table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

        With Controller
        cGameName = "blakpyra"
            .GameName = cGameName
            .SplashInfoLine = "Mopar Madness, R&R 1984" & vbNewLine & "by unclewilly vp9"
            .HandleMechanics = 0
            .HandleKeyboard = 0
            .ShowDMDOnly = 1
            .ShowFrame = 0
            .ShowTitle = 0
            .Hidden = 0
        'LedStartup ' Initialize LED Panel Display
            If Err Then MsgBox Err.Description
        End With
        On Error Goto 0
        Controller.Run

  'Nudging
        vpmNudge.TiltSwitch = 15
        vpmNudge.Sensitivity = 4
        vpmNudge.TiltObj = Array(sw4, sw3, LeftSlingshot, RightSlingShot)

   '**Trough
        Set bsTrough = New cvpmBallStack
        With bsTrough
            .InitSw 0,8,0,0,0,0,0,0
        .InitKick ballrelease, 90, 4
            .InitExitSnd "ballrelease", "Solenoid"
        .BallImage = "ballDark"
           .Balls = 1
        End With

  'Saucer
       Set bsSaucer = New cvpmBallStack
      With bsSaucer
          .InitSaucer sw7, 7, 200, 10
          .InitExitSnd "ballrelease", "Solenoid"
      End With

  '**Targets
        set dtBank = new cvpmdroptarget
        With dtBank
            .InitDrop Array(sw30, sw31, sw32), Array(30, 31, 32)
            .Initsnd "droptargetR", "resetdropR"
            .CreateEvents "dtBank"
        End With

      '**Main Timer init
        PinMAMETimer.Interval = PinMAMEInterval
        PinMAMETimer.Enabled = 1


  'Init Slings & Bumper Rings
    sw3Ra.IsDropped=1:sw3Rb.IsDropped=1:sw3Rc.IsDropped=1
    sw4Ra.IsDropped=1:sw4Rb.IsDropped=1:sw4Rc.IsDropped=1

  'Init Standup Targets and rollovers
  ' sw20a.IsDropped=1:sw20wa.IsDropped=1:sw21a.IsDropped=1:sw21wa.IsDropped=1:sw19a.IsDropped=1:sw19wa.IsDropped=1
   'sw18a.IsDropped=1:sw18wa.IsDropped=1

   sw17a.IsDropped=1:sw24a.IsDropped=1:sw29a.IsDropped=1
   sw22a.IsDropped=1:sw23a.IsDropped=1:sw28a.IsDropped=1

  'init swing target
    'For each xx in sw5a:xx.IsDropped=1:Next
  'For xx = 1 to 10
    'sw5(xx).IsDropped=1
  'Next

  'Init GI
    UpdateGI 0

attractdelay.Interval = 5000
  attractdelay.Enabled = 1

   End Sub

Sub attractdelay_Timer
  Animation AttractVideo(0), AttractVideo(1), 0
  attractdelay.Enabled = 0 'turn off timer
End Sub

    Sub table1_Paused:Controller.Pause = 1:End Sub
    Sub table1_unPaused:Controller.Pause = 0:End Sub
 '***Keyboard handler

Sub table1_KeyUp(ByVal KeyCode)
    If vpmKeyUp(KeyCode) Then Exit Sub
    If KeyCode=PlungerKey Then PlaySoundAtVol"Plunger",plunger,1:Plunger.Fire
End Sub

Sub table1_KeyDown(ByVal KeyCode)
    If vpmKeyDown(KeyCode) Then Exit Sub
    If KeyCode=PlungerKey Then PlaySoundAtVol"PullbackPlunger",plunger,1:Plunger.Pullback
End Sub

    dim PCount:Pcount = 0


  'Switches
   'Slings & Rubbers
    'Dim LStep, RStep




' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
'   Sub LeftSlingshot_Slingshot:lsling.IsDropped=0:PlaySoundAtVol "slingshotL",activeball, 1:vpmTimer.PulseSw 1:LStep=0:ME.TimerEnabled = 1:End Sub
    Sub LeftSlingshot_Timer
        Select Case LStep
           Case 0:'Pause
            Case 1: 'pause
            Case 2:lsling.IsDropped = 1:lsling2.IsDropped = 0
            Case 3:lsling2.IsDropped = 1:lsling3.IsDropped = 0
            Case 4:lsling3.IsDropped = 1:Me.TimerEnabled = 0
        End Select
        LStep = LStep + 1
    End Sub

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
'   Sub RightSlingShot_Slingshot:rsling.IsDropped=0:PlaySoundAtVol "slingshotR",activeball,1:vpmTimer.PulseSw 2:RStep=0:ME.TimerEnabled = 1:End Sub
'    Sub RightSlingShot_Timer
'        Select Case RStep
'            Case 0:'Pause
'            Case 1: 'pause
'            Case 2:rsling.IsDropped = 1:rsling2.IsDropped = 0
'            Case 3:rsling2.IsDropped = 1:rsling3.IsDropped = 0
'            Case 4:rsling3.IsDropped = 1:Me.TimerEnabled = 0
'        End Select
'     End Sub

   Sub sw12_Hit():PlaySoundAtVol "rubber",activeball,1:vpmTimer.PulseSw 12:End Sub
   Sub sw12a_Hit():PlaySoundAtVol "rubber",activeball,1:vpmTimer.PulseSw 12:End Sub

   Sub sw4_Hit:vpmTimer.PulseSw 4:PlaySoundAtVol "BumperRight",sw4,VolBump:End Sub
   Sub sw3_Hit:vpmTimer.PulseSw 3:PlaySoundAtVol "BumperLeft",sw3,VolBump:End Sub



  'Drain
   Sub Drain_Hit():bsTrough.AddBall Me:PlaySoundAtVol "Drain", drain,1: EndMusic:End Sub
   Sub sw7_Hit():bsSaucer.AddBall 0:PlaySoundAtVol "kicker_enter",sw7,1:End Sub

  'Rollovers
    Sub sw17_Hit:sw17a.IsDropped = 0:Controller.Switch(17) = 1:PlaySoundAtVol "rollover",ActiveBall, 1:End Sub
    Sub sw17_UnHit:sw17a.IsDropped = 1:Controller.Switch(17) = 0:End Sub
    Sub sw24_Hit:sw24a.IsDropped = 0:Controller.Switch(24) = 1:PlaySoundAtVol "rollover",ActiveBall, 1:End Sub
    Sub sw24_UnHit:sw24a.IsDropped = 1:Controller.Switch(24) = 0:End Sub
    Sub sw22_Hit:sw22a.IsDropped = 0:Controller.Switch(22) = 1:PlaySoundAtVol "rollover",ActiveBall, 1:End Sub
    Sub sw22_UnHit:sw22a.IsDropped = 1:Controller.Switch(22) = 0:End Sub
    Sub sw23_Hit:sw23a.IsDropped = 0:Controller.Switch(23) = 1:PlaySoundAtVol "rollover",ActiveBall, 1:End Sub
    Sub sw23_UnHit:sw23a.IsDropped = 1:Controller.Switch(23) = 0:End Sub
    Sub sw29_Hit:sw29a.IsDropped = 0:Controller.Switch(29) = 1:PlaySoundAtVol "rollover",ActiveBall, 1:End Sub
    Sub sw29_UnHit:sw29a.IsDropped = 1:Controller.Switch(29) = 0:End Sub
    Sub sw28_Hit:sw28a.IsDropped = 0:Controller.Switch(28) = 1:PlaySoundAtVol "rollover",ActiveBall, 1:End Sub
    Sub sw28_UnHit:sw28a.IsDropped = 1:Controller.Switch(28) = 0:End Sub

    Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAtVol "rollover",ActiveBall, 1:End Sub
    Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub
    Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAtVol "rollover",ActiveBall, 1:End Sub
    Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

  'Standup Targets
    Sub sw20_Hit:vpmTimer.PulseSw 20:sw20.IsDropped = 1:sw20w.IsDropped = 1:sw20a.IsDropped = 0:sw20wa.IsDropped = 0:Me.TimerEnabled = 1:PlaySoundAtVol "target",ActiveBall, 1:End Sub
    Sub sw20_Timer:sw20.IsDropped = 0:sw20w.IsDropped = 0:sw20a.IsDropped = 1:sw20wa.IsDropped = 1:Me.TimerEnabled = 0:End Sub
    Sub sw21_Hit:vpmTimer.PulseSw 21:sw21.IsDropped = 1:sw21w.IsDropped = 1:sw21a.IsDropped = 0:sw21wa.IsDropped = 0:Me.TimerEnabled = 1:PlaySoundAtVol "target",ActiveBall, 1:End Sub
    Sub sw21_Timer:sw21.IsDropped = 0:sw21w.IsDropped = 0:sw21a.IsDropped = 1:sw21wa.IsDropped = 1:Me.TimerEnabled = 0:End Sub
    Sub sw18_Hit:vpmTimer.PulseSw 18:sw18.IsDropped = 1:sw18w.IsDropped = 1:sw18a.IsDropped = 0:sw18wa.IsDropped = 0:Me.TimerEnabled = 1:PlaySoundAtVol "target",ActiveBall, 1:End Sub
    Sub sw18_Timer:sw18.IsDropped = 0:sw18w.IsDropped = 0:sw18a.IsDropped = 1:sw18wa.IsDropped = 1:Me.TimerEnabled = 0:End Sub
    Sub sw19_Hit:vpmTimer.PulseSw 19:sw19.IsDropped = 1:sw19w.IsDropped = 1:sw19a.IsDropped = 0:sw19wa.IsDropped = 0:Me.TimerEnabled = 1:PlaySoundAtVol "target",ActiveBall, 1:End Sub
    Sub sw19_Timer:sw19.IsDropped = 0:sw19w.IsDropped = 0:sw19a.IsDropped = 1:sw19wa.IsDropped = 1:Me.TimerEnabled = 0:End Sub

 '**Swinging target
  Sub sw5p0_Hit():STTimer.Enabled=0:sw5p0.IsDropped=1:sw5p0a.IsDropped=0:vpmTimer.PulseSw 5:PlaySoundAtVol "target",ActiveBall, 1:Me.TimerEnabled = 1:End Sub
  Sub sw5p0_Timer():STTimer.Enabled=1:sw5p0.IsDropped=0:sw5p0a.IsDropped=1:Me.TimerEnabled = 0:End Sub
  Sub sw5p1_Hit():STTimer.Enabled=0:sw5p1.IsDropped=1:sw5p1a.IsDropped=0:vpmTimer.PulseSw 5:PlaySoundAtVol "target",ActiveBall, 1:Me.TimerEnabled = 1:End Sub
  Sub sw5p1_Timer():STTimer.Enabled=1:sw5p1.IsDropped=0:sw5p1a.IsDropped=1:Me.TimerEnabled = 0:End Sub
  Sub sw5p2_Hit():STTimer.Enabled=0:sw5p2.IsDropped=1:sw5p2a.IsDropped=0:vpmTimer.PulseSw 5:PlaySoundAtVol "target",ActiveBall, 1:Me.TimerEnabled = 1:End Sub
  Sub sw5p2_Timer():STTimer.Enabled=1:sw5p2.IsDropped=0:sw5p2a.IsDropped=1:Me.TimerEnabled = 0:End Sub
  Sub sw5p3_Hit():STTimer.Enabled=0:sw5p3.IsDropped=1:sw5p3a.IsDropped=0:vpmTimer.PulseSw 5:PlaySoundAtVol "target",ActiveBall, 1:Me.TimerEnabled = 1:End Sub
  Sub sw5p3_Timer():STTimer.Enabled=1:sw5p3.IsDropped=0:sw5p3a.IsDropped=1:Me.TimerEnabled = 0:End Sub
  Sub sw5p4_Hit():STTimer.Enabled=0:sw5p4.IsDropped=1:sw5p4a.IsDropped=0:vpmTimer.PulseSw 5:PlaySoundAtVol "target",ActiveBall, 1:Me.TimerEnabled = 1:End Sub
  Sub sw5p4_Timer():STTimer.Enabled=1:sw5p4.IsDropped=0:sw5p4a.IsDropped=1:Me.TimerEnabled = 0:End Sub
  Sub sw5p5_Hit():STTimer.Enabled=0:sw5p5.IsDropped=1:sw5p5a.IsDropped=0:vpmTimer.PulseSw 5:PlaySoundAtVol "target",ActiveBall, 1:Me.TimerEnabled = 1:End Sub
  Sub sw5p5_Timer():STTimer.Enabled=1:sw5p5.IsDropped=0:sw5p5a.IsDropped=1:Me.TimerEnabled = 0:End Sub
  Sub sw5p6_Hit():STTimer.Enabled=0:sw5p6.IsDropped=1:sw5p6a.IsDropped=0:vpmTimer.PulseSw 5:PlaySoundAtVol "target",ActiveBall, 1:Me.TimerEnabled = 1:End Sub
  Sub sw5p6_Timer():STTimer.Enabled=1:sw5p6.IsDropped=0:sw5p6a.IsDropped=1:Me.TimerEnabled = 0:End Sub
  Sub sw5p7_Hit():STTimer.Enabled=0:sw5p7.IsDropped=1:sw5p7a.IsDropped=0:vpmTimer.PulseSw 5:PlaySoundAtVol "target",ActiveBall, 1:Me.TimerEnabled = 1:End Sub
  Sub sw5p7_Timer():STTimer.Enabled=1:sw5p7.IsDropped=0:sw5p7a.IsDropped=1:Me.TimerEnabled = 0:End Sub
  Sub sw5p8_Hit():STTimer.Enabled=0:sw5p8.IsDropped=1:sw5p8a.IsDropped=0:vpmTimer.PulseSw 5:PlaySoundAtVol "target",ActiveBall, 1:Me.TimerEnabled = 1:End Sub
  Sub sw5p8_Timer():STTimer.Enabled=1:sw5p8.IsDropped=0:sw5p8a.IsDropped=1:Me.TimerEnabled = 0:End Sub
  Sub sw5p9_Hit():STTimer.Enabled=0:sw5p9.IsDropped=1:sw5p9a.IsDropped=0:vpmTimer.PulseSw 5:PlaySoundAtVol "target",ActiveBall, 1:Me.TimerEnabled = 1:End Sub
  Sub sw5p9_Timer():STTimer.Enabled=1:sw5p9.IsDropped=0:sw5p9a.IsDropped=1:Me.TimerEnabled = 0:End Sub
  Sub sw5p10_Hit():STTimer.Enabled=0:sw5p10.IsDropped=1:sw5p10a.IsDropped=0:vpmTimer.PulseSw 5:PlaySoundAtVol "target",ActiveBall, 1:Me.TimerEnabled = 1:End Sub
  Sub sw5p10_Timer():STTimer.Enabled=1:sw5p10.IsDropped=0:sw5p10a.IsDropped=1:Me.TimerEnabled = 0:End Sub


   Sub Gate1_Hit():PlaysoundAtVol "Gate",ActiveBall, 1:End Sub
   Sub Gate2_Hit():PlaysoundAtVol "Gate",ActiveBall, 1:End Sub
   Sub Gate3_Hit():PlaysoundAtVol "Gate",ActiveBall, 1:End Sub
   Sub tdl_Hit():ActiveBall.Image = "ball":End Sub
   Sub tld_Hit():ActiveBall.Image = "balldark":End Sub

 '****Solenoids

    SolCallback(15) = "vpmSolSound ""Knocker"","
    SolCallback(13) = "dtBank.SolDropUp"
    SolCallback(14) = "bsTrough.SolOut"
    SolCallback(12) = "bsSaucer.SolOut"
    SolCallback(17) = "SolDiv"
    SolCallback(19) = "vpmNudge.SolGameOn"

 'Based on jps diverter code
  Sub SolDiv(Enabled)
     If Enabled Then
         Div.RotateToEnd
         Diva.RotateToEnd
         PlaySoundAtVol "Diverter", Diva, 1
     Else
         vpmTimer.AddTimer 500, "ShutGate"
     End If
 End Sub

  Sub ShutGate(swNo)
     Div.RotateToStart
     Diva.RotateToStart
     PlaySoundAtVol "Diverter", Diva, 1
 End Sub

  '****FlipperSubs
    SolCallback(sLRFlipper) = "SolRFlipper"
    SolCallback(sLLFlipper) = "SolLFlipper"

    Sub SolLFlipper(Enabled)
        If Enabled Then
            PlaySoundAtVol "LFliper", LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
        Else
            PlaySoundAtVol "LFliperD", LeftFlipper, VolFlip:LeftFlipper.RotateToStart
        End If
    End Sub

    Sub SolRFlipper(Enabled)
        If Enabled Then
            PlaySoundAtVol "RFliper", RightFlipper, VolFlip:RightFlipper.RotateToEnd
        Else
            PlaySoundAtVol "RFliperD",RightFlipper, VolFlip:RightFlipper.RotateToStart
        End If
    End Sub

 Sub STTimer_Timer()
  sw5(STPos).IsDropped=1
  sw5(StPos+STDir).IsDropped=0
  STPos=STPos+STDir
  If STPos=10 then STDir=-1
  If StPos=0 then STDir=1
 End Sub

Dim BIP
BIP = 0

' Sub LeftSlingshot_Slingshot()
'   PlaySoundAtVol "fx_bumper4", ActiveBall, 1
' End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 1
    PlaySoundAtVol "left_slingshot", sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  gi10.State = 0':Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:gi10.State = 1':Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 2
    PlaySoundAtVol "right_slingshot", sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  'gi3.State = 0:Gi4.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0':gi3.State = 1:Gi4.State = 1
    End Select
    LStep = LStep + 1
End Sub


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        'Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****GI Lights On
'dim xx

'******************
'   GI effects
'******************

Dim OldGiState
OldGiState = 2 'start witht he Gi off

Sub GIUpdateTimer_Timer
Dim tmp, obj
tmp = Getballs
If UBound(tmp) <> OldGiState Then
OldGiState = Ubound(tmp)
If UBound(tmp) = 1 Then ' since we have 2 captive balls and 1 for the car animation, then Ubound will show 2, so no balls on the table then turn off gi
For each obj in aGILights
  obj.State = 0
Next
Else
For each obj in aGILights
  obj.State = 1
Next
End If
End If
End Sub



 Sub UpdateGI(enabled)
    Dim xx
  If Enabled then

For each xx in GI:xx.State = 1: Next
    'GIB1.State=1
    'GIB2.State=1
    'GIB3.State=1
    'GIB4.State=1
    'GIB5.State=1
    'GIB6.State=1
    'GIB7.State=1
    'GIB8.State=1
    'GIB9.State=1
    'GI.State=1
    'Light1.State=1
    'Light2.State=1
    'GI1.IsDropped=0
    'GI2.IsDropped=0
    'GI3.IsDropped=0
    'GI4.IsDropped=0
    'GI5.IsDropped=0
    'GI6.IsDropped=0
    'GI7.IsDropped=0
    'GI8.IsDropped=0
    'GI9.IsDropped=0
    'GI10.IsDropped=0
    'GI11.IsDropped=0
    'GI12.IsDropped=0
    'GI13.IsDropped=0
    'GIbw.IsDropped=0
  else
    For each xx in GI:  xx.State = 0:Next
    'GIB1.State=0
    'GIB2.State=0
    'GIB3.State=0
    'GIB4.State=0
    'GIB5.State=0
    'GIB6.State=0
    'GIB7.State=0
    'GIB8.State=0
    'GIB9.State=0
    'GI.State=0
    'Light1.State=0
    'Light2.State=0
    'GI1.IsDropped=1
    'GI2.IsDropped=1
    'GI3.IsDropped=1
    'GI4.IsDropped=1
    'GI5.IsDropped=1
    'GI6.IsDropped=1
    'GI7.IsDropped=1
    'GI8.IsDropped=1
    'GI9.IsDropped=1
    'GI10.IsDropped=1
    'GI11.IsDropped=1
    'GI12.IsDropped=1
    'GI13.IsDropped=1
    'GIbw.IsDropped=1
       GiTimer.Enabled=1
  end if
  End Sub



   Sub GiTimer_Timer
  LampTimer.Interval = 10
  LampTimer.Enabled = 1
  updateGi 1
  GiTimer.Enabled=0
  End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()           ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array

            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If

    UpdateLamps
End Sub

Sub UpdateLamps

    NFadeLm 2, B3L1
  NFadeL 2, B3L2
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 15, l15
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 19, l19
  NFadeL 9, l9
    NFadeL 10, l10
  NFadeL 11, l11
    NFadeL 12, l12
    'NFadeL 13, l13
    NFadeL 14, l14
    NFadeL 17, l17
    NFadeLm 18, B2L1
    NFadeL 18, B2L2
    NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    NFadeL 28, l28
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 30, l30
    NFadeL 31, l31
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 46, l46
    NFadeL 47, l47
    NFadeL 49, l49
    NFadeL 50, l50
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57
    NFadeL 58, l58
    NFadeL 60, l60
    NFadeL 62, l62
    NFadeL 63, l63
    NFadeL 65, l65
  NFadeL 66, l66

    NFadeL 81, l81
    NFadeL 82, l82
    NFadeL 83, l83
    NFadeL 97, l97
    NFadeL 98, l98
    NFadeL 99, l99
    NFadeL 113, l113
    NFadeL 115, l115


End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0      ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4     ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2   ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0       ' the minimum value when off, usually 0
        FlashLevel(x) = 0     ' the intensity of the flashers, usually from 0 to 1
    Next
    UpdateLamps
    UpdateLamps
    Updatelamps
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
  Next
    UpdateLamps
    UpdateLamps
    Updatelamps
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6         'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1         'ON
    Case 6,7.8: FadingLevel(nr) =FadingLevel(nr) +1       'wait
        Case 9:object.image = c:FadingLevel(nr) =FadingLevel(nr) +1 'fading...
    Case 10,11,12: FadingLevel(nr) =FadingLevel(nr) +1      'wait
        Case 13:object.image = d:FadingLevel(nr) = 0        'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6        'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1        'ON
    Case 6,7.8: FadingLevel(nr) =FadingLevel(nr) +1       'wait
        Case 9:object.SetValue 2:FadingLevel(nr) =FadingLevel(nr) +1 'fading...
    Case 10,11,12: FadingLevel(nr) =FadingLevel(nr) +1      'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0         'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub


''''''''''''''''''''''''''''''''''''''''''''
'''''''''Targets''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''

Dim TargetTR, TargetTL, TargetML, TargetMR, targetct

Sub loc8_Hit:targetct = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(5):PlaysoundAtVol("fx_target"),ActiveBall, 1:End Sub

Sub loc8_Timer()
           Select Case targetct
               Case 1:PTargetCenter.TransX = -5:targetct = 2
               Case 2:targetct = 3
               Case 3:PTargetCenter.TransX = 0:targetct = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub loc7_Hit:targetct = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(5):PlaysoundAtVol("fx_target"),ActiveBall, 1:End Sub

Sub loc7_Timer()
           Select Case targetct
               Case 1:PTargetCenter.TransX = -5:targetct = 2
               Case 2:targetct = 3
               Case 3:PTargetCenter.TransX = 0:targetct = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub loc6_Hit:targetct = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(5):PlaysoundAtVol("fx_target"),ActiveBall, 1:End Sub

Sub loc6_Timer()
           Select Case targetct
               Case 1:PTargetCenter.TransX = -5:targetct = 2
               Case 2:targetct = 3
               Case 3:PTargetCenter.TransX = 0:targetct = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub loc5_Hit:targetct = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(5):PlaysoundAtVol("fx_target"),ActiveBall, 1:End Sub

Sub loc5_Timer()
           Select Case targetct
               Case 1:PTargetCenter.TransX = -5:targetct = 2
               Case 2:targetct = 3
               Case 3:PTargetCenter.TransX = 0:targetct = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub loc4_Hit:targetct = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(5):PlaysoundAtVol("fx_target"),ActiveBall, 1:End Sub

Sub loc4_Timer()
           Select Case targetct
               Case 1:PTargetCenter.TransX = -5:targetct = 2
               Case 2:targetct = 3
               Case 3:PTargetCenter.TransX = 0:targetct = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub loc3_Hit:targetct = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(5):PlaysoundAtVol("fx_target"),ActiveBall, 1:End Sub

Sub loc3_Timer()
           Select Case targetct
               Case 1:PTargetCenter.TransX = -5:targetct = 2
               Case 2:targetct = 3
               Case 3:PTargetCenter.TransX = 0:targetct = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub loc2_Hit:targetct = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(5):PlaysoundAtVol("fx_target"),ActiveBall, 1:End Sub

       Sub loc2_Timer()
           Select Case targetct
               Case 1:PTargetCenter.TransX = -5:targetct = 2
               Case 2:targetct = 3
               Case 3:PTargetCenter.TransX = 0:targetct = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub loc1_Hit:targetct = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(5):PlaysoundAtVol("fx_target"),ActiveBall, 1:End Sub

       Sub loc1_Timer()
           Select Case targetct
               Case 1:PTargetCenter.TransX = -5:targetct = 2
               Case 2:targetct = 3
               Case 3:PTargetCenter.TransX = 0:targetct = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub S_Top_Y_Hit:TargetTR = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(19):PlaysoundAtVol("fx_target"),ActiveBall, 1:End Sub

       Sub S_Top_Y_Timer()
           Select Case TargetTR
               Case 1:PTargetTopRight.TransX = -5:TargetTR = 2
               Case 2:TargetTR = 3
               Case 3:PTargetTopRight.TransX = 0:TargetTR = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub S_Top_P_Hit:TargetTL = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(18):PlaysoundAtVol("fx_target"),ActiveBall, 1:End Sub

       Sub S_Top_P_Timer()
           Select Case TargetTL
               Case 1:PTargetTopLeft.TransX = -5:TargetTL = 2
               Case 2:TargetTL = 3
               Case 3:PTargetTopLeft.TransX = 0:TargetTL = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub S_L_Target_Hit:TargetML = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(20):PlaysoundAtVol("fx_target"),ActiveBall, 1:End Sub

       Sub S_L_Target_Timer()
           Select Case TargetML
               Case 1:PTargetMidLeft.TransX = -5:TargetML = 2
               Case 2:TargetML = 3
               Case 3:PTargetMidLeft.TransX = 0:TargetML = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub S_R_Target_Hit:TargetMR = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(21):PlaysoundAtVol("fx_target"),ActiveBall, 1:End Sub

       Sub S_R_Target_Timer()
           Select Case TargetMR
               Case 1:PTargetMidRight.TransX = -5:TargetMR = 2
               Case 2:TargetMR = 3
               Case 3:PTargetMidRight.TransX = 0:TargetMR = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Dim omgitmoves:omgitmoves = True
Sub swing_Timer()  'make a timer named targetmove with a interval of 10
        If omgitmoves = True and PTargetCenter.TransZ >= -70 then
                PTargetCenter.TransZ = PTargetCenter.TransZ + 1
        End If
        If omgitmoves = False and PTargetCenter.TransZ <= 70 then
                PTargetCenter.TransZ = PTargetCenter.TransZ - 1
        End If
        If PTargetCenter.TransZ = 70 then omgitmoves = False    'go back left
        If PTargetCenter.TransZ = -70 then omgitmoves = true    'go back right
End Sub

'''''''''''''''''''''''''''''''''
''''''''Center Target''''''''''''
'''''''''''''''''''''''''''''''''
Sub wallfollow_Timer()
  If PTargetCenter.transz <= -51 and PTargetCenter.transz >= -70 then loc1.isdropped = false else loc1.isdropped = true:End If
  If PTargetCenter.transz <= -31 and PTargetCenter.transz >= -50 then loc2.isdropped = false else loc2.isdropped = true:End If
  If PTargetCenter.transz <= -11 and PTargetCenter.transz >= -30 then loc3.isdropped = false else loc3.isdropped = true:End If
  If PTargetCenter.transz <= -1 and PTargetCenter.transz >= -10 then loc4.isdropped = false else loc4.isdropped = true:End If
  If PTargetCenter.transz <= 10 and PTargetCenter.transz >= 0 then loc5.isdropped = false else loc5.isdropped = true:End If
  If PTargetCenter.transz <= 30 and PTargetCenter.transz >= 11 then loc6.isdropped = false else loc6.isdropped = true:End If
  If PTargetCenter.transz <= 50 and PTargetCenter.transz >= 31 then loc7.isdropped = false else loc7.isdropped = true:End If
  If PTargetCenter.transz <= 70 and PTargetCenter.transz >= 51 then loc8.isdropped = false else loc8.isdropped = true:End If
End Sub


Sub Timer1_Timer
  If LampState(13)=1 Then
    If XLocation>7 Then XDir=0
    If XLocation<1 Then XDir=1
    'T(XLocation).IsDropped=1
    If XDir=1 Then XLocation=XLocation+1
    If XDir=0 Then XLocation=XLocation-1
    'T(XLocation).IsDropped=0
    'CheckXLocation   '''  Added to move center drop Primitive --- CP
  End If
End Sub

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub



Sub editDips
Dim vpmDips : Set vpmDips = New cvpmDips
With vpmDips
.AddForm 700,400,"table1 - DIP switches"
.AddFrame 2,0,190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits", &H01000000,"25 credits", &H02000000,"40 credits", &H03000000)'dip 25&26
.AddFrame 2,76,190,"Balls per game",&HC0000000,Array("2 balls",&HC0000000,"3 balls",0,"4 balls",&H80000000,"5 balls",&H40000000)'dip 31&32
.AddFrame 2,154,190,"Bonus special",&H00600000,Array("on after 120K",0,"on with 120K",&H00400000,"on with 60K",&H00600000)'dip 22&23
.AddChk 2,217,100,Array("Match feature",&H08000000)'dip 28
.AddChk 2,232,100,Array("Credits display",&H04000000)'dip 27
.AddChk 2,247,100,Array("Attract sound",&H20000000)'dip 30
.AddFrame 205,0,190,"Left lane X-ball build up",&H000000C0,Array("90000",0,"80000",&H00000040,"70000",&H00000080,"50000",&H000000C0)'dip 7&8
.AddFrame 205,76,190,"Bonus special per game",&H00000020,Array("only 1",0,"unlimited",&H00000020)'dip 6
.AddFrame 205,122,190,"M and I return lanes",&H00002000,Array("lanes separated",0,"lanes tied together",&H00002000)'dip 14
.AddFrame 205,168,190,"Left roll up lane",&H00100000,Array("20000 initially unlit",0,"20000 initially lit",&H00100000)'dip 21
.AddFrame 205,214,190,"Right lane 50000",&H00800000,Array("alternates",0,"stays on",&H00800000)'dip 24
.AddLabel 25,270,350,20,"Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
.AddLabel 25,290,350,20,"After hitting OK, press F3 to reset game with new settings."
.ViewDips
End With
End Sub
Set vpmShowDips = GetRef("editDips")


Dim musicNum
Sub BallRelease_Unhit
  If musicNum = 0 then PlayMusic "startitup.mp3" End If
    If musicNum = 1 then PlayMusic "mopar or nocar.mp3" End If
    If musicNum = 2 then PlayMusic "mopar madness.mp3" End If
  musicNum = (musicNum + 1) mod 3
End Sub

Sub Animation(name, numframes, loops)
  HSteps = numframes
  HPos=0
  Hname = name
  Hloops = loops
  posinc = 1
    holotimer.interval = movieSpeed
  holotimer.enabled=1
End Sub

Dim Hname, Hsteps, Hloops, Hpos, posinc
 Sub holotimer_timer()
    Dim imagename
    HPos=(HPos+posinc) mod AttractVideo(1)
    if HPos = 0 then HPos = 1
    if HPos < 10 then
      imagename = Hname & "00" & Hpos
    elseif HPos < 100 then
      imagename = Hname & "0" & Hpos
    else
      imagename = Hname & Hpos
    end if
    Moviewall1.image = imagename
    Moviewall2.image = imagename
debug.print imagename
end Sub

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub


Sub UpdateFlipperLogo_Timer
    LFLogo.objrotz = LeftFlipper.CurrentAngle + 179
    RFlogo.objrotz = RightFlipper.CurrentAngle + 179
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
  Vol = Csng(BallVel(ball) ^2 / 2000)
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

Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here bacuse of the timing
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  On Error Resume Next
  Controller.Pause = False
  Controller.Stop
End Sub

