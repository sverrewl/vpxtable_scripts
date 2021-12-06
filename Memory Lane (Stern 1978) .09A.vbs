Option Explicit
Randomize

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
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


Const cGameName="memlane",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01000100", "Bally.VBS", 1.2
Dim DesktopMode: DesktopMode = Table1.ShowDT


If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)  = "vpmSolSound ""10"","
SolCallback(2)  = "vpmSolSound ""100"","
SolCallback(3)  = "vpmSolSound ""1000"","
SolCallback(4)  = "vpmSolSound ""10000"","
SolCallback(6)  = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(7)  = "bsTrough.SolOut"
SolCallback(8)  = "bsSaucer1.SolOut" 'Top Kicker
SolCallback(9)  = "bsSaucer.SolOut" 'Side Kicker
SolCallback(10) = "dtLL.SolDropUp" 'Drop Targets
SolCallback(19) = "vpmNudge.SolGameOn"


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next
'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough, dtLL, bsSaucer, bsSaucer1

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Memory Lane (Stern 1978)"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = True

  vpmNudge.TiltSwitch = 7
  vpmNudge.Sensitivity = 1
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot, RightSlingshot)


  Set bsTrough = New cvpmBallStack ' Trough handler
    bsTrough.InitSw 0,8,0,0,0,0,0,0
    bsTrough.InitKick Ballrelease, 90, 5
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls = 1

  'Drop Targets
  Set dtLL = new cvpmDropTarget
    dtLL.InitDrop Array(Sw1, Sw2, Sw3), Array(1, 2, 3)
    dtLL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  Set bsSaucer=New cvpmBallStack
    'bsSaucer.InitSaucer sw39,39,180,5 'NEED it to go down ??????
    bsSaucer.InitSaucer sw39,39,105,15
        bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  Set bsSaucer1=New cvpmBallStack
    'bsSaucer1.InitSaucer sw40,40,180,5
        bsSaucer1.InitSaucer sw40,40,160,15
    bsSaucer1.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
End Sub

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain",drain,1 : End Sub
Sub sw39_Hit:bsSaucer.AddBall 0 : playsoundAtVol "popper_ball",sw39,VolKick: End Sub
Sub sw40_Hit:bsSaucer1.AddBall 0 : playsoundAtVol "popper_ball",sw40,VolKick: End Sub

'Drop Targets
 Sub Sw1_Dropped:dtLL.Hit 1 :End Sub
 Sub Sw2_Dropped:dtLL.Hit 2 :End Sub
 Sub Sw3_Dropped:dtLL.Hit 3 :End Sub

'Wire Triggers
Sub SW4_Hit:Controller.Switch(4)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW4_unHit:Controller.Switch(4)=0:End Sub
Sub SW5_Hit:Controller.Switch(5)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW5_unHit:Controller.Switch(5)=0:End Sub
Sub SW12_Hit:Controller.Switch(12)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW12_unHit:Controller.Switch(12)=0:End Sub
Sub SW13_Hit:Controller.Switch(13)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW13_unHit:Controller.Switch(13)=0:End Sub

'Star Triggers
Sub SW20_Hit:Controller.Switch(20)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW20_unHit:Controller.Switch(20)=0:End Sub
Sub SW21_Hit:Controller.Switch(21)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW21_unHit:Controller.Switch(21)=0:End Sub
Sub SW22_Hit:Controller.Switch(22)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW22_unHit:Controller.Switch(22)=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW23_unHit:Controller.Switch(23)=0:End Sub
Sub SW24_Hit:Controller.Switch(24)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW24_unHit:Controller.Switch(24)=0:End Sub
Sub SW28_Hit:Controller.Switch(28)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW28_unHit:Controller.Switch(28)=0:End Sub
Sub SW29_Hit:Controller.Switch(29)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW29_unHit:Controller.Switch(29)=0:End Sub
Sub SW30_Hit:Controller.Switch(30)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW30_unHit:Controller.Switch(30)=0:End Sub
Sub SW31_Hit:Controller.Switch(31)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW31_unHit:Controller.Switch(31)=0:End Sub
Sub SW32_Hit:Controller.Switch(32)=1 : playsoundAtVol"rollover" , Activeball, 1: End Sub
Sub SW32_unHit:Controller.Switch(32)=0:End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(35) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper1, VolBump: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(36) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper2, VolBump: End Sub

'Stand Up Targets
Sub sw25_hit:vpmTimer.pulseSw 25 : End Sub
Sub sw26_hit:vpmTimer.pulseSw 26 : End Sub
Sub sw27_hit:vpmTimer.pulseSw 27 : End Sub

'Spinners
Sub sw14_Spin:vpmTimer.PulseSw 14 : playsoundAtVol"fx_spinner",sw14,VolSpin : End Sub
Sub sw15_Spin:vpmTimer.PulseSw 15 : playsoundAtVol"fx_spinner",sw15,VolSpin : End Sub

 'Scoring Rubber
Sub sw18a_hit:vpmTimer.pulseSw 18 : playsoundAtVol"flip_hit_3" , Activeball, 1: End Sub
Sub sw18b_hit:vpmTimer.pulseSw 18 : playsoundAtVol"flip_hit_3" , Activeball, 1: End Sub
Sub sw18c_hit:vpmTimer.pulseSw 18 : playsoundAtVol"flip_hit_3" , Activeball, 1: End Sub
Sub sw18d_hit:vpmTimer.pulseSw 18 : playsoundAtVol"flip_hit_3" , Activeball, 1: End Sub

Sub sw19a_hit:vpmTimer.pulseSw 19 : playsoundAtVol"flip_hit_3" , Activeball, 1: End Sub
Sub sw19b_hit:vpmTimer.pulseSw 19 : playsoundAtVol"flip_hit_3" , Activeball, 1: End Sub
Sub sw19c_hit:vpmTimer.pulseSw 19 : playsoundAtVol"flip_hit_3" , Activeball, 1: End Sub
Sub sw19d_hit:vpmTimer.pulseSw 19 : playsoundAtVol"flip_hit_3" , Activeball, 1: End Sub

Sub sw38_hit:vpmTimer.pulseSw 38 : playsoundAtVol"flip_hit_3" , Activeball, 1: End Sub
Sub sw38a_hit:vpmTimer.pulseSw 38 : playsoundAtVol"flip_hit_3" , Activeball, 1: End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 5 'lamp fading speed
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


Sub UpdateLamps()
NfadeL 1, L1     'PinLight1
NfadeL 2, L2     'PinLight5
NfadeL 3, L3     'PinLight9
NfadeL 4, L4     'StrikeLight1
NfadeL 5, L5     'StrikeLight5
NfadeL 6, L6     'TopSpotLight
NfadeL 7, L7     'RightSpinnerLight
NfadeL 11, L11   'ShootAgainLight
NfadeL 17, L17   'PinLight2
NfadeL 18, L18   'PinLight6
NfadeL 19, L19   'PinLight10
NfadeL 20, L20   'StrikeLight2
NfadeL 21, L21   'StrikeLight6
NfadeL 22, L22   'LeftSpotLight
NfadeL 23, L23   'LeftSpinnerLight
NfadeL 26, L26   'Light2X
NfadeL 33, L33   'PinLight3
NfadeL 34, L34   'PinLight7
NfadeL 35, L35   'Left Bumper 1
NfadeL 36, L36   'StrikeLight3
NfadeL 37, L37   'StrikeLight7
NfadeL 38, L38   'RightSpotLight
NfadeL 39, L39   'RightOutlaneLight
NfadeL 42, L42   'Light 3X
NfadeL 43, L43   'DoubleBonusLight
NfadeL 49, L49   'PinLight4
NfadeL 50, L50   'PinLight8
NfadeL 51, L51   'Right bumper 2
NfadeL 52, L52   'StrikeLight4
NfadeL 55, L55   'LeftOutlaneLight
NfadeL 56, L56   'TargetLight
NfadeL 59, L59   'TripleBonusLight

'Unknown ones for testing to identify what lights extra ball lanes
NfadeL 8, L8                              'unknown
NfadeL 9, L9                              'unknown
NfadeL 10, L10                            'unknown

NfadeL 12, L12                            'unknown
NfadeL 13, L13    'BallInPlayLight
NfadeL 14, L14    'OneCanPlayLight
FadeReel 15, L15    'Player 1
NfadeL 16, L16                            'Not used

NfadeL 24, L24                            'unknown
NfadeL 25, L25                            'unknown

NfadeL 27, L27   'MatchLight
NfadeL 28, L28                            'unknown
NfadeL 29, L29    'HighScoreLight
NfadeL 30, L30    'TwoCanPlayLight
FadeReel 31, L31    'Player 2
NfadeL 32, L32                           'Not used

NfadeL 40, L40                            'unknown
NfadeL 41, L41                            'unknown

NfadeL 44, L44                            'unknown
FadeReel 45, L45     'Game Over Light
NfadeL 46, L46     'ThreeCanPlayLight
FadeReel 47, L47     'Player 3
NfadeL 48, L48                           'Not used

NfadeL 53, L53     'StrikeLight8 only used internally
NfadeL 54, L54                            'unknown

NfadeL 57, L57                            'unknown
NfadeL 58, L58                            'unknown

NfadeL 60, L60                            'unknown
NfadeL 61, L61     'TiltLight
NfadeL 62, L62     'FourCan Play
FadeReel 63, L63     'Player 4


End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
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
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
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

 'Reels
Sub FadeReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 0:FadingLevel(nr) = 3
        Case 5:reel.Visible = 1:FadingLevel(nr) = 1
    End Select
End Sub

 'Inverted Reels
Sub FadeIReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 1:FadingLevel(nr) = 3
        Case 5:reel.Visible = 0:FadingLevel(nr) = 1
    End Select
End Sub

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(28)
' 1st Player
'Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
'Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
'Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
'Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
'Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
'Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)

' 2nd Player
'Digits(6) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
'Digits(7) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
'Digits(8) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
'Digits(9) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
'Digits(10) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
'Digits(11) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)

' 3rd Player
'Digits(12) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
'Digits(13) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
'Digits(14) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
'Digits(15) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
'Digits(16) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
'Digits(17) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)

' 4th Player
'Digits(18) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
'Digits(19) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
'Digits(20) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
'Digits(21) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
'Digits(22) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
'Digits(23) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)

' Credits
'Digits(24) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
'Digits(25) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
'Digits(26) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
'Digits(27) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 28) then
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


'**********************************************************************************************************
'**********************************************************************************************************




'Stern Memory Lane
'added by Inkochnito
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"Memory Lane - DIP switches"
    .AddChk 2,10,180,Array("Match feature",&H00100000)'dip 21
    .AddChk 205,10,115,Array("Credits display",&H00080000)'dip 20
    .AddFrame 2,30,190,"Maximum credits",&H00070000,Array("5 credits",0,"10 credits",&H00010000,"15 credits",&H00020000,"20 credits",&H00030000,"25 credits",&H00040000,"30 credits",&H00050000,"35 credits",&H00060000,"40 credits",&H00070000)'dip 17&18&19
    .AddFrame 2,160,190,"High game to date",49152,Array("points",0,"1 free game",&H00004000,"2 free games",32768,"3 free games",49152)'dip 15&16
    .AddFrame 2,235,190,"Special award",&HC0000000,Array("100,000 points",0,"free ball",&H40000000,"free game",&H80000000,"free ball and free game",&HC0000000)'dip 31&32
    .AddFrame 2,310,190,"High score feature",&H00000020,Array("extra ball",0,"replay",&H00000020)'dip 6
    .AddFrame 205,30,190,"Balls per game",&H00000040,Array("3 balls",0,"5 balls",&H00000040)'dip 7
    .AddFrame 205,76,190,"Top eject pocket 3X lamp",&H00002000,Array("no alternation",0,"alternates on 10 points",&H00002000)'dip 14
    .AddFrame 205,122,190,"Number of Strikes needed",&H00200000,Array ("3 Strikes",0,"4 Strikes",&H00200000,"5 Strikes",&H00400000,"6 Strikes",&H00600000)'dip 22&23
    .AddFrame 205,216,190,"Strike lamp lights",&H00800000,Array("every multiple of 20",0,"every multiple of 40",&H00800000)'dip 24
    .AddFrame 205,263,190,"Extra ball",&H02000000,Array("no extra ball (bypass)",0,"award extra ball",&H02000000)'dip 26
    .AddFrame 205,310,190,"Melody option",&H00000080,Array("2 tones only",0,"full melody",&H00000080)'dip 8
    .AddLabel 50,382,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips = GetRef("editDips")



' *********************************************************************
' *********************************************************************

          'Start of VPX call back Functions

' *********************************************************************
' *********************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 36
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 37
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
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

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple PlaysoundAtVol with volume and paning
' depending of the speed of the collision.

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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

