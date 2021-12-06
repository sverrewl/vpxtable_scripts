Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="punk",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01120100","sys80.vbs",3.02

' Thalamus 2019 March : Improved directional sounds
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
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

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

SolCallback(1)="dtT.SolDropUp"
SolCallback(2)="dtL.SolDropUp"
SolCallback(3)="dtCL.SolDropUp"
SolCallback(4)="dtCR.SolDropUp"
SolCallback(5)="bsSaucer.SolOut"
SolCallback(6)="bsSaucer2.SolOut"
SolCallback(8)= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(9)= "SolTrough"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolTrough(Enabled)
  If Enabled Then
    Controller.Switch(36)=0
  End If
End Sub

Sub SolTrough2(Enabled)
  If Enabled Then
    BallRelease.CreateBall
    PlaySound SoundFX("ballrelease",DOFContactors)
    BallRelease.Kick 100,1
    BallsInPlay=BallsInPlay+1
    Controller.Switch(54)=0
  End If
End Sub

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsSaucer, bsSaucer2, dtL, dtT, dtCL, dtCR, BallsInPlay

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = ""&chr(13)&"You Suck"
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

  Controller.SolMask(0)=0
    vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 s

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=57
  vpmNudge.Sensitivity=2
  vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot,Bumper1,Bumper2,Bumper3,Bumper4,Bumper5)

  Set bsSaucer=New cvpmBallStack
    bsSaucer.InitSaucer sw5,5,180,5
    bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  Set bsSaucer2=New cvpmBallStack
    bsSaucer2.InitSaucer sw6,6,180,5
    bsSaucer2.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  set dtT=New cvpmDropTarget
    dtT.InitDrop Array(sw0,sw10,sw20),Array(0,10,20)
    dtT.initsnd  SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  set dtL=New cvpmDropTarget
    dtL.InitDrop Array(sw1,sw11,sw21),Array(1,11,21)
    dtL.initsnd  SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  set dtCL=New cvpmDropTarget
    dtCL.InitDrop Array(sw2,sw12,sw22),Array(2,12,22)
    dtCL.initsnd  SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  set dtCR=New cvpmDropTarget
    dtCR.InitDrop Array(sw3,sw13,sw23),Array(3,13,23)
    dtCR.initsnd  SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  Controller.Switch(36)=1
  Controller.Switch(54)=1

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If keycode=RightFlipperKey Then Controller.Switch(35)=1
  If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If keycode=RightFlipperKey Then Controller.Switch(35)=0
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit
  Me.DestroyBall
  BallsInPlay=BallsInPlay-1
  Controller.Switch(36)=1
  If BallsInPlay=0 Then Controller.Switch(54)=1
End Sub

Sub sw5_Hit:bsSaucer.AddBall 0 : playsoundAtVol "popper_ball", sw5, VolKick: End Sub
Sub sw6_Hit:bsSaucer2.AddBall 0 : playsoundAtVol "popper_ball", sw6, VolKick: End Sub

'Drop Targets
Sub sw0_Dropped:dtT.Hit 1:End Sub
Sub sw10_Dropped:dtT.Hit 2:End Sub
Sub sw20_Dropped:dtT.Hit 3:End Sub

Sub sw1_Dropped:dtL.Hit 1:End Sub
Sub sw11_Dropped:dtL.Hit 2:End Sub
Sub sw21_Dropped:dtL.Hit 3:End Sub

Sub sw2_Dropped:dtCL.Hit 1:End Sub
Sub sw12_Dropped:dtCL.Hit 2:End Sub
Sub sw22_Dropped:dtCL.Hit 3:End Sub

Sub sw3_Dropped:dtCR.Hit 1:End Sub
Sub sw13_Dropped:dtCR.Hit 2:End Sub
Sub sw23_Dropped:dtCR.Hit 3:End Sub

'Wire Triggers
Sub sw4_Hit:Controller.Switch(4)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw4_unHit:Controller.Switch(4)=0:End Sub
Sub sw14_Hit:Controller.Switch(14)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw14_unHit:Controller.Switch(14)=0:End Sub
Sub sw15_Hit:Controller.Switch(15)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw15_unHit:Controller.Switch(15)=0:End Sub
Sub sw16_Hit:Controller.Switch(16)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw16_unHit:Controller.Switch(16)=0:End Sub
Sub sw24_Hit:Controller.Switch(24)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw24_unHit:Controller.Switch(24)=0:End Sub
Sub sw30_Hit:Controller.Switch(30)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw30_unHit:Controller.switch(30)=0:End Sub
Sub sw31_Hit:Controller.Switch(31)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw31_UnHit:Controller.Switch(31)=0:End Sub
Sub sw50_Hit:Controller.Switch(50)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw50_unHit:Controller.Switch(50)=0:End Sub
Sub sw51_Hit:Controller.Switch(51)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw51_unHit:Controller.switch(51)=0:End Sub

 'Scoring Rubber
Sub sw33a_Slingshot:vpmTimer.PulseSw 33 : playsoundAtVol SoundFX("slingshot",DOFContactors), ActiveBall, 1: End Sub
Sub sw33b_Slingshot:vpmTimer.PulseSw 33 : playsoundAtVol SoundFX("slingshot",DOFContactors), ActiveBall, 1: End Sub
Sub sw26_Slingshot:VpmTimer.PulseSw 26 : playsoundAtVol SoundFX("slingshot",DOFContactors), ActiveBall, 1: End Sub

'kick back targets
Sub sw25a_Slingshot:vpmTimer.PulseSw 25:sw25ap.TransX = -8:Me.TimerEnabled = 1 : playsoundAtVol SoundFX("popper",DOFContactors), ActiveBall, 1: End Sub
Sub sw25a_Timer:sw25ap.TransX = 0:Me.TimerEnabled = 0:End Sub
Sub sw25b_Slingshot:vpmTimer.PulseSw 25:sw25bp.TransX = -8:Me.TimerEnabled = 1 : playsoundAtVol SoundFX("popper",DOFContactors), ActiveBall, 1: End Sub
Sub sw25b_Timer:sw25bp.TransX = 0:Me.TimerEnabled = 0:End Sub
Sub sw25c_Slingshot:vpmTimer.PulseSw 25:sw25cp.TransX = -8:Me.TimerEnabled = 1 : playsoundAtVol SoundFX("popper",DOFContactors), ActiveBall, 1: End Sub
Sub sw25c_Timer:sw25cp.TransX = 0:Me.TimerEnabled = 0:End Sub
Sub sw25d_Slingshot:vpmTimer.PulseSw 25:sw25dp.TransX = -8:Me.TimerEnabled = 1 : playsoundAtVol SoundFX("popper",DOFContactors), ActiveBall, 1: End Sub
Sub sw25d_Timer:sw25dp.TransX = 0:Me.TimerEnabled = 0:End Sub
Sub sw25e_Slingshot:vpmTimer.PulseSw 25:sw25ep.TransX = -8:Me.TimerEnabled = 1 : playsoundAtVol SoundFX("popper",DOFContactors), ActiveBall, 1: End Sub
Sub sw25e_Timer:sw25ep.TransX = 0:Me.TimerEnabled = 0:End Sub

 'Stand Up Targets
Sub sw52_Hit:vpmTimer.PulseSw 52:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:End Sub

'Bumpers
Sub Bumper1_Hit:VpmTimer.PulseSw 53 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper2_Hit:VpmTimer.PulseSw 53 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper3_Hit:VpmTimer.PulseSw 34 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper4_Hit:VpmTimer.PulseSw 34 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper5_Hit:VpmTimer.PulseSw 34 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub



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

     'Special Handling
     If chgLamp(ii,0) = 12 Then solTrough2 chgLamp(ii,1)
     'If chgLamp(ii,0) = 4 Then PFGI chgLamp(ii,1)

        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps()
NFadeL 3, Light3
NFadeL 4, Light4
NFadeL 5, Light5
NFadeL 6, Light6
NFadeL 7, Light7
NFadeL 8, Light8
'NFadeL 12, Light12 'ballrelease
NFadeL 13, Light13
NFadeL 14, Light14
NFadeL 15, Light15
NFadeL 16, Light16
NFadeL 17, Light17
NFadeL 18, Light18
NFadeL 19, Light19
NFadeL 20, Light20
NFadeL 21, Light21
NFadeL 22, light22
NFadeL 23, light23
NFadeL 24, light24
NFadeL 25, light25
NFadeL 26, light26
NFadeL 27, light27
NFadeL 28, light28
NFadeL 29, light29
NFadeL 30, light30
NFadeL 31, light31
NFadeL 32, light32
NFadeL 33, light33
NFadeL 34, light34
NFadeL 35, light35
NFadeL 36, light36
NFadeL 37, light37
NFadeL 38, light38
NFadeL 39, light39
NFadeL 40, light40
NFadeL 41, light41
NFadeL 42, light42
NFadeL 43, light43
NFadeL 44, light44
NFadeL 45, light45

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


'**********************************************************************************************************
' Backglass Light Displays (7 digit 7 segment displays)
'**********************************************************************************************************

Dim Digits(32)
Digits(0)=Array(a00,a01,a02,a03,a04,a05,a06,n,a08)
Digits(1)=Array(a10,a11,a12,a13,a14,a15,a16,n,a18)
Digits(2)=Array(a20,a21,a22,a23,a24,a25,a26,n,a28)
Digits(3)=Array(a30,a31,a32,a33,a34,a35,a36,n,a38)
Digits(4)=Array(a40,a41,a42,a43,a44,a45,a46,n,a48)
Digits(5)=Array(a50,a51,a52,a53,a54,a55,a56,n,a58)
Digits(6)=Array(b00,b01,b02,b03,b04,b05,b06,n,b08)

Digits(7)=Array(b10,b11,b12,b13,b14,b15,b16,n,b18)
Digits(8)=Array(b20,b21,b22,b23,b24,b25,b26,n,b28)
Digits(9)=Array(b30,b31,b32,b33,b34,b35,b36,n,b38)
Digits(10)=Array(b40,b41,b42,b43,b44,b45,b46,n,b48)
Digits(11)=Array(b50,b51,b52,b53,b54,b55,b56,n,b58)
Digits(12)=Array(c00,c01,c02,c03,c04,c05,c06,n,c08)
Digits(13)=Array(c10,c11,c12,c13,c14,c15,c16,n,c18)

Digits(14)=Array(c20,c21,c22,c23,c24,c25,c26,n,c28)
Digits(15)=Array(c30,c31,c32,c33,c34,c35,c36,n,c38)
Digits(16)=Array(c40,c41,c42,c43,c44,c45,c46,n,c48)
Digits(17)=Array(c50,c51,c52,c53,c54,c55,c56,n,c58)
Digits(18)=Array(d00,d01,d02,d03,d04,d05,d06,n,d08)
Digits(19)=Array(d10,d11,d12,d13,d14,d15,d16,n,d18)
Digits(20)=Array(d20,d21,d22,d23,d24,d25,d26,n,d28)

Digits(21)=Array(d30,d31,d32,d33,d34,d35,d36,n,d38)
Digits(22)=Array(d40,d41,d42,d43,d44,d45,d46,n,d48)
Digits(23)=Array(d50,d51,d52,d53,d54,d55,d56,n,d58)
Digits(24)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(25)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)
Digits(26)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(27)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)

'credit -- Ball In Play
Digits(28) = Array(e2,e3,e7,e4,e5,e1,e6,n,e23)
Digits(29) = Array(e9,e17,e22,e19,e20,e8,e21,n,e24)
Digits(31) = Array(g2,g3,g7,g4,g5,g1,g6,n,g23)
Digits(32) = Array(g9,g17,g22,g19,g20,g8,g21,n,g24)

Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 32) then
        For Each obj In Digits(num)
          If chg And 1 Then obj.State = stat And 1
          chg = chg\2 : stat = stat\2
        Next
      else
        'if char(stat) > "" then msg(num) = char(stat)
      end if
    next
    end if
end if
End Sub


'**********************************************************************************************************
'**********************************************************************************************************

'Gottlieb Punk
'added by Inkochnito
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"Punk - DIP switches"
    .AddFrame 0,0,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
    .AddFrame 0,76,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 0,122,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
    .AddFrame 0,168,190,"Novelty mode",&H08000000,Array("normal game mode",0,"50K per special/extra ball",&H08000000)'dip 28
    .AddFrame 205,0,190,"High score to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 credits",&H00400000,"displayed and 3 credits",&H00C00000)'dip 23&24
    .AddFrame 205,76,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
    .AddFrame 205,122,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddFrame 205,168,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
    .AddFrame 205,214,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
    .AddFrame 205,260,190,"Out rollovers launch ball on",&H40000000,Array("last ball in play only",0,"each ball in play",&H40000000)'dip 31
    .AddFrameExtra 0,214,190,"Attract Sound adjust",&H000C,Array("off",0,"every 10 seconds",&H0004,"every 2 minutes",&H0008,"every 4 minutes",&H000C)'SS-board dip 3&4
    .AddChkExtra 0,295,190,Array("Background sound",&H0010)'SS-board dip 5
    .AddChk 0,310,190,Array("Match feature",&H02000000)'dip 26
    .AddLabel 50,360,300,20,"After hitting OK, press F3 to reset game with new settings."
  End With
  Dim extra
  extra = Controller.Dip(4) + Controller.Dip(5)*256
  extra = vpmDips.ViewDipsExtra(extra)
  Controller.Dip(4) = extra And 255
  Controller.Dip(5) = (extra And 65280)\256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")

'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 26
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors),sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 26
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
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

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
    PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()

  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  FlipperLSh1.RotZ = LeftFlipper1.currentangle
  FlipperRSh1.RotZ = RightFlipper1.currentangle

  PrimFlipperL.RotZ = LeftFlipper.currentangle + 240
  PrimFlipperR.RotZ = RightFlipper.currentangle + 120
  PrimFlipperL1.RotZ = LeftFlipper1.currentangle + 260
  PrimFlipperR1.RotZ = RightFlipper1.currentangle + 140


  PGateL.RotZ = GateL.currentangle
  PGateR.RotZ = GateR.currentangle

End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
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

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
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
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub
