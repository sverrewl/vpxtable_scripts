Option Explicit
Randomize


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

Dim FlipLag
FlipLag = 1  'Enable/Disable FlipperLag Fix  0 is disable 1 is enable

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="ssvc_a26",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01530000","de.vbs",3.1

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
SolCallback(1)=   "SetLamp 101,"
SolCallback(2)=   "SetLamp 102,"
SolCallback(3)=   "SetLamp 103,"
SolCallback(4)=   "SetLamp 104,"
SolCallback(5)=   "SetLamp 105,"
SolCallback(6)=   "SetLamp 106,"
SolCallback(7)=   "SetLamp 107,"
SolCallback(8)=   "SetLamp 108,"
'qSolCallback(8)=   "SolShieldF"'Shield,Clear Mars
SolCallback(9)=   "SetLamp 109,"
'SolCallback(10)="SolBrelay"   '?????????????????????????
SolCallback(11)=  "PFGI"
'SolCallback(12)="SetLamp 112," 'Backglass Only
'SolCallback(13)="SetLamp 113,"
SolCallback(14)=  "SolDrainPlug" 'Center Post between Flippers
SolCallback(15)=  "kickback"
SolCallback(16)=  "bsVLock.SolExit"
SolCallback(25)=  "SolEatDown"
SolCallback(26)=  "bsTrough.SolIn"
SolCallback(27)=  "bsSaucer.SolOut"
SolCallback(28)=  "dtDrop.SolDropUp"
SolCallback(30)=  "SolEatUp"
'SolCallback(31)= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(32)=  "bsTrough.SolOut"
SolCallback(32)=  "ballreleasehack"



'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
  If FlipLag = 0 then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
  end If
    Else
  If FlipLag = 0 then
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
  end if
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
  If FlipLag = 0 then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), RightFlipper, VolFlip
    PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), RightFlipper1, VolFlip
        RightFlipper.RotateToEnd
    RightFlipper1.RotateToEnd
  end if
    Else
  If FlipLag = 0 then
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), RightFlipper, VolFlip
    PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), RightFlipper1, VolFlip
        RightFlipper.RotateToStart
    RightFlipper1.RotateToStart
  end if
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub




'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'Playfield GI
Sub PFGI(Enabled)
  If Enabled Then
    FlippersEnabled = Enabled
    dim xx
    For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"
  Else
    For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"
  End If
End Sub


' Center Plug Functionality (Ideally setting intitial state is not neccisary, but with repeat triggering it's a safety to assure state initially)
sub SolDrainPlug(enabled)
  If Enabled Then
      ' Establish Initial State
      DrainPlug.IsDropped=1
      DrainPlug.TimerEnabled=0
      L24.BulbHaloHeight = 0
      CenterPost.transz = 0
      ' End Establish Intital State
    DrainPlug.TimerEnabled=1
    DrainPlug.IsDropped=0
    L24.BulbHaloHeight = 27
    CenterPost.transz = 23
    PlaySoundAtVol SoundFX("Centerpost_Up",DOFContactors), L24, VolGates
    Controller.Switch(52)=0
  End If
end sub

Sub DrainPlug_Timer
  DrainPlug.TimerEnabled=0
  DrainPlug.IsDropped=1
  L24.BulbHaloHeight = 2
  CenterPost.transz = 0
  PlaySoundAtVol SoundFX("Centerpost_Down",DOFContactors), L24, VolGates
  Controller.Switch(52)=1
End Sub



Sub Kickback(Enabled)
  If Enabled Then
    plunger1.fire
    PlaySoundAtVol SoundFX("Centerpost_Down",DOFContactors), plunger1, 1.5
  Else
    plunger1.pullback
  End If
End Sub

' ************************
'    KGB Secret Entrance
' ************************

Sub SolEatDown(Enabled)
  If Enabled Then
    Wall_SpySurround.IsDropped = 1
    EatWall.z = -28
    sw11.Enabled=0
    Controller.Switch(12)=1
    'Playsound SoundFX("fx_SolDown",DOFContactors)
    PlaySoundAtVol SoundFX("fx_SolDown",DOFContactors), sw11, VolGates
    ''add primitive animaiton
  End If
End Sub

'Dim hasBall
Sub SolEatUp(Enabled)
  If Enabled Then
    Wall_SpySurround.IsDropped = 0
    EatWall.z = 25
    sw11.Enabled=1
    Controller.Switch(12)=0
    'Playsound SoundFX("fx_SolUp",DOFContactors)
    PlaySoundAtVol SoundFX("fx_SolUp",DOFContactors), sw11, VolGates
    ''add primitive animaiton
    If bsEater.Balls Then bsEater.ExitSol_On
  End If
End Sub


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsSaucer, bsEater, dtDrop, bsVLock

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = ""&chr(13)&"Secret Service By Data East 1988, VPX By Derek Crosby, with help from VPX community, specifically 32Assasin,Baron Shadow, roccodimarco & kiwi  "
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

  PinMAMETimer.Interval=PinMAMEInterval:
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=1:
  vpmNudge.Sensitivity=2
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)
  if vpmNudge.TiltSwitch Then
    FlippersEnabled = 0
  end if

  Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 46,47,48,49,0,0,0,0
    bsTrough.InitKick BallRelease,95,5
    bsTrough.InitExitSnd"ballrelease","SolOn"
    bsTrough.Balls=3

  Set bsSaucer=New cvpmBallStack
    bsSaucer.InitSaucer sw19,19,85,4
    bsSaucer.InitExitSnd"Popper","SolOn"
    bsSaucer.KickForceVar = 3
    bsSaucer.KickAngleVar = 3

  Set bsEater=New cvpmBallStack
    bsEater.InitSw 0,11,0,0,0,0,0,0
    bsEater.InitKick sw11,200,5
    bsEater.InitExitSnd"ballrelease","SolOn"
    bsEater.KickForceVar = 3
    bsEater.KickAngleVar = 3

  Set dtDrop=New cvpmDropTarget
    dtDrop.InitDrop Array(sw36,sw37,sw38,sw39,sw40),Array(36,37,38,39,40)
    'dtDrop.InitSnd"FlapClos","FlapOpen"
    dtDrop.InitSnd"gate4","gate"

  Set bsVLock=New cvpmVLock
    bsVLock.InitVLock Array(Trigger1,Trigger2),Array(Kickbig2,Kickbig1),Array(42,41)
    bsVLock.CreateEvents"bsVLock"

    Controller.Switch(52)=1 'Drop Post
    Controller.Switch(12)=1 'Drop Eater

    plunger1.pullback
    Wall_SpySurround.IsDropped = 1
    DrainPlug.IsDropped=1


End Sub

 Sub ballreleasehack(enabled)
  if (enabled) then
    ' If there is only one ball in trough, then don't kick out ball
    ' unless there are two balls locked at bsVUKLeft
    if bsTrough.balls = 2 then
      if bsVLock.balls = 1 then
        Playsound "Plunger" : bsTrough.SolOut enabled
      else
        ' ooops error case
        'playsound "stern100"
      end if
    else
      playsound "plunger" : bsTrough.SolOut enabled
    end if
  end if
 End Sub


  Sub BKT_Hit
    Set BKTBall=ActiveBall
    InTheZone=1
  End Sub

  Sub BKT_unHit
    InTheZone=0

    EatWall.visible=0
    EatWall.collidable=0

    'Injection_Bar_Back.isDropped=0
    Dim y
    For each y in Injection_Bar_Back
      y.collidable=1
      y.visible=1
    Next

  End Sub

  Sub ModifyBall_Timer
    If InTheZone=1 Then
      BKTBall.VelY=BKTBall.VelY+5
        Dim x
        For each x in Injection_Bar_Fwd
          x.collidable=1
          x.visible=1
        Next
        Dim y
        For each y in Injection_Bar_Back
          y.collidable=0
          y.visible=0
        Next
      If BKTBall.VelX<-2 Then BKTBall.VelX=BKTBall.VelX+5
    End If
  End Sub


'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
        If keycode = LeftFlipperKey Then
      If FlipLag = 1 then flipnf 0, 1
        end if
        If keycode = RightFlipperKey Then
            If FlipLag = 1 then flipnf 1, 1
        end if
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySoundAtVol "fx_PlungerPull", Plunger, 1 :Plunger.Pullback
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
        If keycode = LeftFlipperKey Then
            If FlipLag = 1 then flipnf 0, 0
        end if
        If keycode = RightFlipperKey Then
            If FlipLag = 1 then flipnf 1, 0
        end if
    If keycode = PlungerKey Then PlaySoundAtVol "fx_plunger", Plunger, 1 : Plunger.Fire
    If vpmKeyUp(keycode) Then Exit Sub
End Sub


'*********
' Special Flippers
'*********
dim FlippersEnabled



sub flipnf(LR, DU)
    if LR = 0 Then        'left flipper
        if DU = 1 then
            If FlippersEnabled = True then
                leftflipper.rotatetoend
                LeftFlipperSound 1
            end if
            controller.Switch(swLLFlip) = True
        Elseif DU = 0 then
            If FlippersEnabled = True then
                leftflipper.rotatetoStart
                LeftFlipperSound 0
            end if
            controller.Switch(swLLFlip) = False
        end if
    elseif LR = 1 then        ''right flipper
        if DU = 1 then
            If FlippersEnabled = True then
                RightFlipper.rotatetoend
        RightFlipper1.rotatetoend
                RightFlipperSound 1
            end if
            controller.Switch(swLRFlip) = True
        Elseif DU = 0 then
            If FlippersEnabled = True then
                RightFlipper.rotatetoStart
        RightFlipper1.rotatetoStart
                RightFlipperSound 0
            end if
            controller.Switch(swLRFlip) = False
        end if
    end if
end sub

sub LeftFlipperSound(updown)'called along with the flipper, so feel free to add stuff, EOStorque tweaks, animation updates, upper flippers, whatever.
    if updown = 1 Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), LeftFlipper, VolFlip
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), LeftFlipper, VolFlip
    end if
end sub
sub RightFlipperSound(updown)
    if updown = 1 Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), RightFlipper, VolFlip
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), RightFlipper, VolFlip
    end if
end sub



'**********************************************************************************************************

 ' Drain hole and kickers
'Sub Drain_Hit:bsTrough.addball me : PlaySoundAtBall "drain", ActiveBall : End Sub
Sub Drain_Hit:PlaysoundAtVol "fx_drain", drain, 1:bsTrough.AddBall Me:End Sub
Sub sw11_Hit: bsEater.AddBall me : PlaySoundAtVol "popper_ball", sw11, 1: SolEatDown True : End Sub 'Spy Hideout
Sub sw19_Hit: bsSaucer.AddBall 0 : PlaySoundAtVol "popper_ball", sw19, 1: End Sub

'Drop Targets
Sub sw36_Dropped:dtDrop.Hit 1:End Sub
Sub sw36_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFContactors), sw36, VolTarg:End Sub 'hit event only for the sound
Sub sw37_Dropped:dtDrop.Hit 2:End Sub
Sub sw37_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFContactors), sw37, VolTarg:End Sub 'hit event only for the sound
Sub sw38_Dropped:dtDrop.Hit 3:End Sub
Sub sw38_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFContactors), sw38, VolTarg:End Sub 'hit event only for the sound
Sub sw39_Dropped:dtDrop.Hit 4:End Sub
Sub sw39_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFContactors), sw39, VolTarg:End Sub 'hit event only for the sound
Sub sw40_Dropped:dtDrop.Hit 5:End Sub
Sub sw40_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFContactors), sw40, VolTarg:End Sub 'hit event only for the sound

'Wire Triggers
Sub sw9_Hit:Controller.Switch(9)=1 : PlaySoundAtVol "rollover", sw9, volRol : End Sub
Sub sw9_unHit:Controller.Switch(9)=0:End Sub
Sub sw23_Hit:Controller.Switch(23)=1 : PlaySoundAtVol "rollover", sw23, volRol : End Sub
Sub sw23_unHit:Controller.Switch(23)=0:End Sub
Sub sw24_Hit:Controller.Switch(24)=1 : PlaySoundAtVol "rollover", sw24, volRol : End Sub
Sub sw24_unHit:Controller.Switch(24)=0:End Sub
Sub sw25_Hit:Controller.Switch(25)=1 : PlaySoundAtVol "rollover", sw25, volRol : End Sub
Sub sw25_unHit:Controller.Switch(25)=0:End Sub
Sub sw32_Hit:Controller.Switch(32)=1 : PlaySoundAtVol "rollover", sw32, volRol : End Sub
Sub sw32_unHit:Controller.Switch(32)=0:End Sub
Sub sw33_Hit:Controller.Switch(33)=1 : PlaySoundAtVol "rollover", sw33, volRol : End Sub
Sub sw33_unHit:Controller.Switch(33)=0:End Sub
Sub sw34_Hit:Controller.Switch(34)=1 : PlaySoundAtVol "rollover", sw34, volRol : End Sub
Sub sw34_unHit:Controller.Switch(34)=0:End Sub
Sub sw35_Hit:Controller.Switch(35)=1 : PlaySoundAtVol "rollover", sw35, volRol : End Sub
Sub sw35_unHit:Controller.Switch(35)=0:End Sub

'Ramp Triggers
Sub sw43_Hit:vpmTimer.PulseSw 43 : PlaySoundAtBall "wireramp" : End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44 : PlaySoundAtBall "wireramp" : End Sub
'Sub WireRampStart_Hit : PlaySoundAtVol "wireramp", ActiveBall, 1 : End Sub
'Sub WireRampEnd_Hit : StopSound "wireramp" : End Sub

'Star Trigger
Sub sw50_Hit:Controller.Switch(50)=1 : PlaySoundAtVol "rollover", sw50, volRol : End Sub
Sub sw50_unHit:Controller.Switch(50)=0:End Sub

 'Stand Up Targets
Sub sw10_Hit:vpmTimer.PulseSw 10:End Sub
Sub sw13_Hit:vpmTimer.PulseSw 13:End Sub
Sub sw14_Hit:vpmTimer.PulseSw 14:End Sub
Sub sw15_Hit:vpmTimer.PulseSw 15:End Sub
Sub sw16_Hit:vpmTimer.PulseSw 16:End Sub

'Scoring Rubber
Sub sw18_Hit:vpmTimer.PulseSw 18:End Sub
Sub S27_Hit:vpmTimer.PulseSw 27:End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:End Sub

'Spinners
Sub sw17_Spin:vpmTimer.PulseSw 17 : PlaySoundAtVol "fx_spinner", sw17, volSpin : End Sub
Sub sw26_Spin:vpmTimer.PulseSw 26 : PlaySoundAtVol "fx_spinner", sw26, volSpin : End Sub

'Bumpers
Sub Bumper1_Hit
  vpmTimer.PulseSw 20
  PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper1, volBump
  Bumper1_Light.State = 1
  Me.TimerEnabled = 1
End Sub
Sub Bumper1_Timer()
  Bumper1_Light.State = 0
  Me.TimerEnabled = 0
End Sub

Sub Bumper2_Hit
  vpmTimer.PulseSw 21
  PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper2, volBump
  Bumper2_Light.State = 1
  Me.TimerEnabled = 1
End Sub
Sub Bumper2_Timer()
  Bumper2_Light.State = 0
  Me.TimerEnabled = 0
End Sub

Sub Bumper3_Hit
  vpmTimer.PulseSw 22
  PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper3, volBump
  Bumper3_Light.State = 1
  Me.TimerEnabled = 1
End Sub
Sub Bumper3_Timer()
  Bumper3_Light.State = 0
  Me.TimerEnabled = 0
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
     'If chgLamp(ii,0) = 2 Then solTrough chgLamp(ii,1)
     'If chgLamp(ii,0) = 4 Then PFGI chgLamp(ii,1)

        Next
    End If
    UpdateLamps
End Sub


Sub UpdateLamps()
NFadeL 1, L1
NFadeL 2, L2 'White House
NFadeL 3, L3 'Russian Embassy
NFadeL 4, L4 'Jefferson Memorial
NFadeL 5, L5 'Jefferson Memorial
NFadeL 6, L6
NFadeL 7, L7
NFadeL 8, L8
NFadeL 9, L9
NFadeL 10, L10
NFadeL 11, L11
NFadeL 12, L12
NFadeL 13, L13
NFadeL 14, L14
NFadeL 15, L15
NFadeL 16, L16
NFadeL 17, L17
NFadeL 18, L18
NFadeL 19, L19
NFadeL 20, L20
NFadeL 21, L21
NFadeLm 22, L22
NFadeLm 22, L22a
NFadeL 23, L23
NFadeL 24, L24
NFadeL 25, L25
NFadeL 26, L26
NFadeL 27, L27
NFadeL 28, L28
NFadeL 29, L29
NFadeL 30, L30
NFadeL 31, L31
NFadeL 32, L32
NFadeL 33, L33
NFadeL 34, L34
NFadeL 35, L35
NFadeL 36, L36
NFadeL 37, L37
NFadeL 38, L38
NFadeLm 39, L39
NFadeLm 39, L39a
NFadeLm 39, L39b
NFadeL 40, L40
NFadeL 41, L41
NFadeL 42, L42
NFadeL 43, L43
NFadeL 44, L44
NFadeL 45, L45
NFadeL 46, L46
NFadeL 47, L47
NFadeL 48, L48
NFadeL 49, L49
NFadeL 50, L50
NFadeL 51, L51
NFadeL 52, L52
NFadeL 53, L53
NFadeL 54, L54
NFadeL 55, L55
NFadeL 56, L56
NFadeL 57, L57
NFadeL 58, L58
NFadeL 59, L59
NFadeL 60, L60
NFadeL 61, L61
NFadeL 62, L62
NFadeL 63, L63
NFadeL 64, L64

'Solenoid Controlled

NFadeLm 101, S101
NFadeLm 101, S101a

NFadeLm 102, S102
NFadeLm 102, S102a

Flashm 103, Flasher103a
Flashm 103, Flasher103b
NFadeLm 103, L103a
NFadeLm 103, L103b
NFadeObjm 103, P103a, "dome2_0_blueOn", "dome2_0_blue"  'Dome
NFadeObjm 103, P103b, "dome2_0_blueOn", "dome2_0_blue"  'Dome

NFadeLm 104, S104
NFadeLm 104, S104a

NFadeLm 105, S105
NFadeLm 105, S105a

NFadeLm 106, S106


Flashm 108, Flasher108a
Flashm 108, Flasher108b
NFadeLm 108, L108a
NFadeLm 108, L108b
NFadeObjm 108, P108a, "dome2_0_clearON", "dome2_0_clear"  'Dome
NFadeObjm 108, P108b, "dome2_0_clearON", "dome2_0_clear"  'Dome

NFadeL 109, S109

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
'Digital Display
'**********************************************************************************************************

 Dim Digits(28)

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

Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)


 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
    If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 28) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
      Else
             end if
        Next
     end if
    End If
 End Sub

'**********************************************************************************************************
'**********************************************************************************************************



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, LStep, TStep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 29
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), Sling1, volKick
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
  vpmTimer.PulseSw 28
  PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), Sling2, volKick
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


Sub Top310Point_Slingshot
  vpmTimer.PulseSw 18
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), Sling3, volKick
    LUpSling.Visible = 0
    LUpSling1.Visible = 1
    Sling3.rotx = 20
    TStep = 0
    Top310Point.TimerEnabled = 1
End Sub

Sub Top310Point_Timer
    Select Case TStep
        Case 3:LUpSling1.Visible = 0:LUpSling2.Visible = 1:Sling3.rotx = 10
        Case 4:LUpSling2.Visible = 0:LUpSling.Visible = 1:Sling3.rotx = 0:Top310Point.TimerEnabled = 0:
    End Select
    TStep = TStep + 1
End Sub


' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
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

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

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

'Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here because of the timing

Const tnob = 1 'ninuzzu - why 5 balls? Bad Cats has only one ball
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' disabling these as they has static volume
' Sub LeftFlipper_Collide(parm)
'   RandomSoundFlipper()
' End Sub
'
' Sub RightFlipper_Collide(parm)
'   RandomSoundFlipper()
' End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

' Thalamus - 2021-04-30 : added proper exit

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

