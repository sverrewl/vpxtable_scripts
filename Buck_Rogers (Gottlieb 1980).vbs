Option Explicit
Randomize

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-08-26 : Improved directional sounds
' !! NOTE : Table not verified yet !!

Const VolDiv = 1000    ' Lower number, louder ballrolling/collition sound
Const VolCol    = 3    ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 4    ' Bumpers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 2    ' Metals volume.
Const VolRH     = 2    ' Rubber hits volume.
Const VolPo     = 2    ' Rubber posts volume.
Const VolPi     = 2    ' Rubber pins volume.
Const VolTarg   = 3    ' Targets volume.
Const VolKick   = 3    ' Kicker volume.
Const VolSpin   = 3    ' Spinners volume.
Const VolFlip   = 2    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="buckrgrs",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const SCoin="coin",cCredits=""
Const VT_Delay_Factor = .88   'used to slow down the ball when hitting the vari target, smaller number slows down faster

LoadVPM"01150000","GTS1.VBS",3.22


'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)="bsTrough.SolOut"
SolCallback(2)="vpmSolSound SoundFX(""Knock"",DOFKnocker),"
SolCallback(3)="vpmSolSound SoundFX(""10pts"",DOFChimes),"
SolCallback(4)="vpmSolSound SoundFX(""100pts"",DOFChimes),"
SolCallback(5)="vpmSolSound SoundFX(""1000pts"",DOFChimes),"
SolCallback(6)="sw41exit"

SolCallback(7)="Rraised"  '"dtRight.SolDropUp"
SolCallback(8)="Lraised"  '"dtLeft.SolDropUp"
SolCallback(17)="vpmNudge.SolGameOn"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub Rraised(enabled)
  if enabled then Rreset.enabled=True
End Sub

Sub Rreset_timer
  dtRight.DropSol_On
  For each light in DTRightLights: light.state=0: Next
  Rreset.enabled=false
End Sub

Sub Lraised(enabled)
  if enabled then Lreset.enabled=True
End Sub

Sub Lreset_timer
  dtLeft.DropSol_On
  LDTsw13.state=0
  LDTsw14.state=0
  Lreset.enabled=false
End Sub

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip':LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip':LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("fx_Flipperup",DOFContactors), RightFlipper, VolFlip':RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip':RightFlipper.RotateToStart
     End If
End Sub

'**********************************************************************************************************


'Primitive Flipper Code

Sub FlipperTimer_Timer
  lflip.roty = LeftFlipper.currentangle
  rflip.roty = RightFlipper.currentangle
  Pgate1.rotz = (Gate1.currentangle*.75)+25
  PGate2.rotz = (Gate2.currentangle*.75)+25
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough,dtLeft,dtRight, bsSaucer1, bcK
Dim BPG, hsaward, n, plungerball, xx, Lstep, GOflip, TILTflip

Sub BuckRogers_Init
  vpmInit Me
  plungerball=0
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine="Buck Rogers (Gottlieb 1980)"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 1
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
    Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch=4
  vpmNudge.Sensitivity=3
  vpmNudge.TiltObj=Array(BumperL,BumperR,SlingL)


  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,66,0,0,0,0,0,0
  bsTrough.InitKick DRAIN,60,40
  bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors),SoundFX("drainkick",DOFContactors)
  bsTrough.Balls=1

    Set bsSaucer1=New cvpmBallstack
  bsSaucer1.InitSaucer sw41,42,125,6
  bsSaucer1.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("holekick",DOFContactors)


  bcKicker.CreateBall
  bcKicker.Kick 0,5
  bcKicker.Enabled=False

     set dtLeft = new cvpmdroptarget
     With dtLeft
        .initdrop Array(sw13,sw14,sw23,sw24), array(13,14,23,24)
    .InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
     End With

     set dtRight = new cvpmdroptarget
     With dtRight
        .initdrop Array(sw33,sw34,sw43,sw44), array(33,34,43,44)
    .InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
     End With

  FindDips    'find balls per game and high score reward

  if b2son then: for each xx in backdropstuff: xx.visible=0: next
  startGame.enabled=true

End Sub

Sub BuckRogers_Paused:Controller.Pause = 1:End Sub

Sub BuckRogers_unPaused:Controller.Pause = 0:End Sub

Sub BuckRogers_Exit
  If b2son then controller.stop
End Sub

sub startGame_timer
  playsoundAt "poweron", Drain
  For each xx in GILights:xx.State = 1: Next    '*****GI Lights On
  Lgun.state=2
  me.enabled=false
end sub


Sub BuckRogers_KeyDown(ByVal keycode)
  If GOflip+TILTflip=0 then
    If keycode = LeftFlipperKey Then LeftFlipper.RotateToEnd
    If keycode = RightFlipperKey Then RightFlipper.RotateToEnd
  End If
  If vpmKeyDown(KeyCode) Then Exit Sub
  If keycode=AddCreditKey then playsoundat "coin", Drain: vpmTimer.pulseSW (swCoin1): end if
  If keycode=PlungerKey Then Plunger.Pullback:playsoundat "plungerpull", plunger
End Sub

Sub BuckRogers_KeyUp(ByVal keycode)
  If keycode = LeftFlipperKey Then LeftFlipper.RotateToStart
    If keycode = RightFlipperKey Then RightFlipper.RotateToStart

  If keycode = 61 then FindDips
  If vpmKeyUp(KeyCode) Then Exit Sub
  If keycode=PlungerKey Then
    Plunger.Fire
    if plungerball=1 then
      playsoundat "plungerreleaseball", plunger
      else
      playsoundat "plungerreleasefree", plunger
    end if
  end if
End Sub

Sub Ballhome_hit
  plungerball=1
  DOF 116, 1:
end sub

sub ballhome_unhit
  plungerball=0
  DOF 116, 0
end sub

Sub Drain_Hit
  playsoundat "drain", Drain
  bsTrough.addball me
  DOF 115, 2
End Sub


'bumpers
Sub BumperL_Hit
  vpmTimer.PulseSw 31
  DOF 132, 2
  playsoundAtVol SoundFX("fx_bumper",DOFContactors), BumperL, VolBump
End Sub

Sub BumperR_Hit
  vpmTimer.PulseSw 31
  DOF 131, 2
  playsoundAtVol SoundFX("fx_bumper",DOFContactors), BumperR, VolBump
End Sub

'Hit Target

  Sub sw11_Hit:vpmTimer.PulseSw 11:End Sub


'Rollover wire Tirggers

    Sub sw61a_Hit:Controller.Switch(61) = 1:PlaySoundAt "sensor", sw61a:DOF 109, 1:End Sub
    Sub sw61a_UnHit:Controller.Switch(61) = 0:DOF 109, 0:End Sub
    Sub sw61b_Hit:Controller.Switch(61) = 1:PlaySoundAt "sensor", sw61b:DOF 110, 1:End Sub
    Sub sw61b_UnHit:Controller.Switch(61) = 0:DOF 110, 0:End Sub
    Sub sw63_Hit:Controller.Switch(63) = 1:PlaySoundAt "sensor", sw63:End Sub
    Sub sw63_UnHit:Controller.Switch(63) = 0:End Sub
    Sub sw70_Hit:Controller.Switch(70) = 1:PlaySoundAt "sensor", sw70:DOF 107, 1:End Sub
    Sub sw70_UnHit:Controller.Switch(70) = 0:DOF 107, 0:End Sub
    Sub sw70a_Hit:Controller.Switch(70) = 1:PlaySoundAt "sensor", sw70a:DOF 111, 1:End Sub
    Sub sw70a_UnHit:Controller.Switch(70) = 0:DOF 111, 0:End Sub
    Sub sw71_Hit:Controller.Switch(71) = 1:PlaySoundAt "sensor", sw71:End Sub
    Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub
    Sub sw73_Hit:Controller.Switch(73) = 1:PlaySoundAt "sensor", sw73:End Sub
    Sub sw73_UnHit:Controller.Switch(73) = 0:End Sub
    Sub sw74_Hit:Controller.Switch(74) = 1:PlaySoundAt "sensor", sw74:DOF 108, 1:End Sub
    Sub sw74_UnHit:Controller.Switch(74) = 0:DOF 108, 0:End Sub
    Sub sw74a_Hit:Controller.Switch(74) = 1:PlaySoundAt "sensor", sw74a:DOF 112, 1:End Sub
    Sub sw74a_UnHit:Controller.Switch(74) = 0:DOF 112, 0:End Sub

'Star Rollovers

    Sub sw31_1_Hit:Controller.Switch(31) = 1:PlaySoundAt "sensor", sw31_1:DOF 113,1: End Sub
    Sub sw31_1_UnHit:Controller.Switch(31) = 0:DOF 113,0:End Sub
    Sub sw31_2_Hit:Controller.Switch(31) = 1:PlaySoundAt "sensor", sw31_2:DOF 114, 1:End Sub
    Sub sw31_2_UnHit:Controller.Switch(31) = 0:DOF 114, 0:End Sub

'Slingshot



Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 21
    PlaySoundAt SoundFXDOF("left_slingshot",101,DOFPulse,DOFContactors),slingL
    LSling.Visible = 0
    LSling1.Visible = 1
  slingL.objroty = 15
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:slingL.objroty = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

'Scoring rubbers - ANIMATED!!!!!

  Sub sw54a_hit()
    vpmTimer.PulseSw 54
    R54a.visible=0
    R54a1.visible=1
    me.uservalue=1
    me.timerenabled=1
  End Sub

  sub sw54a_timer
    select case sw54a.uservalue
      Case 1: r54a1.visible=0: r54a.visible=1
      Case 2: r54a.visible=0: r54a2.visible=1
      Case 3: r54a2.visible=0: r54a.visible=1: me.timerenabled=0
    end Select
    sw54a.uservalue=sw54a.uservalue+1
  end sub

  Sub sw54b_hit()
    vpmTimer.PulseSw 54
    R54b.visible=0
    R54b1.visible=1
    me.uservalue=1
    me.timerenabled=1
  End Sub

  sub sw54b_timer
    select case sw54b.uservalue
      Case 1: r54b1.visible=0: r54b.visible=1
      Case 2: r54b.visible=0: r54b2.visible=1
      Case 3: r54b2.visible=0: r54b.visible=1: me.timerenabled=0
    end Select
    sw54b.uservalue=SW54b.uservalue+1
  end sub

  Sub sw54c_hit()
    vpmTimer.PulseSw 54
    R54c.visible=0
    R54c1.visible=1
    me.uservalue=1
    me.timerenabled=1
  End Sub

  sub sw54c_timer
    select case sw54c.uservalue
      Case 1: r54c1.visible=0: r54c.visible=1
      Case 2: r54c.visible=0: r54c2.visible=1
      Case 3: r54c2.visible=0: r54c.visible=1: me.timerenabled=0
    end Select
    sw54c.uservalue=SW54c.uservalue+1
  end sub

  Sub sw54d_hit()
    vpmTimer.PulseSw 54
    R54d.visible=0
    R54d1.visible=1
    me.uservalue=1
    me.timerenabled=1
  End Sub

  sub sw54d_timer
    select case SW54d.uservalue
      Case 1: r54d1.visible=0: r54d.visible=1
      Case 2: r54d.visible=0: r54d2.visible=1
      Case 3: r54d2.visible=0: r54d.visible=1: me.timerenabled=0
    end Select
    sw54d.uservalue=SW54d.uservalue+1
  end sub

  Sub sw21a_hit()
    vpmTimer.PulseSw 21
    R21a.visible=0
    R21a1.visible=1
    me.uservalue=1
    me.timerenabled=1
  End Sub

  sub sw21a_timer
    select case SW21a.uservalue
      Case 1: r21a1.visible=0: r21a.visible=1
      Case 2: r21a.visible=0: r21a2.visible=1
      Case 3: r21a2.visible=0: r21a.visible=1: me.timerenabled=0
    end Select
    sw21a.uservalue=SW21a.uservalue+1
  end sub

  Sub sw21b_hit()
    vpmTimer.PulseSw 21
    R21b.visible=0
    R21b1.visible=1
    me.uservalue=1
    me.timerenabled=1
  End Sub

  sub sw21b_timer
    select case SW21b.uservalue
      Case 1: r21b1.visible=0: r21b.visible=1
      Case 2: r21b.visible=0: r21b2.visible=1
      Case 3: r21b2.visible=0: r21b.visible=1: me.timerenabled=0
    end Select
    SW21b.uservalue=SW21b.uservalue+1
  end sub

  Sub sw21c_hit()
    vpmTimer.PulseSw 21
    R21c.visible=0
    R21c1.visible=1
    me.uservalue=1
    me.timerenabled=1
  End Sub

  sub sw21c_timer
    select case SW21c.uservalue
      Case 1: r21c1.visible=0: r21c.visible=1
      Case 2: r21c.visible=0: r21c2.visible=1
      Case 3: r21c2.visible=0: r21c.visible=1: me.timerenabled=0
    end Select
    SW21c.uservalue=SW21c.uservalue+1
  end sub

'Kickers

   Sub sw41_Hit
    controller.switch(41) = 1
   End Sub

  Dim rkickstep

   Sub sw41exit(enabled)
    if enabled then
    controller.switch(41) = 0
    sw41.kick 219,13
    playsound SoundFX("holekick",DOFContactors)
    PkickarmR.rotz=15
    rkickstep=1
    sw41.timerenabled=1
    end if
  end Sub

  sub sw41_timer
    Select Case rkickstep
      Case 2:
      PkickarmR.rotz=0
      me.timerenabled=0
    End Select
    rkickstep=rkickstep+1
  end Sub

'Drop Targets

    Sub sw13_Hit:dtLeft.Hit 1:LDTsw13.state=1:End Sub
    Sub sw14_Hit:dtLeft.Hit 2:LDTsw14.state=1:End Sub
    Sub sw23_Hit:dtLeft.Hit 3:End Sub
    Sub sw24_Hit:dtLeft.Hit 4:End Sub

    Sub sw33_Hit:dtRight.Hit 1:LDTsw33.state=1:End Sub
    Sub sw34_Hit:dtRight.Hit 2:LDTsw44.state=1:End Sub
    Sub sw43_Hit:dtRight.Hit 3:LDTsw43.state=1:End Sub
    Sub sw44_Hit:dtRight.Hit 4:LDTsw44.state=1:End Sub

''***********************************************************************************
''****                  VariTarget Handling                   ****
''***********************************************************************************
Dim VariNewPos, VariOPos, VariSwitch, i
Dim VariSwitches: VariSwitches = Array(10, 20, 30, 40, 50, 60)
Dim VariPositions: VariPositions = Array(0, 4, 8, 12, 16, 20)



Sub VariTrigger_Hit(vidx)
  VariNewPos=vidx
  If ((ActiveBall.VelY < 0) AND (vidx >= VariOPos)) Then
    ActiveBall.VelY = ActiveBall.VelY * VT_Delay_Factor
    DOF 119,1
    For i = 0 to 5:If vidx >= VariPositions(i) Then VariSwitch = VariSwitches(i):End If:Next
    VariOPos=vidx: VariTargetP.TransX = -(VariOPos * 10)
  End If
  If ActiveBall.VelY >= 0 Then
    vpmTimer.PulseSw(VariSwitch)
  End If
End Sub

Sub VariTrigger_UnHit(idx)
  If ActiveBall.VelY > 0 Then DOF 119,0
End Sub

Sub VariTargetTimer_Timer
  If VariTargetP.TransX < 0 Then
    VariTargetP.TransX = VariTargetP.TransX + 5
  Else
    VariOPos = 0: Me.Enabled = 0
  End If
End Sub

Sub VariTarget_Init
  VariNewPos=0: VariOPos=0: VariSwitch = 0
End Sub



 Dim N1,O1, Light
 N1=0:O1=0:
 Set LampCallback=GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
  If L17.state > 0 Then VariTargetTimer.Enabled = 1
    N1=Controller.Lamp(1) 'Game Over triggers match and BIP

      If N1 then
        EMReelBIP.setvalue 1
        EMReelNTM.setvalue 0
        EMReelGO.setvalue 0
        GOflip=0
        lGI1.state = 1
        else
        EMReelBIP.setvalue 0
        EMReelNTM.setvalue 1
        EMReelGO.setvalue 1
        GOflip=1
        lGI1.state = 0
      end if

    N1=Controller.Lamp(2) 'Tilt
      If N1 then
        EMReelTilt.setvalue 1
        TILTflip=1
        else
        EMReelTilt.setvalue 0
        TILTflip=0
      end if


    N1=Controller.Lamp(4) 'Shoot Again
      If N1 then
        EMReelSPSA.setvalue 1
        else
        EMReelSPSA.setvalue 0
      end if

    N1=Controller.Lamp(3) 'HIGH SCORE TO DATE
      if N1 then
        EMReelHTD.setvalue 1
        else
        EMReelHTD.setvalue 0
      end if

 End Sub



'**********************************************************************************************************

'Map lights to an array
'**********************************************************************************************************
'Set Lights(1) = l1 'Game Over Backbox
'Set Lights(2) = l2 'Tilt Backbox
'Set Lights(3) = l3 'High Score Backbox
Set Lights(4) = l4 ' Shoot Again Playfied and backglass
Set Lights(5) = l5
Set Lights(6) = l6
Set Lights(7) = l7
Set Lights(8) = l8
Set Lights(9) = l9
Set Lights(10) = l10
Set Lights(11) = l11
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
Digits(26)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(27)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)
Digits(24)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(25)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)



Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    If not b2son Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 32) then
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



'Finding an individual dip state based on scapino's Strikes and spares dip code - from unclewillys pinball pool
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
  hsaward = TheDips(22)
  BPG = TheDips(9)
  dim ebplay: ebplay= TheDips(11)
  If BPG = 1 then
    if ebplay = 1 then
      instcard.image="InstCard3Balls"
      else
      instcard.image="InstCard3BallsEB"
    end if
    Else
    if ebplay = 1 then
      instcard.image="InstCard5Balls"
      else
      instcard.image="InstCard5BallsEB"
    end if
  End if
  repcard.image="replaycard"&hsaward
  DipsTimer.enabled=0
 End Sub

 'Gottlieb System 1
 'added by Inkochnito
 Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"System 1 - DIP switches"
    .AddFrame 205,0,190,"Maximum credits",&H00030000,Array("5 credits",0,"8 credits",&H00020000,"10 credits",&H00010000,"15 credits",&H00030000)'dip 17&18
    .AddFrame 0,0,190,"Coin chute control",&H00040000,Array("seperate",0,"same",&H00040000)'dip 19
    .AddFrame 0,46,190,"Game mode",&H00000400,Array("extra ball",0,"replay",&H00000400)'dip 11
    .AddFrame 0,92,190,"High game to date awards",&H00200000,Array("no award",0,"3 replays",&H00200000)'dip 22
    .AddFrame 0,138,190,"Balls per game",&H00000100,Array("5 balls",0,"3 balls",&H00000100)'dip 9
    .AddFrame 0,184,190,"Tilt effect",&H00000800,Array("game over",0,"ball in play only",&H00000800)'dip 12
    .AddChk 205,80,190,Array("Match feature",&H00000200)'dip 10
    .AddChk 205,95,190,Array("Credits displayed",&H00001000)'dip 13
    .AddChk 205,110,190,Array("Play credit button tune",&H00002000)'dip 14
    .AddChk 205,125,190,Array("Play tones when scoring",&H00080000)'dip 20
    .AddChk 205,140,190,Array("Play coin switch tune",&H00400000)'dip 23
    .AddChk 205,155,190,Array("High game to date displayed",&H00100000)'dip 21
    .AddLabel 50,240,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
 End Sub
 Set vpmShowDips = GetRef("editDips")

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
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub a_Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  vpmTimer.PulseSw 10
  PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub a_Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub a_Posts_Hit(idx)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolFlip, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolFlip, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolFlip, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "BuckRogers" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / BuckRogers.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "BuckRogers" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / BuckRogers.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "BuckRogers" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / BuckRogers.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / BuckRogers.height-1
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

Const tnob = 10 ' total number of balls
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*10, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

' Thalamus : Exit in a clean and proper way
Sub BuckRogers_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

