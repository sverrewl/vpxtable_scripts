Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


' Thalamus 2019 August : Improved directional sounds
' !! NOTE : Table not verified yet !!
' changed useSolenoids to 2

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 3000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume
Const VolBump   = 1    ' Bumpers volume.

Const cGameName="sprbrks",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="Flipper2",SFlipperOff="Flipper"
Const SCoin="coin3",cCredits=""

Const Ballsize = 50
'Const BallMass = 1.5


LoadVPM "01200000","sys80.vbs",3.2
Set LampCallback=GetRef("UpdateMultipleLamps")

gion()


sub gion()
gi_1.state=1
gi_2.state=1
gi_5.state=1
GI_6.state=1
GI_7.state=1
gi_9.state=1
gi_11.state=1
GI_12.state=1
gi_13.state=1
gi_15.state=1
gi_20.state=1
gi_21.state=1
GI_22.state=1
GI_25.state=1
GI_28.state=1
GI_29.state=1
GI_30.state=1
gi_31.state=1
GI_32.state=1
GI_33.state=1
GI_34.state=1
GI_35.state=1
GI_36.state=1
GI_43.state=1
GI_48.state=1
GI_49.state=1
GI_50.state=1
GI_53.state=1
GI_54.state=1
End Sub

sub gioff()
gi_1.state=0
gi_2.state=0
gi_5.state=0
GI_6.state=0
GI_7.state=0
gi_9.state=0
gi_11.state=0
GI_12.state=0
gi_13.state=0
gi_15.state=0
gi_20.state=0
gi_21.state=0
GI_22.state=0
GI_25.state=0
GI_28.state=0
GI_29.state=0
GI_30.state=0
gi_31.state=0
GI_32.state=0
GI_33.state=0
GI_34.state=0
GI_35.state=0
GI_36.state=0
GI_43.state=0
GI_48.state=0
GI_49.state=0
GI_50.state=0
GI_53.state=0
GI_54.state=0
End Sub
ball1.CreateSizedBall(25)
  ball1.Kick 180,1
  ball1.enabled=false
ball2.CreateSizedBall(25)
  ball2.Kick 180,1
  ball2.enabled=false
ball3.CreateSizedBall(25)
  ball3.Kick 180,1
  ball3.enabled=false
ball4.CreateSizedBall(25)
  ball4.Kick 180,1
  ball4.enabled=false

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(1)="SolOne" 'ramp flasher
SolCallback(2)="dtC.SolDropUp" 'drop targets
SolCallback(3)="K3" 'top kicker
SolCallback(4)="K2" 'launch ball
SolCallback(5)="dtR.SolDropUp" 'drop targets
SolCallback(6)="dtT.SolDropUp" 'drop targets
SolCallback(7)="vpmFlasher Array(sb1,sb2,sb3,sb4,sb5,sb6,sb7,sb8)," 'flasher 8 lights
SolCallback(8)="vpmSolSound""knocker""," 'knocker
SolCallback(9)="k4" 'outhole


Sub SolOne(Enabled)
  If Enabled Then
  lightsol1.state=1
  lightsol1a.state=1
  Else
  lightsol1.state=0
  lightsol1a.state=0
  End If
End Sub
Sub sollFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub

Sub solrflipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, 1:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, 1:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Dim FlipperActive




'Primitive Flipper
Sub FlipperTimer_Timer
  FlipperT1.roty = LeftFlipper.currentangle  + 237
  FlipperT5.roty = RightFlipper.currentangle + 120
  FlipperT2.roty = LeftFlipper1.currentangle  + 237
  rightFlipper1Bat.roty = RightFlipper1.currentangle + 120
  rightFlipper1Rubber.roty = RightFlipper1.currentangle + 120
  rightFlipper1screw.roty = RightFlipper1.currentangle + 120

End Sub


Dim VarRol, VarHidden

Dim L4,OL4
OL4=0



Sub UpdateMultipleLamps

light15a.State = light15.State





End Sub



'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsKicker, bsHole, DTT, DTR, dtc, bssaucer,k1

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Spring Break"
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

    vpmNudge.TiltSwitch=57
    vpmNudge.Sensitivity=5
  vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3)



  Set dtC=New cvpmDropTarget
   dtC.InitDrop Array(Targ2a,Targ2b,Targ2c),Array(50,60,70)
   dtC.InitSnd "flapclos","flapopen"


   Set dtT=New cvpmDropTarget
   dtT.InitDrop Array(Targ3a,Targ3b,Targ3c),Array(52,62,72)
   dtT.InitSnd "flapclos","flapopen"


   Set dtR=New cvpmDropTarget
   dtR.InitDrop Array(Targ1a,Targ1b,Targ1c),Array(51,61,71)
   dtR.InitSnd "flapclos","flapopen"





End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)

If KeyCode=RightFlipperKey Then Controller.Switch(35)=1

If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)

If KeyCode=RightFlipperKey Then Controller.Switch(35)=0
If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'**********************************************************************************************************




Sub Targ2a_dropped:dtc.Hit 1:End Sub
Sub Targ2b_dropped:dtc.Hit 2:End Sub
Sub Targ2c_dropped:dtc.Hit 3:End Sub
Sub Targ3a_dropped:DTT.Hit 1:End Sub
Sub Targ3b_dropped:DTT.Hit 2:End Sub
Sub Targ3c_dropped:DTT.Hit 3:End Sub
Sub Targ1a_dropped:DTR.Hit 1:End Sub
Sub Targ1b_dropped:DTR.Hit 2:End Sub
Sub Targ1c_dropped:DTR.Hit 3:End Sub
Sub sw34_Hit:Controller.Switch(34)=1:End Sub  'switch 34
Sub sw34_unHit:Controller.Switch(34)=0:End Sub  'switch 34
Sub sw74_Hit:Controller.Switch(74)=1:End Sub  'switch 74
Sub sw74_unHit:Controller.Switch(74)=0:End Sub  'switch 74
Sub sw54_Hit:Controller.Switch(54)=1:End Sub  'switch 54
Sub sw54_unHit:Controller.Switch(54)=0:End Sub  'switch 54
Sub sw64_Hit:Controller.Switch(64)=1:End Sub  'switch 64
Sub sw64_unHit:Controller.Switch(64)=0:End Sub  'switch 64

Sub sw20_Hit:Controller.Switch(20)=1:End Sub  'switch 20
Sub sw20_unHit:Controller.Switch(20)=0:End Sub  'switch 20
Sub sw30_Hit:Controller.Switch(30)=1:End Sub  'switch 30
Sub sw30_unHit:Controller.Switch(30)=0:End Sub  'switch 30
Sub sw40_Hit:Controller.Switch(40)=1:End Sub  'switch 40
Sub sw40_unHit:Controller.Switch(40)=0:End Sub  'switch 40
Sub sw21_Hit:Controller.Switch(21)=1:End Sub'switch 21
Sub sw21_unHit:Controller.Switch(21)=0:End Sub'switch 21

Sub sw55_Hit:Controller.Switch(55)=1:End Sub'switch 55
Sub sw55_unHit:Controller.Switch(55)=0:End Sub'switch 55
Sub sw65_Hit:Controller.Switch(65)=1:End Sub'switch 65
Sub sw65_unHit:Controller.Switch(65)=0:End Sub'switch 65
Sub sw75_Hit:Controller.Switch(75)=1:End Sub'switch 75
Sub sw75_unHit:Controller.Switch(75)=0:End Sub'switch 75
Sub sw66_Hit:Controller.Switch(66)=1:End Sub'switch 66
Sub sw66_unHit:Controller.Switch(66)=0:End Sub'switch 66


Sub sw23_Hit:vpmTimer.PulseSw(23):End Sub 'switch 23
Sub sw31_Hit:vpmTimer.PulseSw(31):End Sub 'switch 31
Sub sw32_Hit:vpmTimer.PulseSw(32):End Sub 'switch 32
Sub sw33_Hit:vpmTimer.PulseSw(33):End Sub 'switch 33
Sub sw41_Spin:vpmTimer.PulseSw(41):End Sub  'switch 41
Sub sw42_Hit:vpmTimer.PulseSw(42):End Sub 'switch 42
Sub sw43_Hit:vpmTimer.PulseSw(43):End Sub 'switch 43
                      'switch 57 Tilt

Sub Drain_Hit:bsTrough.AddBall Me: End Sub
Sub Kicker1_Hit:K1.AddBall 0:End Sub
Sub KickerRight_Hit:bsSaucer.AddBall 0:End Sub


Sub Bumper1_Hit:vpmTimer.PulseSw(44):RandomSoundbumper:End Sub  'switch 44
Sub Bumper2_Hit:vpmTimer.PulseSw(24):RandomSoundbumper:End Sub  'switch 24
Sub Bumper3_Hit:vpmTimer.PulseSw(34):RandomSoundbumper:End Sub  'switch 34

 Plunger1.fire
Sub sw53_Hit:Controller.Switch(53)=1:End Sub'switch 53
Sub sw53_unHit:Controller.Switch(53)=0:End Sub'switch 53
Sub sw63_Hit:Controller.Switch(63)=1:End Sub'switch 63
Sub sw63_unHit:Controller.Switch(63)=0:End Sub'switch 63
Sub sw73_Hit:Controller.Switch(73)=1:End Sub'switch 73
Sub sw73_unHit:Controller.Switch(73)=0:End Sub'switch 73
Sub sw45_Hit:Controller.Switch(45)=1:End Sub'switch 45
Sub sw45_unHit:Controller.Switch(45)=0:End Sub'switch 45
Sub sw22_Hit:Controller.Switch(22)=1:End Sub'switch 22
Sub sw22_unHit:Controller.Switch(22)=0:End Sub'switch 22
sub k2(Enabled)
 Kicker2.kick 0,50
End Sub
sub k3(Enabled)
 Kicker3.kick 0,50
End Sub
sub k4(Enabled)
 Kicker4.kick 70,20
End Sub
Dim BallsInPlay



'**********************************************************************************************************
' Map Lights to Array
'**********************************************************************************************************
set Lights(13)=light13
set lights(12)=light12
set Lights(2)=light2
set lights(3)=Light3
set lights(5)=light5
  set lights(4)=light4
set lights(6)=light6
set lights(7)=light7
set lights(8)=light8
set lights(9)=light9
set lights(10)=light10
set lights(11)=light11
  set lights(1)=light1
  set lights(14)=light14
set lights(15)=Light15
set lights(16)=Light16
set lights(17)=Light17
set lights(18)=Light18
set lights(19)=Light19
set lights(20)=Light20
set lights(21)=Light21
set lights(22)=light22
set lights(23)=light23
set lights(24)=light24
set lights(25)=light25
set lights(26)=light26
set lights(27)=light27
set lights(28)=light28
set lights(29)=light29
set lights(30)=light30
set lights(31)=light31
set lights(32)=light32
set lights(33)=light33
set lights(34)=light34
set lights(35)=light35
set lights(36)=light36
set lights(37)=light37
set lights(38)=light38
set lights(39)=light39
set lights(40)=light40
set lights(41)=light41
set lights(42)=light42
set lights(43)=light43
set lights(44)=light44
set lights(45)=light45
set lights(46)=light46
set lights(47)=light47














'**********************************************************************************************************
' Backglass Light Displays
'**********************************************************************************************************


Sub LampTimer_Timer()

light15a.State = light15.State


light28a.state=light28.state
light29a.state=light29.state

if light13.state =1 Then
loopr1.state=2
loopr2.state=2
loopr3.state=2
loopr4.state=2
loopr5.state=2
loopr6.state=2
loopr7.state=2
loopr8.state=2
loopr9.state=2
loopr10.state=2
Else
loopr1.state=0
loopr2.state=0
loopr3.state=0
loopr4.state=0
loopr5.state=0
loopr6.state=0
loopr7.state=0
loopr8.state=0
loopr9.state=0
loopr10.state=0
end If

if light16.state = 1 Then
light16f.visible=1
light16f1.visible=1
Else
light16f.visible=0
light16f1.visible=0
end If
if light17.state = 1 Then
light17f.visible=1
light17f1.visible=1
Else
light17f.visible=0
light17f1.visible=0
end If
if light18.state = 1 Then
light18f.visible=1
light18f1.visible=1
Else
light18f.visible=0
light18f1.visible=0
end If
if light19.state = 1 Then
light19f.visible=1
light19f1.visible=1
Else
light19f.visible=0
light19f1.visible=0
end If
if light20.state = 1 Then
light20f.visible=1
light20f1.visible=1
Else
light20f.visible=0
light20f1.visible=0
end If

if light2.state = 1 Then
Flasher2a.visible=1
Flasher2.visible=1
Else
Flasher2a.visible=0
Flasher2.visible=0
end If

if light1.state=0 Then
 gion()
Else
 gioff()
end If

L4=light12.State
  If L4<>OL4 Then
    If L4=1 Then

releasein.collidable=1
releaseout.collidable=0
Else
  releaseout.collidable=1
releasein.collidable=0



    End If

  End If


  OL4=L4
If Table1.showDT = True Then Display

end Sub
'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
 Dim Digits(40)
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
 Digits(32)=Array(ac18, ac16, acc1, acd1, ac19, ac17, ac15, acf1, ac11, ac13, ac12, ac14, acb1, aca1, ac10, ace1)
 Digits(33)=Array(ad18, ad16, adc1, add1, ad19, ad17, ad15, adf1, ad11, ad13, ad12, ad14, adb1, ada1, ad10, ade1)
 Digits(34)=Array(ae18, ae16, aec1, aed1, ae19, ae17, ae15, aef1, ae11, ae13, ae12, ae14, aeb1, aea1, ae10, aee1)
 Digits(35)=Array(af18, af16, afc1, afd1, af19, af17, af15, aff1, af11, af13, af12, af14, afb1, afa1, af10, afe1)
 Digits(36)=Array(b9, b7, b0c1, b0d1, b100, b8, b6, b0f1, b2, b4, b3, b5, b0b1, b0a1, b1,b0e1)
 Digits(37)=Array(b109, b107, b1c1, b1d1, b110, b108, b106, b1f1, b102, b104, b103, b105, b1b1, b1a1, b101,b1e1)
 Digits(38)=Array(b119, b117, b2c1, b2d1, b120, b118, b116, b2f1, b112, b114, b113, b115, b2b1, b2a1, b111, b2e1)
 Digits(39)=Array(b129, b127, b3c1, b3d1, b130, b128, b126, b3f1, b122, b3b1, b123, b125, b3b1, b3a1, b121, b3e1)


 Sub Display ()
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 40) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
      Else
             end if
        Next
    End If
 End Sub
'**********************************************************************************************************
'**********************************************************************************************************








'Gottlieb Spring Break
'added by Inkochnito
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm  700,400,"Spring Break - DIP switches"
    .AddFrame 2,4,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"20 credits",49152)'dip 15&16
    .AddFrame 2,80,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 2,126,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddFrame 2,172,190,"High games to date control",&H00000020,Array("no effect",0,"reset high games 2-5 on power off",&H00000020)'dip 6
    '.AddFrame 2,218,190,"Auto-percentage control",&H00000080,Array("disabled (normal high score mode)",0,"enabled",&H00000080)'dip 8
    .AddFrame 2,218,190,"Game start multiball",&H40000000,Array("no multiball at the start",0,"start with multiball",&H40000000)'dip 31
    .AddFrame 2,264,190,"Game option",&H80000000,Array("off",0,"on",&H80000000)'dip 32
    .AddFrame 205,4,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
    .AddFrame 205,80,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
    .AddFrame 205,126,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
    .AddFrame 205,172,190,"Novelty",&H08000000,Array("normal",0,"extra ball and replay scores 500K",&H08000000)'dip 28
    .AddFrame 205,218,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
    .AddFrame 205,264,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
    .AddChk 2,316,180,Array("Match feature",&H02000000)'dip 26
    .AddLabel 50,335,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips = GetRef("editDips")

 Sub Table1_Exit ' in some tables this needs to be Table1_Exit
     Controller.Stop
 End Sub

'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw(25)
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

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
' PlaySoundAtVol sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
    End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.

Sub rollingmetal_Hit (idx)
  PlaySound "Wire Ramp", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub


Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub



Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub


Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub



Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub



Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandomSoundbumper()
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtVol SoundFX("fx_Bumper",DOFContactors), ActiveBall, VolBump
    Case 2 : PlaySoundAtVol SoundFX("fx_Bumper1",DOFContactors), ActiveBall, VolBump
    Case 3 : PlaySoundAtVol SoundFX("fx_Bumper2",DOFContactors), ActiveBall, VolBump
    Case 4 : PlaySoundAtVol SoundFX("fx_bumper3",DOFContactors), ActiveBall, VolBump
    Case 5 : PlaySoundAtVol SoundFX("fx_bumper4",DOFContactors), ActiveBall, VolBump
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub




