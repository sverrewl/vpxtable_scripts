Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0
 
Const cGameName="genie",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const SCoin="coin",cCredits=""

LoadVPM"01120100","GTS1.VBS",3.02
Dim DesktopMode: DesktopMode = Table1.ShowDT

'*************************************************************

'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)="bsTrough.SolOut"
SolCallback(2)= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(3)= "vpmSolSound SoundFX(""Chime1"",DOFChimes),"
SolCallback(4)= "vpmSolSound SoundFX(""Chime2"",DOFChimes),"
SolCallback(5)= "vpmSolSound SoundFX(""Chime3"",DOFChimes),"
SolCallback(6)=  "bsSaucer.SolOut"
SolCallback(7)= "dtDrop1.SolDropUp"
SolCallback(8)= "dtDrop2.SolDropUp"
SolCallback(17)="vpmNudge.SolGameOn"
'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,LeftFlipper1,"
'SolCallback(sLRFlipper)="RightFliper"
 
  SolCallback(sLRFlipper) = "SolRFlipper"
  SolCallback(sLLFlipper) = "SolLFlipper"

  Sub SolLFlipper(Enabled)
     If Enabled Then
       PlaySound SoundFX("fx_FlipperUp",DOFFlippers):LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
       PlaySound SoundFX("fx_FlipperDown",DOFFlippers):LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
    End Sub
    
  Sub SolRFlipper(Enabled)
     If Enabled Then
       PlaySound SoundFX("fx_FlipperUp",DOFFlippers):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd:RightFlipper2.RotateToEnd
     Else
       PlaySound SoundFX("fx_FlipperDown",DOFFlippers):RightFlipper.RotateToStart:RightFlipper1.RotateToStart:RightFlipper2.RotateToStart
     End If
  End Sub

'Sub RightFliper(Enabled)
' If Enabled Then
'   VpmSolFlipper RightFlipper,RightFlipper1,True
'   RightFlipper2.RotateToEnd
' Else
'   VpmSolFlipper RightFlipper,RightFlipper1,False
'   RightFlipper2.RotateToStart
' End If
'End Sub
'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough,dtDrop1,dtDrop2,bsSaucer

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine="Genie (Gottlieb)"&chr(13)&"You Suck"
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
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3)

  Set bsTrough = New cvpmBallStack
  bsTrough.initnotrough BallExit, 66, 60, 14
  bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors),SoundFX("ballrelease",DOFContactors)

  Set bsSaucer=New cvpmBallstack
  bsSaucer.InitSaucer TopHole,41,234+rnd(1)*3,4+rnd(1)*4
  bsSaucer.InitExitSnd SoundFX("HoleKick",DOFContactors),SoundFX("SolOn",DOFContactors)

  Set dtDrop1=New cvpmDropTarget
  dtDrop1.InitDrop Array(sw20,sw21,sw23,sw24),Array(20,21,23,24)
  dtDrop1.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)

  Set dtDrop2=New cvpmDropTarget
  dtDrop2.InitDrop Array(sw30,sw70,sw31,sw71,sw60,sw74,sw61),Array(30,70,31,71,60,74,61)
  dtDrop2.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
End Sub 
'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
 Sub Table1_KeyDown(ByVal keycode)
  If vpmKeyDown(KeyCode) Then Exit Sub
  If keycode=PlungerKey Then Plunger.Pullback:playsound"plungerpull"
  If keycode=AddCreditKey then vpmTimer.pulseSW (swCoin1)
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If vpmKeyUp(KeyCode) Then Exit Sub
  If keycode=PlungerKey Then Plunger.Fire:playsound"plunger"
End Sub
 
'***********************************************************
'****     Rollover Wire Trigger CODE        ****
'***********************************************************
Sub SW10A_Hit:Controller.Switch(10)=1 : playsound"rollover" : DOF 103, DOFOn : End Sub
Sub SW10A_UnHit:Controller.Switch(10)=0: DOF 103, DOFOff : End Sub
Sub SW11A_Hit:Controller.Switch(11)=1 : playsound"rollover" : DOF 105, DOFOn : End Sub
Sub SW11A_UnHit:Controller.Switch(11)=0: DOF 105, DOFOff : End Sub
Sub SW13A_Hit:Controller.Switch(13)=1 : playsound"rollover" : DOF 107, DOFOn : End Sub
Sub SW13A_UnHit:Controller.Switch(13)=0: DOF 107, DOFOff : End Sub
Sub SW14A_Hit:Controller.Switch(14)=1 : playsound"rollover" : DOF 109, DOFOn : End Sub
Sub SW14A_UnHit:Controller.Switch(14)=0: DOF 109, DOFOff : End Sub

Sub SW10B_Hit:Controller.Switch(10)=1 : playsound"rollover" : DOF 104, DOFOn : End Sub
Sub SW10B_UnHit:Controller.Switch(10)=0:DOF 104, DOFOff :End Sub
Sub SW11B_Hit:Controller.Switch(11)=1 : playsound"rollover" : DOF 106, DOFOn : End Sub
Sub SW11B_UnHit:Controller.Switch(11)=0:DOF 106, DOFOff :End Sub
Sub SW13B_Hit:Controller.Switch(13)=1 : playsound"rollover" : DOF 108, DOFOn :End Sub
Sub SW13B_UnHit:Controller.Switch(13)=0:DOF 108, DOFOff :End Sub
Sub SW14B_Hit:Controller.Switch(14)=1 : playsound"rollover" : DOF 110, DOFOn : End Sub
Sub SW14B_UnHit:Controller.Switch(14)=0:DOF 110, DOFOff :End Sub

Sub SW54A_Hit:Controller.Switch(54)=1 : playsound"rollover" : DOF 115, DOFOn :End Sub
Sub SW54A_UnHit:Controller.Switch(54)=0:DOF 115, DOFOff :End Sub
Sub SW54B_Hit:Controller.Switch(54)=1 : playsound"rollover" : DOF 116, DOFOn :End Sub
Sub SW54B_UnHit:Controller.Switch(54)=0:DOF 116, DOFOff :End Sub

Sub SW63A_Hit:Controller.Switch(63)=1 : playsound"rollover" : DOF 112, DOFOn :End Sub
Sub SW63A_UnHit:Controller.Switch(63)=0:DOF 112, DOFOff :End Sub
Sub SW63B_Hit:Controller.Switch(63)=1 : playsound"rollover" : DOF 113, DOFOn :End Sub
Sub SW63B_UnHit:Controller.Switch(63)=0:DOF 113, DOFOff :End Sub
Sub SW63C_Hit:Controller.Switch(63)=1 : playsound"rollover" : DOF 114, DOFOn :End Sub
Sub SW63C_UnHit:Controller.Switch(63)=0:DOF 114, DOFOff :End Sub

Sub SW73A_Hit:Controller.Switch(73)=1 : playsound"rollover" : DOF 117, DOFOn :End Sub
Sub SW73A_UnHit:Controller.Switch(73)=0:DOF 117, DOFOff :End Sub
Sub SW73B_Hit:Controller.Switch(73)=1 : playsound"rollover" : DOF 118, DOFOn :End Sub
Sub SW73B_UnHit:Controller.Switch(73)=0:DOF 118, DOFOff :End Sub

'***********************************************************
'****       Star SWITCH CODE        ****
'***********************************************************
Sub SW53A_Hit:Controller.Switch(53)=1 : playsound"rollover" : DOF 122, DOFOn :End Sub 
Sub SW53A_UnHit:Controller.Switch(53)=0:DOF 122, DOFOff :End Sub
Sub SW53B_Hit:Controller.Switch(53)=1 : playsound"rollover" : DOF 123, DOFOn :End Sub 
Sub SW53B_UnHit:Controller.Switch(53)=0:DOF 123, DOFOff :End Sub
Sub SW53C_Hit:Controller.Switch(53)=1 : playsound"rollover" : DOF 124, DOFOn :End Sub 
Sub SW53C_UnHit:Controller.Switch(53)=0:DOF 124, DOFOff :End Sub
Sub SW64A_Hit:Controller.Switch(64)=1 : playsound"rollover" : DOF 119, DOFOn :End Sub
Sub SW64A_UnHit:Controller.Switch(64)=0:DOF 119, DOFOff :End Sub
Sub SW64B_Hit:Controller.Switch(64)=1 : playsound"rollover" : DOF 120, DOFOn :End Sub
Sub SW64B_UnHit:Controller.Switch(64)=0:DOF 120, DOFOff :End Sub

'***********************************************************
'****         SPINNER CODE          ****
'***********************************************************
Sub SW33_Spin:vpmTimer.PulseSw 33 : playsound"fx_spinner" : End Sub

'***********************************************************
'****       STAND UP TARGET CODE        ****
'***********************************************************
Sub SW40_Hit:vpmTimer.PulseSw 40 : End Sub
Sub SW44_Hit:vpmTimer.PulseSw 44 : End Sub
Sub SW51_Hit:vpmTimer.PulseSw 51 : End Sub
Sub SW63_Hit:vpmTimer.PulseSw 63 : DOF 111, DOFPulse : End Sub

'***********************************************************
'****         SAUCER CODE           ****
'***********************************************************
Sub TopHole_Hit():bsSaucer.Addball 0:End Sub  
Sub Drain_Hit:playsound"drain":bsTrough.addball me:End Sub

'***********************************************************
'****         SLING CODE            ****
'***********************************************************
'Non animated Slingshots
Sub sw50a_Hit:vpmTimer.PulseSw 50 : playsound"slingshot": End Sub 
Sub sw50d_Hit:vpmTimer.PulseSw 50 : playsound"slingshot" : End Sub 
Sub SW50e_Hit:vpmTimer.PulseSw 50 : playsound"slingshot" : End Sub 
Sub SW50f_Hit:vpmTimer.PulseSw 50 : playsound"slingshot" : End Sub 
Sub sw50g_Hit:vpmTimer.PulseSw 50 : playsound"slingshot" : End Sub 
Sub sw50h_Hit:vpmTimer.PulseSw 50 : playsound"slingshot" : End Sub 
Sub sw50i_Hit:vpmTimer.PulseSw 50 : playsound"slingshot" : End Sub 

'***********************************************************
'****         DROP TARGET CODE        ****
'***********************************************************
 Sub Sw20_Hit: dtDrop1.Hit 1: End Sub 
 Sub Sw21_Hit: dtDrop1.Hit 2: End Sub 
 Sub Sw23_Hit: dtDrop1.Hit 3: End Sub 
 Sub Sw24_Hit: dtDrop1.Hit 4: End Sub 


 Sub Sw30_Hit: dtDrop2.Hit 1: End Sub 
 Sub Sw70_Hit: dtDrop2.Hit 2: End Sub 
 Sub Sw31_Hit: dtDrop2.Hit 3: End Sub 
 Sub Sw71_Hit: dtDrop2.Hit 4: End Sub 
 Sub Sw60_Hit: dtDrop2.Hit 5: End Sub 
 Sub Sw61_Hit: dtDrop2.Hit 6: End Sub 
 Sub Sw74_Hit: dtDrop2.Hit 7: End Sub 


'***********************************************************
'****         BUMPER CODE           ****
'***********************************************************
Sub Bumper1_Hit : vpmTimer.PulseSw(43) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(53) : playsound SoundFXDOF("fx_bumper1",101,DOFPulse,DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(53) : playsound SoundFXDOF("fx_bumper1",102,DOFPulse,DOFContactors): End Sub

'******************************************************************************
'Primitive Flipper Code
Sub FlipperTimer_Timer
  FlipperT1.roty = LeftFlipper.currentangle  + 90
  FlipperT5.roty = RightFlipper.currentangle + 90
  FlipperT2.roty = RightFlipper1.currentangle + 90
End Sub
'******************************************************************************

'**********************************************************************************************************
 
'Map lights to an array
'**********************************************************************************************************

Set Lights(4) = l4p ' Shoot Again Playfied and backglass
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
Set Lights(36) = l36
Set Lights(150) = Light1 'Star Trigger 1
Set Lights(151) = Light2 'Star Trigger 2
Set Lights(152) = Light3 'Star Trigger 3
Set Lights(153) = Light4 'Star Trigger 4
Set Lights(154) = Light5 'Star Trigger 5

'Backglass
'Set Lights(1) = l1 'High Score
'Set Lights(2) = l2 'Tilt
'Set Lights(3) = l3 'High Score
'**********************************************************************************************************
' Backglass Light Displays (7 digit 7 segment displays)
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
 Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
    With vpmDips
      .AddForm 700,400,"System 1 (Multi-Mode sound) - DIP switches"
      .AddFrame 0,0,190,"Coin chute control",&H00040000,Array("seperate",0,"same",&H00040000)'dip 19
      .AddFrame 0,46,190,"Game mode",&H00000400,Array("extra ball",0,"replay",&H00000400)'dip 11
      .AddFrame 0,92,190,"High game to date awards",&H00200000,Array("no award",0,"3 replays",&H00200000)'dip 22
      .AddFrame 0,138,190,"Balls per game",&H00000100,Array("5 balls",0,"3 balls",&H00000100)'dip 9
      .AddFrame 0,184,190,"Tilt effect",&H00000800,Array("game over",0,"ball in play only",&H00000800)'dip 12
      .AddFrame 205,0,190,"Maximum credits",&H00030000,Array("5 credits",0,"8 credits",&H00020000,"10 credits",&H00010000,"15 credits",&H00030000)'dip 17&18
      .AddFrame 205,76,190,"Sound settings",&H80000000,Array("sounds",0,"tones",&H80000000)'dip 32
      .AddFrame 205,122,190,"Attract tune",&H10000000,Array("no attract tune",0,"attract tune played every 6 minutes",&H10000000)'dip 29
      .AddChk 205,175,190,Array("Match feature",&H00000200)'dip 10
      .AddChk 205,190,190,Array("Credits displayed",&H00001000)'dip 13
      .AddChk 205,205,190,Array("Play credit button tune",&H00002000)'dip 14
      .AddChk 205,220,190,Array("Play tones when scoring",&H00080000)'dip 20
      .AddChk 205,235,190,Array("Play coin switch tune",&H00400000)'dip 23
      .AddChk 205,250,190,Array("High game to date displayed",&H00100000)'dip 21
      .AddLabel 50,280,300,20,"After hitting OK, press F3 to reset game with new settings."
      .ViewDips
    End With
 End Sub
 Set vpmShowDips = GetRef("editDips")

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

'VPX Functions
'**********************************************************************************************************
'**********************************************************************************************************


'*****GI Lights On
dim xx

For each xx in GI:xx.State = 1: Next

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 50
    PlaySound "left_slingshot", 0, 1, 0.05, 0.05
    'RSling.Visible = 0
    'RSling1.Visible = 1
    'sling1.TransZ = -20
    'RStep = 0
    'RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 50
    PlaySound SoundFXDOF("right_slingshot",121,DOFPulse,DOFcontactors),0,1,-0.05,0.05
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



' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
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
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
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


Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub
