Option Explicit
Randomize
 
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0
 
Const cGameName = "evelknie",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin"
Const BallMass = 1.7
 
LoadVPM"01000100","BALLY.VBS",1.2
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
SolCallback(2)    =  "vpmSolSound SoundFX(""Bell10"",DOFChimes),"
SolCallback(3)    =  "vpmSolSound SoundFX(""Bell100"",DOFChimes),"
SolCallback(4)    =  "vpmSolSound SoundFX(""Bell1000"",DOFChimes),"
SolCallback(5)    =  "vpmSolSound SoundFX(""Bell10000"",DOFChimes),"
SolCallback(6)    =  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(7)    =  "bsTrough.SolOut"
SolCallback(8)    =  "bsSaucer.SolOut"
SolCallback(13)   =  "dtR.SolDropUp"
SolCallback(19)   =  "vpmNudge.SolGameOn"
 
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
 
Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFFlippers):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFFlippers):LeftFlipper.RotateToStart
     End If
  End Sub
 
Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFFlippers):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFFlippers):RightFlipper.RotateToStart
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
 
Dim bsTrough, bsSaucer, dtR
 
Sub Table1_Init
    vpmInit Me
    On Error Resume Next
        With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
        .SplashInfoLine = "Evel Knievel Bally"&chr(13)&"You Suck"
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
 
    PinMAMETimer.Interval=PinMAMEInterval
    PinMAMETimer.Enabled=1
 
    vpmNudge.TiltSwitch = 7
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, Bumper5, LeftSlingshot, RightSlingshot)
 
    Set bsTrough=New cvpmBallStack
        bsTrough.InitNoTrough BallRelease,8,70,8
        bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
 
    Set bsSaucer=New cvpmBallStack
        bsSaucer.InitSaucer SW29,29,180,8
        bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsSaucer.KickAngleVar=10
 
    Set dtR=New cvpmDropTarget
        dtR.InitDrop Array(SW1,SW2,SW3,SW4,SW5),Array(1,2,3,4,5)
        dtR.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
 
End Sub
 
 '**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
 Sub Table1_KeyDown(ByVal keycode)
    If vpmKeyDown(KeyCode) Then Exit Sub
    If keycode=PlungerKey Then Plunger.Pullback:playsound"plungerpull"
End Sub
 
Sub Table1_KeyUp(ByVal keycode)
    If vpmKeyUp(KeyCode) Then Exit Sub
    If keycode=PlungerKey Then Plunger.Fire:playsound"plunger"
End Sub
'**********************************************************************************************************
 
 ' Switches
'**********************************************************************************************************
 
 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub
Sub SW29_Hit:bsSaucer.AddBall 0 : playsound "popper_ball": End Sub
 
 ' Droptargets
Sub SW1_Dropped:dtR.Hit 1:End Sub
Sub SW2_Dropped:dtR.Hit 2:End Sub
Sub SW3_Dropped:dtR.Hit 3:End Sub
Sub SW4_Dropped:dtR.Hit 4:End Sub
Sub SW5_Dropped:dtR.Hit 5:End Sub
 
'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(40) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(39) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(38) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper4_Hit : vpmTimer.PulseSw(25) : DOF 101, DOFPulse : End Sub
Sub Bumper5_Hit : vpmTimer.PulseSw(25) : DOF 102, DOFPulse : End Sub
 
 ' Rollovers
Sub SW20_Hit:Controller.Switch(20)=1 : playsound"rollover" : End Sub
Sub SW20_unHit:Controller.Switch(20)=0:End Sub
Sub SW21_Hit:Controller.Switch(21)=1 : playsound"rollover" : End Sub
Sub SW21_unHit:Controller.Switch(21)=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1 : playsound"rollover" : End Sub
Sub SW23_unHit:Controller.Switch(23)=0:End Sub
Sub SW24_Hit:Controller.Switch(24)=1 : playsound"rollover" : End Sub
Sub SW24_unHit:Controller.Switch(24)=0:End Sub
Sub SW32A_Hit:Controller.Switch(32)=1 : playsound"rollover" : DOF 105, DOFOn : End Sub
Sub SW32A_unHit:Controller.Switch(32)=0: DOF 105, DOFOff : End Sub
Sub SW32B_Hit:Controller.Switch(32)=1 : playsound"rollover" : DOF 106, DOFOn : End Sub
Sub SW32B_unHit:Controller.Switch(32)=0: DOF 106, DOFOff : End Sub
 
 ' Stand Up Targets
Sub SW17_Hit:vpmTimer.PulseSw 17:End Sub
Sub SW18_Hit:vpmTimer.PulseSw 18:End Sub
Sub SW19_Hit:vpmTimer.PulseSw 19:End Sub
Sub SW22_Hit:vpmTimer.PulseSw 22:End Sub
Sub SW26A_Hit:vpmTimer.PulseSw 26:DOF 103, DOFPulse:End Sub
Sub SW26B_Hit:vpmTimer.PulseSw 26:DOF 104, DOFPulse:End Sub
 
'Spinners
Sub sw30_Spin:vpmTimer.PulseSw 30 : playsound"fx_spinner" : End Sub
Sub sw31_Spin:vpmTimer.PulseSw 31 : playsound"fx_spinner" : End Sub
 
'**********************************************************************************************************
'Map lights to an array
'**********************************************************************************************************
 
Set Lights(1)=l1
Set Lights(2)=l2
Set Lights(3)=l3
Lights(4)=array(l4,l4a)
Lights(5)=array(l5,L5A)
Set Lights(7)=l7
Set Lights(8)=l8
Set Lights(9)=l9
Set Lights(10)=l10
Set Lights(17)=l17
Set Lights(18)=l18
Set Lights(19)=l19
Lights(20)=array(l20,L20A)
Set Lights(22)=l22
Set Lights(23)=l23
Set Lights(25)=l25
Set Lights(26)=l26
Set Lights(33)=l33
Set Lights(34)=l34
Lights(36)=array(l36,L36a)
Set Lights(37)=l37 'Bumper 3 Light
Set Lights(38)=l38
Set Lights(39)=l39
Set Lights(40)=l40
Set Lights(41)=l41
Set Lights(42)=l42
Set Lights(43)=l43
Set Lights(49)=l49
Set Lights(50)=l50
Set Lights(51)=l51
Lights(52)=array(l52,L52A)
Lights(53)=array(l53,L53a) 'Bumper 1 and Bumper 2 Light
Set Lights(54)=l54
Set Lights(55)=l55
Set Lights(56)=l56
Set Lights(57)=l57
Set Lights(59)=l59
 
'backglass
'Set Lights(11)=Light11 'Extra Ball
'Set Lights(15)=Light15 'Player 1
'Set Lights(31)=Light31 'Player 2
'Set Lights(47)=Light47 'Player 3
'Set Lights(63)=Light63 'Player 4
'Set Lights(13)=Light13 'Ball in Play
'Set Lights(27)=Light27 'Match
'Set Lights(29)=Light29 'High Score
'Set Lights(43)=Light43 'Same Player Shoot Again
'Set Lights(45)=Light45 'game over
'Set Lights(61)=Light61 'tilt
 
 
'**********************************************************************************************************
' Backglass Light Displays (7 digit 7 segment displays)
Dim Digits(28)
 
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
 
 
' 2nd Player
Digits(6) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(7) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(8) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(9) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(10) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(11) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
 
 
' 3rd Player
Digits(12) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(13) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(14) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(15) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(16) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(17) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
 
 
' 4th Player
Digits(18) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(19) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(20) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(21) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(22) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(23) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
 
 
' Credits
Digits(24) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(25) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(26) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(27) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)
 
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
                if char(stat) > "" then msg(num) = char(stat)
            end if
        next
        end if
end if
End Sub
 
'**********************************************************************************************************
'**********************************************************************************************************
 
'Bally Evel Knievel
'added by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips=New cvpmDips
    With vpmDips
        .AddForm 700,400,"Evel Knievel - DIP switches"
        .AddFrame 2,2,190,"Maximum credits",&H00070000,Array("5 credits",0,"10 credits",&H00010000,"15 credits",&H00020000,"20 credits",&H00030000,"25 credits",&H00040000,"30 credits",&H00050000,"35 credits",&H00060000,"40 credits",&H00070000)'dip 17&18&19
        .AddFrame 2,134,190,"High game to date award",&H00006000,Array("no award",0,"1 credit",&H00002000,"2 credits",&H00004000,"3 credits",&H00006000)'dip 14&15
        .AddFrame 2,211,190,"High score feature",&H00000020,Array("extra ball",0,"replay",&H00000020)'dip 6
        .AddFrame 2,258,190,"Balls per game", 32768,Array("3 balls",0,"5 balls", 32768)'dip 16
        .AddFrame 2,306,190,"CYCLE outlane special",&H00800000,Array("alternating from side to side",0,"both outlanes lit",&H00800000)'dip 24
        .AddFrame 205,2,190,"Completing SUPER award",&H03000000,Array("no award",0,"extra ball",&H01000000,"1 replay",&H02000000,"2 replays",&H03000000)'dip 25&26
        .AddFrame 205,78,190,"Outlane special award",&HC0000000,Array("no award",0,"extra ball",&H40000000,"replay",&H80000000,"replay and extra ball",&HC0000000)'dip 31&32
        .AddFrame 205,154,190,"Drop targets 2nd time down",&H0C000000,Array("no award",0,"5,000 points",&H04000000,"5,000 points, extra ball",&H08000000,"extra ball",&H0C000000)'dip 27&28
        .AddFrame 205,230,190,"Drop targets 3rd time down",&H30000000,Array("no award",0,"5,000 points",&H10000000,"extra ball",&H20000000,"replay",&H30000000)'dip 29&30
        .AddChk 205,310,120,Array("Match feature",&H00100000)'dip 21
        .AddChk 205,325,120,Array("Credits display",&H00080000)'dip 20
        .AddChk 205,340,120,Array("Melody option",&H00000080)'dip 8
        .AddLabel 50,370,300,20,"After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips=GetRef("editDips")
 
'**********************************************************************************************************
'**********************************************************************************************************
 
'**********************************************************************************************************
'**********************************************************************************************************
 
 
'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep
 
Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw 36
    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
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
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
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
 
'*****************************************
'           FLIPPER SHADOWS
'*****************************************
 
sub FlipperTimer_Timer()
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle
 
End Sub
 
'*****************************************
'           BALL SHADOW
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
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