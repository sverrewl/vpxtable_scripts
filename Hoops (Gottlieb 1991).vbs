Option Explicit
Randomize

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
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="hoops",UseSolenoids=1,UseLamps=1,UseGI=0,UseSync=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin"
Const swStartButton=3 'remap start button to 1 key

LoadVPM "01120100","gts3.vbs",3.02

Dim FlipperShadows
Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Ramp17.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Ramp17.visible=0
End if

'**********************************************************************************************************
'Solenoid Call backs
'**********************************************************************************************************

SolCallback(5)="bsL.SolOut"
SolCallback(6)="bsT.SolOut"
SolCallback(7)="bsR.SolOut"
'SolCallback(8)=Not Used
SolCallback(9)="vpmFlasher array(Flasher1,Flasher1a)," 'Bottom Left Flasher
SolCallback(10)="vpmFlasher array(Flasher2,Flasher2a),"'Left Side Flasher
SolCallback(11)="vpmFlasher array(Flasher3,Flasher3a,Flasher3a1),"'Top Left Flasher
SolCallback(12)="vpmFlasher array(Flasher4,Flasher4a,Flasher4a1),"'Top Right Flasher
SolCallback(13)="vpmFlasher array(Flasher5,Flasher5a),"'Right Side Flasher
SolCallback(14)="vpmFlasher array(Flasher6,Flasher6a),"'Bottom Right Flasher
'SolCallback(15)=Not Used
'SolCallback(16)=Not Used
SolCallback(17)="vpmFlasher array(Flasher7,Flasher7a),"'#1 FlashSet (2 Lamps)
SolCallback(18)="vpmFlasher array(Flasher8,Flasher8a),"'#2 FlashSet (2 Lamps)
SolCallback(19)="vpmFlasher array(Flasher9,Flasher9a),"'#3 FlashSet (2 Lamps)
SolCallback(20)="vpmFlasher array(Flasher10,Flasher10a),"'#4 FlashSet (2 Lamps)
SolCallback(21)="vpmFlasher array(Flasher11,Flasher11a),"'#5 FlashSet (2 Lamps)
SolCallback(22)="vpmFlasher array(Flasher12,Flasher12a),"'#6 FlashSet (2 Lamps)
'SolCallback(23)=Not Used
'SolCallback(24)=Not Used
'SolCallback(25)=Not Used
'SolCallback(26)='SolLight Box Relay
'SolCallback(27)='Ticket Dispenser
SolCallback(28)="bsTrough.SolOut"
SolCallback(29)="bsTrough.SolIn"
SolCallback(30)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
'SolCallback(31)='Tilt Relay
SolCallback(32)="vpmNudge.SolGameOn"'Game Over Relay

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
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart
     End If
End Sub

'**********************************************************************************************************
'GI Lights On
'**********************************************************************************************************

sub Encendido_timer
dim xx
playsound "encendido"
For each xx in Ambiente:xx.State = 1: Next
    me.enabled=false
End Sub

'**********************************************************************************************************
'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsL, bsR, bsT

Sub Table1_Init
    vpmInit Me
    On Error Resume Next
    With Controller
        .GameName=cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
        .SplashInfoLine = "Hoops - Gottlieb 1991"&chr(13)&"By Kalavera"
        .HandleKeyboard=0
        .ShowTitle=0
        .ShowDMDOnly=1
        .HandleMechanics=0
        .ShowFrame=0
        .Hidden=1
         On Error Resume Next
        .Run GetPlayerHwnd
        If Err Then MsgBox Err.Description
         On Error Goto 0
    End With
    On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
    PinMAMETimer.Enabled=1
    vpmNudge.TiltSwitch=151
    vpmNudge.Sensitivity=2
    vpmNudge.TiltObj=Array(Bumper1,Bumper2,LeftSlingShot,RightSlingShot)

    Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 47,0,0,37,0,0,0,0
        bsTrough.InitKick BallRelease,62,8
        bsTrough.InitExitSnd SoundFX("Ballrelease",DOFContactors),SoundFX("SolOn",DOFContactors)
        bsTrough.Balls=3

    Set bsL=New cvpmBallStack
        bsL.InitSaucer S14,14,146,8
        bsL.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("SolOn",DOFContactors)
        bsL.KickForceVar=3

    Set bsR=New cvpmBallStack
        bsR.InitSaucer S16,16,247,8
        bsR.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("SolOn",DOFContactors)
        bsR.KickForceVar=2
        bsR.KickAngleVar=3

    Set bsT=New cvpmBallStack
        bsT.InitSaucer S15,15,270,8
        bsT.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("SolOn",DOFContactors)

    Cautiva.CreateBall
    Cautiva.Kick 180,1

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
Sub Table1_KeyDown(ByVal keycode)
    If keycode=LeftFlipperKey Then Controller.Switch(6)=1
	If keycode=RightFlipperkey Then Controller.Switch(7)=1
	If keycode=StartGameKey THEN CONTROLLER.SWITCH(3)=1
	If keycode=KeySlamDoorHit Then Controller.Switch(152) = 1
	If keycode = Leftmagnasave Then controller.switch(4) = 1
	If keycode = rightmagnasave Then controller.switch(5) = 1

	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If keycode=LeftFlipperKey Then Controller.Switch(6)=0
	If keycode=RightFlipperkey Then Controller.Switch(7)=0
	If keycode=StartGameKey THEN CONTROLLER.SWITCH(3)=0
	If keycode=KeySlamDoorHit Then Controller.Switch(152) = 0
	If keycode=Leftmagnasave Then controller.switch(4) = 0
	If keycode=rightmagnasave Then controller.switch(5) = 0

	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
End Sub

'**********************************************************************************************************

' Drain hole
Sub Drain_Hit:bsTrough.addball me:playsoundAtVol"drain",drain,1:End Sub
Sub S14_Hit:bsL.AddBall 0 : playsoundAtVol "popper_ball", ActiveBall, 1: End Sub
Sub S15_Hit:bsT.AddBall 0 : playsoundAtVol "popper_ball", ActiveBall, 1: End Sub
Sub S16_Hit:bsR.AddBall 0 : playsoundAtVol "popper_ball", ActiveBall, 1: End Sub

'spinners
Sub S24_Spin:vpmTimer.PulseSw 24 : playsoundAtVol"fx_spinner",S24,VoLSpin : End Sub

'StandUp Targets
Sub S17_Hit:vpmTimer.PulseSw 17:End Sub
Sub S25_Hit:vpmTimer.PulseSw 25:End Sub
Sub S26_Hit:vpmTimer.PulseSw 26:End Sub
Sub S27_Hit:vpmTimer.PulseSw 27:End Sub
Sub S31_Hit:vpmTimer.PulseSw 31:End Sub
Sub S35_Hit:vpmTimer.PulseSw 35:End Sub
Sub S36_Hit:vpmTimer.PulseSw 36:End Sub
Sub S41_Hit:vpmTimer.PulseSw 41:End Sub
Sub S42_Hit:vpmTimer.PulseSw 42:End Sub
Sub S43_Hit:vpmTimer.PulseSw 43:End Sub
Sub S45_Hit:vpmTimer.PulseSw 45:End Sub
Sub S46_Hit:vpmTimer.PulseSw 46:End Sub

'Wire Triggers
Sub s20_Hit     : Controller.Switch(20) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub s20_UnHit   : Controller.Switch(20) = 0 : End Sub
Sub s21_Hit     : Controller.Switch(21) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub s21_UnHit   : Controller.Switch(21) = 0 : End Sub
Sub s22_Hit     : Controller.Switch(22) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub s22_UnHit   : Controller.Switch(22) = 0 : End Sub
Sub s23_Hit     : Controller.Switch(23) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub s23_UnHit   : Controller.Switch(23) = 0 : End Sub
Sub s30_Hit     : Controller.Switch(30) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub s30_UnHit   : Controller.Switch(30) = 0 : End Sub
Sub s32_Hit     : Controller.Switch(32) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub s32_UnHit   : Controller.Switch(32) = 0 : End Sub
Sub s33_Hit     : Controller.Switch(33) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub s33_UnHit   : Controller.Switch(33) = 0 : End Sub
Sub s34_Hit     : Controller.Switch(34) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub s34_UnHit   : Controller.Switch(34) = 0 : End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(10) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(11) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub

'**********************************************************************************************************
' Lights
'**********************************************************************************************************

Set Lights(0)=L0
Lights(1)=Array(L1,L1a)
Set Lights(2)=L2
Set Lights(3)=L3
Set Lights(4)=L4
Set Lights(5)=L5
Set Lights(6)=L6
Set Lights(7)=L7
Set Lights(8)=L8
Set Lights(9)=L9
Set Lights(10)=L10
Set Lights(11)=L11
Set Lights(12)=L12
Set Lights(13)=L13
Set Lights(14)=L14
Set Lights(15)=L15
Set Lights(16)=L16
Set Lights(17)=L17
Set Lights(18)=L18
Set Lights(19)=L19
Set Lights(20)=L20
Set Lights(21)=L21
Set Lights(22)=L22
Set Lights(23)=L23
Set Lights(24)=L24
Set Lights(25)=L25
Set Lights(26)=L26
Set Lights(27)=L27
Set Lights(28)=L28
Set Lights(29)=L29
Set Lights(30)=L30
Set Lights(31)=L31
Set Lights(32)=L32
Set Lights(33)=L33
Set Lights(34)=L34
Set Lights(35)=L35
Set Lights(36)=L36
Set Lights(37)=L37
Set Lights(38)=L38
Set Lights(39)=L39
Set Lights(40)=L40
Set Lights(41)=L41
Set Lights(42)=L42
Lights(43)=Array(L43,L43a)
Lights(44)=Array(L44,L44a)
Lights(45)=Array(L45,L45a)
Lights(46)=Array(L46,L46a)
Set Lights(47)=L47
Set Lights(48)=L48
Set Lights(49)=L49
Set Lights(50)=L50
Set Lights(51)=L51
Lights(52)=Array(L52,L52a,L52a1)   'bumper
Lights(53)=Array(L53,L53a,L53a1)   'bumper
Set Lights(54)=L54
Set Lights(55)=L55
Set Lights(56)=L56
Set Lights(57)=L57
Set Lights(58)=L58
Set Lights(59)=L59
Set Lights(60)=L60
Set Lights(61)=L61
Set Lights(62)=L62
Set Lights(63)=L63
Set Lights(64)=L64
Set Lights(65)=L65
Set Lights(66)=L66
Set Lights(67)=L67
Set Lights(68)=L68
Set Lights(69)=L69
Set Lights(70)=L70
Set Lights(71)=L71
Set Lights(72)=L72
Set Lights(73)=L73
Set Lights(74)=L74
Set Lights(75)=L75
Set Lights(76)=L76
Set Lights(77)=L77
Set Lights(78)=L78
Set Lights(79)=L79
Set Lights(80)=L80
Set Lights(81)=L81
Set Lights(82)=L82
Set Lights(83)=L83
Set Lights(84)=L84
Set Lights(85)=L85
Set Lights(86)=L86
Set Lights(87)=L87
Set Lights(88)=L88
Set Lights(89)=L89
Set Lights(90)=L90
Set Lights(91)=L91
Set Lights(92)=L92
Set Lights(93)=L93
Set Lights(94)=L94
Set Lights(95)=L95
Set Lights(96)=L96
Set Lights(97)=L97
Set Lights(98)=L98
Set Lights(99)=L99
Set Lights(100)=L100
Set Lights(101)=L101
Set Lights(102)=L102
Set Lights(103)=L103
Set Lights(104)=L104
Set Lights(105)=L105

'**********************************************************************************************************
' Display
'**********************************************************************************************************

Dim Digits(51)
Digits(0)=Array(Q1,Q2,Q3,Q4,Q5,Q6,Q7,Q8,Q9,Q10,Q11,Q12,Q13,Q14,Q15,Q16)
Digits(1)=Array(Q17,Q18,Q19,Q20,Q21,Q22,Q23,Q24,Q25,Q26,Q27,Q28,Q29,Q30,Q31,Q32)
Digits(2)=Array(Q33,Q34,Q35,Q36,Q37,Q38,Q39,Q40,Q41,Q42,Q43,Q44,Q45,Q46,Q47,Q48)
Digits(3)=Array(Q49,Q50,Q51,Q52,Q53,Q54,Q55,Q56,Q57,Q58,Q59,Q60,Q61,Q62,Q63,Q64)
Digits(4)=Array(Q65,Q66,Q67,Q68,Q69,Q70,Q71,Q72,Q73,Q74,Q75,Q76,Q77,Q78,Q79,Q80)
Digits(5)=Array(Q81,Q82,Q83,Q84,Q85,Q86,Q87,Q88,Q89,Q90,Q91,Q92,Q93,Q94,Q95,Q96)
Digits(6)=Array(Q97,Q98,Q99,Q100,Q101,Q102,Q103,Q104,Q105,Q106,Q107,Q108,Q109,Q110,Q111,Q112)
Digits(7)=Array(Q113,Q114,Q115,Q116,Q117,Q118,Q119,Q120,Q121,Q122,Q123,Q124,Q125,Q126,Q127,Q128)
Digits(8)=Array(Q129,Q130,Q131,Q132,Q133,Q134,Q135,Q136,Q137,Q138,Q139,Q140,Q141,Q142,Q143,Q144)
Digits(9)=Array(Q145,Q146,Q147,Q148,Q149,Q150,Q151,Q152,Q153,Q154,Q155,Q156,Q157,Q158,Q159,Q160)
Digits(10)=Array(Q161,Q162,Q163,Q164,Q165,Q166,Q167,Q168,Q169,Q170,Q171,Q172,Q173,Q174,Q175,Q176)
Digits(11)=Array(Q177,Q178,Q179,Q180,Q181,Q182,Q183,Q184,Q185,Q186,Q187,Q188,Q189,Q190,Q191,Q192)
Digits(12)=Array(Q193,Q194,Q195,Q196,Q197,Q198,Q199,Q200,Q201,Q202,Q203,Q204,Q205,Q206,Q207,Q208)
Digits(13)=Array(Q209,Q210,Q211,Q212,Q213,Q214,Q215,Q216,Q217,Q218,Q219,Q220,Q221,Q222,Q223,Q224)
Digits(14)=Array(Q225,Q226,Q227,Q228,Q229,Q230,Q231,Q232,Q233,Q234,Q235,Q236,Q237,Q238,Q239,Q240)
Digits(15)=Array(Q241,Q242,Q243,Q244,Q245,Q246,Q247,Q248,Q249,Q250,Q251,Q252,Q253,Q254,Q255,Q256)
Digits(16)=Array(Q257,Q258,Q259,Q260,Q261,Q262,Q263,Q264,Q265,Q266,Q267,Q268,Q269,Q270,Q271,Q272)
Digits(17)=Array(Q273,Q274,Q275,Q276,Q277,Q278,Q279,Q280,Q281,Q282,Q283,Q284,Q285,Q286,Q287,Q288)
Digits(18)=Array(Q289,Q290,Q291,Q292,Q293,Q294,Q295,Q296,Q297,Q298,Q299,Q300,Q301,Q302,Q303,Q304)
Digits(19)=Array(Q305,Q306,Q307,Q308,Q309,Q310,Q311,Q312,Q313,Q314,Q315,Q316,Q317,Q318,Q319,Q320)
Digits(20)=Array(P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16)
Digits(21)=Array(P17,P18,P19,P20,P21,P22,P23,P24,P25,P26,P27,P28,P29,P30,P31,P32)
Digits(22)=Array(P33,P34,P35,P36,P37,P38,P39,P40,P41,P42,P43,P44,P45,P46,P47,P48)
Digits(23)=Array(P49,P50,P51,P52,P53,P54,P55,P56,P57,P58,P59,P60,P61,P62,P63,P64)
Digits(24)=Array(P65,P66,P67,P68,P69,P70,P71,P72,P73,P74,P75,P76,P77,P78,P79,P80)
Digits(25)=Array(P81,P82,P83,P84,P85,P86,P87,P88,P89,P90,P91,P92,P93,P94,P95,P96)
Digits(26)=Array(P97,P98,P99,P100,P101,P102,P103,P104,P105,P106,P107,P108,P109,P110,P111,P112)
Digits(27)=Array(P113,P114,P115,P116,P117,P118,P119,P120,P121,P122,P123,P124,P125,P126,P127,P128)
Digits(28)=Array(P129,P130,P131,P132,P133,P134,P135,P136,P137,P138,P139,P140,P141,P142,P143,P144)
Digits(29)=Array(P145,P146,P147,P148,P149,P150,P151,P152,P153,P154,P155,P156,P157,P158,P159,P160)
Digits(30)=Array(P161,P162,P163,P164,P165,P166,P167,P168,P169,P170,P171,P172,P173,P174,P175,P176)
Digits(31)=Array(P177,P178,P179,P180,P181,P182,P183,P184,P185,P186,P187,P188,P189,P190,P191,P192)
Digits(32)=Array(P193,P194,P195,P196,P197,P198,P199,P200,P201,P202,P203,P204,P205,P206,P207,P208)
Digits(33)=Array(P209,P210,P211,P212,P213,P214,P215,P216,P217,P218,P219,P220,P221,P222,P223,P224)
Digits(34)=Array(P225,P226,P227,P228,P229,P230,P231,P232,P233,P234,P235,P236,P237,P238,P239,P240)
Digits(35)=Array(P241,P242,P243,P244,P245,P246,P247,P248,P249,P250,P251,P252,P253,P254,P255,P256)
Digits(36)=Array(P257,P258,P259,P260,P261,P262,P263,P264,P265,P266,P267,P268,P269,P270,P271,P272)
Digits(37)=Array(P273,P274,P275,P276,P277,P278,P279,P280,P281,P282,P283,P284,P285,P286,P287,P288)
Digits(38)=Array(P289,P290,P291,P292,P293,P294,P295,P296,P297,P298,P299,P300,P301,P302,P303,P304)
Digits(39)=Array(P305,P306,P307,P308,P309,P310,P311,P312,P313,P314,P315,P316,P317,P318,P319,P320)
Digits(40)=Array(Z1,Z2,Z3,Z4,Z5,Z6,Z7)
Digits(41)=Array(Z8,Z9,Z10,Z11,Z12,Z13,Z14)
Digits(42)=Array(Z15,Z16,Z17,Z18,Z19,Z20,Z21)
Digits(43)=Array(Z22,Z23,Z24,Z25,Z26,Z27,Z28)
Digits(44)=Array(Z29,Z30,Z31,Z32,Z33,Z34,Z35)
Digits(45)=Array(Z36,Z37,Z38,Z39,Z40,Z41,Z42)
Digits(46)=Array(Z43,Z44,Z45,Z46,Z47,Z48,Z49)
Digits(47)=Array(Z50,Z51,Z52,Z53,Z54,Z55,Z56)
Digits(48)=Array(Z57,Z58,Z59,Z60,Z61,Z62,Z63)
Digits(49)=Array(Z64,Z65,Z66,Z67,Z68,Z69,Z70)
Digits(50)=Array(Z71,Z72,Z73,Z74,Z75,Z76,Z77)
Digits(51)=Array(Z78,Z79,Z80,Z81,Z82,Z83,Z84)

Sub DisplayTimer_Timer
    Dim ChgLED,ii,num,chg,stat,obj
    ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
        If DesktopMode = True Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
            if (num < 52) then
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

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw 13
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
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
    vpmTimer.PulseSw 12
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
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
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

Const tnob = 4 ' total number of balls
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
'   ninuzzu's   FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle

    Flipperizq.roty=leftFlipper.CurrentAngle +270
    Flipperder.roty=rightFlipper.CurrentAngle +270

End Sub

'*****************************************
'   ninuzzu's   BALL SHADOW
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
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
        PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VoLRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub Posts_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  On Error Resume Next
  Controller.Pause = False
  Controller.Stop
End Sub

