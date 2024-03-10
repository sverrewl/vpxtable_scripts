Option Explicit
On Error Resume Next

ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01500000", "Bally.VBS", 3.10

Const cGameName="startrek",cCredits="Dark Rider",UseSolenoids=2,UseLamps=1,UseGI=0,UseSync=0
Const SSolenoidOn="solon",SSolenoidOff="soloff",SFlipperOn="Flipper2",SFlipperOff="Flipper",sCoin="Coin3"



Const sKnocker=6
Const sBallRelease=7
Const sSaucer=8
Const sLeftJet=9
Const sRightJet=10
Const sBottomJet=11
Const sRSling=13
Const sLSling=12
Const sTargetReset=14
Const scoinlockout=18
Const sEnable=19

SolCallback(2)    = "vpmSolSound ""Chime1"","
SolCallback(3)    = "vpmSolSound ""Chime2"","
SolCallback(4)    = "vpmSolSound ""Chime3"","
SolCallback(5)    = "vpmSolSound ""Chime4"","

SolCallback(sBallRelease)="bsTrough.SolOut"
SolCallback(sKnocker)="vpmSolSound""Knocker"","
SolCallback(sLSling)="vpmSolSound""sling"","
SolCallback(sRSling)="vpmSolSound""sling"","
SolCallback(sLeftJet)="vpmSolSound""jet3"","
SolCallback(sRightJet)="vpmSolSound""jet3"","
SolCallback(sBottomJet)="vpmSolSound""jet3"","
SolCallback(sSaucer)="bsSaucer.SolOut"
SolCallback(sTargetReset) = "dtT.SolDropUp"
SolCallback(sEnable)="vpmNudge.SolGameOn"
SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,nothing,"
SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,nothing,"


Dim bsTrough,Bump1,bump2,bump3,bsSaucer,dtT

Sub Table1_Init()
  On Error Resume Next
  With Controller
    .GameName=cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine=cCredits
        '.Games(cGameName).Settings.Value("rol")=0
    .Games(cGameName).Settings.Value("rol")=1
    .HandleMechanics=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .Run
        .Hidden=1
        .SetDisplayPosition 0,0
  If Err Then MsgBox Err.Description
  End With
  On Error Goto 0

' Thalamus : Was missing 'vpminit me'
vpminit me
'  PinMAMETimer.Interval=PinMAMEInterval
'  PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch=7
  vpmNudge.Sensitivity=3
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)


  Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 0,8,0,0,0,0,0,0
    bsTrough.InitKick BallRelease,90,10
    bsTrough.InitExitSnd "ballrel","solon"
    bsTrough.Balls=1

  Set bsSaucer = New cvpmBallStack
  bsSaucer.InitSaucer Kicker1,32,215,5
  bsSaucer.InitExitSnd "Solon","Solon"

  Set dtT = New cvpmDropTarget
  dtT.InitDrop Array(DT1,DT2,DT3,DT4),Array(1,2,3,4)
  dtT.InitSnd " ","FlapOpen"

  vpmMapLights AllLights
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If vpmKeyUp(KeyCode) Then Exit Sub
    If KeyCode=PlungerKey Then
      PlaySound"EmptyPlunger"
      Plunger.Fire
    End If
End Sub

Sub Table1_KeyDown(ByVal KeyCode)
    If vpmKeyDown(KeyCode) Then Exit Sub
    If KeyCode=PlungerKey Then
      Plunger.Pullback
      PlaySound"PullbackPlunger"
    End If
End Sub

Sub Drain_Hit:bsTrough.AddBall Me:End Sub
Sub Kicker1_Hit : bsSaucer.AddBall 0 : End Sub
Sub DT1_Hit:PlaySound "FlapClose":vpmTimer.addtimer 300, "dtT.Hit 1'" : End Sub
Sub DT2_Hit:PlaySound "FlapClose":vpmTimer.addtimer 300, "dtT.Hit 2'" : End Sub
Sub DT3_Hit:PlaySound "FlapClose":vpmTimer.addtimer 300, "dtT.Hit 3'" : End Sub
Sub DT4_Hit:PlaySound "FlapClose":vpmTimer.addtimer 300, "dtT.Hit 4'" : End Sub

Sub Bumper1_Hit:vpmTimer.PulseSw 40:BL1.duration 1, 150, 0:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 39:BL2.duration 1, 150, 0:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 38:BL3.duration 1, 150, 0:End Sub

Sub LeftSlingshot_Slingshot:vpmTimer.PulseSw 37:End Sub
Sub RightSlingshot_Slingshot:vpmTimer.PulseSw 36:End Sub
Sub Spinner1_Spin:vpmTimer.PulseSw 23:End Sub
Sub SW23a_Slingshot:PlaySound "sling":vpmTimer.PulseSw 23:End Sub
Sub SW23b_Slingshot:PlaySound "sling":vpmTimer.PulseSw 23:End Sub
 Sub SW23c_Slingshot:PlaySound "sling":vpmTimer.PulseSw 23:End Sub
Sub SW5_Hit:Controller.Switch(5)=1:End Sub
Sub SW5_unHit:Controller.Switch(5)=0:End Sub
 Sub SW18_Hit:Controller.Switch(18)=1:End Sub
Sub SW18_unHit:Controller.Switch(18)=0:End Sub
 Sub SW19_Hit:Controller.Switch(19)=1:End Sub
Sub SW19_unHit:Controller.Switch(19)=0:End Sub
 Sub SW20_Hit:Controller.Switch(20)=1:End Sub
Sub SW20_unHit:Controller.Switch(20)=0:End Sub
 Sub SW22_Hit:Controller.Switch(22)=1:End Sub
Sub SW22_unHit:Controller.Switch(22)=0:End Sub
 Sub SW25_Hit:Controller.Switch(25)=1:End Sub
Sub SW25_unHit:Controller.Switch(25)=0:End Sub
 Sub SW26_Hit:Controller.Switch(26)=1:End Sub
Sub SW26_unHit:Controller.Switch(26)=0:End Sub
 Sub SW30_Hit:Controller.Switch(30)=1:End Sub
Sub SW30_unHit:Controller.Switch(30)=0:End Sub
 Sub SW31_Hit:Controller.Switch(31)=1:End Sub
Sub SW31_unHit:Controller.Switch(31)=0:End Sub
Sub SW34_Hit:Controller.Switch(34)=1:End Sub
Sub SW34_unHit:Controller.Switch(34)=0:End Sub
Sub SW35_Hit:Controller.Switch(35)=1:End Sub
Sub SW35_unHit:Controller.Switch(35)=0:End Sub
Sub T21_Hit:vpmTimer.PulseSw 21:End Sub
Sub T24_Hit:vpmTimer.PulseSw 24:End Sub
Sub T27_Hit:vpmTimer.PulseSw 27:End Sub
Sub T28_Hit:vpmTimer.PulseSw 28:End Sub
Sub T29_Hit:vpmTimer.PulseSw 29:End Sub
Sub Gate1_Hit : playsound "sgate" : End Sub
Sub Gate2_Hit : playsound "sgate" : End Sub
Sub Gate3_Hit : playsound "sgate" : End Sub
Sub Trigger1_Hit: PlaySound "Plungere" : End Sub

'=========================================================
'                    LED Handling
'=========================================================
'Modified version of Scapino's LED code for Fathom
'
Dim SevenDigitOutput(28)
Dim DisplayPatterns(11)
Dim DigStorage(28)

dim ledstatus : ledstatus = 2

'Binary/Hex Pattern Recognition Array
DisplayPatterns(0) = 0    '0000000 Blank
DisplayPatterns(1) = 63   '0111111 zero
DisplayPatterns(2) = 6    '0000110 one
DisplayPatterns(3) = 91   '1011011 two
DisplayPatterns(4) = 79   '1001111 three
DisplayPatterns(5) = 102  '1100110 four
DisplayPatterns(6) = 109  '1101101 five
DisplayPatterns(7) = 125  '1111101 six
DisplayPatterns(8) = 7    '0000111 seven
DisplayPatterns(9) = 127  '1111111 eight
DisplayPatterns(10)= 111  '1101111 nine

'Assign 7-digit output to reels
Set SevenDigitOutput(12)  = P3D1
Set SevenDigitOutput(13)  = P3D2
Set SevenDigitOutput(14)  = P3D3
Set SevenDigitOutput(15)  = P3D4
Set SevenDigitOutput(16)  = P3D5
Set SevenDigitOutput(17)  = P3D6


Set SevenDigitOutput(18)  = P4D1
Set SevenDigitOutput(19)  = P4D2
Set SevenDigitOutput(20)  = P4D3
Set SevenDigitOutput(21)  = P4D4
Set SevenDigitOutput(22) = P4D5
Set SevenDigitOutput(23) = P4D6


Set SevenDigitOutput(0) = P1D1
Set SevenDigitOutput(1) = P1D2
Set SevenDigitOutput(2) = P1D3
Set SevenDigitOutput(3) = P1D4
Set SevenDigitOutput(4) = P1D5
Set SevenDigitOutput(5) = P1D6


Set SevenDigitOutput(6) = P2D1
Set SevenDigitOutput(7) = P2D2
Set SevenDigitOutput(8) = P2D3
Set SevenDigitOutput(9) = P2D4
Set SevenDigitOutput(10) = P2D5
Set SevenDigitOutput(11) = P2D6


Set SevenDigitOutput(24) = CrD1
Set SevenDigitOutput(25) = CrD2
Set SevenDigitOutput(26) = BaD1
Set SevenDigitOutput(27) = BaD2


Sub DisplayTimer_Timer
  On Error Resume Next
  Dim ChgLED,ii,chg,stat,obj,TempCount,temptext,adj

  ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF) 'hex of binary (display 111111, or first 6 digits)

  If Not IsEmpty(ChgLED) Then
    For ii = 0 To UBound(ChgLED)
      chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      For TempCount = 0 to 10
        If stat = DisplayPatterns(TempCount) then
          If LedStatus = 2 Then SevenDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
          DigStorage(chgLED(ii, 0)) = TempCount
        End If
        If stat = (DisplayPatterns(TempCount) + 128) then
          If LedStatus = 2 Then SevenDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
          DigStorage(chgLED(ii, 0)) = TempCount
        End If
      Next
    Next
  End IF
End Sub

 Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 320,258,"Dark Rider - DIP switch settings"
    .AddChk 214,218,95,Array("Match feature",&H00100000)
    .AddChk 214,234,95,Array("Credits display",&H00080000)
    .AddFrame 2,5,88,"SECRET WAY lane",&H00200000,Array("start at 2K",0,"start at 4K",&H200000)
    .AddFrame 2,55,88,"RIDER value",&H10000000,Array("start at 10K",0,"start at 25K",&H10000000)
    .AddFrame 2,105,88,"RIDER Special",&H00400000,Array("each time",&H400000,"only once",0)
    .AddFrame 2,155,88,"Center target",&H00800000,Array("always lit",&H800000,"alternating",0)
    .AddFrame 2,205,88,"Outlanes",&H20000000,Array("both lit",&H20000000,"alternating",0)
    .AddFrame 105,5,88,"Balls per game",&H0F009F1F,Array("3 balls",&H01000A04,"5 balls",&H01008A04)
    .AddFrame 105,53,88,"Play mode",&H00006000,Array("replay",&H6000,"extra ball",&H4000,"novelty",0,"unknown",&H2000)
    .AddFrame 105,129,88,"Sound settings",&H80000080,Array("few",0,"more",&H80,"most",&H80000000,"full",&H80000080)
    .AddFrame 105,205,88,"Return lanes",&H40000000,Array("both lit",&H40000000,"alternating",0)
    .AddFrame 208,5,93,"High game to date",&H00000060,Array("no award",0,"1 credit",&H20,"2 credits",&H40,"3 credits",&H60)
    .AddFrame 208,83,93,"Max. credits",&H00070000,Array("5 credits",0,"10 credits",&H10000,"15 credits",&H20000,"20 credits",&H30000,"25 credits",&H40000,"30 credits",&H50000,"35 credits",&H60000,"40 credits",&H70000)
    .ViewDips
  End With
End Sub
Set vpmShowDips = GetRef("editDips")

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
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

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 5000)
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

Const tnob = 2 ' total number of balls
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
 ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If
    Next
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub
