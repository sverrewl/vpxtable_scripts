'***************************************************************
'*            Front Door Game Adjustments                *
'***************************************************************
 '                High Score Feature Adjustments

'The game is designed to award Laser Cannons Panic Buttons and Base Stations at each of three score levels.
'The recommended levels are on the score card in game.
'Any level from 50,000 to 9,950,000 can be set as desired. It is also possible to reset 00 any or all of
'the levels if desired.
'
'   1.Push and releas self test button (Key 7 in VP) at one second intervals approximately six times or
'     until identification number 01 appears on the base station display.
'
'   2.The number on the player score displays is the score level. It can be increased if desired by holding
'     the credit button in. To decrease the score level hold the credit button in and depress and release The
'     self test button. Release the credit button when the desired number appears. Note that the level changes
'     50,000 points at a time. If the number 00 is left on the displays the high score feature is eliminated.
'Setting the first level automatically sets the second and third.

' Thalamus 2018-07-24
' This table doesn't have "Positional Sound Playback Functions" or "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

option explicit
Randomize

' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

Const cGameName="rapidfir"
Const UseSolenoids=1,UseLamps=True,UseGI=0,UseSyn=1,SSolenoidOn="SolOn",SSolenoidOff="Soloff"
Const SCoin="coin3",cCredits="Rapid Fire"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0
LoadVPM "01560000","Bally.vbs", 3.36

SolCallback(1)="SolShootBall"'Ball Shooter/Ball Feed Motor
SolCallback(17) = "vpmFlasher array(Flasher1,Flasher2,Flasher3,Flasher4,Flasher5,Flasher6),"

'***************************************************************
'*                   Cannon                        *
'***************************************************************

Dim Angle,Position,GunWalls
Angle=0:Position=30
GunWalls=Array(S1,S2,S3,S4,S5,S6,S7,S8,S9,S10,S11,S12,S13,S14,S15,S16,S17,S18,S19,S20,S21,S22,S23,S24,S25,S26,S27,S28,S29,S30,S31,S32,S33,S34,S35,S36,S37,S38,S39,S40,S41,S42,S43,S44,S45,S46,S47,S48,S49,S50,S51,S52,S53,S54,S55,S56,S57,S58,S59,S60)

Sub SolShootBall(Enabled)
  If Enabled Then
    Shooter.CreateSizedBall(12)
    Select Case Position
      Case 0:Shooter.Kick 331,50
      Case 1:Shooter.Kick 331,50
      Case 2:Shooter.Kick 332,50
      Case 3:Shooter.Kick 333,50
      Case 4:Shooter.Kick 334,50
      Case 5:Shooter.Kick 335,50
      Case 6:Shooter.Kick 336,50
      Case 7:Shooter.Kick 337,50
      Case 8:Shooter.Kick 338,50
      Case 9:Shooter.Kick 339,50
      Case 10:Shooter.Kick 340,50
      Case 11:Shooter.Kick 341,50
      Case 12:Shooter.Kick 342,50
      Case 13:Shooter.Kick 343,50
      Case 14:Shooter.Kick 344,50
      Case 15:Shooter.Kick 345,50
      Case 16:Shooter.Kick 346,50
      Case 17:Shooter.Kick 347,50
      Case 18:Shooter.Kick 348,50
      Case 19:Shooter.Kick 349,50
      Case 20:Shooter.Kick 350,50
      Case 21:Shooter.Kick 351,50
      Case 22:Shooter.Kick 352,50
      Case 23:Shooter.Kick 353,50
      Case 24:Shooter.Kick 354,50
      Case 25:Shooter.Kick 355,50
      Case 26:Shooter.Kick 356,50
      Case 27:Shooter.Kick 357,50
      Case 28:Shooter.Kick 358,50
      Case 29:Shooter.Kick 359,50
      Case 30:Shooter.Kick 0,50
      Case 31:Shooter.Kick 1,50
      Case 32:Shooter.Kick 2,50
      Case 33:Shooter.Kick 3,50
      Case 34:Shooter.Kick 4,50
      Case 35:Shooter.Kick 5,50
      Case 36:Shooter.Kick 6,50
      Case 37:Shooter.Kick 7,50
      Case 38:Shooter.Kick 8,50
      Case 39:Shooter.Kick 9,50
      Case 40:Shooter.Kick 10,50
      Case 41:Shooter.Kick 11,50
      Case 42:Shooter.Kick 12,50
      Case 43:Shooter.Kick 13,50
      Case 44:Shooter.Kick 14,50
      Case 45:Shooter.Kick 15,50
      Case 46:Shooter.Kick 16,50
      Case 47:Shooter.Kick 17,50
      Case 48:Shooter.Kick 18,50
      Case 49:Shooter.Kick 19,50
      Case 50:Shooter.Kick 20,50
      Case 51:Shooter.Kick 21,50
      Case 52:Shooter.Kick 22,50
      Case 53:Shooter.Kick 23,50
      Case 54:Shooter.Kick 24,50
      Case 55:Shooter.Kick 25,50
      Case 56:Shooter.Kick 26,50
      Case 57:Shooter.Kick 27,50
      Case 58:Shooter.Kick 28,50
      Case 59:Shooter.Kick 29,50
    End Select
  End If
End Sub

Dim LT,RT
LT=0:RT=0

Sub UpdatePos_Timer
  If LT=1 And RT=0 And Position>0 Then
    GunWalls(Position).IsDropped=0
    Position=Position-1
    GunWalls(Position).IsDropped=1
        Cannon.Rotz = (Position)
  End If
  If LT=0 And RT=1 And Position<59 Then
    GunWalls(Position).IsDropped=0
    Position=Position+1
    GunWalls(Position).IsDropped=1
        Cannon.Rotz = (Position)
  End If
End Sub


'***************************************************************
'*                Table Init                       *
'***************************************************************

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

GunWalls(Position).IsDropped=1
  On Error Resume Next
  With Controller
    .GameName=cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine="Rapid Fire" & vbnewline & "Rapid Fire"
    .ShowFrame=0
    .ShowTitle=0
    .ShowDMDOnly=1
    .HandleMechanics=0
    .HandleKeyboard=0
        .hidden=1
  End With
  Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0
End Sub
  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1
  vpmMapLights AllLights


'ExtraKeyHelp=KeyName(KeyUpperLeft)&vbTab&"Left Fire"&vbNewLine&KeyName(KeyUpperRight)&vbTab&"Right Fire"_
'&vbNewLine&KeyName(LeftTiltKey)&vbTab&"Panic Button"&vbNewLine&KeyName(RightTiltKey)&vbTab&"Laser Button"_
'&vbNewLine&"2"&vbTab&"Special Credit Game"&vbNewLine&"k"&vbTab&"AutoFire On/Off"

Sub Table1_KeyDown(ByVal KeyCode)
If KeyName(KeyCode)="K" Then
  If AATimer.Enabled Then
    AATimer.Enabled=0
  Else
    AATimer.Enabled=1
  End If
End If
If KeyCode=LeftMagnaSave Then
'If KeyCode=LeftTiltKey Then
  Controller.Switch(7)=1
  Exit Sub
End If
If KeyCode=RightMagnaSave Then
'If KeyCode=RightTiltKey Then
  Controller.Switch(5)=1
  Exit Sub
End If
If KeyName(KeyCode)="1" Then
  Controller.Switch(6)=1
  Exit Sub
End If
If KeyName(KeyCode)="2" Then
  Controller.Switch(8)=1
  Exit Sub
End If
  If KeyCode=LeftFlipperKey Then
    LT=1
    RT=0
    Controller.Switch(3)=1
    Controller.Switch(4)=0
        Controller.Switch(1)=1
    Exit Sub
  End If
  If KeyCode=RightFlipperKey Then
    LT=0
    RT=1
    Controller.Switch(3)=0
    Controller.Switch(4)=1
        Controller.Switch(1)=1
    Exit Sub
  End If
    If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub AATimer_Timer
       vpmTimer.PulseSw 1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
If KeyCode=LeftMagnaSave Then
'If KeyCode=LeftTiltKey Then
  Controller.Switch(7)=0
  Exit Sub
End If
If KeyCode=RightMagnaSave Then
'If KeyCode=RightTiltKey Then
  Controller.Switch(5)=0
  Exit Sub
End If
If KeyName(KeyCode)="1" Then
  Controller.Switch(6)=0
  Exit Sub
End If
If KeyName(KeyCode)="2" Then
  Controller.Switch(8)=0
  Exit Sub
End If
  If KeyCode=LeftFlipperKey Then
    LT=0
    Controller.Switch(3)=0
    Controller.Switch(4)=0
        Controller.Switch(1)=0
    Exit Sub
  End If
  If KeyCode=RightFlipperKey Then
    RT=0
    Controller.Switch(3)=0
    Controller.Switch(4)=0
        Controller.Switch(1)=0
    Exit Sub
  End If
    If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'*****************************************************
'*               Flap And Tank Switches              *
'*****************************************************

Sub sw17_Hit:Me.DestroyBall:vpmTimer.PulseSw 17:End Sub
Sub sw18_Hit:Me.DestroyBall:vpmTimer.PulseSw 18:End Sub
Sub sw19_Hit:Me.DestroyBall:vpmTimer.PulseSw 19:End Sub
Sub sw20_Hit:Me.DestroyBall:vpmTimer.PulseSw 20:End Sub
Sub sw21_Hit:Me.DestroyBall:vpmTimer.PulseSw 21:End Sub
Sub sw22_Hit:Me.DestroyBall:vpmTimer.PulseSw 22:End Sub
Sub sw23_Hit:Me.DestroyBall:vpmTimer.PulseSw 23:End Sub
Sub sw24_Hit:Me.DestroyBall:vpmTimer.PulseSw 24:End Sub
Sub sw25_Hit:Me.DestroyBall:End Sub
Sub sw26_Hit:Me.DestroyBall:End Sub
Sub sw27_Hit:Me.DestroyBall:End Sub
Sub sw28_Hit:Me.DestroyBall:End Sub
Sub sw29_Hit:Me.DestroyBall:End Sub
Sub sw30_Hit:Me.DestroyBall:End Sub
Sub sw31_Hit:Me.DestroyBall:End Sub
Sub sw32_Hit:Me.DestroyBall:End Sub
Sub sw33_Hit:Me.DestroyBall:End Sub
Sub sw34_Hit:Me.DestroyBall:End Sub
Sub sw35_Hit:Me.DestroyBall:End Sub
Sub sw36_Hit:Me.DestroyBall:End Sub
Sub sw37_Hit:Me.DestroyBall:End Sub
Sub sw38_Hit:Me.DestroyBall:End Sub
Sub sw39_Hit:Me.DestroyBall:End Sub
Sub Drain_Hit:Me.DestroyBall:End Sub


Sub sw25gate_Hit:vpmTimer.PulseSw 25:End Sub
Sub sw26gate_Hit:vpmTimer.PulseSw 26:End Sub
Sub sw27gate_Hit:vpmTimer.PulseSw 27:End Sub
Sub sw28gate_Hit:vpmTimer.PulseSw 28:End Sub
Sub sw29gate_Hit:vpmTimer.PulseSw 29:End Sub
Sub sw30gate_Hit:vpmTimer.PulseSw 30:End Sub
Sub sw31gate_Hit:vpmTimer.PulseSw 31:End Sub
Sub sw32gate_Hit:vpmTimer.PulseSw 32:End Sub
Sub sw33gate_Hit:vpmTimer.PulseSw 33:End Sub
Sub sw34gate_Hit:vpmTimer.PulseSw 34:End Sub
Sub sw35gate_Hit:vpmTimer.PulseSw 35:End Sub
Sub sw36gate_Hit:vpmTimer.PulseSw 36:End Sub
Sub sw37gate_Hit:vpmTimer.PulseSw 37:End Sub
Sub sw38gate_Hit:vpmTimer.PulseSw 38:End Sub
Sub sw39gate_Hit:vpmTimer.PulseSw 39:End Sub



'***************************************************************
'*                Flap Gate Primitives                   *
'***************************************************************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates()
  Flap39.Rotz = sw39gate.currentangle
    Flap38.Rotz = sw38gate.currentangle
    Flap37.Rotz = sw37gate.currentangle
    Flap36.Rotz = sw36gate.currentangle
    Flap35.Rotz = sw35gate.currentangle
    Flap34.Rotz = sw34gate.currentangle
    Flap33.Rotz = sw33gate.currentangle
    Flap32.Rotz = sw32gate.currentangle
    Flap31.Rotz = sw31gate.currentangle
    Flap30.Rotz = sw30gate.currentangle
    Flap29.Rotz = sw29gate.currentangle
    Flap28.Rotz = sw28gate.currentangle
    Flap27.Rotz = sw27gate.currentangle
    Flap26.Rotz = sw26gate.currentangle
    Flap25.Rotz = sw25gate.currentangle
    Cannon.Rotz = (Position)
End Sub

'***************************************************************
'*             Dip Switches By Inkochnito                  *
'***************************************************************

'Bally Rapid Fire
 'added by Inkochnito
 Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 700,400,"Rapid Fire - DIP switches"
    .AddFrame 2,0,135,"Threshold cannon awards",&H00006000,Array("1 laser cannon",0,"2 laser cannons",&H00002000,"3 laser cannons",&H00004000,"4 laser cannons",&H00006000)'dip 14&15
    .AddFrame 2,76,135,"Panic button adjustment",&H00000020,Array("1 panic button",0,"2 panic buttons",&H00000020)'dip 6
    .AddFrame 2,122,135,"Laser cannon adjustment",&H00000040,Array("2 laser cannons",0,"4 laser cannons",&H00000040)'dip 7
    .AddChk   2,175,115,Array("Credits displayed",&H04000000)'dip 27
    .AddFrame 155,0,230,"Bases per game",&HC0000000,Array("1 base",&H40000000,"2 bases",0,"3 bases",&H80000000,"4 bases",&HC0000000)'dip 31&32
    .AddFrame 155,76,230,"Forcefield laser cannon switch",&H00000080,Array("stops shot and player loses 1 laser cannon",0,"stops shot",&H00000080)'dip 8
    .AddFrame 155,122,230,"Knocking out sets of invaders will",32768,Array("not recall for next base",0,"recall each set knocked out for next base",32768)'dip 16
    .AddLabel 50,200,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
 End Sub

Set vpmShowDips=GetRef("editDips")

Sub DisplayTimer_Timer
  Dim ChgLED,ii,jj,num,chg,stat,obj,b,x
  ChgLED = Controller.ChangedLEDs(0, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      For Each obj In Digits(num)
        If chg And 1 Then obj.State = stat And 1
        chg = chg\2 : stat = stat\2
      Next
    Next
  End If
End Sub

Dim Digits(20)
Digits(0)=Array(BLight1,BLight2,BLight3,BLight4,BLight5,BLight6,BLight7,BLight8)
Digits(1)=Array(BLight10,BLight11,BLight12,BLight13,BLight14,BLight15,BLight16)
Digits(2)=Array(BLight17,BLight18,BLight19,BLight20,BLight21,BLight22,BLight23)
Digits(3)=Array(BLight24,BLight25,BLight26,BLight27,BLight28,BLight29,BLight30,BLight31)
Digits(4)=Array(BLight33,BLight34,BLight35,BLight36,BLight37,BLight38,BLight39)
Digits(5)=Array(BLight40,BLight41,BLight42,BLight43,BLight44,BLight45,BLight46)
Digits(6)=Array(BLight47,BLight48,BLight49,BLight50,BLight51,BLight52,BLight53)
Digits(7)=Array(BLight54,BLight55,BLight56,BLight57,BLight58,BLight59,BLight60,BLight61)
Digits(8)=Array(BLight63,BLight64,BLight65,BLight66,BLight67,BLight68,BLight69)
Digits(9)=Array(BLight70,BLight71,BLight72,BLight73,BLight74,BLight75,BLight76)
Digits(10)=Array(BLight77,BLight78,BLight79,BLight80,BLight81,BLight82,BLight83,Blight84)
Digits(11)=Array(BLight86,BLight87,BLight88,BLight89,BLight90,BLight91,BLight92)
Digits(12)=Array(BLight93,BLight94,BLight95,BLight96,BLight97,BLight98,BLight99)
Digits(13)=Array(BLight100,BLight101,BLight102,BLight103,BLight104,BLight105,BLight106)
Digits(14)=Array(BLight107,BLight108,BLight109,BLight110,BLight111,BLight112,BLight113)
Digits(15)=Array(BLight114,BLight115,BLight116,BLight117,BLight118,BLight119,BLight120)
Digits(16)=Array(BLight128,BLight129,BLight130,BLight131,BLight132,BLight133,BLight134)
Digits(17)=Array(BLight121,BLight122,BLight123,BLight124,BLight125,BLight126,BLight127)
Digits(18)=Array(Light10,Light11,Light12,Light13,Light14,Light15,Light16)
Digits(19)=Array(Light26,Light27,Light28,Light29,Light30,Light31,Light32)
Digits(20)=Array(Light42,Light43,Light44,Light45,Light46,Light47,Light48)

  Set LampCallback=GetRef("UpdateMultipleLamps")

 Sub UpdateMultipleLamps
 EMReel1.SetValue L61.State
 EMReel2.SetValue L125B.State
 End Sub

'***************************************************************
'*                 Desktop Mode                      *
'***************************************************************
Dim xx
Dim DesktopMode:DesktopMode = Table1.ShowDT
If DesktopMode = True Then
    For each xx in DTVisible
      xx.visible = 1
    Next
        l61.visible = 0
        l125a.visible = 0
        l61b.visible = 1
        l125b.visible = 1
        Light1.visible = 1
        Light2.visible = 1
        Light3.visible = 1
        BLight1.visible = 1
        BLight2.visible = 1
        BLight3.visible = 1
        BLight4.visible = 1
        BLight5.visible = 1
        BLight6.visible = 1
        BLight7.visible = 1
        BLight8.visible = 1
        BLight10.visible = 1
        BLight11.visible = 1
        BLight12.visible = 1
        BLight13.visible = 1
        BLight14.visible = 1
        BLight15.visible = 1
        BLight16.visible = 1
        BLight17.visible = 1
        BLight18.visible = 1
        BLight19.visible = 1
        BLight20.visible = 1
        BLight21.visible = 1
        BLight22.visible = 1
        BLight23.visible = 1
        BLight24.visible = 1
        BLight25.visible = 1
        BLight26.visible = 1
        BLight27.visible = 1
        BLight28.visible = 1
        BLight29.visible = 1
        BLight30.visible = 1
        BLight31.visible = 1
        BLight33.visible = 1
        BLight34.visible = 1
        BLight35.visible = 1
        BLight36.visible = 1
        BLight37.visible = 1
        BLight38.visible = 1
        BLight39.visible = 1
        BLight40.visible = 1
        BLight41.visible = 1
        BLight42.visible = 1
        BLight43.visible = 1
        BLight44.visible = 1
        BLight45.visible = 1
        BLight46.visible = 1
        BLight47.visible = 1
        BLight48.visible = 1
        BLight49.visible = 1
        BLight50.visible = 1
        BLight51.visible = 1
        BLight52.visible = 1
        BLight53.visible = 1
        BLight54.visible = 1
        BLight55.visible = 1
        BLight56.visible = 1
        BLight57.visible = 1
        BLight58.visible = 1
        BLight59.visible = 1
        BLight60.visible = 1
        BLight61.visible = 1
        BLight63.visible = 1
        BLight64.visible = 1
        BLight65.visible = 1
        BLight66.visible = 1
        BLight67.visible = 1
        BLight68.visible = 1
        BLight69.visible = 1
        BLight70.visible = 1
        BLight71.visible = 1
        BLight72.visible = 1
        BLight73.visible = 1
        BLight74.visible = 1
        BLight75.visible = 1
        BLight76.visible = 1
        BLight77.visible = 1
        BLight78.visible = 1
        BLight79.visible = 1
        BLight80.visible = 1
        BLight81.visible = 1
        BLight82.visible = 1
        BLight83.visible = 1
        BLight84.visible = 1
        BLight86.visible = 1
        BLight87.visible = 1
        BLight88.visible = 1
        BLight89.visible = 1
        BLight90.visible = 1
        BLight91.visible = 1
        BLight92.visible = 1
        BLight93.visible = 1
        BLight94.visible = 1
        BLight95.visible = 1
        BLight96.visible = 1
        BLight97.visible = 1
        BLight98.visible = 1
        BLight99.visible = 1
        BLight100.visible = 1
        BLight101.visible = 1
        BLight102.visible = 1
        BLight103.visible = 1
        BLight104.visible = 1
        BLight105.visible = 1
        BLight106.visible = 1
        BLight107.visible = 1
        BLight108.visible = 1
        BLight109.visible = 1
        BLight110.visible = 1
        BLight111.visible = 1
        BLight112.visible = 1
        BLight113.visible = 1
        BLight114.visible = 1
        BLight115.visible = 1
        BLight116.visible = 1
        BLight117.visible = 1
        BLight118.visible = 1
        BLight119.visible = 1
        BLight120.visible = 1
        BLight121.visible = 1
        BLight122.visible = 1
        BLight123.visible = 1
        BLight124.visible = 1
        BLight125.visible = 1
        BLight126.visible = 1
        BLight127.visible = 1
        BLight128.visible = 1
        BLight129.visible = 1
        BLight130.visible = 1
        BLight131.visible = 1
        BLight132.visible = 1
        BLight133.visible = 1
        BLight134.visible = 1
        Light10.visible = 1
        Light11.visible = 1
        Light12.visible = 1
        Light13.visible = 1
        Light14.visible = 1
        Light15.visible = 1
        Light16.visible = 1
        Light26.visible = 1
        Light27.visible = 1
        Light28.visible = 1
        Light29.visible = 1
        Light30.visible = 1
        Light31.visible = 1
        Light32.visible = 1
        Light42.visible = 1
        Light43.visible = 1
        Light44.visible = 1
        Light45.visible = 1
        Light46.visible = 1
        Light47.visible = 1
        Light48.visible = 1




End If
If DesktopMode = False Then
    For each xx in DTVisible
      xx.visible = 0
    Next

        l61.visible = 1
        l125a.visible = 1
        l61b.visible = 0
        l125b.visible = 0
    Light1.visible = 0
    Light2.visible = 0
    Light3.visible = 0
        BLight1.visible = 0
        BLight2.visible = 0
        BLight3.visible = 0
        BLight4.visible = 0
        BLight5.visible = 0
        BLight6.visible = 0
        BLight7.visible = 0
        BLight8.visible = 0
        BLight10.visible = 0
        BLight11.visible = 0
        BLight12.visible = 0
        BLight13.visible = 0
        BLight14.visible = 0
        BLight15.visible = 0
        BLight16.visible = 0
        BLight17.visible = 0
        BLight18.visible = 0
        BLight19.visible = 0
        BLight20.visible = 0
        BLight21.visible = 0
        BLight22.visible = 0
        BLight23.visible = 0
        BLight24.visible = 0
        BLight25.visible = 0
        BLight26.visible = 0
        BLight27.visible = 0
        BLight28.visible = 0
        BLight29.visible = 0
        BLight30.visible = 0
        BLight31.visible = 0
        BLight33.visible = 0
        BLight34.visible = 0
        BLight35.visible = 0
        BLight36.visible = 0
        BLight37.visible = 0
        BLight38.visible = 0
        BLight39.visible = 0
        BLight40.visible = 0
        BLight41.visible = 0
        BLight42.visible = 0
        BLight43.visible = 0
        BLight44.visible = 0
        BLight45.visible = 0
        BLight46.visible = 0
        BLight47.visible = 0
        BLight48.visible = 0
        BLight49.visible = 0
        BLight50.visible = 0
        BLight51.visible = 0
        BLight52.visible = 0
        BLight53.visible = 0
        BLight54.visible = 0
        BLight55.visible = 0
        BLight56.visible = 0
        BLight57.visible = 0
        BLight58.visible = 0
        BLight59.visible = 0
        BLight60.visible = 0
        BLight61.visible = 0
        BLight63.visible = 0
        BLight64.visible = 0
        BLight65.visible = 0
        BLight66.visible = 0
        BLight67.visible = 0
        BLight68.visible = 0
        BLight69.visible = 0
        BLight70.visible = 0
        BLight71.visible = 0
        BLight72.visible = 0
        BLight73.visible = 0
        BLight74.visible = 0
        BLight75.visible = 0
        BLight76.visible = 0
        BLight77.visible = 0
        BLight78.visible = 0
        BLight79.visible = 0
        BLight80.visible = 0
        BLight81.visible = 0
        BLight82.visible = 0
        BLight83.visible = 0
        BLight84.visible = 0
        BLight86.visible = 0
        BLight87.visible = 0
        BLight88.visible = 0
        BLight89.visible = 0
        BLight90.visible = 0
        BLight91.visible = 0
        BLight92.visible = 0
        BLight93.visible = 0
        BLight94.visible = 0
        BLight95.visible = 0
        BLight96.visible = 0
        BLight97.visible = 0
        BLight98.visible = 0
        BLight99.visible = 0
        BLight100.visible = 0
        BLight101.visible = 0
        BLight102.visible = 0
        BLight103.visible = 0
        BLight104.visible = 0
        BLight105.visible = 0
        BLight106.visible = 0
        BLight107.visible = 0
        BLight108.visible = 0
        BLight109.visible = 0
        BLight110.visible = 0
        BLight111.visible = 0
        BLight112.visible = 0
        BLight113.visible = 0
        BLight114.visible = 0
        BLight115.visible = 0
        BLight116.visible = 0
        BLight117.visible = 0
        BLight118.visible = 0
        BLight119.visible = 0
        BLight120.visible = 0
        BLight121.visible = 0
        BLight122.visible = 0
        BLight123.visible = 0
        BLight124.visible = 0
        BLight125.visible = 0
        BLight126.visible = 0
        BLight127.visible = 0
        BLight128.visible = 0
        BLight129.visible = 0
        BLight130.visible = 0
        BLight131.visible = 0
        BLight132.visible = 0
        BLight133.visible = 0
        BLight134.visible = 0
        Light10.visible = 0
        Light11.visible = 0
        Light12.visible = 0
        Light13.visible = 0
        Light14.visible = 0
        Light15.visible = 0
        Light16.visible = 0
        Light26.visible = 0
        Light27.visible = 0
        Light28.visible = 0
        Light29.visible = 0
        Light30.visible = 0
        Light31.visible = 0
        Light32.visible = 0
        Light42.visible = 0
        Light43.visible = 0
        Light44.visible = 0
        Light45.visible = 0
        Light46.visible = 0
        Light47.visible = 0
        Light48.visible = 0
End If

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

Const tnob = 13 ' total number of balls
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

