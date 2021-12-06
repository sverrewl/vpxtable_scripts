'NFL Stern 2001 for VPX by Sliderpoint


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

Const VolRH     = 1    ' Rubber hits volume.
Const VolKick   = 1    ' Kicker volume.
Const VolFlip   = 1    ' Flipper volume.


'******************* Options *********************
' DMD/Backglass Controller Setting
Const cController = 1 '0=Use value defined in cController.txt, 1=VPinMAME,  2=UVP(don't use, update already) 3=B2S server, 4=B2S with DOF (disable VP mech sounds)
'*************************************************

'*******table option
toggleModSounds = 0 '1 extra sound clips are on, 0 are off
'*********

Dim DesktopMode: DesktopMode = Table1.ShowDT

DIM UseVPMDMD
UseVPMDMD = DesktopMode


LoadVPM "01520000","Sega.VBS",3.02

Dim cNewController
Sub LoadVPM(VPMver, VBSfile, VBSver)
Dim FileObj, ControllerFile, TextStr
 'Sub LoadVPM(VPMver,VBSfile,VBSver)
  On Error Resume Next
    If ScriptEngineMajorVersion<5 Then MsgBox"VB Script Engine 5.0 or higher required"
    ExecuteGlobal GetTextFile(VBSfile)
    If Err Then MsgBox "Unable to open "&VBSfile&". Ensure that it is in the same folder as this table. "&vbNewLine&Err.Description
    Set Controller=CreateObject("b2s.server")
    'Set Controller = CreateObject("VPinMAME.Controller")
    If Err Then MsgBox "Unable to load VPinMAME."&vbNewLine&Err.Description
    If VPMver>"" Then
      If Controller.Version<VPMver Or Err Then MsgBox"This table requires VPinMAME ver "&VPMver&" or higher."
    End If
    If VPinMAMEDriverVer<VBSver Or Err Then MsgBox"This table requires "&VBSFile&" ver "&VBSver&" or higher."
  On Error Goto 0
End Sub
' On Error Resume Next
' If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
' ExecuteGlobal GetTextFile(VBSfile)
'   If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
'
'   InitializeOptions
'
'   cNewController = 1
'   If cController = 0 then
'   Set FileObj=CreateObject("Scripting.FileSystemObject")
'     If Not FileObj.FolderExists(UserDirectory) then
'     Msgbox "Visual Pinball\User directory does not exist. Defaulting to vPinMame"
'     ElseIf Not FileObj.FileExists(UserDirectory & "cController.txt") then
'     Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
'     ControllerFile.WriteLine 1: ControllerFile.Close
'     Else
'     Set ControllerFile=FileObj.GetFile(UserDirectory & "cController.txt")
'     Set TextStr=ControllerFile.OpenAsTextStream(1,0)
'       If (TextStr.AtEndOfStream=True) then
'       Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
'       ControllerFile.WriteLine 1: ControllerFile.Close
'       Else
'       cNewController=Textstr.ReadLine: TextStr.Close
'       End If
'     End If
'     Else
'     cNewController = cController
'     End If
'
' Select Case cNewController
' Case 1
' Set Controller = CreateObject("VPinMAME.Controller")
'   If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
'   If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
'   If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
' Case 2
'   Set Controller = CreateObject("UltraVP.BackglassServ")
' Case 3,4
'   Set Controller = CreateObject("B2S.Server")
' End Select
' On Error Goto 0
'End Sub

'*************************************************************
'Toggle DOF sounds on/off based on cController value
'*************************************************************
Dim ToggleMechSounds
Function SoundFX (sound)
    If cController= 4 and ToggleMechSounds = 0 Then
        SoundFX = ""
    Else
        SoundFX = sound
    End If
End Function

Sub DOF(dofevent, dofstate)
If cController>1 Then
If dofstate = 2 Then
Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
Else
Controller.B2SSetData dofevent, dofstate
End If
End If
End Sub

Dim toggleModSounds
Function ModSound(sound)
  If toggleModSounds = 0 Then
    ModSound = ""
  Else
    ModSound = sound
  End If
End Function

If toggleModSounds = 1 Then
  PlayMusic "NFL 49ers 17.mp3"
end If

Const UseSolenoids=1,UseLamps=1,UseSync=1, SCoin="coin3"

SolCallback(1)="SolTrough"
SolCallback(2)="SolShooter"
SolCallback(3)="pVUK"
SolCallback(4)="uVUK"
SolCallback(5)="vpmSolSound ""Sling"","
SolCallback(6)="dtDrop.SolDropUp"
SolCallback(7)="dtDrop.SolHit 1,"
SolCallback(8)="StLock"
SolCallback(9)="vpmSolSound ""Jet3"","
SolCallback(10)="vpmSolSound ""Jet3"","
SolCallback(11)="vpmSolSound ""Jet3"","
SolCallback(sLLFlipper)="SolLFlipper"
SolCallback(sLRFlipper)="SolRFlipper"
SolCallback(17)="vpmSolSound ""Sling"","
SolCallback(18)="vpmSolSound ""Sling"","
SolCallback(19)="SolBallDeflector"
SolCallback(20)="SFF" ' flasher under the stadium seats
SolCallback(21)="dtDrop.SolHit 2,"
SolCallback(22)="dtDrop.SolHit 3,"
SolCallback(23)="dtDrop.SolHit 4,"
'24=coin meter optional
SolCallBack(27)="UFF" 'left flasher cap
SolCallBack(28)="Flasher28" '"vpmFlasher Flasher6," 'left ramp flasher cap
SolCallBack(29)="flasher29" ' "vpmFlasher Flasher7," 'upper flasher cap
'SolCallBack(30)="vpmFlasher Array(FlasherZ1,FlasherZ2,FlasherZ3,FlasherZ4)," 'Back Panel Flashers x4
SolCallback(31)="PopsFlash"
SolCallback(32)="SlingFlash"
SolCallback(33)="SolLeftPost" ' AUX1 not sure these are used in this game
SolCallback(34)="SolMidPost" ' AUX2 not sure these are used in this game
SolCallback(35)="SolRightPost" ' AUX3 not sure these are used in this game

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("FlipperUp"),LeftFlipper, VolFlip:LeftFlipper.RotateToEnd:Flipper1.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("FlipperDown"),LeftFlipper,VolFlip:LeftFlipper.RotateToStart:Flipper1.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("FlipperUp"),RightFlipper,VolFlip:RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("FlipperDown"),RightFlipper,VolFlip:RightFlipper.RotateToStart
    End If
End Sub

Sub SolLeftPost(Enabled)
  If Enabled Then
    LP.IsDropped=0
  Else
    LP.IsDropped=1
  End If
End Sub
Sub SolRightPost(Enabled)
  If Enabled Then
    RP.IsDropped=0
  Else
    RP.IsDropped=1
  End If
End Sub
Sub SolMidPost(Enabled)
  If Enabled Then
    MP.IsDropped=0
  Else
    MP.IsDropped=1
  End If
End Sub

' add any new GI lights to the collection "GI" to control them together.
Dim ig
Sub UpdateGITimer
  For each ig in GI
    If updateGI = 1 then
    ig.state = 1
    ElseIf UpdateGI = 2 Then
    ig.state = 2
    Else
    ig.state = 0
    End If
  Next
  GILite.enabled = 1
End Sub

Sub GILite_Timer
  UpdateGI = 1
  UpdateGITimer
  me.enabled = 0
End Sub

Sub StLock(enabled)
  if Enabled Then
  Stadiumlock.isDropped = 1
  Else
  Stadiumlock.isDropped = 0
  end If
End Sub



Sub pVUK(Enabled)
  If Enabled Then
  Kicker1.Kickz 180, 12, 5, 120
  PlaysoundAtVol SoundFX("Solenoid"), Kicker1, VolKick
  Kicker1.TimerEnabled = 1
  End If
End Sub

Sub uVUK(Enabled)
  If Enabled Then
  Kicker2.Kickz 0, 70,1, 135
  PlaysoundAtVol SoundFX("Solenoid"), Kicker2, VolKick
  Kicker2.TimerEnabled = 1
  End If
End Sub

Sub Kicker1_Timer
  Controller.Switch(45) = 0
  Kicker1.Timerenabled = 0
End Sub

Sub Kicker2_Timer
  Controller.Switch(46) = 0
  Kicker2.Timerenabled = 0
End Sub

Sub SFF(Enabled)
  If Enabled Then
  Light20a.state=1
  Light20b.state=1
  Light20c.state=1
  Light20d.state=1
  Else
  Light20a.state=0
  Light20b.state=0
  Light20c.state=0
  Light20d.state=0
  End If
End Sub

Sub UFF(Enabled)
  If Enabled Then
    Light27a.state = 1
    Light27b.state = 1
    Light27c.state = 1
  Else
    Light27a.state = 0
    Light27b.state = 0
    Light27c.state = 0
  End If
End Sub

Sub SlingFlash(Enabled)
  If Enabled Then
  Light28.state = 1
  Light36.State = 1
  Light77.State = 1
  Light78.State = 1
  Else
  Light28.state = 0
  Light36.State = 0
  Light77.State = 0
  Light78.State = 0
  End If
End Sub

Sub PopsFlash(enabled)
  If Enabled Then
    Light75.state = 1
    Light76.state = 1
    Light79.State = 1
    Light80.State = 1
  Else
    Light75.state = 0
    Light76.State = 0
    Light79.State = 0
    Light80.State = 0
  End If
End Sub

Sub Flasher29(enabled)
  If Enabled Then
    Light29c.state = 1
    Light29a.state = 1
    Light29b.state = 1
  Else
    Light29c.state = 0
    Light29a.state = 0
    Light29b.state = 0
  End If
End Sub

Sub Flasher28(enabled)
  If Enabled Then
    Light28a.state = 1
    Light28b.state = 1
  Else
    Light28a.state = 0
    Light28b.state = 0
  End If
End Sub

Sub LightTimer_Timer
  F25.visible = Light25.State
  F26.visible = Light26.State
  F27.visible = Light27.State
End Sub

Sub SolBallDeflector(Enabled)
  If Enabled Then
    Deflector.IsDropped=0
  Else
    Deflector.IsDropped=1
  End If
End Sub

Sub SolTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 15
  End If
End Sub

Sub SolShooter(Enabled)
    If Enabled Then
  PlaySoundAtVol SoundFX("SolOn"), plunger1, 1
  Plunger1.fire
  Plunger1.pullback
  end if
End Sub

Dim bsTrough, Magnet1,VLLock,dtDrop,mGoalie,Magnet2, UpdateGI

Sub Table1_Init
  Plunger1.Pullback
    LP.IsDropped=1
    MP.IsDropped=1
    RP.IsDropped=1
  Deflector.IsDropped=1
  For X=0 To 7:LBPlace(X).IsDropped=1:Next
  vpmInit Me
  Controller.GameName="nfl"
  NVOffset (29)
    Controller.Games("nfl").Settings.Value("dmd_red") = 255
    Controller.Games("nfl").Settings.Value("dmd_green") = 255
    Controller.Games("nfl").Settings.Value("dmd_blue") = 255
    Controller.Games("nfl").Settings.Value("dmd_red66") = 191
    Controller.Games("nfl").Settings.Value("dmd_green66") = 156
    Controller.Games("nfl").Settings.Value("dmd_blue66") = 75
    Controller.Games("nfl").Settings.Value("dmd_red33") = 199
    Controller.Games("nfl").Settings.Value("dmd_green33") = 0
    Controller.Games("nfl").Settings.Value("dmd_blue33") = 31
    Controller.Games("nfl").Settings.Value("dmd_red0") = 0
    Controller.Games("nfl").Settings.Value("dmd_green0") = 0
    Controller.Games("nfl").Settings.Value("dmd_blue0") = 0
  Controller.SplashInfoLine="NFL San Francisco 49ers"&vbNewLine&"Stern 2001"
  Controller.HandleKeyboard=0
  Controller.ShowTitle=0
  Controller.ShowDMDOnly=1
  Controller.ShowFrame=0
  Controller.HandleMechanics=0
  On Error Resume Next
    Controller.Run GetPlayerHwnd
    If Err Then MsgBox Err.Description
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1:vpmNudge.TiltSwitch=56:vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,14,13,12,11,0,0,0
  bsTrough.InitKick BallRelease,95,4
  bsTrough.InitEntrySnd "Solenoid","Solenoid"
  bsTrough.InitExitSnd "BallRel","Solenoid"
  bsTrough.Balls=4

    Set Magnet1=New cvpmMagnet
  with Magnet1
    .InitMagnet Trigger8,40
    .Solenoid=12
    .CreateEvents "Magnet1"
    .Grabcenter = 1
  end With

    Set Magnet2=New cvpmMagnet
  with magnet2
    .InitMagnet RampMag,40
    .Solenoid=13
    .CreateEvents "Magnet2"
    .GrabCenter = 0
  End With

  Set dtDrop=New cvpmDropTarget
  dtDrop.InitDrop Array(Drop1,Drop2,Drop3,Drop4),Array(20,19,18,17)
  dtDrop.InitSnd "flapclos","flapopen"
  dtDrop.CreateEvents "dtDrop"

  Set mGoalie=New cvpmMech
  mGoalie.MType=vpmMechOneDirSol+vpmMechReverse+vpmMechLinear
  mGoalie.Sol1=25
  mGoalie.Sol2=26
  mGoalie.Length=20
  mGoalie.Steps=8
  mGoalie.AddSw 41,0,0
  mGoalie.AddSw 47,3,4
  mGoalie.AddSw 42,7,7
  mGoalie.Callback=GetRef("UpdateGoalie")
  mGoalie.Start

  vpmMapLights AllLights

' Controller.Switch(24) = 1  'switch testing


  Dim iw
    If DesktopMode = True Then
      For each iw in BGstuff: iw.Visible = True:Next
    Else
      For each iw in BGstuff: iw.Visible = False:Next
    End If
End Sub

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode=LeftMagnaSave Then Controller.Switch(1)=1
  If KeyCode=RightMagnaSave Then Controller.Switch(8)=1
  If KeyDownHandler(KeyCode) Then Exit Sub
  If KeyCode=PlungerKey Then
    PlaySoundAtVol"Plunger",plunger1, 1
        PlaySound ModSound("football grunt 03")
      if toggleModSounds = 1 Then
      Dim x
      x = INT(11 * RND(1) )
      Select Case x
      Case 1:PlayMusic "NFL 49ers 01.mp3"
      Case 2:PlayMusic "NFL 49ers 02.mp3"
      Case 3:PlayMusic "NFL 49ers 03.mp3"
      Case 4:PlayMusic "NFL 49ers 04.mp3"
      Case 5:PlayMusic "NFL 49ers 05.mp3"
      Case 6:PlayMusic "NFL 49ers 06.mp3"
      Case 7:PlayMusic "NFL 49ers 07.mp3"
      Case 8:PlayMusic "NFL 49ers 08.mp3"
      Case 9:PlayMusic "NFL 49ers 09.mp3"
      Case 10:PlayMusic "NFL 49ers 10.mp3"
      End Select
      end if
    Plunger.Pullback
  End If

End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode=LeftMagnaSave Then Controller.Switch(1)=0
  If KeyCode=RightMagnaSave Then Controller.Switch(8)=0
  If KeyUpHandler(KeyCode) Then Exit Sub
  If KeyCode=PlungerKey Then Plunger.Fire
End Sub


Dim LBPlace,X
LBPlace=Array(L0,L1,L2,L3,L4,L5,L6,L7)

Sub UpdateGoalie(aNewPos,aSpeed,aLastPos)
If aNewPos>-1 And aNewPos<8 Then For X=0 To 7:LBPlace(X).IsDropped=1:Next
  Select Case aNewPos
    Case 0:Magnet1.X=155:Magnet1.Y=275:LBPlace(0).IsDropped=0:Goalie.ObjRotZ = 38
    Case 1:Magnet1.X=170:Magnet1.Y=285:LBPlace(1).IsDropped=0:Goalie.ObjRotZ = 26
    Case 2:Magnet1.X=195:Magnet1.Y=295:LBPlace(2).IsDropped=0:Goalie.ObjRotZ = 15
    Case 3:Magnet1.X=215:Magnet1.Y=297:LBPlace(3).IsDropped=0:Goalie.ObjRotZ = 0
    Case 4:Magnet1.X=230:Magnet1.Y=295:LBPlace(4).IsDropped=0:Goalie.ObjRotZ = -15
    Case 5:Magnet1.X=250:Magnet1.Y=285:LBPlace(5).IsDropped=0:Goalie.ObjRotZ = -26
    Case 6:Magnet1.X=270:Magnet1.Y=275:LBPlace(6).IsDropped=0:Goalie.ObjRotZ = -38
    Case 7:Magnet1.X=280:Magnet1.Y=265:LBPlace(7).IsDropped=0:Goalie.ObjRotZ = -48
  End Select
  Magnet1.Size=40
End Sub

Sub Spinner1_Spin:vpmTimer.PulseSw 9:End Sub
Sub Drain_Hit:bsTrough.AddBall Me
  If toggleModSounds = 1 Then
    Dim x
    x = INT(23 * RND(1) )
    Select Case x
    Case 1:PlayMusic "NFL 49ers 11.mp3"
    Case 2:PlayMusic "NFL 49ers 12.mp3"
    Case 3:PlayMusic "NFL 49ers 13.mp3"
    Case 4:PlayMusic "NFL 49ers 14.mp3"
    Case 5:PlayMusic "NFL 49ers 15.mp3"
    Case 6:PlayMusic "NFL 49ers 16.mp3"
    Case 7:PlayMusic "NFL 49ers 17.mp3"
    Case 8:PlayMusic "NFL 49ers 18.mp3"
    Case 9:PlayMusic "NFL 49ers 19.mp3"
    Case 10:PlayMusic "NFL 49ers 20.mp3"
    Case 11:PlayMusic "NFL 49ers 21.mp3"
    Case 12:PlayMusic "NFL 49ers 22.mp3"
    Case 13:PlayMusic "NFL 49ers 23.mp3"
    Case 14:PlayMusic "NFL 49ers 24.mp3"
    Case 15:PlayMusic "NFL 49ers 25.mp3"
    Case 16:PlayMusic "NFL 49ers 26.mp3"
    Case 17:PlayMusic "NFL 49ers 27.mp3"
    Case 18:PlayMusic "NFL 49ers 28.mp3"
    Case 19:PlayMusic "NFL 49ers 29.mp3"
    Case 20:PlayMusic "NFL 49ers 15.mp3"
    Case 21:PlayMusic "NFL 49ers 15.mp3"
    Case 22:PlayMusic "NFL 49ers 17.mp3"
    End Select
  End If
UpdateGI = 0
UpdateGITimer
End Sub



Sub Trigger2_Hit:Controller.Switch(16)=1:UpdateGI = 1:updateGITimer:End Sub
Sub Trigger2_UnHit:Controller.Switch(16)=0:UpdateGI = 2: updateGITimer:End Sub
Sub SW21_Hit:Controller.Switch(21)=1:End Sub
Sub SW21_unHit:Controller.Switch(21)=0:End Sub
Sub SW22_Hit:Controller.Switch(22)=1:End Sub
Sub SW22_unHit:Controller.Switch(22)=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1:End Sub
Sub SW23_unHit:Controller.Switch(23)=0:End Sub
Sub SW24_Hit:Controller.Switch(24)=1:End Sub
Sub SW24_unHit:Controller.Switch(24)=0:End Sub
Sub L0_Hit:vpmTimer.PulseSw 25:End Sub
Sub L1_Hit:vpmTimer.PulseSw 25:End Sub
Sub L2_Hit:vpmTimer.PulseSw 25:End Sub
Sub L3_Hit:vpmTimer.PulseSw 25:End Sub
Sub L4_Hit:vpmTimer.PulseSw 25:End Sub
Sub L5_Hit:vpmTimer.PulseSw 25:End Sub
Sub L6_Hit:vpmTimer.PulseSw 25:End Sub
Sub L7_Hit:vpmTimer.PulseSw 25:End Sub
Sub T26_Hit:vpmTimer.PulseSw 26:End Sub
Sub T27_Hit:vpmTimer.PulseSw 27:End Sub
Sub T28_Hit:vpmTimer.PulseSw 28:End Sub
Sub T29_Hit:vpmTimer.PulseSw 29:End Sub
Sub Trigger4_Hit:Controller.Switch(31)=1:End Sub
Sub Trigger4_UnHit:Controller.Switch(31)=0:End Sub
Sub Trigger1_Hit:Controller.Switch(32)=1:End Sub
Sub Trigger1_UnHit:Controller.Switch(32)=0:End Sub
Sub A_Hit:Controller.Switch(33)=1:End Sub
Sub A_UnHit:Controller.Switch(33)=0:End Sub
Sub B_Hit:Controller.Switch(34)=1:End Sub
Sub B_UnHit:Controller.Switch(34)=0:End Sub
Sub C_Hit:Controller.Switch(35)=1:End Sub
Sub C_UnHit:Controller.Switch(35)=0:End Sub
Sub Gate1_Hit:vpmTimer.PulseSw(38):End Sub
Sub Gate39_Hit:vpmTimer.PulseSw (39): End Sub
Sub Trigger8_Hit:Magnet1.AddBall ActiveBall:Controller.Switch(43)=1:End Sub 'switch 43 is goalie magnet
Sub Trigger8_UnHit:Magnet1.RemoveBall ActiveBall:Controller.Switch(43)=0:End Sub
Sub Kicker7_Hit:vpmTimer.PulseSw 44:PlaySound ModSound("NFL 49ers 00b"):End Sub 'Me.DestroyBall:vpmTimer.PulseSwitch 44,100,"HandleUnder"
Sub Kicker8_Hit:vpmTimer.PulseSw 44:PlaySound ModSound("NFL 49ers 00a"):End Sub
Sub kicker9_Hit:vpmTimer.PulseSw 44:PlaySound ModSound("NFL 49ers 00b"):End Sub
Sub Kicker10_Hit:vpmTimer.PulseSw 44:PlaySound ModSound("NFL 49ers 00a"):End Sub
Sub Kicker4_Hit:vpmTimer.PulseSw 44:PlaySound ModSound("NFL 49ers 00b"):End Sub
Sub Kicker1_Hit:Controller.Switch(45)=1:Playsound ModSound("Defence Chant"):End Sub
Sub Kicker2_Hit:Controller.Switch(46)=1:Playsound ModSound("Defence Chant"):End Sub
Sub Trigger3_Hit:Controller.Switch(48)=1:End Sub
Sub Trigger3_UnHit:Controller.Switch(48)=0:End Sub
Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySound ModSound("football grunt 01"):End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySound ModSound("football grunt 02"):End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySound ModSound("football grunt 03"):End Sub
Sub Target1_Hit:vpmTimer.PulseSw 52:End Sub
Sub LeftOutlane_Hit:Controller.Switch(57)=1:PlaySound ModSound("Whistle"):End Sub
Sub LeftOutlane_UnHit:Controller.Switch(57)=0:End Sub
Sub LeftInlane_Hit:Controller.Switch(58)=1:End Sub
Sub LeftInLane_UnHit:Controller.Switch(58)=0:End Sub
Sub LeftSlingShot_SlingShot:vpmTimer.PulseSw 59:End Sub
Sub RightOutlane_Hit:Controller.Switch(60)=1:PlaySound ModSound("Whistle"):End Sub
Sub RightOutlane_UnHit:Controller.Switch(60)=0:End Sub
Sub RightInlane_Hit:Controller.Switch(61)=1:End Sub
Sub RightInlane_UnHit:Controller.Switch(61)=0:End Sub
Sub RightSlingShot_SlingShot:vpmTimer.PulseSw 62:End Sub
sub Gate5_Hit:UpdateGI=1:UpdateGITimer:End Sub
Sub RampMag_Hit():Magnet2.addball activeball:End Sub
Sub RampMag_Unhit():Magnet2.removeball activeball:activeball.vely = activeball.vely * 1.5:End Sub


' Controller.switch(54) = 1' start button
' Controller.switch(55) = 1' slam tilt
' Controller.switch(56) = 1' plumbbob tilt

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "rubber", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

