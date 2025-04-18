'Spectrum (Bally 1981)
'https://ipdb.org/machine.cgi?2274
' _______  _______  _______  _______ _________ _______           _______
'(  ____ \(  ____ )(  ____ \(  ____ \\__   __/(  ____ )|\     /|(       )
'| (    \/| (    )|| (    \/| (    \/   ) (   | (    )|| )   ( || () () |
'| (_____ | (____)|| (__    | |         | |   | (____)|| |   | || || || |
'(_____  )|  _____)|  __)   | |         | |   |     __)| |   | || ||_|| |
'      ) || (      | (      | |         | |   | (\ (   | |   | || |   | |
'/\____) || )      | (____/\| (____/\   | |   | ) \ \__| (___) || )   ( |
'\_______)|/       (_______/(_______/   )_(   |/   \__/(_______)|/     \|
'
'VPW Masterminds
'===============
'Full Table Build & Playfield Redraw by Sixtoe
'Plastics Redraw by CyberPez
'VPX Toolkit and Part Library by Niwak
'New Star Rollover by Flux and Bord
'Bell Standoff and Kicker Arm Primitives by Bord
'New Arrow Primitive by Flux
'Script Help by Apophis & Iaakki
'Blender help by Tomate
'Lamp and Solonoid Assignments by 32Assassin (The Manual was useless!)
'VR Backglass by leojreimroc (including finding backglass lamp assignments with Hauntfreaks).
'Comforting noises by Apophis, Flux and Bord.
'
'How To Play
'===========
'Spectrum is a pinball version of the boardgame "Mastermind"
'You have to guess 4 colours in a row, hit all 3 drop targets to guess that colour.
'If you guessed right the colour will flash, if you guess it wrong it will be solid.
'If you hit a lit saucer or the center top saucer when lit they will give you a clue.
'If you guess all 4 in the correctly it will increase your superstar status.
'Full Rules: https://www.ipdb.org/rulesheets/2274/SPECTRUM.HTM
'Youtube Rule Explaination: https://www.youtube.com/watch?v=4rkgSyzDJXE

Option Explicit
Randomize
SetLocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


'**********************************************************************
' Constants and Global Variables
'**********************************************************************

Const BallSize = 50     'Ball size must be 50
Const BallMass = 1      'Ball mass must be 1
Const tnob = 3      'Total number of balls on the playfield including captive balls.
Const lob = 1     'Total number of locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Const cGameName="spectru4",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="",SSolenoidOff="", SCoin=""

LoadVPM "03060000", "BALLY.VBS", 3.26
Dim VRMode, DesktopMode: DesktopMode = Table1.ShowDT
Const ForceVR = False   'Forces VR room to be active

'**************
' Timers
'**************
Sub CorTimer_Timer()
  Cor.Update
End Sub

Dim FrameTime, InitFrameTime
InitFrameTime = 0
FrameTimer.Interval = -1
Sub FrameTimer_Timer()
  FrameTime = gametime - InitFrameTime 'Calculate FrameTime as some animuations could use this
  InitFrameTime = gametime  'Count frametime
  RollingUpdate       'Update rolling sounds
  DoDTAnim          'Drop target animations
  UpdateDropTargets
  BSUpdate
  SetupVRRoom
End Sub

'**********************************************************************
' Initiate Table
'**********************************************************************

Dim SpecBall1, SpecBall2, SpecBall3, gBOT, bsD

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Spectrum - Bally 1981" & vbNewLine & "VPW"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 0
        '.Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  vpmMapLights AllLamps

    vpmNudge.TiltSwitch = 15
    vpmNudge.Sensitivity = 4
    vpmNudge.Tiltobj = Array()

  'Balls
  Set SpecBall1 = ball1.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SpecBall2 = ball2.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SpecBall3 = ball3.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBOT = Array(SpecBall1,SpecBall2,SpecBall3)

  Cor.Update

  ball1.kick 180,1
  ball1.enabled= 0
  ball2.kick 180,1
  ball2.enabled= 0
  ball3.kick 180,1
  ball3.enabled= 0


    ' Drain
    Set bsD = New cvpmBallStack
        bsD.InitSaucer Drain, 8, 352, 17.6
        bsD.CreateEvents "bsD", Drain

  SetupVRRoom

End Sub


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.5       ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 Then DisableStaticPreRendering = True

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

  SetupVRRoom

    If eventId = 3 Then DisableStaticPreRendering = False
End Sub

'****************************
'   ZBRI: Room Brightness
'****************************

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","Plastic with an image")

Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0
  'lvl = lvl^2

  ' Lighting level
  Dim v: v=(lvl * 240 + 15)/255

  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    ModulateMaterialBaseColor RoomBrightnessMtlArray(i), i, v
  Next
End Sub

Dim SavedMtlColorArray
SaveMtlColors
Sub SaveMtlColors
  ReDim SavedMtlColorArray(UBound(RoomBrightnessMtlArray))
  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    SaveMaterialBaseColor RoomBrightnessMtlArray(i), i
  Next
End Sub

Sub SaveMaterialBaseColor(name, idx)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  SavedMtlColorArray(idx) = round(base,0)
End Sub


Sub ModulateMaterialBaseColor(name, idx, val)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  Dim red, green, blue, saved_base, new_base

  'First get the existing material properties
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

  'Get saved color
  saved_base = SavedMtlColorArray(idx)

  'Next extract the r,g,b values from the base color
  red = saved_base And &HFF
  green = (saved_base \ &H100) And &HFF
  blue = (saved_base \ &H10000) And &HFF
  'msgbox red & " " & green & " " & blue

  'Create new color scaled down by 'val', and update the material
  new_base = RGB(red*val, green*val, blue*val)
  UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, new_base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub





'**********************************
'   ZMAT: General Math Functions
'**********************************
' These get used throughout the script.

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
  dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
  dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    End If
  ElseIf dx = 0 Then
    If dy = 0 Then
      Atn2 = 0
    Else
      Atn2 = Sgn(dy) * pi / 2
    End If
  End If
End Function

Function ArcCos(x)
  If x = 1 Then
    ArcCos = 0/180*PI
  ElseIf x = -1 Then
    ArcCos = 180/180*PI
  Else
    ArcCos = Atn(-x/Sqr(-x * x + 1)) + 2 * Atn(1)
  End If
End Function

Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

Function min(a,b)
  If a > b Then
    min = b
  Else
    min = a
  End If
End Function

' Used for drop targets
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy) 'Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
  Dim AB, BC, CD, DA
  AB = (bx * py) - (by * px) - (ax * py) + (ay * px) + (ax * by) - (ay * bx)
  BC = (cx * py) - (cy * px) - (bx * py) + (by * px) + (bx * cy) - (by * cx)
  CD = (dx * py) - (dy * px) - (cx * py) + (cy * px) + (cx * dy) - (cy * dx)
  DA = (ax * py) - (ay * px) - (dx * py) + (dy * px) + (dx * ay) - (dy * ax)

  If (AB <= 0 And BC <= 0 And CD <= 0 And DA <= 0) Or (AB >= 0 And BC >= 0 And CD >= 0 And DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
  Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
  Dim rotxy
  rotxy = RotPoint(ax,ay,angle)
  rax = rotxy(0) + px
  ray = rotxy(1) + py
  rotxy = RotPoint(bx,by,angle)
  rbx = rotxy(0) + px
  rby = rotxy(1) + py
  rotxy = RotPoint(cx,cy,angle)
  rcx = rotxy(0) + px
  rcy = rotxy(1) + py
  rotxy = RotPoint(dx,dy,angle)
  rdx = rotxy(0) + px
  rdy = rotxy(1) + py

  InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
  Dim rx, ry
  rx = x * dCos(angle) - y * dSin(angle)
  ry = x * dSin(angle) + y * dCos(angle)
  RotPoint = Array(rx,ry)
End Function

'**********************************************************************
' Controls
'**********************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
    If keycode = RightFlipperKey Then Controller.Switch(1) = False : FlipperActivate RightFlipper, RFPress : VR_ButtonRight.TransX = -8
    If keycode = LeftFlipperKey Then Controller.Switch(2) = False : FlipperActivate LeftFlipper, LFPress : VR_ButtonLeft.TransX = 8
  If keycode = LeftTiltKey Then Nudge 90, 1 : SoundNudgeLeft      ' Sets the nudge angle and power
  If keycode = RightTiltKey Then Nudge 270, 1 : SoundNudgeRight   ' ^
  If keycode = CenterTiltKey Then Nudge 0, 1 : SoundNudgeCenter   ' ^
  If keycode = StartGameKey Then SoundStartButton : VR_StartButton.TransY = -4
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If keycode = RightFlipperKey Then Controller.Switch(1) = True : FlipperDeActivate RightFlipper, RFPress : VR_ButtonRight.TransX = 0
    If keycode = LeftFlipperKey Then Controller.Switch(2) = True : FlipperDeActivate LeftFlipper, LFPress : VR_ButtonLeft.TransX = 0
  If keycode = StartGameKey Then VR_StartButton.TransY = 0
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'**********************************************************************
' Solenoid Call backs
'**********************************************************************

SolCallback(1) = "kicker_topleft"    'Top Kicker Kick Left
SolCallback(2) = "kicker_topright"     'Top Kicker kick Right
SolCallback(3) = "kicker_midleft"
SolCallback(4) = "kicker_midright"
SolCallback(5) = "DrainSolOut"
SolCallback(6) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(8) = "kicker_botleft"
SolCallback(9) = "kicker_botright"
SolCallback(10) = "DT_B_Reset"
SolCallback(11) = "DT_G_Reset"
SolCallback(12) = "DT_Y_Reset"
SolCallback(13) = "DT_R_Reset"
SolCallback(18) = "SolGI"
SolCallback(19) = "vpmNudge.SolGameOn"

'GI
Sub SolGI(Enabled)
  dim xx
  For each xx in GI:xx.State = Enabled: Next
End Sub

'Drain
Sub DrainSolOut(Enabled)
  bsD.SolOut Enabled
  SoundSaucerKick 1,Drain
End Sub

Sub Drain_Hit
  SoundSaucerLock
End Sub


'**********************************************************************
' SWITCHES
'**********************************************************************

'Star Triggers
Sub sw19_Hit:Controller.Switch(19) = 1 : End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub
Sub sw19a_Hit:Controller.Switch(19) = 1 : End Sub
Sub sw19a_UnHit:Controller.Switch(19) = 0:End Sub
Sub sw19b_Hit:Controller.Switch(19) = 1 : End Sub
Sub sw19b_UnHit:Controller.Switch(19) = 0:End Sub
Sub sw20_Hit:Controller.Switch(20) = 1 : End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub
Sub sw21_Hit:Controller.Switch(21) = 1 : End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub
Sub sw21a_Hit:Controller.Switch(21) = 1 : End Sub
Sub sw21a_UnHit:Controller.Switch(21) = 0:End Sub
Sub sw21b_Hit:Controller.Switch(21) = 1 : End Sub
Sub sw21b_UnHit:Controller.Switch(21) = 0:End Sub
Sub sw27_Hit:Controller.Switch(27) = 1 : End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1 : End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1 : End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw30_Hit:Controller.Switch(30) = 1 : End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1 : End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw33a_Hit:Controller.Switch(33) = 1 : End Sub
Sub sw33a_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1 : End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub
Sub sw40a_Hit:Controller.Switch(40) = 1 : End Sub
Sub sw40a_UnHit:Controller.Switch(40) = 0:End Sub

' Rollovers
Sub sw18_Hit:Controller.Switch(18) = 1 : End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub
Sub sw22_Hit:Controller.Switch(22) = 1 : End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub
Sub sw23_Hit:Controller.Switch(23) = 1 : End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
Sub sw25a_Hit:Controller.Switch(25) = 1 : End Sub
Sub sw25a_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw25b_Hit:Controller.Switch(25) = 1 : End Sub
Sub sw25b_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw25c_Hit:Controller.Switch(25) = 1 : End Sub
Sub sw25c_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw32a_Hit:Controller.Switch(32) = 1 : End Sub
Sub sw32a_UnHit:Controller.Switch(32) = 0:End Sub
Sub sw32b_Hit:Controller.Switch(32) = 1 : End Sub
Sub sw32b_UnHit:Controller.Switch(32) = 0:End Sub
Sub sw32c_Hit:Controller.Switch(32) = 1 : End Sub
Sub sw32c_UnHit:Controller.Switch(32) = 0:End Sub

' Spinners
Sub sw26_Spin:vpmTimer.PulseSw 26:End Sub
Sub sw31_Spin:vpmTimer.PulseSw 31:End Sub

'Drop Targets
Sub sw34_Hit: DTHit 34: TargetBouncer Activeball, 1.5: End Sub
Sub sw35_Hit: DTHit 35: TargetBouncer Activeball, 1.5: End Sub
Sub sw36_Hit: DTHit 36: TargetBouncer Activeball, 1.5: End Sub

Sub sw37_Hit: DTHit 37: TargetBouncer Activeball, 1.5: End Sub
Sub sw38_Hit: DTHit 38: TargetBouncer Activeball, 1.5: End Sub
Sub sw39_Hit: DTHit 39: TargetBouncer Activeball, 1.5: End Sub

Sub sw42_Hit: DTHit 42: TargetBouncer Activeball, 1.5: End Sub
Sub sw43_Hit: DTHit 43: TargetBouncer Activeball, 1.5: End Sub
Sub sw44_Hit: DTHit 44: TargetBouncer Activeball, 1.5: End Sub

Sub sw45_Hit: DTHit 45: TargetBouncer Activeball, 1.5: End Sub
Sub sw46_Hit: DTHit 46: TargetBouncer Activeball, 1.5: End Sub
Sub sw47_Hit: DTHit 47: TargetBouncer Activeball, 1.5: End Sub

Sub DT_G_Reset(enabled)
  if enabled then
    PlaySoundAt SoundFX(DTResetSound,DOFContactors),BM_sw35
    DTRaise 34
    DTRaise 35
    DTRaise 36
  end if
End Sub

Sub DT_Y_Reset(enabled)
  if enabled then
    PlaySoundAt SoundFX(DTResetSound,DOFContactors),BM_sw38
    DTRaise 37
    DTRaise 38
    DTRaise 39
  end if
End Sub

Sub DT_B_Reset(enabled)
  if enabled then
    PlaySoundAt SoundFX(DTResetSound,DOFContactors),BM_sw44
    DTRaise 42
    DTRaise 43
    DTRaise 44
  end if
End Sub

Sub DT_R_Reset(enabled)
  if enabled then
    PlaySoundAt SoundFX(DTResetSound,DOFContactors),BM_sw46
    DTRaise 45
    DTRaise 46
    DTRaise 47
  end if
End Sub




'**********************************************************************
' KICKERS
'**********************************************************************

Dim KickerBall17, KickerBall24, KickerBall41, KickerBall48
Dim KickerBall7

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

'Top Kicker ==========================================================
Sub sw7_Hit
' msgbox activeball.z
    set KickerBall7 = activeball
    Controller.Switch(7) = 1
    SoundSaucerLock
  sw7.timerinterval=-1
  sw7.timerenabled=true
End Sub

Sub sw7_unHit
  Controller.Switch(7) = 0
  sw7.timerenabled=false
end sub

sub sw7_timer
' debug.print "vuk left: " & KickerBall7.z
  'prevents ball wiggle in saucer
  KickerBall7.x=sw7.x
  KickerBall7.y=sw7.y
  KickerBall7.z=5.5 'a bit higher so it wont touch the collidable bottom
  KickerBall7.angmomx=0
  KickerBall7.angmomy=0
  KickerBall7.angmomz=0
  KickerBall7.velx=0
  KickerBall7.vely=0
  KickerBall7.velz=0
  If Controller.Switch(7) = 0 Then me.timerenabled = false
  Dim BP
    If topright_state then
    If Controller.Switch(7) <> 0 Then
      KickBall KickerBall7, 70, 15, 5, 10
      SoundSaucerKick 1, sw7
      For Each BP in BP_kick_sw7r : BP.rotx = -20: Next
'     Controller.Switch(7) = 0 'this may not be safe here
    End If
    me.timerenabled = false
  End If

    If topleft_state then
    If Controller.Switch(7) <> 0 Then
      KickBall KickerBall7, 290, 15, 5, 10
      SoundSaucerKick 1, sw7
      For Each BP in BP_kick_sw7l : BP.rotx = -20: Next
'     Controller.Switch(7) = 0 'this may not be safe here
    End If
    me.timerenabled = false
  End If

end sub

dim topright_state : topright_state = False
Sub kicker_topright(Enable)
    If Enable then
    topright_state = True
  Else
    topright_state = False
    Dim BP : For Each BP in BP_kick_sw7r : BP.rotx = 0: Next
  End If
End Sub

dim topleft_state : topleft_state = False
Sub kicker_topleft(Enable)
    If Enable then
    topleft_state = True
  Else
    topleft_state = False
    Dim BP : For Each BP in BP_kick_sw7l : BP.rotx = 0: Next
  End If
End Sub


'Bottom Left Kicker ==========================================================
Sub sw17_Hit
' debug.print "hit: " & activeball.id
    set KickerBall17 = activeball
    Controller.Switch(17) = 1
    SoundSaucerLock
  sw17.timerinterval=-1
  sw17.timerenabled=true
End Sub

Sub sw17_unHit
' debug.print "unhit: " & activeball.id
  Controller.Switch(17) = 0
  sw17.timerenabled=false
end sub

sub sw17_timer
' debug.print "vuk left: " & KickerBall17.z
  'prevents ball wiggle in saucer
  KickerBall17.x=sw17.x
  KickerBall17.y=sw17.y
  KickerBall17.z=5.5 'a bit higher so it wont touch the collidable bottom
  KickerBall17.angmomx=0
  KickerBall17.angmomy=0
  KickerBall17.angmomz=0
  KickerBall17.velx=0
  KickerBall17.vely=0
  KickerBall17.velz=0
  If Controller.Switch(17) = 0 Then me.timerenabled = false

    If botleft_state then
    If Controller.Switch(17) <> 0 Then
      KickBall KickerBall17, 135, rndint(17,19), 15, 10
      SoundSaucerKick 1, sw17
      Dim BP : For Each BP in BP_kick_sw17 : BP.rotx = -20: Next
'     Controller.Switch(17) = 0 'this may not be safe here
    End If
    me.timerenabled = false
  End If
end sub

dim botleft_state : botleft_state = False

Sub kicker_botleft(Enable)
    If Enable then
    botleft_state = True
  Else
    botleft_state = False
    Dim BP : For Each BP in BP_kick_sw17 : BP.rotx = 0: Next
  End If
End Sub

'Bottom Right Kicker ==========================================================
Sub sw24_Hit
' msgbox activeball.z
    set KickerBall24 = activeball
    Controller.Switch(24) = 1
    SoundSaucerLock
  sw24.timerinterval=-1
  sw24.timerenabled=true
End Sub

Sub sw24_unHit
  Controller.Switch(24) = 0
  sw24.timerenabled=false
end sub

sub sw24_timer
' debug.print "vuk left: " & KickerBall24.z
  'prevents ball wiggle in saucer
  KickerBall24.x=sw24.x
  KickerBall24.y=sw24.y
  KickerBall24.z=5.5 'a bit higher so it wont touch the collidable bottom
  KickerBall24.angmomx=0
  KickerBall24.angmomy=0
  KickerBall24.angmomz=0
  KickerBall24.velx=0
  KickerBall24.vely=0
  KickerBall24.velz=0
  If Controller.Switch(24) = 0 Then me.timerenabled = false

    If botright_state then
    If Controller.Switch(24) <> 0 Then
      KickBall KickerBall24, 230, rndint(17,19), 15, 10
      SoundSaucerKick 1, sw24
      Dim BP : For Each BP in BP_kick_sw24 : BP.rotx = -20: Next
'     Controller.Switch(24) = 0 'this may not be safe here
    End If
    me.timerenabled = false
  End If

end sub

dim botright_state : botright_state = False
Sub kicker_botright(Enable)
    If Enable then
    botright_state = True
  Else
    botright_state = False
    Dim BP : For Each BP in BP_kick_sw24 : BP.rotx = 0: Next
  End If
End Sub

'Mid Left Kicker ==========================================================
Sub sw41_Hit
' msgbox activeball.z
    set KickerBall41 = activeball
    Controller.Switch(41) = 1
    SoundSaucerLock
  sw41.timerinterval=-1
  sw41.timerenabled=true
End Sub

Sub sw41_unHit
  Controller.Switch(41) = 0
  sw41.timerenabled=false
end sub

sub sw41_timer
' debug.print "vuk left: " & KickerBall41.z
  'prevents ball wiggle in saucer
  KickerBall41.x=sw41.x
  KickerBall41.y=sw41.y
  KickerBall41.z=5.5 'a bit higher so it wont touch the collidable bottom
  KickerBall41.angmomx=0
  KickerBall41.angmomy=0
  KickerBall41.angmomz=0
  KickerBall41.velx=0
  KickerBall41.vely=0
  KickerBall41.velz=0
  If Controller.Switch(41) = 0 Then me.timerenabled = false

    If midleft_state then
    If Controller.Switch(41) <> 0 Then
      KickBall KickerBall41, 250, 10, 5, 10
      SoundSaucerKick 1, sw41
      Dim BP : For Each BP in BP_kick_sw41 : BP.rotx = -20: Next
'     Controller.Switch(41) = 0 'this may not be safe here
    End If
    me.timerenabled = false
  End If

end sub

dim midleft_state : midleft_state = False
Sub kicker_midleft(Enable)
    If Enable then
    midleft_state = True
  Else
    midleft_state = False
    Dim BP : For Each BP in BP_kick_sw41 : BP.rotx = 0: Next
  End If
End Sub

'Mid Right Kicker ==========================================================
Sub sw48_Hit
' msgbox activeball.z
    set KickerBall48 = activeball
    Controller.Switch(48) = 1
    SoundSaucerLock
  sw48.timerinterval=-1
  sw48.timerenabled=true
End Sub

Sub sw48_unHit
  Controller.Switch(48) = 0
  sw48.timerenabled=false
end sub

sub sw48_timer
' debug.print "vuk left: " & KickerBall48.z
  'prevents ball wiggle in saucer
  KickerBall48.x=sw48.x
  KickerBall48.y=sw48.y
  KickerBall48.z=5.5 'a bit higher so it wont touch the collidable bottom
  KickerBall48.angmomx=0
  KickerBall48.angmomy=0
  KickerBall48.angmomz=0
  KickerBall48.velx=0
  KickerBall48.vely=0
  KickerBall48.velz=0
  If Controller.Switch(48) = 0 Then me.timerenabled = false

    If midright_state then
    If Controller.Switch(48) <> 0 Then
      KickBall KickerBall48, 110, 10, 5, 10
      SoundSaucerKick 1, sw48
'     BP_kick_sw48.rotx = -20
      Dim BP : For Each BP in BP_kick_sw48 : BP.rotx = -20: Next
'     Controller.Switch(48) = 0 'this may not be safe here
    End If
    me.timerenabled = false
  End If
end sub

dim midright_state : midright_state = False
Sub kicker_midright(Enable)
    If Enable then
    midright_state = True
  Else
    midright_state = False
    Dim BP : For Each BP in BP_kick_sw48 : BP.rotx = 0: Next
  End If
End Sub

'******************************************************
' ZFLP: FLIPPERS
'******************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
    If Enabled Then
    LF.Fire
        If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
            RandomSoundReflipUpLeft LeftFlipper
        Else
            SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
        End If
    Else
        LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
    RF.Fire
        If rightflipper.currentangle < rightflipper.endangle + ReflipAngle Then
            RandomSoundReflipUpRight RightFlipper
        Else
            SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
        End If
    Else
    RightFlipper.RotateToStart
        If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
    End If
End Sub

'Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub



'******************************************************
' ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 To tnob
    rolling(i) = False
  Next
End Sub

Sub RollingUpdate()
  Dim b
  '   Dim BOT
  '   BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 To tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(gBOT)
    If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 And Not InRect(gBOT(b).x,gBOT(b).y,579,2015,845,1870,845,1959,590,2090) Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ <  - 1 And gBOT(b).z < 55 And gBOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz >  - 7 Then
          RandomSoundBallBouncePlayfieldSoft gBOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard gBOT(b)
        End If
      End If
    End If

    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If

  Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************

'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'     * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'     * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'     * Create a Timer called RampRoll, that is enabled, with a interval of 100
'     * Set RampBalls and RampType variable to Total Number of Balls
' Usage:
'     * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'     * To stop tracking ball
'        * call WireRampOff
'        * Otherwise, the ball will auto remove if it's below 30 vp units
'

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
' Setup:  Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RampBalls(6,2)
Dim RampBalls(8,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(8)

Sub WireRampOn(input)
  Waddball ActiveBall, input
  RampRollUpdate
End Sub

Sub WireRampOff()
  WRemoveBall ActiveBall.ID
End Sub

' WaddBall (Active Ball, Boolean)
Sub Waddball(input, RampInput) 'This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  Dim x
  For x = 1 To UBound(RampBalls)  'Check, don't add balls twice
    If RampBalls(x, 1) = input.id Then
      If Not IsEmpty(RampBalls(x,1) ) Then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 To UBound(RampBalls)
    If IsEmpty(RampBalls(x, 1)) Then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      Exit Sub
    End If
    If x = UBound(RampBalls) Then  'debug
      Debug.print "WireRampOn error, ball queue is full: " & vbNewLine & _
      RampBalls(0, 0) & vbNewLine & _
      TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbNewLine & _
      TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbNewLine & _
      TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbNewLine & _
      TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbNewLine & _
      TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbNewLine & _
      " "
    End If
  Next
End Sub

' WRemoveBall (BallId)
Sub WRemoveBall(ID) 'This subroutine is called from the RampRollUpdate subroutine and is used to remove and stop the ball rolling sounds
  '   Debug.Print "In WRemoveBall() + Remove ball from loop array"
  Dim ballcount
  ballcount = 0
  Dim x
  For x = 1 To UBound(RampBalls)
    If ID = RampBalls(x, 1) Then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    If Not IsEmpty(Rampballs(x,1)) Then ballcount = ballcount + 1
  Next
  If BallCount = 0 Then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer()
  RampRollUpdate
End Sub

Sub RampRollUpdate() ' Timer update
  Dim x
  For x = 1 To UBound(RampBalls)
    If Not IsEmpty(RampBalls(x,1) ) Then
      If BallVel(RampBalls(x,0) ) > 1 Then ' if ball is moving, play rolling sound
        If RampType(x) Then
          PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      End If
      If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
  Next
  If Not RampBalls(0,0) Then RampRoll.enabled = 0
End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  Me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbNewLine & _
  "1 " & TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbNewLine & _
  "2 " & TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbNewLine & _
  "3 " & TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbNewLine & _
  "4 " & TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbNewLine & _
  "5 " & TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbNewLine & _
  "6 " & TypeName(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbNewLine & _
  " "
End Sub

Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
  BallPitch = pSlope(BallVel(ball), 1, - 1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, - 4000, 60, 7000)
End Function

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************


'******************************************************
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'  Metals (all metal objects, metal walls, metal posts, metal wire guides)
'  Apron (the apron walls and plunger wall)
'  Walls (all wood or plastic walls)
'  Rollovers (wire rollover triggers, star triggers, or button triggers)
'  Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'  Gates (plate gates)
'  GatesWire (wire gates)
'  Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
' - Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
' - Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
' - Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
' - Ball rolling sounds sub
'
' Tutorial videos by Apophis
' Audio : Adding Fleep Part 1         https://youtu.be/rG35JVHxtx4?si=zdN9W4cZWEyXbOz_
' Audio : Adding Fleep Part 2         https://youtu.be/dk110pWMxGo?si=2iGMImXXZ0SFKVCh
' Audio : Adding Fleep Part 3         https://youtu.be/ESXWGJZY_EI?si=6D20E2nUM-xAw7xy


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1        'volume level; range [0, 1]
NudgeRightSoundLevel = 1        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1        'volume level; range [0, 1]
StartButtonSoundLevel = 0.1      'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1        'volume level; range [0, 1]
RollingSoundFactor = 1.1 / 5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010    'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635    'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0            'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45          'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel    'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel    'sound helper; not configurable
SlingshotSoundLevel = 0.95            'volume level; range [0, 1]
BumperSoundFactor = 4.25            'volume multiplier; must not be zero
KnockerSoundLevel = 1              'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2      'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5      'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075 / 5      'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025      'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025      'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8    'volume level; range [0, 1]
WallImpactSoundFactor = 0.075          'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075 / 3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5      'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10  'volume multiplier; must not be zero
DTSoundLevel = 0.25        'volume multiplier; must not be zero
RolloverSoundLevel = 0.25      'volume level; range [0, 1]
SpinnerSoundLevel = 0.5      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8          'volume level; range [0, 1]
BallReleaseSoundLevel = 1        'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015  'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025 / 5      'volume multiplier; must not be zero

'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / tableheight - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

' Thalamus, AudioFade - Patched
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 5) 'was 10
	Else
		AudioFade = CSng( - (( - tmp) ^ 5) ) 'was 10
	End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / tablewidth - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

' Thalamus, AudioPan - Patched
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 5) 'was 10
	Else
		AudioPan = CSng( - (( - tmp) ^ 5) ) 'was 10
	End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = CSng(BallVel(ball) ^ 2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = CSng((ball.velz) ^ 2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
End Function

Function RndInt(min, max) ' Sets a random number integer between min and max
  RndInt = Int(Rnd() * (max - min + 1) + min)
End Function

Function RndNum(min, max) ' Sets a random number between min and max
  RndNum = Rnd() * (max - min) + min
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////

Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * VolumeDial, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub

Sub SoundPlungerPull()
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////

Sub KnockerSolenoid()
  PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 8) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////

Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////

Sub RandomSoundFlipperUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm / 10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm / 10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd * 4) + 1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////

Sub Rubbers_Hit(idx)
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 5 Then
    RandomSoundRubberStrong 1
  End If
  If finalspeed <= 5 Then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////

Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd * 10) + 1
    Case 1
      PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 2
      PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 3
      PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 4
      PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 5
      PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 6
      PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 7
      PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 8
      PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 9
      PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 10
      PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd * 9) + 1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
  RandomSoundWall()
End Sub


Sub RandomSoundWall()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 5) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 4) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 3) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

Sub RandomSoundMetal()
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
  RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
  RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////

Sub RandomSoundBottomArchBallGuide()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Bounce_" & Int(Rnd * 2) + 1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd * 3) + 1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(ActiveBall.id) < 4) And cor.ballvely(ActiveBall.id) > 7 Then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd * 3) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd * 7) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

Sub RandomSoundTargetHitStrong()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 10 Then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft ActiveBall
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////

Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd * 9) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd * 7) + 1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////

Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd * 5) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, ActiveBall
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub Arch1_hit()
  If ActiveBall.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If ActiveBall.velx <  - 8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If ActiveBall.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If ActiveBall.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0
      PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1
      PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////

Sub OnBallBallCollision(ball1, ball2, velocity)

  FlipperCradleCollision ball1, ball2, velocity

  Dim snd
  Select Case Int(Rnd * 7) + 1
    Case 1
      snd = "Ball_Collide_1"
    Case 2
      snd = "Ball_Collide_2"
    Case 3
      snd = "Ball_Collide_3"
    Case 4
      snd = "Ball_Collide_4"
    Case 5
      snd = "Ball_Collide_5"
    Case 6
      snd = "Ball_Collide_6"
    Case 7
      snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05    'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025 * RelayGISoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025 * RelayGISoundLevel, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025 * RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025 * RelayFlashSoundLevel, obj
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////


'******************************************************
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************







'******************************************************
' ZPHY:  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Adding nFozzy roth physics : pt1 rubber dampeners         https://youtu.be/AXX3aen06FM?si=Xqd-rcaqTlgEd_wx
' Adding nFozzy roth physics : pt2 flipper physics          https://youtu.be/VSBFuK2RCPE?si=i8ne8Ao2co8rt7fy
' Adding nFozzy roth physics : pt3 other elements           https://youtu.be/JN8HEJapCvs?si=hvgMOk-ej1BEYjJv
'
' Note: BallMass must be set to 1. BallSize should be set to 50 (in other words the ball radius is 25)
'
' Recommended Table Physics Settings
' | Gravity Constant             | 0.97      |
' | Playfield Friction           | 0.15-0.25 |
' | Playfield Elasticity         | 0.25      |
' | Playfield Elasticity Falloff | 0         |
' | Playfield Scatter            | 0         |
' | Default Element Scatter      | 2         |
'

'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. and, special scripting
'
' TriggerLF and RF should now be 27 vp units from the flippers. In addition, 3 degrees should be added to the end angle
' when creating these triggers.
'
' RF.ReProcessBalls Activeball and LF.ReProcessBalls Activeball must be added the flipper_collide subs.
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.4            | 0.4                   | 0.375                  | 0.375              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity
Dim URF : Set URF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger


        x.AddPt "Polarity", 0, 0, 0
        x.AddPt "Polarity", 1, 0.05, - 2.7
        x.AddPt "Polarity", 2, 0.16, - 2.7
        x.AddPt "Polarity", 3, 0.22, - 0
        x.AddPt "Polarity", 4, 0.25, - 0
        x.AddPt "Polarity", 5, 0.3, - 1
        x.AddPt "Polarity", 6, 0.4, - 2
        x.AddPt "Polarity", 7, 0.5, - 2.7
        x.AddPt "Polarity", 8, 0.65, - 1.8
        x.AddPt "Polarity", 9, 0.75, - 0.5
        x.AddPt "Polarity", 10, 0.81, - 0.5
        x.AddPt "Polarity", 11, 0.88, 0
        x.AddPt "Polarity", 12, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.15, 0.85
    x.AddPt "Velocity", 2, 0.2, 0.9
    x.AddPt "Velocity", 3, 0.23, 0.95
    x.AddPt "Velocity", 4, 0.41, 0.95
    x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
    x.AddPt "Velocity", 6, 0.62, 1.0
    x.AddPt "Velocity", 7, 0.702, 0.968
    x.AddPt "Velocity", 8, 0.95,  0.968
    x.AddPt "Velocity", 9, 1.03,  0.945
    x.AddPt "Velocity", 10, 1.5,  0.945

  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
    LF.SetObjects "LF", LeftFlipper, TriggerLF
    RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub

'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before\
' modified 2024 by rothbauerw
' Added Reprocessballs for flipper collisions (LF.Reprocessballs Activeball and RF.Reprocessballs Activeball must be added to the flipper collide subs
' Improved handling to remove correction for backhand shots when the flipper is raised

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef, FlipStartAngle
  Private Balls(20), balldata(20)
  Private Name

  Dim PolarityIn, PolarityOut
  Dim VelocityIn, VelocityOut
  Dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    ReDim PolarityIn(0)
    ReDim PolarityOut(0)
    ReDim VelocityIn(0)
    ReDim VelocityOut(0)
    ReDim YcoefIn(0)
    ReDim YcoefOut(0)
    Enabled = True
    TimeDelay = 50
    LR = 1
    Dim x
    For x = 0 To UBound(balls)
      balls(x) = Empty
      Set Balldata(x) = new SpoofBall
    Next
  End Sub

  Public Sub SetObjects(aName, aFlipper, aTrigger)

    If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: .SetObjects error: first argument must be a String (And name of Object). Found:" & TypeName(aName) End If
    If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: .SetObjects error: Second argument must be a flipper. Found:" & TypeName(aFlipper) End If
    If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & TypeName(aTrigger) End If
    If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
    Name = aName
    Set Flipper = aFlipper
    FlipperStart = aFlipper.x
    FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
    FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y

    Dim str
    str = "Sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
    ExecuteGlobal(str)
    str = "Sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
    ExecuteGlobal(str)

  End Sub

  ' Legacy: just no op
  Public Property Let EndPoint(aInput)

  End Property

  Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      Case "Polarity"
        ShuffleArrays PolarityIn, PolarityOut, 1
        PolarityIn(aIDX) = aX
        PolarityOut(aIDX) = aY
        ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity"
        ShuffleArrays VelocityIn, VelocityOut, 1
        VelocityIn(aIDX) = aX
        VelocityOut(aIDX) = aY
        ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef"
        ShuffleArrays YcoefIn, YcoefOut, 1
        YcoefIn(aIDX) = aX
        YcoefOut(aIDX) = aY
        ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
  End Sub

  Public Sub AddBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If IsEmpty(balls(x)) Then
        Set balls(x) = aBall
        Exit Sub
      End If
    Next
  End Sub

  Private Sub RemoveBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If TypeName(balls(x) ) = "IBall" Then
        If aBall.ID = Balls(x).ID Then
          balls(x) = Empty
          Balldata(x).Reset
        End If
      End If
    Next
  End Sub

  Public Sub Fire()
    Flipper.RotateToEnd
    processballs
  End Sub

  Public Property Get Pos 'returns % position a ball. For debug stuff.
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x)) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x)) Then
        balldata(x).Data = balls(x)
      End If
    Next
    FlipStartAngle = Flipper.currentangle
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub

  Public Sub ReProcessBalls(aBall) 'save data of balls in flipper range
    If FlipperOn() Then
      Dim x
      For x = 0 To UBound(balls)
        If Not IsEmpty(balls(x)) Then
          if balls(x).ID = aBall.ID Then
            If isempty(balldata(x).ID) Then
              balldata(x).Data = balls(x)
            End If
          End If
        End If
      Next
    End If
  End Sub

  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef, BalltoFlip, BalltoBase, NoCorrection, checkHit
      Ycoef = 1

      'y safety Exit
      If aBall.VelY > -8 Then 'ball going down
        RemoveBall aBall
        Exit Sub
      End If

      'Find balldata. BallPos = % on Flipper
      For x = 0 To UBound(Balls)
        If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          BalltoFlip = DistanceFromFlipperAngle(BallData(x).x, BallData(x).y, Flipper, FlipStartAngle)
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
        NoCorrection = 1
      Else
        checkHit = 50 + (20 * BallPos)

        If BalltoFlip > checkHit or (PartialFlipCoef < 0.5 and BallPos > 0.22) Then
          NoCorrection = 1
        Else
          NoCorrection = 0
        End If
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        'If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled and NoCorrection = 0 Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef*VelCoef)
      End If
      If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  Dim x, aCount
  aCount = 0
  ReDim a(UBound(aArray) )
  For x = 0 To UBound(aArray)   'Shuffle objects in a temp array
    If Not IsEmpty(aArray(x) ) Then
      If IsObject(aArray(x)) Then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  If offset < 0 Then offset = 0
  ReDim aArray(aCount-1+offset)   'Resize original array
  For x = 0 To aCount-1       'set objects back into original array
    If IsObject(a(x)) Then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
  BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)    'Set up line via two points, no clamping. Input X, output Y
  Dim x, y, b, m
  x = input
  m = (Y2 - Y1) / (X2 - X1)
  b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x
      y = .y
      z = .z
      velx = .velx
      vely = .vely
      velz = .velz
      id = .ID
      mass = .mass
      radius = .radius
    End With
  End Property
  Public Sub Reset()
    x = Empty
    y = Empty
    z = Empty
    velx = Empty
    vely = Empty
    velz = Empty
    id = Empty
    mass = Empty
    radius = Empty
  End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  Dim y 'Y output
  Dim L 'Line
  'find active line
  Dim ii
  For ii = 1 To UBound(xKeyFrame)
    If xInput <= xKeyFrame(ii) Then
      L = ii
      Exit For
    End If
  Next
  If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)    'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )     'Clamp lower
  If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )    'Clamp upper

  LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************
' To add the flipper tricks you must
'  - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'  - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines
'  - Include FlipperActivate and FlipperDeactivate in the Flipper solenoid subs

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
  '   Dim BOT
  '   BOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
          gBOT(b).velx = gBOT(b).velx / 1.3
          gBOT(b).vely = gBOT(b).vely - 0.5
        End If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
  End If
End Sub


Dim FCCDamping: FCCDamping = 0.4

Sub FlipperCradleCollision(ball1, ball2, velocity)
  if velocity < 0.7 then exit sub   'filter out gentle collisions
    Dim DoDamping, coef
    DoDamping = false
    'Check left flipper
    If LeftFlipper.currentangle = LFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, LeftFlipper) OR FlipperTrigger(ball2.x, ball2.y, LeftFlipper) Then DoDamping = true
    End If
    'Check right flipper
    If RightFlipper.currentangle = RFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, RightFlipper) OR FlipperTrigger(ball2.x, ball2.y, RightFlipper) Then DoDamping = true
    End If
    If DoDamping Then
    coef = FCCDamping
        ball1.velx = ball1.velx * coef: ball1.vely = ball1.vely * coef: ball1.velz = ball1.velz * coef
        ball2.velx = ball2.velx * coef: ball2.vely = ball2.vely * coef: ball2.velz = ball2.velz * coef
    End If
End Sub



'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
  DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function DistanceFromFlipperAngle(ballx, bally, Flipper, Angle)
  DistanceFromFlipperAngle = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Angle + 90)) + Flipper.x, Sin(Radians(angle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function

'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, ULFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
'Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0
    SOSRampup = 2.5
  Case 1
    SOSRampup = 6
  Case 2
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
Const EOSReturn = 0.035  'mid 80's to early 90's
'   Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
  FlipperPress = 0
  Flipper.eostorqueangle = EOSA
  Flipper.eostorque = EOST * EOSReturn / FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b', BOT
    '   BOT = GetBalls

    For b = 0 To UBound(gBOT)
      If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3 * Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
    If FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If
  End If
End Sub

Const LiveDistanceMin = 5  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)
Const BaseDampen = 0.55

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
    Dim Dir, LiveDist
    Dir = Flipper.startangle / Abs(Flipper.startangle)    '-1 for Right Flipper
    Dim LiveCatchBounce   'If live catch is not perfect, it won't freeze ball totally
    Dim CatchTime
    CatchTime = GameTime - FCount
    LiveDist = Abs(Flipper.x - ball.x)

    If CatchTime <= LiveCatch And parm > 3 And LiveDist > LiveDistanceMin And LiveDist < LiveDistanceMax Then
        If CatchTime <= LiveCatch * 0.5 Then   'Perfect catch only when catch time happens in the beginning of the window
            LiveCatchBounce = 0
        Else
            LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)  'Partial catch when catch happens a bit late
        End If

        If LiveCatchBounce = 0 And ball.velx * Dir > 0 And LiveDist > 30 Then ball.velx = 0

        If ball.velx * Dir > 0 And LiveDist < 30 Then
            ball.velx = BaseDampen * ball.velx
            ball.vely = BaseDampen * ball.vely
            ball.angmomx = BaseDampen * ball.angmomx
            ball.angmomy = BaseDampen * ball.angmomy
            ball.angmomz = BaseDampen * ball.angmomz
        Elseif LiveDist > 30 Then
            ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
            ball.angmomx = 0
            ball.angmomy = 0
            ball.angmomz = 0
        End If
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
    End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************





'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

' Put all the Post and Pin objects in dPosts collection. Make sure dPosts fires hit events.
Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

' This collection contains the bottom sling posts. They are not in the dPosts collection so that the TargetBouncer is not applied to them, but they should still have dampening applied
' If you experience airballs with posts or targets, consider adding them to this collection
Sub NoTargetBouncer_Hit
    RubbersD.dampen ActiveBall
End Sub

' Put all the Sleeve objects in dSleeves collection. Make sure dSleeves fires hit events.
Sub dSleeves_Hit(idx)
  SleevesD.Dampen ActiveBall
  TargetBouncer ActiveBall, 0.7
End Sub

Dim RubbersD        'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False    'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1    'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64    'there's clamping so interpolate up to 56 at least

Dim SleevesD  'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False    'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn   'tbpOut.text
  Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
  End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
  End Sub

  Public Sub Dampen(aBall)
    If threshold Then
      If BallSpeed(aBall) < threshold Then Exit Sub
    End If
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
    "actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
    If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)

    aBall.velx = aBall.velx * coef
    aBall.vely = aBall.vely * coef
    aBall.velz = aBall.velz * coef
    If debugOn Then TBPout.text = str
  End Sub

  Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
      aBall.velx = aBall.velx * coef
      aBall.vely = aBall.vely * coef
      aBall.velz = aBall.velz * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    Dim x
    For x = 0 To UBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
    Next
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub
End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
  Public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize
    ReDim ballvel(0)
    ReDim ballvelx(0)
    ReDim ballvely(0)
  End Sub

  Public Sub Update() 'tracks in-ball-velocity
    Dim str, b, AllBalls, highestID
    allBalls = GetBalls

    For Each b In allballs
      If b.id >= HighestID Then highestID = b.id
    Next

    If UBound(ballvel) < highestID Then ReDim ballvel(highestID)  'set bounds
    If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)  'set bounds
    If UBound(ballvely) < highestID Then ReDim ballvely(highestID)  'set bounds

    For Each b In allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class



'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************


'******************************************************
'   ZRDT:  DROP TARGETS by Rothbauerw
'******************************************************
' The Stand Up and Drop Target solutions improve the physics for targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full target animation, switch handling and deflection on hit. For drop targets there is also a slight lift when
' the drop targets raise, bricking, and popping the ball up if it's over the drop target when it raises.
'
' Add a Timers named DTAnim and STAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
' DTAnim.interval = 10
' DTAnim.enabled = True

' Sub DTAnim_Timer
'   DoDTAnim
' DoSTAnim
' End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.
'
' For each stand up target we'll use a vp target, a laid back collidable primitive, and one primitive for visuals and animation.
' The visual primitive should should have it's pivot point centered on the x and y axis and the z should be at or just below the playfield.
' The target should animate backwards using transy.
'
' To create visual target primitives that work with the stand up and drop target code, follow the below instructions:
' (Other methods will work as well, but this is easy for even non-blender users to do)
' 1) Open a new blank table. Delete everything off the table in editor.
' 2) Copy and paste the VP target from your table into this blank table.
' 3) Place the target at x = 0, y = 0  (upper left hand corner) with an orientation of 0 (target facing the front of the table)
' 4) Under the file menu, select Export "OBJ Mesh"
' 5) Go to "https://threejs.org/editor/". Here you can modify the exported obj file. When you export, it exports your target and also
'    the playfield mesh. You need to delete the playfield mesh here. Under the file menu, chose import, and select the obj you exported
'    from VPX. In the right hand panel, find the Playfield object and click on it and delete. Then use the file menu to Export OBJ.
' 6) In VPX, you can add a primitive and use "Import Mesh" to import the exported obj from the previous step. X,Y,Z scale should be 1.
'    The primitive will use the same target texture as the VP target object.
'
' * Note, each target must have a unique switch number. If they share a same number, add 100 to additional target with that number.
' For example, three targets with switch 32 would use 32, 132, 232 for their switch numbers.
' The 100 and 200 will be removed when setting the switch value for the target.

'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

Class DropTarget
  Private m_primary, m_secondary, m_prim, m_sw, m_animate, m_isDropped

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Secondary(): Set Secondary = m_secondary: End Property
  Public Property Let Secondary(input): Set m_secondary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
  Public Property Let IsDropped(input): m_isDropped = input: End Property

  Public default Function init(primary, secondary, prim, sw, animate, isDropped)
    Set m_primary = primary
    Set m_secondary = secondary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate
    m_isDropped = isDropped

    Set Init = Me
  End Function
End Class

'Define a variable for each drop target
Dim DT34, DT35, DT36, DT37, DT38, DT39, DT42, DT43, DT44, DT45, DT46, DT47

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:      primary target wall to determine drop
' secondary:      wall used to simulate the ball striking a bent or offset target after the initial Hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             rotz must be used for orientation
'             rotx to bend the target back
'             transz to move it up and down
'             the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
' switch:       ROM switch number
' animate:      Array slot for handling the animation instrucitons, set to 0
'           Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target
'   isDropped:      Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.

Set DT34 = (new DropTarget)(sw34, sw34a, BM_sw34, 34, 0, false)
Set DT35 = (new DropTarget)(sw35, sw35a, BM_sw35, 35, 0, false)
Set DT36 = (new DropTarget)(sw36, sw36a, BM_sw36, 36, 0, false)

Set DT37 = (new DropTarget)(sw37, sw37a, BM_sw37, 37, 0, false)
Set DT38 = (new DropTarget)(sw38, sw38a, BM_sw38, 38, 0, false)
Set DT39 = (new DropTarget)(sw39, sw39a, BM_sw39, 39, 0, false)

Set DT42 = (new DropTarget)(sw42, sw42a, BM_sw42, 42, 0, false)
Set DT43 = (new DropTarget)(sw43, sw43a, BM_sw43, 43, 0, false)
Set DT44 = (new DropTarget)(sw44, sw44a, BM_sw44, 44, 0, false)

Set DT45 = (new DropTarget)(sw45, sw45a, BM_sw45, 45, 0, false)
Set DT46 = (new DropTarget)(sw46, sw46a, BM_sw46, 46, 0, false)
Set DT47 = (new DropTarget)(sw47, sw47a, BM_sw47, 47, 0, false)

Dim DTArray
DTArray = Array(DT34, DT35, DT36, DT37, DT38, DT39, DT42, DT43, DT44, DT45, DT46, DT47)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
  Dim i
  i = DTArrayID(switch)

  PlayTargetSound
  DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
  If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
    DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate =  - 1
  DoDTAnim
End Sub

Sub DTDrop(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate = 1
  DoDTAnim
End Sub

Function DTArrayID(switch)
  Dim i
  For i = 0 To UBound(DTArray)
    If DTArray(i).sw = switch Then
      DTArrayID = i
      Exit Function
    End If
  Next
End Function

Sub DTBallPhysics(aBall, angle, mass)
  Dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * Cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Cos(rangle + 4 * Atn(1) / 2)
  calc3 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Sin(rangle + 4 * Atn(1) / 2)

  aBall.velx = calc1 * Cos(rangle) + calc2
  aBall.vely = calc1 * Sin(rangle) + calc3
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  Dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - Tan(bangle) * aball.x + Tan(rangle2) * dtprim.x) / (Tan(rangle2) - Tan(bangle))
  Yintersect = Tan(rangle2) * Xintersect + (dtprim.y - Tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
  paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

  perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

  If perpvel > 0 And  perpvelafter <= 0 Then
    If DTEnableBrick = 1 And  perpvel > DTBrickVel And DTBrickVel <> 0 And cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
    DTCheckBrick = 4
  Else
    DTCheckBrick = 0
  End If
End Function

Sub DoDTAnim()
  Dim i
  For i = 0 To UBound(DTArray)
    DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
  Next
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
  Dim transz, switchid
  Dim animtime, rangle

  switchid = switch

  Dim ind
  ind = DTArrayID(switchid)

  rangle = prim.rotz * PI / 180

  DTAnimate = animate

  If animate = 0 Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  ElseIf primary.uservalue = 0 Then
    primary.uservalue = GameTime
  End If

  animtime = GameTime - primary.uservalue

  If (animate = 1 Or animate = 4) And animtime < DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    DTAnimate = animate
    Exit Function
  ElseIf (animate = 1 Or animate = 4) And animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 1 'If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0 'updated by rothbauerw to account for edge case
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
  End If

  If animate = 2 Then
    transz = (animtime - DTDropDelay) / DTDropSpeed * DTDropUnits *  - 1
    If prim.transz >  - DTDropUnits  Then
      prim.transz = transz
    End If

    prim.rotx = DTMaxBend * Cos(rangle) / 2
    prim.roty = DTMaxBend * Sin(rangle) / 2

    If prim.transz <= - DTDropUnits Then
      prim.transz =  - DTDropUnits
      secondary.collidable = 0
      DTArray(ind).isDropped = True 'Mark target as dropped
      controller.Switch(Switchid mod 100) = 1
      primary.uservalue = 0
      DTAnimate = 0
      Exit Function
    Else
      DTAnimate = 2
      Exit Function
    End If
  End If

  If animate = 3 And animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
  ElseIf animate = 3 And animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  If animate =  - 1 Then
    transz = (1 - (animtime) / DTDropUpSpeed) * DTDropUnits *  - 1

    If prim.transz =  - DTDropUnits Then
      Dim b
      'Dim gBOT
      'gBOT = GetBalls

      For b = 0 To UBound(gBOT)
        If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < prim.z + DTDropUnits + 25 Then
          gBOT(b).velz = 20
        End If
      Next
    End If

    If prim.transz < 0 Then
      prim.transz = transz
    ElseIf transz > 0 Then
      prim.transz = transz
    End If

    If prim.transz > DTDropUpUnits Then
      DTAnimate =  - 2
      prim.transz = DTDropUpUnits
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = GameTime
    End If
    primary.collidable = 0
    secondary.collidable = 1
    DTArray(ind).isDropped = False 'Mark target as not dropped
    controller.Switch(Switchid mod 100) = 0
  End If

  If animate =  - 2 And animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay) / DTDropSpeed * DTDropUnits *  - 1 + DTDropUpUnits
    If prim.transz < 0 Then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    End If
  End If
End Function

Function DTDropped(switchid)
  Dim ind
  ind = DTArrayID(switchid)

  DTDropped = DTArray(ind).isDropped
End Function


'******************************************************
'****  END DROP TARGETS
'******************************************************


'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7-1.0

sub TargetBouncer(aBall,defvalue)
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled = 1 and aball.z < 30 then
        'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        vel = BallSpeed(aBall)
        if aBall.velx = 0 then vratio = 1 else vratio = aBall.vely/aBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1: zMultiplier = 0.2*defvalue
      Case 2: zMultiplier = 0.25*defvalue
            Case 3: zMultiplier = 0.3*defvalue
      Case 4: zMultiplier = 0.4*defvalue
            Case 5: zMultiplier = 0.45*defvalue
            Case 6: zMultiplier = 0.5*defvalue
        End Select
        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
        aBall.vely = aBall.velx * vratio
        'debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        'debug.print "conservation check: " & BallSpeed(aBall)/vel
  end if
end sub

'****************************************************************
'****  END VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'****************************************************************



'***************************************************************
' ZSHA: Ambient ball shadows
'***************************************************************

' For dynamic ball shadows, Check the "Raytraced ball shadows" box for the specific light.
' Also make sure the light's z position is around 25 (mid ball)

'Ambient (Room light source)
Const AmbientBSFactor = 0.9    '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(7)

'Initialization
BSInit

Sub BSInit()
  Dim iii
  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 3 + iii / 1000
    objBallShadow(iii).visible = 0
  Next
End Sub


Sub BSUpdate
  Dim s: For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow

    'Primitive shadow on playfield, flasher shadow in ramps
    '** If on main and upper pf
    If gBOT(s).Z > 20 And gBOT(s).Z < 30 Then
      objBallShadow(s).visible = 1
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
      'objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25

    '** No shadow if ball is off the main playfield (this may need to be adjusted per table)
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub







'*********************
' Digital Display
'*********************
Dim Digits(32)
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

' Credits
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)


Dim VRDigits(32)
' 1st Player
VRDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6)
VRDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)
VRDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)
VRDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6)
VRDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)
VRDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)
VRDigits(6) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6)

' 2nd Player
VRDigits(7) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6)
VRDigits(8) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)
VRDigits(9) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)
VRDigits(10) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6)
VRDigits(11) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)
VRDigits(12) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)
VRDigits(13) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6)

' 3rd Player
VRDigits(14) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006)
VRDigits(15) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106)
VRDigits(16) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206)
VRDigits(17) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306)
VRDigits(18) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406)
VRDigits(19) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506)
VRDigits(20) = Array(LED1x600,LED1x601,LED1x602,LED1x603,LED1x604,LED1x605,LED1x606)

' 4th Player
VRDigits(21) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006)
VRDigits(22) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106)
VRDigits(23) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206)
VRDigits(24) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306)
VRDigits(25) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406)
VRDigits(26) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506)
VRDigits(27) = Array(LED2x600,LED2x601,LED2x602,LED2x603,LED2x604,LED2x605,LED2x606)

'Credits
VRDigits(28) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306)
VRDigits(29) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406)
'Balls
VRDigits(30) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506)
VRDigits(31) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606)


Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    If Renderingmode = 2 Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        if (num < 32) then
        For Each obj In VRDigits(num)
          If chg And 1 Then obj.visible=stat And 1
            chg=chg\2 : stat=stat\2
                Next
        Else
        end if
      Next
    Else

      If DesktopMode = True Then
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
    End if
  end if
End Sub

'**********************************************************************
' Dip Switches by Inkochnito
'**********************************************************************
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Spectrum - DIP switches"
        .AddFrame 2, 0, 190, "Maximum credits", &H03000000, Array("10 credits", 0, "15 credits", &H01000000, "25 credits", &H02000000, "40 credits", &H03000000)         'dip 25&26
        .AddFrame 2, 74, 190, "Multiplier will step when making", &H00000080, Array("4 green buttons from left to right", 0, "4 green buttons in any order", &H00000080) 'dip 8
        .AddFrame 2, 120, 190, "Outlane special will lite after", &H00200000, Array("making 4 stars", 0, "making 3 stars", &H00200000)                                   'dip 22
        .AddFrame 205, 0, 190, "Balls per game", &HC0000000, Array("2 balls", &HC0000000, "3 balls", 0, "4 balls", &H80000000, "5 balls", &H40000000)                    'dip 31&32
        .AddFrame 205, 74, 190, "4 circling lites will step down after", &H00800000, Array("hitting 3 same targets", 0, "hitting 4 same targets", &H00800000)            'dip 24
        .AddFrame 205, 120, 190, "Replay limit", &H10000000, Array("1 replay per game", 0, "unlimited", &H10000000)                                                      'dip 29
        .AddChk 2, 170, 120, Array("Attract sound", &H20000000)                                                                                                          'dip 30
        .AddChk 2, 185, 240, Array("Any color drop target lites out memory", &H00004000)                                                                                 'dip 15
        .AddChk 2, 200, 240, Array("Any color drop target flashing lites memory", 32768)                                                                                 'dip 16
        .AddChk 2, 215, 190, Array("Left && right spinners memory", &H00000020)                                                                                          'dip 6
        .AddChk 2, 230, 398, Array("Hitting 2 or 3 same targets to step 4 circling computer lites down 1 step", &H00400000)                                              'dip 23
        .AddChk 250, 170, 140, Array("Outlane special memory", &H00002000)                                                                                               'dip 14
        .AddChk 250, 185, 120, Array("Match feature", &H08000000)                                                                                                        'dip 28
        .AddChk 250, 200, 120, Array("Credits displayed", &H04000000)                                                                                                    'dip 27
        .AddChk 250, 215, 150, Array("2X, 3X, 4X lite memory", &H00000040)                                                                                               'dip 7
        .AddLabel 50, 260, 320, 20, "Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
        .AddLabel 50, 280, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")

'**********************************************************************
'**********************************************************************

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BM_PincabRails: BP_BM_PincabRails=Array(BM_BM_PincabRails, LM_GI_GI_16_BM_PincabRails, LM_GI_GI_17_BM_PincabRails, LM_GI_GI_23_BM_PincabRails, LM_GI_GI_24_BM_PincabRails, LM_GI_GI_27_BM_PincabRails, LM_GI_GI_5_BM_PincabRails, LM_GI_GI_7_BM_PincabRails, LM_GI_GI_8_BM_PincabRails, LM_IN_below_l38_BM_PincabRails)
Dim BP_Layer0: BP_Layer0=Array(BM_Layer0, LM_GI_GI_19_Layer0, LM_GI_GI_2_Layer0, LM_GI_GI_20_Layer0, LM_GI_GI_23_Layer0, LM_GI_GI_24_Layer0, LM_GI_GI_25_Layer0, LM_GI_GI_26_Layer0, LM_GI_GI_27_Layer0, LM_GI_GI_28_Layer0, LM_GI_GI_29_Layer0, LM_GI_GI_3_Layer0, LM_GI_GI_30_Layer0, LM_GI_GI_5_Layer0, LM_GI_GI_6_Layer0, LM_IN_below_l1_Layer0, LM_IN_below_l150_Layer0, LM_IN_above_l150_Layer0, LM_IN_below_l151_Layer0, LM_IN_above_l151_Layer0, LM_IN_below_l152_Layer0, LM_IN_above_l152_Layer0, LM_IN_below_l153_Layer0, LM_IN_above_l153_Layer0, LM_IN_below_l154_Layer0, LM_IN_above_l154_Layer0, LM_IN_below_l155_Layer0, LM_IN_above_l155_Layer0, LM_IN_below_l156_Layer0, LM_IN_above_l156_Layer0, LM_IN_below_l157_Layer0, LM_IN_above_l157_Layer0, LM_IN_below_l158_Layer0, LM_IN_above_l158_Layer0, LM_IN_below_l159_Layer0, LM_IN_above_l159_Layer0, LM_IN_below_l19_Layer0, LM_IN_below_l20_Layer0, LM_IN_below_l33_Layer0, LM_IN_below_l36_Layer0, LM_IN_above_l49_Layer0)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_GI_GI_001_Parts, LM_GI_GI_002_Parts, LM_GI_GI_10_Parts, LM_GI_GI_11_Parts, LM_GI_GI_12_Parts, LM_GI_GI_13_Parts, LM_GI_GI_14_Parts, LM_GI_GI_15_Parts, LM_GI_GI_16_Parts, LM_GI_GI_17_Parts, LM_GI_GI_18_Parts, LM_GI_GI_19_Parts, LM_GI_GI_2_Parts, LM_GI_GI_20_Parts, LM_GI_GI_22222_Parts, LM_GI_GI_23_Parts, LM_GI_GI_24_Parts, LM_GI_GI_25_Parts, LM_GI_GI_26_Parts, LM_GI_GI_27_Parts, LM_GI_GI_28_Parts, LM_GI_GI_29_Parts, LM_GI_GI_3_Parts, LM_GI_GI_30_Parts, LM_GI_GI_34_Parts, LM_GI_GI_35_Parts, LM_GI_GI_36_Parts, LM_GI_GI_37_Parts, LM_GI_GI_38_Parts, LM_GI_GI_39_Parts, LM_GI_GI_5_Parts, LM_GI_GI_6_Parts, LM_GI_GI_7_Parts, LM_GI_GI_8_Parts, LM_GI_GI_9_Parts, LM_IN_above_l1_Parts, LM_IN_below_l150_Parts, LM_IN_above_l150_Parts, LM_IN_below_l151_Parts, LM_IN_above_l151_Parts, LM_IN_below_l152_Parts, LM_IN_above_l152_Parts, LM_IN_below_l153_Parts, LM_IN_above_l153_Parts, LM_IN_below_l154_Parts, LM_IN_above_l154_Parts, LM_IN_below_l155_Parts, LM_IN_above_l155_Parts, _
  LM_IN_below_l156_Parts, LM_IN_above_l156_Parts, LM_IN_below_l157_Parts, LM_IN_above_l157_Parts, LM_IN_below_l158_Parts, LM_IN_above_l158_Parts, LM_IN_below_l159_Parts, LM_IN_above_l159_Parts, LM_IN_below_l17_Parts, LM_IN_above_l17_Parts, LM_IN_below_l19_Parts, LM_IN_below_l2_Parts, LM_IN_below_l20_Parts, LM_IN_below_l21_Parts, LM_IN_above_l21_Parts, LM_IN_below_l22_Parts, LM_IN_below_l26_Parts, LM_IN_above_l26_Parts, LM_IN_below_l33_Parts, LM_IN_above_l33_Parts, LM_IN_below_l36_Parts, LM_IN_below_l37_Parts, LM_IN_below_l38_Parts, LM_IN_above_l39_Parts, LM_IN_below_l42_Parts, LM_IN_above_l49_Parts, LM_IN_below_l4_Parts, LM_IN_below_l5_Parts, LM_IN_below_l50_Parts, LM_IN_above_l50_Parts, LM_IN_below_l52_Parts, LM_IN_below_l53_Parts, LM_IN_below_l54_Parts, LM_IN_below_l58_Parts, LM_IN_above_l58_Parts, LM_IN_below_l59_Parts, LM_IN_above_l81_Parts, LM_IN_above_l83_Parts, LM_IN_above_l84_Parts, LM_IN_above_l89_Parts, LM_IN_above_l92_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GI_GI_001_Playfield, LM_GI_GI_002_Playfield, LM_GI_GI_10_Playfield, LM_GI_GI_11_Playfield, LM_GI_GI_12_Playfield, LM_GI_GI_13_Playfield, LM_GI_GI_14_Playfield, LM_GI_GI_15_Playfield, LM_GI_GI_16_Playfield, LM_GI_GI_17_Playfield, LM_GI_GI_18_Playfield, LM_GI_GI_19_Playfield, LM_GI_GI_2_Playfield, LM_GI_GI_20_Playfield, LM_GI_GI_22222_Playfield, LM_GI_GI_23_Playfield, LM_GI_GI_24_Playfield, LM_GI_GI_25_Playfield, LM_GI_GI_26_Playfield, LM_GI_GI_27_Playfield, LM_GI_GI_28_Playfield, LM_GI_GI_29_Playfield, LM_GI_GI_3_Playfield, LM_GI_GI_30_Playfield, LM_GI_GI_34_Playfield, LM_GI_GI_35_Playfield, LM_GI_GI_36_Playfield, LM_GI_GI_37_Playfield, LM_GI_GI_38_Playfield, LM_GI_GI_39_Playfield, LM_GI_GI_5_Playfield, LM_GI_GI_6_Playfield, LM_GI_GI_7_Playfield, LM_GI_GI_8_Playfield, LM_GI_GI_9_Playfield, LM_IN_above_l1_Playfield, LM_IN_above_l10_Playfield, LM_IN_above_l100_Playfield, LM_IN_above_l101_Playfield, LM_IN_above_l102_Playfield, LM_IN_above_l103_Playfield, _
  LM_IN_above_l104_Playfield, LM_IN_above_l105_Playfield, LM_IN_above_l106_Playfield, LM_IN_above_l107_Playfield, LM_IN_above_l108_Playfield, LM_IN_above_l113_Playfield, LM_IN_above_l114_Playfield, LM_IN_above_l115_Playfield, LM_IN_above_l116_Playfield, LM_IN_above_l117_Playfield, LM_IN_above_l118_Playfield, LM_IN_above_l119_Playfield, LM_IN_above_l120_Playfield, LM_IN_above_l121_Playfield, LM_IN_above_l122_Playfield, LM_IN_above_l123_Playfield, LM_IN_above_l124_Playfield, LM_IN_below_l150_Playfield, LM_IN_above_l150_Playfield, LM_IN_below_l151_Playfield, LM_IN_above_l151_Playfield, LM_IN_below_l152_Playfield, LM_IN_above_l152_Playfield, LM_IN_below_l153_Playfield, LM_IN_above_l153_Playfield, LM_IN_below_l154_Playfield, LM_IN_above_l154_Playfield, LM_IN_below_l155_Playfield, LM_IN_above_l155_Playfield, LM_IN_below_l156_Playfield, LM_IN_above_l156_Playfield, LM_IN_below_l157_Playfield, LM_IN_above_l157_Playfield, LM_IN_below_l158_Playfield, LM_IN_above_l158_Playfield, LM_IN_below_l159_Playfield, _
  LM_IN_above_l159_Playfield, LM_IN_above_l17_Playfield, LM_IN_above_l18_Playfield, LM_IN_above_l19_Playfield, LM_IN_above_l2_Playfield, LM_IN_below_l20_Playfield, LM_IN_above_l21_Playfield, LM_IN_above_l22_Playfield, LM_IN_above_l23_Playfield, LM_IN_above_l24_Playfield, LM_IN_above_l25_Playfield, LM_IN_above_l26_Playfield, LM_IN_above_l3_Playfield, LM_IN_above_l33_Playfield, LM_IN_above_l34_Playfield, LM_IN_above_l35_Playfield, LM_IN_below_l36_Playfield, LM_IN_above_l37_Playfield, LM_IN_above_l38_Playfield, LM_IN_above_l39_Playfield, LM_IN_above_l40_Playfield, LM_IN_above_l41_Playfield, LM_IN_below_l42_Playfield, LM_IN_above_l49_Playfield, LM_IN_below_l4_Playfield, LM_IN_above_l5_Playfield, LM_IN_above_l50_Playfield, LM_IN_above_l51_Playfield, LM_IN_below_l52_Playfield, LM_IN_above_l53_Playfield, LM_IN_above_l54_Playfield, LM_IN_above_l55_Playfield, LM_IN_above_l56_Playfield, LM_IN_above_l57_Playfield, LM_IN_above_l58_Playfield, LM_IN_below_l59_Playfield, LM_IN_above_l6_Playfield, LM_IN_above_l65_Playfield, _
  LM_IN_above_l66_Playfield, LM_IN_above_l67_Playfield, LM_IN_above_l68_Playfield, LM_IN_above_l69_Playfield, LM_IN_above_l7_Playfield, LM_IN_above_l70_Playfield, LM_IN_above_l71_Playfield, LM_IN_above_l72_Playfield, LM_IN_above_l73_Playfield, LM_IN_above_l74_Playfield, LM_IN_above_l75_Playfield, LM_IN_above_l76_Playfield, LM_IN_above_l8_Playfield, LM_IN_above_l81_Playfield, LM_IN_above_l82_Playfield, LM_IN_above_l83_Playfield, LM_IN_above_l84_Playfield, LM_IN_above_l85_Playfield, LM_IN_above_l86_Playfield, LM_IN_above_l87_Playfield, LM_IN_above_l88_Playfield, LM_IN_above_l89_Playfield, LM_IN_above_l9_Playfield, LM_IN_above_l90_Playfield, LM_IN_above_l91_Playfield, LM_IN_above_l92_Playfield, LM_IN_above_l97_Playfield, LM_IN_above_l98_Playfield, LM_IN_above_l99_Playfield)
Dim BP_flip_l: BP_flip_l=Array(BM_flip_l, LM_GI_GI_001_flip_l, LM_GI_GI_002_flip_l, LM_GI_GI_10_flip_l, LM_GI_GI_11_flip_l, LM_GI_GI_13_flip_l, LM_GI_GI_14_flip_l, LM_GI_GI_22222_flip_l, LM_GI_GI_34_flip_l, LM_GI_GI_35_flip_l, LM_GI_GI_36_flip_l, LM_GI_GI_37_flip_l, LM_GI_GI_38_flip_l, LM_GI_GI_39_flip_l, LM_GI_GI_9_flip_l, LM_IN_below_l42_flip_l, LM_IN_below_l59_flip_l)
Dim BP_flip_l_up: BP_flip_l_up=Array(BM_flip_l_up, LM_GI_GI_001_flip_l_up, LM_GI_GI_10_flip_l_up, LM_GI_GI_11_flip_l_up, LM_GI_GI_12_flip_l_up, LM_GI_GI_13_flip_l_up, LM_GI_GI_22222_flip_l_up, LM_GI_GI_34_flip_l_up, LM_GI_GI_35_flip_l_up, LM_GI_GI_36_flip_l_up, LM_GI_GI_37_flip_l_up, LM_GI_GI_38_flip_l_up, LM_GI_GI_39_flip_l_up, LM_GI_GI_9_flip_l_up, LM_IN_below_l23_flip_l_up, LM_IN_above_l23_flip_l_up, LM_IN_above_l39_flip_l_up, LM_IN_above_l55_flip_l_up, LM_IN_below_l59_flip_l_up)
Dim BP_flip_r: BP_flip_r=Array(BM_flip_r, LM_GI_GI_001_flip_r, LM_GI_GI_002_flip_r, LM_GI_GI_10_flip_r, LM_GI_GI_12_flip_r, LM_GI_GI_14_flip_r, LM_GI_GI_22222_flip_r, LM_GI_GI_34_flip_r, LM_GI_GI_35_flip_r, LM_GI_GI_36_flip_r, LM_GI_GI_37_flip_r, LM_GI_GI_38_flip_r, LM_GI_GI_39_flip_r, LM_GI_GI_9_flip_r, LM_IN_below_l42_flip_r, LM_IN_below_l59_flip_r)
Dim BP_flip_r_up: BP_flip_r_up=Array(BM_flip_r_up, LM_GI_GI_002_flip_r_up, LM_GI_GI_10_flip_r_up, LM_GI_GI_11_flip_r_up, LM_GI_GI_12_flip_r_up, LM_GI_GI_22222_flip_r_up, LM_GI_GI_34_flip_r_up, LM_GI_GI_35_flip_r_up, LM_GI_GI_36_flip_r_up, LM_GI_GI_37_flip_r_up, LM_GI_GI_38_flip_r_up, LM_GI_GI_39_flip_r_up, LM_GI_GI_9_flip_r_up, LM_IN_above_l23_flip_r_up, LM_IN_above_l39_flip_r_up, LM_IN_below_l42_flip_r_up, LM_IN_below_l55_flip_r_up, LM_IN_above_l55_flip_r_up)
Dim BP_flipr_l: BP_flipr_l=Array(BM_flipr_l, LM_GI_GI_001_flipr_l, LM_GI_GI_10_flipr_l, LM_GI_GI_11_flipr_l, LM_GI_GI_22222_flipr_l, LM_GI_GI_34_flipr_l, LM_GI_GI_35_flipr_l, LM_GI_GI_37_flipr_l, LM_GI_GI_38_flipr_l, LM_IN_below_l59_flipr_l)
Dim BP_flipr_l_up: BP_flipr_l_up=Array(BM_flipr_l_up, LM_GI_GI_001_flipr_l_up, LM_GI_GI_10_flipr_l_up, LM_GI_GI_11_flipr_l_up, LM_GI_GI_22222_flipr_l_up, LM_GI_GI_34_flipr_l_up, LM_GI_GI_35_flipr_l_up, LM_GI_GI_37_flipr_l_up, LM_GI_GI_38_flipr_l_up, LM_IN_above_l23_flipr_l_up, LM_IN_below_l59_flipr_l_up, LM_IN_above_l7_flipr_l_up)
Dim BP_flipr_r: BP_flipr_r=Array(BM_flipr_r, LM_GI_GI_002_flipr_r, LM_GI_GI_12_flipr_r, LM_GI_GI_22222_flipr_r, LM_GI_GI_34_flipr_r, LM_GI_GI_35_flipr_r, LM_GI_GI_36_flipr_r, LM_GI_GI_39_flipr_r, LM_GI_GI_9_flipr_r, LM_IN_below_l42_flipr_r)
Dim BP_flipr_r_up: BP_flipr_r_up=Array(BM_flipr_r_up, LM_GI_GI_002_flipr_r_up, LM_GI_GI_12_flipr_r_up, LM_GI_GI_22222_flipr_r_up, LM_GI_GI_34_flipr_r_up, LM_GI_GI_35_flipr_r_up, LM_GI_GI_36_flipr_r_up, LM_GI_GI_39_flipr_r_up, LM_GI_GI_9_flipr_r_up, LM_IN_below_l42_flipr_r_up, LM_IN_below_l55_flipr_r_up, LM_IN_above_l55_flipr_r_up, LM_IN_above_l8_flipr_r_up)
Dim BP_kick_sw17: BP_kick_sw17=Array(BM_kick_sw17, LM_GI_GI_001_kick_sw17, LM_GI_GI_10_kick_sw17, LM_GI_GI_11_kick_sw17, LM_GI_GI_13_kick_sw17, LM_GI_GI_8_kick_sw17)
Dim BP_kick_sw24: BP_kick_sw24=Array(BM_kick_sw24, LM_GI_GI_002_kick_sw24, LM_GI_GI_12_kick_sw24, LM_GI_GI_14_kick_sw24, LM_GI_GI_7_kick_sw24, LM_GI_GI_9_kick_sw24)
Dim BP_kick_sw41: BP_kick_sw41=Array(BM_kick_sw41, LM_GI_GI_12_kick_sw41, LM_GI_GI_14_kick_sw41, LM_GI_GI_15_kick_sw41, LM_GI_GI_17_kick_sw41, LM_GI_GI_8_kick_sw41, LM_GI_GI_9_kick_sw41)
Dim BP_kick_sw48: BP_kick_sw48=Array(BM_kick_sw48, LM_GI_GI_10_kick_sw48, LM_GI_GI_11_kick_sw48, LM_GI_GI_14_kick_sw48, LM_GI_GI_16_kick_sw48, LM_GI_GI_18_kick_sw48, LM_GI_GI_7_kick_sw48, LM_IN_above_l90_kick_sw48)
Dim BP_kick_sw7l: BP_kick_sw7l=Array(BM_kick_sw7l, LM_GI_GI_25_kick_sw7l, LM_GI_GI_26_kick_sw7l, LM_IN_below_l20_kick_sw7l, LM_IN_below_l36_kick_sw7l)
Dim BP_kick_sw7r: BP_kick_sw7r=Array(BM_kick_sw7r, LM_GI_GI_23_kick_sw7r, LM_GI_GI_25_kick_sw7r, LM_GI_GI_26_kick_sw7r)
Dim BP_sw18: BP_sw18=Array(BM_sw18, LM_GI_GI_11_sw18, LM_GI_GI_13_sw18, LM_GI_GI_17_sw18, LM_GI_GI_8_sw18, LM_GI_GI_9_sw18)
Dim BP_sw19: BP_sw19=Array(BM_sw19, LM_GI_GI_23_sw19, LM_GI_GI_25_sw19, LM_GI_GI_26_sw19, LM_GI_GI_27_sw19, LM_GI_GI_5_sw19, LM_IN_below_l150_sw19, LM_IN_above_l150_sw19, LM_IN_below_l20_sw19, LM_IN_below_l36_sw19)
Dim BP_sw19a: BP_sw19a=Array(BM_sw19a, LM_GI_GI_23_sw19a, LM_GI_GI_26_sw19a, LM_GI_GI_27_sw19a, LM_GI_GI_29_sw19a, LM_GI_GI_3_sw19a, LM_GI_GI_5_sw19a, LM_IN_below_l151_sw19a, LM_IN_above_l151_sw19a, LM_IN_below_l20_sw19a, LM_IN_below_l36_sw19a)
Dim BP_sw19b: BP_sw19b=Array(BM_sw19b, LM_GI_GI_19_sw19b, LM_GI_GI_23_sw19b, LM_GI_GI_26_sw19b, LM_GI_GI_27_sw19b, LM_GI_GI_29_sw19b, LM_IN_below_l152_sw19b, LM_IN_above_l152_sw19b, LM_IN_below_l20_sw19b)
Dim BP_sw20: BP_sw20=Array(BM_sw20, LM_GI_GI_002_sw20, LM_GI_GI_10_sw20, LM_GI_GI_11_sw20, LM_GI_GI_13_sw20, LM_GI_GI_8_sw20, LM_GI_GI_9_sw20)
Dim BP_sw21: BP_sw21=Array(BM_sw21, LM_GI_GI_24_sw21, LM_GI_GI_25_sw21, LM_GI_GI_26_sw21, LM_GI_GI_28_sw21, LM_GI_GI_30_sw21, LM_IN_below_l155_sw21, LM_IN_above_l155_sw21, LM_IN_below_l20_sw21, LM_IN_below_l36_sw21)
Dim BP_sw21a: BP_sw21a=Array(BM_sw21a, LM_GI_GI_2_sw21a, LM_GI_GI_23_sw21a, LM_GI_GI_24_sw21a, LM_GI_GI_25_sw21a, LM_GI_GI_28_sw21a, LM_GI_GI_30_sw21a, LM_GI_GI_6_sw21a, LM_IN_below_l156_sw21a, LM_IN_above_l156_sw21a, LM_IN_below_l159_sw21a, LM_IN_below_l36_sw21a)
Dim BP_sw21b: BP_sw21b=Array(BM_sw21b, LM_GI_GI_20_sw21b, LM_GI_GI_24_sw21b, LM_GI_GI_25_sw21b, LM_GI_GI_28_sw21b, LM_GI_GI_6_sw21b, LM_IN_below_l157_sw21b, LM_IN_above_l157_sw21b, LM_IN_below_l36_sw21b)
Dim BP_sw22: BP_sw22=Array(BM_sw22, LM_GI_GI_10_sw22, LM_GI_GI_12_sw22, LM_GI_GI_14_sw22, LM_GI_GI_18_sw22, LM_GI_GI_7_sw22, LM_GI_GI_9_sw22)
Dim BP_sw23: BP_sw23=Array(BM_sw23, LM_GI_GI_10_sw23, LM_GI_GI_14_sw23, LM_GI_GI_7_sw23)
Dim BP_sw26: BP_sw26=Array(BM_sw26, LM_GI_GI_19_sw26, LM_GI_GI_23_sw26, LM_GI_GI_25_sw26, LM_GI_GI_26_sw26, LM_GI_GI_27_sw26, LM_GI_GI_28_sw26, LM_GI_GI_3_sw26, LM_IN_below_l1_sw26, LM_IN_above_l1_sw26, LM_IN_above_l17_sw26, LM_IN_below_l18_sw26, LM_IN_below_l2_sw26, LM_IN_above_l2_sw26, LM_IN_below_l20_sw26, LM_IN_below_l21_sw26, LM_IN_below_l36_sw26, LM_IN_below_l5_sw26)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_GI_GI_15_sw27, LM_GI_GI_16_sw27, LM_GI_GI_17_sw27, LM_GI_GI_19_sw27, LM_GI_GI_20_sw27, LM_GI_GI_26_sw27, LM_GI_GI_3_sw27, LM_IN_below_l2_sw27, LM_IN_above_l2_sw27, LM_IN_below_l20_sw27, LM_IN_below_l36_sw27, LM_IN_below_l4_sw27, LM_IN_below_l52_sw27, LM_IN_above_l57_sw27)
Dim BP_sw28: BP_sw28=Array(BM_sw28, LM_GI_GI_15_sw28, LM_GI_GI_16_sw28, LM_GI_GI_17_sw28, LM_GI_GI_18_sw28, LM_GI_GI_19_sw28, LM_GI_GI_25_sw28, LM_GI_GI_28_sw28, LM_GI_GI_3_sw28, LM_IN_below_l18_sw28, LM_IN_above_l18_sw28, LM_IN_below_l20_sw28, LM_IN_below_l36_sw28, LM_IN_below_l4_sw28, LM_IN_below_l52_sw28)
Dim BP_sw29: BP_sw29=Array(BM_sw29, LM_GI_GI_15_sw29, LM_GI_GI_16_sw29, LM_GI_GI_17_sw29, LM_GI_GI_18_sw29, LM_GI_GI_20_sw29, LM_GI_GI_27_sw29, LM_GI_GI_7_sw29, LM_IN_below_l20_sw29, LM_IN_below_l34_sw29, LM_IN_above_l34_sw29, LM_IN_below_l36_sw29, LM_IN_below_l4_sw29, LM_IN_below_l50_sw29, LM_IN_below_l52_sw29)
Dim BP_sw30: BP_sw30=Array(BM_sw30, LM_GI_GI_16_sw30, LM_GI_GI_17_sw30, LM_GI_GI_18_sw30, LM_GI_GI_19_sw30, LM_GI_GI_2_sw30, LM_GI_GI_20_sw30, LM_GI_GI_25_sw30, LM_IN_below_l20_sw30, LM_IN_below_l36_sw30, LM_IN_above_l49_sw30, LM_IN_below_l4_sw30, LM_IN_below_l50_sw30, LM_IN_above_l50_sw30, LM_IN_below_l52_sw30)
Dim BP_sw31: BP_sw31=Array(BM_sw31, LM_GI_GI_2_sw31, LM_GI_GI_20_sw31, LM_GI_GI_24_sw31, LM_GI_GI_25_sw31, LM_GI_GI_26_sw31, LM_GI_GI_28_sw31, LM_IN_below_l20_sw31, LM_IN_above_l33_sw31, LM_IN_below_l34_sw31, LM_IN_below_l36_sw31, LM_IN_below_l37_sw31, LM_IN_above_l37_sw31, LM_IN_below_l49_sw31, LM_IN_above_l49_sw31, LM_IN_below_l50_sw31, LM_IN_below_l53_sw31)
Dim BP_sw33: BP_sw33=Array(BM_sw33, LM_GI_GI_23_sw33, LM_GI_GI_24_sw33, LM_GI_GI_25_sw33, LM_GI_GI_26_sw33, LM_GI_GI_27_sw33, LM_GI_GI_5_sw33, LM_IN_below_l153_sw33, LM_IN_above_l153_sw33, LM_IN_below_l20_sw33, LM_IN_below_l36_sw33)
Dim BP_sw33a: BP_sw33a=Array(BM_sw33a, LM_GI_GI_23_sw33a, LM_GI_GI_24_sw33a, LM_GI_GI_25_sw33a, LM_GI_GI_26_sw33a, LM_GI_GI_27_sw33a, LM_IN_below_l150_sw33a, LM_IN_below_l154_sw33a, LM_IN_above_l154_sw33a, LM_IN_below_l20_sw33a, LM_IN_below_l36_sw33a)
Dim BP_sw34: BP_sw34=Array(BM_sw34, LM_GI_GI_23_sw34, LM_GI_GI_26_sw34, LM_GI_GI_27_sw34, LM_GI_GI_3_sw34, LM_IN_below_l20_sw34)
Dim BP_sw35: BP_sw35=Array(BM_sw35, LM_GI_GI_26_sw35, LM_GI_GI_27_sw35, LM_IN_below_l20_sw35)
Dim BP_sw36: BP_sw36=Array(BM_sw36, LM_GI_GI_26_sw36, LM_IN_below_l19_sw36, LM_IN_below_l20_sw36)
Dim BP_sw37: BP_sw37=Array(BM_sw37, LM_GI_GI_19_sw37, LM_GI_GI_2_sw37, LM_GI_GI_20_sw37, LM_GI_GI_24_sw37, LM_GI_GI_25_sw37, LM_GI_GI_26_sw37, LM_GI_GI_28_sw37, LM_GI_GI_3_sw37, LM_IN_below_l158_sw37, LM_IN_above_l158_sw37, LM_IN_below_l159_sw37, LM_IN_below_l19_sw37, LM_IN_below_l20_sw37, LM_IN_below_l36_sw37, LM_IN_below_l4_sw37, LM_IN_below_l52_sw37)
Dim BP_sw38: BP_sw38=Array(BM_sw38, LM_GI_GI_19_sw38, LM_GI_GI_2_sw38, LM_GI_GI_20_sw38, LM_GI_GI_24_sw38, LM_GI_GI_25_sw38, LM_GI_GI_26_sw38, LM_GI_GI_28_sw38, LM_IN_below_l158_sw38, LM_IN_above_l158_sw38, LM_IN_below_l159_sw38, LM_IN_above_l159_sw38, LM_IN_below_l36_sw38, LM_IN_below_l52_sw38)
Dim BP_sw39: BP_sw39=Array(BM_sw39, LM_GI_GI_19_sw39, LM_GI_GI_2_sw39, LM_GI_GI_20_sw39, LM_GI_GI_24_sw39, LM_GI_GI_25_sw39, LM_GI_GI_28_sw39, LM_IN_below_l158_sw39, LM_IN_below_l159_sw39, LM_IN_above_l159_sw39, LM_IN_below_l36_sw39, LM_IN_above_l49_sw39)
Dim BP_sw40: BP_sw40=Array(BM_sw40, LM_GI_GI_23_sw40, LM_GI_GI_24_sw40, LM_GI_GI_25_sw40, LM_GI_GI_26_sw40, LM_GI_GI_28_sw40, LM_GI_GI_30_sw40, LM_IN_below_l159_sw40, LM_IN_above_l159_sw40, LM_IN_below_l20_sw40, LM_IN_below_l36_sw40)
Dim BP_sw40a: BP_sw40a=Array(BM_sw40a, LM_GI_GI_2_sw40a, LM_GI_GI_23_sw40a, LM_GI_GI_24_sw40a, LM_GI_GI_25_sw40a, LM_GI_GI_26_sw40a, LM_GI_GI_28_sw40a, LM_IN_below_l155_sw40a, LM_IN_below_l158_sw40a, LM_IN_above_l158_sw40a, LM_IN_below_l20_sw40a, LM_IN_below_l36_sw40a, LM_IN_below_l4_sw40a)
Dim BP_sw42: BP_sw42=Array(BM_sw42, LM_GI_GI_15_sw42, LM_GI_GI_17_sw42, LM_IN_below_l4_sw42)
Dim BP_sw43: BP_sw43=Array(BM_sw43, LM_GI_GI_15_sw43, LM_GI_GI_17_sw43, LM_IN_below_l4_sw43)
Dim BP_sw44: BP_sw44=Array(BM_sw44, LM_GI_GI_15_sw44, LM_GI_GI_17_sw44, LM_IN_below_l4_sw44)
Dim BP_sw45: BP_sw45=Array(BM_sw45, LM_GI_GI_15_sw45, LM_GI_GI_16_sw45, LM_GI_GI_17_sw45, LM_GI_GI_18_sw45, LM_IN_below_l52_sw45)
Dim BP_sw46: BP_sw46=Array(BM_sw46, LM_GI_GI_16_sw46, LM_GI_GI_17_sw46, LM_GI_GI_18_sw46, LM_IN_below_l52_sw46)
Dim BP_sw47: BP_sw47=Array(BM_sw47, LM_GI_GI_16_sw47, LM_GI_GI_18_sw47, LM_GI_GI_20_sw47, LM_IN_below_l4_sw47, LM_IN_below_l52_sw47)
Dim BP_underPF: BP_underPF=Array(BM_underPF, LM_GI_GI_11_underPF, LM_GI_GI_13_underPF, LM_GI_GI_14_underPF, LM_GI_GI_19_underPF, LM_GI_GI_2_underPF, LM_GI_GI_20_underPF, LM_GI_GI_23_underPF, LM_GI_GI_24_underPF, LM_GI_GI_25_underPF, LM_GI_GI_26_underPF, LM_GI_GI_27_underPF, LM_GI_GI_28_underPF, LM_GI_GI_29_underPF, LM_GI_GI_3_underPF, LM_GI_GI_30_underPF, LM_GI_GI_5_underPF, LM_GI_GI_6_underPF, LM_GI_GI_7_underPF, LM_GI_GI_8_underPF, LM_IN_below_l1_underPF, LM_IN_above_l1_underPF, LM_IN_below_l10_underPF, LM_IN_above_l10_underPF, LM_IN_below_l100_underPF, LM_IN_above_l100_underPF, LM_IN_below_l101_underPF, LM_IN_above_l101_underPF, LM_IN_below_l102_underPF, LM_IN_above_l102_underPF, LM_IN_below_l103_underPF, LM_IN_above_l103_underPF, LM_IN_below_l104_underPF, LM_IN_below_l105_underPF, LM_IN_below_l106_underPF, LM_IN_above_l106_underPF, LM_IN_below_l107_underPF, LM_IN_below_l108_underPF, LM_IN_below_l113_underPF, LM_IN_below_l114_underPF, LM_IN_below_l115_underPF, LM_IN_below_l116_underPF, _
  LM_IN_below_l117_underPF, LM_IN_below_l118_underPF, LM_IN_below_l119_underPF, LM_IN_below_l120_underPF, LM_IN_below_l121_underPF, LM_IN_below_l122_underPF, LM_IN_below_l123_underPF, LM_IN_below_l124_underPF, LM_IN_below_l150_underPF, LM_IN_above_l150_underPF, LM_IN_below_l151_underPF, LM_IN_above_l151_underPF, LM_IN_below_l152_underPF, LM_IN_above_l152_underPF, LM_IN_below_l153_underPF, LM_IN_above_l153_underPF, LM_IN_below_l154_underPF, LM_IN_above_l154_underPF, LM_IN_below_l155_underPF, LM_IN_above_l155_underPF, LM_IN_below_l156_underPF, LM_IN_above_l156_underPF, LM_IN_below_l157_underPF, LM_IN_above_l157_underPF, LM_IN_below_l158_underPF, LM_IN_above_l158_underPF, LM_IN_below_l159_underPF, LM_IN_above_l159_underPF, LM_IN_below_l17_underPF, LM_IN_above_l17_underPF, LM_IN_below_l18_underPF, LM_IN_above_l18_underPF, LM_IN_below_l19_underPF, LM_IN_above_l19_underPF, LM_IN_below_l2_underPF, LM_IN_above_l2_underPF, LM_IN_below_l20_underPF, LM_IN_below_l21_underPF, LM_IN_above_l21_underPF, _
  LM_IN_below_l22_underPF, LM_IN_above_l22_underPF, LM_IN_below_l23_underPF, LM_IN_below_l24_underPF, LM_IN_below_l25_underPF, LM_IN_above_l25_underPF, LM_IN_below_l26_underPF, LM_IN_above_l26_underPF, LM_IN_below_l3_underPF, LM_IN_above_l3_underPF, LM_IN_below_l33_underPF, LM_IN_above_l33_underPF, LM_IN_below_l34_underPF, LM_IN_above_l34_underPF, LM_IN_below_l35_underPF, LM_IN_above_l35_underPF, LM_IN_below_l36_underPF, LM_IN_below_l37_underPF, LM_IN_above_l37_underPF, LM_IN_below_l38_underPF, LM_IN_above_l38_underPF, LM_IN_below_l39_underPF, LM_IN_below_l40_underPF, LM_IN_below_l41_underPF, LM_IN_above_l41_underPF, LM_IN_below_l49_underPF, LM_IN_above_l49_underPF, LM_IN_below_l5_underPF, LM_IN_above_l5_underPF, LM_IN_below_l50_underPF, LM_IN_above_l50_underPF, LM_IN_below_l51_underPF, LM_IN_above_l51_underPF, LM_IN_below_l52_underPF, LM_IN_below_l53_underPF, LM_IN_above_l53_underPF, LM_IN_below_l54_underPF, LM_IN_above_l54_underPF, LM_IN_below_l55_underPF, LM_IN_below_l56_underPF, LM_IN_below_l57_underPF, _
  LM_IN_above_l57_underPF, LM_IN_below_l58_underPF, LM_IN_above_l58_underPF, LM_IN_below_l6_underPF, LM_IN_above_l6_underPF, LM_IN_below_l65_underPF, LM_IN_below_l66_underPF, LM_IN_below_l67_underPF, LM_IN_below_l68_underPF, LM_IN_below_l69_underPF, LM_IN_below_l7_underPF, LM_IN_below_l70_underPF, LM_IN_below_l71_underPF, LM_IN_below_l72_underPF, LM_IN_below_l73_underPF, LM_IN_below_l74_underPF, LM_IN_below_l75_underPF, LM_IN_below_l76_underPF, LM_IN_below_l8_underPF, LM_IN_below_l81_underPF, LM_IN_below_l82_underPF, LM_IN_above_l82_underPF, LM_IN_below_l83_underPF, LM_IN_below_l84_underPF, LM_IN_above_l84_underPF, LM_IN_below_l85_underPF, LM_IN_above_l85_underPF, LM_IN_below_l86_underPF, LM_IN_above_l86_underPF, LM_IN_below_l87_underPF, LM_IN_below_l88_underPF, LM_IN_above_l88_underPF, LM_IN_below_l89_underPF, LM_IN_above_l89_underPF, LM_IN_below_l9_underPF, LM_IN_above_l9_underPF, LM_IN_below_l90_underPF, LM_IN_above_l90_underPF, LM_IN_below_l91_underPF, LM_IN_above_l91_underPF, LM_IN_below_l92_underPF, _
  LM_IN_above_l92_underPF, LM_IN_below_l97_underPF, LM_IN_below_l98_underPF, LM_IN_above_l98_underPF, LM_IN_below_l99_underPF, LM_IN_above_l99_underPF)
' Arrays per lighting scenario
Dim BL_GI_GI_001: BL_GI_GI_001=Array(LM_GI_GI_001_Parts, LM_GI_GI_001_Playfield, LM_GI_GI_001_flip_l, LM_GI_GI_001_flip_l_up, LM_GI_GI_001_flip_r, LM_GI_GI_001_flipr_l, LM_GI_GI_001_flipr_l_up, LM_GI_GI_001_kick_sw17)
Dim BL_GI_GI_002: BL_GI_GI_002=Array(LM_GI_GI_002_Parts, LM_GI_GI_002_Playfield, LM_GI_GI_002_flip_l, LM_GI_GI_002_flip_r, LM_GI_GI_002_flip_r_up, LM_GI_GI_002_flipr_r, LM_GI_GI_002_flipr_r_up, LM_GI_GI_002_kick_sw24, LM_GI_GI_002_sw20)
Dim BL_GI_GI_10: BL_GI_GI_10=Array(LM_GI_GI_10_Parts, LM_GI_GI_10_Playfield, LM_GI_GI_10_flip_l, LM_GI_GI_10_flip_l_up, LM_GI_GI_10_flip_r, LM_GI_GI_10_flip_r_up, LM_GI_GI_10_flipr_l, LM_GI_GI_10_flipr_l_up, LM_GI_GI_10_kick_sw17, LM_GI_GI_10_kick_sw48, LM_GI_GI_10_sw20, LM_GI_GI_10_sw22, LM_GI_GI_10_sw23)
Dim BL_GI_GI_11: BL_GI_GI_11=Array(LM_GI_GI_11_Parts, LM_GI_GI_11_Playfield, LM_GI_GI_11_flip_l, LM_GI_GI_11_flip_l_up, LM_GI_GI_11_flip_r_up, LM_GI_GI_11_flipr_l, LM_GI_GI_11_flipr_l_up, LM_GI_GI_11_kick_sw17, LM_GI_GI_11_kick_sw48, LM_GI_GI_11_sw18, LM_GI_GI_11_sw20, LM_GI_GI_11_underPF)
Dim BL_GI_GI_12: BL_GI_GI_12=Array(LM_GI_GI_12_Parts, LM_GI_GI_12_Playfield, LM_GI_GI_12_flip_l_up, LM_GI_GI_12_flip_r, LM_GI_GI_12_flip_r_up, LM_GI_GI_12_flipr_r, LM_GI_GI_12_flipr_r_up, LM_GI_GI_12_kick_sw24, LM_GI_GI_12_kick_sw41, LM_GI_GI_12_sw22)
Dim BL_GI_GI_13: BL_GI_GI_13=Array(LM_GI_GI_13_Parts, LM_GI_GI_13_Playfield, LM_GI_GI_13_flip_l, LM_GI_GI_13_flip_l_up, LM_GI_GI_13_kick_sw17, LM_GI_GI_13_sw18, LM_GI_GI_13_sw20, LM_GI_GI_13_underPF)
Dim BL_GI_GI_14: BL_GI_GI_14=Array(LM_GI_GI_14_Parts, LM_GI_GI_14_Playfield, LM_GI_GI_14_flip_l, LM_GI_GI_14_flip_r, LM_GI_GI_14_kick_sw24, LM_GI_GI_14_kick_sw41, LM_GI_GI_14_kick_sw48, LM_GI_GI_14_sw22, LM_GI_GI_14_sw23, LM_GI_GI_14_underPF)
Dim BL_GI_GI_15: BL_GI_GI_15=Array(LM_GI_GI_15_Parts, LM_GI_GI_15_Playfield, LM_GI_GI_15_kick_sw41, LM_GI_GI_15_sw27, LM_GI_GI_15_sw28, LM_GI_GI_15_sw29, LM_GI_GI_15_sw42, LM_GI_GI_15_sw43, LM_GI_GI_15_sw44, LM_GI_GI_15_sw45)
Dim BL_GI_GI_16: BL_GI_GI_16=Array(LM_GI_GI_16_BM_PincabRails, LM_GI_GI_16_Parts, LM_GI_GI_16_Playfield, LM_GI_GI_16_kick_sw48, LM_GI_GI_16_sw27, LM_GI_GI_16_sw28, LM_GI_GI_16_sw29, LM_GI_GI_16_sw30, LM_GI_GI_16_sw45, LM_GI_GI_16_sw46, LM_GI_GI_16_sw47)
Dim BL_GI_GI_17: BL_GI_GI_17=Array(LM_GI_GI_17_BM_PincabRails, LM_GI_GI_17_Parts, LM_GI_GI_17_Playfield, LM_GI_GI_17_kick_sw41, LM_GI_GI_17_sw18, LM_GI_GI_17_sw27, LM_GI_GI_17_sw28, LM_GI_GI_17_sw29, LM_GI_GI_17_sw30, LM_GI_GI_17_sw42, LM_GI_GI_17_sw43, LM_GI_GI_17_sw44, LM_GI_GI_17_sw45, LM_GI_GI_17_sw46)
Dim BL_GI_GI_18: BL_GI_GI_18=Array(LM_GI_GI_18_Parts, LM_GI_GI_18_Playfield, LM_GI_GI_18_kick_sw48, LM_GI_GI_18_sw22, LM_GI_GI_18_sw28, LM_GI_GI_18_sw29, LM_GI_GI_18_sw30, LM_GI_GI_18_sw45, LM_GI_GI_18_sw46, LM_GI_GI_18_sw47)
Dim BL_GI_GI_19: BL_GI_GI_19=Array(LM_GI_GI_19_Layer0, LM_GI_GI_19_Parts, LM_GI_GI_19_Playfield, LM_GI_GI_19_sw19b, LM_GI_GI_19_sw26, LM_GI_GI_19_sw27, LM_GI_GI_19_sw28, LM_GI_GI_19_sw30, LM_GI_GI_19_sw37, LM_GI_GI_19_sw38, LM_GI_GI_19_sw39, LM_GI_GI_19_underPF)
Dim BL_GI_GI_2: BL_GI_GI_2=Array(LM_GI_GI_2_Layer0, LM_GI_GI_2_Parts, LM_GI_GI_2_Playfield, LM_GI_GI_2_sw21a, LM_GI_GI_2_sw30, LM_GI_GI_2_sw31, LM_GI_GI_2_sw37, LM_GI_GI_2_sw38, LM_GI_GI_2_sw39, LM_GI_GI_2_sw40a, LM_GI_GI_2_underPF)
Dim BL_GI_GI_20: BL_GI_GI_20=Array(LM_GI_GI_20_Layer0, LM_GI_GI_20_Parts, LM_GI_GI_20_Playfield, LM_GI_GI_20_sw21b, LM_GI_GI_20_sw27, LM_GI_GI_20_sw29, LM_GI_GI_20_sw30, LM_GI_GI_20_sw31, LM_GI_GI_20_sw37, LM_GI_GI_20_sw38, LM_GI_GI_20_sw39, LM_GI_GI_20_sw47, LM_GI_GI_20_underPF)
Dim BL_GI_GI_22222: BL_GI_GI_22222=Array(LM_GI_GI_22222_Parts, LM_GI_GI_22222_Playfield, LM_GI_GI_22222_flip_l, LM_GI_GI_22222_flip_l_up, LM_GI_GI_22222_flip_r, LM_GI_GI_22222_flip_r_up, LM_GI_GI_22222_flipr_l, LM_GI_GI_22222_flipr_l_up, LM_GI_GI_22222_flipr_r, LM_GI_GI_22222_flipr_r_up)
Dim BL_GI_GI_23: BL_GI_GI_23=Array(LM_GI_GI_23_BM_PincabRails, LM_GI_GI_23_Layer0, LM_GI_GI_23_Parts, LM_GI_GI_23_Playfield, LM_GI_GI_23_kick_sw7r, LM_GI_GI_23_sw19, LM_GI_GI_23_sw19a, LM_GI_GI_23_sw19b, LM_GI_GI_23_sw21a, LM_GI_GI_23_sw26, LM_GI_GI_23_sw33, LM_GI_GI_23_sw33a, LM_GI_GI_23_sw34, LM_GI_GI_23_sw40, LM_GI_GI_23_sw40a, LM_GI_GI_23_underPF)
Dim BL_GI_GI_24: BL_GI_GI_24=Array(LM_GI_GI_24_BM_PincabRails, LM_GI_GI_24_Layer0, LM_GI_GI_24_Parts, LM_GI_GI_24_Playfield, LM_GI_GI_24_sw21, LM_GI_GI_24_sw21a, LM_GI_GI_24_sw21b, LM_GI_GI_24_sw31, LM_GI_GI_24_sw33, LM_GI_GI_24_sw33a, LM_GI_GI_24_sw37, LM_GI_GI_24_sw38, LM_GI_GI_24_sw39, LM_GI_GI_24_sw40, LM_GI_GI_24_sw40a, LM_GI_GI_24_underPF)
Dim BL_GI_GI_25: BL_GI_GI_25=Array(LM_GI_GI_25_Layer0, LM_GI_GI_25_Parts, LM_GI_GI_25_Playfield, LM_GI_GI_25_kick_sw7l, LM_GI_GI_25_kick_sw7r, LM_GI_GI_25_sw19, LM_GI_GI_25_sw21, LM_GI_GI_25_sw21a, LM_GI_GI_25_sw21b, LM_GI_GI_25_sw26, LM_GI_GI_25_sw28, LM_GI_GI_25_sw30, LM_GI_GI_25_sw31, LM_GI_GI_25_sw33, LM_GI_GI_25_sw33a, LM_GI_GI_25_sw37, LM_GI_GI_25_sw38, LM_GI_GI_25_sw39, LM_GI_GI_25_sw40, LM_GI_GI_25_sw40a, LM_GI_GI_25_underPF)
Dim BL_GI_GI_26: BL_GI_GI_26=Array(LM_GI_GI_26_Layer0, LM_GI_GI_26_Parts, LM_GI_GI_26_Playfield, LM_GI_GI_26_kick_sw7l, LM_GI_GI_26_kick_sw7r, LM_GI_GI_26_sw19, LM_GI_GI_26_sw19a, LM_GI_GI_26_sw19b, LM_GI_GI_26_sw21, LM_GI_GI_26_sw26, LM_GI_GI_26_sw27, LM_GI_GI_26_sw31, LM_GI_GI_26_sw33, LM_GI_GI_26_sw33a, LM_GI_GI_26_sw34, LM_GI_GI_26_sw35, LM_GI_GI_26_sw36, LM_GI_GI_26_sw37, LM_GI_GI_26_sw38, LM_GI_GI_26_sw40, LM_GI_GI_26_sw40a, LM_GI_GI_26_underPF)
Dim BL_GI_GI_27: BL_GI_GI_27=Array(LM_GI_GI_27_BM_PincabRails, LM_GI_GI_27_Layer0, LM_GI_GI_27_Parts, LM_GI_GI_27_Playfield, LM_GI_GI_27_sw19, LM_GI_GI_27_sw19a, LM_GI_GI_27_sw19b, LM_GI_GI_27_sw26, LM_GI_GI_27_sw29, LM_GI_GI_27_sw33, LM_GI_GI_27_sw33a, LM_GI_GI_27_sw34, LM_GI_GI_27_sw35, LM_GI_GI_27_underPF)
Dim BL_GI_GI_28: BL_GI_GI_28=Array(LM_GI_GI_28_Layer0, LM_GI_GI_28_Parts, LM_GI_GI_28_Playfield, LM_GI_GI_28_sw21, LM_GI_GI_28_sw21a, LM_GI_GI_28_sw21b, LM_GI_GI_28_sw26, LM_GI_GI_28_sw28, LM_GI_GI_28_sw31, LM_GI_GI_28_sw37, LM_GI_GI_28_sw38, LM_GI_GI_28_sw39, LM_GI_GI_28_sw40, LM_GI_GI_28_sw40a, LM_GI_GI_28_underPF)
Dim BL_GI_GI_29: BL_GI_GI_29=Array(LM_GI_GI_29_Layer0, LM_GI_GI_29_Parts, LM_GI_GI_29_Playfield, LM_GI_GI_29_sw19a, LM_GI_GI_29_sw19b, LM_GI_GI_29_underPF)
Dim BL_GI_GI_3: BL_GI_GI_3=Array(LM_GI_GI_3_Layer0, LM_GI_GI_3_Parts, LM_GI_GI_3_Playfield, LM_GI_GI_3_sw19a, LM_GI_GI_3_sw26, LM_GI_GI_3_sw27, LM_GI_GI_3_sw28, LM_GI_GI_3_sw34, LM_GI_GI_3_sw37, LM_GI_GI_3_underPF)
Dim BL_GI_GI_30: BL_GI_GI_30=Array(LM_GI_GI_30_Layer0, LM_GI_GI_30_Parts, LM_GI_GI_30_Playfield, LM_GI_GI_30_sw21, LM_GI_GI_30_sw21a, LM_GI_GI_30_sw40, LM_GI_GI_30_underPF)
Dim BL_GI_GI_34: BL_GI_GI_34=Array(LM_GI_GI_34_Parts, LM_GI_GI_34_Playfield, LM_GI_GI_34_flip_l, LM_GI_GI_34_flip_l_up, LM_GI_GI_34_flip_r, LM_GI_GI_34_flip_r_up, LM_GI_GI_34_flipr_l, LM_GI_GI_34_flipr_l_up, LM_GI_GI_34_flipr_r, LM_GI_GI_34_flipr_r_up)
Dim BL_GI_GI_35: BL_GI_GI_35=Array(LM_GI_GI_35_Parts, LM_GI_GI_35_Playfield, LM_GI_GI_35_flip_l, LM_GI_GI_35_flip_l_up, LM_GI_GI_35_flip_r, LM_GI_GI_35_flip_r_up, LM_GI_GI_35_flipr_l, LM_GI_GI_35_flipr_l_up, LM_GI_GI_35_flipr_r, LM_GI_GI_35_flipr_r_up)
Dim BL_GI_GI_36: BL_GI_GI_36=Array(LM_GI_GI_36_Parts, LM_GI_GI_36_Playfield, LM_GI_GI_36_flip_l, LM_GI_GI_36_flip_l_up, LM_GI_GI_36_flip_r, LM_GI_GI_36_flip_r_up, LM_GI_GI_36_flipr_r, LM_GI_GI_36_flipr_r_up)
Dim BL_GI_GI_37: BL_GI_GI_37=Array(LM_GI_GI_37_Parts, LM_GI_GI_37_Playfield, LM_GI_GI_37_flip_l, LM_GI_GI_37_flip_l_up, LM_GI_GI_37_flip_r, LM_GI_GI_37_flip_r_up, LM_GI_GI_37_flipr_l, LM_GI_GI_37_flipr_l_up)
Dim BL_GI_GI_38: BL_GI_GI_38=Array(LM_GI_GI_38_Parts, LM_GI_GI_38_Playfield, LM_GI_GI_38_flip_l, LM_GI_GI_38_flip_l_up, LM_GI_GI_38_flip_r, LM_GI_GI_38_flip_r_up, LM_GI_GI_38_flipr_l, LM_GI_GI_38_flipr_l_up)
Dim BL_GI_GI_39: BL_GI_GI_39=Array(LM_GI_GI_39_Parts, LM_GI_GI_39_Playfield, LM_GI_GI_39_flip_l, LM_GI_GI_39_flip_l_up, LM_GI_GI_39_flip_r, LM_GI_GI_39_flip_r_up, LM_GI_GI_39_flipr_r, LM_GI_GI_39_flipr_r_up)
Dim BL_GI_GI_5: BL_GI_GI_5=Array(LM_GI_GI_5_BM_PincabRails, LM_GI_GI_5_Layer0, LM_GI_GI_5_Parts, LM_GI_GI_5_Playfield, LM_GI_GI_5_sw19, LM_GI_GI_5_sw19a, LM_GI_GI_5_sw33, LM_GI_GI_5_underPF)
Dim BL_GI_GI_6: BL_GI_GI_6=Array(LM_GI_GI_6_Layer0, LM_GI_GI_6_Parts, LM_GI_GI_6_Playfield, LM_GI_GI_6_sw21a, LM_GI_GI_6_sw21b, LM_GI_GI_6_underPF)
Dim BL_GI_GI_7: BL_GI_GI_7=Array(LM_GI_GI_7_BM_PincabRails, LM_GI_GI_7_Parts, LM_GI_GI_7_Playfield, LM_GI_GI_7_kick_sw24, LM_GI_GI_7_kick_sw48, LM_GI_GI_7_sw22, LM_GI_GI_7_sw23, LM_GI_GI_7_sw29, LM_GI_GI_7_underPF)
Dim BL_GI_GI_8: BL_GI_GI_8=Array(LM_GI_GI_8_BM_PincabRails, LM_GI_GI_8_Parts, LM_GI_GI_8_Playfield, LM_GI_GI_8_kick_sw17, LM_GI_GI_8_kick_sw41, LM_GI_GI_8_sw18, LM_GI_GI_8_sw20, LM_GI_GI_8_underPF)
Dim BL_GI_GI_9: BL_GI_GI_9=Array(LM_GI_GI_9_Parts, LM_GI_GI_9_Playfield, LM_GI_GI_9_flip_l, LM_GI_GI_9_flip_l_up, LM_GI_GI_9_flip_r, LM_GI_GI_9_flip_r_up, LM_GI_GI_9_flipr_r, LM_GI_GI_9_flipr_r_up, LM_GI_GI_9_kick_sw24, LM_GI_GI_9_kick_sw41, LM_GI_GI_9_sw18, LM_GI_GI_9_sw20, LM_GI_GI_9_sw22)
Dim BL_IN_above_l1: BL_IN_above_l1=Array(LM_IN_above_l1_Parts, LM_IN_above_l1_Playfield, LM_IN_above_l1_sw26, LM_IN_above_l1_underPF)
Dim BL_IN_above_l10: BL_IN_above_l10=Array(LM_IN_above_l10_Playfield, LM_IN_above_l10_underPF)
Dim BL_IN_above_l100: BL_IN_above_l100=Array(LM_IN_above_l100_Playfield, LM_IN_above_l100_underPF)
Dim BL_IN_above_l101: BL_IN_above_l101=Array(LM_IN_above_l101_Playfield, LM_IN_above_l101_underPF)
Dim BL_IN_above_l102: BL_IN_above_l102=Array(LM_IN_above_l102_Playfield, LM_IN_above_l102_underPF)
Dim BL_IN_above_l103: BL_IN_above_l103=Array(LM_IN_above_l103_Playfield, LM_IN_above_l103_underPF)
Dim BL_IN_above_l104: BL_IN_above_l104=Array(LM_IN_above_l104_Playfield)
Dim BL_IN_above_l105: BL_IN_above_l105=Array(LM_IN_above_l105_Playfield)
Dim BL_IN_above_l106: BL_IN_above_l106=Array(LM_IN_above_l106_Playfield, LM_IN_above_l106_underPF)
Dim BL_IN_above_l107: BL_IN_above_l107=Array(LM_IN_above_l107_Playfield)
Dim BL_IN_above_l108: BL_IN_above_l108=Array(LM_IN_above_l108_Playfield)
Dim BL_IN_above_l113: BL_IN_above_l113=Array(LM_IN_above_l113_Playfield)
Dim BL_IN_above_l114: BL_IN_above_l114=Array(LM_IN_above_l114_Playfield)
Dim BL_IN_above_l115: BL_IN_above_l115=Array(LM_IN_above_l115_Playfield)
Dim BL_IN_above_l116: BL_IN_above_l116=Array(LM_IN_above_l116_Playfield)
Dim BL_IN_above_l117: BL_IN_above_l117=Array(LM_IN_above_l117_Playfield)
Dim BL_IN_above_l118: BL_IN_above_l118=Array(LM_IN_above_l118_Playfield)
Dim BL_IN_above_l119: BL_IN_above_l119=Array(LM_IN_above_l119_Playfield)
Dim BL_IN_above_l120: BL_IN_above_l120=Array(LM_IN_above_l120_Playfield)
Dim BL_IN_above_l121: BL_IN_above_l121=Array(LM_IN_above_l121_Playfield)
Dim BL_IN_above_l122: BL_IN_above_l122=Array(LM_IN_above_l122_Playfield)
Dim BL_IN_above_l123: BL_IN_above_l123=Array(LM_IN_above_l123_Playfield)
Dim BL_IN_above_l124: BL_IN_above_l124=Array(LM_IN_above_l124_Playfield)
Dim BL_IN_above_l150: BL_IN_above_l150=Array(LM_IN_above_l150_Layer0, LM_IN_above_l150_Parts, LM_IN_above_l150_Playfield, LM_IN_above_l150_sw19, LM_IN_above_l150_underPF)
Dim BL_IN_above_l151: BL_IN_above_l151=Array(LM_IN_above_l151_Layer0, LM_IN_above_l151_Parts, LM_IN_above_l151_Playfield, LM_IN_above_l151_sw19a, LM_IN_above_l151_underPF)
Dim BL_IN_above_l152: BL_IN_above_l152=Array(LM_IN_above_l152_Layer0, LM_IN_above_l152_Parts, LM_IN_above_l152_Playfield, LM_IN_above_l152_sw19b, LM_IN_above_l152_underPF)
Dim BL_IN_above_l153: BL_IN_above_l153=Array(LM_IN_above_l153_Layer0, LM_IN_above_l153_Parts, LM_IN_above_l153_Playfield, LM_IN_above_l153_sw33, LM_IN_above_l153_underPF)
Dim BL_IN_above_l154: BL_IN_above_l154=Array(LM_IN_above_l154_Layer0, LM_IN_above_l154_Parts, LM_IN_above_l154_Playfield, LM_IN_above_l154_sw33a, LM_IN_above_l154_underPF)
Dim BL_IN_above_l155: BL_IN_above_l155=Array(LM_IN_above_l155_Layer0, LM_IN_above_l155_Parts, LM_IN_above_l155_Playfield, LM_IN_above_l155_sw21, LM_IN_above_l155_underPF)
Dim BL_IN_above_l156: BL_IN_above_l156=Array(LM_IN_above_l156_Layer0, LM_IN_above_l156_Parts, LM_IN_above_l156_Playfield, LM_IN_above_l156_sw21a, LM_IN_above_l156_underPF)
Dim BL_IN_above_l157: BL_IN_above_l157=Array(LM_IN_above_l157_Layer0, LM_IN_above_l157_Parts, LM_IN_above_l157_Playfield, LM_IN_above_l157_sw21b, LM_IN_above_l157_underPF)
Dim BL_IN_above_l158: BL_IN_above_l158=Array(LM_IN_above_l158_Layer0, LM_IN_above_l158_Parts, LM_IN_above_l158_Playfield, LM_IN_above_l158_sw37, LM_IN_above_l158_sw38, LM_IN_above_l158_sw40a, LM_IN_above_l158_underPF)
Dim BL_IN_above_l159: BL_IN_above_l159=Array(LM_IN_above_l159_Layer0, LM_IN_above_l159_Parts, LM_IN_above_l159_Playfield, LM_IN_above_l159_sw38, LM_IN_above_l159_sw39, LM_IN_above_l159_sw40, LM_IN_above_l159_underPF)
Dim BL_IN_above_l17: BL_IN_above_l17=Array(LM_IN_above_l17_Parts, LM_IN_above_l17_Playfield, LM_IN_above_l17_sw26, LM_IN_above_l17_underPF)
Dim BL_IN_above_l18: BL_IN_above_l18=Array(LM_IN_above_l18_Playfield, LM_IN_above_l18_sw28, LM_IN_above_l18_underPF)
Dim BL_IN_above_l19: BL_IN_above_l19=Array(LM_IN_above_l19_Playfield, LM_IN_above_l19_underPF)
Dim BL_IN_above_l2: BL_IN_above_l2=Array(LM_IN_above_l2_Playfield, LM_IN_above_l2_sw26, LM_IN_above_l2_sw27, LM_IN_above_l2_underPF)
Dim BL_IN_above_l21: BL_IN_above_l21=Array(LM_IN_above_l21_Parts, LM_IN_above_l21_Playfield, LM_IN_above_l21_underPF)
Dim BL_IN_above_l22: BL_IN_above_l22=Array(LM_IN_above_l22_Playfield, LM_IN_above_l22_underPF)
Dim BL_IN_above_l23: BL_IN_above_l23=Array(LM_IN_above_l23_Playfield, LM_IN_above_l23_flip_l_up, LM_IN_above_l23_flip_r_up, LM_IN_above_l23_flipr_l_up)
Dim BL_IN_above_l24: BL_IN_above_l24=Array(LM_IN_above_l24_Playfield)
Dim BL_IN_above_l25: BL_IN_above_l25=Array(LM_IN_above_l25_Playfield, LM_IN_above_l25_underPF)
Dim BL_IN_above_l26: BL_IN_above_l26=Array(LM_IN_above_l26_Parts, LM_IN_above_l26_Playfield, LM_IN_above_l26_underPF)
Dim BL_IN_above_l3: BL_IN_above_l3=Array(LM_IN_above_l3_Playfield, LM_IN_above_l3_underPF)
Dim BL_IN_above_l33: BL_IN_above_l33=Array(LM_IN_above_l33_Parts, LM_IN_above_l33_Playfield, LM_IN_above_l33_sw31, LM_IN_above_l33_underPF)
Dim BL_IN_above_l34: BL_IN_above_l34=Array(LM_IN_above_l34_Playfield, LM_IN_above_l34_sw29, LM_IN_above_l34_underPF)
Dim BL_IN_above_l35: BL_IN_above_l35=Array(LM_IN_above_l35_Playfield, LM_IN_above_l35_underPF)
Dim BL_IN_above_l37: BL_IN_above_l37=Array(LM_IN_above_l37_Playfield, LM_IN_above_l37_sw31, LM_IN_above_l37_underPF)
Dim BL_IN_above_l38: BL_IN_above_l38=Array(LM_IN_above_l38_Playfield, LM_IN_above_l38_underPF)
Dim BL_IN_above_l39: BL_IN_above_l39=Array(LM_IN_above_l39_Parts, LM_IN_above_l39_Playfield, LM_IN_above_l39_flip_l_up, LM_IN_above_l39_flip_r_up)
Dim BL_IN_above_l40: BL_IN_above_l40=Array(LM_IN_above_l40_Playfield)
Dim BL_IN_above_l41: BL_IN_above_l41=Array(LM_IN_above_l41_Playfield, LM_IN_above_l41_underPF)
Dim BL_IN_above_l49: BL_IN_above_l49=Array(LM_IN_above_l49_Layer0, LM_IN_above_l49_Parts, LM_IN_above_l49_Playfield, LM_IN_above_l49_sw30, LM_IN_above_l49_sw31, LM_IN_above_l49_sw39, LM_IN_above_l49_underPF)
Dim BL_IN_above_l5: BL_IN_above_l5=Array(LM_IN_above_l5_Playfield, LM_IN_above_l5_underPF)
Dim BL_IN_above_l50: BL_IN_above_l50=Array(LM_IN_above_l50_Parts, LM_IN_above_l50_Playfield, LM_IN_above_l50_sw30, LM_IN_above_l50_underPF)
Dim BL_IN_above_l51: BL_IN_above_l51=Array(LM_IN_above_l51_Playfield, LM_IN_above_l51_underPF)
Dim BL_IN_above_l53: BL_IN_above_l53=Array(LM_IN_above_l53_Playfield, LM_IN_above_l53_underPF)
Dim BL_IN_above_l54: BL_IN_above_l54=Array(LM_IN_above_l54_Playfield, LM_IN_above_l54_underPF)
Dim BL_IN_above_l55: BL_IN_above_l55=Array(LM_IN_above_l55_Playfield, LM_IN_above_l55_flip_l_up, LM_IN_above_l55_flip_r_up, LM_IN_above_l55_flipr_r_up)
Dim BL_IN_above_l56: BL_IN_above_l56=Array(LM_IN_above_l56_Playfield)
Dim BL_IN_above_l57: BL_IN_above_l57=Array(LM_IN_above_l57_Playfield, LM_IN_above_l57_sw27, LM_IN_above_l57_underPF)
Dim BL_IN_above_l58: BL_IN_above_l58=Array(LM_IN_above_l58_Parts, LM_IN_above_l58_Playfield, LM_IN_above_l58_underPF)
Dim BL_IN_above_l6: BL_IN_above_l6=Array(LM_IN_above_l6_Playfield, LM_IN_above_l6_underPF)
Dim BL_IN_above_l65: BL_IN_above_l65=Array(LM_IN_above_l65_Playfield)
Dim BL_IN_above_l66: BL_IN_above_l66=Array(LM_IN_above_l66_Playfield)
Dim BL_IN_above_l67: BL_IN_above_l67=Array(LM_IN_above_l67_Playfield)
Dim BL_IN_above_l68: BL_IN_above_l68=Array(LM_IN_above_l68_Playfield)
Dim BL_IN_above_l69: BL_IN_above_l69=Array(LM_IN_above_l69_Playfield)
Dim BL_IN_above_l7: BL_IN_above_l7=Array(LM_IN_above_l7_Playfield, LM_IN_above_l7_flipr_l_up)
Dim BL_IN_above_l70: BL_IN_above_l70=Array(LM_IN_above_l70_Playfield)
Dim BL_IN_above_l71: BL_IN_above_l71=Array(LM_IN_above_l71_Playfield)
Dim BL_IN_above_l72: BL_IN_above_l72=Array(LM_IN_above_l72_Playfield)
Dim BL_IN_above_l73: BL_IN_above_l73=Array(LM_IN_above_l73_Playfield)
Dim BL_IN_above_l74: BL_IN_above_l74=Array(LM_IN_above_l74_Playfield)
Dim BL_IN_above_l75: BL_IN_above_l75=Array(LM_IN_above_l75_Playfield)
Dim BL_IN_above_l76: BL_IN_above_l76=Array(LM_IN_above_l76_Playfield)
Dim BL_IN_above_l8: BL_IN_above_l8=Array(LM_IN_above_l8_Playfield, LM_IN_above_l8_flipr_r_up)
Dim BL_IN_above_l81: BL_IN_above_l81=Array(LM_IN_above_l81_Parts, LM_IN_above_l81_Playfield)
Dim BL_IN_above_l82: BL_IN_above_l82=Array(LM_IN_above_l82_Playfield, LM_IN_above_l82_underPF)
Dim BL_IN_above_l83: BL_IN_above_l83=Array(LM_IN_above_l83_Parts, LM_IN_above_l83_Playfield)
Dim BL_IN_above_l84: BL_IN_above_l84=Array(LM_IN_above_l84_Parts, LM_IN_above_l84_Playfield, LM_IN_above_l84_underPF)
Dim BL_IN_above_l85: BL_IN_above_l85=Array(LM_IN_above_l85_Playfield, LM_IN_above_l85_underPF)
Dim BL_IN_above_l86: BL_IN_above_l86=Array(LM_IN_above_l86_Playfield, LM_IN_above_l86_underPF)
Dim BL_IN_above_l87: BL_IN_above_l87=Array(LM_IN_above_l87_Playfield)
Dim BL_IN_above_l88: BL_IN_above_l88=Array(LM_IN_above_l88_Playfield, LM_IN_above_l88_underPF)
Dim BL_IN_above_l89: BL_IN_above_l89=Array(LM_IN_above_l89_Parts, LM_IN_above_l89_Playfield, LM_IN_above_l89_underPF)
Dim BL_IN_above_l9: BL_IN_above_l9=Array(LM_IN_above_l9_Playfield, LM_IN_above_l9_underPF)
Dim BL_IN_above_l90: BL_IN_above_l90=Array(LM_IN_above_l90_Playfield, LM_IN_above_l90_kick_sw48, LM_IN_above_l90_underPF)
Dim BL_IN_above_l91: BL_IN_above_l91=Array(LM_IN_above_l91_Playfield, LM_IN_above_l91_underPF)
Dim BL_IN_above_l92: BL_IN_above_l92=Array(LM_IN_above_l92_Parts, LM_IN_above_l92_Playfield, LM_IN_above_l92_underPF)
Dim BL_IN_above_l97: BL_IN_above_l97=Array(LM_IN_above_l97_Playfield)
Dim BL_IN_above_l98: BL_IN_above_l98=Array(LM_IN_above_l98_Playfield, LM_IN_above_l98_underPF)
Dim BL_IN_above_l99: BL_IN_above_l99=Array(LM_IN_above_l99_Playfield, LM_IN_above_l99_underPF)
Dim BL_IN_below_l1: BL_IN_below_l1=Array(LM_IN_below_l1_Layer0, LM_IN_below_l1_sw26, LM_IN_below_l1_underPF)
Dim BL_IN_below_l10: BL_IN_below_l10=Array(LM_IN_below_l10_underPF)
Dim BL_IN_below_l100: BL_IN_below_l100=Array(LM_IN_below_l100_underPF)
Dim BL_IN_below_l101: BL_IN_below_l101=Array(LM_IN_below_l101_underPF)
Dim BL_IN_below_l102: BL_IN_below_l102=Array(LM_IN_below_l102_underPF)
Dim BL_IN_below_l103: BL_IN_below_l103=Array(LM_IN_below_l103_underPF)
Dim BL_IN_below_l104: BL_IN_below_l104=Array(LM_IN_below_l104_underPF)
Dim BL_IN_below_l105: BL_IN_below_l105=Array(LM_IN_below_l105_underPF)
Dim BL_IN_below_l106: BL_IN_below_l106=Array(LM_IN_below_l106_underPF)
Dim BL_IN_below_l107: BL_IN_below_l107=Array(LM_IN_below_l107_underPF)
Dim BL_IN_below_l108: BL_IN_below_l108=Array(LM_IN_below_l108_underPF)
Dim BL_IN_below_l113: BL_IN_below_l113=Array(LM_IN_below_l113_underPF)
Dim BL_IN_below_l114: BL_IN_below_l114=Array(LM_IN_below_l114_underPF)
Dim BL_IN_below_l115: BL_IN_below_l115=Array(LM_IN_below_l115_underPF)
Dim BL_IN_below_l116: BL_IN_below_l116=Array(LM_IN_below_l116_underPF)
Dim BL_IN_below_l117: BL_IN_below_l117=Array(LM_IN_below_l117_underPF)
Dim BL_IN_below_l118: BL_IN_below_l118=Array(LM_IN_below_l118_underPF)
Dim BL_IN_below_l119: BL_IN_below_l119=Array(LM_IN_below_l119_underPF)
Dim BL_IN_below_l120: BL_IN_below_l120=Array(LM_IN_below_l120_underPF)
Dim BL_IN_below_l121: BL_IN_below_l121=Array(LM_IN_below_l121_underPF)
Dim BL_IN_below_l122: BL_IN_below_l122=Array(LM_IN_below_l122_underPF)
Dim BL_IN_below_l123: BL_IN_below_l123=Array(LM_IN_below_l123_underPF)
Dim BL_IN_below_l124: BL_IN_below_l124=Array(LM_IN_below_l124_underPF)
Dim BL_IN_below_l150: BL_IN_below_l150=Array(LM_IN_below_l150_Layer0, LM_IN_below_l150_Parts, LM_IN_below_l150_Playfield, LM_IN_below_l150_sw19, LM_IN_below_l150_sw33a, LM_IN_below_l150_underPF)
Dim BL_IN_below_l151: BL_IN_below_l151=Array(LM_IN_below_l151_Layer0, LM_IN_below_l151_Parts, LM_IN_below_l151_Playfield, LM_IN_below_l151_sw19a, LM_IN_below_l151_underPF)
Dim BL_IN_below_l152: BL_IN_below_l152=Array(LM_IN_below_l152_Layer0, LM_IN_below_l152_Parts, LM_IN_below_l152_Playfield, LM_IN_below_l152_sw19b, LM_IN_below_l152_underPF)
Dim BL_IN_below_l153: BL_IN_below_l153=Array(LM_IN_below_l153_Layer0, LM_IN_below_l153_Parts, LM_IN_below_l153_Playfield, LM_IN_below_l153_sw33, LM_IN_below_l153_underPF)
Dim BL_IN_below_l154: BL_IN_below_l154=Array(LM_IN_below_l154_Layer0, LM_IN_below_l154_Parts, LM_IN_below_l154_Playfield, LM_IN_below_l154_sw33a, LM_IN_below_l154_underPF)
Dim BL_IN_below_l155: BL_IN_below_l155=Array(LM_IN_below_l155_Layer0, LM_IN_below_l155_Parts, LM_IN_below_l155_Playfield, LM_IN_below_l155_sw21, LM_IN_below_l155_sw40a, LM_IN_below_l155_underPF)
Dim BL_IN_below_l156: BL_IN_below_l156=Array(LM_IN_below_l156_Layer0, LM_IN_below_l156_Parts, LM_IN_below_l156_Playfield, LM_IN_below_l156_sw21a, LM_IN_below_l156_underPF)
Dim BL_IN_below_l157: BL_IN_below_l157=Array(LM_IN_below_l157_Layer0, LM_IN_below_l157_Parts, LM_IN_below_l157_Playfield, LM_IN_below_l157_sw21b, LM_IN_below_l157_underPF)
Dim BL_IN_below_l158: BL_IN_below_l158=Array(LM_IN_below_l158_Layer0, LM_IN_below_l158_Parts, LM_IN_below_l158_Playfield, LM_IN_below_l158_sw37, LM_IN_below_l158_sw38, LM_IN_below_l158_sw39, LM_IN_below_l158_sw40a, LM_IN_below_l158_underPF)
Dim BL_IN_below_l159: BL_IN_below_l159=Array(LM_IN_below_l159_Layer0, LM_IN_below_l159_Parts, LM_IN_below_l159_Playfield, LM_IN_below_l159_sw21a, LM_IN_below_l159_sw37, LM_IN_below_l159_sw38, LM_IN_below_l159_sw39, LM_IN_below_l159_sw40, LM_IN_below_l159_underPF)
Dim BL_IN_below_l17: BL_IN_below_l17=Array(LM_IN_below_l17_Parts, LM_IN_below_l17_underPF)
Dim BL_IN_below_l18: BL_IN_below_l18=Array(LM_IN_below_l18_sw26, LM_IN_below_l18_sw28, LM_IN_below_l18_underPF)
Dim BL_IN_below_l19: BL_IN_below_l19=Array(LM_IN_below_l19_Layer0, LM_IN_below_l19_Parts, LM_IN_below_l19_sw36, LM_IN_below_l19_sw37, LM_IN_below_l19_underPF)
Dim BL_IN_below_l2: BL_IN_below_l2=Array(LM_IN_below_l2_Parts, LM_IN_below_l2_sw26, LM_IN_below_l2_sw27, LM_IN_below_l2_underPF)
Dim BL_IN_below_l20: BL_IN_below_l20=Array(LM_IN_below_l20_Layer0, LM_IN_below_l20_Parts, LM_IN_below_l20_Playfield, LM_IN_below_l20_kick_sw7l, LM_IN_below_l20_sw19, LM_IN_below_l20_sw19a, LM_IN_below_l20_sw19b, LM_IN_below_l20_sw21, LM_IN_below_l20_sw26, LM_IN_below_l20_sw27, LM_IN_below_l20_sw28, LM_IN_below_l20_sw29, LM_IN_below_l20_sw30, LM_IN_below_l20_sw31, LM_IN_below_l20_sw33, LM_IN_below_l20_sw33a, LM_IN_below_l20_sw34, LM_IN_below_l20_sw35, LM_IN_below_l20_sw36, LM_IN_below_l20_sw37, LM_IN_below_l20_sw40, LM_IN_below_l20_sw40a, LM_IN_below_l20_underPF)
Dim BL_IN_below_l21: BL_IN_below_l21=Array(LM_IN_below_l21_Parts, LM_IN_below_l21_sw26, LM_IN_below_l21_underPF)
Dim BL_IN_below_l22: BL_IN_below_l22=Array(LM_IN_below_l22_Parts, LM_IN_below_l22_underPF)
Dim BL_IN_below_l23: BL_IN_below_l23=Array(LM_IN_below_l23_flip_l_up, LM_IN_below_l23_underPF)
Dim BL_IN_below_l24: BL_IN_below_l24=Array(LM_IN_below_l24_underPF)
Dim BL_IN_below_l25: BL_IN_below_l25=Array(LM_IN_below_l25_underPF)
Dim BL_IN_below_l26: BL_IN_below_l26=Array(LM_IN_below_l26_Parts, LM_IN_below_l26_underPF)
Dim BL_IN_below_l3: BL_IN_below_l3=Array(LM_IN_below_l3_underPF)
Dim BL_IN_below_l33: BL_IN_below_l33=Array(LM_IN_below_l33_Layer0, LM_IN_below_l33_Parts, LM_IN_below_l33_underPF)
Dim BL_IN_below_l34: BL_IN_below_l34=Array(LM_IN_below_l34_sw29, LM_IN_below_l34_sw31, LM_IN_below_l34_underPF)
Dim BL_IN_below_l35: BL_IN_below_l35=Array(LM_IN_below_l35_underPF)
Dim BL_IN_below_l36: BL_IN_below_l36=Array(LM_IN_below_l36_Layer0, LM_IN_below_l36_Parts, LM_IN_below_l36_Playfield, LM_IN_below_l36_kick_sw7l, LM_IN_below_l36_sw19, LM_IN_below_l36_sw19a, LM_IN_below_l36_sw21, LM_IN_below_l36_sw21a, LM_IN_below_l36_sw21b, LM_IN_below_l36_sw26, LM_IN_below_l36_sw27, LM_IN_below_l36_sw28, LM_IN_below_l36_sw29, LM_IN_below_l36_sw30, LM_IN_below_l36_sw31, LM_IN_below_l36_sw33, LM_IN_below_l36_sw33a, LM_IN_below_l36_sw37, LM_IN_below_l36_sw38, LM_IN_below_l36_sw39, LM_IN_below_l36_sw40, LM_IN_below_l36_sw40a, LM_IN_below_l36_underPF)
Dim BL_IN_below_l37: BL_IN_below_l37=Array(LM_IN_below_l37_Parts, LM_IN_below_l37_sw31, LM_IN_below_l37_underPF)
Dim BL_IN_below_l38: BL_IN_below_l38=Array(LM_IN_below_l38_BM_PincabRails, LM_IN_below_l38_Parts, LM_IN_below_l38_underPF)
Dim BL_IN_below_l39: BL_IN_below_l39=Array(LM_IN_below_l39_underPF)
Dim BL_IN_below_l4: BL_IN_below_l4=Array(LM_IN_below_l4_Parts, LM_IN_below_l4_Playfield, LM_IN_below_l4_sw27, LM_IN_below_l4_sw28, LM_IN_below_l4_sw29, LM_IN_below_l4_sw30, LM_IN_below_l4_sw37, LM_IN_below_l4_sw40a, LM_IN_below_l4_sw42, LM_IN_below_l4_sw43, LM_IN_below_l4_sw44, LM_IN_below_l4_sw47)
Dim BL_IN_below_l40: BL_IN_below_l40=Array(LM_IN_below_l40_underPF)
Dim BL_IN_below_l41: BL_IN_below_l41=Array(LM_IN_below_l41_underPF)
Dim BL_IN_below_l42: BL_IN_below_l42=Array(LM_IN_below_l42_Parts, LM_IN_below_l42_Playfield, LM_IN_below_l42_flip_l, LM_IN_below_l42_flip_r, LM_IN_below_l42_flip_r_up, LM_IN_below_l42_flipr_r, LM_IN_below_l42_flipr_r_up)
Dim BL_IN_below_l49: BL_IN_below_l49=Array(LM_IN_below_l49_sw31, LM_IN_below_l49_underPF)
Dim BL_IN_below_l5: BL_IN_below_l5=Array(LM_IN_below_l5_Parts, LM_IN_below_l5_sw26, LM_IN_below_l5_underPF)
Dim BL_IN_below_l50: BL_IN_below_l50=Array(LM_IN_below_l50_Parts, LM_IN_below_l50_sw29, LM_IN_below_l50_sw30, LM_IN_below_l50_sw31, LM_IN_below_l50_underPF)
Dim BL_IN_below_l51: BL_IN_below_l51=Array(LM_IN_below_l51_underPF)
Dim BL_IN_below_l52: BL_IN_below_l52=Array(LM_IN_below_l52_Parts, LM_IN_below_l52_Playfield, LM_IN_below_l52_sw27, LM_IN_below_l52_sw28, LM_IN_below_l52_sw29, LM_IN_below_l52_sw30, LM_IN_below_l52_sw37, LM_IN_below_l52_sw38, LM_IN_below_l52_sw45, LM_IN_below_l52_sw46, LM_IN_below_l52_sw47, LM_IN_below_l52_underPF)
Dim BL_IN_below_l53: BL_IN_below_l53=Array(LM_IN_below_l53_Parts, LM_IN_below_l53_sw31, LM_IN_below_l53_underPF)
Dim BL_IN_below_l54: BL_IN_below_l54=Array(LM_IN_below_l54_Parts, LM_IN_below_l54_underPF)
Dim BL_IN_below_l55: BL_IN_below_l55=Array(LM_IN_below_l55_flip_r_up, LM_IN_below_l55_flipr_r_up, LM_IN_below_l55_underPF)
Dim BL_IN_below_l56: BL_IN_below_l56=Array(LM_IN_below_l56_underPF)
Dim BL_IN_below_l57: BL_IN_below_l57=Array(LM_IN_below_l57_underPF)
Dim BL_IN_below_l58: BL_IN_below_l58=Array(LM_IN_below_l58_Parts, LM_IN_below_l58_underPF)
Dim BL_IN_below_l59: BL_IN_below_l59=Array(LM_IN_below_l59_Parts, LM_IN_below_l59_Playfield, LM_IN_below_l59_flip_l, LM_IN_below_l59_flip_l_up, LM_IN_below_l59_flip_r, LM_IN_below_l59_flipr_l, LM_IN_below_l59_flipr_l_up)
Dim BL_IN_below_l6: BL_IN_below_l6=Array(LM_IN_below_l6_underPF)
Dim BL_IN_below_l65: BL_IN_below_l65=Array(LM_IN_below_l65_underPF)
Dim BL_IN_below_l66: BL_IN_below_l66=Array(LM_IN_below_l66_underPF)
Dim BL_IN_below_l67: BL_IN_below_l67=Array(LM_IN_below_l67_underPF)
Dim BL_IN_below_l68: BL_IN_below_l68=Array(LM_IN_below_l68_underPF)
Dim BL_IN_below_l69: BL_IN_below_l69=Array(LM_IN_below_l69_underPF)
Dim BL_IN_below_l7: BL_IN_below_l7=Array(LM_IN_below_l7_underPF)
Dim BL_IN_below_l70: BL_IN_below_l70=Array(LM_IN_below_l70_underPF)
Dim BL_IN_below_l71: BL_IN_below_l71=Array(LM_IN_below_l71_underPF)
Dim BL_IN_below_l72: BL_IN_below_l72=Array(LM_IN_below_l72_underPF)
Dim BL_IN_below_l73: BL_IN_below_l73=Array(LM_IN_below_l73_underPF)
Dim BL_IN_below_l74: BL_IN_below_l74=Array(LM_IN_below_l74_underPF)
Dim BL_IN_below_l75: BL_IN_below_l75=Array(LM_IN_below_l75_underPF)
Dim BL_IN_below_l76: BL_IN_below_l76=Array(LM_IN_below_l76_underPF)
Dim BL_IN_below_l8: BL_IN_below_l8=Array(LM_IN_below_l8_underPF)
Dim BL_IN_below_l81: BL_IN_below_l81=Array(LM_IN_below_l81_underPF)
Dim BL_IN_below_l82: BL_IN_below_l82=Array(LM_IN_below_l82_underPF)
Dim BL_IN_below_l83: BL_IN_below_l83=Array(LM_IN_below_l83_underPF)
Dim BL_IN_below_l84: BL_IN_below_l84=Array(LM_IN_below_l84_underPF)
Dim BL_IN_below_l85: BL_IN_below_l85=Array(LM_IN_below_l85_underPF)
Dim BL_IN_below_l86: BL_IN_below_l86=Array(LM_IN_below_l86_underPF)
Dim BL_IN_below_l87: BL_IN_below_l87=Array(LM_IN_below_l87_underPF)
Dim BL_IN_below_l88: BL_IN_below_l88=Array(LM_IN_below_l88_underPF)
Dim BL_IN_below_l89: BL_IN_below_l89=Array(LM_IN_below_l89_underPF)
Dim BL_IN_below_l9: BL_IN_below_l9=Array(LM_IN_below_l9_underPF)
Dim BL_IN_below_l90: BL_IN_below_l90=Array(LM_IN_below_l90_underPF)
Dim BL_IN_below_l91: BL_IN_below_l91=Array(LM_IN_below_l91_underPF)
Dim BL_IN_below_l92: BL_IN_below_l92=Array(LM_IN_below_l92_underPF)
Dim BL_IN_below_l97: BL_IN_below_l97=Array(LM_IN_below_l97_underPF)
Dim BL_IN_below_l98: BL_IN_below_l98=Array(LM_IN_below_l98_underPF)
Dim BL_IN_below_l99: BL_IN_below_l99=Array(LM_IN_below_l99_underPF)
Dim BL_World: BL_World=Array(BM_BM_PincabRails, BM_Layer0, BM_Parts, BM_Playfield, BM_flip_l, BM_flip_l_up, BM_flip_r, BM_flip_r_up, BM_flipr_l, BM_flipr_l_up, BM_flipr_r, BM_flipr_r_up, BM_kick_sw17, BM_kick_sw24, BM_kick_sw41, BM_kick_sw48, BM_kick_sw7l, BM_kick_sw7r, BM_sw18, BM_sw19, BM_sw19a, BM_sw19b, BM_sw20, BM_sw21, BM_sw21a, BM_sw21b, BM_sw22, BM_sw23, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw33, BM_sw33a, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw38, BM_sw39, BM_sw40, BM_sw40a, BM_sw42, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw47, BM_underPF)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BM_PincabRails, BM_Layer0, BM_Parts, BM_Playfield, BM_flip_l, BM_flip_l_up, BM_flip_r, BM_flip_r_up, BM_flipr_l, BM_flipr_l_up, BM_flipr_r, BM_flipr_r_up, BM_kick_sw17, BM_kick_sw24, BM_kick_sw41, BM_kick_sw48, BM_kick_sw7l, BM_kick_sw7r, BM_sw18, BM_sw19, BM_sw19a, BM_sw19b, BM_sw20, BM_sw21, BM_sw21a, BM_sw21b, BM_sw22, BM_sw23, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw33, BM_sw33a, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw38, BM_sw39, BM_sw40, BM_sw40a, BM_sw42, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw47, BM_underPF)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_GI_001_Parts, LM_GI_GI_001_Playfield, LM_GI_GI_001_flip_l, LM_GI_GI_001_flip_l_up, LM_GI_GI_001_flip_r, LM_GI_GI_001_flipr_l, LM_GI_GI_001_flipr_l_up, LM_GI_GI_001_kick_sw17, LM_GI_GI_002_Parts, LM_GI_GI_002_Playfield, LM_GI_GI_002_flip_l, LM_GI_GI_002_flip_r, LM_GI_GI_002_flip_r_up, LM_GI_GI_002_flipr_r, LM_GI_GI_002_flipr_r_up, LM_GI_GI_002_kick_sw24, LM_GI_GI_002_sw20, LM_GI_GI_10_Parts, LM_GI_GI_10_Playfield, LM_GI_GI_10_flip_l, LM_GI_GI_10_flip_l_up, LM_GI_GI_10_flip_r, LM_GI_GI_10_flip_r_up, LM_GI_GI_10_flipr_l, LM_GI_GI_10_flipr_l_up, LM_GI_GI_10_kick_sw17, LM_GI_GI_10_kick_sw48, LM_GI_GI_10_sw20, LM_GI_GI_10_sw22, LM_GI_GI_10_sw23, LM_GI_GI_11_Parts, LM_GI_GI_11_Playfield, LM_GI_GI_11_flip_l, LM_GI_GI_11_flip_l_up, LM_GI_GI_11_flip_r_up, LM_GI_GI_11_flipr_l, LM_GI_GI_11_flipr_l_up, LM_GI_GI_11_kick_sw17, LM_GI_GI_11_kick_sw48, LM_GI_GI_11_sw18, LM_GI_GI_11_sw20, LM_GI_GI_11_underPF, LM_GI_GI_12_Parts, LM_GI_GI_12_Playfield, LM_GI_GI_12_flip_l_up, _
  LM_GI_GI_12_flip_r, LM_GI_GI_12_flip_r_up, LM_GI_GI_12_flipr_r, LM_GI_GI_12_flipr_r_up, LM_GI_GI_12_kick_sw24, LM_GI_GI_12_kick_sw41, LM_GI_GI_12_sw22, LM_GI_GI_13_Parts, LM_GI_GI_13_Playfield, LM_GI_GI_13_flip_l, LM_GI_GI_13_flip_l_up, LM_GI_GI_13_kick_sw17, LM_GI_GI_13_sw18, LM_GI_GI_13_sw20, LM_GI_GI_13_underPF, LM_GI_GI_14_Parts, LM_GI_GI_14_Playfield, LM_GI_GI_14_flip_l, LM_GI_GI_14_flip_r, LM_GI_GI_14_kick_sw24, LM_GI_GI_14_kick_sw41, LM_GI_GI_14_kick_sw48, LM_GI_GI_14_sw22, LM_GI_GI_14_sw23, LM_GI_GI_14_underPF, LM_GI_GI_15_Parts, LM_GI_GI_15_Playfield, LM_GI_GI_15_kick_sw41, LM_GI_GI_15_sw27, LM_GI_GI_15_sw28, LM_GI_GI_15_sw29, LM_GI_GI_15_sw42, LM_GI_GI_15_sw43, LM_GI_GI_15_sw44, LM_GI_GI_15_sw45, LM_GI_GI_16_BM_PincabRails, LM_GI_GI_16_Parts, LM_GI_GI_16_Playfield, LM_GI_GI_16_kick_sw48, LM_GI_GI_16_sw27, LM_GI_GI_16_sw28, LM_GI_GI_16_sw29, LM_GI_GI_16_sw30, LM_GI_GI_16_sw45, LM_GI_GI_16_sw46, LM_GI_GI_16_sw47, LM_GI_GI_17_BM_PincabRails, LM_GI_GI_17_Parts, LM_GI_GI_17_Playfield, _
  LM_GI_GI_17_kick_sw41, LM_GI_GI_17_sw18, LM_GI_GI_17_sw27, LM_GI_GI_17_sw28, LM_GI_GI_17_sw29, LM_GI_GI_17_sw30, LM_GI_GI_17_sw42, LM_GI_GI_17_sw43, LM_GI_GI_17_sw44, LM_GI_GI_17_sw45, LM_GI_GI_17_sw46, LM_GI_GI_18_Parts, LM_GI_GI_18_Playfield, LM_GI_GI_18_kick_sw48, LM_GI_GI_18_sw22, LM_GI_GI_18_sw28, LM_GI_GI_18_sw29, LM_GI_GI_18_sw30, LM_GI_GI_18_sw45, LM_GI_GI_18_sw46, LM_GI_GI_18_sw47, LM_GI_GI_19_Layer0, LM_GI_GI_19_Parts, LM_GI_GI_19_Playfield, LM_GI_GI_19_sw19b, LM_GI_GI_19_sw26, LM_GI_GI_19_sw27, LM_GI_GI_19_sw28, LM_GI_GI_19_sw30, LM_GI_GI_19_sw37, LM_GI_GI_19_sw38, LM_GI_GI_19_sw39, LM_GI_GI_19_underPF, LM_GI_GI_2_Layer0, LM_GI_GI_2_Parts, LM_GI_GI_2_Playfield, LM_GI_GI_2_sw21a, LM_GI_GI_2_sw30, LM_GI_GI_2_sw31, LM_GI_GI_2_sw37, LM_GI_GI_2_sw38, LM_GI_GI_2_sw39, LM_GI_GI_2_sw40a, LM_GI_GI_2_underPF, LM_GI_GI_20_Layer0, LM_GI_GI_20_Parts, LM_GI_GI_20_Playfield, LM_GI_GI_20_sw21b, LM_GI_GI_20_sw27, LM_GI_GI_20_sw29, LM_GI_GI_20_sw30, LM_GI_GI_20_sw31, LM_GI_GI_20_sw37, LM_GI_GI_20_sw38, _
  LM_GI_GI_20_sw39, LM_GI_GI_20_sw47, LM_GI_GI_20_underPF, LM_GI_GI_22222_Parts, LM_GI_GI_22222_Playfield, LM_GI_GI_22222_flip_l, LM_GI_GI_22222_flip_l_up, LM_GI_GI_22222_flip_r, LM_GI_GI_22222_flip_r_up, LM_GI_GI_22222_flipr_l, LM_GI_GI_22222_flipr_l_up, LM_GI_GI_22222_flipr_r, LM_GI_GI_22222_flipr_r_up, LM_GI_GI_23_BM_PincabRails, LM_GI_GI_23_Layer0, LM_GI_GI_23_Parts, LM_GI_GI_23_Playfield, LM_GI_GI_23_kick_sw7r, LM_GI_GI_23_sw19, LM_GI_GI_23_sw19a, LM_GI_GI_23_sw19b, LM_GI_GI_23_sw21a, LM_GI_GI_23_sw26, LM_GI_GI_23_sw33, LM_GI_GI_23_sw33a, LM_GI_GI_23_sw34, LM_GI_GI_23_sw40, LM_GI_GI_23_sw40a, LM_GI_GI_23_underPF, LM_GI_GI_24_BM_PincabRails, LM_GI_GI_24_Layer0, LM_GI_GI_24_Parts, LM_GI_GI_24_Playfield, LM_GI_GI_24_sw21, LM_GI_GI_24_sw21a, LM_GI_GI_24_sw21b, LM_GI_GI_24_sw31, LM_GI_GI_24_sw33, LM_GI_GI_24_sw33a, LM_GI_GI_24_sw37, LM_GI_GI_24_sw38, LM_GI_GI_24_sw39, LM_GI_GI_24_sw40, LM_GI_GI_24_sw40a, LM_GI_GI_24_underPF, LM_GI_GI_25_Layer0, LM_GI_GI_25_Parts, LM_GI_GI_25_Playfield, LM_GI_GI_25_kick_sw7l, _
  LM_GI_GI_25_kick_sw7r, LM_GI_GI_25_sw19, LM_GI_GI_25_sw21, LM_GI_GI_25_sw21a, LM_GI_GI_25_sw21b, LM_GI_GI_25_sw26, LM_GI_GI_25_sw28, LM_GI_GI_25_sw30, LM_GI_GI_25_sw31, LM_GI_GI_25_sw33, LM_GI_GI_25_sw33a, LM_GI_GI_25_sw37, LM_GI_GI_25_sw38, LM_GI_GI_25_sw39, LM_GI_GI_25_sw40, LM_GI_GI_25_sw40a, LM_GI_GI_25_underPF, LM_GI_GI_26_Layer0, LM_GI_GI_26_Parts, LM_GI_GI_26_Playfield, LM_GI_GI_26_kick_sw7l, LM_GI_GI_26_kick_sw7r, LM_GI_GI_26_sw19, LM_GI_GI_26_sw19a, LM_GI_GI_26_sw19b, LM_GI_GI_26_sw21, LM_GI_GI_26_sw26, LM_GI_GI_26_sw27, LM_GI_GI_26_sw31, LM_GI_GI_26_sw33, LM_GI_GI_26_sw33a, LM_GI_GI_26_sw34, LM_GI_GI_26_sw35, LM_GI_GI_26_sw36, LM_GI_GI_26_sw37, LM_GI_GI_26_sw38, LM_GI_GI_26_sw40, LM_GI_GI_26_sw40a, LM_GI_GI_26_underPF, LM_GI_GI_27_BM_PincabRails, LM_GI_GI_27_Layer0, LM_GI_GI_27_Parts, LM_GI_GI_27_Playfield, LM_GI_GI_27_sw19, LM_GI_GI_27_sw19a, LM_GI_GI_27_sw19b, LM_GI_GI_27_sw26, LM_GI_GI_27_sw29, LM_GI_GI_27_sw33, LM_GI_GI_27_sw33a, LM_GI_GI_27_sw34, LM_GI_GI_27_sw35, LM_GI_GI_27_underPF, _
  LM_GI_GI_28_Layer0, LM_GI_GI_28_Parts, LM_GI_GI_28_Playfield, LM_GI_GI_28_sw21, LM_GI_GI_28_sw21a, LM_GI_GI_28_sw21b, LM_GI_GI_28_sw26, LM_GI_GI_28_sw28, LM_GI_GI_28_sw31, LM_GI_GI_28_sw37, LM_GI_GI_28_sw38, LM_GI_GI_28_sw39, LM_GI_GI_28_sw40, LM_GI_GI_28_sw40a, LM_GI_GI_28_underPF, LM_GI_GI_29_Layer0, LM_GI_GI_29_Parts, LM_GI_GI_29_Playfield, LM_GI_GI_29_sw19a, LM_GI_GI_29_sw19b, LM_GI_GI_29_underPF, LM_GI_GI_3_Layer0, LM_GI_GI_3_Parts, LM_GI_GI_3_Playfield, LM_GI_GI_3_sw19a, LM_GI_GI_3_sw26, LM_GI_GI_3_sw27, LM_GI_GI_3_sw28, LM_GI_GI_3_sw34, LM_GI_GI_3_sw37, LM_GI_GI_3_underPF, LM_GI_GI_30_Layer0, LM_GI_GI_30_Parts, LM_GI_GI_30_Playfield, LM_GI_GI_30_sw21, LM_GI_GI_30_sw21a, LM_GI_GI_30_sw40, LM_GI_GI_30_underPF, LM_GI_GI_34_Parts, LM_GI_GI_34_Playfield, LM_GI_GI_34_flip_l, LM_GI_GI_34_flip_l_up, LM_GI_GI_34_flip_r, LM_GI_GI_34_flip_r_up, LM_GI_GI_34_flipr_l, LM_GI_GI_34_flipr_l_up, LM_GI_GI_34_flipr_r, LM_GI_GI_34_flipr_r_up, LM_GI_GI_35_Parts, LM_GI_GI_35_Playfield, LM_GI_GI_35_flip_l, _
  LM_GI_GI_35_flip_l_up, LM_GI_GI_35_flip_r, LM_GI_GI_35_flip_r_up, LM_GI_GI_35_flipr_l, LM_GI_GI_35_flipr_l_up, LM_GI_GI_35_flipr_r, LM_GI_GI_35_flipr_r_up, LM_GI_GI_36_Parts, LM_GI_GI_36_Playfield, LM_GI_GI_36_flip_l, LM_GI_GI_36_flip_l_up, LM_GI_GI_36_flip_r, LM_GI_GI_36_flip_r_up, LM_GI_GI_36_flipr_r, LM_GI_GI_36_flipr_r_up, LM_GI_GI_37_Parts, LM_GI_GI_37_Playfield, LM_GI_GI_37_flip_l, LM_GI_GI_37_flip_l_up, LM_GI_GI_37_flip_r, LM_GI_GI_37_flip_r_up, LM_GI_GI_37_flipr_l, LM_GI_GI_37_flipr_l_up, LM_GI_GI_38_Parts, LM_GI_GI_38_Playfield, LM_GI_GI_38_flip_l, LM_GI_GI_38_flip_l_up, LM_GI_GI_38_flip_r, LM_GI_GI_38_flip_r_up, LM_GI_GI_38_flipr_l, LM_GI_GI_38_flipr_l_up, LM_GI_GI_39_Parts, LM_GI_GI_39_Playfield, LM_GI_GI_39_flip_l, LM_GI_GI_39_flip_l_up, LM_GI_GI_39_flip_r, LM_GI_GI_39_flip_r_up, LM_GI_GI_39_flipr_r, LM_GI_GI_39_flipr_r_up, LM_GI_GI_5_BM_PincabRails, LM_GI_GI_5_Layer0, LM_GI_GI_5_Parts, LM_GI_GI_5_Playfield, LM_GI_GI_5_sw19, LM_GI_GI_5_sw19a, LM_GI_GI_5_sw33, LM_GI_GI_5_underPF, LM_GI_GI_6_Layer0, _
  LM_GI_GI_6_Parts, LM_GI_GI_6_Playfield, LM_GI_GI_6_sw21a, LM_GI_GI_6_sw21b, LM_GI_GI_6_underPF, LM_GI_GI_7_BM_PincabRails, LM_GI_GI_7_Parts, LM_GI_GI_7_Playfield, LM_GI_GI_7_kick_sw24, LM_GI_GI_7_kick_sw48, LM_GI_GI_7_sw22, LM_GI_GI_7_sw23, LM_GI_GI_7_sw29, LM_GI_GI_7_underPF, LM_GI_GI_8_BM_PincabRails, LM_GI_GI_8_Parts, LM_GI_GI_8_Playfield, LM_GI_GI_8_kick_sw17, LM_GI_GI_8_kick_sw41, LM_GI_GI_8_sw18, LM_GI_GI_8_sw20, LM_GI_GI_8_underPF, LM_GI_GI_9_Parts, LM_GI_GI_9_Playfield, LM_GI_GI_9_flip_l, LM_GI_GI_9_flip_l_up, LM_GI_GI_9_flip_r, LM_GI_GI_9_flip_r_up, LM_GI_GI_9_flipr_r, LM_GI_GI_9_flipr_r_up, LM_GI_GI_9_kick_sw24, LM_GI_GI_9_kick_sw41, LM_GI_GI_9_sw18, LM_GI_GI_9_sw20, LM_GI_GI_9_sw22, LM_IN_above_l1_Parts, LM_IN_above_l1_Playfield, LM_IN_above_l1_sw26, LM_IN_above_l1_underPF, LM_IN_above_l10_Playfield, LM_IN_above_l10_underPF, LM_IN_above_l100_Playfield, LM_IN_above_l100_underPF, LM_IN_above_l101_Playfield, LM_IN_above_l101_underPF, LM_IN_above_l102_Playfield, LM_IN_above_l102_underPF, _
  LM_IN_above_l103_Playfield, LM_IN_above_l103_underPF, LM_IN_above_l104_Playfield, LM_IN_above_l105_Playfield, LM_IN_above_l106_Playfield, LM_IN_above_l106_underPF, LM_IN_above_l107_Playfield, LM_IN_above_l108_Playfield, LM_IN_above_l113_Playfield, LM_IN_above_l114_Playfield, LM_IN_above_l115_Playfield, LM_IN_above_l116_Playfield, LM_IN_above_l117_Playfield, LM_IN_above_l118_Playfield, LM_IN_above_l119_Playfield, LM_IN_above_l120_Playfield, LM_IN_above_l121_Playfield, LM_IN_above_l122_Playfield, LM_IN_above_l123_Playfield, LM_IN_above_l124_Playfield, LM_IN_above_l150_Layer0, LM_IN_above_l150_Parts, LM_IN_above_l150_Playfield, LM_IN_above_l150_sw19, LM_IN_above_l150_underPF, LM_IN_above_l151_Layer0, LM_IN_above_l151_Parts, LM_IN_above_l151_Playfield, LM_IN_above_l151_sw19a, LM_IN_above_l151_underPF, LM_IN_above_l152_Layer0, LM_IN_above_l152_Parts, LM_IN_above_l152_Playfield, LM_IN_above_l152_sw19b, LM_IN_above_l152_underPF, LM_IN_above_l153_Layer0, LM_IN_above_l153_Parts, LM_IN_above_l153_Playfield, _
  LM_IN_above_l153_sw33, LM_IN_above_l153_underPF, LM_IN_above_l154_Layer0, LM_IN_above_l154_Parts, LM_IN_above_l154_Playfield, LM_IN_above_l154_sw33a, LM_IN_above_l154_underPF, LM_IN_above_l155_Layer0, LM_IN_above_l155_Parts, LM_IN_above_l155_Playfield, LM_IN_above_l155_sw21, LM_IN_above_l155_underPF, LM_IN_above_l156_Layer0, LM_IN_above_l156_Parts, LM_IN_above_l156_Playfield, LM_IN_above_l156_sw21a, LM_IN_above_l156_underPF, LM_IN_above_l157_Layer0, LM_IN_above_l157_Parts, LM_IN_above_l157_Playfield, LM_IN_above_l157_sw21b, LM_IN_above_l157_underPF, LM_IN_above_l158_Layer0, LM_IN_above_l158_Parts, LM_IN_above_l158_Playfield, LM_IN_above_l158_sw37, LM_IN_above_l158_sw38, LM_IN_above_l158_sw40a, LM_IN_above_l158_underPF, LM_IN_above_l159_Layer0, LM_IN_above_l159_Parts, LM_IN_above_l159_Playfield, LM_IN_above_l159_sw38, LM_IN_above_l159_sw39, LM_IN_above_l159_sw40, LM_IN_above_l159_underPF, LM_IN_above_l17_Parts, LM_IN_above_l17_Playfield, LM_IN_above_l17_sw26, LM_IN_above_l17_underPF, LM_IN_above_l18_Playfield, _
  LM_IN_above_l18_sw28, LM_IN_above_l18_underPF, LM_IN_above_l19_Playfield, LM_IN_above_l19_underPF, LM_IN_above_l2_Playfield, LM_IN_above_l2_sw26, LM_IN_above_l2_sw27, LM_IN_above_l2_underPF, LM_IN_above_l21_Parts, LM_IN_above_l21_Playfield, LM_IN_above_l21_underPF, LM_IN_above_l22_Playfield, LM_IN_above_l22_underPF, LM_IN_above_l23_Playfield, LM_IN_above_l23_flip_l_up, LM_IN_above_l23_flip_r_up, LM_IN_above_l23_flipr_l_up, LM_IN_above_l24_Playfield, LM_IN_above_l25_Playfield, LM_IN_above_l25_underPF, LM_IN_above_l26_Parts, LM_IN_above_l26_Playfield, LM_IN_above_l26_underPF, LM_IN_above_l3_Playfield, LM_IN_above_l3_underPF, LM_IN_above_l33_Parts, LM_IN_above_l33_Playfield, LM_IN_above_l33_sw31, LM_IN_above_l33_underPF, LM_IN_above_l34_Playfield, LM_IN_above_l34_sw29, LM_IN_above_l34_underPF, LM_IN_above_l35_Playfield, LM_IN_above_l35_underPF, LM_IN_above_l37_Playfield, LM_IN_above_l37_sw31, LM_IN_above_l37_underPF, LM_IN_above_l38_Playfield, LM_IN_above_l38_underPF, LM_IN_above_l39_Parts, _
  LM_IN_above_l39_Playfield, LM_IN_above_l39_flip_l_up, LM_IN_above_l39_flip_r_up, LM_IN_above_l40_Playfield, LM_IN_above_l41_Playfield, LM_IN_above_l41_underPF, LM_IN_above_l49_Layer0, LM_IN_above_l49_Parts, LM_IN_above_l49_Playfield, LM_IN_above_l49_sw30, LM_IN_above_l49_sw31, LM_IN_above_l49_sw39, LM_IN_above_l49_underPF, LM_IN_above_l5_Playfield, LM_IN_above_l5_underPF, LM_IN_above_l50_Parts, LM_IN_above_l50_Playfield, LM_IN_above_l50_sw30, LM_IN_above_l50_underPF, LM_IN_above_l51_Playfield, LM_IN_above_l51_underPF, LM_IN_above_l53_Playfield, LM_IN_above_l53_underPF, LM_IN_above_l54_Playfield, LM_IN_above_l54_underPF, LM_IN_above_l55_Playfield, LM_IN_above_l55_flip_l_up, LM_IN_above_l55_flip_r_up, LM_IN_above_l55_flipr_r_up, LM_IN_above_l56_Playfield, LM_IN_above_l57_Playfield, LM_IN_above_l57_sw27, LM_IN_above_l57_underPF, LM_IN_above_l58_Parts, LM_IN_above_l58_Playfield, LM_IN_above_l58_underPF, LM_IN_above_l6_Playfield, LM_IN_above_l6_underPF, LM_IN_above_l65_Playfield, LM_IN_above_l66_Playfield, _
  LM_IN_above_l67_Playfield, LM_IN_above_l68_Playfield, LM_IN_above_l69_Playfield, LM_IN_above_l7_Playfield, LM_IN_above_l7_flipr_l_up, LM_IN_above_l70_Playfield, LM_IN_above_l71_Playfield, LM_IN_above_l72_Playfield, LM_IN_above_l73_Playfield, LM_IN_above_l74_Playfield, LM_IN_above_l75_Playfield, LM_IN_above_l76_Playfield, LM_IN_above_l8_Playfield, LM_IN_above_l8_flipr_r_up, LM_IN_above_l81_Parts, LM_IN_above_l81_Playfield, LM_IN_above_l82_Playfield, LM_IN_above_l82_underPF, LM_IN_above_l83_Parts, LM_IN_above_l83_Playfield, LM_IN_above_l84_Parts, LM_IN_above_l84_Playfield, LM_IN_above_l84_underPF, LM_IN_above_l85_Playfield, LM_IN_above_l85_underPF, LM_IN_above_l86_Playfield, LM_IN_above_l86_underPF, LM_IN_above_l87_Playfield, LM_IN_above_l88_Playfield, LM_IN_above_l88_underPF, LM_IN_above_l89_Parts, LM_IN_above_l89_Playfield, LM_IN_above_l89_underPF, LM_IN_above_l9_Playfield, LM_IN_above_l9_underPF, LM_IN_above_l90_Playfield, LM_IN_above_l90_kick_sw48, LM_IN_above_l90_underPF, LM_IN_above_l91_Playfield, _
  LM_IN_above_l91_underPF, LM_IN_above_l92_Parts, LM_IN_above_l92_Playfield, LM_IN_above_l92_underPF, LM_IN_above_l97_Playfield, LM_IN_above_l98_Playfield, LM_IN_above_l98_underPF, LM_IN_above_l99_Playfield, LM_IN_above_l99_underPF, LM_IN_below_l1_Layer0, LM_IN_below_l1_sw26, LM_IN_below_l1_underPF, LM_IN_below_l10_underPF, LM_IN_below_l100_underPF, LM_IN_below_l101_underPF, LM_IN_below_l102_underPF, LM_IN_below_l103_underPF, LM_IN_below_l104_underPF, LM_IN_below_l105_underPF, LM_IN_below_l106_underPF, LM_IN_below_l107_underPF, LM_IN_below_l108_underPF, LM_IN_below_l113_underPF, LM_IN_below_l114_underPF, LM_IN_below_l115_underPF, LM_IN_below_l116_underPF, LM_IN_below_l117_underPF, LM_IN_below_l118_underPF, LM_IN_below_l119_underPF, LM_IN_below_l120_underPF, LM_IN_below_l121_underPF, LM_IN_below_l122_underPF, LM_IN_below_l123_underPF, LM_IN_below_l124_underPF, LM_IN_below_l150_Layer0, LM_IN_below_l150_Parts, LM_IN_below_l150_Playfield, LM_IN_below_l150_sw19, LM_IN_below_l150_sw33a, LM_IN_below_l150_underPF, _
  LM_IN_below_l151_Layer0, LM_IN_below_l151_Parts, LM_IN_below_l151_Playfield, LM_IN_below_l151_sw19a, LM_IN_below_l151_underPF, LM_IN_below_l152_Layer0, LM_IN_below_l152_Parts, LM_IN_below_l152_Playfield, LM_IN_below_l152_sw19b, LM_IN_below_l152_underPF, LM_IN_below_l153_Layer0, LM_IN_below_l153_Parts, LM_IN_below_l153_Playfield, LM_IN_below_l153_sw33, LM_IN_below_l153_underPF, LM_IN_below_l154_Layer0, LM_IN_below_l154_Parts, LM_IN_below_l154_Playfield, LM_IN_below_l154_sw33a, LM_IN_below_l154_underPF, LM_IN_below_l155_Layer0, LM_IN_below_l155_Parts, LM_IN_below_l155_Playfield, LM_IN_below_l155_sw21, LM_IN_below_l155_sw40a, LM_IN_below_l155_underPF, LM_IN_below_l156_Layer0, LM_IN_below_l156_Parts, LM_IN_below_l156_Playfield, LM_IN_below_l156_sw21a, LM_IN_below_l156_underPF, LM_IN_below_l157_Layer0, LM_IN_below_l157_Parts, LM_IN_below_l157_Playfield, LM_IN_below_l157_sw21b, LM_IN_below_l157_underPF, LM_IN_below_l158_Layer0, LM_IN_below_l158_Parts, LM_IN_below_l158_Playfield, LM_IN_below_l158_sw37, _
  LM_IN_below_l158_sw38, LM_IN_below_l158_sw39, LM_IN_below_l158_sw40a, LM_IN_below_l158_underPF, LM_IN_below_l159_Layer0, LM_IN_below_l159_Parts, LM_IN_below_l159_Playfield, LM_IN_below_l159_sw21a, LM_IN_below_l159_sw37, LM_IN_below_l159_sw38, LM_IN_below_l159_sw39, LM_IN_below_l159_sw40, LM_IN_below_l159_underPF, LM_IN_below_l17_Parts, LM_IN_below_l17_underPF, LM_IN_below_l18_sw26, LM_IN_below_l18_sw28, LM_IN_below_l18_underPF, LM_IN_below_l19_Layer0, LM_IN_below_l19_Parts, LM_IN_below_l19_sw36, LM_IN_below_l19_sw37, LM_IN_below_l19_underPF, LM_IN_below_l2_Parts, LM_IN_below_l2_sw26, LM_IN_below_l2_sw27, LM_IN_below_l2_underPF, LM_IN_below_l20_Layer0, LM_IN_below_l20_Parts, LM_IN_below_l20_Playfield, LM_IN_below_l20_kick_sw7l, LM_IN_below_l20_sw19, LM_IN_below_l20_sw19a, LM_IN_below_l20_sw19b, LM_IN_below_l20_sw21, LM_IN_below_l20_sw26, LM_IN_below_l20_sw27, LM_IN_below_l20_sw28, LM_IN_below_l20_sw29, LM_IN_below_l20_sw30, LM_IN_below_l20_sw31, LM_IN_below_l20_sw33, LM_IN_below_l20_sw33a, _
  LM_IN_below_l20_sw34, LM_IN_below_l20_sw35, LM_IN_below_l20_sw36, LM_IN_below_l20_sw37, LM_IN_below_l20_sw40, LM_IN_below_l20_sw40a, LM_IN_below_l20_underPF, LM_IN_below_l21_Parts, LM_IN_below_l21_sw26, LM_IN_below_l21_underPF, LM_IN_below_l22_Parts, LM_IN_below_l22_underPF, LM_IN_below_l23_flip_l_up, LM_IN_below_l23_underPF, LM_IN_below_l24_underPF, LM_IN_below_l25_underPF, LM_IN_below_l26_Parts, LM_IN_below_l26_underPF, LM_IN_below_l3_underPF, LM_IN_below_l33_Layer0, LM_IN_below_l33_Parts, LM_IN_below_l33_underPF, LM_IN_below_l34_sw29, LM_IN_below_l34_sw31, LM_IN_below_l34_underPF, LM_IN_below_l35_underPF, LM_IN_below_l36_Layer0, LM_IN_below_l36_Parts, LM_IN_below_l36_Playfield, LM_IN_below_l36_kick_sw7l, LM_IN_below_l36_sw19, LM_IN_below_l36_sw19a, LM_IN_below_l36_sw21, LM_IN_below_l36_sw21a, LM_IN_below_l36_sw21b, LM_IN_below_l36_sw26, LM_IN_below_l36_sw27, LM_IN_below_l36_sw28, LM_IN_below_l36_sw29, LM_IN_below_l36_sw30, LM_IN_below_l36_sw31, LM_IN_below_l36_sw33, LM_IN_below_l36_sw33a, _
  LM_IN_below_l36_sw37, LM_IN_below_l36_sw38, LM_IN_below_l36_sw39, LM_IN_below_l36_sw40, LM_IN_below_l36_sw40a, LM_IN_below_l36_underPF, LM_IN_below_l37_Parts, LM_IN_below_l37_sw31, LM_IN_below_l37_underPF, LM_IN_below_l38_BM_PincabRails, LM_IN_below_l38_Parts, LM_IN_below_l38_underPF, LM_IN_below_l39_underPF, LM_IN_below_l4_Parts, LM_IN_below_l4_Playfield, LM_IN_below_l4_sw27, LM_IN_below_l4_sw28, LM_IN_below_l4_sw29, LM_IN_below_l4_sw30, LM_IN_below_l4_sw37, LM_IN_below_l4_sw40a, LM_IN_below_l4_sw42, LM_IN_below_l4_sw43, LM_IN_below_l4_sw44, LM_IN_below_l4_sw47, LM_IN_below_l40_underPF, LM_IN_below_l41_underPF, LM_IN_below_l42_Parts, LM_IN_below_l42_Playfield, LM_IN_below_l42_flip_l, LM_IN_below_l42_flip_r, LM_IN_below_l42_flip_r_up, LM_IN_below_l42_flipr_r, LM_IN_below_l42_flipr_r_up, LM_IN_below_l49_sw31, LM_IN_below_l49_underPF, LM_IN_below_l5_Parts, LM_IN_below_l5_sw26, LM_IN_below_l5_underPF, LM_IN_below_l50_Parts, LM_IN_below_l50_sw29, LM_IN_below_l50_sw30, LM_IN_below_l50_sw31, _
  LM_IN_below_l50_underPF, LM_IN_below_l51_underPF, LM_IN_below_l52_Parts, LM_IN_below_l52_Playfield, LM_IN_below_l52_sw27, LM_IN_below_l52_sw28, LM_IN_below_l52_sw29, LM_IN_below_l52_sw30, LM_IN_below_l52_sw37, LM_IN_below_l52_sw38, LM_IN_below_l52_sw45, LM_IN_below_l52_sw46, LM_IN_below_l52_sw47, LM_IN_below_l52_underPF, LM_IN_below_l53_Parts, LM_IN_below_l53_sw31, LM_IN_below_l53_underPF, LM_IN_below_l54_Parts, LM_IN_below_l54_underPF, LM_IN_below_l55_flip_r_up, LM_IN_below_l55_flipr_r_up, LM_IN_below_l55_underPF, LM_IN_below_l56_underPF, LM_IN_below_l57_underPF, LM_IN_below_l58_Parts, LM_IN_below_l58_underPF, LM_IN_below_l59_Parts, LM_IN_below_l59_Playfield, LM_IN_below_l59_flip_l, LM_IN_below_l59_flip_l_up, LM_IN_below_l59_flip_r, LM_IN_below_l59_flipr_l, LM_IN_below_l59_flipr_l_up, LM_IN_below_l6_underPF, LM_IN_below_l65_underPF, LM_IN_below_l66_underPF, LM_IN_below_l67_underPF, LM_IN_below_l68_underPF, LM_IN_below_l69_underPF, LM_IN_below_l7_underPF, LM_IN_below_l70_underPF, LM_IN_below_l71_underPF, _
  LM_IN_below_l72_underPF, LM_IN_below_l73_underPF, LM_IN_below_l74_underPF, LM_IN_below_l75_underPF, LM_IN_below_l76_underPF, LM_IN_below_l8_underPF, LM_IN_below_l81_underPF, LM_IN_below_l82_underPF, LM_IN_below_l83_underPF, LM_IN_below_l84_underPF, LM_IN_below_l85_underPF, LM_IN_below_l86_underPF, LM_IN_below_l87_underPF, LM_IN_below_l88_underPF, LM_IN_below_l89_underPF, LM_IN_below_l9_underPF, LM_IN_below_l90_underPF, LM_IN_below_l91_underPF, LM_IN_below_l92_underPF, LM_IN_below_l97_underPF, LM_IN_below_l98_underPF, LM_IN_below_l99_underPF)
Dim BG_All: BG_All=Array(BM_BM_PincabRails, BM_Layer0, BM_Parts, BM_Playfield, BM_flip_l, BM_flip_l_up, BM_flip_r, BM_flip_r_up, BM_flipr_l, BM_flipr_l_up, BM_flipr_r, BM_flipr_r_up, BM_kick_sw17, BM_kick_sw24, BM_kick_sw41, BM_kick_sw48, BM_kick_sw7l, BM_kick_sw7r, BM_sw18, BM_sw19, BM_sw19a, BM_sw19b, BM_sw20, BM_sw21, BM_sw21a, BM_sw21b, BM_sw22, BM_sw23, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw33, BM_sw33a, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw38, BM_sw39, BM_sw40, BM_sw40a, BM_sw42, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw47, BM_underPF, LM_GI_GI_001_Parts, LM_GI_GI_001_Playfield, LM_GI_GI_001_flip_l, LM_GI_GI_001_flip_l_up, LM_GI_GI_001_flip_r, LM_GI_GI_001_flipr_l, LM_GI_GI_001_flipr_l_up, LM_GI_GI_001_kick_sw17, LM_GI_GI_002_Parts, LM_GI_GI_002_Playfield, LM_GI_GI_002_flip_l, LM_GI_GI_002_flip_r, LM_GI_GI_002_flip_r_up, LM_GI_GI_002_flipr_r, LM_GI_GI_002_flipr_r_up, LM_GI_GI_002_kick_sw24, LM_GI_GI_002_sw20, LM_GI_GI_10_Parts, LM_GI_GI_10_Playfield, LM_GI_GI_10_flip_l, _
  LM_GI_GI_10_flip_l_up, LM_GI_GI_10_flip_r, LM_GI_GI_10_flip_r_up, LM_GI_GI_10_flipr_l, LM_GI_GI_10_flipr_l_up, LM_GI_GI_10_kick_sw17, LM_GI_GI_10_kick_sw48, LM_GI_GI_10_sw20, LM_GI_GI_10_sw22, LM_GI_GI_10_sw23, LM_GI_GI_11_Parts, LM_GI_GI_11_Playfield, LM_GI_GI_11_flip_l, LM_GI_GI_11_flip_l_up, LM_GI_GI_11_flip_r_up, LM_GI_GI_11_flipr_l, LM_GI_GI_11_flipr_l_up, LM_GI_GI_11_kick_sw17, LM_GI_GI_11_kick_sw48, LM_GI_GI_11_sw18, LM_GI_GI_11_sw20, LM_GI_GI_11_underPF, LM_GI_GI_12_Parts, LM_GI_GI_12_Playfield, LM_GI_GI_12_flip_l_up, LM_GI_GI_12_flip_r, LM_GI_GI_12_flip_r_up, LM_GI_GI_12_flipr_r, LM_GI_GI_12_flipr_r_up, LM_GI_GI_12_kick_sw24, LM_GI_GI_12_kick_sw41, LM_GI_GI_12_sw22, LM_GI_GI_13_Parts, LM_GI_GI_13_Playfield, LM_GI_GI_13_flip_l, LM_GI_GI_13_flip_l_up, LM_GI_GI_13_kick_sw17, LM_GI_GI_13_sw18, LM_GI_GI_13_sw20, LM_GI_GI_13_underPF, LM_GI_GI_14_Parts, LM_GI_GI_14_Playfield, LM_GI_GI_14_flip_l, LM_GI_GI_14_flip_r, LM_GI_GI_14_kick_sw24, LM_GI_GI_14_kick_sw41, LM_GI_GI_14_kick_sw48, LM_GI_GI_14_sw22, _
  LM_GI_GI_14_sw23, LM_GI_GI_14_underPF, LM_GI_GI_15_Parts, LM_GI_GI_15_Playfield, LM_GI_GI_15_kick_sw41, LM_GI_GI_15_sw27, LM_GI_GI_15_sw28, LM_GI_GI_15_sw29, LM_GI_GI_15_sw42, LM_GI_GI_15_sw43, LM_GI_GI_15_sw44, LM_GI_GI_15_sw45, LM_GI_GI_16_BM_PincabRails, LM_GI_GI_16_Parts, LM_GI_GI_16_Playfield, LM_GI_GI_16_kick_sw48, LM_GI_GI_16_sw27, LM_GI_GI_16_sw28, LM_GI_GI_16_sw29, LM_GI_GI_16_sw30, LM_GI_GI_16_sw45, LM_GI_GI_16_sw46, LM_GI_GI_16_sw47, LM_GI_GI_17_BM_PincabRails, LM_GI_GI_17_Parts, LM_GI_GI_17_Playfield, LM_GI_GI_17_kick_sw41, LM_GI_GI_17_sw18, LM_GI_GI_17_sw27, LM_GI_GI_17_sw28, LM_GI_GI_17_sw29, LM_GI_GI_17_sw30, LM_GI_GI_17_sw42, LM_GI_GI_17_sw43, LM_GI_GI_17_sw44, LM_GI_GI_17_sw45, LM_GI_GI_17_sw46, LM_GI_GI_18_Parts, LM_GI_GI_18_Playfield, LM_GI_GI_18_kick_sw48, LM_GI_GI_18_sw22, LM_GI_GI_18_sw28, LM_GI_GI_18_sw29, LM_GI_GI_18_sw30, LM_GI_GI_18_sw45, LM_GI_GI_18_sw46, LM_GI_GI_18_sw47, LM_GI_GI_19_Layer0, LM_GI_GI_19_Parts, LM_GI_GI_19_Playfield, LM_GI_GI_19_sw19b, LM_GI_GI_19_sw26, _
  LM_GI_GI_19_sw27, LM_GI_GI_19_sw28, LM_GI_GI_19_sw30, LM_GI_GI_19_sw37, LM_GI_GI_19_sw38, LM_GI_GI_19_sw39, LM_GI_GI_19_underPF, LM_GI_GI_2_Layer0, LM_GI_GI_2_Parts, LM_GI_GI_2_Playfield, LM_GI_GI_2_sw21a, LM_GI_GI_2_sw30, LM_GI_GI_2_sw31, LM_GI_GI_2_sw37, LM_GI_GI_2_sw38, LM_GI_GI_2_sw39, LM_GI_GI_2_sw40a, LM_GI_GI_2_underPF, LM_GI_GI_20_Layer0, LM_GI_GI_20_Parts, LM_GI_GI_20_Playfield, LM_GI_GI_20_sw21b, LM_GI_GI_20_sw27, LM_GI_GI_20_sw29, LM_GI_GI_20_sw30, LM_GI_GI_20_sw31, LM_GI_GI_20_sw37, LM_GI_GI_20_sw38, LM_GI_GI_20_sw39, LM_GI_GI_20_sw47, LM_GI_GI_20_underPF, LM_GI_GI_22222_Parts, LM_GI_GI_22222_Playfield, LM_GI_GI_22222_flip_l, LM_GI_GI_22222_flip_l_up, LM_GI_GI_22222_flip_r, LM_GI_GI_22222_flip_r_up, LM_GI_GI_22222_flipr_l, LM_GI_GI_22222_flipr_l_up, LM_GI_GI_22222_flipr_r, LM_GI_GI_22222_flipr_r_up, LM_GI_GI_23_BM_PincabRails, LM_GI_GI_23_Layer0, LM_GI_GI_23_Parts, LM_GI_GI_23_Playfield, LM_GI_GI_23_kick_sw7r, LM_GI_GI_23_sw19, LM_GI_GI_23_sw19a, LM_GI_GI_23_sw19b, LM_GI_GI_23_sw21a, _
  LM_GI_GI_23_sw26, LM_GI_GI_23_sw33, LM_GI_GI_23_sw33a, LM_GI_GI_23_sw34, LM_GI_GI_23_sw40, LM_GI_GI_23_sw40a, LM_GI_GI_23_underPF, LM_GI_GI_24_BM_PincabRails, LM_GI_GI_24_Layer0, LM_GI_GI_24_Parts, LM_GI_GI_24_Playfield, LM_GI_GI_24_sw21, LM_GI_GI_24_sw21a, LM_GI_GI_24_sw21b, LM_GI_GI_24_sw31, LM_GI_GI_24_sw33, LM_GI_GI_24_sw33a, LM_GI_GI_24_sw37, LM_GI_GI_24_sw38, LM_GI_GI_24_sw39, LM_GI_GI_24_sw40, LM_GI_GI_24_sw40a, LM_GI_GI_24_underPF, LM_GI_GI_25_Layer0, LM_GI_GI_25_Parts, LM_GI_GI_25_Playfield, LM_GI_GI_25_kick_sw7l, LM_GI_GI_25_kick_sw7r, LM_GI_GI_25_sw19, LM_GI_GI_25_sw21, LM_GI_GI_25_sw21a, LM_GI_GI_25_sw21b, LM_GI_GI_25_sw26, LM_GI_GI_25_sw28, LM_GI_GI_25_sw30, LM_GI_GI_25_sw31, LM_GI_GI_25_sw33, LM_GI_GI_25_sw33a, LM_GI_GI_25_sw37, LM_GI_GI_25_sw38, LM_GI_GI_25_sw39, LM_GI_GI_25_sw40, LM_GI_GI_25_sw40a, LM_GI_GI_25_underPF, LM_GI_GI_26_Layer0, LM_GI_GI_26_Parts, LM_GI_GI_26_Playfield, LM_GI_GI_26_kick_sw7l, LM_GI_GI_26_kick_sw7r, LM_GI_GI_26_sw19, LM_GI_GI_26_sw19a, LM_GI_GI_26_sw19b, _
  LM_GI_GI_26_sw21, LM_GI_GI_26_sw26, LM_GI_GI_26_sw27, LM_GI_GI_26_sw31, LM_GI_GI_26_sw33, LM_GI_GI_26_sw33a, LM_GI_GI_26_sw34, LM_GI_GI_26_sw35, LM_GI_GI_26_sw36, LM_GI_GI_26_sw37, LM_GI_GI_26_sw38, LM_GI_GI_26_sw40, LM_GI_GI_26_sw40a, LM_GI_GI_26_underPF, LM_GI_GI_27_BM_PincabRails, LM_GI_GI_27_Layer0, LM_GI_GI_27_Parts, LM_GI_GI_27_Playfield, LM_GI_GI_27_sw19, LM_GI_GI_27_sw19a, LM_GI_GI_27_sw19b, LM_GI_GI_27_sw26, LM_GI_GI_27_sw29, LM_GI_GI_27_sw33, LM_GI_GI_27_sw33a, LM_GI_GI_27_sw34, LM_GI_GI_27_sw35, LM_GI_GI_27_underPF, LM_GI_GI_28_Layer0, LM_GI_GI_28_Parts, LM_GI_GI_28_Playfield, LM_GI_GI_28_sw21, LM_GI_GI_28_sw21a, LM_GI_GI_28_sw21b, LM_GI_GI_28_sw26, LM_GI_GI_28_sw28, LM_GI_GI_28_sw31, LM_GI_GI_28_sw37, LM_GI_GI_28_sw38, LM_GI_GI_28_sw39, LM_GI_GI_28_sw40, LM_GI_GI_28_sw40a, LM_GI_GI_28_underPF, LM_GI_GI_29_Layer0, LM_GI_GI_29_Parts, LM_GI_GI_29_Playfield, LM_GI_GI_29_sw19a, LM_GI_GI_29_sw19b, LM_GI_GI_29_underPF, LM_GI_GI_3_Layer0, LM_GI_GI_3_Parts, LM_GI_GI_3_Playfield, LM_GI_GI_3_sw19a, _
  LM_GI_GI_3_sw26, LM_GI_GI_3_sw27, LM_GI_GI_3_sw28, LM_GI_GI_3_sw34, LM_GI_GI_3_sw37, LM_GI_GI_3_underPF, LM_GI_GI_30_Layer0, LM_GI_GI_30_Parts, LM_GI_GI_30_Playfield, LM_GI_GI_30_sw21, LM_GI_GI_30_sw21a, LM_GI_GI_30_sw40, LM_GI_GI_30_underPF, LM_GI_GI_34_Parts, LM_GI_GI_34_Playfield, LM_GI_GI_34_flip_l, LM_GI_GI_34_flip_l_up, LM_GI_GI_34_flip_r, LM_GI_GI_34_flip_r_up, LM_GI_GI_34_flipr_l, LM_GI_GI_34_flipr_l_up, LM_GI_GI_34_flipr_r, LM_GI_GI_34_flipr_r_up, LM_GI_GI_35_Parts, LM_GI_GI_35_Playfield, LM_GI_GI_35_flip_l, LM_GI_GI_35_flip_l_up, LM_GI_GI_35_flip_r, LM_GI_GI_35_flip_r_up, LM_GI_GI_35_flipr_l, LM_GI_GI_35_flipr_l_up, LM_GI_GI_35_flipr_r, LM_GI_GI_35_flipr_r_up, LM_GI_GI_36_Parts, LM_GI_GI_36_Playfield, LM_GI_GI_36_flip_l, LM_GI_GI_36_flip_l_up, LM_GI_GI_36_flip_r, LM_GI_GI_36_flip_r_up, LM_GI_GI_36_flipr_r, LM_GI_GI_36_flipr_r_up, LM_GI_GI_37_Parts, LM_GI_GI_37_Playfield, LM_GI_GI_37_flip_l, LM_GI_GI_37_flip_l_up, LM_GI_GI_37_flip_r, LM_GI_GI_37_flip_r_up, LM_GI_GI_37_flipr_l, LM_GI_GI_37_flipr_l_up, _
  LM_GI_GI_38_Parts, LM_GI_GI_38_Playfield, LM_GI_GI_38_flip_l, LM_GI_GI_38_flip_l_up, LM_GI_GI_38_flip_r, LM_GI_GI_38_flip_r_up, LM_GI_GI_38_flipr_l, LM_GI_GI_38_flipr_l_up, LM_GI_GI_39_Parts, LM_GI_GI_39_Playfield, LM_GI_GI_39_flip_l, LM_GI_GI_39_flip_l_up, LM_GI_GI_39_flip_r, LM_GI_GI_39_flip_r_up, LM_GI_GI_39_flipr_r, LM_GI_GI_39_flipr_r_up, LM_GI_GI_5_BM_PincabRails, LM_GI_GI_5_Layer0, LM_GI_GI_5_Parts, LM_GI_GI_5_Playfield, LM_GI_GI_5_sw19, LM_GI_GI_5_sw19a, LM_GI_GI_5_sw33, LM_GI_GI_5_underPF, LM_GI_GI_6_Layer0, LM_GI_GI_6_Parts, LM_GI_GI_6_Playfield, LM_GI_GI_6_sw21a, LM_GI_GI_6_sw21b, LM_GI_GI_6_underPF, LM_GI_GI_7_BM_PincabRails, LM_GI_GI_7_Parts, LM_GI_GI_7_Playfield, LM_GI_GI_7_kick_sw24, LM_GI_GI_7_kick_sw48, LM_GI_GI_7_sw22, LM_GI_GI_7_sw23, LM_GI_GI_7_sw29, LM_GI_GI_7_underPF, LM_GI_GI_8_BM_PincabRails, LM_GI_GI_8_Parts, LM_GI_GI_8_Playfield, LM_GI_GI_8_kick_sw17, LM_GI_GI_8_kick_sw41, LM_GI_GI_8_sw18, LM_GI_GI_8_sw20, LM_GI_GI_8_underPF, LM_GI_GI_9_Parts, LM_GI_GI_9_Playfield, LM_GI_GI_9_flip_l, _
  LM_GI_GI_9_flip_l_up, LM_GI_GI_9_flip_r, LM_GI_GI_9_flip_r_up, LM_GI_GI_9_flipr_r, LM_GI_GI_9_flipr_r_up, LM_GI_GI_9_kick_sw24, LM_GI_GI_9_kick_sw41, LM_GI_GI_9_sw18, LM_GI_GI_9_sw20, LM_GI_GI_9_sw22, LM_IN_above_l1_Parts, LM_IN_above_l1_Playfield, LM_IN_above_l1_sw26, LM_IN_above_l1_underPF, LM_IN_above_l10_Playfield, LM_IN_above_l10_underPF, LM_IN_above_l100_Playfield, LM_IN_above_l100_underPF, LM_IN_above_l101_Playfield, LM_IN_above_l101_underPF, LM_IN_above_l102_Playfield, LM_IN_above_l102_underPF, LM_IN_above_l103_Playfield, LM_IN_above_l103_underPF, LM_IN_above_l104_Playfield, LM_IN_above_l105_Playfield, LM_IN_above_l106_Playfield, LM_IN_above_l106_underPF, LM_IN_above_l107_Playfield, LM_IN_above_l108_Playfield, LM_IN_above_l113_Playfield, LM_IN_above_l114_Playfield, LM_IN_above_l115_Playfield, LM_IN_above_l116_Playfield, LM_IN_above_l117_Playfield, LM_IN_above_l118_Playfield, LM_IN_above_l119_Playfield, LM_IN_above_l120_Playfield, LM_IN_above_l121_Playfield, LM_IN_above_l122_Playfield, _
  LM_IN_above_l123_Playfield, LM_IN_above_l124_Playfield, LM_IN_above_l150_Layer0, LM_IN_above_l150_Parts, LM_IN_above_l150_Playfield, LM_IN_above_l150_sw19, LM_IN_above_l150_underPF, LM_IN_above_l151_Layer0, LM_IN_above_l151_Parts, LM_IN_above_l151_Playfield, LM_IN_above_l151_sw19a, LM_IN_above_l151_underPF, LM_IN_above_l152_Layer0, LM_IN_above_l152_Parts, LM_IN_above_l152_Playfield, LM_IN_above_l152_sw19b, LM_IN_above_l152_underPF, LM_IN_above_l153_Layer0, LM_IN_above_l153_Parts, LM_IN_above_l153_Playfield, LM_IN_above_l153_sw33, LM_IN_above_l153_underPF, LM_IN_above_l154_Layer0, LM_IN_above_l154_Parts, LM_IN_above_l154_Playfield, LM_IN_above_l154_sw33a, LM_IN_above_l154_underPF, LM_IN_above_l155_Layer0, LM_IN_above_l155_Parts, LM_IN_above_l155_Playfield, LM_IN_above_l155_sw21, LM_IN_above_l155_underPF, LM_IN_above_l156_Layer0, LM_IN_above_l156_Parts, LM_IN_above_l156_Playfield, LM_IN_above_l156_sw21a, LM_IN_above_l156_underPF, LM_IN_above_l157_Layer0, LM_IN_above_l157_Parts, LM_IN_above_l157_Playfield, _
  LM_IN_above_l157_sw21b, LM_IN_above_l157_underPF, LM_IN_above_l158_Layer0, LM_IN_above_l158_Parts, LM_IN_above_l158_Playfield, LM_IN_above_l158_sw37, LM_IN_above_l158_sw38, LM_IN_above_l158_sw40a, LM_IN_above_l158_underPF, LM_IN_above_l159_Layer0, LM_IN_above_l159_Parts, LM_IN_above_l159_Playfield, LM_IN_above_l159_sw38, LM_IN_above_l159_sw39, LM_IN_above_l159_sw40, LM_IN_above_l159_underPF, LM_IN_above_l17_Parts, LM_IN_above_l17_Playfield, LM_IN_above_l17_sw26, LM_IN_above_l17_underPF, LM_IN_above_l18_Playfield, LM_IN_above_l18_sw28, LM_IN_above_l18_underPF, LM_IN_above_l19_Playfield, LM_IN_above_l19_underPF, LM_IN_above_l2_Playfield, LM_IN_above_l2_sw26, LM_IN_above_l2_sw27, LM_IN_above_l2_underPF, LM_IN_above_l21_Parts, LM_IN_above_l21_Playfield, LM_IN_above_l21_underPF, LM_IN_above_l22_Playfield, LM_IN_above_l22_underPF, LM_IN_above_l23_Playfield, LM_IN_above_l23_flip_l_up, LM_IN_above_l23_flip_r_up, LM_IN_above_l23_flipr_l_up, LM_IN_above_l24_Playfield, LM_IN_above_l25_Playfield, LM_IN_above_l25_underPF, _
  LM_IN_above_l26_Parts, LM_IN_above_l26_Playfield, LM_IN_above_l26_underPF, LM_IN_above_l3_Playfield, LM_IN_above_l3_underPF, LM_IN_above_l33_Parts, LM_IN_above_l33_Playfield, LM_IN_above_l33_sw31, LM_IN_above_l33_underPF, LM_IN_above_l34_Playfield, LM_IN_above_l34_sw29, LM_IN_above_l34_underPF, LM_IN_above_l35_Playfield, LM_IN_above_l35_underPF, LM_IN_above_l37_Playfield, LM_IN_above_l37_sw31, LM_IN_above_l37_underPF, LM_IN_above_l38_Playfield, LM_IN_above_l38_underPF, LM_IN_above_l39_Parts, LM_IN_above_l39_Playfield, LM_IN_above_l39_flip_l_up, LM_IN_above_l39_flip_r_up, LM_IN_above_l40_Playfield, LM_IN_above_l41_Playfield, LM_IN_above_l41_underPF, LM_IN_above_l49_Layer0, LM_IN_above_l49_Parts, LM_IN_above_l49_Playfield, LM_IN_above_l49_sw30, LM_IN_above_l49_sw31, LM_IN_above_l49_sw39, LM_IN_above_l49_underPF, LM_IN_above_l5_Playfield, LM_IN_above_l5_underPF, LM_IN_above_l50_Parts, LM_IN_above_l50_Playfield, LM_IN_above_l50_sw30, LM_IN_above_l50_underPF, LM_IN_above_l51_Playfield, LM_IN_above_l51_underPF, _
  LM_IN_above_l53_Playfield, LM_IN_above_l53_underPF, LM_IN_above_l54_Playfield, LM_IN_above_l54_underPF, LM_IN_above_l55_Playfield, LM_IN_above_l55_flip_l_up, LM_IN_above_l55_flip_r_up, LM_IN_above_l55_flipr_r_up, LM_IN_above_l56_Playfield, LM_IN_above_l57_Playfield, LM_IN_above_l57_sw27, LM_IN_above_l57_underPF, LM_IN_above_l58_Parts, LM_IN_above_l58_Playfield, LM_IN_above_l58_underPF, LM_IN_above_l6_Playfield, LM_IN_above_l6_underPF, LM_IN_above_l65_Playfield, LM_IN_above_l66_Playfield, LM_IN_above_l67_Playfield, LM_IN_above_l68_Playfield, LM_IN_above_l69_Playfield, LM_IN_above_l7_Playfield, LM_IN_above_l7_flipr_l_up, LM_IN_above_l70_Playfield, LM_IN_above_l71_Playfield, LM_IN_above_l72_Playfield, LM_IN_above_l73_Playfield, LM_IN_above_l74_Playfield, LM_IN_above_l75_Playfield, LM_IN_above_l76_Playfield, LM_IN_above_l8_Playfield, LM_IN_above_l8_flipr_r_up, LM_IN_above_l81_Parts, LM_IN_above_l81_Playfield, LM_IN_above_l82_Playfield, LM_IN_above_l82_underPF, LM_IN_above_l83_Parts, LM_IN_above_l83_Playfield, _
  LM_IN_above_l84_Parts, LM_IN_above_l84_Playfield, LM_IN_above_l84_underPF, LM_IN_above_l85_Playfield, LM_IN_above_l85_underPF, LM_IN_above_l86_Playfield, LM_IN_above_l86_underPF, LM_IN_above_l87_Playfield, LM_IN_above_l88_Playfield, LM_IN_above_l88_underPF, LM_IN_above_l89_Parts, LM_IN_above_l89_Playfield, LM_IN_above_l89_underPF, LM_IN_above_l9_Playfield, LM_IN_above_l9_underPF, LM_IN_above_l90_Playfield, LM_IN_above_l90_kick_sw48, LM_IN_above_l90_underPF, LM_IN_above_l91_Playfield, LM_IN_above_l91_underPF, LM_IN_above_l92_Parts, LM_IN_above_l92_Playfield, LM_IN_above_l92_underPF, LM_IN_above_l97_Playfield, LM_IN_above_l98_Playfield, LM_IN_above_l98_underPF, LM_IN_above_l99_Playfield, LM_IN_above_l99_underPF, LM_IN_below_l1_Layer0, LM_IN_below_l1_sw26, LM_IN_below_l1_underPF, LM_IN_below_l10_underPF, LM_IN_below_l100_underPF, LM_IN_below_l101_underPF, LM_IN_below_l102_underPF, LM_IN_below_l103_underPF, LM_IN_below_l104_underPF, LM_IN_below_l105_underPF, LM_IN_below_l106_underPF, LM_IN_below_l107_underPF, _
  LM_IN_below_l108_underPF, LM_IN_below_l113_underPF, LM_IN_below_l114_underPF, LM_IN_below_l115_underPF, LM_IN_below_l116_underPF, LM_IN_below_l117_underPF, LM_IN_below_l118_underPF, LM_IN_below_l119_underPF, LM_IN_below_l120_underPF, LM_IN_below_l121_underPF, LM_IN_below_l122_underPF, LM_IN_below_l123_underPF, LM_IN_below_l124_underPF, LM_IN_below_l150_Layer0, LM_IN_below_l150_Parts, LM_IN_below_l150_Playfield, LM_IN_below_l150_sw19, LM_IN_below_l150_sw33a, LM_IN_below_l150_underPF, LM_IN_below_l151_Layer0, LM_IN_below_l151_Parts, LM_IN_below_l151_Playfield, LM_IN_below_l151_sw19a, LM_IN_below_l151_underPF, LM_IN_below_l152_Layer0, LM_IN_below_l152_Parts, LM_IN_below_l152_Playfield, LM_IN_below_l152_sw19b, LM_IN_below_l152_underPF, LM_IN_below_l153_Layer0, LM_IN_below_l153_Parts, LM_IN_below_l153_Playfield, LM_IN_below_l153_sw33, LM_IN_below_l153_underPF, LM_IN_below_l154_Layer0, LM_IN_below_l154_Parts, LM_IN_below_l154_Playfield, LM_IN_below_l154_sw33a, LM_IN_below_l154_underPF, LM_IN_below_l155_Layer0, _
  LM_IN_below_l155_Parts, LM_IN_below_l155_Playfield, LM_IN_below_l155_sw21, LM_IN_below_l155_sw40a, LM_IN_below_l155_underPF, LM_IN_below_l156_Layer0, LM_IN_below_l156_Parts, LM_IN_below_l156_Playfield, LM_IN_below_l156_sw21a, LM_IN_below_l156_underPF, LM_IN_below_l157_Layer0, LM_IN_below_l157_Parts, LM_IN_below_l157_Playfield, LM_IN_below_l157_sw21b, LM_IN_below_l157_underPF, LM_IN_below_l158_Layer0, LM_IN_below_l158_Parts, LM_IN_below_l158_Playfield, LM_IN_below_l158_sw37, LM_IN_below_l158_sw38, LM_IN_below_l158_sw39, LM_IN_below_l158_sw40a, LM_IN_below_l158_underPF, LM_IN_below_l159_Layer0, LM_IN_below_l159_Parts, LM_IN_below_l159_Playfield, LM_IN_below_l159_sw21a, LM_IN_below_l159_sw37, LM_IN_below_l159_sw38, LM_IN_below_l159_sw39, LM_IN_below_l159_sw40, LM_IN_below_l159_underPF, LM_IN_below_l17_Parts, LM_IN_below_l17_underPF, LM_IN_below_l18_sw26, LM_IN_below_l18_sw28, LM_IN_below_l18_underPF, LM_IN_below_l19_Layer0, LM_IN_below_l19_Parts, LM_IN_below_l19_sw36, LM_IN_below_l19_sw37, _
  LM_IN_below_l19_underPF, LM_IN_below_l2_Parts, LM_IN_below_l2_sw26, LM_IN_below_l2_sw27, LM_IN_below_l2_underPF, LM_IN_below_l20_Layer0, LM_IN_below_l20_Parts, LM_IN_below_l20_Playfield, LM_IN_below_l20_kick_sw7l, LM_IN_below_l20_sw19, LM_IN_below_l20_sw19a, LM_IN_below_l20_sw19b, LM_IN_below_l20_sw21, LM_IN_below_l20_sw26, LM_IN_below_l20_sw27, LM_IN_below_l20_sw28, LM_IN_below_l20_sw29, LM_IN_below_l20_sw30, LM_IN_below_l20_sw31, LM_IN_below_l20_sw33, LM_IN_below_l20_sw33a, LM_IN_below_l20_sw34, LM_IN_below_l20_sw35, LM_IN_below_l20_sw36, LM_IN_below_l20_sw37, LM_IN_below_l20_sw40, LM_IN_below_l20_sw40a, LM_IN_below_l20_underPF, LM_IN_below_l21_Parts, LM_IN_below_l21_sw26, LM_IN_below_l21_underPF, LM_IN_below_l22_Parts, LM_IN_below_l22_underPF, LM_IN_below_l23_flip_l_up, LM_IN_below_l23_underPF, LM_IN_below_l24_underPF, LM_IN_below_l25_underPF, LM_IN_below_l26_Parts, LM_IN_below_l26_underPF, LM_IN_below_l3_underPF, LM_IN_below_l33_Layer0, LM_IN_below_l33_Parts, LM_IN_below_l33_underPF, LM_IN_below_l34_sw29, _
  LM_IN_below_l34_sw31, LM_IN_below_l34_underPF, LM_IN_below_l35_underPF, LM_IN_below_l36_Layer0, LM_IN_below_l36_Parts, LM_IN_below_l36_Playfield, LM_IN_below_l36_kick_sw7l, LM_IN_below_l36_sw19, LM_IN_below_l36_sw19a, LM_IN_below_l36_sw21, LM_IN_below_l36_sw21a, LM_IN_below_l36_sw21b, LM_IN_below_l36_sw26, LM_IN_below_l36_sw27, LM_IN_below_l36_sw28, LM_IN_below_l36_sw29, LM_IN_below_l36_sw30, LM_IN_below_l36_sw31, LM_IN_below_l36_sw33, LM_IN_below_l36_sw33a, LM_IN_below_l36_sw37, LM_IN_below_l36_sw38, LM_IN_below_l36_sw39, LM_IN_below_l36_sw40, LM_IN_below_l36_sw40a, LM_IN_below_l36_underPF, LM_IN_below_l37_Parts, LM_IN_below_l37_sw31, LM_IN_below_l37_underPF, LM_IN_below_l38_BM_PincabRails, LM_IN_below_l38_Parts, LM_IN_below_l38_underPF, LM_IN_below_l39_underPF, LM_IN_below_l4_Parts, LM_IN_below_l4_Playfield, LM_IN_below_l4_sw27, LM_IN_below_l4_sw28, LM_IN_below_l4_sw29, LM_IN_below_l4_sw30, LM_IN_below_l4_sw37, LM_IN_below_l4_sw40a, LM_IN_below_l4_sw42, LM_IN_below_l4_sw43, LM_IN_below_l4_sw44, _
  LM_IN_below_l4_sw47, LM_IN_below_l40_underPF, LM_IN_below_l41_underPF, LM_IN_below_l42_Parts, LM_IN_below_l42_Playfield, LM_IN_below_l42_flip_l, LM_IN_below_l42_flip_r, LM_IN_below_l42_flip_r_up, LM_IN_below_l42_flipr_r, LM_IN_below_l42_flipr_r_up, LM_IN_below_l49_sw31, LM_IN_below_l49_underPF, LM_IN_below_l5_Parts, LM_IN_below_l5_sw26, LM_IN_below_l5_underPF, LM_IN_below_l50_Parts, LM_IN_below_l50_sw29, LM_IN_below_l50_sw30, LM_IN_below_l50_sw31, LM_IN_below_l50_underPF, LM_IN_below_l51_underPF, LM_IN_below_l52_Parts, LM_IN_below_l52_Playfield, LM_IN_below_l52_sw27, LM_IN_below_l52_sw28, LM_IN_below_l52_sw29, LM_IN_below_l52_sw30, LM_IN_below_l52_sw37, LM_IN_below_l52_sw38, LM_IN_below_l52_sw45, LM_IN_below_l52_sw46, LM_IN_below_l52_sw47, LM_IN_below_l52_underPF, LM_IN_below_l53_Parts, LM_IN_below_l53_sw31, LM_IN_below_l53_underPF, LM_IN_below_l54_Parts, LM_IN_below_l54_underPF, LM_IN_below_l55_flip_r_up, LM_IN_below_l55_flipr_r_up, LM_IN_below_l55_underPF, LM_IN_below_l56_underPF, LM_IN_below_l57_underPF, _
  LM_IN_below_l58_Parts, LM_IN_below_l58_underPF, LM_IN_below_l59_Parts, LM_IN_below_l59_Playfield, LM_IN_below_l59_flip_l, LM_IN_below_l59_flip_l_up, LM_IN_below_l59_flip_r, LM_IN_below_l59_flipr_l, LM_IN_below_l59_flipr_l_up, LM_IN_below_l6_underPF, LM_IN_below_l65_underPF, LM_IN_below_l66_underPF, LM_IN_below_l67_underPF, LM_IN_below_l68_underPF, LM_IN_below_l69_underPF, LM_IN_below_l7_underPF, LM_IN_below_l70_underPF, LM_IN_below_l71_underPF, LM_IN_below_l72_underPF, LM_IN_below_l73_underPF, LM_IN_below_l74_underPF, LM_IN_below_l75_underPF, LM_IN_below_l76_underPF, LM_IN_below_l8_underPF, LM_IN_below_l81_underPF, LM_IN_below_l82_underPF, LM_IN_below_l83_underPF, LM_IN_below_l84_underPF, LM_IN_below_l85_underPF, LM_IN_below_l86_underPF, LM_IN_below_l87_underPF, LM_IN_below_l88_underPF, LM_IN_below_l89_underPF, LM_IN_below_l9_underPF, LM_IN_below_l90_underPF, LM_IN_below_l91_underPF, LM_IN_below_l92_underPF, LM_IN_below_l97_underPF, LM_IN_below_l98_underPF, LM_IN_below_l99_underPF)
' VLM  Arrays - End


'******************************************************
'   ZANI: Misc Animations
'******************************************************


''''' Flippers

Sub LeftFlipper_Animate
  Dim a : a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a

  Dim v, BP
  v = 255.0 * (121.0 - LeftFlipper.CurrentAngle) / (121.0 -  70.0)

  For each BP in BP_flip_l
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_flipr_l
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_flip_l_up
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
  For each BP in BP_flipr_l_up
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub RightFlipper_Animate
  Dim a : a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a

  Dim v, BP
  v = 255.0 * (-121.0 - RightFlipper.CurrentAngle) / (-121.0 +  70.0)

  For each BP in BP_flip_r
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_flipr_r
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_flip_r_up
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
  For each BP in BP_flipr_r_up
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub


''''' Spinner Animations

Sub sw26_Animate
  Dim a : a = sw26.CurrentAngle
  Dim BP : For Each BP in BP_sw26 : BP.rotx = a: Next
End Sub

Sub sw31_Animate
  Dim a : a = sw31.CurrentAngle
  Dim BP : For Each BP in BP_sw31 : BP.rotx = a: Next
End Sub


''''' Switch Animations

Sub sw18_Animate
  Dim z : z = sw18.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw18 : BP.transz = z: Next
End Sub

Sub sw20_Animate
  Dim z : z = sw20.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw20 : BP.transz = z: Next
End Sub


Sub sw22_Animate
  Dim z : z = sw22.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw22 : BP.transz = z: Next
End Sub

Sub sw23_Animate
  Dim z : z = sw23.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw23 : BP.transz = z: Next
End Sub


Sub sw27_Animate
  Dim z : z = sw27.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw27 : BP.transz = z: Next
End Sub

Sub sw28_Animate
  Dim z : z = sw28.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw28 : BP.transz = z: Next
End Sub

Sub sw29_Animate
  Dim z : z = sw29.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw29 : BP.transz = z: Next
End Sub

Sub sw30_Animate
  Dim z : z = sw30.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw30 : BP.transz = z: Next
End Sub


''''' Drop Targets

Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_sw34.transz
  rx = BM_sw34.rotx
  ry = BM_sw34.roty
  For each BP in BP_sw34: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw35.transz
  rx = BM_sw35.rotx
  ry = BM_sw35.roty
  For each BP in BP_sw35: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw36.transz
  rx = BM_sw36.rotx
  ry = BM_sw36.roty
  For each BP in BP_sw36: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw37.transz
  rx = BM_sw37.rotx
  ry = BM_sw37.roty
  For each BP in BP_sw37: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw38.transz
  rx = BM_sw38.rotx
  ry = BM_sw38.roty
  For each BP in BP_sw38: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw39.transz
  rx = BM_sw39.rotx
  ry = BM_sw39.roty
  For each BP in BP_sw39: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw42.transz
  rx = BM_sw42.rotx
  ry = BM_sw42.roty
  For each BP in BP_sw42: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw43.transz
  rx = BM_sw43.rotx
  ry = BM_sw43.roty
  For each BP in BP_sw43: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw44.transz
  rx = BM_sw44.rotx
  ry = BM_sw44.roty
  For each BP in BP_sw44: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw45.transz
  rx = BM_sw45.rotx
  ry = BM_sw45.roty
  For each BP in BP_sw45: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw46.transz
  rx = BM_sw46.rotx
  ry = BM_sw46.roty
  For each BP in BP_sw46: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw47.transz
  rx = BM_sw47.rotx
  ry = BM_sw47.roty
  For each BP in BP_sw47: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub

'******************************
' ZVRR - Setup VR Room
'******************************

Sub SetupVRRoom
  dim VRThing, BP
  For Each BP in BP_BM_PincabRails : BP.Visible = 0: Next
  If RenderingMode = 2 or ForceVR = True Then
    VRMode = True
    For Each VRThing in VRThings: VRThing.visible = 1: Next
    For Each BP in BP_BM_PincabRails : BP.Visible = 1: Next
  Else
    VRMode = False
    For Each VRThing in VRThings: VRThing.visible = 0: Next
    If DesktopMode = True Then For Each BP in BP_BM_PincabRails : BP.Visible = 1: Next
  End If
End Sub

'******************* VR Backglass **********************

If RenderingMode = 2 or ForceVR = True Then
  dim VRobj
  For each VRobj in VRBackglass : VRobj.visible = true: Next
  SetBackglass
End If


Sub SetBackglass()
  Dim obj
  bgdark.visible = True
  For Each obj In VRBackglass
    obj.x = obj.x
    obj.height = - obj.y - 130
    obj.y = 112 'adjusts the distance from the backglass towards the user
    obj.rotx=-90
  Next

  For Each obj In VRAlphanumerics
    obj.x = obj.x
    obj.height = - obj.y - 130
    obj.y = 112 'adjusts the distance from the backglass towards the user
    obj.rotx=-90
  Next
'
' For Each obj In VRBackglassSpeaker
'   obj.x = obj.x
'   obj.height = - obj.y - 75
'   obj.y = 110 'adjusts the distance from the backglass towards the user
'   obj.rotx=-92
' Next
End Sub

' CHANGE LOG
' 000 - Sixtoe - Initial Build and Toolkit Render
' 001 - Sixtoe - New 1k render, fixed playfield, preliminary animations (which are wrong)
' 002 - Sixtoe - More work, forgot tbh.
' 003 - apophis - Fixed timers, DT animation, flipper animations. Added star trig animations. Added SolGI. Updated lights to use Incandescent fader. Updated flipper physics and triggers. Lowered table slope to 5.5
' 004 - apophis - Added fleep sound files.
' 005 - Sixtoe - New 2k render, numerous fixes and tweaks, various physics issues fixed.
' 006 - apophis - Added tweak menu. Added ambient ball shadows. Made triggers invisible. Adjusted flipper triggers per new flipper shapes. Flipper strength set to 1900. Added Metals collection for fleep sounds. Drain sound effects. Reflection probe on BM_Playfield. Enabled reflections for BM_Parts. Enable insert light reflections on ball. Trying out Tony McMapFace.
' 007 - Sixtoe - New 2k render, numerous fixes, added credit and start lights and cab which is bugged?), played with saucer physics, added physical saucer deflectors
' 008 - Sixtoe - New 1k render, Improved physics on saucers, improved top centre shot, numerous Blender improvements, animated saucer kickers
' 009 - Sixtoe - 4K bake, fixed camera port
' 010 - Sixtoe - new 4k bake with a few fixes, added VR cab and room (not scripted yet), set up the flipper physics (thanks Iaakki), some other changes.
' 011 - iaakki - saucer timers and strengths reworked
' 012 - Sixtoe - New 4k bake, added instruction cards, updated some physics objects, slowed down lower inlane kickers fractionally, tuned drain kicker, hooked up VR room auto code and animated buttons, hooked up room brightness, fixed coin door cabinet decal.
' 013 - leojreimroc - VRBackglass
' 014 - Sixtoe - Tweaks to VR cab and room, updated rubber physics, reinforced layout, added rules


' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub
