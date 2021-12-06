'Spider-Man - IPDB No. 5237
'? Stern 2007
'Table Recreation by ninuzzu & Alessio for Visual Pinball 10
'Credits/Thanks (in no particular order):
'85Vett and Alessio for the VP9 script
'Dark and Zany for some primitive templates I used and edited
'Javier1515,JPSalas,GTXJoe and Shoopity for some code and sounds effects I borrowed from their tables
'VPDev Team for the new freaking amazing VPX

'******************************************************
'* ROM VERSION 2.61 (sman_261.bin)                    *
'******************************************************
'* LAYERS *********************************************
'******************************************************
'* LAYER 1 = COLLISION OBJECTS ************************
'* LAYER 2 = PLASTICS AND DECALS **********************
'* LAYER 3 = RAMPS  ***********************************
'* LAYER 4 = 3D OBJECTS *******************************
'* LAYER 5 = LIGHTS REFLECTIONS ***********************
'* LAYER 6 = GI LIGHTS ********************************
'* LAYER 7 = FLASHERS *********************************
'* LAYER 8 = INSERTS AND SPOTLIGHTS *******************
'******************************************************

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

Const BallSize = 51
Const BallMass = 1.3

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const UseVPMModSol = 1

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMColoredDMD

UseVPMColoredDMD = DesktopMode

LoadVPM "01560000", "sam.VBS", 3.10

lamptimer.interval = -1


'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = "FlipperUp"
Const SFlipperOff = "FlipperDown"
Const SCoin = "coin"

'************
' Table init.
'************

Const cGameName = "sman_261" 'ENGLISH VERSION

'Const cGameName = "sman_210ai" 'ITALIAN VERSION

'Variables
Dim bsTrough, bsSandman, bsDocOck, x
Dim DocMagnet
Dim PlungerIM
Dim PalleInGioco
Dim Attendi
'Dim Magneth


Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Spider Man (Stern 2007) - Ninuzzu & Alessio"
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = DesktopMode
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With


    'Trough
  Set bsTrough = New cvpmTrough
  bsTrough.Size = 4
  bsTrough.InitSwitches Array(21, 20, 19, 18)
' bsTrough.EntrySw = 0
  bsTrough.InitExit BallRelease, 90, 8
  bsTrough.InitExitSounds SoundFX("Solenoid",DOFContactors), SoundFX("ballrelease",DOFContactors)
  bsTrough.Balls = 4

  'Sandman VUK
  Set bsSandman = New cvpmSaucer
  bsSandman.InitKicker sw59, 59, 0, 35, 1.56
  bsSandman.InitSounds "kicker_enter", SoundFX("Solenoid",DOFContactors), SoundFX("ExitSandman",DOFContactors)

  'Doc Ock VUK
  Set bsDocOck = New cvpmSaucer
  bsDocOck.InitKicker sw36, 36, 0, 35, 1.56
  bsDocOck.InitSounds "kicker_enter", SoundFX("Solenoid",DOFContactors), SoundFX("ExitDoc",DOFContactors)




  'Doc Ock Magmet
  Set DocMagnet = New cvpmMagnet
  DocMagnet.InitMagnet DocOckMagnet, 50
  DocMagnet.Solenoid = 3
  'DocMagnet.GrabCenter = True
  DocMagnet.CreateEvents "DocMagnet"




  'Loop Diverter
  diverter.IsDropped = 1

    'Nudging
  vpmNudge.TiltSwitch=-7
    vpmNudge.Sensitivity=3
    vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  'Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

    ' Impulse Plunger
    Const IMPowerSetting = 55   '55
    Const IMTime = 0.6
    Set plungerIM = New cvpmImpulseP
    With plungerIM
    .InitImpulseP swplunger, IMPowerSetting, IMTime
    .Switch 23
    .Random 0.3
    .InitExitSnd SoundFX("solenoid",DOFContactors), ""
        .CreateEvents "plungerIM"
    End With
  Attendi=1
  PausaAnimazione.Enabled=1
  Controller.Switch(50)=1
  Controller.Switch(53)=1
  Controller.Switch(57)=1
  BankAlto=1
  SandmanPronto=0
  PalleInGioco=0
  SandmanAlto=0
  PallaBucoOctopus=0
  PallaBucoSandman=0
  PallaSuMagnete=0
  sw16Premuto=0
  sw17Premuto=0
  'AlzaSandman
  sw36.Enabled=0
  sw59.Enabled=0
 End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit():Controller.Stop:End Sub

'************************************
' PAUSA ANIMAZIONE DOPO AVVIO TAVOLO
'************************************

Sub PausaAnimazione_Timer()
  Attendi=0
  sw63.IsDropped=1
  Me.Enabled=0
End Sub

Dim bulb

Set GICallback = GetRef("UpdateGI")

Sub UpdateGI(nr,enabled)
 DOF 200, enabled*-1
 Select Case nr
 Case 0
 For each bulb in GI
 bulb.state=enabled
 GestioneGIWall
 next
 End Select
End Sub

Sub GestioneGIWall
  If giPiano.state= 0 Then
  SpegniLuciWall
  Else
  AccendiLuciWall
  End If
End Sub

Sub LHP1_Hit()
If ActiveBall.velY < 0  Then    'ball is going up
PlaySound "rrenter"
Else
StopSound "rrenter"
End If
End Sub

Sub LHP2_Hit()
If ActiveBall.velY < 0  Then    'ball is going up
PlaySound "rrenter"
Else
StopSound "rrenter"
End If
End Sub

Sub RHP1_Hit()
If ActiveBall.velY < 0  Then    'ball is going up
PlaySound "rrenter"
Else
StopSound "rrenter"
End If
End Sub


'********************************************
'  Depotenzia magnete dopo presa della palla
'********************************************

Dim PallaSuMagnete

Sub PresenzaSuMagnete_Hit() 'CON PALLA SUL MAGNETE RIDUCE LA POTENZA DELLA CALAMITA
  If DocMagnet.MagnetON= True Then DepotenziaMagnete.Enabled=True: Playsound "Magnete"
  PallaSuMagnete=1
End Sub


Sub PresenzaSuMagnete_UnHit() 'IN USCITA DAL MAGNETE VIENE RIPRISTINATA LA POTENZA DELLA CALAMITA
  If DocMagnet.MagnetON= False Then DocMagnet.InitMagnet DocOckMagnet, 50
  PallaSuMagnete=0
End Sub


Sub DepotenziaMagnete_timer() 'TIMER PER RIDURRE LA POTENZA DEL MAGNETE
  DocMagnet.MagnetON= False
  DocMagnet.InitMagnet DocOckMagnet, 2
  DocMagnet.MagnetON= True
  Me.Enabled = 0
End Sub


'************************************
'  Lancio palla dopo sgancio magnete
'************************************

Dim Pausa

Sub solDocMagnet(enabled)
    MagnetOffTimer.Enabled = Not enabled
    If enabled Then DocMagnet.MagnetOn = True
End Sub

Sub MagnetOffTimer_Timer
    Dim ball
    For Each ball In DocMagnet.Balls
        With ball
      .VelX = 15: .VelY = -8: Pausa= 1  'ERA VelX = 15: .VelY = -7
        End With
    Next
    Me.Enabled = False:DocMagnet.MagnetOn = False
End Sub

Sub DocOckMagnet_UnHit()
  If Pausa=1 Then
  DocMagnet.MagnetON= False
  DocMagnet.InitMagnet DocOckMagnet, 0
  DocOckMagnet.Enabled=0
  MagneteDisabilitato.Enabled=1
  End If
End Sub

Sub MagneteDisabilitato_Timer()
  DocMagnet.InitMagnet DocOckMagnet, 50
  Pausa= 0
  DocOckMagnet.Enabled=1
  Me.Enabled=0
End Sub

Sub Table1_KeyDown(ByVal Keycode)

  If keycode = PlungerKey Then
    Plunger.PullBack:Playsound "plungerpull"
  End If
  If Keycode = LeftFlipperKey then
    Controller.Switch(84)=1
    Exit Sub
  End If
  If Keycode = RightFlipperKey then
    Controller.Switch(86)=1
    Controller.Switch(82)=1
    Exit Sub
  End If

  If keycode = LeftTiltKey Then PlaySound SoundFX("fx_nudge",0)
  If keycode = RightTiltKey Then PlaySound SoundFX("fx_nudge",0)
  If keycode = CenterTiltKey Then PlaySound SoundFX("fx_nudge",0)
  If vpmKeyDown(Keycode) Then Exit Sub
 End Sub

Sub Table1_KeyUp(ByVal Keycode)
  If keycode = PlungerKey Then
    Plunger.Fire: Playsound "Plunger"
  End If
  If vpmKeyUp(Keycode) Then Exit Sub
  If Keycode = LeftFlipperKey then
    Controller.Switch(84)=0
    Exit Sub
  End If
  If Keycode = RightFlipperKey then
    Controller.Switch(86)=0
    Controller.Switch(82)=0
    Exit Sub
  End If
 End Sub

'Realtime updates

Sub GatesTimer_Timer()
  GateSWsx.RotZ= -Gate2.currentangle
  GateSWdx.RotZ= -Gate3.currentangle
  GateP0.RotX = -Gate4.currentangle + 90
  GateP1.RotX = -Gate5.currentangle + 90
  GateP2.RotX = -LeftRampEnd.currentangle + 90    'gate V shape
  GateP3.RotX = -LeftRampStart.currentangle + 90
  GateP4.RotX = -RightRampEnd.currentangle +90    'gate V shape
  GateP5.RotX = -RightRampStart.currentangle +90
  UpdateFlipperLogos
  RollingSoundUpdate
  BallShadowUpdate
End Sub

Sub UpdateFlipperLogos
    flipperl.RotY = LeftFlipper.CurrentAngle
    flipperr.RotY = RightFlipper.CurrentAngle
  flipperr1.RotY = RightFlipper2.CurrentAngle
End Sub

'Solenoids
SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
SolCallback(3) = "solDocMagnet"
SolModCallback(3) = "SetModLampmm 0, 200,"

SolCallback(4) = "solDocVUK"
SolCallback(5) = "solDocMotor"
'SolCallback(6) = "ShakerMotor"
SolCallback(7) = "Gate2.open ="
SolCallback(8) = "Gate3.open ="
SolCallback(9) = "solLBump"
SolCallback(10) = "solRBump"
SolCallback(11) = "solBBump"
SolCallback(12) = "solSandVUK"
SolCallback(13) = "solSandMotor"
SolCallback(14) = "solURFlipper"
SolCallback(15) = "solLFlipper"
SolCallback(16) = "solRFlipper"
SolCallback(17) = "solLSling"
SolCallback(18) = "solRSling"
SolCallback(19) = "solGoblin"
SolCallback(20) = "sol3Bank"
SolCallback(22) = "solDivert"
SolCallback(24) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"

'Flashers
'SolCallback(21) = "setlamp 121,"
'SolCallback(23) = "setlamp 123,"
'SolCallback(25) = "setlamp 125,"
'SolCallback(26) = "setlamp 126,"
'SolCallback(27) = "setlamp 127,"
'SolCallback(28) = "setlamp 128,"
'SolCallback(29) = "setlamp 129,"
'SolCallback(30) = "SetLamp 130,"
'SolCallback(31) = "setlamp 131,"
'SolModCallback(3) = "solModDocMagnet"  'TODO
'SolModCallback(3) = "SetModLampmm 0, 200,"

SolModCallBack(21) = "SetModLamp 121,"
SolModCallBack(23) = "SetModLamp 123,"
SolModCallBack(25) = "SetModLamp 125,"
SolModCallBack(26) = "SetModLamp 126,"
SolModCallBack(27) = "SetModLamp 127,"
SolModCallBack(28) = "SetModLamp 128,"
SolModCallBack(29) = "SetModLamp 129,"
SolModCallBack(30) = "SetModLamp 130,"
SolModCallBack(31) = "SetModLamp 131,"



'Solenoid Functions
Sub solTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 22
  End If
 End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
  End If
 End Sub


Sub solDocMotor(Enabled)
  If Enabled Then
    If sw63.IsDropped Then
    Controller.Switch(58) = 0
    Controller.Switch(57) = 1
    sw63.IsDropped=0
    AbbassaOctopus
  Else
    Controller.Switch(57) = 0
    Controller.Switch(58) = 1
    If Attendi=0 Then AlzaOctopus
    End If
  End If
 End Sub

Sub solSandMotor(Enabled)
  If Enabled Then
    If sw42.IsDropped Then
    Controller.Switch(54) =0
    Controller.Switch(53) =1
    sw42.IsDropped=0
    AbbassaSandman
  Else
    Controller.Switch(53) =0
    Controller.Switch(54) =1
    AlzaSandman
    End If
  End If
 End Sub


Sub solGoblin(Enabled)
  If Enabled Then
    ShakeGoblin
  End If
 End Sub


Sub Sol3Bank(Enabled)
  If Enabled Then
    If sw11.IsDropped Then
    Controller.Switch(49)=0
    Controller.Switch(50)=1
    ParetiBankSu
    AlzaBank
  Else
    Controller.Switch(50)=0
    Controller.Switch(49)=1
    AbbassaBank
    End If
  End If
End Sub

Sub solSandVUK(Enabled)
  If Enabled Then
    bsSandman.ExitSol_On
    SolenoideSandmanAbilitato
    Playsound "SolenoideFuori"
  End If
 End Sub

Sub SolenoideSandmanAbilitato
  SolenoideSandman.TransZ= -50
  SolenoideUscitaSandman.enabled=1
End Sub

Sub SolenoideUscitaSandman_Timer 'GESTIONE PISTONE LANCIO PALLA Octopus
  SolenoideSandman.TransZ= -59
  Playsound "SolenoideDentro"
  Me.Enabled=0
End Sub

Sub solDocVUK(Enabled)
  If Enabled Then
    bsDocOck.ExitSol_On
    SolenoideOctopusAbilitato
    Playsound "SolenoideFuori"
  End If
 End Sub

Sub SolenoideOctopusAbilitato
  SolenoideOctopus.TransZ= -50
  SolenoideUscitaOctopus.enabled=1
End Sub

Sub SolenoideUscitaOctopus_Timer 'GESTIONE PISTONE LANCIO PALLA Octopus
  SolenoideOctopus.TransZ= -59
  Playsound "SolenoideDentro"
  Me.Enabled=0
End Sub

Dim LockAttivo

Sub solDivert(Enabled)
  If Enabled Then
    Diverter.IsDropped = 0
    Playsound SoundFX("diverter",DOFContactors)
    LockAttivo = 1
  Else
    Diverter.IsDropped = 1
    Playsound SoundFX("diverter",DOFContactors)
    LockAttivo=0
  End If
 End Sub

'Drains and Kickers
Sub drain_Hit()
  PalleInGioco = PalleInGioco - 1
  bsTrough.AddBall Me
  PlaySound "drain"

 End Sub

Sub BallRelease_UnHit()
  PalleInGioco = PalleInGioco + 1
 End Sub

'************************************************
'************Slingshots Animation****************
'************************************************

Dim RStep, Lstep

Sub LeftSlingShot_Slingshot: vpmTimer.PulseSw 26: End Sub
Sub RightSlingShot_Slingshot: vpmTimer.PulseSw 27: End Sub

Sub solLSling(enabled)
  If enabled then
    PlaySound SoundFX ("SlingshotLeft",DOFContactors)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling1.TransZ = -27
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  End If
End Sub

Sub solRSling(enabled)
  If enabled then
    PlaySound SoundFX ("SlingshotRight",DOFContactors)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling2.TransZ = -27
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  End If
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -15
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -15
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'************************************************
'**************Bumpers Animation*****************
'************************************************

Dim dirRing1:dirRing1 = -1
Dim dirRing2:dirRing2 = -1
Dim dirRing3:dirRing3 = -1

Sub Bumper1_Hit:vpmTimer.PulseSw 30:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 31:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 32:End Sub

Sub solLBump (enabled): If enabled then Bumper1.TimerEnabled = 1: PlaySound SoundFX("BumperSinistro",DOFContactors): End If: End Sub
Sub solRBump (enabled): If enabled then Bumper2.TimerEnabled = 1: PlaySound SoundFX("BumperDestro",DOFContactors): End If: End Sub
Sub solBBump (enabled): If enabled then Bumper3.TimerEnabled = 1: PlaySound SoundFX("BumperCentrale",DOFContactors): End If: End Sub

Sub Bumper1_timer()
  BumperRing1.Z = BumperRing1.Z + (5 * dirRing1)
  If BumperRing1.Z <= -40 Then dirRing1 = 1
  If BumperRing1.Z >= 0 Then
    dirRing1 = -1
    BumperRing1.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper2_timer()
  BumperRing2.Z = BumperRing2.Z + (5 * dirRing2)
  If BumperRing2.Z <= -40 Then dirRing2 = 1
  If BumperRing2.Z >= 0 Then
    dirRing2 = -1
    BumperRing2.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper3_timer()
  BumperRing3.Z = BumperRing3.Z + (5 * dirRing3)
  If BumperRing3.Z <= -40 Then dirRing3 = 1
  If BumperRing3.Z >= 0 Then
    dirRing3 = -1
    BumperRing3.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

'Rollovers
Sub sw23_Hit:Controller.Switch(23) = 1:PlaySound "rollover":End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:Playsound "plunger2":End Sub

'Lower Lanes
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySound "rollover":End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySound "rollover":End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:PlaySound "rollover":End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:PlaySound "rollover":End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

'Upper Lanes
Sub sw8_Hit:Controller.Switch(8) = 1:PlaySound "rollover"
End Sub
Sub sw8_UnHit:Controller.Switch(8) = 0:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1:PlaySound "rollover":End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw34_Hit:Controller.Switch(34) = 1:PlaySound "rollover":End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySound "rollover":End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

'Right
Sub sw37_Hit:Controller.Switch(37) = 1:PlaySound "rollover":End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySound "rollover"
End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

'Right Under Flipper
Sub sw46_Hit:Controller.Switch(46) = 1:PlaySound "rollover":End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

'Spinner
Sub sw7_Spin:vpmTimer.PulseSw 7:PlaySound "spinner":End Sub

'Right Ramp
Sub sw44_Hit:Controller.Switch(44) = 1:Playsound "gate":End Sub

Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:Playsound "gate"
End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

'Left Ramp
Sub sw47_Hit:Controller.Switch(47) = 1:Playsound "gate":End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:Playsound "gate"
End Sub

Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw49_Hit
  Playsound "gate"
End Sub

'Venom da rampa venom
Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySound SoundFX("target",DOFContactors)
End Sub

Sub sw43a_Hit
  Playsound "target"
End Sub


'Doc Ock
Sub S63_Hit:Controller.Switch(63)=1:End Sub

Sub S63_unHit:Controller.Switch(63)=0:End Sub

'Sandman
Sub s42_Hit:Controller.Switch(42)=1:End Sub
Sub s42_unHit:Controller.Switch(42)=0:End Sub
'Lock Opto
Sub sw6_Hit:vpmTimer.PulseSw 6
  sw6p.TransX = -5:sw6.TimerEnabled = 1: PlaySound SoundFX("target",DOFContactors)
End Sub

Sub sw6_Timer:sw6p.TransX = 0: Me.TimerEnabled = 0: End Sub

Dim SandmanPronto

Sub SandmanOK_Hit
  SandmanPronto=1
  If ActiveBall.velY > 2  Then    'ball is going up
  sw59.enabled=1
  Else
  sw59.enabled=0
  End If
End Sub

Sub OctopusOK_Hit
  If ActiveBall.velY > 2  Then    'ball is going up
  sw36.enabled=1
  Else
  sw36.enabled=0
  End If
End Sub

Sub RHP1_Hit()
If ActiveBall.velY < 0  Then    'ball is going up
PlaySound "rrenter"
Else
StopSound "rrenter"
End If
End Sub

'Sandman Optos

Sub sw9_Hit
  PlaySound SoundFX("target",DOFContactors)
  If SandmanPronto=0 Then vpmTimer.PulseSw 9
  Bank.TransX = 5
  BankColpito.Enabled = 1
End Sub

Sub sw10_Hit
  PlaySound SoundFX("target",DOFContactors)
  If SandmanPronto=0 Then vpmTimer.PulseSw 10
  Bank.TransX = 5
  BankColpito.Enabled = 1
End Sub

Sub sw11_Hit
  PlaySound SoundFX("target",DOFContactors)
  If SandmanPronto=0 Then vpmTimer.PulseSw 11
  Bank.TransX = 5
  BankColpito.Enabled = 1
End Sub

Sub BankColpito_Timer()
  Bank.TransX = -0
  Me.Enabled = 0
End Sub

Sub sw12_Hit:vpmTimer.PulseSw 12
  sw12p.TransX = -5:sw12.TimerEnabled = 1:PlaySound SoundFX("target",DOFContactors)
End Sub

Sub sw12_Timer:sw12p.TransX = 0:Me.TimerEnabled = 0: End Sub

Sub sw13_Hit:vpmTimer.PulseSw 13
  sw13p.TransX = -5:sw13.TimerEnabled = 1:PlaySound SoundFX("target",DOFContactors)
End Sub

Sub sw13_Timer:sw13p.TransX = 0:Me.TimerEnabled = 0:End Sub

'Green Goblin Optos
Sub sw1_Hit:vpmTimer.PulseSw 1
  sw1p.TransX = -5: sw1.TimerEnabled = 1:PlaySound SoundFX("target",DOFContactors)
End Sub

Sub sw1_Timer:sw1p.TransX = 0:Me.TimerEnabled = 0:End Sub

Sub sw2_Hit:vpmTimer.PulseSw 2
  sw2p.TransX = -5: sw2.TimerEnabled = 1:PlaySound SoundFX("target",DOFContactors)
End Sub

Sub sw2_Timer:sw2p.TransX = 0:Me.TimerEnabled = 0:End Sub

Sub sw3_Hit:vpmTimer.PulseSw 3
  sw3p.TransX = -5: sw3.TimerEnabled = 1:PlaySound SoundFX("target",DOFContactors)
End Sub

Sub sw3_Timer:sw3p.TransX = 0:Me.TimerEnabled = 0:End Sub

Sub sw4_Hit:vpmTimer.PulseSw 4
  sw4p.TransX = -5: sw4.TimerEnabled = 1:PlaySound SoundFX("target",DOFContactors)
End Sub

Sub sw4_Timer:sw4p.TransX = 0:Me.TimerEnabled = 0:End Sub

Sub sw5_Hit:vpmTimer.PulseSw 5
  sw5p.TransX = -5: sw5.TimerEnabled = 1:PlaySound SoundFX("target",DOFContactors)
End Sub

Sub sw5_Timer:sw5p.TransX = 0:Me.TimerEnabled = 0:End Sub

'Right 3Bank Optos
Sub sw39_Hit:vpmTimer.PulseSw 39
  sw39p.TransX = -5: sw39.TimerEnabled = 1:PlaySound SoundFX("target",DOFContactors)
End Sub

Sub sw39_Timer:sw39p.TransX = 0:sw39.TimerEnabled = 0:End Sub

Sub sw40_Hit:vpmTimer.PulseSw 40
  sw40p.TransX = -5: sw40.TimerEnabled = 1:PlaySound SoundFX("target",DOFContactors)
End Sub

Sub sw40_Timer:sw40p.TransX = 0:sw40.TimerEnabled = 0:End Sub

Sub sw41_Hit:vpmTimer.PulseSw 41
  sw41p.TransX = -5: sw41.TimerEnabled = 1:PlaySound SoundFX("target",DOFContactors)
End Sub

Sub sw41_Timer:sw41p.TransX = 0:sw41.TimerEnabled = 0:End Sub

'Sub PallaBloccata_Hit: vpmTimer.PulseSw 9:vpmTimer.PulseSw 10:vpmTimer.PulseSw 11:End Sub

'Switch 14
Sub RPostColl21_Hit():vpmTimer.PulseSw 14:End Sub

'Sandman VUK

Sub MuroSandman_Hit()
  Playsound "Parete"
End Sub

Sub sw59Abilitato_Hit()
  Playsound "EnterHole"
  If PallaBucoSandman=0 Then
  sw59.Enabled=1
  End If
End Sub

Sub sw59Abilitato_UnHit()
  If PallaBucoSandman=0 Then
  sw59.Enabled=1
  End If
End Sub

Sub sw59_UnHit()  'Uscita buco sandman
  SandmanPronto=0
  PallaBucoSandman=0
  Playsound "popper"
  sw59.Enabled=0
End Sub

Sub sw59_Hit()  'Entrata buco sandman
  bsSandman.AddBall Me
  PallaBucoSandman=1
End Sub

'DocOck VUK

Dim PallaBucoOctopus, PallaBucoSandman

Sub MuroOctopus_Hit()
  Playsound "Parete"
End Sub

Sub sw36Abilitato_UnHit()
  sw36.Enabled=1
End Sub

Sub sw36Abilitato_Hit()
  Playsound "EnterHole"
End Sub

Sub sw36_UnHit()  'Buco doc uscita
  PallaBucoOctopus=0
  sw36.Enabled=0
End Sub

Sub sw36_Hit()  'Buco doc entrata
  bsDocOck.AddBall Me
  PallaBucoOctopus=1
 End Sub

 '**************
 ' Flipper Subs
 '**************

 Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX (SFlipperOn,DOFContactors):LeftFlipper.RotateToEnd

     Else
         PlaySound SoundFX (SFlipperOff,DOFContactors):LeftFlipper.RotateToStart

     End If
 End Sub

 Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX (SFlipperOn,DOFContactors):RightFlipper.RotateToEnd

     Else
         PlaySound SoundFX (SFlipperOff,DOFContactors):RightFlipper.RotateToStart

     End If
 End Sub

 Sub SolURFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX (SFlipperOn,DOFContactors):RightFlipper2.RotateToEnd

     Else
         PlaySound SoundFX (SFlipperOff,DOFContactors):RightFlipper2.RotateToStart

     End If
 End Sub

'*********************************
' GESTIONE MOVIMENTAZIONE GOBLIN
'*********************************

Dim GoblinPos

Sub ShakeGoblin
GoblinPos = 8
GoblinShakeTimer.Enabled = 1
End Sub

Sub GoblinShakeTimer_Timer
  Playsound "shake"
    Goblin.TransZ = GoblinPos
    If GoblinPos = 0 Then GoblinShakeTimer.Enabled = 0:Exit Sub
    If GoblinPos < 0 Then
        GoblinPos = ABS(GoblinPos) - 1
    Else
        GoblinPos = - GoblinPos + 1
    End If
End Sub

Sub AccendiGoblin
  If LampState (76) = 0 Then
  IlluminaGoblin.Visible=0
  Else
  If LampState (76) > 0 Then
  IlluminaGoblin.Visible=1
  End If
  End If
End Sub

Sub AccendiVenom
  If LampState (75) = 0 Then
   IlluminaVenom.Visible= 0
  Else
  If LampState (75) > 0 Then
   IlluminaVenom.Visible=1
  End If
  End If
End Sub

'*********************************
' GESTIONE MOVIMENTAZIONE OCTOPUS
'*********************************

Dim OctopusDir, OctopusPos

Sub sw63_Hit()
  Playsound "bersaglio"
  Octopus.TransZ = 5
  BracketOctopus.TransZ = 5
  OctopusColpito.Enabled = 1
End Sub

Sub Fotocellula1Disattiva_Hit()
  FotocellulaOctopus.Visible=0
End Sub

Sub Fotocellula1Disattiva_UnHit()
  TimerFotocellula1.enabled=1
End Sub

Sub TimerFotocellula1_Timer()
  FotocellulaOctopus.Visible=1
  Me.Enabled=0
End Sub

Sub Fotocellula2Disattiva_Hit()
  FotocellulaSandmanDx.Visible=0
  FotocellulaSandmanSx.Visible=0
End Sub

Sub Fotocellula2Disattiva_UnHit()
  TimerFotocellula2.enabled=1
End Sub

Sub TimerFotocellula2_Timer()
  FotocellulaSandmanDx.Visible=1
  FotocellulaSandmanSx.Visible=1
  Me.Enabled=0
End Sub

Sub Fotocellula3Disattiva_Hit()
  FotocellulaVenom.Visible=0
End Sub

Sub Fotocellula3Disattiva_UnHit()
  TimerFotocellula3.enabled=1
End Sub

Sub TimerFotocellula3_Timer()
  FotocellulaVenom.Visible=1
  Me.Enabled=0
End Sub

Sub OctopusColpito_Timer()
  Octopus.TransZ = 0
  BracketOctopus.TransZ = 0
  Me.Enabled = 0
End Sub

Sub AlzaOctopus
  OctopusDir = -1 ' removing 1 will make Oct go up
  OctopusTimer.Enabled = 1
End Sub

Sub AbbassaOctopus
  OctopusDir = 1 ' adding 1 will make Oct to go down by one step
  OctopusTimer.Enabled = 1
End Sub


Sub OctopusTimer_Timer
  Playsound "Motor"
  Octopus.TransY = -OctopusPos
  BracketOctopus.TransY = -OctopusPos
  OctopusPos = OctopusPos + OctopusDir
  If OctopusPos < 0 Then OctopusPos=0: sw63.IsDropped=1: Me.Enabled = 0
  If OctopusPos > 50 Then OctopusPos = 50: Me.Enabled = 0
End Sub


Sub AccendiOctopus
  If LampState (77) = 0  Then
  IlluminaOctopus.Visible=0
  Else
  If LampState (77) > 0 Then
  IlluminaOctopus.Visible=1
  End If
  End If
End Sub


'*********************************
' GESTIONE MOVIMENTAZIONE SANDMAN
'*********************************

Dim SandmanDir, SandmanPos, SandmanAlto

Sub sw42_Hit()
  Playsound "bersaglio"
  Sandman.TransZ = 5
  BracketSandman.TransZ = 5
  SandmanColpito.Enabled = 1
End Sub

Sub SandmanColpito_Timer()
  Sandman.TransZ = 0
  BracketSandman.TransZ = 0
  Me.Enabled = 0
End Sub

Sub AlzaSandman
  SandmanDir = 1 ' removing 1 will make Oct go up
  SandmanTimer.Enabled = 1
  SandmanAlto=1
End Sub

Sub AbbassaSandman
  SandmanDir = -1 ' adding 1 will make Oct to go down by one step
  SandmanTimer.Enabled = 1
  SandmanAlto=0
End Sub


Sub SandmanTimer_Timer
  Playsound "Motor"
  Sandman.TransY = SandmanPos
  BracketSandman.TransY = SandmanPos
  SandmanPos = SandmanPos + SandmanDir
  If SandmanPos < 0 Then SandmanPos=0: Me.Enabled = 0
  If SandmanPos > 50 Then SandmanPos = 50: sw42.IsDropped=1: Me.Enabled = 0
End Sub

Sub AccendiSandman
  If LampState (74) = 0 Then
  IlluminaSandman.Visible=0
  Else
  If LampState (74) > 0 Then
  IlluminaSandman.Visible=1
  End If
  End If
End Sub


'***********************************
' GESTIONE MOVIMENTAZIONE MURO BABK
'***********************************

Dim BankDir, BankPos

Sub AlzaBank
  BankDir = -1
  BankTimer.Enabled = 1
End Sub

Sub AbbassaBank
  BankDir = 1
  BankTimer.Enabled = 1
End Sub


Sub BankTimer_Timer
  Bank.TransY = -BankPos
  BankPos = BankPos + BankDir
  If BankPos < 0 Then BankPos = 0: Me.Enabled = 0
  If BankPos > 52 Then BankPos = 52: ParetiBankGiu: Me.Enabled = 0
End Sub

Dim BankAlto

Sub PallaBloccata_Hit()
  If BankAlto=1 AND Attendi=0 Then
  Controller.Switch(50)=0
  Controller.Switch(49)=1
  AbbassaBank
  ParetiBankGiu
  SbloccaPallaBank.Enabled= True
  End If
End Sub

Sub SbloccaPallaBank_Timer()
  If Attendi=0 Then
  Controller.Switch(49)=0
  Controller.Switch(50)=1
  AlzaBank
  ParetiBankSu
  Me.Enabled=0
  End If
End Sub

Sub ParetiBankSu
  sw9.IsDropped=0
  sw10.IsDropped=0
  sw11.IsDropped=0
  BankAlto=1
End Sub

Sub ParetiBankGiu
  sw9.IsDropped=1
  sw10.IsDropped=1
  sw11.IsDropped=1
  BankAlto=0
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************


Sub SetModLamp(nr, value)
    If value <> SolModValue(nr) Then
    SolModValue(nr) = value
    if value > 0 then LampState(nr) = 1 else LampState(nr) = 0
        FadingLevel(nr) = LampState(nr) + 4
    End If
End Sub

Sub SetModLampM(nr, value)  'setlamp NR, but also NR + 50
    If value <> SolModValue(nr) Then
    SolModValue(nr) = value
    if value > 0 then LampState(nr) = 1 else LampState(nr) = 0
        FadingLevel(nr) = LampState(nr) + 4
    End If
    If value <> SolModValue(nr+50) Then
    SolModValue(nr+50) = value
    if value > 0 then LampState(nr+50) = 1 else LampState(nr+50) = 0
        FadingLevel(nr+50) = LampState(nr+50) + 4
    End If
End Sub


Sub SetModLampMM(nr, nr2, value)  'setlamp two NRs
    If value <> SolModValue(nr) Then
    SolModValue(nr) = value
    if value > 0 then LampState(nr) = 1 else LampState(nr) = 0
        FadingLevel(nr) = LampState(nr) + 4
    End If
    If value <> SolModValue(nr2) Then
    SolModValue(nr2) = value
    if value > 0 then LampState(nr2) = 1 else LampState(nr2) = 0
        FadingLevel(nr2) = LampState(nr2) + 4
    End If
End Sub

Sub SetModLampMMM(nr, nr2, nr3, value)  'setlamp 3 NRs
    If value <> SolModValue(nr) Then
    SolModValue(nr) = value
    if value > 0 then LampState(nr) = 1 else LampState(nr) = 0
        FadingLevel(nr) = LampState(nr) + 4
    End If
    If value <> SolModValue(nr2) Then
    SolModValue(nr2) = value
    if value > 0 then LampState(nr2) = 1 else LampState(nr2) = 0
        FadingLevel(nr2) = LampState(nr2) + 4
    End If
    If value <> SolModValue(nr3) Then
    SolModValue(nr3) = value
    if value > 0 then LampState(nr3) = 1 else LampState(nr3) = 0
        FadingLevel(nr3) = LampState(nr3) + 4
    End If
End Sub



Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)
            ' turn off the lights and flashers and reset them to the default parameters
dim SolModValue(200)
dim LightFallOff(200, 4)  '2d array to hold alt falloff values in different columns
dim FlashersOpacity(200)
dim FlashersFalloff(200)  '??? (could use multiply? or some other kind of mixing?...)
dim GIscale(200)
InitLamps()
Sub InitLamps()
  Dim x
  For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.01    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.008 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    giscale(x) = 1

    SolModValue(x) = 0
    FlashersOpacity(x) = 0
'   FlashersFalloff(x) = 0  '????
    LightFallOff(x, 0) = 0
    LightFallOff(x, 1) = 0
    LightFallOff(x, 2) = 0
    LightFallOff(x, 3) = 0
    Next
  for x = 100 to 200  'flashers
        FlashSpeedUp(x) = 1.1 '0.4  ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.9 '0.2 ' slower speed when turning off the flasher
  next

  'SetMagnet
  'LightFallOff(123, 0) = f23.falloff
  'LightFallOff(123, 1) = f23a.falloff

  LightFallOff(125, 0) = f25.falloff
  LightFallOff(125, 1) = f25a.falloff

  'LightFallOff(126, 0) = f26.falloff

  'f23.state = 1
  'f23a.state = 1
  f25.state = 1
  f25a.state = 1
  'f26.state = 1

End Sub


'NmodLight subs:  'lampnumber, object, falloff column, ScaleType (see Function ScaleLights below)
dim CGT 'compensated game time
Sub LampTimer_Timer()
  cgt = gametime - InitFadeTime(0)
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next

    End If
  UpdateLamps

  nModFlash 121, FlashOctopus, 0, 9, 1

  'nModLight 123, f23, 0, 0, 1
  'nModLightm 123, F23a, 0, 0

  nModLight 125, f25, 0, 0, 1
  nModLightm 125, f25a, 0, 0

  'nModLight 126, f26, 0, 0, 1



  nModFlash 127, FlashGiallo,   0, 9, 1 'nr, object, offset (not used for Flashers), scaletype, offscale
  nModFlash 128, FlasherGoblin,   0, 0, 1
  nModFlash 129, FlashBlu,    0, 9, 1
  nModFlash 130, FlashRosso,    0, 9, 1
  nModFlash 131, FlasherBumpers,  0, 0, 1

  InitFadeTime(0) = gametime
End Sub

sub f127t_timer 'debug
  me.text = "SolModValue" & SolModValue(127) & vbnewline & _
    "fadinglevel" & FadingLevel(127) & vbnewline & _
    "FlashLevel" & FlashLevel(127) & vbnewline & _
    "Intensityscale" & FlashGiallo.IntensityScale & vbnewline & _
    " "
end sub


Function ScaleLights(value, scaletype)  'returns an intensityscale-friendly 0->100% value out of 255
  dim i
  Select Case scaletype 'select case because bad at maths   'TODO: Simplify these functions. B/c this is absurdly bad.
    case 0
      i = value * (1 / 255) '0 to 1
    case 6  '0.0625 to 1
      i = (value + 17)/272
    case 9  '0.089 to 1
      i = (value + 25)/280
    case 15
      i = (value / 300) + 0.15
    case 20
      i = (4 * value)/1275 + (1/5)
    case 25
      i = (value + 85) / 340
    case 37 '0.375 to 1
      i = (value+153) / 408
    case 40
      i = (value + 170) / 425
    case 50
      i = (value + 255) / 510 '0.5 to 1
    case 75
      i = (value + 765) / 1020  '0.75 to 1
    case Else
      i = 10
  End Select
  ScaleLights = i
End Function

Function ScaleByte(value, scaletype)  'returns a number between 1 and 255
  dim i
  Select Case scaletype
    case 0
      i = value * 1 '0 to 1
    case 9  'ugh
      i = (5*(200*value + 1887))/1037
    case 15
      i = (16*value)/17 + 15
    case else
      i = (3*(value + 85))/4  '63.75 to 255
  End Select
  ScaleByte = i
End Function

Function ScaleGI(value, scaletype)  'returns an intensityscale-friendly 0->100% value out of 1>8 'it does go to 8
  dim i
  Select Case scaletype 'select case because bad at maths
    case 0
      i = value * (1/8) '0 to 1
    case 25
      i = (1/28)*(3*value + 4)
    case 50
      i = (value+5)/12
    case else
'     x = (4*value)/3 - 85  '63.75 to 255

  End Select
  ScaleGI = i
End Function


Function ScaleFalloff(value, nr)  'TODO make more options here
  if nr > 128 then 'do not scale special bulb NRs
    ScaleFalloff = 1
  Else
'   ScaleFalloff = (value + 255) / 510  '0.5 to 1
    ScaleFalloff = (value + 765) / 1020 '0.75 to 1
  end if
End Function

dim InitFadeTime(200)

Sub nModLight(nr, object, offset, scaletype, offscale)  'Fading using intensityscale with modulated callbacks
  dim DesiredFading
  Select Case FadingLevel(nr)
    case 3  'workaround - wait a frame to let M sub finish fading
      FadingLevel(nr) = 0
    Case 4  'off
'     FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)*offscale
      FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * cgt ) * offscale
      If FlashLevel(nr) < 0 then FlashLevel(nr) = 0 : FadingLevel(nr) = 3
      Object.IntensityScale = ScaleLights(FlashLevel(nr),0 ) * GIscale(nr)
      Object.Falloff = LightFallOff(nr, offset) * ScaleFalloff(FlashLevel(nr), nr)
'     InitFadeTime(nr) = gametime
'     tbt.text = (cgt - InitFadeTime(0) )
    Case 5 ' Fade (Dynamic)
      DesiredFading = ScaleByte(SolModValue(nr), scaletype)

      if FlashLevel(nr) < DesiredFading Then
'       tb5.text = "+"
        'FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
        FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * cgt )
        If FlashLevel(nr) >= DesiredFading Then FlashLevel(nr) = DesiredFading : FadingLevel(nr) = 1
      elseif FlashLevel(nr) > DesiredFading Then
'       tb5.text = "-"
'       FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
        FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * cgt )
        If FlashLevel(nr) <= DesiredFading Then FlashLevel(nr) = DesiredFading : FadingLevel(nr) = 1
      End If
      Object.Intensityscale = ScaleLights(FlashLevel(nr),0 ) * GIscale(nr)
      Object.Falloff = LightFallOff(nr, offset) * ScaleFalloff(FlashLevel(nr), nr)
'     tbt.text = (cgt - InitFadeTime(0) )
'     InitFadeTime(nr) = gametime
'     tbt.text = (FlashSpeedDown(nr) * cgt  ) & vbnewline & (FlashSpeedup(nr) * cgt ) & "cgt:" & cgt
  End Select
End Sub

Sub nModFlash(nr, object, offset, scaletype, offscale)  'Fading using intensityscale with modulated callbacks 'gametime compensated
  dim DesiredFading
  Select Case FadingLevel(nr)
    case 3  'workaround - wait a frame to let M sub finish fading
      FadingLevel(nr) = 0
    Case 4  'off
'     FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)*offscale
      FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * cgt ) * offscale
      If FlashLevel(nr) < 0 then FlashLevel(nr) = 0 : FadingLevel(nr) = 3
      Object.IntensityScale = ScaleLights(FlashLevel(nr),0 ) * GIscale(nr)'     Object.Falloff = LightFallOff(nr, offset) * ScaleFalloff(FlashLevel(nr), nr)
'     InitFadeTime(nr) = gametime
'     tbt.text = (cgt - InitFadeTime(0) )
    Case 5 ' Fade (Dynamic)
      DesiredFading = ScaleByte(SolModValue(nr), scaletype)
'     tb1.text = DesiredFading & " ...5"
      if FlashLevel(nr) < DesiredFading Then
'       tb5.text = "+"
        'FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
        FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * cgt )
        If FlashLevel(nr) >= DesiredFading Then FlashLevel(nr) = DesiredFading : FadingLevel(nr) = 1
      elseif FlashLevel(nr) > DesiredFading Then
'       tb5.text = "-"
'       FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
        FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * cgt )
        If FlashLevel(nr) <= DesiredFading Then FlashLevel(nr) = DesiredFading : FadingLevel(nr) = 1
      End If
      Object.Intensityscale = ScaleLights(FlashLevel(nr),0 ) * GIscale(nr)
'     Object.Falloff = LightFallOff(nr, offset) * ScaleFalloff(FlashLevel(nr), nr)
'     tbt.text = (cgt - InitFadeTime(0) )
'     InitFadeTime(nr) = gametime
'     tbt.text = (FlashSpeedDown(nr) * cgt  ) & vbnewline & (FlashSpeedup(nr) * cgt ) & "cgt:" & cgt
'     tbt.text = DesiredFading
  End Select
End Sub

Sub nModLightM(nr, Object, offset, scaletype) 'uses offset to store different falloff values in a unused lamp number. default 0
  Select Case FadingLevel(nr)
    Case 3, 4, 5
      Object.Intensityscale = ScaleLights(FlashLevel(nr),0 ) * GIscale(nr)
'     Object.IntensityScale = ScaleLights(FlashLevel(nr),scaletype ) * GIscale(nr)
      Object.Falloff = LightFallOff(nr, offset) * ScaleFalloff(FlashLevel(nr), nr)
  End Select
End Sub

Sub nModFlashM(nr, Object, offset, scaletype)
  Select Case FadingLevel(nr)
    Case 3, 4, 5
      Object.Intensityscale = ScaleLights(FlashLevel(nr),0 ) * GIscale(nr)
  End Select
End Sub

Sub Flashc(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * CGT)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
               FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * CGT)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Gate2_Hit()
  Playsound "gate"
End Sub

Sub Gate3_Hit()
  Playsound "gate"
End Sub


Sub AccendiLuciWall
  LuceWall1.Visible=1
  LuceWall2.Visible=1
  LuceWall3.Visible=1
  LuceWall4.Visible=1
  LuceWall5.Visible=1
  LuceWall6.Visible=1
  LuceWall7.Visible=1
  LuceWall8.Visible=1
End Sub

Sub SpegniLuciWall
  LuceWall1.Visible=0
  LuceWall2.Visible=0
  LuceWall3.Visible=0
  LuceWall4.Visible=0
  LuceWall5.Visible=0
  LuceWall6.Visible=0
  LuceWall7.Visible=0
  LuceWall8.Visible=0
End Sub


 Sub UpdateLamps
  NFadeL 3, l3
  NFadeL 4, l4
  NFadeL 5, l5
  NFadeL 6, l6
  NFadeL 7, l7
  NFadeL 8, l8
  NFadeL 9, l9
  NFadeL 10, l10
  NFadeL 11, l11
  NFadeL 12, l12
  NFadeL 13, l13
  NFadeL 14, l14
  NFadeL 15, l15
  NFadeL 16, l16
  NFadeL 17, l17
  NFadeL 18, l18
  NFadeL 19, l19
  NFadeL 20, l20
  NFadeL 21, l21
  NFadeL 22, l22
  NFadeL 23, l23
  NFadeL 24, l24
  NFadeL 25, l25
  NFadeL 26, l26
  NFadeL 27, l27
  NFadeL 28, l28
  NFadeL 29, l29
  NFadeL 30, l30
  NFadeL 31, l31
  NFadeL 32, l32
  NFadeL 33, l33
  NFadeL 34, l34
  NFadeL 35, l35
  NFadeL 36, l36
  NFadeL 37, l37
  NFadeL 38, l38
  NFadeL 39, l39
  NFadeL 40, l40
  NFadeLm 41, l41
  Flashc 41, l41r
  NFadeL 42, l42
  NFadeL 43, l43
  NFadeLm 44, l44
  Flashc 44, l44r
  NFadeL 45, l45
  NFadeL 46, l46
  NFadeL 47, l47
  NFadeL 48, l48
  NFadeL 49, l49
  NFadeL 50, l50
  NFadeL 51, l51
  NFadeLm 52, l52
  Flashc 52, l52r
  NFadeLm 53, l53
  Flashc 53, l53r
  NFadeLm 54, l54
  Flashc 54, l54r
  NFadeL 57, l57
  NFadeL 58, l58
  NFadeL 59, l59
  NFadeLm 60, l60
  NFadeL 60, l60a
  NFadeLm 61, l61
  NFadeL 61, l61a
  NFadeLm 62, l62
  NFadeL 62, l62a
  NFadeL 63, l63
  NFadeL 64, l64
  NFadeL 65, l65
  Flashc 66, MissioneVenom3
  Flashc 67, MissioneVenom2
  Flashc 68, MissioneVenom1
  Flashc 69, MissioneSandman3
  Flashc 70, MissioneSandman2
  Flashc 71, MissioneSandman1
  NFadeLm 72, l72
  Flashc 72, l72r
  NFadeLm 74, l74: AccendiSandman
  NFadeLm 75, l75: AccendiVenom
  NFadeLm 76, l76: AccendiGoblin
  NFadeLm 77, l77: AccendiOctopus
  NFadeL 78, l78
'Flashers
' Flash 121, FlashOctopus
' NfadeLm 123, F23
' NfadeL 123, F23a
' NfadeLm 125, F25
' NfadeL 125, F25a
' NfadeL 126, F26
' 'NfadeLm 127, F27a
' 'NfadeLm 127, F27b
' Flash 127, FlashGiallo

' Flash 128, FlasherGoblin
' 'NfadeLm 129, F29a
' 'NfadeLm 129, F29b
' Flash 129, FlashBlu
' Flash 130, FlashRosso
' Flash 131, FlasherBumpers
' 'NfadeLm 131, F31
' 'NfadeLm 131, F31a
' 'NfadeL 131, F31b
 End Sub

Sub SetLamp(nr, value)
  If value = 0 AND LampState(nr) = 0 Then Exit Sub
  If value = 1 AND LampState(nr) = 1 Then Exit Sub
  If value <> LampState(nr) Then
    LampState(nr) = abs(value)
    FadingLevel(nr) = abs(value) + 4
    End If
End Sub


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

' *********************************************************************
'                JP's Supporting Ball & Sound Functions
' *********************************************************************

Sub ShooterEnd_UnHit():If activeball.z > 30  Then vpmTimer.AddTimer 150, "BallHitSound":End If:End Sub

Sub BallHitSound(dummy):PlaySound "ballhit":End Sub

Sub Rhelp1_Hit()
   ActiveBall.VelZ = -2
     ActiveBall.VelY = 0
     ActiveBall.VelX = 0
   StopSound "metalrolling"
   StopSound "metalrolling"
   sw16Premuto=0
   Playsound "ballrampdrop"
 End Sub

Sub Rhelp2_Hit()
   ActiveBall.VelZ = -2
     ActiveBall.VelY = 0
     ActiveBall.VelX = 0
   StopSound "metalrolling"
   StopSound "metalrolling"
   sw17Premuto=0
   Playsound "ballrampdrop"
 End Sub

Dim sw16Premuto

Sub sw16_Hit 'Venom da rampa sinistra
  Playsound "metalrolling"
  sw16Premuto= 1
End Sub

Dim sw17Premuto

Sub sw17_Hit
  Playsound "metalrolling"
  sw17Premuto=1
End Sub

Sub DocVUKexit_Hit
  If sw17Premuto= 0 Then Playsound "metalrolling"
End Sub

Sub VenomOK_Hit
  If sw16Premuto= 0 Then Playsound "metalrolling"
End Sub


Sub sw100_Hit()
  Playsound "muro"
End Sub

Sub sw101_Hit()
  Playsound "muro"
End Sub

Sub sw102_Hit()
  Playsound "muro"
End Sub

Sub sw103_Hit()
  Playsound "muro"
End Sub

Sub sw104_Hit()
  Playsound "muro"
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

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
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

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper2_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1)
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


'*********** ROLLING SOUND *********************************
Const tnob = 4            ' total number of balls : 4 (trough)
ReDim rolling(tnob-1)
InitRolling

Sub InitRolling:Dim i:For i=0 to (tnob-1):rolling(i) = False:Next:End Sub


'*********** BALL SHADOW *********************************
ReDim BallShadow(tnob-1)
InitBallShadow

Sub InitBallShadow
  Dim i: For i=0 to tnob-1
    ExecuteGlobal "Set BallShadow(" & i & ")=BallShadow" & (i+1) & ":"
  Next
End Sub

Sub BallShadowUpdate()
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
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) - 10
    End If
      ballShadow(b).Y = BOT(b).Y + 20
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

Function RndNum(min,max)
  RndNum = Int(Rnd()*(max-min+1))+min     ' Sets a random number between min and max
End Function

' *********************************************************************
'             Other Sound FX
' *********************************************************************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound "fx_collide", 0, Csng(velocity) ^2 / 50, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

Sub LeftFlipper_Collide(parm)
    RandomSoundFlipper
End Sub

Sub Rightflipper_Collide(parm)
    RandomSoundFlipper
End Sub

Sub Posts_Hit(idx)
  RandomSoundRubber()
End Sub

Sub Rubbers_Hit(idx)
  RandomSoundRubber()
End Sub

Sub RandomSoundRubber()
  Select Case RndNum(1,3)
    Case 1 : PlaySoundAtBallVol "fx_rubber_hit_1",10
    Case 2 : PlaySoundAtBallVol "fx_rubber_hit_2",10
    Case 3 : PlaySoundAtBallVol "fx_rubber_hit_3",10
  End Select
End Sub

Sub RandomSoundFlipper()
  Select Case RndNum(1,3)
    Case 1 : PlaySoundAtBallVol "flip_hit_1", 20
    Case 2 : PlaySoundAtBallVol "flip_hit_2", 20
    Case 3 : PlaySoundAtBallVol "flip_hit_3", 20
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
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Sub RollingSoundUpdate()
    Dim BOT, b
    BOT = GetBalls
  ' stop the sound of deleted balls
  If UBound(BOT)<(tnob - 1) Then
    For b = (UBound(BOT) + 1) to (tnob-1)
      rolling(b) = False
      StopSound("fx_ballrolling" & b+1)
    Next
  End If
  ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
       ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 Then
      rolling(b) = True
      if BOT(b).z < 30 Then ' Ball on playfield
            PlaySound("fx_ballrolling" & b+1), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
      Else ' Ball on raised ramp
            PlaySound("fx_ballrolling" & b+1), -1, Vol(BOT(b) )/5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
    Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b+1)
                rolling(b) = False
            End If
        End If
    Next
End Sub
