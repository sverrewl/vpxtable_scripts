'88888bab  8888888b .d888888  dP     dP dP .d88888b                             dP
'88    `8b 88       d8'    88 88     88 88 88.    "'                            88
'88aaaa8P' 88aaaa   88aaaaa8a 88    .8P 88 `Y88888b.    .d8888b. 88d888b. .d888b88
'88   `8b. 88       88     88 88    d8' 88       `8b    88'  `88 88'  `88 88'  `88
'88    .88 88       88     88 88  .d8P  88 d8'   .8P    88.  .88 88    88 88.  .88
'88888888P 88888888 88     88 888888'   dP  Y88888P     `88888P8 dP    dP `88888P8

'888888ba  dP     dP dd888888PP dd888888PP dP     dP 88888888b  .d888888 888888ba
'88    `8b 88     88     88         88     88     88 88        d8'    88 88    `8b
'88aaaa8P' 88     88     88         88     88aaaaa88 88aaaa    88aaaaa88 88     88
'88   `8b. 88     88     88         88     88     88 88        88     88 88     88
'88    .88 Y8.   .8P     88         88     88     88 88        88     88 88    .8P
'88888888P `Y88888P'     dP         dP     dP     dP 88888888P 88     88 8888888P

' ItItIttIIIIIIIIIIIYIt+tIIIIIYIYIIIYIIIIYYYIIIYIIYYIIIIYIIIIIIIYYIIIYIYi=+RWMWW
' VIIItttItYItIIIIII;,....;IYYIIYIIIItIIIItIIIIIIIIi+ttiii+ttitIIItIIYYYi;+VWWBR
' +tItttiIIIIIYIYI=........,IiitItIIII BY IIIIIII+tRBBBBBBXMBRtIitIYYYYYt;=RWWMB
' iititttttIIIIYY,.............iII watacaractr +tVBBBMBBMBMMBBBBXV+tIIYYt=:RRXRB
' ttttitIIIYYIYI:..............,IItttIIIttIIIIt;RBBBBBBBBBBBBBBBBBYiYYYY+=;VXXXR
' ttItttIIIYYYIi.....,..........=tiii 2022 tII=tBBBBBBBBBBBBBMBBBBBtIYVYi;:++=i+
' ttttttIIIIIIt....;VR+....;i+,..+tttttttittIti+RBXBBMBMBBBMBMBBBBBIIYVYt:+ii+++
' ttttttttttttt,..=BBBBIi,YRBRV:.;ittttttttttIiiiVYIRIYYYiXRBBBBBBBX+IIIItItIIII
' tttttittttIIt,.;RBBBRBBBBBBBBV.=ttttttttttttittt=YVYIYYYIYRRBBBBBYiIIIIIIIItIt
' tttittttIIIII+.=RBBXXRBBBBXXBB+=ttttttttItIttii=RRBBBBBBBXtBBBBBBIttttIIIItIti
' ttttttIIIItIII:iRBBRtiRRX+IXRB:tttItttttIIItIti+RBBBRBBBBBtBBBBBB+tttItttIttIt
' IttttIYIIIItIti+RBBBRRBRYRBBBR=tttttIttIIIItItt=BBBBBXYXRRIBBBBBX=ttttttttittt
' tttttItIIIIIIt=+XRBRi+RBt=YRBt+tItIIItIIIIttItt;RRBVYYBBBVYBBBBB+ttItttttttItt
' ttiittttIIIIItItiBRBYXIYIRBBVitIIIIIItItIIttttt++VIIIXBBBR++iBRIitttttttitItit
' tiiittttttIIItIt=RRBBBBRBBRR+IIIIIYIItttttttttti=iiVVRBBBBBt+RR+tttIttitittttt
' ti++titittItIIIIttRIXYi+iYIViIIYYIYYYIIIItItttt+;YYVRBBRRRBYRttttttttIIttttttt
' ti++i++ttii+i++++=VBRBXVBBR+iiiiiiiti+ittIiii++=Yi=tVRBX+VRBBi=tt+tiitttItitII
' YitIt+++==++++=++++VRXVVYR+==++++++++++i=i+iii++i;YI++i+IRRRR+++++ii+=++++++iY
' Ytt;:=======+==++++;RBBBBX;=++++++++++++=:++++++:+tIVRRBBBBBRi+++++++====+=+iX
' YI+====+=======+=++=IBBBBi==+++++++++++==,++==++;IVIii++IRRRBt::++++++i=+=++tY
' YY+;=====++===;==+;,;RRRV=,:=++;+;++++++=,+++++++=+=++,:,=It=:,,,=++;++i+i=itI
' VV;;=;==+===:;,;;;===;;;====::+:+:+iii+i+:i++++++=:;:.:::::::::::.;::++++==++i
' IV;;;====+===+;:=++=======+==,;;;++++++i;:+++++=++;;;...,.,...,,+.;:=+=+ii;+=i

Option Explicit
Randomize

Const BallSize = 52
Const BallMass = 1.2

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT

Const cGameName= "beav_butt"

LoadVPM "02800000", "beav_butt.VBS", 3.50

'********************
'Standard definitions
'********************
Const UseSolenoids  = 1
Const UseLamps      = 1
Const UseSync   = 1
Const UseGI     = 0

' Standard Sounds
Const SSolenoidOn = "SolenoidOn"
Const SSolenoidOff = "SolenoidOff"
Const SCoin = "fx_Coin"

'******************************************************
'         Initialize the table
'******************************************************

Dim bsSaucer,bsScoop, bsTK, dtBankL, dtBankR, xx

Sub Table1_Init
         Controller.Games("beav_butt").Settings.Value("sound_mode") =1
         With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Beavis & Butt-head (Bally 1993)"
        .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .dip(0) = &H00  'Set DIP to USA
    .Hidden = 0
    On Error Resume Next
    .Run 'GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
    End With

' Thalamus : Was missing 'vpminit me'
vpminit me
    '** Main Timer init
'  PinMAMETimer.Interval = PinMAMEInterval
'  PinMAMETimer.Enabled = 1
    Playsound "B&BWarning"

    '** Nudging
  vpmNudge.TiltSwitch = 151
  vpmNudge.Sensitivity = 4
  vpmNudge.TiltObj = Array(sw10, sw11, sw12, sw13, sw14)

    Set bsScoop = New cvpmBallStack
    With bsScoop
        .InitSaucer Sw52, 52, 120, 0
        .InitExitSnd SoundFX("CenterEject", DOFContactors), SoundFX("CenterEject", DOFContactors)
    End With

    Set bsTK = New cvpmBallStack
    With bsTK
        .InitSaucer Sw53k, 53, 20, 50
        .InitExitSnd SoundFX("LeftEject", DOFContactors), SoundFX("LeftEject", DOFContactors)
        .KickZ = 3.1415926/3
        .CreateEvents "bsTK", sw53k
    End With

    '** DropTargets Bank (Left)
    Set dtBankL = new cvpmDropTarget
  With dtBankL
    .InitDrop Array(sw16,sw26,sw36,sw46),Array(16,26,36,46)
    .Initsnd SoundFX("fx_drop_target_down",DOFDropTargets), SoundFX("fx_drop_target_up",DOFDropTargets)
  End With

    '** DropTargets Bank (Right)
    Set dtBankR = new cvpmDropTarget
  With dtBankR
    .InitDrop Array(sw17,sw27,sw37,sw47),Array(17,27,37,47)
    .Initsnd SoundFX("fx_drop_target_down",DOFDropTargets), SoundFX("fx_drop_target_up",DOFDropTargets)
  End With

    '** Other Suff
  InitRolling:InitLamps:InitTrough:DiverterOn.Collidable=0
  SideRails.visible=DesktopMode
End Sub

Sub kicker1_Hit
   PlaySound "fx_subwayramproll"
    kicker1.Destroyball
    sw53k.CreateBall
    bsTK.AddBall 0
    mtvlight.state = 0
    beavisrampplasticlight.state = 0
    buttheadrampplasticlight.state = 0
End Sub

Sub kicker2_Hit
   PlaySound "fx_subwayramproll"
   vpmtimer.addtimer 1500, "MyLightACounter = 0:Timer001.Enabled = 1:bsscoop.AddBall kicker2 '"
End Sub

'******************************************************
'             Keys
'******************************************************

Sub Table1_KeyDown(ByVal keycode)
  If keycode = LeftTiltKey Then Nudge 90, 5:PlaySoundAt SoundFX("fx_nudge",0), InstructionCard
  If keycode = RightTiltKey Then Nudge 270, 5:PlaySoundAt SoundFX("fx_nudge",0), PriceCard
  If keycode = CenterTiltKey Then Nudge 0, 6:PlaySoundAt SoundFX("fx_nudge",0), Drain
  If keycode = PlungerKey Then Plunger.PullBack:PlaySoundAt "fx_plungerpull",Plunger
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then Plunger.Fire:StopSound "plungerpull":PlaySoundAt "fx_plunger",Plunger
  if vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Table1_Paused:Controller.Pause = True:End Sub
Sub Table1_unPaused:Controller.Pause = False:End Sub
Sub Table1_exit():Controller.Pause = False:Controller.Stop:Controller.Games("beav_butt").Settings.Value("sound_mode") =0
End Sub

'******************************************************
'         Solenoids Map
'******************************************************

'(**) Used in backbox (translite)

'SolCallback(1) = ""            '1 - Left Bumper
'SolCallback(2) = ""            '2 - Right Bumper
'SolCallback(3) = ""            '3 - Bottom Bumper
'SolCallback(4) = ""            '4 - Left Slingshot
'SolCallback(5) = ""            '5 - Right Slingshot
SolCallback(6) = "SolVUK"         '6 - Bottom Scoop
SolCallback(7)  = "bsTK.SolOut"       '7 - Top Upkicker
SolCallback(8)  = "dtBankL.SolDropUp"   '8 - Left Bank Reset
SolCallback(9)  = "dtBankR.SolDropUp"   '9 - Right Bank Reset
SolCallback(10) = "SolDiv"          '10 - Diverter
SolCallback(11)  = "SolHeadbangersSpecial"  '11 - Headbangers Special
SolCallback(12)  = "SolHeadbangers"     '12 - Headbangers

SolCallback(22) = "SetLamp 152,"          '22 - Flasher:Ramps #1 (2)
SolCallback(23) = "SetLamp 153,"          '23 - Flasher:Ramps #2 (2)
SolCallback(24) = "SetLamp 154,"          '24 - Flasher:Ramps #3 (2)
SolCallback(25) = "SetLamp 155,"          '25 - Flasher:Devil Horns

SolCallback(26) = "GIRelay"         '26 - LightBox Relay
SolCallback(27) = ""            '27 - Ticket/Coin Meter
SolCallback(28) = "ReleaseBall"       '28 - Ball Release
SolCallback(29) = "SolOuthole"        '29 - Outhole.
SolCallback(30) = "SolKnocker"        '30 - Knocker
SolCallback(31) = "TiltRelay"       '31 - Tilt Relay
SolCallback(32) = "SolRun"                '32 - Game Over Relay

Sub SolRun(Enabled)
  vpmNudge.SolGameOn Enabled
  lightDMDborder.State = Enabled
  Dim GOoffon
  GOoffon = ABS(ABS(Enabled))
  If B2SOn = True Then
    Controller.B2SSetData 98, GOoffon
  End If
End Sub

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'******************************************************
'     TROUGH BASED ON FOZZY'S
'******************************************************

'Init Trough
Sub InitTrough
  Slot1.CreateSizedballWithMass Ballsize/2,Ballmass
  Slot2.CreateSizedballWithMass Ballsize/2,Ballmass
  Controller.Switch(44) = 1
  Controller.Switch(54) = 0
End Sub

'Handle Trough
Sub Slot2_Hit():UpdateTrough:End Sub
Sub Slot1_Hit():Controller.Switch(44) = 1:UpdateTrough:End Sub
Sub Slot1_UnHit():Controller.Switch(44) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  Slot1.TimerInterval = 300
  Slot1.TimerEnabled = 1
End Sub

Sub Slot1_Timer()
  Me.TimerEnabled = 0
  If Slot1.BallCntOver = 0 Then Slot2.kick 60, 9
End Sub

'Drain & Release
Sub Drain_Hit()
  PlaySoundAt "",Drain
  UpdateTrough
  Controller.Switch(54) = 1
End Sub

Sub swdrain_hit
    if l50.state = 1 Then
    Playsound "fx_drain"
    Else
    Playsound "fx_drain_laughs"
    End If
End Sub

Sub Drain_UnHit()
  Controller.Switch(54) = 0
End Sub

Sub SolOuthole(enabled)
  If enabled Then
    Drain.kick 60,20
    PlaySoundAt SoundFX(SSolenoidOn, DOFContactors), Drain
  Else
    PlaySoundAt SSolenoidOff, Drain
  End If
End Sub

Sub ReleaseBall(enabled)
  If enabled Then
    PlaySoundAt SoundFX("fx_ballrel",DOFContactors), Drain
    Slot1.kick 60, 12
    UpdateTrough
  End If
End Sub

'******************************************************
'         Solenoids Routines
'******************************************************

'*********** Knocker
Sub SolKnocker(enabled)
  If Enabled Then
    PlaySoundAt SoundFX("Knocker",DOFKnocker), l82
  End If
End Sub

'*********** Flippers
Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
    End If
End Sub

'*********** Scoop
Sub SolVUK(enabled)
If Enabled Then
If bsScoop.Balls > 0 Then
sw52.timerinterval=500:sw52.timerEnabled=1
CenterHole.Enabled=0
bsScoop.ExitSol_On
          If MyLightACounter > 25 Then
              tvlightboobtubeA.duration 1, 2300, 0
              tvlightboobtubeB.duration 1, 2300, 0
              tablelamplightA.duration 0, 2300, 1
              tablelamplightB.duration 0, 2300, 1
              Playsound "fx_light_switch"
              Timer001.Enabled = 0
              Dim bpw
            bpw = INT(3 * RND(1) )
            Select Case bpw
            Case 0:Playsound "fx_boobie_prize_win_boing"
            Case 1:Playsound "fx_boobie_prize_win_hey_baby"
              Case 2:Playsound "fx_boobie_prize_win_big_boobies"
            End Select
          Else
              tvlightA.duration 1, 2300, 0
              tvlightB.duration 1, 2300, 0
              tablelamplightA.duration 0, 2300, 1
              tablelamplightB.duration 0, 2300, 1
              Playsound "fx_light_switch"
              Timer001.Enabled = 0
              Dim bpl
            bpl = INT(4 * RND(1) )
            Select Case bpl
            Case 0:Playsound "fx_beavis_this_sucks"
            Case 1:Playsound "fx_uhh_change_it"
              Case 2:Playsound "fx_dammit_beavis_remote"
              Case 3:Playsound "fx_whats_wrong_with_the_tv"
              End Select
End If
End If
End If
End Sub

sub sw52_timer:me.timerEnabled=0:CenterHole.Enabled=1:end Sub

Dim MyLightACounter
MyLightACounter = 0

Sub Timer001_Timer
    MyLightACounter = MyLightACounter +1
End Sub

'*********** Kickback
Dim kickbackenabled: kickbackenabled = True: kickbacklight.duration 2, 100, 2
Dim sw57Hits

Sub kickbackstartswitch_Hit
    kickbackenabled = True
    kickbacklight.duration 0, 1000, 1:l99.state=0
End Sub

Sub sw57_Hit
    if Activeball.vely < 0 then
    sw57Hits = sw57Hits + 1
    If sw57Hits = 2 Then
        kickbackenabled = True
        kickbacklight.state=1:l99.state=0:
        sw57Hits = 0
    End If
End If
End Sub

Sub sw56_Hit
    If kickbackenabled = True Then
    kickbackplunger.Fire
    PlaySoundAt SoundFX("CenterEject",DOFContactors), kickbackplunger
    Playsound "fx_beavis_kick"
    kickburstlight.duration 2, 450, 0
    kickbackenabled = False
    kickbacklight.state=0:l99.state=1
    kickbackplunger.Pullback
End If
End Sub

'*********** Diverter
Dim DiverterDir

Sub SolDiv(Enabled)
  If Enabled Then
    DiverterOn.Collidable=1
    DiverterDir = -1
    Diverter.Interval = 5:Diverter.Enabled = 1
    PlaySoundAt SoundFX("DiverterOn",DOFContactors),DiverterP
    Else
    DiverterOn.Collidable=0
    DiverterDir = +1
    Diverter.Interval = 5:Diverter.Enabled = 1
    PlaySoundAt SoundFX("DiverterOn",DOFContactors),DiverterP
    End If
End Sub

Sub Diverter_Timer()
DiverterP.RotZ=DiverterP.RotZ+DiverterDir
If DiverterP.RotZ>-45 AND DiverterDir=1 Then Me.Enabled=0:DiverterP.RotZ=+90
If DiverterP.RotZ<-0 AND DiverterDir=-1 Then Me.Enabled=0:DiverterP.RotZ=-0
End Sub

'********* B&B Headbangers Special
Sub SolHeadbangersSpecial(enabled)
  If enabled Then
        Buttheadarm.RotZ=20
    Buttheadhead.RotZ=-18
        Buttheadneck.RotZ=-18
        BeavisArm.RotZ=20
        Beavishead.RotZ=18
        Beavisneck.RotZ=18
    Else
        Buttheadarm.RotZ=0
    Buttheadhead.RotZ=0
        Buttheadneck.RotZ=0
        BeavisArm.RotZ=0
        Beavishead.RotZ=0
        Beavisneck.RotZ=0
End If
End Sub

'********* Beavis & Butt-head Prims
Sub SolHeadbangers(enabled)
  If enabled Then
        Buttheadhead.RotZ=-5
        Buttheadneck.RotZ=-5
        Beavishead.RotZ=-5
        Beavisneck.RotZ=-5
    Else
        Buttheadhead.RotZ=0
        Buttheadneck.RotZ=0
        Beavishead.RotZ=0
        Beavisneck.RotZ=0
End If
End Sub

'******************************************************
'             Switches
'******************************************************

'Bumpers
Sub sw10_Hit() : vpmTimer.PulseSw 10 : PlaySoundAt SoundFX("LeftBumper_Hit",DOFContactors), ActiveBall:light14.duration 0, 200, 1: Playsound "bumper_1_hit":End Sub
Sub sw11_Hit() : vpmTimer.PulseSw 11 : PlaySoundAt SoundFX("RightBumper_Hit",DOFContactors), ActiveBall:light16.duration 0, 200, 1: Playsound "bumper_2_hit":End Sub
Sub sw12_Hit() : vpmTimer.PulseSw 12 : PlaySoundAt SoundFX("CenterBumper_Hit",DOFContactors), ActiveBall:light13.duration 0, 200, 1: Playsound "bumper_3_hit":End Sub

'Slings
Dim LStep,RStep
Sub sw13_Slingshot
  PlaySoundAt SoundFX("LeftSlingShot",DOFContactors), ActiveBall: Playsound "fx_butthead_sling_grunt": Playsound "fx_richochet":
     light1.duration 0, 100, 1
     light2.duration 0, 100, 1
     light3.duration 0, 100, 1
     light4.duration 0, 100, 1
    vpmTimer.PulseSw 13
  LSling.Visible = 0
  LSling1.Visible = 1
  sling1.TransZ = -20:SlingHornsLeft.RotX = 15:hornsLEFTscrews.RotX = 15
  LStep = 3
  Me.TimerInterval = 75
  Me.TimerEnabled = 1
End Sub

Sub sw13_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -10:SlingHornsLeft.RotX = 15:hornsLEFTscrews.RotX = 15:hornsLEFTshaft.RotX = 15:hornsLEFTshaftSHADOW.RotX = 15
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:SlingHornsLeft.RotX = -10:hornsLEFTscrews.RotX = -10:hornsLEFTshaft.RotX = -10:hornsLEFTshaftSHADOW.RotX = -10:Me.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub sw14_Slingshot
  PlaySoundAt SoundFX("RightSlingShot",DOFContactors), ActiveBall: Playsound "fx_beavis_sling_grunt": Playsound "fx_richochet_2":
  light19.duration 0, 100, 1
    light20.duration 0, 100, 1
    light21.duration 0, 100, 1
    light22.duration 0, 100, 1
    vpmTimer.PulseSw 14
  RSling.Visible = 0
  RSling1.Visible = 1
  sling2.TransZ = -20:SlingHornsRight.RotX = 15:hornsRIGHTscrews.RotX = 15
  RStep = 3
  Me.TimerInterval = 75
  Me.TimerEnabled = 1
End Sub

Sub sw14_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -10:SlingHornsRight.RotX = 15:hornsRIGHTscrews.RotX = 15:hornsRIGHTshaft.RotX = 15:hornsRIGHTshaftSHADOW.RotX = 15
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:SlingHornsRight.RotX = -10:hornsRIGHTscrews.RotX = -10:hornsRIGHTshaft.RotX = -10:hornsRIGHTshaftSHADOW.RotX = -10:Me.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'Frog Pinball Spinner
Sub sw15_Spin: vpmtimer.pulsesw 15:PlaySoundAt "fx_spinner", sw15: End Sub
Sub swbuttheadstare_hit
if Activeball.vely < 0 then
Playsound "frogspinnerlaughs"
if ButtheadHead.roty=0 then
            me.uservalue = -1
            me.timerinterval = 10
            me.timerenabled = 1
    end if
    end if
if l59.state = 1 Then
   Stopsound "frogspinnerlaughs"
if arrow2.visible = 1 Then
   Stopsound "frogspinnerlaughs"
End If
End If
if l58.state = 1 Then
   Stopsound "frogspinnerlaughs"
if arrow1.visible = 1 Then
   Stopsound "frogspinnerlaughs"
End If
End If
End Sub
Sub swbuttheadstare_timer
    ButtheadHead.roty = ButtheadHead.roty + (me.uservalue * 4.0)
    ButtheadArm.roty = ButtheadHead.roty
    ButtheadBody.roty = ButtheadHead.roty
    ButtheadNeck.roty = ButtheadHead.roty
    ButtheadPrimShadow.roty = ButtheadHead.roty
    if ButtheadHead.roty < - 40 Then
        me.uservalue = 1
        swbuttheadstare.timerenabled = 0
        vpmtimer.addtimer 3000, "swbuttheadstare.timerenabled = 1 '"
    end If
    if ButtheadHead.roty > 0 Then
       ButtheadHead.roty = 0
        me.timerenabled = 0
    end if
end sub

Sub swbeavisstare_hit
if BeavisHead.roty=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
        end if
End Sub
Sub swbeavisstare_timer
    BeavisHead.roty = BeavisHead.roty + (me.uservalue * 4.0)
    BeavisArm.roty = BeavisHead.roty
    BeavisBody.roty = BeavisHead.roty
    BeavisNeck.roty = BeavisHead.roty
    BeavisPrimShadow.roty = BeavisHead.roty
    if BeavisHead.roty > 45 Then
        me.uservalue = -1
        swbeavisstare.timerenabled = 0
        vpmtimer.addtimer 3000, "swbeavisstare.timerenabled = 1 '"
    end If
    if BeavisHead.roty < 0 Then
       BeavisHead.roty = 0
        me.timerenabled = 0
    end if
end sub

'StewartSpinner
Sub stewartspinner_Spin:if Activeball.vely > 0 then:PlaySound "fx_spinner": End If:End Sub

'Drop Targets
Sub sw16_dropped:dtBankL.Hit 1: Playsound "fx_butthead_drop_guitar_2":
if ButtheadHead.rotz=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
    end if
End Sub
Sub sw16_timer
    ButtheadHead.rotz = ButtheadHead.rotz + (me.uservalue * 3)
    ButtheadNeck.rotz = ButtheadHead.rotz
    ButtheadArm.rotz = ButtheadHead.rotz
    if ButtheadHead.rotz > 10 Then
        me.uservalue = -1
        sw16.timerenabled = 0
        vpmtimer.addtimer 350, "sw16.timerenabled = 1 '"
    end If
    if ButtheadHead.rotz < 0 Then
        ButtheadHead.rotz = 0
        me.timerenabled = 0
    end if
End Sub

Sub sw26_dropped:dtBankL.Hit 2: Playsound "fx_butthead_drop_guitar_4":
if ButtheadHead.rotz=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
    end if
End Sub
Sub sw26_timer
    ButtheadHead.rotz = ButtheadHead.rotz + (me.uservalue * 3)
    ButtheadNeck.rotz = ButtheadHead.rotz
    ButtheadArm.rotz = ButtheadHead.rotz
    if ButtheadHead.rotz > 10 Then
        me.uservalue = -1
        sw26.timerenabled = 0
        vpmtimer.addtimer 350, "sw26.timerenabled = 1 '"
    end If
    if ButtheadHead.rotz < 0 Then
        ButtheadHead.rotz = 0
        me.timerenabled = 0
    end if
End Sub

Sub sw36_dropped:dtBankL.Hit 3: Playsound "fx_butthead_drop_guitar_3":
if ButtheadHead.rotz=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
    end if
End Sub
Sub sw36_timer
    ButtheadHead.rotz = ButtheadHead.rotz + (me.uservalue * 3)
    ButtheadNeck.rotz = ButtheadHead.rotz
    ButtheadArm.rotz = ButtheadHead.rotz
    if ButtheadHead.rotz > 10 Then
        me.uservalue = -1
        sw36.timerenabled = 0
        vpmtimer.addtimer 350, "sw36.timerenabled = 1 '"
    end If
    if ButtheadHead.rotz < 0 Then
        ButtheadHead.rotz = 0
        me.timerenabled = 0
    end if
End Sub

Sub sw46_dropped:dtBankL.Hit 4: Playsound "fx_butthead_drop_guitar_1":
if ButtheadHead.rotz=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
    end if
End Sub
Sub sw46_timer
    ButtheadHead.rotz = ButtheadHead.rotz + (me.uservalue * 3)
    ButtheadNeck.rotz = ButtheadHead.rotz
    ButtheadArm.rotz = ButtheadHead.rotz
    if ButtheadHead.rotz > 10 Then
        me.uservalue = -1
        sw46.timerenabled = 0
        vpmtimer.addtimer 350, "sw46.timerenabled = 1 '"
    end If
    if ButtheadHead.rotz < 0 Then
        ButtheadHead.rotz = 0
        me.timerenabled = 0
    end if
End Sub

Sub sw17_dropped:dtBankR.Hit 1: Playsound "fx_beavis_drink_1":
if BeavisHead.rotz=0 then
            me.uservalue = -1
            me.timerinterval = 10
            me.timerenabled = 1
    end if
End Sub
Sub sw17_timer
    BeavisHead.rotz = BeavisHead.rotz + (me.uservalue * 3)
    BeavisNeck.rotz = BeavisHead.rotz
    BeavisArm.rotz = BeavisHead.rotz
    if BeavisHead.rotz < - 10 Then
        me.uservalue = 1
        sw17.timerenabled = 0
        vpmtimer.addtimer 750, "sw17.timerenabled = 1 '"
    end If
    if BeavisHead.rotz > 0 Then
        BeavisHead.rotz = 0
        me.timerenabled = 0
    end if
End Sub

Sub sw27_dropped:dtBankR.Hit 2: Playsound "fx_beavis_eat_1":
if BeavisHead.rotz=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
    end if
End Sub
Sub sw27_timer
    BeavisHead.rotz = BeavisHead.rotz + (me.uservalue * 3)
    BeavisNeck.rotz = BeavisHead.rotz
    BeavisArm.rotz = BeavisHead.rotz
    if BeavisHead.rotz > 10 Then
        me.uservalue = -1
        sw27.timerenabled = 0
        vpmtimer.addtimer 750, "sw27.timerenabled = 1 '"
    end If
    if BeavisHead.rotz < 0 Then
        BeavisHead.rotz = 0
        me.timerenabled = 0
    end if
End Sub

Sub sw37_dropped:dtBankR.Hit 3: Playsound "fx_beavis_drink_2":
if BeavisHead.rotz=0 then
            me.uservalue = -1
            me.timerinterval = 10
            me.timerenabled = 1
    end if
End Sub
Sub sw37_timer
    BeavisHead.rotz = BeavisHead.rotz + (me.uservalue * 3)
    BeavisNeck.rotz = BeavisHead.rotz
    BeavisArm.rotz = BeavisHead.rotz
    if BeavisHead.rotz < - 10 Then
        me.uservalue = 1
        sw37.timerenabled = 0
        vpmtimer.addtimer 750, "sw37.timerenabled = 1 '"
    end If
    if BeavisHead.rotz > 0 Then
        BeavisHead.rotz = 0
        me.timerenabled = 0
    end if
End Sub

Sub sw47_dropped:dtBankR.Hit 4: Playsound "fx_beavis_eat_2":
if BeavisHead.rotz=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
    end if
End Sub
Sub sw47_timer
    BeavisHead.rotz = BeavisHead.rotz + (me.uservalue * 3)
    BeavisNeck.rotz = BeavisHead.rotz
    BeavisArm.rotz = BeavisHead.rotz
    if BeavisHead.rotz > 10 Then
        me.uservalue = -1
        sw47.timerenabled = 0
        vpmtimer.addtimer 750, "sw47.timerenabled = 1 '"
    end If
    if BeavisHead.rotz < 0 Then
        BeavisHead.rotz = 0
        me.timerenabled = 0
    end if
End Sub

'Standup Targets
Sub Sw20_Hit():vpmTimer.PulseSw 20: PlaysoundAt SoundFX("fx_target",DOFTargets), ActiveBall: Playsound "fx_butthead_hit_3":
if ButtheadHead.roty=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
        end if
End Sub
Sub sw20_timer
    ButtheadHead.roty = ButtheadHead.roty + (me.uservalue * 3)
    if ButtheadHead.roty > 40 Then
        me.uservalue = -1
        sw20.timerenabled = 0
        vpmtimer.addtimer 400, "sw20.timerenabled = 1 '"
    end If
    if ButtheadHead.roty < 0 Then
        ButtheadHead.roty = 0
        me.timerenabled = 0
    end if
End Sub

Sub Sw30_Hit():vpmTimer.PulseSw 30: PlaysoundAt SoundFX("fx_target",DOFTargets), ActiveBall: Playsound "fx_butthead_hit_2":
if ButtheadHead.rotz=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
        end if
End Sub
Sub sw30_timer
    ButtheadHead.rotz = ButtheadHead.rotz + (me.uservalue * 3)
    ButtheadNeck.rotz = ButtheadHead.rotz
    if ButtheadHead.rotz > 14 Then
        me.uservalue = -1
        sw30.timerenabled = 0
        vpmtimer.addtimer 400, "sw30.timerenabled = 1 '"
    end If
    if ButtheadHead.rotz < 0 Then
        ButtheadHead.rotz = 0
        me.timerenabled = 0
    end if
End Sub

Sub Sw40_Hit():vpmTimer.PulseSw 40: PlaysoundAt SoundFX("fx_target",DOFTargets), ActiveBall: Playsound "fx_butthead_hit_1":
if ButtheadHead.roty=0 then
            me.uservalue = -1
            me.timerinterval = 10
            me.timerenabled = 1
        end if
End Sub
Sub sw40_timer
    ButtheadHead.roty = ButtheadHead.roty + (me.uservalue * 3)
    if ButtheadHead.roty < - 40 Then
        me.uservalue = 1
        sw40.timerenabled = 0
        vpmtimer.addtimer 400, "sw40.timerenabled = 1 '"
    end If
    if ButtheadHead.roty > 0 Then
        ButtheadHead.roty = 0
        me.timerenabled = 0
    end if
End Sub

Sub Sw21_Hit():vpmTimer.PulseSw 21: PlaysoundAt SoundFX("fx_target",DOFTargets), ActiveBall:
Dim fft1
  fft1 = INT(5 * RND(1) )
  Select Case fft1
  Case 0:Playsound "fx_fast_food_splotch_hit_1"
  Case 1:Playsound "fx_fast_food_splotch_hit_2"
    Case 2:Playsound "fx_fast_food_splotch_hit_3"
    Case 3:Playsound "fx_burp"
    Case 4:Playsound ""
  End Select
End Sub

Sub Sw31_Hit():vpmTimer.PulseSw 31: PlaysoundAt SoundFX("fx_target",DOFTargets), ActiveBall:
Dim fft2
  fft2 = INT(5 * RND(1) )
  Select Case fft2
  Case 0:Playsound "fx_fast_food_splotch_hit_1"
  Case 1:Playsound "fx_fast_food_splotch_hit_2"
    Case 2:Playsound "fx_fast_food_splotch_hit_3"
    Case 3:Playsound "fx_burp"
    Case 4:Playsound ""
  End Select
End Sub

Sub Sw41_Hit():vpmTimer.PulseSw 41: PlaysoundAt SoundFX("fx_target",DOFTargets), ActiveBall:
Dim fft3
  fft3 = INT(5 * RND(1) )
  Select Case fft3
  Case 0:Playsound "fx_fast_food_splotch_hit_1"
  Case 1:Playsound "fx_fast_food_splotch_hit_2"
    Case 2:Playsound "fx_fast_food_splotch_hit_3"
    Case 3:Playsound "fx_burp"
    Case 4:Playsound ""
  End Select
End Sub

Sub Sw22_Hit():vpmTimer.PulseSw 22: PlaysoundAt SoundFX("fx_target",DOFTargets), ActiveBall: Playsound "fx_beavis_hit_3":
    if BeavisHead.roty=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
        end if
End Sub
Sub sw22_timer
    BeavisHead.roty = BeavisHead.roty + (me.uservalue * 3)
    if BeavisHead.roty > 40 Then
        me.uservalue = -1
        sw22.timerenabled = 0
        vpmtimer.addtimer 400, "sw22.timerenabled = 1 '"
    end If
    if BeavisHead.roty < 0 Then
        BeavisHead.roty = 0
        me.timerenabled = 0
    end if
End Sub

Sub Sw32_Hit():vpmTimer.PulseSw 32: PlaysoundAt SoundFX("fx_target",DOFTargets), ActiveBall: Playsound "fx_beavis_hit_2":
if BeavisHead.rotz=0 then
            me.uservalue = -1
            me.timerinterval = 10
            me.timerenabled = 1
        end if
End Sub
Sub sw32_timer
    BeavisHead.rotz = BeavisHead.rotz + (me.uservalue * 3)
    BeavisNeck.rotz = BeavisHead.rotz
    if BeavisHead.rotz < - 14 Then
        me.uservalue = 1
        sw32.timerenabled = 0
        vpmtimer.addtimer 400, "sw32.timerenabled = 1 '"
    end If
    if BeavisHead.rotz > 0 Then
        BeavisHead.rotz = 0
        me.timerenabled = 0
    end if
End Sub

Sub Sw42_Hit():vpmTimer.PulseSw 42: PlaysoundAt SoundFX("fx_target",DOFTargets), ActiveBall: Playsound "fx_beavis_hit_1":
if BeavisHead.roty=0 then
            me.uservalue = -1
            me.timerinterval = 10
            me.timerenabled = 1
        end if
End Sub
Sub sw42_timer
    BeavisHead.roty = BeavisHead.roty + (me.uservalue * 3)
    if BeavisHead.roty < - 40 Then
        me.uservalue = 1
        sw42.timerenabled = 0
        vpmtimer.addtimer 400, "sw42.timerenabled = 1 '"
    end If
    if BeavisHead.roty > 0 Then
        BeavisHead.roty = 0
        me.timerenabled = 0
    end if
 End Sub

'Rollovers
Sub Sw23_Hit:Controller.Switch(23)=1: PlaysoundAt "fx_sensor", ActiveBall:
    if l66.state=1 then
    Dim u
  u = INT(5 * RND(1) )
  Select Case u
  Case 0:Playsound "fx_butthead_get_me_some_nachos"
  Case 1:Playsound "fx_butthead_get_ready_for_nachos"
    Case 2:Playsound "fx_beavis_we_want_nachos"
    Case 3:Playsound ""
    Case 4:Playsound ""
    End Select
End If
End Sub

Sub Sw23_UnHit:Controller.Switch(23)=0: End Sub

Sub Sw24_Hit:Controller.Switch(24)=1: PlaysoundAt "fx_sensor", ActiveBall:
    if l56.state=1 then
    Dim i
  i = INT(6 * RND(1) )
  Select Case i
  Case 0:Playsound "fx_butthead_go_to_burger_world":BWSignLight.duration 1,3500,0
  Case 1:Playsound "fx_butthead_burger_world_open":BWSignLight.duration 1,2700,0
    Case 2:Playsound "fx_beavis_back_to_burger_world":BWSignLight.duration 1,3500,0
    Case 3:Playsound ""
    Case 4:Playsound ""
    Case 5:Playsound ""
    End Select
End If
End Sub

Sub Sw24_UnHit:Controller.Switch(24)=0: End Sub

Sub Sw25_Hit:Controller.Switch(25)=1:PlaysoundAt "fx_sensor", ActiveBall:PlaysoundAt "fx_top_rollovers_sound", ActiveBall:End Sub
Sub Sw25_UnHit:Controller.Switch(25)=0:End Sub
Sub Sw33_Hit:if Activeball.vely > 0 then:Controller.Switch(33)=1: PlaysoundAt "fx_sensor", ActiveBall:End If
    if Activeball.vely > 0 then
    if kickbacklight.state = 1 Then
    Dim kb
  kb = INT(2 * RND(1) )
  Select Case kb
  Case 0:Playsound "fx_beavis_lets_kick_it"
  Case 1:Playsound "fx_beavis_yeah_kick_it"
  End Select
    Stopsound "fx_butthead_whats_your_problem"
  Stopsound "fx_butthead_you_lost_your_ball"
    Stopsound "fx_butthead_you_monkeyspank"
end if
end if
End Sub
Sub Sw33_UnHit:Controller.Switch(33)=0:End Sub
Sub Sw34_Hit:Controller.Switch(34)=1: PlaysoundAt "fx_sensor", ActiveBall:End Sub
Sub Sw34_UnHit:Controller.Switch(34)=0: End Sub
Sub Sw35_Hit:Controller.Switch(35)=1: PlaysoundAt "fx_sensor", ActiveBall:PlaysoundAt "fx_top_rollovers_sound", ActiveBall:End Sub
Sub Sw35_UnHit:Controller.Switch(35)=0: End Sub
Sub Sw43_Hit:Controller.Switch(43)=1: PlaysoundAt "fx_sensor", ActiveBall: kickbackplunger.Pullback:beavisrampplasticlight.state = 0:buttheadrampplasticlight.state = 0: End Sub
Sub Sw43_UnHit:Controller.Switch(43)=0: End Sub
Sub Sw45_Hit:Controller.Switch(45)=1: PlaysoundAt "fx_sensor", ActiveBall:PlaysoundAt "fx_top_rollovers_sound", ActiveBall:End Sub
Sub Sw45_UnHit:Controller.Switch(45)=0: End Sub
Sub Sw55_Hit:Controller.Switch(55)=1: PlaysoundAt "fx_sensor", ActiveBall: End Sub
Sub Sw55_UnHit:Controller.Switch(55)=0: End Sub
Sub Sw55b_Hit:Controller.Switch(55)=1: PlaysoundAt "fx_sensor", ActiveBall: Playsound "fx_BallDropDelayed" End Sub
Sub Sw55b_UnHit:Controller.Switch(55)=0: End Sub

'Optos
Sub Sw50_Hit():vpmTimer.PulseSw 50: PlaysoundAt "", ActiveBall:
    if l46.state=1 Then
    Dim bb
  bb = INT(3 * RND(1) )
  Select Case bb
  Case 0:Playsound "fx_beavis_nachos_rule":Stopsound "fx_stewart_hit_spin_scream":Stopsound "fx_stewart_hit_spin":Stopsound "mtvintro4b&b"
  Case 1:Playsound "fx_beavis_nachos":Stopsound "fx_stewart_hit_spin_scream":Stopsound "fx_stewart_hit_spin":Stopsound "mtvintro4b&b"
    Case 2:Playsound "fx_beavis_yeeah_nachos":Stopsound "fx_stewart_hit_spin_scream":Stopsound "fx_stewart_hit_spin":Stopsound "mtvintro4b&b"
    End Select
End If
    if l44.state=1 Then
    Stopsound "fx_beavis_nachos_rule":Stopsound "fx_beavis_nachos":Stopsound "fx_stewart_hit_spin_scream":Stopsound "fx_stewart_hit_spin":Stopsound "mtvintro4b&b"
End if
End Sub

Sub Sw51_Hit():vpmTimer.PulseSw 51: PlaysoundAt "", ActiveBall:
    Dim bwsfx
  bwsfx = INT(8 * RND(1) )
  Select Case bwsfx
  Case 0:Playsound "fx_burger_world_crash"
    Case 1:Playsound "fx_burger_world_crash"
    Case 2:Playsound "fx_burger_world_crash"
    Case 3:Playsound "fx_carl_what_the_hell":carllight.duration 1, 2000, 0
    Case 4:Playsound "fx_carl_what_are_you_doing":carllight.duration 1, 2000, 0
    Case 5:Playsound "fx_carl_where_are_those_idiots":carllight.duration 1, 2000, 0
    Case 6:Playsound "fx_butthead_testes_testes":speakerlight.duration 1, 2700, 0
    Case 7:Playsound "fx_butthead_customers_suck":speakerlight.duration 1, 2000, 0
  End Select
    if l36.state=1 Then
    Stopsound "fx_burger_world_crash"
    Stopsound "fx_carl_what_the_hell":carllight.state = 0
    Stopsound "fx_carl_what_are_you_doing":carllight.state = 0
    Stopsound "fx_carl_where_are_those_idiots":carllight.state = 0
    Stopsound "fx_butthead_testes_testes":speakerlight.state = 0
    Stopsound "fx_butthead_customers_suck":speakerlight.state = 0
End If
    if l37.state=1 Then
    Stopsound "fx_burger_world_crash"
    Stopsound "fx_carl_what_the_hell":carllight.state = 0
    Stopsound "fx_carl_what_are_you_doing":carllight.state = 0
    Stopsound "fx_carl_where_are_those_idiots":carllight.state = 0
    Stopsound "fx_butthead_testes_testes":speakerlight.state = 0
    Stopsound "fx_butthead_customers_suck":speakerlight.state = 0
End If
    if l32.state=1 Then
    Stopsound "fx_burger_world_crash"
    Stopsound "fx_carl_what_the_hell":carllight.state = 0
    Stopsound "fx_carl_what_are_you_doing":carllight.state = 0
    Stopsound "fx_carl_where_are_those_idiots":carllight.state = 0
    Stopsound "fx_butthead_testes_testes":speakerlight.state = 0
    Stopsound "fx_butthead_customers_suck":speakerlight.state = 0
End If
    if l31.state=1 Then
    Stopsound "fx_burger_world_crash"
    Stopsound "fx_carl_what_the_hell":carllight.state = 0
    Stopsound "fx_carl_what_are_you_doing":carllight.state = 0
    Stopsound "fx_carl_where_are_those_idiots":carllight.state = 0
    Stopsound "fx_butthead_testes_testes":speakerlight.state = 0
    Stopsound "fx_butthead_customers_suck":speakerlight.state = 0
End If
if l44.state=1 Then
    Stopsound "fx_carl_what_the_hell":carllight.state = 0
    Stopsound "fx_carl_what_are_you_doing":carllight.state = 0
    Stopsound "fx_carl_where_are_those_idiots":carllight.state = 0
    Stopsound "fx_butthead_testes_testes":speakerlight.state = 0
    Stopsound "fx_butthead_customers_suck":speakerlight.state = 0
End If
End Sub

' Burger World Sign
Sub BWSignSpin_Hit
if Activeball.vely < 0 then
    Playsound "fx_bw_sign_spin"
    BurgerWorldSignLIT.visible = 1
    RotBWSignStep = 7
    BWSignTimer.Enabled = 1
End If
End Sub

Dim RotBWSign, RotBWSignStep
RotBWSign = 0
Sub BWSignTimer_Timer
    RotBWSign = (RotBWSign + RotBWSignStep)MOD 360
    BurgerWorldSign.RotY = RotBWSign
    BurgerWorldSignLIT.RotY = RotBWSign
    BurgerWorldSignLIT2.RotY = RotBWSign
    RotBWSignStep = RotBWSignStep - 0.05
    If RotBWSignStep <1 Then BWSignTimer.Enabled = 0
End Sub

Sub BWSignLitOFFtop_Hit
    BurgerWorldSignLIT.visible = 0
End Sub

Sub BWSignLitOFFbottom_Hit
    BurgerWorldSignLIT.visible = 0
End Sub

' Couch Cushions
Sub sw58_hit()
  Playsound "fx_BallDropMuffled":
Dim cb
  cb = INT(2 * RND(1) )
  Select Case cb
  Case 0:Playsound "couchboing"
  Case 1:Playsound "couchboing2"
  End Select
    if Couchcushions.rotx=0 then
      me.uservalue = -1
      me.timerinterval = 10
      me.timerenabled = 1
  end if
end sub

Sub sw58_timer
  Couchcushions.rotx = Couchcushions.rotx + (me.uservalue * 10)
  if Couchcushions.rotx < - 75 Then
    me.uservalue = 1
  end If
  if Couchcushions.rotx > 0 Then
    Couchcushions.rotx = 0
    me.timerenabled = 0
  end if
End Sub

Sub livingroomwirewall_Hit
    Playsound "fx_metal_hit_1"
End Sub

' Stewartspinner
Sub stewartspinnerlightSW_hit
    stewartspotlightLIGHT.duration 1, 2000, 0
End Sub

' Mtv Logo
Dim TargetHits: TargetHits = 0

Sub SWmtvlogo_Hit
Select Case TargetHits
    Case 0: PlaySound "mtvintro1":mtvlight.duration 1, 2500, 0:beavisrampplasticlight.duration 2, 2500, 0:buttheadrampplasticlight.duration 2, 2500, 0
    Case 1: PlaySound ""
    Case 2: PlaySound "mtvintro2":mtvlight.duration 1, 2500, 0:beavisrampplasticlight.duration 2, 2500, 0:buttheadrampplasticlight.duration 2, 2500, 0
    Case 3: PlaySound ""
    Case 4: PlaySound "mtvintro3":mtvlight.duration 1, 2500, 0:beavisrampplasticlight.duration 2, 2500, 0:buttheadrampplasticlight.duration 2, 2500, 0
    Case 5: PlaySound ""
    Case 6: PlaySound "mtvintro4b&b":mtvlight.duration 1, 2500, 0:beavisrampplasticlight.duration 2, 2500, 0:buttheadrampplasticlight.duration 2, 2500, 0
    Case 7: PlaySound ""
End Select
TargetHits = (TargetHits + 1) MOD 8
'mtvlight.duration 1, 2500, 0
End Sub

Sub mtvlightoff_hit
'    mtvlight.duration 1, 850, 0
     mtvlight.state = 0
     beavisrampplasticlight.state = 0
     buttheadrampplasticlight.state = 0
End Sub

Sub swskull3light_hit
    skull3light.duration 1, 150, 0:
    Dim sb
  sb = INT(4 * RND(1) )
  Select Case sb
  Case 0:Playsound "fx_skull_bonks"
  Case 1:Playsound "fx_skull_bonks_2"
    Case 2:Playsound ""
    Case 3:Playsound ""
    End Select
End Sub

Sub swskull2light_hit
    skull2light.duration 1, 150, 0:End Sub

Sub swskull1light_hit
    skull1light.duration 1, 150, 0:End Sub

Sub balltroughlighton_hit
    balltroughlight. state = 1
End Sub

Sub balltroughlightoff_hit
    balltroughlight. state = 0
End Sub

Sub swfrogjump_hit()
    if Activeball.vely < 0 then
    frogspotlight.duration 1, 3150, 0
        if frog.transy=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
    Dim w
  w = INT(2 * RND(1) )
  Select Case w
  Case 0:Playsound "fx_frog_ribbit"
  Case 1:Playsound "fx_frog_jump"
  End Select
End If
End if
End Sub

Sub swfrogjump_timer
    frog.transy = frog.transy + (me.uservalue * 1.5)
    frogbracket.transy = frog.transy
    if frog.transy > 35 Then
        me.uservalue = -1
    end If
    if frog.transy < 0 Then
        frog.transy = 0
        me.timerenabled = 0
    end if
end sub

Sub swbeavisloseball_hit
if l4.state = 0 then
if BeavisHead.roty=0 then
            me.uservalue = -1
            me.timerinterval = 10
            me.timerenabled = 1
        end if
    Dim belb
  belb = INT(3 * RND(1) )
  Select Case belb
  Case 0:Playsound "fx_beavis_settle_down":spotlightRIGHTspecial.duration 1, 1800, 0
  Case 1:Playsound "fx_beavis_choke_choke":spotlightRIGHTspecial.duration 1, 1800, 0
    Case 2:Playsound "fx_beavis_fire_fire":spotlightRIGHTspecial.duration 1, 1800, 0:light21.duration 0, 1800, 1:light22.duration 0, 1800, 1:lightFIRE.duration 2,1800,0
    End Select
end if
End Sub

Sub swbeavisloseball_timer
    BeavisHead.roty = BeavisHead.roty + (me.uservalue * 2.6)
    BeavisArm.roty = BeavisHead.roty
    BeavisBody.roty = BeavisHead.roty
    BeavisNeck.roty = BeavisHead.roty
    BeavisPrimShadow.roty = BeavisHead.roty
    if BeavisHead.roty < - 25 Then
        me.uservalue = 1
        swbeavisloseball.timerenabled = 0
        vpmtimer.addtimer 1500, "swbeavisloseball.timerenabled = 1 '"
    end If
    if BeavisHead.roty > 0 Then
       BeavisHead.roty = 0
        me.timerenabled = 0
    end if
end sub

Sub swbuttheadloseball_hit
if Activeball.vely > 0 then
if l3.state = 0 then
if kickbacklight.state = 0 then
    if ButtheadHead.roty=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
    Dim bulb
  bulb = INT(3 * RND(1) )
  Select Case bulb
  Case 0:Playsound "fx_butthead_whats_your_problem":spotlightLEFTspecial.duration 1, 1900, 0
  Case 1:Playsound "fx_butthead_you_lost_your_ball":spotlightLEFTspecial.duration 1, 2500, 0
    Case 2:Playsound "fx_butthead_you_monkeyspank":spotlightLEFTspecial.duration 1, 2600, 0
    End Select
end If
end if
end if
end if
End Sub

Sub swbuttheadloseball_timer
    ButtheadHead.roty = ButtheadHead.roty + (me.uservalue * 2.6)
    ButtheadArm.roty = ButtheadHead.roty
    ButtheadBody.roty = ButtheadHead.roty
    ButtheadNeck.roty = ButtheadHead.roty
    ButtheadPrimShadow.roty = ButtheadHead.roty
    if ButtheadHead.roty > 25 Then
        me.uservalue = -1
        swbuttheadloseball.timerenabled = 0
        vpmtimer.addtimer 1500, "swbuttheadloseball.timerenabled = 1 '"
    end If
    if ButtheadHead.roty < 0 Then
       ButtheadHead.roty = 0
        me.timerenabled = 0
    end if
end sub

Sub swbuttheadtvlook_hit
if ButtheadHead.roty=0 then
            me.uservalue = -1
            me.timerinterval = 10
            me.timerenabled = 1
        end if
End Sub

Sub swbuttheadtvlook_timer
    ButtheadHead.roty = ButtheadHead.roty + (me.uservalue * 4.0)
    ButtheadArm.roty = ButtheadHead.roty
    ButtheadBody.roty = ButtheadHead.roty
    ButtheadNeck.roty = ButtheadHead.roty
    ButtheadPrimShadow.roty = ButtheadHead.roty
    if ButtheadHead.roty < - 80 Then
        me.uservalue = 1
        swbuttheadtvlook.timerenabled = 0
        vpmtimer.addtimer 2200, "swbuttheadtvlook.timerenabled = 1 '"
    end If
    if ButtheadHead.roty > 0 Then
       ButtheadHead.roty = 0
        me.timerenabled = 0
    end if
end sub

Sub swbeavistvlook_hit
if BeavisHead.roty=0 then
            me.uservalue = 1
            me.timerinterval = 10
            me.timerenabled = 1
        end if
End Sub

Sub swbeavistvlook_timer
    BeavisHead.roty = BeavisHead.roty + (me.uservalue * 4.0)
    BeavisArm.roty = BeavisHead.roty
    BeavisBody.roty = BeavisHead.roty
    BeavisNeck.roty = BeavisHead.roty
    BeavisPrimShadow.roty = BeavisHead.roty
    if BeavisHead.roty > 100 Then
        me.uservalue = -1
        swbeavistvlook.timerenabled = 0
        vpmtimer.addtimer 1800, "swbeavistvlook.timerenabled = 1 '"
    end If
    if BeavisHead.roty < 0 Then
       BeavisHead.roty = 0
        me.timerenabled = 0
    end if
end sub

Sub stewartspinnerSFX_hit
    if Activeball.vely < 0 then
    Dim r
  r = INT(3 * RND(1) )
  Select Case r
  Case 0:Playsound "fx_stewart_hit_spin"
  Case 1:Playsound "fx_stewart_hit_spin_scream"
    Case 2:Playsound ""
  End Select
End If
End Sub

Sub mtvmusicstop_hit
    if Activeball.vely > 0 then
Stopsound "mtvintro1"
Stopsound "mtvintro2"
Stopsound "mtvintro3"
Stopsound "mtvintro4b&b"
mtvlight.state = 0
beavisrampplasticlight.state = 0
buttheadrampplasticlight.state = 0
End If
End Sub

Sub stewartspinnerSFXstop_hit
    if Activeball.vely > 0 then
Stopsound "fx_stewart_hit_spin"
Stopsound "fx_stewart_hit_spin_scream"
stewartspotlightLIGHT.state = 0
End If
End Sub

Sub swMcVicker_hit
if Activeball.vely > 0 then
    if l66.state = 0 Then
    Dim pm
  pm = INT(8 * RND(1) )
  Select Case pm
  Case 0:if l86.state = 1 Then: Playsound "": mcvickerlight.state = 0: else: Playsound "fx_mcvicker_1": mcvickerlight.duration 2, 3200, 0: End If
    Case 1:if l86.state = 1 Then: Playsound "": mcvickerlight.state = 0: else: Playsound "fx_mcvicker_2": mcvickerlight.duration 2, 3000, 0: End If
    Case 2:if l86.state = 1 Then: Playsound "": mcvickerlight.state = 0: else: Playsound "fx_mcvicker_3": mcvickerlight.duration 2, 2600, 0: End If
    Case 3:if l86.state = 1 Then: Playsound "": else: Playsound "fx_b&b_short_laugh": End If
    Case 4:if l86.state = 1 Then: Playsound "": else: Playsound "fx_b&b_short_laugh": End If
    Case 5:Playsound ""
    Case 6:Playsound ""
    Case 7:Playsound ""
    End Select
End If
End If
End Sub

Sub swBuzzcut_hit
if Activeball.vely > 0 then
    if l56.state = 0 Then
    Dim cb
  cb = INT(11 * RND(1) )
  Select Case cb
  Case 0:if l86.state = 1 Then: Playsound "": buzzcutlight.state = 0: else: Playsound "fx_buzzcut_1": buzzcutlight.duration 2, 3000, 0: End If
  Case 1:if l86.state = 1 Then: Playsound "": buzzcutlight.state = 0: else: Playsound "fx_buzzcut_2": buzzcutlight.duration 2, 3000, 0: End If
    Case 2:if l86.state = 1 Then: Playsound "": buzzcutlight.state = 0: else: Playsound "fx_buzzcut_3": buzzcutlight.duration 2, 3000, 0: End If
    Case 3:if l86.state = 1 Then: Playsound "": light18.state = 0: else: Playsound "fx_daria_get_a_life": light18.duration 2, 700, 1: End If
    Case 4:if l86.state = 1 Then: Playsound "": light18.state = 0: else: Playsound "fx_diaharrea_cha_cha_cha": light18.duration 2, 1500, 1: End If
    Case 5:if l86.state = 1 Then: Playsound "": else: Playsound "fx_b&b_short_laugh": End If
    Case 6:if l86.state = 1 Then: Playsound "": else: Playsound "fx_b&b_short_laugh": End If
    Case 7:Playsound ""
    Case 8:Playsound ""
    Case 9:Playsound ""
    Case 10:Playsound ""
    End Select
End If
End If
End Sub

Sub swgwar_hit
if Activeball.vely < 0 then
    Dim gwar
  gwar = INT(8 * RND(1) )
  Select Case gwar
  Case 0:Playsound "fx_beavis_gwar":l98.duration 2, 1700, 0
  Case 1:Playsound "fx_butthead_gwar_rules":l98.duration 2, 2300, 0
    Case 2:Playsound "fx_metal_metal_land":l98.duration 2, 2100, 0
    Case 3:Playsound "fx_beavis_heavy_metal_rules":l98.duration 2, 1700, 0
    Case 4:Playsound ""
    Case 5:Playsound ""
    Case 6:Playsound ""
    Case 7:Playsound ""
    End Select
End If
End Sub

Sub swtoiletflush_hit
    Dim tf
  tf = INT(7 * RND(1) )
  Select Case tf
  Case 0:Playsound "fx_toilet_flush":ltoilet.duration 2, 2200, 0
  Case 1:Playsound "fx_toilet_flush_2":ltoilet.duration 2, 2500, 0
    Case 2:Playsound ""
    Case 3:Playsound ""
    Case 4:Playsound ""
    Case 5:Playsound ""
    Case 6:Playsound ""
    End Select
    Dim honk
  honk = INT(3 * RND(1) )
  Select Case honk
  Case 0:Playsound "fx_honk"
  Case 1:Playsound ""
    Case 2:Playsound ""
  End Select
End Sub

Sub swleftloseball_hit
if kickbacklight.state = 1 then
    Stopsound "fx_butthead_whats_your_problem"
    Stopsound "fx_butthead_you_lost_your_ball"
    Stopsound "fx_butthead_you_monkeyspank"
end if
End Sub

Sub swLRballdrop_hit
    Playsound "fx_lr_ball_drop"
End Sub

' Middle Scoop Trapdoor
Sub swtrapdoor_hit()
    Playsound "fx_metal_scoop_hit":
    if trapdoor.rotx=0 then
            me.uservalue = -1
            me.timerinterval = 10
            me.timerenabled = 1
        end if
End Sub
Sub swtrapdoor_timer
    trapdoor.rotx = trapdoor.rotx + (me.uservalue * 3)
    if trapdoor.rotx < - 83.5 Then
        me.uservalue = 1
        swtrapdoor.timerenabled = 0
        trapdoorwall.collidable = 1
        vpmtimer.addtimer 4500, "swtrapdoor.timerenabled = 1 '"
    end If
    if trapdoor.rotx > 0 Then
        trapdoor.rotx = 0
        me.timerenabled = 0
        trapdoorwall.collidable = 0
    end if
End Sub

Sub swscoophitMIDDLE_hit
    Playsound "fx_scoop_hit"
End Sub

Sub swscoophitTOP_hit
    Playsound "fx_scoop_hit"
End Sub

Sub livingroomwirewall_hit
    Playsound "fx_metal_hit_4"
End Sub

'******************************************************
'             GENERAL ILLUMINATION
'******************************************************

Sub GIRelay(enabled)
    If Enabled Then
         for each xx in GI:xx.State=0:Next
'         for each xx in GIbulbs:xx.intensityscale=0:Next
        Table1.ColorGradeImage = "ColorGrade_1"
        PlaysoundAt "fx_relay_off", BackWall
        kickbacklight.state= 0:l99.state= 0:mcvickerlight.state= 0:buzzcutlight.state= 0:buttheadpflight.state= 0:beavispflight.state= 0:beavisrampplasticlight.state= 0:buttheadrampplasticlight.state= 0:BWSignLight.state= 0:Stopsound "B&BWarning"
    Else
         for each xx in GI:xx.State=1:Next
'         for each xx in GIbulbs:xx.intensityscale=1:Next
        Table1.ColorGradeImage = "ColorGrade_8"
        PlaysoundAt "fx_relay_on", BackWall
        If kickbackenabled Then kickbacklight.state=1:l99.state=0
    End If
End Sub

Sub TiltRelay(enabled)
    If Enabled Then
         for each xx in GI:xx.State=0:Next
'         for each xx in GIbulbs:xx.intensityscale=0:Next
        Table1.ColorGradeImage = "ColorGrade_1"
        PlaysoundAt "fx_relay_off", BackWall
        kickbacklight.state= 0:l99.state= 0:mcvickerlight.state= 0:buzzcutlight.state= 0:buttheadpflight.state= 0:beavispflight.state= 0:beavisrampplasticlight.state= 0:buttheadrampplasticlight.state= 0:BWSignLight.state= 0:Stopsound "B&BWarning"
    Else
         for each xx in GI:xx.State=1:Next
'         for each xx in GIbulbs:xx.intensityscale=1:Next
        Table1.ColorGradeImage = "ColorGrade_8"
        PlaysoundAt "fx_relay_on", BackWall
        If kickbackenabled Then kickbacklight.state=1:l99.state=0
    End If
End Sub

' *********************************************************************
'           Lighting
' *********************************************************************

Dim LampState(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashLevel(200)

Sub InitLamps
  On Error Resume Next
  Dim i
  For i=0 To 200: Execute "Lights(" & i & ")  = Array (L" & i & ",L" & i & "a)": Next
    For i = 0 to 200
        LampState(i) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FlashSpeedUp(i) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(i) = 0.35 ' slower speed when turning off the flasher
        FlashLevel(i) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
  LampTimer.Interval = 30:LampTimer.Enabled = 1
End Sub

Sub LampTimer_Timer()
  Dim i
  For i=0 to 128:LampState(i) = ABS(Controller.Lamp(i)):Next
  'Bulbs
  AddLamp 2, flasherspotlightRIGHT   'Beavis Spotlight
    AddLamp 2, spotlightRIGHTlight     'Beavis Spotlight
    AddLamp 2, flasherspotlightLEFT    'Butt-head Spotlight
    AddLamp 2, spotlightLEFTlight      'Butt-head Spotlight


    AddLamp 10, arrow5    'Burger World Menu "5.000"
    AddLamp 10, flasherorderhere      'Burger World "Order Here" Sign Reflection
  AddLamp 11, arrow4    'Burger World Menu "10.000"
  AddLamp 12, arrow3    'Burger World Menu "20.000"
  AddLamp 13, arrow2    'Burger World Menu "EXTRA BALL"
  AddLamp 14, arrow1    'Burger World Menu "SPECIAL"

    AddLamp 10, insert5   'Burger World Menu "5.000"
  AddLamp 11, insert4   'Burger World Menu "10.000"
  AddLamp 12, insert3   'Burger World Menu "20.000"
  AddLamp 13, insert2   'Burger World Menu "EXTRA BALL"
  AddLamp 14, insert1   'Burger World Menu "SPECIAL"

  AddLamp 13, L59       'Burger World Menu PF "EXTRA BALL"
  AddLamp 14, L58       'Burger World Menu PF "SPECIAL"

  AddLamp 36, drivethrugateaddletterflasher   'Right Ramp Add Letter
  AddLamp 37, drivethrugatemillionflasher         'Right Ramp Million
  AddLamp 46, L46flasher  'Left Ramp Add Letter
    AddLamp 47, L47flasher  'Left Ramp Lock Ball
    AddLamp 67, couchfishingbulbunlit  '"Couch Fishing Unlit"

    'Flashers
    AddLamp 5, L5flasher
    AddLamp 6, L6flasher
    AddLamp 7, L7flasher

    'Beavis TP Lights
    AddLamp 51, L51
    AddLamp 51, tplight1
    AddLamp 51, tplight1baseflasher
    AddLamp 52, L52
    AddLamp 52, tplight2
    AddLamp 52, tplight2baseflasher
    AddLamp 53, L53
    AddLamp 53, tplight3
    AddLamp 53, tplight3baseflasher

    'Butthead Skull Lights
    AddLamp 63, L63
    AddLamp 63, sw40skull
    AddLamp 63, sw40skullwallflasher
    AddLamp 63, sw40skullbaseflasher
    AddLamp 62, L62
    AddLamp 62, sw30skull
    AddLamp 62, sw30skullwallflasher
    AddLamp 62, sw30skullbaseflasher
    AddLamp 61, L61
    AddLamp 61, sw20skull
    AddLamp 61, sw20skullwallflasher
    AddLamp 61, sw20skullbaseflasher

    AddLamp 152, domeG    'Green Dome
  AddLamp 152, F22a
  AddLamp 152, F22r3
    AddLamp 152, F22ta
    AddLamp 152, middomebottomflasher
  AddLamp 153, domeR    'Red Dome
  AddLamp 153, F23
    AddLamp 153, rightdomebottomflasher
    AddLamp 153, F23a       'Under The Right Ramp
  AddLamp 153, F23r
  AddLamp 154, F24    'Under The Left Ramp
  AddLamp 154, domeB    'Blue Dome
  AddLamp 154, F24a
    AddLamp 154, F24r
    AddLamp 154, leftdomebottomflasher
  AddLamp 155, F25butthead
    AddLamp 155, F25beavis
  If LampState(152)=1 AND LampState(153)=1 Then HideLamp True
  If LampState(152)=0 OR LampState(153)=0 Then HideLamp False
    If f23.state = 1 then
        domeR.image = "domeredlit"
    else
        domeR.image = "domeredunlit"
End If
    If f22a.state = 1 then
        domeG.image = "domegreenlit"
        signshadow.visible = 1
        screwTAlit.visible = 1
        screwTA.visible = 0
        f22b.state = 1
        f22c.state = 1
    else
        domeG.image = "domegreenunlit"
        signshadow.visible = 0
        screwTAlit.visible = 0
        screwTA.visible = 1
        f22b.state = 0
        f22c.state = 0
End If
    If f24a.state = 1 then
        domeB.image = "domebluelit"
    else
        domeB.image = "domeblueunlit"
End If
    If light4.state = 1 Then
        hornsLEFTshaftSHADOW.visible = 1
        flasherSlingHornsLeft1.visible = 1
        flasherSlingHornsLeft2.visible = 1
    else
        hornsLEFTshaftSHADOW.visible = 0
        flasherSlingHornsLeft1.visible = 0
        flasherSlingHornsLeft2.visible = 0
End If
    If light14.state = 1 Then
        sw10bumperflasherA.visible = 1
        sw10bumperflasherB.visible = 1
        bumperbulb14on.visible = 1
        bumperbulb14off.visible = 0
    else
        sw10bumperflasherA.visible = 0
        sw10bumperflasherB.visible = 0
        bumperbulb14on.visible = 0
        bumperbulb14off.visible = 1
End If
    If light12.state = 1 Then
        sw12bumperflasherA.visible = 1
        sw12bumperflasherB.visible = 1
        bumperbulb13on.visible = 1
        bumperbulb13off.visible = 0
    else
        sw12bumperflasherA.visible = 0
        sw12bumperflasherB.visible = 0
        bumperbulb13on.visible = 0
        bumperbulb13off.visible = 1
End If
    If light16.state = 1 Then
        sw16bumperflasherA.visible = 1
        sw16bumperflasherB.visible = 1
        bumperbulb16on.visible = 1
        bumperbulb16off.visible = 0
    else
        sw16bumperflasherA.visible = 0
        sw16bumperflasherB.visible = 0
        bumperbulb16on.visible = 0
        bumperbulb16off.visible = 1
End If
    If light19.state = 1 Then
        hornsRIGHTshaftSHADOW.visible = 1
        flasherSlingHornsRight1.visible = 1
        flasherSlingHornsRight2.visible = 1
    else
        hornsRIGHTshaftSHADOW.visible = 0
        flasherSlingHornsRight1.visible = 0
        flasherSlingHornsRight2.visible = 0
End If
    If remotesidelight.state = 1 Then
        remotesidelightflasher.visible = 1
        remotetbuttheadflasher.visible = 1
    else
        remotesidelightflasher.visible = 0
        remotetbuttheadflasher.visible = 0
End If
    If l67.state = 1 Then
        l67arrow.state = 1
        l67hook.state = 1
    else
        l67arrow.state = 0
        l67hook.state = 0
End If
    If l67a.state = 1 Then
        l67b.state = 1
        couchfishingbulblit.visible = 1
        L67c.visible = 1
        l67d.state = 1
        L67e.visible = 1
        L67f.visible = 1
    else
        l67b.state = 0
        couchfishingbulblit.visible = 0
        L67c.visible = 0
        l67d.state = 0
        L67e.visible = 0
        L67f.visible = 0
End If
    If stewartspotlightLIGHT.state = 1 Then
        stewartspinnerflasher.visible = 1
        flasherstewartglow.visible = 1
    else
        stewartspinnerflasher.visible = 0
        flasherstewartglow.visible = 0
End If
    If billboardlight1.state = 1 Then
        flasherbillboard1.visible = 1
        backwallhillarrowLEFT.Image = "backwallhillarrowLEFTshadow"
    else
        flasherbillboard1.visible = 0
        backwallhillarrowLEFT.Image = "backwallhillarrows"
End If
    If billboardlight2.state = 1 Then
        flasherbillboard2.visible = 1
        backwallhillarrowMID.Image = "backwallhillarrowMIDDLEshadow"
    else
        flasherbillboard2.visible = 0
        backwallhillarrowMID.Image = "backwallhillarrows"
End If
    If billboardlight3.state = 1 Then
        flasherbillboard3.visible = 1
        backwallhillarrowRIGHT.Image = "backwallhillarrowRIGHTshadow"
    else
        flasherbillboard3.visible = 0
        backwallhillarrowRIGHT.Image = "backwallhillarrows"
End If
    If l46.state = 1 Then
        nachosgateleftbulblit.visible = 1
    else
        nachosgateleftbulblit.visible = 0
End If
    If l47.state = 1 Then
        nachosgaterightbulblit.visible = 1
    else
        nachosgaterightbulblit.visible = 0
End If
    If l36.state = 1 Then
        drivethrugateaddletter.visible = 1
    else
        drivethrugateaddletter.visible = 0
End If
    If l37.state = 1 Then
        drivethrugatemillion.visible = 1
    else
        drivethrugatemillion.visible = 0
End If
    If tvlightboobtubeA.state = 1 Then
        tvscreenboobs.visible = 1
        flashertvboobtubebottom.visible = 1
        flashertvboobtubewall.visible = 1
        tvframelitboobtube.visible = 1
        tvscreenreflectionB.visible = 1
        tvscreenboobwatchreflection.visible = 1
        tvcouchreflectionboobtube.visible = 1
        flashertvexitboobtube.visible = 1
        railinsidetvglowA.visible = 1
    else
        tvscreenboobs.visible = 0
        flashertvboobtubebottom.visible = 0
        flashertvboobtubewall.visible = 0
        tvframelitboobtube.visible = 0
        tvscreenreflectionB.visible = 0
        tvscreenboobwatchreflection.visible = 0
        tvcouchreflectionboobtube.visible = 0
        flashertvexitboobtube.visible = 0
        railinsidetvglowA.visible = 0
End If
    If tvlightA.state = 1 Then
        tvscreenstatic.visible = 1
        flashertvbottom.visible = 1
        flashertvwall.visible = 1
        tvframelitstatic.visible = 1
        tvscreenreflectionA.visible = 1
        tvscreenstaticreflection.visible = 1
        tvcouchreflectionstatic.visible = 1
        flashertvexitStatic.visible = 1
        railinsidetvglowB.visible = 1
     else
        tvscreenstatic.visible = 0
        flashertvbottom.visible = 0
        flashertvwall.visible = 0
        tvframelitstatic.visible = 0
        tvscreenreflectionA.visible = 0
        tvscreenstaticreflection.visible = 0
        tvcouchreflectionstatic.visible = 0
        flashertvexitStatic.visible = 0
        railinsidetvglowB.visible = 0
End If
    If tablelamplightA.state = 1 Then
       tablelamplightB.state = 1
       lampshadelit.visible = 1
       lampcolumnlit.visible = 1
       couchlamplight.visible = 1
    else
       tablelamplightB.state = 0
       lampshadelit.visible = 0
       lampcolumnlit.visible = 0
       couchlamplight.visible = 0
End If
    If kickbacklight.state = 1 Then
        beaviskickinglight.visible = 1
        beaviskickingflasher.visible = 1
    else
        beaviskickinglight.visible = 0
        beaviskickingflasher.visible = 0
End If
    If light2.state = 1 Then
        beaviskickingscoopshadow.visible = 1
    else
        beaviskickingscoopshadow.visible = 0

End If
    If mtvlight.state = 1 Then
        mtvlogowallflasher.visible = 1
        mtvlogofrontflasher.visible = 1
        mtvlogoLIT.visible = 1
        mtvlogoleftwallflasher.visible = 1
        mtvlogobracketLIT.visible = 1
        mtvlogobracket.visible = 0
        billboardmtvglow.visible = 1
    else
        mtvlogowallflasher.visible = 0
        mtvlogofrontflasher.visible = 0
        mtvlogoLIT.visible = 0
        mtvlogoleftwallflasher.visible = 0
        mtvlogobracketLIT.visible = 0
        mtvlogobracket.visible = 1
        billboardmtvglow.visible = 0
End If
    If skull3light.state = 1 Then
       skull3lightbaseflasher.visible = 1
       skull3lightwallflasher.visible = 1
       skull3LIT.visible = 1
       sw40skullwireLIT2.visible = 1
       sw40skullwire.visible = 0
    else
       skull3lightbaseflasher.visible = 0
       skull3lightwallflasher.visible = 0
       skull3LIT.visible = 0
       sw40skullwireLIT2.visible = 0
       sw40skullwire.visible = 1
End If
    If l63.state = 1 Then
       sw40skullwireLIT.visible = 1
       sw40skullwire.visible = 0
    else
       sw40skullwireLIT.visible = 0
       sw40skullwire.visible = 1
End If
    If skull2light.state = 1 Then
       skull2lightbaseflasher.visible = 1
       skull2lightwallflasher.visible = 1
       skull2LIT.visible = 1
       sw30skullwireLIT2.visible = 1
       sw30skullwire.visible = 0
    else
       skull2lightbaseflasher.visible = 0
       skull2lightwallflasher.visible = 0
       skull2LIT.visible = 0
       sw30skullwireLIT2.visible = 0
       sw30skullwire.visible = 1
End If
    If l62.state = 1 Then
       sw30skullwireLIT.visible = 1
       sw30skullwire.visible = 0
    else
       sw30skullwireLIT.visible = 0
       sw30skullwire.visible = 1
End If
    If skull1light.state = 1 Then
       skull1lightbaseflasher.visible = 1
       skull1lightwallflasher.visible = 1
       skull1LIT.visible = 1
       sw20skullwireLIT2.visible = 1
       sw20skullwire.visible = 0
    else
       skull1lightbaseflasher.visible = 0
       skull1lightwallflasher.visible = 0
       skull1LIT.visible = 0
       sw20skullwireLIT2.visible = 0
       sw20skullwire.visible = 1
End If
    If l61.state = 1 Then
       sw20skullwireLIT.visible = 1
       sw20skullwire.visible = 0
    else
       sw20skullwireLIT.visible = 0
       sw20skullwire.visible = 1
End If
    If swburgerworldlights.state = 1 Then
       bwwindowsleftflasher.visible = 1
       bwwindowsmiddleflasher.visible = 1
       bwwindowsrightflasher.visible = 1
       flasherbwleft.visible = 1
       flasherbwleft2.visible = 1
       flasherbwright.visible = 1
       flasherbwright2.visible = 1
       bwplasticbandbflasher.visible = 1
       flasherbwtree.visible = 1
       flasherBWtallbushRwallSHADOW.visible = 1
    else
       bwwindowsleftflasher.visible = 0
       bwwindowsmiddleflasher.visible = 0
       bwwindowsrightflasher.visible = 0
       flasherbwleft.visible = 0
       flasherbwleft2.visible = 0
       flasherbwright.visible = 0
       flasherbwright2.visible = 0
       bwplasticbandbflasher.visible = 0
       flasherbwtree.visible = 0
       flasherBWtallbushRwallSHADOW.visible = 0
End If

'Frog Spotlights
if frogspotlight.state = 1 then
        flasherFROGspotlightTOP.visible = 1
        flasherFROGspotlightBOTTOM.visible = 1
        frogshadow.visible = 0
    else
        flasherFROGspotlightTOP.visible = 0
        flasherFROGspotlightBOTTOM.visible = 0
        frogshadow.visible = 1
    end If

'Burger World Sign - Beavis and Butthead Sounds
if BWSignLight.state = 1 then
    BurgerWorldSignLIT2.visible = 1
    else
    BurgerWorldSignLIT2.visible = 0
    end if

'Left, Right and Center Spot and Drop Target Bank Shadows
if sw16.isdropped then
        sw16shadow.visible = 0
    else
        sw16shadow.visible = 1
    end If

if sw26.isdropped then
        sw26shadow.visible = 0
    else
        sw26shadow.visible = 1
    end If

if sw36.isdropped then
        sw36shadow.visible = 0
    else
        sw36shadow.visible = 1
    end If

if sw46.isdropped then
        sw46shadow.visible = 0
    else
        sw46shadow.visible = 1
    end If

if sw17.isdropped then
        sw17shadow.visible = 0
    else
        sw17shadow.visible = 1
    end If

if sw27.isdropped then
        sw27shadow.visible = 0
    else
        sw27shadow.visible = 1
    end If

if sw37.isdropped then
        sw37shadow.visible = 0
    else
        sw37shadow.visible = 1
    end If

if sw47.isdropped then
        sw47shadow.visible = 0
    else
        sw47shadow.visible = 1
    end If

if light6.state = 0 then
    sw16shadow.visible = 0
    sw26shadow.visible = 0
    sw36shadow.visible = 0
    sw46shadow.visible = 0
    end If

if light17.state = 0 then
    sw17shadow.visible = 0
    sw27shadow.visible = 0
    sw37shadow.visible = 0
    sw47shadow.visible = 0
    end If

if light6.state = 0 then
        leftspottargetshadows.visible = 0
    else
        leftspottargetshadows.visible = 1
    end If

if light17.state = 0 then
        rightspottargetshadows.visible = 0
    else
        rightspottargetshadows.visible = 1
    end If

if light23.state = 0 then
        sw31shadow.visible = 0
    else
        sw31shadow.visible = 1
    end If

if l47.state = 1 then
        sw21.Image = "target-green-reflect"
    else
        sw21.Image = "target-green"
    end If

if spotlightLEFTspecial.state = 1 then
        flasherspotlightLEFTloseball.visible = 1
    Else
        flasherspotlightLEFTloseball.visible = 0
end if

if spotlightRIGHTspecial.state = 1 then
        flasherspotlightRIGHTloseball.visible = 1
        tplight1baseflasher.visible = 0
        tplight2baseflasher.visible = 0
        tplight3baseflasher.visible = 0
        tplight1.visible = 0
        tplight1OFF.visible = 1
        tplight2.visible = 0
        tplight2OFF.visible = 1
        tplight3.visible = 0
        tplight3OFF.visible = 1
    else
        flasherspotlightRIGHTloseball.visible = 0
        tplight1.visible = 1
        tplight1OFF.visible = 0
        tplight2.visible = 1
        tplight2OFF.visible = 0
        tplight3.visible = 1
        tplight3OFF.visible = 0
end if

if spotlightRIGHTlight.state = 1 then
        tplight1baseflasher.visible = 0
        tplight2baseflasher.visible = 0
        tplight3baseflasher.visible = 0
        tplight1.visible = 0
        tplight1OFF.visible = 1
        tplight2.visible = 0
        tplight2OFF.visible = 1
        tplight3.visible = 0
        tplight3OFF.visible = 1
end if

if l46.state = 1 then
        flasherleftbulbonplastic.visible = 1
    else
        flasherleftbulbonplastic.visible = 0
    end If

drivethrumenuspeakerlight.visible = 1

if speakerlight.state = 1 then
        drivethrumenuspeakerlight.visible = 1
    else
        drivethrumenuspeakerlight.visible = 0
    end If

'DT vs. CAB Flashers
If Table1.ShowDT = true then
if swburgerworldlights.state = 1 then
        flasherBWleftwindowreflectDT.visible = 1
        flasherBWmidwindowreflectDT.visible = 1
        flasherBWrightwindowreflectDT.visible = 1
        mtvlogoplasticreflectionDT.visible = 1
    else
        flasherBWleftwindowreflectDT.visible = 0
        flasherBWmidwindowreflectDT.visible = 0
        flasherBWrightwindowreflectDT.visible = 0
        mtvlogoplasticreflectionDT.visible = 0
    end If

if F24a.state = 1 then
        leftdomebulbflasherDT.visible = 1
    else
        leftdomebulbflasherDT.visible = 0
    end If

if F22a.state = 1 then
        middomebulbflasherDT.visible = 1
    else
        middomebulbflasherDT.visible = 0
    end If

if F23.state = 1 then
        rightdomebulbflasherDT.visible = 1
    else
        rightdomebulbflasherDT.visible = 0
    end If

if mtvlight.state = 1 then
        mtvlogoplasticreflectionDT.visible = 1
    else
        mtvlogoplasticreflectionDT.visible = 0
    end If

if l36.state = 1 then
        flasherLetterDT.visible = 1
    else
        flasherLetterDT.visible = 0
    end If

if l37.state = 1 then
        flasherMillionDT.visible = 1
    else
        flasherMillionDT.visible = 0
    end If

if l46.state = 1 then
        L46flasherPFreflectDT.visible = 1
    else
        L46flasherPFreflectDT.visible = 0
    end If

if l47.state = 1 then
        L47flasherPFreflectDT.visible = 1
    else
        L47flasherPFreflectDT.visible = 0
    end If

if l67arrow.state = 1 then
        flasherarrowhookDT.visible = 1
    else
        flasherarrowhookDT.visible = 0
    end If

if l3.state = 1 then
        flasherEBleftDT.visible = 1
    else
        flasherEBleftDT.visible = 0
    end If

if l4.state = 1 then
        flasherEBrightDT.visible = 1
    else
        flasherEBrightDT.visible = 0
    end If

if l5.state = 1 then
        flasher1BackHillDT.visible = 1
    else
        flasher1BackHillDT.visible = 0
    end If

if l6.state = 1 then
        flasher2BackHillDT.visible = 1
    else
        flasher2BackHillDT.visible = 0
    end If

if l7.state = 1 then
        flasher3BackHillDT.visible = 1
    else
        flasher3BackHillDT.visible = 0
    end If

if l15.state = 1 then
        flasherButtheadDT.visible = 1
    else
        flasherButtheadDT.visible = 0
    end If

if l16.state = 1 then
        flasherAndDT.visible = 1
    else
        flasherAndDT.visible = 0
    end If

if l17.state = 1 then
        flasherBeavisDT.visible = 1
    else
        flasherBeavisDT.visible = 0
    end If

if carllight.state = 1 then
        flasher1CARLDT.visible = 1
        flasher1CARLwindow.visible = 1
        flasher1CARLbw.visible = 1
        flasher1CARLbw2.visible = 1
    else
        flasher1CARLDT.visible = 0
        flasher1CARLwindow.visible = 0
        flasher1CARLbw.visible = 0
        flasher1CARLbw2.visible = 0
    end If

Else 'if in FS mode

if swburgerworldlights.state = 1 then
        flasherBWleftwindowreflectCAB.visible = 1
        flasherBWmidwindowreflectCAB.visible = 1
        flasherBWrightwindowreflectCAB.visible = 1
        mtvlogoplasticreflectionCAB.visible = 1
    else
        flasherBWleftwindowreflectCAB.visible = 0
        flasherBWmidwindowreflectCAB.visible = 0
        flasherBWrightwindowreflectCAB.visible = 0
        mtvlogoplasticreflectionCAB.visible = 0
    end If

if F24a.state = 1 then
        leftdomebulbflasherCAB.visible = 1
    else
        leftdomebulbflasherCAB.visible = 0
    end If

if F22a.state = 1 then
        middomebulbflasherCAB.visible = 1
    else
        middomebulbflasherCAB.visible = 0
    end If

if F23.state = 1 then
        rightdomebulbflasherCAB.visible = 1
    else
        rightdomebulbflasherCAB.visible = 0
    end If

if l36.state = 1 then
        flasherLetterCAB.visible = 1
    else
        flasherLetterCAB.visible = 0
    end If

if l37.state = 1 then
        flasherMillionCAB.visible = 1
    else
        flasherMillionCAB.visible = 0
    end If

if mtvlight.state = 1 then
        mtvlogoplasticreflectionCAB.visible = 1
    else
        mtvlogoplasticreflectionCAB.visible = 0
    end If

if l46.state = 1 then
        L46flasherPFreflectCAB.visible = 1
    else
        L46flasherPFreflectCAB.visible = 0
    end If

if l47.state = 1 then
        L47flasherPFreflectCAB.visible = 1
    else
        L47flasherPFreflectCAB.visible = 0
    end If

if l67arrow.state = 1 then
        flasherarrowhookCAB.visible = 1
    else
        flasherarrowhookCAB.visible = 0
    end If

if l5.state = 1 then
        flasher1BackHillCAB.visible = 1
    else
        flasher1BackHillCAB.visible = 0
    end If

if l6.state = 1 then
        flasher2BackHillCAB.visible = 1
    else
        flasher2BackHillCAB.visible = 0
    end If

if l7.state = 1 then
        flasher3BackHillCAB.visible = 1
    else
        flasher3BackHillCAB.visible = 0
    end If

if carllight.state = 1 then
        flasher1CARLCAB.visible = 1
        flasher1CARLwindow.visible = 1
        flasher1CARLbw.visible = 1
        flasher1CARLbw2.visible = 1
    else
        flasher1CARLCAB.visible = 0
        flasher1CARLwindow.visible = 0
        flasher1CARLbw.visible = 0
        flasher1CARLbw2.visible = 0
    end If
End If
End Sub

'Billboard Lights
Sub swbillboardlight1_hit
    billboardlight1.state = 0
End Sub

Sub swbillboardlight2_hit
    billboardlight2.state = 0
End Sub

Sub swbillboardlight3_hit
    billboardlight3.state = 0
End Sub

Sub swbillboardlightson_hit
    billboardlight1.state = 1
    billboardlight2.state = 1
    billboardlight3.state = 1
    Playsound "fx_metal_hit_3"
End Sub

'Burger World Lights
Sub swBWlightson_hit
    swburgerworldlights.state = 1
    carllight.state = 0
End Sub

Sub swBWlightsoff_hit
    swburgerworldlights.state = 0
if l36.state=1 Then
    Dim bwdt
  bwdt = INT(3 * RND(1) )
  Select Case bwdt
  Case 0:Playsound "fx_butthead_drive_thru_please":speakerlight.duration 1, 1600, 0
  Case 1:Playsound "fx_beavis_thank_you_drive_thru":speakerlight.duration 1, 1400, 0
    Case 2:Playsound "fx_butthead_drive_thru_dumbass":speakerlight.duration 1, 3300, 0
    End Select
End If
if l37.state=1 Then
    Stopsound "fx_butthead_drive_thru_please":speakerlight.state = 0
  Stopsound "fx_beavis_thank_you_drive_thru":speakerlight.state = 0
    Stopsound "fx_butthead_drive_thru_dumbass":speakerlight.state = 0
End If
if l44.state=1 Then
    Stopsound "fx_butthead_drive_thru_please":speakerlight.state = 0
  Stopsound "fx_beavis_thank_you_drive_thru":speakerlight.state = 0
    Stopsound "fx_butthead_drive_thru_dumbass":speakerlight.state = 0
End If
End Sub

Sub SetLamp(nr, enabled)
    If enabled Then
    LampState(nr) = 1
  Else
    LampState(nr) = 0
  End If
End Sub

Sub HideLamp (enabled)
End Sub

Sub AddLamp(nr, object)
    Select Case LampState(nr)

        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < 0 Then FlashLevel(nr) = 0
      If TypeName(object) = "Light" Then
        Object.State = 0
      End If
      If TypeName(object) = "Flasher" Then
        Object.IntensityScale = FlashLevel(nr)
      End If
      If TypeName(object) = "Primitive" Then
        Object.DisableLighting = 0
      End If

        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > 1 Then FlashLevel(nr) = 1
      If TypeName(object) = "Light" Then
        Object.State = 1
      End If
      If TypeName(object) = "Flasher" Then
        Object.IntensityScale = FlashLevel(nr)
      End If
      If TypeName(object) = "Primitive" Then
        Object.DisableLighting = 1
      End If

    End Select
End Sub

Sub SwapTex (nr,object, texoff, texon)
    Select Case LampState(nr)
        Case 0 'off
      Object.DisableLighting = 0
      Object.image = texoff
        Case 1 ' on
      Object.DisableLighting = 1
      Object.image = texon
    End Select
End Sub

' Remote Control Buttons & PF Bulbs
Sub FlasherTimer_Timer
    l85pf.state = l85.state
    remoteGIbuttons.visible = l001.state
    button85.visible = l85.state
    button86.visible = l86.state
    button90.visible = l90.state
    button91.visible = l91.state
    button92.visible = l92.state
    button93.visible = l93.state
    button94.visible = l94.state
    button95.visible = l95.state
    button96.visible = l96.state
    button97.visible = l97.state
    button100.visible = l100.state
    button101.visible = l101.state
    button102.visible = l102.state
    button103.visible = l103.state
    button104.visible = l104.state
    button105.visible = l105.state
    button106.visible = l106.state
    button107.visible = l107.state
    button110.visible = l110.state
    button111.visible = l111.state
    button112.visible = l112.state
    button113.visible = l113.state
    button114.visible = l114.state
    button115.visible = l115.state
    button116.visible = l116.state
    button117.visible = l117.state

' Playfield Bulb Prims
     light1bulbon.visible = light1.state
     light2bulbon.visible = light2.state
     light3bulbon.visible = light3.state
     light4bulbon.visible = light4.state
     light5bulbon.visible = light5.state
     light6bulbon.visible = light6.state
     light7bulbon.visible = light7.state
     light8bulbon.visible = light8.state
     light8bulbon.visible = f24.state
     light9bulbon.visible = light9.state
     light12bulbon.visible = light12.state
     light12bulbon.visible = f23a.state
     light15bulbon.visible = light15.state
     light17bulbon.visible = light17.state
     light18bulbon.visible = light18.state
     light19bulbon.visible = light19.state
     light20bulbon.visible = light20.state
     light21bulbon.visible = light21.state
     light22bulbon.visible = light22.state
     light23bulbon.visible = light23.state
     light24bulbon.visible = light24.state

' Pop Bumper Bulbs
     bumperbulb14on.visible = light14.state
     bumperbulb16on.visible = light16.state
     bumperbulb13on.visible = light13.state
End Sub

' *********************************************************************
'         Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function AudioPan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function AudioFade(ball)
    Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

'*********************************************************************
'                 Positional Sound Playback Routines
'*********************************************************************

Sub PlaySoundXY(soundname, loopcount, volume, useexisting, tableobj)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound soundname, loopcount, volume, AudioPan(tableobj), 0, Pitch(tableobj), useexisting, 0, AudioFade(tableobj)
  Else
    PlaySound soundname, loopcount, volume, AudioPan(tableobj), 0, Pitch(tableobj), useexisting, 0
  End If
End Sub

Sub PlaySoundAt(soundname, tableobj)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
  Else
    PlaySound soundname, 1, 1, AudioPan(tableobj)
  End If
End Sub

Sub PlaySoundAtBall(soundname)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound soundname, 0, 10*Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  Else
    PlaySound soundname, 0, 10*Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End If
End Sub

' *********************************************************************
'               Maths
' *********************************************************************
Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

' *********************************************************************
'             Other Sound FX
' *********************************************************************
'Ball Collision
Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySoundXY "fx_collide", 0, Csng(velocity) ^2 / 500, 0, ball1
End Sub

'Gates Collision
Sub LGate_hit():PlaysoundAt "fx_gate4", ActiveBall:End Sub
Sub RGate_hit():PlaysoundAt "fx_gate4", ActiveBall:End Sub

'Flippers Collision
Sub LeftFlipper_Collide(parm):RandomSoundFlipper:End Sub
Sub Rightflipper_Collide(parm):RandomSoundFlipper:End Sub

'Rubber Posts Collision
Sub RubberPosts_Hit()
  dim finalspeed
    finalspeed=SQR((activeball.velx ^2) + (activeball.vely ^2))
  If finalspeed > 20 then
    PlaySoundAtBall "fx_rubber"
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

'Rubber Collision
Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR((activeball.velx ^2) + (activeball.vely ^2))
  If finalspeed > 20 then
    PlaySoundAtBall "fx_rubber"
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case RndNum(1,3)
    Case 1 : PlaySoundAtBall "fx_rubber_hit_1"
    Case 2 : PlaySoundAtBall "fx_rubber_hit_2"
    Case 3 : PlaySoundAtBall "fx_rubber_hit_3"
  End Select
End Sub

Sub RandomSoundFlipper()
  Select Case RndNum(1,3)
    Case 1 : PlaySoundAtBall "fx_flip_hit_1"
    Case 2 : PlaySoundAtBall "fx_flip_hit_2"
    Case 3 : PlaySoundAtBall "fx_flip_hit_3"
  End Select
End Sub

' *********************************************************************
'           Ball Drop & Ramp Sounds
' *********************************************************************
'left ramp
Sub LREnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAt "fx_rampenterplasticrolling", ActiveBall:PlaySoundAt "fx_leftrampenter", ActiveBall:End If:End Sub
Sub LREnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_rampenterplasticrolling":StopSound "fx_leftrampenter":End If:End Sub
Sub LRail_hit():StopSound "fx_rampenterplasticrolling":PlaySoundAt "fx_wireramp", ActiveBall:End Sub
Sub LRailHelp1_hit:StopSound "fx_wireramp":PlaySoundAt "", ActiveBall:End Sub
Sub LRExit_Hit:StopSound "fx_wireramp":PlaysoundAt "fx_wireramp_exit", ActiveBall:Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub LRExit_timer:Me.TimerEnabled=0:PlaysoundAt "fx_balldrop", LRExit:End Sub

'middle ramp
Sub MRail_hit():PlaySoundAt "fx_wireramp", ActiveBall:End Sub

'right ramp
Sub RREnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAt "fx_rampenterplasticrolling", ActiveBall:End If:
    If ActiveBall.VelY < 0 Then
    Playsound "fx_car_accelerate"
End If
End Sub
Sub RREnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_rampenterplasticrolling":StopSound "fx_car_accelerate":End If:End Sub
Sub RRail_Hit:StopSound "fx_rampenterplasticrolling":PlaySoundAt "fx_wireramp", ActiveBall:End Sub
Sub RRExit_Hit:StopSound "fx_wireramp":PlaysoundAt "fx_wireramp_exit", ActiveBall:Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub RRExit_timer:Me.TimerEnabled=0:PlaysoundAt "fx_balldrop", RRExit:End Sub
Sub RRailHit_Hit():PlaySoundAt "fx_metal_hit_2", ActiveBall:End Sub

'******************************************************
'         RealTime Updates
'******************************************************
Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates()
  RollingSoundUpdate
  BallShadowUpdate
  Prim_LeftFlipper.RotY=LeftFlipper.currentangle-90
  Prim_RightFlipper.RotY=RightFlipper.currentangle-90
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  RightGate.RotX = - RGate.currentangle
  sw15s.RotX = - sw15.currentangle
    stewartspinnerprim.RotX = - stewartspinner.currentangle
End Sub

'*********** ROLLING SOUND *********************************
Const tnob = 2            ' total number of balls : 2 (trough)
ReDim rolling(tnob)

Sub InitRolling:Dim i:For i=0 to (tnob-1):rolling(i) = False:Next:End Sub

Sub RollingSoundUpdate()
    Dim BOT,b
    BOT = GetBalls
  ' play the rolling sound for each ball
  For b = 0 to Ubound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySoundXY "fx_ballrolling" & (b+1), -1, Vol(BOT(b)), 1, BOT(b)
    Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & (b+1))
                rolling(b) = False
            End If
        End If
    Next
End Sub

'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2)

Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls
  ' render the shadow for each ball
    For b = 0 to Ubound(BOT)
    If BOT(b).X < Table1.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) - 10
    End If
      ballShadow(b).Y = BOT(b).Y + 15
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(40)
Digits(0) = Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
Digits(1) = Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
Digits(2) = Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
Digits(3) = Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
Digits(4) = Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
Digits(5) = Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
Digits(6) = Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
Digits(7) = Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
Digits(8) = Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
Digits(9) = Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
Digits(10) = Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
Digits(11) = Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
Digits(12) = Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
Digits(13) = Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
Digits(14) = Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
Digits(15) = Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)

Digits(16) = Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
Digits(17) = Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
Digits(18) = Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
Digits(19) = Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
Digits(20) = Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
Digits(21) = Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
Digits(22) = Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
Digits(23) = Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
Digits(24) = Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
Digits(25) = Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
Digits(26) = Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
Digits(27) = Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
Digits(28) = Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
Digits(29) = Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
Digits(30) = Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
Digits(31) = Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)

Digits(32) = Array(c00, c05, c0c, c0d, c08, c01, c06, c0f, c02, c03, c04, c07, c0b, c0a, c09, c0e)
Digits(33) = Array(c10, c15, c1c, c1d, c18, c11, c16, c1f, c12, c13, c14, c17, c1b, c1a, c19, c1e)
Digits(34) = Array(c20, c25, c2c, c2d, c28, c21, c26, c2f, c22, c23, c24, c27, c2b, c2a, c29, c2e)
Digits(35) = Array(c30, c35, c3c, c3d, c38, c31, c36, c3f, c32, c33, c34, c37, c3b, c3a, c39, c3e)
Digits(36) = Array(c40, c45, c4c, c4d, c48, c41, c46, c4f, c42, c43, c44, c47, c4b, c4a, c49, c4e)
Digits(37) = Array(c50, c55, c5c, c5d, c58, c51, c56, c5f, c52, c53, c54, c57, c5b, c5a, c59, c5e)
Digits(38) = Array(c60, c65, c6c, c6d, c68, c61, c66, c6f, c62, c63, c64, c67, c6b, c6a, c69, c6e)
Digits(39) = Array(c70, c75, c7c, c7d, c78, c71, c76, c7f, c72, c73, c74, c77, c7b, c7a, c79, c7e)

 Sub DisplayTimer_Timer
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If DesktopMode Then
    If Not IsEmpty(ChgLED)Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        For Each obj In Digits(num)
          If chg And 1 Then obj.State=stat And 1
          chg=chg\2 : stat=stat\2
        Next
      Next
    End If
  End If
 End Sub
