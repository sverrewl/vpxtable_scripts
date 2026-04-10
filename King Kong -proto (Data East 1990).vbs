Option Explicit
Randomize

'==============================================================================================='
'                                               '
'                     King Kong                       '
'                                      DataEast (1990)                                    '
'                  https://www.ipdb.org/machine.cgi?id=3194                     '
'                                               '
'                             Created by: cyberpez                                '
'                                               '
'==============================================================================================='


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0


Dim PrototypeVersion, FlipperBatMod, SolidStateSticker, KingFlipperBatMod, KingFlipperSticker, InstructionCards, ApronType, RubberColor, VUKMod, MainFlipperMod, MainFlipperKingSticker, BallMod, EnableBallControl, FlasherMod, MysteryTarget, RubberPostColor, PegColor, VUKPlastics, CenterPlastic, LeftPlastic, LeftWireRampColor, RightWireRampColor, BlackRampDecal, FlasherMod2, LeftWireRampOption, LeftTargets, DTMod, KSRPlastics, WackerDrivePlastic, DropTargetDecal, SlingKongVersion, KingSizedDELogo, DateEastLogoColorMatch, UPFGate, PlasticRampWedge, BlackRampWedge1, BlackRampWedge2, DoubleRubbers, ApronHide, DPP, KSRSwitchPlastic, RightWireRampGuide, LeftSBLPPlastic, LeftSBLP, LeftWireRampSupportConfig, FlipperColorOverride, KSBRAorB, OptinalRails1, OptinalRails2, RandomInsertGI, BlackRampColor, BlackRampColorOverRide, Plastic8Height, GIColorMod, WackerPost, CabVisible, RailsVisible, WallsVisible, FloorVisible, BackGlassVisible, LockBarVisible, InsertDecals, BRGate, BackboardPlasticAorB
Dim DesktopMode: DesktopMode = KK.ShowDT



''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''' Options''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''



'''''Prototype version (my guesses)
'0=random
'1=Proto1 - VUK1
'2=Proto2 - VUK2
'3=Proto3 - VUK3
'4=Proto4 - VUK4
'5=Proto5 - KSR1
'6=Proto6 - KSR2
'7=Proto7 - KSR3
'8=Proto8 - KSR4
'9=Proto9 - KSR5
'10= CP's Production - VUK
'11= CP's Production - KSR
'12=custom (you set the options)
'13=MangaProtoType

PrototypeVersion = 5



''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''VR Settings''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Cab Visible
'0=no
'1=yes

CabVisible = 0
BackGlassVisible = 0

WallsVisible = 0
FloorVisible = 0

'Rails and LockBar
'0=no
'1=yes

RailsVisible = 0
LockBarVisible = 0

''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Table Customimization options'''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'GI ColorMod
'
'1=Normal
'2=ColorMod

GIColorMod = 1


'Ball Mod
'0=Normal balls
'1=Yellow/Orange/Red balls
'2=Marbled Balls

BallMod = 0


'''DropTargetDecal
'1=none
'2=Dino
'3=Kong-"original"
'4=Kong2
'5=Kong3

DropTargetDecal = 3



'Random Insert GI
'0=off
'1=Tied to GI

RandomInsertGI = 1


'BlackRamp Color
'0=Random
'1=Black
'2=Clear
'3=Damaged
'4=TigerStriped
'5=BlackMetal
'6=Orange
'7=Cream

BlackRampColor = 3

'This needs to be set to 1 if you want the Black Ramp color to be something besides black
BlackRampColorOverRide = 0



''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''VUK and KSR option'''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'0=KingSizedRamp
'1=VUK

VUKMod = 1

'Black Ramp Version  (King Sized Ramp Protos)
'1=Option A
'2=Option B
'3=Option C

KSBRAorB = 2

''''DropTarget Mod
'0=Standup
'1=Drop Targets

DTMod = 1


''FlasherMods
'Left
'1=Front
'2=Rear
FlasherMod = 2

'Right
'1=Red
'2=Yellow
FlasherMod2 = 2


''Mystery Target (as seen in one of the prototypes
''''' May or may not do anything, currently tied to sw44
''0=no target, only rubber
''1=Target

MysteryTarget = 1



''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''MainFlipperMods'''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''FlipperColor Override
'0=Normal Proto colors
'1=Overrides normal colors with custom Bellow

FlipperColorOverride = 0


'Main Flipper Size
'0 = Normal 1 = KingSized
MainFlipperMod = 1


'''Solid State Sticker (only normal sized flipper)
'0=Random
'1=not visible
'2=visible
SolidStateSticker = 1

'''Main Flipper (King Sized) ... King Sized Decal
'0=not visible   1=visible
MainFlipperKingSticker = 0


'''Main Fipper Colors
'1=white/black  2=yellow/orange   3=black/dark orange  4=Yellow/Red
FlipperBatMod = 3




''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''KingSizedFlipperMods''''''''''''''''''
'''''''''''''''''''''''Third Flipper''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Color
'1=white/black  2=yellow/orange   3=black/dark orange  4=Yellow/Red
KingFlipperBatMod = 2


'KS Flipper Sticker
'0=not visible   1=visible
KingFlipperSticker = 1



''''King Sized Data East log

'0=none
'1=King Sized Main Flipper
'2=King Sized Flipper (not active is KS Decal is enabled)
'3=Visible on all three King Sized Flippers

KingSizedDELogo = 1


''Date East logo color match
'1=Logo matches Flipper Color
'2=Logo Matches Flipper Rubber

DateEastLogoColorMatch = 1



''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''Apron''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''Apron
'1=DataEast
'2=King Kong

ApronType = 2


'''Instruction Cards
'1=normal
'2=custom

InstructionCards = 2




''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''Alt Plastics''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''SlingKong
'1=Proto
'2=Final
SlingKongVersion = 1


''Left Plastic
'1=Proto
'2=Final

LeftPlastic = 2

''Center Plastic
'1=None
'2=Final

CenterPlastic = 2


'VUK Plastics
'1=none
'2=almost none
'3=BlueConfig1
'4=BlueConfig2
'5=jungle

VUKPlastics = 2


'KingSizedRamp Plastics
'1=none
'2=option1
'3=option2
'4=jungle
'5=Proto9

KSRPlastics = 2


'WackerDrivePlastic
'0=none
'1=WackerDrive1
'2=WackerDrive2
WackerDrivePlastic = 2


'Backboard Plastic
'1=A
'2=B

BackboardPlasticAorB = 2


'Plastic8 - plastic by right back bumper - height.
'1=Normal
'2=A little higher than nomal

Plastic8Height = 2



''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Posts and Rubber'''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Rubber Color
'1=white  2=black  3=Yellow
RubberColor = 2


'Rubber Post color
'1=All Black
'2=All Yellow
'3=All Orange
'4=Color Option 1
'5=Color Option 2

RubberPostColor = 4


'Peg color
'0=
'1=Yellow
'2=Orange

PegColor = 2


'WackerDrivePost
'0=none
'1=Post
'2=Post+Rubber

WackerPost = 2


'Double Rubber (back Wall)
'0=single Rubbers
'1=Double Rubbers

DoubleRubbers = 1


''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Other Options''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''Insert Decals
'proto 9 had sticker decals over a few inserts
'1=None
'2=Visible

InsertDecals = 1


''''''Left Targets
'1=Round
'2=Triangle

LeftTargets = 2


'''WireRampColor
''Left
'1=Chrome
'2=Yellow
'3=TriColor

LeftWireRampColor = 3


''Ball Lock Switch protecter plastic

'0=none
'1=Black
'2=orange 1
'3=orange 2
'4=orange 3


LeftSBLPPlastic = 0


''Ball Lock Switch Protector
'0=none
'1=option1 - Shield
'2=option2 - Fome Blocks
'3=option3 - Posts and Rubbers

LeftSBLP = 4


'Left Wire Ramp Support Config
'1=Option1
'2=Option2
'3=Option3

LeftWireRampSupportConfig = 1


''Right Wire Ramp Color
'1=Chrome
'2=Yellow

RightWireRampColor = 2


''Right Wire Ramp Guide
'1=Wire Guide
'2=Plastic Guide

RightWireRampGuide = 1


'''Black Ramp Gate - Found on proto 9.. tied to switch....?
'0=no Gate
'1=Gate

BRGate = 1


'''UPF Gate
'0=no Gate
'1=Gate

UPFGate = 1


'''BlackRamp decal
'1=none
'2=visible

BlackRampDecal = 2


'RampWedge
'0=Hidden
'1=Visible
'2=Alt Visible
'3=Alt2 Visible

PlasticRampWedge = 0

BlackRampWedge1 = 0

BlackRampWedge2 = 0


'KSR Switch Plastic
KSRSwitchPlastic = 0


'Differnt Post Posistions
'1=
'2=
'3=

DPP = 2


'Rail options around Radar Lock
'1=Rail Option
'2=Wall Option
OptinalRails1 = 2


'Rail Behind KSFlipper
'1=Rail Option
'2=Wall Option
OptinalRails2 = 1




' DMD rotation
Const cDMDRotation      = 0         ' 0 or 1 for a DMD rotation of 90°

'Hide DMD




''''Testing
ApronHide = 0




'Ball Control
EnableBallControl = 1


' ball size
Const BallSize = 25
Const BallMass = 1

' VPinMAME ROM name

Const cGameName="kiko_a10"


' ***********************************************************************************************
' OPTIONS END
' ***********************************************************************************************



' ===============================================================================================
' some general constants and variables
' ===============================================================================================

'Const UseVPMModSol = 1

Const UseSolenoids    = 2
Const UseLamps      = 0
Const UseGI       = 1
Const UseSync       = 0
Const HandleMech    = 0

Const SSolenoidOn     = "SolOn"
Const SSolenoidOff    = "SolOff"
Const SCoin       = "Coin"
Const SKnocker      = "Knocker"

Dim I, x, obj, bsTrough, bsSaucer, dtRBank


LoadVPM "01560000", "DE.VBS", 3.26

' ===============================================================================================
' solenoids
' ===============================================================================================


'Solenoids
SolCallBack(16) = "MissileKick"
SolCallBack(25) = "kisort"
SolCallBack(26) = "KickBallToLane"
SolCallBack(27) = "CRBallLock"
SolCallBack(28) = "RadarKick"
SolCallBack(29) = "KickBallUp"  'VUK VERSION
SolCallBack(30) = "ResetDrops"  'TARGET BANK USED IN THIS VERSION
SolCallBack(32) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"    'Need to confirm
SolCallBack(1) = "SetLamp 101," 'King Shooter
SolCallBack(2) = "SetLamp 102," 'King Shooter
SolCallback(3) = "SetLamp 103," 'Ape Ramp BG Skull
SolCallback(4) = "SetLamp 104," 'BG NDSRS Shooter
SolCallback(5) = "SetLamp 105," 'BG Tower 25k
SolCallback(6) = "SetLamp 106," 'BG Tower 50k
SolCallback(7) = "SetLamp 107," 'BG Tower 100k
SolCallback(8) = "SetLamp 108," 'BG Tower Extra Ball
SolCallback(9) = "SetLamp 109," 'BG Tower Million
'SolCallback(10) = "SetLamp 110,"  'Was Not Used and Still Not
SolCallback(11) = "SetGI " ' General Illum
SolCallback(12) = "Lamp112 " 'BG Girl Missile
SolCallback(13) = "SetLamp 113," 'BG Kong Explode
SolCallback(14) = "Lamp114 " 'Tower Flasher
SolCallback(15) = "SetLamp 115," 'Skull Island
'SolCallback(22) = NOT USED
'SolCallback(23) = NOT USED
'SolCallback(24) = NOT USED


SolCallback(46) = "SolRFlipper"
SolCallback(45) = "SolR2Flipper"
SolCallback(48) = "SolLFlipper"

Function SoundFX (sound)
    If cController = 4 Then 'Table sounds disabled
        SoundFX = ""
    Else                    'Table sounds enabled
        SoundFX = sound
    End If
End Function


Dim Ball(6)
Dim InitTime
Dim TroughTime
Dim EjectTime
Dim MaxBalls
Dim TroughCount
Dim TroughBall(7)
Dim TroughEject
Dim Momentum
Dim UpperGIon
Dim Multiball
Dim BallsInPlay
Dim iBall
Dim fgBall

Dim dtBank
Dim bsRadarEject


 Sub InitVPM()
     With Controller
         .GameName = cGameName
         .SplashInfoLine = "King Kong" & vbNewLine & "VPX - cyberpez"
         .HandleMechanics = 0
         .HandleKeyboard = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .ShowTitle = 0
'    .hidden = 0
''     If DesktopMode = true then .hidden = true Else .hidden = false End If
'         If Err Then MsgBox Err.Description
     End With
     On Error Goto 0
     Controller.SolMask(0) = 0
     vpmTimer.AddTimer 4000, "Controller.SolMask(0)=&Hffffffff'"
     Controller.Run
End Sub


Sub KK_Init
  ' table initialization

  vpmInit me
  InitVPM



  ' basic pinmame timer
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled  = True

  ' nudging
  vpmNudge.TiltSwitch   = 1
  vpmNudge.Sensitivity  = 3
' vpmNudge.TiltObj    = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)


    ' Drop Targets
    Set DTBank = New cvpmDropTarget
      With DTBank
      .InitDrop Array(Array(sw41vuk),Array(sw42vuk),Array(sw43vuk)), Array(41,42,43)
    .InitSnd SoundFX("_droptarget",DOFDropTargets),SoundFX("resetdrop",DOFContactors)
       End With
'
  Backdrop_Init
  If PrototypeVersion = 13 then MangaProtoType = 1 End If
  SetPrototype

' ball through system
  MaxBalls=3
  InitTime=61
  EjectTime=0
  TroughEject=1
  TroughCount=0
  iBall = 3
  fgBall = false

    CreatBalls

  setup_backglass()

End Sub


Sub KK_Exit()
  Controller.Stop
End Sub

Sub KK_Paused
  Controller.Pause = True
End Sub
Sub KK_UnPaused
  Controller.Pause = False
End Sub


' **** Key_Up ***
Sub KK_KeyUp(ByVal keyCode)

  If keycode = LeftFlipperKey Then
    Controller.Switch(15) = False
    lfpress = 0
    lfpress2 = 0
    leftflipper.eostorqueangle = EOSA
    leftflipper.eostorque = EOST
    LeftFlipperKS.eostorqueangle = EOSA2
    LeftFlipperKS.eostorque = EOST2
  End If
  If keycode = RightFlipperKey Then
    Controller.Switch(16) = False
    rfpress = 0
    rfpress2 = 0
    rightflipper.eostorqueangle = EOSA
    rightflipper.eostorque = EOST
    RightFlipperKS.eostorqueangle = EOSA2
    RightFlipperKS.eostorque = EOST2
  End If
  If KeyUpHandler(keyCode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt"plunger",Plunger

    'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If

End Sub

' **** Key_Down ****
Sub KK_KeyDown(ByVal keyCode)
    If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAt"plungerpull",Plunger
  If keycode = LeftFlipperKey Then  Controller.Switch(15) = True:LFPress = 1:LFPress2 = 1 End If
  If keycode = RightFlipperKey Then Controller.Switch(16) = True:RFPress = 1:RFPress2 = 1 End If
    If KeyDownHandler(keyCode) Then Exit Sub

    ' Manual Ball Control
  If keycode = 46 Then          ' C Key
    If EnableBallControl = 1 Then
      EnableBallControl = 0
    Else
      EnableBallControl = 1
    End If
  End If
    If EnableBallControl = 1 Then
    If keycode = 48 Then        ' B Key
      If BCboost = 1 Then
        BCboost = BCboostmulti
      Else
        BCboost = 1
      End If
    End If
    If keycode = 203 Then BCleft = 1  ' Left Arrow
    If keycode = 200 Then BCup = 1    ' Up Arrow
    If keycode = 208 Then BCdown = 1  ' Down Arrow
    If keycode = 205 Then BCright = 1 ' Right Arrow
  End If

  If keycode = LeftMagnaSave then

    MangaProtoType = MangaProtoType - 1
    If MangaProtoType = 0 then MangaProtoType = 11 End If
    SetPrototype

  End If

  If keycode = RightMagnaSave then
    MangaProtoType = MangaProtoType + 1
    If MangaProtoType = 12 then MangaProtoType = 1 End If
    SetPrototype

  End If


End Sub

Sub TimerPlunger_Timer()
 'debug.print plunger.position
  vr_PlungerRod1.Y = 1105.09 + (5* Plunger.Position) -20
  vr_PlungerRod2.Y = 1105.09 + (5* Plunger.Position) -20
End Sub

'******************************
'  Setup Desktop
'******************************

Sub Backdrop_Init

  If DesktopMode = True then

    If ShowFSS = true then
      lbg53.visible = false
      lbg52.visible = false
      lbg51.visible = false
      lbg50.visible = false
      lbg49.visible = false
    Else

      If BackGlassVisible = 1 then
        lbg53.visible = false
        lbg52.visible = false
        lbg51.visible = false
        lbg50.visible = false
        lbg49.visible = false
      Else
        lbg53.visible = true
        lbg52.visible = true
        lbg51.visible = true
        lbg50.visible = true
        lbg49.visible = true
      End If
    End If

    l57.BulbHaloHeight = 220
    l58.BulbHaloHeight = 220
    l59.BulbHaloHeight = 220
    l60.BulbHaloHeight = 220
    l61.BulbHaloHeight = 220

  Else
    lbg53.visible = false
    lbg52.visible = false
    lbg51.visible = false
    lbg50.visible = false
    lbg49.visible = false

    l57.BulbHaloHeight = 190
    l58.BulbHaloHeight = 190
    l59.BulbHaloHeight = 190
    l60.BulbHaloHeight = 190
    l61.BulbHaloHeight = 190

  End If
End Sub


Dim xxRubbers
Dim PrototypeVersionType, VUKModType, FlasherModType, FlasherMod2Type, MysteryTargetType, MainFlipperModType, SolidStateStickerType, MainFlipperKingStickerType, FlipperBatModType, KingFlipperBatModType, KingFlipperStickerType, LeftWireRampColorType, RightWireRampColorType, ApronTypeType, InstructionCardsType, RubberColorType, RubberPostColorType, PegColorType, BlackRampDecalType, LeftPlasticType, CenterPlasticType, VUKPlasticsType, LeftTargetsType, DTModType, KSRPlasticsType, MangaProtoType, WackerDrivePlasticType, DropTargetDecalType, SlingKongVersionType, KingSizedDELogoType, DateEastLogoColorMatchType, UPFGateType, PlasticRampWedgeType, BlackRampWedge1type, BlackRampWedge2type, DoubleRubbersType, DPPType, KSRSwitchPlasticType, RightWireRampGuideType, LeftSBLPPlasticType, LeftSBLPType, LeftWireRampSupportConfigType, KSBRAorBType, OptinalRails1Type, OptinalRails2Type, Plastic8HeightType, WackerPostType, BlackRampColorType, InsertDecalsType, BRGateType, BackboardPlasticAorBType

Sub SetPrototype ()

'''''Prototype version (my guesses)
'0=random
'1=
'2=
'3=
'4=
'5=
'6=
'7=
'8=
'9=
'10= CP's Production
'11=custom

  If PrototypeVersion = 0 Then
    PrototypeVersionType = Int(Rnd*11)
  Else
    If PrototypeVersion = 13 then
      PrototypeVersionType = MangaProtoType
    Else
      PrototypeVersionType = PrototypeVersion
    End If
  End If

  If PrototypeVersionType = 1 then
    DTModType = 1
    VUKModType = 1
    FlasherModType = 2
    FlasherMod2Type = 2
    MysteryTargetType = 0
    MainFlipperModType = 1
    SolidStateStickerType = 1
    MainFlipperKingStickerType = 0
    FlipperBatModType = 1
    KingFlipperBatModType = 1
    KingFlipperStickerType = 0
    LeftWireRampColorType = 3
    RightWireRampColorType = 1
    ApronTypeType = 2
    InstructionCardsType = 2
    RubberColorType = 2
    RubberPostColorType = 1
    PegColorType = 1
    BlackRampDecalType = 1
    LeftPlasticType = 1
    CenterPlasticType = 1
    VUKPlasticsType = 1
    LeftTargetsType = 1
    KSRPlasticsType = 1
    WackerDrivePlasticType = 0
    SlingKongVersionType = 1
    KingSizedDELogoType = 0
    DateEastLogoColorMatchType = 1
    UPFGateType = 0
    BRGateType = 0
    PlasticRampWedgeType = 1
    BlackRampWedge1type = 0
    BlackRampWedge2type = 0
    DoubleRubbersType = 1
    KSRSwitchPlasticType = 0
    RightWireRampGuideType = 1
    LeftSBLPPlasticType = 1
    LeftSBLPType = 1
    LeftWireRampSupportConfigType = 1
    OptinalRails1Type = 2
    OptinalRails2Type = 2
    Plastic8HeightType = 1
    DropTargetDecalType = 1
    WackerPostType = 0
    InsertDecalsType = 1
    BackboardPlasticAorBType = 1
    If BlackRampColorOverRide = 1 Then
      BlackRampColorType = BlackRampColor
    Else
      BlackRampColorType = 5
    End If
  End If

  If PrototypeVersionType = 2 then
    DTModType = 1
    VUKModType = 1
    FlasherModType = 2
    FlasherMod2Type = 1
    MysteryTargetType = 0
    MainFlipperModType = 1
    SolidStateStickerType = 1
    MainFlipperKingStickerType = 0
    FlipperBatModType = 1
    KingFlipperBatModType = 2
    KingFlipperStickerType = 0
    LeftWireRampColorType = 1
    RightWireRampColorType = 1
    ApronTypeType = 1
    InstructionCardsType = 2
    RubberColorType = 2
    RubberPostColorType = 1
    PegColorType = 2
    BlackRampDecalType = 1
    LeftPlasticType = 1
    CenterPlasticType = 2
    VUKPlasticsType = 2
    LeftTargetsType = 1
    KSRPlasticsType = 1
    WackerDrivePlasticType = 0
    SlingKongVersionType = 1
    KingSizedDELogoType = 3
    DateEastLogoColorMatchType = 1
    UPFGateType = 0
    BRGateType = 0
    PlasticRampWedgeType = 1
    BlackRampWedge1type = 1
    BlackRampWedge2type = 1
    DoubleRubbersType = 1
    KSRSwitchPlasticType = 0
    RightWireRampGuideType = 1
    LeftSBLPPlasticType = 2
    LeftSBLPType = 1
    LeftWireRampSupportConfigType = 1
    OptinalRails1Type = 2
    OptinalRails2Type = 2
    Plastic8HeightType = 1
    DropTargetDecalType = 3
    WackerPostType = 0
    InsertDecalsType = 1
    BackboardPlasticAorBType = 1
    If BlackRampColorOverRide = 1 Then
      BlackRampColorType = BlackRampColor
    Else
      BlackRampColorType = 5
    End If
  End If

  If PrototypeVersionType = 3 then
    DTModType = 1
    VUKModType = 1
    FlasherModType = 2
    FlasherMod2Type = 1
    MysteryTargetType = 0
    MainFlipperModType = 1
    SolidStateStickerType = 1
    MainFlipperKingStickerType = 0
    FlipperBatModType = 1
    KingFlipperBatModType = 2
    KingFlipperStickerType = 1
    LeftWireRampColorType = 2
    RightWireRampColorType = 2
    ApronTypeType = 2
    InstructionCardsType = 2
    RubberColorType = 2
    RubberPostColorType = 1
    PegColorType = 1
    BlackRampDecalType = 2
    LeftPlasticType = 2
    CenterPlasticType = 2
    VUKPlasticsType = 3
    LeftTargetsType = 1
    KSRPlasticsType = 1
    WackerDrivePlasticType = 2
    SlingKongVersionType = 1
    KingSizedDELogoType = 1
    DateEastLogoColorMatchType = 1
    UPFGateType = 0
    BRGateType = 0
    PlasticRampWedgeType = 2
    BlackRampWedge1type = 2
    BlackRampWedge2type = 2
    DoubleRubbersType = 1
    KSRSwitchPlasticType = 0
    RightWireRampGuideType = 1
    LeftSBLPPlasticType = 2
    LeftSBLPType = 0
    LeftWireRampSupportConfigType = 1
    OptinalRails1Type = 2
    OptinalRails2Type = 2
    Plastic8HeightType = 1
    DropTargetDecalType = 3
    WackerPostType = 1
    InsertDecalsType = 1
    BackboardPlasticAorBType = 1
    If BlackRampColorOverRide = 1 Then
      BlackRampColorType = BlackRampColor
    Else
      BlackRampColorType = 5
    End If
  End If

  If PrototypeVersionType = 4 then
    DTModType = 1
    VUKModType = 1
    FlasherModType = 2
    FlasherMod2Type = 1
    MysteryTargetType = 0
    MainFlipperModType = 1
    SolidStateStickerType = 1
    MainFlipperKingStickerType = 0
    FlipperBatModType = 1
    KingFlipperBatModType = 2
    KingFlipperStickerType = 1
    LeftWireRampColorType = 2
    RightWireRampColorType = 2
    ApronTypeType = 2
    InstructionCardsType = 2
    RubberColorType = 2
    RubberPostColorType = 1
    PegColorType = 1
    BlackRampDecalType = 2
    LeftPlasticType = 2
    CenterPlasticType = 2
    VUKPlasticsType = 4
    LeftTargetsType = 1
    KSRPlasticsType = 1
    WackerDrivePlasticType = 2
    SlingKongVersionType = 1
    KingSizedDELogoType = 1
    DateEastLogoColorMatchType = 1
    UPFGateType = 1
    BRGateType = 0
    PlasticRampWedgeType = 3
    BlackRampWedge1type = 3
    BlackRampWedge2type = 3
    DoubleRubbersType = 1
    KSRSwitchPlasticType = 0
    RightWireRampGuideType = 1
    LeftSBLPPlasticType = 2
    LeftSBLPType = 2
    LeftWireRampSupportConfigType = 1
    OptinalRails1Type = 2
    OptinalRails2Type = 2
    Plastic8HeightType = 1
    DropTargetDecalType = 2
    WackerPostType = 1
    InsertDecalsType = 1
    BackboardPlasticAorBType = 1
    If BlackRampColorOverRide = 1 Then
      BlackRampColorType = BlackRampColor
    Else
      BlackRampColorType = 1
    End If
  End If

  If PrototypeVersionType = 5 then
    KSBRAorBType = 1
    DTModType = 1
    VUKModType = 0
    FlasherModType = 1
    FlasherMod2Type = 2
    MysteryTargetType = 1
    MainFlipperModType = 0
    SolidStateStickerType = 1
    MainFlipperKingStickerType = 0
    FlipperBatModType = 1
    KingFlipperBatModType = 2
    KingFlipperStickerType = 1
    LeftWireRampColorType = 2
    RightWireRampColorType = 2
    ApronTypeType = 2
    InstructionCardsType = 2
    RubberColorType = 2
    RubberPostColorType = 1
    PegColorType = 1
    BlackRampDecalType = 1
    LeftPlasticType = 2
    CenterPlasticType = 2
    VUKPlasticsType = 4
    LeftTargetsType = 1
    KSRPlasticsType = 1
    WackerDrivePlasticType = 1
    SlingKongVersionType = 2
    KingSizedDELogoType = 1
    DateEastLogoColorMatchType = 1
    UPFGateType = 1
    BRGateType = 0
    PlasticRampWedgeType = 0
    BlackRampWedge1type = 1
    BlackRampWedge2type = 1
    DoubleRubbersType = 1
    KSRSwitchPlasticType = 1
    RightWireRampGuideType = 1
    LeftSBLPPlasticType = 2
    LeftSBLPType = 2
    LeftWireRampSupportConfigType = 2
    OptinalRails1Type = 2
    OptinalRails2Type = 2
    Plastic8HeightType = 1
    DropTargetDecalType = 2
    WackerPostType = 2
    InsertDecalsType = 1
    BackboardPlasticAorBType = 1
    If BlackRampColorOverRide = 1 Then
      BlackRampColorType = BlackRampColor
    Else
      BlackRampColorType = 1
    End If
  End If

  If PrototypeVersionType = 6 then
    KSBRAorBType = 1
    DTModType = 1
    VUKModType = 0
    FlasherModType = 1
    FlasherMod2Type = 2
    MysteryTargetType = 1
    MainFlipperModType = 0
    SolidStateStickerType = 1
    MainFlipperKingStickerType = 0
    FlipperBatModType = 1
    KingFlipperBatModType = 2
    KingFlipperStickerType = 1
    LeftWireRampColorType = 2
    RightWireRampColorType = 2
    ApronTypeType = 1
    InstructionCardsType = 2
    RubberColorType = 2
    RubberPostColorType = 1
    PegColorType = 1
    BlackRampDecalType = 1
    LeftPlasticType = 2
    CenterPlasticType = 2
    VUKPlasticsType = 1
    LeftTargetsType = 1
    KSRPlasticsType = 1
    WackerDrivePlasticType = 1
    SlingKongVersionType = 2
    KingSizedDELogoType = 1
    DateEastLogoColorMatchType = 1
    UPFGateType = 1
    BRGateType = 0
    PlasticRampWedgeType = 0
    BlackRampWedge1type = 2
    BlackRampWedge2type = 2
    DoubleRubbersType = 0
    KSRSwitchPlasticType = 1
    RightWireRampGuideType = 2
    LeftSBLPPlasticType = 3
    LeftSBLPType = 3
    LeftWireRampSupportConfigType = 2
    OptinalRails1Type = 1
    OptinalRails2Type = 1
    Plastic8HeightType = 2
    DropTargetDecalType = 2
    WackerPostType = 2
    InsertDecalsType = 1
    BackboardPlasticAorBType = 1
    If BlackRampColorOverRide = 1 Then
      BlackRampColorType = BlackRampColor
    Else
      BlackRampColorType = 1
    End If
  End If

  If PrototypeVersionType = 7 then
    KSBRAorBType = 2
    DTModType = 0
    VUKModType = 0
    FlasherModType = 1
    FlasherMod2Type = 2
    MysteryTargetType = 1
    MainFlipperModType = 0
    SolidStateStickerType = 1
    MainFlipperKingStickerType = 0
    FlipperBatModType = 1
    KingFlipperBatModType = 2
    KingFlipperStickerType = 1
    LeftWireRampColorType = 2
    RightWireRampColorType = 2
    ApronTypeType = 1
    InstructionCardsType = 2
    RubberColorType = 2
    RubberPostColorType = 1
    PegColorType = 1
    BlackRampDecalType = 2
    LeftPlasticType = 2
    CenterPlasticType = 2
    VUKPlasticsType = 1
    LeftTargetsType = 1
    KSRPlasticsType = 2
    WackerDrivePlasticType = 1
    SlingKongVersionType = 2
    KingSizedDELogoType = 1
    DateEastLogoColorMatchType = 1
    UPFGateType = 1
    BRGateType = 0
    PlasticRampWedgeType = 0
    BlackRampWedge1type = 1
    BlackRampWedge2type = 2
    DoubleRubbersType = 0
    KSRSwitchPlasticType = 2
    RightWireRampGuideType = 2
    LeftSBLPPlasticType = 4
    LeftSBLPType = 2
    LeftWireRampSupportConfigType = 2
    OptinalRails1Type = 1
    OptinalRails2Type = 1
    Plastic8HeightType = 2
    DropTargetDecalType = 2
    WackerPostType = 2
    InsertDecalsType = 1
    BackboardPlasticAorBType = 1
    If BlackRampColorOverRide = 1 Then
      BlackRampColorType = BlackRampColor
    Else
      BlackRampColorType = 1
    End If
  End If

  If PrototypeVersionType = 8 then
    KSBRAorBType = 2
    DTModType = 0
    VUKModType = 0
    FlasherModType = 1
    FlasherMod2Type = 2
    MysteryTargetType = 1
    MainFlipperModType = 0
    SolidStateStickerType = 1
    MainFlipperKingStickerType = 0
    FlipperBatModType = 4
    KingFlipperBatModType = 2
    KingFlipperStickerType = 1
    LeftWireRampColorType = 2
    RightWireRampColorType = 2
    ApronTypeType = 1
    InstructionCardsType = 2
    RubberColorType = 2
    RubberPostColorType = 1
    PegColorType = 1
    BlackRampDecalType = 2
    LeftPlasticType = 2
    CenterPlasticType = 2
    VUKPlasticsType = 1
    LeftTargetsType = 1
    KSRPlasticsType = 3
    WackerDrivePlasticType = 1
    SlingKongVersionType = 2
    KingSizedDELogoType = 1
    DateEastLogoColorMatchType = 1
    UPFGateType = 1
    BRGateType = 1
    PlasticRampWedgeType = 3
    BlackRampWedge1type = 3
    BlackRampWedge2type = 3
    DoubleRubbersType = 0
    KSRSwitchPlasticType = 2
    RightWireRampGuideType = 2
    LeftSBLPPlasticType = 4
    LeftSBLPType = 2
    LeftWireRampSupportConfigType = 2
    OptinalRails1Type = 1
    OptinalRails2Type = 1
    Plastic8HeightType = 2
    DropTargetDecalType = 2
    WackerPostType = 2
    InsertDecalsType = 1
    BackboardPlasticAorBType = 2
    If BlackRampColorOverRide = 1 Then
      BlackRampColorType = BlackRampColor
    Else
      BlackRampColorType = 1
    End If
  End If

  If PrototypeVersionType = 9 then
    KSBRAorBType = 3
    DTModType = 0
    VUKModType = 0
    FlasherModType = 1
    FlasherMod2Type = 2
    MysteryTargetType = 1
    MainFlipperModType = 0
    SolidStateStickerType = 1
    MainFlipperKingStickerType = 0
    FlipperBatModType = 4
    KingFlipperBatModType = 2
    KingFlipperStickerType = 1
    LeftWireRampColorType = 2
    RightWireRampColorType = 2
    ApronTypeType = 2
    InstructionCardsType = 2
    RubberColorType = 2
    RubberPostColorType = 1
    PegColorType = 1
    BlackRampDecalType = 2
    LeftPlasticType = 2
    CenterPlasticType = 2
    VUKPlasticsType = 1
    LeftTargetsType = 1
    KSRPlasticsType = 5
    WackerDrivePlasticType = 2
    SlingKongVersionType = 2
    KingSizedDELogoType = 1
    DateEastLogoColorMatchType = 1
    UPFGateType = 1
    BRGateType = 1
    PlasticRampWedgeType = 1
    BlackRampWedge1type = 1
    BlackRampWedge2type = 1
    DoubleRubbersType = 0
    KSRSwitchPlasticType = 3
    RightWireRampGuideType = 2
    LeftSBLPPlasticType = 5
    LeftSBLPType = 0
    LeftWireRampSupportConfigType = 2
    OptinalRails1Type = 1
    OptinalRails2Type = 1
    Plastic8HeightType = 2
    DropTargetDecalType = 2
    WackerPostType = 2
    InsertDecalsType = 2
    BackboardPlasticAorBType = 2
    If BlackRampColorOverRide = 1 Then
      BlackRampColorType = BlackRampColor
    Else
      BlackRampColorType = 7
    End If
  End If

  If PrototypeVersionType = 10 then
    DTModType = 1
    VUKModType = 1
    FlasherModType = 2
    FlasherMod2Type = 1
    MysteryTargetType = 0
    MainFlipperModType = 1
    SolidStateStickerType = 1
    MainFlipperKingStickerType = 0
    FlipperBatModType = 4
    KingFlipperBatModType = 2
    KingFlipperStickerType = 1
    LeftWireRampColorType = 2
    RightWireRampColorType = 2
    ApronTypeType = 2
    InstructionCardsType = 2
    RubberColorType = 2
    RubberPostColorType = 1
    PegColorType = 1
    BlackRampDecalType = 2
    KSBRAorBType = 1
    LeftPlasticType = 2
    CenterPlasticType = 2
    VUKPlasticsType = 5
    LeftTargetsType = 1
    KSRPlasticsType = 1
    WackerDrivePlasticType = 1
    SlingKongVersionType = 1
    KingSizedDELogoType = 1
    DateEastLogoColorMatchType = 2
    UPFGateType = 1
    BRGateType = 1
    PlasticRampWedgeType = 0
    BlackRampWedge1type = 3
    BlackRampWedge2type = 3
    DoubleRubbersType = 1
    KSRSwitchPlasticType = 0
    RightWireRampGuideType = 1
    LeftSBLPPlasticType = 1
    LeftSBLPType = 1
    LeftWireRampSupportConfigType = 2
    OptinalRails1Type = 2
    OptinalRails2Type = 2
    Plastic8HeightType = 2
    DropTargetDecalType = 4
    WackerPostType = 2
    InsertDecalsType = 1
    BackboardPlasticAorBType = 2
    If BlackRampColorOverRide = 1 Then
      BlackRampColorType = BlackRampColor
    Else
      BlackRampColorType = 5
    End If
  End If

  If PrototypeVersionType = 11 then
    DTModType = 0
    VUKModType = 0
    FlasherModType = 1
    FlasherMod2Type = 1
    MysteryTargetType = 1
    MainFlipperModType = 1
    SolidStateStickerType = 1
    MainFlipperKingStickerType = 0
    FlipperBatModType = 4
    KingFlipperBatModType = 2
    KingFlipperStickerType = 1
    LeftWireRampColorType = 2
    RightWireRampColorType = 2
    ApronTypeType = 2
    InstructionCardsType = 2
    RubberColorType = 2
    RubberPostColorType = 1
    PegColorType = 1
    BlackRampDecalType = 2
    KSBRAorBType = 1
    LeftPlasticType = 2
    CenterPlasticType = 2
    VUKPlasticsType = 1
    LeftTargetsType = 1
    KSRPlasticsType = 4
    WackerDrivePlasticType = 1
    SlingKongVersionType = 2
    KingSizedDELogoType = 1
    DateEastLogoColorMatchType = 2
    UPFGateType = 1
    BRGateType = 1
    PlasticRampWedgeType = 0
    BlackRampWedge1type = 3
    BlackRampWedge2type = 3
    DoubleRubbersType = 1
    KSRSwitchPlasticType = 1
    RightWireRampGuideType = 2
    LeftSBLPPlasticType = 4
    LeftSBLPType = 2
    LeftWireRampSupportConfigType = 2
    OptinalRails1Type = 1
    OptinalRails2Type = 1
    Plastic8HeightType = 2
    DropTargetDecalType = 4
    WackerPostType = 2
    InsertDecalsType = 1
    BackboardPlasticAorBType = 2
    If BlackRampColorOverRide = 1 Then
      BlackRampColorType = BlackRampColor
    Else
      BlackRampColorType = 6
    End If
  End If

  If PrototypeVersionType = 12 Then

    If FlipperBatMod = 0 Then
      FlipperBatModType = Int(Rnd*4)
    Else
      FlipperBatModType = FlipperBatMod
    End If

    If SolidStateSticker = 0 Then
      SolidStateStickerType = Int(Rnd*2)
    Else
      SolidStateStickerType = SolidStateSticker
    End If


    DTModType = DTMod
    VUKModType = VUKMod
    FlasherModType = FlasherMod
    FlasherMod2Type = FlasherMod2
    MysteryTargetType = MysteryTarget
    MainFlipperModType = MainFlipperMod
    MainFlipperKingStickerType = MainFlipperKingSticker
    KingFlipperBatModType = KingFlipperBatMod
    KingFlipperStickerType = KingFlipperSticker
    LeftWireRampColorType = LeftWireRampColor
    RightWireRampColorType = RightWireRampColor
    ApronTypeType = ApronType
    InstructionCardsType = InstructionCards
    RubberColorType = RubberColor
    RubberPostColorType = RubberPostColor
    PegColorType = PegColor
    BlackRampDecalType = BlackRampDecal
    LeftPlasticType = LeftPlastic
    CenterPlasticType = CenterPlastic
    VUKPlasticsType = VUKPlastics
    LeftTargetsType = LeftTargets
    KSRPlasticsType = KSRPlastics
    WackerDrivePlasticType = WackerDrivePlastic
    DropTargetDecalType = DropTargetDecal
    SlingKongVersionType = SlingKongVersion
    KingSizedDELogoType = KingSizedDELogo
    DateEastLogoColorMatchType = DateEastLogoColorMatch
    UPFGateType = UPFGate
    BRGateType = BRGate
    PlasticRampWedgeType = PlasticRampWedge
    BlackRampWedge1type = BlackRampWedge1
    BlackRampWedge2type = BlackRampWedge2
    DoubleRubbersType = DoubleRubbers
    KSRSwitchPlasticType = KSRSwitchPlastic
    RightWireRampGuideType = RightWireRampGuide
    LeftSBLPPlasticType = LeftSBLPPlastic
    LeftSBLPType = LeftSBLP
    LeftWireRampSupportConfigType = LeftWireRampSupportConfig
    KSBRAorBType = KSBRAorB
    OptinalRails1Type = OptinalRails1
    OptinalRails2Type = OptinalRails2
    Plastic8HeightType = Plastic8Height
    WackerPostType = WackerPost
    BlackRampColorType = BlackRampColor
    InsertDecalsType = InsertDecals
    BackboardPlasticAorBType = BackboardPlasticAorB
  Else
    DTModType = DTModType
    VUKModType = VUKModType
    FlasherModType = FlasherModType
    FlasherMod2Type = FlasherMod2Type
    MysteryTargetType = MysteryTargetType
    MainFlipperModType = MainFlipperModType
    SolidStateStickerType = SolidStateStickerType
    MainFlipperKingStickerType = MainFlipperKingStickerType
    If FlipperColorOverride = 1 Then
      FlipperBatModType = FlipperBatMod
      KingFlipperBatModType = KingFlipperBatMod
    Else
      FlipperBatModType = FlipperBatModType
      KingFlipperBatModType = KingFlipperBatModType
    End If
    KingFlipperStickerType = KingFlipperStickerType
    LeftWireRampColorType = LeftWireRampColorType
    RightWireRampColorType = RightWireRampColorType
    ApronTypeType = ApronTypeType
    InstructionCardsType = InstructionCardsType
    RubberColorType = RubberColorType
    RubberPostColorType = RubberPostColorType
    PegColorType = PegColorType
    BlackRampDecalType = BlackRampDecalType
    LeftPlasticType = LeftPlasticType
    CenterPlasticType = CenterPlasticType
    VUKPlasticsType = VUKPlasticsType
    LeftTargetsType = LeftTargetsType
    KSRPlasticsType = KSRPlasticsType
    WackerDrivePlasticType = WackerDrivePlasticType
    DropTargetDecalType = DropTargetDecalType
    SlingKongVersionType = SlingKongVersionType
    KingSizedDELogoType = KingSizedDELogoType
    DateEastLogoColorMatchType = DateEastLogoColorMatchType
    UPFGateType = UPFGateType
    BRGateType = BRGateType
    PlasticRampWedgeType = PlasticRampWedgeType
    BlackRampWedge1type = BlackRampWedge1type
    BlackRampWedge2type = BlackRampWedge2type
    DoubleRubbersType = DoubleRubbersType
    KSRSwitchPlasticType = KSRSwitchPlasticType
    RightWireRampGuideType = RightWireRampGuideType
    LeftSBLPPlasticType = LeftSBLPPlasticType
    LeftSBLPType = LeftSBLPType
    LeftWireRampSupportConfigType = LeftWireRampSupportConfigType
    KSBRAorBType = KSBRAorBType
    OptinalRails1Type = OptinalRails1Type
    OptinalRails2Type = OptinalRails2Type
    Plastic8HeightType = Plastic8HeightType
    WackerPostType = WackerPostType
    BlackRampColorType = BlackRampColorType
    InsertDecalsType = InsertDecalsType
    BackboardPlasticAorBType = BackboardPlasticAorBType
  End If

  SetOptions

End Sub

Sub SetOptions ()
  If FlipperBatModType = 1 Then
    batright.image = "Flipperbat WhiteBlack"
    batleft.image = "Flipperbat WhiteBlack"

    batleftKS2.image = "Flipperbat WhiteBlack"
    batrightKS.image = "Flipperbat WhiteBlack"
    batLeftKS.image = "Flipperbat WhiteBlack"

    If DateEastLogoColorMatch = 1 Then
      pLeftFlipperLogo.image = "Flipperbat WhiteBlack"
      pRightFlipperLogo.image = "Flipperbat WhiteBlack"
    Else
      pLeftFlipperLogo.image = "flipperbat BlackWhite"
      pRightFlipperLogo.image = "flipperbat BlackWhite"
    End If
  End If

  If FlipperBatModType = 2 Then
    batright.image = "Flipperbat OrangeYellow"
    batleft.image = "Flipperbat OrangeYellow"
    batleftKS2.image = "Flipperbat OrangeYellow"
    batrightKS.image = "Flipperbat OrangeYellow"
    batLeftKS.image = "Flipperbat OrangeYellow"
    If DateEastLogoColorMatch = 1 Then
      pLeftFlipperLogo.image = "Flipperbat OrangeYellow"
      pRightFlipperLogo.image = "Flipperbat OrangeYellow"
    Else
      pLeftFlipperLogo.image = "flipperbat YellowRed"
      pRightFlipperLogo.image = "flipperbat YellowRed"
    End If
  End If

  If FlipperBatModType = 3 Then
    batright.image = "Flipperbat DarkOrangeBlack"
    batleft.image = "Flipperbat DarkOrangeBlack"
    batleftKS2.image = "Flipperbat DarkOrangeBlack"
    batrightKS.image = "Flipperbat DarkOrangeBlack"
    batLeftKS.image = "Flipperbat DarkOrangeBlack"
    If DateEastLogoColorMatch = 1 Then
      pLeftFlipperLogo.image = "Flipperbat DarkOrangeBlack"
      pRightFlipperLogo.image = "Flipperbat DarkOrangeBlack"
    Else
      pLeftFlipperLogo.image = "flipperbat BlackDarkOrange"
      pRightFlipperLogo.image = "flipperbat BlackDarkOrange"
    End If

  End If

  If FlipperBatModType = 4 Then
    batright.image = "flipperbat YellowRed"
    batleft.image = "flipperbat YellowRed"
    batleftKS2.image = "flipperbat YellowRed"
    batrightKS.image = "flipperbat YellowRed"
    batLeftKS.image = "flipperbat YellowRed"
    If DateEastLogoColorMatch = 1 Then
      pLeftFlipperLogo.image = "flipperbat YellowRed"
      pRightFlipperLogo.image = "flipperbat YellowRed"
    Else
      pLeftFlipperLogo.image = "flipperbat RedYellow"
      pRightFlipperLogo.image = "flipperbat RedYellow"
    End If
  End If


  If SolidStateStickerType = 1 Then
    pLSS.visible = false
    pRSS.visible = false
  End If

  If SolidStateStickerType = 2 Then
    pLSS.visible = True
    pRSS.visible = True
  End If

  If KingFlipperBatModType = 1 Then
    batright1.image = "Flipperbat WhiteBlack"
    If DateEastLogoColorMatch = 1 Then
      pRightFlipper2Logo.image = "Flipperbat WhiteBlack"
    Else
      pRightFlipper2Logo.image = "flipperbat BlackWhite"
    End If
  End If

  If KingFlipperBatModType = 2 Then
    batright1.image = "Flipperbat OrangeYellow"
    If DateEastLogoColorMatch = 1 Then
      pRightFlipper2Logo.image = "Flipperbat OrangeYellow"
    Else
      pRightFlipper2Logo.image = "flipperbat YellowRed"
    End If
  End If

  If KingFlipperBatModType = 3 Then
    batright1.image = "Flipperbat DarkOrangeBlack"
    If DateEastLogoColorMatch = 1 Then
      pRightFlipper2Logo.image = "Flipperbat DarkOrangeBlack"
    Else
      pRightFlipper2Logo.image = "flipperbat BlackDarkOrange"
    End If
  End If

  If KingFlipperStickerType= 0 Then
    batright2.visible = false
    If KingSizedDELogoType = 0 Then
      pRightFlipper2Logo.visible = False
      batright1.visible = True
    End If
    If KingSizedDELogoType = 1 Then
      pRightFlipper2Logo.visible = False
      batright1.visible = True
    End If
    If KingSizedDELogoType = 2 Then
      pRightFlipper2Logo.visible = True
      batright1.visible = False
    End If
    If KingSizedDELogoType = 3 Then
      pRightFlipper2Logo.visible = True
      batright1.visible = False
    End If
  Else
    pRightFlipper2Logo.visible = False
    batright2.visible = True
    batright1.visible = True
  End If

  If ApronHide = 1 Then
    pApron.Visible = False
    pApron2.Visible = False
  Else
    If ApronTypeType = 1 Then
      pApron.Visible = False
      pApron2.Visible = True
    End If

    If ApronTypeType = 2 Then
      pApron.Visible = True
      pApron2.Visible = False
    End If

    If ApronTypeType = 3 Then
      pApron.Visible = False
      pApron2.Visible = False
    End If
  End If


  If ApronHide = 1 Then
    IC_Right.Visible = False
    IC_Left.Visible = False
  Else
    If InstructionCardsType = 1 Then
      IC_Right.image = "DE_IC_3ball"
      IC_Left.image = "KK_IC_Left"
    End If

    If InstructionCardsType = 2 Then
      IC_Right.image = "KK_IC_Right2"
      IC_Left.image = "KK_IC_Left2"
    End If

    If InstructionCardsType = 3 Then
      IC_Right.Visible = False
      IC_Left.Visible = False
    End If
  End If

  If RubberColorType = 1 Then
    for each xxRubbers in allRubbers
      xxRubbers.material = "Rubber White"
      next
  End If

  If RubberColorType = 2 Then
    for each xxRubbers in allRubbers
      xxRubbers.material = "Black Rubber"
      next
  End If

  If RubberColorType = 3 Then
    for each xxRubbers in allRubbers
      xxRubbers.material = "Yellow Rubber"
      next
  End If


If DTModType = 0 Then
  sw41p.visible = False
  sw42p.visible = False
  sw43p.visible = False


  SW41.isDropped = False
  SW42.isDropped = False
  SW43.isDropped = False

  SW41vuk.isDropped = true
  SW42vuk.isDropped = true
  SW43vuk.isDropped = true

  SW41vuk001.isDropped = true
  SW42vuk001.isDropped = true
  SW43vuk001.isDropped = true


  psw41.visible = true
  pT41b.visible = true
  pT41c.visible = true
  psw42.visible = true
  pT42b.visible = true
  pT42c.visible = true
  psw43.visible = true
  pT43b.visible = true
  pT43c.visible = true

  PegMetalT4.visible = False
  PegMetalT10.visible = False
  Rubber10.visible = False
  Rubber10b.visible = False

  Wall004.Collidable = False
  Primitive031.Collidable = False
  Primitive030.Collidable = False

  Wall013.Visible = False
  Wall013.SideVisible = false
  Wall012.Visible = False
  Wall012.SideVisible = False

  Wall008.Visible = True
  Wall008.SideVisible = True

  Wall016.IsDropped = True
  Wall16.IsDropped = False


  Primitive011b.collidable = False
  PrimRubberPostT015.Visible = False
  Primitive011.Collidable = True
  PrimRubberPostT4.Visible = True
  nut016.visible = False
  nut011.visible = True
  PegPlasticT1Mini2.Visible = False

  screw026.Visible = False
  screw027.Visible = False
  screw028.Visible = False
  screw029.Visible = False

End If


If DTModType = 1 Then
  psw41.visible = false
  pT41b.visible = false
  pT41c.visible = false
  psw42.visible = false
  pT42b.visible = false
  pT42c.visible = false
  psw43.visible = false
  pT43b.visible = false
  pT43c.visible = false

  SW41.isDropped = true
  SW42.isDropped = true
  SW43.isDropped = true

  SW41vuk.isDropped = false
  SW42vuk.isDropped = false
  SW43vuk.isDropped = false

  SW41vuk001.isDropped = false
  SW42vuk001.isDropped = false
  SW43vuk001.isDropped = false

  sw41p.visible = True
  sw42p.visible = True
  sw43p.visible = True


  Wall004.Collidable = True
  Primitive031.Collidable = True
  Primitive030.Collidable = True

  PegMetalT4.visible = True
  PegMetalT10.visible = True
  Rubber10.visible = True
  Rubber10b.visible = True

  Wall013.Visible = True
  Wall013.SideVisible = True
  Wall012.Visible = True
  Wall012.SideVisible = True

  Wall008.Visible = False
  Wall008.SideVisible = False

  Wall016.IsDropped = false
  Wall16.IsDropped = True

  Primitive011b.collidable = True
  PrimRubberPostT015.Visible = True
  Primitive011.Collidable = False
  PrimRubberPostT4.Visible = False
  nut016.visible = True
  nut011.visible = False
  If CenterPlasticType = 2 then
    PegPlasticT1Mini2.Visible = True
    screw026.Visible = False
    screw027.Visible = False
    screw028.Visible = False
    screw029.Visible = False
  Else
    PegPlasticT1Mini2.Visible = False
    screw026.Visible = True
    screw027.Visible = True
    screw028.Visible = True
    screw029.Visible = True
  End If


End If


If VUKModType = 1 Then

  pBlackRampPrettyKSR.Visible = False
  pBlackRampPrettyKSR2.Visible = False
  pBlackRampColSideTopWallKSR.Collidable = False
  pBlackRampColKSR.Collidable = False
  pBlackRampPrettyKSR3.Visible = False
  pBlackRampColKSR3.Collidable = False
  pBlackRampDecal2KSR3.Visible = False

  pBlackRampPretty.Visible = True
  pBlackRampColSideTopWall.Collidable = True
  pBlackRampCol.Collidable = True

  screw054.Visible = True
  screw055.Visible = True
  screw056.Visible = False
  screw057.Visible = False

  pMetalRampCol.Collidable = False
  pLeftWireRampKSR.Visible = False
  pMetalRamp.Visible = False
  Ramp4.Collidable = False
  Ramp001.Collidable = False
  Ramp003.Collidable = False
  Ramp004.Collidable = False

  Ramp5.Collidable = True
  pColorRamp.Visible = True
  Wall323.IsDropped = false

  RailVertKicker.Visible = true
  WallBracket3.Visible = true
  WallBracket3a.Visible = true
  WallBracket4.Visible = true
  WallBracket4a.Visible = true

  RampKSR.Collidable = False
  RampVUK.Collidable = True

  Rivit1.z = 15
  Rivit2.z = 12
  pRampWedge1.TransY = 0

If LeftWireRampColorType = 1 Then
  pColorRamp.Material = "Chrome"
  pColorRamp.Image = ""
  Primitive44.Material = "Chrome"
  Primitive44.Image = ""
End If

If LeftWireRampColorType = 2 Then
  pColorRamp.Material = "Yellow Wire Ramp"
  pColorRamp.Image = ""
  Primitive44.Material = "Yellow Wire Ramp"
  Primitive44.Image = ""
End If

If LeftWireRampColorType = 3 Then
  pColorRamp.Material = "Plastic with an image"
  pColorRamp.Image = "ColorRamp_texture7"
  Primitive44.Material = "Plastic with an image"
  Primitive44.Image = "ColorRampOrange"
End If

  pWiresB.Visible = False
  pWiresZipTiesB.Visible = False
  KSRswitchArm.Visible = False
  KSRswitch.Visible = False
Else
  If KSBRAorBType = 1 then
    pBlackRampPrettyKSR.Visible = True
    pBlackRampPrettyKSR2.Visible = False
    pBlackRampPrettyKSR3.Visible = False
    pBlackRampColKSR.Collidable = True
    pBlackRampColKSR3.Collidable = False

    pBlackRampDecal2KSR3.Visible = False

  screw056.Visible = True
  screw057.Visible = True

  Rivit1.z = 21
  Rivit2.z = 15
  pRampWedge1.TransY = 3.6

  End If
  If KSBRAorBType = 2 then
    pBlackRampPrettyKSR.Visible = False
    pBlackRampPrettyKSR2.Visible = True
    pBlackRampPrettyKSR3.Visible = False
    pBlackRampColKSR.Collidable = True
    pBlackRampColKSR3.Collidable = False

    pBlackRampDecal2KSR3.Visible = False

  screw056.Visible = True
  screw057.Visible = True

  Rivit1.z = 21
  Rivit2.z = 15
  pRampWedge1.TransY = 3.6

  End If
  If KSBRAorBType = 3 then
    pBlackRampPrettyKSR.Visible = False
    pBlackRampPrettyKSR2.Visible = False
    pBlackRampPrettyKSR3.Visible = True
    pBlackRampColKSR.Collidable = False
    pBlackRampColKSR3.Collidable = True

    pBlackRampDecal2KSR3.Visible = True

  screw056.Visible = false
  screw057.Visible = false

  Rivit1.z = 20
  Rivit2.z = 13
  pRampWedge1.TransY = 2

  End If

  screw054.Visible = False
  screw055.Visible = False

  pBlackRampColSideTopWallKSR.Collidable = True

  pBlackRampPretty.Visible = False
  pBlackRampColSideTopWall.Collidable = False
  pBlackRampCol.Collidable = False

  pMetalRampCol.Collidable = True
  pLeftWireRampKSR.Visible = True
  pMetalRamp.Visible = True
  Ramp4.Collidable = True
  Ramp001.Collidable = True
  Ramp003.Collidable = True
  Ramp004.Collidable = True
  Primitive44.Material = "Chrome"
  Primitive44.Image = ""

  Ramp5.Collidable = False
  pColorRamp.Visible = False
  Wall323.IsDropped = True

  RailVertKicker.Visible = False
  WallBracket3.Visible = False
  WallBracket3a.Visible = False
  WallBracket4.Visible = False
  WallBracket4a.Visible = False

  RampKSR.Collidable = True
  RampVUK.Collidable = False

  pWiresB.Visible = True
  pWiresZipTiesB.Visible = True
  KSRswitchArm.Visible = True
  KSRswitch.Visible = True

End If



Dim RocketLaunchTexture

RocketLaunchTexture = Int(Rnd*3) + 1

If RocketLaunchTexture = 1 Then
  pRocketLaunch.image = "MissleLauncher2"
End If

If RocketLaunchTexture = 2 Then
  pRocketLaunch.image = "MissleLauncher"
End If

If RocketLaunchTexture = 3 Then
  pRocketLaunch.image = "MissleLauncher2"
End If



  If BlackRampColorType = 0 Then
    BlackRampColorType = Int(Rnd*7) + 1
  Else
    BlackRampColorType = BlackRampColorType
  End If

  If BlackRampColorType = 2 Then
    pBlackRampPrettyKSR.image = "ramptex1"
    pBlackRampPrettyKSR.material = "ramptex1"
    pBlackRampPrettyKSR2.image = "ramptex1"
    pBlackRampPrettyKSR2.material = "ramptex1"
    pBlackRampPrettyKSR3.image = "ramptex1"
    pBlackRampPrettyKSR3.material = "ramptex1"
    pBlackRampPretty.image = "ramptex1"
    pBlackRampPretty.material = "ramptex1"
  End If

  If BlackRampColorType = 1 Then
    pBlackRampPrettyKSR.image = "ramptex2"
    pBlackRampPrettyKSR.material = "ramptex22"
    pBlackRampPrettyKSR2.image = "ramptex2"
    pBlackRampPrettyKSR2.material = "ramptex22"
    pBlackRampPrettyKSR3.image = "ramptex2"
    pBlackRampPrettyKSR3.material = "ramptex22"
    pBlackRampPretty.image = "ramptex2"
    pBlackRampPretty.material = "ramptex22"
  End If

  If BlackRampColorType = 3 Then
    pBlackRampPrettyKSR.image = "ramptex3"
    pBlackRampPrettyKSR.material = "ramptex3"
    pBlackRampPrettyKSR2.image = "ramptex3"
    pBlackRampPrettyKSR2.material = "ramptex3"
    pBlackRampPrettyKSR3.image = "ramptex3"
    pBlackRampPrettyKSR3.material = "ramptex3"
    pBlackRampPretty.image = "ramptex3"
    pBlackRampPretty.material = "ramptex3"
  End If

  If BlackRampColorType = 4 Then
    pBlackRampPrettyKSR.image = "ramptex4"
    pBlackRampPrettyKSR.material = "ramptex3"
    pBlackRampPrettyKSR2.image = "ramptex4"
    pBlackRampPrettyKSR2.material = "ramptex4"
    pBlackRampPrettyKSR3.image = "ramptex4"
    pBlackRampPrettyKSR3.material = "ramptex4"
    pBlackRampPretty.image = "ramptex4"
    pBlackRampPretty.material = "ramptex4"
  End If

  If BlackRampColorType = 5 Then
    pBlackRampPrettyKSR.image = "ramptex5"
    pBlackRampPrettyKSR.material = "ramptex3"
    pBlackRampPrettyKSR2.image = "ramptex5"
    pBlackRampPrettyKSR2.material = "ramptex3"
    pBlackRampPrettyKSR3.image = "ramptex5"
    pBlackRampPrettyKSR3.material = "ramptex3"
    pBlackRampPretty.image = "ramptex5"
    pBlackRampPretty.material = "ramptex3"
  End If

  If BlackRampColorType = 6 Then
    pBlackRampPrettyKSR.image = "ramptex6"
    pBlackRampPrettyKSR.material = "ramptex3"
    pBlackRampPrettyKSR2.image = "ramptex6"
    pBlackRampPrettyKSR2.material = "ramptex3"
    pBlackRampPrettyKSR3.image = "ramptex6"
    pBlackRampPrettyKSR3.material = "ramptex3"
    pBlackRampPretty.image = "ramptex6"
    pBlackRampPretty.material = "ramptex3"
  End If

  If BlackRampColorType = 7 Then
    pBlackRampPrettyKSR.image = "ramptex7"
    pBlackRampPrettyKSR.material = "ramptex3"
    pBlackRampPrettyKSR2.image = "ramptex7"
    pBlackRampPrettyKSR2.material = "ramptex3"
    pBlackRampPrettyKSR3.image = "ramptex7"
    pBlackRampPrettyKSR3.material = "ramptex3"
    pBlackRampPretty.image = "ramptex7"
    pBlackRampPretty.material = "ramptex3"
  End If

If MainFlipperModType = 1 Then
  If MainFlipperKingStickerType = 1 Then
    batrightKS2.visible = True
    batleftKS2.visible = True
  Else
    batrightKS2.visible = False
    batleftKS2.visible = False

    pLeftFlipperLogo.visible = False
    pRightFlipperLogo.visible = False

    If KingSizedDELogoType = 0 Then
      batleftKS.visible = true
      batRightKS.visible = true
      pLeftFlipperLogo.visible = false
      pRightFlipperLogo.visible = false
    End If

    If KingSizedDELogoType = 1 Then
      batleftKS.visible = False
      batRightKS.visible = False
      pLeftFlipperLogo.visible = true
      pRightFlipperLogo.visible = true
    End If

    If KingSizedDELogoType = 2 Then
      batleftKS.visible = true
      batRightKS.visible = true
      pLeftFlipperLogo.visible = false
      pRightFlipperLogo.visible = false
    End If

    If KingSizedDELogoType = 3 Then
      batleftKS.visible = False
      batRightKS.visible = False
      pLeftFlipperLogo.visible = True
      pRightFlipperLogo.visible = True
    End If

  End If

  pRSS.visible = false
  batright.visible = false

  pLSS.visible = false
  batleft.visible = false

  RightFlipper.Enabled = false
  LeftFlipper.Enabled = false

  RightFlipperKS.Enabled = True
  LeftFlipperKS.Enabled = True

  pLanePlasticsBig.visible = false
  pLaneGuidesBig.visible = false
  pLanePlasticsSmall.visible = True
  pLaneGuidesSmall.visible = True

  Primitive13.Visible = false
  Primitive5.Visible = True

  Primitive14.Visible = false
  Primitive1.Visible = True

  TriggerLF.enabled = False
  TriggerRF.enabled = False
  TriggerLKSF.enabled = True
  TriggerRKSF.enabled = True

  batrightshadow.Visible = False
  batleftshadow.Visible = False
  batrightshadowKS.Visible = True
  batleftshadowKS.Visible = True

Else

  pRSS.visible = True
  batright.visible = True

  pLSS.visible = True
  batleft.visible = True

  batleftKS2.visible = false
  batleftKS.visible = false

  batrightKS.visible = false
  batrightKS2.visible = false

  pLeftFlipperLogo.visible = false
  pRightFlipperLogo.visible = false

  pLeftFlipperLogo.visible = false
  pRightFlipperLogo.visible = false

  RightFlipper.Enabled = True
  LeftFlipper.Enabled = True

  RightFlipperKS.Enabled = False
  LeftFlipperKS.Enabled = False

  pLanePlasticsBig.visible = True
  pLaneGuidesBig.visible = True
  pLanePlasticsSmall.visible = False
  pLaneGuidesSmall.visible = False

  Primitive13.Visible = True
  Primitive5.Visible = False

  Primitive14.Visible = True
  Primitive1.Visible = False

  TriggerLF.enabled = True
  TriggerRF.enabled = True
  TriggerLKSF.enabled = False
  TriggerRKSF.enabled = False

  batrightshadow.Visible = True
  batleftshadow.Visible = True
  batrightshadowKS.Visible = False
  batleftshadowKS.Visible = False


End If

If FlasherModType = 1 Then
  pBulbFlasher1.visible = True
  pBulbFilament1.visible = True
  pBulbBase1.visible = True
  f112X.visible = True
  pWiresC.visible = True
  Primitive7.visible = false
  Cap003.visible = false
  Cap004.visible = false
  pBulbFlasher003.visible = false
  pBulbBase003.visible = false
  pBulbFilament003.visible = false
  Primitive2.visible = True
  Cap005.visible = True
  Cap006.visible = True

End If


If FlasherModType = 2 Then
  pBulbFlasher1.visible = false
  pBulbFilament1.visible = false
  pBulbBase1.visible = false
  f112X.visible = false
  pWiresC.visible = False
  Primitive2.visible = false
  Cap005.visible = False
  Cap006.visible = False
  Primitive7.visible = true
  Cap003.visible = True
  Cap004.visible = True
  pBulbFlasher003.visible = True
  pBulbBase003.visible = True
  pBulbFilament003.visible = True

End If


If FlasherMod2Type = 1 Then
  f104a.Color=RGB(255,255,0)
  f104a.ColorFull=RGB(255,255,255)
  f104b.Color=RGB(255,128,0)
  f104b.ColorFull=RGB(255,255,255)
  Primitive8.visible = False
  Primitive8b.visible = True
End If

If FlasherMod2Type = 2 Then
  f104a.Color=RGB(255,0,0)
  f104a.ColorFull=RGB(255,72,72)
  f104b.Color=RGB(255,72,72)
  f104b.ColorFull=RGB(255,255,255)
  Primitive8.visible = True
  Primitive8b.visible = False
End If



If MysteryTargetType = 1 Then
  sw44.collidable = True
  sw44.Visible = True
  Wall33.collidable = False
  pThickRubber1a.Visible = False
  pThickRubber1b.Visible = False
  Primitive015.collidable = True
  PegPlasticT6.Visible = true
  Rubber1.Visible = true
  Rubber4.Visible = false
  PegMetalT001.Visible = False
  PrimRubberPostT2.Visible = true
Else
  sw44.collidable = False
  sw44.Visible = false
  Wall33.collidable = False
  pThickRubber1a.Visible = True
  pThickRubber1b.Visible = True
  Primitive015.collidable = False
  PegPlasticT6.Visible = False
  Rubber1.Visible = False
' Rubber4.Visible = True
  PegMetalT001.Visible = true
  PrimRubberPostT2.Visible = False
End if



''''Backboard Plastic A or B

If BackboardPlasticAorBType = 1 Then
  pBackBoard1A.Visible = True
  pBackBoard1B.Visible = False
End If

If BackboardPlasticAorBType = 2 Then
  pBackBoard1A.Visible = False
  pBackBoard1B.Visible = True
End If



'''DropTargetDecal

If DropTargetDecalType = 1 Then
  sw41p.image = ""
  sw42p.image = ""
  sw43p.image = ""
End If

If DropTargetDecalType = 3 Then
  sw41p.image = "drop_texture"
  sw42p.image = "drop_texture"
  sw43p.image = "drop_texture"
End If

If DropTargetDecalType = 2 Then
  sw41p.image = "drop_texture2"
  sw42p.image = "drop_texture2"
  sw43p.image = "drop_texture2"
End If

If DropTargetDecalType = 4 Then
  sw41p.image = "drop_texture3"
  sw42p.image = "drop_texture3"
  sw43p.image = "drop_texture3"
End If

If DropTargetDecalType = 5 Then
  sw41p.image = "drop_texture4"
  sw42p.image = "drop_texture4"
  sw43p.image = "drop_texture4"
End If

If LeftTargetsType = 1 Then
  psw33.Visible = true
  psw34.Visible = true
  psw35.Visible = true
  psw36.Visible = true
  psw37.Visible = true

  psw33t.Visible = False
  psw34t.Visible = False
  psw35t.Visible = False
  psw36t.Visible = False
  psw37t.Visible = False
End If

If LeftTargetsType = 2 Then
  psw33t.Visible = true
  psw34t.Visible = true
  psw35t.Visible = true
  psw36t.Visible = true
  psw37t.Visible = true

  psw33.Visible = False
  psw34.Visible = False
  psw35.Visible = False
  psw36.Visible = False
  psw37.Visible = False
End If

If UPFGateType = 1 then
  pUPFGate.visible = True
  UPFGateBracket.visible = True
  GateUPF.collidable = True
  PrimRubberPostT012.visible = True
  PrimRubberPostT011.visible = True
  screw023.visible = True
  screw024.visible = True
Else
  pUPFGate.visible = False
  UPFGateBracket.visible = False
  GateUPF.collidable = False
  PrimRubberPostT012.visible = False
  PrimRubberPostT011.visible = False
  screw023.visible = False
  screw024.visible = False
End If


If BRGateType = 0 Then
  GateBR.Collidable = False
  pGateBR.visible = False
  pGateBRBracket.visible = False
End If

If BRGateType = 1 Then
  GateBR.Collidable = true
  pGateBR.visible = true
  pGateBRBracket.visible = true
End If


If RubberPostColorType = 1 Then
  PrimRubberPostT10.image = "rubber-post_Black"
  PrimRubberPostT9.image = "rubber-post_Black"
  PrimRubberPostT12.image = "rubber-post_Black"
  PrimRubberPostT6.image = "rubber-post_Black"
  PrimRubberPostT3.image = "rubber-post_Black"
  PrimRubberPostT5.image = "rubber-post_Black"
  PrimRubberPostT17.image = "rubber-post_Black"
  PrimRubberPostT11.image = "rubber-post_Black"
  PrimRubberPostT4.image = "rubber-post_Black"
  PrimRubberPostT015.image = "rubber-post_Black"
  PrimRubberPostT2.image = "rubber-post_Black"
  PrimRubberPostT024.image = "rubber-post_Black"
End If

If RubberPostColorType = 2 Then
  PrimRubberPostT10.image = "rubber-post_Yellow"
  PrimRubberPostT9.image = "rubber-post_Yellow"
  PrimRubberPostT12.image = "rubber-post_Yellow"
  PrimRubberPostT6.image = "rubber-post_Yellow"
  PrimRubberPostT3.image = "rubber-post_Yellow"
  PrimRubberPostT5.image = "rubber-post_Yellow"
  PrimRubberPostT17.image = "rubber-post_Yellow"
  PrimRubberPostT11.image = "rubber-post_Yellow"
  PrimRubberPostT4.image = "rubber-post_Yellow"
  PrimRubberPostT015.image = "rubber-post_Yellow"
  PrimRubberPostT2.image = "rubber-post_Yellow"
  PrimRubberPostT024.image = "rubber-post_Yellow"
End If

If RubberPostColorType = 3 Then
  PrimRubberPostT10.image = "rubber-post_Orange"
  PrimRubberPostT9.image = "rubber-post_Orange"
  PrimRubberPostT12.image = "rubber-post_Orange"
  PrimRubberPostT6.image = "rubber-post_Orange"
  PrimRubberPostT3.image = "rubber-post_Orange"
  PrimRubberPostT5.image = "rubber-post_Orange"
  PrimRubberPostT17.image = "rubber-post_Orange"
  PrimRubberPostT11.image = "rubber-post_Orange"
  PrimRubberPostT4.image = "rubber-post_Orange"
  PrimRubberPostT015.image = "rubber-post_Orange"
  PrimRubberPostT2.image = "rubber-post_Orange"
  PrimRubberPostT024.image = "rubber-post_Orange"
End If

If RubberPostColorType = 4 Then
  PrimRubberPostT10.image = "rubber-post_Green"
  PrimRubberPostT9.image = "rubber-post_Green"
  PrimRubberPostT12.image = "rubber-post_Blue"
  PrimRubberPostT6.image = "rubber-post_Blue"
  PrimRubberPostT3.image = "rubber-post_Blue"
  PrimRubberPostT5.image = "rubber-post_Blue"
  PrimRubberPostT17.image = "rubber-post_Blue"
  PrimRubberPostT11.image = "rubber-post_Blue"
  PrimRubberPostT4.image = "rubber-post_Blue"
  PrimRubberPostT015.image = "rubber-post_Blue"
  PrimRubberPostT2.image = "rubber-post_Blue"
  PrimRubberPostT024.image = "rubber-post_Blue"
End If

If PegColorType = 1 Then
  PegPlasticT001.material = "Peg"
  PegPlasticT002.material = "Peg"
  PegPlasticT003.material = "Peg"
  PegPlasticT2.material = "Peg"
  PegPlasticT3.material = "Peg"
  PegPlasticT5.material = "Peg"
  PegPlasticT6.material = "Peg"
  PegPlasticT7.material = "Peg"
  PegPlasticT8.material = "Peg"
  PegPlasticT9.material = "Peg"
  PegPlasticT10.material = "Peg"
  PegPlasticT11.material = "Peg"
  PegPlasticT12.material = "Peg"
End If

If PegColorType = 2 Then
  PegPlasticT001.material = "PegOrange"
  PegPlasticT002.material = "PegOrange"
  PegPlasticT003.material = "PegOrange"
  PegPlasticT2.material = "PegOrange"
  PegPlasticT3.material = "PegOrange"
  PegPlasticT5.material = "PegOrange"
  PegPlasticT6.material = "PegOrange"
  PegPlasticT7.material = "PegOrange"
  PegPlasticT8.material = "PegOrange"
  PegPlasticT9.material = "PegOrange"
  PegPlasticT10.material = "PegOrange"
  PegPlasticT11.material = "PegOrange"
  PegPlasticT12.material = "PegOrange"
End If


If BlackRampDecalType = 1 Then
  pBlackRampDecalsKSR.Visible = False
  pBlackRampDecals.Visible = False
  pBlackRampDecalsKSR3.Visible = False
End If

If BlackRampDecalType = 2 Then
  If VUKModType = 1 then
    pBlackRampDecals.Visible = True
    pBlackRampDecalsKSR.Visible = False
    pBlackRampDecalsKSR3.Visible = False
  Else
    If KSBRAorBType = 3 then
      pBlackRampDecalsKSR.Visible = False
      pBlackRampDecals.Visible = False
      pBlackRampDecalsKSR3.Visible = True
    Else
      pBlackRampDecalsKSR.Visible = True
      pBlackRampDecals.Visible = False
      pBlackRampDecalsKSR3.Visible = False
    End If
  End If
End If


If SlingKongVersionType = 1 Then
  pSlingKongProtoLeft.Visible = True
  pSlingKongLProtoBracket.Visible = True
  pSlingKongLeft.Visible = False
  pSlingKongLBracket.Visible = False
  SlingKongProtoRight.Visible = True
  pSlingRProtoBracket.Visible = True
  SlingKongRight.Visible = False
  pSlingRBracket.Visible = False
End If

If SlingKongVersionType = 2 Then
  pSlingKongProtoLeft.Visible = False
  pSlingKongLProtoBracket.Visible = False
  pSlingKongLeft.Visible = True
  pSlingKongLBracket.Visible = True
  SlingKongProtoRight.Visible = False
  pSlingRProtoBracket.Visible = False
  SlingKongRight.Visible = True
  pSlingRBracket.Visible = True
End If


'''''Alt Plastics
''Left Plastic

If LeftPlasticType = 1 Then
  Wall37.Image = "KK Plastics3"
End If

If LeftPlasticType = 2 Then
  Wall37.Image = "KK Plastics"
End If

''Center Plastic

If CenterPlasticType = 1 Then
  pPlastics14.Visible = False
  nut010.visible = false
  nut011.visible = false
End If

If CenterPlasticType = 2 Then
  pPlastics14.Visible = True
  nut010.visible = True
  If DTModType = 1 Then
    nut011.visible = False
  Else
    nut011.visible = True
  End If
End If

''VUK Plastics

If VUKModType = 1 Then

If VUKPlasticsType = 1 Then

  pPlastic6.visible = false  'KSR and VUK bottom
  Wall31.collidable = false

  screw050.visible = true
  screw051.visible = true


  pPlastic7.visible = false  'KSR mid
  Wall22a.collidable = false
  screw008.visible = False
  screw006.visible = False
  Primitive043.visible = False
  pPlastic9.visible = false  'VUK mid
  Wall22.collidable = false
  Primitive045.visible = false
  screw009.visible = false
  screw010.visible = false
  pPlastic10.visible = false 'VUK Top
  Wall009.collidable = false
  Primitive044.visible = false
  screw007.visible = false
  pBlackRampShiled.visible = false
  pBlackRampShiledRivits.visible = false

  screw011.visible = false
  screw012.visible = True

  PegMetalT011.visible = False
  Pin002.visible = False

  screw013.visible = false
  screw014.visible = True
  screw015.visible = false

  pPlastic7b.visible = False  'KSR mid Clear
  screw052.visible = False
  screw053.visible = False
  Primitive056.visible = False
  pBRBB.visible = False
  pPlastic11.visible = False  'KSR Protector
  pPlasticCopter.visible = False  'KSR Copter
  pPlasticCopterRivits.visible = False  'KSR Copter Rivits
End If

If VUKPlasticsType = 2 Then
  screw050.visible = False
  screw051.visible = False

  pPlastic6.visible = True  'KSR and VUK bottom
  pPlastic6.image = ""
  pPlastic6.material = "Plastic Black"
  Wall31.collidable = True
  screw011.visible = True
  screw012.visible = True

  PegMetalT011.visible = False
  Pin002.visible = False

  screw013.visible = True
  screw014.visible = True
  screw015.visible = True
  pPlastic7.visible = false  'KSR mid
  pPlastic7.image = ""
  pPlastic7.material = ""
  Wall22a.collidable = false
  screw008.visible = False
  screw006.visible = False
  Primitive043.visible = False
  pPlastic9.visible = false  'VUK mid
  pPlastic9.image = ""
  pPlastic9.material = ""
  Wall22.collidable = false
  Primitive045.visible = false
  screw009.visible = false
  screw010.visible = false
  pPlastic10.visible = false 'VUK Top
  pPlastic10.image = ""
  pPlastic10.material = ""
  Wall009.collidable = false
  Primitive044.visible = false
  screw007.visible = false
  pBlackRampShiled.visible = false
  pBlackRampShiledRivits.visible = false

  pPlastic7b.visible = False  'KSR mid Clear
  screw052.visible = False
  screw053.visible = False
  Primitive056.visible = False
  pBRBB.visible = False
  pPlastic11.visible = False  'KSR Protector
  pPlasticCopter.visible = False  'KSR Copter
  pPlasticCopterRivits.visible = False  'KSR Copter Rivits
End If

If VUKPlasticsType = 3 Then
  screw050.visible = False
  screw051.visible = False

  pPlastic6.visible = True  'KSR and VUK bottom
  pPlastic6.image = "KK Plastics2"
  pPlastic6.material = "Plastic with an image"
  Wall31.collidable = True
  screw011.visible = True
  screw012.visible = True

  PegMetalT011.visible = False
  Pin002.visible = False

  screw013.visible = True
  screw014.visible = false
  screw015.visible = false
  pPlastic7.visible = False  'KSR mid
  pPlastic7.image = ""
  pPlastic7.material = ""
  Wall22a.collidable = False
  screw008.visible = False
  screw006.visible = False
  Primitive043.visible = False
  pPlastic9.visible = True  'VUK mid
  pPlastic9.image = "KK Plastics"
  pPlastic9.material = "Plastic with an image"
  Wall22.collidable = True
  Primitive045.visible = True
  screw009.visible = True
  screw010.visible = True
  pPlastic10.visible = false 'VUK Top
  pPlastic10.image = ""
  pPlastic10.material = ""
  Wall009.collidable = false
  Primitive044.visible = false
  screw007.visible = false
  pBlackRampShiled.visible = true
  pBlackRampShiledRivits.visible = True

  pPlastic7b.visible = False  'KSR mid Clear
  screw052.visible = False
  screw053.visible = False
  Primitive056.visible = False
  pBRBB.visible = False
  pPlastic11.visible = False  'KSR Protector
  pPlasticCopter.visible = False  'KSR Copter
  pPlasticCopterRivits.visible = False  'KSR Copter Rivits
End If

If VUKPlasticsType = 4 Then
  screw050.visible = False
  screw051.visible = False

  pPlastic6.visible = True  'KSR and VUK bottom
  pPlastic6.image = "KK Plastics"
  pPlastic6.material = "Plastic with an image"
  Wall31.collidable = True
  screw011.visible = True
  screw012.visible = True

  PegMetalT011.visible = False
  Pin002.visible = False

  screw013.visible = True
  screw014.visible = false
  screw015.visible = false
  pPlastic7.visible = False  'KSR mid
  pPlastic7.image = ""
  pPlastic7.material = ""
  Wall22a.collidable = False
  screw008.visible = False
  screw006.visible = False
  Primitive043.visible = False
  pPlastic9.visible = True  'VUK mid
  pPlastic9.image = "KK Plastics"
  pPlastic9.material = "Plastic with an image"
  Wall22.collidable = True
  Primitive045.visible = True
  screw009.visible = True
  screw010.visible = false
  pPlastic10.visible = True 'VUK Top
  pPlastic10.image = "KK Plastics"
  pPlastic10.material = "Plastic with an image"
  Wall009.collidable = True
  Primitive044.visible = True
  screw007.visible = True
  pBlackRampShiled.visible = true
  pBlackRampShiledRivits.visible = True

  pPlastic7b.visible = False  'KSR mid Clear
  screw052.visible = False
  screw053.visible = False
  Primitive056.visible = False
  pBRBB.visible = False
  pPlastic11.visible = False  'KSR Protector
  pPlasticCopter.visible = False  'KSR Copter
  pPlasticCopterRivits.visible = False  'KSR Copter Rivits
End If

If VUKPlasticsType = 5 Then
  screw050.visible = False
  screw051.visible = False

  pPlastic6.visible = True  'KSR and VUK bottom
  pPlastic6.image = "KK Plastics3"
  pPlastic6.material = "Plastic with an image"
  Wall31.collidable = True
  screw011.visible = True
  screw012.visible = True

  PegMetalT011.visible = False
  Pin002.visible = False

  screw013.visible = True
  screw014.visible = false
  screw015.visible = false
  pPlastic7.visible = False  'KSR mid
  pPlastic7.image = ""
  pPlastic7.material = ""
  Wall22a.collidable = False
  screw008.visible = False
  screw006.visible = False
  Primitive043.visible = False
  pPlastic9.visible = True  'VUK mid
  pPlastic9.image = "KK Plastics3"
  pPlastic9.material = "Plastic with an image"
  Wall22.collidable = True
  Primitive045.visible = True
  screw009.visible = True
  screw010.visible = false
  pPlastic10.visible = True 'VUK Top
  pPlastic10.image = "KK Plastics3"
  pPlastic10.material = "Plastic with an image"
  Wall009.collidable = True
  Primitive044.visible = True
  screw007.visible = True
  pBlackRampShiled.visible = true
  pBlackRampShiledRivits.visible = True

  pPlastic7b.visible = False  'KSR mid Clear
  screw052.visible = False
  screw053.visible = False
  Primitive056.visible = False
  pBRBB.visible = False
  pPlastic11.visible = False  'KSR Protector
  pPlasticCopter.visible = False  'KSR Copter
  pPlasticCopterRivits.visible = False  'KSR Copter Rivits
End If

Else

If KSRPlasticsType = 1 Then
  screw050.visible = True
  screw051.visible = True

  pPlastic6.visible = False  'KSR and VUK bottom
  pPlastic6.image = "KK Plastics"
  pPlastic6.material = "Plastic with an image"
  Wall31.collidable = False
  screw011.visible = false
  screw012.visible = false

  PegMetalT011.visible = False
  Pin002.visible = False

  screw013.visible = false
  screw014.visible = false
  screw015.visible = false
  pPlastic7.visible = false  'KSR mid
  pPlastic7.image = "Plastic with an image"
  pPlastic7.material = ""
  screw008.visible = False
  screw006.visible = False
  Wall22a.collidable = false
  screw008.visible = False
  screw006.visible = False
  Primitive043.visible = False
  pPlastic9.visible = false  'VUK mid
  pPlastic9.image = ""
  pPlastic9.material = ""
  Wall22.collidable = false
  Primitive045.visible = False
  screw009.visible = False
  screw010.visible = false
  pPlastic10.visible = false 'VUK Top
  pPlastic10.image = ""
  pPlastic10.material = ""
  Primitive044.visible = false
  screw007.visible = false
  Primitive043.visible = false
  Wall009.collidable = false
  pBlackRampShiled.visible = false
  pBlackRampShiledRivits.visible = false

  pPlastic7b.visible = False  'KSR mid Clear
  screw052.visible = False
  screw053.visible = False
  Primitive056.visible = False
  pBRBB.visible = False
  pPlastic11.visible = False  'KSR Protector
  pPlasticCopter.visible = False  'KSR Copter
  pPlasticCopterRivits.visible = False  'KSR Copter Rivits
End If

If KSRPlasticsType = 2 Then
  screw050.visible = False
  screw051.visible = False

  pPlastic6.visible = True  'KSR and VUK bottom
  pPlastic6.image = "KK Plastics2"
  pPlastic6.material = "Plastic with an image"
  Wall31.collidable = True
  screw011.visible = True
  screw012.visible = True

  PegMetalT011.visible = False
  Pin002.visible = False

  screw013.visible = True
  screw014.visible = True
  screw015.visible = True
  pPlastic7.visible = false  'KSR mid
  pPlastic7.image = "KK Plastics"
  pPlastic7.material = "Plastic with an image"
  Wall22a.collidable = false
  screw008.visible = False
  screw006.visible = False
  Primitive043.visible = False
  pPlastic9.visible = false  'VUK mid
  pPlastic9.image = ""
  pPlastic9.material = ""
  Wall22.collidable = false
  Primitive045.visible = False
  screw009.visible = False
  screw010.visible = false
  pPlastic10.visible = false 'VUK Top
  pPlastic10.image = ""
  pPlastic10.material = ""
  Wall009.collidable = false
  Primitive044.visible = false
  screw007.visible = false
  pBlackRampShiled.visible = false
  pBlackRampShiledRivits.visible = false

  pPlastic7b.visible = False  'KSR mid Clear
  screw052.visible = False
  screw053.visible = False
  Primitive056.visible = False
  pBRBB.visible = False
  pPlastic11.visible = False  'KSR Protector
  pPlasticCopter.visible = False  'KSR Copter
  pPlasticCopterRivits.visible = False  'KSR Copter Rivits
End If

If KSRPlasticsType = 3 Then
  screw050.visible = False
  screw051.visible = False

  pPlastic6.visible = True  'KSR and VUK bottom
  pPlastic6.image = "KK Plastics"
  pPlastic6.material = "Plastic with an image"
  Wall31.collidable = True
  screw011.visible = True
  screw012.visible = True

  PegMetalT011.visible = False
  Pin002.visible = False

  screw013.visible = True
  screw014.visible = false
  screw015.visible = false
  pPlastic7.visible = True  'KSR mid
  pPlastic7.image = "KK Plastics"
  pPlastic7.material = "Plastic with an image"
  Wall22a.collidable = True
  screw008.visible = True
  screw006.visible = True
  Primitive043.visible = True
  pPlastic9.visible = false  'VUK mid
  pPlastic9.image = ""
  pPlastic9.material = ""
  Wall22.collidable = false
  Primitive045.visible = False
  screw009.visible = False
  screw010.visible = false
  pPlastic10.visible = false 'VUK Top
  pPlastic10.image = ""
  pPlastic10.material = ""
  Wall009.collidable = false
  Primitive044.visible = false
  screw007.visible = false
  pBlackRampShiled.visible = false
  pBlackRampShiledRivits.visible = false

  pPlastic7b.visible = False  'KSR mid Clear
  screw052.visible = False
  screw053.visible = False
  Primitive056.visible = False
  pBRBB.visible = False
  pPlastic11.visible = False  'KSR Protector
  pPlasticCopter.visible = False  'KSR Copter
  pPlasticCopterRivits.visible = False  'KSR Copter Rivits
End If

If KSRPlasticsType = 4 Then
  screw050.visible = False
  screw051.visible = False

  pPlastic6.visible = True  'KSR and VUK bottom
  pPlastic6.image = "KK Plastics3"
  pPlastic6.material = "Plastic with an image"
  Wall31.collidable = True
  screw011.visible = True
  screw012.visible = True

  PegMetalT011.visible = False
  Pin002.visible = False

  screw013.visible = True
  screw014.visible = false
  screw015.visible = false
  pPlastic7.visible = True  'KSR mid
  pPlastic7.image = "KK Plastics3"
  pPlastic7.material = "Plastic with an image"
  Wall22a.collidable = True
  screw008.visible = True
  screw006.visible = True
  Primitive043.visible = True
  pPlastic9.visible = false  'VUK mid
  pPlastic9.image = ""
  pPlastic9.material = ""
  Wall22.collidable = false
  Primitive045.visible = False
  screw009.visible = False
  screw010.visible = false
  pPlastic10.visible = false 'VUK Top
  pPlastic10.image = ""
  pPlastic10.material = ""
  Wall009.collidable = false
  screw007.visible = false
  pBlackRampShiled.visible = false
  Primitive044.visible = false
  pBlackRampShiledRivits.visible = false

  pPlastic7b.visible = False  'KSR mid Clear
  screw052.visible = False
  screw053.visible = False
  Primitive056.visible = False
  pBRBB.visible = False
  pPlastic11.visible = False  'KSR Protector
  pPlasticCopter.visible = False  'KSR Copter
  pPlasticCopterRivits.visible = False  'KSR Copter Rivits
End If


If KSRPlasticsType = 5 Then
  screw050.visible = False
  screw051.visible = False

  pPlastic6.visible = True  'KSR and VUK bottom
  pPlastic6.image = "KK Plastics2"
  pPlastic6.material = "Plastic with an image"
  Wall31.collidable = True
  screw011.visible = True
  screw012.visible = False

  PegMetalT011.visible = True
  Pin002.visible = True

  screw013.visible = True
  screw014.visible = True
  screw015.visible = True
  pPlastic7.visible = false  'KSR mid
  pPlastic7.image = "KK Plastics"
  pPlastic7.material = "Plastic with an image"
  Wall22a.collidable = false
  screw008.visible = True
  screw006.visible = True
  Primitive043.visible = True
  pPlastic9.visible = false  'VUK mid
  pPlastic9.image = ""
  pPlastic9.material = ""
  Wall22.collidable = false
  Primitive045.visible = False
  screw009.visible = False
  screw010.visible = false
  pPlastic10.visible = false 'VUK Top
  pPlastic10.image = ""
  pPlastic10.material = ""
  Wall009.collidable = false
  Primitive044.visible = false
  screw007.visible = false
  pBlackRampShiled.visible = false
  pBlackRampShiledRivits.visible = false

  pPlastic7b.visible = True  'KSR mid Clear
  screw052.visible = True
  screw053.visible = True
  Primitive056.visible = True
  pBRBB.visible = True
  pPlastic11.visible = True  'KSR Protector
  pPlasticCopter.visible = True  'KSR Copter
  pPlasticCopterRivits.visible = True  'KSR Copter Rivits
End If

End If

  If WackerDrivePlasticType = 0 Then
    pWackerDrive.Visible = False
    pWackerDrive2.Visible = False
    pWackerDrive_bracket.Visible = False
  End If

  If WackerDrivePlasticType = 1 Then
    pWackerDrive.Visible = False
    pWackerDrive2.Visible = True
    pWackerDrive_bracket.Visible = True
  End If

  If WackerDrivePlasticType = 2 Then
    pWackerDrive.Visible = True
    pWackerDrive2.Visible = False
    pWackerDrive_bracket.Visible = True
  End If

  If WackerPostType = 0 then
    PegMetalT010.Visible = False
    Pin001.Visible = False
  End If

  If WackerPostType = 1 then
    PegMetalT010.Visible = True
    Pin001.Visible = False
  End If

  If WackerPostType = 2 then
    PegMetalT010.Visible = True
    Pin001.Visible = True
  End If

If LeftSBLPPlasticType = 0 Then
  Primitive051.Visible = False
  Primitive046.Visible = False
  Primitive047.Visible = False
  pProto9RampProtector.Visible = False
  screw036.Visible = False
  PrimRubberPostT026.Visible = False
  screw058.Visible = False
  PrimRubberPostT027.Visible = False
  screw032.Visible = False
  screw033.Visible = False
  screw034.Visible = False
  screw035.Visible = False
  PrimRubberPostT018.Visible = False
  PrimRubberPostT019.Visible = False
  PrimRubberPostT020.Visible = False
  PrimRubberPostT021.Visible = False
  PrimRubberPostT022.Visible = False
  PrimRubberPostT023.Visible = False
  Primitive048.Visible = False
  screw036.Visible = False
End If

If LeftSBLPPlasticType = 1 Then

  Primitive051.Visible = False
  Primitive046.Visible = True
  pProto9RampProtector.Visible = False
  screw036.Visible = False
  PrimRubberPostT026.Visible = False
  screw058.Visible = False
  PrimRubberPostT027.Visible = False
  Primitive046.Material = "Plastic Black"
  Primitive046.DisableLighting = 0
  Primitive047.Visible = False
  screw032.Visible = True
  screw033.Visible = True
  screw034.Visible = False
  screw035.Visible = False
  PrimRubberPostT018.Visible = False
  PrimRubberPostT019.Visible = False
  PrimRubberPostT020.Visible = False
  PrimRubberPostT021.Visible = False
  PrimRubberPostT022.Visible = False
  PrimRubberPostT023.Visible = False
  Primitive048.Visible = False
  screw036.Visible = False
End If

If LeftSBLPPlasticType = 2 Then
  Primitive051.Visible = False
  Primitive046.Visible = True
  pProto9RampProtector.Visible = False
  screw036.Visible = False
  PrimRubberPostT026.Visible = False
  screw058.Visible = False
  PrimRubberPostT027.Visible = False
  Primitive046.Material = "AcrylicBLOrange"
  Primitive046.DisableLighting = .85
  Primitive047.Visible = False
  screw032.Visible = True
  screw033.Visible = True
  screw034.Visible = False
  screw035.Visible = False
  PrimRubberPostT018.Visible = True
  PrimRubberPostT019.Visible = True
  PrimRubberPostT020.Visible = False
  PrimRubberPostT021.Visible = False
  PrimRubberPostT022.Visible = False
  PrimRubberPostT023.Visible = False
  Primitive048.Visible = False
  screw036.Visible = False

End If

If LeftSBLPPlasticType = 3 Then
  Primitive046.Visible = False
  Primitive047.Visible = True
  pProto9RampProtector.Visible = False
  screw036.Visible = False
  PrimRubberPostT026.Visible = False
  screw058.Visible = False
  PrimRubberPostT027.Visible = False
  screw032.Visible = True
  screw033.Visible = True
  screw034.Visible = False
  screw035.Visible = False
  PrimRubberPostT018.Visible = False
  PrimRubberPostT019.Visible = False
  PrimRubberPostT020.Visible = False
  PrimRubberPostT021.Visible = False
  PrimRubberPostT022.Visible = True
  PrimRubberPostT023.Visible = True
  Primitive051.Visible = False
  Primitive048.Visible = False
  screw036.Visible = False
End If

If LeftSBLPPlasticType = 4 Then
  Primitive051.Visible = True
  Primitive046.Visible = False
  Primitive047.Visible = True
  pProto9RampProtector.Visible = False
  screw036.Visible = False
  PrimRubberPostT026.Visible = False
  screw058.Visible = False
  PrimRubberPostT027.Visible = False
  screw032.Visible = False
  screw033.Visible = False
  PrimRubberPostT018.Visible = False
  PrimRubberPostT019.Visible = False
  PrimRubberPostT020.Visible = True
  PrimRubberPostT021.Visible = True
  PrimRubberPostT022.Visible = True
  PrimRubberPostT023.Visible = True
  screw034.Visible = True
  screw035.Visible = True
  Primitive048.Visible = True
  screw036.Visible = True
End If

If LeftSBLPPlasticType = 5 Then
  Primitive051.Visible = True
  Primitive046.Visible = False
  Primitive047.Visible = False
  pProto9RampProtector.Visible = True
  screw036.Visible = True
  PrimRubberPostT026.Visible = True
  screw058.Visible = True
  PrimRubberPostT027.Visible = True
  screw032.Visible = False
  screw033.Visible = False
  PrimRubberPostT018.Visible = False
  PrimRubberPostT019.Visible = False
  PrimRubberPostT020.Visible = True
  PrimRubberPostT021.Visible = True
  PrimRubberPostT022.Visible = True
  PrimRubberPostT023.Visible = True
  screw034.Visible = True
  screw035.Visible = True
  Primitive048.Visible = False
  screw036.Visible = False
End If



If LeftSBLPType = 0 Then
  Primitive049.Visible = False
  screw049.Visible = False
  Primitive050.Visible = False
  Primitive050.Collidable = False

  PrimitiveFomeBlock1.Visible = False
  PrimitiveFomeBlock2.Visible = False
  PrimitiveFomeBlock1.Collidable = False
  PrimitiveFomeBlock2.Collidable = False

  DoublePegPlastic1.Visible = False
  DoublePegPlastic2.Visible = False

  PegPlasticT004.Visible = False
  PegPlasticT005.Visible = False
  Rubber001.Visible = False
  screw044.Visible = False
  screw043.Visible = False

  PegPlasticT006.Visible = False
  screw045.Visible = False
  PrimRubberPostT025.Visible = False
  screw047.Visible = False
End If

If LeftSBLPType = 1 Then
  Primitive049.Visible = True
  screw049.Visible = True
  Primitive050.Visible = True
  Primitive050.Collidable = True

  PrimitiveFomeBlock1.Visible = False
  PrimitiveFomeBlock2.Visible = False
  PrimitiveFomeBlock1.Collidable = False
  PrimitiveFomeBlock2.Collidable = False

  DoublePegPlastic1.Visible = False
  DoublePegPlastic2.Visible = False

  PegPlasticT004.Visible = False
  PegPlasticT005.Visible = False
  Rubber001.Visible = False
  screw044.Visible = False
  screw043.Visible = False

  Wall021.IsDropped = True

  PegPlasticT006.Visible = False
  screw045.Visible = False
  PrimRubberPostT025.Visible = False
  screw047.Visible = False
End If

If LeftSBLPType = 2 Then
  Primitive049.Visible = False
  screw049.Visible = False
  Primitive050.Visible = False
  Primitive050.Collidable = False

  PrimitiveFomeBlock1.Visible = True
  PrimitiveFomeBlock2.Visible = True
  PrimitiveFomeBlock1.Collidable = True
  PrimitiveFomeBlock2.Collidable = True

  DoublePegPlastic1.Visible = False
  DoublePegPlastic2.Visible = False

  PegPlasticT004.Visible = False
  PegPlasticT005.Visible = False
  Rubber001.Visible = False
  screw044.Visible = False
  screw043.Visible = False

  Wall021.IsDropped = True

  PegPlasticT006.Visible = True
  screw045.Visible = True
  PrimRubberPostT025.Visible = True
  screw047.Visible = False
End If

If LeftSBLPType = 3 Then
  Primitive049.Visible = False
  screw049.Visible = False
  Primitive050.Visible = False
  Primitive050.Collidable = False

  PrimitiveFomeBlock1.Visible = False
  PrimitiveFomeBlock2.Visible = False
  PrimitiveFomeBlock1.Collidable = False
  PrimitiveFomeBlock2.Collidable = False

  DoublePegPlastic1.Visible = True
  DoublePegPlastic2.Visible = True

  PegPlasticT004.Visible = True
  PegPlasticT005.Visible = True
  Rubber001.Visible = True
  screw044.Visible = True
  screw043.Visible = True

  Wall021.IsDropped = False

  PegPlasticT006.Visible = False
  screw045.Visible = True
  PrimRubberPostT025.Visible = False
  screw047.Visible = True
End If

If LeftSBLPType = 4 Then
  Primitive049.Visible = True
  screw049.Visible = True
  Primitive050.Visible = True
  Primitive050.Collidable = True

  PrimitiveFomeBlock1.Visible = False
  PrimitiveFomeBlock2.Visible = False
  PrimitiveFomeBlock1.Collidable = False
  PrimitiveFomeBlock2.Collidable = False

  DoublePegPlastic1.Visible = True
  DoublePegPlastic2.Visible = True

  PegPlasticT004.Visible = False
  PegPlasticT005.Visible = False
  Rubber001.Visible = False
  screw044.Visible = False
  screw043.Visible = False

  Wall021.IsDropped = True

  PegPlasticT006.Visible = False
  screw045.Visible = True
  PrimRubberPostT025.Visible = False
  screw047.Visible = True
End If

If Plastic8HeightType = 1 Then
  pPlastic8.TransY = 0
  Cap1.TransZ = 0
  Cap007.TransZ = 0
  Primitive054.Visible = False
End If


If Plastic8HeightType = 2 Then
  pPlastic8.TransY = 35
  Cap1.TransZ = 35
  Cap007.TransZ = 35
  Primitive054.Visible = True
End If

If LeftWireRampSupportConfigType = 1 Then
  PrimRubberPostT024.Visible = True
  MetalPost2.Visible = False
  Rubber9.Visible = False
  Wall003.Collidable = False
  Rubber003.Visible = True
  Primitive3.Material = "Plastic Black"
  Primitive052.Visible = True
End If


If LeftWireRampSupportConfigType = 2 Then
  PrimRubberPostT024.Visible = False
  MetalPost2.Visible = True
  Rubber9.Visible = True
  Wall003.Collidable = True
  Rubber003.Visible = False
  Primitive3.Material = "Plastic with an image"
  Primitive052.Visible = False
End If


If RightWireRampGuideType = 1 Then
  pRightWireRampB.Visible = True
  pRightWireRampC.Visible = False
  pRightWireRampC.Visible = False
  Primitive006.Visible = false
  PrimRubberPostT017.Visible = False
  screw031.Visible = False
  PrimRubberPostT016.Visible = False
  screw030.Visible = False
End If

If RightWireRampGuideType = 2 Then
  pRightWireRampB.Visible = False
  pRightWireRampC.Visible = True
  Primitive006.Visible = True
  PrimRubberPostT017.Visible = True
  screw031.Visible = True
  PrimRubberPostT016.Visible = True
  screw030.Visible = True
End If



''Right

If RightWireRampColorType = 1 Then
  pRightWireRamp.Material = "Chrome"
  pRightWireRampB.Material = "Chrome"
  pRightWireRampC.Material = "Chrome"
End If

If RightWireRampColorType = 2 Then
  pRightWireRamp.Material = "Yellow Wire Ramp"
  pRightWireRampB.Material = "Yellow Wire Ramp"
  pRightWireRampC.Material = "Yellow Wire Ramp"
End If

If PlasticRampWedgeType = 0 Then
  Rivit5.visible = False
  Rivit6.visible = False
  FlatScrew1.Visible = True
  FlatScrew2.Visible = True
  pRampWedge3.Visible = False
End If

If PlasticRampWedgeType = 1 Then
  Rivit5.visible = True
  Rivit6.visible = True
  FlatScrew1.Visible = False
  FlatScrew2.Visible = False
  pRampWedge3.Visible = True
End If

If PlasticRampWedgeType = 2 Then
  Rivit5.visible = True
  Rivit6.visible = True
  FlatScrew1.Visible = False
  FlatScrew2.Visible = False
  pRampWedge3.Visible = True
  pRampWedge3.Material = "Plastic with an image"
End If

If PlasticRampWedgeType = 3 Then
  Rivit5.visible = True
  Rivit6.visible = True
  FlatScrew1.Visible = False
  FlatScrew2.Visible = False
  pRampWedge3.Visible = True
  pRampWedge3.Material = "Plastic with an image dark"
End If

If BlackRampWedge1type = 0 Then
  Rivit1.visible = False
  Rivit2.visible = False
  pRampWedge1.Material = "Plastic with an image"
  FlatScrew3.Visible = True
  FlatScrew4.Visible = True
  pRampWedge1.visible = False

End If

If BlackRampWedge1type = 1 Then
  Rivit1.Visible = True
  Rivit2.Visible = True
  pRampWedge1.visible = True
  pRampWedge1.Material = "Plastic with an image"
  FlatScrew3.Visible = False
  FlatScrew4.Visible = False
End If

If BlackRampWedge1type = 2 Then
  Rivit1.Visible = True
  Rivit2.Visible = True
  pRampWedge1.visible = True
  pRampWedge1.Material = "Plastic with an image"
  FlatScrew3.Visible = True
  FlatScrew4.Visible = True
End If

If BlackRampWedge1type = 3 Then
  pRampWedge1.visible = True
  pRampWedge1.Material = "Plastic with an image dark"
  FlatScrew3.Visible = True
  FlatScrew4.Visible = True
End If



If BlackRampWedge2type = 0 Then
  pRampWedge2.visible = False
  pRampWedge2Screws.Visible = True
  pRampWedge2Rivits.Visible = False
  pRampWedge2b.visible = False
  pRampWedge2bScrews.Visible = False
  pRampWedge2bRivits.Visible = False
End If

If BlackRampWedge2type = 1 Then
  pRampWedge2.visible = False
  pRampWedge2Screws.Visible = False
  pRampWedge2Rivits.Visible = False
  pRampWedge2b.visible = True
  pRampWedge2bScrews.Visible = True
  pRampWedge2bRivits.Visible = True
End If

If BlackRampWedge2type = 2 Then
  pRampWedge2.visible = True
  pRampWedge2.Material = "Plastic with an image"
  pRampWedge2Screws.Visible = True
  pRampWedge2Rivits.Visible = True
  pRampWedge2b.visible = False
  pRampWedge2bScrews.Visible = False
  pRampWedge2bRivits.Visible = False
End If

If BlackRampWedge2type = 3 Then
  pRampWedge2.visible = True
  pRampWedge2.Material = "Plastic with an image dark"
  pRampWedge2Screws.Visible = True
  pRampWedge2Rivits.Visible = True
  pRampWedge2b.visible = False
  pRampWedge2bScrews.Visible = False
  pRampWedge2bRivits.Visible = False
End If

If  KSRSwitchPlasticType = 0 Then
  KSRswitchProtectorScrews.Visible = False

  KSRswitchProtectorA.Visible = False
  KSRswitchProtectorSupportA.Visible = False

  KSRswitchProtectorB.Visible = False
  KSRswitchProtectorSupportB.Visible = False

  KSRswitchProtectorScrews2.Visible = false
  KSRswitchProtectorSupportc.Visible = false
End If

If  KSRSwitchPlasticType = 1 Then
  KSRswitchProtectorScrews.Visible = True

  KSRswitchProtectorA.Visible = True
  KSRswitchProtectorSupportA.Visible = True

  KSRswitchProtectorB.Visible = False
  KSRswitchProtectorSupportB.Visible = False

  KSRswitchProtectorScrews2.Visible = false
  KSRswitchProtectorSupportc.Visible = false
End If

If  KSRSwitchPlasticType = 2 Then
  KSRswitchProtectorScrews.Visible = True

  KSRswitchProtectorA.Visible = False
  KSRswitchProtectorSupportA.Visible = False

  KSRswitchProtectorB.Visible = True
  KSRswitchProtectorSupportB.Visible = True

  KSRswitchProtectorScrews2.Visible = false
  KSRswitchProtectorSupportc.Visible = false
End If

If  KSRSwitchPlasticType = 3 Then
  KSRswitchProtectorScrews.Visible = False

  KSRswitchProtectorA.Visible = False
  KSRswitchProtectorSupportA.Visible = False

  KSRswitchProtectorB.Visible = True
  KSRswitchProtectorSupportB.Visible = False

  KSRswitchProtectorScrews2.Visible = True
  KSRswitchProtectorSupportc.Visible = True
End If

If  DoubleRubbersType = 1 Then
  Rubber12.Visible = False
  Rubber12A.visible = True
  Rubber12B.visible = True
  Rubber11.visible = False
  Rubber11A.Visible = True
  Rubber11B.Visible = True
  pRails_back.visible = false
Else
  Rubber12.Visible = True
  Rubber12A.visible = False
  Rubber12B.visible = False
  Rubber11.visible = True
  Rubber11A.Visible = False
  Rubber11B.Visible = False
  pRails_back.visible = True
End If

If OptinalRails1Type = 1 Then
  pRailsOptinal1.Visible = True
  Rubber13.Visible = True
  PegMetalT002.Visible = True
  PegMetalT003.Visible = True
  Wall014.IsDropped = False
  pWall021.visible = false
  WallBracket1.visible = false
  WallBracket2.visible = false
  WallBracket1a.visible = false
  WallBracket2a.visible = false
End If

If OptinalRails1Type = 2 Then
  pRailsOptinal1.Visible = False
  pRailsOptinal1.Visible = False
  Rubber13.Visible = False
  PegMetalT002.Visible = False
  PegMetalT003.Visible = False
  Wall014.IsDropped = True
  pWall021.visible = True
  WallBracket1.visible = True
  WallBracket2.visible = True
  WallBracket1a.visible = True
  WallBracket2a.visible = True
End If

If OptinalRails2Type = 1 Then
  pRailsOptinal2.Visible = True
  Wall019.Visible = False
  Wall019.SideVisible = False
End If
If OptinalRails2Type = 2 Then
  pRailsOptinal2.Visible = False
  Wall019.Visible = True
  Wall019.SideVisible = True
End If

If InsertDecalsType = 1 Then
  pInsertDecalSpecialLeft.Visible = False
  pInsertDecalSpecialRight.Visible = False
  pInsertDecalEverythingLit.Visible = False
End If

If InsertDecalsType = 2 Then
  pInsertDecalSpecialLeft.Visible = True
  pInsertDecalSpecialRight.Visible = True
  pInsertDecalEverythingLit.Visible = True
End If




'''''''''''''''''''''''''
'''''  GI Color Mod '''''
'''''''''''''''''''''''''
Dim Red, RedFull, RedI, Pink, PinkFull, PinkI, White, WhiteFull, WhiteI, Blue, BlueFull, BlueI, Yellow, YellowFull, YellowI, Green, GreenFull, GreenI, Orange, OrangeFull, OrangeI, Purple, PurpleFull, PurpleI

RedFull = rgb(255,0,0)
Red = rgb(255,0,0)
RedI = 5
PinkFull = rgb(255,0,128)
Pink = rgb(255,0,255)
PinkI = 5
WhiteFull = rgb(255,255,128)
White = rgb(255,255,255)
WhiteI = 7
BlueFull = rgb(0,75,255)
Blue = rgb(0,100,255)
BlueI = 30
YellowFull = rgb(255,255,128)
Yellow = rgb(255,255,0)
YellowI = 15
GreenFull = rgb(128,255,128)
Green = rgb(0,255,0)
GreenI = 20
PurpleFull = rgb(128,0,255)
Purple = rgb(64,0,128)
PurpleI = 2
OrangeFull = rgb(172,63,4)
Orange = rgb(172,63,4)
OrangeI = 10


If GIColorMod = 1 Then

    gi1a.Color=White
    gi1a.ColorFull=WhiteFull
    gi1b.Color=Yellow
    gi1b.ColorFull=YellowFull
    gi1c.Color=White
    gi1c.ColorFull=WhiteFull
    gi1a.Intensity = WhiteI
    pBulb001.Material = "BulbGIOff"

    gi2a.Color=White
    gi2a.ColorFull=WhiteFull
    gi2b.Color=Yellow
    gi2b.ColorFull=YellowFull
    gi2c.Color=White
    gi2c.ColorFull=WhiteFull
    gi2a.Intensity = WhiteI
    pBulb002.Material = "BulbGIOff"

    gi3a.Color=White
    gi3a.ColorFull=WhiteFull
    gi3b.Color=Yellow
    gi3b.ColorFull=YellowFull
    gi3c.Color=White
    gi3c.ColorFull=WhiteFull
    gi3a.Intensity = WhiteI
    pBulb003.Material = "BulbGIOff"

    gi4a.Color=White
    gi4a.ColorFull=WhiteFull
    gi4b.Color=Yellow
    gi4b.ColorFull=YellowFull
    gi4c.Color=White
    gi4c.ColorFull=WhiteFull
    gi4a.Intensity = WhiteI
    pBulb004.Material = "BulbGIOff"

    gi5a.Color=White
    gi5a.ColorFull=WhiteFull
    gi5b.Color=Yellow
    gi5b.ColorFull=YellowFull
    gi5c.Color=White
    gi5c.ColorFull=WhiteFull
    gi5a.Intensity = WhiteI
    pBulb005.Material = "BulbGIOff"

    gi6a.Color=White
    gi6a.ColorFull=WhiteFull
    gi6b.Color=Yellow
    gi6b.ColorFull=YellowFull
    gi6c.Color=White
    gi6c.ColorFull=WhiteFull
    gi6a.Intensity = WhiteI
    pBulb006.Material = "BulbGIOff"

    gi7a.Color=White
    gi7a.ColorFull=WhiteFull
    gi7b.Color=Yellow
    gi7b.ColorFull=YellowFull
    gi7c.Color=White
    gi7c.ColorFull=WhiteFull
    gi7a.Intensity = WhiteI
    pBulb007.Material = "BulbGIOff"

'   gi8a.Color=White
'   gi8a.ColorFull=WhiteFull
''    gi8b.Color=Yellow
''    gi8b.ColorFull=YellowFull
''    gi8c.Color=White
''    gi8c.ColorFull=WhiteFull
'   gi8a.Intensity = WhiteI
''    pBulb008.Material = "BulbGIOff"

'   gi9a.Color=White
'   gi9a.ColorFull=WhiteFull
    gi9b.Color=Yellow
    gi9b.ColorFull=YellowFull
    gi9c.Color=White
    gi9c.ColorFull=WhiteFull
'   gi9a.Intensity = WhiteI
    pBulb009.Material = "BulbGIOff"

    gi10a.Color=White
    gi10a.ColorFull=WhiteFull
    gi10b.Color=Yellow
    gi10b.ColorFull=YellowFull
    gi10c.Color=White
    gi10c.ColorFull=WhiteFull
    gi10a.Intensity = WhiteI
    pBulb010.Material = "BulbGIOff"

    gi11a.Color=White
    gi11a.ColorFull=WhiteFull
    gi11b.Color=Yellow
    gi11b.ColorFull=YellowFull
    gi11c.Color=White
    gi11c.ColorFull=WhiteFull
    gi11a.Intensity = WhiteI
    pBulb011.Material = "BulbGIOff"

    gi12a.Color=White
    gi12a.ColorFull=WhiteFull
    gi12b.Color=Yellow
    gi12b.ColorFull=YellowFull
    gi12c.Color=White
    gi12c.ColorFull=WhiteFull
    gi12a.Intensity = WhiteI
    pBulb012.Material = "BulbGIOff"

    gi13a.Color=White
    gi13a.ColorFull=WhiteFull
    gi13b.Color=Yellow
    gi13b.ColorFull=YellowFull
    gi13c.Color=White
    gi13c.ColorFull=WhiteFull
    gi13a.Intensity = WhiteI
    pBulb013.Material = "BulbGIOff"

    gi14a.Color=White
    gi14a.ColorFull=WhiteFull
    gi14b.Color=Yellow
    gi14b.ColorFull=YellowFull
    gi14c.Color=White
    gi14c.ColorFull=WhiteFull
    gi14a.Intensity = WhiteI
    pBulb014.Material = "BulbGIOff"

    gi15a.Color=White
    gi15a.ColorFull=WhiteFull
    gi15b.Color=Yellow
    gi15b.ColorFull=YellowFull
    gi15c.Color=White
    gi15c.ColorFull=WhiteFull
    gi15a.Intensity = WhiteI
    pBulb015.Material = "BulbGIOff"

    gi16a.Color=White
    gi16a.ColorFull=WhiteFull
    gi16b.Color=Yellow
    gi16b.ColorFull=YellowFull
    gi16c.Color=White
    gi16c.ColorFull=WhiteFull
    gi16a.Intensity = WhiteI
    pBulb016.Material = "BulbGIOff"

    gi17a.Color=White
    gi17a.ColorFull=WhiteFull
    gi17b.Color=Yellow
    gi17b.ColorFull=YellowFull
    gi17c.Color=White
    gi17c.ColorFull=WhiteFull
    gi17a.Intensity = WhiteI
    pBulb017.Material = "BulbGIOff"

    gi18a.Color=White
    gi18a.ColorFull=WhiteFull
    gi18b.Color=Yellow
    gi18b.ColorFull=YellowFull
    gi18c.Color=White
    gi18c.ColorFull=WhiteFull
    gi18a.Intensity = WhiteI
    pBulb018.Material = "BulbGIOff"
End If

If GIColorMod = 2 Then
    gi1a.Color=Yellow
    gi1a.ColorFull=YellowFull
    gi1b.Color=Yellow
    gi1b.ColorFull=YellowFull
    gi1c.Color=Yellow
    gi1c.ColorFull=YellowFull
    gi1a.Intensity = YellowI
    pBulb001.Material = "BulbYellowGIOff"

    gi2a.Color=Orange
    gi2a.ColorFull=OrangeFull
    gi2b.Color=Orange
    gi2b.ColorFull=OrangeFull
    gi2c.Color=Orange
    gi2c.ColorFull=OrangeFull
    gi2a.Intensity = OrangeI
    pBulb002.Material = "BulbOrangeGIOff"

    gi3a.Color=Orange
    gi3a.ColorFull=OrangeFull
    gi3b.Color=Orange
    gi3b.ColorFull=OrangeFull
    gi3c.Color=Orange
    gi3c.ColorFull=OrangeFull
    gi3a.Intensity = OrangeI
    pBulb003.Material = "BulbOrangeGIOff"

    gi4a.Color=Orange
    gi4a.ColorFull=OrangeFull
    gi4b.Color=Orange
    gi4b.ColorFull=OrangeFull
    gi4c.Color=Orange
    gi4c.ColorFull=OrangeFull
    gi4a.Intensity = OrangeI
    pBulb004.Material = "BulbOrangeGIOff"

    gi5a.Color=Orange
    gi5a.ColorFull=OrangeFull
    gi5b.Color=Orange
    gi5b.ColorFull=OrangeFull
    gi5c.Color=Orange
    gi5c.ColorFull=OrangeFull
    gi5a.Intensity = OrangeI
    pBulb005.Material = "BulbOrangeGIOff"

    gi6a.Color=Blue
    gi6a.ColorFull=BlueFull
    gi6b.Color=Blue
    gi6b.ColorFull=BlueFull
    gi6c.Color=Blue
    gi6c.ColorFull=BlueFull
    gi6a.Intensity = BlueI
    pBulb006.Material = "BulbBlueGIOff"

    gi7a.Color=Blue
    gi7a.ColorFull=BlueFull
    gi7b.Color=Blue
    gi7b.ColorFull=BlueFull
    gi7c.Color=Blue
    gi7c.ColorFull=BlueFull
    gi7a.Intensity = BlueI
    pBulb007.Material = "BulbBlueGIOff"

'   gi8a.Color=White
'   gi8a.ColorFull=WhiteFull
''    gi8b.Color=Yellow
''    gi8b.ColorFull=YellowFull
''    gi8c.Color=White
''    gi8c.ColorFull=WhiteFull
'   gi8a.Intensity = WhiteI

    gi9a.Color=Blue
    gi9a.ColorFull=BlueFull
    gi9b.Color=Blue
    gi9b.ColorFull=BlueFull
    gi9c.Color=Blue
    gi9c.ColorFull=BlueFull
    gi9a.Intensity = BlueI
    pBulb009.Material = "BulbBlueGIOff"

    gi10a.Color=Blue
    gi10a.ColorFull=BlueFull
    gi10b.Color=Blue
    gi10b.ColorFull=BlueFull
    gi10c.Color=Blue
    gi10c.ColorFull=BlueFull
    gi10a.Intensity = BlueI
    pBulb010.Material = "BulbBlueGIOff"

    gi11a.Color=Blue
    gi11a.ColorFull=BlueFull
    gi11b.Color=Blue
    gi11b.ColorFull=BlueFull
    gi11c.Color=Blue
    gi11c.ColorFull=BlueFull
    gi11a.Intensity = BlueI
    pBulb011.Material = "BulbBlueGIOff"

    gi12a.Color=Orange
    gi12a.ColorFull=OrangeFull
    gi12b.Color=Orange
    gi12b.ColorFull=OrangeFull
    gi12c.Color=Orange
    gi12c.ColorFull=OrangeFull
    gi12a.Intensity = OrangeI
    pBulb012.Material = "BulbOrangeGIOff"

    gi13a.Color=Orange
    gi13a.ColorFull=OrangeFull
    gi13b.Color=Orange
    gi13b.ColorFull=OrangeFull
    gi13c.Color=Orange
    gi13c.ColorFull=OrangeFull
    gi13a.Intensity = OrangeI
    pBulb013.Material = "BulbOrangeGIOff"

    gi14a.Color=Green
    gi14a.ColorFull=GreenFull
    gi14b.Color=Green
    gi14b.ColorFull=GreenFull
    gi14c.Color=Green
    gi14c.ColorFull=GreenFull
    gi14a.Intensity = GreenI
    pBulb014.Material = "BulbGreenGIOff"

    gi15a.Color=Orange
    gi15a.ColorFull=OrangeFull
    gi15b.Color=Orange
    gi15b.ColorFull=OrangeFull
    gi15c.Color=Orange
    gi15c.ColorFull=OrangeFull
    gi15a.Intensity = OrangeI
    pBulb015.Material = "BulbOrangeGIOff"

    gi16a.Color=Orange
    gi16a.ColorFull=OrangeFull
    gi16b.Color=Orange
    gi16b.ColorFull=OrangeFull
    gi16c.Color=Orange
    gi16c.ColorFull=OrangeFull
    gi16a.Intensity = OrangeI
    pBulb016.Material = "BulbOrangeGIOff"

    gi17a.Color=Yellow
    gi17a.ColorFull=YellowFull
    gi17b.Color=Yellow
    gi17b.ColorFull=YellowFull
    gi17c.Color=Yellow
    gi17c.ColorFull=YellowFull
    gi17a.Intensity = YellowI
    pBulb017.Material = "BulbYellowGIOff"

    gi18a.Color=Green
    gi18a.ColorFull=GreenFull
    gi18b.Color=Green
    gi18b.ColorFull=GreenFull
    gi18c.Color=Green
    gi18c.ColorFull=GreenFull
    gi18a.Intensity = GreenI
    pBulb018.Material = "BulbGreenGIOff"
End If




''''''''''''''VR Settings''''''''''''''''

If CabVisible = 1 then
  Primary_cab_left.visible = true
  Primary_cab_right.visible = true
  Primary_cab_front.visible = true
  pDmdArea.visible = true
  fDMD.visible = true
  fDMD_reflection.Visible = true
  Primary_box_support.visible = true
  pBackBox.visible = true
  Primary_legs.visible = true
  Primary_glass_support.visible = true
  pBackGlassHolder.visible = true

  Primary_flipper_button_left.visible = true
  Primary_flipper_button_rings.visible = true
  Primary_flipper_button_right.visible = true

  Primary_PlungerBracket.visible = true
  vr_PlungerRod1.visible = true
  vr_PlungerRod2.visible = true
  pCoinDoorA.Visible = True
  pCoinDoorB.Visible = True
  Cabinet.RotX = 88
Else
  Primary_cab_left.visible = false
  Primary_cab_right.visible = false
  Primary_cab_front.visible = false
  pDmdArea.visible = false
  fDMD.visible = false
  fDMD_reflection.visible = false
  Primary_box_support.visible = false
  pBackBox.visible = false
  Primary_legs.visible = false
  Primary_glass_support.visible = false
  pBackGlassHolder.visible = false

  Primary_flipper_button_left.visible = false
  Primary_flipper_button_rings.visible = false
  Primary_flipper_button_right.visible = false

  Primary_PlungerBracket.visible = false
  vr_PlungerRod1.visible = false
  vr_PlungerRod2.visible = false
  pCoinDoorA.Visible = false
  pCoinDoorB.Visible = false
  Cabinet.RotX = 90
End If


If BackGlassVisible = 1 Then
  DisplayTimer.enabled = true
  BGHigh.visible = true
  BGHigh1.visible = true
  BGHigh2.visible = true
  BGDark.visible = true
  BGSol11.visible = true
  BGSolKong.visible = true
  BGSolGirl.visible = true
  BGSol49.visible = true
  BGSol50.visible = true
  BGSol51.visible = true
  BGSol52.visible = true
  BGSol53.visible = true
  BGSol54.visible = true
  BGSolDNSRS.visible = true
  BGSolSkull.visible = true
  BGSolKingShooter.visible = true
  BGSolKongShooter.visible = true
  BGSolTower25.visible = true
  BGSolTower50.visible = true
  BGSolTower100.visible = true
  BGSolTowerEB.visible = true
  BGSolTowerMil.visible = true
  BGLight55.visible = true
Else
  DisplayTimer.enabled = false
  BGHigh.visible = false
  BGHigh1.visible = false
  BGHigh2.visible = false
  BGDark.visible = false
  BGSol11.visible = false
  BGSolKong.visible = false
  BGSolGirl.visible = false
  BGSol49.visible = false
  BGSol50.visible = false
  BGSol51.visible = false
  BGSol52.visible = false
  BGSol53.visible = false
  BGSol54.visible = false
  BGSolDNSRS.visible = false
  BGSolSkull.visible = false
  BGSolKingShooter.visible = false
  BGSolKongShooter.visible = false
  BGSolTower25.visible = false
  BGSolTower50.visible = false
  BGSolTower100.visible = false
  BGSolTowerEB.visible = false
  BGSolTowerMil.visible = false
  BGLight55.visible = false
End If

If RailsVisible = 1 Then
  pSideRails.visible = true
Else
  pSideRails.visible = false
End If

If LockBarVisible = 1 Then
  pLockdownBar.visible = true
Else
  pLockdownBar.visible = false
End If

If WallsVisible = 1 Then
  VRWalls.visible = true
Else
  VRWalls.visible = false
End If

If FloorVisible = 1 Then
  VRfloor.visible = true
Else
  VRfloor.visible = false
End If

End Sub

'**************
' Flipper Subs
'**************

dim LF : Set LF = New FlipperPolarity 'Left Flipper
dim RF : Set RF = New FlipperPolarity 'Right Flipper
dim LFKS : Set LFKS = New FlipperPolarity 'Left Flipper King Sized
dim RFKS : Set RFKS = New FlipperPolarity 'Right Flipper King Sized
dim RFKS2 : Set RFKS2 = New FlipperPolarity 'Right Flipper 2 King Sized

InitPolarity

Sub InitPolarity()
dim x, a, b : a = Array(LF, RF)
for each x in a
'safety coefficient (diminishes polarity correction only)
  x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1
  x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
  x.enabled = True
  x.TimeDelay = 80
Next

b = Array(LFKS, RFKS)
for each x in b
'safety coefficient (diminishes polarity correction only)
  x.AddPoint "Ycoef", 0, RightFlipperKS.Y-65, 1
  x.AddPoint "Ycoef", 1, RightFlipperKS.Y-11, 1
  x.enabled = True
  x.TimeDelay = 80
Next

'rf.report "Velocity"
addpt "Velocity", 0, 0, 1
addpt "Velocity", 1, 0.2, 1.07
addpt "Velocity", 2, 0.41, 1.05
addpt "Velocity", 3, 0.44, 1
addpt "Velocity", 4, 0.65, 1.0'0.982
addpt "Velocity", 5, 0.702, 0.968
addpt "Velocity", 6, 0.95, 0.968
addpt "Velocity", 7, 1.03, 0.945

'rf.report "Polarity"
AddPt "Polarity", 0, 0, -4.7
AddPt "Polarity", 1, 0.16, -4.7
AddPt "Polarity", 2, 0.33, -4.7
AddPt "Polarity", 3, 0.37, -4.7
AddPt "Polarity", 4, 0.41, -4.7
AddPt "Polarity", 5, 0.45, -4.7
AddPt "Polarity", 6, 0.576,-4.7
AddPt "Polarity", 7, 0.66, -2.8
AddPt "Polarity", 8, 0.743, -1.5
AddPt "Polarity", 9, 0.81, -1.5
AddPt "Polarity", 10, 0.88, 0

LF.Object = LeftFlipper
LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property.
RF.Object = RightFlipper
RF.EndPoint = EndPointRp

LFKS.Object = LeftFlipperKS
LFKS.EndPoint = EndPointLKSFp  'you can use just a coordinate, or an object with a .x property.
RFKS.Object = RightFlipperKS
RFKS.EndPoint = EndPointRKSFp

RFKS2.Object = RightFlipper2
RFKS2.EndPoint = EndPointRKSF2p
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

Sub TriggerLKSF_Hit() : LFKS.Addball activeball : End Sub
Sub TriggerLKSF_UnHit() : LFKS.PolarityCorrect activeball: End Sub
Sub TriggerRKSF_Hit() : RFKS.Addball activeball : End Sub
Sub TriggerRKSF_UnHit() : RFKS.PolarityCorrect activeball: End Sub

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers),LeftFlipper
    LF.fire 'LeftFlipper.RotateToEnd
    LFKS.fire 'LeftFlipperKS.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers),LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperKS.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers),RightFlipper
    RF.fire 'RightFlipper.RotateToEnd
    RFKS.fire 'RightFlipperKS.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers),RightFlipper
        RightFlipper.RotateToStart
        RightFlipperKS.RotateToStart
    End If
End Sub

Sub SolR2Flipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers),RightFlipper2
        RightFlipper2.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers),RightFlipper2
        RightFlipper2.RotateToStart
    End If
End Sub

Sub RotateThings_timer()
    batleft.objrotz = LeftFlipper.CurrentAngle + 1
    batright.objrotz = RightFlipper.CurrentAngle - 1

    batleftKS.objrotz = LeftFlipperKS.CurrentAngle + 1
    batrightKS.objrotz = RightFlipperKS.CurrentAngle - 1

    pLeftFlipperLogo.objrotz = LeftFlipperKS.CurrentAngle + 1
    pRightFlipperLogo.objrotz = RightFlipperKS.CurrentAngle - 1
    pRightFlipper2Logo.objrotz = RightFlipper2.CurrentAngle - 1

    batleftKS2.objrotz = LeftFlipperKS.CurrentAngle + 1
    batrightKS2.objrotz = RightFlipperKS.CurrentAngle - 1

    pLSS.Roty = LeftFlipper.Currentangle - 90
    pRSS.Roty = RightFlipper.Currentangle - 90

    batright1.objrotz = RightFlipper2.CurrentAngle - 1
    batright2.objrotz = RightFlipper2.CurrentAngle - 1

    pCageDoor.RotX = CageDoor.currentAngle * -1

    pUPFGate.RotX = GateUPF.currentAngle * -1 + -30
    p_gate2.RotX = Gate2.currentAngle * -1 + -30
    pGateBR.RotX = GateBR.currentAngle * -1 + 90

    batrightshadow.objrotz = RightFlipper.CurrentAngle
    batleftshadow.objrotz = LeftFlipper.CurrentAngle
    batrightshadowKS.objrotz = RightFlipperKS.CurrentAngle
    batleftshadowKS.objrotz = LeftFlipperKS.CurrentAngle
    batrightshadow2.objrotz = RightFlipper2.CurrentAngle
End Sub


'Other code (If you have low strength, you may want to increase EOSTNew to a higher number to prevent a springy flipper when in the up position):

''''''''''Normal Flippers

RightFlipper.timerinterval=1
rightflipper.timerenabled=True

sub RightFlipper_timer()

  If leftflipper.currentangle = leftflipper.endangle and LFPress = 1 then
    leftflipper.eostorqueangle = EOSAnew
    leftflipper.eostorque = EOSTnew
    LeftFlipper.rampup = EOSRampup
    if LFCount = 0 Then LFCount = GameTime
    if GameTime - LFCount < LiveCatch Then
      leftflipper.Elasticity = 0.1
      If LeftFlipper.endangle <> LFEndAngle Then leftflipper.endangle = LFEndAngle
    Else
      leftflipper.Elasticity = FElasticity
    end if
  elseif leftflipper.currentangle > leftflipper.startangle - 0.05  Then
    leftflipper.rampup = SOSRampup
    leftflipper.endangle = LFEndAngle - 3
    leftflipper.Elasticity = FElasticity
    LFCount = 0
  elseif leftflipper.currentangle > leftflipper.endangle + 0.01 Then
    leftflipper.eostorque = EOST
    leftflipper.eostorqueangle = EOSA
    LeftFlipper.rampup = Frampup
    leftflipper.Elasticity = FElasticity
  end if

  If rightflipper.currentangle = rightflipper.endangle and RFPress = 1 then
    rightflipper.eostorqueangle = EOSAnew
    rightflipper.eostorque = EOSTnew
    RightFlipper.rampup = EOSRampup
    if RFCount = 0 Then RFCount = GameTime
    if GameTime - RFCount < LiveCatch Then
      rightflipper.Elasticity = 0.1
      If RightFlipper.endangle <> RFEndAngle Then rightflipper.endangle = RFEndAngle
    Else
      rightflipper.Elasticity = FElasticity
    end if
  elseif rightflipper.currentangle < rightflipper.startangle + 0.05 Then
    rightflipper.rampup = SOSRampup
    rightflipper.endangle = RFEndAngle + 3
    rightflipper.Elasticity = FElasticity
    RFCount = 0
  elseif rightflipper.currentangle < rightflipper.endangle - 0.01 Then
    rightflipper.eostorque = EOST
    rightflipper.eostorqueangle = EOSA
    RightFlipper.rampup = Frampup
    rightflipper.Elasticity = FElasticity
  end if

end sub

dim LFPress, RFPress, EOST, EOSA, EOSTnew, EOSAnew
dim FStrength, Frampup, FElasticity, EOSRampup, SOSRampup
dim RFEndAngle, LFEndAngle, LFCount, RFCount, LiveCatch

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
FStrength = LeftFlipper.strength
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
EOSTnew = 1.0 'FEOST
EOSAnew = 0.2
EOSRampup = 1.5
SOSRampup = 8.5
LiveCatch = 8

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

''''''''''KingSized Flippers

RightFlipperKS.timerinterval=1
rightflipperKS.timerenabled=True

sub RightFlipperKS_timer()

  If leftflipperKS.currentangle = leftflipperKS.endangle and LFPress2 = 1 then
    leftflipperKS.eostorqueangle = EOSAnew2
    leftflipperKS.eostorque = EOSTnew2
    LeftFlipperKS.rampup = EOSRampup2
    if LFCount2 = 0 Then LFCount2 = GameTime
    if GameTime - LFCount2 < LiveCatch2 Then
      leftflipperKS.Elasticity = 0.1
      If LeftFlipperKS.endangle <> LFEndAngle2 Then leftflipperKS.endangle = LFEndAngle2
    Else
      leftflipperKS.Elasticity = FElasticity2
    end if
  elseif leftflipperKS.currentangle > leftflipperKS.startangle - 0.05  Then
    leftflipperKS.rampup = SOSRampup2
    leftflipperKS.endangle = LFEndAngle2 - 3
    leftflipperKS.Elasticity = FElasticity2
    LFCount = 0
  elseif leftflipperKS.currentangle > leftflipperKS.endangle + 0.01 Then
    leftflipperKS.eostorque = EOST2
    leftflipperKS.eostorqueangle = EOSA2
    LeftFlipperKS.rampup = Frampup2
    leftflipperKS.Elasticity = FElasticity2
  end if

  If rightflipperKS.currentangle = rightflipperKS.endangle and RFPress2 = 1 then
    rightflipperKS.eostorqueangle = EOSAnew2
    rightflipperKS.eostorque = EOSTnew2
    RightFlipperKS.rampup = EOSRampup2
    if RFCount2 = 0 Then RFCount2 = GameTime
    if GameTime - RFCount2 < LiveCatch2 Then
      rightflipperKS.Elasticity = 0.1
      If RightFlipperKS.endangle <> RFEndAngle2 Then rightflipperKS.endangle = RFEndAngle2
    Else
      rightflipperKS.Elasticity = FElasticity2
    end if
  elseif rightflipperKS.currentangle < rightflipperKS.startangle + 0.05 Then
    rightflipperKS.rampup = SOSRampup2
    rightflipperKS.endangle = RFEndAngle2 + 3
    rightflipperKS.Elasticity = FElasticity2
    RFCount2 = 0
  elseif rightflipperKS.currentangle < rightflipperKS.endangle - 0.01 Then
    rightflipperKS.eostorque = EOST2
    rightflipperKS.eostorqueangle = EOSA2
    RightFlipperKS.rampup = Frampup2
    rightflipperKS.Elasticity = FElasticity2
  end if

end sub

dim LFPress2, RFPress2, EOST2, EOSA2, EOSTnew2, EOSAnew2
dim FStrength2, Frampup2, FElasticity2, EOSRampup2, SOSRampup2
dim RFEndAngle2, LFEndAngle2, LFCount2, RFCount2, LiveCatch2

EOST2 = leftflipperKS.eostorque
EOSA2 = leftflipperKS.eostorqueangle
FStrength2 = LeftFlipperKS.strength
Frampup2 = LeftFlipperKS.rampup
FElasticity2 = LeftFlipperKS.elasticity
EOSTnew2 = 1.0 'FEOST
EOSAnew2 = 0.2
EOSRampup2 = 1.5
SOSRampup2 = 8.5
LiveCatch2 = 8

LFEndAngle2 = LeftflipperKS.endangle
RFEndAngle2 = RightFlipperKS.endangle

'''''''''''''''''''''''''''''''
''''''''''Rollovers
'''''''''''''''''''''''''''''''

Sub sw14_Hit()
  Switch14dir = 1
  Sw14Move = 1
  Me.TimerEnabled = true
  Controller.Switch(14) = 1
  PlaySoundAt "rollover",sw14
End Sub

Sub sw14_unHit()
  Switch14dir = -1
  Sw14Move = 5
  Me.TimerEnabled = true
  Controller.Switch(14) = 0
End Sub


Dim Switch14dir, SW14Move

Sub sw14_timer()
Select case Sw14Move
  Case 0:me.TimerEnabled = false:pRollover9.RotX = 90
  Case 1:pRollover9.RotX = 95
  Case 2:pRollover9.RotX = 100
  Case 3:pRollover9.RotX = 105
  Case 4:pRollover9.RotX = 110
  Case 5:pRollover9.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover9.RotX = 120
End Select

SW14Move = SW14Move + Switch14dir

End Sub

Sub sw17_Hit()
  Switch17dir = 1
  Sw17Move = 1
  Me.TimerEnabled = true
  Controller.Switch(17) = 1
  PlaySoundAt "rollover",sw17
End Sub

Sub sw17_unHit()
  Switch17dir = -1
  Sw17Move = 5
  Me.TimerEnabled = true
  Controller.Switch(17) = 0
End Sub


Dim Switch17dir, SW17Move

Sub sw17_timer()
Select case Sw17Move
  Case 0:me.TimerEnabled = false:pRollover5.RotX = 90
  Case 1:pRollover5.RotX = 95
  Case 2:pRollover5.RotX = 100
  Case 3:pRollover5.RotX = 105
  Case 4:pRollover5.RotX = 110
  Case 5:pRollover5.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover5.RotX = 120
End Select

SW17Move = SW17Move + Switch17dir

End Sub

Sub sw18_Hit()
  Switch18dir = 1
  Sw18Move = 1
  Me.TimerEnabled = true
  Controller.Switch(18) = 1
  PlaySoundAt "rollover",sw18
End Sub

Sub sw18_unHit()
  Switch18dir = -1
  Sw18Move = 5
  Me.TimerEnabled = true
  Controller.Switch(18) = 0
End Sub


Dim Switch18dir, SW18Move

Sub sw18_timer()
Select case Sw18Move
  Case 0:me.TimerEnabled = false:pRollover6.RotX = 90
  Case 1:pRollover6.RotX = 95
  Case 2:pRollover6.RotX = 100
  Case 3:pRollover6.RotX = 105
  Case 4:pRollover6.RotX = 110
  Case 5:pRollover6.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover6.RotX = 120
End Select

SW18Move = SW18Move + Switch18dir

End Sub

Sub sw19_Hit()
  Switch19dir = 1
  Sw19Move = 1
  Me.TimerEnabled = true
  Controller.Switch(19) = 1
  PlaySoundAt "rollover",sw14
End Sub

Sub sw19_unHit()
  Switch19dir = -1
  Sw19Move = 5
  Me.TimerEnabled = true
  Controller.Switch(19) = 0
End Sub

Dim Switch19dir, SW19Move

Sub sw19_timer()
Select case Sw19Move
  Case 0:me.TimerEnabled = false:pRollover8.RotX = 90
  Case 1:pRollover8.RotX = 95
  Case 2:pRollover8.RotX = 100
  Case 3:pRollover8.RotX = 105
  Case 4:pRollover8.RotX = 110
  Case 5:pRollover8.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover8.RotX = 120
End Select

SW19Move = SW19Move + Switch19dir

End Sub

Sub sw20_Hit()
  Switch20dir = 1
  Sw20Move = 1
  Me.TimerEnabled = true
  Controller.Switch(20) = 1
  PlaySoundAt "rollover",sw20
End Sub

Sub sw20_unHit()
  Switch20dir = -1
  Sw20Move = 5
  Me.TimerEnabled = true
  Controller.Switch(20) = 0
End Sub


Dim Switch20dir, SW20Move

Sub sw20_timer()
Select case Sw20Move
  Case 0:me.TimerEnabled = false:pRollover7.RotX = 90
  Case 1:pRollover7.RotX = 95
  Case 2:pRollover7.RotX = 100
  Case 3:pRollover7.RotX = 105
  Case 4:pRollover7.RotX = 110
  Case 5:pRollover7.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover7.RotX = 120
End Select

SW20Move = SW20Move + Switch20dir

End Sub

Sub sw23_Hit()
  Switch23dir = 1
  Sw23Move = 1
  Me.TimerEnabled = true
  Controller.Switch(23) = 1
  PlaySoundAt "rollover",sw23
End Sub

Sub sw23_unHit()
  Switch23dir = -1
  Sw23Move = 5
  Me.TimerEnabled = true
  Controller.Switch(23) = 0
End Sub


Dim Switch23dir, SW23Move

Sub sw23_timer()
Select case Sw23Move
  Case 0:me.TimerEnabled = false:pRollover10.RotX = 90
  Case 1:pRollover10.RotX = 95
  Case 2:pRollover10.RotX = 100
  Case 3:pRollover10.RotX = 105
  Case 4:pRollover10.RotX = 110
  Case 5:pRollover10.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover10.RotX = 120
End Select

SW23Move = SW23Move + Switch23dir

End Sub

Sub sw25_Hit()
  Switch25dir = 1
  Sw25Move = 1
  Me.TimerEnabled = true
  Controller.Switch(25) = 1
  PlaySoundAt "rollover",sw25
End Sub

Sub sw25_unHit()
  Switch25dir = -1
  Sw25Move = 5
  Me.TimerEnabled = true
  Controller.Switch(25) = 0
End Sub


Dim Switch25dir, SW25Move

Sub sw25_timer()
Select case Sw25Move
  Case 0:me.TimerEnabled = false:pRollover2.RotX = 90
  Case 1:pRollover2.RotX = 95
  Case 2:pRollover2.RotX = 100
  Case 3:pRollover2.RotX = 105
  Case 4:pRollover2.RotX = 110
  Case 5:pRollover2.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover2.RotX = 120
End Select

SW25Move = SW25Move + Switch25dir

End Sub


Sub sw26_Hit()
  Switch26dir = 1
  Sw26Move = 1
  Me.TimerEnabled = true
  Controller.Switch(26) = 1
  PlaySoundAt "rollover",sw26
End Sub

Sub sw26_unHit()
  Switch26dir = -1
  Sw26Move = 5
  Me.TimerEnabled = true
  Controller.Switch(26) = 0
End Sub


Dim Switch26dir, SW26Move

Sub sw26_timer()
Select case Sw26Move
  Case 0:me.TimerEnabled = false:pRollover3.RotX = 90
  Case 1:pRollover3.RotX = 95
  Case 2:pRollover3.RotX = 100
  Case 3:pRollover3.RotX = 105
  Case 4:pRollover3.RotX = 110
  Case 5:pRollover3.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover3.RotX = 120
End Select

SW26Move = SW26Move + Switch26dir

End Sub

Sub sw27_Hit()
  Switch27dir = 1
  Sw27Move = 1
  Me.TimerEnabled = true
  Controller.Switch(27) = 1
  PlaySoundAt "rollover",sw27
End Sub

Sub sw27_unHit()
  Switch27dir = -1
  Sw27Move = 5
  Me.TimerEnabled = true
  Controller.Switch(27) = 0
End Sub


Dim Switch27dir, SW27Move

Sub sw27_timer()
Select case Sw27Move
  Case 0:me.TimerEnabled = false:pRollover4.RotX = 90
  Case 1:pRollover4.RotX = 95
  Case 2:pRollover4.RotX = 100
  Case 3:pRollover4.RotX = 105
  Case 4:pRollover4.RotX = 110
  Case 5:pRollover4.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover4.RotX = 120
End Select

SW27Move = SW27Move + Switch27dir

End Sub


Sub sw28_Hit()
  Switch28dir = 1
  Sw28Move = 1
  Me.TimerEnabled = true
  Controller.Switch(28) = 1
  PlaySoundAt "rollover",sw28
End Sub

Sub sw28_unHit()
  Switch28dir = -1
  Sw28Move = 5
  Me.TimerEnabled = true
  Controller.Switch(28) = 0
End Sub


Dim Switch28dir, SW28Move

Sub sw28_timer()
Select case Sw28Move
  Case 0:me.TimerEnabled = false:pRollover1.RotX = 90
  Case 1:pRollover1.RotX = 95
  Case 2:pRollover1.RotX = 100
  Case 3:pRollover1.RotX = 105
  Case 4:pRollover1.RotX = 110
  Case 5:pRollover1.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover1.RotX = 120
End Select

SW28Move = SW28Move + Switch28dir

End Sub


Sub sw31_Hit()
  Switch31dir = 1
  Sw31Move = 1
  Me.TimerEnabled = true
  Controller.Switch(31) = 1
  PlaySoundAt "rollover",sw31
End Sub

Sub sw31_unHit()
  Switch31dir = -1
  Sw31Move = 5
  Me.TimerEnabled = true
  Controller.Switch(31) = 0
End Sub


Dim Switch31dir, SW31Move

Sub sw31_timer()
Select case Sw31Move
  Case 0:me.TimerEnabled = false:pRollover11.RotX = 90
  Case 1:pRollover11.RotX = 95
  Case 2:pRollover11.RotX = 100
  Case 3:pRollover11.RotX = 105
  Case 4:pRollover11.RotX = 110
  Case 5:pRollover11.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover11.RotX = 120
End Select

SW31Move = SW31Move + Switch31dir

End Sub


Sub sw32_Hit()
  Switch32dir = 1
  Sw32Move = 1
  Me.TimerEnabled = true
  Controller.Switch(32) = 1
  PlaySoundAt "rollover",sw32
End Sub

Sub sw32_unHit()
  Switch32dir = -1
  Sw32Move = 5
  Me.TimerEnabled = true
  Controller.Switch(32) = 0
End Sub


Dim Switch32dir, SW32Move

Sub sw32_timer()
Select case Sw32Move
  Case 0:me.TimerEnabled = false:pRollover12.RotX = 90
  Case 1:pRollover12.RotX = 95
  Case 2:pRollover12.RotX = 100
  Case 3:pRollover12.RotX = 105
  Case 4:pRollover12.RotX = 110
  Case 5:pRollover12.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover12.RotX = 120
End Select

SW32Move = SW32Move + Switch32dir

End Sub



''Targets


Dim Target33Step, Target34Step, Target35Step, Target36Step, Target37Step, Target41Step, Target42Step, Target43Step, Target49Step, Target50Step, Target51Step


''Tower

Sub sw33_Hit:vpmTimer.PulseSw(33):psw33.RotZ = 5:psw33t.RotZ = 5:Target33Step = 0:sw33.TimerEnabled = True:PlaySoundAt SoundFX("target",DOFTargets),psw33:End Sub
Sub sw33_timer()
  Select Case Target33Step
    Case 1:psw33.RotZ = 3:psw33t.RotZ = 3
        Case 2:psw33.RotZ = -2:psw33t.RotZ = -2
        Case 3:psw33.RotZ = 1:psw33t.RotZ = 1
        Case 4:psw33.RotZ = 0:psw33t.RotZ = 0:sw33.TimerEnabled = False:Target33Step = 0
     End Select
  Target33Step = Target33Step + 1
End Sub

Sub sw34_Hit:vpmTimer.PulseSw(34):psw34.RotZ = 5:psw34t.RotZ = 5:Target34Step = 0:sw34.TimerEnabled = True:PlaySoundAt SoundFX("target",DOFTargets),psw34:End Sub
Sub sw34_timer()
  Select Case Target34Step
    Case 1:psw34.RotZ = 3:psw34t.RotZ = 3
        Case 2:psw34.RotZ = -2:psw34t.RotZ = -2
        Case 3:psw34.RotZ = 1:psw34t.RotZ = 1
        Case 4:psw34.RotZ = 0:psw34t.RotZ = 0:sw34.TimerEnabled = False:Target34Step = 0
     End Select
  Target34Step = Target34Step + 1
End Sub

Sub sw35_Hit:vpmTimer.PulseSw(35):psw35.RotZ = 5:psw35t.RotZ = 5:Target35Step = 0:sw35.TimerEnabled = True:PlaySoundAt SoundFX("target",DOFTargets),psw35:End Sub
Sub sw35_timer()
  Select Case Target35Step
    Case 1:psw35.RotZ = 3:psw35t.RotZ = 3
        Case 2:psw35.RotZ = -2:psw35t.RotZ = -2
        Case 3:psw35.RotZ = 1:psw35t.RotZ = 1
        Case 4:psw35.RotZ = 0:psw35t.RotZ = 0:sw35.TimerEnabled = False:Target35Step = 0
     End Select
  Target35Step = Target35Step + 1
End Sub

Sub sw36_Hit:vpmTimer.PulseSw(36):psw36.RotZ = 5:psw36t.RotZ = 5:Target36Step = 0:sw36.TimerEnabled = True:PlaySoundAt SoundFX("target",DOFTargets),psw36:End Sub
Sub sw36_timer()
  Select Case Target36Step
    Case 1:psw36.RotZ = 3:psw36t.RotZ = 3
        Case 2:psw36.RotZ = -2:psw36t.RotZ = -2
        Case 3:psw36.RotZ = 1:psw36t.RotZ = 1
        Case 4:psw36.RotZ = 0:psw36t.RotZ = 0:sw36.TimerEnabled = False:Target36Step = 0
     End Select
  Target36Step = Target36Step + 1
End Sub

Sub sw37_Hit:vpmTimer.PulseSw(37):psw37.RotZ = 5:psw37t.RotZ = 5:Target37Step = 0:sw37.TimerEnabled = True:PlaySoundAt SoundFX("target",DOFTargets),psw37:End Sub
Sub sw37_timer()
  Select Case Target37Step
    Case 1:psw37.RotZ = 3:psw37t.RotZ = 3
        Case 2:psw37.RotZ = -2:psw37t.RotZ = -2
        Case 3:psw37.RotZ = 1:psw37t.RotZ = 1
        Case 4:psw37.RotZ = 0:psw37t.RotZ = 0:sw37.TimerEnabled = False:Target37Step = 0
     End Select
  Target37Step = Target37Step + 1
End Sub

''Dino

Sub sw41_Hit:vpmTimer.PulseSw(41):psw41.RotZ = 5:Target41Step = 0:sw41.TimerEnabled = True:PlaySoundAt SoundFX("target",DOFTargets),psw41:End Sub
Sub sw41_timer()
  Select Case Target41Step
    Case 1:psw41.RotZ = 3
        Case 2:psw41.RotZ = -2
        Case 3:psw41.RotZ = 1
        Case 4:psw41.RotZ = 0:sw41.TimerEnabled = False:Target41Step = 0
     End Select
  Target41Step = Target41Step + 1
End Sub

Sub sw42_Hit:vpmTimer.PulseSw(42):psw42.RotZ = 5:Target42Step = 0:sw42.TimerEnabled = True:PlaySoundAt SoundFX("target",DOFTargets),psw42:End Sub
Sub sw42_timer()
  Select Case Target42Step
    Case 1:psw42.RotZ = 3
        Case 2:psw42.RotZ = -2
        Case 3:psw42.RotZ = 1
        Case 4:psw42.RotZ = 0:sw42.TimerEnabled = False:Target42Step = 0
     End Select
  Target42Step = Target42Step + 1
End Sub

Sub sw43_Hit:vpmTimer.PulseSw(43):psw43.RotZ = 5:Target43Step = 0:sw43.TimerEnabled = True:PlaySoundAt SoundFX("target",DOFTargets),psw43:End Sub
Sub sw43_timer()
  Select Case Target43Step
    Case 1:psw43.RotZ = 3
        Case 2:psw43.RotZ = -2
        Case 3:psw43.RotZ = 1
        Case 4:psw43.RotZ = 0:sw43.TimerEnabled = False:Target43Step = 0
     End Select
  Target43Step = Target43Step + 1
End Sub


'''''''''''Mystery target by Million Bananas lamp, seen in one of protorypes (Kong Ramp...)  No idea on switch number

Sub sw44_Hit:vpmTimer.PulseSw(44):PlaySoundAt SoundFX("target",DOFTargets),sw44:End Sub
Sub sw44_timer()
  Select Case Target44Step
    Case 1:psw44.RotY = -3
        Case 2:psw44.RotY = 2
        Case 3:psw44.RotY = -1
        Case 4:psw44.RotY = 0:sw44.TimerEnabled = False:Target44Step = 0
     End Select
  Target44Step = Target44Step + 1
End Sub


''Teams

Sub sw49_Hit:vpmTimer.PulseSw(49):psw49.RotZ = 5:Target49Step = 0:sw49.TimerEnabled = True:PlaySoundAt SoundFX("standup_hit",DOFTargets),psw49:End Sub
Sub sw49_timer()
  Select Case Target49Step
    Case 1:psw49.RotZ = 3
        Case 2:psw49.RotZ = -2
        Case 3:psw49.RotZ = 1
        Case 4:psw49.RotZ = 0:sw49.TimerEnabled = False:Target49Step = 0
     End Select
  Target49Step = Target49Step + 1
End Sub

Sub sw50_Hit:vpmTimer.PulseSw(50):psw50.RotZ = 5:Target50Step = 0:sw50.TimerEnabled = True:PlaySoundAt SoundFX("standup_hit",DOFTargets),psw50:End Sub
Sub sw50_timer()
  Select Case Target50Step
    Case 1:psw50.RotZ = 3
        Case 2:psw50.RotZ = -2
        Case 3:psw50.RotZ = 1
        Case 4:psw50.RotZ = 0:sw50.TimerEnabled = False:Target50Step = 0
     End Select
  Target50Step = Target50Step + 1
End Sub

Sub sw51_Hit:vpmTimer.PulseSw(51):psw51.RotZ = 5:Target51Step = 0:sw51.TimerEnabled = True:PlaySoundAt SoundFX("standup_hit",DOFTargets),psw51:End Sub
Sub sw51_timer()
  Select Case Target51Step
    Case 1:psw51.RotZ = 3
        Case 2:psw51.RotZ = -2
        Case 3:psw51.RotZ = 1
        Case 4:psw51.RotZ = 0:sw51.TimerEnabled = False:Target51Step = 0
     End Select
  Target51Step = Target51Step + 1
End Sub


'***Slings and rubbers
  ' Slings
 Dim LSStep, RSStep

 Sub SlingLeft_Slingshot
  PlaySoundAt SoundFX("LeftSlingShot",DOFContactors),pSlingL
    LeftSlingA.Visible = False:LeftSlingB.Visible = True
    pSlingL.TransZ = -10:pSlingKongLeft.ObjRotY = 2:pSlingKongLBracket.ObjRotY = -2:pSlingKongProtoLeft.ObjRotY = -2:pSlingKongLProtoBracket.ObjRotY = -2
    LSStep = 0
    SlingLeft.TimerEnabled = True
    vpmTimer.PulseSw 21
End Sub


Sub SlingLeft_Timer
    Select Case LSStep
        Case 0:LeftSlingB.Visible = false:LeftSlingC.Visible = True:pSlingL.TransZ = -18:pSlingKongLeft.ObjRotY = 0
        Case 1:LeftSlingC.Visible = false:LeftSlingD.Visible = True:pSlingL.TransZ = -27:pSlingKongLeft.ObjRotY = 2
        Case 2:LeftSlingD.Visible = false:LeftSlingC.Visible = True:pSlingL.TransZ = -18:pSlingKongLeft.ObjRotY = 5
        Case 3:LeftSlingC.Visible = false:LeftSlingB.Visible = True:pSlingL.TransZ = -10:pSlingKongLeft.ObjRotY = 0
        Case 4:LeftSlingB.Visible = false:LeftSlingA.Visible = True:pSlingL.TransZ = 0:pSlingKongLeft.ObjRotY = -3:Me.TimerEnabled = 0 '
    End Select

pSlingKongLBracket.ObjRotY = pSlingKongLeft.ObjRotY
pSlingKongProtoLeft.ObjRotY = pSlingKongLeft.ObjRotY
pSlingKongLProtoBracket.ObjRotY = pSlingKongLeft.ObjRotY


    LSStep = LSStep + 1
End Sub


Sub SlingRight_Slingshot()
  PlaySoundAt SoundFX("RightSlingShot",DOFContactors),pSlingR
    RightSlingA.Visible = False:RightSlingB.Visible = True
    pSlingR.TransZ = -10:SlingKongRight.ObjRotY = 2:pSlingRBracket.ObjRotY = 2:SlingKongProtoRight.ObjRotY = 2:pSlingRProtoBracket.ObjRotY = 2
    RSStep = 0
    SlingRight.TimerEnabled = True
    vpmTimer.PulseSw 22
End Sub

Sub SlingRight_Timer
    Select Case RSStep

        Case 0:RightSlingB.Visible = false:RightSlingC.Visible = True:pSlingR.TransZ = -18:SlingKongRight.ObjRotY = 0
        Case 1:RightSlingC.Visible = false:RightSlingD.Visible = True:pSlingR.TransZ = -27:SlingKongRight.ObjRotY = -2
        Case 2:RightSlingD.Visible = false:RightSlingC.Visible = True:pSlingR.TransZ = -18:SlingKongRight.ObjRotY = -5
        Case 3:RightSlingC.Visible = false:RightSlingB.Visible = True:pSlingR.TransZ = -10:SlingKongRight.ObjRotY = 0
        Case 4:RightSlingB.Visible = false:RightSlingA.Visible = True::pSlingR.TransZ = 0:SlingKongRight.ObjRotY = 3:Me.TimerEnabled = 0 '
    End Select

pSlingRBracket.ObjRotY = SlingKongRight.ObjRotY
SlingKongProtoRight.ObjRotY = SlingKongRight.ObjRotY
pSlingRProtoBracket.ObjRotY = SlingKongRight.ObjRotY

    RSStep = RSStep + 1
End Sub


'************************************
'''''''' Bumpers
'************************************

Sub Bumper1_hit:vpmTimer.PulseSw(48):PlaySoundAtBumperVol SoundFX("bumper1",DOFContactors),Bumper1,1:End Sub

Sub Bumper2_hit:vpmTimer.PulseSw(47):PlaySoundAtBumperVol SoundFX("bumper2",DOFContactors),Bumper2,1:End Sub

Sub Bumper3_hit:vpmTimer.PulseSw(46):PlaySoundAtBumperVol SoundFX("bumper3",DOFContactors),Bumper3,1:End Sub



''''King Sized Ramp

'sw29
Sub sw29ksr_Hit()
  sw29ksrdir = 1
  sw29ksrMove = 1
  Me.TimerEnabled = true
  Controller.Switch(29) = 1
  PlaySoundAt "rollover",KSRswitchArm
End Sub

Sub sw29ksr_unHit()
  sw29ksrdir = -1
  sw29ksrMove = 4
  Me.TimerEnabled = true
  Controller.Switch(29) = 0
End Sub


Dim sw29ksrdir, sw29ksrMove

Sub sw29ksr_timer()
Select case sw29ksrMove
  Case 0:me.TimerEnabled = false:KSRswitchArm.RotY = -25
  Case 1:KSRswitchArm.RotY = -20
  Case 2:KSRswitchArm.RotY = -15
  Case 3:KSRswitchArm.RotY = -10
  Case 4:KSRswitchArm.RotY = -5
  Case 5:me.TimerEnabled = false:KSRswitchArm.RotY = -0
End Select

sw29ksrMove = sw29ksrMove + sw29ksrdir

End Sub

''''''''''''''''''''''''''''''''''''''''
'VUK MOD Extras
''''''''''''''''''''''''''''''''''''''''

'VUK Lock
Sub sw29_Hit()
  If VUKModType = 1 then
    PlaySoundat "kicker_enter",pUpKicker
  End If
End Sub

Sub KickBallUp(Enabled)
  If VUKModType Then
    sw29.timerenabled = 1
  End If
End Sub

Dim sw29step

Sub sw29_timer()
  Select Case sw29step
    Case 0:pUpKicker.TransY = 10:Playsoundat SoundFX("Solenoid",DOFContactors),sw29
    Case 1:pUpKicker.TransY = 20:If BallInKicker1 = 1 then BallSaucer1.x = 181:BallSaucer1.y = 371:BallSaucer1.vely = 0:BallSaucer1.velx = 0:BallSaucer1.velz = 60 End If
    Case 2:pUpKicker.TransY = 30:
    Case 3:
    Case 4:
    Case 5:pUpKicker.TransY = 25
    Case 6:pUpKicker.TransY = 20
    Case 7:pUpKicker.TransY = 15
    Case 8:pUpKicker.TransY = 10:Controller.Switch(29) = 0
    Case 9:pUpKicker.TransY = 5
    Case 10:pUpKicker.TransY = 0:sw29.timerEnabled = 0:sw29step = 0
  End Select
  sw29step = sw29step + 1
End Sub


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'   Drop Targets
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

  dim sw41Dir, sw42Dir, sw43Dir
  dim sw41Pos, sw42Pos, sw43Pos
  Dim sw41step, sw42step, sw43step

  sw41Dir = 1:sw42Dir = 1:sw43Dir = 1
  sw41Pos = 0:sw42Pos = 0:sw43Pos = 0

  'Targets Init
  sw41.TimerEnabled = 1:sw42.timerEnabled = 1:sw43.TimerEnabled = 1

'Sub DoubleDrop1_HIt:sw41.timerenabled = True:sw42.timerenabled = True: End Sub
'Sub DoubleDrop2_HIt:sw42.timerenabled = True:sw43.timerenabled = True: End Sub

  Sub sw41vuk_Hit
    me.timerEnabled = 1
    sw41vuk.Collidable = false
    SW41vuk001.IsDropped = True
    sw41vuk001.Collidable = false
  End Sub
  Sub sw42vuk_Hit
    me.timerEnabled = 1
    sw42vuk.Collidable = false
    SW42vuk001.IsDropped = True
    sw42vuk001.Collidable = false
  End Sub

  Sub sw43vuk_Hit
    me.timerEnabled = 1
    sw43vuk.Collidable = false
    SW43vuk001.IsDropped = True
    sw43vuk001.Collidable = false
  End Sub


Sub sw41vuk_timer()
  Select Case sw41step
    Case 0:
    Case 1:sw41P.RotX = 2
    Case 2:sw41P.RotX = 5
    Case 3:DTBank.Hit 1:sw41Dir = 0:sw41vuka.Enabled = 1
    Case 4:sw41P.RotX = 3
    Case 5:sw41P.RotX = 0:me.timerEnabled = 0:sw41step = 0
  End Select
  sw41step = sw41step + 1
End Sub

'''Target animation

 Sub sw41vuka_Timer()
  Select Case sw41Pos
        Case 0: sw41P.TransZ=0
         If sw41Dir = 1 then
          sw41vuka.Enabled = 0
         else
           end if
        Case 1: sw41P.TransZ=0
        Case 2: sw41P.TransZ=-6
        Case 3: sw41P.TransZ=-8
        Case 4: sw41P.TransZ=-18
        Case 5: sw41P.TransZ=-24
        Case 6: sw41P.TransZ=-30
        Case 7: sw41P.TransZ=-36
        Case 8: sw41P.TransZ=-42
        Case 9: sw41P.TransZ=-48
        Case 10: sw41P.TransZ=-52
         If sw41Dir = 1 then
         else
          sw41vuka.Enabled = 0
           end if


End Select
  If sw41Dir = 1 then
    If sw41pos>0 then sw41pos=sw41pos-2
  else
    If sw41pos<10 then sw41pos=sw41pos+2
  end if
  End Sub


Sub sw42vuk_timer()
  Select Case sw42step
    Case 0:
    Case 1:sw42P.RotX = 2
    Case 2:sw42P.RotX = 5
    Case 3:DTBank.Hit 2:sw42Dir = 0:sw42vuka.Enabled = 1
    Case 4:sw42P.RotX = 3
    Case 5:sw42P.RotX = 0:me.timerEnabled = 0:sw42step = 0
  End Select
  sw42step = sw42step + 1
End Sub


 Sub sw42vuka_Timer()
  Select Case sw42Pos
        Case 0: sw42P.TransZ=0
         If sw42Dir = 1 then
          sw42vuka.Enabled = 0
         else
           end if
        Case 1: sw42P.TransZ=0
        Case 2: sw42P.TransZ=-6
        Case 3: sw42P.TransZ=-12
        Case 4: sw42P.TransZ=-18
        Case 5: sw42P.TransZ=-24
        Case 6: sw42P.TransZ=-30
        Case 7: sw42P.TransZ=-36
        Case 8: sw42P.TransZ=-42
        Case 9: sw42P.TransZ=-48
        Case 10: sw42P.TransZ=-52
         If sw42Dir = 1 then
         else
          sw42vuka.Enabled = 0
           end if


End Select
  If sw42Dir = 1 then
    If sw42pos>0 then sw42pos=sw42pos-2
  else
    If sw42pos<10 then sw42pos=sw42pos+2
  end if
  End Sub

Sub sw43vuk_timer()
  Select Case sw43step
    Case 0:
    Case 1:sw43P.RotX = 2
    Case 2:sw43P.RotX = 5
    Case 3:DTBank.Hit 3:sw43Dir = 0:sw43vuka.Enabled = 1
    Case 4:sw43P.RotX = 3
    Case 5:sw43P.RotX = 0:me.timerEnabled = 0:sw43step = 0
  End Select
  sw43step = sw43step + 1
End Sub


Sub sw43vuka_Timer()
  Select Case sw43Pos
        Case 0: sw43P.TransZ=0
         If sw43Dir = 1 then
          sw43vuka.Enabled = 0
         else
           end if
        Case 1: sw43P.TransZ=0
        Case 2: sw43P.TransZ=-6
        Case 3: sw43P.TransZ=-12
        Case 4: sw43P.TransZ=-18
        Case 5: sw43P.TransZ=-24
        Case 6: sw43P.TransZ=-30
        Case 7: sw43P.TransZ=-36
        Case 8: sw43P.TransZ=-42
        Case 9: sw43P.TransZ=-48
        Case 10: sw43P.TransZ=-52
         If sw43Dir = 1 then
         else
          sw43vuka.Enabled = 0
           end if
  End Select
  If sw43Dir = 1 then
    If sw43pos>0 then sw43pos=sw43pos-2
  else
    If sw43pos<10 then sw43pos=sw43pos+2
  end if
End Sub


Sub RDampen_Timer()
Cor.Update
End Sub



'DT Subs
   Sub ResetDrops(Enabled)
    If Enabled Then
    If DTModType = 1 Then
      sw41Dir = 1:sw42Dir = 1:sw43Dir = 1
      SW41vuk.Collidable = True:SW42vuk.Collidable = True:SW43vuk.Collidable = True
      SW41vuk001.IsDropped = False:SW42vuk001.IsDropped = False:SW43vuk001.IsDropped = False
      SW41vuk001.Collidable = True:SW42vuk001.Collidable = True:SW43vuk001.Collidable = True
      sw41vuka.Enabled = 1:sw42vuka.Enabled = 1:sw43vuka.Enabled = 1
      DTBank.DropSol_On
    End if
    End If
   End Sub

''''''''''''''''''''''''''''''''''''''''
'End VUK MOD Extras
''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''
''''Color Ramp Ball Lock
'''''''''''''''''''''''''''''

Dim CRBLStep, WPStep, WP2Step, PlasticWobbling

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
' Sub sw40_Hit()
'   Psw40.rotY = 20
'   Controller.Switch(40) = 1
' End Sub
Sub sw40_UnHit:Psw40.rotY = 0:Controller.Switch(40) = 0:End Sub

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
' Sub sw39_Hit:Psw39.rotY = 20:Controller.Switch(39) = 1:End Sub
Sub sw39_UnHit:Psw39.rotY = 0:Controller.Switch(39) = 0:End Sub


' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
' Sub sw38_Hit()
'   Psw38.rotY = 20
'   Controller.Switch(38) = 1
'   PlaySoundAt "metalhit_medium",sw38
' End Sub
Sub sw38_UnHit:Psw38.rotY = 0:Controller.Switch(38) = 0:End Sub

Sub sw38_Hit()
  sw38dir = 1
  sw38Move = 1
  Me.TimerEnabled = true
  Controller.Switch(38) = 1
  PlaySoundAt "rollover",psw38
End Sub

Sub sw38_unHit()
  sw38dir = -1
  sw38Move = 4
  Me.TimerEnabled = true
  Controller.Switch(38) = 0
End Sub


Dim sw38dir, sw38Move

Sub sw38_timer()
Select case sw38Move
  Case 0:me.TimerEnabled = false:psw38.RotY = -25
  Case 1:psw38.RotY = -20
  Case 2:psw38.RotY = -15
  Case 3:psw38.RotY = -10
  Case 4:psw38.RotY = -5
  Case 5:me.TimerEnabled = false:psw38.RotY = -0
End Select

sw38Move = sw38Move + sw38dir

End Sub

Sub sw39_Hit()
  sw39dir = 1
  sw39Move = 1
  Me.TimerEnabled = true
  Controller.Switch(39) = 1
  PlaySoundAt "rollover",psw39
End Sub

Sub sw39_unHit()
  sw39dir = -1
  sw39Move = 4
  Me.TimerEnabled = true
  Controller.Switch(39) = 0
End Sub


Dim sw39dir, sw39Move

Sub sw39_timer()
Select case sw39Move
  Case 0:me.TimerEnabled = false:psw39.RotY = -25
  Case 1:psw39.RotY = -20
  Case 2:psw39.RotY = -15
  Case 3:psw39.RotY = -10
  Case 4:psw39.RotY = -5
  Case 5:me.TimerEnabled = false:psw39.RotY = -0
End Select

sw39Move = sw39Move + sw39dir

End Sub

Sub sw40_Hit()
  sw40dir = 1
  sw40Move = 1
  Me.TimerEnabled = true
  Controller.Switch(40) = 1
  PlaySoundAt "rollover",psw40
End Sub

Sub sw40_unHit()
  sw40dir = -1
  sw40Move = 4
  Me.TimerEnabled = true
  Controller.Switch(40) = 0
End Sub


Dim sw40dir, sw40Move

Sub sw40_timer()
Select case sw40Move
  Case 0:me.TimerEnabled = false:psw40.RotY = -25
  Case 1:psw40.RotY = -20
  Case 2:psw40.RotY = -15
  Case 3:psw40.RotY = -10
  Case 4:psw40.RotY = -5
  Case 5:me.TimerEnabled = false:psw40.RotY = -0
End Select

sw40Move = sw40Move + sw40dir

End Sub

Sub CRBallLock(Enabled)
    PlaySound SoundFX("fx_Rudysol1",DOFContactors),0,1,-.3
    pCRLock.collidable = False
    pCRLock.RotZ = 50
    CRBallLockTimer.Enabled = True
End Sub


Sub CRBallLockTimer_Timer()
  Select Case CRBLStep
    Case 0:
    Case 1:
    Case 2:
    Case 3:
    Case 4:
    Case 5:pCRLock.collidable = true:pCRLock.RotZ = 0:CRBallLockTimer.Enabled = false:CRBLStep = 0
  End Select
  CRBLStep = CRBLStep + 1

End Sub


Sub sw30_hit()
  set MissleBall = activeball
  Controller.Switch(30) = 1
End Sub

Sub sw30_unhit()
  MissleBall = Empty
  Controller.Switch(30) = 0
End Sub

Sub sw30a_hit()
  sw30dir = 1
  sw30Move = 1
  set MissleBall = activeball
' Controller.Switch(30) = 1
  me.TimerEnabled = true
End Sub

Sub sw30a_unhit()
  sw30dir = -1
  sw30Move = 4
' MissleBall = Empty
' Controller.Switch(30) = 0
  me.TimerEnabled = true
End Sub


Dim sw30dir, sw30Move

Sub sw30a_timer()
Select case sw30Move
  Case 0:me.TimerEnabled = false:MKSwitchArm.RotY = -25
  Case 1:MKSwitchArm.RotY = -20
  Case 2:MKSwitchArm.RotY = -15
  Case 3:MKSwitchArm.RotY = -10
  Case 4:MKSwitchArm.RotY = -5
  Case 5:me.TimerEnabled = false:MKSwitchArm.RotY = -0
End Select

sw30Move = sw30Move + sw30dir

End Sub


Dim MissleBall


Sub MissileKick(Enabled)
  If Enabled = True Then
    missilekicker.timerenabled = True
    PlaySound SoundFX("Solenoid",DOFContactors),0,1,-.3
  If Not isEmpty(MissleBall) then
    KickBall MissleBall, 90, 75, 0, 0
  End If
  End If
End Sub

Dim RRStep

Sub missilekicker_Timer()
  Select Case RRStep
    Case 0:
    Case 1:pRocketRod.TransX = 5
    Case 2:pRocketRod.TransX = 10
    Case 3:pRocketRod.TransX = 15
    Case 4:pRocketRod.TransX = 20
    Case 5:pRocketRod.TransX = 10:
    Case 6:pRocketRod.TransX = 0:missilekicker.timerenabled = false:RRStep = 0
  End Select
  RRStep = RRStep + 1

End Sub


Dim KickerBall

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub


'''''''''''''''Radar kicker


Dim SW45Step, sw45ark, BallInKicker2, BallSaucer2

Sub sw45_Hit:Playsoundat ("kicker_enter"),sw45:Controller.Switch(45) = 1:set KickerBall = activeball:End Sub

Sub sw45_unHit:Controller.Switch(45) = 0:KickerBall = Empty:End Sub

Sub RadarKick(enabled)
    If Enabled = True then
  If Not isEmpty(KickerBall) then
    KickBall KickerBall, 90, 9, 5, 30
    PlaySound SoundFX("Solenoid",DOFContactors),0,1,-.3
  End If
    End If
End Sub

'''''''''''''''''Cage

Sub CageKicker_Hit()
  PlaySoundat "fx_Drain",CageKicker
  cagekicker.kick 0, 10
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''  Ball Through system''''''''''''''''''''''''''
'''''''''''''''''''''by cyberpez''''''''''''''''''''''''''''''''
''''''''''''''''based off of EalaDubhSidhe's''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim BallCount
Dim cBall1, cBall2, cBall3

dim bstatus

Sub CreatBalls()
  Controller.Switch(11) = 1
  Controller.Switch(12) = 1
  Controller.Switch(13) = 1
  Set cBall1 = Kicker1.CreateSizedballWithMass(BallSize,Ballmass)
  Set cBall2 = Kicker2.CreateSizedballWithMass(BallSize,Ballmass)
  Set cBall3 = Kicker3.CreateSizedballWithMass(BallSize,Ballmass)


  If BallMod = 1 Then
    cBall1.Image = "PinballLaserLemon"
    cBall2.Image = "PinballOutrageousOrange"
    cBall3.Image = "PinballRadicalRed"
  End If


  If BallMod = 2 Then
    cBall1.Image = "Chrome_Ball_29"
    cBall1.FrontDecal = "MarbleBallRYO1"
    cBall2.Image = "Chrome_Ball_29"
    cBall2.FrontDecal = "MarbleBallRYO2"
    cBall3.Image = "Chrome_Ball_29"
    cBall3.FrontDecal = "MarbleBallRYO3"
  End If
End Sub


Sub Kicker3_Hit():Controller.Switch(11) = 1:UpdateTrough:End Sub
Sub Kicker3_UnHit():Controller.Switch(11) = 0:UpdateTrough:End Sub
Sub Kicker2_Hit():Controller.Switch(12) = 1:UpdateTrough:End Sub
Sub Kicker2_UnHit():Controller.Switch(12) = 0:UpdateTrough:End Sub
Sub Kicker1_Hit():Controller.Switch(13) = 1:UpdateTrough:End Sub
Sub Kicker1_UnHit():Controller.Switch(13) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  CheckBallStatus.Interval = 300
  CheckBallStatus.Enabled = 1
End Sub

Sub CheckBallStatus_timer()
  If Kicker1.BallCntOver = 0 Then Kicker2.kick 60, 12
  If Kicker2.BallCntOver = 0 Then Kicker3.kick 60, 9
  Me.Enabled = 0
End Sub

Dim Kicker1active, Kicker2active, Kicker3active, Kicker4active, Kicker5active, Kicker6active

dim DontKickAnyMoreBalls


Sub KickBallToLane(Enabled)
  If DontKickAnyMoreBalls = 0 then
    Kicker1.Kick 70,12
    bstatus = 2
    Kicker1active = 0
    iBall = iBall - 1
    fgBall = false
    UpperGIon = 1
    Controller.Switch(13)=0
    DontKickAnyMoreBalls = 1
    DKTMstep = 1
    DontKickToMany.enabled = true
    BallsInPlay = BallsInPlay + 1
  End If
End Sub


Dim DKTMstep

Sub DontKickToMany_timer ()
  Select Case DKTMstep
  Case 1:
  Case 2:
  Case 3: DontKickAnyMoreBalls = 0:DontKickToMany.Enabled = False: DontKickAnyMoreBalls = 0
  End Select
  DKTMstep = DKTMstep + 1
End Sub


sub kisort(enabled)
  if fgBall then
    Drain.Kick 70,20
    iBall = iBall + 1
    fgBall = false

  end if

end sub


Sub Drain_hit()
  controller.switch(10) = true
  fgBall = true
  iBall = iBall + 1
  BallsInPlay = BallsInPlay - 1
End Sub

Sub Drain_UnHit()
  Controller.Switch(10) = 0
End Sub


'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1       'Do Not Change - default setting
BCvel = 4       'Controls the speed of the ball movement
BCyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
  Set ControlActiveBall = ActiveBall
  ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
  ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
  If EnableBallControl and ControlBallInPlay then
    If BCright = 1 Then
      ControlActiveBall.velx =  BCvel*BCboost
    ElseIf BCleft = 1 Then
      ControlActiveBall.velx = -BCvel*BCboost
    Else
      ControlActiveBall.velx = 0
    End If

    If BCup = 1 Then
      ControlActiveBall.vely = -BCvel*BCboost
    ElseIf BCdown = 1 Then
      ControlActiveBall.vely =  BCvel*BCboost
    Else
      ControlActiveBall.vely = bcyveloffset
    End If
  End If
End Sub


'***************************************
''***Prim Image Swaps***
''***************************************

'
''***************************************
''***Prim Material Swaps***
''***************************************
'''

Dim TextureArray1: TextureArray1 = Array("Plastic with an image1", "Plastic with an image1", "Plastic with an image trans","Plastic with an image trans")
Dim TextureArray2: TextureArray2 = Array("Plastic with an image1", "Plastic with an image1", "Plastic with an image2","Plastic with an image2")
Dim FilamentArray: FilamentArray = Array("WireDT_off", "WireDT_33", "WireDT_66", "WireDT_on")
Dim ClearBulbArray: ClearBulbArray = Array("BulbGIoff", "BulbGIoff", "BulbGIOn","BulbGIOn")
Dim YellowDomeArray: YellowDomeArray = Array("dome3_yellow", "dome3_yellowBright", "dome3_yellowBright", "dome3_yellowBright")
Dim RedDomeArray: RedDomeArray = Array("dome3_red", "dome3_RedBright", "dome3_RedBright", "dome3_RedBright")
Dim OrangeDomeArray: OrangeDomeArray = Array("dome3_orange2", "dome3_orange2_bright", "dome3_orange2_bright", "dome3_orange2_bright")
Dim BackdropLightArray: BackdropLightArray = Array("BackdropOff", "BackdropOff", "BackdropOn", "BackdropOn")


Dim DLintensity
Sub MatSwap(pri, group, DLintensity, ByVal aLvl)  'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
    Select case FlashLevelToIndex(aLvl, 3)
    Case 1:pri.Material = group(0) 'Full
    Case 2:pri.Material = group(1) 'Fading...
    Case 3:pri.Material = group(2) 'Fading...
        Case 4:pri.Material = group(3) 'Off
    End Select
pri.blenddisablelighting = aLvl * DLintensity
End Sub


Sub ImageSwap(pri, group, DLintensity, ByVal aLvl)  'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
    Select case FlashLevelToIndex(aLvl, 3)
    Case 1:pri.Image = group(0) 'Full
    Case 2:pri.Image = group(1) 'Fading...
    Case 3:pri.Image = group(2) 'Fading...
        Case 4:pri.Image = group(3) 'Off
    End Select
pri.blenddisablelighting = aLvl * DLintensity
End Sub


Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically

pri.blenddisablelighting = aLvl * DLintensity
End Sub

    Sub FadeBGL(pri, group, DLintensity, ByVal aLvl) ' used for multiple lights
        Select Case FlashLevelToIndex(aLvl, 3)
    Case 1:pri.Image = group(0) 'Full
    Case 2:pri.Image = group(1) 'Fading...
    Case 3:pri.Image = group(2) 'Fading...
        Case 4:pri.Image = group(3) 'Off
        End Select
  End Sub

'*********************************************************************************************************************************************************
'Begin nfozzy lamp handling
'*********************************************************************************************************************************************************

redim GILamps(99) : redim GIFlashers(99)  'new arrays
SortGI GILamps, GIFlashers, GILighting
dim TestString, TestStringAll 'debug strings
Sub SortGI(ByRef aLight,aFlasher, GImixed) 'different method using Arrays instead of scripting dictionary objects
  dim x, CountMe: CountMe = 0
  for x = 0 to (GImixed.Count-1)
    if TypeName(GImixed(x) ) = "Light" Then
      Set aLight(CountMe) = GImixed(x)
      TestString = TestString & "assigned " & GImixed(x).Name & " to aLight(" & CountMe & ")" & vbnewline 'debug
      CountMe = CountMe+1
      redim Preserve aLight(CountMe)
    end if
  Next
  CountMe = 0
  for x = 0 to (GImixed.Count-1)  '(note: this sub assumes there ARE flashers in the collection!)
    if TypeName(GImixed(x) ) = "Flasher" Then
      Set aFlasher(CountMe) = GImixed(x)
      TestString = TestString & "assigned " & GImixed(x).Name & " to aFlasher(" & CountMe & ")" & vbnewline 'debug
      CountMe = CountMe+1
      redim Preserve aFlasher(CountMe)
    end if
  Next
  redim Preserve aLight(uBound(aLight)-1) 'final trim of the arrays
  redim Preserve aFlasher(uBound(aFlasher)-1)
  'TestSTR(0) = TestSTR(0) & "ubound aLight: " & uBound(aLight) & " uBound aFlashers:" & uBound(aFlasher) 'debug
  'Debug.Print TestString
End Sub

'These arrays contain the following info of all non-GI lights (collected from GetElements via SortLamps sub)
Redim LightsA(999)' Object references
Redim LightsB(999)' Opacity / Intensity
Redim LightsC(999)' Fade Up (Light objects)
Redim LightsD(999)' Fade Down(Light Objects)

SortLamps GILighting, GIOffFlasherCorrection
Sub SortLamps(ByVal GI, aExclude) 'Sorts remaining light and flashers objects (EXCLUDES those in the GI collection)
  dim Counter,x,xx,skipme : skipme = False:Counter = 0 : TestStringAll = "Test String 2"
  for each x in GetElements 'now we're cooking
    'if TypeName(x) = "IDecal" then Continue For 'Decals don't have names. Evil imo D:
    if TypeName(x) = "Light" or TypeName(x) = "Flasher" Then
      SkipMe = False
      for each xx in GI 'Find duplicates and Skip them
        if x.Name = xx.Name then
          TestStringAll = TestStringAll & x.Name & "found in GI collection, Disregarding & Continuing..." & vbnewline 'debug
          SkipMe = True'Continue For
        End If
      next
      for each xx in aExclude 'Exclude collection
        if x.Name = xx.Name then
          TestStringAll = TestStringAll & x.Name & "found in exclude collection, Disregarding & Continuing..." & vbnewline 'debug
          SkipMe = True'Continue For
        End If
      next
      If Not SkipMe Then
        On Error Resume Next
        'LightsA(Counter) = x.name  'name
        Set LightsA(Counter) = x  'ref
        LightsB(Counter) = x.Opacity
        LightsB(Counter) = x.Intensity
        LightsC(Counter) = x.FadeSpeedUp
        LightsD(Counter) = x.FadeSpeedDown
        On Error Goto 0
        Counter = Counter + 1
        redim Preserve LightsA(Counter)
        redim Preserve LightsB(Counter)
        redim Preserve LightsC(Counter)
        redim Preserve LightsD(Counter)
      End If
    End If
  next
  redim Preserve LightsA(uBound(LightsA)-1) 'final trim of the arrays
  redim Preserve LightsB(uBound(LightsB)-1)
  redim Preserve LightsC(uBound(LightsC)-1)
  redim Preserve LightsD(uBound(LightsD)-1)

  TestStringAll = TestStringAll & "Ubound LightsA = " & UBound(LightsA) 'Debug
  'debug.print TestTwo
End Sub

' Lamp & Flasher Updates
' LampFader object (Lampz) updates on two timers: Logic on 1, Game updates on -1
' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
' Sub LampTimer_Timer()
'   dim x, chglamp
'   chglamp = Controller.ChangedLamps
'   If Not IsEmpty(chglamp) Then
'     For x = 0 To UBound(chglamp)      'nmbr = chglamp(x, 0), state = chglamp(x, 1)
'       Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
'     next
'   End If
'   'Lampz.Update1  'update (fading logic only)
'   Lampz.Update2 'update (Pinmame and Fading (for -1, lower latency)
' End Sub

function FlashLevelToIndex(Input, MaxSize)
  'FlashLevelToIndex = cInt(Input * (MaxSize-1)+.5)+1
     FlashLevelToIndex = cInt(MaxSize * Input)
end function

'Lamp Filter
Function LampFilter(aLvl)
  LampFilter = aLvl^1.6 'exponential curve?
End Function

'Collections to arrays
Function ColtoArray(aDict)  'converts a collection to an indexed array. Indexes will come out random probably.
  redim a(999)
  dim count : count = 0
  dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
  redim preserve a(count-1) : ColtoArray = a
End Function

'Setlamp, etc
  'Solenoid pipeline looks like this:
  'Pinmame Controller -> UseSolenoids -> Solcallback -> intermediate subs (here) -> Lampz fading object -> object updates / more callbacks

  'Lamps, for reference:
  'Pinmame Controller -> UpdateLamps sub -> Lampz Fading Object -> Object Updates / callbacks

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
' Sub SetLamp(aNr, aOn)
'   Lampz.state(aNr) = abs(aOn)
' End Sub

Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New LampFader
InitLamps
LampTimer.Interval = -1 '1
LampTimer.Enabled = 1

' Lamp & Flasher Updates
' LampFader object (Lampz) updates on two timers: Logic on 1, Game updates on -1
Sub LampTimer_Timer()
  dim x, chglamp
  chglamp = Controller.ChangedLamps
  If Not IsEmpty(chglamp) Then
    For x = 0 To UBound(chglamp)      'nmbr = chglamp(x, 0), state = chglamp(x, 1)
      Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
    next
  End If
  'Lampz.Update1  'update (fading logic only)
  Lampz.Update2 'update (Pinmame and Fading (for -1, lower latency)
End Sub

Sub InitLamps()
  'Filtering (comment out to disable)
  Lampz.Filter = "LampFilter" 'Puts all lamp intensity scale output (no callbacks) through this function before updating
  'Adjust fading speeds (1 / full MS fading time)
  dim x : for x = 0 to 140 : Lampz.FadeSpeedUp(x) = 1/90 : Lampz.FadeSpeedDown(x) = 1/100 : next
  Lampz.FadeSpeedUp(110) = 1/64 'GI

'*********************************************************************************************************************************************************
'End nfozzy lamp handling
'*********************************************************************************************************************************************************
  'Lamp Assignments
  Lampz.Callback(1) = "  DisableLighting pl1, 200,"
  Lampz.MassAssign(1) = l1
  Lampz.MassAssign(1) = l1a
  Lampz.MassAssign(1) = l1d

  Lampz.Callback(2) = "  DisableLighting pl2, 200,"
  Lampz.MassAssign(2) = l2
  Lampz.MassAssign(2) = l2a
  Lampz.MassAssign(2) = l2d

  Lampz.Callback(3) = "  DisableLighting pl3, 200,"
  Lampz.MassAssign(3) = l3
  Lampz.MassAssign(3) = l3a
  Lampz.MassAssign(3) = l3d

  Lampz.Callback(4) = "  DisableLighting pl4, 200,"
  Lampz.MassAssign(4) = l4
  Lampz.MassAssign(4) = l4a
  Lampz.MassAssign(4) = l4d

  Lampz.Callback(5) = "  DisableLighting pl5, 200,"
  Lampz.MassAssign(5) = l5
  Lampz.MassAssign(5) = l5a
  Lampz.MassAssign(5) = l5d

  Lampz.Callback(6) = "  DisableLighting pl6, 200,"
  Lampz.MassAssign(6) = l6
  Lampz.MassAssign(6) = l6a
  Lampz.MassAssign(6) = l6d

  Lampz.Callback(7) = "  DisableLighting pl7, 200,"
  Lampz.MassAssign(7) = l7
  Lampz.MassAssign(7) = l7a
  Lampz.MassAssign(7) = l7d

  Lampz.Callback(8) = "  DisableLighting pl8, 200,"
  Lampz.MassAssign(8) = l8
  Lampz.MassAssign(8) = l8a
  Lampz.MassAssign(8) = l8d

  Lampz.Callback(9) = "  DisableLighting pl9, 50,"
  Lampz.MassAssign(9) = l9
  Lampz.MassAssign(9) = l9a
  Lampz.MassAssign(9) = l9d

  Lampz.Callback(10) = "  DisableLighting pl10, 50,"
  Lampz.MassAssign(10) = l10
  Lampz.MassAssign(10) = l10a
  Lampz.MassAssign(10) = l10d

  Lampz.Callback(11) = "  DisableLighting pl11, 50,"
  Lampz.MassAssign(11) = l11
  Lampz.MassAssign(11) = l11a
  Lampz.MassAssign(11) = l11d

  Lampz.Callback(12) = "  DisableLighting pl12, 50,"' Lampz.MassAssign(12) = l12
  Lampz.MassAssign(12) = l12a
  Lampz.MassAssign(12) = l12d

  Lampz.Callback(13) = "  DisableLighting pl13, 50,"
  Lampz.MassAssign(13) = l13
  Lampz.MassAssign(13) = l13a
  Lampz.MassAssign(13) = l13d

  Lampz.Callback(14) = "  DisableLighting pl14, 25,"
  Lampz.MassAssign(14) = l14
  Lampz.MassAssign(14) = l14a
  Lampz.MassAssign(14) = l14d

  Lampz.Callback(15) = "  DisableLighting pl15, 25,"
  Lampz.MassAssign(15) = l15
  Lampz.MassAssign(15) = l15a
  Lampz.MassAssign(15) = l15d

  Lampz.Callback(16) = "  DisableLighting pl16, 25,"
  Lampz.MassAssign(16) = l16
  Lampz.MassAssign(16) = l16a
  Lampz.MassAssign(16) = l16d

  Lampz.Callback(17) = "  DisableLighting pl17, 150,"'  Lampz.MassAssign(17) = l17
  Lampz.MassAssign(17) = l17a
  Lampz.MassAssign(17) = l17d

  Lampz.Callback(18) = "  DisableLighting pl18, 150,"
  Lampz.MassAssign(18) = l18
  Lampz.MassAssign(18) = l18a
  Lampz.MassAssign(18) = l18d

  Lampz.Callback(19) = "  DisableLighting pl19, 150,"
  Lampz.MassAssign(19) = l19
  Lampz.MassAssign(19) = l19a
  Lampz.MassAssign(19) = l19d

  Lampz.Callback(20) = "  DisableLighting pl20, 500,"
  Lampz.MassAssign(20) = l20
  Lampz.MassAssign(20) = l20a
  Lampz.MassAssign(20) = l20d

  Lampz.Callback(21) = "  DisableLighting pl21, 5,"
  Lampz.MassAssign(21) = l21
  Lampz.MassAssign(21) = l21a
  Lampz.MassAssign(21) = l21d

  Lampz.Callback(22) = "  DisableLighting pl22, 5,"
  Lampz.MassAssign(22) = l22
  Lampz.MassAssign(22) = l22a
  Lampz.MassAssign(22) = l22d

  Lampz.Callback(23) = "  DisableLighting pl23, 5,"
  Lampz.MassAssign(23) = l23
  Lampz.MassAssign(23) = l23a
  Lampz.MassAssign(23) = l23d

  Lampz.Callback(24) = "  DisableLighting pl24, 5,"
  Lampz.MassAssign(24) = l24
  Lampz.MassAssign(24) = l24a
  Lampz.MassAssign(24) = l24d

  Lampz.Callback(25) = "  DisableLighting pl25, 300,"
  Lampz.MassAssign(25) = l25
  Lampz.MassAssign(25) = l25a
  Lampz.MassAssign(25) = l25d

  Lampz.Callback(26) = "  DisableLighting pl26, 400,"
  Lampz.MassAssign(26) = l26
  Lampz.MassAssign(26) = l26a
  Lampz.MassAssign(26) = l26d

  Lampz.Callback(27) = "  DisableLighting pl27, 400,"
  Lampz.MassAssign(27) = l27
  Lampz.MassAssign(27) = l27a
  Lampz.MassAssign(27) = l27d

  Lampz.Callback(28) = "  DisableLighting pl28, 50,"
  Lampz.MassAssign(28) = l28
  Lampz.MassAssign(28) = l28a
  Lampz.MassAssign(28) = l28d

  Lampz.Callback(29) = "  DisableLighting pl29, 100,"
  Lampz.MassAssign(29) = l29
  Lampz.MassAssign(29) = l29a
  Lampz.MassAssign(29) = l29d

  Lampz.Callback(30) = "  DisableLighting pl30, 100,"
  Lampz.MassAssign(30) = l30
  Lampz.MassAssign(30) = l30a
  Lampz.MassAssign(30) = l30d

  Lampz.Callback(31) = "  DisableLighting pl31, 100,"
  Lampz.MassAssign(31) = l31
  Lampz.MassAssign(31) = l31a
  Lampz.MassAssign(31) = l31d

  Lampz.Callback(32) = "  DisableLighting pl32, 100,"
  Lampz.MassAssign(32) = l32
  Lampz.MassAssign(32) = l32a
  Lampz.MassAssign(32) = l32d

  Lampz.Callback(33) = "  DisableLighting pl33, 25,"
  Lampz.MassAssign(33) = l33
  Lampz.MassAssign(33) = l33a
  Lampz.MassAssign(33) = l33d

  Lampz.Callback(34) = "  DisableLighting pl34, 25,"
  Lampz.MassAssign(34) = l34
  Lampz.MassAssign(34) = l34a
  Lampz.MassAssign(34) = l34d

  Lampz.Callback(35) = "  DisableLighting pl35, 25,"
  Lampz.MassAssign(35) = l35
  Lampz.MassAssign(35) = l35a
  Lampz.MassAssign(35) = l35d

  Lampz.Callback(36) = "  DisableLighting pl36, 25,"
  Lampz.MassAssign(36) = l36
  Lampz.MassAssign(36) = l36a
  Lampz.MassAssign(36) = l36d

  Lampz.Callback(37) = "  DisableLighting pl37, 150,"
  Lampz.MassAssign(37) = l37
  Lampz.MassAssign(37) = l37a
  Lampz.MassAssign(37) = l37d

  Lampz.Callback(38) = "  DisableLighting pl38, 350,"
  Lampz.MassAssign(38) = l38
  Lampz.MassAssign(38) = l38a
  Lampz.MassAssign(38) = l38d

  Lampz.Callback(39) = "  DisableLighting pl39, 350,"
  Lampz.MassAssign(39) = l39
  Lampz.MassAssign(39) = l39a
  Lampz.MassAssign(39) = l39d

  Lampz.Callback(40) = "  DisableLighting pl40, 350,"
  Lampz.MassAssign(40) = l40
  Lampz.MassAssign(40) = l40a
  Lampz.MassAssign(40) = l40d

  Lampz.Callback(41) = "  DisableLighting pl41, 100,"
  Lampz.MassAssign(41) = l41
  Lampz.MassAssign(41) = l41a
  Lampz.MassAssign(41) = l41d

  Lampz.Callback(42) = "  DisableLighting pl42, 150,"
  Lampz.MassAssign(42) = l42
  Lampz.MassAssign(42) = l42a
  Lampz.MassAssign(42) = l42d

  Lampz.Callback(43) = "  DisableLighting pl43, 150,"
  Lampz.MassAssign(43) = l43
  Lampz.MassAssign(43) = l43a
  Lampz.MassAssign(43) = l43d

  Lampz.Callback(44) = "  DisableLighting pl44, 100,"
  Lampz.MassAssign(44) = l44
  Lampz.MassAssign(44) = l44a
  Lampz.MassAssign(44) = l44d

  Lampz.Callback(45) = "  DisableLighting pl45, 5,"
  Lampz.MassAssign(45) = l45
  Lampz.MassAssign(45) = l45a
  Lampz.MassAssign(45) = l45d

  Lampz.Callback(46) = "  DisableLighting pl46, 250,"
  Lampz.MassAssign(46) = l46
  Lampz.MassAssign(46) = l46a
  Lampz.MassAssign(46) = l46d

  Lampz.Callback(47) = "  DisableLighting pl47, 350,"
  Lampz.MassAssign(47) = l47
  Lampz.MassAssign(47) = l47a
  Lampz.MassAssign(47) = l47d

  Lampz.Callback(48) = "  DisableLighting pl48, 150,"
  Lampz.MassAssign(48) = l48
  Lampz.MassAssign(48) = l48a
  Lampz.MassAssign(48) = l48d

If BackGlassVisible = 1 then
  Lampz.MassAssign(49) = BGSol49
End If
  Lampz.Callback(49) = " FadeBGL lbg49, BackdropLightArray, .15,"
  Lampz.MassAssign(49) = lbg49
If BackGlassVisible = 1 then
  Lampz.MassAssign(50) = BGSol50
End If
  Lampz.Callback(50) = " FadeBGL lbg50, BackdropLightArray, .15,"
  Lampz.MassAssign(50) = lbg50
If BackGlassVisible = 1 then
  Lampz.MassAssign(50) = BGSol51
End If
  Lampz.Callback(51) = " FadeBGL lbg51, BackdropLightArray, .15,"
  Lampz.MassAssign(51) = lbg51
If BackGlassVisible = 1 then
  Lampz.MassAssign(51) = BGSol52
End If
  Lampz.Callback(52) = " FadeBGL lbg52, BackdropLightArray, .15,"
  Lampz.MassAssign(52) = lbg52
If BackGlassVisible = 1 then
  Lampz.MassAssign(52) = BGSol53
End If

  Lampz.Callback(53) = " FadeBGL lbg53, BackdropLightArray, .15,"
  Lampz.MassAssign(53) = lbg53
If BackGlassVisible = 1 then
  Lampz.MassAssign(54) = BGSol54
End If

If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(55) = BGLight55
Else
  Lampz.MassAssign(55) = BGLight55
End If
End If

  Lampz.Callback(56) = "  DisableLighting pl56, 150,"
  Lampz.MassAssign(56) = l56
  Lampz.MassAssign(56) = l56a
  Lampz.MassAssign(56) = l56d
  Lampz.MassAssign(57) = l57
  Lampz.MassAssign(58) = l58
  Lampz.MassAssign(59) = l59
  Lampz.MassAssign(60) = l60
  Lampz.MassAssign(61) = l61
  Lampz.MassAssign(62) = l62
  Lampz.MassAssign(62) = l62a
  Lampz.MassAssign(63) = l63
  Lampz.MassAssign(63) = l63a
  Lampz.MassAssign(64) = l64
  Lampz.MassAssign(64) = l64a



'*****************
'Flasher Assignments
'*****************

If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(101) = BGSolKingShooter
Else
  Lampz.MassAssign(101) = BGSolKingShooter
End If
End If
  Lampz.MassAssign(101) = l101a
  Lampz.MassAssign(101) = l101b
  Lampz.Callback(101) = "  DisableLighting pl101, 1000,"
' Lampz.MassAssign(101) = l101



If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(102) = BGSolKongShooter
Else
  Lampz.MassAssign(102) = BGSolKongShooter
End If
End If

  Lampz.MassAssign(102) = l102a
  Lampz.MassAssign(102) = l102b
  Lampz.Callback(102) = "  DisableLighting pl102, 1000,"
' Lampz.MassAssign(102) = l102


If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(103) = BGSolSkull
Else
  Lampz.MassAssign(103) = BGSolSkull
End If
End If

  Lampz.MassAssign(103) = l103b
  Lampz.Callback(103) = "  DisableLighting pl103b, 100,"
  Lampz.MassAssign(103) = f103b
  Lampz.Callback(103) = "  DisableLighting pl103a, 100,"
  Lampz.MassAssign(103) = f103a
  Lampz.MassAssign(103) = l103a


If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(104) = BGSolDNSRS
Else
  Lampz.MassAssign(104) = BGSolDNSRS
End If
End If
  Lampz.Callback(104) = "  DisableLighting pBulbFlasher004, 5,"
  Lampz.Callback(104) = " MatSwap Primitive8, TextureArray2, .15,"
  Lampz.Callback(104) = " MatSwap Primitive8b, TextureArray2, .15,"

' If FlasherMod2Type = 1 then
    Lampz.Callback(104) = "ImageSwap Primitive8b, YellowDomeArray, 5,"
' Else
    Lampz.Callback(104) = "ImageSwap Primitive8, RedDomeArray, 5,"
' End If
  Lampz.MassAssign(104) = f104b
  Lampz.MassAssign(104) = f104a


If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(105) = BGSolTower25
Else
  Lampz.MassAssign(105) = BGSolTower25
End If
End If

  Lampz.MassAssign(105) = f105
  Lampz.Callback(105) = "  DisableLighting pl21, 5,"
  Lampz.MassAssign(105) = f105d
  Lampz.MassAssign(105) = f105a

If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(106) = BGSolTower50
Else
  Lampz.MassAssign(106) = BGSolTower50
End If
End If

  Lampz.MassAssign(106) = f106
  Lampz.Callback(106) = "  DisableLighting pl22, 5,"
  Lampz.MassAssign(106) = f106d
  Lampz.MassAssign(106) = f106a

If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(107) = BGSolTower100
Else
  Lampz.MassAssign(107) = BGSolTower100
End If
End If

  Lampz.MassAssign(107) = f107
  Lampz.Callback(107) = "  DisableLighting pl23, 5,"
  Lampz.MassAssign(107) = f107d
  Lampz.MassAssign(107) = f107a

If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(108) = BGSolTowerEB
Else
  Lampz.MassAssign(108) = BGSolTowerEB
End If
End If

  Lampz.MassAssign(108) = f108
  Lampz.Callback(108) = "  DisableLighting pl24, 5,"
  Lampz.MassAssign(108) = f108d
  Lampz.MassAssign(108) = f108a

If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(109) = BGSolTowerMil
Else
  Lampz.MassAssign(109) = BGSolTowerMil
End If
End If

  Lampz.MassAssign(109) = f109
  Lampz.Callback(109) = "  DisableLighting pl45, 5,"
  Lampz.MassAssign(109) = f109d
  Lampz.MassAssign(109) = f109a

If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(113) = BGSolKong
Else
  Lampz.MassAssign(113) = BGSolKong
End If
End If

  Lampz.MassAssign(113) = l113a
  Lampz.MassAssign(113) = l113b


  Lampz.Callback(115) = "  DisableLighting pBulbFilament2, 5,"
  Lampz.Callback(115) = "  DisableLighting pBulbFlasher2, 5,"
  Lampz.Callback(115) = "ImageSwap pBulbFilament2, FilamentArray, 5,"
  Lampz.Callback(115) = "MatSwap pBulbFlasher2, ClearBulbArray, .15,"
  Lampz.MassAssign(115) = F115X
  Lampz.MassAssign(115) = F115X1
  Lampz.MassAssign(115) = f115xx


If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(122) = BGSolGirl
Else
  Lampz.MassAssign(122) = BGSolGirl
End If
End If

  Lampz.MassAssign(122) = f112
  Lampz.Callback(122) = "  DisableLighting pl112, 5,"
  Lampz.MassAssign(122) = l112f
  Lampz.MassAssign(122) = Flasher1
  Lampz.MassAssign(122) = l112
  Lampz.MassAssign(122) = l112a
  Lampz.MassAssign(122) = l112b
  Lampz.MassAssign(122) = l112c


  Lampz.Callback(122) = "  DisableLighting pBulbFilament1, 5,"
  Lampz.Callback(122) = "ImageSwap pBulbFilament1, FilamentArray, 5,"
  Lampz.Callback(122) = "MatSwap pBulbFlasher1, ClearBulbArray, .15,"
  Lampz.MassAssign(122) = F112X

  Lampz.Callback(124) = "  DisableLighting Primitive2, 5,"
  Lampz.Callback(124) = "ImageSwap Primitive2, YellowDomeArray, 5,"
  Lampz.Callback(124) = "MatSwap Primitive2, TextureArray2, .15,"
  Lampz.MassAssign(124) = f114b2
  Lampz.MassAssign(124) = f114a2

If BackGlassVisible = 1 then
If GION = 1 Then
  Lampz.MassAssign(112) = BGSolGirl
Else
  Lampz.MassAssign(112) = BGSolGirl
End If
End If

  Lampz.MassAssign(112) = f112
  Lampz.Callback(112) = "  DisableLighting pl112, 5,"
  Lampz.MassAssign(112) = l112f
  Lampz.MassAssign(112) = Flasher1
  Lampz.MassAssign(112) = l112
  Lampz.MassAssign(112) = l112a
  Lampz.MassAssign(112) = l112b
  Lampz.MassAssign(112) = l112c

  Lampz.Callback(114) = "  DisableLighting pBulbFlasher003, 15,"
  Lampz.Callback(114) = "  DisableLighting Primitive7, 5,"
  Lampz.Callback(114) = "ImageSwap Primitive7, OrangeDomeArray, 5,"
  Lampz.Callback(114) = "MatSwap Primitive7, TextureArray2, .15,"
  Lampz.MassAssign(114) = f114b
  Lampz.MassAssign(114) = f114a



'*****************
'GI Assignments
'*****************


  Lampz.Callback(110) = "  DisableLighting plLCB, 100,"
  Lampz.MassAssign(110) = lLCB
  Lampz.MassAssign(110) = lLCBa
  Lampz.MassAssign(110) = lLCBd
  Lampz.Callback(110) = "  DisableLighting plLJB, 100,"
  Lampz.MassAssign(110) = lLJB
  Lampz.MassAssign(110) = lLJBa
  Lampz.MassAssign(110) = lLJBd

  If BackGlassVisible = 1 then
    Lampz.MassAssign(110) = BGSol11
  End If

  Lampz.Callback(110) = "  DisableLighting pbcap1, 1,"
  Lampz.MassAssign(110) = bLight1
  Lampz.Callback(110) = "  DisableLighting pbcap2, 1,"
  Lampz.MassAssign(110) = bLight2
  Lampz.Callback(110) = "  DisableLighting pbcap3, 1,"
  Lampz.MassAssign(110) = bLight3

  Lampz.Callback(110) = "GIUpdates"
  Lampz.obj(110) = ColtoArray(GILighting)
  Lampz.state(110) = 1    'Turn on GI to Start

  'Turn off all lamps on startup
  lampz.TurnOnStates  'Set any lamps state to 1. (Object handles fading!)
  lampz.update

End Sub


'*********************************************************************************************************************************************************
'Begin lamp helper functions
'*********************************************************************************************************************************************************

'***************************************
'System 11 GI On/Off
'***************************************
'Sub GIOn  : SetGI False: End Sub 'These are just debug commands now
'Sub GIOff : SetGI True : End Sub

'***************************************
'GI off insert lamps intensity boost
'***************************************
Dim GIoffMult : GIoffMult = 3 'Multiplies all non-GI inserts lights opacities when the GI is off
Sub GIupdates(ByVal aLvl) 'GI update odds and ends go here

  dim giAvg
  if Lampz.UseFunction then
    aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
    giAvg = aLvl'(LampFilter(Lampz.Lvl(1)) + LampFilter(Lampz.Lvl(4)) )/2
' Else
'   giAvg = (Lampz.Lvl(1) + Lampz.Lvl(4) )/2
  end if

  'Lut Fading
  dim LutName, LutCount, GoLut
  LutName = "LutCont_"
  LutCount = 27
  GoLut = cInt(LutCount * giAvg )'+1  '+1 if no 0 with these luts
  GoLut = LutName & GoLut
  if kk.ColorGradeImage <> GoLut then kk.ColorGradeImage = GoLut ':   tb.text = golut

  'Fade lamps up when GI is off
  dim GIscale
  GiScale = (GIoffMult-1) * (ABS(aLvl-1 )  ) + 1  'invert
  dim x : for x = 0 to uBound(LightsA)
    On Error Resume Next
''    LightsA(x).Opacity = LightsB(x) * GIscale
'   LightsA(x).Intensity = LightsB(x) * GIscale
    'LightsA(x).FadeSpeedUp = LightsC(x) * GIscale
    'LightsA(x).FadeSpeedDown = LightsD(x) * GIscale
    On Error Goto 0
  Next
End Sub

Sub Lamp114(aOn)
        Select Case aOn
    Case True  'GI off
      Playsound "fx_FlashRelayOff"
      If flashermodtype = 1 then
        SetLamp 124, 5  'Inverted, Solenoid cuts GI circuit on this era of game
      Else
        SetLamp 114, 5  'Inverted, Solenoid cuts GI circuit on this era of game
      End If
    Case False
      Playsound "fx_FlashRelayOn"
      If flashermodtype = 1 then
        SetLamp 124, 0  'Inverted, Solenoid cuts GI circuit on this era of game
      Else
        SetLamp 114, 0  'Inverted, Solenoid cuts GI circuit on this era of game
      End If
  End Select
End Sub





Sub Lamp112(aOn)
        Select Case aOn
    Case True  'GI off
      Playsound "fx_FlashRelayOff"
      If flashermodtype = 1 then
        SetLamp 122, 5  'Inverted, Solenoid cuts GI circuit on this era of game
      Else
        SetLamp 112, 5  'Inverted, Solenoid cuts GI circuit on this era of game
      End If
    Case False
      Playsound "fx_FlashRelayOn"
      If flashermodtype = 1 then
        SetLamp 122, 0  'Inverted, Solenoid cuts GI circuit on this era of game
      Else
        SetLamp 112, 0  'Inverted, Solenoid cuts GI circuit on this era of game
      End If
  End Select
End Sub


'***************************************
'GI off finsert light falloff-power boost
'***************************************
Dim GiOffFOP, GION, itemLightUpPegs, itemGIBulbs, itemGIBulbFil
Sub SetGI(aOn)
        Select Case aOn
    Case True  'GI off
      Playsound "fx_FlashRelayOff"
      SetLamp 110, 0  'Inverted, Solenoid cuts GI circuit on this era of game
            For each GiOffFOP in LampsInserts 'increases falloff power for lamps in "lampinserts" collection, when GI is turned off
        GIOffFOP.falloffpower = 1.25 'Sets falloff power to 1.25 when GI is ON
            next

  If  GION = 1 then playsound "flasher_relay_off", 0
  GION = 0

  Shadow.opacity = 50
  Shadow.imageA = "KK - Shadow GI-OFF"

  for each itemLightUpPegs in LightUpPegs
    itemLightUpPegs.BlendDisableLighting = 0
    next

  for each itemGIBulbs in GIBulbs
    itemGIBulbs.BlendDisableLighting = 0
    next

  for each itemGIBulbFil in GIBulbFil
    itemGIBulbFil.BlendDisableLighting = 0
    next

  pPlastics13.BlendDisableLightingFromBelow = 1
  pWall021.image = "Wall021 GI-OFF"
  RampClearMeS34.image = "KK - Ramp clear Metal-OFF"

  pRailsOptinal1.BlendDisableLighting = 0
' Flasher2.visible = 0
  pkkwallshooterlan1.image = "KK - Wall shooter lane"

  fDMD.ImageA = "DMD Display_cropped_s_unlit"
  fDMD_reflection.ImageA = "DMD Display_reflection_unlit"

    pRI001.DisableLighting = 0
    fRI001.Visible = False
    pRI002.DisableLighting = 0
    fRI002.Visible = False
    pRI003.DisableLighting = 0
    fRI003.Visible = False
    pRI004.DisableLighting = 0
    fRI004.Visible = False

    Case False
      Playsound "fx_FlashRelayOn"
      SetLamp 110, 5
            For each GiOffFOP in LampsInserts  'reduces falloff power for lamps in "lampinserts" collection, when GI is turned off
        GiOffFOP.falloffpower = 4 'Sets falloff power to 4 when GI is OFF
            next

  If  GION = 0 then playsound "flasher_relay_on", 0
  GION = 1

  Shadow.opacity = 95
  If OptinalRails1Type = 1 Then
    Shadow.imageA = "KK - Shadow GI-ON-2"
  Else
    Shadow.imageA = "KK - Shadow GI-ON"
  End If

  for each itemLightUpPegs in LightUpPegs
    itemLightUpPegs.BlendDisableLighting = 1
    next

  for each itemGIBulbs in GIBulbs
    itemGIBulbs.BlendDisableLighting = 1.5
    next

  for each itemGIBulbFil in GIBulbFil
    itemGIBulbFil.BlendDisableLighting = 5
    next

  pPlastics13.BlendDisableLightingFromBelow = 0
  pWall021.image = "Wall021 GI-ON"
  RampClearMeS34.image = "KK - Ramp clear Metal-ON"
  pRailsOptinal1.BlendDisableLighting = .3
  pkkwallshooterlan1.image = "KK - Wall shooter lane-Dim"

  fDMD.ImageA = "DMD Display_cropped_s_lit"
  fDMD_reflection.ImageA = "DMD Display_reflection_lit"

  If RandomInsertGI = 1 Then
    pRI001.DisableLighting = 5
    fRI001.Visible = True
    pRI002.DisableLighting = 5
    fRI002.Visible = True
    pRI003.DisableLighting = 5
    fRI003.Visible = True
    pRI004.DisableLighting = 5
    fRI004.Visible = True

  Else
    pRI001.DisableLighting = 0
    fRI001.Visible = False
    pRI002.DisableLighting = 0
    fRI002.Visible = False
    pRI003.DisableLighting = 0
    fRI003.Visible = False
    pRI004.DisableLighting = 0
    fRI004.Visible = False
  End If


  End Select
End Sub

'*********************************************************************************************************************************************************
'End lamp helper functions
'*********************************************************************************************************************************************************




'====================
'Class jungle nf
'=============

'No-op object instead of adding more conditionals to the main loop
'It also prevents errors if empty lamp numbers are called, and it's only one object
'should be g2g?

Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class

'version 0.11 - Mass Assign, Changed modulate style
'version 0.12 - Update2 (single -1 timer update) update method for core.vbs
'Version 0.12a - Filter can now be accessed via 'FilterOut'
'Version 0.12b - Changed MassAssign from a sub to an indexed property (new syntax: lampfader.MassAssign(15) = Light1 )
'Version 0.13 - No longer requires setlocale. Callback() can be assigned multiple times per index
' Note: if using multiple 'LampFader' objects, set the 'name' variable to avoid conflicts with callbacks

Class LampFader
  Public FadeSpeedDown(140), FadeSpeedUp(140)
  Private Lock(140), Loaded(140), OnOff(140)
  Public UseFunction
  Private cFilter
  Public UseCallback(140), cCallback(140)
  Public Lvl(140), Obj(140)
  Private Mult(140)
  Public FrameTime
  Private InitFrame
  Public Name

  Sub Class_Initialize()
    InitFrame = 0
    dim x : for x = 0 to uBound(OnOff)  'Set up fade speeds
      FadeSpeedDown(x) = 1/100  'fade speed down
      FadeSpeedUp(x) = 1/80   'Fade speed up
      UseFunction = False
      lvl(x) = 0
      OnOff(x) = False
      Lock(x) = True : Loaded(x) = False
      Mult(x) = 1
    Next
    Name = "LampFaderNF" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    for x = 0 to uBound(OnOff)    'clear out empty obj
      if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
    Next
  End Sub

  Public Property Get Locked(idx) : Locked = Lock(idx) : End Property   'debug.print Lampz.Locked(100)  'debug
  Public Property Get state(idx) : state = OnOff(idx) : end Property
  Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
  Public Function FilterOut(aInput) : if UseFunction Then FilterOut = cFilter(aInput) Else FilterOut = aInput End If : End Function
  'Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property
  Public Property Let Callback(idx, String)
    UseCallBack(idx) = True
    'cCallback(idx) = String 'old execute method
    'New method: build wrapper subs using ExecuteGlobal, then call them
    cCallback(idx) = cCallback(idx) & "___" & String  'multiple strings dilineated by 3x _

    dim tmp : tmp = Split(cCallback(idx), "___")

    dim str, x : for x = 0 to uBound(tmp) 'build proc contents
      'If Not tmp(x)="" then str = str & "  " & tmp(x) & " aLVL" & "  '" & x & vbnewline  'more verbose
      If Not tmp(x)="" then str = str & tmp(x) & " aLVL:"
    Next

    dim out : out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    'if idx = 132 then msgbox out 'debug
    ExecuteGlobal Out

  End Property

  Public Property Let state(ByVal idx, input) 'Major update path
    if Input <> OnOff(idx) then  'discard redundant updates
      OnOff(idx) = input
      Lock(idx) = False
      Loaded(idx) = False
    End If
  End Property

  'Mass assign, Builds arrays where necessary
  'Sub MassAssign(aIdx, aInput)
  Public Property Let MassAssign(aIdx, aInput)
    If typename(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
      if IsArray(aInput) then
        obj(aIdx) = aInput
      Else
        Set obj(aIdx) = aInput
      end if
    Else
      Obj(aIdx) = AppendArray(obj(aIdx), aInput)
    end if
  end Property

  Sub SetLamp(aIdx, aOn) : state(aIdx) = aOn : End Sub  'Solenoid Handler

  Public Sub TurnOnStates() 'If obj contains any light objects, set their states to 1 (Fading is our job!)
    dim debugstr
    dim idx : for idx = 0 to uBound(obj)
      if IsArray(obj(idx)) then
        'debugstr = debugstr & "array found at " & idx & "..."
        dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
        for x = 0 to uBound(tmp)
          if typename(tmp(x)) = "Light" then DisableState tmp(x)' : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline
          tmp(x).intensityscale = 0.001 ' this can prevent init stuttering
        Next
      Else
        if typename(obj(idx)) = "Light" then DisableState obj(idx)' : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline
        obj(idx).intensityscale = 0.001 ' this can prevent init stuttering
      end if
    Next
    'debug.print debugstr
  End Sub
  Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub 'turn state to 1

  Public Sub Init() 'Just runs TurnOnStates right now
    TurnOnStates
  End Sub

  Public Property Let Modulate(aIdx, aCoef) : Mult(aIdx) = aCoef : Lock(aIdx) = False : Loaded(aIdx) = False: End Property
  Public Property Get Modulate(aIdx) : Modulate = Mult(aIdx) : End Property

  Public Sub Update1()   'Handle all boolean numeric fading. If done fading, Lock(x) = True. Update on a '1' interval Timer!
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x)
          if Lvl(x) >= 1 then Lvl(x) = 1 : Lock(x) = True
        elseif Not OnOff(x) then 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x)
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
  End Sub

  Public Sub Update2()   'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
    FrameTime = gametime - InitFrame : InitFrame = GameTime 'Calculate frametime
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= 1 then Lvl(x) = 1 : Lock(x) = True
        elseif Not OnOff(x) then 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
    Update
  End Sub

  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    dim x,xx : for x = 0 to uBound(OnOff)
      if not Loaded(x) then
        if IsArray(obj(x) ) Then  'if array
          If UseFunction then
            for each xx in obj(x) : xx.IntensityScale = cFilter(Lvl(x)*Mult(x)) : Next
          Else
            for each xx in obj(x) : xx.IntensityScale = Lvl(x)*Mult(x) : Next
          End If
        else            'if single lamp or flasher
          If UseFunction then
            obj(x).Intensityscale = cFilter(Lvl(x)*Mult(x))
          Else
            obj(x).Intensityscale = Lvl(x)
          End If
        end if
        if TypeName(lvl(x)) <> "Double" and typename(lvl(x)) <> "Integer" then msgbox "uhh " & 2 & " = " & lvl(x)
        'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
        If UseCallBack(x) then Proc name & x,Lvl(x)*mult(x) 'Proc
        If Lock(x) Then
          if Lvl(x) = 1 or Lvl(x) = 0 then Loaded(x) = True 'finished fading
        end if
      end if
    Next
  End Sub
End Class


'Helper functions

Sub Proc(string, Callback)  'proc using a string and one argument
  'On Error Resume Next
  dim p : Set P = GetRef(String)
  P Callback
  If err.number = 13 then  msgbox "Proc error! No such procedure: " & vbnewline & string
  if err.number = 424 then msgbox "Proc error! No such Object"
End Sub


Function AppendArray(ByVal aArray, aInput)  'append one value, object, or Array onto the end of a 1 dimensional array
  if IsArray(aInput) then 'Input is an array...
    dim tmp : tmp = aArray
    If not IsArray(aArray) Then 'if not array, create an array
      tmp = aInput
    Else            'Append existing array with aInput array
      Redim Preserve tmp(uBound(aArray) + uBound(aInput)+1) 'If existing array, increase bounds by uBound of incoming array
      dim x : for x = 0 to uBound(aInput)
        if isObject(aInput(x)) then
          Set tmp(x+uBound(aArray)+1 ) = aInput(x)
        Else
          tmp(x+uBound(aArray)+1 ) = aInput(x)
        End If
      Next
    AppendArray = tmp  'return new array
    End If
  Else 'Input is NOT an array...
    If not IsArray(aArray) Then 'if not array, create an array
      aArray = Array(aArray, aInput)
    Else
      Redim Preserve aArray(uBound(aArray)+1) 'If array, increase bounds by 1
      if isObject(aInput) then
        Set aArray(uBound(aArray)) = aInput
      Else
        aArray(uBound(aArray)) = aInput
      End If
    End If
    AppendArray = aArray 'return new array
  End If
End Function


dim GameOnFF
GameOnFF = 0



'****************************************************************************
'PHYSICS DAMPENERS
'****************************************************************************

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub


dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False  'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    if gametime > 100 then Report
  End Sub

  public sub Dampen(aBall)
    if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

' Thalamus - patched : ' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()   'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


End Class

'Tracks ball velocity for judging bounce calculations & angle
'apologies to JimmyFingers is this is what his script does. I know his tracks ball velocity too but idk how it works in particular
'dim cor : set cor = New CoRTracker
'cor.debugOn = False
''cor.update() - put this on a low interval timer
'Class CoRTracker
' public DebugOn 'tbpIn.text
' public ballvel
'
' Private Sub Class_Initialize : redim ballvel(0) : End Sub
' 'TODO this would be better if it didn't do the sorting every ms, but instead every time it's pulled for COR stuff
' Public Sub Update() 'tracks in-ball-velocity
'   dim str, b, AllBalls, highestID : allBalls = getballs
'   if uBound(allballs) < 0 then if DebugOn then str = "no balls" : TBPin.text = str : exit Sub else exit sub end if: end if
'   for each b in allballs
'     if b.id >= HighestID then highestID = b.id
'   Next
'
'   if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds
'
'   for each b in allballs
'     ballvel(b.id) = BallSpeed(b)
'     if DebugOn then
'       dim s, bs 'debug spacer, ballspeed
'       bs = round(BallSpeed(b),1)
'       if bs < 10 then s = " " else s = "" end if
'       str = str & b.id & ": " & s & bs & vbnewline
'       'str = str & b.id & ": " & s & bs & "z:" & b.z & vbnewline
'     end if
'   Next
'   if DebugOn then str = "ubound ballvels: " & ubound(ballvel) & vbnewline & str : if TBPin.text <> str then TBPin.text = str : end if
' End Sub
'End Class




'******************************************************
'   Drop Target SUPPORTING FUNCTIONS
'******************************************************

Sub dDropTargetsR_Hit(idx)
  DropsR.DTHit Activeball
End Sub

dim DropsR: Set DropsR = new DTPhysics  'this is just rubber but negative 85%...
DropsR.angle  = sw41vuk.Orientation
DropsR.mass = 0.1

Class DTPhysics
  public angle, mass

  public sub DTHit(aBall)
    dim rangle,bangle,calc1, calc2, calc3
    rangle = (angle - 90) * 3.1416 / 180
    bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

    calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
    calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
    calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

    aBall.velx = calc1 * cos(rangle) + calc2
    aBall.vely = calc1 * sin(rangle) + calc3
  End Sub
End Class

Function Atn2(dy, dx)
  dim pi
  pi = 4*Atn(1)

  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    end if
  ElseIf dx = 0 Then
    if dy = 0 Then
      Atn2 = 0
    else
      Atn2 = Sgn(dy) * pi / 2
    end if
  End If
End Function

dim cor : set cor = New CoRTracker

Class CoRTracker
  public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs

    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds
    if uBound(ballvelx) < highestID then redim ballvelx(highestID)  'set bounds
    if uBound(ballvely) < highestID then redim ballvely(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class




'******************************************************
'   FLIPPER CORRECTION SUPPORTING FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF, LFKS, RFKS)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

'Methods:
'.TimeDelay - Delay before trigger shuts off automatically. Default = 80 (ms)
'.AddPoint - "Polarity", "Velocity", "Ycoef" coordinate points. Use one of these 3 strings, keep coordinates sequential. x = %position on the flipper, y = output
'.Object - set to flipper reference. Optional.
'.StartPoint - set start point coord. Unnecessary, if .object is used.

'Called with flipper -
'ProcessBalls - catches ball data.
' - OR -
'.Fire - fires flipper.rotatetoend automatically + processballs. Requires .Object to be set to the flipper.

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart, FlipperEnd, LR, PartialFlipCoef
  Private Balls(20), balldata(20)

  dim PolarityIn, PolarityOut
  dim VelocityIn, VelocityOut
  dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
    Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
  End Sub

  Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
  Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
  Public Property Get StartPoint : StartPoint = FlipperStart : End Property
  Public Property Let EndPoint(aInput) : if IsObject(aInput) then FlipperEnd = aInput.x else FlipperEnd = aInput : end if : End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property

  Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
    if gametime > 100 then Report aChooseArray
  End Sub

  Public Sub Report(aChooseArray)   'debug, reports all coords in tbPL.text
    if not DebugOn then exit sub
    dim a1, a2 : Select Case aChooseArray
      case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
      Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
      Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut
      case else :tbpl.text = "wrong string" : exit sub
    End Select
    dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    tbpl.text = str
  End Sub

  Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

  Private Sub RemoveBall(aBall)
    dim x : for x = 0 to uBound(balls)
      if TypeName(balls(x) ) = "IBall" then
        if aBall.ID = Balls(x).ID Then
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
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        balldata(x).Data = balls(x)
        if DebugOn then StickL.visible = True : StickL.x = balldata(x).x    'debug TODO
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
    if abs(Flipper.currentAngle - Flipper.EndAngle) < 30 Then
      PartialFlipCoef = 0
    End If
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1
      dim teststr : teststr = "Cutoff"
      tmp = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
      if tmp < 0.1 then 'if real ball position is behind flipper, exit Sub to prevent stucks  'Disabled 1.03, I think it's the Mesh that's causing stucks, not this
        if DebugOn then TestStr = "real pos < 0.1 ( " & round(tmp,2) & ")" : tbpl.text = Teststr
        'RemoveBall aBall
        'Exit Sub
      end if

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        if DebugOn then teststr = "y velocity: " & round(aBall.vely, 3) & "exit sub" : tbpl.text = teststr
        RemoveBall aBall
        exit Sub
      end if
      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          'TB.TEXT = balldata(x).id & " " & BALLDATA(X).X & VBNEWLINE & FLIPPERSTART & " " & FLIPPEREND
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        if DebugOn then set tmp = new spoofball : tmp.data = aBall : End If
        if IsEmpty(BallData(idx).id) and aBall.VelY < -12 then 'if tip hit with no collected data, do vel correction anyway
          if PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1) > 1.1 then 'adjust plz
            VelCoef = LinearEnvelope(5, VelocityIn, VelocityOut)
            if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
            if Enabled then aBall.Velx = aBall.Velx*VelCoef'VelCoef
            if Enabled then aBall.Vely = aBall.Vely*VelCoef'VelCoef
            if DebugOn then teststr = "tip protection" & vbnewline & "velcoef: " & round(velcoef,3) & vbnewline & round(PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1),3) & vbnewline
            'debug.print teststr
          end if
        Else
     :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
          if Enabled then aBall.Velx = aBall.Velx*VelCoef
          if Enabled then aBall.Vely = aBall.Vely*VelCoef
        end if
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
        'playsound "knocker"
      End If
      'debug
      if DebugOn then
        TestStr = teststr & "%pos:" & round(BallPos,2)
        if IsEmpty(PolarityOut(0) ) then
          teststr = teststr & vbnewline & "(Polarity Disabled)" & vbnewline
        else
          teststr = teststr & "+" & round(1 *(AddX*ycoef*PartialFlipcoef),3)
          if BallPos >= PolarityOut(uBound(PolarityOut) ) then teststr = teststr & "(MAX)" & vbnewline else teststr = teststr & vbnewline end if
          if Ycoef < 1 then teststr = teststr &  "ycoef: " & ycoef & vbnewline
          if PartialFlipcoef < 1 then teststr = teststr & "PartialFlipcoef: " & round(PartialFlipcoef,4) & vbnewline
        end if

        teststr = teststr & vbnewline & "Vel: " & round(BallSpeed(tmp),2) & " -> " & round(ballspeed(aBall),2) & vbnewline
        teststr = teststr & "%" & round(ballspeed(aBall) / BallSpeed(tmp),2)
        tbpl.text = TestSTR
      end if
    Else
      'if DebugOn then tbpl.text = "td" & timedelay
    End If
    RemoveBall aBall
  End Sub
End Class

'================================
'Helper Functions


Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
    if not IsEmpty(aArray(x) ) Then
      if IsObject(aArray(x)) then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  if offset < 0 then offset = 0
  redim aArray(aCount-1+offset) 'Resize original array
  for x = 0 to aCount-1   'set objects back into original array
    if IsObject(a(x)) then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub


Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

Function NullFunctionZ(aEnabled):End Function '1 argument null function placeholder  TODO move me or replac eme

Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
      id = .ID : mass = .mass : radius = .radius
    end with
  End Property
  Public Sub Reset()
    x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty
    id = Empty : mass = Empty : radius = Empty
  End Sub
End Class


Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  'Clamp if on the boundry lines
  'if L=1 and Y < yLvl(LBound(yLvl) ) then Y = yLvl(lBound(yLvl) )
  'if L=uBound(xKeyFrame) and Y > yLvl(uBound(yLvl) ) then Y = yLvl(uBound(yLvl) )
  'clamp 2.0
  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 200)*1.2 'was 1.2
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
    VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / KK.width-1
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

function AudioFade(ball)
    Dim tmp
    tmp = ball.y * 2 / KK.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
    BallVelZ = INT((ball.VelZ) * -1 )
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

Function PI()

    PI = 4*Atn(1)

End Function


'*****************************************
'     BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7)

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
        If BOT(b).X < KK.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (KK.Width/2))/7)) '+ 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (KK.Width/2))/7)) '- 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************
Dim BallInKicker1, BallSaucer1
Const tnob = 7 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub CollisionTimer_Timer()
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


               ' ***Ball on WOOD playfield***
            if BOT(b).z < 30 Then
                        StopSound("fx_Rolling_Metal" & b):StopSound("fx_Rolling_Plastic" & b):PlaySound("fx_Rolling_Wood" & b), -1, Vol(BOT(b) )/10, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )

            Else
        ' ***Ball on METAL ramp***
                 If InRect(BOT(b).x, BOT(b).y, 100,345,262,320,270,1385,25,1400) and BOT(b).z > 100 Then
                        StopSound("fx_Rolling_Plastic" & b):StopSound("fx_Rolling_Wood" & b):PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
                 ElseIf InRect(BOT(b).x, BOT(b).y, 0,93,320,93,250,875,40,900) and BOT(b).z > 150 Then
                       StopSound("fx_Rolling_Plastic" & b):StopSound("fx_Rolling_Wood" & b):PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
                ElseIf InRect(BOT(b).x, BOT(b).y, 700,600,950,600,950,1433,680,1433) Then
                       StopSound("fx_Rolling_Plastic" & b):StopSound("fx_Rolling_Wood" & b):PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        ' ***Ball on PLASTIC ramp***
                Else
                        StopSound("fx_Rolling_Metal" & b):StopSound("fx_Rolling_Wood" & b):PlaySound("fx_Rolling_Plastic" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
                End If
            End If



        Else
            If rolling(b) = True Then
                StopSound("fx_Rolling_Wood" & b)
                StopSound("fx_Rolling_Plastic" & b)
                StopSound("fx_Rolling_Metal" & b)
                rolling(b) = False
            End If


        End If

    'Ball Drop Hits

    If BOT(b).VelZ < 0 and BOT(b).z < 90 and BOT(b).z > 80 Then
      If InRect(BOT(b).x,BOT(b).y,(pRI004.x-40),(pRI004.y-40),(pRI004.x+40),(pRI004.y-40),(pRI004.x+40),(pRI004.y+40),(pRI004.x-40),(pRI004.y+40)) Then
        PlaySoundAtBOTBallZ "metalhit_medium", BOT(b)
        debug.print BOT(b).z
      End If
    End If


    If BOT(b).VelZ < 0 and BOT(b).z < 60 and BOT(b).z > 40 Then
      If InRect(BOT(b).x,BOT(b).y,(pRI001.x-100),(pRI001.y-200),(pRI001.x+100),(pRI001.y-200),(pRI001.x+100),(pRI001.y+200),(pRI001.x-100),(pRI001.y+200)) Then
      ElseIf InRect(BOT(b).x,BOT(b).y,(pRampWedge3.x-40),(pRampWedge3.y-140),(pRampWedge3.x+40),(pRampWedge3.y-140),(pRampWedge3.x+40),(pRampWedge3.y+40),(pRampWedge3.x-40),(pRampWedge3.y+40)) Then
      Else
        PlaySoundAtBOTBallZ "Ball_Bounce", BOT(b)
        debug.print BOT(b).z
      End If
    End If

    ' ********* Saucer Code sw29

    If VUKModType = 1 Then

    If InRect(BOT(b).x,BOT(b).y,(pUpKicker.x-30),(pUpKicker.y-30),(pUpKicker.x+30),(pUpKicker.y-30),(pUpKicker.x+30),(pUpKicker.y+30),(pUpKicker.x-30),(pUpKicker.y+30)) then
      If ABS(BOT(b).velx) < 1 and ABS(BOT(b).vely) < 1 Then
        Controller.Switch(29) = 1
        sw29.enabled = true
        BallInKicker1 = 1
        Set BallSaucer1 = BOT(b)
      Else
        BallInKicker1 = 0
      End If
    End If

    End If

    ' ********* End Saucer Code

    Next


End Sub


'**********************
' Ball Collision Sound
'**********************

'Sub OnBallBallCollision(ball1, ball2, velocity)
' PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
'End Sub


'Fix by Roth to fix ball chatter on lock ramp
Sub OnBallBallCollision(ball1, ball2, velocity)
  If velocity > 0.5 then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  End If
End Sub


'******************************************************
'         JP's Sound Routines
'******************************************************

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub MetalWalls_Hit (idx)
  PlaySound "metal_hit1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub MetalRails_Hit (idx)
  PlaySound "metal_hit1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub aRubbers_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, 1
End Sub

Sub RubberWalls_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, 1
End Sub

Sub RubberBands_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, 1
End Sub

Sub Posts_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, .8
End Sub

Sub RubberPosts_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, .8
End Sub


Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 2
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 2
End Sub

Sub RightFlipper2_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 2
End Sub

Sub LeftFlipperKS_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 2
End Sub

Sub RightFlipperKS_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 2
End Sub

Sub plastic_Hit(idx)
    PlaySoundAtBallVol "plastichit" & Int(Rnd*1)+1, 1
End Sub



'''''''''''''''''''''''''''other sounds

'''''''Random Metal and plastic sounds

Sub RandomMetalHitSound()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtBallVol "metal_hit1",1
    Case 2 : PlaySoundAtBallVol "metal_hit2",1
    Case 3 : PlaySoundAtBallVol "metal_hit3",1
  End Select
End Sub

Sub RandomPlasticHitSound()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtBallVol "plastic_hit1",1
    Case 2 : PlaySoundAtBallVol "plastic_hit2",1
    Case 3 : PlaySoundAtBallVol "plastic_hit2",1
  End Select
End Sub


'''''''Subway sounds

Sub RandomTopSubwayEnterSound()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtBallVol "topsubwayenter1",1
    Case 2 : PlaySoundAtBallVol "topsubwayenter2",1
    Case 3 : PlaySoundAtBallVol "topsubwayenter3",1
    Case 4 : PlaySoundAtBallVol "topsubwayenter4",1
  End Select
End Sub


'**************************************************************************
'                 Positional Sound Playback Functions by DJRobX
'**************************************************************************

'Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(sound, tableobj)
    PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'Set all as per ball position & speed.

Sub PlaySoundAtBall(sound)
    PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub


'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Vol)
    PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
    PlaySound sound, 0, VolZ(BOT), Pan(BOT), 0, Pitch(BOT), 1, 1, AudioFade(BOT)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
    PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub


'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
    PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*****************************
'Random Ramp and Orbit Sounds
'*****************************

Dim NextOrbitHit:NextOrbitHit = 0

Sub WireRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump3 .1, Pitch(ActiveBall)+5
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .4 + (Rnd * .2)
  end if
End Sub

Sub PlasticRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .4 and Timer > NextOrbitHit then
    RandomBump 5, Pitch(ActiveBall)
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .2 + (Rnd * .2)
  end if
End Sub


Sub MetalGuideBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump2 2, Pitch(ActiveBall)
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .2 + (Rnd * .2)
  end if
End Sub

Sub MetalWallBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump 2, 20000 'Increased pitch to simulate metal wall
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .2 + (Rnd * .2)
  end if
End Sub


' Requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
  dim BumpSnd:BumpSnd= "fx_rampbump" & CStr(Int(Rnd*7)+1)
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Requires metalguidebump1 to 2 in Sound Manager
Sub RandomBump2(voladj, freq)
  dim BumpSnd:BumpSnd= "metalguidebump" & CStr(Int(Rnd*2)+1)
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Requires WireRampBump1 to 5 in Sound Manage
Sub RandomBump3(voladj, freq)
  dim BumpSnd:BumpSnd= "WireRampBump" & CStr(Int(Rnd*5)+1)
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub


' Stop Bump Sounds
Sub BumpSTOP1_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOP2_Hit ()
dim i:for i=1 to 4:StopSound "fx_rampbump" & i:next
NextOrbitHit = Timer + 1
End Sub



' ***************************************************************************
'                    BASIC FSS(DMD,SS,EM) SETUP CODE
' ****************************************************************************

' add next two lines to end of init subroutine, uncomment those that apply.
'For each nxx in aBGLights:nxx.visible = 0:Next
'setup_backglass()




Const cUSESHADOW = 1
Const cUSECOLORGRADE = 1
Const cUSECCUPBOOST = 1
Const cUSELITEBOOST = 0
Const cUSEBACKGLASS = 1
Const cDYNPINBALL = 2 ' 1:(inverted) easy to play light to Dark, 2: (normal) ball dark to light matches table
Const cDYNGAMEBLADES =0
Const cSHADOWBLADES =0
Const cUSEBGLASSHIGH =0
Const cUSETOPPER =0
Const cUSEBACKLIGHT =1
Const cUSEBGLASSCCX =1
Const cUSEBGLASSREFL =1
Const cUSETRANSLIST =1 ' manage transparent light to dark object list
Const cUSEOPAQUELIST =1 ' 0=OFF, 1=ON, 2=switched(ON/OFF) opaque transparent light to dark object list
Const cUSESEMITRANSLIST =1 ' manage semi-transparent light to dark object list





const USEEM = 0 ' 0=no EM, 1 = EM enabled, 2 = EM SysGFX enabled ,3=Reels coded, 4 = Start SysGFX Timer
const USESS = 2 ' 0=no SS, 1 = SS enabled, 2 = SS LED's
const nUSEDMD =0 ' 0=no DMD, 1= DMD 2= DMD +mirror DMD

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen

sub setup_backglass()

xoff =473
yoff = 100
zoff =970
xrot = -90

if cUSEBACKGLASS then
bgdark.x = xoff
bgdark.y = yoff
bgdark.height = zoff
bgdark.rotx = xrot

bgHigh.x = xoff
bgHigh.y = yoff
bgHigh.height = zoff
bgHigh.rotx = xrot

bgHigh1.x = xoff
bgHigh1.y = yoff
bgHigh1.height = zoff
bgHigh1.rotx = xrot

if 0 then
bgGrill.x = xoff
bgGrill.y = yoff
bgGrill.height = zoff
bgGrill.rotx = xrot
end if

if 1  then
BGHigh2.x = xoff
BGHigh2.y = yoff
BGHigh2.height = zoff
BGHigh2.rotx = xrot
end If
'
'
'bgFrame.x = xoff
'bgFrame.y = yoff
'bgFrame.height = zoff +40
'bgFrame.rotx = xrot
'
'if 1 then
'BGFrameMask.x = xoff
'BGFrameMask.y = yoff
'BGFrameMask.height = zoff +40
'BGFrameMask.rotx = xrot
'
'
'BGFrameMaskFill.x = xoff
'BGFrameMaskFill.y = yoff
'BGFrameMaskFill.height = zoff
'BGFrameMaskFill.rotx = xrot
'end if
'
'if cUSEBGLASSREFL then
'
'' the mask fill image
'if NOT IsEmpty(tMaskFill(0)) then
'BGFrameMaskFill1.x = xoff
'BGFrameMaskFill1.y = yoff
'BGFrameMaskFill1.height = zoff+0
'BGFrameMaskFill1.rotx = xrot
'end If
'' the background reflective image usually __default_screen_space_reflection.png
'if NOT IsEmpty(tMaskFill(4)) then
'BGFrameMaskFill5.x = xoff
'BGFrameMaskFill5.y = yoff
'BGFrameMaskFill5.height = zoff-100 ' adjust value to correct vertical
'BGFrameMaskFill5.rotx = xrot
'end if
'end if 'cUSEBGLASSREFL
'
'
'if cUSEBGLASSCCX then
'FlBGRedGreenSpec.x = xoff
'FlBGRedGreenSpec.y = yoff
'FlBGRedGreenSpec.height = zoff
'FlBGRedGreenSpec.rotx = xrot
'
'FlBGBlueSpec.x = xoff
'FlBGBlueSpec.y = yoff
'FlBGBlueSpec.height = zoff
'FlBGBlueSpec.rotx = xrot
'end if ' cUSEBGLASSCCX
end if  'cUSEBACKGLASS

' ********************************
'           TOPPER SCRIPT
' ********************************

if cUSETOPPER then

Dim topoff

' the topper
  topoff = 590
If Table1.inclination >= 65.0 then

  TopDark.x = xoff
  TopDark.y = yoff
  TopDark.height = zoff +topoff
  TopDark.rotx = xrot

  TopHigh.x = xoff
  TopHigh.y = yoff
  TopHigh.height = zoff +topoff
  TopHigh.rotx = xrot

  TopHigh1.x = xoff
  TopHigh1.y = yoff
  TopHigh1.height = zoff +topoff
  TopHigh1.rotx = xrot


  TopHigh2.x = xoff
  TopHigh2.y = yoff
  TopHigh2.height = zoff +(topoff)
  TopHigh2.rotx = xrot

else
end if
end if ' cUSETOPPER

  if nUSEDMD >= 1 then
  DMD.x = xoff
  DMD.y = yoff
  DMD.height = zoff - 310
  DMD.rotx = xrot

  'DMD mirror
    if nUSEDMD > 1 then
    DMD1.x = xoff
    DMD1.y = yoff + 250
    DMD1.height = zoff - 500
    DMD1.rotx = xrot +180
    DMD1.visible =1
    end if
  end if 'nUSEDMD

center_graphix()

if USESS > 1 then
center_digits()
end if

  if USEEM > 0 then
  if  USEEM > 1 then
  center_sysgfx()
  end if


  if USESS = 0 then
  'center_objects() 'always calllast
  'or
  center_objects_em()

  if  USEEM > 2 then
  For  ix =0 to 4
  SetDrum ix, 1 , 0
  SetDrum ix, 2 , 0
  SetDrum ix, 3 , 0
  SetDrum ix, 4 , 0
  SetDrum ix, 5 , 0
  Next

  cred =reels(4, 0)
  reels(4, 0) = 0
  SetDrum -1,0,  0

  SetReel 0,-1,  Credit
  reels(4, 0) = Credit
  end if 'USEEM > 2
  end if 'USESS = 0

  if  USEEM > 3 then
  EMReelTimer.enabled = 1
  end if 'USEEM > 3

  end if 'USEEM

end sub

' ********************* POSITION IMAGES(flashers) ON BACKGLASS *************************





Dim BGArr

If BackGlassVisible = 1 then

BGArr=Array (BGSol11, BGSolKong, BGSolGirl, BGSol49, BGSol50, BGSol51, BGSol52, BGSol53, BGSol54, BGSolDNSRS, BGSolSkull, BGSolKingShooter, BGSolKongShooter, BGSolTower25, BGSolTower50, BGSolTower100, BGSolTowerEB, BGSolTowerMil, BGLight55 )


Else

BGArr=Array ()


End If

Sub center_graphix()
Dim xx,yy,yfact,xfact,xobj
zscale = 0.0000001

xcen =(1050 /2) - (73 / 2)
ycen = (969 /2 ) + (177 /2)


yfact =0 'y fudge factor (ycen was wrong so fix)
xfact =0

  For Each xobj In BGArr
  if Not IsEmpty(xobj) then
  xx =xobj.x

  xobj.x = (xoff -xcen) + xx +xfact
  yy = xobj.y ' get the yoffset before it is changed
  xobj.y =yoff

    If(yy < 0.) then
    yy = yy * -1
    end if


  xobj.height =( zoff - ycen) + yy - (yy * zscale) + yfact

  xobj.rotx = xrot
  xobj.visible =1 ' for testing
  end if
  Next
end sub

' ***************************************************************************
'       BASIC FSS(SS TYPE0) 2x16 solid state character display SETUP CODE
' ****************************************************************************

Sub center_digits()
Dim ix, xx, yy, yfact, xfact, xobj

zscale = 0.0000001

xcen =(1050 /2) - (73 / 2)
ycen = (969 /2 ) + (177 /2)

yfact =0 'y fudge factor (ycen was wrong so fix)
xfact =0


for ix =0 to 31
  For Each xobj In LEDFSS(ix)

  'if obj NOT n then

  xx =xobj.x

  xobj.x = (xoff -xcen) + xx +xfact
  yy = xobj.y ' get the yoffset before it is changed
  xobj.y =yoff

    If(yy < 0.) then
    yy = yy * -1
    end if

  xobj.height =( zoff - ycen) + yy - (yy * (zscale)) + yfact

  xobj.rotx = xrot

  'xobj.visible = 0
  'end if
  Next
  Next
end sub


Sub hide_elem(elem)
Dim objx

  For Each objx In elem ' hide all elements
  objx.visible =0
  Next
end sub


Sub show_elem(elem)
Dim objx

  For Each objx In elem ' hide all elements
  objx.visible =1
  Next
end sub


 Dim LEDFSS(32)
 LEDFSS(0)=Array(ax00, ax05, ax0c, ax0d, ax08, ax01, ax06, ax0f, ax02, ax03, ax04, ax07, ax0b, ax0a, ax09, ax0e)
 LEDFSS(1)=Array(ax10, ax15, ax1c, ax1d, ax18, ax11, ax16, ax1f, ax12, ax13, ax14, ax17, ax1b, ax1a, ax19, ax1e)
 LEDFSS(2)=Array(ax20, ax25, ax2c, ax2d, ax28, ax21, ax26, ax2f, ax22, ax23, ax24, ax27, ax2b, ax2a, ax29, ax2e)
 LEDFSS(3)=Array(ax30, ax35, ax3c, ax3d, ax38, ax31, ax36, ax3f, ax32, ax33, ax34, ax37, ax3b, ax3a, ax39, ax3e)
 LEDFSS(4)=Array(ax40, ax45, ax4c, ax4d, ax48, ax41, ax46, ax4f, ax42, ax43, ax44, ax47, ax4b, ax4a, ax49, ax4e)
 LEDFSS(5)=Array(ax50, ax55, ax5c, ax5d, ax58, ax51, ax56, ax5f, ax52, ax53, ax54, ax57, ax5b, ax5a, ax59, ax5e)
 LEDFSS(6)=Array(ax60, ax65, ax6c, ax6d, ax68, ax61, ax66, ax6f, ax62, ax63, ax64, ax67, ax6b, ax6a, ax69, ax6e)
 LEDFSS(7)=Array(ax70, ax75, ax7c, ax7d, ax78, ax71, ax76, ax7f, ax72, ax73, ax74, ax77, ax7b, ax7a, ax79, ax7e)
 LEDFSS(8)=Array(ax80, ax85, ax8c, ax8d, ax88, ax81, ax86, ax8f, ax82, ax83, ax84, ax87, ax8b, ax8a, ax89, ax8e)
 LEDFSS(9)=Array(ax90, ax95, ax9c, ax9d, ax98, ax91, ax96, ax9f, ax92, ax93, ax94, ax97, ax9b, ax9a, ax99, ax9e)
 LEDFSS(10)=Array(axa0, axa5, axac, axad, axa8, axa1, axa6, axaf, axa2, axa3, axa4, axa7, axab, axaa, axa9, axae)
 LEDFSS(11)=Array(axb0, axb5, axbc, axbd, axb8, axb1, axb6, axbf, axb2, axb3, axb4, axb7, axbb, axba, axb9, axbe)
 LEDFSS(12)=Array(axc0, axc5, axcc, axcd, axc8, axc1, axc6, axcf, axc2, axc3, axc4, axc7, axcb, axca, axc9, axce)
 LEDFSS(13)=Array(axd0, axd5, axdc, axdd, axd8, axd1, axd6, axdf, axd2, axd3, axd4, axd7, axdb, axda, axd9, axde)
 LEDFSS(14)=Array(axe0, axe5, axec, axed, axe8, axe1, axe6, axef, axe2, axe3, axe4, axe7, axeb, axea, axe9, axee)
 LEDFSS(15)=Array(axf0, axf5, axfc, axfd, axf8, axf1, axf6, axff, axf2, axf3, axf4, axf7, axfb, axfa, axf9, axfe)

 LEDFSS(16)=Array(bx00, bx05, bx0c, bx0d, bx08, bx01, bx06, bx0f, bx02, bx03, bx04, bx07, bx0b, bx0a, bx09, bx0e)
 LEDFSS(17)=Array(bx10, bx15, bx1c, bx1d, bx18, bx11, bx16, bx1f, bx12, bx13, bx14, bx17, bx1b, bx1a, bx19, bx1e)
 LEDFSS(18)=Array(bx20, bx25, bx2c, bx2d, bx28, bx21, bx26, bx2f, bx22, bx23, bx24, bx27, bx2b, bx2a, bx29, bx2e)
 LEDFSS(19)=Array(bx30, bx35, bx3c, bx3d, bx38, bx31, bx36, bx3f, bx32, bx33, bx34, bx37, bx3b, bx3a, bx39, bx3e)
 LEDFSS(20)=Array(bx40, bx45, bx4c, bx4d, bx48, bx41, bx46, bx4f, bx42, bx43, bx44, bx47, bx4b, bx4a, bx49, bx4e)
 LEDFSS(21)=Array(bx50, bx55, bx5c, bx5d, bx58, bx51, bx56, bx5f, bx52, bx53, bx54, bx57, bx5b, bx5a, bx59, bx5e)
 LEDFSS(22)=Array(bx60, bx65, bx6c, bx6d, bx68, bx61, bx66, bx6f, bx62, bx63, bx64, bx67, bx6b, bx6a, bx69, bx6e)
 LEDFSS(23)=Array(bx70, bx75, bx7c, bx7d, bx78, bx71, bx76, bx7f, bx72, bx73, bx74, bx77, bx7b, bx7a, bx79, bx7e)
 LEDFSS(24)=Array(bx80, bx85, bx8c, bx8d, bx88, bx81, bx86, bx8f, bx82, bx83, bx84, bx87, bx8b, bx8a, bx89, bx8e)
 LEDFSS(25)=Array(bx90, bx95, bx9c, bx9d, bx98, bx91, bx96, bx9f, bx92, bx93, bx94, bx97, bx9b, bx9a, bx99, bx9e)
 LEDFSS(26)=Array(bxa0, bxa5, bxac, bxad, bxa8, bxa1, bxa6, bxaf, bxa2, bxa3, bxa4, bxa7, bxab, bxaa, bxa9, bxae)
 LEDFSS(27)=Array(bxb0, bxb5, bxbc, bxbd, bxb8, bxb1, bxb6, bxbf, bxb2, bxb3, bxb4, bxb7, bxbb, bxba, bxb9, bxbe)
 LEDFSS(28)=Array(bxc0, bxc5, bxcc, bxcd, bxc8, bxc1, bxc6, bxcf, bxc2, bxc3, bxc4, bxc7, bxcb, bxca, bxc9, bxce)
 LEDFSS(29)=Array(bxd0, bxd5, bxdc, bxdd, bxd8, bxd1, bxd6, bxdf, bxd2, bxd3, bxd4, bxd7, bxdb, bxda, bxd9, bxde)
 LEDFSS(30)=Array(bxe0, bxe5, bxec, bxed, bxe8, bxe1, bxe6, bxef, bxe2, bxe3, bxe4, bxe7, bxeb, bxea, bxe9, bxee)
 LEDFSS(31)=Array(bxf0, bxf5, bxfc, bxfd, bxf8, bxf1, bxf6, bxff, bxf2, bxf3, bxf4, bxf7, bxfb, bxfa, bxf9, bxfe)

 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then

       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      'if (num < 32) then
      if (num < 32) then
              For Each obj In LEDFSS(num)
                   If chg And 1 Then obj.visible=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
      Else
             end if
        Next

    End If
 End Sub







'''''''''dup3d



'Switches:
'1 Plumb Tilt
'3 Credit Button
'4 Right Coin
'5 Center Coin
'6 Left Coin
'7 Tilt Slam
'8 Ticket
'10 Outhole
'11 Trough Left
'12 Trough Center
'13 Trough Right
'14 Shooter Lane
'15 Left Flipper
'16 Right Flipper
'17 Left Outlane
'18 Left Return
'19 Right Outlane
'20 Right Return
'21 Left Sling
'22 Right Sling
'23 Wacker Drive
'24
'25 Ape Lane A
'26 Ape Lane P
'27 Ape Lane E
'28 Ape Lane Ramp
'29 Lock VUK
'30 Missile Kicker?
'31 Loop Left
'32 Loop Right
'33 Tower T
'34 Tower o
'35 Tower W
'36 Tower E
'37 Tower R
'38 Lock Bottom
'39 Lock Center
'40 Lock Top
'41 DT Left
'42 DT Center
'43 DT Right
'44 Not Used - Million Banana's target - owner of proto 9
'45 Radar Eject
'46 Bumper Left
'47 Bumper Center
'48 Bumper Right
'49 Cubs Target
'50 Bears Target
'51 Bulls Target
'
'Lamps:
'1 King K
'2 King I
'3 King N
'4 King G
'5 Kong K
'6 Kong O
'7 Kong N
'8 Kong G
'9 Tower T
'10 Tower O
'11 Tower W
'12 Towet E
'13 Tower R
'14 Cubs
'15 Bears
'16 Bulls
'17 DT Left
'18 DT Center
'19 DT Right
'20 Not Used
'21 Tower 25000
'22 Tower 50000
'23 Tower 100000
'24 Tower Ex.Ball
'25 Lock Red
'26 Lock Yellow
'27 Lock Green
'''''''''''''28 City Bonus
'29 Ape Ramp 25000
'30 Ape Ramp 50000
'31 Ape Ramp 100000
'32 Million Bananas
'33 2x
'34 3x
'35 4x
'36 5x
'37 Go Ape Again
'38 Eject Ex. Ball
'39 Everything Lit
''''''''''''''''40 Jungle Bonus
'41 Loop 25000
'42 Loop 50000
'43 Loop 100000
'44 Loop Ex.Ball
'45 S Replay Arrow
'46 Not Used
'47 Tower 3000000
'''''''''''''''48 Special Left
'49 Jackpot 1000000
'50 Jackpot 2000000
'51 Jackpot 3000000
'52 Jackpot 4000000
'53 Jackpot 5000000
'54 Jackpot S Replay
'55 Not Used
''''''''''''''''''56 Special Right
'57 Missile 25000
'58 Missile 50000
'59 Missile 75000
'60 Missile 100000
'61 Missile 1000000
'62 Ape Lane A
'63 Ape Lane P
'64 Ape Lane E
'
'Solenoids:
'1 King Shooter
'2 Kong Shooter
'3 Ape Ramp BG Skul
'4 BG DNSRS Shooter
'5 BG Tower 25000
'6 BG Tower 50000
'7 BG Tower 100000
'8 BG Tower Ex. Ball
'9 BG Tower Million
'11 Gen Illm
'12 BG Girl Missile
'13 BG Kong Explod??e
'14 Tower Flasher
'15 Skull Island
'16 Missile Launch
'17 Bumper Left
'18 Bumper Center
'19 Bumper Right
'20 Sling Left
'21 Sling Right
'25 Outhole
'26 Shooter Eject
'27 Lock Release
'28 Radar Eject
'29 Ball Lock
'30 Drop Target
'32 Ticket
'45 Flipper RightU
'46 Flipper Right
'47 Flipper LeftU
'48 Flipper Left
