'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////

' SpongeBob's Bikini Bottom Pinball 2.0

'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////
'
'
' 1.0 Skillshot
' 2.0 VPW team
'
' VPW TEAM
' Original author: Skillshot
' Modded from JPSalas' AFM
' Blender work: Gedankekojote
' 2d art: Astronasty
' Coding: Oqqsan
' Vr room: Rawd
' Toolkit scripting: Apophis
' Light controller: Flux
' Sound update: Friscopinball
' Audio fixes: MrGrynch
' Backglass stuff: Leojreimroc, Remdwaas
' Table cleanup and support: Sixtoe
' Testing: VPW jellyfishuP
' Special thanks to Niwak.
'


Option Explicit
Randomize
SetLocale 1033
Const TableVersion = "2.1"

'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////
'// OPTIONS
'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////


' rest of options is on the Option DMD ingame, use both magnabuttons to start.
'                                       ^^^^
'// Music
Const SongVolume = 1          ' to just dampen all music playing and nothing else .... range 0=off 1=max

'// VR-Room
VRRoomchoice = 1              ' 0 = Desktop/FS  1 = SpongeBob Room  2 = Minimal room  3 = Ultra-Minimal Room
Scratches = 0               ' 0 = Glass scratches OFF  1 = Glass scratches ON   'Only works on VRroom 1 or 2
VR_Backglass = 0            ' 0 = Main Backglass, 1 = Alt Backglass

'// DMD
Dim FlexOnFlasher : FlexOnFlasher = 0 ' 1 = Force VR DMD to be on ! Vrroom turns this on by auto

Dim FlexDMDScorbit : FlexDMDScorbit = Null
Dim FlexDMDScorbitClaim : FlexDMDScorbitClaim = Null


' FOrce DMD on flasher for GL + flasher DMD   1=on 0=off ' vrroom overrides this
' Freezy is hiding on the latest GL versions of vpx, temp fix here ! needed this workaround for GL 10.8 1483
Const ForceDMD_Flasher = 0
' ^^^for desktop mostly


'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////
'// Initializations
'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////



Const BallSize = 50
Const BallMass = 1

'/////////////////////-----Scorbit Options-----////////////////////
dim TablesDir : TablesDir = GetTablesFolder

Const     ScorbitAlternateUUID  = 0   ' Force Alternate UUID from Windows Machine and saves it in VPX Users directory (C:\Visual Pinball\User\ScorbitUUID.dat)


ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0


Dim DesktopMode, CabMode, VRRoom, VRRoomChoice, Scratches, VR_Backglass
DesktopMode = Table1.ShowDT
If RenderingMode = 2 Then VRRoom = VRRoomChoice Else VRRoom = 0
If Not DesktopMode And VRRoom=0 Then CabMode = 1 Else CabMode = 0



'If DesktopMode = False Or B2SonDesktop = 1 Then
' BallsText.Y = BallsText.Y - 500
' DMDText.Y = DMDText.Y - 500
' ScoreText.Y = ScoreText.Y - 500
' ScoreText2.Y = ScoreText2.Y - 500
' ScoreText3.Y = ScoreText3.Y - 500
' ScoreText4.Y = ScoreText4.Y - 500
'End If
'Const B2SonDesktop  =  0            '0 - No B2S in Dekstop mode, 1 - Use B2S in Dekstop mode
' use ingame options to set these ( changing here will not do anything )
Dim RoomBrightness : RoomBrightness =  70         'Room brightness - Value between 0 and 100 (0=Dark ... 100=Bright)
Dim MechVolume : MechVolume = 0.8    ' Level of mechanical volume.  Value between 0 and 1
Dim BallRollVolume : BallRollVolume = 0.5         'Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5       'Level of ramp rolling volume. Value between 0 and 1



Dim Controller
' Sub startcontroller
'   If vrroom < 1 Then
'   On Error Resume Next
'     Set Controller = CreateObject("B2S.Server")
'   If Controller Is Nothing Then
'     Err.Clear
'     B2SOn = False
'   Else
'     Controller.B2SName = "Spongebob"
'     Controller.Run
'     On Error Goto 0
'     B2SOn = True
'   End If
'   Else
'     B2SOn = False
'   End If
'End Sub

Sub startB2S(aB2S)
    If B2SOn Then
    Controller.B2SSetData 1,0
    Controller.B2SSetData 2,0
    Controller.B2SSetData 3,0
    Controller.B2SSetData 4,0
    Controller.B2SSetData aB2S,1
    End If
End Sub

sub b2snormal_timer
    startB2S(1)
    b2snormal.enabled=false
end sub


Sub DOF(DOFevent, State)
  If B2SOn Then
    If State = 2 Then
      Controller.B2SSetData DOFevent, 1:Controller.B2SSetData DOFevent, 0
    Else
      Controller.B2SSetData DOFevent, State
    End If
  End If
End Sub

'DOF id
'E101 0/1 LeftFlipper
'E102 0/1 RightFlipper
'E103 2 Leftslingshot
'E105 2 Rightslingshot
'E107 2 Leftslingshot1
'E108 2 Bumper3
'E109 2 Bumper1
'E111 2 Ballrelease & DrainShooter
'E112 2 Right standup targets hit
'E116 2 Middle standup and drop targets hit
'E117 2 Left standup targets hit
'E118 2 Right scoop eject & Right VUK eject

'E120 0/1 GI On
'E121 0/1 Startgamebutton ready
'E122 0/1 Ball in plunger lane
'E123 2 Plunger fired
'E124 2 drain Hit
'E125 2 ballsaved
'E126 2 Left Spinner
'E127 2 Gate2 hit
'E128 2 Gate3 hit
'E129 2 award extraball
'E130 2 mode start 1-9 and wizard
'E131 0/1 Multiball
'E132 2 Jackpot
'E133 0/1 Do Or Die mode on
'E134 2 shaker for plankton and wiz jackpots
'E135 2 Sponge blinker
'E136 2 Party blinker

'E141 2 Flasher bottom left (f141)
'E142 2 Flasher bottom right (f142)
'E143 2 Flasher top left (f143)
'E144 2 Flasher top (f144)
'E145 2 Flasher top  (f145)
'E146 2 Flasher right  (f146)
'E147 2 Flasher plunger (f147)



'//////////////////////////////////////////////////////////////////////
'// VLM Arrays
'//////////////////////////////////////////////////////////////////////


' ===============================================================
' The following code can be copy/pasted to have premade array for
' movable objects:
' - _LM suffixed arrays contains the lightmaps
' - _BM suffixed arrays contains the bakemap
' - _BL suffixed arrays contains both the bakemap & the lightmaps
Dim BR1_LM: BR1_LM=Array(BR1_LM_Flashers_L110, BR1_LM_Flashers_f143, BR1_LM_Flashers_f144, BR1_LM_Flashers_f145, BR1_LM_GIString11, BR1_LM_GIString5, BR1_LM_GIString6, BR1_LM_Spotlights) ' VLM.Array;LM;BR1
Dim BR1_BM: BR1_BM=Array(BR1_BM_Lit_Room) ' VLM.Array;BM;BR1
Dim BR1_BL: BR1_BL=Array(BR1_BM_Lit_Room, BR1_LM_Flashers_L110, BR1_LM_Flashers_f143, BR1_LM_Flashers_f144, BR1_LM_Flashers_f145, BR1_LM_GIString11, BR1_LM_GIString5, BR1_LM_GIString6, BR1_LM_Spotlights) ' VLM.Array;BL;BR1
Dim BR3_LM: BR3_LM=Array(BR3_LM_Flashers_f143, BR3_LM_Flashers_f144, BR3_LM_GIString11, BR3_LM_GIString5, BR3_LM_GIString6, BR3_LM_Inserts_L35, BR3_LM_Inserts_L38, BR3_LM_Spotlights) ' VLM.Array;LM;BR3
Dim BR3_BM: BR3_BM=Array(BR3_BM_Lit_Room) ' VLM.Array;BM;BR3
Dim BR3_BL: BR3_BL=Array(BR3_BM_Lit_Room, BR3_LM_Flashers_f143, BR3_LM_Flashers_f144, BR3_LM_GIString11, BR3_LM_GIString5, BR3_LM_GIString6, BR3_LM_Inserts_L35, BR3_LM_Inserts_L38, BR3_LM_Spotlights) ' VLM.Array;BL;BR3
Dim Diverter1_LM: Diverter1_LM=Array(Diverter1_LM_Flashers_f143, Diverter1_LM_Flashers_f144, Diverter1_LM_Flashers_f145, Diverter1_LM_Flashers_f146, Diverter1_LM_GIString5, Diverter1_LM_GIString6, Diverter1_LM_Inserts_L34, Diverter1_LM_Inserts_L35, Diverter1_LM_Inserts_L39, Diverter1_LM_Spotlights) ' VLM.Array;LM;Diverter1
Dim Diverter1_BM: Diverter1_BM=Array(Diverter1_BM_Lit_Room) ' VLM.Array;BM;Diverter1
Dim Diverter1_BL: Diverter1_BL=Array(Diverter1_BM_Lit_Room, Diverter1_LM_Flashers_f143, Diverter1_LM_Flashers_f144, Diverter1_LM_Flashers_f145, Diverter1_LM_Flashers_f146, Diverter1_LM_GIString5, Diverter1_LM_GIString6, Diverter1_LM_Inserts_L34, Diverter1_LM_Inserts_L35, Diverter1_LM_Inserts_L39, Diverter1_LM_Spotlights) ' VLM.Array;BL;Diverter1
Dim Diverter2_LM: Diverter2_LM=Array(Diverter2_LM_Flashers_f143, Diverter2_LM_Flashers_f145, Diverter2_LM_Flashers_f146, Diverter2_LM_GIString5, Diverter2_LM_GIString6, Diverter2_LM_GIString7, Diverter2_LM_Inserts_L30, Diverter2_LM_Inserts_L31, Diverter2_LM_Inserts_L32, Diverter2_LM_Inserts_L33, Diverter2_LM_Inserts_L34, Diverter2_LM_Inserts_L35, Diverter2_LM_Inserts_L36, Diverter2_LM_Inserts_L37, Diverter2_LM_Inserts_L38, Diverter2_LM_Inserts_L39, Diverter2_LM_Inserts_L49, Diverter2_LM_Inserts_L50, Diverter2_LM_Spotlights) ' VLM.Array;LM;Diverter2
Dim Diverter2_BM: Diverter2_BM=Array(Diverter2_BM_Lit_Room) ' VLM.Array;BM;Diverter2
Dim Diverter2_BL: Diverter2_BL=Array(Diverter2_BM_Lit_Room, Diverter2_LM_Flashers_f143, Diverter2_LM_Flashers_f145, Diverter2_LM_Flashers_f146, Diverter2_LM_GIString5, Diverter2_LM_GIString6, Diverter2_LM_GIString7, Diverter2_LM_Inserts_L30, Diverter2_LM_Inserts_L31, Diverter2_LM_Inserts_L32, Diverter2_LM_Inserts_L33, Diverter2_LM_Inserts_L34, Diverter2_LM_Inserts_L35, Diverter2_LM_Inserts_L36, Diverter2_LM_Inserts_L37, Diverter2_LM_Inserts_L38, Diverter2_LM_Inserts_L39, Diverter2_LM_Inserts_L49, Diverter2_LM_Inserts_L50, Diverter2_LM_Spotlights) ' VLM.Array;BL;Diverter2
Dim Gate2_LM: Gate2_LM=Array() ' VLM.Array;LM;Gate2
Dim Gate2_BM: Gate2_BM=Array(Gate2_BM_Lit_Room) ' VLM.Array;BM;Gate2
Dim Gate2_BL: Gate2_BL=Array(Gate2_BM_Lit_Room) ' VLM.Array;BL;Gate2
Dim Gate3_LM: Gate3_LM=Array(Gate3_LM_Spotlights) ' VLM.Array;LM;Gate3
Dim Gate3_BM: Gate3_BM=Array(Gate3_BM_Lit_Room) ' VLM.Array;BM;Gate3
Dim Gate3_BL: Gate3_BL=Array(Gate3_BM_Lit_Room, Gate3_LM_Spotlights) ' VLM.Array;BL;Gate3
Dim LEMK1_LM: LEMK1_LM=Array(LEMK1_LM_Flashers_f143, LEMK1_LM_GIString6) ' VLM.Array;LM;LEMK1
Dim LEMK1_BM: LEMK1_BM=Array(LEMK1_BM_Lit_Room) ' VLM.Array;BM;LEMK1
Dim LEMK1_BL: LEMK1_BL=Array(LEMK1_BM_Lit_Room, LEMK1_LM_Flashers_f143, LEMK1_LM_GIString6) ' VLM.Array;BL;LEMK1
Dim LFMesh_LM: LFMesh_LM=Array(LFMesh_LM_Flashers_f141, LFMesh_LM_Flashers_f148, LFMesh_LM_GIString1_gi001, LFMesh_LM_GIString1_gi005, LFMesh_LM_GIString1_gi26, LFMesh_LM_GIString2_gi002, LFMesh_LM_GIString2_gi004, LFMesh_LM_Inserts_L01, LFMesh_LM_Inserts_L04) ' VLM.Array;LM;LFMesh
Dim LFMesh_BM: LFMesh_BM=Array(LFMesh_BM_Lit_Room) ' VLM.Array;BM;LFMesh
Dim LFMesh_BL: LFMesh_BL=Array(LFMesh_BM_Lit_Room, LFMesh_LM_Flashers_f141, LFMesh_LM_Flashers_f148, LFMesh_LM_GIString1_gi001, LFMesh_LM_GIString1_gi005, LFMesh_LM_GIString1_gi26, LFMesh_LM_GIString2_gi002, LFMesh_LM_GIString2_gi004, LFMesh_LM_Inserts_L01, LFMesh_LM_Inserts_L04) ' VLM.Array;BL;LFMesh
Dim LGate_LM: LGate_LM=Array(LGate_LM_GIString6, LGate_LM_GIString7, LGate_LM_Spotlights) ' VLM.Array;LM;LGate
Dim LGate_BM: LGate_BM=Array(LGate_BM_Lit_Room) ' VLM.Array;BM;LGate
Dim LGate_BL: LGate_BL=Array(LGate_BM_Lit_Room, LGate_LM_GIString6, LGate_LM_GIString7, LGate_LM_Spotlights) ' VLM.Array;BL;LGate
Dim LeftSling12_LM: LeftSling12_LM=Array(LeftSling12_LM_Flashers_f143, LeftSling12_LM_GIString11, LeftSling12_LM_GIString6, LeftSling12_LM_Spotlights) ' VLM.Array;LM;LeftSling12
Dim LeftSling12_BM: LeftSling12_BM=Array(LeftSling12_BM_Lit_Room) ' VLM.Array;BM;LeftSling12
Dim LeftSling12_BL: LeftSling12_BL=Array(LeftSling12_BM_Lit_Room, LeftSling12_LM_Flashers_f143, LeftSling12_LM_GIString11, LeftSling12_LM_GIString6, LeftSling12_LM_Spotlights) ' VLM.Array;BL;LeftSling12
Dim LeftSling13_LM: LeftSling13_LM=Array(LeftSling13_LM_Flashers_f143, LeftSling13_LM_GIString11, LeftSling13_LM_GIString6, LeftSling13_LM_Spotlights) ' VLM.Array;LM;LeftSling13
Dim LeftSling13_BM: LeftSling13_BM=Array(LeftSling13_BM_Lit_Room) ' VLM.Array;BM;LeftSling13
Dim LeftSling13_BL: LeftSling13_BL=Array(LeftSling13_BM_Lit_Room, LeftSling13_LM_Flashers_f143, LeftSling13_LM_GIString11, LeftSling13_LM_GIString6, LeftSling13_LM_Spotlights) ' VLM.Array;BL;LeftSling13
Dim LeftSling14_LM: LeftSling14_LM=Array(LeftSling14_LM_Flashers_f143, LeftSling14_LM_GIString11, LeftSling14_LM_GIString6, LeftSling14_LM_Spotlights) ' VLM.Array;LM;LeftSling14
Dim LeftSling14_BM: LeftSling14_BM=Array(LeftSling14_BM_Lit_Room) ' VLM.Array;BM;LeftSling14
Dim LeftSling14_BL: LeftSling14_BL=Array(LeftSling14_BM_Lit_Room, LeftSling14_LM_Flashers_f143, LeftSling14_LM_GIString11, LeftSling14_LM_GIString6, LeftSling14_LM_Spotlights) ' VLM.Array;BL;LeftSling14
Dim LeftSling2_LM: LeftSling2_LM=Array(LeftSling2_LM_Flashers_f141, LeftSling2_LM_Flashers_f142, LeftSling2_LM_Flashers_f148, LeftSling2_LM_GIString1_gi001, LeftSling2_LM_GIString1_gi005, LeftSling2_LM_GIString1_gi26, LeftSling2_LM_Inserts_L15) ' VLM.Array;LM;LeftSling2
Dim LeftSling2_BM: LeftSling2_BM=Array(LeftSling2_BM_Lit_Room) ' VLM.Array;BM;LeftSling2
Dim LeftSling2_BL: LeftSling2_BL=Array(LeftSling2_BM_Lit_Room, LeftSling2_LM_Flashers_f141, LeftSling2_LM_Flashers_f142, LeftSling2_LM_Flashers_f148, LeftSling2_LM_GIString1_gi001, LeftSling2_LM_GIString1_gi005, LeftSling2_LM_GIString1_gi26, LeftSling2_LM_Inserts_L15) ' VLM.Array;BL;LeftSling2
Dim LeftSling3_LM: LeftSling3_LM=Array(LeftSling3_LM_Flashers_f141, LeftSling3_LM_Flashers_f142, LeftSling3_LM_Flashers_f148, LeftSling3_LM_GIString1_gi001, LeftSling3_LM_GIString1_gi005, LeftSling3_LM_GIString1_gi26, LeftSling3_LM_Inserts_L15) ' VLM.Array;LM;LeftSling3
Dim LeftSling3_BM: LeftSling3_BM=Array(LeftSling3_BM_Lit_Room) ' VLM.Array;BM;LeftSling3
Dim LeftSling3_BL: LeftSling3_BL=Array(LeftSling3_BM_Lit_Room, LeftSling3_LM_Flashers_f141, LeftSling3_LM_Flashers_f142, LeftSling3_LM_Flashers_f148, LeftSling3_LM_GIString1_gi001, LeftSling3_LM_GIString1_gi005, LeftSling3_LM_GIString1_gi26, LeftSling3_LM_Inserts_L15) ' VLM.Array;BL;LeftSling3
Dim LeftSling4_LM: LeftSling4_LM=Array(LeftSling4_LM_Flashers_f141, LeftSling4_LM_Flashers_f142, LeftSling4_LM_Flashers_f148, LeftSling4_LM_GIString1_gi001, LeftSling4_LM_GIString1_gi005, LeftSling4_LM_GIString1_gi26, LeftSling4_LM_Inserts_L15) ' VLM.Array;LM;LeftSling4
Dim LeftSling4_BM: LeftSling4_BM=Array(LeftSling4_BM_Lit_Room) ' VLM.Array;BM;LeftSling4
Dim LeftSling4_BL: LeftSling4_BL=Array(LeftSling4_BM_Lit_Room, LeftSling4_LM_Flashers_f141, LeftSling4_LM_Flashers_f142, LeftSling4_LM_Flashers_f148, LeftSling4_LM_GIString1_gi001, LeftSling4_LM_GIString1_gi005, LeftSling4_LM_GIString1_gi26, LeftSling4_LM_Inserts_L15) ' VLM.Array;BL;LeftSling4
Dim Lemk_LM: Lemk_LM=Array(Lemk_LM_Flashers_f141, Lemk_LM_Flashers_f148, Lemk_LM_GIString1_gi001, Lemk_LM_GIString1_gi005) ' VLM.Array;LM;Lemk
Dim Lemk_BM: Lemk_BM=Array(Lemk_BM_Lit_Room) ' VLM.Array;BM;Lemk
Dim Lemk_BL: Lemk_BL=Array(Lemk_BM_Lit_Room, Lemk_LM_Flashers_f141, Lemk_LM_Flashers_f148, Lemk_LM_GIString1_gi001, Lemk_LM_GIString1_gi005) ' VLM.Array;BL;Lemk
Dim PincabRails_LM: PincabRails_LM=Array(PincabRails_LM_Flashers_f141, PincabRails_LM_Flashers_f142, PincabRails_LM_Flashers_f147, PincabRails_LM_GIString11, PincabRails_LM_GIString8) ' VLM.Array;LM;PincabRails
Dim PincabRails_BM: PincabRails_BM=Array(PincabRails_BM_Lit_Room) ' VLM.Array;BM;PincabRails
Dim PincabRails_BL: PincabRails_BL=Array(PincabRails_BM_Lit_Room, PincabRails_LM_Flashers_f141, PincabRails_LM_Flashers_f142, PincabRails_LM_Flashers_f147, PincabRails_LM_GIString11, PincabRails_LM_GIString8) ' VLM.Array;BL;PincabRails
Dim RFMesh_LM: RFMesh_LM=Array(RFMesh_LM_Flashers_f142, RFMesh_LM_Flashers_f148, RFMesh_LM_GIString1_gi005, RFMesh_LM_GIString2_gi002, RFMesh_LM_GIString2_gi003, RFMesh_LM_GIString2_gi004, RFMesh_LM_Inserts_L01) ' VLM.Array;LM;RFMesh
Dim RFMesh_BM: RFMesh_BM=Array(RFMesh_BM_Lit_Room) ' VLM.Array;BM;RFMesh
Dim RFMesh_BL: RFMesh_BL=Array(RFMesh_BM_Lit_Room, RFMesh_LM_Flashers_f142, RFMesh_LM_Flashers_f148, RFMesh_LM_GIString1_gi005, RFMesh_LM_GIString2_gi002, RFMesh_LM_GIString2_gi003, RFMesh_LM_GIString2_gi004, RFMesh_LM_Inserts_L01) ' VLM.Array;BL;RFMesh
Dim RGate_LM: RGate_LM=Array(RGate_LM_Flashers_f143, RGate_LM_GIString11, RGate_LM_GIString6, RGate_LM_Spotlights) ' VLM.Array;LM;RGate
Dim RGate_BM: RGate_BM=Array(RGate_BM_Lit_Room) ' VLM.Array;BM;RGate
Dim RGate_BL: RGate_BL=Array(RGate_BM_Lit_Room, RGate_LM_Flashers_f143, RGate_LM_GIString11, RGate_LM_GIString6, RGate_LM_Spotlights) ' VLM.Array;BL;RGate
Dim RGate001_LM: RGate001_LM=Array(RGate001_LM_Flashers_f143) ' VLM.Array;LM;RGate001
Dim RGate001_BM: RGate001_BM=Array(RGate001_BM_Lit_Room) ' VLM.Array;BM;RGate001
Dim RGate001_BL: RGate001_BL=Array(RGate001_BM_Lit_Room, RGate001_LM_Flashers_f143) ' VLM.Array;BL;RGate001
Dim Remk_LM: Remk_LM=Array(Remk_LM_Flashers_f142, Remk_LM_Flashers_f148, Remk_LM_GIString2_gi002, Remk_LM_GIString2_gi004) ' VLM.Array;LM;Remk
Dim Remk_BM: Remk_BM=Array(Remk_BM_Lit_Room) ' VLM.Array;BM;Remk
Dim Remk_BL: Remk_BL=Array(Remk_BM_Lit_Room, Remk_LM_Flashers_f142, Remk_LM_Flashers_f148, Remk_LM_GIString2_gi002, Remk_LM_GIString2_gi004) ' VLM.Array;BL;Remk
Dim RightSling2_LM: RightSling2_LM=Array(RightSling2_LM_Flashers_f142, RightSling2_LM_Flashers_f148, RightSling2_LM_GIString2_gi002, RightSling2_LM_GIString2_gi003, RightSling2_LM_GIString2_gi004, RightSling2_LM_GIString8, RightSling2_LM_Inserts_L17) ' VLM.Array;LM;RightSling2
Dim RightSling2_BM: RightSling2_BM=Array(RightSling2_BM_Lit_Room) ' VLM.Array;BM;RightSling2
Dim RightSling2_BL: RightSling2_BL=Array(RightSling2_BM_Lit_Room, RightSling2_LM_Flashers_f142, RightSling2_LM_Flashers_f148, RightSling2_LM_GIString2_gi002, RightSling2_LM_GIString2_gi003, RightSling2_LM_GIString2_gi004, RightSling2_LM_GIString8, RightSling2_LM_Inserts_L17) ' VLM.Array;BL;RightSling2
Dim RightSling3_LM: RightSling3_LM=Array(RightSling3_LM_Flashers_f142, RightSling3_LM_Flashers_f148, RightSling3_LM_GIString2_gi002, RightSling3_LM_GIString2_gi003, RightSling3_LM_GIString2_gi004, RightSling3_LM_GIString8, RightSling3_LM_Inserts_L17) ' VLM.Array;LM;RightSling3
Dim RightSling3_BM: RightSling3_BM=Array(RightSling3_BM_Lit_Room) ' VLM.Array;BM;RightSling3
Dim RightSling3_BL: RightSling3_BL=Array(RightSling3_BM_Lit_Room, RightSling3_LM_Flashers_f142, RightSling3_LM_Flashers_f148, RightSling3_LM_GIString2_gi002, RightSling3_LM_GIString2_gi003, RightSling3_LM_GIString2_gi004, RightSling3_LM_GIString8, RightSling3_LM_Inserts_L17) ' VLM.Array;BL;RightSling3
Dim RightSling4_LM: RightSling4_LM=Array(RightSling4_LM_Flashers_f142, RightSling4_LM_Flashers_f148, RightSling4_LM_GIString2_gi004, RightSling4_LM_GIString8) ' VLM.Array;LM;RightSling4
Dim RightSling4_BM: RightSling4_BM=Array(RightSling4_BM_Lit_Room) ' VLM.Array;BM;RightSling4
Dim RightSling4_BL: RightSling4_BL=Array(RightSling4_BM_Lit_Room, RightSling4_LM_Flashers_f142, RightSling4_LM_Flashers_f148, RightSling4_LM_GIString2_gi004, RightSling4_LM_GIString8) ' VLM.Array;BL;RightSling4
Dim Spinner1_LM: Spinner1_LM=Array(Spinner1_LM_Flashers_f143, Spinner1_LM_Flashers_f144, Spinner1_LM_Flashers_f145, Spinner1_LM_Flashers_f146, Spinner1_LM_GIString3, Spinner1_LM_GIString5, Spinner1_LM_GIString6, Spinner1_LM_GIString7, Spinner1_LM_Inserts_L24, Spinner1_LM_Inserts_L25, Spinner1_LM_Inserts_L26, Spinner1_LM_Spotlights) ' VLM.Array;LM;Spinner1
Dim Spinner1_BM: Spinner1_BM=Array(Spinner1_BM_Lit_Room) ' VLM.Array;BM;Spinner1
Dim Spinner1_BL: Spinner1_BL=Array(Spinner1_BM_Lit_Room, Spinner1_LM_Flashers_f143, Spinner1_LM_Flashers_f144, Spinner1_LM_Flashers_f145, Spinner1_LM_Flashers_f146, Spinner1_LM_GIString3, Spinner1_LM_GIString5, Spinner1_LM_GIString6, Spinner1_LM_GIString7, Spinner1_LM_Inserts_L24, Spinner1_LM_Inserts_L25, Spinner1_LM_Inserts_L26, Spinner1_LM_Spotlights) ' VLM.Array;BL;Spinner1
Dim TLFMesh_LM: TLFMesh_LM=Array(TLFMesh_LM_GIString3, TLFMesh_LM_Inserts_L27, TLFMesh_LM_Inserts_L28, TLFMesh_LM_Spotlights) ' VLM.Array;LM;TLFMesh
Dim TLFMesh_BM: TLFMesh_BM=Array(TLFMesh_BM_Lit_Room) ' VLM.Array;BM;TLFMesh
Dim TLFMesh_BL: TLFMesh_BL=Array(TLFMesh_BM_Lit_Room, TLFMesh_LM_GIString3, TLFMesh_LM_Inserts_L27, TLFMesh_LM_Inserts_L28, TLFMesh_LM_Spotlights) ' VLM.Array;BL;TLFMesh
Dim TLFMeshU_LM: TLFMeshU_LM=Array(TLFMeshU_LM_Flashers_f141, TLFMeshU_LM_GIString3, TLFMeshU_LM_Inserts_L20, TLFMeshU_LM_Inserts_L24, TLFMeshU_LM_Inserts_L27, TLFMeshU_LM_Inserts_L28, TLFMeshU_LM_Spotlights) ' VLM.Array;LM;TLFMeshU
Dim TLFMeshU_BM: TLFMeshU_BM=Array(TLFMeshU_BM_Lit_Room) ' VLM.Array;BM;TLFMeshU
Dim TLFMeshU_BL: TLFMeshU_BL=Array(TLFMeshU_BM_Lit_Room, TLFMeshU_LM_Flashers_f141, TLFMeshU_LM_GIString3, TLFMeshU_LM_Inserts_L20, TLFMeshU_LM_Inserts_L24, TLFMeshU_LM_Inserts_L27, TLFMeshU_LM_Inserts_L28, TLFMeshU_LM_Spotlights) ' VLM.Array;BL;TLFMeshU
Dim Ufo1_LM: Ufo1_LM=Array(Ufo1_LM_Flashers_L110, Ufo1_LM_Flashers_f143, Ufo1_LM_Flashers_f144, Ufo1_LM_Flashers_f145, Ufo1_LM_Flashers_f146, Ufo1_LM_GIString11, Ufo1_LM_GIString5, Ufo1_LM_GIString6, Ufo1_LM_GIString7, Ufo1_LM_Inserts_L39, Ufo1_LM_Spotlights) ' VLM.Array;LM;Ufo1
Dim Ufo1_BM: Ufo1_BM=Array(Ufo1_BM_Lit_Room) ' VLM.Array;BM;Ufo1
Dim Ufo1_BL: Ufo1_BL=Array(Ufo1_BM_Lit_Room, Ufo1_LM_Flashers_L110, Ufo1_LM_Flashers_f143, Ufo1_LM_Flashers_f144, Ufo1_LM_Flashers_f145, Ufo1_LM_Flashers_f146, Ufo1_LM_GIString11, Ufo1_LM_GIString5, Ufo1_LM_GIString6, Ufo1_LM_GIString7, Ufo1_LM_Inserts_L39, Ufo1_LM_Spotlights) ' VLM.Array;BL;Ufo1
Dim alien14_LM: alien14_LM=Array(alien14_LM_Flashers_f142, alien14_LM_GIString4, alien14_LM_GIString8, alien14_LM_Spotlights) ' VLM.Array;LM;alien14
Dim alien14_BM: alien14_BM=Array(alien14_BM_Lit_Room) ' VLM.Array;BM;alien14
Dim alien14_BL: alien14_BL=Array(alien14_BM_Lit_Room, alien14_LM_Flashers_f142, alien14_LM_GIString4, alien14_LM_GIString8, alien14_LM_Spotlights) ' VLM.Array;BL;alien14
Dim alien5_LM: alien5_LM=Array(alien5_LM_Flashers_f141, alien5_LM_GIString3, alien5_LM_Inserts_L19, alien5_LM_Inserts_L20, alien5_LM_Inserts_L27, alien5_LM_Spotlights) ' VLM.Array;LM;alien5
Dim alien5_BM: alien5_BM=Array(alien5_BM_Lit_Room) ' VLM.Array;BM;alien5
Dim alien5_BL: alien5_BL=Array(alien5_BM_Lit_Room, alien5_LM_Flashers_f141, alien5_LM_GIString3, alien5_LM_Inserts_L19, alien5_LM_Inserts_L20, alien5_LM_Inserts_L27, alien5_LM_Spotlights) ' VLM.Array;BL;alien5
Dim alien6_LM: alien6_LM=Array(alien6_LM_Flashers_f143, alien6_LM_Flashers_f144, alien6_LM_Flashers_f146, alien6_LM_GIString5, alien6_LM_GIString6, alien6_LM_Inserts_L31, alien6_LM_Spotlights) ' VLM.Array;LM;alien6
Dim alien6_BM: alien6_BM=Array(alien6_BM_Lit_Room) ' VLM.Array;BM;alien6
Dim alien6_BL: alien6_BL=Array(alien6_BM_Lit_Room, alien6_LM_Flashers_f143, alien6_LM_Flashers_f144, alien6_LM_Flashers_f146, alien6_LM_GIString5, alien6_LM_GIString6, alien6_LM_Inserts_L31, alien6_LM_Spotlights) ' VLM.Array;BL;alien6
Dim alien8_LM: alien8_LM=Array(alien8_LM_Flashers_f143, alien8_LM_Flashers_f144, alien8_LM_Flashers_f146, alien8_LM_GIString4, alien8_LM_GIString5, alien8_LM_Inserts_L33, alien8_LM_Inserts_L34, alien8_LM_Inserts_L35, alien8_LM_Inserts_L38, alien8_LM_Spotlights) ' VLM.Array;LM;alien8
Dim alien8_BM: alien8_BM=Array(alien8_BM_Lit_Room) ' VLM.Array;BM;alien8
Dim alien8_BL: alien8_BL=Array(alien8_BM_Lit_Room, alien8_LM_Flashers_f143, alien8_LM_Flashers_f144, alien8_LM_Flashers_f146, alien8_LM_GIString4, alien8_LM_GIString5, alien8_LM_Inserts_L33, alien8_LM_Inserts_L34, alien8_LM_Inserts_L35, alien8_LM_Inserts_L38, alien8_LM_Spotlights) ' VLM.Array;BL;alien8
Dim sw16_LM: sw16_LM=Array(sw16_LM_Flashers_f141, sw16_LM_GIString1_gi26) ' VLM.Array;LM;sw16
Dim sw16_BM: sw16_BM=Array(sw16_BM_Lit_Room) ' VLM.Array;BM;sw16
Dim sw16_BL: sw16_BL=Array(sw16_BM_Lit_Room, sw16_LM_Flashers_f141, sw16_LM_GIString1_gi26) ' VLM.Array;BL;sw16
Dim sw17_LM: sw17_LM=Array(sw17_LM_Flashers_f142, sw17_LM_GIString2_gi004, sw17_LM_Inserts_L17) ' VLM.Array;LM;sw17
Dim sw17_BM: sw17_BM=Array(sw17_BM_Lit_Room) ' VLM.Array;BM;sw17
Dim sw17_BL: sw17_BL=Array(sw17_BM_Lit_Room, sw17_LM_Flashers_f142, sw17_LM_GIString2_gi004, sw17_LM_Inserts_L17) ' VLM.Array;BL;sw17
Dim sw26_LM: sw26_LM=Array(sw26_LM_Flashers_f141, sw26_LM_GIString1_gi001, sw26_LM_GIString1_gi26) ' VLM.Array;LM;sw26
Dim sw26_BM: sw26_BM=Array(sw26_BM_Lit_Room) ' VLM.Array;BM;sw26
Dim sw26_BL: sw26_BL=Array(sw26_BM_Lit_Room, sw26_LM_Flashers_f141, sw26_LM_GIString1_gi001, sw26_LM_GIString1_gi26) ' VLM.Array;BL;sw26
Dim sw27_LM: sw27_LM=Array(sw27_LM_Flashers_f142, sw27_LM_GIString2_gi003) ' VLM.Array;LM;sw27
Dim sw27_BM: sw27_BM=Array(sw27_BM_Lit_Room) ' VLM.Array;BM;sw27
Dim sw27_BL: sw27_BL=Array(sw27_BM_Lit_Room, sw27_LM_Flashers_f142, sw27_LM_GIString2_gi003) ' VLM.Array;BL;sw27
Dim sw38_LM: sw38_LM=Array(sw38_LM_GIString6) ' VLM.Array;LM;sw38
Dim sw38_BM: sw38_BM=Array(sw38_BM_Lit_Room) ' VLM.Array;BM;sw38
Dim sw38_BL: sw38_BL=Array(sw38_BM_Lit_Room, sw38_LM_GIString6) ' VLM.Array;BL;sw38
Dim sw41_LM: sw41_LM=Array(sw41_LM_Flashers_f142, sw41_LM_GIString4, sw41_LM_Inserts_L03, sw41_LM_Inserts_L16, sw41_LM_Inserts_L21, sw41_LM_Inserts_L22, sw41_LM_Inserts_L23, sw41_LM_Inserts_L46, sw41_LM_Inserts_L47, sw41_LM_Inserts_L48, sw41_LM_Spotlights) ' VLM.Array;LM;sw41
Dim sw41_BM: sw41_BM=Array(sw41_BM_Lit_Room) ' VLM.Array;BM;sw41
Dim sw41_BL: sw41_BL=Array(sw41_BM_Lit_Room, sw41_LM_Flashers_f142, sw41_LM_GIString4, sw41_LM_Inserts_L03, sw41_LM_Inserts_L16, sw41_LM_Inserts_L21, sw41_LM_Inserts_L22, sw41_LM_Inserts_L23, sw41_LM_Inserts_L46, sw41_LM_Inserts_L47, sw41_LM_Inserts_L48, sw41_LM_Spotlights) ' VLM.Array;BL;sw41
Dim sw42_LM: sw42_LM=Array(sw42_LM_Flashers_f142, sw42_LM_GIString4, sw42_LM_Inserts_L02, sw42_LM_Inserts_L03, sw42_LM_Inserts_L21, sw42_LM_Inserts_L22, sw42_LM_Inserts_L23, sw42_LM_Inserts_L40, sw42_LM_Inserts_L46, sw42_LM_Inserts_L47, sw42_LM_Inserts_L48, sw42_LM_Spotlights) ' VLM.Array;LM;sw42
Dim sw42_BM: sw42_BM=Array(sw42_BM_Lit_Room) ' VLM.Array;BM;sw42
Dim sw42_BL: sw42_BL=Array(sw42_BM_Lit_Room, sw42_LM_Flashers_f142, sw42_LM_GIString4, sw42_LM_Inserts_L02, sw42_LM_Inserts_L03, sw42_LM_Inserts_L21, sw42_LM_Inserts_L22, sw42_LM_Inserts_L23, sw42_LM_Inserts_L40, sw42_LM_Inserts_L46, sw42_LM_Inserts_L47, sw42_LM_Inserts_L48, sw42_LM_Spotlights) ' VLM.Array;BL;sw42
Dim sw43_LM: sw43_LM=Array(sw43_LM_Flashers_f143, sw43_LM_Flashers_f144, sw43_LM_GIString5, sw43_LM_GIString6, sw43_LM_Inserts_L30, sw43_LM_Inserts_L31, sw43_LM_Inserts_L32, sw43_LM_Inserts_L33, sw43_LM_Inserts_L34, sw43_LM_Inserts_L36, sw43_LM_Inserts_L37, sw43_LM_Inserts_L40, sw43_LM_Inserts_L41, sw43_LM_Inserts_L49, sw43_LM_Spotlights) ' VLM.Array;LM;sw43
Dim sw43_BM: sw43_BM=Array(sw43_BM_Lit_Room) ' VLM.Array;BM;sw43
Dim sw43_BL: sw43_BL=Array(sw43_BM_Lit_Room, sw43_LM_Flashers_f143, sw43_LM_Flashers_f144, sw43_LM_GIString5, sw43_LM_GIString6, sw43_LM_Inserts_L30, sw43_LM_Inserts_L31, sw43_LM_Inserts_L32, sw43_LM_Inserts_L33, sw43_LM_Inserts_L34, sw43_LM_Inserts_L36, sw43_LM_Inserts_L37, sw43_LM_Inserts_L40, sw43_LM_Inserts_L41, sw43_LM_Inserts_L49, sw43_LM_Spotlights) ' VLM.Array;BL;sw43
Dim sw44_LM: sw44_LM=Array(sw44_LM_GIString4, sw44_LM_Inserts_L02, sw44_LM_Inserts_L03, sw44_LM_Inserts_L21, sw44_LM_Inserts_L22, sw44_LM_Inserts_L23, sw44_LM_Inserts_L40, sw44_LM_Inserts_L43, sw44_LM_Inserts_L46, sw44_LM_Inserts_L47, sw44_LM_Inserts_L48, sw44_LM_Spotlights) ' VLM.Array;LM;sw44
Dim sw44_BM: sw44_BM=Array(sw44_BM_Lit_Room) ' VLM.Array;BM;sw44
Dim sw44_BL: sw44_BL=Array(sw44_BM_Lit_Room, sw44_LM_GIString4, sw44_LM_Inserts_L02, sw44_LM_Inserts_L03, sw44_LM_Inserts_L21, sw44_LM_Inserts_L22, sw44_LM_Inserts_L23, sw44_LM_Inserts_L40, sw44_LM_Inserts_L43, sw44_LM_Inserts_L46, sw44_LM_Inserts_L47, sw44_LM_Inserts_L48, sw44_LM_Spotlights) ' VLM.Array;BL;sw44
Dim sw45_LM: sw45_LM=Array(sw45_LM_Flashers_f143, sw45_LM_Flashers_f144, sw45_LM_Flashers_f146, sw45_LM_Inserts_L28, sw45_LM_Inserts_L30, sw45_LM_Inserts_L31, sw45_LM_Inserts_L33, sw45_LM_Inserts_L34, sw45_LM_Inserts_L35, sw45_LM_Inserts_L36, sw45_LM_Inserts_L37, sw45_LM_Inserts_L38, sw45_LM_Inserts_L39, sw45_LM_Inserts_L40, sw45_LM_Inserts_L41, sw45_LM_Inserts_L43, sw45_LM_Inserts_L50, sw45_LM_Spotlights) ' VLM.Array;LM;sw45
Dim sw45_BM: sw45_BM=Array(sw45_BM_Lit_Room) ' VLM.Array;BM;sw45
Dim sw45_BL: sw45_BL=Array(sw45_BM_Lit_Room, sw45_LM_Flashers_f143, sw45_LM_Flashers_f144, sw45_LM_Flashers_f146, sw45_LM_Inserts_L28, sw45_LM_Inserts_L30, sw45_LM_Inserts_L31, sw45_LM_Inserts_L33, sw45_LM_Inserts_L34, sw45_LM_Inserts_L35, sw45_LM_Inserts_L36, sw45_LM_Inserts_L37, sw45_LM_Inserts_L38, sw45_LM_Inserts_L39, sw45_LM_Inserts_L40, sw45_LM_Inserts_L41, sw45_LM_Inserts_L43, sw45_LM_Inserts_L50, sw45_LM_Spotlights) ' VLM.Array;BL;sw45
Dim sw46_LM: sw46_LM=Array(sw46_LM_Flashers_f144, sw46_LM_Flashers_f146, sw46_LM_GIString5, sw46_LM_GIString6, sw46_LM_Inserts_L28, sw46_LM_Inserts_L29, sw46_LM_Inserts_L30, sw46_LM_Inserts_L32, sw46_LM_Inserts_L33, sw46_LM_Inserts_L34, sw46_LM_Inserts_L35, sw46_LM_Inserts_L36, sw46_LM_Inserts_L37, sw46_LM_Inserts_L38, sw46_LM_Inserts_L39, sw46_LM_Inserts_L40, sw46_LM_Inserts_L41, sw46_LM_Inserts_L42, sw46_LM_Inserts_L43, sw46_LM_Inserts_L44, sw46_LM_Spotlights) ' VLM.Array;LM;sw46
Dim sw46_BM: sw46_BM=Array(sw46_BM_Lit_Room) ' VLM.Array;BM;sw46
Dim sw46_BL: sw46_BL=Array(sw46_BM_Lit_Room, sw46_LM_Flashers_f144, sw46_LM_Flashers_f146, sw46_LM_GIString5, sw46_LM_GIString6, sw46_LM_Inserts_L28, sw46_LM_Inserts_L29, sw46_LM_Inserts_L30, sw46_LM_Inserts_L32, sw46_LM_Inserts_L33, sw46_LM_Inserts_L34, sw46_LM_Inserts_L35, sw46_LM_Inserts_L36, sw46_LM_Inserts_L37, sw46_LM_Inserts_L38, sw46_LM_Inserts_L39, sw46_LM_Inserts_L40, sw46_LM_Inserts_L41, sw46_LM_Inserts_L42, sw46_LM_Inserts_L43, sw46_LM_Inserts_L44, sw46_LM_Spotlights) ' VLM.Array;BL;sw46
Dim sw47_LM: sw47_LM=Array(sw47_LM_Flashers_f144, sw47_LM_Flashers_f146, sw47_LM_GIString5, sw47_LM_GIString6, sw47_LM_Inserts_L28, sw47_LM_Inserts_L32, sw47_LM_Inserts_L33, sw47_LM_Inserts_L34, sw47_LM_Inserts_L35, sw47_LM_Inserts_L36, sw47_LM_Inserts_L37, sw47_LM_Inserts_L38, sw47_LM_Inserts_L39, sw47_LM_Inserts_L40, sw47_LM_Inserts_L41, sw47_LM_Inserts_L42, sw47_LM_Inserts_L43, sw47_LM_Inserts_L44, sw47_LM_Spotlights) ' VLM.Array;LM;sw47
Dim sw47_BM: sw47_BM=Array(sw47_BM_Lit_Room) ' VLM.Array;BM;sw47
Dim sw47_BL: sw47_BL=Array(sw47_BM_Lit_Room, sw47_LM_Flashers_f144, sw47_LM_Flashers_f146, sw47_LM_GIString5, sw47_LM_GIString6, sw47_LM_Inserts_L28, sw47_LM_Inserts_L32, sw47_LM_Inserts_L33, sw47_LM_Inserts_L34, sw47_LM_Inserts_L35, sw47_LM_Inserts_L36, sw47_LM_Inserts_L37, sw47_LM_Inserts_L38, sw47_LM_Inserts_L39, sw47_LM_Inserts_L40, sw47_LM_Inserts_L41, sw47_LM_Inserts_L42, sw47_LM_Inserts_L43, sw47_LM_Inserts_L44, sw47_LM_Spotlights) ' VLM.Array;BL;sw47
Dim sw48_LM: sw48_LM=Array(sw48_LM_Flashers_f143, sw48_LM_Flashers_f145, sw48_LM_GIString6) ' VLM.Array;LM;sw48
Dim sw48_BM: sw48_BM=Array(sw48_BM_Lit_Room) ' VLM.Array;BM;sw48
Dim sw48_BL: sw48_BL=Array(sw48_BM_Lit_Room, sw48_LM_Flashers_f143, sw48_LM_Flashers_f145, sw48_LM_GIString6) ' VLM.Array;BL;sw48
Dim sw56_LM: sw56_LM=Array(sw56_LM_Flashers_f141, sw56_LM_GIString3, sw56_LM_Inserts_L14, sw56_LM_Inserts_L18, sw56_LM_Inserts_L19, sw56_LM_Inserts_L20, sw56_LM_Spotlights) ' VLM.Array;LM;sw56
Dim sw56_BM: sw56_BM=Array(sw56_BM_Lit_Room) ' VLM.Array;BM;sw56
Dim sw56_BL: sw56_BL=Array(sw56_BM_Lit_Room, sw56_LM_Flashers_f141, sw56_LM_GIString3, sw56_LM_Inserts_L14, sw56_LM_Inserts_L18, sw56_LM_Inserts_L19, sw56_LM_Inserts_L20, sw56_LM_Spotlights) ' VLM.Array;BL;sw56
Dim sw57_LM: sw57_LM=Array(sw57_LM_Flashers_f141, sw57_LM_GIString3, sw57_LM_Inserts_L03, sw57_LM_Inserts_L18, sw57_LM_Inserts_L19, sw57_LM_Inserts_L20, sw57_LM_Inserts_L24, sw57_LM_Spotlights) ' VLM.Array;LM;sw57
Dim sw57_BM: sw57_BM=Array(sw57_BM_Lit_Room) ' VLM.Array;BM;sw57
Dim sw57_BL: sw57_BL=Array(sw57_BM_Lit_Room, sw57_LM_Flashers_f141, sw57_LM_GIString3, sw57_LM_Inserts_L03, sw57_LM_Inserts_L18, sw57_LM_Inserts_L19, sw57_LM_Inserts_L20, sw57_LM_Inserts_L24, sw57_LM_Spotlights) ' VLM.Array;BL;sw57
Dim sw58_LM: sw58_LM=Array(sw58_LM_Flashers_f141, sw58_LM_GIString3, sw58_LM_Inserts_L03, sw58_LM_Inserts_L18, sw58_LM_Inserts_L19, sw58_LM_Inserts_L20, sw58_LM_Inserts_L24, sw58_LM_Spotlights) ' VLM.Array;LM;sw58
Dim sw58_BM: sw58_BM=Array(sw58_BM_Lit_Room) ' VLM.Array;BM;sw58
Dim sw58_BL: sw58_BL=Array(sw58_BM_Lit_Room, sw58_LM_Flashers_f141, sw58_LM_GIString3, sw58_LM_Inserts_L03, sw58_LM_Inserts_L18, sw58_LM_Inserts_L19, sw58_LM_Inserts_L20, sw58_LM_Inserts_L24, sw58_LM_Spotlights) ' VLM.Array;BL;sw58
Dim sw62_LM: sw62_LM=Array(sw62_LM_Flashers_f146, sw62_LM_GIString11, sw62_LM_GIString5, sw62_LM_GIString6) ' VLM.Array;LM;sw62
Dim sw62_BM: sw62_BM=Array(sw62_BM_Lit_Room) ' VLM.Array;BM;sw62
Dim sw62_BL: sw62_BL=Array(sw62_BM_Lit_Room, sw62_LM_Flashers_f146, sw62_LM_GIString11, sw62_LM_GIString5, sw62_LM_GIString6) ' VLM.Array;BL;sw62
Dim sw71_LM: sw71_LM=Array(sw71_LM_Flashers_f143, sw71_LM_GIString6, sw71_LM_GIString7) ' VLM.Array;LM;sw71
Dim sw71_BM: sw71_BM=Array(sw71_BM_Lit_Room) ' VLM.Array;BM;sw71
Dim sw71_BL: sw71_BL=Array(sw71_BM_Lit_Room, sw71_LM_Flashers_f143, sw71_LM_GIString6, sw71_LM_GIString7) ' VLM.Array;BL;sw71
Dim sw72_LM: sw72_LM=Array(sw72_LM_Flashers_f143) ' VLM.Array;LM;sw72
Dim sw72_BM: sw72_BM=Array(sw72_BM_Lit_Room) ' VLM.Array;BM;sw72
Dim sw72_BL: sw72_BL=Array(sw72_BM_Lit_Room, sw72_LM_Flashers_f143) ' VLM.Array;BL;sw72
Dim sw73_LM: sw73_LM=Array(sw73_LM_GIString11, sw73_LM_GIString5) ' VLM.Array;LM;sw73
Dim sw73_BM: sw73_BM=Array(sw73_BM_Lit_Room) ' VLM.Array;BM;sw73
Dim sw73_BL: sw73_BL=Array(sw73_BM_Lit_Room, sw73_LM_GIString11, sw73_LM_GIString5) ' VLM.Array;BL;sw73
Dim sw74_LM: sw74_LM=Array() ' VLM.Array;LM;sw74
Dim sw74_BM: sw74_BM=Array(sw74_BM_Lit_Room) ' VLM.Array;BM;sw74
Dim sw74_BL: sw74_BL=Array(sw74_BM_Lit_Room) ' VLM.Array;BL;sw74
Dim sw75_LM: sw75_LM=Array(sw75_LM_Flashers_f143, sw75_LM_Flashers_f144, sw75_LM_GIString5, sw75_LM_Inserts_L32, sw75_LM_Inserts_L33, sw75_LM_Inserts_L34, sw75_LM_Inserts_L35, sw75_LM_Inserts_L36, sw75_LM_Inserts_L37, sw75_LM_Inserts_L38, sw75_LM_Inserts_L41, sw75_LM_Inserts_L42, sw75_LM_Inserts_L44, sw75_LM_Inserts_L45, sw75_LM_Spotlights) ' VLM.Array;LM;sw75
Dim sw75_BM: sw75_BM=Array(sw75_BM_Lit_Room) ' VLM.Array;BM;sw75
Dim sw75_BL: sw75_BL=Array(sw75_BM_Lit_Room, sw75_LM_Flashers_f143, sw75_LM_Flashers_f144, sw75_LM_GIString5, sw75_LM_Inserts_L32, sw75_LM_Inserts_L33, sw75_LM_Inserts_L34, sw75_LM_Inserts_L35, sw75_LM_Inserts_L36, sw75_LM_Inserts_L37, sw75_LM_Inserts_L38, sw75_LM_Inserts_L41, sw75_LM_Inserts_L42, sw75_LM_Inserts_L44, sw75_LM_Inserts_L45, sw75_LM_Spotlights) ' VLM.Array;BL;sw75
Dim sw76_001_LM: sw76_001_LM=Array(sw76_001_LM_Flashers_f143, sw76_001_LM_Flashers_f144, sw76_001_LM_Flashers_f146, sw76_001_LM_GIString5, sw76_001_LM_GIString6, sw76_001_LM_Inserts_L29, sw76_001_LM_Inserts_L30, sw76_001_LM_Inserts_L31, sw76_001_LM_Inserts_L32, sw76_001_LM_Inserts_L33, sw76_001_LM_Inserts_L34, sw76_001_LM_Inserts_L35, sw76_001_LM_Inserts_L36, sw76_001_LM_Inserts_L37, sw76_001_LM_Inserts_L41, sw76_001_LM_Inserts_L42, sw76_001_LM_Inserts_L44, sw76_001_LM_Inserts_L45, sw76_001_LM_Spotlights) ' VLM.Array;LM;sw76_001
Dim sw76_001_BM: sw76_001_BM=Array(sw76_001_BM_Lit_Room) ' VLM.Array;BM;sw76_001
Dim sw76_001_BL: sw76_001_BL=Array(sw76_001_BM_Lit_Room, sw76_001_LM_Flashers_f143, sw76_001_LM_Flashers_f144, sw76_001_LM_Flashers_f146, sw76_001_LM_GIString5, sw76_001_LM_GIString6, sw76_001_LM_Inserts_L29, sw76_001_LM_Inserts_L30, sw76_001_LM_Inserts_L31, sw76_001_LM_Inserts_L32, sw76_001_LM_Inserts_L33, sw76_001_LM_Inserts_L34, sw76_001_LM_Inserts_L35, sw76_001_LM_Inserts_L36, sw76_001_LM_Inserts_L37, sw76_001_LM_Inserts_L41, sw76_001_LM_Inserts_L42, sw76_001_LM_Inserts_L44, sw76_001_LM_Inserts_L45, sw76_001_LM_Spotlights) ' VLM.Array;BL;sw76_001
Dim sw77_LM: sw77_LM=Array(sw77_LM_Flashers_f144, sw77_LM_Flashers_f146, sw77_LM_GIString5, sw77_LM_Inserts_L32, sw77_LM_Inserts_L33, sw77_LM_Inserts_L34, sw77_LM_Inserts_L35, sw77_LM_Inserts_L36, sw77_LM_Inserts_L37, sw77_LM_Inserts_L38, sw77_LM_Inserts_L41, sw77_LM_Inserts_L42, sw77_LM_Inserts_L44, sw77_LM_Spotlights) ' VLM.Array;LM;sw77
Dim sw77_BM: sw77_BM=Array(sw77_BM_Lit_Room) ' VLM.Array;BM;sw77
Dim sw77_BL: sw77_BL=Array(sw77_BM_Lit_Room, sw77_LM_Flashers_f144, sw77_LM_Flashers_f146, sw77_LM_GIString5, sw77_LM_Inserts_L32, sw77_LM_Inserts_L33, sw77_LM_Inserts_L34, sw77_LM_Inserts_L35, sw77_LM_Inserts_L36, sw77_LM_Inserts_L37, sw77_LM_Inserts_L38, sw77_LM_Inserts_L41, sw77_LM_Inserts_L42, sw77_LM_Inserts_L44, sw77_LM_Spotlights) ' VLM.Array;BL;sw77
Dim swPlunger_LM: swPlunger_LM=Array(swPlunger_LM_Flashers_f142, swPlunger_LM_Flashers_f147, swPlunger_LM_GIString8) ' VLM.Array;LM;swPlunger
Dim swPlunger_BM: swPlunger_BM=Array(swPlunger_BM_Lit_Room) ' VLM.Array;BM;swPlunger
Dim swPlunger_BL: swPlunger_BL=Array(swPlunger_BM_Lit_Room, swPlunger_LM_Flashers_f142, swPlunger_LM_Flashers_f147, swPlunger_LM_GIString8) ' VLM.Array;BL;swPlunger


'//////////////////////////////////////////////////////////////////////
'// Constants and Variables
'//////////////////////////////////////////////////////////////////////

TiltLevel = .15 'Threshold for Analog Nudge. Raise value to tilt less often. Values are small, only adjust by .01 at a time.

Dim RomSoundVolume : RomSoundVolume = 0.1
Dim TiltLevel

Dim lightCtrl : Set lightCtrl = New LStateController

Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = ""

Const tnob = 5
Const lob = 0

Dim AmbientBallShadowOn : AmbientBallShadowOn= 1
Dim DynamicBallShadowsOn : DynamicBallShadowsOn = 0
Const RubberizerEnabled = 1
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height


Dim bsTrough, bsL, bsR, dtDrop, x, BallFrame, plungerIM, Mech3bank
Dim BankMove, MoveSpeed, currpos, NewBankDown, ShakeThree, ShakeMain, ShakeSponge, ShakePatty, ShakePatrick, ShakeSquidward, i
Dim HSI1, HSI2, HSI3
Dim StartGame
Dim CurrentBall
Dim Players
Dim CurrentPlayer
Dim PreviousPlayer
Dim ROrbitHit
Dim LOrbitHit
Dim RRampHit
Dim LRampHit
Dim PlayTime
Dim OrbitSaved
Dim RoundScore
Dim BallInLane
Dim SkillShot
Dim SkillShot2
Dim SSText
Dim LEDMessage2
Dim AreLightsOff
Dim DelayMusic
Dim DScoreValue
Dim SongNumber
Dim LaneTune
Dim HSDisplay
Dim HSDisplay2
Dim ScoreNumber
Dim HSDigit
Dim HSChr
Dim DrainSave
Dim SideSave
Dim LTiltPower
Dim RTiltPower
Dim CTiltPower
Dim LeftPress
Dim RightPress
Dim CenterPress
Dim WarningCount
Dim TiltGame
Dim SkillSelect
Dim SeqNum
Dim SpongeBalls
Dim soundcount
Dim TiltTime
Dim HSTextFile
Dim PartyMode
Dim FanfareSelect
Dim HighScores(5,1)
Dim MusicInterval(11)
Dim DMDDisplay(20,20)
Dim MultiBallReady(4)
Dim ModeText(4)
Dim TotalScore(4)
Dim KickerSave(4)
Dim BallSaveTime(4)
Dim ExtraBall
Dim BonusRounds
Dim BankDown(4)
Dim TriggerWallCollide(4)
Dim TriggerDown(4)
Dim ModeReady(4)
Dim CurrentMode
Dim NextMode(4)
Dim ModesComplete(4)
Dim BallsLocked(4)
Dim SJPReady(4)

Dim Tilt
Dim MechTilt
Dim TiltSensitivity
Dim bTilted
Dim bMechTiltJustHit

Dim L01State(4)
Dim L02State(4)
Dim L03State(4)
Dim L04State(4)
Dim L05State(4)
Dim L06State(4)
Dim L07State(4)
Dim L08State(4)
Dim L09State(4)
Dim L10State(4)
Dim L11State(4)
Dim L12State(4)
Dim L13State(4)
Dim L14State(4)
Dim L15State(4)
Dim L16State(4)
Dim L17State(4)
Dim L18State(4)
Dim L19State(4)
Dim L20State(4)
Dim L21State(4)
Dim L22State(4)
Dim L23State(4)
Dim L24State(4)
Dim L25State(4)
Dim L26State(4)
Dim L27State(4)
Dim L28State(4)
Dim L29State(4)
Dim L30State(4)
Dim L31State(4)
Dim L32State(4)
Dim L33State(4)
Dim L34State(4)
Dim L35State(4)
Dim L36State(4)
Dim L37State(4)
Dim L38State(4)
Dim L39State(4)
Dim L40State(4)
Dim L41State(4)
Dim L42State(4)
Dim L43State(4)
Dim L44State(4)
Dim L45State(4)
Dim L46State(4)
Dim L47State(4)
Dim L48State(4)
Dim L49State(4)
Dim L50State(4)
Wall002.IsDropped = True
Wall003.IsDropped = True
Wall004.IsDropped = True
Wall005.IsDropped = True

ScoreText.Text = ""
ScoreText2.Text = ""
ScoreText3.Text = ""
ScoreText4.Text = ""
BallsText.Text = ""
InsertSequence.Play SeqCircleOutOn,50,2
lightCtrl.SyncWithVpxLights GISequence
GISequence.Play SeqCircleOutOn,50,100
DMDText.Text = "Bikini Bottom Pinball: Press Start" : LEDText("Bikini Bottom Pinball Press Start")

'If Cabmode=1 then
'rrail.visible = 0
'lrail.visible = 0
'Else
'rrail.visible = 1
'lrail.visible = 1
'End if


' Culture neutral string to double conversion (handles situation where you don't know how the string was written)
Function CNCDbl(str)
    Dim strt, Sep, i
    If IsNumeric(str) Then
        CNCDbl = CDbl(str)
    Else
        Sep = Mid(CStr(0.5), 2, 1)
        Select Case Sep
        Case "."
            i = InStr(1, str, ",")
        Case ","
            i = InStr(1, str, ".")
        End Select
        If i = 0 Then
            CNCDbl = Empty
        Else
            strt = Mid(str, 1, i - 1) & Sep & Mid(str, i + 1)
            If IsNumeric(strt) Then
                CNCDbl = CDbl(strt)
            Else
                CNCDbl = Empty
            End If
        End If
    End If
End Function

'************
' Table init.
'************

Const cGameName = "spongebob"


Dim gioriginal(11)

Sub Table1_Init
  LoadEM
  If Vrroom = 0 then
'   startcontroller
    LeftFlipper1_Timer
  Else
    LeftFlipper1.timerenabled = True
  End If
End Sub

Sub LeftFlipper1_Timer
  LeftFlipper1.timerenabled = False

  If ForceDMD_Flasher = 1 And vrroom = 0 Then FlexOnFlasher = 1 : VRDMD001.visible = True ': VRDMD5.visible = True : VRDMD5.x = 1600 : VRDMD5.y = 220 : VRDMD5.height = 100 : VRDMD5.RotX = -50 : VRDMD5.opacity = 150

  flex_init

  vpmMapLights(AllLights)
  lightCtrl.RegisterLights "VPX"


    ' Impulse Plunger
    Const IMPowerSetting = 50 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 3
        .InitExitSnd SoundFX("ShooterLanefx",DOFContactors), SoundFX("popper_ball",DOFContactors)
        .CreateEvents "plungerIM"
    End With
  Options_Load
  resetnewgame
  bTilted = True
  DMDtimer.enabled = True
  gioriginal(1) = l101.color
  gioriginal(2) = l102.color
  gioriginal(3) = l103.color
  gioriginal(4) = l104.color
  gioriginal(5) = l105.color
  gioriginal(6) = l106.color
  gioriginal(7) = l107.color
  gioriginal(8) = l108.color
  gioriginal(9) = l109.color
  gioriginal(10) = l110.color
  gioriginal(11) = l111.color


  LoadHighScores

  Call PlayTableMusic(12)

  SetRoomBrightness RoomBrightness/100

  FlipperHeight

End Sub

Sub Table1_Exit
  If B2SOn Then Controller.Stop
End Sub


Dim SSmade(4)
Dim modedone1(4)
Dim modedone2(4)
Dim modedone3(4)
Dim modedone4(4)
Dim modedone5(4)
Dim modedone6(4)
Dim modedone7(4)
Dim modedone8(4)
Dim modedone9(4)
Dim InlaneEB(4)
Dim RandomEB(4)
Dim MBaddaball(4)

Dim SpinnerCount1(4)
Dim SpinnerCount2(4)
Dim SpinnerLevel1(4)
Dim SpinnerLevel2(4)
Dim HurryUpCollected(4)
Dim RampsForMystery(4)
Dim MysteryReady(4)
Dim RampsMysteryCounter(4)
Dim CompletedCombos(4)
Dim MB_SJP(4)
Dim SecretshotDone(4)
Dim modeEBgotten(4)
Dim Wizard1Done(4)
Sub resetnewgame

  'CreateGameDMD
  ModeJP_Done = True
  ModeJP_Count = 0

  LStep = 0:LeftSlingShot.TimerEnabled=1
  L1Step = 0:LeftSlingShot1.TimerEnabled=1
  RStep = 0:RightSlingShot.TimerEnabled=1

  'Preload flashers
  'Flash1 :Flash2 :Flash3 :Flash4 :Flash5
  lightCtrl.Pulse f141, 0: DOF 141,2
  lightCtrl.Pulse f142, 0: DOF 142,2
  lightCtrl.Pulse f143, 0: DOF 143,2
  lightCtrl.Pulse f144, 0: DOF 144,2
  lightCtrl.Pulse f145, 0: DOF 145,2

  Warnings = 0
  BankMove = 0
  NewBankDown = 0
  StartGame = 0
  CurrentBall = 3
  Players = 0
  CurrentPlayer = 0
  CurrentMode = 0
  ROrbitHit = 0
  LOrbitHit = 0
  RRampHit = 0
  LRampHit = 0
  PlayTime = 0
  BallInLane = 0
  ExtraBall = 0
  OrbitSaved = 0
  RoundScore = 0
  BonusRounds = 0
  SkillShot = 0
  SSText = 0
  AreLightsOff = 0
  DelayMusic = 0
  DScoreValue = 0
  HSDisplay2 = 0
  ScoreNumber = ""
  HSDigit = 0
  HsChr = 65
  DrainSave = 0
  SideSave = 0
  SeqNum = 0
  SpongeBalls = 0
  soundcount = 0
  WarningCount = 0

  FanfareSelect = int(rnd*5)
  SongNumber = int(rnd*7)

' TiltGame = 1
  Tilt = 0
  MechTilt = 0
  TiltSensitivity = 5
' bTilted = False
  bMechTiltJustHit = False
  StartGame = 0

  for i = 0 To 4
    Wizard1Done(i) = 0
    MBaddaball(i) = 0
    modeEBgotten(i) = 0
    SecretshotDone(i) = 0
    MB_SJP(i) = 0
    CompletedCombos(i) = 0
    modedone1(i) = 0
    modedone2(i) = 0
    modedone3(i) = 0
    modedone4(i) = 0
    modedone5(i) = 0
    modedone6(i) = 0
    modedone7(i) = 0
    modedone8(i) = 0
    modedone9(i) = 0
    InlaneEB(i) = 0
    SpinnerCount1(i) = 0
    SpinnerCount2(i) = 0
    SpinnerLevel1(i) = 0
    SpinnerLevel2(i) = 0
    HurryUpCollected(i) = 0
    RampsForMystery(i) = 2
    MysteryReady(i) = 0
    RampsMysteryCounter(i) = 0
    RandomEB(i) = 0
    SSmade(i) = 0
    BankDown(i)=0
    TriggerWallCollide(i)=False
    TriggerDown(i)=True
    ModeReady(i)=0
    MultiBallReady(i) = 0
    ModesComplete(i)=0
    NextMode(i)=0
    BallsLocked(i)=0
    BallSaveTime(i)=9
    KickerSave(i) = 2
    TotalScore(i) = 0
    SJPReady(i) = 0
    ModeText(i) = ""
     L01State(i)=0
     L02State(i)=0
     L03State(i)=0
     L04State(i)=0
     L05State(i)=0
     L06State(i)=0
     L07State(i)=0
     L08State(i)=0
     L09State(i)=0
     L10State(i)=0
     L11State(i)=0
     L12State(i)=0
     L13State(i)=0
     L14State(i)=0
     L15State(i)=0
     L16State(i)=0
     L17State(i)=0
     L18State(i)=0
     L19State(i)=0
     L20State(i)=0
     L21State(i)=0
     L22State(i)=0
     L23State(i)=0
     L24State(i)=0
     L25State(i)=0
     L26State(i)=0
     L27State(i)=0
     L28State(i)=0
     L29State(i)=0
     L30State(i)=0
     L31State(i)=0
     L32State(i)=0
     L33State(i)=0
     L34State(i)=0
     L35State(i)=0
     L36State(i)=0
     L37State(i)=0
     L38State(i)=0
     L39State(i)=0
     L40State(i)=0
     L41State(i)=0
     L42State(i)=0
     L43State(i)=0
     L44State(i)=0
     L45State(i)=0
     L46State(i)=0
     L47State(i)=0
     L48State(i)=0
     L49State(i)=0
     L50State(i)=0
  Next

  For i = 0 To 9
    DMDDisplay(i,0) = ""
    DMDDisplay(i,1) = ""
    DMDDisplay(i,2) = ""
    DMDDisplay(i,3) = ""
    DMDDisplay(i,4) = ""
    DMDDisplay(i,5) = ""
  Next

  For i = 0 To 11
    MusicInterval(i) = 2000
  Next

  MusicInterval(1) = 6096 '17948
  MusicInterval(2) = 6372
  MusicInterval(3) = 3930 '5312
  MusicInterval(4) = 5928
  MusicInterval(5) = 9148
  MusicInterval(6) = 8668 '16210
  MusicInterval(7) = 6524
  MusicInterval(8) = 9854
  MusicInterval(9) = 8284
  MusicInterval(10) = 9376
  MusicInterval(11) = 9376



  UpdateLights
  Call ChangeGILights(0)
  DiverterON = 1

  HSDisplay = 1 : DisplayHS.Enabled = True

  TriggerWall.Collidable = False  : TriggerWallCollide(CurrentPlayer)=False
  TriggerWall.IsDropped = Not TriggerWall.Collidable
  TriggerDown(CurrentPlayer) = True
  sw77MovableHelper

  StartupScreen = 4


End Sub

Sub ResetPlayer
  ROrbitHit = 0
  LOrbitHit = 0
  LRampHit = 0
  RRampHit = 0
  WarningCount = 0
  LGate.Open=False
  RGate.Open=False
  PreviousPlayer = CurrentPlayer - 1
  If CurrentPlayer = 0 OR CurrentPlayer = 1 Then PreviousPlayer = Players
  If CurrentBall > 0 Then
    If BankDown(CurrentPlayer) = 1 AND BankDown(PreviousPlayer) = 0 Then currpos = 50 : Call Update3Bank(50,0)
    If BankDown(CurrentPlayer) = 0 AND BankDown(PreviousPlayer) = 1 Then currpos = 0 : Call Update3Bank(0,50)
  End If
  'sw77.IsDropped = TriggerDown(CurrentPlayer)
  TriggerWall.Collidable = TriggerWallCollide(CurrentPlayer)
  TriggerWall.IsDropped = Not TriggerWall.Collidable
  sw77MovableHelper
  UpdateLights
  DiverterON = 1

  If ModeReady(CurrentPlayer) = 1 And MultiBallReady(CurrentPlayer) = 0 Then DiverterON = 2
End Sub

'//////////////////////////////////////////////////////////////////////
'// TIMERS
'//////////////////////////////////////////////////////////////////////

Dim TriggerWallPos : TriggerWallPos = 15
Dim TriggerGate : triggergate = 45
Dim Triggerbounce
' The game timer interval is 10 ms
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking
  RollingUpdate         'update rolling sounds
  DoSTAnim
  UpdateBallBrightness
  Primitive001.blenddisablelighting = Light001.GetInPlayIntensity


  Diverterupdate

End Sub


' The frame timer interval is -1, so executes at the display frame rate
dim FrameTime, InitFrameTime : InitFrameTime = 0
Sub FrameTimer_Timer()
  FrameTime = gametime - InitFrameTime : InitFrameTime = gametime 'Count frametime. Unused atm?
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub

Sub LampTimer_Timer()
  lightCtrl.Update
End Sub


'**********
' Keys
'**********
Dim ssREADY
Dim plungerwait1
dim cancelL
dim cancelR


Sub table1_KeyDown(ByVal Keycode)



  'VR Specific...
  If keycode = StartGameKey and VRRoom > 0  Then
    VRStartButton.y = VRStartButton.y - 5
    if VRRoom = 1 then  StartButtonBubbleOn = true:   VRBubbleStartButton.visible = true
  End If

  If keycode = PlungerKey and VRRoom >0 Then
    VRLaunchButtonText.y = VRLaunchButtonText.y -5
    VRLaunchButton.y = VRLaunchButton.y -5
    If VRRoom =1 then LaunchButtonBubbleOn = true: VRBubbleLaunchButton.visible = true
  End if

  If keycode = LeftFlipperKey and VRRoom > 0 Then
    VR_FlipperLeft.x = VR_FlipperLeft.x +5
    If VRRoom = 1 then LeftFlipperBubbleOn = true: VRBubbleLeftFlipper.visible = true
  End If

  If keycode = RightFlipperKey and VRRoom >0 Then
    VR_FlipperRight.x = VR_FlipperRight.x -5
    If VRRoom = 1 then RightFlipperBubbleOn = true: VRBubbleRightFlipper.visible = true
  End if

  If bInOptions = True Then
    Options_KeyDown keycode
    Exit Sub
  End If

  If keycode = LeftMagnaSave Or keycode = RightMagnaSave  Then
    If bOptionsMagna = False And bInOptions = False And inQRClaimDMD = False Then StartQRclaim
  End If


  If keycode = LeftMagnaSave Then
    If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
    ElseIf keycode = RightMagnaSave Then
    If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
  End If
  If keycode = LeftMagnaSave Or keycode = RightMagnaSave  Then
    If bInOptions And inQRClaimDMD Then StopQRClaim
  End If


  If keycode = PlungerKey And Plungerwait1 < Frame And BallInLane > 0 Then
    DOf 123,2
    If ssREADY < 1 Then plungerwait1 = frame + 70
    If DoOrDieMode = 1 Then Doordiemode = 2 : DOF 133,0
    If DoOrDieMode = 3 Then Doordiemode = 4 : DOF 133,1 : CurrentBall = 1 : InlaneEB(CurrentPlayer) = 5
    If ssREADY = 1 And SkillShot = 1 Then ssREADY = 0 : SkillShot = 0 : Auto_Plunger : playtime = -2
  End If

    If keycode = LeftTiltKey Then
    LeftPress = 1
    Nudge 90, 1.5
    SoundNudgeLeft
    ToyShake
    CheckTilt
  End If
    If keycode = RightTiltKey Then
    RightPress = 1
    Nudge 270, 1.5
    SoundNudgeRight
    ToyShake
    CheckTilt
  End If
    If keycode = CenterTiltKey Then
    CenterPress = 1
    Nudge 0, 1.5
    SoundNudgeCenter
    ToyShake
    CheckTilt
  End If
    If (keycode = MechanicalTilt) And (Not bTilted) And (ScoreNumber = "") Then
    CheckMechTilt
  End If

  If Keycode = 30 Then
    If StagedFlipperMod = 1 and Not bTilted Then
      SolULFlipper True
    End If
  End If

  If Keycode = LeftFlipperKey Then
    If DoOrDieMode = 1 Then Dodnext = 1 : Doordiemode = 3 : ClearDMD : FlexDMD.Stage.GetFrame("VSeparator").Thickness = 0 : FlexDMD.Stage.GetLabel("Score_1").visible = False : Stopsound "slothhey" : PlaySound "slothhey",0,RomSoundVolume : Light001.state = 2 : Primitive001.visible = True
    If PartyMode = 1 Then StopSound "Sfx_Airhorn" : PlaySound "Sfx_Airhorn",0,RomSoundVolume
    If FanfareTimer.Enabled = True Then FanfareSfx : cancelL = 1 : If cancelL = 1 And CancelR = 1 Then FanfareTimer.Enabled = False : ClearDMD : sw78_Timer

    If Not bTilted Then
      If ScoreNumber = "" Then
        'If CurrentMode <> 2 Then    'Commented this out because if feels wrong with the flippers die in this mode. -apophis
          FlipperActivate LeftFlipper, LFPress
          SolLFlipper True
          If StagedFlipperMod <> 1 Then
            SolULFlipper True
          End If
        'End If
        RotateLaneLights
      End If
    Elseif startgame = 0 Then
      PlayQuote(3)
    End If
    If EnterHighscore > 101 And EnterHighscore < 127 Then
      SkillShotSfx
      EnterGoleft
      EnterKey = 1
      EnterRepeat = 0
      RepeatTimer.enabled = False
      RepeatTimer.enabled = True
    End If



  End If
  If Keycode = RightFlipperKey Then
    If DoOrDieMode = 3 Then Doordiemode = 1 : Stopsound "slothhey" : FlexDMD.Stage.GetLabel("Score_1").visible = True : FlexDMD.Stage.GetFrame("VSeparator").Thickness = 1 : FlexDMD.Stage.Getimage("dod1").visible = False : FlexDMD.Stage.Getimage("dod2").visible = False : FlexDMD.Stage.Getimage("dod3").visible = False : Light001.state = 0 : Primitive001.visible = False
    If PartyMode = 1 Then StopSound "Sfx_Airhorn" : PlaySound "Sfx_Airhorn",0,RomSoundVolume
    If FanfareTimer.Enabled = True Then FanfareSfx : cancelR = 1 : If cancelL = 1 And CancelR = 1 Then FanfareTimer.Enabled = False : ClearDMD : sw78_Timer
    If Not bTilted Then
      If ScoreNumber = "" Then
        'If CurrentMode <> 2 Then     'Commented this out because if feels wrong with the flippers die in this mode. -apophis
          FlipperActivate RightFlipper, RFPress
          SolRFlipper True
        'End If
        RotateLaneLights
      End If
    Elseif startgame = 0 Then
      PlayQuote(3)
    End If
    If EnterHighscore > 101 And EnterHighscore < 127 Then
      SkillShotSfx
      EnterGoright
      EnterKey = 2
      EnterRepeat = 0
      RepeatTimer.enabled = False
      RepeatTimer.enabled = True
    End If
  End If
  If keycode = StartGameKey And CurrentBall < 3 Then StartbuttonHold = frame + 62
End Sub
Dim StartbuttonHold
Dim RestartGame



Sub SetGI(col,time)
  If AreLightsOff = 1 Then ChangeGILights(1)
  lightCtrl.LightColor l101, col
  lightCtrl.LightColor l102, col
  lightCtrl.LightColor l103, col
  lightCtrl.LightColor l104, col
  lightCtrl.LightColor l105, col
  lightCtrl.LightColor l106, col
  lightCtrl.LightColor l107, col
  lightCtrl.LightColor l108, col
  lightCtrl.LightColor l109, col
  lightCtrl.LightColor l110, col
  lightCtrl.LightColor l111, col
  gitimer.interval = time
  gitimer.enabled = False
  gitimer.enabled = True

End Sub


Sub gitimer_Timer
  If AreLightsOff = 1 Then ChangeGILights(0)
  gitimer.enabled = False
  lightCtrl.LightColor l101, gioriginal(1)
  lightCtrl.LightColor l102, gioriginal(2)
  lightCtrl.LightColor l103, gioriginal(3)
  lightCtrl.LightColor l104, gioriginal(4)
  lightCtrl.LightColor l105, gioriginal(5)
  lightCtrl.LightColor l106, gioriginal(6)
  lightCtrl.LightColor l107, gioriginal(7)
  lightCtrl.LightColor l108, gioriginal(8)
  lightCtrl.LightColor l109, gioriginal(9)
  lightCtrl.LightColor l110, gioriginal(10)
  lightCtrl.LightColor l111, gioriginal(11)
End Sub

Dim CreditBasher : CreditBasher = 0
Dim attractcreditNO
Dim Credits
Sub table1_KeyUp(ByVal Keycode)

   If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    If UseCredits Then
      Credits = Credits + 1

      If Credits > 12 Then
        If QuoteTimer.Enabled = False Then
                Credits = 10
                CreditBasher = CreditBasher + 1
                If CreditBasher > 4 Then CreditBasher = 1
                Dampenmusic2 = 0.2
                Select Case CreditBasher
                  Case 1
                    PlaySound "Sfx_BobParty",0,RomSoundVolume
                    QuoteTimer.Interval = 3364 : QuoteTimer.Enabled = True
                  Case 2
                    PlaySound "Sfx_KrabbyLand",0,RomSoundVolume
                    QuoteTimer.Interval = 9270 : QuoteTimer.Enabled = True
                  Case 3
                    PlaySound "Sfx_looser",0,RomSoundVolume
                    QuoteTimer.Interval = 8047 : QuoteTimer.Enabled = True
                  Case 4
                    PlaySound "Sfx_moneyfish",0,RomSoundVolume
                    QuoteTimer.Interval = 6795 : QuoteTimer.Enabled = True
                End Select
        End If
      End If

      Select Case Int(rnd*3)
        Case 0: PlaySound ("Coin_In_1"), 0, RomSoundVolume, 0, 0.25
        Case 1: PlaySound ("Coin_In_2"), 0, RomSoundVolume, 0, 0.25
        Case 2: PlaySound ("Coin_In_3"), 0, RomSoundVolume, 0, 0.25
      End Select
      If  startupscreen = 2 And attractcreditNO = 0 Then If attractcredit > 2 Then attractcredit = 2 Else attractcredit = 1
    End If
  End If


  'VR Specific...
  If keycode = StartGameKey and VRRoom > 0 Then VRStartButton.y = VRStartButton.y + 5
  If keycode = PlungerKey and VRRoom > 0 Then VRLaunchButtonText.y = VRLaunchButtonText.y +5: VRLaunchButton.y = VRLaunchButton.y +5
  If keycode = LeftFlipperKey and VRRoom > 0 Then VR_FlipperLeft.x = VR_FlipperLeft.x -5
  If keycode = RightFlipperKey and VRRoom > 0 Then VR_FlipperRight.x = VR_FlipperRight.x +5

    If keycode = LeftMagnaSave And Not bInOptions Then bOptionsMagna = False
  If keycode = LeftMagnaSave And inQRClaimDMD Then StopQRclaim
  If keycode = RightMagnaSave And inQRClaimDMD Then StopQRclaim
    If keycode = RightMagnaSave And Not bInOptions Then bOptionsMagna = False

' If keycode = PlungerKey And Plungerpush = 1 Then
' End If

  If keycode = LeftTiltKey Then LeftPress = 0
  If keycode = RightTiltKey Then RightPress = 0
  If keycode = CenterTiltKey Then CenterPress = 0

  If Keycode = 30 Then
    If StagedFlipperMod = 1 Then
      SolULFlipper False
    End If
  End If

  If keycode = LeftFlipperKey Then
    cancelL = 0
    EnterKey = 0
    RepeatTimer.enabled = False

'   If Not bTilted Then
'     If ScoreNumber = "" Then
'       If CurrentMode <> 2 Then
          FlipperDeActivate LeftFlipper, LFPress
          SolLFlipper False
          If StagedFlipperMod <> 1 Then
            SolULFlipper False
          End If
'       End If
'     If ScoreNumber <> "" Then
'       If HSDigit = 0 Then HSDigit = 1
'       HSChr = HSChr - 1
'       If HSChr = 64 Then HSChr = 90
'       EnterHS
'     End If
'   End If
  End If

  If keycode = RightFlipperKey Then
    cancelR = 0
    EnterKey = 0
    RepeatTimer.enabled = False

'   If Not bTilted Then
'     If ScoreNumber = "" Then
'       If CurrentMode <> 2 Then
          FlipperDeActivate RightFlipper, RFPress
          SolRFlipper False
'       End If
'     End If
'     If ScoreNumber <> "" Then
'       If HSDigit = 0 Then HSDigit = 1
'       HSChr = HSChr + 1
'       If HSChr = 91 Then HSChr = 65
'       EnterHS
'     End If
'   End If
  End If
  If keycode = StartGameKey Then
    If CurrentBall < 3 And Currentball > 0 And StartbuttonHold < frame Then
      RestartGame = 1
      DMDMessage1 "RESTART"   ,""                   ,FontSponge16bb ,52,16  ,55,16  ,"noslide"      ,5  ,"BG001"  ,"novideo"
      TiltMachine
    End If


    If EnterHighscore > 101 And EnterHighscore < 127 Then
      EnterKey = 0
      EnterRepeat = 0
      RepeatTimer.enabled = False

      If EnterText(1) = "" Then
        EnterHighblinking = 2
        EnterText(1) = mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789",EnterHighPos+18,1)
      Elseif EnterText(2) = "" Then
        EnterHighblinking = 3
        EnterText(2) = mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789",EnterHighPos+18,1)
      Elseif EnterText(3) = "" Then
        EnterHighblinking = 4
        EnterText(3) = mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789",EnterHighPos+18,1)
      End If
    End If

    If StartupScreen = 0 Or StartupScreen = 2 Then
      If ScoreNumber = "" Then
        If CurrentBall = 3 and Players < 4  and DoOrDieMode < 3 Then


          If Players < 1 Then
            bTilted = false
          End If

          If UseCredits Then
            If Credits < 1 Then
              BumperSfx
              PlaySound "Sfx_Steel_Sting",0,RomSoundVolume
              If attractcredit = 0 And attractcreditNO = 0 Then attractcreditNO = 1
              Exit Sub
            End If
            Credits = Credits - 1
            If Credits > 9 Then Credits = 9
          End If

          Players = Players + 1
          If Players = 1 Then
            ScoreText.Text = 0
            startupscreen = 3 ' turn off attractmode
            DMD_ResetAll
            DoOrDieScore = 0
          End If
          If Players = 2 Then
            FlexDMD.Stage.GetLabel("Score_2").visible = True
            PlaySound "Sfx_wow",0,RomSoundVolume
            DoOrDieMode = 2
            ScoreText2.Text = 0
            StopSound "Sfx_NewPlayer": PlaySound "Sfx_NewPlayer",0,RomSoundVolume
          End If
          If Players = 3 Then
            FlexDMD.Stage.GetLabel("Score_3").visible = True
            PlaySound "Sfx_Coolestparty",0,RomSoundVolume
            ScoreText3.Text = 0
            StopSound "Sfx_NewPlayer": PlaySound "Sfx_NewPlayer",0,RomSoundVolume
          End If
          If Players = 4 Then
            FlexDMD.Stage.GetLabel("Score_4").visible = True
            Stopsound "Sfx_Coolestparty" : PlaySound "Sfx_Coolestparty",0,RomSoundVolume
            ScoreText4.Text = 0
            StopSound "Sfx_NewPlayer": PlaySound "Sfx_NewPlayer",0,RomSoundVolume
          End If
          If StartGame = 0 Then
            If RoosterTimer.Enabled = False Then
              StopTableMusic
              StopTableSequences()
              HSDisplay = 0 : DisplayHS.Enabled = False
              HSMusicPlaying.Enabled = False
              DMDMessage1 "GAMESTART"     ,""       ,FontSponge16bb ,52,16  ,52,16  ,"solid"  ,2  ,"BG001"  ,"novideo"
              PlaySound "Sfx_StartGame",0,RomSoundVolume
              RoosterTimer.Enabled = True
              DoOrDieMode = 1
              FlexDMD.Stage.GetLabel("Score_1").visible = True
              FlexDMD.Stage.GetLabel("Score_2").visible = False
              FlexDMD.Stage.GetLabel("Score_3").visible = False
              FlexDMD.Stage.GetLabel("Score_4").visible = False

              FlexDMD.Stage.GetFrame("VSeparator").Thickness = 1
              FlexDMD.Stage.Getimage("dod1").visible = False
              FlexDMD.Stage.Getimage("dod2").visible = False
              FlexDMD.Stage.Getimage("dod3").visible = False
              Light001.state = 0 : Primitive001.visible = False
            End If
          End If
        End If
      End If
    End If
'   If HSDigit >= 1 Then
'     HSDigit = HSDigit + 1
'     HSChr = 65
'     If HSDigit > 4 Then HSDigit = 4
'     EnterHS
'   End If
    UpdateText
  End If
End Sub
Dim DoOrDieMode
DoOrDieMode = 0


Sub RoosterTimer_Timer
  RoosterTimer.Enabled = False
  BallRelease.CreateBall
  BallRelease.Kick 80,6
      DOF 111,2
  PlaySoundAt SoundFX("popper",DOFContactors), BallRelease
  ssREADY = 1
  bTilted = False
  PlayTableMusic(13)
  Call ChangeGILights(1)
  StartGame = 1
  DOF 121,0
  CurrentPlayer = 1
  ClearDMD
  DMDMessage1 "PLAYER 1"      ,""       ,FontSponge16bb ,65,16  ,65,16  ,"blink"  ,2  ,"BG001"  ,"novideo"
  UpdateText
  SkillshotWait.enabled = True
  SJPgotten = 0

    If Not IsNull(Scorbit) Then
        If ScorbitActive = 1 And Scorbit.bNeedsPairing = False Then
            Scorbit.StartSession()
        End If
    End If

End Sub

Sub SkillshotWait_Timer
  SkillshotWait.enabled = False
  lightCtrl.Blink f147: DOF 147,2
  SkillShot = 1
  SkillShotMessage.Enabled = False : SkillShotTimer.Enabled = True
  SkillShot2 = Int(rnd(1)*6)
End Sub


Sub PlungerTrigger_hit()
  BallinLane = BallinLane + 1
  DOF 122,1
' lightCtrl.Blink f147
End Sub
Sub PlungerTrigger_unhit()
  DOF 122,0
  BallinLane = BallinLane - 1
' lightCtrl.Blink f147
End Sub


'********************
'Current Ball Text
'********************
Sub UpdateText
  If StartGame = 1 And 4 - CurrentBall < 4 Then
    BallsText.Text = 4 - CurrentBall
  Else
    BallsText.Text = ""
  End If
End Sub

'********************
'Lights
'********************
Dim collect1 : collect1 = 0
Dim collect2 : collect2 = 0
Dim collect3 : collect3 = 0
Dim collect4 : collect4 = 0
Dim collect5 : collect5 = 0
Dim collect6 : collect6 = 0
Dim collect7 : collect7 = 0
Dim collect8 : collect8 = 0
Dim collect9 : collect9 = 0


Sub UpdateLights
  If AreLightsOff = 0 Then
    'l01.State = L01State(CurrentPlayer) done in ballsave timer
    l02.State = L02State(CurrentPlayer)
    l03.State = L03State(CurrentPlayer)
    l04.State = L04State(CurrentPlayer)
    l05.State = L05State(CurrentPlayer)
    l06.State = L06State(CurrentPlayer)
    l07.State = L07State(CurrentPlayer)
    l08.State = L08State(CurrentPlayer)
    l09.State = L09State(CurrentPlayer)
    l10.State = L10State(CurrentPlayer)
    l11.State = L11State(CurrentPlayer)
    l12.State = L12State(CurrentPlayer)
    l13.State = L13State(CurrentPlayer)
    If KickerSave(CurrentPlayer) < 1 Then l14.State = 0 Else l14.state = 2
    l15.State = L15State(CurrentPlayer)
    If KickerSave(CurrentPlayer) < 1 Then l16.State = 0 Else l16.state = 2
    l17.State = L17State(CurrentPlayer)
    l33.State = L33State(CurrentPlayer)
    l34.State = L34State(CurrentPlayer)
    l35.State = L35State(CurrentPlayer)

    If currentmode = 11 then
      l18.State = 2 * collect1 : l19.State = 2 * collect1 : l20.State = 2 * collect1
      l24.State = 2 * collect2 : l25.State = 2 * collect2 : l26.State = 2 * collect2
      l27.State = 2 * collect3 : l28.State = 2 * collect3

      If AddABallReady = 1 then
        l29.State = 2 : l29.blinkinterval = 60
        l45.State = 2 : l45.blinkinterval = 60
      Elseif MBRampsforAddaBall < 3 + MBaddaball(CurrentPlayer) then
          l29.State = 2 : l29.blinkinterval = 90
          l45.State = 2 : l45.blinkinterval = 90
      Else
          l29.State = 2 * collect3 : l29.blinkinterval = 123
          l45.State = 2 * collect7 : l45.blinkinterval = 123
      End If


      If SpinnerHurryUp > 0 Then ' krusty
        l30.State = 2
        l31.State = 2
        l32.State = 2
        l30.blinkinterval = 77 : l31.blinkinterval = 77 : l32.blinkinterval = 77
      Else
        l30.State = 2 * collect4
        l31.State = 2 * collect4
        l32.State = 2 * collect4
        l30.blinkinterval = 125 : l31.blinkinterval = 125 : l32.blinkinterval = 125
      End If


      l37.State = 2 * collect5 : l38.State = 2 * collect5 : l39.State = 2 * collect5
      l40.state = 2 * collect6    ' sponge
      l41.state = 2 * collect6
      l42.state = 2 * collect6
      l43.State = 2 * collect7 : l44.State = 2 * collect7  ' 45 moved Up


      l46.State = 2 * collect8: l47.State = 2 * collect8 : l48.State = 2 * collect8
      l21.State = 2 * collect9 : l22.State = 2 * collect9 : l23.State = 2 * collect9


      If MB_blinkbankLights = 1 Then If L33State(CurrentPlayer) = 1 or L34State(CurrentPlayer) = 1 or L35State(CurrentPlayer) = 1 Then MB_blinkbankLights = 0
      If MB_blinkbankLights = 1 Then
        l33.State = 2 ' bank blink
        l34.State = 2
        l35.State = 2
        l36.State = 2 ' lock light
      Else
        If collect1 + collect2 + collect3 + collect4 + collect5 + collect6 + collect7 + collect8 + collect9 = 0 Then
          If L33State(CurrentPlayer) = 1 or L34State(CurrentPlayer) = 1 or L35State(CurrentPlayer) = 1 Then
            If L33State(CurrentPlayer) = 1 Then l33.State = 1 Else  l33.State = 2
            If L34State(CurrentPlayer) = 1 Then l34.State = 1 Else  l34.State = 2
            If L35State(CurrentPlayer) = 1 Then l35.State = 1 Else  l35.State = 2
            l36.State = 2 ' lock light
          Else
            l33.State = 2 ' bank blink
            l34.State = 2
            l35.State = 2
            l36.State = 2 ' lock light
          End If
        Else
          l33.State = 0
          l34.State = 0
          l35.State = 0
          l36.State = 1
        End If
      End If

    Elseif CurrentMode = 21 Then
        If StopPinappleWiz2 = 0 Then
          If wiz2Addaball = 4 Or wiz2Addaball = 14 Then
            l33.State = 2 : L33.blinkinterval = 60
            l34.State = 2 : L34.blinkinterval = 60
            l35.State = 2 : L35.blinkinterval = 60
          Elseif wiz2Addaball < 15 Then
            l33.State = 2 : L33.blinkinterval = 220
            l34.State = 2 : L34.blinkinterval = 220
            l35.State = 2 : L35.blinkinterval = 220
          Else
            l33.State = 0 : L33.blinkinterval = 220
            l34.State = 0 : L34.blinkinterval = 220
            l35.State = 0 : L35.blinkinterval = 220
          End If

        Else
          l33.State = 0 : L33.blinkinterval = 125
          l34.State = 0 : L34.blinkinterval = 125
          l35.State = 0 : L35.blinkinterval = 125

        End If

      Select Case collect1
        case 0 : l18.State = 0 : l19.State = 0 : l20.State = 0
        case 1 : l18.State = 2 : l19.State = 2 : l20.State = 2 : L18.blinkinterval = 60 : L20.blinkinterval = 60 : L20.blinkinterval = 60
        case 2 : l18.State = 2 : l19.State = 2 : l20.State = 2 : L18.blinkinterval = 110 : L20.blinkinterval = 110 : L20.blinkinterval = 110
      End Select
      Select Case collect2
        case 0 : l24.State = 0 : l25.State = 0 : l26.State = 0
        case 1 : l24.State = 2 : l25.State = 2 : l26.State = 2 : L24.blinkinterval = 60 : L25.blinkinterval = 60 : L26.blinkinterval = 60
        case 2 : l24.State = 2 : l25.State = 2 : l26.State = 2 : L24.blinkinterval = 110 : L25.blinkinterval = 110 : L26.blinkinterval = 110
      End Select
      Select Case collect3
        case 0 : l27.State = 0 : l28.State = 0 : l29.State = 0
        case 1 : l27.State = 2 : l28.State = 2 : l29.State = 2 : L27.blinkinterval = 60 : L28.blinkinterval = 60 : L29.blinkinterval = 60
        case 2 : l27.State = 2 : l28.State = 2 : l29.State = 2 : L27.blinkinterval = 110 : L28.blinkinterval = 110 : L29.blinkinterval = 110
      End Select
      If wizaddaball < 12 Then L29.blinkinterval = 60 : l29.state = 2
      Select Case collect4
        case 0 : l30.State = 0 : l31.State = 0 : l32.State = 0
        case 1 : l30.State = 2 : l31.State = 2 : l32.State = 2 : L30.blinkinterval = 60 : L31.blinkinterval = 60 : L32.blinkinterval = 60
        case 2 : l30.State = 2 : l31.State = 2 : l32.State = 2 : L30.blinkinterval = 110 : L31.blinkinterval = 110 : L32.blinkinterval = 110
      End Select
      Select Case collect5
        case 0 : l37.State = 0 : l38.State = 0 : l39.State = 0
        case 1 : l37.State = 2 : l38.State = 2 : l39.State = 2 : L37.blinkinterval = 60 : L38.blinkinterval = 60 : L39.blinkinterval = 60
        case 2 : l37.State = 2 : l38.State = 2 : l39.State = 2 : L37.blinkinterval = 110 : L38.blinkinterval = 110 : L39.blinkinterval = 110
      End Select
      Select Case collect6
        case 0 : l40.State = 0 : l41.State = 0 : l42.State = 0
        case 1 : l40.State = 2 : l41.State = 2 : l42.State = 2 : L40.blinkinterval = 60 : L41.blinkinterval = 60 : L42.blinkinterval = 60
        case 2 : l40.State = 2 : l41.State = 2 : l42.State = 2 : L40.blinkinterval = 110 : L41.blinkinterval = 110 : L42.blinkinterval = 110
      End Select
      Select Case collect7
        case 0 : l43.State = 0 : l44.State = 0 : l45.State = 0
        case 1 : l43.State = 2 : l44.State = 2 : l45.State = 2 : L43.blinkinterval = 60 : L44.blinkinterval = 60 : L45.blinkinterval = 60
        case 2 : l43.State = 2 : l44.State = 2 : l45.State = 2 : L43.blinkinterval = 110 : L44.blinkinterval = 110 : L45.blinkinterval = 110
      End Select
      If wizaddaball < 12 Then L45.blinkinterval = 60 : l45.state = 2
      Select Case collect8
        case 0 : l46.State = 0 : l47.State = 0 : l48.State = 0
        case 1 : l46.State = 2 : l47.State = 2 : l48.State = 2 : L46.blinkinterval = 60 : L47.blinkinterval = 60 : L48.blinkinterval = 60
        case 2 : l46.State = 2 : l47.State = 2 : l48.State = 2 : L46.blinkinterval = 110 : L47.blinkinterval = 110 : L48.blinkinterval = 110
      End Select
      Select Case collect9
        case 0 : l21.State = 0 : l22.State = 0 : l23.State = 0
        case 1 : l21.State = 2 : l22.State = 2 : l23.State = 2 : L21.blinkinterval = 60 : L22.blinkinterval = 60 : L23.blinkinterval = 60
        case 2 : l21.State = 2 : l22.State = 2 : l23.State = 2 : L21.blinkinterval = 110 : L22.blinkinterval = 110 : L23.blinkinterval = 110
      End Select

      If AddABallReady = 1 then
        l29.State = 2 : l29.blinkinterval = 60
        l45.State = 2 : l45.blinkinterval = 60
      End If

    Else


      l18.State = L18State(CurrentPlayer)     'tiki
      l19.State = L19State(CurrentPlayer)
      If (ModeReady(CurrentPlayer) = 1 And l20state(currentplayer) = 2) or ModeJP_Done = False Then l20.state = 0 Else l20.State = L20State(CurrentPlayer)
      l24.State = L24State(CurrentPlayer)     ' boating
      l25.State = L25State(CurrentPlayer)
      If RampComboSelect> 0 And RampComboSelect <> 3 Then
        L26.blinkinterval = 60 : L26.state = 2
      Else
        l26.blinkinterval = 110
        If (ModeReady(CurrentPlayer) = 1 And l26state(currentplayer) = 2 ) or ModeJP_Done = False Then
          l26.state = 0
        Else
          l26.state = L26State(CurrentPlayer)
        End If
      End If


      l27.State = L27State(CurrentPlayer)     ' plankton
      l28.State = L28State(CurrentPlayer)
      If RampComboSelect> 0 And RampComboSelect <> 2 Then
        L29.state = 2 : L29.blinkinterval = 60
      Else
        L29.blinkinterval = 110
        If (ModeReady(CurrentPlayer) = 1 And l29state(currentplayer) = 2 ) or ModeJP_Done = False Then
          l29.state = 0
        Else
          l29.state = L29State(CurrentPlayer)
        End If
      End If

      If SpinnerHurryUp < 1 Then
        l30.blinkinterval = 125 : l31.blinkinterval = 125 : l32.blinkinterval = 125
        l30.State = L30State(CurrentPlayer)     ' krusty
        l31.State = L31State(CurrentPlayer)
        If (ModeReady(CurrentPlayer) = 1 And l32state(currentplayer) = 2) or ModeJP_Done = False  Then
          l32.state = 0
        Else
          l32.state = L32State(CurrentPlayer)
        End If
      Else
        l30.State = 2
        l31.State = 2
        l32.State = 2
        l30.blinkinterval = 77 : l31.blinkinterval = 77 : l32.blinkinterval = 77
      End If




      If ModeReady(CurrentPlayer)=0 OR CurrentMode = 11 OR MultiBallReady(CurrentPlayer) = 1 then
        l37.State = L37State(CurrentPlayer) : l37.BlinkPattern = "10" 'mr krab
        l38.State = L38State(CurrentPlayer) : l38.BlinkPattern = "10"
        If RampComboSelect> 0 And RampComboSelect <> 5 Then
          L39.blinkinterval =  60 : L39.state = 2
        Else
          l39.blinkinterval = 110
          If (ModeReady(CurrentPlayer) = 1 And l39state(currentplayer) = 2 ) or ModeJP_Done = False Then
            l39.state = 0
          Else
            l39.state = L39State(CurrentPlayer) : l39.BlinkPattern = "10"
          End If
        End If
      End If


      If ScoopRestart > frame Then ' sponge
        l40.state = 2 : l40.BlinkPattern = "100"
        l41.state = 2 : l41.BlinkPattern = "010"
        l42.state = 2 : l42.BlinkPattern = "001"
        l40.blinkinterval = 110
        l41.blinkinterval = 110
        l42.blinkinterval = 110
      Else
        If MysteryReady(CurrentPlayer) = 1 Then
          l40.state = 2 : l40.BlinkPattern = "10"
          l41.state = 2 : l41.BlinkPattern = "10"
          l42.state = 2 : l42.BlinkPattern = "10"
          l40.blinkinterval = 60
          l41.blinkinterval = 60
          l42.blinkinterval = 60
        Else
          l40.State = L40State(CurrentPlayer) : l40.BlinkPattern = "10"
          l41.State = L41State(CurrentPlayer) : l41.BlinkPattern = "10"

          If( ModeReady(CurrentPlayer) = 1 And l42state(currentplayer) = 2) or ModeJP_Done = False  Then
            l42.State = 0
          Else
            l42.State = L42State(CurrentPlayer)
          End If
          l42.BlinkPattern = "10"
          l40.blinkinterval = 110
          l41.blinkinterval = 110
          l42.blinkinterval = 110

        End If
      End If


      l43.State = L43State(CurrentPlayer)   ' sandy
      l44.State = L44State(CurrentPlayer)
      If RampComboSelect> 0 And RampComboSelect <> 1 Then
        L45.blinkinterval = 60 : L45.state = 2
      Else
        l45.blinkinterval = 110

        If ( ModeReady(CurrentPlayer) = 1 And l45state(currentplayer) = 2 ) or ModeJP_Done = False Then
          l45.state = 0
        Else
          l45.state = L45State(CurrentPlayer)
        End If
      End If

      l46.State = L46State(CurrentPlayer)   ' yellyfish
      l47.State = L47State(CurrentPlayer)
      If RampComboSelect> 0 And RampComboSelect <> 4 Then
        L48.blinkinterval = 60
        L48.state = 2
      Else
        l48.blinkinterval = 110
        If ( ModeReady(CurrentPlayer) = 1 And l48state(currentplayer) = 2 ) or ModeJP_Done = False Then
          l48.state = 0
        Else
          l48.state = L48State(CurrentPlayer)
        End If
      End If


      l21.State = L21State(CurrentPlayer)     ' goobers
      l22.State = L22State(CurrentPlayer)
      If (ModeReady(CurrentPlayer) = 1 And l23State(currentplayer) = 2 ) or ModeJP_Done = False Then
        l23.state = 0
      Else
        l23.state = L23State(CurrentPlayer)
      End If

      l33.State = L33State(CurrentPlayer)
      l34.State = L34State(CurrentPlayer)
      l35.State = L35State(CurrentPlayer)
      L33.blinkinterval = 125
      L34.blinkinterval = 125
      L35.blinkinterval = 125


    End If

    l36.State = L36State(CurrentPlayer)

    If ModeReady(CurrentPlayer) = 0 OR CurrentMode = 11 OR MultiBallReady(CurrentPlayer) = 1 Or CurrentMode > 19 then lightCtrl.LightOff f146
    l49.State = L49State(CurrentPlayer)
    l50.State = L50State(CurrentPlayer)
    If ModeReady(CurrentPlayer) = 1 Then
      If CurrentMode = 4 Then
        DMDMessage1 "MR KRABS"    ,"" ,FontSponge16bb ,54,17  ,54,17  ,"shake"  ,2  ,"BG001"  ,"novideo"
        DMDMessage1 "MODE"  ,"COMPLETE" ,FontSponge12bb ,54,17  ,54,17  ,"blink"  ,2  ,"BG001"  ,"novideo"
        CurrentMode = 0 : IntroMusic : PlaySound "Sfx_ModeReady",0,RomSoundVolume
        modedone4(CurrentPlayer) = 1
        l37state(CurrentPlayer) = 1 : l37.state = 1
        l38state(CurrentPlayer) = 1 : l38.state = 1
        l39state(CurrentPlayer) = 1 : l39.state = 1
        modetimecounter = 0
      End If
      ModeReadyLights

    End If
    If CurrentMode <> 0 Then BonusLights
    If SJPReady(CurrentPlayer) = 1 Then
      If CurrentMode = 11 Then L33.State = 2 : L34.State = 2 : L35.State = 2 : L36.State = 2
      If CurrentMode = 0 Then ModeReadyLights
    End If

  Else
    lightCtrl.LightOff f146
  End If
End Sub

Sub ModeReadyLights
  If currentmode = 11 Then Exit Sub
  If ModeReady(CurrentPlayer) = 1 Then
    If CurrentMode < 11 And MultiBallReady(CurrentPlayer) = 0 Then
      l04.state = 2
      l37.state = 2 : l37.BlinkPattern = "100"
      l38.state = 2 : l38.BlinkPattern = "010"
      l39.state = 2 : l39.BlinkPattern = "001"
      l39.blinkinterval = 110
      lightCtrl.Blink f146 : DOF 146,2
    End If
    If SJPReady(CurrentPlayer) = 1 Then l03.state = 2
  End If
  ModeLights
End Sub


Sub ModeLights
  If MysteryReady(CurrentPlayer) = 1 Then
    l40.state = 2 : l40.blinkinterval = 60
    l41.state = 2 : l41.blinkinterval = 60
    l42.state = 2 : l42.blinkinterval = 60
  Else
    l40.blinkinterval = 60
    l41.blinkinterval = 60
    l42.blinkinterval = 60
  End If
  If modedone1(CurrentPlayer) = 1 Then L12State(CurrentPlayer) = 1
  If modedone2(CurrentPlayer) = 1 Then L06State(CurrentPlayer) = 1
  If modedone3(CurrentPlayer) = 1 Then L08State(CurrentPlayer) = 1
  If modedone4(CurrentPlayer) = 1 Then L07State(CurrentPlayer) = 1
  If modedone5(CurrentPlayer) = 1 Then L13State(CurrentPlayer) = 1
  If modedone6(CurrentPlayer) = 1 Then L10State(CurrentPlayer) = 1
  If modedone7(CurrentPlayer) = 1 Then L05State(CurrentPlayer) = 1
  If modedone8(CurrentPlayer) = 1 Then L09State(CurrentPlayer) = 1
  If modedone9(CurrentPlayer) = 1 Then L11State(CurrentPlayer) = 1

  If CurrentMode = 1 Then L12State(CurrentPlayer) = 2
  If CurrentMode = 2 Then L06State(CurrentPlayer) = 2   ' CHANGED
  If CurrentMode = 3 Then L08State(CurrentPlayer) = 2
  If CurrentMode = 4 Then L07State(CurrentPlayer) = 2
  If CurrentMode = 5 Then L13State(CurrentPlayer) = 2
  If CurrentMode = 6 Then L10State(CurrentPlayer) = 2   ' CHANGED
  If CurrentMode = 7 Then L05State(CurrentPlayer) = 2
  If CurrentMode = 8 Then L09State(CurrentPlayer) = 2
  If CurrentMode = 9 Then L11State(CurrentPlayer) = 2

End Sub

Sub ResetModeLights
    modedone1(CurrentPlayer)=0
    modedone2(CurrentPlayer)=0
    modedone3(CurrentPlayer)=0
    modedone4(CurrentPlayer)=0
    modedone5(CurrentPlayer)=0
    modedone6(CurrentPlayer)=0
    modedone7(CurrentPlayer)=0
    modedone8(CurrentPlayer)=0
    modedone9(CurrentPlayer)=0
    L05State(CurrentPlayer) = 0
    L06State(CurrentPlayer) = 0
    L07State(CurrentPlayer) = 0
    L08State(CurrentPlayer) = 0
    L09State(CurrentPlayer) = 0
    L10State(CurrentPlayer) = 0
    L11State(CurrentPlayer) = 0
    L12State(CurrentPlayer) = 0
    L13State(CurrentPlayer) = 0
End Sub

Sub BonusLights
    If CurrentMode = 1 Then l40.State = 2 : l41.State = 2 : l42.State = 2
    If CurrentMode = 2 Then l27.State = 2 : l28.State = 2 : l29.State = 2
    If CurrentMode = 3 Then l43.State = 2 : l44.State = 2 : l45.State = 2
    If CurrentMode = 4 Then l37.State = 2 : l38.State = 2 : l39.State = 2
    If CurrentMode = 5 Then l46.State = 2 : l47.State = 2 : l48.State = 2
    If CurrentMode = 6 Then l24.State = 2 : l25.State = 2 : l26.State = 2
    If CurrentMode = 7 Then l18.State = 2 : l19.State = 2 : l20.State = 2
    If CurrentMode = 8 Then l21.State = 2 : l22.State = 2 : l23.State = 2
    If CurrentMode = 9 Then l30.State = 2 : l31.State = 2 : l32.State = 2

End Sub

dim gilvl:gilvl = 1

Sub TurnLightsOff
    startB2S(4)
  ChangeGILights(0)
  For each i in InsertLights
    i.State = 0
  Next
' fixing .. skip on warning/tilt ?
End Sub





Sub ChangeGILights(GILightState)
    Dim light
    gilvl = GILightState
  Sound_GI_Relay GILightState, LeftFlipper
    If GILightState = 1 Then
    DOF 120,1
    startB2S(1)
        For each light in GiStrings
            lightCtrl.LightOn light
        Next
        lightCtrl.LightOn L115
    If VRRoom > 0 Then
      BGBright.visible = 1
    End IF
    Light003.state = 1
    Else
    Dof 120,0
        For each light in GiStrings
            lightCtrl.LightOff light
        Next
        lightCtrl.LightOff L115
    If VRRoom > 0 Then
      BGBright.visible = 0
    End IF
    Light003.state = 0
    End If

End Sub

Sub JackpotSequence(SequenceTime)
  InsertSequence.Play SeqRandom,10,,SequenceTime
  GISequence.Play SeqRandom,10,,SequenceTime
End Sub


Sub SJPSequence
  If SJPReady(CurrentPlayer) = 2 Then
    If SeqNum = 1 Then
      InsertSequence.UpdateInterval = 10 : GISequence.UpdateInterval = 10
      SJPLight.Interval = 9603 : SJPLight.Enabled = True
      InsertSequence.Play SeqHatch1VertOn,10,100
      GISequence.Play SeqHatch1VertOn,10,100
    End If
    If SeqNum = 2 Then
            StopTableSequences()
      SJPLight.Interval = 5830 : SJPLight.Enabled = True
      InsertSequence.Play SeqCircleInOn,10,100
      GISequence.Play SeqCircleInOn,10,100
    End If
    If SeqNum = 3 Then
      StopTableSequences()
      InsertSequence.UpdateInterval = 20 : GISequence.UpdateInterval = 20
      InsertSequence.Play SeqRandom,10,,3177
      GISequence.Play SeqRandom,10,,3177
      SeqNum = 0
    End If
  Else
    If SeqNum = 1 Then
      PlaySound "Sfx_SJP2",0,RomSoundVolume
      InsertSequence.UpdateInterval = 10 : GISequence.UpdateInterval = 10
      SJPLight.Interval = 7085 : SJPLight.Enabled = True
      InsertSequence.Play SeqCircleInOn,10,100
      GISequence.Play SeqCircleInOn,10,100
    End If
    If SeqNum = 2 Then
            StopTableSequences()
      SJPLight.Interval = 10719 : SJPLight.Enabled = True
      InsertSequence.Play SeqHatch1VertOn,10,100
      GISequence.Play SeqHatch1VertOn,10,100
    End If
    If SeqNum = 3 Then
      StopTableSequences()
      InsertSequence.UpdateInterval = 20 : GISequence.UpdateInterval = 20
      InsertSequence.Play SeqRandom,10,,3000
      GISequence.Play SeqRandom,10,,3000
      SeqNum = 0
    End If
  End If
End Sub

Sub SJPLight_Timer
  SJPLight.Enabled = False
  SeqNum = SeqNum + 1
  Call SJPSequence
End Sub

Sub gi038_Timer
  gi038.TimerEnabled = False
  If CurrentMode <> 2 Then gi038.State = 1
  If CurrentMode = 2 Or SJPReady(CurrentPlayer) > 0 Then gi038.State = 0
End Sub

Sub ThunderTimer_Timer
  ThunderTimer.Enabled = False
  If CurrentMode = 2 Then
    PlaySound "Sfx_Thunder",0,RomSoundVolume
    gi038.State = 1 : gi038.TimerEnabled = True
'   Nudge int(rnd*360), 5
    DOF 134,2 ' shaker
    ThunderTimer.Interval = int(rnd*10000)
    ThunderTimer.Enabled = True
  End If
End Sub

Sub MBIntroLights_Timer
  MBIntroLights.Enabled = False
  JackpotSequence(1394)
End Sub

Sub RotateLaneLights
  dim L49s
  L49s = L49State(CurrentPlayer)
  L49State(CurrentPlayer) = L50State(CurrentPlayer)
  L50State(CurrentPlayer) = L49s
  UpdateLights
End Sub


'********************
'Slings & div switches
'********************

Dim LStep, RStep, L1Step

Sub LeftSlingShot_Slingshot
  LightSeqmodes.play SeqBlinking,,1,30
  DOF 103,2

  SetGI Rgb(55,10,10),99
  If Slinger < 1 Then Slinger = 1
  startB2S(3)
  LS.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotLeft Lemk

  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else
    AddScore(101230)
  End If
  SlingSfx
'    LeftSling4.Visible = 1
'    Lemk.RotX = 26
    LStep = 0
  LeftSlingShot_Timer
    LeftSlingShot.TimerEnabled = 1
    b2snormal.Enabled=True
  If VRRoom > 0 Then
    VRBGFLPat
  End If
End Sub

Sub LeftSlingShot_Timer
  Dim y2, y3, y4, ty, bl
    Select Case LStep
    Case 0: y2=0: y3=0: y4=1: ty=-20
        Case 1: y2=0: y3=1: y4=0: ty=-10
        Case 2: y2=1: y3=0: y4=0: ty=0
        Case 3: y2=0: y3=0: y4=0: ty=0   :LeftSlingShot.TimerEnabled = 0 : lightCtrl.Pulse f141, 0 : DOF 141,2
    End Select
  For each bl in LeftSling2_bl : bl.visible = y2 : Next
  For each bl in LeftSling3_bl : bl.visible = y3 : Next
  For each bl in LeftSling4_bl : bl.visible = y4 : Next
  For each bl in Lemk_bl : bl.transy = -ty : Next
    LStep = LStep + 1
End Sub

'Sub LeftSlingShot_Timer
'    Select Case LStep
'        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
'        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
'        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
'    End Select
'
'    LStep = LStep + 1
'End Sub

Sub RightSlingShot_Slingshot
  DOF 105,2

  LightSeqmodes.play SeqBlinking,,1,30

  SetGI Rgb(55,10,10),99
  If Slinger < 1 Then Slinger = 1

  startB2S(2)
  RS.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotRight Remk

  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else
    AddScore(101230)
  End If
  SlingSfx

'    RightSling4.Visible = 1
'    Remk.RotX = 26
    RStep = 0
  RightSlingShot_Timer
    RightSlingShot.TimerEnabled = 1
    b2snormal.Enabled=True
  If VRRoom > 0 Then
    VRBGFLSponge
  End If
End Sub


Sub RightSlingShot_Timer
  Dim y2, y3, y4, ty, bl
    Select Case RStep
    Case 0: y2=0: y3=0: y4=1: ty=-20
        Case 1: y2=0: y3=1: y4=0: ty=-10
        Case 2: y2=1: y3=0: y4=0: ty=0
        Case 3: y2=0: y3=0: y4=0: ty=0   :RightSlingShot.TimerEnabled = 0 : lightCtrl.Pulse f142, 0: DOF 142,2
    End Select
  For each bl in RightSling2_bl : bl.visible = y2 : Next
  For each bl in RightSling3_bl : bl.visible = y3 : Next
  For each bl in RightSling4_bl : bl.visible = y4 : Next
  For each bl in Remk_bl : bl.transy = -ty : Next
    RStep = RStep + 1
End Sub

'
'Sub RightSlingShot_Timer
'    Select Case RStep
'        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
'        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
'        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
'    End Select
'
'    RStep = RStep + 1
'End Sub

Sub LeftSlingShot1_Slingshot
  DOF 107,2

    startB2S(4)
    b2snormal.Enabled=True
  lightCtrl.Pulse f143, 0: DOF 143,2
  BumperSfx
  SetGI Rgb(0,0,255),76

  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else

    If CurrentMode <> 2 Then gi038.State = 0
    If CurrentMode = 2 Or SJPReady(CurrentPlayer) > 0 Then gi038.State = 1

    If CurrentMode = 5 Then
      AddScore(500000)
      bumperjackpotcount = bumperjackpotcount + 1
    Else
      AddScore(50110)
    End If
  End If

  gi038.TimerEnabled = True
' LS1.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotLeft Lemk1
  SlingSfx
'    LeftSling14.Visible = 1
'    Lemk1.RotX = 26
    L1Step = 0
  LeftSlingShot1_Timer
    LeftSlingShot1.TimerEnabled = 1
End Sub

Sub LeftSlingShot1_Timer
  Dim y2, y3, y4, ty, bl
    Select Case L1Step
    Case 0: y2=0: y3=0: y4=1: ty=-20
        Case 1: y2=0: y3=1: y4=0: ty=-10
        Case 2: y2=1: y3=0: y4=0: ty=0
        Case 3: y2=0: y3=0: y4=0: ty=0   :LeftSlingShot1.TimerEnabled = 0
    End Select
  For each bl in LeftSling12_bl : bl.visible = y2 : Next
  For each bl in LeftSling13_bl : bl.visible = y3 : Next
  For each bl in LeftSling14_bl : bl.visible = y4 : Next
  For each bl in LEMK1_bl : bl.transy = ty : Next
    L1Step = L1Step + 1
End Sub

'Sub LeftSlingShot1_Timer
'    Select Case L1Step
'        Case 1:LeftSLing14.Visible = 0:LeftSLing13.Visible = 1:Lemk1.RotX = 14
'        Case 2:LeftSLing13.Visible = 0:LeftSLing12.Visible = 1:Lemk1.RotX = 2
'        Case 3:LeftSLing12.Visible = 0:Lemk1.RotX = -10:LeftSlingShot1.TimerEnabled = 0
'    End Select
'
'    L1Step = L1Step + 1
'End Sub


'********************
'Bumpers
'********************
Dim bumperjackpot
Dim bumperjackpotcount

Sub Bumper1_Hit

  DOF 109,2


    startB2S(4)
  If VRRoom > 0 Then
    VRBGFLBump1
  End If
    b2snormal.Enabled=True
  lightCtrl.Pulse f145, 0: DOF 145,2
  SetGI Rgb(0,0,255),76

'    dim light
'    For each light in GiStrings
'        lightCtrl.PulseWithProfile light, Array(0,0,0,0,0,0,0,0), 1
'    Next

  RandomSoundBumperTop Bumper1
  dim bl
  For each bl in BR1_bl : bl.transz = -20 : Next

  Bumper1.Timerenabled = true
  BumperSfx


  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else

    If CurrentMode <> 2 Then gi038.State = 0
    If CurrentMode = 2 Or SJPReady(CurrentPlayer) > 0 Then gi038.State = 1
    gi038.TimerEnabled = True
    If CurrentMode = 5 Then
      AddScore(500000)
      bumperjackpotcount = bumperjackpotcount + 1
    Else
      AddScore(50110)
    End If
  End If

End Sub

sub Bumper1_timer
  dim bl
  Bumper1.Timerenabled = false
  For each bl in BR1_bl : bl.transz = 0 : Next
end sub

Sub Bumper3_Hit
  DOF 108,2

    startB2S(4)
  If VRRoom > 0 Then
    VRBGFLBump2
  End If
    b2snormal.Enabled=True
  lightCtrl.Pulse f144, 0: DOF 144,2
  SetGI Rgb(0,0,255),76

'    dim light
'    For each light in GiStrings
'        lightCtrl.PulseWithProfile light, Array(0,0,0,0,0,0,0,0), 1
'    Next
  RandomSoundBumpertop Bumper3

  dim bl
  For each bl in BR3_bl : bl.transz = -20 : Next

  Bumper3.Timerenabled = true
  BumperSfx

  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else

    If CurrentMode <> 2 Then gi038.State = 0
    If CurrentMode = 2 Or SJPReady(CurrentPlayer) > 0 Then gi038.State = 1
    gi038.TimerEnabled = True
    If CurrentMode = 5 Then
      AddScore(500000)
      bumperjackpotcount = bumperjackpotcount + 1
    Else
      AddScore(50110)
    End If
  End If
End Sub

sub Bumper3_timer
  dim bl
  Bumper3.Timerenabled = false
  For each bl in BR3_bl : bl.transz = 0 : Next
end sub


Dim BallsInPlay
Dim BallsToShoot
Dim Ballmustwait
Dim ScoopRestart
'********************
'Drain holes, vuks & saucers
'********************
Sub f148_timer
  f148.timerenabled = False
  f148.state = 0
  lightCtrl.Pulse f141, 0 : DOF 141,2
End Sub


Sub Drain_Hit
  Drain.DestroyBall
  RandomSoundDrain drain
' debug.print "Drain: WIZBiP=" & wizBiP
  if inlaneEBcount > 0 Then inlaneEBcount = 150



  If WizardWaitForNoBalls = 1 Then
    PlaySound "wiz1",0,RomSoundVolume
'   debug.print "Drain WizardWaitForNoBalls"
'   debug.print "WIZBiP=" & wizBiP
    wizBiP = wizBiP - 1
    If wizBiP < 1 Then WizardWaitForNoBalls = 0
    Exit Sub
  End If



  If SJPReady(CurrentPlayer) = 2 Or bTilted Then AddABallReady = 0 : Exit Sub

  SideSave = 0
  If OrbitSaved > 0 Then
    OrbitSaved = OrbitSaved - 1
    RandomKicker(2)
    f148.timerenabled = True
    lightCtrl.Pulse f141, 0 : DOF 141,2
'   BallsToShoot = BallsToShoot + 1
'   If Ballmustwait < frame + 40 Then Ballmustwait = frame + 40
    Underfire = 1 : firetime = 80
    StopSound "Sfx_BallSaved" : PlaySound "Sfx_BallSaved",0,RomSoundVolume
    DOF 125,2
    Exit Sub
  End If


' If OrbitSaved = 1 Then OrbitSaved = 0 : Exit Sub
  If CurrentMode <> 11 Then L49State(CurrentPlayer) = 0 : L50State(CurrentPlayer) = 0

  If Playtime < BallSaveTime(CurrentPlayer) Then
'   debug.print "drain  playtime check=" & PlayTime
    If CurrentMode <> 11 Then BallSaverTimer.Enabled = False
    BallsToShoot = BallsToShoot + 1
'   If Ballmustwait < frame + 40 Then Ballmustwait = frame + 40
    Underfire = 1 : firetime = 80
    StopSound "Sfx_BallSaved" : PlaySound "Sfx_BallSaved",0,RomSoundVolume
    DMDMessage1 "BALL     SAVE"   ,""       ,FontSponge16black  ,64,16  ,64,16  ,"shake"  ,2  ,"BG001"  ,"VIDdance"
    DOF 125,2
    UpdateLights
  End If

  If Playtime >= BallSaveTime(CurrentPlayer) Then
    DOF 124,2

    If AddABallReady = 2 Then AddABallReady = 0 : Exit Sub

    If CurrentMode <> 11 Then
      ClearDMD
      If VRroom = 1 then EndBOBAnim = 1 'setup VR Bob to start leaving

      ScoopRestart = 0



      If CurrentMode > 19 Then
        wizBiP = wizBiP - 1
        If wizBiP = 0 Then
          WizardTimer = 0
          ResetModeLights
          DMDMessage1 "OOOOOH"    ,"NOOOOO"     ,FontSponge12bb ,52,16  ,52,16  ,"blink"  ,1.0  ,"BG001"  ,"novideo"
          PlaySound "Sfx_SkillMiss",1,RomSoundVolume
          L18.blinkinterval = 110 : L20.blinkinterval = 110 : L20.blinkinterval = 110
          L24.blinkinterval = 110 : L25.blinkinterval = 110 : L26.blinkinterval = 110
          L27.blinkinterval = 110 : L28.blinkinterval = 110 : L29.blinkinterval = 110
          L30.blinkinterval = 110 : L31.blinkinterval = 110 : L32.blinkinterval = 110
          L37.blinkinterval = 110 : L38.blinkinterval = 110 : L39.blinkinterval = 110
          L40.blinkinterval = 110 : L41.blinkinterval = 110 : L42.blinkinterval = 110
          L43.blinkinterval = 110 : L44.blinkinterval = 110 : L45.blinkinterval = 110
          L46.blinkinterval = 110 : L47.blinkinterval = 110 : L48.blinkinterval = 110
          L21.blinkinterval = 110 : L22.blinkinterval = 110 : L23.blinkinterval = 110
          L33.blinkinterval = 125
          L34.blinkinterval = 125
          L35.blinkinterval = 125

        Else
          Exit Sub
        End If
      End If

      drain_FX

'Fixing need one for mode 21 too
      line4=0
      line4wait = 0

      lightCtrl.LightOff f146
      CurrentMode = 0
      StopTableMusic
      UpdateLights
      If modetimecounter > 0 Then
        DMDMessage1 "MODE"    ,"COMPLETED"      ,FontSponge12bb ,52,16  ,52,16  ,"blink"  ,1.0  ,"BG001"  ,"novideo"
        UpdateLights
        InsertSequence.Play SeqRandom,10,,1000
        modetimecounter = 0
      End If
      EndSpinnerHurryUp
      EobContinue.Enabled = True


    Else
      FlexDMD.Stage.GetLabel("eb1").visible = False
      FlexDMD.Stage.GetLabel("eb2").visible = False
      FlexDMD.Stage.GetLabel("eb3").visible = False
      FlexDMD.Stage.GetLabel("eb4").visible = False
      FlexDMD.Stage.GetLabel("eb5").visible = False
      FlexDMD.Stage.GetLabel("eb6").visible = False
      FlexDMD.Stage.GetImage("BG014").Visible=False
      DOF 131,0
      ModeJP_Done = True
      CurrentMode = 0
      Call ChangeGILights(1)
      ModeReady(CurrentPlayer) = 0
      SJPReady(CurrentPlayer) = 0
      If BankDown(CurrentPlayer) = 1 Then
        If MoveSpeed < 48 Then NewBankDown = 0
        If MoveSpeed > 47 Then currpos = 0 : Call Update3Bank(0,50)
      End If
      'sw77.IsDropped = False
      TriggerDown(CurrentPlayer) = False
      TriggerWall.Collidable = True : TriggerWallCollide(CurrentPlayer)=True
      TriggerWall.IsDropped = Not TriggerWall.Collidable
      sw77MovableHelper
      L33State(CurrentPlayer) = 0 : L34State(CurrentPlayer) = 0 : L35State(CurrentPlayer) = 0 : L36State(CurrentPlayer) = 0
      If NextMode(CurrentPlayer) > 0 Then
        ModeReady(CurrentPlayer) = 1
        DiverterON = 2
      End If
      Call ChangeGILights(1)
      UpdateLights
      If SJPgotten > 0 Then
        DMDMessage1 "MULTIBALL"   ,""       ,FontSponge16bb ,59,16  ,59,16  ,"solid"  ,0.7  ,"BG001"  ,"novideo"
        DMDMessage1 "COMPLETE"    ,""       ,FontSponge16bb ,59,16  ,59,16  ,"solid"  ,0.7  ,"BG001"  ,"novideo"
        SJPgotten = 0
        If VRroom = 1 then EndBOBAnim = 1 'setup VR Bob to start leaving

      Elseif SJPgotten = 0 then
        DMDMessage1 "RESTART"   ,"MULTIBALL"    ,FontSponge12bb ,59,16  ,59,16  ,"solid"  ,1.5  ,"BG001"  ,"novideo"
        DMDMessage1 "13 SECONDS TO"   ,"HIT SCOOP"  ,FontSponge12bb ,59,16  ,59,16  ,"solid"  ,1  ,"BG001"  ,"novideo"
        DMDMessage1 "12 SECONDS TO"   ,"HIT SCOOP"  ,FontSponge12bb ,59,16  ,59,16  ,"solid"  ,1  ,"BG001"  ,"novideo"
        DMDMessage1 "11 SECONDS TO"   ,"HIT SCOOP"  ,FontSponge12bb ,59,16  ,59,16  ,"solid"  ,1  ,"BG001"  ,"novideo"
        DMDMessage1 "10 SECONDS TO"   ,"HIT SCOOP"  ,FontSponge12bb ,59,16  ,59,16  ,"solid"  ,1  ,"BG001"  ,"novideo"
        DMDMessage1 "9 SECONDS TO"    ,"HIT SCOOP"  ,FontSponge12bb ,59,16  ,59,16  ,"solid"  ,1  ,"BG001"  ,"novideo"
        DMDMessage1 "8 SECONDS TO"    ,"HIT SCOOP"  ,FontSponge12bb ,59,16  ,59,16  ,"solid"  ,1  ,"BG001"  ,"novideo"
        DMDMessage1 "7 SECONDS TO"    ,"HIT SCOOP"  ,FontSponge12bb ,59,16  ,59,16  ,"solid"  ,1  ,"BG001"  ,"novideo"

        ScoopRestart = frame + 900
        SJPgotten = 2
        playsound "Hurry_up",1,RomSoundVolume
        Dampenmusic2 = 0.2
        QuoteTimer.Interval = 2712 : QuoteTimer.Enabled = True

      End If
      collect1 = 0 : collect2 = 0 : collect3 = 0 : collect4 = 0 : collect5 = 0 : collect6 = 0 : collect7 = 0 : collect8 = 0 : collect9 = 0
      IntroMusic
                  'VR Specific


    End If
  End If
End Sub


Sub EobContinue_Timer
      EobContinue.Enabled = False
      ModeLights
      BallSaverTimer.Enabled = False
      TurnLightsOff
      StopTableMusic
      SongTimer.Enabled = False
      EndingMusic
      Drain.TimerEnabled = True
      RoundEndTimer.Enabled = True
      'TiltGame = 1
      bTilted = True
      FlipperDeActivate LeftFlipper, LFPress
      SolLFlipper False
      SolULFlipper False
      FlipperDeActivate RightFlipper, RFPress
      SolRFlipper False
      ClearDMD
      If skipEOB = false And RestartGame < 1 Then
        DMDMessage1 "ROUND OVER"    ,""                   ,FontSponge16 ,52,16  ,55,16  ,"noslide"      ,0.5  ,"BG001"  ,"novideo"
        DMDMessage1 "BONUS SCORE"   ,""                   ,FontSponge12 ,56,16  ,60,16  ,"noslide2"     ,0.5  ,"BG001"  ,"novideo"
        DMDMessage1 FormatNumber(RoundScore/10,0,,,-1)  ,""           ,FontHighscore2 ,56,16  ,60,16  ,"noslide2"     ,0.5  ,"BG001"  ,"novideo"
        DMDMessage1 "MULTIPLIER X" & BonusRounds + 1  ,""           ,FontHighscore2 ,56,16  ,60,16  ,"noslide2"     ,0.5  ,"BG001"  ,"novideo"
        DMDMessage1 "TOTAL BONUS"   ,""                   ,FontSponge12 ,56,16  ,60,16  ,"noslide2"     ,0.5  ,"BG001"  ,"novideo"
        DMDMessage1 FormatNumber((RoundScore/10) * (BonusRounds + 1),0,,,-1),"" ,FontHighscore2 ,56,16  ,60,16  ,"noslide2blink"  ,1.0  ,"BG001"  ,"novideo"
        AddScore ((RoundScore/10) * (BonusRounds + 1))
      Elseif RestartGame < 1 Then
        DMDMessage1 "TILTED"    ,""                   ,FontSponge16bb ,52,16  ,55,16  ,"noslide"      ,0.5  ,"BG001"  ,"novideo"
        DMDMessage1 "NO BONUS"    ,""                   ,FontSponge16bb ,52,16  ,55,16  ,"noslide"      ,0.5  ,"BG001"  ,"novideo"
      End If
      skipEOB = False
      Warnings = 0
End Sub


Dim DoEball
Sub Drain_Timer
  Drain.TimerEnabled = False

      RoundScore = 0
      BonusRounds = 0

      Tilt = 0
      MechTilt = 0
      bTilted = False

      If RestartGame = 1 Then
        CurrentBall = 0
        ClearDMD
        StartNewPlayer
        Exit Sub
      End If

      If ExtraBall = 0 then
        If CurrentBall = 1 Then
          CheckHighScores : EnterPlayerNR = CurrentPlayer
        End If

        If CurrentPlayer = Players Then
          CurrentBall = CurrentBall - 1
          CurrentPlayer = 1

        Else
          CurrentPlayer = CurrentPlayer +1
        End If
        PlayTime = 0
        CurrentMode = 0
        ClearDMD
        StartNewPlayer
      Else
        PlayTime = 0
        ExtraBall = ExtraBall - 1
        ClearDMD
        DMDMessage1 "EXTRA BALL"    ,""       ,FontSponge16bb ,64,16  ,64,16  ,"shake"  ,1  ,"BG001"  ,"novideo"
        DMDMessage1 "SHOOT AGAIN"   ,""       ,FontSponge12bb ,54,16  ,54,16  ,"solid"  ,1  ,"BG001"  ,"novideo"
        QuoteTimer.Enabled = False
        PlaySound "Sfx_Sponge_EB",0,RomSoundVolume
        QuoteTimer.Interval =  2141 : QuoteTimer.Enabled = True

        CurrentMode = 0
        UpdateLights

        StartNewPlayer
      End If

End Sub


Dim AttractScore(4)
AttractScore(1) = 0
AttractScore(2) = 0
AttractScore(3) = 0
AttractScore(4) = 0
Dim ModeJP_Done
Dim ModeJP_Count
Sub StartNewPlayer
  ModeJP_Count = 0
  ModeJP_Done = True
  MBRampsforAddaBall = 1000
  Stopbob = 0
  If ScoreNumber = "" Then
    If CurrentBall > 0 Then
      If ModeReady(CurrentPlayer) = 0 Then
        DiverterON = 1
      Else
        DiverterON = 2
      End If

      ssREADY = 1
      StopSound "HSMusic"
      HSMusicPlaying.Enabled = False
      BallRelease.CreateBall
      DOF 111,2
      PlaySoundAt SoundFX("popper",DOFContactors), BallRelease
      BallRelease.Kick 80,6
      BallRelease.TimerEnabled = False
      SkillshotWait.enabled = True
      Call ChangeGILights(1)

      DMDMessage1 "PLAYER "& CurrentPlayer      ,""       ,FontSponge16BB ,65,16  ,65,16  ,"shake"  ,2  ,"BG001"  ,"novideo"
      SongNumber = int(rnd*7)
      PlayTableMusic(13)
      bTilted = False
    End If

    ResetPlayer
    SJPgotten = 0

    If MysteryReady(CurrentPlayer) = 1 Then
        l40.blinkinterval = 60 : l40.state = 2
        l41.blinkinterval = 60 : l41.state = 2
        l42.blinkinterval = 50 : l42.state = 2
    Else
        l40.blinkinterval = 125
        l41.blinkinterval = 125
        l42.blinkinterval = 125
    End If


    If CurrentBall = 0 Then
      lightCtrl.LightOff f146
      If BankDown(PreviousPlayer) = 1 Then currpos = 0 : Call Update3Bank(0,50)
      For x = 1 to 4
        AttractScore(x) = TotalScore(x)
      Next
      If RestartGame = 1 Then
        gameoverwaiting.interval = 999
      Else
        gameoverwaiting.interval = 9999
      End If
      gameoverwaiting.enabled = true

      RestartGame = 0
      ClearDMD
      DMDMessage1 "GAME OVER"       ,""       ,FontSponge16BB ,53,16  ,53,16  ,"fastblink"  ,3.5  ,"BGtilt" ,"novideo"
      EndingMusic
      Endofgamescoreblink = 1
      If Bubbles = 0 Then Bubbles = 1
      PlaySound "bubbles",0,RomSoundVolume
      bTilted = True

      If Not IsNull(Scorbit) Then
        Scorbit.StopSession TotalScore(1), TotalScore(2), TotalScore(3), TotalScore(4), Players
      End If

      DodNext = 2
      InsertSequence.Play SeqUpOn,50,1
      InsertSequence.Play SeqDownOn,50,1
      InsertSequence.Play SeqUpOn,50,1
      InsertSequence.Play SeqDownOn,50,1
      InsertSequence.Play SeqUpOn,50,1
      InsertSequence.Play SeqDownOn,50,1
      InsertSequence.Play SeqUpOn,50,1
      InsertSequence.Play SeqDownOn,50,1
      InsertSequence.Play SeqUpOn,50,1
      InsertSequence.Play SeqDownOn,50,1
    End If
    UpdateText
  End If
End Sub
Dim Endofgamescoreblink

Sub gameoverwaiting_Timer
  gameoverwaiting.enabled = false
  Endofgamescoreblink = 0
  DOF 121,1
  StartGame = 0
  EndingMusic
  resetnewgame
  DoOrDieMode = 0
  If Not IsNull(Scorbit) Then
    Scorbit.StopSession TotalScore(1), TotalScore(2), TotalScore(3), TotalScore(4), Players
  End If
End Sub




Sub sw36a_Hit

  DOF 130,2
  DodNext = 5
  PlaySoundAt "VUKEnter", sw36a : StopQuotes
  sw36a.DestroyBall
' TiltGame = 1
  bTilted = True
  ClearDMD
  BallSaverTimer.Enabled = False
  If ModesComplete(CurrentPlayer) = 4 And DoOrDieMode < 4 And modeEBgotten(CurrentPlayer) = 0 Then
      modeEBgotten(CurrentPlayer) = 1
      ExtraBall = ExtraBall + 1 : DoEball = 1
      DOF 129,2
      DMDMessage1 "EXTRA BALL"    ,""       ,FontSponge16bb ,53,16  ,53,16  ,"solid"  ,2.2  ,"BG001"  ,"novideo"
      Call JackpotSequence(750) : StopTableMusic : PlaySound "Sfx_ExtraBall",0,RomSoundVolume
  End If
' If ModesComplete(CurrentPlayer) = 9 Then
'   SJPReady(CurrentPlayer) = 0
''    TiltGame = 1
'   bTilted = True
'   FanfareTimer.Interval = 21026 : FanfareTimer.Enabled = True
'
'   Call DelayScore(20000000+HurryUpCollected(CurrentPlayer)*5000000 , 19.304)
'   DMDMessage1 " "   ,""       ,FontSponge16 ,69,16  ,69,16  ,"solid"  ,1  ,"BG001"  ,"novideo"
'   DMDMessage1 "ALL MODES "    ,"OPENED"       ,FontSponge12bb ,69,16  ,69,16  ,"solid"  ,7.085  ,"BG001"  ,"novideo"
'   DMDMessage1 "SUPER"       ,"JACKPOT"        ,FontSponge12bb ,69,16  ,69,16  ,"solid"  ,10.719 ,"BG001"  ,"novideo"
'   DMDMessage1 HurryUpCollected(CurrentPlayer)*5+20  , " MILLION"            ,FontSponge16bb ,54,16  ,54,16  ,"solid"  ,3.222  ,"BG001"  ,"novideo"
'
'
'   SJPLight.Interval = 1000 : SJPLight.Enabled = True
'   StopTableMusic
' End If


  If ModesComplete(CurrentPlayer) < 9 Then
    DMDMessage1 "NEW MODE"    ,""       ,FontSponge16bb ,50,16  ,55,16  ,"blink"  ,1.2  ,"BG001"  ,"novideo"


    Select Case ModeText(CurrentPlayer)
      Case "Do the Sponge Mode"
      DMDMessage1 "DO THE"    ,"SPONGE"     ,FontSponge16black  ,52,16  ,52,16  ,"blink"  ,4  ,"noimage"  ,"VIDsponge"
      Case "Plankton has Taken Over"
      DMDMessage1 "PLANKTON"  ,"TAKEOVER"       ,FontSponge16   ,52,16  ,52,16  ,"blink"  ,4  ,"noimage"  ,"VIDplankt"
      Case "Sandy's Rodeo Mode"
      DMDMessage1 "SANDY'S" ,"RODEO"        ,FontSponge16black  ,52,16  ,52,16  ,"blink"  ,4  ,"noimage"  ,"VIDsandys"
      Case "Mr Krabs Money Mode"
      DMDMessage1 "MR KRABS"  ,"MONEY"        ,FontSponge16black  ,52,16  ,52,16  ,"blink"  ,4  ,"noimage"  ,"VIDmoney"
      Case "Jellyfish Party Mode"
      DMDMessage1 "JELLYFISH" ,"PARTY"        ,FontSponge16   ,52,16  ,52,16  ,"blink"  ,4  ,"noimage"  ,"VIDjelly"
      Case "Boating School Mode"
      DMDMessage1 "BOATING" ,"SCHOOL"       ,FontSponge16black  ,52,16  ,52,16  ,"blink"  ,4  ,"noimage"  ,"VIDboating"
      Case "Squidward's Tiki Mode"
      DMDMessage1 "SQUIDWARD" ,"TIKI"         ,FontSponge16black  ,52,16  ,52,16  ,"blink"  ,4  ,"noimage"  ,"VIDtikiland"
      Case "Goofy Goobers Mode"
      DMDMessage1 "GOOFY"   ,"GOOBERS"        ,FontSponge16black  ,52,16  ,52,16  ,"blink"  ,4  ,"noimage"  ,"VIDgoofy"
      Case "Krusty Krab Mode"
      DMDMessage1 "KRUSTY"    ,"KRAB"       ,FontSponge16black  ,52,16  ,52,16  ,"blink"  ,4  ,"noimage"  ,"VIDcrabs"
    End select
  End If

  If ModesComplete(CurrentPlayer) = 3 And DoOrDieMode < 4 And modeEBgotten(CurrentPlayer) = 0 Then  DMDMessage1 "NEXT MODE"   ,"GIVE EXTRABALL"     ,FontSponge12bb ,52,16  ,52,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"


' If ModesComplete(CurrentPlayer) = 7 Then  DMDMessage1 "NEXT MODE"   ,"START WIZARD"       ,FontSponge12 ,52,16  ,52,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"


  AddScore(500000)
  showtimestopp = True
  DiverterON = 1
  sw36a.TimerInterval = MusicInterval(NextMode(CurrentPlayer))
  If ModesComplete(CurrentPlayer) = 4 Then sw36a.TimerInterval = sw36a.TimerInterval + 2000
' If ModesComplete(CurrentPlayer) = 9 Then sw36a.TimerInterval = sw36a.TimerInterval + 22026

  ModeReady(CurrentPlayer)=0
  ModeJP_Done = False
  ModeJP_Count = 0
  BonusRounds = BonusRounds + 1
  ModesComplete(CurrentPlayer) = ModesComplete(CurrentPlayer) + 1

'debug.print "modesComplete=" & ModesComplete(CurrentPlayer)

  If ModesComplete(CurrentPlayer) < 10 And ModeText(CurrentPlayer) = "Mr Krabs Money Mode" Then  modetimecounter = 100 : Else modetimecounter = 60


  If ModesComplete(CurrentPlayer) = 10 Then sw36a.Timerenabled = False : SJPReady(CurrentPlayer) = 0 : StopTableMusic : ResetModeLights : ModesComplete(CurrentPlayer) = 0 : WizardStart = 1 : Exit Sub '  wizard start

  Call ChangeGILights(0)

  sw36a.TimerEnabled = True

  CurrentMode = NextMode(CurrentPlayer)

  If CurrentMode = 1 Then modedone1(CurrentPlayer) = 1
  If CurrentMode = 2 Then modedone2(CurrentPlayer) = 1
  If CurrentMode = 3 Then modedone3(CurrentPlayer) = 1
  If CurrentMode = 4 Then modedone4(CurrentPlayer) = 1
  If CurrentMode = 5 Then modedone5(CurrentPlayer) = 1
  If CurrentMode = 6 Then modedone6(CurrentPlayer) = 1
  If CurrentMode = 7 Then modedone7(CurrentPlayer) = 1
  If CurrentMode = 8 Then modedone8(CurrentPlayer) = 1
  If CurrentMode = 9 Then modedone9(CurrentPlayer) = 1

  If ModesComplete(CurrentPlayer) <> 5 AND ModesComplete(CurrentPlayer) < 10 Then IntroMusic
  If ModesComplete(CurrentPlayer) = 5 Then DelayMusic = 1 : SongTimer.Interval = 1000 : SongTimer.Enabled = True
' If ModesComplete(CurrentPlayer) = 9 Then DelayMusic = 1 : SongTimer.Interval = 22026 : SongTimer.Enabled = True  ' wizardstart
    FlipperDeActivate LeftFlipper, LFPress : SolLFlipper False : SolULFlipper false: FlipperDeActivate RightFlipper, RFPress : SolRFlipper False
  NextMode(CurrentPlayer) = 0
  ModeLights



' DEBUG.PRINT "MODE12345=" & modedone1(CurrentPlayer) & modedone2(CurrentPlayer) & modedone3(CurrentPlayer) & modedone4(CurrentPlayer) & modedone5(CurrentPlayer) & " 6789=" & modedone6(CurrentPlayer) & modedone7(CurrentPlayer) & modedone8(CurrentPlayer) & modedone9(CurrentPlayer)
' DEBUG.PRINT "CURRENTMODDE=" & CURRENTMODE

  UpdateLights
  TurnLightsOff
  lightCtrl.AddLightSeq "lSeqRunnerModeSelect", lSeqModeSelect
End Sub

Dim modetimecounter
Sub ModeTimer_Timer
  StopTableMusic : EndingMusic : SongTimer.Interval = 3000 : SongTimer.Enabled = False: SongTimer.Enabled = True

  CurrentMode = 0
  ModeLights

  DMDMessage1 "MODE"    ,"COMPLETED"      ,FontSponge12bb ,52,16  ,52,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
  UpdateLights
  InsertSequence.Play SeqRandom,10,,1000
  modetimecounter = 0
End Sub


Sub sw36a_Timer
  sw36a.TimerEnabled = False
  lightCtrl.RemoveLightSeq "lSeqRunnerModeSelect", lSeqModeSelect
' TiltGame = 0
  bTilted = False
  Call RandomKicker(3)
' debug.print "sw36a_timer"

  PlayTime = 0
  BallSaverTimer.Enabled = True
  If CurrentMode <> 2 Then Call ChangeGILights(1)
  If CurrentMode = 2 Then ThunderTimer.Interval = int(rnd*10000) : ThunderTimer.Enabled = True
  UpdateLights
End Sub


Sub sw37a_Hit

  PlaySoundAt "VUKEnter", sw37a
  sw37a.DestroyBall
  sw37a.TimerEnabled = True

  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
    Exit Sub
  Elseif CurrentMode = 21 Then
    addscore (25000)
    Exit Sub
  End If

  If CurrentMode <> 11 Then
    StopTableMusic
    StopQuotes
'   TiltGame = 1
    bTilted = True
    FlipperDeActivate LeftFlipper, LFPress
    SolLFlipper False
    SolULFlipper False
    FlipperDeActivate RightFlipper, RFPress
    SolRFlipper False
  End If
  If CurrentMode <> 11 Then ChangeGILights(0)
  PlaySound "Sfx_SecretShot",0,RomSoundVolume

  SecretshotDone(CurrentPlayer) = SecretshotDone(CurrentPlayer) + 1
  If SecretshotDone(CurrentPlayer) > 9 Then SecretshotDone(CurrentPlayer) = 9
  DMDMessage1 "SECRET SHOT"   ,""       ,FontSponge12bb ,54,16  ,54,16  ,"solid"  ,1.0  ,"BG001"  ,"novideo"
  DMDMessage1 SecretshotDone(CurrentPlayer) * 2 + 8 & " MILLION"      ,""       ,FontSponge12bb ,63,16  ,63,16  ,"solid"  ,0.8  ,"BG001"  ,"novideo"
  Call DelayScore(SecretshotDone(CurrentPlayer)*2000000+8000000, 1)
  InsertSequence.Play SeqRandom,10,,1500
End Sub
Sub sw37a_Timer
  sw37a.TimerEnabled = False

  RandomKicker(3)

' debug.print "sw37a_timer"

  If CurrentMode <> 2 And SJPReady(CurrentPlayer) = 0 Then ChangeGILights(1)
  If CurrentMode <> 11 Then
    Call PlayTableMusic(CurrentMode)
'   TiltGame = 0
    bTilted = False
  End If
End Sub

Dim aBall


Sub sw78b_hit
  sw78b.timerenabled = True
    PlaySoundAt "VUKEnter", sw78 : StopQuotes
End Sub

Sub sw78b_timer
  sw78b.timerenabled = False
  sw78b.DestroyBall

    PlaySoundAt "Metal_Touch_1", sw78
    PlaySoundAt "VUKOut", sw78

  sw78_hit
End Sub

Dim rampsforaddaball
Dim MB_blinkbankLights
Sub sw78_Hit
  showtimestopp = True
  if CurrentMode = 11 Then Partyblinker.Interval = 255 : Partyblinker_Timer : ShakeSponge = 3 : SpongeTimer.Interval = 250 : SpongeTimer.Enabled = True


  ShakeMain=1
  ToyShake

' TiltGame = 1
  bTilted = True
  FlipperDeActivate LeftFlipper, LFPress : SolLFlipper False : SolULFlipper False: FlipperDeActivate RightFlipper, RFPress : SolRFlipper False
  ClearDMD
  BallSaverTimer.Enabled = False
  ChangeGILights(0)
  If CurrentMode <> 11 Then AddScore(500000)
  If CurrentMode = 11 Then


    addscore( 50000000+ HurryUpCollected(CurrentPlayer)*10000000 + MB_SJP(CurrentPlayer) * 20000000 )

    SJPgotten = 1
    DMDMessage1 "SUPER"   ,"JACKPOT"        ,FontSponge12bb ,69,17  ,69,17  ,"solid"  ,13 ,"BG001"  ,"novideo"
    DMDMessage1 formatscore (50 + HurryUpCollected(CurrentPlayer)* 10 + MB_SJP(CurrentPlayer) * 20 )    ,"MILLION"        ,FontSponge12bb ,69,16  ,69,16  ,"solid"  ,4  ,"BG001"  ,"novideo"
    DMDMessage1 "DO IT"   ,"AGAIN"        ,FontSponge12bb ,69,17  ,69,17  ,"solid"  ,2  ,"BG001"  ,"novideo"
    DMDMessage1 " "   ,""       ,FontSponge12bb ,69,16  ,69,16  ,"solid"  ,2  ,"party"  ,"novideo"
      DOF 132,2


    If MB_SJP(CurrentPlayer) < 9 Then MB_SJP(CurrentPlayer) = MB_SJP(CurrentPlayer) + 1
    Select Case MB_SJP(CurrentPlayer)
      Case 1 : collect1 = 1 : collect2 = 0 : collect3 = 0 : collect4 = 0 : collect5 = 0 : collect6 = 0 : collect7 = 0 : collect8 = 0 : collect9 = 0
      Case 2 : collect1 = 1 : collect2 = 0 : collect3 = 0 : collect4 = 0 : collect5 = 0 : collect6 = 0 : collect7 = 0 : collect8 = 0 : collect9 = 1
      Case 3 : collect1 = 1 : collect2 = 0 : collect3 = 0 : collect4 = 0 : collect5 = 0 : collect6 = 1 : collect7 = 0 : collect8 = 0 : collect9 = 1
      Case 4 : collect1 = 1 : collect2 = 0 : collect3 = 1 : collect4 = 0 : collect5 = 0 : collect6 = 1 : collect7 = 0 : collect8 = 0 : collect9 = 1
      Case 5 : collect1 = 1 : collect2 = 0 : collect3 = 1 : collect4 = 0 : collect5 = 0 : collect6 = 1 : collect7 = 0 : collect8 = 1 : collect9 = 1
      Case 6 : collect1 = 1 : collect2 = 1 : collect3 = 1 : collect4 = 0 : collect5 = 0 : collect6 = 1 : collect7 = 0 : collect8 = 1 : collect9 = 1
      Case 7 : collect1 = 1 : collect2 = 1 : collect3 = 1 : collect4 = 0 : collect5 = 1 : collect6 = 1 : collect7 = 0 : collect8 = 1 : collect9 = 1
      Case 8 : collect1 = 1 : collect2 = 1 : collect3 = 1 : collect4 = 0 : collect5 = 1 : collect6 = 1 : collect7 = 1 : collect8 = 1 : collect9 = 1
      Case 9 : collect1 = 1 : collect2 = 1 : collect3 = 1 : collect4 = 1 : collect5 = 1 : collect6 = 1 : collect7 = 1 : collect8 = 1 : collect9 = 1
    End Select
    If MBRampsforAddaBall = 999 And MBaddaball(CurrentPlayer) > 0 Then MBaddaball(CurrentPlayer) = MBaddaball(CurrentPlayer) - 1
    MBRampsforAddaBall = 0
    rampsforaddaball = 1
    StopTableMusic
    PlaySound "Sfx_SJP1",0,RomSoundVolume
    FanfareTimer.Interval = 22111 : FanfareTimer.Enabled = True
    SJPLight.Interval = 2622 : SJPLight.Enabled = True
    SJPReady(CurrentPlayer) = 2
    SongTimer.Enabled = False
    sw78.TimerInterval = 22141
'   TiltGame = 1
    bTilted = True
    FlipperDeActivate LeftFlipper, LFPress : SolLFlipper False : SolULFlipper False: FlipperDeActivate RightFlipper, RFPress : SolRFlipper False
  End If
  'sw77.IsDropped = False
  TriggerDown(CurrentPlayer) = False
  TriggerWall.Collidable = True : TriggerWallCollide(CurrentPlayer)=True
  TriggerWall.IsDropped = Not TriggerWall.Collidable
  sw77MovableHelper

  If CurrentMode < 10 Then BallsLocked(CurrentPlayer) = BallsLocked(CurrentPlayer)+1
  If BallsLocked(CurrentPlayer) = 1 Then StopTableMusic : PlaySound "Sfx_Lock1",0,RomSoundVolume : SongTimer.Enabled = False : sw78.TimerInterval = 2500 :    DMDMessage1 "BALL"    ,"LOCKED"       ,FontSponge16bb ,69,16  ,69,16  ,"solid"  ,2.5  ,"BG001"  ,"novideo"

  If BallsLocked(CurrentPlayer) = 2 Then
    CurrentMode = 10
    modetimecounter = 0
    MultiBallReady(CurrentPlayer) = 0
    BallsLocked(CurrentPlayer) = 0
    StopTableMusic
    PlaySound "Sfx_Lock2",0,RomSoundVolume
    DelayMusic = 1
    PartyMode = 1
'   TiltGame = 1
    bTilted = True
    SongTimer.Interval = 3380
    SongTimer.Enabled = True
    MBIntroLights.Enabled = True

    DMDMessage1 "BALL"    ,"LOCKED"         ,FontSponge16bb ,54,17  ,54,17  ,"solid"  ,3.38 ,"BG001"  ,"novideo"
    DMDMessage1 "ARE YOU"     ,"READY KIDS"   ,FontSponge12bb ,64,17  ,64,17  ,"solid"  ,3.00 ,"BG001"  ,"novideo"
    DMDMessage1 "I CAN'T"     ,"HEAR YOU"     ,FontSponge12bb ,69,17  ,69,17  ,"solid"  ,4.00 ,"BG001"  ,"novideo"
    DMDMessage1 "OHHHHH"      ,""         ,FontSponge12bb ,69,17  ,69,17  ,"solid"  ,2.00 ,"BG001"  ,"novideo"
    DMDMessage1 "MULTIBALL"     ,""         ,FontSponge12bb ,69,17  ,69,17  ,"solid"  ,2.00 ,"BG001"  ,"novideo"
    DMDMessage1 " "     ,""     ,FontSponge12bb ,50,18  ,58,18  ,"blink"  ,2.00 ,"party"  ,"novideo"

    rampsforaddaball = 1


    ShakeSponge = 7
    SpongeTimer.Interval = 250
    SpongeTimer.Enabled = True
    ShakePatty=1   : PlaySoundAt "AlienShake2", alien6_BM_Lit_Room
    ShakePatrick=1   : PlaySoundAt "AlienShake1", alien14_BM_Lit_Room
    ShakeSquidward=1 :PlaySoundAt "AlienShake1", alien5_BM_Lit_Room
    ToyShake



    sw78.TimerInterval = MusicInterval(11) + 3380
    FlipperDeActivate LeftFlipper, LFPress
    SolLFlipper False
    SolULFlipper False
    FlipperDeActivate RightFlipper, RFPress
    SolRFlipper False
    'VR Specific...
    If VRroom = 1 And Stopbob = 0 And BobAnimation = 0 then BobAnimation = 1 'Start VRBob Animation
    Stopbob = 0
  End If

  sw78.TimerEnabled = True

  If BankDown(CurrentPlayer) = 1 Then currpos = 0 : Call Update3Bank(0,50)

  BankMove = 0

  L33State(CurrentPlayer) = 0
  L34State(CurrentPlayer) = 0
  L35State(CurrentPlayer) = 0
  L36State(CurrentPlayer) = 0 : L03State(CurrentPlayer) = 0
  If BallsLocked(CurrentPlayer) = 0 Then L02State(CurrentPlayer) = 0
  If BallsLocked(CurrentPlayer) = 1 Then L02State(CurrentPlayer) = 1
  UpdateLights
  TurnLightsOff : AreLightsOff = 1
  showtimestopp = True

End Sub


'lightCtrl.AddTableLightSeq "GI", lSeqRainbow
'lightCtrl.RemoveTableLightSeq "GI", lSeqRainbow
Dim SJPgotten
Sub sw78_Timer
  showtimestopp = True

  sw78.TimerInterval = 1000
  sw78.TimerEnabled = False

  If SJPReady(CurrentPlayer) = 2 Then
    If ubound(getballs) > -1 Then
      sw78.TimerEnabled = True
      Exit Sub
    End If
  End If

    PlaySoundAt "Metal_Touch_1", sw78
' PlaySoundAt "VUKEnter", sw78

  StopSound "Sfx_SJP1"
  SJPLight.Enabled = False
  StopTableSequences()

  BallSaverTimer.Enabled = True

  If CurrentMode = 10 Then

    ShakeSponge = 2
    SpongeTimer.Interval = 250
    SpongeTimer.Enabled = True
    ShakePatty=1   : PlaySoundAt "AlienShake2", alien6_BM_Lit_Room
    ShakePatrick=1   : PlaySoundAt "AlienShake1", alien14_BM_Lit_Room
    ShakeSquidward=1 :PlaySoundAt "AlienShake1", alien5_BM_Lit_Room
    ToyShake
    CurrentMode = 11
    DOF 131,1
    MBRampsforAddaBall = 0
    Select Case MB_SJP(CurrentPlayer)
      case 0 : collect1 = 0 : collect2 = 0 : collect3 = 0 : collect4 = 0 : collect5 = 0 : collect6 = 0 : collect7 = 0 : collect8 = 0 : collect9 = 0
           MB_blinkbankLights = 1

      Case 1 : collect1 = 1 : collect2 = 0 : collect3 = 0 : collect4 = 0 : collect5 = 0 : collect6 = 0 : collect7 = 0 : collect8 = 0 : collect9 = 0
      Case 2 : collect1 = 1 : collect2 = 0 : collect3 = 0 : collect4 = 0 : collect5 = 0 : collect6 = 0 : collect7 = 0 : collect8 = 0 : collect9 = 1
      Case 3 : collect1 = 1 : collect2 = 0 : collect3 = 0 : collect4 = 0 : collect5 = 0 : collect6 = 1 : collect7 = 0 : collect8 = 0 : collect9 = 1
      Case 4 : collect1 = 1 : collect2 = 0 : collect3 = 1 : collect4 = 0 : collect5 = 0 : collect6 = 1 : collect7 = 0 : collect8 = 0 : collect9 = 1
      Case 5 : collect1 = 1 : collect2 = 0 : collect3 = 1 : collect4 = 0 : collect5 = 0 : collect6 = 1 : collect7 = 0 : collect8 = 1 : collect9 = 1
      Case 6 : collect1 = 1 : collect2 = 1 : collect3 = 1 : collect4 = 0 : collect5 = 0 : collect6 = 1 : collect7 = 0 : collect8 = 1 : collect9 = 1
      Case 7 : collect1 = 1 : collect2 = 1 : collect3 = 1 : collect4 = 0 : collect5 = 1 : collect6 = 1 : collect7 = 0 : collect8 = 1 : collect9 = 1
      Case 8 : collect1 = 1 : collect2 = 1 : collect3 = 1 : collect4 = 0 : collect5 = 1 : collect6 = 1 : collect7 = 1 : collect8 = 1 : collect9 = 1
      Case 9 : collect1 = 1 : collect2 = 1 : collect3 = 1 : collect4 = 1 : collect5 = 1 : collect6 = 1 : collect7 = 1 : collect8 = 1 : collect9 = 1
    End Select


    ModeReady(CurrentPlayer) = 1
    PlayTime = -10
    KickerSave(CurrentPlayer) = 2
    DiverterON = 1
    RandomKicker(4)
' debug.print "sw768_timer"

    sw78.TimerEnabled = True
    UpdateLights
    PartyMode = 0
    bTilted = False
    ModeLights
    Exit Sub
  End If
  ModeLights
' TiltGame = 0
  bTilted = False

  If Spongeshoootonmystery = 1 Then
' debug.print "sw78_timer spongeshootonmystery"
    If rampsforaddaball = 1 Then
      rampsforaddaball = 0
      DMDMessage1 "RAMPS FOR"     ,"ADDABALL"         ,FontSponge12bb ,59,17  ,59,17  ,"solid"  ,1.7  ,"BG001"  ,"novideo"
    End If
    Spongeshoootonmystery = 0
    SpongeBalls = SpongeBalls + 1
    sw37part2
  Else
    RandomKicker(3)
  End If

  If CurrentMode <> 11 Or SJPReady(CurrentPlayer) = 2 Then StopTableMusic : Call PlayTableMusic(CurrentMode)
  If SJPReady(CurrentPlayer) = 2 Then
    RandomKicker(4)
' debug.print "sw78_timer sjpready=2"
    SJPReady(CurrentPlayer) = 0
'   TiltGame = 0
    bTilted = False
  End If
  If CurrentMode <> 2 Then Call ChangeGILights(1)
  AreLightsOff = 0
  UpdateLights
End Sub

' Dim bBall

Dim Rampblinkcount
Sub Rampblinker_Timer
  If Rampblinker.enabled = False Then
    Rampblinker.enabled = True
    Rampblinkcount = 0
    SetGI Rgb(123,33,0),99
  Else
    Rampblinkcount = Rampblinkcount + 1
    Select Case Rampblinkcount
      case 1,3,5,7,9 : SetGI Rgb(255,77,0),99 : lightCtrl.LightOn L115
      case 2,4,6,8,10 : SetGI Rgb(123,33,0),99 : lightCtrl.LightOff L115
      case 12 : Rampblinker.enabled = False
          If gilvl = 1 Then lightCtrl.LightOn L115
    End Select

  End If
End Sub




Dim Spongeblinkercout
Sub Spongeblinker_Timer
  If Spongeblinker.enabled = False Then
    Spongeblinker.enabled = True
    DOF 135,2
    Spongeblinkercout = 0
    SetGI Rgb(123,123,10),99
  Else
    Spongeblinkercout = Spongeblinkercout + 1
    Select Case Spongeblinkercout
      case 1,3,5,7,9,11,13,15 : SetGI Rgb(244,222,10),99 : lightCtrl.LightOn L115
      case 2,4,6,8,10,12,14,16 : SetGI Rgb(123,123,10),99 : lightCtrl.LightOff L115
      case 17 : Spongeblinker.enabled = False
          If gilvl = 1 Then lightCtrl.LightOn L115
    End Select

  End If
End Sub

Dim partyblinkercout
Sub Partyblinker_Timer
  If Partyblinker.enabled = False Then
    Partyblinker.enabled = True
    DOF 136,2
    partyblinkercout = 0
    SetGI Rgb(123,123,10),99

  Else
    partyblinkercout = partyblinkercout + 1
    Select Case partyblinkercout
      case 1,4,7,10,13 : SetGI Rgb(0,255,11),99 : lightCtrl.LightOn L115
      case 2,5,8,11,14 : SetGI Rgb(0,11,255),99 : lightCtrl.LightOff L115
      case 3,6,9,12,15 : SetGI Rgb(0,77,77),99
      case 17 : Partyblinker.enabled = False
          If gilvl = 1 Then lightCtrl.LightOn L115
    End Select

  End If
End Sub

Sub sw37b_hit
  sw37b.timerinterval = 34
  sw37b.timerenabled = True
    PlaySoundAt "VUKEnter", sw37
End Sub

Sub sw37b_Timer
  sw37b.timerenabled = False
  sw37b.DestroyBall
  sw37_hit
End Sub

Sub Wall012_hit
  activeball.velx = activeball.velx / 3
  activeball.vely = activeball.vely / 3
  activeball.velz = - 1
End Sub

Sub RestartMultiball_Timer



End Sub


Dim Stopbob
Sub sw37_Hit
  showtimestopp = True
  DodNext = 7
  Spongeblinker_Timer
' debug.print "ScoopRestart" & ScoopRestart & "   frame" & frame
  If ScoopRestart > frame Then
    BallsLocked(CurrentPlayer) = 1
    ScoopRestart = 0
    Spongeshoootonmystery = 1
    DOF 131,1
    sw78_hit
    Stopbob = 1
    rampsforaddaball = 1
    Exit Sub
  End If


  If currentmode = 21 Then
    If collect6 > 0 Then collect6 = collect6 - 1 : WizardJPUpdate
  Else
    If collect6 = 1 Then collect6 = 0 : CheckCollected
  End If

  If MysteryReady(CurrentPlayer) = 1 and currentmode < 10 Then
    MysteryReady(CurrentPlayer) = 0
    l40.blinkinterval = 125 : l40.State = L40State(CurrentPlayer)
    l41.blinkinterval = 125 : l41.State = L41State(CurrentPlayer)
    l42.blinkinterval = 125 : l42.State = L42State(CurrentPlayer)
    RampsMysteryCounter(CurrentPlayer) = 0
    RampsForMystery(CurrentPlayer) = RampsForMystery(CurrentPlayer) + 1
    If RampsForMystery(CurrentPlayer) > 5 Then RampsForMystery(CurrentPlayer) = 5
    DMDMessage1 "MYSTERY"   ,""     ,FontSponge16bb ,-30,16 ,55,16  ,"noslide"  ,2.0  ,"BG001"  ,"novideo"

    MysteryTimer_Timer

    StopTableMusic
    PlaySound "Mystery", 1, RomSoundVolume
    SongTimer.interval = 8000
    SongTimer.enabled = False
    SongTimer.enabled = True
    Exit Sub
  End If

  SpongeBalls = SpongeBalls + 1
  sw37part2
End Sub



Dim Spongeshoootonmystery
Dim MysteryCounter
Dim MysterySelect
Sub MysteryTimer_Timer
  If MysteryTimer.enabled = False Then
    MysteryTimer.interval = 2000
    MysteryTimer.enabled = True
    MysteryCounter = 0
  Else
    MysteryTimer.interval = 230
    MysteryCounter = MysteryCounter + 1
    If MysteryCounter < 10 Then
      Select Case int(rnd(1)*10)

        Case 0 : DMDMessage1 "500 K"      ,""         ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"

        Case 1 : DMDMessage1 "1 MILLION"  ,""         ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"

        Case 2 : DMDMessage1 "JACKPOT"    ,""         ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"

        Case 3 : DMDMessage1 "BALL LOCK"  ,""         ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"

        Case 4 : If BallSaveTime(CurrentPlayer) < 18 Then
              DMDMessage1 "BALLSAVE +1" ,""     ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"
             Else
              DMDMessage1 "500 K" ,""         ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"
             End If

        Case 5 : DMDMessage1 "NO WORRIES" ,""         ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"

        Case 6 : DMDMessage1 "ADVANCE JP" ,""         ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"

        Case 7 : If InlaneEB(CurrentPlayer) < 5 Then
              DMDMessage1 "EB LETTER" ,""         ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"
             Else
              DMDMessage1 "BIG POINTS"  ,""         ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"
             End If


        Case 8 : If ModesComplete(CurrentPlayer) < 7 Then
              DMDMessage1 "MODECOMPLETE"  ,""       ,FontSponge12bb ,56,16  ,56,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"
             Else
              DMDMessage1 "JACKPOT"   ,""       ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"
             End If

        Case 9 : If RandomEB(CurrentPlayer) = 0 And DoOrDieMode < 4 Then
              DMDMessage1 "EXTRA BALL"  ,""       ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"
             Else
              DMDMessage1 "1 MILLION" ,""         ,FontSponge12bb ,55,16  ,55,16  ,"noslide2" ,0.2  ,"BG001"  ,"novideo"
             End If
      End Select
    Else
      Select Case MysteryCounter
        case 11 :
          MysterySelect = Int(Rnd(1)*100)
          If MysterySelect < 15 Then mysteryselect = 88
          If BallSaveTime(CurrentPlayer) > 17 And MysterySelect >58 And MysterySelect < 66 Then MysterySelect = Int(Rnd(1)*53)
          If InlaneEB(CurrentPlayer) > 4 And MysterySelect >92 And MysterySelect < 98 Then MysterySelect = Int(Rnd(1)*53)
          If RandomEB(CurrentPlayer) = 1 And MysterySelect >72 And MysterySelect < 79 Then MysterySelect = Int(Rnd(1)*53)
          If ModesComplete(CurrentPlayer) > 7 And MysterySelect >84 And MysterySelect < 93 Then MysterySelect = Int(Rnd(1)*53)
          If CurrentMode = 11 And MysterySelect > 52 And MysterySelect < 59 Then MysterySelect = Int(Rnd(1)*53)
          ClearDMD
          Select Case MysterySelect

            Case 0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23
              Addscore (500000)
              DMDMessage1 "500.000"     ,""           ,FontSponge12bb ,55,16  ,55,16  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"

            Case 24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40
              Addscore (1000000)
              DMDMessage1 "1.000.000"   ,""             ,FontSponge12bb ,55,16  ,55,16  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"

            Case 41,42,43,44,45,46,47,48,49,50,51,52
              Addscore ((HurryUpCollected(CurrentPlayer) + 2)*1000000)
              DMDMessage1 Formatscore ((HurryUpCollected(CurrentPlayer) + 2)*1000000) ,"" ,FontSponge12bb ,55,16  ,55,16  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"

            Case 53,54,55,56,57,58      ' lock ball
              Addscore (100000)
              MysteryTimer.enabled = False
              Spongeshoootonmystery = 1
              sw78_Hit
              Exit Sub

            Case 59,60,61,62,63,64,65 ' ballsave +1
              Addscore (100000)
              If BallSaveTime(CurrentPlayer) < 18 Then BallSaveTime(CurrentPlayer) = BallSaveTime(CurrentPlayer) + 1
              DMDMessage1 "BALLSAVE"      ,"UPGRADE"        ,FontSponge12bb ,55,17  ,55,17  ,"noslide2blink"  ,1  ,"BG001"  ,"novideo"
              DMDMessage1 "IS NOW"    ,BallSaveTime(CurrentPlayer) & " SEC"   ,FontSponge12bb ,69,17  ,69,17  ,"blink"  ,2  ,"BG001"  ,"novideo"
'             Displayballsave = "BALL SAVE IS NOW " & BallSaveTime(CurrentPlayer) & " SEC"

            Case 66,67,68,69,70,71,72,98,99 ' ballsavestart to the max +3
              Addscore (100000)
              DMDMessage1 "BALLSAVE"      ,"ACTIVATED"      ,FontSponge12bb ,55,17  ,55,17  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
              PlayTime = -3

            Case 73,74,75,76,77,78    'Extraball Letter
              Addscore (100000)
              InlaneEB(CurrentPlayer) = InlaneEB(CurrentPlayer) + 1
              ShakeSponge = 1 : SpongeTimer.Interval = 250 : SpongeTimer.Enabled = True
              ClearDMD
'             DMDMessage1 " "   ,""       ,FontSponge16bb ,55,16  ,55,16  ,"extraball"  ,7  ,"BG014"  ,"novideo"
              inlaneEBcount = 1
              If InlaneEB(CurrentPlayer) = 5 Then
                DoEball = 1
                ExtraBall = ExtraBall + 1
                DOF 129,2
                Call JackpotSequence(750)
                PlaySound "Sfx_ExtraBall",0,RomSoundVolume
              End If
              L15State(CurrentPlayer) = 2 : L17State(CurrentPlayer) = 2
              inlanetimer.enabled =  True
              If PlayTime > BallSaveTime(CurrentPlayer)-3 Then
                If ExtraBall > 0 Then l01.State=1 Else l01.State=0
              End If

            Case 79,80,81,82,83,84    'advance JP level
              Addscore (100000)
              FanfareSfx
              Call JackpotSequence(1000)
              HurryUpCollected(CurrentPlayer) = HurryUpCollected(CurrentPlayer) + 1
              DMDMessage1 "JP LEVEL"    ,"IS NOW " & HurryUpCollected(CurrentPlayer) ,FontSponge12bb  ,59,18  ,59,18  ,"solid"  ,1.5  ,"BG009"  ,"novideo"


            Case 85,86,87,88,89,90,91,92      'complete a random mode
              Addscore (100000)
              dim tmp,x
              tmp = 0
'             FindMissionToComplete
              Select Case int(rnd(1)*9) + 1
                Case 1 : If L05State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 7 Then tmp = 1
                Case 2 : If L06State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 2 Then tmp = 2
                Case 3 : If L07State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 4 Then tmp = 3
                Case 4 : If L08State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 3 Then tmp = 4
                Case 5 : If L09State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 8 Then tmp = 5
                Case 6 : If L10State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 6 Then tmp = 6
                Case 7 : If L11State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 9 Then tmp = 7
                Case 8 : If L12State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 1 Then tmp = 8
                Case 9 : If L13State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 5 Then tmp = 9
              End Select
' If modedone1(CurrentPlayer) = 1 Then L12State(CurrentPlayer) = 1 X sponge
' If modedone2(CurrentPlayer) = 1 Then L06State(CurrentPlayer) = 1 x plankton
' If modedone3(CurrentPlayer) = 1 Then L08State(CurrentPlayer) = 1 x sandy
' If modedone4(CurrentPlayer) = 1 Then L07State(CurrentPlayer) = 1 x mr krab
' If modedone5(CurrentPlayer) = 1 Then L13State(CurrentPlayer) = 1 x yellyfish
' If modedone6(CurrentPlayer) = 1 Then L10State(CurrentPlayer) = 1 x boating
' If modedone7(CurrentPlayer) = 1 Then L05State(CurrentPlayer) = 1 x tiki
' If modedone8(CurrentPlayer) = 1 Then L09State(CurrentPlayer) = 1 x goofyt
' If modedone9(CurrentPlayer) = 1 Then L11State(CurrentPlayer) = 1 x formula

              If tmp = 0 Then '2nd random try
                Select Case int(rnd(1)*9) + 1
                Case 1 : If L05State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 7 Then tmp = 1
                Case 2 : If L06State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 2 Then tmp = 2
                Case 3 : If L07State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 4 Then tmp = 3
                Case 4 : If L08State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 3 Then tmp = 4
                Case 5 : If L09State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 8 Then tmp = 5
                Case 6 : If L10State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 6 Then tmp = 6
                Case 7 : If L11State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 9 Then tmp = 7
                Case 8 : If L12State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 1 Then tmp = 8
                Case 9 : If L13State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 5 Then tmp = 9
                End Select
              End If
              ' force it
              If tmp = 0 And L05State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 7 Then tmp = 1
              If tmp = 0 And L06State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 2 Then tmp = 2
              If tmp = 0 And L07State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 4 Then tmp = 3
              If tmp = 0 And L08State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 3 Then tmp = 4
              If tmp = 0 And L09State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 8 Then tmp = 5
              If tmp = 0 And L10State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 6 Then tmp = 6
              If tmp = 0 And L11State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 9 Then tmp = 7
              If tmp = 0 And L12State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 1 Then tmp = 8
              If tmp = 0 And L13State(CurrentPlayer) = 0 And NextMode(CurrentPlayer) <> 5 Then tmp = 9

              Select Case tmp
                Case 1 :
                  DMDMessage1 "TIKILAND"      ,"COMPLETE"           ,FontSponge12bb ,55,17  ,55,17  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
                  L18State(CurrentPlayer) = 1 : L19State(CurrentPlayer) = 1 : L20State(CurrentPlayer) = 1
                  modedone7(CurrentPlayer) = 1

                Case 2 :
                  DMDMessage1 "PLANKTON"      ,"COMPLETE"           ,FontSponge12bb ,55,17  ,55,17  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
                  L27State(CurrentPlayer) = 1 : L28State(CurrentPlayer) = 1 : L29State(CurrentPlayer) = 1
                  modedone2(CurrentPlayer) = 1

                Case 3 :
                  DMDMessage1 "MR KRABS"      ,"COMPLETE"           ,FontSponge12bb ,55,17  ,55,17  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
                  L37State(CurrentPlayer) = 1 : L38State(CurrentPlayer) = 1 : L39State(CurrentPlayer) = 1
                  modedone4(CurrentPlayer) = 1
                Case 4 :
                  DMDMessage1 "RODEO"     ,"COMPLETE"           ,FontSponge12bb ,55,17  ,55,17  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
                  L43State(CurrentPlayer) = 1 : L44State(CurrentPlayer) = 1 : L45State(CurrentPlayer) = 1
                  modedone3(CurrentPlayer) = 1

                Case 5 :
                  DMDMessage1 "GOOFY"     ,"COMPLETE"           ,FontSponge12bb ,55,17  ,55,17  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
                  L21State(CurrentPlayer) = 1 : L22State(CurrentPlayer) = 1 : L23State(CurrentPlayer) = 1
                  modedone8(CurrentPlayer) = 1

                Case 6 :
                  DMDMessage1 "SCHOOL"      ,"COMPLETE"           ,FontSponge12bb ,55,17  ,55,17  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
                  L24State(CurrentPlayer) = 1 : L25State(CurrentPlayer) = 1 : L26State(CurrentPlayer) = 1
                  modedone6(CurrentPlayer) = 1

                Case 7 :
                  DMDMessage1 "FORMULA"     ,"COMPLETE"           ,FontSponge12bb ,55,17  ,55,17  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
                  L30State(CurrentPlayer) = 1 : L31State(CurrentPlayer) = 1 : L32State(CurrentPlayer) = 1
                  modedone9(CurrentPlayer) = 1

                Case 8 :
                  DMDMessage1 "SPONGE"      ,"COMPLETE"           ,FontSponge12bb ,55,17  ,55,17  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
                  L40State(CurrentPlayer) = 1 : L41State(CurrentPlayer) = 1 : L42State(CurrentPlayer) = 1
                  modedone1(CurrentPlayer) = 1

                Case 9 :
                  DMDMessage1 "JELLYFISH"     ,"COMPLETE"           ,FontSponge12bb ,55,17  ,55,17  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
                  L46State(CurrentPlayer) = 1 : L47State(CurrentPlayer) = 1 : L48State(CurrentPlayer) = 1
                  modedone5(CurrentPlayer) = 1

              End Select




              ModesComplete(CurrentPlayer) = ModesComplete(CurrentPlayer) + 1

              If ModesComplete(CurrentPlayer) = 5 And DoOrDieMode < 4 And modeEBgotten(CurrentPlayer) = 0 Then
                modeEBgotten(CurrentPlayer) = 1
                DOF 129,2
                ExtraBall = ExtraBall + 1 : DoEball = 1 : DMDMessage1 "EXTRA BALL"    ,""       ,FontSponge12bb ,53,16  ,53,16  ,"solid"  ,2.2  ,"BG001"  ,"novideo" : Call JackpotSequence(750) : StopTableMusic : PlaySound "Sfx_ExtraBall",0,RomSoundVolume
                sw37.TimerInterval = sw37.TimerInterval + 2000
              End If

              ModeLights
              UpdateLights

            Case 93,94,95,96,97       'very smal chance of a EB
              Addscore (100000)
              ExtraBall = ExtraBall + 1
              DOF 129,2
              Call JackpotSequence(750)
              PlaySound "Sfx_ExtraBall",0,RomSoundVolume
              If PlayTime > BallSaveTime(CurrentPlayer)-3 Then
                If ExtraBall > 0 Then l01.State=1 Else l01.State=0
              End If
              DMDMessage1 "EXTRA"     ,"BALL"     ,FontSponge12bb ,55,16  ,55,16  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
              RandomEB(CurrentPlayer) = 1
              DoEball = 1
          End Select

        Case 30,31 : MysteryTimer.enabled = False : SpongeBalls = SpongeBalls + 1 : sw37part2

      End Select
    End If


  End If
End Sub


' x = "noslide2"
' If time = 0.8 Then x = "noslide3"


Sub sw37part2
  If not bTilted Then
    If CurrentMode = 20 Then
      AddScore (100000)
        Call PlayQuote(3)
        Select Case int(rnd(1)*4)
          case 0 : DMDMessage1 "HERE IT IS"   ,""       ,FontSponge16bb ,53,17  ,57,17  ,"solid"  ,1.7  ,"noimage"  ,"VIDhouse"
          case 1 : DMDMessage1 "CAREFUL"    ,""         ,FontSponge16bb ,53,17  ,57,17  ,"solid"  ,1.7  ,"noimage"  ,"VIDhouse"
          case 2 : DMDMessage1 "INCOMING"   ,""         ,FontSponge16bb ,53,17  ,57,17  ,"solid"  ,1.7  ,"noimage"  ,"VIDhouse"
          case 3 : DMDMessage1 "WATCHOUT"   ,""         ,FontSponge16bb ,53,17  ,57,17  ,"solid"  ,1.7  ,"noimage"  ,"VIDhouse"
        End Select
      WizardJPUpdate
    Elseif CurrentMode = 21 Then
      addscore (25000)
        Call PlayQuote(3)
        Select Case int(rnd(1)*4)
          case 0 : DMDMessage1 "HERE IT IS"   ,""       ,FontSponge16bb ,53,17  ,57,17  ,"solid"  ,1.7  ,"noimage"  ,"VIDhouse"
          case 1 : DMDMessage1 "CAREFUL"    ,""         ,FontSponge16bb ,53,17  ,57,17  ,"solid"  ,1.7  ,"noimage"  ,"VIDhouse"
          case 2 : DMDMessage1 "INCOMING"   ,""         ,FontSponge16bb ,53,17  ,57,17  ,"solid"  ,1.7  ,"noimage"  ,"VIDhouse"
          case 3 : DMDMessage1 "WATCHOUT"   ,""         ,FontSponge16bb ,53,17  ,57,17  ,"solid"  ,1.7  ,"noimage"  ,"VIDhouse"
        End Select
    Else


      If CurrentMode = 1 Then
        ModeJP_Done = True
        DOF 132,2

        ModeJP_Count = ModeJP_Count + 1
        If ModeJP_Count = 3 Then

          AddScore(10000000+HurryUpCollected(CurrentPlayer)*1000000)
          DMDMessage1 "MEGA"      ,"JACKPOT"          ,FontSponge12bb ,69,17  ,69,17  ,"noslide"  ,2.00 ,"BG001"  ,"novideo"
          DMDMessage1 formatscore ( 10 + HurryUpCollected(CurrentPlayer)*1) & "MILLION"   ,""         ,FontSponge12bb ,54,14  ,54,14  ,"noslide2blink"  ,2.00 ,"BG001"  ,"novideo"
          HurryUpCollected(CurrentPlayer) = HurryUpCollected(CurrentPlayer) + 1
          DMDMessage1 "JP LEVEL"    ,"IS NOW " & HurryUpCollected(CurrentPlayer),FontSponge12bb ,55,18  ,55,18  ,"solid"  ,1  ,"BG009"  ,"novideo"
          JackpotSequence(1000) : JackpotSfx
          sw37.TimerInterval = sw37.TimerInterval + 2000
        Else
          AddScore(2500000+HurryUpCollected(CurrentPlayer)*500000)
          DMDMessage1 "JACKPOT"     ,""         ,FontSponge16bb ,69,16  ,69,16  ,"noslide"  ,2.00 ,"BG001"  ,"novideo"
          DMDMessage1 formatscore ( 2.5 + HurryUpCollected(CurrentPlayer)*0.5) & "MILLION"    ,""         ,FontSponge12bb ,54,14  ,54,14  ,"noslide2blink"  ,2.00 ,"BG001"  ,"novideo"
          JackpotSequence(1000) : JackpotSfx
          sw37.TimerInterval = sw37.TimerInterval + 2000
        End If

      Else
        AddScore(350000)
        Call PlayQuote(3)
        Select Case int(rnd(1)*4)
          case 0 : DMDMessage1 "HERE IT IS"   ,""       ,FontSponge16bb ,53,17  ,57,17  ,"solid"  ,1.7  ,"noimage"  ,"VIDhouse"
          case 1 : DMDMessage1 "CAREFUL"    ,""         ,FontSponge16bb ,53,17  ,57,17  ,"solid"  ,1.7  ,"noimage"  ,"VIDhouse"
          case 2 : DMDMessage1 "INCOMING"   ,""         ,FontSponge16bb ,53,17  ,57,17  ,"solid"  ,1.7  ,"noimage"  ,"VIDhouse"
          case 3 : DMDMessage1 "WATCHOUT"   ,""         ,FontSponge16bb ,53,17  ,57,17  ,"solid"  ,1.7  ,"noimage"  ,"VIDhouse"
        End Select
      End If

      If CurrentMode < 20 Then
        If L42State(CurrentPlayer) = 2 And ModeReady(CurrentPlayer)= 0 And MultiBallReady(CurrentPlayer) = 0 And ModeJP_Done Then
          L42State(CurrentPlayer) = 1
          DiverterON = 2
          ModeReady(CurrentPlayer)=1 : NextMode(CurrentPlayer)=1
          DMDMessage1 "MODE "   ,"READY"        ,FontSponge16bb ,57,16  ,57,16  ,"solid"  , 2 ,"noimage"  ,"VIDmememe" : ModeText(CurrentPlayer) = "Do the Sponge Mode" : PlaySound "Sfx_ModeReady",0,RomSoundVolume
          UpdateLights
        End If

        If ModesComplete(CurrentPlayer) = 9 And SJPReady(CurrentPlayer) = 0 And ModeJP_Done Then
          SJPReady(CurrentPlayer) = 1 : PlaySound "Sfx_Cheer1",0,RomSoundVolume
          DiverterON = 2
          ModeReady(CurrentPlayer)=1
        End If

        If L40State(CurrentPlayer) = 1 AND L41State(CurrentPlayer) = 0 Then L41State(CurrentPlayer) = 1 : L42State(CurrentPlayer) = 2 : UpdateLights
        If L40State(CurrentPlayer) = 0 Then L40State(CurrentPlayer) = 1 : UpdateLights
      End If
    End If
  End If
  sw37.TimerEnabled = False
  sw37.TimerEnabled = True
  ShakeSponge = 2
  SpongeTimer.Interval = 250
  SpongeTimer.Enabled = True
End Sub

Sub sw37_Timer
  If btilted Then sw37.TimerEnabled = False : SpongeBalls = 0 : Exit Sub

  DodNext = 9
  sw37.TimerEnabled = False
  sw37.timerinterval = 1500
  SpongeBalls = SpongeBalls - 1
  RandomKicker(int(rnd(1)*3))
' debug.print "sw37_timer"

  If SpongeBalls > 0 Then sw37.TimerEnabled = True

End Sub

Sub RandomKicker(KickerSelect)
  If KickerSelect < 3 And Not Btilted Then
    BallSaverTimer.enabled = False
    BallSaverTimer.enabled = True
    If Playtime > BallSaveTime(CurrentPlayer) - 3 And WizardWaitForNoBalls = 0 Then Playtime = BallSaveTime(CurrentPlayer) - 3
  End If
  Select Case KickerSelect
    Case 0 : sw37b.CreateBall : sw37b.Kickz int(rnd(1)*2 + 199), int(rnd(1)*2 + 23),13,60 : PlaySoundAt SoundFX("popper",DOFContactors), sw37:      DOF 118,2
    Case 1 : sw37b.CreateBall : sw37b.Kickz int(rnd(1)*2 + 190), int(rnd(1)*2 + 14),13,60 : PlaySoundAt SoundFX("popper",DOFContactors), sw37:      DOF 118,2
    Case 2 : DrainShooter.CreateBall : DrainShooter.Kick int(rnd(1)*30 - 15), int(rnd(1)*2 + 19) : PlaySoundAt SoundFX("popper",DOFContactors), DrainShooter  :   DOF 111,2
    Case 3 : sw36.CreateBall : sw36.Kick int(rnd(1)*2 - 1), int(rnd(1)*2 + 1) : PlaySoundAt SoundFX("KickandWire",DOFContactors), sw36 : blockmysterycount = 1 :      DOF 118,2
    Case 4 : BallRelease.CreateBall : BallRelease.Kick 80,6 : BallRelease.TimerEnabled = True : PlaySoundAt SoundFX("popper",DOFContactors), BallRelease :      DOF 111,2
  End Select
End Sub


Sub BallRelease_Timer
  BallRelease.TimerEnabled = False
  Auto_Plunger
End Sub


'********************
'Rollovers & Ramp Switches
'********************


Sub sw16_Hit
  lightCtrl.Pulse f141, 3 : DOF 141,2
  If btilted then exit Sub
  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else
    AddScore (100000)
  End If


' If PlayTime < BallSaveTime(CurrentPlayer) Then BallSaverTimer.Enabled = False
  If PlayTime > BallSaveTime(CurrentPlayer) Then
    If KickerSave(CurrentPlayer) > 0 Then
      If SJPReady(CurrentPlayer) <> 2 Then
        KickerSave(CurrentPlayer) = KickerSave(CurrentPlayer) - 1
        OrbitSaved = OrbitSaved + 1
        f148.state = 2
        ClearDMD
        DMDMessage1 "BALL     SAVE"   ,""       ,FontSponge16black  ,64,16  ,64,16  ,"shake"  ,2  ,"BG001"  ,"VIDdance"
      End If
    End If
  Else
    If SJPReady(CurrentPlayer) <> 2 Then
'       KickerSave(CurrentPlayer) = KickerSave(CurrentPlayer) - 1
      OrbitSaved = OrbitSaved + 1
      f148.state = 2
      ClearDMD
      DMDMessage1 "BALL     SAVE"   ,""       ,FontSponge16black  ,64,16  ,64,16  ,"shake"  ,2  ,"BG001"  ,"VIDdance"
    End If
  End If
  UpdateLights
End Sub


Sub sw26_Hit
  activeball.angmomz = 0
  activeball.vely = 0.7*activeball.vely
  If LRampHit = 0 Then L15State(CurrentPlayer) = 1
  LRampHit = 0

  If btilted then exit Sub

  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else
    AddScore (50000)
  End If


  If L15State(CurrentPlayer) = 1 And L17State(CurrentPlayer) = 1 And InlaneEB(CurrentPlayer) = 5 Then
    bumperjackpotcount = bumperjackpotcount + 1
    addscore 500000
    L15State(CurrentPlayer) = 2 : L17State(CurrentPlayer) = 2
    inlanetimer.enabled =  True
  End If

  If L15State(CurrentPlayer) = 1 And L17State(CurrentPlayer) = 1 And InlaneEB(CurrentPlayer) < 5 Then
    InlaneEB(CurrentPlayer) = InlaneEB(CurrentPlayer) + 1
    ShakeSponge = 1 : SpongeTimer.Interval = 250 : SpongeTimer.Enabled = True

'   DMDMessage1 " "   ,""       ,FontSponge16bb ,55,16  ,63,16  ,"extraball"  ,7  ,"BG014"  ,"novideo"
    ClearDMD
    inlaneEBcount = 1

    If InlaneEB(CurrentPlayer) = 5 Then
      ExtraBall = ExtraBall + 1
      DOF 129,2
      Call JackpotSequence(750)
      PlaySound "Sfx_ExtraBall",0,RomSoundVolume
    End If
    L15State(CurrentPlayer) = 2 : L17State(CurrentPlayer) = 2
    inlanetimer.enabled =  True
    If PlayTime > BallSaveTime(CurrentPlayer)-3 Then
      If ExtraBall > 0 Then l01.State=1 Else l01.State=0
    End If

  End If

  UpdateLights
End Sub
Sub inlanetimer_Timer
  inlanetimer.enabled =  False
  L15State(CurrentPlayer) = 0 : L17State(CurrentPlayer) = 0
  UpdateLights
End Sub


Sub sw17_Hit
  activeball.angmomz = 0
  activeball.vely = 0.7*activeball.vely
  If RRampHit = 0 Then L17State(CurrentPlayer) = 1
  RRampHit = 0

  If btilted then exit Sub

  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else
    AddScore (50000)
  End If

  If L15State(CurrentPlayer) = 1 And L17State(CurrentPlayer) = 1 And InlaneEB(CurrentPlayer) = 5 Then
    bumperjackpotcount = bumperjackpotcount + 1
    addscore 500000
    L15State(CurrentPlayer) = 2 : L17State(CurrentPlayer) = 2
    inlanetimer.enabled =  True
  End If

  If L15State(CurrentPlayer) = 1 And L17State(CurrentPlayer) = 1 And InlaneEB(CurrentPlayer) < 5 Then
    InlaneEB(CurrentPlayer) = InlaneEB(CurrentPlayer) + 1
    ShakeSponge = 1 : SpongeTimer.Interval = 250 : SpongeTimer.Enabled = True
    ClearDMD
'   DMDMessage1 " "   ,""       ,FontSponge16bb ,55,16  ,63,16  ,"extraball"  ,7  ,"BG014"  ,"novideo"
    inlaneEBcount = 1

    If InlaneEB(CurrentPlayer) = 5 Then
      ExtraBall = ExtraBall + 1
      DOF 129,2
      Call JackpotSequence(750)
      PlaySound "Sfx_ExtraBall",0,RomSoundVolume
    End If
    L15State(CurrentPlayer) = 2 : L17State(CurrentPlayer) = 2
    inlanetimer.enabled =  True
    If PlayTime > BallSaveTime(CurrentPlayer)-3 Then
      If ExtraBall > 0 Then l01.State=1 Else l01.State=0
    End If
  End If
  UpdateLights
End Sub

Sub sw27_Hit
  lightCtrl.Pulse f142, 3 : DOF 142,2

  If btilted then exit Sub

  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else
    AddScore (100000)
  End If

' If PlayTime < BallSaveTime(CurrentPlayer) Then BallSaverTimer.Enabled = False
  If PlayTime > BallSaveTime(CurrentPlayer) Then
    If KickerSave(CurrentPlayer) > 0 Then
      If SJPReady(CurrentPlayer) <> 2 Then
        KickerSave(CurrentPlayer) = KickerSave(CurrentPlayer) - 1
        OrbitSaved = OrbitSaved + 1
        f148.state = 2
        ClearDMD
        DMDMessage1 "BALL     SAVE"   ,""       ,FontSponge16black  ,64,16  ,64,16  ,"shake"  ,2  ,"BG001"  ,"VIDdance"
      End If
    End If
  Else
    If SJPReady(CurrentPlayer) <> 2 Then
      OrbitSaved = OrbitSaved + 1
      f148.state = 2
      ClearDMD
      DMDMessage1 "BALL     SAVE"   ,""       ,FontSponge16black  ,64,16  ,64,16  ,"shake"  ,2  ,"BG001"  ,"VIDdance"
    End If
  End If
  UpdateLights
End Sub


Dim Displayballsave : Displayballsave = ""
Sub sw38_Hit
  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else

    addscore 25000 : If L49State(CurrentPlayer)=0 Then addscore 150000
    L49State(CurrentPlayer)=1
    If L49State(CurrentPlayer) = 1 AND L50State(CurrentPlayer) = 1 Then

      If BallSaveTime(CurrentPlayer) < 18 Then
        BallSaveTime(CurrentPlayer) = BallSaveTime(CurrentPlayer) + 1

'       DMDMessage1 "BALLSAVE"    ,BallSaveTime(CurrentPlayer) & " SEC"   ,FontSponge12bb ,69,16  ,69,16  ,"blink"  ,1.2  ,"BG001"  ,"novideo"
'       Displayballsave = "BALL SAVE IS NOW " & BallSaveTime(CurrentPlayer) & " SEC"
        DMDMessage1 "BALLSAVE"            ,"UPGRADE"                ,FontSponge12bb    ,55,17    ,55,17    ,"noslide2blink"    ,1    ,"BG001"    ,"novideo"
                DMDMessage1 "IS NOW"        ,BallSaveTime(CurrentPlayer) & " SEC"        ,FontSponge12bb    ,69,17    ,69,17    ,"blink"    ,2    ,"BG001"    ,"novideo"

      Else
        addscore 300000
      End If
      StopSound "Sfx_Timer" : PlaySound "Sfx_Timer",0,RomSoundVolume

      L49State(CurrentPlayer) = 0 : L50State(CurrentPlayer) = 0
    End If
    UpdateLights
  End If
End Sub

Sub sw48_Hit
  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else

    addscore 25000 : If L50State(CurrentPlayer)=0 Then addscore 100000
    L50State(CurrentPlayer)=1
    If L49State(CurrentPlayer) = 1 AND L50State(CurrentPlayer) = 1 Then

      If BallSaveTime(CurrentPlayer) < 18 Then
        BallSaveTime(CurrentPlayer) = BallSaveTime(CurrentPlayer) + 1
'       DMDMessage1 "BALLSAVE"    ,BallSaveTime(CurrentPlayer) & " SEC"   ,FontSponge12bb ,69,16  ,69,16  ,"blink"  ,1.2  ,"BG001"  ,"novideo"
'       Displayballsave = "BALL SAVE IS NOW " & BallSaveTime(CurrentPlayer) & " SEC"
        DMDMessage1 "BALLSAVE"            ,"UPGRADE"                ,FontSponge12bb    ,55,17    ,55,17    ,"noslide2blink"    ,1    ,"BG001"    ,"novideo"
                DMDMessage1 "IS NOW"        ,BallSaveTime(CurrentPlayer) & " SEC"        ,FontSponge12bb    ,69,17    ,69,17    ,"blink"    ,2    ,"BG001"    ,"novideo"
      Else
        addscore 300000
      End If

      StopSound "Sfx_Timer" : PlaySound "Sfx_Timer",0,RomSoundVolume

      L49State(CurrentPlayer) = 0 : L50State(CurrentPlayer) = 0
    End If
    UpdateLights
  End If
End Sub

Sub sw71_Hit
  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else
    AddScore (10000)
  End If

  If Bubbles = 0 Then Bubbles = 1
  PlaySound "bubbles",0,RomSoundVolume
  LGate.Open=True : sw71.TimerEnabled=True
End Sub

Sub sw71_Timer
  sw71.TimerEnabled=False
  LGate.Open=False
End Sub

Sub sw72_Hit
  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else
    AddScore (10000)
  End If

  StopLOrbitSfx
  If ActiveBall.VelY < 0 Then
    LOrbitSfx
    sw72.TimerEnabled = False
    sw72.TimerEnabled = True
  End If
End Sub

Sub sw72_Timer
  sw72.TimerEnabled = False
End Sub

Sub sw73_Hit
  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else
    AddScore (10000)
  End If

  If Bubbles = 0 Then Bubbles = 1
  PlaySound "bubbles",0,RomSoundVolume
End Sub


Sub sw74_Hit
  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else
    AddScore (10000)
  End If

  StopSound "Sfx_ROrbit"
  If ActiveBall.VelY < 0 Then
    PlaySound "Sfx_ROrbit",0,RomSoundVolume
    sw74.TimerEnabled = False
    sw74.TimerEnabled = True
  End If
End Sub

Sub sw74_Timer
  sw74.TimerEnabled = False
End Sub

Sub sw61_Hit
  RRampHit = RRampHit + 1
  If RRampHit > 1 Then RRampHit = 0
  If ActiveBall.VelY < 0 Then
    RRampSfx
  Else
    StopRRampSfx
  End If
End Sub
Sub sw61_Unhit:End Sub

Sub sw62_Hit
  If ModeReady(CurrentPlayer) = 0 OR CurrentMode = 11 Then RGate.Open=True : sw62.TimerEnabled=True
End Sub

Sub sw62_Timer
  sw62.TimerEnabled=False
  If ModeReady(CurrentPlayer)=1 AND CurrentMode <> 11 then
    DiverterON = 2 : RGate.Open=False
  End If
  If ModeReady(CurrentPlayer)=0 OR CurrentMode = 11 then RGate.Open=False
End Sub
Sub DiverterStuck_Hit
  If ModeReady(CurrentPlayer)=1 AND CurrentMode <> 11 And MultiBallReady(CurrentPlayer) = 0 Then
    DiverterON = 1
    sw62.TimerEnabled=True
  End If
End Sub

Sub sw63_Hit
  If ActiveBall.VelY < 0 Then
    If CurrentMode <> 2 Then LRampSfx
    If CurrentMode = 2 Then Call PlayQuote(5)
  Else
    StopLRampSfx
  End If
End Sub

Sub sw63_Unhit:End Sub

Sub LeftRamp_Hit
  LRampHit = LRampHit + 1
  StopLRampSfx


  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else

    If CurrentMode <> 2 Then AddScore(500000)

    If CurrentMode = 2 Then
      ModeJP_Done = True
      ModeJP_Count = ModeJP_Count + 1
      DOF 132,2

      If ModeJP_Count = 3 Then
        AddScore(10000000+HurryUpCollected(CurrentPlayer)*1000000)
        DMDMessage1 "MEGA"      ,"JACKPOT"          ,FontSponge12bb ,69,17  ,69,17  ,"noslide"  ,2.00 ,"BG001"  ,"novideo"
        DMDMessage1 formatscore ( 10 + HurryUpCollected(CurrentPlayer)*1) & "MILLION"   ,""         ,FontSponge12bb ,54,14  ,54,14  ,"noslide2blink"  ,2.00 ,"BG001"  ,"novideo"
        HurryUpCollected(CurrentPlayer) = HurryUpCollected(CurrentPlayer) + 1
        DMDMessage1 "JP LEVEL"    ,"IS NOW " & HurryUpCollected(CurrentPlayer),FontSponge12bb ,55,18  ,55,18  ,"solid"  ,1  ,"BG009"  ,"novideo"
        JackpotSequence(1000) : JackpotSfx
        sw37.TimerInterval = sw37.TimerInterval + 2000
      Else
        AddScore(3000000+HurryUpCollected(CurrentPlayer)*1000000)
        JackpotSequence(1000) : JackpotSfx
        DMDMessage1 "JACKPOT"   ,""       ,FontSponge16bb ,69,16  ,69,16  ,"noslide"  ,0.75 ,"BG001"  ,"novideo"
        DMDMessage1 formatscore ( 3 + HurryUpCollected(CurrentPlayer)*1 )   ,"MILLION"      ,FontSponge16bb ,69,14  ,69,14  ,"noslide2blink"  ,2  ,"BG001"  ,"novideo"
      End If
    End If

    If L29State(CurrentPlayer) = 2 And ModeReady(CurrentPlayer)= 0 And MultiBallReady(CurrentPlayer) = 0 And ModeJP_Done Then
      L29State(CurrentPlayer) = 1
      DiverterON = 2
      ModeReady(CurrentPlayer)=1 : NextMode(CurrentPlayer)=2
      DMDMessage1 "MODE "   ,"READY"        ,FontSponge16bb ,69,16  ,69,16  ,"solid"  , 2 ,"BG001"  ,"novideo"
      ModeText(CurrentPlayer) = "Plankton has Taken Over" : PlaySound "Sfx_ModeReady",0,RomSoundVolume
    End If

        If ModesComplete(CurrentPlayer) = 9 And SJPReady(CurrentPlayer) = 0 And ModeJP_Done Then
      SJPReady(CurrentPlayer) = 1 : PlaySound "Sfx_Cheer1",0,RomSoundVolume
      DiverterON = 2
      ModeReady(CurrentPlayer)=1
    End If

    If L27State(CurrentPlayer) = 1 AND L28State(CurrentPlayer) = 0 Then L28State(CurrentPlayer) = 1 : L29State(CurrentPlayer) = 2
    If L27State(CurrentPlayer) = 0 Then L27State(CurrentPlayer) = 1

    UpdateLights
  End If

End Sub

Sub RightRamp_Hit
  RRampHit = RRampHit + 1
  StopRRampSfx

  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    addscore (25000)
  Else

    If rampsforaddaball = 1 Then
      rampsforaddaball = 0
      DMDMessage1 "RAMPS FOR"     ,"ADDABALL"         ,FontSponge12bb ,59,17  ,59,17  ,"solid"  ,1.7  ,"BG001"  ,"novideo"
    End If


    If RRampHit = 2 Then
      If CurrentMode <> 3 Then AddScore(500000)
      If CurrentMode = 3 Then
        ModeJP_Done = True
      DOF 132,2

        ModeJP_Count = ModeJP_Count + 1
        If ModeJP_Count = 3 Then
          AddScore(10000000+HurryUpCollected(CurrentPlayer)*1000000)
          DMDMessage1 "MEGA"      ,"JACKPOT"          ,FontSponge12bb ,69,17  ,69,17  ,"noslide"  ,2.00 ,"BG001"  ,"novideo"
          DMDMessage1 formatscore ( 10 + HurryUpCollected(CurrentPlayer)*1) & "MILLION"   ,""         ,FontSponge12bb ,54,14  ,54,14  ,"noslide2blink"  ,2.00 ,"BG001"  ,"novideo"
          HurryUpCollected(CurrentPlayer) = HurryUpCollected(CurrentPlayer) + 1
          DMDMessage1 "JP LEVEL"    ,"IS NOW " & HurryUpCollected(CurrentPlayer),FontSponge12bb ,55,18  ,55,18  ,"solid"  ,1  ,"BG009"  ,"novideo"
          JackpotSequence(1000) : JackpotSfx
          sw37.TimerInterval = sw37.TimerInterval + 2000
        Else
          AddScore(5000000+HurryUpCollected(CurrentPlayer)*1000000)
          Call JackpotSequence(1000) : JackpotSfx
          DMDMessage1 "JACKPOT"   ,""     ,FontSponge16bb ,69,16  ,69,16  ,"noslide"  ,0.75 ,"BG001"  ,"novideo"
          DMDMessage1 formatscore ( 5 + HurryUpCollected(CurrentPlayer)*1 )   ,"MILLION"      ,FontSponge16bb ,69,14  ,69,14  ,"noslide2blink"  ,2  ,"BG001"  ,"novideo"
        End If
      End If

      If L45State(CurrentPlayer) = 2 And ModeReady(CurrentPlayer)= 0 And MultiBallReady(CurrentPlayer) = 0 And ModeJP_Done Then
        L45State(CurrentPlayer) = 1
        DiverterON = 2
        ModeReady(CurrentPlayer)=1 : NextMode(CurrentPlayer)=3
        DMDMessage1 "MODE "   ,"READY"        ,FontSponge16bb ,69,16  ,69,16  ,"solid"  , 2 ,"BG001"  ,"novideo"
        ModeText(CurrentPlayer) = "Sandy's Rodeo Mode" : PlaySound "Sfx_ModeReady",0,RomSoundVolume
      End If

      If ModesComplete(CurrentPlayer) = 9 And SJPReady(CurrentPlayer) = 0 And ModeJP_Done Then
        SJPReady(CurrentPlayer) = 1 : PlaySound "Sfx_Cheer1",0,RomSoundVolume
DiverterON = 2
        ModeReady(CurrentPlayer)=1
      End If

      If L43State(CurrentPlayer) = 1 AND L44State(CurrentPlayer) = 0 Then L44State(CurrentPlayer) = 1 : L45State(CurrentPlayer) = 2
      If L43State(CurrentPlayer) = 0 Then L43State(CurrentPlayer) = 1
    End If

    UpdateLights
  End If
End Sub



Sub sw70_Hit()
    lightCtrl.Pulse L110, 0
End Sub

'Dim SpinnerCount1(4)
'Dim SpinnerCount2(4)
'Dim SpinnerLevel1(4)
'Dim SpinnerLevel2(4)
'Dim HurryUpCollected(4)


Const sc_SpinnerSpin      =   10000
Const sc_SpinnerLevelUp     = 1500000
Const sc_SpinnerLevelBonus    =  500000
Const sc_SpinneMaxLevel     =      20
Const sc_SpinnsNeededFirstLvl   =      20
Const sc_SpinnsAdded      =       5

Dim SpinnersIgnore
Dim SpinnerHurryUp
Dim SpinnerTimer


Sub EndSpinnerHurryUp
  If SpinnerHurryUp > 0 Then
    ClearDMD
    SpinnerCount1(CurrentPlayer) = 0
    SpinnerTimer = 0
    L30.state = 0 : L31.state = 0 : L32.state = 0
    l30.blinkinterval = 125 : l31.blinkinterval = 125 : l32.blinkinterval = 125
    SpinnerHurryUp = 0
  End If
End Sub


Sub Spinner1_Animate
  Dim spinangle:spinangle = Spinner1.currentangle
  Dim BL : For Each BL in Spinner1_BL : BL.RotX = spinangle: Next
End Sub


Sub Spinner1_spin
  dim tmp
  DOF 126,2
  SoundSpinner(spinner1)

  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
    Exit Sub
  Elseif CurrentMode = 21 Then
    addscore (25000)
    Exit Sub
  End If

  AddScore(( sc_SpinnerSpin * ( SpinnerLevel2(CurrentPlayer) +1) ) + 10)
  If DMDDisplay(20,0) = "" then FlexMsg = 0


  If CurrentMode > 0 Then  ' And CurrentMode < 10 Then

    SpinnerCount1(CurrentPlayer) = SpinnerCount1(CurrentPlayer) + 1
    If SpinnerCount1(CurrentPlayer) = sc_SpinnsNeededFirstLvl - 1 + ( spinnerlevel1 (CurrentPlayer) * sc_SpinnsAdded)  Then
      SpinnerHurryUp = 1

      PlaySound "Sfx_Airhorn",0,RomSoundVolume
      SpinnerTimer = 27
      UpdateLights
      ShakePatty=1 : Toyshake : PlaySoundAt "AlienShake2", alien6_BM_Lit_Room : noshakebigone = 1
      ClearDMD
      DMDMessage1 "HURRY UP"      , ""    ,FontSponge12BB ,54,19  ,54,19  ,"fastblink"  ,1.5  ,"BG010"  ,"novideo"
      DMDMessage1 "SECRET FORMULA"  , "24 SECONDS"    ,FontSponge12 ,54,17  ,54,17  ,"fastblink"  ,2  ,"BG010"  ,"novideo"
    Else
      If CurrentMSGdone = 0 And DMD_Slide < 1 And SpinnersIgnore < frame And SpinnerTimer = 0 Then
        SpinnersIgnore = frame + 456
        tmp = sc_SpinnsNeededFirstLvl + (SpinnerLevel1(CurrentPlayer) * sc_SpinnsAdded ) - SpinnerCount1(CurrentPlayer)
        If tmp >= 0 Then DMDMessage1 "HURRYUP"  ,"NEED " & tmp & " SPINS"     ,FontSponge12bb ,54,18  ,54,18  ,"spinner2" , 0.8 ,"BG010"  ,"novideo"

      ElseIf DMDDisplay(20,7) = "spinner2" And DMD_Slide < 1 And SpinnerTimer = 0 Then ' reset displaytim
        FlexMsgTime = Frame + 0.7 * 63 + 5
      End If

    End If

  Else

    SpinnerCount2(CurrentPlayer) = SpinnerCount2(CurrentPlayer) + 1
    If SpinnerCount2(CurrentPlayer) > sc_SpinnsNeededFirstLvl - 1 + (SpinnerLevel2(CurrentPlayer)*sc_SpinnsAdded) Then

      SpinnerCount2(CurrentPlayer) = 0
      SpinnerLevel2(CurrentPlayer) = SpinnerLevel2(CurrentPlayer) + 1
      If SpinnerLevel2(CurrentPlayer) > sc_SpinneMaxLevel Then SpinnerLevel2(CurrentPlayer) = sc_SpinneMaxLevel
      AddScore( sc_SpinnerLevelUp + SpinnerLevel2(CurrentPlayer) * sc_SpinnerLevelBonus )
      If DMDDisplay(20,7) = "spinner" Then ClearDMD
      DMDMessage1 "SPINLEVEL " & SpinnerLevel2(CurrentPlayer)   , formatscore( sc_SpinnerLevelUp + ( SpinnerLevel2(CurrentPlayer) * sc_SpinnerLevelBonus))      ,FontSponge12BB ,54,19  ,54,19  ,"blink"  ,2.0  ,"BG006"  ,"novideo"
      PlaySound "Sfx_Airhorn",0,RomSoundVolume
    Else

      If CurrentMSGdone = 0 And DMD_Slide < 1 And SpinnersIgnore < frame Then
        SpinnersIgnore = frame + 456
        tmp = sc_SpinnsNeededFirstLvl + (SpinnerLevel2(CurrentPlayer) * sc_SpinnsAdded ) - SpinnerCount2(CurrentPlayer)
        DMDMessage1 "SPINNER-UP"  ,"NEED " & tmp & " SPINS"     ,FontSponge12bb ,54,18  ,54,18  ,"spinner"  , 0.8 ,"BG005"  ,"novideo"
      ElseIf DMDDisplay(20,7) = "spinner" And DMD_Slide < 1 Then ' reset displaytim
        FlexMsgTime = Frame + 0.7 * 63 + 5
      End If

    End If
  End If

End Sub


Sub Gate001_Hit
  If ActiveBall.VelY < 0 Then


    If CurrentMode = 20 Then
      AddScore (100000)
      WizardJPUpdate
      Exit Sub
    Elseif CurrentMode = 21 Then
      addscore (25000)
      If collect5 > 0 Then collect5 = collect5 - 1 : WizardJPUpdate
      Exit Sub
    Else
      If RampComboSelect <> 5  Then CheckCombo
      RampCombo = frame + 500
      RampComboSelect = 5
    End If

    If collect5 = 1 Then collect5 = 0 : CheckCollected

    If CurrentMode <> 4 Then AddScore(100000) : Call PlayQuote(2)

    If CurrentMode = 4 Then
      ModeJP_Done = True
      ModeJP_Count = ModeJP_Count + 1
      DOF 132,2

      If ModeJP_Count = 3 Then
        AddScore(10000000+HurryUpCollected(CurrentPlayer)*1000000)
        DMDMessage1 "MEGA"      ,"JACKPOT"          ,FontSponge12bb ,69,17  ,69,17  ,"noslide"  ,2.00 ,"BG001"  ,"novideo"
        DMDMessage1 formatscore ( 10 + HurryUpCollected(CurrentPlayer)*1) & "MILLION"   ,""         ,FontSponge12bb ,54,14  ,54,14  ,"noslide2blink"  ,2.00 ,"BG001"  ,"novideo"
        HurryUpCollected(CurrentPlayer) = HurryUpCollected(CurrentPlayer) + 1
        DMDMessage1 "JP LEVEL"    ,"IS NOW " & HurryUpCollected(CurrentPlayer),FontSponge12bb ,55,18  ,55,18  ,"solid"  ,1  ,"BG009"  ,"novideo"
        JackpotSequence(1000) : JackpotSfx
        sw37.TimerInterval = sw37.TimerInterval + 2000
      Else
        AddScore(7000000+HurryUpCollected(CurrentPlayer)*1000000)
        Call JackpotSequence(1000) : JackpotSfx
        DMDMessage1 "JACKPOT"   ,""     ,FontSponge16bb ,69,16  ,69,16  ,"noslide"  ,0.75 ,"BG001"  ,"novideo"
        DMDMessage1 formatscore ( 7 + HurryUpCollected(CurrentPlayer)*1 )   ,"MILLION"      ,FontSponge16bb ,69,14  ,69,14  ,"noslide2blink"  ,0.75 ,"BG001"  ,"novideo"
      End If
    End If

    If L39State(CurrentPlayer) = 2 And ModeReady(CurrentPlayer)= 0 And MultiBallReady(CurrentPlayer) = 0 And ModeJP_Done Then
      L39State(CurrentPlayer) = 1
DiverterON = 2
      ModeReady(CurrentPlayer)=1 : NextMode(CurrentPlayer)=4
      DMDMessage1 "MODE "   ,"READY"        ,FontSponge16bb ,69,16  ,69,16  ,"solid"  , 2 ,"BG001"  ,"novideo"
      ModeText(CurrentPlayer) = "Mr Krabs Money Mode" : PlaySound "Sfx_ModeReady",0,RomSoundVolume
    End If

    If ModesComplete(CurrentPlayer) = 9 And SJPReady(CurrentPlayer) = 0 And ModeJP_Done Then
      SJPReady(CurrentPlayer) = 1 : PlaySound "Sfx_Cheer1",0,RomSoundVolume
DiverterON = 2
      ModeReady(CurrentPlayer)=1
    End If

    If L37State(CurrentPlayer) = 1 AND L38State(CurrentPlayer) = 0 Then L38State(CurrentPlayer) = 1 : L39State(CurrentPlayer) = 2
    If L37State(CurrentPlayer) = 0 Then L37State(CurrentPlayer) = 1

    UpdateLights

  End If
End Sub

Sub LGateTrigger_Hit


  If sw74.TimerEnabled = True AND ActiveBall.VelX < 0 Then 'ROrbitHit = 3 Then


    If CurrentMode = 20 Then
      AddScore (100000)
      WizardJPUpdate
      Exit Sub
    Elseif CurrentMode = 21 Then
      addscore (25000)
      If collect8 > 0 Then collect8 = collect8 - 1 : WizardJPUpdate

      Exit Sub
    Else

      If RampComboSelect <> 4  Then CheckCombo
      RampCombo = frame + 500
      RampComboSelect = 4
    End If

    If collect8 = 1 Then collect8 = 0 : CheckCollected
    AddScore(100000)
    If CurrentMode = 5 Then
      ModeJP_Done = True
      ModeJP_Count = ModeJP_Count + 1
      DOF 132,2

      If ModeJP_Count = 3 Then
        AddScore(10000000+HurryUpCollected(CurrentPlayer)*1000000)
        DMDMessage1 "MEGA"      ,"JACKPOT"          ,FontSponge12bb ,69,17  ,69,17  ,"noslide"  ,2.00 ,"BG001"  ,"novideo"
        DMDMessage1 formatscore ( 10 + HurryUpCollected(CurrentPlayer)*1) & "MILLION"   ,""         ,FontSponge12bb ,54,14  ,54,14  ,"noslide2blink"  ,2.00 ,"BG001"  ,"novideo"
        HurryUpCollected(CurrentPlayer) = HurryUpCollected(CurrentPlayer) + 1
        DMDMessage1 "JP LEVEL"    ,"IS NOW " & HurryUpCollected(CurrentPlayer),FontSponge12bb ,55,18  ,55,18  ,"solid"  ,1  ,"BG009"  ,"novideo"
        JackpotSequence(1000) : JackpotSfx
        sw37.TimerInterval = sw37.TimerInterval + 2000
      Else
        AddScore(500000+HurryUpCollected(CurrentPlayer)*500000)
        Call JackpotSequence(1000) : JackpotSfx
        DMDMessage1 "JACKPOT"   ,""     ,FontSponge16bb ,69,16  ,69,16  ,"noslide"  ,0.75 ,"BG001"  ,"novideo"
      End If
    End If

    If L48State(CurrentPlayer) = 2 And ModeReady(CurrentPlayer)= 0 And MultiBallReady(CurrentPlayer) = 0 And ModeJP_Done Then
      L48State(CurrentPlayer) = 1
DiverterON = 2
      ModeReady(CurrentPlayer)=1 : NextMode(CurrentPlayer)=5
      DMDMessage1 "MODE "   ,"READY"        ,FontSponge16bb ,69,16  ,69,16  ,"solid"  , 2 ,"BG001"  ,"novideo"
      ModeText(CurrentPlayer) = "Jellyfish Party Mode" : PlaySound "Sfx_ModeReady",0,RomSoundVolume
    End If

  If ModesComplete(CurrentPlayer) = 9 And SJPReady(CurrentPlayer) = 0 And ModeJP_Done Then
    SJPReady(CurrentPlayer) = 1 : PlaySound "Sfx_Cheer1",0,RomSoundVolume
  DiverterON = 2
    ModeReady(CurrentPlayer)=1
  End If

    If L46State(CurrentPlayer) = 1 AND L47State(CurrentPlayer) = 0 Then L47State(CurrentPlayer) = 1 : L48State(CurrentPlayer) = 2
    If L46State(CurrentPlayer) = 0 Then L46State(CurrentPlayer) = 1
  End If
  UpdateLights
End Sub

Sub RGate_Hit
  If sw72.TimerEnabled = True AND ActiveBall.VelX > 0 Then 'LOrbitHit = 3 Then  fixing  was vely > 0

    If CurrentMode = 20 Then
      AddScore (100000)
      WizardJPUpdate
      Exit Sub
    Elseif CurrentMode = 21 Then
      addscore (25000)
      If collect2 > 0 Then collect2 = collect2 - 1 : WizardJPUpdate
      Exit Sub
    Else
      If RampComboSelect <> 3  Then CheckCombo
      RampCombo = frame + 500
      RampComboSelect = 3
    End If

    If collect2 = 1 Then collect2 = 0 : CheckCollected

    If CurrentMode <> 6 Then AddScore(100000)
    If CurrentMode = 6 Then
      ModeJP_Done = True
      ModeJP_Count = ModeJP_Count + 1
      DOF 132,2

      If ModeJP_Count = 3 Then
        AddScore(10000000+HurryUpCollected(CurrentPlayer)*1000000)
        DMDMessage1 "MEGA"      ,"JACKPOT"          ,FontSponge12bb ,69,17  ,69,17  ,"noslide"  ,2.00 ,"BG001"  ,"novideo"
        DMDMessage1 formatscore ( 10 + HurryUpCollected(CurrentPlayer)*1) & "MILLION"   ,""         ,FontSponge12bb ,54,14  ,54,14  ,"noslide2blink"  ,2.00 ,"BG001"  ,"novideo"
        HurryUpCollected(CurrentPlayer) = HurryUpCollected(CurrentPlayer) + 1
        DMDMessage1 "JP LEVEL"    ,"IS NOW " & HurryUpCollected(CurrentPlayer),FontSponge12bb ,55,18  ,55,18  ,"solid"  ,1  ,"BG009"  ,"novideo"
        JackpotSequence(1000) : JackpotSfx
        sw37.TimerInterval = sw37.TimerInterval + 2000
      Else
        AddScore(5000000+HurryUpCollected(CurrentPlayer)*1000000)
        Call JackpotSequence(1000) : JackpotSfx
        DMDMessage1 "JACKPOT"   ,""     ,FontSponge16bb ,69,16  ,69,16  ,"noslide"  ,0.75 ,"BG001"  ,"novideo"
        DMDMessage1 formatscore ( 5 + HurryUpCollected(CurrentPlayer)*1 ) ,"MILLION"      ,FontSponge16bb ,69,14  ,69,14  ,"noslide2blink"  ,0.75 ,"BG001"  ,"novideo"
      End If
    End If
    If L26State(CurrentPlayer) = 2 And ModeReady(CurrentPlayer)= 0 And MultiBallReady(CurrentPlayer) = 0 And ModeJP_Done Then
      L26State(CurrentPlayer) = 1 : DiverterON = 2
      ModeReady(CurrentPlayer)=1 : NextMode(CurrentPlayer)=6
      DMDMessage1 "MODE "   ,"READY"        ,FontSponge16bb ,69,16  ,69,16  ,"solid"  , 2 ,"BG001"  ,"novideo"
      ModeText(CurrentPlayer) = "Boating School Mode" : PlaySound "Sfx_ModeReady",0,RomSoundVolume
    End If

  If ModesComplete(CurrentPlayer) = 9 And SJPReady(CurrentPlayer) = 0 And ModeJP_Done Then
    SJPReady(CurrentPlayer) = 1 : PlaySound "Sfx_Cheer1",0,RomSoundVolume
    DiverterON = 2
    ModeReady(CurrentPlayer)=1
  End If

    If L24State(CurrentPlayer) = 1 AND L25State(CurrentPlayer) = 0 Then L25State(CurrentPlayer) = 1 : L26State(CurrentPlayer) = 2
    If L24State(CurrentPlayer) = 0 Then L24State(CurrentPlayer) = 1
  End If
  UpdateLights
End Sub



'********************
'Targets
'********************
Sub CheckCollected
  If collect1 + collect2 + collect3 + collect4 + collect5 + collect6 + collect7 + collect8 + collect9 = 0 Then
    FanfareSfx
  Else
    PlaySound "Sfx_Airhorn",0,RomSoundVolume
  End If
  UpdateLights
End Sub

Sub sw56_Hit:STHit 56:ShakeSquidward=1:ToyShake:PlaySoundAt "AlienShake1", alien5_BM_Lit_Room:LeftTargetsHit:End Sub
Sub sw57_Hit:STHit 57:ShakeSquidward=1:ToyShake:PlaySoundAt "AlienShake2", alien5_BM_Lit_Room:LeftTargetsHit:End Sub
Sub sw58_Hit:STHit 58:ShakeSquidward=1:ToyShake:PlaySoundAt "AlienShake3", alien5_BM_Lit_Room:LeftTargetsHit:End Sub
Sub LeftTargetsHit
 DOF 117,2

    dim light
    For each light in GiStrings
        If light.name = "l103" Then
            lightCtrl.Pulse light, 1
        Else
            lightCtrl.PulseWithProfile light, Array(0,0,0,0,0,0,0,0), 1
        End If
    Next



  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
    PlayQuote(1)
    Exit Sub
  Elseif CurrentMode = 21 Then
    addscore (25000)
    If collect1 > 0 Then collect1 = collect1 - 1 : WizardJPUpdate
    PlayQuote(1)
    Exit Sub
  End If

  If collect1 = 1 Then collect1 = 0 : CheckCollected

  If CurrentMode <> 7 Then AddScore(200000) : Call PlayQuote(1)
  If CurrentMode = 7 Then
      ModeJP_Done = True
      ModeJP_Count = ModeJP_Count + 1
      DOF 132,2

      If ModeJP_Count = 3 Then
        AddScore(10000000+HurryUpCollected(CurrentPlayer)*1000000)
        DMDMessage1 "MEGA"      ,"JACKPOT"          ,FontSponge12bb ,69,17  ,69,17  ,"noslide"  ,2.00 ,"BG001"  ,"novideo"
        DMDMessage1 formatscore ( 10 + HurryUpCollected(CurrentPlayer)*1) & "MILLION"   ,""         ,FontSponge12bb ,54,14  ,54,14  ,"noslide2blink"  ,2.00 ,"BG001"  ,"novideo"
        HurryUpCollected(CurrentPlayer) = HurryUpCollected(CurrentPlayer) + 1
        DMDMessage1 "JP LEVEL"    ,"IS NOW " & HurryUpCollected(CurrentPlayer),FontSponge12bb ,55,18  ,55,18  ,"solid"  ,1  ,"BG009"  ,"novideo"
        JackpotSequence(1000) : JackpotSfx
        sw37.TimerInterval = sw37.TimerInterval + 2000
      Else
        AddScore(2500000+HurryUpCollected(CurrentPlayer)*1000000)
        Call JackpotSequence(1000) : JackpotSfx
        DMDMessage1 "JACKPOT"   ,""     ,FontSponge16bb ,69,16  ,69,16  ,"noslide"  ,0.75 ,"BG001"  ,"novideo"
        DMDMessage1 formatscore ( 2.5 + HurryUpCollected(CurrentPlayer)*1)    ,"MILLION"      ,FontSponge16bb ,69,14  ,69,14  ,"noslide2blink"  ,0.75 ,"BG001"  ,"novideo"
    End If
  End If

  If L20State(CurrentPlayer) = 2 And ModeReady(CurrentPlayer)= 0 And MultiBallReady(CurrentPlayer) = 0 And ModeJP_Done Then
    L20State(CurrentPlayer) = 1 : DiverterON = 2
    ModeReady(CurrentPlayer)=1 : NextMode(CurrentPlayer)=7
    DMDMessage1 "MODE "   ,"READY"        ,FontSponge16bb ,69,16  ,69,16  ,"solid"  , 2 ,"BG001"  ,"novideo"
    ModeText(CurrentPlayer) = "Squidward's Tiki Mode" : PlaySound "Sfx_ModeReady",0,RomSoundVolume
  End If

  If ModesComplete(CurrentPlayer) = 9 And SJPReady(CurrentPlayer) = 0 And ModeJP_Done Then
    SJPReady(CurrentPlayer) = 1 : PlaySound "Sfx_Cheer1",0,RomSoundVolume
    DiverterON = 2
    ModeReady(CurrentPlayer)=1

  End If

  If L18State(CurrentPlayer) = 1 AND L19State(CurrentPlayer) = 0 Then L19State(CurrentPlayer) = 1 : L20State(CurrentPlayer) = 2
  If L18State(CurrentPlayer) = 0 Then L18State(CurrentPlayer) = 1

  UpdateLights

End Sub




Sub sw41_Hit:STHit 41:ShakePatrick=1:ToyShake:PlaySoundAt "AlienShake1", alien14_BM_Lit_Room:RightTargetsHit:End Sub
Sub sw42_Hit:STHit 42:ShakePatrick=1:ToyShake:PlaySoundAt "AlienShake2", alien14_BM_Lit_Room:RightTargetsHit:End Sub
Sub sw44_Hit:STHit 44:ShakePatrick=1:ToyShake:PlaySoundAt "AlienShake3", alien14_BM_Lit_Room:RightTargetsHit:End Sub
Sub RightTargetsHit
  DOF 112,2
    dim light
    For each light in GiStrings
        If light.name = "l104" Then
            lightCtrl.Pulse light, 1
        Else
            lightCtrl.PulseWithProfile light, Array(0,0,0,0,0,0,0,0), 1
        End If
    Next



  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
    PlayQuote(4)
    Exit Sub
  Elseif CurrentMode = 21 Then
    addscore (25000)
    If collect9 > 0 Then collect9 = collect9 - 1 : WizardJPUpdate
    PlayQuote(4)
    Exit Sub
  End If

  If collect9 = 1 Then collect9 = 0 : CheckCollected

  If CurrentMode <> 8 Then AddScore(200000) : Call PlayQuote(4)
  If CurrentMode = 8 Then
    ModeJP_Done = True
    ModeJP_Count = ModeJP_Count + 1
      DOF 132,2

    If ModeJP_Count = 3 Then
      AddScore(10000000+HurryUpCollected(CurrentPlayer)*1000000)
      DMDMessage1 "MEGA"      ,"JACKPOT"          ,FontSponge12bb ,69,17  ,69,17  ,"noslide"  ,2.00 ,"BG001"  ,"novideo"
      DMDMessage1 formatscore ( 10 + HurryUpCollected(CurrentPlayer)*1) & "MILLION"   ,""         ,FontSponge12bb ,54,14  ,54,14  ,"noslide2blink"  ,2.00 ,"BG001"  ,"novideo"
      HurryUpCollected(CurrentPlayer) = HurryUpCollected(CurrentPlayer) + 1
      DMDMessage1 "JP LEVEL"    ,"IS NOW " & HurryUpCollected(CurrentPlayer),FontSponge12bb ,55,18  ,55,18  ,"solid"  ,1  ,"BG009"  ,"novideo"
      JackpotSequence(1000) : JackpotSfx
      sw37.TimerInterval = sw37.TimerInterval + 2000
    Else
      AddScore(2500000+HurryUpCollected(CurrentPlayer)*1000000)
      Call JackpotSequence(1000) : JackpotSfx
      DMDMessage1 "JACKPOT"   ,""     ,FontSponge16bb ,69,17  ,69,17  ,"noslide"  ,0.75 ,"BG001"  ,"novideo"
      DMDMessage1 formatscore ( 2.5 + HurryUpCollected(CurrentPlayer)*1 )   ,"MILLION"      ,FontSponge16bb ,69,14  ,69,14  ,"noslide2blink"  ,0.75 ,"BG001"  ,"novideo"
    End If
  End If

  If L23State(CurrentPlayer) = 2 And ModeReady(CurrentPlayer)= 0 And MultiBallReady(CurrentPlayer) = 0 And ModeJP_Done Then
    L23State(CurrentPlayer) = 1 : DiverterON = 2
    ModeReady(CurrentPlayer)=1 : NextMode(CurrentPlayer)=8 :  DMDMessage1 "MODE "   ,"READY"        ,FontSponge16bb ,69,16  ,69,16  ,"solid"  , 2 ,"BG001"  ,"novideo" : ModeText(CurrentPlayer) = "Goofy Goobers Mode" : PlaySound "Sfx_ModeReady",0,RomSoundVolume

  End If

  If ModesComplete(CurrentPlayer) = 9 And SJPReady(CurrentPlayer) = 0 And ModeJP_Done Then
    SJPReady(CurrentPlayer) = 1 : PlaySound "Sfx_Cheer1",0,RomSoundVolume
    DiverterON = 2
    ModeReady(CurrentPlayer)=1
  End If


  If L21State(CurrentPlayer) = 1 AND L22State(CurrentPlayer) = 0 Then L22State(CurrentPlayer) = 1 : L23State(CurrentPlayer) = 2
  If L21State(CurrentPlayer) = 0 Then L21State(CurrentPlayer) = 1

  UpdateLights
End Sub





Sub sw43_Hit
  DOF 116,2
  STHit 43
  TargetSfx

  If CurrentMode = 20 Then
    AddScore (100000)
    WizardJPUpdate
    ShakePatty=1 : ToyShake : PlaySoundAt "AlienShake2", alien6_BM_Lit_Room
    Exit Sub
  Elseif CurrentMode = 21 Then
    addscore (25000)
    If collect4 > 0 Then collect4 = collect4 - 1 : WizardJPUpdate

    ShakePatty=1 : ToyShake : PlaySoundAt "AlienShake2", alien6_BM_Lit_Room
    Exit Sub
  End If

  If collect4 = 1 Then collect4 = 0 : CheckCollected

  If SpinnerTimer > 0 Then
    FanfareSfx
    Call JackpotSequence(1000)
    HurryUpCollected(CurrentPlayer) = HurryUpCollected(CurrentPlayer) + 2
    DMDMessage1 "HURRYUP"   ,"COMPLETE "                ,FontSponge12bb ,55,18  ,55,18  ,"solid"  ,1  ,"BG001"  ,"novideo"
    DMDMessage1 "JP LEVEL"    ,"IS NOW " & HurryUpCollected(CurrentPlayer),FontSponge12bb ,55,18  ,55,18  ,"solid"  ,1  ,"BG009"  ,"novideo"
    DMDMessage1 "25"    ,"MILLIONS " & HurryUpCollected(CurrentPlayer),FontSponge12bb ,55,18  ,55,18  ,"solid"  ,1  ,"BG001"  ,"novideo"
    SpinnerLevel1(CurrentPlayer) = 0
    SpinnerTimer = 0
    EndSpinnerHurryUp

    ShakePatty=4 : ToyShake : PlaySoundAt "AlienShake2", alien6_BM_Lit_Room
  Else
    ShakePatty=1 : ToyShake : PlaySoundAt "AlienShake2", alien6_BM_Lit_Room
  End If


  If CurrentMode <> 9 Then AddScore(250000) :


  If CurrentMode = 9 Then
    ModeJP_Done = True
    ModeJP_Count = ModeJP_Count + 1
      DOF 132,2

    If ModeJP_Count = 3 Then
      AddScore(10000000+HurryUpCollected(CurrentPlayer)*1000000)
      DMDMessage1 "MEGA"      ,"JACKPOT"          ,FontSponge12bb ,69,17  ,69,17  ,"noslide"  ,2.00 ,"BG001"  ,"novideo"
      DMDMessage1 formatscore ( 10 + HurryUpCollected(CurrentPlayer)*1) & "MILLION"   ,""         ,FontSponge12bb ,54,14  ,54,14  ,"noslide2blink"  ,2.00 ,"BG001"  ,"novideo"
      HurryUpCollected(CurrentPlayer) = HurryUpCollected(CurrentPlayer) + 1
      DMDMessage1 "JP LEVEL"    ,"IS NOW " & HurryUpCollected(CurrentPlayer),FontSponge12bb ,55,18  ,55,18  ,"solid"  ,1  ,"BG009"  ,"novideo"
      JackpotSequence(1000) : JackpotSfx
      sw37.TimerInterval = sw37.TimerInterval + 2000
    Else
      AddScore(5000000+HurryUpCollected(CurrentPlayer)*1000000)
      Call JackpotSequence(1000) : JackpotSfx
      DMDMessage1 "JACKPOT"   ,""     ,FontSponge16bb ,69,16  ,69,16  ,"noslide"  ,0.75 ,"BG001"  ,"novideo"
      DMDMessage1 formatscore ( 5 + HurryUpCollected(CurrentPlayer)*1 )   ,"MILLION"      ,FontSponge16bb ,69,14  ,69,14  ,"noslide2blink"  ,0.75 ,"BG001"  ,"novideo"
    End If
  End If


  If L32State(CurrentPlayer) = 2 And ModeReady(CurrentPlayer)= 0 And MultiBallReady(CurrentPlayer) = 0 And ModeJP_Done Then
    L32State(CurrentPlayer) = 1 : DiverterON = 2
    ModeReady(CurrentPlayer)=1 : NextMode(CurrentPlayer)=9 :  DMDMessage1 "MODE "   ,"READY"        ,FontSponge16bb ,69,16  ,69,16  ,"solid"  , 2 ,"BG001"  ,"novideo" : ModeText(CurrentPlayer) = "Krusty Krab Mode" : PlaySound "Sfx_ModeReady",0,RomSoundVolume
  End If

  If ModesComplete(CurrentPlayer) = 9 And SJPReady(CurrentPlayer) = 0 And ModeJP_Done Then
    SJPReady(CurrentPlayer) = 1 : PlaySound "Sfx_Cheer1",0,RomSoundVolume
    DiverterON = 2
    ModeReady(CurrentPlayer)=1
  End If

  If L30State(CurrentPlayer) = 1 AND L31State(CurrentPlayer) = 0 Then L31State(CurrentPlayer) = 1 : L32State(CurrentPlayer) = 2
  If L30State(CurrentPlayer) = 0 Then L30State(CurrentPlayer) = 1

  UpdateLights
End Sub


Sub sw75_Hit
 DOF 116,2

  if CurrentMode = 11 Then Partyblinker.Interval = 150 : Partyblinker_Timer : ShakeSponge = 1 : SpongeTimer.Interval = 250 : SpongeTimer.Enabled = True : DMDMessage1 " "     ,""     ,FontSponge12bb ,50,18  ,58,18  ,"blink"  ,1  ,"party"  ,"novideo"

  STHit 75
    TargetSfx
  ShakeMain=1
  ToyShake
  AddScore(100000)

  If L33State(CurrentPlayer) = 2 And SJPReady(CurrentPlayer) = 0 Then L33State(CurrentPlayer) = 1

  If L33State(CurrentPlayer) = 1 AND L34State(CurrentPlayer) = 1 AND L35State(CurrentPlayer) = 1 Then
    TriggerWall.Collidable = False : TriggerWallCollide(CurrentPlayer)=False
    TriggerWall.IsDropped = Not TriggerWall.Collidable
    'sw77.IsDropped = True
    TriggerDown(CurrentPlayer) = True
    sw77MovableHelper
L33State(CurrentPlayer) = 0
L35State(CurrentPlayer) = 0
    L34State(CurrentPlayer) = 2
    L36State(CurrentPlayer) = 2
    If BallsLocked(CurrentPlayer) = 0 AND CurrentMode <> 11 Then L02State(CurrentPlayer) = 2 : DMDMessage1 "LOCK" ,"READY"    ,FontSponge16bb ,69,16  ,69,16  ,"blink"  ,2  ,"BG001"  ,"novideo"
    If BallsLocked(CurrentPlayer) = 1 AND CurrentMode <> 11 Then

      L02State(CurrentPlayer) = 2
      L03State(CurrentPlayer) = 2

      L33State(CurrentPlayer) = 2 : L34State(CurrentPlayer) = 2 : L35State(CurrentPlayer) = 2
      DMDMessage1 "MULTIBALL" ,"READY"    ,FontSponge12bb ,69,16  ,69,16  ,"blink"  ,2  ,"BG001"  ,"novideo"
      Call ChangeGILights(1) : AreLightsOff = 0
      MultiBallReady(CurrentPlayer) = 1
      CurrentMode = 0 : IntroMusic : ModeReady(CurrentPlayer) = 1 : DiverterON = 1

    End If
    If CurrentMode = 11 Then
      L33State(CurrentPlayer) = 2 : L34State(CurrentPlayer) = 2 : L35State(CurrentPlayer) = 2
      SJPReady(CurrentPlayer) = 1
      IntroMusic
    End If
  End If
  UpdateLights
End Sub


Sub sw76_Hit
  DOF 116,2

    if CurrentMode = 11 Then Partyblinker.Interval = 150 : Partyblinker_Timer : ShakeSponge = 1 : SpongeTimer.Interval = 250 : SpongeTimer.Enabled = True  : DMDMessage1 " "      ," "      ,FontSponge12bb ,50,18  ,58,18  ,"blink"  ,1  ,"party"  ,"novideo"
  STHit 76
  TargetSfx
  ShakeMain=1
  ToyShake
  AddScore(100000)
  If L35State(CurrentPlayer) = 2 And SJPReady(CurrentPlayer) = 0 Then L35State(CurrentPlayer) = 1
  If L33State(CurrentPlayer) = 1 AND L34State(CurrentPlayer) = 1 AND L35State(CurrentPlayer) = 1 Then
    TriggerWall.Collidable = False : TriggerWallCollide(CurrentPlayer)=False
    TriggerWall.IsDropped = Not TriggerWall.Collidable
    'sw77.IsDropped = True
    TriggerDown(CurrentPlayer) = True
    sw77MovableHelper
L33State(CurrentPlayer) = 0
L35State(CurrentPlayer) = 0
    L34State(CurrentPlayer) = 2
    L36State(CurrentPlayer) = 2
    If BallsLocked(CurrentPlayer) = 0 AND CurrentMode <> 11 Then L02State(CurrentPlayer) = 2 : DMDMessage1 "LOCK" ,"READY"    ,FontSponge16bb ,69,16  ,69,16  ,"blink"  ,2  ,"BG001"  ,"novideo"
    If BallsLocked(CurrentPlayer) = 1 AND CurrentMode <> 11 Then
      L03State(CurrentPlayer) = 2
      L02State(CurrentPlayer) = 2

      L33State(CurrentPlayer) = 2 : L34State(CurrentPlayer) = 2 : L35State(CurrentPlayer) = 2
      DMDMessage1 "MULTIBALL" ,"READY"    ,FontSponge12bb ,69,16  ,69,16  ,"blink"  ,2  ,"BG001"  ,"novideo"
      Call ChangeGILights(1) : AreLightsOff = 0
      MultiBallReady(CurrentPlayer) = 1
      CurrentMode = 0 : IntroMusic : ModeReady(CurrentPlayer) = 1 : DiverterON = 1

    End If
    If CurrentMode = 11 Then
      L33State(CurrentPlayer) = 2 : L34State(CurrentPlayer) = 2 : L35State(CurrentPlayer) = 2
      SJPReady(CurrentPlayer) = 1
      IntroMusic
    End If
  End If
  UpdateLights
End Sub

'********************
'3 Bank
'********************
Dim StopPinappleWiz2
Sub checkaddaball2
  wiz2Addaball = wiz2Addaball + 1
  If wiz2Addaball = 4 Or wiz2Addaball = 14 Then
    Playsound "Sfx_ModeReady" ,1,RomSoundVolume
  End If

  If wiz2Addaball = 5 Then
    ballstoshoot = ballstoshoot + 1
    wizBiP = wizBiP + 1
    PlaySound "Sfx_SecretShot",1,RomSoundVolume
  End If

  If wiz2Addaball = 15 Then
    ballstoshoot = ballstoshoot + 1 : wizBiP = wizBiP + 1 : StopPinappleWiz2 = 1
    l33.State = 0 : L33.blinkinterval = 125
    l34.State = 0 : L34.blinkinterval = 125
    l35.State = 0 : L35.blinkinterval = 125
    PlaySound "Sfx_SecretShot",1,RomSoundVolume
  End If
  UpdateLights
End Sub

'   If wizAddaBall = 2 Or wizAddaBall = 11 Then
'     Playsound "Sfx_ModeReady" ,1,RomSoundVolume
'     AddABallReady = 1
'     DMDMessage1 "ADDABALL"      ,"READY"        ,FontSponge12bb ,54,16  ,69,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
'   End If
'
'   If wizAddaBall = 3 Then
'     AddABallReady = 0
'     DMDMessage1 "ADDABALL"      ,""       ,FontSponge16bb ,54,16  ,69,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
'     ballstoshoot = ballstoshoot + 1 : wizBiP = wizBiP + 1
'     PlaySound "Sfx_SecretShot",1,RomSoundVolume


Sub sw45_Hit
 DOF 116,2

  If MoveSpeed > 47 Or MoveSpeed = 0 Then

    ShakeThree=1
    ShakeMain=.5
    ToyShake
    BankWobble
    TargetSfx

    If CurrentMode = 20 Then
      AddScore (100000)
      WizardJPUpdate
      Exit Sub
    Elseif CurrentMode = 21 Then
      addscore (25000)
      checkaddaball2
      Exit Sub
    Else
      AddScore (100000)
    End If


    if CurrentMode = 11 Then Partyblinker.Interval = 150 : Partyblinker_Timer : ShakeSponge = 1 : SpongeTimer.Interval = 250 : SpongeTimer.Enabled = True : DMDMessage1 " "     ,""     ,FontSponge12bb ,50,18  ,58,18  ,"blink"  ,1  ,"party"  ,"novideo"
    If ScoopRestart > frame Then
      Exit Sub
    End If

    If CurrentMode = 11 And collect1 + collect2 + collect3 + collect4 + collect5 + collect6 + collect7 + collect8 + collect9 > 0 Then
      ClearDMD
      DMDMessage1 "DOORS"     ,"CLOSED"     ,FontSponge12bb ,60,16  ,54,16  ,"blink"  ,1.3  ,"BG001"  ,"novideo"
      DMDMessage1 "PARTY NEED"  ,"MORE PEOPLE"    ,FontSponge12bb ,45,16  ,54,16  ,"blink"  ,1.3  ,"BG001"  ,"novideo"
      PlaySound "Sfx_SkillMiss",1,RomSoundVolume
      Exit Sub
    End If

'   AddScore(250000)
    BankMove = BankMove + 1
    If L33State(CurrentPlayer) = 0 Then
      L33State(CurrentPlayer) = 1
    Else
      If L34State(CurrentPlayer) = 0 Then
        L34State(CurrentPlayer) = 1
      Else
        L35State(CurrentPlayer) = 1
      End If
    End If

    currpos = 50
    If L33State(CurrentPlayer) = 1 AND L34State(CurrentPlayer) = 1 AND L35State(CurrentPlayer) = 1 Then
      Call Update3Bank(50, 0)
      if BallsLocked(CurrentPlayer) = 1 Then
        L03State(CurrentPlayer) = 1
        If not TriggerDown(CurrentPlayer) then
          L33State(CurrentPlayer) = 2
          L34State(CurrentPlayer) = 2
          L35State(CurrentPlayer) = 2
        Else
          L33State(CurrentPlayer) = 0
          L34State(CurrentPlayer) = 0
          L35State(CurrentPlayer) = 0
        End If

      End If
      L36State(CurrentPlayer) = 1
      L02State(CurrentPlayer) = 1
      If TriggerDown(CurrentPlayer) Then L36State(CurrentPlayer) = 2 : L02State(CurrentPlayer) = 2 : L34State(CurrentPlayer) = 2

    End If
    UpdateLights
  End If
End Sub

Sub sw46_Hit
 DOF 116,2

  If MoveSpeed > 47 Or MoveSpeed = 0 Then

    ShakeThree=1
    ShakeMain=.5
    ToyShake
    BankWobble
    TargetSfx

    If CurrentMode = 20 Then
      AddScore (100000)
      WizardJPUpdate
      Exit Sub
    Elseif CurrentMode = 21 Then
      addscore (25000)
      checkaddaball2
      Exit Sub
    Else
      AddScore (100000)
    End If

    if CurrentMode = 11 Then Partyblinker.Interval = 150 : Partyblinker_Timer : ShakeSponge = 1 : SpongeTimer.Interval = 250 : SpongeTimer.Enabled = True : DMDMessage1 " "     ,""     ,FontSponge12bb ,50,18  ,58,18  ,"blink"  ,1  ,"party"  ,"novideo"
    If ScoopRestart > frame Then
      Exit Sub
    End If
    If CurrentMode = 11 And collect1 + collect2 + collect3 + collect4 + collect5 + collect6 + collect7 + collect8 + collect9 > 0 Then
      ClearDMD
      DMDMessage1 "DOORS"     ,"CLOSED"     ,FontSponge12bb ,60,16  ,54,16  ,"blink"  ,1.3  ,"BG001"  ,"novideo"
      DMDMessage1 "PARTY NEED"  ,"MORE PEOPLE"    ,FontSponge12bb ,45,16  ,54,16  ,"blink"  ,1.3  ,"BG001"  ,"novideo"
      PlaySound "Sfx_SkillMiss",1,RomSoundVolume
      Exit Sub
    End If

    AddScore(250000)
    BankMove = BankMove + 1
    If L34State(CurrentPlayer) = 0 Then
      L34State(CurrentPlayer) = 1
    Else
      If RndInt(1,2) = 1 Then
        If L33State(CurrentPlayer) = 0 Then
          L33State(CurrentPlayer) = 1
        Else
          L35State(CurrentPlayer) = 1
        End If
      Else
        If L35State(CurrentPlayer) = 0 Then
          L35State(CurrentPlayer) = 1
        Else
          L33State(CurrentPlayer) = 1
        End If
      End If
    End If

    currpos = 50
    If L33State(CurrentPlayer) = 1 AND L34State(CurrentPlayer) = 1 AND L35State(CurrentPlayer) = 1 Then
      Call Update3Bank(50, 0)
      if BallsLocked(CurrentPlayer) = 1 Then
        L03State(CurrentPlayer) = 1
        If not TriggerDown(CurrentPlayer) then
          L33State(CurrentPlayer) = 2
          L34State(CurrentPlayer) = 2
          L35State(CurrentPlayer) = 2
        Else
          L33State(CurrentPlayer) = 0
          L34State(CurrentPlayer) = 0
          L35State(CurrentPlayer) = 0
        End If
      End If
      L36State(CurrentPlayer) = 1
      L02State(CurrentPlayer) = 1
      If TriggerDown(CurrentPlayer) Then L36State(CurrentPlayer) = 2 : L02State(CurrentPlayer) = 2 : L34State(CurrentPlayer) = 2
    End If
    UpdateLights
  End If
End Sub

Sub sw47_Hit
 DOF 116,2
  If MoveSpeed > 47 Or MoveSpeed = 0 Then

    ShakeThree=1
    ShakeMain=.5
    ToyShake
    BankWobble
    TargetSfx

    If CurrentMode = 20 Then
      AddScore (100000)
      WizardJPUpdate
      Exit Sub
    Elseif CurrentMode = 21 Then
      addscore (25000)
      checkaddaball2
      Exit Sub
    Else
      AddScore (100000)
    End If

    if CurrentMode = 11 Then Partyblinker.Interval = 150 : Partyblinker_Timer : ShakeSponge = 1 : SpongeTimer.Interval = 250 : SpongeTimer.Enabled = True : DMDMessage1 " "     ,""     ,FontSponge12bb ,50,18  ,58,18  ,"blink"  ,1  ,"party"  ,"novideo"
    If ScoopRestart > frame Then
      Exit Sub
    End If
    If CurrentMode = 11 And collect1 + collect2 + collect3 + collect4 + collect5 + collect6 + collect7 + collect8 + collect9 > 0 Then
      ClearDMD
      DMDMessage1 "DOORS"     ,"CLOSED"     ,FontSponge12bb ,60,16  ,54,16  ,"blink"  ,1.3  ,"BG001"  ,"novideo"
      DMDMessage1 "PARTY NEED"  ,"MORE PEOPLE"    ,FontSponge12bb ,45,16  ,54,16  ,"blink"  ,1.3  ,"BG001"  ,"novideo"
      PlaySound "Sfx_SkillMiss",1,RomSoundVolume
      Exit Sub
    End If

    AddScore(250000)
    BankMove = BankMove + 1
    If L35State(CurrentPlayer) = 0 Then
      L35State(CurrentPlayer) = 1
    Else
      If L34State(CurrentPlayer) = 0 Then
        L34State(CurrentPlayer) = 1
      Else
        L33State(CurrentPlayer) = 1
      End If
    End If
    currpos = 50
    If L33State(CurrentPlayer) = 1 AND L34State(CurrentPlayer) = 1 AND L35State(CurrentPlayer) = 1 Then
      Call Update3Bank(50, 0)
      if BallsLocked(CurrentPlayer) = 1 Then
        L03State(CurrentPlayer) = 1
        If not TriggerDown(CurrentPlayer) then
          L33State(CurrentPlayer) = 2
          L34State(CurrentPlayer) = 2
          L35State(CurrentPlayer) = 2
        Else
          L33State(CurrentPlayer) = 0
          L34State(CurrentPlayer) = 0
          L35State(CurrentPlayer) = 0
        End If
      End If
      L36State(CurrentPlayer) = 1
      L02State(CurrentPlayer) = 1
      If TriggerDown(CurrentPlayer) Then L36State(CurrentPlayer) = 2 : L02State(CurrentPlayer) = 2 : L34State(CurrentPlayer) = 2
    End If
    UpdateLights
  End If
End Sub

'********************
'Droptarget
'********************
Sub TriggerWall_Hit
 DOF 116,2

  TargetSfx
  ShakeMain= 1
  ToyShake
  AddScore(100000)

  Partyblinker.Interval = 150 : Partyblinker_Timer


  If SJPReady(CurrentPlayer) = 0 Then L34State(CurrentPlayer) = 1


  If L33State(CurrentPlayer) = 1 AND L34State(CurrentPlayer) = 1 AND L35State(CurrentPlayer) = 1 Then



      TriggerWall.Collidable = False  : TriggerWallCollide(CurrentPlayer)=False
      TriggerWall.IsDropped = Not TriggerWall.Collidable
      'sw77.IsDropped = True
      TriggerDown(CurrentPlayer) = True
      sw77MovableHelper
      L33State(CurrentPlayer) = 0
      L35State(CurrentPlayer) = 0
      L34State(CurrentPlayer) = 2
      L36State(CurrentPlayer) = 2
      If BallsLocked(CurrentPlayer) = 0 AND CurrentMode <> 11 Then L02State(CurrentPlayer) = 2 : DMDMessage1 "LOCK" ,"READY"    ,FontSponge16bb ,69,16  ,69,16  ,"blink"  ,2  ,"BG001"  ,"novideo"
      If BallsLocked(CurrentPlayer) = 1 AND CurrentMode <> 11 Then

        L03State(CurrentPlayer) = 2
        L02State(CurrentPlayer) = 2

        L33State(CurrentPlayer) = 0 : L34State(CurrentPlayer) = 2 : L35State(CurrentPlayer) = 0
        DMDMessage1 "MULTIBALL" ,"READY"    ,FontSponge12bb ,69,16  ,69,16  ,"blink"  ,2  ,"BG001"  ,"novideo"
        Call ChangeGILights(1)  : AreLightsOff = 0
        MultiBallReady(CurrentPlayer) = 1
        CurrentMode = 0 : IntroMusic : ModeReady(CurrentPlayer) = 1 :DiverterON = 1

      End If
      If CurrentMode = 11 Then
        L33State(CurrentPlayer) = 0 : L34State(CurrentPlayer) = 2 : L35State(CurrentPlayer) = 0
        SJPReady(CurrentPlayer) = 1
        IntroMusic
      End If


  End If
  UpdateLights
End Sub

Sub sw77_Hit
    if CurrentMode = 11 Then ShakeSponge = 1 : SpongeTimer.Interval = 250 : SpongeTimer.Enabled = True : DMDMessage1 " "      ,""     ,FontSponge12bb ,50,18  ,58,18  ,"blink"  ,1  ,"party"  ,"novideo"
End Sub

Sub PlungerLaneTrigger_Hit
  BallSaverTimer.Enabled = True
  If PlayTime < BallSaveTime(CurrentPlayer)-2 Then
    l01.State=2
  Else
    If ExtraBall > 0 Then l01.State=1
  End If
' If BallInLane = 1 Then
    SkillShot = 0
    SSText = 0
' End If

' BallinLane = 0
  lightCtrl.LightOff f147
  DrainSave = 0
End Sub

'********************
'Ball Save Timer
'********************
Dim noshakebigone
Dim LastScoreTime
Dim showtimestopp : showtimestopp = False
Sub BallSaverTimer_Timer()

  If CurrentMode = 20 Then LightSeqWiz1.play SeqBlinking,,1,35 : LightSeqWiz1.play SeqBlinking,,1,60

  If LastScoreTime > frame Then showtimestopp = False

  If LastScoreTime + 100 < frame And showtimestopp = False And not bTilted And CurrentMSGdone = 0 Then
    showtimestopp = True
    If CurrentMode = 0 Then
      If ModeReady(CurrentPlayer) = 1 Then
        DMDMessage1 "START","NEXT MODE"   ,FontSponge12bb ,54,17  ,54,17  ,"solid"  ,1.5  ,"BG001"  ,"novideo"
      Else
        DMDMessage1 "KEEP","SHOOTING"   ,FontSponge12bb ,54,17  ,54,17  ,"solid"  ,1.5  ,"BG001"  ,"novideo"
      End If
    Elseif CurrentMode < 10 Then
      If ModeJP_Count = 0 Then
        DMDMessage1 "GO FOR","JACKPOT"    ,FontSponge12bb ,54,17  ,54,17  ,"solid"  ,1.5  ,"BG001"  ,"novideo"
      Elseif ModeJP_Count < 3 Then
        DMDMessage1 "GO FOR","THREE JACKPOTS"   ,FontSponge12bb ,54,17  ,54,17  ,"solid"  ,1.5  ,"BG001"  ,"novideo"
      Else
        DMDMessage1 "KEEP","SHOOTING"   ,FontSponge12bb ,54,17  ,54,17  ,"solid"  ,1.5  ,"BG001"  ,"novideo"
      End If
    Elseif currentmode = 11 Then
      DMDMessage1 "LOCK","THE BALL"   ,FontSponge12bb ,54,17  ,54,17  ,"solid"  ,1.5,"BG001"  ,"novideo"
    Else
      DMDMessage1 "ITS NOT","OVER"    ,FontSponge12bb ,54,17  ,54,17  ,"solid"  ,1.5  ,"BG001"  ,"novideo"
    End If
  End If


  If WizardTimer > 0 And LastScoreTime > frame Then
    WizardTimer = WizardTimer - 1
    If WizardTimer < 1 Then
      StopTableMusic
      songtimer.enabled = False
      DMDMessage1 "TIME"      ,"I S   U P"    ,FontSponge12bb ,45,17  ,69,17  ,"solid"  ,5  ,"BG001"  ,"novideo"

      PlaySound "wiz3",0,RomSoundVolume

      bTilted = True
      FlipperDeActivate LeftFlipper, LFPress : SolLFlipper False : SolULFlipper False: FlipperDeActivate RightFlipper, RFPress : SolRFlipper False
      PlaySound "Sfx_SuperJackpot",0,RomSoundVolume
      ResetModeLights
      UpdateLights
      EndWizard1.Interval = 7000
      EndWizard1.Enabled = True
      lightCtrl.AddTableLightSeq "GI", lSeqRainbow
      InsertSequence.play SeqBlinking,,11,150
      playtime = 100
      WizardTimer = 0
      WizardWaitForNoBalls = 1
    End If

  End If

  If modetimecounter > 0 And LastScoreTime > frame Then
    modetimecounter = modetimecounter - 1
'   debug.print "modetimevounter:" & modetimecounter
    If modetimecounter < 1 Then
      modetimecounter = 0
      ModeJP_Done = True
      ModeTimer_Timer
    End If
  End If

  PlayTime = PlayTime + 1

'debug.print "currentmode" & CurrentMode & "playtime " & playtime & "  ballsave =" & BallSaveTime(CurrentPlayer)

  If SpinnerTimer > 0 And LastScoreTime > frame Then
    SpinnerTimer = SpinnerTimer - 1
    If SpinnerTimer < 1 Then
      SpinnerTimer = 0
      PlaySound "Sfx_SkillMiss",0,RomSoundVolume
      EndSpinnerHurryUp
      CurrentMSGdone = 0 : DMD_Display_Timer
    End If
  End If




' If playtime > 8 then l01.blinkinterval = 125 else l01.blinkinterval = BallSaveTime(CurrentPlayer)*18 + 40 - (playtime*18)

  If PlayTime > BallSaveTime(CurrentPlayer)-7 Then l01.blinkinterval = 60 Else l01.blinkinterval = 125

  If PlayTime < BallSaveTime(CurrentPlayer)-2 Then l01.State=2

  If PlayTime > BallSaveTime(CurrentPlayer)-3 Then
    If ExtraBall > 0 Then l01.State=1 Else l01.State=0
  End If


End Sub

'**********************
' Skill Shot
'**********************
Sub SkillShotDMD(nr,time)
  dim x
  x = "noslide2"
  If time = 0.8 Then x = "noslide3"
  If nr = 6 Then x = "fastblink"
    If DoOrDieMode > 2 Then
      Select Case nr
        case 0 : DMDMessage1 "WHO"      ,""     ,FontSponge12bb   ,54,16  ,54,16  ,x  ,time ,"dodss"  ,"novideo"
        case 1 : DMDMessage1 "LIVES"    ,""     ,FontSponge12bb   ,54,16  ,54,16  ,x  ,time ,"dodss"  ,"novideo"
        case 2 : DMDMessage1 "IN A"     ,""     ,FontSponge12bb   ,54,16  ,54,16  ,x  ,time ,"dodss"  ,"novideo"
        case 3 : DMDMessage1 "PINEAPPLE"  ,""     ,FontSponge12bb2    ,54,16  ,54,16  ,x  ,time ,"dodss"  ,"novideo"
        case 4 : DMDMessage1 "UNDER"    ,""     ,FontSponge12bb   ,54,16  ,54,16  ,x  ,time ,"dodss"  ,"novideo"
        case 5 : DMDMessage1 "THE SEA"    ,""     ,FontSponge12bb   ,54,16  ,54,16  ,x  ,time ,"dodss"  ,"novideo"
        case 6 : CLEARDMD: DMDMessage1 "PINEAPPLE"  ,"" ,FontSponge12bb2  ,54,16  ,54,16  ,x  ,time ,"dodss"  ,"novideo"
      End Select
    Else
      Select Case nr
        case 0 : DMDMessage1 "WHO"      ,""     ,FontSponge12bb   ,54,16  ,54,16  ,x  ,time ,"BG001"  ,"novideo"
        case 1 : DMDMessage1 "LIVES"    ,""     ,FontSponge12bb   ,54,16  ,54,16  ,x  ,time ,"BG001"  ,"novideo"
        case 2 : DMDMessage1 "IN A"     ,""     ,FontSponge12bb   ,54,16  ,54,16  ,x  ,time ,"BG001"  ,"novideo"
        case 3 : DMDMessage1 "PINEAPPLE"  ,""     ,FontSponge12bb2    ,54,16  ,54,16  ,x  ,time ,"BG001"  ,"novideo"
        case 4 : DMDMessage1 "UNDER"    ,""     ,FontSponge12bb   ,54,16  ,54,16  ,x  ,time ,"BG001"  ,"novideo"
        case 5 : DMDMessage1 "THE SEA"    ,""     ,FontSponge12bb   ,54,16  ,54,16  ,x  ,time ,"BG001"  ,"novideo"
        case 6 : CLEARDMD: DMDMessage1 "PINEAPPLE"  ,"" ,FontSponge12bb2  ,54,16  ,54,16  ,x  ,time ,"BG001"  ,"novideo"
      End Select
    End If
End Sub

Sub SkillShotTimer_Timer
  SkillShotTimer.Enabled = False
  If SkillShot = 1 Then
    If Not bInOptions Then SkillShotSfx
    If SkillShot2 = 0 Then SkillShotDMD 0,0.16
    If SkillShot2 = 1 Then SkillShotDMD 1,0.16
    If SkillShot2 = 2 Then SkillShotDMD 2,0.16
    If SkillShot2 = 3 Then SkillShotDMD 3,0.16
    If SkillShot2 = 4 Then SkillShotDMD 4,0.16
    If SkillShot2 = 5 Then SkillShotDMD 5,0.16

    SkillShot2 = SkillShot2 + 1
    If SkillShot2 = 6 Then SkillShot2 = 0
    SkillShotTimer.Enabled = True
  Else
    SkillShot2 = SkillShot2 - 1
    If SkillShot2 = -1 Then SkillShot2 = 5

    SkillShotDMD SkillShot2,0.8
    If SkillShot2 = 3 Then
      SkillShotDMD 6,1.1
      IntroMusic
      InsertSequence.Play SeqRandom,10,,750

      DMDMessage1 "SKILLSHOT"     ,""     ,FontSponge16bb ,54,16  ,54,16  ,"blink"  ,1  ,"BG001"  ,"novideo"
      Select Case SSmade(CurrentPlayer)
        case 0
          DMDMessage1 "10 MILLION"      ,""     ,FontSponge16bb ,54,16  ,54,16  ,"blink"  ,1  ,"BG001"  ,"novideo"
          AddScore 10000000
        case 1
          DMDMessage1 "20 MILLION"      ,""     ,FontSponge16bb ,54,16  ,54,16  ,"blink"  ,1  ,"BG001"  ,"novideo"
          AddScore 20000000
        case 2
          DMDMessage1 "25 MILLION"      ,""     ,FontSponge16bb ,54,16  ,54,16  ,"blink"  ,1  ,"BG001"  ,"novideo"
          AddScore 25000000
        case 3
          DMDMessage1 "30 MILLION"      ,""     ,FontSponge16bb ,54,16  ,54,16  ,"blink"  ,1  ,"BG001"  ,"novideo"
          AddScore 30000000
        case 4
          DMDMessage1 "35 MILLION"      ,""     ,FontSponge16bb ,54,16  ,54,16  ,"blink"  ,1  ,"BG001"  ,"novideo"
          AddScore 30000000
        case 5
          DMDMessage1 "40 MILLION"      ,""     ,FontSponge16bb ,54,16  ,54,16  ,"blink"  ,1  ,"BG001"  ,"novideo"
          AddScore 40000000
        case 6,7,8
          DMDMessage1 "45 MILLION"      ,""     ,FontSponge16bb ,54,16  ,54,16  ,"blink"  ,1  ,"BG001"  ,"novideo"
          AddScore 50000000

      End Select
      SSmade(CurrentPlayer) = SSmade(CurrentPlayer) + 1
      PlaySound "Sfx_Airhorn",0,RomSoundVolume
    Else
      SkillShotDMD SkillShot2,0.8
      StopTableMusic  : PlaySound "Sfx_SkillMiss",0,RomSoundVolume : DelayMusic = 0 : SongTimer.Interval = 459 : SongTimer.Enabled = True
      DMDMessage1 "M I S S"     ,""     ,FontSponge16bb ,68,16  ,68,16  ,"blink"  ,1  ,"BG001"  ,"novideo"
      PlaySound "Sfx_SkillMiss",0,RomSoundVolume
    End If
    SkillShot = 0
  End If
End Sub


'





Sub SkillShotMessage_Timer


  SkillShotMessage.Enabled = False : SkillShotMessage.Interval = 10

' If SSText = 0 AND SkillShot = 0 AND BallInLane = 1 Then DMDText.Text = "HOLD PLUNGER BUTTON" : LEDText("HOLD PLUNGER BUTTON")
' If SSText = 1 AND SkillShot = 0 AND BallInLane = 1 Then DMDText.Text = "STOP ON PINEAPPLE FOR SKILL SHOT" : LEDText("STOP ON PINEAPPLE FOR SKILL SHOT")
' If SSText = 2 AND SkillShot = 0 AND BallInLane = 1 Then DMDText.Text = "PLAYER " & CurrentPlayer : LEDText("PLAYER " & CurrentPlayer)
' If SkillShot = 0 AND BallInLane = 1 Then SSText = SSText + 1
' If SSText = 3 Then SSText = 0
' If BallInLane = 1 AND SkillShot = 0 Then SkillShotMessage.Enabled = True : SkillShotMessage.Interval = 5000

  If SkillShot = 0 And BallInLane = 1 Then
    Select Case SSText
      case 0 : SSText = 1 : line4 = 1 : msg4 = "USE PLUNGER BUTTON"
      case 1 : SSText = 2 : line4 = 1 : msg4 = "STOP ON PINEAPPLE"
      case 2 : SSText = 3 : line4 = 1 : msg4 = "FOR SKILLSHOT"
      case 3 : SSText = 4
      case 4 : SSText = 5 : line4 = 1 : msg4 = "PLAYER " & CurrentPlayer
      case 5 : SSText = 0


    End Select
    : SkillShotMessage.Interval = 3000 : SkillShotMessage.Enabled = True
  End If

End Sub

'*********
'Solenoids
'*********

Sub Auto_Plunger
        PlungerIM.AutoFire
End Sub

'******************
'Motor Bank Up Down
'******************

Sub Update3Bank(currpos, lastpos)
    If currpos <> lastpos Then
    PlaySoundAt "TargetBank" ,BankSound
    MoveSpeed = 0
    c3BankTimer.Enabled = True
    End If
    If currpos > 40 Then
    NewBankDown = 1
        sw45.Isdropped = 1
        sw46.Isdropped = 1
        sw47.Isdropped = 1
        'UFORotSpeedMedium
    End If
    If currpos < 10 Then
    NewBankDown = 0
        sw45.Isdropped = 0
        sw46.Isdropped = 0
        sw47.Isdropped = 0
        'UFORotSpeedSlow
    End If
End Sub

Sub c3BankTimer_Timer
  Dim bl, a

  MoveSpeed = MoveSpeed + 1
  If currpos = 50 AND MoveSpeed < 48 Then

    BackBank.Z = BackBank.Z - 1
    swp45_BM_Lit_Room_Z = swp45_BM_Lit_Room_Z - 1
    swp46_BM_Lit_Room_Z = swp46_BM_Lit_Room_Z - 1
    swp47_BM_Lit_Room_Z = swp47_BM_Lit_Room_Z - 1
    BankDown(CurrentPlayer) = 1
  End If
  If currpos = 0 AND MoveSpeed < 48 Then
    BackBank.Z = BackBank.Z + 1
    swp45_BM_Lit_Room_Z = swp45_BM_Lit_Room_Z + 1
    swp46_BM_Lit_Room_Z = swp46_BM_Lit_Room_Z + 1
    swp47_BM_Lit_Room_Z = swp47_BM_Lit_Room_Z + 1
    BankDown(CurrentPlayer) = 0
  End If
  If MoveSpeed > 47 Then
    c3BankTimer.Enabled = False
    UpdateLights
    If BankDown(CurrentPlayer) = 1 AND NewBankDown = 0 Then currpos = 0 : Call Update3Bank(0,50)
  End If

  'Move the bank targets
  a = swp45_BM_Lit_Room_z - 17.03
  if a < -58 then a = -58
  For each bl in sw45_bl : bl.z = a : Next

  a = swp46_BM_Lit_Room_z - 17.03
  if a < -58 then a = -58
  For each bl in sw46_bl : bl.z = a : Next

  a = swp47_BM_Lit_Room_z - 17.03
  if a < -58 then a = -58
  For each bl in sw47_bl : bl.z = a : Next


End Sub

'*************
' Toy and Target Bank Shake
'*************

Dim ccBall
Const cMod = .65 'percentage of hit power transfered to the 3 Bank of targets
dim ShakeCount
dim WobbleValue
Const WobbleScale = 1

ShakeMain = 0
ShakeThree = 0
ShakeSponge = 0
ShakePatty = 0
ShakePatrick = 0
ShakeSquidward = 0

ShakeCount = 0
WobbleValue = 0

Sub BankWobble
    WobbleValue = cor.ballvel(activeball.id)/10
    BankWobbleTimer.enabled = True
End Sub

Sub ToyShake 'when nudging
  ShakeThree=1
  ShakeMain=.5
  ShakeCount = 0
    ToyShakeTimer.enabled = True
End Sub



BankWobbleTimer.interval = 34 ' Controls the speed of the wobble
Sub BankWobbleTimer_timer
  Dim bl, a, x, y, z

    if WobbleValue < 0 then
        WobbleValue = abs(WobbleValue) * 0.9 - 0.1
    Else
        WobbleValue = -abs(WobbleValue) * 0.9 + 0.1
    end if

    if abs(WobbleValue) < 0.1 Then
        WobbleValue = 0
    swp45_BM_Lit_Room_transy = WobbleValue
    swp46_BM_Lit_Room_transy = WobbleValue
    swp47_BM_Lit_Room_transy = WobbleValue
        BankWobbleTimer.Enabled = False
    end If

  swp45_BM_Lit_Room_transy = WobbleValue
  swp46_BM_Lit_Room_transy = WobbleValue
  swp47_BM_Lit_Room_transy = WobbleValue

  'Move the bank targets
  x = swp45_BM_Lit_Room_transx
  y = -swp45_BM_Lit_Room_transy
  z = swp45_BM_Lit_Room_transz
  For each bl in sw45_bl : bl.transx = x : bl.transy = y : bl.transz = z : Next

  x = swp46_BM_Lit_Room_transx
  y = -swp46_BM_Lit_Room_transy
  z = swp46_BM_Lit_Room_transz
  For each bl in sw46_bl : bl.transx = x : bl.transy = y : bl.transz = z : Next

  x = swp47_BM_Lit_Room_transx
  y = -swp47_BM_Lit_Room_transy
  z = swp47_BM_Lit_Room_transz
  For each bl in sw47_bl : bl.transx = x : bl.transy = y : bl.transz = z : Next

End Sub



Sub ToyShakeTimer_Timer            'start animation
    Dim x, y
  ShakeCount = ShakeCount + 1
  If ShakeCount < 7 Then
    If noshakebigone = 0 Then Ufo1_BM_Lit_Room_transx = ((rnd(1) * 20) - 10) * ShakeMain
    If noshakebigone = 0 Then Ufo1_BM_Lit_Room_transz = rnd(1)  * 10 * ShakeMain
    alien5_BM_Lit_Room_Transy = 20 * ShakeSquidward : If ShakeSquidward = 1 Then Wall002.IsDropped = False : alien5shake2 = 12
    alien6_BM_Lit_Room_Transy = 20 * ShakePatty  : If ShakePatty = 1 Then Wall004.IsDropped = False : alien6shake2 = 12
    alien14_BM_Lit_Room_Transy = 20 * ShakePatrick  : If ShakePatrick = 1 Then Wall003.IsDropped = False : alien14shake2 = 12
  Else
    ToyShakeTimer.enabled = False
    noshakebigone = 0
    Ufo1_BM_Lit_Room_transx = 0
    Ufo1_BM_Lit_Room_transz = 0
    alien5_BM_Lit_Room_Transy = 0 : Wall002.IsDropped = True : If ShakeSquidward = 1 Then alien5shake2 = 9
    alien6_BM_Lit_Room_Transy = 0 : Wall004.IsDropped = True : If ShakePatty = 1 Then alien6shake2 = 9
    alien14_BM_Lit_Room_Transy = 0 : Wall003.IsDropped = True : If ShakePatrick = 1 Then alien14shake2 = 9
    ShakeMain = 0
    ShakeThree = 0
    ShakePatty = 0
    ShakePatrick = 0
    ShakeSquidward = 0
    ShakeCount = 0
  End If
  CharactersMovableHelper
End Sub




Sub SpongeTimer_Timer
  SpongeTimer.Enabled = False
  If ShakeSponge > 0 Then
    alien8_BM_Lit_Room_Transy = 20 : Wall005.IsDropped = False
    PlaySoundAt "AlienShake1", alien8_BM_Lit_Room
    ShakeSponge = ShakeSponge - 1
    ShakeDown.Enabled = True
    spongeshaker2 = 12
  Else
    alien8_BM_Lit_Room_Transy = 0 :  : Wall005.IsDropped = True
  End If
  CharactersMovableHelper
End Sub

Sub ShakeDown_Timer
  ShakeDown.Enabled = False
  alien8_BM_Lit_Room_Transy = 0
  spongeshaker2 = 9
  CharactersMovableHelper
  SpongeTimer.Interval = 125
  SpongeTimer.Enabled = True
End Sub

'********************
'Adding Score
'********************
Dim scoreblink

Sub AddScore(points)

  LastScoreTime = frame + 180

  DodNext = 9
  dim x
  If points > 12345 Then
    For x = 0 to 10
      If SmalScoring(x) = 0 Then SmalScoring(x)=points : Exit for
    Next
  End If
  scoreblink = 1
  If Drain.TimerEnabled = False Then RoundScore = RoundScore + points
  TotalScore(CurrentPlayer) = TotalScore(CurrentPlayer) + Int(points/10) * 10

' If Int(points/10) * 10 <> points Then Debug.print "scoring=" & points

  If miniscoreblinks < 14 and points > 900000 Then miniscoreblinks = 14 : Exit Sub
  If miniscoreblinks < 10 and points > 222000 Then miniscoreblinks = 8 : Exit Sub
  If points > 75000 Then miniscoreblinks = 4 : Exit Sub
  If miniscoreblinks < 2 Then miniscoreblinks = 2
End Sub


Sub DelayScore(dscore, dtime)
  DScoreValue = dscore
  DelayScoreTimer.Interval = dtime * 1000
  DelayScoreTimer.Enabled = True
End Sub

Sub DelayScoreTimer_Timer
  DelayScoreTimer.Enabled = False
  AddScore(DScoreValue)
End Sub

'**********************
' DMD Display
'**********************
Dim h, j



' DMDMessage1 "WARNING"     ,""       ,FontSponge16 ,64,16  ,64,16  ,"shake"  ,5  ,"BGwarn" ,"novideo"
' DMDMessage1 "TILT"        ,""       ,FontSponge32 ,64,16  ,64,16  ,"solid"  ,5  ,"BGtilt" ,"novideo"

'max 11 options
Sub DMDMessage1 (message,message2,Font,posX,posY,endX,endY,state,time,image,video)
  Dim x
  For h = 0 to 9
    If DMDDisplay(h,0) = "" Then
      DMDDisplay(h,0) = Message
      DMDDisplay(h,1) = Message2
      DMDDisplay(h,2) = Font
      DMDDisplay(h,3) = posX
      DMDDisplay(h,4) = posY
      DMDDisplay(h,5) = endX
      DMDDisplay(h,6) = endY
      DMDDisplay(h,7) = state
      DMDDisplay(h,8) = time
      DMDDisplay(h,9) = image
      DMDDisplay(h,10) = video
      if message2 ="" then DMDDisplay(h,11)= 1 Else DMDDisplay(h,11) = 2

      If h = 0 Then
        FlexMsgTime = Frame + time * 63 + 5 ' little bit extra so dont blink
        FlexMsg  = 1   ' START POS ! frame 1
        For x = 0 to 12
          DMDDisplay(20,x) = DMDDisplay(0,x)
        Next
      End If
      Exit For
    End If
  Next

' If DMD_Display.Enabled = False Then
'   DMD_Display.Interval = time * 1000
'   DMD_Display.Enabled = True
' If FlexMsg = 0 Then
'   FlexMsgTime = Frame + time * 63 + 33 ' little bit extra so dont blink
'   FlexMsg  = 1   ' START POS ! frame 1
'   For h = 0 to 12
'     DMDDisplay(20,h) = DMDDisplay(0,h)
'   Next
' End If
End Sub

Sub DMD_Display_Timer
' DMD_Display.Enabled = False
  dim x

  For j = 0 to 8
    For x = 0 to 12
      DMDDisplay(j,x) = DMDDisplay(j+1,x)
    Next
  Next
  For x = 0 to 12
    DMDDisplay(9,x) = ""
  Next
  If DMDDisplay(0,0) <> "" Then
'   DMD_Display.Interval = DMDDisplay(0,8) * 1000
'   DMD_Display.Enabled = True
    FlexMsgTime = Frame + DMDDisplay(0,8) * 62.5 ' little bit extra
    FlexMsg = 1
    For h = 0 to 12
      DMDDisplay(20,h) = DMDDisplay(0,h)
    Next
  Else

  End If
End Sub

Sub DMDMessage(Message , DisplayTime)
  Exit Sub
  'debug.print Message
  For h = 0 to 9
    If DMDDisplay(h,0) = "" Then
      DMDDisplay(h,0) = Message
'     DMDDisplay(h,2) = Line2
'     DMDDisplay(h,3) = Line3
      DMDDisplay(h,1) = DisplayTime

    End If
  Next

  If DMDText.TimerEnabled = False Then
    DMDText.Text = Message : LEDText(Message)
    DMDText.TimerInterval = DisplayTime * 1000
    DMDText.TimerEnabled = True

'   FlexMsgTime = Frame + DisplayTime * 60 + 10
'   FlexMsg = 1
''    If line2 <> "" then FlexMsg = 2
''    If line3 <> "" then FlexMsg = 3
'   FlexText1=Message
''    FlexText2=Line2
''    FlexText3=line3
  End If


End Sub




Sub DMDText_Timer

  DMDText.TimerEnabled = False
  Exit Sub

  For j = 0 to 8
    DMDDisplay(j,0) = DMDDisplay(j+1,0)
    DMDDisplay(j,1) = DMDDisplay(j+1,1)
    DMDDisplay(j,2) = DMDDisplay(j+1,2)
    DMDDisplay(j,3) = DMDDisplay(j+1,3)
  Next
  DMDDisplay(9,0) = ""
  DMDDisplay(9,1) = ""
  DMDDisplay(9,2) = ""
  DMDDisplay(9,3) = ""

  If DMDDisplay(0,0) <> "" Then
    DMDText.Text = DMDDisplay(0,0) : LEDText(DMDDisplay(0,0))
    DMDText.TimerInterval = DMDDisplay(0,1) * 1000
    DMDText.TimerEnabled = True
  Else
    If StartGame = 0 Then
      ClearDMD
      If ScoreNumber = "" Then DMDText.Text = "Game Over" : LEDText("Game Over")
'     If ScoreNumber <> "" Then EnterHS
    Else
      Select Case CurrentMode
        Case 0
          If SkillShot = 0 Then
            If ScoreNumber = "" Then DMDText.Text = "Player " & CurrentPlayer : LEDText("Player " & CurrentPlayer)
            If ModesComplete(CurrentPlayer) = 9 Then DMDText.Text = "Shoot Vault for Super Jackpot" : LEDText("Shoot Vault for Super Jackpot")
  '         If ScoreNumber <> "" Then EnterHS
            If BallInLane = 1 Then SkillShotMessage.Enabled = True
          End If
        Case 1 : DMDText.Text = "Shoot Spongebob for Jackpot" : LEDText("Shoot Spongebob for Jackpot")
        Case 2 : DMDText.Text = "Shoot Left Ramp for Jackpot" : LEDText("Shoot Left Ramp for Jackpot")
        Case 3 : DMDText.Text = "Shoot Right Ramp for Jackpot" : LEDText("Shoot Right Ramp for Jackpot")
        Case 4 : DMDText.Text = "Shoot Mr Krab's Vault for Jackpot" : LEDText("Shoot Mr Krab's Vault for Jackpot")
        Case 5 : DMDText.Text = "BUMPERS SCORE 500 THOUSAND" : LEDText("BUMPERS SCORE 500 THOUSAND")
        Case 6 : DMDText.Text = "Shoot Left Orbit for Jackpot" : LEDText("Shoot Left Orbit for Jackpot")
        Case 7 : DMDText.Text = "Shoot Squidward for Jackpot" : LEDText("Shoot Squidward for Jackpot")
        Case 8 : DMDText.Text = "Shoot Patrick for Jackpot" : LEDText("Shoot Patrick for Jackpot")
        Case 9 : DMDText.Text = "Shoot Krabby Patty for Jackpot" : LEDText("Shoot Krabby Patty for Jackpot")
        Case 11 : DMDText.Text = "Shoot Pineapple for Super Jackpot" : LEDText("Shoot Pineapple for Super Jackpot")
      End Select
    End If
  End If
End Sub


Sub modetexttimer_timer

' Dim GameModeStr : GameModeStr = "NA:"

' If bTilted Then
'   Scorbit.SetGameMode(GameModeStr)
'   Exit Sub
' End If
  If StartGame = 0 Then
    ClearDMD
    If ScoreNumber = "" Then DMDText.Text = "Game Over" : LEDText("Game Over")
'     If ScoreNumber <> "" Then EnterHS
  Else
    Select Case CurrentMode
      Case 0

        If SkillShot = 0 Then

          If BallInLane = 1 Then SkillShotMessage.Enabled = True
          If ModesComplete(CurrentPlayer) = 9 Then
            line4 = 1 : msg4 = "VAULT FOR WIZARD MODE"
          Elseif NextMode(CurrentPlayer) > 0 Then
            line4 = 1
            If ModesComplete(CurrentPlayer) = 0 Then
              msg4 = "VAULT TO START FIRST MODE"
            Else
              msg4 = "VAULT TO START NEXT MODE"
            End If
            FlexDMD.Stage.GetLabel("TextSmalLine4").text = msg4
          End If
        End If

      case 1  : line4 = 1 : msg4 = "SHOOT SPONGEBOB FOR JACKPOT" ':     GameModeStr= "NA{yellow}:TIKILAND"
      case 2  : line4 = 1 : msg4 = "SHOOT LEFT RAMP FOR JACKPOT" ':     GameModeStr= "NA{yellow}:PLANKTON"
      case 3  : line4 = 1 : msg4 = "SHOOT RIGHT RAMP FOR JACKPOT" ':    GameModeStr= "NA{yellow}:MR KRAB"
      case 4  : line4 = 1 : msg4 = "SHOOT MR KRABS VAULT FOR JACKPOT" ':  GameModeStr= "NA{yellow}:RODEO"
      case 5  : line4 = 1 : msg4 = "SUPERBUMPERS IS 500.000 POINTS" ':  GameModeStr= "NA{yellow}:GOOFY"
      case 6  : line4 = 1 : msg4 = "SHOOT LEFT ORBIT FOR JACKPOT" ':    GameModeStr= "NA{yellow}:SCHOOL"
      case 7  : line4 = 1 : msg4 = "SHOOT SQUIDWARD FOR JACKPOT" ':     GameModeStr= "NA{yellow}:FORMULA"
      case 8  : line4 = 1 : msg4 = "SHOOT PATRICK FOR JACKPOT" ':       GameModeStr= "NA{yellow}:SPONGE"
      case 9  : line4 = 1 : msg4 = "SHOOT KRABBY PATTY FOR JACKPOT" ':  GameModeStr= "NA{yellow}:JELLYFISH"
      case 11 : line4 = 1 : msg4 = "ENTER PINEAPPLE FOR SUPERJACKPOT" ':  GameModeStr= GameModeStr + ";NA{red}:MULTIBALL"
'     case 20 : GameModeStr= "NA{green}:WIZARD 1"
'     case 21 : GameModeStr= "NA{green}:WIZARD 2"

    End Select
  End If

' If DoOrDieMode = 4 Then
'   GameModeStr= GameModeStr + ";NA{red}:DOD"
' End If

' If IsNull(Scorbit) OR Scorbit.bSessionActive=False then Exit Sub
' Scorbit.SetGameMode(GameModeStr)
End Sub


Sub LEDText(LEDMessage)
' If DesktopMode = False Or B2SonDesktop = 1 Then
'   LEDMessage2 = ""
'
'   If Len(LEDMessage) > 20 Then
'     LEDMessage2 = Mid(LEDMessage, InStrRev(LEDMessage," ",20) + 1, Len(LEDMessage))
'     LEDMessage = Mid(LEDMessage, 1, InStrRev(LEDMessage," ",20))
'   End If
'
'   For i = 41 to 80
'   Next
'
'   For i = 1 to Len(LEDMessage)
'   Next
'
'   If Len(LEDMessage2) > 0 Then
'     For i = 1 to Len(LEDMessage2)
'     Next
'   End If
' End If
End Sub

Sub ClearDMD
  Dim x
  For h = 0 to 11
    For x = 0 to 20
      DMDDisplay(h,x) = ""
    Next
  Next
DMD_ResetAll
End Sub

'**********************
' Music & Sound Routines
'**********************

Dim MusicPlaying : MusicPlaying = 0
Dim Dampenmusic : Dampenmusic = 1
Sub PlayTableMusic(SongSelect)

  MusicPlaying = SongSelect
  Select Case SongSelect
    Case 0
      If MultiBallReady(CurrentPlayer) = 1 Then
        PlaySound "Mode0d", -1 , RomSoundVolume * SongVolume * Dampenmusic , 0, 0, 0, True, 0, 0
      ElseIf SJPReady(CurrentPlayer) = 1 Then
        PlaySound "Sfx_SJPReady",-1,RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
      Else
      If SongNumber = 0 Then PlaySound "Mode0", -1, RomSoundVolume * SongVolume * Dampenmusic , 0, 0, 0, True, 0, 0
      If SongNumber = 1 Then PlaySound "Mode0a", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
      If SongNumber = 2 Then PlaySound "Mode0b", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
      If SongNumber = 3 Then PlaySound "Mode0c", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
      If SongNumber = 4 Then PlaySound "Mode0e", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
      If SongNumber = 5 Then PlaySound "Sfx_Greenhorn", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
      If SongNumber = 6 Then PlaySound "Sfx_Clownfish", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
      End If
    Case 1 : PlaySound "Mode1", -1, RomSoundVolume * SongVolume * Dampenmusic , 0, 0, 0, True, 0, 0
    Case 2 : PlaySound "Mode2", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
    Case 3 : PlaySound "Mode3", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
    Case 4 : PlaySound "Mode4", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
    Case 5 : PlaySound "Mode5", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
    Case 6 : PlaySound "Mode6", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
    Case 7 : PlaySound "Mode7", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
    Case 8 : PlaySound "Mode8", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
    Case 9 : PlaySound "Mode9", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
    Case 10 : PlaySound "Mode11", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
    Case 11
      If SJPReady(CurrentPlayer) = 1 Then
        PlaySound "Sfx_SJPReady",-1,RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
      Else
        PlaySound "Mode11", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
      End If
    Case 12: PlaySound "Opening", 0, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
    Case 13
      If MultiBallReady(CurrentPlayer) = 1 Then
        PlaySound "ShooterLaned", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
      ElseIf SJPReady(CurrentPlayer) = 1 Then
        PlaySound "Sfx_SJPReadyShooter",-1,RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
      Else
        If SongNumber = 0 Then PlaySound "ShooterLane", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
        If SongNumber = 1 Then PlaySound "ShooterLanea", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
        If SongNumber = 2 Then PlaySound "ShooterLaneb", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
        If SongNumber = 3 Then PlaySound "ShooterLanec", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
        If SongNumber = 4 Then PlaySound "ShooterLanee", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
        If SongNumber = 5 Then PlaySound "Sfx_GreenhornShooter", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
        If SongNumber = 6 Then PlaySound "Sfx_ClownfishShooter", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
      End If
    case 20,21 :

        songtimer.interval = 8000 : SongTimer.Enabled = True
        StopTableMusic
        Select Case int(rnd(1)*9) + 1
              Case 1 : PlaySound "Mode1", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
              Case 2 : PlaySound "rock", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
              Case 3 : PlaySound "Mode3", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
              Case 4 : PlaySound "Mode4", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
              Case 5 : PlaySound "Mode5", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
              Case 6 : PlaySound "Mode6", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
              Case 7 : PlaySound "Mode5", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
              Case 8 : PlaySound "Mode8", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
              Case 9 : PlaySound "Mode9", -1, RomSoundVolume * SongVolume * Dampenmusic, 0, 0, 0, True, 0, 0
        End Select
        DelayMusic = 0

  End Select

End Sub

Sub StopTableMusic
  MusicPlaying = -1
  Dampenmusic = 1
  Stopsound "rock"
  StopSound "Opening"
  StopSound "ShooterLane"
  StopSound "ShooterLanea"
  StopSound "ShooterLaneb"
  StopSound "ShooterLanec"
  StopSound "ShooterLaned"
  StopSound "ShooterLanee"
  StopSound "Mode0"
  StopSound "Mode0a"
  StopSound "Mode0b"
  StopSound "Mode0c"
  StopSound "Mode0d"
  StopSound "Mode0e"
  StopSound "IntroMode0"
  StopSound "IntroMode0a"
  StopSound "IntroMode0b"
  StopSound "IntroMode0c"
  StopSound "IntroMode0d"
  StopSound "IntroMode0e"
  StopSound "IntroMode1"
  StopSound "Mode1"
  StopSound "Mode2"
  StopSound "Mode3"
  StopSound "Mode4"
  StopSound "Mode5"
  StopSound "Mode6"
  StopSound "IntroMode6"
  StopSound "Mode7"
  StopSound "Mode8"
  StopSound "Mode9"
  StopSound "Mode11"
  StopSound "GameOverMusic"
  StopSound "HSMusic"
  StopSound "Sfx_ROrbit"
  StopSound "Sfx_ModeReady"
  StopSound "Sfx_Greenhorn"
  StopSound "Sfx_GreenhornShooter"
  StopSound "Sfx_GreenhornIntro"
  StopSound "Sfx_Clownfish"
  StopSound "Sfx_ClownfishShooter"
  StopSound "Sfx_ClownfishIntro"
  StopSound "Sfx_SJPReady"
  StopSound "Sfx_SJPReadyShooter"
  StopSound "Sfx_SJPReadyIntro"
  StopJackpotSfx
  StopLOrbitSfx
End Sub

Sub IntroMusic
  StopTableMusic
  Select Case CurrentMode
    Case 0
      If MultiBallReady(CurrentPlayer) = 1 Then
        PlaySound "IntroMode0d",0,RomSoundVolume * SongVolume : SongTimer.Interval = 1628 : SongTimer.Enabled = True
      ElseIf SJPReady(CurrentPlayer) = 1 Then
        PlaySound "IntroModeSJackpot",0,RomSoundVolume : SongTimer.Interval = 779 : SongTimer.Enabled = True
      Else
        If SongNumber = 0 Then PlaySound "IntroMode0",0,RomSoundVolume : SongTimer.Interval = 947 : SongTimer.Enabled = True
        If SongNumber = 1 Then PlaySound "IntroMode0a",0,RomSoundVolume : SongTimer.Interval = 1287 : SongTimer.Enabled = True
        If SongNumber = 2 Then PlaySound "IntroMode0b",0,RomSoundVolume : SongTimer.Interval = 1607 : SongTimer.Enabled = True
        If SongNumber = 3 Then PlaySound "IntroMode0c",0,RomSoundVolume : SongTimer.Interval = 2082 : SongTimer.Enabled = True
        If SongNumber = 4 Then PlaySound "IntroMode0e",0,RomSoundVolume : SongTimer.Interval = 951 : SongTimer.Enabled = True
        If SongNumber = 5 Then PlaySound "IntroModeGreenhorn",0,RomSoundVolume : SongTimer.Interval = 1525 : SongTimer.Enabled = True
        If SongNumber = 6 Then PlaySound "IntroModeClownfish",0,RomSoundVolume : SongTimer.Interval = 1608 : SongTimer.Enabled = True
      End If
    Case 1 : PlaySound "IntroMode1",0,RomSoundVolume : SongTimer.Interval = 17969 : SongTimer.Enabled = True
    Case 2 : PlaySound "IntroMode2",0,RomSoundVolume : SongTimer.Interval = 6411 : SongTimer.Enabled = True
    Case 3 : PlaySound "IntroMode3",0,RomSoundVolume : SongTimer.Interval = 5315 : SongTimer.Enabled = True
    Case 4 : PlaySound "IntroMode4",0,RomSoundVolume : SongTimer.Interval = 5966 : SongTimer.Enabled = True
    Case 5 : PlaySound "IntroMode5",0,RomSoundVolume : SongTimer.Interval = 9199 : SongTimer.Enabled = True
    Case 6 : PlaySound "IntroMode6",0,RomSoundVolume : SongTimer.Interval = 16247 : SongTimer.Enabled = True
    Case 7 : PlaySound "IntroMode7",0,RomSoundVolume : SongTimer.Interval = 6533 : SongTimer.Enabled = True
    Case 8 : PlaySound "IntroMode8",0,RomSoundVolume : SongTimer.Interval = 9877 : SongTimer.Enabled = True
    Case 9 : PlaySound "IntroMode9",0,RomSoundVolume : SongTimer.Interval = 8297 : SongTimer.Enabled = True
    Case 10 : PlaySound "IntroMode11",0,RomSoundVolume : SongTimer.Interval = 9413 : SongTimer.Enabled = True
    Case 11
      If SJPReady(CurrentPlayer) = 1 Then
        PlaySound "Sfx_SJPReadyIntro",0,RomSoundVolume : SongTimer.Interval = 779 : SongTimer.Enabled = True
      Else
        PlaySound "IntroMode11",0,RomSoundVolume : SongTimer.Interval = 9413 : SongTimer.Enabled = True
      End If
  End Select
End Sub


Sub EndingMusic
  If StartGame = 0 Then InsertSequence.Play SeqCircleOutOn,50,2
  If HSMusicPlaying.Enabled = True Then Exit Sub
  If StartGame = 0 Then PlaySound "GameOverMusic",0,RomSoundVolume : Exit Sub
  Select Case CurrentMode
    Case 0
      If MultiBallReady(CurrentPlayer) = 1 Then
        PlaySound "EndingMode0b",0,RomSoundVolume
      ElseIf SJPReady(CurrentPlayer) = 1 Then
        PlaySound "EndingModeSJackpot",0,RomSoundVolume
      Else
        If SongNumber = 0 Then PlaySound "EndingMode0",0,RomSoundVolume
        If SongNumber = 1 Then PlaySound "EndingMode0a",0,RomSoundVolume
        If SongNumber = 2 Then PlaySound "EndingMode0b",0,RomSoundVolume
        If SongNumber = 3 Then PlaySound "EndingMode0c",0,RomSoundVolume
        If SongNumber = 4 Then PlaySound "EndingMode0e",0,RomSoundVolume
        If SongNumber = 5 Then PlaySound "EndingModeGreenhorn",0,RomSoundVolume
        If SongNumber = 6 Then PlaySound "EndingModeClownfish",0,RomSoundVolume
        SongNumber = SongNumber +1 : If SongNumber > 6 Then SongNumber = 0
      End If
    Case 1 : PlaySound "EndingMode1",0,RomSoundVolume
    Case 2 : PlaySound "EndingMode2",0,RomSoundVolume
    Case 3 : PlaySound "EndingMode3",0,RomSoundVolume
    Case 4 : PlaySound "EndingMode4",0,RomSoundVolume
    Case 5 : PlaySound "EndingMode5",0,RomSoundVolume
    Case 6 : PlaySound "EndingMode6",0,RomSoundVolume
    Case 7 : PlaySound "EndingMode7",0,RomSoundVolume
    Case 8 : PlaySound "EndingMode8",0,RomSoundVolume
    Case 9 : PlaySound "EndingMode9",0,RomSoundVolume
    Case 10 : PlaySound "EndingMode11",0,RomSoundVolume
    Case 11
      If SJPReady(CurrentPlayer) = 1 Then
        PlaySound "Sfx_SJPReadyEnding",0,RomSoundVolume
      Else
        PlaySound "EndingMode11",0,RomSoundVolume
      End If
  End Select
End Sub

Sub SongTimer_Timer
  SongTimer.Enabled = False
  If DelayMusic = 0 Then Call PlayTableMusic(CurrentMode)
  If DelayMusic = 1 Then IntroMusic : DelayMusic = 0
End Sub

Sub HSMusicPlaying_Timer
  HSMusicPlaying.Enabled = False
End Sub



Sub drain_FX
  QuoteTimer.Enabled = False
  Select Case CurrentMode
    Case 0,4,5,9,10,11,20,21
      Select Case int(Rnd(1)*8)
        Case 0 : PlaySound "Sfx_Drain1",0,RomSoundVolume    : QuoteTimer.Interval = 2794 : QuoteTimer.Enabled = True
        Case 1 : PlaySound "Sfx_Drain3",0,RomSoundVolume    : QuoteTimer.Interval = 1216 : QuoteTimer.Enabled = True
        Case 2 : PlaySound "Sfx_Drain4",0,RomSoundVolume    : QuoteTimer.Interval = 1056 : QuoteTimer.Enabled = True
        Case 3 : PlaySound "Sfx_Drain5",0,RomSoundVolume    : QuoteTimer.Interval =  929 : QuoteTimer.Enabled = True
        Case 4 : PlaySound "Sfx_Drain6",0,RomSoundVolume    : QuoteTimer.Interval = 1153 : QuoteTimer.Enabled = True
        Case 5 : PlaySound "Sfx_Drain7",0,RomSoundVolume    : QuoteTimer.Interval = 1124 : QuoteTimer.Enabled = True
        Case 6 : PlaySound "Sfx_Drain8",0,RomSoundVolume    : QuoteTimer.Interval = 2132 : QuoteTimer.Enabled = True
        Case 7 : PlaySound "Sfx_Drain9",0,RomSoundVolume    : QuoteTimer.Interval =  985 : QuoteTimer.Enabled = True
      End Select

    Case 1 : PlaySound "Sfx_Sponge18",0,RomSoundVolume    : QuoteTimer.Interval = 2387 : QuoteTimer.Enabled = True    ' spnge
    Case 2 : PlaySound "Sfx_Drain2",0,RomSoundVolume    : QuoteTimer.Interval =  777 : QuoteTimer.Enabled = True    ' plankton
    Case 3 : PlaySound "Sfx_Squid12",0,RomSoundVolume   : QuoteTimer.Interval = 2179 : QuoteTimer.Enabled = True    ' sandy
    Case 6 : PlaySound "Sfx_Drain_Puff",0,RomSoundVolume  : QuoteTimer.Interval = 3083 : QuoteTimer.Enabled = True    ' puffschool
    Case 7 : PlaySound "Sfx_Drain_Sandy",0,RomSoundVolume : QuoteTimer.Interval = 1723 : QuoteTimer.Enabled = True    ' tikiland
    Case 8 : PlaySound "Sfx_Patrick12",0,RomSoundVolume   : QuoteTimer.Interval = 2287 : QuoteTimer.Enabled = True    ' patric
'                 DMDMessage1 "MR KRAB"     ,"COMPLETE"           ,FontSponge12bb ,55,16  ,55,16  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
'                 modedone4(CurrentPlayer) = 1
'                 DMDMessage1 "FORMULA"     ,"COMPLETE"           ,FontSponge12bb ,55,16  ,55,16  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
'                 modedone9(CurrentPlayer) = 1
'                 DMDMessage1 "JELLYFISH"     ,"COMPLETE"           ,FontSponge12bb ,55,16  ,55,16  ,"noslide2blink"  ,3  ,"BG001"  ,"novideo"
'                 modedone5(CurrentPlayer) = 1

  End Select
  Dampenmusic2 = 0.2
End Sub



Sub PlayQuote(QuoteNumber)
  If wizBiP > 1 Then Exit Sub

  Dim QuoteSelect
  If CurrentMode = 2 Then QuoteNumber = 5
  Select Case QuoteNumber
    Case 1
      If QuoteTimer.Enabled = True Then SlingSfx
      If QuoteTimer.Enabled = False Then
        QuoteSelect = int(rnd*14)
        Select Case QuoteSelect
          Case 0 : PlaySound "Sfx_Squid1",0,RomSoundVolume : QuoteTimer.Interval = 5060 : QuoteTimer.Enabled = True
          Case 1 : PlaySound "Sfx_Squid2",0,RomSoundVolume : QuoteTimer.Interval = 4265 : QuoteTimer.Enabled = True
          Case 2 : PlaySound "Sfx_Squid3",0,RomSoundVolume : QuoteTimer.Interval = 2547 : QuoteTimer.Enabled = True
          Case 3 : PlaySound "Sfx_Squid4",0,RomSoundVolume : QuoteTimer.Interval = 1892 : QuoteTimer.Enabled = True
          Case 4 : PlaySound "Sfx_Squid5",0,RomSoundVolume : QuoteTimer.Interval = 2100 : QuoteTimer.Enabled = True
          Case 5 : PlaySound "Sfx_Squid6",0,RomSoundVolume : QuoteTimer.Interval = 3337 : QuoteTimer.Enabled = True
          Case 6 : PlaySound "Sfx_Squid7",0,RomSoundVolume : QuoteTimer.Interval = 1464 : QuoteTimer.Enabled = True
          Case 7 : PlaySound "Sfx_Squid8",0,RomSoundVolume : QuoteTimer.Interval = 3253 : QuoteTimer.Enabled = True
          Case 8 : PlaySound "Sfx_Squid9",0,RomSoundVolume : QuoteTimer.Interval = 2302 : QuoteTimer.Enabled = True
          Case 9 : PlaySound "Sfx_Squid10",0,RomSoundVolume : QuoteTimer.Interval = 2692 : QuoteTimer.Enabled = True
          Case 10 : PlaySound "Sfx_Squid11",0,RomSoundVolume : QuoteTimer.Interval = 2695 : QuoteTimer.Enabled = True
          Case 11 : PlaySound "Sfx_Squid13",0,RomSoundVolume : QuoteTimer.Interval = 2595 : QuoteTimer.Enabled = True
          Case 12 : PlaySound "Sfx_Squid14",0,RomSoundVolume : QuoteTimer.Interval = 2204 : QuoteTimer.Enabled = True
          Case 13 : PlaySound "Sfx_Squid15",0,RomSoundVolume : QuoteTimer.Interval = 1730 : QuoteTimer.Enabled = True


        End Select
      End If
    Case 2
      If QuoteTimer.Enabled = False Then
        QuoteSelect = int(rnd*10)
        Select Case QuoteSelect
          Case 0 : PlaySound "Sfx_Krabs1",0,RomSoundVolume : QuoteTimer.Interval = 4322 : QuoteTimer.Enabled = True
          Case 1 : PlaySound "Sfx_Krabs2",0,RomSoundVolume : QuoteTimer.Interval = 2338 : QuoteTimer.Enabled = True
          Case 2 : PlaySound "Sfx_Krabs3",0,RomSoundVolume : QuoteTimer.Interval = 4970 : QuoteTimer.Enabled = True
          Case 3 : PlaySound "Sfx_Krabs4",0,RomSoundVolume : QuoteTimer.Interval = 3282 : QuoteTimer.Enabled = True
          Case 4 : PlaySound "Sfx_Krabs5",0,RomSoundVolume : QuoteTimer.Interval = 1943 : QuoteTimer.Enabled = True
          Case 5 : PlaySound "Sfx_Krabs6",0,RomSoundVolume : QuoteTimer.Interval = 3601 : QuoteTimer.Enabled = True
          Case 6 : PlaySound "Sfx_Krabs7",0,RomSoundVolume : QuoteTimer.Interval = 4690 : QuoteTimer.Enabled = True
          Case 7 : PlaySound "Sfx_Krabs8",0,RomSoundVolume : QuoteTimer.Interval = 4426 : QuoteTimer.Enabled = True
          Case 8 : PlaySound "Sfx_Krabs9",0,RomSoundVolume : QuoteTimer.Interval = 3542 : QuoteTimer.Enabled = True
          Case 9 : PlaySound "Sfx_Krabs10",0,RomSoundVolume : QuoteTimer.Interval = 2270 : QuoteTimer.Enabled = True
        End Select
      End If
    Case 3
      If QuoteTimer.Enabled = False Then
        QuoteSelect = int(rnd*20)
        Select Case QuoteSelect
          Case 0 : PlaySound "Sfx_Sponge1",0,RomSoundVolume : QuoteTimer.Interval = 3947 : QuoteTimer.Enabled = True
          Case 1 : PlaySound "Sfx_Sponge2",0,RomSoundVolume : QuoteTimer.Interval = 3098 : QuoteTimer.Enabled = True
          Case 2 : PlaySound "Sfx_Sponge3",0,RomSoundVolume : QuoteTimer.Interval = 3144 : QuoteTimer.Enabled = True
          Case 3 : PlaySound "Sfx_Sponge4",0,RomSoundVolume : QuoteTimer.Interval = 3046 : QuoteTimer.Enabled = True
          Case 4 : PlaySound "Sfx_Sponge5",0,RomSoundVolume : QuoteTimer.Interval = 3170 : QuoteTimer.Enabled = True
          Case 5 : PlaySound "Sfx_Sponge6",0,RomSoundVolume : QuoteTimer.Interval = 4823 : QuoteTimer.Enabled = True
          Case 6 : PlaySound "Sfx_Sponge7",0,RomSoundVolume : QuoteTimer.Interval = 7004 : QuoteTimer.Enabled = True
          Case 7 : PlaySound "Sfx_Sponge8",0,RomSoundVolume : QuoteTimer.Interval = 2804 : QuoteTimer.Enabled = True
          Case 8 : PlaySound "Sfx_Sponge9",0,RomSoundVolume : QuoteTimer.Interval = 3932 : QuoteTimer.Enabled = True
          Case 9 : PlaySound "Sfx_Sponge10",0,RomSoundVolume : QuoteTimer.Interval = 3462 : QuoteTimer.Enabled = True
          Case 10 : PlaySound "Sfx_Sponge11",0,RomSoundVolume : QuoteTimer.Interval = 3136 : QuoteTimer.Enabled = True
          Case 11 : PlaySound "Sfx_Sponge12",0,RomSoundVolume : QuoteTimer.Interval = 5201 : QuoteTimer.Enabled = True
          Case 12 : PlaySound "Sfx_Sponge13",0,RomSoundVolume : QuoteTimer.Interval = 6203 : QuoteTimer.Enabled = True
          Case 13 : PlaySound "Sfx_Sponge14",0,RomSoundVolume : QuoteTimer.Interval = 3639 : QuoteTimer.Enabled = True
          Case 14 : PlaySound "Sfx_Sponge15",0,RomSoundVolume : QuoteTimer.Interval = 5010 : QuoteTimer.Enabled = True
          Case 15 : PlaySound "Sfx_Sponge16",0,RomSoundVolume : QuoteTimer.Interval = 5172 : QuoteTimer.Enabled = True
          Case 16 : PlaySound "Sfx_Sponge17",0,RomSoundVolume : QuoteTimer.Interval = 2475 : QuoteTimer.Enabled = True
          Case 17 : PlaySound "Sfx_Sponge19",0,RomSoundVolume : QuoteTimer.Interval = 2359 : QuoteTimer.Enabled = True
          Case 18 : PlaySound "Sfx_Sponge20",0,RomSoundVolume : QuoteTimer.Interval = 1955 : QuoteTimer.Enabled = True
          Case 19 : PlaySound "Sfx_Sponge21",0,RomSoundVolume : QuoteTimer.Interval = 2808 : QuoteTimer.Enabled = True


        End Select
      End If
    Case 4
      If QuoteTimer.Enabled = True Then SlingSfx
      If QuoteTimer.Enabled = False Then
        QuoteSelect = int(rnd*16)
        Select Case QuoteSelect
          Case 0 : PlaySound "Sfx_Patrick1",0,RomSoundVolume : QuoteTimer.Interval = 4319 : QuoteTimer.Enabled = True
          Case 1 : PlaySound "Sfx_Patrick2",0,RomSoundVolume : QuoteTimer.Interval = 2843 : QuoteTimer.Enabled = True
          Case 2 : PlaySound "Sfx_Patrick3",0,RomSoundVolume : QuoteTimer.Interval = 3714 : QuoteTimer.Enabled = True
          Case 3 : PlaySound "Sfx_Patrick4",0,RomSoundVolume : QuoteTimer.Interval = 3946 : QuoteTimer.Enabled = True
          Case 4 : PlaySound "Sfx_Patrick5",0,RomSoundVolume : QuoteTimer.Interval = 2635 : QuoteTimer.Enabled = True
          Case 5 : PlaySound "Sfx_Patrick6",0,RomSoundVolume : QuoteTimer.Interval = 5504 : QuoteTimer.Enabled = True
          Case 6 : PlaySound "Sfx_Patrick7",0,RomSoundVolume : QuoteTimer.Interval = 1625 : QuoteTimer.Enabled = True
          Case 7 : PlaySound "Sfx_Patrick8",0,RomSoundVolume : QuoteTimer.Interval = 5052 : QuoteTimer.Enabled = True
          Case 8 : PlaySound "Sfx_Patrick9",0,RomSoundVolume : QuoteTimer.Interval = 3648 : QuoteTimer.Enabled = True
          Case 9 : PlaySound "Sfx_Patrick10",0,RomSoundVolume : QuoteTimer.Interval = 1784 : QuoteTimer.Enabled = True
          Case 10 : PlaySound "Sfx_Patrick11",0,RomSoundVolume : QuoteTimer.Interval = 1466 : QuoteTimer.Enabled = True
          Case 11 : PlaySound "Sfx_Patrick13",0,RomSoundVolume : QuoteTimer.Interval = 1790 : QuoteTimer.Enabled = True
          Case 12 : PlaySound "Sfx_Patrick14",0,RomSoundVolume : QuoteTimer.Interval = 2470 : QuoteTimer.Enabled = True
          Case 13 : PlaySound "Sfx_Patrick15",0,RomSoundVolume : QuoteTimer.Interval = 3383 : QuoteTimer.Enabled = True
          Case 14 : PlaySound "Sfx_Patrick16",0,RomSoundVolume : QuoteTimer.Interval = 1449 : QuoteTimer.Enabled = True
          Case 15 : PlaySound "Sfx_Patrick17",0,RomSoundVolume : QuoteTimer.Interval = 1632 : QuoteTimer.Enabled = True


        End Select
      End If
    Case 5
      If QuoteTimer.Enabled = False Then
        QuoteSelect = int(rnd*10)
        Select Case QuoteSelect
          Case 0 : PlaySound "Sfx_Plankton1",0,RomSoundVolume : QuoteTimer.Interval = 3944 : QuoteTimer.Enabled = True
          Case 1 : PlaySound "Sfx_Plankton2",0,RomSoundVolume : QuoteTimer.Interval = 2959 : QuoteTimer.Enabled = True
          Case 2 : PlaySound "Sfx_Plankton3",0,RomSoundVolume : QuoteTimer.Interval = 3887 : QuoteTimer.Enabled = True
          Case 3 : PlaySound "Sfx_Plankton4",0,RomSoundVolume : QuoteTimer.Interval = 1463 : QuoteTimer.Enabled = True
          Case 4 : PlaySound "Sfx_Plankton5",0,RomSoundVolume : QuoteTimer.Interval = 2095 : QuoteTimer.Enabled = True
          Case 5 : PlaySound "Sfx_Plankton6",0,RomSoundVolume : QuoteTimer.Interval = 6339 : QuoteTimer.Enabled = True
          Case 6 : PlaySound "Sfx_Plankton7",0,RomSoundVolume : QuoteTimer.Interval = 2156 : QuoteTimer.Enabled = True
          Case 7 : PlaySound "Sfx_Plankton8",0,RomSoundVolume : QuoteTimer.Interval = 3437 : QuoteTimer.Enabled = True
          Case 8 : PlaySound "Sfx_Plankton9",0,RomSoundVolume : QuoteTimer.Interval = 4067 : QuoteTimer.Enabled = True
          Case 9 : PlaySound "Sfx_Plankton10",0,RomSoundVolume : QuoteTimer.Interval = 2725 : QuoteTimer.Enabled = True
        End Select
      End If
  End Select
' if Dampenmusic = "Mode1" Then PlaySound "Mode1", -1, RomSoundVolume * SongVolume/2, 0, 0, 0, True, 0, 0: debug.print"low"

  Dampenmusic2 = 0.2 ': If MusicPlaying <> - 1 Then PlayTableMusic(MusicPlaying) : debug.print"low"
End Sub

Sub SlingSfx
  Dim SfxSelect
  SfxSelect = int(rnd*6)
  StopSfx
  Select Case SfxSelect
    Case 0 : PlaySound "Sfx_Sling1",0,RomSoundVolume
    Case 1 : PlaySound "Sfx_Sling2",0,RomSoundVolume
    Case 2 : PlaySound "Sfx_Sling3",0,RomSoundVolume
    Case 3 : PlaySound "Sfx_Sling4",0,RomSoundVolume
    Case 4 : PlaySound "Sfx_Sling5",0,RomSoundVolume
    Case 5 : PlaySound "Sfx_Sling6",0,RomSoundVolume
  End Select
End Sub

Sub TargetSfx
  Dim SfxSelect
  SfxSelect = int(rnd*7)
  StopSfx
  Select Case SfxSelect
    Case 0 : PlaySound "Sfx_Target1",0,RomSoundVolume
    Case 1 : PlaySound "Sfx_Target2",0,RomSoundVolume
    Case 2 : PlaySound "Sfx_Target3",0,RomSoundVolume
    Case 3 : PlaySound "Sfx_Target4",0,RomSoundVolume
    Case 4 : PlaySound "Sfx_Target5",0,RomSoundVolume
    Case 5 : PlaySound "Sfx_Target6",0,RomSoundVolume
    Case 6 : PlaySound "Sfx_Target7",0,RomSoundVolume
  End Select
End Sub

Sub BumperSfx
  Dim SfxSelect
  SfxSelect = int(rnd*3)
  StopSound "Sfx_Bumper1" : StopSound "Sfx_Bumper2" : StopSound "Sfx_Bumper3"
  Select Case SfxSelect
    Case 0 : PlaySound "Sfx_Bumper1",0,RomSoundVolume
    Case 1 : PlaySound "Sfx_Bumper2",0,RomSoundVolume
    Case 2 : PlaySound "Sfx_Bumper3",0,RomSoundVolume
  End Select
End Sub

Sub LOrbitSfx
  Dim SfxSelect
  SfxSelect = int(rnd*4)
  Select Case SfxSelect
    Case 0 : PlaySound "Sfx_LOrbit",0,RomSoundVolume
    Case 1 : PlaySound "Sfx_LOrbit2",0,RomSoundVolume
    Case 2 : PlaySound "Sfx_LOrbit3",0,RomSoundVolume
    Case 3 : PlaySound "Sfx_LOrbit4",0,RomSoundVolume
  End Select
End Sub

Sub LRampSfx
  Dim SfxSelect
  SfxSelect = int(rnd*2)
  Select Case SfxSelect
    Case 0 : PlaySound "Sfx_LRamp1",0,RomSoundVolume
    Case 1 : PlaySound "Sfx_LRamp2",0,RomSoundVolume
  End Select
End Sub

Sub RRampSfx
  Dim SfxSelect
  SfxSelect = int(rnd*2)
  Select Case SfxSelect
    Case 0 : PlaySound "Sfx_RRamp1",0,RomSoundVolume
    Case 1 : PlaySound "Sfx_RRamp2",0,RomSoundVolume
  End Select
End Sub

Sub JackpotSfx
  Dim SfxSelect
  SfxSelect = int(rnd*4)
  StopJackpotSfx
  Select Case SfxSelect
    Case 0 : PlaySound "Sfx_Jackpot",0,RomSoundVolume
    Case 1 : PlaySound "Sfx_Jackpot2",0,RomSoundVolume
    Case 2 : PlaySound "Sfx_Jackpot3",0,RomSoundVolume
    Case 3 : PlaySound "Sfx_Jackpot4",0,RomSoundVolume
  End Select
End Sub

Sub SkillShotSfx
  Select Case SkillSelect
    Case 0 : PlaySound "Sfx_Skill7",0,RomSoundVolume *.15
    Case 1 : PlaySound "Sfx_Skill9",0,RomSoundVolume *.15
    Case 2 : PlaySound "Sfx_Skill7",0,RomSoundVolume *.15
    Case 3 : PlaySound "Sfx_Skill4",0,RomSoundVolume *.15
    Case 4 : PlaySound "Sfx_Skill0",0,RomSoundVolume *.15
    Case 5 : PlaySound "Sfx_Skill4",0,RomSoundVolume *.15
    Case 6 : PlaySound "Sfx_Skill7",0,RomSoundVolume *.15
    Case 7 : PlaySound "Sfx_Skill9",0,RomSoundVolume *.15
    Case 8 : PlaySound "Sfx_Skill7",0,RomSoundVolume *.15
    Case 9 : PlaySound "Sfx_Skill4",0,RomSoundVolume *.15
    Case 10 : PlaySound "Sfx_Skill4",0,RomSoundVolume *.15
    Case 11 : PlaySound "Sfx_Skill4",0,RomSoundVolume *.15
  End Select
  SkillSelect = SkillSelect + 1
  If SkillSelect > 11 Then SkillSelect = 0
End Sub

Sub RoundBonusSfx
  Dim SfxSelect
  SfxSelect = int(rnd*7)
  Select Case SfxSelect
    Case 0 : PlaySound "Sfx_Bonus1",0,RomSoundVolume
    Case 1 : PlaySound "Sfx_Bonus2",0,RomSoundVolume
    Case 2 : PlaySound "Sfx_Bonus3",0,RomSoundVolume
    Case 3 : PlaySound "Sfx_Bonus4",0,RomSoundVolume
    Case 4 : PlaySound "Sfx_Bonus5",0,RomSoundVolume
    Case 5 : PlaySound "Sfx_Bonus6",0,RomSoundVolume
    Case 6 : PlaySound "Sfx_Bonus7",0,RomSoundVolume
  End Select
End Sub

Sub StopQuotes
  DampenMusic2 = 1
  StopSound "Sfx_Squid1"
  StopSound "Sfx_Squid2"
  StopSound "Sfx_Squid3"
  StopSound "Sfx_Squid4"
  StopSound "Sfx_Squid5"
  StopSound "Sfx_Squid6"
  StopSound "Sfx_Squid7"
  StopSound "Sfx_Squid8"
  StopSound "Sfx_Squid9"
  StopSound "Sfx_Squid10"
  StopSound "Sfx_Squid11"
    StopSound "Sfx_Krabs1"
  StopSound "Sfx_Krabs2"
  StopSound "Sfx_Krabs3"
  StopSound "Sfx_Krabs4"
  StopSound "Sfx_Krabs5"
  StopSound "Sfx_Krabs6"
  StopSound "Sfx_Krabs7"
  StopSound "Sfx_Krabs8"
  StopSound "Sfx_Krabs9"
  StopSound "Sfx_Krabs10"
  StopSound "Sfx_Sponge1"
  StopSound "Sfx_Sponge2"
  StopSound "Sfx_Sponge3"
  StopSound "Sfx_Sponge4"
  StopSound "Sfx_Sponge5"
  StopSound "Sfx_Sponge6"
  StopSound "Sfx_Sponge7"
  StopSound "Sfx_Sponge8"
  StopSound "Sfx_Sponge9"
  StopSound "Sfx_Sponge10"
  StopSound "Sfx_Sponge11"
  StopSound "Sfx_Sponge12"
  StopSound "Sfx_Sponge13"
  StopSound "Sfx_Sponge14"
  StopSound "Sfx_Sponge15"
  StopSound "Sfx_Sponge16"
  StopSound "Sfx_Patrick1"
  StopSound "Sfx_Patrick2"
  StopSound "Sfx_Patrick3"
  StopSound "Sfx_Patrick4"
  StopSound "Sfx_Patrick5"
  StopSound "Sfx_Patrick6"
  StopSound "Sfx_Patrick7"
  StopSound "Sfx_Patrick8"
  StopSound "Sfx_Patrick9"
  StopSound "Sfx_Patrick10"
  QuoteTimer.Enabled = False
End Sub

Sub StopSfx
  StopSound "Sfx_Sling1"
  StopSound "Sfx_Sling2"
  StopSound "Sfx_Sling3"
  StopSound "Sfx_Sling4"
  StopSound "Sfx_Sling5"
  StopSound "Sfx_Sling6"
  StopSound "Sfx_Target1"
  StopSound "Sfx_Target2"
  StopSound "Sfx_Target3"
  StopSound "Sfx_Target4"
  StopSound "Sfx_Target5"
  StopSound "Sfx_Target6"
  StopSound "Sfx_Target7"
End Sub

Sub StopLOrbitSfx
  StopSound "Sfx_LOrbit"
  StopSound "Sfx_LOrbit2"
  StopSound "Sfx_LOrbit3"
  StopSound "Sfx_LOrbit4"
End Sub

Sub StopLRampSfx
  StopSound "Sfx_LRamp1"
  StopSound "Sfx_LRamp2"
End Sub

Sub StopRRampSfx
  StopSound "Sfx_RRamp1"
  StopSound "Sfx_RRamp2"
End Sub

Sub StopJackpotSfx
  StopSound "Sfx_Jackpot"
  StopSound "Sfx_Jackpot2"
  StopSound "Sfx_Jackpot3"
  StopSound "Sfx_Jackpot4"
End Sub

Dim Doublequote : Doublequote = 0
Sub QuoteTimer_Timer

  Dampenmusic2 = 1 ': If MusicPlaying <> - 1 Then PlayTableMusic(MusicPlaying) : debug.print"high"

' if Dampenmusic = "Mode1" Then PlaySound "Mode1", -1, RomSoundVolume * SongVolume/2, 0, 0, 0, True, 0, 0:: debug.print"high"
  If CurrentMode = 11 Or wizBiP = 2 And Doublequote = 0 Then Doublequote = 1 :QuoteTimer.Interval = 7000 : Exit Sub
  QuoteTimer.Enabled = False
  Doublequote = 0
End Sub

Sub RoundEndTimer_Timer
  RoundEndTimer.Enabled = False
  If soundcount < 3 Then RoundBonusSfx : soundcount = soundcount + 1
  If soundcount < 3 Then RoundEndTimer.Enabled = True
  If soundcount = 3 Then soundcount = 0
End Sub

Sub FanfareSfx
  FanfareSelect = FanfareSelect + 1
  If FanfareSelect > 4 Then FanfareSelect = 0
  StopFanfareSfx
  Select Case FanfareSelect
    Case 0 : PlaySound "Sfx_Cheer1",0,RomSoundVolume
    Case 1 : PlaySound "Sfx_Cheer4",0,RomSoundVolume
    Case 2 : PlaySound "Sfx_Cheer2",0,RomSoundVolume
    Case 3 : PlaySound "Sfx_Cheer5",0,RomSoundVolume
    Case 4 : PlaySound "Sfx_Cheer3",0,RomSoundVolume
  End Select
End Sub

Sub StopFanfareSfx
  StopSound "Sfx_Cheer1"
  StopSound "Sfx_Cheer2"
  StopSound "Sfx_Cheer3"
  StopSound "Sfx_Cheer4"
  StopSound "Sfx_Cheer5"
End Sub

Sub FanfareTimer_Timer
  FanfareTimer.Enabled = False
' TiltGame = 0
  bTilted = False

  ShakeSponge = 3
  SpongeTimer.Interval = 250
  SpongeTimer.Enabled = True
  ShakePatty=1   : PlaySoundAt "AlienShake2", alien6_BM_Lit_Room
  ShakePatrick=1   : PlaySoundAt "AlienShake1", alien14_BM_Lit_Room
  ShakeSquidward=1 :PlaySoundAt "AlienShake1", alien5_BM_Lit_Room
  ToyShake

End Sub

'**********************
' High Scores
'**********************

Sub ResetHighscores
  HighScores(0,0)=200000000
  HighScores(0,1)="PLK"
  HighScores(1,0)=150000000
  HighScores(1,1)="SPB"
  HighScores(2,0)=100000000
  HighScores(2,1)="PAT"
  HighScores(3,0)=70000000
  HighScores(3,1)="SQD"
  HighScores(4,0)=50000000
  HighScores(4,1)="KRB"
  HighScores(5,0)=30000000
  HighScores(5,1)="SDY"
End Sub

Sub SaveHighScores
  Dim FileObj, ScoreFile
  Dim P1,P2,P3,P4,P5,P6

  'Save HS Text Files
  Set FileObj=CreateObject("Scripting.FileSystemObject")

  If Not FileObj.FolderExists(UserDirectory) then
        ResetHighscores
    Exit Sub
  End if

  P1 = "1.:" & HighScores(0,1) & ":" & HighScores(0,0)
  P2 = "2.:" & HighScores(1,1) & ":" & HighScores(1,0)
  P3 = "3.:" & HighScores(2,1) & ":" & HighScores(2,0)
  P4 = "4.:" & HighScores(3,1) & ":" & HighScores(3,0)
  P5 = "5.:" & HighScores(4,1) & ":" & HighScores(4,0)
  P6 = "6.:" & HighScores(5,1) & ":" & HighScores(5,0)

  Set ScoreFile = FileObj.CreateTextFile(UserDirectory & cGameName & "_hiscores.txt",True)
    ScoreFile.WriteLine "High Scores 2.1"
    ScoreFile.WriteLine P1
    ScoreFile.WriteLine P2
    ScoreFile.WriteLine P3
    ScoreFile.WriteLine P4
    ScoreFile.WriteLine P5
    ScoreFile.WriteLine P6
    ScoreFile.Close

  Set ScoreFile=Nothing

  Dim dd, mm, yy, hh, nn, ss
  Dim datevalue, timevalue, dtsnow, dtsvalue

  'Store DateTimeStamp once.
  dtsnow = Now()

  'Individual date components
  dd = Right("00" & Day(dtsnow), 2)
  mm = Right("00" & Month(dtsnow), 2)
  yy = Year(dtsnow)
  hh = Right("00" & Hour(dtsnow), 2)
  nn = Right("00" & Minute(dtsnow), 2)
  ss = Right("00" & Second(dtsnow), 2)

  'Build the date string in the format mm/dd/yyyy
  datevalue = mm & "/" & dd & "/" & yy

  'Build the time string in the format hh:mm:ss
  timevalue = hh & ":" & nn & ":" & ss

  'Concatenate both together to build the timestamp mm/dd/yyyy hh:mm:ss
  dtsvalue = datevalue & " " & timevalue

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "log.txt",8,False)
    ScoreFile.WriteLine dtsvalue & ", Changed, " & cGameName & ".nv"
    ScoreFile.close

  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub

Sub LoadHighScores
  Dim FileObj, ScoreFile
  Dim P1,P2,P3,P4,P5
  Dim lineData, parts
  Dim tmp
  ResetHighscores
  ' Save High Score Text Files
  Set FileObj = CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then Exit Sub
    ' Read High Scores
    If FileObj.FileExists(UserDirectory & cGameName & "_hiscores.txt") Then
        Set ScoreFile = FileObj.OpenTextFile(UserDirectory & cGameName & "_hiscores.txt", 1)
    Linedata = ScoreFile.ReadLine ' Skip the first line ("High Scores")
    If lineData <> "High Scores 2.1" Then
      ResetHighscores
          ScoreFile.Close
      Set ScoreFile = Nothing
      Set FileObj = Nothing
      Exit Sub
    End If

        Dim i
    For i = 0 to 5
      If ScoreFile.AtEndOfStream Then Exit Sub
            lineData = ScoreFile.ReadLine
            parts = Split(lineData, ":")

            If UBound(parts) >= 2 Then
        HighScores(i, 1) = parts(1) ' Name
        If len(HighScores(i, 1)) <> 3 Then HighScores(i, 1) = Left(HighScores(i, 1) & "   ",3)
        HighScores(i, 0) = Clng( parts(2)) ' Score
            End If
    Next
        ScoreFile.Close
    End If
    ' Clean up
    Set ScoreFile = Nothing
    Set FileObj = Nothing
End Sub

Sub CheckHighScores
  If DoOrDieMode = 4 Then
    DoOrDieScore = TotalScore(CurrentPlayer)
    If TotalScore(CurrentPlayer) > HighScores(5,0) Then
      HighScores(5,0) = TotalScore(CurrentPlayer)
      ScoreNumber = 5
      EnterHighscore = 1
    End If
    Exit Sub
  End If


    Dim i, j
    ScoreNumber = ""

    ' Loop through the existing high scores
    For i = 0 To 4
        If TotalScore(CurrentPlayer) > HighScores(i, 0) Then
            ' Shift down lower scores to make room for the new score
            For j = 4 To i + 1 Step -1
                HighScores(j, 0) = HighScores(j - 1, 0)
                HighScores(j, 1) = HighScores(j - 1, 1)
            Next

            ' Insert the new score and name
            HighScores(i, 0) = TotalScore(CurrentPlayer)

            ' Capture position of new score
            ScoreNumber = i
      EnterHighscore = 1
            ' Exit the loop as we've found our spot
            Exit For
        End If
    Next
End Sub


Sub EnterHS
' debug.pring "EnterHS called"
  Exit Sub

  HSDisplay = 0 : DisplayHS.Enabled = False
  If HSDigit = 0 Then DMDText.Text = "HIGH SCORE " & ScoreNumber+1 & " -- ENTER INITIALS WITH FLIPPERS" : LEDText("HIGH SCORE " & ScoreNumber+1 & " -- ENTER INITIALS") : PlaySound "HSMusic",0,RomSoundVolume : HSMusicPlaying.Enabled = True
  If HSDigit = 1 Then DMDText.Text = Chr(HSChr) : LEDText(Chr(HSChr)) : HSI1 = Chr(HSChr)
  If HSDigit = 2 Then DMDText.Text = HSI1 & Chr(HSChr) : LEDText(HSI1 & Chr(HSChr)) : HSI2 = Chr(HSChr)
  If HSDigit = 3 Then DMDText.Text = HSI1 & HSI2 & Chr(HSChr) : LEDText(HSI1 & HSI2 & Chr(HSChr)) : HSI3 = Chr(HSChr) : HighScores(ScoreNumber, 1) = HSI1 & HSI2 & HSI3
  If HSDigit = 4 Then
    SaveHighScores
    DMDText.Text = "NUMBER " & ScoreNumber+1 & " SCORE -- CONGRATULATIONS" : LEDText("NUMBER " & ScoreNumber+1 & " SCORE -- CONGRATULATIONS")
    HSDigit = 0
    ScoreNumber = ""

    If StartGame = 1 Then ClearDMD : StartNewPlayer

    If StartGame = 0 Then
      HSDisplay = 1 : DisplayHS.Interval = 5000 : DisplayHS.Enabled = True
    End If
  End If
End Sub

Sub DisplayHS_Timer
  DisplayHS.Enabled = False
  DisplayHS.Interval = 2000
  If HSDisplay = 1 Then
    If HSDisplay2 = 0 Then DMDText.Text = "Bikini Bottom Pinball: Press Start" : LEDText("Bikini Bottom Pinball Press Start")
    If HSDisplay2 = 1 Then DMDText.Text = "HS 1- " & FormatNumber(HighScores(0,0),0,,,-1) & " " & HighScores(0,1) : LEDText("HS 1- " & FormatNumber(HighScores(0,0),0,,,-1) & "  " & HighScores(0,1))
    If HSDisplay2 = 2 Then DMDText.Text = "HS 2- " & FormatNumber(HighScores(1,0),0,,,-1) & " " & HighScores(1,1) : LEDText("HS 2- " & FormatNumber(HighScores(1,0),0,,,-1) & "  " & HighScores(1,1))
    If HSDisplay2 = 3 Then DMDText.Text = "HS 3- " & FormatNumber(HighScores(2,0),0,,,-1) & " " & HighScores(2,1) : LEDText("HS 3- " & FormatNumber(HighScores(2,0),0,,,-1) & "  " & HighScores(2,1))
    If HSDisplay2 = 4 Then DMDText.Text = "HS 4- " & FormatNumber(HighScores(3,0),0,,,-1) & " " & HighScores(3,1) : LEDText("HS 4- " & FormatNumber(HighScores(3,0),0,,,-1) & "  " & HighScores(3,1))
    If HSDisplay2 = 5 Then DMDText.Text = "HS 5- " & FormatNumber(HighScores(4,0),0,,,-1) & " " & HighScores(4,1) : LEDText("HS 5- " & FormatNumber(HighScores(4,0),0,,,-1) & "  " & HighScores(4,1))
    HSDisplay2 = HSDisplay2 + 1
    If HSDisplay2 = 6 Then HSDisplay2 = 0
    DisplayHS.Enabled = True
  End If
End Sub

Sub VolumeTimer_Timer
  VolumeTimer.Enabled = False
  DisplayHS_Timer
End Sub




'********************************************
' Tilt
'********************************************

'Sub CheckTilt
' If StartGame = 1 And TiltGame = 0 And BallInLane = 0 And Drain.TimerEnabled = False Then
'   If LTiltPower > 25 Or RTiltPower > 25 Or CTiltPower > 25 Then
'     WarningCount = WarningCount + 1
'     If WarningCount = 1 Then ClearDMD : Call DMDMessage("WARNING",1) : Playsound "Sfx_Warning", 0, RomSoundVolume
'   End If
'   If WarningCount > 1 Then ClearDMD : Call DMDMessage("TILT",2) : TiltGame = 1 : Playtime = 100 : StopTableMusic : PlaySound "Sfx_Tilt",0,RomSoundVolume : WarningCount = 0 :FlipperDeActivate LeftFlipper, LFPress : SolLFlipper False : FlipperDeActivate RightFlipper, RFPress : SolRFlipper False
' End If
' LTiltPower = 0 : RTiltPower = 0 : CTiltPower = 0
'End Sub


'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt
  If CurrentBall = 0 Then Exit Sub
            'Called when table is nudged
  SetGI Rgb(255,123,10),1555

  If StartGame = 1 And Not bTilted And BallInLane = 0 And Drain.TimerEnabled = False Then
    Tilt = Tilt + TiltSensitivity                 'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt > TiltSensitivity) AND (Tilt <= 15) Then ShowTiltWarning 'show a warning
    If Tilt > 15 Then TiltMachine           'If more than 15 then TILT the table
  End If
End Sub

Sub CheckMechTilt                           'Called when mechanical tilt bob switch closed
  If CurrentBall = 0 Then Exit Sub
  ToyShake
  If StartGame = 1 And Not bTilted And BallInLane = 0 And Drain.TimerEnabled = False Then
    If Not bMechTiltJustHit Then
      MechTilt = MechTilt + 1                   'Add to tilt count
'     If(MechTilt > 0) AND (MechTilt <= 2) Then
      ShowTiltWarning 'show a warning
'     If MechTilt > 2 Then TiltMachine        'If more than 2 then TILT the table
      bMechTiltJustHit = True
      TiltDebounceTimer.Enabled = True
    End If
  End If
End Sub

'
'
Dim Warnings : Warnings = 0
Sub ShowTiltWarning

  Warnings = Warnings + 1
  If Warnings > 2 Then
    TiltMachine
  Else
    SetGI Rgb(255,0,0),2222
    ClearDMD
    If Warnings = 1 Then
      DMDMessage1 "WARNING"     ,""       ,FontSponge16bb ,69,16  ,69,16  ,"shake"  ,3  ,"BGwarn" ,"novideo"
    Else
      DMDMessage1 "WARNING"     ,"WARNING"    ,FontSponge12bb ,69,16  ,69,16  ,"shake"  ,3  ,"BGwarn" ,"novideo"
      Playsound "Sfx_Warning", 0, RomSoundVolume
    End If
    Playsound "Sfx_Warning", 0, RomSoundVolume
  End If
End Sub

Dim skipEOB : skipEOB = False

Sub TiltMachine
  BallsToShoot = 0
  EobContinue.Enabled = False
  drain.timerenabled = False
  lightCtrl.LightOff f146
  lightCtrl.AddTableLightSeq "GI", lSeqRainbow
  CurrentMode = 0
  skipEOB = True
  bTilted = True
  FlipperDeActivate LeftFlipper, LFPress
  SolLFlipper False
  SolULFlipper False
  FlipperDeActivate RightFlipper, RFPress
  SolRFlipper False
  If RestartGame < 1 Then ClearDMD
  If RestartGame < 1 Then DMDMessage1 "TILT"        ,""       ,FontSponge32BB ,57,16  ,57,16  ,"tilt" ,99 ,"noimage"  ,"VIDchaos"
  Playtime = 100
  StopTableMusic
  PlaySound "Sfx_Tilt",0,RomSoundVolume
  DisableTable True
  TurnLightsOff
  TiltRecoveryTimer2.interval = 1600
  TiltRecoveryTimer2.Enabled = True 'start the Tilt delay to check for all the balls to be drained

  FlexDMD.Stage.GetLabel("eb1").visible = False
  FlexDMD.Stage.GetLabel("eb2").visible = False
  FlexDMD.Stage.GetLabel("eb3").visible = False
  FlexDMD.Stage.GetLabel("eb4").visible = False
  FlexDMD.Stage.GetLabel("eb5").visible = False
  FlexDMD.Stage.GetLabel("eb6").visible = False
  FlexDMD.Stage.GetImage("BG014").Visible=False

End Sub


Sub TiltDecreaseTimer_Timer
  ' DecreaseTilt
  If Tilt> 0 Then
    Tilt = Tilt - 0.1
  Else
    TiltDecreaseTimer.Enabled = False
  End If
End Sub


Sub TiltDebounceTimer_Timer
  bMechTiltJustHit = False
  TiltDebounceTimer.Enabled = False
End Sub


Sub DisableTable(Enabled)
  If Enabled Then
    'GiOff
    ' Add tilt messages and sounds
    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart
    LeftSlingshot.Disabled = 1
    RightSlingshot.Disabled = 1
  Else
    'GiOn
    LeftSlingshot.Disabled = 0
    RightSlingshot.Disabled = 0
  End If
End Sub

Sub TiltRecoveryTimer2_Timer
  TiltRecoveryTimer2.Enabled = False

  lightCtrl.AddTableLightSeq "GI", lSeqRainbow

  TiltRecoveryTimer.interval = 1600
  TiltRecoveryTimer.Enabled = True
End Sub

Sub TiltRecoveryTimer_Timer()
  TiltRecoveryTimer.interval = 1234
  If ubound(getballs) > -1 Then Exit Sub
  TiltRecoveryTimer.Enabled = False
  ClearDMD
  FlexMsg = 1
  FlexMsgTime = 1
  lightCtrl.RemoveTableLightSeq "GI", lSeqRainbow
  lightCtrl.RemoveTableLightSeq "GI", lSeqRainbow
  skipEOB = True
  EobContinue_Timer
End Sub



'//////////////////////////////////////////////////////////////////////
'// Dynamic Ball Shadows
'//////////////////////////////////////////////////////////////////////

Const fovY          = 0   'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const DynamicBSFactor     = 0.95  '0 to 1, higher is darker
Const AmbientBSFactor     = 0.7 '0 to 1, higher is darker
Const AmbientMovement   = 2   '1 to 4, higher means more movement as the ball moves left and right
Const Wideness        = 20  'Sets how wide the dynamic ball shadows can get (20 +5 thinness should be most realistic for a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source


' *** This segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
' ' stop the sound of deleted balls
' For b = UBound(BOT) + 1 to tnob
'   If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
'   ...rolling(b) = False
'   ...StopSound("BallRoll_" & b)
' Next
'
'...rolling and drop sounds...

'   If DropCount(b) < 5 Then
'     DropCount(b) = DropCount(b) + 1
'   End If
'
'   ' "Static" Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     If BOT(b).Z > 30 Then
'       BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
'     Else
'       BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
'     End If
'     BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
'     BallShadowA(b).X = BOT(b).X
'     BallShadowA(b).visible = 1
'   End If

' *** Required Functions, enable these if they are not already present elswhere in your table
Function DistanceFast(x, y)
  dim ratio, ax, ay
  ax = abs(x)         'Get absolute value of each vector
  ay = abs(y)
  ratio = 1 / max(ax, ay)   'Create a ratio
  ratio = ratio * (1.29289 - (ax + ay) * ratio * 0.29289)
  if ratio > 0 then     'Quickly determine if it's worth using
    DistanceFast = 1/ratio
  Else
    DistanceFast = 0
  End if
end Function

'Function max(a,b)
' if a > b then
'   max = a
' Else
'   max = b
' end if
'end Function

'Dim PI: PI = 4*Atn(1)

'Function Atn2(dy, dx)
' If dx > 0 Then
'   Atn2 = Atn(dy / dx)
' ElseIf dx < 0 Then
'   If dy = 0 Then
'     Atn2 = pi
'   Else
'     Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
'   end if
' ElseIf dx = 0 Then
'   if dy = 0 Then
'     Atn2 = 0
'   else
'     Atn2 = Sgn(dy) * pi / 2
'   end if
' End If
'End Function
'
'Function AnglePP(ax,ay,bx,by)
' AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

'****** End Part B:  Code and Functions ******




'****** Part C:  The Magic ******
Dim sourcenames, currentShadowCount, DSSources(30), numberofsources, numberofsources_hold
sourcenames = Array ("","","","","","","","","","","","")
currentShadowCount = Array (0,0,0,0,0,0,0,0,0,0,0,0)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objrtx1(12), objrtx2(12)
dim objBallShadow(12)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5,BallShadowA6,BallShadowA7,BallShadowA8,BallShadowA9,BallShadowA10,BallShadowA11)

DynamicBSInit

sub DynamicBSInit()
  Dim iii, source

  for iii = 0 to tnob                 'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = iii/1000 + 1.01
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = (iii)/1000 + 1.02
    objrtx2(iii).visible = 0

    currentShadowCount(iii) = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = iii/1000 + 1.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100*AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source in DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
    iii = iii + 1
  Next
  numberofsources = iii
  numberofsources_hold = iii
end sub


Sub DynamicBSUpdate
  Dim falloff:  falloff = 150     'Max distance to light sources, can be changed if you have a reason
  Dim ShadowOpacity, ShadowOpacity2
  Dim s, Source, LSd, currentMat, AnotherSource, BOT, iii
  BOT = GetBalls

  'Hide shadow of deleted balls
  For s = UBound(BOT) + 1 to tnob
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(BOT) < lob Then Exit Sub    'No balls in play, exit

'The Magic happens now

  For s = lob to UBound(BOT)
' *** Normal "ambient light" ball shadow
  'Layered from top to bottom. If you had an upper pf at for example 80 and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

    If AmbientBallShadowOn = 1 Then     'Primitive shadow on playfield, flasher shadow in ramps
      If BOT(s).Z > 30 Then             'The flasher follows the ball up ramps while the primitive is on the pf
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + BallSize/10 + fovY
        objBallShadow(s).visible = 1

        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf, primitive only
        objBallShadow(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + fovY
        BallShadowA(s).visible = 0
      Else                      'Under pf, no shadows
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 0
      end if

    Elseif AmbientBallShadowOn = 2 Then   'Flasher shadow everywhere
      If BOT(s).Z > 30 Then             'In a ramp
        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf
        BallShadowA(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          BallShadowA(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          BallShadowA(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        BallShadowA(s).Y = BOT(s).Y + Ballsize/10 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/2 + 5
      Else                      'Under pf
        BallShadowA(s).visible = 0
      End If
    End If

' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If BOT(s).Z < 30 Then 'And BOT(s).Y < (TableHeight - 200) Then 'Or BOT(s).Z > 105 Then    'Defining when and where (on the table) you can have dynamic shadows
        For iii = 0 to numberofsources - 1
          LSd=DistanceFast((BOT(s).x-DSSources(iii)(0)),(BOT(s).y-DSSources(iii)(1))) 'Calculating the Linear distance to the Source
          If LSd < falloff And gilvl > 0 Then               'If the ball is within the falloff range of a light and light is on (we will set numberofsources to 0 when GI is off)
            currentShadowCount(s) = currentShadowCount(s) + 1   'Within range of 1 or 2
            if currentShadowCount(s) = 1 Then           '1 dynamic shadow source
              sourcenames(s) = iii
              currentMat = objrtx1(s).material
              objrtx2(s).visible = 0 : objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01            'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-LSd)/falloff                 'Sets opacity/darkness of shadow by distance to light
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness           'Scales shape of shadow with distance/opacity
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^2,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-ShadowOpacity),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-ShadowOpacity)
              End If

            Elseif currentShadowCount(s) = 2 Then
                                  'Same logic as 1 shadow, but twice
              currentMat = objrtx1(s).material
              AnotherSource = sourcenames(s)
              objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(AnotherSource)(0),DSSources(AnotherSource)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-DistanceFast((BOT(s).x-DSSources(AnotherSource)(0)),(BOT(s).y-DSSources(AnotherSource)(1))))/falloff
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0

              currentMat = objrtx2(s).material
              objrtx2(s).visible = 1 : objrtx2(s).X = BOT(s).X : objrtx2(s).Y = BOT(s).Y + fovY
  '           objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx2(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity2 = (falloff-LSd)/falloff
              objrtx2(s).size_y = Wideness*ShadowOpacity2+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity2*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2))
              End If
            end if
          Else
            currentShadowCount(s) = 0
            BallShadowA(s).Opacity = 100*AmbientBSFactor
          End If
        Next
      Else                  'Hide dynamic shadows everywhere else
        objrtx2(s).visible = 0 : objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub

'//////////////////////////////////////////////////////////////////////
'// FLIPPERS
'//////////////////////////////////////////////////////////////////////

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
  If Enabled Then
    DOF 101,1
    LF.Fire
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    DOF 101,0
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolULFlipper(Enabled)
  If Enabled Then
    LeftFlipper1.RotateToEnd
    If leftflipper1.currentangle < leftflipper1.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper1
    Else
      SoundFlipperUpAttackLeft LeftFlipper1
      RandomSoundFlipperUpLeft LeftFlipper1
    End If
  Else
    LeftFlipper1.RotateToStart
    If LeftFlipper1.currentangle < LeftFlipper1.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper1
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub


Sub SolRFlipper(Enabled)
  If Enabled Then
    DOF 102,1
    RF.Fire 'rightflipper.rotatetoend

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    DOF 102,0
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub


' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub


dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -2.7
'        AddPt "Polarity", 2, 0.33, -2.7
'        AddPt "Polarity", 3, 0.37, -2.7
'        AddPt "Polarity", 4, 0.41, -2.7
'        AddPt "Polarity", 5, 0.45, -2.7
'        AddPt "Polarity", 6, 0.576,-2.7
'        AddPt "Polarity", 7, 0.66, -1.8
'        AddPt "Polarity", 8, 0.743, -0.5
'        AddPt "Polarity", 9, 0.81, -0.5
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -3.7
'        AddPt "Polarity", 2, 0.33, -3.7
'        AddPt "Polarity", 3, 0.37, -3.7
'        AddPt "Polarity", 4, 0.41, -3.7
'        AddPt "Polarity", 5, 0.45, -3.7
'        AddPt "Polarity", 6, 0.576,-3.7
'        AddPt "Polarity", 7, 0.66, -2.3
'        AddPt "Polarity", 8, 0.743, -1.5
'        AddPt "Polarity", 9, 0.81, -1
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub
'
'


'*******************************************
'  Late 80's early 90's

'Sub InitPolarity()
' dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'   x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
' Next
'
' AddPt "Polarity", 0, 0, 0
' AddPt "Polarity", 1, 0.05, -5
' AddPt "Polarity", 2, 0.4, -5
' AddPt "Polarity", 3, 0.6, -4.5
' AddPt "Polarity", 4, 0.65, -4.0
' AddPt "Polarity", 5, 0.7, -3.5
' AddPt "Polarity", 6, 0.75, -3.0
' AddPt "Polarity", 7, 0.8, -2.5
' AddPt "Polarity", 8, 0.85, -2.0
' AddPt "Polarity", 9, 0.9,-1.5
' AddPt "Polarity", 10, 0.95, -1.0
' AddPt "Polarity", 11, 1, -0.5
' AddPt "Polarity", 12, 1.1, 0
' AddPt "Polarity", 13, 1.3, 0
'
' addpt "Velocity", 0, 0,         1
' addpt "Velocity", 1, 0.16, 1.06
' addpt "Velocity", 2, 0.41,         1.05
' addpt "Velocity", 3, 0.53,         1'0.982
' addpt "Velocity", 4, 0.702, 0.968
' addpt "Velocity", 5, 0.95,  0.968
' addpt "Velocity", 6, 1.03,         0.945
'
' LF.Object = LeftFlipper
' LF.EndPoint = EndPointLp
' RF.Object = RightFlipper
' RF.EndPoint = EndPointRp
'End Sub



'
''*******************************************
'' Early 90's and after
'
Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
                x.enabled = True
                x.TimeDelay = 60
        Next

        AddPt "Polarity", 0, 0, 0
        AddPt "Polarity", 1, 0.05, -5.5
        AddPt "Polarity", 2, 0.4, -5.5
        AddPt "Polarity", 3, 0.6, -5.0
        AddPt "Polarity", 4, 0.65, -4.5
        AddPt "Polarity", 5, 0.7, -4.0
        AddPt "Polarity", 6, 0.75, -3.5
        AddPt "Polarity", 7, 0.8, -3.0
        AddPt "Polarity", 8, 0.85, -2.5
        AddPt "Polarity", 9, 0.9,-2.0
        AddPt "Polarity", 10, 0.95, -1.5
        AddPt "Polarity", 11, 1, -1.0
        AddPt "Polarity", 12, 1.05, -0.5
        AddPt "Polarity", 13, 1.1, 0
        AddPt "Polarity", 14, 1.3, 0

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        LF.Object = LeftFlipper
        LF.EndPoint = EndPointLp
        RF.Object = RightFlipper
        RF.EndPoint = EndPointRp
End Sub


' Flipper trigger hit subs
Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub




'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
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
  Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property
  Public Property Get EndPointY: EndPointY = FlipperEndY : End Property

  Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
    if gametime > 100 then Report aChooseArray
  End Sub

  Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
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
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        RemoveBall aBall
        exit Sub
      end if

      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
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
  redim aArray(aCount-1+offset)        'Resize original array
  for x = 0 to aCount-1                'set objects back into original array
    if IsObject(a(x)) then
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
  BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
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

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

  LinearEnvelope = Y
End Function


'******************************************************
'  FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b, BOT
  BOT = GetBalls

  If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          exit Sub
        end If
      Next
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
        end If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 then EOSNudge1 = 0
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

'*****************
' Maths
'*****************
Dim PI: PI = 4*Atn(1)

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

Function Atn2(dy, dx)
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

'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function


'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0:
    SOSRampup = 2.5
  Case 1:
    SOSRampup = 6
  Case 2:
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

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
  Flipper.eostorque = EOST*EOSReturn/FReturn


  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b, BOT
    BOT = GetBalls

    For b = 0 to UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3*Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
    if FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If

  End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
  Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime : CatchTime = GameTime - FCount

  if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
    if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    else
      LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
    end If

    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
  End If
End Sub

'//////////////////////////////////////////////////////////////////////
'// TargetBounce
'//////////////////////////////////////////////////////////////////////

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

' Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit
  TargetBouncer activeball, 1
End Sub

'//////////////////////////////////////////////////////////////////////
'// PHYSICS DAMPENERS
'//////////////////////////////////////////////////////////////////////

' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR



Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
  TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
  TargetBouncer Activeball, 0.7
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'#########################    Adjust these values to increase or lessen the elasticity

dim FlippersD : Set FlippersD = new Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold         'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    if debugOn then TBPout.text = str
  End Sub

  public sub Dampenf(aBall, parm) 'Rubberizer is handle here
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
      aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()         'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub

End Class

'//////////////////////////////////////////////////////////////////////
'// TRACK ALL BALL VELOCITIES FOR RUBBER DAMPENER AND DROP TARGETS
'//////////////////////////////////////////////////////////////////////

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

'//////////////////////////////////////////////////////////////////////
'// RAMP ROLLING SFX
'//////////////////////////////////////////////////////////////////////

'Ball tracking ramp SFX 1.0
'   Reqirements:
'          * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'          * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'          * Create a Timer called RampRoll, that is enabled, with a interval of 100
'          * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'          * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'          * To stop tracking ball
'                 * call WireRampOff
'                 * Otherwise, the ball will auto remove if it's below 30 vp units
'

dim RampMinLoops : RampMinLoops = 4

' RampBalls
'      Setup:        Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RammBalls(6,2)
'      Description:
dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(6)

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID : End Sub


' WaddBall (Active Ball, Boolean)
'     Description: This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
Sub Waddball(input, RampInput)  'Add ball
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  dim x : for x = 1 to uBound(RampBalls)  'Check, don't add balls twice
    if RampBalls(x, 1) = input.id then
      if Not IsEmpty(RampBalls(x,1) ) then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 to uBound(RampBalls)
    if IsEmpty(RampBalls(x, 1)) then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      exit Sub
    End If
    if x = uBound(RampBalls) then   'debug
      Debug.print "WireRampOn error, ball queue is full: " & vbnewline & _
      RampBalls(0, 0) & vbnewline & _
      Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbnewline & _
      Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbnewline & _
      Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbnewline & _
      Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbnewline & _
      Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbnewline & _
      " "
    End If
  next
End Sub

' WRemoveBall (BallId)
'    Description: This subroutine is called from the RampRollUpdate subroutine
'                 and is used to remove and stop the ball rolling sounds
Sub WRemoveBall(ID)   'Remove ball
  'Debug.Print "In WRemoveBall() + Remove ball from loop array"
  dim ballcount : ballcount = 0
  dim x : for x = 1 to Ubound(RampBalls)
    if ID = RampBalls(x, 1) then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    if not IsEmpty(Rampballs(x,1)) then ballcount = ballcount + 1
  next
  if BallCount = 0 then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer():RampRollUpdate:End Sub

Sub RampRollUpdate()    'Timer update
  dim x : for x = 1 to uBound(RampBalls)
    if Not IsEmpty(RampBalls(x,1) ) then
      if BallVel(RampBalls(x,0) ) > 1 then ' if ball is moving, play rolling sound
        If RampType(x) then
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * MechVolume, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * MechVolume, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end if
  next
  if not RampBalls(0,0) then RampRoll.enabled = 0

End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbnewline & _
  "1 " & Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbnewline & _
  "2 " & Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbnewline & _
  "3 " & Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbnewline & _
  "4 " & Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbnewline & _
  "5 " & Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbnewline & _
  "6 " & Typename(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbnewline & _
  " "
End Sub


Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(ball), 1, -1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function

'//////////////////////////////////////////////////////////////////////
'// RAMP TRIGGERS
'//////////////////////////////////////////////////////////////////////

Sub Gate1_hit ()
RandomSoundBallRelease BallRelease
End Sub

Sub RRSound001_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub RRSound002_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Dim RampCombo
Dim RampComboSelect

Dim blockmysterycount
Sub RRSound002_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
  DodNext = 5

  If blockmysterycount = 0 Then Rampblinker_Timer
  InsertSequence.Play SeqBlinking,,3,44


  If CurrentMode = 20 Then
    If blockmysterycount = 1 Then blockmysterycount = 0 : Exit Sub
    wizAddaBall = wizAddaBall + 1
    If wizAddaBall = 2 Or wizAddaBall = 11 Then
      Playsound "Sfx_ModeReady" ,1,RomSoundVolume
      AddABallReady = 1
      DMDMessage1 "ADDABALL"      ,"READY"        ,FontSponge12bb ,54,16  ,59,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
    End If

    If wizAddaBall = 3 Then
      AddABallReady = 0
      DMDMessage1 "ADDABALL"      ,""       ,FontSponge16bb ,54,16  ,59,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
      ballstoshoot = ballstoshoot + 1 : wizBiP = wizBiP + 1
      PlaySound "Sfx_SecretShot",1,RomSoundVolume
    End If
    If wizAddaBall = 12 Then
      AddABallReady = 0
      DMDMessage1 "ADDABALL"      ,""       ,FontSponge16bb ,54,16  ,59,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
      ballstoshoot = ballstoshoot + 1 : wizBiP = wizBiP + 1
      PlaySound "Sfx_SecretShot",1,RomSoundVolume
    End If
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then
    If blockmysterycount = 1 Then blockmysterycount = 0 : Exit Sub

    addscore (25000)
    If collect7 > 0 Then collect7 = collect7 - 1 : WizardJPUpdate

  Else


    addscore (2000)
    If blockmysterycount = 1 Then
      blockmysterycount = 0
    Else

      If currentmode <> 21 Then If collect7 = 1 Then collect7 = 0 : CheckCollected

      If CurrentMode <> 11 then
        RampsMysteryCounter(CurrentPlayer) = RampsMysteryCounter(CurrentPlayer) + 1
        If RampsMysteryCounter(CurrentPlayer) = RampsForMystery(CurrentPlayer) Then
          DMDMessage1 "MYSTERY"     ,"READY"        ,FontSponge12bb ,54,16  ,59,16  ,"blink"  ,.9 ,"BG001"  ,"novideo"
          l40.blinkinterval = 60 : l40.state = 2
          l41.blinkinterval = 60 : l41.state = 2
          l42.blinkinterval = 50 : l42.state = 2
          MysteryReady(CurrentPlayer) = 1
        End If
      End If

      If CurrentMode = 11 Then
        MBRampsforAddaBall = MBRampsforAddaBall + 1
        If MBRampsforAddaBall = 2 + MBaddaball(CurrentPlayer) Then
          Playsound "Sfx_ModeReady" ,1,RomSoundVolume
          DMDMessage1 "ADDABALL"      ,"READY"        ,FontSponge12bb ,54,16  ,59,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
          AddABallReady = 1
        End If
        If MBRampsforAddaBall = 3 + MBaddaball(CurrentPlayer) Then
          MBaddaball(CurrentPlayer) = MBaddaball(CurrentPlayer) + 1
          AddABallReady = 2
          RandomKicker(4)
          DMDMessage1 "ADDABALL"      ,""       ,FontSponge16bb ,54,16  ,59,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
          PlaySound "Sfx_SecretShot",1,RomSoundVolume
          MBRampsforAddaBall = 999
        End If
      End If

      If RampComboSelect <> 1 Then

        If CurrentMode = 11 Then ' rampmillion
          If RampCombo > frame Then
            ComboCount = ComboCount + 1
            CompletedCombos(CurrentPlayer) = CompletedCombos(CurrentPlayer) + 1 ': Debug.print "completed Combos=" & CompletedCombos(CurrentPlayer)
            addscore (3000000)
            DMDMessage1 "RAMP COMBO"      ,"3.000.000"        ,FontSponge12bb ,64,17  ,69,17  ,"shake"  ,.9 ,"BG001"  ,"novideo"
          Else
            ComboCount = 0
            addscore (1500000)
            DMDMessage1 "1.500.000"     ,""       ,FontSponge12bb ,64,16  ,69,16  ,"shake"  ,.9 ,"BG001"  ,"novideo"
          End If
        Else
          CheckCombo
          RampCombo = frame + 500
          RampComboSelect = 1

        End If
      End If
      UpdateLights
    End If
  End If


End Sub
Dim ComboCount



Sub RRSound003_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
    'Flash3
  'Flash4
End Sub

Sub RRSound003_unhit()
  PlaySoundAt "WireRamp_Stop", RRSound003
End Sub



Dim MBRampsforAddaBall
Dim AddABallReady

Sub LRSound001_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub LRSound002_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
  DodNext = 5

  Rampblinker_Timer
' SetGI Rgb(0,88,255),900
' LightSeqWiz1.play SeqBlinking,,3,35
  InsertSequence.Play SeqBlinking,,7,44
  If CurrentMode = 20 Then
    wizAddaBall = wizAddaBall + 1
    If wizAddaBall = 2 Or wizAddaBall = 11 Then
      Playsound "Sfx_ModeReady" ,1,RomSoundVolume
      AddABallReady = 1
      DMDMessage1 "ADDABALL"      ,"READY"        ,FontSponge12bb ,54,16  ,59,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
    End If

    If wizAddaBall = 3 Then
      AddABallReady = 0
      DMDMessage1 "ADDABALL"      ,""       ,FontSponge16bb ,54,16  ,59,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
      ballstoshoot = ballstoshoot + 1 : wizBiP = wizBiP + 1
    End If
    If wizAddaBall = 12 Then
      AddABallReady = 0
      DMDMessage1 "ADDABALL"      ,""       ,FontSponge16bb ,54,16  ,59,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
      ballstoshoot = ballstoshoot + 1 : wizBiP = wizBiP + 1
    End If
    AddScore (100000)
    WizardJPUpdate
  Elseif CurrentMode = 21 Then

    addscore (25000)
    If collect3 > 0 Then collect3 = collect3 - 1 : WizardJPUpdate

  Else

    if currentmode <> 21 then If collect3 = 1 Then collect3 = 0 : CheckCollected

    addscore (2000)
    If blockmysterycount = 1 Then
      blockmysterycount = 0
    Else
      If CurrentMode <> 11 then
        RampsMysteryCounter(CurrentPlayer) = RampsMysteryCounter(CurrentPlayer) + 1
        If RampsMysteryCounter(CurrentPlayer) = RampsForMystery(CurrentPlayer) Then
          DMDMessage1 "MYSTERY"     ,"READY"        ,FontSponge12bb ,54,16  ,59,16  ,"blink"  ,.9 ,"BG001"  ,"novideo"
          l40.blinkinterval = 60 : l40.state = 2
          l41.blinkinterval = 60 : l41.state = 2
          l42.blinkinterval = 50 : l42.state = 2
          MysteryReady(CurrentPlayer) = 1
        End If
      End If

      If CurrentMode = 11 Then
        MBRampsforAddaBall = MBRampsforAddaBall + 1
        If MBRampsforAddaBall = 2 + MBaddaball(CurrentPlayer) Then
          Playsound "Sfx_ModeReady" ,1,RomSoundVolume
          DMDMessage1 "ADDABALL"      ,"READY"        ,FontSponge12bb ,54,16  ,59,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
          AddABallReady = 1
        End If
        If MBRampsforAddaBall = 3 + MBaddaball(CurrentPlayer) Then
          MBaddaball(CurrentPlayer) = MBaddaball(CurrentPlayer) + 1
          AddABallReady = 2
          RandomKicker(4)
          DMDMessage1 "ADDABALL"      ,""       ,FontSponge16bb ,54,16  ,59,16  ,"blink"  ,1.5  ,"BG001"  ,"novideo"
          PlaySound "Sfx_SecretShot",1,RomSoundVolume
MBRampsforAddaBall = 999
        End If
      End If

      If RampComboSelect <> 2 Then
        If CurrentMode = 11 Then ' rampmillion
          If RampCombo > frame Then
            ComboCount = ComboCount + 1
            CompletedCombos(CurrentPlayer) = CompletedCombos(CurrentPlayer) + 1 ': Debug.print "completed Combos=" & CompletedCombos(CurrentPlayer)
            addscore (3000000)
            DMDMessage1 "RAMP COMBO"      ,"3.000.000"        ,FontSponge12bb ,64,17  ,69,17  ,"shake"  ,.9 ,"BG001"  ,"novideo"
          Else
            ComboCount = 0
            addscore (1500000)
            DMDMessage1 "1.500.000"     ,""       ,FontSponge12bb ,64,16  ,69,16  ,"shake"  ,.9 ,"BG001"  ,"novideo"
          End If
        Else
          CheckCombo
          RampCombo = frame + 500
          RampComboSelect = 2
        End If
      End If
      UpdateLights
    End If
  End If
End Sub


Sub CheckCombo

  If RampCombo > frame Then
    ComboCount = ComboCount + 1

    CompletedCombos(CurrentPlayer) = CompletedCombos(CurrentPlayer) + 1
    If CompletedCombos(CurrentPlayer) = 20 Then
      ClearDMD
      ExtraBall = ExtraBall + 1 : DoEball = 1
      DOF 129,2
      DMDMessage1 "EXTRA BALL"    ,""       ,FontSponge16bb ,53,16  ,53,16  ,"solid"  ,2.2  ,"BG001"  ,"novideo"
      Call JackpotSequence(750) : StopTableMusic : PlaySound "Sfx_ExtraBall",0,RomSoundVolume
      songtimer.interval = 4040
      songtimer.enabled = True
      UpdateLights
      Exit Sub
    End If
    Select Case ComboCount
      case 1 : addscore ( 250000) : DMDMessage1 "COMBO"     ,"250.000"        ,FontSponge12bb ,64,17  ,64,17  ,"shake"  ,1.2  ,"BG001"  ,"novideo"
      case 2 : addscore ( 500000) : DMDMessage1 "2 COMBOS"    ,"500.000"        ,FontSponge12bb ,64,17  ,64,17  ,"shake"  ,1.2  ,"BG001"  ,"novideo"
      case 3 : addscore ( 750000) : DMDMessage1 "3 COMBOS"    ,"750.000"        ,FontSponge12bb ,64,17  ,64,17  ,"shake"  ,1.2  ,"BG001"  ,"novideo"
      case 4,5 : addscore (1000000) : DMDMessage1 "SUPER COMBO" ,"1.000.000"      ,FontSponge12bb ,45,17  ,64,17  ,"shake"  ,1.5  ,"BG001"  ,"novideo"
          ComboCount = 3
    End Select
    If CompletedCombos(CurrentPlayer) = 3 or CompletedCombos(CurrentPlayer) = 7 or CompletedCombos(CurrentPlayer) = 13 or CompletedCombos(CurrentPlayer) = 18 or CompletedCombos(CurrentPlayer) = 19 Then
      DMDMessage1 CompletedCombos(CurrentPlayer) & "/20 COMBOS"     ,"FOR EXTRABALL"        ,FontSponge12bb ,55,16  ,55,16  ,"solid"  ,1.2  ,"BG001"  ,"novideo"
    End If
  Else
    ComboCount = 0
  End If
  UpdateLights
End Sub



Sub LRSound002_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub LRSound003_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
    'Flash2
  'Flash5
End Sub

Sub LRSound003_unhit()
  PlaySoundAt "WireRamp_Stop", LRSound003
End Sub

Sub FlasherTrigger_hit()
  'Flash3
End Sub



'//////////////////////////////////////////////////////////////////////
'// Ball Rolling
'//////////////////////////////////////////////////////////////////////

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 to tnob
    rolling(i) = False
  Next
End Sub

Sub RollingUpdate()
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
  If UBound(BOT) > 4 Then Exit Sub
  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * MechVolume, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz > -7 Then
          RandomSoundBallBouncePlayfieldSoft BOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard BOT(b)
        End If
      End If
    End If
    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If BOT(b).Z > 30 Then
        BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
      Else
        BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
      End If
      BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
      BallShadowA(b).X = BOT(b).X
      BallShadowA(b).visible = 1
    End If

    ' Ramp Boost
        If BOT(b).z > 120 and inRect(BOT(b).x,BOT(b).y,490,450,490,200,633,200,633,450) and BOT(b).vely < 1 Then
      BOT(b).velx = BOT(b).velx - 0.1
      BOT(b).vely = BOT(b).vely + 0.1
      'debug.print BOT(b).x & " " & BOT(b).y
        End If

  Next
End Sub

'//////////////////////////////////////////////////////////////////////
'// Mechanic Sounds
'//////////////////////////////////////////////////////////////////////

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'   Metals (all metal objects, metal walls, metal posts, metal wire guides)
'   Apron (the apron walls and plunger wall)
'   Walls (all wood or plastic walls)
'   Rollovers (wire rollover triggers, star triggers, or button triggers)
'   Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'   Gates (plate gates)
'   GatesWire (wire gates)
'   Rubbers (all rubbers including posts, sleeves, pegs, and bands)
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
' Tutorial vides by Apophis
' Part 1:   https://youtu.be/PbE2kNiam3g
' Part 2:   https://youtu.be/B5cm1Y8wQsk
' Part 3:   https://youtu.be/eLhWyuYOyGg


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                         'volume level; range [0, 1]
NudgeRightSoundLevel = 1                        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                       'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                       'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                     'volume level; range [0, 1]
PlungerPullSoundLevel = 1                       'volume level; range [0, 1]
RollingSoundFactor = 1.1/5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                      'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                  'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel               'sound helper; not configurable
SlingshotSoundLevel = 0.95                        'volume level; range [0, 1]
BumperSoundFactor = 8                       'volume multiplier; must not be zero
KnockerSoundLevel = 1                           'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                  'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                     'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                     'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                    'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                 'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                 'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                     'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                          'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                     'volume multiplier; must not be zero
DTSoundLevel = 0.25                           'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                       'volume level; range [0, 1]
SpinnerSoundLevel = 0.5                                       'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                           'volume level; range [0, 1]
BallReleaseSoundLevel = 1                       'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                   'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                         'volume multiplier; must not be zero


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
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, -1, min(aVol,1) * MechVolume, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, -1, min(aVol,1) * MechVolume, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * MechVolume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, min(aVol,1) * MechVolume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  Playsound soundname, 1,min(aVol,1) * MechVolume, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,min(aVol,1) * MechVolume, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * MechVolume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / tableheight-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = Csng((ball.velz) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  PitchPlayfieldRoll = BallVel(ball) ^2 * 15
End Function

Function RndInt(min, max)
  RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
  RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * MechVolume, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * MechVolume, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * MechVolume, 0, 0.25
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
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*6)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*7)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*7)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////
Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub


'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm/10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm/10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong 1
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
  RandomSoundWall()
End Sub

Sub RandomSoundWall()
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
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
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft Activeball
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
  If Activeball.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If activeball.velx < -8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If Activeball.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If activeball.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Activeball
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  Dim snd

  FlipperCradleCollision ball1, ball2, velocity

  Select Case Int(Rnd*7)+1
    Case 1 : snd = "Ball_Collide_1"
    Case 2 : snd = "Ball_Collide_2"
    Case 3 : snd = "Ball_Collide_3"
    Case 4 : snd = "Ball_Collide_4"
    Case 5 : snd = "Ball_Collide_5"
    Case 6 : snd = "Ball_Collide_6"
    Case 7 : snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * MechVolume, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd*6)+1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315                  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05                  'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_GI_On"), RelayGISoundLevel*MechVolume, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_GI_Off"), RelayGISoundLevel*MechVolume, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025*RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025*RelayFlashSoundLevel, obj
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************
' To add these slingshot corrections:
'   - On the table, add the endpoint primitives that define the two ends of the Slingshot
' - Initialize the SlingshotCorrection objects in InitSlingCorrection
'   - Call the .VelocityCorrect methods from the respective _Slingshot event sub


dim LS : Set LS = New SlingshotCorrection
dim LS1 : Set LS1 = New SlingshotCorrection
dim RS : Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection

  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  LS1.Object = LeftSlingshot1
  LS1.EndPoint1 = EndPoint1LS1
  LS1.EndPoint2 = EndPoint2LS1

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00,  -4
  AddSlingsPt 1, 0.45,  -7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4

End Sub


Sub AddSlingsPt(idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LS, RS)
  dim x : for each x in a
    x.addpoint idx, aX, aY
  Next
End Sub

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function
'
Function RotPoint(x,y,angle)
    dim rx, ry
    rx = x*dCos(angle) - y*dSin(angle)
    ry = x*dSin(angle) + y*dCos(angle)
    RotPoint = Array(rx,ry)
End Function

Class SlingshotCorrection
  Public DebugOn, Enabled
  private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): Enabled = True : End Sub

  Public Property let Object(aInput) : Set Slingshot = aInput : End Property
  Public Property Let EndPoint1(aInput) : SlingX1 = aInput.x: SlingY1 = aInput.y: End Property
  Public Property Let EndPoint2(aInput) : SlingX2 = aInput.x: SlingY2 = aInput.y: End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    If gametime > 100 then Report
  End Sub

  Public Sub Report()         'debug, reports all coords in tbPL.text
    If not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1 : YL = SlingY1 : XR = SlingX2 : YR = SlingY2
    Else
      XL = SlingX2 : YL = SlingY2 : XR = SlingX1 : YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If ABS(XR-XL) > ABS(YR-YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If not IsEmpty(ModIn(0) ) then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      'debug.print " BallPos=" & BallPos &" Angle=" & Angle
      'debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled then aBall.Velx = RotVxVy(0)
      If Enabled then aBall.Vely = RotVxVy(1)
      'debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      'debug.print " "
    End If
  End Sub

End Class


' FlexDMD constants
Const   FlexDMD_RenderMode_DMD_GRAY = 0, _
    FlexDMD_RenderMode_DMD_GRAY_4 = 1, _
    FlexDMD_RenderMode_DMD_RGB = 2, _
    FlexDMD_RenderMode_SEG_2x16Alpha = 3, _
    FlexDMD_RenderMode_SEG_2x20Alpha = 4, _
    FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num = 5, _
    FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num_4x1Num = 6, _
    FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num = 7, _
    FlexDMD_RenderMode_SEG_2x7Num_2x7Num_10x1Num = 8, _
    FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num_gen7 = 9, _
    FlexDMD_RenderMode_SEG_2x7Num10_2x7Num10_4x1Num = 10, _
    FlexDMD_RenderMode_SEG_2x6Num_2x6Num_4x1Num = 11, _
    FlexDMD_RenderMode_SEG_2x6Num10_2x6Num10_4x1Num = 12, _
    FlexDMD_RenderMode_SEG_4x7Num10 = 13, _
    FlexDMD_RenderMode_SEG_6x4Num_4x1Num = 14, _
    FlexDMD_RenderMode_SEG_2x7Num_4x1Num_1x16Alpha = 15, _
    FlexDMD_RenderMode_SEG_1x16Alpha_1x16Num_1x7Num = 16

Const   FlexDMD_Align_TopLeft = 0, _
    FlexDMD_Align_Top = 1, _
    FlexDMD_Align_TopRight = 2, _
    FlexDMD_Align_Left = 3, _
    FlexDMD_Align_Center = 4, _
    FlexDMD_Align_Right = 5, _
    FlexDMD_Align_BottomLeft = 6, _
    FlexDMD_Align_Bottom = 7, _
    FlexDMD_Align_BottomRight = 8


Sub flex_init
    Dim fso,curdir

    Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
    If FlexDMD is Nothing Then
      MsgBox "No FlexDMD found. This table will NOT run without it."
      Exit Sub
    End If
    SetLocale(1033)
    With FlexDMD
      .GameName = cGameName
      .TableFile = Table1.Filename & ".vpx"
      .Color = RGB(255, 88, 32)
      .RenderMode = FlexDMD_RenderMode_DMD_RGB
      .Width = 128
      .Height = 32
      .Clear = True
      .ProjectFolder = "./SpongebobDMD/"
      '.Run = True
    End With


    CreateGameDMD
end sub


Dim FontScoreActive
Dim FontScoreBlack
Dim FontScoreInactive
Dim FontScoreInactive2
Dim FontBig
Dim FontBig2
Dim FontBig3
Dim FontBig4
Dim FontSponge12
Dim FontSponge16
Dim FontSponge16A
Dim FontSponge16B
Dim FontSponge16BB
Dim FontSponge16black
Dim FontSponge32
Dim FontSponge32black
Dim FontSponge32BB
Dim FontSponge32RB
Dim FlexDMD
Dim FontScoreCredits
Dim FontHighscore
Dim FontHighscore2
Dim FontSponge12bb
Dim FontSponge12bb2
Dim FontSponge12h
Dim FontSponge12black

sub CreateGameDMD

  Dim title, af,list

  Set FontScoreActive = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, RGB(0, 0, 0), 0)

  Set FontScoreBlack = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbBlack, RGB(0, 0, 0), 0)
  Set FontScoreInactive = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(64, 64, 64), RGB(0, 0, 0), 0)
  Set FontScoreInactive2 = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(80, 80, 80), RGB(0, 0, 0), 0)
  Set FontScoreCredits = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(16, 118, 27),  RGB(0, 0, 0), 0)
  Set FontBig = FlexDMD.NewFont("FlexDMD.Resources.bm_army-12.fnt", RGB(255, 155, 0), vbWhite, 0)
  Set FontBig2 = FlexDMD.NewFont("FlexDMD.Resources.bm_army-12.fnt", RGB(155, 85, 0), vbWhite, 0)

  Set FontBig3 = FlexDMD.NewFont("Spongebob12.fnt", RGB(255, 155, 0), vbWhite, 0)
  Set FontBig4 = FlexDMD.NewFont("Spongebob12.fnt", RGB(155, 85, 0), vbWhite, 0)

  Set FontSponge12    = FlexDMD.NewFont("Spongebob12.fnt", RGB(222,155, 0), RGB(0, 0, 0), 0)
  Set FontSponge12bb    = FlexDMD.NewFont("Spongebob12.fnt", RGB(222,155, 0), RGB(0, 0, 0), 1)
  Set FontSponge12bb2   = FlexDMD.NewFont("Spongebob12.fnt", RGB(222,55, 0), RGB(0, 0, 0), 1)

  Set FontSponge12h   = FlexDMD.NewFont("Spongebob12.fnt", RGB(88,33, 11), RGB(0, 0, 0), 0)
  Set FontSponge12black = FlexDMD.NewFont("Spongebob12.fnt", RGB(0,0, 0), RGB(0, 0, 0), 0)

  Set FontSponge16  = FlexDMD.NewFont("Spongebobx16.fnt", RGB(255, 222, 0), RGB(0, 0, 0), 0 )
  Set FontSponge16A = FlexDMD.NewFont("Spongebobx16.fnt", RGB(255, 150, 10), RGB(0, 0, 0), 0 )
  Set FontSponge16B = FlexDMD.NewFont("Spongebobx16.fnt", RGB(10, 255, 10), RGB(0, 0, 0), 0 )
  Set FontSponge16BB  = FlexDMD.NewFont("Spongebobx16.fnt", RGB(255, 222, 0), RGB(0, 0, 0), 1 )
  Set FontSponge16black = FlexDMD.NewFont("Spongebobx16.fnt", RGB(0, 0, 0), RGB(0, 0, 0), 0 )

  Set FontSponge32  = FlexDMD.NewFont("Spongebobx32.fnt", RGB(255, 255, 255), RGB(0, 0, 0), 0)
  Set FontSponge32black  = FlexDMD.NewFont("Spongebobx32.fnt", RGB(85, 85, 145), RGB(0, 0, 0), 1)
  Set FontSponge32BB  = FlexDMD.NewFont("Spongebobx32.fnt", RGB(255, 255, 255), RGB(0, 0, 0), 1)
  Set FontSponge32RB  = FlexDMD.NewFont("Spongebobx32.fnt", RGB(255, 255, 255), RGB(255, 150,50), 1)

  Set FontHighscore = FlexDMD.NewFont("FlexDMD.Resources.bm_army-12.fnt", RGB(11,11,11), vbWhite, 0)
  Set FontHighscore2 = FlexDMD.NewFont("FlexDMD.Resources.bm_army-12.fnt", RGB(255,140,11), vbWhite, 0)


  Dim scene


  Set scene = FlexDMD.NewGroup("Score")
  Set title = FlexDMD.NewImage("bubbles", "bubblesblue.png") : title.Visible = false : scene.AddActor title
  Set title = FlexDMD.NewImage("sling1", "slinger5.png") : title.Visible = false : scene.AddActor title
  Set title = FlexDMD.NewImage("sling2", "slinger4.png") : title.Visible = false : scene.AddActor title
  Set title = FlexDMD.NewImage("sling3", "slinger3.png") : title.Visible = false : scene.AddActor title
  Set title = FlexDMD.NewImage("sling4", "slinger2.png") : title.Visible = false : scene.AddActor title
  Set title = FlexDMD.NewImage("sling5", "slinger1.png") : title.Visible = false : scene.AddActor title
  Set title = FlexDMD.NewImage("fire1", "fire1.png") : title.Visible = false : scene.AddActor title
  Set title = FlexDMD.NewImage("fire2", "fire2.png") : title.Visible = false : scene.AddActor title
  Set title = FlexDMD.NewImage("fire3", "fire3.png") : title.Visible = false : scene.AddActor title
  Set title = FlexDMD.NewImage("fire4", "fire4.png") : title.Visible = false : scene.AddActor title
  Set title = FlexDMD.NewImage("fire5", "fire5.png") : title.Visible = false : scene.AddActor title
  Set title = FlexDMD.NewImage("fire6", "fire6.png") : title.Visible = false : scene.AddActor title
  Set title = FlexDMD.NewImage("fire7", "fire7.png") : title.Visible = false : scene.AddActor title

  For i = 1 to 4
    scene.Addactor FlexDMD.NewLabel("Score_" & i, FontScoreInactive, " ")
  Next
  scene.AddActor FlexDMD.NewLabel("comma1", FontScoreActive, ".")
  scene.AddActor FlexDMD.NewLabel("comma1b", FontScoreActive, ".")
  scene.AddActor FlexDMD.NewLabel("comma1c", FontScoreActive, ".")

  scene.AddActor FlexDMD.NewFrame("VSeparator")
  scene.GetFrame("VSeparator").Thickness = 1
  scene.GetFrame("VSeparator").SetBounds 40, 0, 1, 32

  Dim scene1
  Set scene1 = FlexDMD.NewGroup("Content")
  scene1.Clip = True
  scene1.SetBounds 41, 0, 128-41, 32
  scene1.AddActor FlexDMD.NewLabel("Ball", FontScoreCredits, " ")
  scene1.AddActor FlexDMD.NewLabel("Credit", FontScoreCredits, " ")


  scene1.GetLabel("Credit").SetAlignedPosition 52, 27, FlexDMD_Align_TopRight
  scene1.GetLabel("Ball").SetAlignedPosition 1, 27, FlexDMD_Align_TopLeft
  Set title = FlexDMD.NewLabel("Content_1", FontBig, " ")
' Set af = title.ActionFactory
' Set list = af.Sequence()
' title.AddAction af.Repeat(list, 1)
  scene1.AddActor title
  Set title = FlexDMD.NewLabel("TextSmalLine5", FontScoreInactive , " ")  : title.Visible = False : scene1.AddActor title
  Set title = FlexDMD.NewLabel("TextSmalLine6", FontScoreInactive , " ")  : title.Visible = False : scene1.AddActor title  ' timer mode+wiz

  Set title = FlexDMD.NewLabel("TextSmalLine7", FontSponge16 , " ") : title.Visible = False : scene1.AddActor title ' rl scoring number 1
  Set title = FlexDMD.NewLabel("TextSmalLine8", FontSponge16 , " ") : title.Visible = False : scene1.AddActor title ' rl scoring number 1
  Set title = FlexDMD.NewLabel("TextSmalLine9", FontSponge16 , " ") : title.Visible = False : scene1.AddActor title ' rl scoring number 1
  Set title = FlexDMD.NewLabel("TextSmalLine10",  FontSponge16 , " ") : title.Visible = False : scene1.AddActor title ' rl scoring number 1
  Set title = FlexDMD.NewLabel("TextSmalLine11",  FontSponge16 , " ") : title.Visible = False : scene1.AddActor title ' rl scoring number 1

  Dim scene2
  Set scene2 = FlexDMD.NewGroup("Overlay")
  Set title = FlexDMD.NewImage("dod1",    "doordie1.png") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("dod2",    "doordie2.png") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("dod3",    "doordie3.png") : title.Visible = False : scene2.AddActor title


  Set title = FlexDMD.NewImage("five1",     "500k1.png")    : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("five2",     "500k2.png")    : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("five3",     "500k3.png")    : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("five4",     "500k4.png")    : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("five5",     "500k5.png")    : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("Title",     "Title.png")    : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("BG001",   "SquareBG001.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("dodss",   "doordie.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("party",   "pineappleparty2.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("BG004",   "SquareBG004.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("BG005",   "SquareBG005.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("BG006",   "SquareBG006.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("BG007",   "SquareBG007.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("BG008",   "SquareBG008.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("BG009",   "SquareBG009.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("BG010",   "SquareBG002.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("BG011",   "SquareBG011.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("BG012",   "SquareBG012.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("BG013",   "SquareBG013.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title



' bottom tiny text line
  Set title = FlexDMD.NewImage("BG002",   "yellodmdline.png") : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  scene2.AddActor FlexDMD.NewLabel("TextSmalLine4", FontSponge16, " ")



  Set title = FlexDMD.NewVideo("VIDchaos",  "caos-bob.gif")         : title.SetBounds 0, 0, 128, 44 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title
  Set title = FlexDMD.NewVideo("VIDsponge", "Game_Mode_Spongebob.gif")    : title.SetBounds 0, 0, 128, 44 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title
  Set title = FlexDMD.NewVideo("VIDplankt", "Game_Mode_Plankton.gif")   : title.SetBounds 0, 0, 128, 44 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title
  Set title = FlexDMD.NewVideo("VIDgoofy",  "Game_Mode_Goofy_Goobers.gif")  : title.SetBounds 0, 0, 128, 44 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title
  Set title = FlexDMD.NewVideo("VIDcrabs",  "krustykrab.gif")       : title.SetBounds 0, 0, 128, 32 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title
  Set title = FlexDMD.NewVideo("VIDjelly",  "Game_Mode_Jellyfish Jam.gif")  : title.SetBounds 0, 0, 128, 44 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title
  Set title = FlexDMD.NewVideo("VIDmoney",  "Game_Mode_Mr Crabs.gif")   : title.SetBounds 0, 0, 150, 44 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title
  Set title = FlexDMD.NewVideo("VIDtikiland", "Game_Mode_Tiki Land.gif")    : title.SetBounds 0, 0, 128, 44 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title
  Set title = FlexDMD.NewVideo("VIDboating",  "Game_Mode_Boating School.gif") : title.SetBounds 0, 0, 200, 33 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title
  Set title = FlexDMD.NewVideo("VIDsandys", "Game_Mode_Sandys Rodeo.gif") : title.SetBounds 0, 0, 128, 44 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title
  Set title = FlexDMD.NewVideo("VIDmememe", "spongemememe.gif")       : title.SetBounds 0, 0, 128, 44 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title
  Set title = FlexDMD.NewVideo("VIDhouse",  "spongehouse.gif")        : title.SetBounds 0, 0, 128, 44 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title
  Set title = FlexDMD.NewVideo("VIDdance",  "Sponge_Dance.gif")       : title.SetBounds 0, 0, 128, 32 : title.Visible = False : title.SetAlignedPosition 64, 12 , FlexDMD_Align_Center : scene2.AddActor title



  Set title = FlexDMD.NewImage("ebroll",    "spongeball.png") : title.Visible = False : scene2.AddActor title


  Set title = FlexDMD.NewImage("BGtilt",    "SquareBGtilt.png") : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewImage("BGwarn",    "SquareBGwarn.png") : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title


  Set title = FlexDMD.NewLabel("TextSmalLine1", FontSponge16 , " ") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("TextSmalLine2", FontSponge16 , " ") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("TextSmalLine3", FontSponge16 , " ") : title.Visible = False : scene2.AddActor title
'inlane eb
  Set title = FlexDMD.NewImage("BG014",   "SquareBG014.png")  : title.SetBounds 0, 0, 128, 32 : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("eb1", FontSponge32 , "E") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("eb2", FontSponge32 , "X") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("eb3", FontSponge32 , "T") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("eb4", FontSponge32 , "R") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("eb5", FontSponge32 , "A") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("eb6", FontSponge32 , "BALL")  : title.Visible = False : scene2.AddActor title

'highscores use their own
  Set title = FlexDMD.NewImage("BG003",   "SquareBG002.png")  : title.SetAlignedPosition 0, 0 , FlexDMD_Align_Topleft : title.Visible = false : scene2.AddActor title

  Set title = FlexDMD.NewLabel("hs1", FontSponge16bb , " ") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("hs2", FontSponge16bb , " ") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("hs3", FontSponge16bb , " ") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("hs4", FontSponge16bb , " ") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("hs5", FontSponge16bb , " ") : title.Visible = False : scene2.AddActor title
  Set title = FlexDMD.NewLabel("hs6", FontSponge16bb , " ") : title.Visible = False : scene2.AddActor title





  FlexDMD.LockRenderThread
  FlexDMD.RenderMode =  FlexDMD_RenderMode_DMD_RGB
  FlexDMD.Stage.RemoveAll
  FlexDMD.Stage.AddActor scene
  FlexDMD.Stage.AddActor scene1
  FlexDMD.Stage.AddActor scene2
  If VRRoom = 0 And FlexOnFlasher = 0 Then FlexDMD.Show = True Else FlexDMD.Show = False
  FlexDMD.Run = True
  FlexDMD.UnlockRenderThread

End Sub

Dim Bubbles : Bubbles = 0
Dim bubblespos : bubblespos = 0
Dim Slinger : Slinger = 0
Dim Underfire : Underfire = 0
Dim FirePos : FirePos = 0
Dim FireTime : FireTime = 0
Dim FireImg : FireImg = 0

Dim FlexMsgTime
Dim FlexMsg
Dim FlexText1
Dim FlexText2
Dim FlexText3
Dim frame
Dim StartupScreen : StartupScreen = 1
Dim AttractCount : AttractCount = 0
Dim spongeshaker2 ' scoop
Dim alien6shake2 ' right
Dim alien14shake2 ' burger
Dim alien5shake2 ' left

'InsertSequence.Play SeqCircleOutOn,50,2
Dim Attractlightsequence
Sub InsertSequence_playdone
  If StartGame = 0 Then Attractlightsequence = 2
  attractspeed = int(rnd(1)*2) + 1
End Sub


Dim WizardStart : WizardStart = 0
Dim WizardPos : WizardPos = 0
Dim WizardDir : WizardDir = 1
Dim WizardShake : WizardShake = 0
Dim DodAnim : DodAnim = 0
Dim DodPos  : DodPos = 0
Dim DodGoal : DodGoal = 0
Dim DodFeat : DodFeat = 0
Dim DodWait : DodWait = 0
Dim DodNext : Dodnext = 0
Dim WizText(7)
Dim SmalScoring(10)
Dim Smalscorecounter
Dim smalline7
Dim smalline8
Dim smalline9
Dim smalline10
Dim nextsmalline : nextsmalline = 7
Dim WizardTimer : WizardTimer = 0
Dim FreshWizSwitch
Dim DampenMusic2 : DampenMusic2 = 1
Dim miniscoreblinks
Dim attractcredit
Sub DMDTimer_Timer
  dim tmp

  If frame mod 180 = 1 Then
    Dim GameModeStr : GameModeStr = "NA:"
    If Not bTilted Then
      Select Case CurrentMode
        case 1  : GameModeStr= "NA{yellow}:SPONGE"
        case 2  : GameModeStr= "NA{yellow}:PLANKTON"
        case 3  : GameModeStr= "NA{yellow}:RODEO"
        case 4  : GameModeStr= "NA{yellow}:MR KRAB"
        case 6  : GameModeStr= "NA{yellow}:SCHOOL"
        case 5  : GameModeStr= "NA{yellow}:JELLYFISH"
        case 7  : GameModeStr= "NA{yellow}:TIKILAND"
        case 8  : GameModeStr= "NA{yellow}:GOOFY"
        case 9  : GameModeStr= "NA{yellow}:FORMULA":
        case 11 : GameModeStr= "NA{red}:MULTIBALL"
        case 20 : GameModeStr= "NA{green}:WIZARD 1"
        case 21 : GameModeStr= "NA{green}:WIZARD 2"
      End Select
      If DoOrDieMode = 4 Then GameModeStr= GameModeStr + ";NA{red}:DoOrDie"
    End If
    If Not IsNull(Scorbit) Then If Scorbit.bSessionActive=True then Scorbit.SetGameMode(GameModeStr)
  End If




  If VRroom = 1 And ScoopRestart < frame And ScoopRestart > 0  then EndBOBAnim = 1'setup VR Bob to start leaving

  If EndBOBAnim = 1 Then
    If BobAnimation > 5 Or BobAnimation = 0 Then
      EndBOBAnim = 0

    Elseif BobAnimation = 5 Then
      EndBOBAnim = 0
      BobAnimation = 6
      BobWalking.PlayAnimEndless(0.09)
      ClothesWalking.PlayAnimEndless(0.09)
      BobJumping.StopAnim()
      ClothesJumping.StopAnim()
      BobWalking.visible = true
      ClothesWalking.visible = true
      BobJumping.visible = false
      ClothesJumping.visible = false
      ClothesWalking.image = "EyesClosed2"
      ScoopRestart = 0
    End If
  End if
  If SpinnerHurryUp > 0 And frame mod 66 = 1 Then ShakePatty=1 : PlaySoundAt "AlienShake2", alien6_BM_Lit_Room : ToyShake : noshakebigone = 1

  If FlexOnFlasher = 1 Then DMD_Flasher

  If frame mod 3 = 1 And DampenMusic > DampenMusic2 Then DampenMusic = DampenMusic - 0.1 : If MusicPlaying <> - 1 Then PlayTableMusic(MusicPlaying)
  If frame mod 6 = 1 And DampenMusic < DampenMusic2 Then DampenMusic = DampenMusic + 0.1 : If MusicPlaying <> - 1 Then PlayTableMusic(MusicPlaying)

'Debug.print "currentmsgdone=" & CurrentMSGdone

  If RampCombo < frame And Rampcombo <> 0 Then RampCombo = 0 : ComboCount = 0 : RampComboSelect = 0 : UpdateLights


  If Ballmustwait < Frame And BallsToShoot > 0 Then ' drain kickers fast timers
    If btilted And WizardWaitForNoBalls Then
      BallsToShoot = BallsToShoot - 1
      wizBiP = wizBiP -1
      If wizBiP < 1 Then WizardWaitForNoBalls = 0
'     debug.print "DMDtimer ballsetoshoot=0 tilted"
    Else
      Ballmustwait = Frame + 50
      BallsToShoot = BallsToShoot - 1

      If DrainSave = 0 Then
        Call RandomKicker(4)
        DrainSave = 1
      Else
        Call RandomKicker(3)
      End If
    End If
  End If
  If BallInLane > 0 And Ballmustwait < frame And BallsToShoot = 0 And ssREADY = 0 Then
    Auto_Plunger
    Ballmustwait = Frame + 50
  End If


  If Attractlightsequence > 0 Then attractmodeseq: DOF 121,1

  tmp = 0
  If spongeshaker2 <> 0 Then
    tmp = 1
    If spongeshaker2 < 0 Then spongeshaker2 = ABS(spongeshaker2) - 1 Else spongeshaker2 = - spongeshaker2 + 1
  End If
  If alien5shake2 <> 0 Then
    tmp = 1
    If alien5shake2 < 0 Then alien5shake2 = ABS(alien5shake2) - 1 Else alien5shake2 = - alien5shake2 + 1
  End If
  If alien6shake2 <> 0 Then
    tmp = 1
    If alien6shake2 < 0 Then alien6shake2 = ABS(alien6shake2) - 1 Else alien6shake2 = - alien6shake2 + 1
  End If
  If alien14shake2 <> 0 Then
    tmp = 1
    If alien14shake2 < 0 Then alien14shake2 = ABS(alien14shake2) - 1 Else alien14shake2 = - alien14shake2 + 1
  End If
  If tmp = 1 Then CharactersMovableHelper


  Dim i, n, x, y, label
  Dim title, af,list



  Frame=Frame+1

  FlexDMD.LockRenderThread



  If attractcredit > 0 Then
    attractcredit = attractcredit + 1
    Select Case attractcredit
      case 2
        FlexDMD.Stage.Addactor FlexDMD.NewImage("addcredit",    "SquareBG013.png")
        FlexDMD.Stage.Addactor FlexDMD.Newlabel("coin1",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(222,166,11), RGB(11,11,11),1)  , "CREDITS")
        FlexDMD.stage.GetLabel("coin1").SetAlignedPosition 17,11, FlexDMD_Align_Topleft
        tmp = Credits : If tmp > 10 Then tmp = 10
        FlexDMD.Stage.Addactor FlexDMD.Newlabel("coin2",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(222,166,11), RGB(11,11,11),1)  , tmp )
        FlexDMD.stage.GetLabel("coin2").SetAlignedPosition 92,11, FlexDMD_Align_Topleft
      case 10,30,50,70,90,110
        tmp = Credits : If tmp > 10 Then tmp = 10
        FlexDMD.stage.GetLabel("coin2").text =  tmp
        FlexDMD.stage.GetLabel("coin2").visible = False
      case 3,20,40,60,80,120
        tmp = Credits : If tmp > 10 Then tmp = 10
        FlexDMD.stage.GetLabel("coin2").text =  tmp
        FlexDMD.stage.GetLabel("coin2").visible = True
      case 130
        FlexDMD.stage.GetLabel("coin1").Remove
        FlexDMD.stage.GetLabel("coin2").Remove
        FlexDMD.stage.GetImage("addcredit").Remove
        attractcredit = 0
    End Select
  End If

  If attractcreditNO > 0 Then
    attractcreditNO = attractcreditNO + 1
    Select Case attractcreditNO
      case 2
        FlexDMD.Stage.Addactor FlexDMD.NewImage("addcredit",    "SquareBG013.png")
        FlexDMD.Stage.Addactor FlexDMD.Newlabel("coin1",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(222,166,11), RGB(11,11,11),1)  , "INSERT COIN")
        FlexDMD.stage.GetLabel("coin1").SetAlignedPosition 10,11, FlexDMD_Align_Topleft
      case 10,20,30,40,50,60,70,80,90
        FlexDMD.stage.GetLabel("coin1").visible = False
      case 5,15,25,35,45,55,65,75,85,95
        FlexDMD.stage.GetLabel("coin1").visible = True
      case 100
        FlexDMD.stage.GetLabel("coin1").Remove
        FlexDMD.stage.GetImage("addcredit").Remove
        attractcreditNO = 0
    End Select
  End If

  If QRClaim_Time > 0 Then
    QRClaim_Time = QRClaim_Time + 1
    '60 = 1s 4sec = 240
    If QRClaim_Time < 250 Then FlexDMD.Stage.GetLabel("MidLine").text = "HOLD FOR QRCLAIM " & INT(QRClaim_Time/60)  & "/4 SEC"

    If ScorbitActive <> 1 or isnull(scorbit) then
      FlexDMD.Stage.GeTLabel("MidLine").text = "NEED PAIRING IN OPTIONS"
    Elseif CurrentBall <> 3 or StartGame = 0 Then
      FlexDMD.Stage.GetLabel("MidLine").text = "QRCLAIM ONLY ON BALL 1"
    Elseif noQRimage Then
      FlexDMD.Stage.GetLabel("MidLine").text = "QR IMAGE NOT FOUND"
      FlexDMD.Stage.GetLabel("MidLine").SetAlignedPosition 64, 16, FlexDMD_Align_Center
    Else
      If QRClaim_Time = 240 Then InitFlexScorbitClaimDMD
    End If
    FlexDMD.Stage.GetLabel("MidLine").SetAlignedPosition 64, 16, FlexDMD_Align_Center

  End If

  If bInOptions = True Then Options_UpdateDMD

  If startupscreen > 0 Then
    If frame = 2 then FlexDMD.Stage.GetImage("Title").Visible = True
    If frame = 250 then startupscreen = 2 : FlexDMD.Stage.GetImage("Title").Visible = False : FlexDMD.Stage.GetImage("BG012").Visible = True
    Select Case startupscreen
      Case 1 : FlexDMD.UnLockRenderThread : Exit Sub
      Case 2 : AttractModeUpdate
           FlexDMD.UnLockRenderThread
           Exit Sub
      Case 3 : startupscreen = 0 : AttractCount = 0 :

          If AttractvideoON then FlexDMD.stage.GetVideo("VIDattr").remove : AttractvideoON = False
          FlexDMD.Stage.GetImage("BG012").visible = false

          FlexDMD.UnLockRenderThread : Exit Sub

      Case 4 :
          startupscreen = 2 : AttractCount = 90 + 52 * 35 : FlexDMD.UnLockRenderThread :  Exit sub

    End Select
  End If

  If WizardShake <> 0 And Frame mod 3 = 1 Then
    If WizardShake < 0 Then WizardShake = ABS(WizardShake) - 1 Else WizardShake = - WizardShake + 1
  End If

  If WizardStart > 0 Then DoTheWizIntro

  If SkillShot = 1 And DMDDisplay(0,0) <> "" Then DMD_Display_Timer

  If DoEball > 0 Then
    If DoEball = 1 Then FlexDMD.Stage.Getimage("ebroll").visible = true
    DoEball = DoEball + 1
    FlexDMD.Stage.Getimage("ebroll").SetAlignedPosition -70 + DoEball * 2, 15 , FlexDMD_Align_Center
    If DoEball > 100 Then FlexDMD.Stage.Getimage("ebroll").visible = false : DoEball = 0
  End If

  If Bubbles > 0 And Underfire = 0 Then
    If Bubbles = 1 Then  bubblespos = -64 : FlexDMD.Stage.Getimage("bubbles").visible = true
    Bubbles = Bubbles + 1
    bubblespos = bubblespos + Int(rnd(1)*3)-1

    If bubbles > 287 Then
      bubbles = 0 : FlexDMD.Stage.Getimage("bubbles").visible = False
    Else
      FlexDMD.Stage.Getimage("bubbles").SetAlignedPosition bubblespos, 33 - bubbles , FlexDMD_Align_TopLeft
    End If
  Elseif Underfire > 0 Then
    If Underfire = 1 Then FirePos = 30
    Underfire = 2
    If FireTime < 1 Then
      FlexDMD.Stage.Getimage("fire1").visible = False
      FlexDMD.Stage.Getimage("fire2").visible = False
      FlexDMD.Stage.Getimage("fire3").visible = False
      FlexDMD.Stage.Getimage("fire4").visible = False
      FlexDMD.Stage.Getimage("fire5").visible = False
      FlexDMD.Stage.Getimage("fire6").visible = False
      FlexDMD.Stage.Getimage("fire7").visible = False
      Underfire = 0
    Else
      If frame mod 3 = 1 Then
        FireTime = FireTime - 1
        If FireTime > 20 And FirePos >  0 Then FirePos = FirePos - 1
        If FireTime < 21 And FirePos < 32 Then FirePos = FirePos + 1
        Select Case FireImg
          Case 0 : Fireimg = 1 : FlexDMD.Stage.Getimage("fire1").visible = True : FlexDMD.Stage.Getimage("fire7").visible = False
          Case 1 : Fireimg = 2 : FlexDMD.Stage.Getimage("fire2").visible = True : FlexDMD.Stage.Getimage("fire1").visible = False
          Case 2 : Fireimg = 3 : FlexDMD.Stage.Getimage("fire3").visible = True : FlexDMD.Stage.Getimage("fire2").visible = False
          Case 3 : Fireimg = 4 : FlexDMD.Stage.Getimage("fire4").visible = True : FlexDMD.Stage.Getimage("fire3").visible = False
          Case 4 : Fireimg = 5 : FlexDMD.Stage.Getimage("fire5").visible = True : FlexDMD.Stage.Getimage("fire4").visible = False
          Case 5 : Fireimg = 6 : FlexDMD.Stage.Getimage("fire6").visible = True : FlexDMD.Stage.Getimage("fire5").visible = False
          Case 6 : Fireimg = 0 : FlexDMD.Stage.Getimage("fire7").visible = True : FlexDMD.Stage.Getimage("fire6").visible = False
        End Select
        FlexDMD.Stage.Getimage("fire1").SetAlignedPosition 0, Firepos+1 , FlexDMD_Align_TopLeft
        FlexDMD.Stage.Getimage("fire2").SetAlignedPosition 0, Firepos+1 , FlexDMD_Align_TopLeft
        FlexDMD.Stage.Getimage("fire3").SetAlignedPosition 0, Firepos+1 , FlexDMD_Align_TopLeft
        FlexDMD.Stage.Getimage("fire4").SetAlignedPosition 0, Firepos+1 , FlexDMD_Align_TopLeft
        FlexDMD.Stage.Getimage("fire5").SetAlignedPosition 0, Firepos+1 , FlexDMD_Align_TopLeft
        FlexDMD.Stage.Getimage("fire6").SetAlignedPosition 0, Firepos+1 , FlexDMD_Align_TopLeft
        FlexDMD.Stage.Getimage("fire7").SetAlignedPosition 0, Firepos+1 , FlexDMD_Align_TopLeft
      End If
    End If
  End If

  If Slinger > 0 Then
    Slinger = Slinger + 1
    Select Case Slinger
      Case 2  : FlexDMD.Stage.Getimage("sling1").visible = True
      Case 4  : FlexDMD.Stage.Getimage("sling2").visible = True  : FlexDMD.Stage.Getimage("sling1").visible = False
      Case 6  : FlexDMD.Stage.Getimage("sling3").visible = True  : FlexDMD.Stage.Getimage("sling2").visible = False
      Case 8  : FlexDMD.Stage.Getimage("sling4").visible = True  : FlexDMD.Stage.Getimage("sling3").visible = False
      Case 10 : FlexDMD.Stage.Getimage("sling5").visible = True  : FlexDMD.Stage.Getimage("sling4").visible = False
    End Select
    If Slinger > 11 Then Slinger = 0 : FlexDMD.Stage.Getimage("sling5").visible = False
  End If
  If bumperjackpotcount > 0 And bumperjackpot = 0 Then bumperjackpot = 1
  If bumperjackpot > 0 Then
    bumperjackpot = bumperjackpot + 1
    Select case bumperjackpot
      case 2 : FlexDMD.Stage.Getimage("five1").visible = True
      case 3 : FlexDMD.Stage.Getimage("five2").visible = True : FlexDMD.Stage.Getimage("five1").visible = False
      case 4 : FlexDMD.Stage.Getimage("five3").visible = True : FlexDMD.Stage.Getimage("five2").visible = False
      case 5 : FlexDMD.Stage.Getimage("five4").visible = True : FlexDMD.Stage.Getimage("five3").visible = False
      case 6 : FlexDMD.Stage.Getimage("five5").visible = True : FlexDMD.Stage.Getimage("five4").visible = False
    End Select
    If bumperjackpot > 6 Then
      bumperjackpot = 0
      FlexDMD.Stage.Getimage("five5").visible = False
      bumperjackpotcount = bumperjackpotcount - 1
    End If
  End If

  If DoOrDieMode > 2 Then
    FlexDMD.Stage.GetLabel("Ball").Text = "One Ball"
  Elseif CurrentBall = 0 Then
    FlexDMD.Stage.GetLabel("Ball").Text = "Ball 3"
  Else
    FlexDMD.Stage.GetLabel("Ball").Text = "Ball "& 4 - CurrentBall 'BallInPlay
  End If

  If UseCredits Then
    tmp = Credits : If tmp > 10 Then tmp = 10
    FlexDMD.Stage.GetLabel("Credit").Text = "Credits " & tmp
  Else
    FlexDMD.Stage.GetLabel("Credit").Text = "FreePlay" 'Credits go here if added to game
  End If


  If DoOrDieMode > 2 Then
    DodAnim = DodAnim + 1
    Select Case DodAnim
      Case  2 : FlexDMD.Stage.Getimage("dod1").visible = True : FlexDMD.Stage.Getimage("dod3").visible = False
            FlexDMD.Stage.Getimage("dod1").setposition 0,DodPos
      Case  6 : FlexDMD.Stage.Getimage("dod2").visible = True : FlexDMD.Stage.Getimage("dod1").visible = False
            FlexDMD.Stage.Getimage("dod2").setposition 0,DodPos
      Case 10 : FlexDMD.Stage.Getimage("dod3").visible = True : FlexDMD.Stage.Getimage("dod2").visible = False
            FlexDMD.Stage.Getimage("dod3").setposition 0,DodPos
      Case 14 : FlexDMD.Stage.Getimage("dod2").visible = True : FlexDMD.Stage.Getimage("dod3").visible = False
            FlexDMD.Stage.Getimage("dod2").setposition 0,DodPos
      Case 16 : DodAnim = 0
    End Select

    If DodNext = 0 Then
      Select Case DodFeat
        Case 1 :          DodGoal =  0  : DodNext = 0     ' 1 goto top and stop !         startball
        Case 2 :          DodGoal = -9  : DodNext = 3       ' 2 middle surfing            gameover ?
        Case 3 : DodWait =  3  :  DodGoal = -15 : DodNext = 4
        Case 4 : DodWait =  11 :  DodGoal = -7  : DodNext = 3
        Case 5 :          DodGoal =  0  : DodNext = 6     ' 5 goto top wait 111 goto surfonce   topramps+startmode
        Case 6 : DodWait = 111 :  DodGoal = -9  : DodNext = 9
        Case 7 :          DodGoal = -15 : DodNext = 8     ' 7 goto middle and stop        scoop enter
        Case 8 : DodWait =  55 :  DodGoal = -12 : DodNext = 0
        Case 9 :          DodGoal = -9  : DodNext = 10      ' 9 middle surfing once and Stop    addscore exit scoop
        Case 10: DodWait =  3  :  DodGoal = -15 : DodNext = 11
        Case 11: DodWait =  11 :  DodGoal = -7  : DodNext = 0
      End Select
    End If
    If DodWait > 0 Then
      DodWait = DodWait -1
    Else
        If DodPos > DodGoal Then DodPos = DodPos - 1
        If DodPos < DodGoal Then DodPos = DodPos + 1
        If DodPos = DodGoal Then DodFeat = DodNext : DodNext = 0
    End If
  Else
    For i = 1 to Players
      Set label = FlexDMD.Stage.GetLabel("Score_" & i)
      If Endofgamescoreblink = 1 Then
        If frame mod 30 > 15 Then
          label.Font = FontScoreInactive
          FlexDMD.Stage.GetLabel("comma1").font = FontScoreInactive
          FlexDMD.Stage.GetLabel("comma1b").font = FontScoreInactive
          FlexDMD.Stage.GetLabel("comma1c").font = FontScoreInactive
        Else
          label.Font = FontScoreInactive2
          FlexDMD.Stage.GetLabel("comma1").font = FontScoreInactive2
          FlexDMD.Stage.GetLabel("comma1b").font = FontScoreInactive2
          FlexDMD.Stage.GetLabel("comma1c").font = FontScoreInactive2
        End If
      Else
        If i = currentplayer Then
          label.Font = FontScoreActive
          FlexDMD.Stage.GetLabel("comma1").font = FontScoreactive
          FlexDMD.Stage.GetLabel("comma1b").font = FontScoreactive
          FlexDMD.Stage.GetLabel("comma1c").font = FontScoreactive
        Else
          label.Font = FontScoreInactive
        End If
      End If
      label.SetAlignedPosition 40, 1 + (i - 1) * 8, FlexDMD_Align_TopRight
      label.Text = totalscore(i)

      If miniscoreblinks > 0 And frame mod 3 = 1 Then miniscoreblinks = miniscoreblinks - 1
      Select Case miniscoreblinks
        case 1,3,5,7,9,11,13
          FlexDMD.Stage.GetLabel("Score_" & i).visible = False
          FlexDMD.Stage.GetLabel("comma1").visible = False
          FlexDMD.Stage.GetLabel("comma1b").visible = False
          FlexDMD.Stage.GetLabel("comma1c").visible = False
        case 0,2,4,6,8,10,12
          FlexDMD.Stage.GetLabel("Score_" & i).visible = True
          FlexDMD.Stage.GetLabel("comma1").SetAlignedPosition  30, 2+ (i - 1) * 8 , FlexDMD_Align_TopRight
          FlexDMD.Stage.GetLabel("comma1b").SetAlignedPosition 18, 2+ (i - 1) * 8 , FlexDMD_Align_TopRight
          FlexDMD.Stage.GetLabel("comma1c").SetAlignedPosition  6, 2+ (i - 1) * 8 , FlexDMD_Align_TopRight
            FlexDMD.Stage.GetLabel("comma1").visible = False
            FlexDMD.Stage.GetLabel("comma1b").visible = False
            FlexDMD.Stage.GetLabel("comma1c").visible = False
            If TotalScore(i) > 999 Then FlexDMD.Stage.GetLabel("comma1").visible = True Else FlexDMD.Stage.GetLabel("comma1").visible = False
            If TotalScore(i) > 999999 Then FlexDMD.Stage.GetLabel("comma1b").visible = True Else FlexDMD.Stage.GetLabel("comma1b").visible = False
            If TotalScore(i) > 999999999 Then FlexDMD.Stage.GetLabel("comma1c").visible = True Else FlexDMD.Stage.GetLabel("comma1c").visible = False
      End Select
    Next
  End If


  If SmalScoring(0) > 0 And Frame mod 16 = 1 Then
    Set label = FlexDMD.Stage.GetLabel("TextSmalLine" & nextsmalline)
    label.Font = FontScoreInactive
    label.Text = FormatScore( SmalScoring(0) )
    label.SetAlignedPosition 39, 23 , FlexDMD_Align_Center
    label.visible = True
    Select Case nextsmalline
      case 7 : nextsmalline = 8 : smalline7 = 1
      case 8 : nextsmalline = 9 : smalline8 = 1
      case 9 : nextsmalline = 10 : smalline9 = 1
      case 10: nextsmalline = 7 : smalline10 = 1
    End Select

    For x = 0 To 9
      SmalScoring(x) = SmalScoring(x+1)
    Next
    SmalScoring(10) = 0
  End If

  If smalline7 > 0 And frame mod 3 = 1 Then
    smalline7 = smalline7 + 1
    If smalline7 > 30 Then
      smalline7 = 0 : FlexDMD.Stage.GetLabel("TextSmalLine7").visible = False
    Else
      Set label = FlexDMD.Stage.GetLabel("TextSmalLine7")
      label.font =        FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(64+smalline7*2, 64-smalline7*2, 64-smalline7*2), RGB(0, 0, 0), 0)
      label.SetAlignedPosition 39, 22+smalline7 , FlexDMD_Align_Center
    End If
  End If
  If smalline8 > 0 And frame mod 3 = 1 Then
    smalline8 = smalline8 + 1
    If smalline8 > 30 Then
      smalline8 = 0 : FlexDMD.Stage.GetLabel("TextSmalLine8").visible = False
    Else
      Set label = FlexDMD.Stage.GetLabel("TextSmalLine8")
      label.font =        FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(64+smalline8, 64-smalline8*2, 64-smalline8*2), RGB(0, 0, 0), 0)
      label.SetAlignedPosition 39, 22+smalline8 , FlexDMD_Align_Center
    End If
  End If
  If smalline9 > 0 And frame mod 3 = 1 Then
    smalline9 = smalline9 + 1
    If smalline9 > 30 Then
      smalline9 = 0 : FlexDMD.Stage.GetLabel("TextSmalLine9").visible = False
    Else
      Set label = FlexDMD.Stage.GetLabel("TextSmalLine9")
      label.font =        FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(64+smalline9, 64-smalline9*2, 64-smalline9*2), RGB(0, 0, 0), 0)
      label.SetAlignedPosition 39, 22+smalline9 , FlexDMD_Align_Center
    End If
  End If
  If smalline10 > 0 And frame mod 3 = 1 Then
    smalline10 = smalline10 + 1
    If smalline10 > 30 Then
      smalline10 = 0 : FlexDMD.Stage.GetLabel("TextSmalLine10").visible = False
    Else
      Set label = FlexDMD.Stage.GetLabel("TextSmalLine10")
      label.font =        FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(64+smalline10, 64-smalline10*2, 64-smalline10*2), RGB(0, 0, 0), 0)
      label.SetAlignedPosition 39, 22+smalline10 , FlexDMD_Align_Center
    End If
  End If


  If Endofgamescoreblink = 1 Or scoreblink > 0 Then
    If frame mod 30 > 15 Then
      FlexDMD.Stage.GetLabel("Content_1").Font = fontbig4
      If scoreblink = 1 Then scoreblink = 2
    Else
      FlexDMD.Stage.GetLabel("Content_1").Font = fontbig3
      If scoreblink = 2 Then scoreblink = 0
    End If
  Else
    FlexDMD.Stage.GetLabel("Content_1").Font = fontbig3
  End If

  FlexDMD.Stage.GetLabel("Content_1").Text = formatscore(totalscore(currentplayer))
  FlexDMD.Stage.GetLabel("Content_1").SetAlignedPosition 45, 21 , FlexDMD_Align_Center


  If WizardTimer = 0 Then
    FlexDMD.Stage.GetLabel("TextSmalLine11").visible = False
  Else
    If Frame mod 8 = 1 Then
      Set label = FlexDMD.Stage.GetLabel("TextSmalLine11")
      If FreshWizSwitch <> WizardSwitches Then
        FreshWizSwitch = WizardSwitches
        label.Font = FontScoreActive
      Else
        label.Font = FontScoreinActive
      End If
      If CurrentMode = 20 Then
        label.Text =  WizardSwitches & "/" & "240"
      Else
        label.Text =  WizardSwitches & "/" & "18"
      End If
      label.SetAlignedPosition 2,1,FlexDMD_Align_topleft
      label.visible = True
    End If
  End If


  If SpinnerTimer > 0 Then
    Set label = FlexDMD.Stage.GetLabel("TextSmalLine5")
    label.Font = FontScoreinActive
    label.Text =  "HURRYUP " & SpinnerTimer & " s"
    label.SetAlignedPosition 85,7,FlexDMD_Align_topright
    label.visible = True
  Else
    FlexDMD.Stage.GetLabel("TextSmalLine5").visible = False ' If CurrentMode = 0 Or CurrentMode > 9 then
  End If
  If CurrentMode > 0 Then ' And CurrentMode < 10
    If frame mod 150 > 75 Or SpinnerTimer < 1 Then
      Set label = FlexDMD.Stage.GetLabel("TextSmalLine5")
      label.Font = FontScoreinActive
      label.Text =  "JP LVL = " & HurryUpCollected(CurrentPlayer)
      label.SetAlignedPosition 85,7,FlexDMD_Align_topright
      label.visible = True
    End If
  Else
    If SpinnerTimer < 1 Then FlexDMD.Stage.GetLabel("TextSmalLine5").visible = False
  End If


  If WizardTimer > 0 Then
    Set label = FlexDMD.Stage.GetLabel("TextSmalLine6")
    label.Font = FontScoreinActive
    label.Text =  "TIME "& WizardTimer & " s"
    label.SetAlignedPosition 85,1,FlexDMD_Align_topright
    label.visible = True
  Elseif modetimecounter > 0 Then
    Set label = FlexDMD.Stage.GetLabel("TextSmalLine6")
    label.Font = FontScoreinActive
    label.Text =  "MODE "& modetimecounter &" s"
    label.SetAlignedPosition 85,1,FlexDMD_Align_topright
    label.visible = True
  Elseif DoOrDieMode > 2 then
    Set label = FlexDMD.Stage.GetLabel("TextSmalLine6")
    label.Font = FontScoreinActive
    label.Text =  "DO OR DIE"
    label.SetAlignedPosition 85,1,FlexDMD_Align_topright
    label.visible = True
  Else
    FlexDMD.Stage.GetLabel("TextSmalLine6").visible = False
  End If



  If FlexMsg > 1 Then FlexMsg = FlexMsg + 1 : DMD_Update
  If FlexMsg = 1 Then DMD_ResetAll : DMD_NewOverlay

  If FlexMsg > 0 And FlexMsgTime < frame Then DMD_SlideOff


  If EnterHighscore > 0 Then

    Select Case EnterHighscore
      Case 1                                'startup hs
        PlaySound "HSMusic",0,RomSoundVolume : HSMusicPlaying.Enabled = True
        FlexDMD.Stage.GetImage("BG003").Visible = True          'background

        FlexDMD.Stage.GetLabel("hs1").font = FontSponge12bb       'line1
        FlexDMD.Stage.GetLabel("hs1").text = "PLAYER " & EnterPlayerNR
        FlexDMD.Stage.GetLabel("hs1").SetAlignedPosition 66, 10 , FlexDMD_Align_Center
        FlexDMD.Stage.GetLabel("hs1").Visible = True

        FlexDMD.Stage.GetLabel("hs2").font = FontSponge12bb       'line2
        FlexDMD.Stage.GetLabel("hs2").text = "ENTER INITIALS"
        FlexDMD.Stage.GetLabel("hs2").SetAlignedPosition 66, 26 , FlexDMD_Align_Center
        FlexDMD.Stage.GetLabel("hs2").Visible = True
      case 10,25,40,55,70,80  : FlexDMD.Stage.GetLabel("hs1").Visible = True            'blinking
      case 15,30,45,60,75,90  : FlexDMD.Stage.GetLabel("hs1").Visible = False
      case 95         : FlexDMD.Stage.GetLabel("hs2").Visible = False
      case 101                                          'enter letterstart
        EnterHighblinking = 1
        EnterHighPos = 20
        EnterText(1) = ""
        EnterText(2) = ""
        EnterText(3) = ""

        FlexDMD.Stage.GetLabel("hs1").font = FontSponge16bb
        FlexDMD.Stage.GetLabel("hs1").text = " "
        FlexDMD.Stage.GetLabel("hs1").SetAlignedPosition 43,  7 , FlexDMD_Align_Center
        FlexDMD.Stage.GetLabel("hs1").Visible = True

        FlexDMD.Stage.GetLabel("hs2").font = FontSponge16bb
        FlexDMD.Stage.GetLabel("hs2").text = "-"
        FlexDMD.Stage.GetLabel("hs2").SetAlignedPosition 60,  7 , FlexDMD_Align_Center
        FlexDMD.Stage.GetLabel("hs2").Visible = True

        FlexDMD.Stage.GetLabel("hs3").font = FontSponge16bb
        FlexDMD.Stage.GetLabel("hs3").text = " "
        FlexDMD.Stage.GetLabel("hs3").SetAlignedPosition 77,  7 , FlexDMD_Align_Center
        FlexDMD.Stage.GetLabel("hs3").Visible = True

        FlexDMD.Stage.GetLabel("hs4").font = FontSponge12bb
        FlexDMD.Stage.GetLabel("hs4").text = "W"
        FlexDMD.Stage.GetLabel("hs4").SetAlignedPosition 61,  27 , FlexDMD_Align_Center
        FlexDMD.Stage.GetLabel("hs4").Visible = True

        FlexDMD.Stage.GetLabel("hs5").font = FontScoreBlack
        FlexDMD.Stage.GetLabel("hs5").text = "ABCDEFGHIJKLMNOPQR"
        FlexDMD.Stage.GetLabel("hs5").SetAlignedPosition 54, 25 , FlexDMD_Align_TopRight
        FlexDMD.Stage.GetLabel("hs5").Visible = True

        FlexDMD.Stage.GetLabel("hs6").font = FontScoreBlack
        FlexDMD.Stage.GetLabel("hs6").text = "TUVWXYZ 0123456789"
        FlexDMD.Stage.GetLabel("hs6").SetAlignedPosition 66, 25 , FlexDMD_Align_TopLeft
        FlexDMD.Stage.GetLabel("hs6").Visible = True

      case 110
        Select Case EnterHighblinking
          Case 1
            FlexDMD.Stage.GetLabel("hs1").text = mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789",EnterHighPos+18,1)
            FlexDMD.Stage.GetLabel("hs1").font = FontSponge16bb
            FlexDMD.Stage.GetLabel("hs1").SetAlignedPosition 43,  7 , FlexDMD_Align_Center

          Case 2
            FlexDMD.Stage.GetLabel("hs2").text = mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789",EnterHighPos+18,1)
            FlexDMD.Stage.GetLabel("hs2").font = FontSponge16bb
            FlexDMD.Stage.GetLabel("hs2").SetAlignedPosition 60,  7 , FlexDMD_Align_Center
          Case 3
            FlexDMD.Stage.GetLabel("hs3").text = mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789",EnterHighPos+18,1)
            FlexDMD.Stage.GetLabel("hs3").font = FontSponge16bb
            FlexDMD.Stage.GetLabel("hs3").SetAlignedPosition 77,  7 , FlexDMD_Align_Center
        End Select
      case 118
        Select Case EnterHighblinking
          Case 1
            FlexDMD.Stage.GetLabel("hs1").font = FontBig
            FlexDMD.Stage.GetLabel("hs1").text = "_"
            FlexDMD.Stage.GetLabel("hs1").SetAlignedPosition 43,  15 , FlexDMD_Align_Center
          Case 2
            FlexDMD.Stage.GetLabel("hs2").font = FontBig
            FlexDMD.Stage.GetLabel("hs2").text = "_"
            FlexDMD.Stage.GetLabel("hs2").SetAlignedPosition 60,  15 , FlexDMD_Align_Center
          Case 3
            FlexDMD.Stage.GetLabel("hs3").font = FontBig
            FlexDMD.Stage.GetLabel("hs3").text = "_"
            FlexDMD.Stage.GetLabel("hs3").SetAlignedPosition 77,  15 , FlexDMD_Align_Center
        End Select

      case 125
        EnterHighscore = 109

      case 140,150,160,170,180,190,200,210,220,230,240
        FlexDMD.Stage.GetLabel("hs1").visible = False             'final blinking
        FlexDMD.Stage.GetLabel("hs2").visible = False
        FlexDMD.Stage.GetLabel("hs3").visible = False
      case 147,157,167,177,187,197,207,217,227,237,247
        FlexDMD.Stage.GetLabel("hs1").visible = True
        FlexDMD.Stage.GetLabel("hs2").visible = True
        FlexDMD.Stage.GetLabel("hs3").visible = True

      case 333,334
        FlexDMD.Stage.GetImage("BG003").Visible = False
        FlexDMD.Stage.GetLabel("hs1").visible = False
        FlexDMD.Stage.GetLabel("hs2").visible = False
        FlexDMD.Stage.GetLabel("hs3").visible = False
        FlexDMD.Stage.GetLabel("hs4").visible = False
        FlexDMD.Stage.GetLabel("hs5").visible = False
        FlexDMD.Stage.GetLabel("hs6").visible = False
        EnterHighscore = 0
        FlexDMD.UnLockRenderThread
        HighScores(ScoreNumber, 1) = EnterText(1) & EnterText(2) & EnterText(3)

        SaveHighScores
        ScoreNumber = ""
        highscoredelay.enabled = True

        Exit Sub

    End Select

    If EnterHighscore > 101 And EnterHighscore < 125 Then

      Select Case EnterHighblinking
        Case 1
'         FlexDMD.Stage.GetLabel("hs1").text = mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789",EnterHighPos+18,1)
        Case 2
          FlexDMD.Stage.GetLabel("hs1").text = Entertext(1)
          FlexDMD.Stage.GetLabel("hs1").font = FontSponge16bb
          FlexDMD.Stage.GetLabel("hs1").SetAlignedPosition 43,  7 , FlexDMD_Align_Center
'         FlexDMD.Stage.GetLabel("hs2").text = mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789",EnterHighPos+18,1)
        Case 3
          FlexDMD.Stage.GetLabel("hs2").text = Entertext(2)
          FlexDMD.Stage.GetLabel("hs2").font = FontSponge16bb
          FlexDMD.Stage.GetLabel("hs2").SetAlignedPosition 60,  7 , FlexDMD_Align_Center
'         FlexDMD.Stage.GetLabel("hs3").text = mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789",EnterHighPos+18,1)
        case 4
          FlexDMD.Stage.GetLabel("hs3").text = Entertext(3)
          FlexDMD.Stage.GetLabel("hs3").font = FontSponge16bb
          FlexDMD.Stage.GetLabel("hs3").SetAlignedPosition 77,  7 , FlexDMD_Align_Center
      End Select

      FlexDMD.Stage.GetLabel("hs4").text = mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789",EnterHighPos+18,1)
      FlexDMD.Stage.GetLabel("hs5").text = mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789",EnterHighPos,18)
      FlexDMD.Stage.GetLabel("hs6").text = mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789",EnterHighPos+19,18)
        ' 20 = "A"
        ' 37 = R    1 = S
      If EnterText(3) <> "" Then EnterHighscore = 130
    End If


    EnterHighscore = EnterHighscore + 1

  End If
  FlexDMD.UnLockRenderThread
End Sub

highscoredelay.interval = 2000
Sub highscoredelay_Timer
  highscoredelay.enabled = False
  ClearDMD
  StartNewPlayer
End Sub


Dim EnterHighscore
Dim EnterPlayerNR
Dim EnterHighblinking
Dim EnterHighPos
Dim EnterText(3)
EnterHighPos = 1

Dim EnterKey : EnterKey = 0

Sub EnterGoRight
  EnterHighPos = EnterHighPos + 1
  If EnterHighPos > 37 Then EnterHighPos = 1
End Sub

Sub EnterGoLeft
  EnterHighPos = EnterHighPos - 1
  If EnterHighPos < 1 Then EnterHighPos = 37
End Sub

Dim EnterRepeat
Sub RepeatTimer_Timer
  If EnterRepeat = 0 Then EnterRepeat = 1 : Exit Sub
  If EnterRepeat = 1 Then EnterRepeat = 2 : Exit Sub
  If EnterKey = 1 Then EnterGoLeft
  If EnterKey = 2 Then EnterGoRight
  SkillShotSfx
End Sub


Dim WizardSwitches
Sub WizardJPUpdate
  dim tmp
  LightSeqmodes.stopplay
  LightSeqmodes.play SeqBlinking,,2,20
  LightSeqmodes.play SeqBlinking,,1,30
  LightSeqmodes.play SeqBlinking,,1,40
  LightSeqmodes.play SeqBlinking,,1,50

  WizardSwitches = WizardSwitches + 1

  If CurrentMode = 20 Then
    Playsound "fx_Solenoidon",1,mechvolume

    Select Case WizardSwitches
    Case 30
      tmp = 10000000 + HurryUpCollected(CurrentPlayer) * 4
      addscore (tmp) : DMDMessage1 "JACKPOT"  ,FormatScore (tmp)      ,FontSponge12bb ,45,16  ,69,16  ,"shake"  ,2  ,"BG001"  ,"novideo"
      PlaySound "wiz2",0,RomSoundVolume
      DOF 132,2
      DOF 134,2 'shaker
    Case 60
      tmp = 10000000 + HurryUpCollected(CurrentPlayer) * 4
      addscore (tmp) : DMDMessage1 "JACKPOT"  ,FormatScore (tmp)      ,FontSponge12bb ,45,16  ,69,16  ,"shake"  ,2  ,"BG001"  ,"novideo"
      PlaySound "wiz1",0,RomSoundVolume
      DOF 132,2
      DOF 134,2 'shaker
    Case 90
      tmp = 10000000 + HurryUpCollected(CurrentPlayer) * 4
      addscore (tmp) : DMDMessage1 "JACKPOT"  ,FormatScore (tmp)      ,FontSponge12bb ,45,16  ,69,16  ,"shake"  ,2  ,"BG001"  ,"novideo"
      PlaySound "wiz2",0,RomSoundVolume
      DOF 132,2
      DOF 134,2 'shaker
    Case 120
      tmp = 15000000 + HurryUpCollected(CurrentPlayer) * 4
      addscore (tmp) : DMDMessage1 "JACKPOT"  ,FormatScore (tmp)      ,FontSponge12bb ,45,16  ,69,16  ,"shake"  ,2  ,"BG001"  ,"novideo"
      PlaySound "wiz1",0,RomSoundVolume
      DOF 132,2
      DOF 134,2 'shaker
    Case 150
      tmp = 15000000 + HurryUpCollected(CurrentPlayer) * 4
      addscore (tmp) : DMDMessage1 "JACKPOT"  ,FormatScore (tmp)      ,FontSponge12bb ,45,16  ,69,16  ,"shake"  ,2  ,"BG001"  ,"novideo"
      PlaySound "wiz2",0,RomSoundVolume
      DOF 132,2
      DOF 134,2 'shaker
    Case 180
      tmp = 15000000 + HurryUpCollected(CurrentPlayer) * 4
      addscore (tmp) : DMDMessage1 "JACKPOT"  ,FormatScore (tmp)      ,FontSponge12bb ,45,16  ,69,16  ,"shake"  ,2  ,"BG001"  ,"novideo"
      PlaySound "wiz2",0,RomSoundVolume
      DOF 132,2
      DOF 134,2 'shaker
    Case 210
      tmp = 25000000 + HurryUpCollected(CurrentPlayer) * 4
      addscore (tmp) : DMDMessage1 "JACKPOT"  ,FormatScore (tmp)      ,FontSponge12bb ,45,16  ,69,16  ,"shake"  ,2  ,"BG001"  ,"novideo"
      PlaySound "wiz2",0,RomSoundVolume
      DOF 132,2
      DOF 134,2 'shaker
    Case 240
      DOF 132,2
      DOF 134,2 'shaker
      StopTableMusic
      songtimer.enabled = False
      tmp = 50000000 + HurryUpCollected(CurrentPlayer) * 10
                     : DMDMessage1 "LEVEL 1"      ,"COMPLETE"   ,FontSponge12bb ,45,17  ,69,17  ,"solid"  ,2  ,"BG001"  ,"novideo"
      addscore (tmp) : DMDMessage1 FormatScore(tmp) ,""       ,FontSponge12bb ,45,17  ,69,17  ,"blink"  ,4  ,"BG001"  ,"novideo"
      PlaySound "wiz3",0,RomSoundVolume

      bTilted = True
      FlipperDeActivate LeftFlipper, LFPress : SolLFlipper False : SolULFlipper False: FlipperDeActivate RightFlipper, RFPress : SolRFlipper False
      PlaySound "Sfx_SuperJackpot",0,RomSoundVolume
      ResetModeLights
      UpdateLights
      EndWizard1.Interval = 7000
      EndWizard1.Enabled = True
      lightCtrl.AddTableLightSeq "GI", lSeqRainbow
      InsertSequence.play SeqBlinking,,11,150
      playtime = 100
      WizardTimer = 0
      WizardWaitForNoBalls = 1
      BonusRounds = BonusRounds + 1

    End Select
  Else
    ' currentmode 21
    SetGI Rgb(0,255,10),600
    DOF 132,2
    DOF 134,2 'shaker
    If WizardSwitches < 18 Then
      tmp = 14000000 + HurryUpCollected(CurrentPlayer) * 4 + WizardSwitches * 1000000
      addscore (tmp) : DMDMessage1 "JACKPOT"  ,FormatScore (tmp)      ,FontSponge12bb ,45,16  ,69,16  ,"shake"  ,2  ,"BG001"  ,"novideo"
      PlaySound "wiz2",0,RomSoundVolume

    Else
      StopTableMusic
      songtimer.enabled = False
      tmp = 50000000 + HurryUpCollected(CurrentPlayer) * 10
                     : DMDMessage1 "LEVEL 2"      ,"COMPLETE"   ,FontSponge12bb ,45,17  ,69,17  ,"solid"  ,2  ,"BG001"  ,"novideo"
      addscore (tmp) : DMDMessage1 FormatScore(tmp) ,""       ,FontSponge12bb ,45,17  ,69,17  ,"blink"  ,4  ,"BG001"  ,"novideo"
      PlaySound "wiz3",0,RomSoundVolume

      bTilted = True
      FlipperDeActivate LeftFlipper, LFPress : SolLFlipper False : SolULFlipper False: FlipperDeActivate RightFlipper, RFPress : SolRFlipper False
      PlaySound "Sfx_SuperJackpot",0,RomSoundVolume
      ResetModeLights
      UpdateLights
      EndWizard1.Interval = 7000
      EndWizard1.Enabled = True
      lightCtrl.AddTableLightSeq "GI", lSeqRainbow
      InsertSequence.play SeqBlinking,,11,150
      playtime = 100
      WizardTimer = 0
      WizardWaitForNoBalls = 1
      BonusRounds = BonusRounds + 1
    End If
  End If
End Sub
Dim WizardWaitForNoBalls


Sub EndWizard1_Timer
  EndWizard1.Interval = 1000
  If WizardWaitForNoBalls = 0 Then
    EndWizard1.Enabled = False
    ProperEndWiz1
  End If

End Sub

Sub ProperEndWiz1
  L18.blinkinterval = 110 : L20.blinkinterval = 110 : L20.blinkinterval = 110
  L24.blinkinterval = 110 : L25.blinkinterval = 110 : L26.blinkinterval = 110
  L27.blinkinterval = 110 : L28.blinkinterval = 110 : L29.blinkinterval = 110
  L30.blinkinterval = 110 : L31.blinkinterval = 110 : L32.blinkinterval = 110
  L37.blinkinterval = 110 : L38.blinkinterval = 110 : L39.blinkinterval = 110
  L40.blinkinterval = 110 : L41.blinkinterval = 110 : L42.blinkinterval = 110
  L43.blinkinterval = 110 : L44.blinkinterval = 110 : L45.blinkinterval = 110
  L46.blinkinterval = 110 : L47.blinkinterval = 110 : L48.blinkinterval = 110
  L21.blinkinterval = 110 : L22.blinkinterval = 110 : L23.blinkinterval = 110
' debug.print "tilted? " & btilted
' debug.print "Endwizard1:  wizard reset now"

  WizardTimer = 0
  WizardSwitches = 0
  InsertSequence.play SeqBlinking,,3,33
  AreLightsOff = 0
  currentmode = 0
  bTilted = False
  playtime = 0
  UpdateLights
ClearDMD
  startnewplayer

End Sub



Sub DoTheWizIntro
    AddABallReady = 0

    WizardStart = WizardStart + 1

    Select Case WizardStart
      Case 2
        InsertSequence.Play SeqDownOn,70,5
        FlexDMD.Stage.Addactor FlexDMD.NewImage("wizimage",   "SpongeWizard.png")
        WizardPos = 0
      case 30
        FlexDMD.Stage.Addactor FlexDMD.Newlabel("wiz_w",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(11,11,11), RGB(255,255,255),1)  , "W")
        FlexDMD.stage.GetLabel("wiz_w").SetAlignedPosition 33,26, FlexDMD_Align_Center
        WizText(0)= 255
        WizardShake = 5
        PlaySound "wizshort",0,RomSoundVolume


      case 50
        FlexDMD.Stage.Addactor FlexDMD.Newlabel("wiz_i",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(11,11,11), RGB(255,255,255),1)  , "I")
        FlexDMD.stage.GetLabel("wiz_i").SetAlignedPosition 46,26, FlexDMD_Align_Center
        WizText(1)= 255
        WizardShake = 5
        PlaySound "wizshort",0,RomSoundVolume

      case 70
        FlexDMD.Stage.Addactor FlexDMD.Newlabel("wiz_z",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(11,11,11), RGB(255,255,255),1)  , "Z")
        FlexDMD.stage.GetLabel("wiz_z").SetAlignedPosition 59,26, FlexDMD_Align_Center
        WizText(2)= 255
        WizardShake = 5
        PlaySound "wizshort",0,RomSoundVolume

      case 90
        FlexDMD.Stage.Addactor FlexDMD.Newlabel("wiz_a",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(11,11,11), RGB(255,255,255),1)  , "A")
        FlexDMD.stage.GetLabel("wiz_a").SetAlignedPosition 72,26, FlexDMD_Align_Center
        WizText(3)= 255
        WizardShake = 5
        PlaySound "wizshort",0,RomSoundVolume

      case 110
        FlexDMD.Stage.Addactor FlexDMD.Newlabel("wiz_r",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(11,11,11), RGB(255,255,255),1)  , "R")
        FlexDMD.stage.GetLabel("wiz_r").SetAlignedPosition 85,26, FlexDMD_Align_Center
        WizText(4)= 255
        WizardShake = 5
        PlaySound "wizshort",0,RomSoundVolume

      case 130
        FlexDMD.Stage.Addactor FlexDMD.Newlabel("wiz_d",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(11,11,11), RGB(255,255,255),1)  , "D")
        FlexDMD.stage.GetLabel("wiz_d").SetAlignedPosition 98,26, FlexDMD_Align_Center
        WizText(5)= 255
        WizardShake = 5
        PlaySound "wizshort",0,RomSoundVolume
      Case 177
        CurrentMode = 20
        StopTableMusic
        PlaySound "rock", 1, RomSoundVolume  'for intro
        songtimer.interval = 9000 : SongTimer.Enabled = False : SongTimer.Enabled = True
        WizardShake = 5

      case 208
        wizardshake = 10

      case 222
        wizardshake = 5

      case 160,180,200,220,240,260,280,300,320,340
        FlexDMD.stage.GetLabel("wiz_w").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(0, 0, 0), RGB( WizText(0) , WizText(0) , WizText(0)) , 1)
        FlexDMD.stage.GetLabel("wiz_i").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(0, 0, 0), RGB( WizText(1) , WizText(1) , WizText(1)) , 1)
        FlexDMD.stage.GetLabel("wiz_z").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(0, 0, 0), RGB( WizText(2) , WizText(2) , WizText(2)) , 1)
        FlexDMD.stage.GetLabel("wiz_a").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(0, 0, 0), RGB( WizText(3) , WizText(3) , WizText(3)) , 1)
        FlexDMD.stage.GetLabel("wiz_r").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(0, 0, 0), RGB( WizText(4) , WizText(4) , WizText(4)) , 1)
        FlexDMD.stage.GetLabel("wiz_d").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(0, 0, 0), RGB( WizText(5) , WizText(5) , WizText(5)) , 1)
      case 166,186,206,226,246,266,286,306,326,346
        FlexDMD.stage.GetLabel("wiz_w").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 222, 222), RGB( WizText(0) , WizText(0) , WizText(0)) , 1)
        FlexDMD.stage.GetLabel("wiz_i").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 222, 222), RGB( WizText(1) , WizText(1) , WizText(1)) , 1)
        FlexDMD.stage.GetLabel("wiz_z").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 222, 222), RGB( WizText(2) , WizText(2) , WizText(2)) , 1)
        FlexDMD.stage.GetLabel("wiz_a").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 222, 222), RGB( WizText(3) , WizText(3) , WizText(3)) , 1)
        FlexDMD.stage.GetLabel("wiz_r").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 222, 222), RGB( WizText(4) , WizText(4) , WizText(4)) , 1)
        FlexDMD.stage.GetLabel("wiz_d").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 222, 222), RGB( WizText(5) , WizText(5) , WizText(5)) , 1)
      case 355

        If Wizard1Done(CurrentPlayer) = 1 Then
          FlexDMD.Stage.Addactor FlexDMD.Newlabel("wiz_1",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(255,33,0), RGB(11,11,11),0)  , "ALL SWITCHES IS 25.000")
          FlexDMD.stage.GetLabel("wiz_1").SetAlignedPosition 64,4, FlexDMD_Align_Center
          wizardshake = 3
        Else
          FlexDMD.Stage.Addactor FlexDMD.Newlabel("wiz_1",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(255,33,0), RGB(11,11,11),0)  , "ALL SWITCHES IS 100.000")
          FlexDMD.stage.GetLabel("wiz_1").SetAlignedPosition 64,4, FlexDMD_Align_Center
          wizardshake = 3
        End If

      case 385
        If Wizard1Done(CurrentPlayer) = 1 Then
          FlexDMD.Stage.Addactor FlexDMD.Newlabel("wiz_2",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(255,255,0), RGB(11,11,11),0)  , "COLLECT ALL 9 SHOTS TWICE")
          FlexDMD.stage.GetLabel("wiz_2").SetAlignedPosition 64,9, FlexDMD_Align_Center
          wizardshake = 3
        Else
          FlexDMD.Stage.Addactor FlexDMD.Newlabel("wiz_2",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(255,255,0), RGB(11,11,11),0)  , "COL2LECT 240 FOR MAX PAYOUT")
          FlexDMD.stage.GetLabel("wiz_2").SetAlignedPosition 64,9, FlexDMD_Align_Center
          wizardshake = 3
        End If
      case 415
        If Wizard1Done(CurrentPlayer) = 1 Then
          FlexDMD.Stage.Addactor FlexDMD.Newlabel("wiz_3",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(255,33,0), RGB(11,11,11),0)  , "TIMED EVENT 180 SECONDS")
          FlexDMD.stage.GetLabel("wiz_3").SetAlignedPosition 64,14, FlexDMD_Align_Center
          wizardshake = 3
        Else
          FlexDMD.Stage.Addactor FlexDMD.Newlabel("wiz_3",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(255,33,0), RGB(11,11,11),0)  , "TRIBALL MULTIBALL 180 SECONDS")
          FlexDMD.stage.GetLabel("wiz_3").SetAlignedPosition 64,14, FlexDMD_Align_Center
          wizardshake = 3
        End If


    End Select

    If WizText(0) > 15 Then WizText(0) = WizText(0) - 5 : FlexDMD.stage.GetLabel("wiz_w").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 188, 188), RGB( WizText(0) , WizText(0) , WizText(0)) , 1)
    If WizText(1) > 15 Then WizText(1) = WizText(1) - 5 : FlexDMD.stage.GetLabel("wiz_i").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 188, 188), RGB( WizText(1) , WizText(1) , WizText(1)) , 1)
    If WizText(2) > 15 Then WizText(2) = WizText(2) - 5 : FlexDMD.stage.GetLabel("wiz_z").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 188, 188), RGB( WizText(2) , WizText(2) , WizText(2)) , 1)
    If WizText(3) > 15 Then WizText(3) = WizText(3) - 5 : FlexDMD.stage.GetLabel("wiz_a").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 188, 188), RGB( WizText(3) , WizText(3) , WizText(3)) , 1)
    If WizText(4) > 15 Then WizText(4) = WizText(4) - 5 : FlexDMD.stage.GetLabel("wiz_r").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 188, 188), RGB( WizText(4) , WizText(4) , WizText(4)) , 1)
    If WizText(5) > 15 Then WizText(5) = WizText(5) - 5 : FlexDMD.stage.GetLabel("wiz_d").font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 188, 188), RGB( WizText(5) , WizText(5) , WizText(5)) , 1)



    If WizardPos < -100 Then WizardDir = 1
    If WizardPos > -35 Then WizardDir = -1
    WizardPos = WizardPos + WizardDir
    FlexDMD.stage.Getimage("wizimage").SetPosition  -10+WizardShake  , WizardPos
    FlexDMD.stage.Getimage("wizimage").visible = True

    If WizardStart > 700 Then
      InsertSequence.stopPlay
      InsertSequence.Play SeqBlinking,,3,55
      FlexDMD.stage.GetImage("wizimage").remove
      FlexDMD.stage.GetLabel("wiz_w").remove
      FlexDMD.stage.GetLabel("wiz_i").remove
      FlexDMD.stage.GetLabel("wiz_z").remove
      FlexDMD.stage.GetLabel("wiz_a").remove
      FlexDMD.stage.GetLabel("wiz_r").remove
      FlexDMD.stage.GetLabel("wiz_d").remove
      FlexDMD.stage.GetLabel("wiz_1").remove
      FlexDMD.stage.GetLabel("wiz_2").remove
      FlexDMD.stage.GetLabel("wiz_3").remove
      WizardStart = 0
      StopPinappleWiz2 = 0

      MysteryReady(i) = 0
      BankDown(i)=0
      spinnerhurryup=0
      MultiBallReady(i) = 0
      ModesComplete(i)=0
      NextMode(i)=0
      KickerSave(i) = 2
      SJPReady(i) = 0
      ModeText(i) = ""

      dim i
      i=CurrentPlayer
      L01State(i)=0
      L02State(i)=0
      L03State(i)=0
      L04State(i)=0

      L05State(i)=1 ' mode inserts on ?
      L06State(i)=1
      L07State(i)=1
      L08State(i)=1
      L09State(i)=1
      L10State(i)=1
      L11State(i)=1
      L12State(i)=1
      L13State(i)=1

      L14State(i)=0
      L15State(i)=0
      L16State(i)=0
      L17State(i)=0
      L18State(i)=0
      L19State(i)=0
      L20State(i)=0
      L21State(i)=0
      L22State(i)=0
      L23State(i)=0
      L24State(i)=0
      L25State(i)=0
      L26State(i)=0
      L27State(i)=0
      L28State(i)=0
      L29State(i)=0
      L30State(i)=0
      L31State(i)=0
      L32State(i)=0
      L33State(i)=0
      L34State(i)=0
      L35State(i)=0
      L36State(i)=0
      L37State(i)=0
      L38State(i)=0
      L39State(i)=0
      L40State(i)=0
      L41State(i)=0
      L42State(i)=0
      L43State(i)=0
      L44State(i)=0
      L45State(i)=0
      L46State(i)=0
      L47State(i)=0
      L48State(i)=0
      L49State(i)=0
      L50State(i)=0

      If BankDown(CurrentPlayer) = 1 Then currpos = 0 : Call Update3Bank(0,50)      ' close pineapple
      BankMove = 0
      TriggerDown(CurrentPlayer) = False
      TriggerWall.Collidable = True : TriggerWallCollide(CurrentPlayer)=True
      TriggerWall.IsDropped = Not TriggerWall.Collidable
      sw77MovableHelper


      ModeText(currentplayer) = ""
      NextMode(CurrentPlayer) = 0
      ModeReady(CurrentPlayer) = 0
      modetimecounter = 0
      ResetModeLights

      BallSaverTimer.Enabled = True
      bTilted = False

      WizAddaball = 0
      lightCtrl.AddTableLightSeq "GI", lSeqRainbow
      If Wizard1Done(CurrentPlayer) = 1 Then
        playtime = - 14
        Wizard1Done(CurrentPlayer) = 0
        CurrentMode = 21
        wizBiP = 1
        Call RandomKicker(3)
      Else
        playtime = - BallSaveTime(CurrentPlayer) - 14
        Wizard1Done(CurrentPlayer) = 1
        CurrentMode = 20
        WizballsTolaunch = 3
        wizBiP = 3
        Wizmultiballstart.interval = 1000
        Wizmultiballstart.Enabled = True
      End If
      wiz2Addaball = 0
      UpdateLights
      WizardTimer = 180
      WizardSwitches = 0
      collect1 = 2
      collect2 = 2
      collect3 = 2
      collect4 = 2
      collect5 = 2
      collect6 = 2
      collect7 = 2
      collect8 = 2
      collect9 = 2
      rampsforaddaball = 1
    End If
End Sub
Dim wiz2Addaball
Dim wizBiP
Dim wizAddaBall  ' fixing adding this to the ramps 3 = addaball   mby more





Dim WizballsTolaunch
Sub Wizmultiballstart_Timer '1000
  If WizballsTolaunch > 0 Then
    WizballsTolaunch = WizballsTolaunch - 1
    If WizballsTolaunch = 2 Then
      Call RandomKicker(3)'
' debug.print "Wizmultiballstart_Timer"
      Wizmultiballstart.interval = 1700
    Else
      Call RandomKicker(4)
    End If
  Else
    Wizmultiballstart.Enabled = False
  End If
End Sub




Sub DMD_ResetAll
    FlexMsg = 0
    FlexDMD.Stage.GetImage("Title").Visible = False
    FlexDMD.Stage.GetImage("dodss").Visible = False
    FlexDMD.Stage.GetImage("BG001").Visible=False   ' need all of them
  ' FlexDMD.Stage.GetImage("BG002").Visible=False ' skip this one
  ' FlexDMD.Stage.GetImage("BG003").Visible=False  ' highscore dont hide
    FlexDMD.Stage.GetImage("BG004").Visible=False
    FlexDMD.Stage.GetImage("BG005").Visible=False
    FlexDMD.Stage.GetImage("party").Visible=False
    FlexDMD.Stage.GetImage("BG006").Visible=False
    FlexDMD.Stage.GetImage("BG007").Visible=False
    FlexDMD.Stage.GetImage("BG008").Visible=False
    FlexDMD.Stage.GetImage("BG009").Visible=False
    FlexDMD.Stage.GetImage("BG010").Visible=False
    FlexDMD.Stage.GetImage("BG011").Visible=False
'   FlexDMD.Stage.GetImage("BG012").Visible=False
    FlexDMD.Stage.GetImage("BG013").Visible=False
'   FlexDMD.Stage.GetImage("BG014").Visible=False
    FlexDMD.Stage.GetImage("BGwarn").Visible=False
    FlexDMD.Stage.GetImage("BGtilt").Visible=False

    FlexDMD.Stage.GetVideo("VIDchaos").Visible=False
    FlexDMD.Stage.GetVideo("VIDsponge").Visible=False
    FlexDMD.Stage.GetVideo("VIDplankt").Visible=False
    FlexDMD.Stage.GetVideo("VIDgoofy").Visible=False
    FlexDMD.Stage.GetVideo("VIDcrabs").Visible=False
    FlexDMD.Stage.GetVideo("VIDjelly").Visible=False
    FlexDMD.Stage.GetVideo("VIDmoney").Visible=False
    FlexDMD.Stage.GetVideo("VIDtikiland").Visible=False
    FlexDMD.Stage.GetVideo("VIDboating").Visible=False
    FlexDMD.Stage.GetVideo("VIDsandys").Visible=False
    FlexDMD.Stage.GetVideo("VIDmememe").Visible=False
    FlexDMD.Stage.GetVideo("VIDhouse").Visible=False
    FlexDMD.Stage.GetVideo("VIDdance").Visible=False


    FlexDMD.Stage.GetLabel("TextSmalLine1").Visible=False
    FlexDMD.Stage.GetLabel("TextSmalLine2").Visible=False
    FlexDMD.Stage.GetLabel("TextSmalLine3").Visible=False
    FlexDMD.Stage.GetLabel("TextSmalLine4").Visible=False
End Sub

'DMDMessage1 "WARNING"      ,""       ,FontSponge16 ,64,16  ,64,16  ,"shake"  ,5  ,"BGwarn" ,"novideo"
'DMDMessage1 "TILT"       ,""       ,FontSponge32 ,64,16  ,64,16  ,"blink"  ,11 ,"noimage"  ,"VIDchaos"
'FlexMsg = 1
'FlexMsgTime = 1 ' turnoff tilt ( or stop midanim )


Dim CurrentMSGdone
CurrentMSGdone = 0
Sub DMD_SlideOff
  if DMDDisplay(20,7) = "noslide" Or DMDDisplay(20,7) = "noslide2" Then CurrentMSGdone = 0
  If CurrentMSGdone > 0 Then
    DMD_Slide = DMD_Slide - 3
    If DMD_Slide < -32 Then CurrentMSGdone = 0
    DMD_Update
  Else
    DMD_Display_Timer
  End If


End Sub

' : If DoOrDieMode < 3 Then FlexDMD.Stage.GetImage("dodss").visible = False

Sub DMD_NewOverlay
  CurrentMSGdone = 1 ' to get it off screen
  If DMDDisplay(20,0) = "" then FlexMsg = 0 :Exit Sub
  Dim label
  FlexMsg = 2
  DMD_ShakePos = 12
  DMD_slide = 32

  if DMDDisplay(20,7) = "noslide2" Then DMD_slide = 0
  if DMDDisplay(20,7) = "noslide2blink" Then DMD_slide = 0

  if DMDDisplay(20,7) = "noslide3" Then DMD_Slide = 0

  If DMDDisplay(20,9)  <> "noimage" Then FlexDMD.Stage.GetImage(DMDDisplay(20,9)).Visible=True : FlexDMD.Stage.GetImage(DMDDisplay(20,9)).SetPosition 0, - DMD_slide
  If DMDDisplay(20,10) <> "novideo" Then FlexDMD.Stage.GetVideo(DMDDisplay(20,10)).Visible=True : FlexDMD.Stage.GetVideo(DMDDisplay(20,10)).SetPosition 0, - DMD_slide * 1.5
  If DMDDisplay(20,11) = 1 Then
    Set label = FlexDMD.Stage.GetLabel("TextSmalLine1")
    label.Font = DMDDisplay(20,2)
    label.Text = DMDDisplay(20,0)
    label.SetAlignedPosition DMDDisplay(20,3),DMDDisplay(20,4) - DMD_slide ,FlexDMD_Align_Center
    label.visible = True
  End If
  If DMDDisplay(20,11) = 2 Then

    Set label = FlexDMD.Stage.GetLabel("TextSmalLine1")
    label.Font = DMDDisplay(20,2)
    label.Text = DMDDisplay(20,0)
    label.SetAlignedPosition DMDDisplay(20,3),DMDDisplay(20,4)-8 - DMD_slide ,FlexDMD_Align_Center
    label.visible = True
'   If DMDDisplay(20,7) = "hurry" Then
'     Set label = FlexDMD.Stage.GetLabel("TextSmalLine2")
'     label.Font = DMDDisplay(20,2)
'     label.Text = DMDDisplay(20,1)
'     label.SetAlignedPosition DMDDisplay(20,3),DMDDisplay(20,4) - DMD_slide ,FlexDMD_Align_Center
'     label.visible = True
'     Set label = FlexDMD.Stage.GetLabel("TextSmalLine3")
'     label.Font = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, RGB(0, 0, 0), 1)
'     label.Text =  formatscore(totalscore(currentplayer))
'     label.SetAlignedPosition 60,29,FlexDMD_Align_Center
'     label.visible = True
'   Else
      Set label = FlexDMD.Stage.GetLabel("TextSmalLine2")
      label.Font = DMDDisplay(20,2)
      label.Text = DMDDisplay(20,1)
      label.SetAlignedPosition DMDDisplay(20,3),DMDDisplay(20,4)+8 - DMD_slide ,FlexDMD_Align_Center
      label.visible = True
'   End If

  End If

'     DMDDisplay(h,0) = Message
'     DMDDisplay(h,1) = Message2
'     DMDDisplay(h,2) = Font
'     DMDDisplay(h,3) = posX
'     DMDDisplay(h,4) = posY
'     DMDDisplay(h,5) = endX
'     DMDDisplay(h,6) = endY
'     DMDDisplay(h,7) = state
'     DMDDisplay(h,8) = time
'     DMDDisplay(h,9) = image
'     DMDDisplay(h,10) = video
'     if message2 ="" then DMDDisplay(h,11)= 1 Else DMDDisplay(h,11) = 2

'   FlexDMD.Stage.GetImage("BG001").Visible=True
'   Set label = FlexDMD.Stage.GetLabel("TextSmalLine1")
'   label.Font = FontSponge12
'   label.Text = FlexText1
'   label.SetAlignedPosition 64,16,FlexDMD_Align_Center
'   label.visible = True
End Sub

Dim DMD_Slide
Dim DMD_ShakePos
Dim inlaneEBcount
Dim inlane1
Dim inlane2
Dim inlane3
Dim inlane4
Dim inlane5
Dim inlane6


'     DMDMessage1 "testing testing again" ,"" ,FontScoreActive  ,54,28  ,54,28  ,"line4only"  ,2  ,"noimage"  ,"novideo"
' TextSmalLine4
Dim line4
Dim msg4
' line4 = 1 : msg4 = "TEST 1 TEST 2 TEST 3 TEST 4"

Sub DMD_Update
  dim tmp
  If DMD_Slide > 0 Then DMD_slide = DMD_slide - 2
  If frame mod 3 = 1 Then
    If DMDDisplay(20,7) = "shake" And Not DMD_ShakePos = 0 Then ' shaking
      If DMD_ShakePos < 0 Then DMD_ShakePos = ABS(DMD_ShakePos) - 1 Else DMD_ShakePos = - DMD_ShakePos + 1

      Playsound "AlienShake4",0,RomSoundVolume

    End If
  End If
  If DMDDisplay(20,7) = "blink" or DMDDisplay(20,7) = "noslide2blink" Then ' blinking
    If frame mod 30 > 15 Then
      FlexDMD.Stage.GetLabel("TextSmalLine1").visible = False
      If DMDDisplay(20,11) > 1 Then  FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False
    Else
      FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
      If DMDDisplay(20,11) > 1 Then  FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
    End If
  End If
  If DMDDisplay(20,7) = "fastblink" Then ' fblinking
    If frame mod 18 > 8 Then
      FlexDMD.Stage.GetLabel("TextSmalLine1").visible = False
      If DMDDisplay(20,11) > 1 Then  FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False
    Else
      FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
      If DMDDisplay(20,11) > 1 Then  FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
    End If
  End If


'     DMDMessage1 "SPINLEVEL " & SpinnerLevel2(CurrentPlayer)   ,"" & formatscore( 1000000+SpinnerLevel2(CurrentPlayer)*250000)     ,FontSponge12BB ,54,20  ,54,20  ,"blink"  ,2  ,"BG006"  ,"novideo"
'   Else
'     If CurrentMSGdone = 0 Then
'       dim tmp
'       tmp = 20 + (SpinnerLevel2(CurrentPlayer)*5) - SpinnerCount2(CurrentPlayer)
'       if tmp < 0 Then tmp = 0
'       DMDMessage1 "SPINLEVELUP" ,"NEED  " & tmp & "  SPINS"   ,FontSponge12BB ,54,20  ,54,20  ,"spinner"  ,2  ,"BG006"  ,"novideo"
'
  If DMDDisplay(20,7) = "spinner" Then
    tmp = sc_SpinnsNeededFirstLvl + (SpinnerLevel2(CurrentPlayer) * sc_SpinnsAdded ) - SpinnerCount2(CurrentPlayer)
    If tmp < 1 Then CurrentMSGdone = 0 : DMD_Display_Timer
    FlexDMD.Stage.GetLabel("TextSmalLine2").Text = "NEED  " & tmp & "  SPINS"
  End If
  If DMDDisplay(20,7) = "spinner2" Then
    tmp = sc_SpinnsNeededFirstLvl + (SpinnerLevel1(CurrentPlayer) * sc_SpinnsAdded ) - SpinnerCount1(CurrentPlayer)
    If tmp < 1 Then CurrentMSGdone = 0 : DMD_Display_Timer
    If tmp > 0 Then FlexDMD.Stage.GetLabel("TextSmalLine2").Text = "NEED  " & tmp & "  SPINS"
  End If





  If inlaneEBcount > 0 Then
    inlaneEBcount = inlaneEBcount + 1
    Select Case inlaneEBcount
      case 2 :
        FlexDMD.Stage.GetLabel("eb1").font = FontSponge32BB
        FlexDMD.Stage.GetLabel("eb2").font = FontSponge32Black
        FlexDMD.Stage.GetLabel("eb3").font = FontSponge32Black
        FlexDMD.Stage.GetLabel("eb4").font = FontSponge32Black
        FlexDMD.Stage.GetLabel("eb5").font = FontSponge32Black
        FlexDMD.Stage.GetLabel("eb6").font = FontSponge32Black
        Select Case InlaneEB(CurrentPlayer)
          case 3 : FlexDMD.Stage.GetLabel("eb2").font = FontSponge32BB
          case 4 : FlexDMD.Stage.GetLabel("eb2").font = FontSponge32BB
               FlexDMD.Stage.GetLabel("eb3").font = FontSponge32BB
          case 5 : FlexDMD.Stage.GetLabel("eb2").font = FontSponge32BB
               FlexDMD.Stage.GetLabel("eb3").font = FontSponge32BB
               FlexDMD.Stage.GetLabel("eb4").font = FontSponge32BB
        End Select

        FlexDMD.Stage.GetLabel("eb1").visible = True : inlane1 = -24 : FlexDMD.Stage.GetLabel("eb1").SetAlignedPosition  20,inlane1,FlexDMD_Align_Center
      case 10
        FlexDMD.Stage.GetLabel("eb2").visible = True : inlane2 = -24 : FlexDMD.Stage.GetLabel("eb2").SetAlignedPosition  42,inlane2,FlexDMD_Align_Center
      case 18
        FlexDMD.Stage.GetLabel("eb3").visible = True : inlane3 = -24 : FlexDMD.Stage.GetLabel("eb3").SetAlignedPosition  65,inlane3,FlexDMD_Align_Center
      case 26
        FlexDMD.Stage.GetLabel("eb4").visible = True : inlane4 = -24 : FlexDMD.Stage.GetLabel("eb4").SetAlignedPosition  86,inlane4,FlexDMD_Align_Center
      case 34
        FlexDMD.Stage.GetLabel("eb5").visible = True :  inlane5 = -24 : FlexDMD.Stage.GetLabel("eb5").SetAlignedPosition  108,inlane5,FlexDMD_Align_Center
        If InlaneEB(CurrentPlayer) > 4 Then
          FlexDMD.Stage.GetLabel("eb6").visible = True : inlane6 = 50  : FlexDMD.Stage.GetLabel("eb6").SetAlignedPosition  68,inlane6,FlexDMD_Align_Center
        End If
      case 60

    End Select
    If inlaneEBcount < 80 Then
      If inlane1 < 14 Then inlane1 = inlane1 + 2
      If inlane2 < 14 Then inlane2 = inlane2 + 2
      If inlane3 < 14 Then inlane3 = inlane3 + 2
      If inlane4 < 14 Then inlane4 = inlane4 + 2
      If inlane5 < 14 Then inlane5 = inlane5 + 2
      FlexDMD.Stage.GetLabel("eb1").SetAlignedPosition  20,inlane1,FlexDMD_Align_Center
      FlexDMD.Stage.GetLabel("eb2").SetAlignedPosition  42,inlane2,FlexDMD_Align_Center
      FlexDMD.Stage.GetLabel("eb3").SetAlignedPosition  64,inlane3,FlexDMD_Align_Center
      FlexDMD.Stage.GetLabel("eb4").SetAlignedPosition  86,inlane4,FlexDMD_Align_Center
      FlexDMD.Stage.GetLabel("eb5").SetAlignedPosition 108,inlane5,FlexDMD_Align_Center
    Elseif inlaneEBcount < 147 Then
      If inlaneEBcount = 115 And InlaneEB(CurrentPlayer) < 5 Then inlaneEBcount = 151
      inlane1 = inlane1 - 1
      inlane2 = inlane2 - 1
      inlane3 = inlane3 - 1
      inlane4 = inlane4 - 1
      inlane5 = inlane5 - 1
      FlexDMD.Stage.GetLabel("eb1").SetAlignedPosition  20,inlane1,FlexDMD_Align_Center
      FlexDMD.Stage.GetLabel("eb2").SetAlignedPosition  42,inlane2,FlexDMD_Align_Center
      FlexDMD.Stage.GetLabel("eb3").SetAlignedPosition  64,inlane3,FlexDMD_Align_Center
      FlexDMD.Stage.GetLabel("eb4").SetAlignedPosition  86,inlane4,FlexDMD_Align_Center
      FlexDMD.Stage.GetLabel("eb5").SetAlignedPosition 108,inlane5,FlexDMD_Align_Center
      inlane6 = inlane6 - 1
      FlexDMD.Stage.GetLabel("eb6").SetAlignedPosition  65,inlane6,FlexDMD_Align_Center
    Elseif inlaneEBcount > 150 then
      FlexDMD.Stage.GetLabel("eb1").visible = False
      FlexDMD.Stage.GetLabel("eb2").visible = False
      FlexDMD.Stage.GetLabel("eb3").visible = False
      FlexDMD.Stage.GetLabel("eb4").visible = False
      FlexDMD.Stage.GetLabel("eb5").visible = False
      FlexDMD.Stage.GetLabel("eb6").visible = False
      FlexDMD.Stage.GetImage("BG014").Visible=False
      inlaneEBcount = 0
      DMDDisplay(20,7) = ""
      FlexMsgTime = 1
      CurrentMSGdone = 0 : DMD_Display_Timer
    End If
    If inlaneEBcount = 140 Then Underfire = 1 : firetime = 70 : ShakeSponge = 5 : SpongeTimer.Interval = 250 : SpongeTimer.Enabled = True



    If frame mod 16 > 7 Then
      Select Case InlaneEB(CurrentPlayer)
        case 1 : FlexDMD.Stage.GetLabel("eb1").font = FontSponge32BB
        case 2 : FlexDMD.Stage.GetLabel("eb2").font = FontSponge32BB
        case 3 : FlexDMD.Stage.GetLabel("eb3").font = FontSponge32BB
        case 4 : FlexDMD.Stage.GetLabel("eb4").font = FontSponge32BB
        case 5 : FlexDMD.Stage.GetLabel("eb5").font = FontSponge32BB : FlexDMD.Stage.GetLabel("eb6").font = FontSponge32BB
      End Select
    Else
      Select Case InlaneEB(CurrentPlayer)
        case 1 : FlexDMD.Stage.GetLabel("eb1").font = FontSponge32black
        case 2 : FlexDMD.Stage.GetLabel("eb2").font = FontSponge32black
        case 3 : FlexDMD.Stage.GetLabel("eb3").font = FontSponge32black
        case 4 : FlexDMD.Stage.GetLabel("eb4").font = FontSponge32black
        case 5 : FlexDMD.Stage.GetLabel("eb5").font = FontSponge32black : FlexDMD.Stage.GetLabel("eb6").font = FontSponge32black
      End Select
    End If
  End If

  If DMDDisplay(20,7) = "tilt" Then ' borderblink
    If frame mod 20 > 9 Then
      FlexDMD.Stage.GetLabel("TextSmalLine1").font = FontSponge32BB
    Else
      FlexDMD.Stage.GetLabel("TextSmalLine1").font = FontSponge32RB
    End If
  End If

  If DMDDisplay(20,3) < DMDDisplay(20,5) Then DMDDisplay(20,3) = DMDDisplay(20,3) + 2
  If DMDDisplay(20,3) > DMDDisplay(20,5) Then DMDDisplay(20,3) = DMDDisplay(20,3) - 2
  If DMDDisplay(20,4) < DMDDisplay(20,6) Then DMDDisplay(20,4) = DMDDisplay(20,4) + 2
  If DMDDisplay(20,4) > DMDDisplay(20,6) Then DMDDisplay(20,4) = DMDDisplay(20,4) - 2

  If DMDDisplay(20,11) = 1 Then
    FlexDMD.Stage.GetLabel("TextSmalLine1").SetAlignedPosition DMDDisplay(20,3) + DMD_ShakePos ,DMDDisplay(20,4) - DMD_Slide , FlexDMD_Align_Center
  Elseif DMDDisplay(20,11) = 2 Then
    FlexDMD.Stage.GetLabel("TextSmalLine1").SetAlignedPosition DMDDisplay(20,3) + DMD_ShakePos ,DMDDisplay(20,4)-8 - DMD_Slide , FlexDMD_Align_Center
    FlexDMD.Stage.GetLabel("TextSmalLine2").SetAlignedPosition DMDDisplay(20,3) + DMD_ShakePos ,DMDDisplay(20,4)+8 - DMD_Slide , FlexDMD_Align_Center
  End If


  If DMDDisplay(20, 9) <> "noimage"  Then FlexDMD.Stage.GetImage(DMDDisplay(20, 9)).SetPosition 0, - DMD_slide
  If DMDDisplay(20,10) <> "novideo" Then FlexDMD.Stage.GetVideo(DMDDisplay(20,10)).SetPosition 0, - DMD_slide * 1.5

  If CurrentMSGdone = 1 or currentball = 0 Or bTilted Or line4=0 or ScoopRestart < frame Then
    FlexDMD.Stage.GetImage("BG002").visible = False
    FlexDMD.Stage.GetLabel("TextSmalLine4").visible = False
  Else
    If Displayballsave <> "" And line4 = 0 Then line4 = 1
    If line4 > 0 Then
      line4 = line4 + 1
      If line4 = 2 And MysteryTimer.enabled = True Then Line4 = 1
      If line4 > 1 Then
        tmp = 9
        If line4 < 10 then tmp = line4
        If line4 > 115 Then tmp = 125 - line4
        If line4 = 15 And line4wait = 1 Then line4 = 40
        If line4 = 20 And Displayballsave <> "" Then line4 = 40
        if line4 = 2 Then
          FlexDMD.Stage.GetLabel("TextSmalLine4").font = FontScoreBlack

          If Displayballsave <>"" Then
            FlexDMD.Stage.GetLabel("TextSmalLine4").text = Displayballsave
            Displayballsave = ""
            line4wait = 1
          Else
            FlexDMD.Stage.GetLabel("TextSmalLine4").text = msg4
          End if

        End If
        FlexDMD.Stage.GetImage("BG002").SetAlignedPosition 64,50 - tmp, FlexDMD_Align_Center
        FlexDMD.Stage.GetLabel("TextSmalLine4").SetAlignedPosition 64,39 - tmp, FlexDMD_Align_Center
        FlexDMD.Stage.GetImage("BG002").visible = True
        FlexDMD.Stage.GetLabel("TextSmalLine4").visible = true

        Select Case line4
          Case 65,75,85,95,105,115  : FlexDMD.Stage.GetLabel("TextSmalLine4").visible = False
          Case 68,78,88,98,108,118 : FlexDMD.Stage.GetLabel("TextSmalLine4").visible = True
        End Select

        If line4 > 125 Then
          If ModeReady(CurrentPlayer) = 1 And CurrentMode <> 11 And BallInLane=0 And CurrentMode > 0 And Msg4 <> "NEXT MODE IS READY" And ScoopRestart < frame Then
            line4 = 45
            msg4 = "NEXT MODE IS READY"
            FlexDMD.Stage.GetLabel("TextSmalLine4").text = msg4
            line4wait = 1

          Else
            line4wait = 0
            FlexDMD.Stage.GetImage("BG002").visible = False
            FlexDMD.Stage.GetLabel("TextSmalLine4").visible = False
            line4 = 0
          End If
        End If
      End If
    End If
  End If

End Sub
Dim line4wait : line4wait = 0
Dim attractspeed : attractspeed = 1
Sub attractmodeseq
    If StartGame <> 0 Then
      Attractlightsequence = 0
      InsertSequence.UpdateInterval = 10
    Else

      Attractlightsequence = Attractlightsequence + attractspeed

      Select Case Attractlightsequence
        case  10 : l05.state = 2 : l18.state = 2 : l19.state = 2 : l20.state = 2 : l01.state = 0
        case  60 : l05.state = 0 : l18.state = 0 : l19.state = 0 : l20.state = 0

        case  80 : l10.state = 2 : l24.state = 2 : l25.state = 2 : l26.state = 2
        case 130 : l10.state = 0 : l24.state = 0 : l25.state = 0 : l26.state = 0

        case 150 : l06.state = 2 : l27.state = 2 : l28.state = 2 : l29.state = 2
        case 200 : l06.state = 0 : l27.state = 0 : l28.state = 0 : l29.state = 0

        case 220 : l11.state = 2 : l30.state = 2 : l31.state = 2 : l32.state = 2
        case 270 : l11.state = 0 : l30.state = 0 : l31.state = 0 : l32.state = 0

        case 290 : l07.state = 2 : l37.state = 2 : l38.state = 2 : l39.state = 2
        case 340 : l07.state = 0 : l37.state = 0 : l38.state = 0 : l39.state = 0

        case 360 : l12.state = 2 : l40.state = 2 : l41.state = 2 : l42.state = 2
        case 410 : l12.state = 0 : l40.state = 0 : l41.state = 0 : l42.state = 0

        case 430 : l08.state = 2 : l43.state = 2 : l44.state = 2 : l45.state = 2
        case 480 : l08.state = 0 : l43.state = 0 : l44.state = 0 : l45.state = 0

        case 500 : l13.state = 2 : l46.state = 2 : l47.state = 2 : l48.state = 2
        case 550 : l13.state = 0 : l46.state = 0 : l47.state = 0 : l48.state = 0

        case 570 : l09.state = 2 : l21.state = 2 : l22.state = 2 : l23.state = 2
        case 620 : l09.state = 0 : l21.state = 0 : l22.state = 0 : l23.state = 0
        case 630 :
              Attractlightsequence = 0
              InsertSequence.UpdateInterval = 5
              InsertSequence.Play SeqCircleOutOn,50,2
              InsertSequence.Play SeqUpOn,50,1
              l01.state = 2
              Select Case int(rnd(1)*13)
                case 1 : PlaySound "IntroMode0",0,RomSoundVolume
                case 2 : PlaySound "IntroMode0a",0,RomSoundVolume
                case 3 : PlaySound "IntroMode0b",0,RomSoundVolume
                case 4 : PlaySound "IntroMode0c",0,RomSoundVolume
                case 5 : PlaySound "IntroMode0d",0,RomSoundVolume
                case 6 : PlaySound "IntroMode0e",0,RomSoundVolume
                case 8 : InsertSequence.Play SeqRandom,10,,1000
                case 9 : InsertSequence.Play SeqRandom,10,,1500
                case 10 : InsertSequence.Play SeqBlinking,,6,35
                case 11 : InsertSequence.Play SeqBlinking,,3,77

              End Select

              InsertSequence.UpdateInterval = 10
              InsertSequence.Play SeqCircleOutOn,50,int(rnd(1)*5)+2

      End Select
    End If
End Sub

Dim LinePos1
Dim LinePos2
Dim LinePos3
Dim LinePos4
Dim DoOrDieScore

Sub AttractModeUpdate
  Dim label,x
  x = 52

  If Not bInOptions Then AttractCount = AttractCount + 1
  Select Case AttractCount
    case   2 :
      Set label = FlexDMD.Stage.GetLabel("TextSmalLine1") : label.Font = FontSponge16A : label.Text = " " : label.SetAlignedPosition 64,40,FlexDMD_Align_Center : label.visible = True
      Set label = FlexDMD.Stage.GetLabel("TextSmalLine2") : label.Font = FontSponge16A : label.Text = " " : label.SetAlignedPosition 64,40,FlexDMD_Align_Center : label.visible = True
      Set label = FlexDMD.Stage.GetLabel("TextSmalLine3") : label.Font = FontSponge16A : label.Text = " " : label.SetAlignedPosition 64,40,FlexDMD_Align_Center : label.visible = True
      Set label = FlexDMD.Stage.GetLabel("TextSmalLine4") : label.Font = FontSponge16A : label.Text = " " : label.SetAlignedPosition 64,40,FlexDMD_Align_Center : label.visible = True
      stoponlast = 0
      ChangeGILights(1)
    case 10 + x   : LinePos1 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine1").Text = "SCORES"                       : FlexDMD.Stage.GetLabel("TextSmalLine1").font = FontSponge16A
    case 10 + x *2  :
  FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine3").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine4").visible = True
            LinePos2 = 60
            If DoOrDieScore > 0 then
              FlexDMD.Stage.GetLabel("TextSmalLine2").Text = "DOD " & formatscore2(DoOrDieScore)                      : FlexDMD.Stage.GetLabel("TextSmalLine2").font = FontHighScore

            Else
              If AttractScore(1)=0 then
                FlexDMD.Stage.GetLabel("TextSmalLine2").Text = "PLAY A GAME"                            : FlexDMD.Stage.GetLabel("TextSmalLine2").font = FontHighScore
                AttractCount = 10 + x *5
              Else
                FlexDMD.Stage.GetLabel("TextSmalLine2").Text = formatscore2(AttractScore(1))                      : FlexDMD.Stage.GetLabel("TextSmalLine2").font = FontHighScore
              End if
            End If
    case 10 + x *3  :
            If AttractScore(2)=0 then
              AttractCount = 10 + x *5
            Else
              LinePos3 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine3").Text = formatscore2(AttractScore(2))              : FlexDMD.Stage.GetLabel("TextSmalLine3").font = FontHighScore
            End If
    case 10 + x *4  :
            If AttractScore(3)=0 then
              AttractCount = 10 + x *5
            Else
              LinePos4 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine4").Text = formatscore2(AttractScore(3))              : FlexDMD.Stage.GetLabel("TextSmalLine4").font = FontHighScore
            End If
    case 10 + x *5  :
            If AttractScore(4) > 0 then LinePos1 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine1").Text = formatscore2(AttractScore(4))  : FlexDMD.Stage.GetLabel("TextSmalLine1").font = FontHighScore

    case 10 + x *8  :
            FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
            FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
            FlexDMD.Stage.GetLabel("TextSmalLine3").visible = True
            FlexDMD.Stage.GetLabel("TextSmalLine4").visible = True
            LinePos2 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine2").Text = "HIGHSCORES"                     : FlexDMD.Stage.GetLabel("TextSmalLine2").font = FontSponge16A
case 10 + x *8+5  : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False
case 10 + x *8+10 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
case 10 + x *8+15 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False
case 10 + x *8+20 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
case 10 + x *8+25 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False
case 10 + x *8+30 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
case 10 + x *8+35 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False
case 10 + x *8+40 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
case 10 + x *8+45 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False
case 10 + x *8+50 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True

    case 10 + x *9  : LinePos3 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine3").Text = HighScores(0,1) & " " & formatscore2(HighScores(0,0))  : FlexDMD.Stage.GetLabel("TextSmalLine3").font = FontHighScore
case 10 + x *9+5  : FlexDMD.Stage.GetLabel("TextSmalLine3").visible = False
case 10 + x *9+10 : FlexDMD.Stage.GetLabel("TextSmalLine3").visible = True
case 10 + x *9+15 : FlexDMD.Stage.GetLabel("TextSmalLine3").visible = False
case 10 + x *9+20 : FlexDMD.Stage.GetLabel("TextSmalLine3").visible = True
case 10 + x *9+25 : FlexDMD.Stage.GetLabel("TextSmalLine3").visible = False
case 10 + x *9+30 : FlexDMD.Stage.GetLabel("TextSmalLine3").visible = True
case 10 + x *9+35 : FlexDMD.Stage.GetLabel("TextSmalLine3").visible = False
case 10 + x *9+40 : FlexDMD.Stage.GetLabel("TextSmalLine3").visible = True
case 10 + x *9+45 : FlexDMD.Stage.GetLabel("TextSmalLine3").visible = False
case 10 + x *9+50 : FlexDMD.Stage.GetLabel("TextSmalLine3").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine3").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine4").visible = True

    case 10 + x *10 : LinePos4 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine4").Text = HighScores(1,1) & " " & formatscore2(HighScores(1,0))  : FlexDMD.Stage.GetLabel("TextSmalLine4").font = FontHighScore
    case 10 + x *11 : LinePos1 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine1").Text = HighScores(2,1) & " " & formatscore2(HighScores(2,0))  : FlexDMD.Stage.GetLabel("TextSmalLine1").font = FontHighScore
    case 10 + x *12 : LinePos2 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine2").Text = HighScores(3,1) & " " & formatscore2(HighScores(3,0))  : FlexDMD.Stage.GetLabel("TextSmalLine2").font = FontHighScore
    case 10 + x *13 : LinePos3 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine3").Text = HighScores(4,1) & " " & formatscore2(HighScores(4,0))  : FlexDMD.Stage.GetLabel("TextSmalLine3").font = FontHighScore


    Case 10 + x *15 : LinePos4 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine4").Text = "DO OR DIE"                      : FlexDMD.Stage.GetLabel("TextSmalLine4").font = FontSponge16A
    Case 10 + x *16 : LinePos1 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine1").Text = HighScores(5,1) & " " & formatscore2(HighScores(5,0))  : FlexDMD.Stage.GetLabel("TextSmalLine1").font = FontHighScore
case 10 + x *16+5 : FlexDMD.Stage.GetLabel("TextSmalLine1").visible = False
case 10 + x *16+10  : FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
case 10 + x *16+15  : FlexDMD.Stage.GetLabel("TextSmalLine1").visible = False
case 10 + x *16+20  : FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
case 10 + x *16+25  : FlexDMD.Stage.GetLabel("TextSmalLine1").visible = False
case 10 + x *16+30  : FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
case 10 + x *16+35  : FlexDMD.Stage.GetLabel("TextSmalLine1").visible = False
case 10 + x *16+40  : FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
case 10 + x *16+45  : FlexDMD.Stage.GetLabel("TextSmalLine1").visible = False
case 10 + x *16+50  : FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine3").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine4").visible = True

    Case 10 + x *18 : LinePos2 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine2").Text = "CREDITS"                        : FlexDMD.Stage.GetLabel("TextSmalLine2").font = FontSponge16A
    Case 10 + x *19 : LinePos3 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine3").Text = "SKILLSHOT"                      : FlexDMD.Stage.GetLabel("TextSmalLine3").font = FontHighScore
    Case 10 + x *20 : LinePos4 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine4").Text = "GEDANKEKOJOTE"                    : FlexDMD.Stage.GetLabel("TextSmalLine4").font = FontHighScore
    Case 10 + x *21 : LinePos1 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine1").Text = "APOPHIS"                        : FlexDMD.Stage.GetLabel("TextSmalLine1").font = FontHighScore
    Case 10 + x *22 : LinePos2 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine2").Text = "FLUX"                         : FlexDMD.Stage.GetLabel("TextSmalLine2").font = FontHighScore
    Case 10 + x *23 : LinePos3 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine3").Text = "RAWD"                         : FlexDMD.Stage.GetLabel("TextSmalLine3").font = FontHighScore
    Case 10 + x *24 : LinePos4 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine4").Text = "OQQSAN"                       : FlexDMD.Stage.GetLabel("TextSmalLine4").font = FontHighScore
    Case 10 + x *25 : LinePos1 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine1").Text = "FRISCOPINBALL"                    : FlexDMD.Stage.GetLabel("TextSmalLine1").font = FontHighScore
    Case 10 + x *26 : LinePos2 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine2").Text = "ASTRONASTY"                     : FlexDMD.Stage.GetLabel("TextSmalLine2").font = FontHighScore
    Case 10 + x *27 : LinePos3 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine3").Text = "REMDWAAS"                       : FlexDMD.Stage.GetLabel("TextSmalLine3").font = FontHighScore
    Case 10 + x *28 : LinePos4 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine4").Text = "NIWAK"                        : FlexDMD.Stage.GetLabel("TextSmalLine4").font = FontHighScore
    Case 10 + x *29 : LinePos1 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine1").Text = "LEOJREIMROC"                      : FlexDMD.Stage.GetLabel("TextSmalLine1").font = FontHighScore
    Case 10 + x *30 : LinePos2 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine2").Text = "MRGRYNCH"                       : FlexDMD.Stage.GetLabel("TextSmalLine2").font = FontHighScore

  FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine3").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine4").visible = True

    Case 10 + x *32 : LinePos4 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine4").Text = "PLAY PINBALL"                     : FlexDMD.Stage.GetLabel("TextSmalLine4").font = FontHighScore
    Case 10 + x *33 : LinePos1 = 60 : FlexDMD.Stage.GetLabel("TextSmalLine1").Text = "LIVE LONGER"                      : FlexDMD.Stage.GetLabel("TextSmalLine1").font = FontHighScore

    Case 10 + x *35 : LinePos2 = 60 :
          If UseCredits And Credits < 1 Then
            FlexDMD.Stage.GetLabel("TextSmalLine2").Text = "INSERT COIN"                      : FlexDMD.Stage.GetLabel("TextSmalLine2").font = FontSponge16B
          Else
            FlexDMD.Stage.GetLabel("TextSmalLine2").Text = "PRESS START"                      : FlexDMD.Stage.GetLabel("TextSmalLine2").font = FontSponge16B
          End If
  FlexDMD.Stage.GetLabel("TextSmalLine1").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine3").visible = True
  FlexDMD.Stage.GetLabel("TextSmalLine4").visible = True

    case 10 + x *35 + 40  : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False
    case 10 + x *35 + 50  : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
    case 10 + x *35 + 60  : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False
    case 10 + x *35 + 70  : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True  : stoponlast = 1
    case 10 + x *35 + 80  : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False
    case 10 + x *35 + 90  : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
    case 10 + x *35 + 100 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False
    case 10 + x *35 + 110 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
    case 10 + x *35 + 120 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False
    case 10 + x *35 + 130 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = True
    case 10 + x *35 + 140 : FlexDMD.Stage.GetLabel("TextSmalLine2").visible = False

    case 10 + x *35 + 150 :
      ChangeGILights(0)


      AttractvideoON = true
      FlexDMD.Stage.Addactor FlexDMD.Newvideo ("VIDattr","Attract2.gif")
      FlexDMD.Stage.GetVideo ("VIDattr").SetBounds 0, 0, 128, 64 :
      FlexDMD.Stage.GetVideo ("VIDattr").SetAlignedPosition 64, 12 , FlexDMD_Align_Center


    case 650 + x *35: AttractCount = 1 : FlexDMD.stage.GetVideo("VIDattr").remove : AttractvideoON = False
          FlexDMD.Stage.GetImage("BG012").visible = true

  End Select
  If frame mod 3 = 1 And Stoponlast = 0 then
    If LinePos1 > 0 Then LinePos1 = LinePos1 - 1
    If LinePos2 > 0 Then LinePos2 = LinePos2 - 1
    If LinePos3 > 0 Then LinePos3 = LinePos3 - 1
    If LinePos4 > 0 Then LinePos4 = LinePos4 - 1
  End If
  FlexDMD.Stage.GetLabel("TextSmalLine1").SetAlignedPosition 64,LinePos1 - 20 ,FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("TextSmalLine2").SetAlignedPosition 64,LinePos2 - 20 ,FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("TextSmalLine3").SetAlignedPosition 64,LinePos3 - 20 ,FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("TextSmalLine4").SetAlignedPosition 64,LinePos4 - 20 ,FlexDMD_Align_Center



End Sub
dim stoponlast
dim AttractvideoON



Sub DMD_Flasher
  Dim DMDp
    DMDp = FlexDMD.DmdColoredPixels
    If Not IsEmpty(DMDp) Then
    If VRRoom > 0 Then
        VRDMD5.DMDWidth = FlexDMD.Width
        VRDMD5.DMDHeight = FlexDMD.Height
        VRDMD5.DMDColoredPixels = DMDp
    Else
        VRDMD001.DMDWidth = FlexDMD.Width
        VRDMD001.DMDHeight = FlexDMD.Height
        VRDMD001.DMDColoredPixels = DMDp
    End If
    End If


End Sub



Function FormatScore(ByVal Num)
    dim NumString
    NumString = CStr(abs(Num) )
  If len(NumString)> 9 then NumString = left(NumString, Len(NumString)-9) & "," & right(NumString,9)
  If len(NumString)> 6 then NumString = left(NumString, Len(NumString)-6) & "," & right(NumString,6)
  If len(NumString)> 3 then NumString = left(NumString, Len(NumString)-3) & "," & right(NumString,3)


' If len(NumString)> 6 then NumString = left(NumString, Len(NumString)-6) & "," & right(NumString,6)
' If len(NumString)>3 then NumString = left(NumString, Len(NumString)-3) & "," & right(NumString,3)
  FormatScore = NumString
'len(NumString)< 11 And
End function

Function FormatScore2(ByVal Num) ' for highscores with all commas !
    dim NumString
    NumString = CStr(abs(Num) )
  If len(NumString)> 9 then NumString = left(NumString, Len(NumString)-9) & "," & right(NumString,9)
  If len(NumString)> 6 then NumString = left(NumString, Len(NumString)-6) & "," & right(NumString,6)
  If len(NumString)> 3 then NumString = left(NumString, Len(NumString)-3) & "," & right(NumString,3)
  FormatScore2 = NumString
End function


Sub Gate2_Hit
  DOF 127,2
End Sub

Sub Gate3_Hit
  DOF 128,2
End Sub


'******************************************************
'   STAND-UP TARGET INITIALIZATION
'******************************************************

Class StandupTarget
  Private m_primary, m_prim, m_sw, m_animate

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(value): Set m_primary = value: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(value): Set m_prim = value: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(asw): m_sw = asw: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(value): m_animate = value: End Property

  Public default Function init(primary, prim, sw, animate)
    Set m_primary = primary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate

    Set Init = Me
  End Function
End Class

'Define a variable for each stand-up target
Dim ST41, ST42, ST43, ST44, ST56, ST57, ST58, ST75, ST76

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:      vp target to determine target hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             transy must be used to offset the target animation
' switch:       ROM switch number
' animate:      Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST41 = (new StandupTarget)(sw41, sw41_BM_Lit_Room, 41, 0)
Set ST42 = (new StandupTarget)(sw42, sw42_BM_Lit_Room, 42, 0)
Set ST43 = (new StandupTarget)(sw43, sw43_BM_Lit_Room, 43, 0)
Set ST44 = (new StandupTarget)(sw44, sw44_BM_Lit_Room, 44, 0)
Set ST56 = (new StandupTarget)(sw56, sw56_BM_Lit_Room, 56, 0)
Set ST57 = (new StandupTarget)(sw57, sw57_BM_Lit_Room, 57, 0)
Set ST58 = (new StandupTarget)(sw58, sw58_BM_Lit_Room, 58, 0)
Set ST75 = (new StandupTarget)(sw75, sw75_BM_Lit_Room, 75, 0)
Set ST76 = (new StandupTarget)(sw76, sw76_001_BM_Lit_Room, 76, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST41, ST42, ST43, ST44, ST56, ST57, ST58, ST75, ST76)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5         'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.2        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance


'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
  Dim i
  i = STArrayID(switch)

  PlayTargetSound
  STArray(i).animate = STCheckHit(Activeball,STArray(i).primary)

  If STArray(i).animate <> 0 Then
    DTBallPhysics Activeball, STArray(i).primary.orientation, STMass
  End If
  DoSTAnim
End Sub

Function STArrayID(switch)
  Dim i
  For i = 0 To UBound(STArray)
    If STArray(i).sw = switch Then
      STArrayID = i
      Exit Function
    End If
  Next
End Function

Function STCheckHit(aBall, target) 'Check if target is hit on it's face
  Dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
  rangle = (target.orientation - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
  paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

  perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

  If perpvel > 0 And  perpvelafter <= 0 Then
    STCheckHit = 1
  ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
    STCheckHit = 1
  Else
    STCheckHit = 0
  End If
End Function

Sub DoSTAnim()
  Dim i
  For i = 0 To UBound(STArray)
    STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prim,STArray(i).sw,STArray(i).animate)
  Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
  Dim animtime

  STAnimate = animate

  If animate = 0  Then
    primary.uservalue = 0
    STAnimate = 0
    Exit Function
  ElseIf primary.uservalue = 0 Then
    primary.uservalue = gametime
  End If

  animtime = gametime - primary.uservalue

  If animate = 1 Then
    primary.collidable = 0
    prim.transy =  - STMaxOffset
    STAction switch
    STAnimate = 2
    Exit Function
  ElseIf animate = 2 Then
    prim.transy = prim.transy + STAnimStep
    If prim.transy >= 0 Then
      prim.transy = 0
      primary.collidable = 1
      STAnimate = 0
      Exit Function
    Else
      STAnimate = 2
    End If
  End If
End Function

Sub STAction(Switch)
' Select Case Switch
'   Case 41:
'   Case 42:
'   Case 43:
'   Case 44:
'   Case 56:
'   Case 57:
'   Case 58:
'   Case 75:
'   Case 76:
'   Case 77:
' End Select
End Sub

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


'******************************************************
'****   END STAND-UP TARGETS
'******************************************************



dim bl
If CabMode = 1 Then
  For each bl in PincabRails_BL : bl.visible = true : Next
Else
  For each bl in PincabRails_BL : bl.visible = false : Next
  If DesktopMode Then PinCab_Rails.visible = True
End If





' Flipper animations

Sub LeftFlipper_Animate
  Dim a, v, BL
  a = LeftFlipper.CurrentAngle
    FlipperLSh.RotZ = a
  For each BL in LFMesh_BL : BL.RotZ = a: Next
End Sub

Sub RightFlipper_Animate
  Dim a, v, BL
  a = RightFlipper.CurrentAngle
    FlipperRSh.RotZ = a
  For each BL in RFMesh_BL : BL.RotZ = a: Next
End Sub

Sub LeftFlipper1_Animate
  Dim a : a = LeftFlipper1.CurrentAngle
    FlipperLSh1.RotZ = a

  ' Darken light from lane bulbs when bats are up
  Dim v, BL
  v = 255.0 * (158.0 -  LeftFlipper1.CurrentAngle) / (158.0 -  110.0)

  For each BL in TLFMesh_BL
    BL.Rotz = a
    BL.visible = v < 128.0
  Next
  For each BL in TLFMeshU_BL
    BL.Rotz = a
    BL.visible = v >= 128.0
  Next
End Sub

Sub FlipperHeight
  Dim BL
  LFMesh_BM_Lit_Room.size_z = 1.3 : LFMesh_BM_Lit_Room.z = 2
  For each BL in LFMesh_BL : BL.size_z = 1.3 : BL.z = 2 : Next
  For each BL in RFMesh_BL : BL.size_z = 1.3 : BL.z = 2 :  Next
  For each BL in TLFMesh_BL : BL.size_z = 1.3 : BL.z = 2 :  Next
  For each BL in TLFMeshU_BL : BL.size_z = 1.3 : BL.z = 2 : Next
End Sub


' Gate animations



Sub Gate2_Animate
  Dim a : a = Gate2.CurrentAngle
  Dim BL : For Each BL in Gate2_BL : BL.rotx = -a: Next
End Sub

Sub Gate3_Animate
  Dim a : a = Gate3.CurrentAngle
  Dim BL : For Each BL in Gate3_BL : BL.rotx = -a: Next
End Sub

Sub LGate_Animate
  Dim a : a = LGate.CurrentAngle
  Dim BL : For Each BL in LGate_BL : BL.rotx = a: Next
End Sub

Sub RGate_Animate
  Dim a : a = RGate.CurrentAngle
  Dim BL : For Each BL in RGate_BL : BL.rotx = a: Next
End Sub

Sub RGate001_Animate
  Dim a : a = RGate001.CurrentAngle
  Dim BL : For Each BL in RGate001_BL : BL.rotx = a: Next
End Sub


' Switch animations

Sub sw16_Animate
  Dim z : z = sw16.CurrentAnimOffset
  Dim BL : For Each BL in sw16_BL : BL.transz = z: Next
End Sub

Sub sw17_Animate
  Dim z : z = sw17.CurrentAnimOffset
  Dim BL : For Each BL in sw17_BL : BL.transz = z: Next
End Sub

Sub sw26_Animate
  Dim z : z = sw26.CurrentAnimOffset
  Dim BL : For Each BL in sw26_BL : BL.transz = z: Next
End Sub

Sub sw27_Animate
  Dim z : z = sw27.CurrentAnimOffset
  Dim BL : For Each BL in sw27_BL : BL.transz = z: Next
End Sub

Sub sw38_Animate
  Dim z : z = sw38.CurrentAnimOffset
  Dim BL : For Each BL in sw38_BL : BL.transz = z: Next
End Sub

Sub sw48_Animate
  Dim z : z = sw48.CurrentAnimOffset
  Dim BL : For Each BL in sw48_BL : BL.transz = z: Next
End Sub

Sub sw62_Animate
  Dim z : z = sw62.CurrentAnimOffset
  Dim BL : For Each BL in sw62_BL : BL.transz = z: Next
End Sub

Sub sw71_Animate
  Dim z : z = sw71.CurrentAnimOffset
  Dim BL : For Each BL in sw71_BL : BL.transz = z: Next
End Sub

Sub sw72_Animate
  Dim z : z = sw72.CurrentAnimOffset
  Dim BL : For Each BL in sw72_BL : BL.transz = z: Next
End Sub

Sub sw73_Animate
  Dim z : z = sw73.CurrentAnimOffset
  Dim BL : For Each BL in sw73_BL : BL.transz = z: Next
End Sub

Sub sw74_Animate
  Dim z : z = sw74.CurrentAnimOffset
  Dim BL : For Each BL in sw74_BL : BL.transz = z: Next
End Sub

Sub swPlunger_Animate
  Dim z : z = swPlunger.CurrentAnimOffset
  Dim BL : For Each BL in swPlunger_BL : BL.transz = z: Next
End Sub


' Target animations


Dim DiverterPos1 : DiverterPos1 = 0
Dim DiverterPos2 : DiverterPos2 = 0
Dim DiverterON : DiverterON = 1


Sub Diverterupdate
  Dim lm
  'Diverters

  If DiverterON = 1 Then
    Diverter1.collidable = 1 : Diverter2.collidable = 0
    If DiverterPos1 > 0 Then DiverterPos1 = DiverterPos1 - 5
    If DiverterPos2 < 45 Then DiverterPos2 = DiverterPos2 + 5
  Else
    Diverter1.collidable = 0 : Diverter2.collidable = 1
    If DiverterPos2 > 0 Then DiverterPos2 = DiverterPos2 - 5
    If DiverterPos1 < 45 Then DiverterPos1 = DiverterPos1 + 5
  End If
  Diverter1_BM_Lit_Room.transz = -DiverterPos1
  Diverter2_BM_Lit_Room.transz = -DiverterPos2
  For each lm in Diverter1_LM : lm.transz = -DiverterPos1 : Next
  For each lm in Diverter2_LM : lm.transz = -DiverterPos2 : Next




' triggerwall up or down
  If TriggerWallPos > TriggerGate Then TriggerWallPos = TriggerWallPos - 5 : if TriggerWallPos = 0 Then Triggerbounce = 3
  If TriggerWallPos < TriggerGate Then TriggerWallPos = TriggerWallPos + 5

  If Triggerbounce > 0 Then Triggerbounce = Triggerbounce - 1

  If TriggerWallPos > 35 Then TriggerWall.collidable = False Else TriggerWall.collidable = True

  sw77_BM_Lit_Room.transz = - TriggerWallPos + Triggerbounce
  For each lm in sw77_lm : lm.transz = - TriggerWallPos + Triggerbounce: Next

End Sub

Sub sw77MovableHelper
  Dim lm, v
  'If sw77.IsDropped=True Then sw77_BM_Lit_Room.visible = False Else sw77_BM_Lit_Room.visible = True
  If TriggerWallCollide(CurrentPlayer) Then
'   sw77_BM_Lit_Room.visible = True
    TriggerGate = 0
  Else
    triggergate = 45
'   sw77_BM_Lit_Room.visible = False
  End If
' v = sw77_BM_Lit_Room.visible
' For each lm in sw77_lm : lm.visible = v : Next
End Sub

Sub turnon77
  dim lm
  sw77_BM_Lit_Room.visible = True
  For each lm in sw77_lm : lm.visible = True : Next
  Diverter1_BM_Lit_Room.visible = True
  For each lm in Diverter1_LM : lm.visible = True : Next
  Diverter2_BM_Lit_Room.visible = True
  For each lm in Diverter2_LM : lm.visible = True : Next
Diverter1_BM_Lit_Room.visible = 1:Diverter1.collidable = 1:Diverter2_BM_Lit_Room.visible = 1:Diverter2.collidable = 0
End Sub
turnon77


Sub TargetMovableHelper(Switch)
  Dim lm, y, z

  'Standup Targets
  Select Case switch
    Case 41
    y = sw41_BM_Lit_Room.transy
    For each lm in sw41_lm : lm.transy = y : Next

    Case 42
    y = sw42_BM_Lit_Room.transy
    For each lm in sw42_lm : lm.transy = y : Next

    Case 43
    y = sw43_BM_Lit_Room.transy
    For each lm in sw43_lm : lm.transy = y : Next

    Case 44
    y = sw44_BM_Lit_Room.transy
    For each lm in sw44_lm : lm.transy = y : Next

    Case 56
    y = sw56_BM_Lit_Room.transy
    For each lm in sw56_lm : lm.transy = y : Next

    Case 57
    y = sw57_BM_Lit_Room.transy
    For each lm in sw57_lm : lm.transy = y : Next

    Case 58
    y = sw58_BM_Lit_Room.transy
    For each lm in sw58_lm : lm.transy = y : Next

    Case 75
    y = sw75_BM_Lit_Room.transy
    For each lm in sw75_lm : lm.transy = y : Next

    Case 76
    y = sw76_001_BM_Lit_Room.transy
    For each lm in sw76_001_lm : lm.transy = y : Next

    Case 77
    y = sw77_BM_Lit_Room.transy
    For each lm in sw77_lm : lm.transy = y : Next
  End Select
End Sub


' Character animations


dim Ufo1_BM_Lit_Room_transx
dim Ufo1_BM_Lit_Room_transz
dim alien14_BM_Lit_Room_transy
dim alien5_BM_Lit_Room_transy
dim alien6_BM_Lit_Room_transy
dim alien8_BM_Lit_Room_transy

dim swp45_BM_Lit_Room_transx
dim swp45_BM_Lit_Room_transy
dim swp45_BM_Lit_Room_transz
dim swp45_BM_Lit_Room_z

dim swp46_BM_Lit_Room_transx
dim swp46_BM_Lit_Room_transy
dim swp46_BM_Lit_Room_transz
dim swp46_BM_Lit_Room_z

dim swp47_BM_Lit_Room_transx
dim swp47_BM_Lit_Room_transy
dim swp47_BM_Lit_Room_transz
dim swp47_BM_Lit_Room_z



Sub CharactersMovableHelper
  Dim bl, x, y, z
  'Characters
  x = Ufo1_BM_Lit_Room_transx
  z = Ufo1_BM_Lit_Room_transz
  For each bl in Ufo1_bl : bl.transx = x : bl.transz = z : Next

  y = alien14_BM_Lit_Room_transy
  For each bl in alien14_bl : bl.transz = y+alien14shake2/5 : bl.transx=alien14shake2/15 : : Next

  y = alien5_BM_Lit_Room_transy
  For each bl in alien5_bl : bl.transz = y+alien5shake2/5 : bl.transx=alien5shake2/15 : : Next

  y = alien6_BM_Lit_Room_transy
  For each bl in alien6_bl : bl.transz = y+alien6shake2/5 : bl.transx=alien6shake2/15 : Next

  y = alien8_BM_Lit_Room_transy
  For each bl in alien8_bl : bl.transz = y+spongeshaker2/5 : bl.transx=spongeshaker2/15 : Next
End Sub




'****************
'  Light Seqs
'****************

Sub StopTableSequences()
    InsertSequence.StopPlay()
    GISequence.StopPlay()
  lightCtrl.StopSyncWithVpxLights()
End Sub

Dim lSeqModeSelect: Set lSeqModeSelect = new LCSeq
lSeqModeSelect.Name = "lSeqModeSelect"


lSeqModeSelect.Sequence = Array( Array(), _
Array(), _
Array("l18|100","l19|100","l20|100", "l101|100", "l103|100"), _
Array("l101|0", "l103|0","l106|100", "l110|100","l18|0","l19|0","l20|0", "l24|100", "l25|100", "l26|100"),_
Array("l110|0", "l24|0","l25|0","l26|0", "l27|100", "l28|100", "l29|100"),_
Array("l105|100","l27|0","l28|0","l29|0", "l30|100", "l31|100", "l32|100"),_
Array("l107|100","l30|0","l31|0","l32|0", "l40|100", "l41|100", "l42|100"),_
Array("l105|0","l40|0","l41|0","l42|0", "l43|100", "l44|100", "l45|100"),_
Array("l111|100","l43|0","l44|0","l45|0", "l46|100", "l47|100", "l48|100"),_
Array("l107|0","l111|0","l102|100","l104|100","l46|0","l47|0","l48|0", "l21|100", "l22|100", "l23|100"),_
Array("l107|100","l111|100","l102|0","l104|0","l21|0","l23|0","l24|0", "l46|100", "l47|100", "l48|100"),_
Array("l111|0","l46|0","l47|0","l48|0", "l43|100", "l44|100", "l46|100"),_
Array("l105|100","l43|0","l44|0","l46|0", "l40|100", "l41|100", "l42|100"),_
Array("l107|0","l106|100","l40|0","l41|0","l42|0", "l30|100", "l31|100", "l32|100"),_
Array("l105|0","l30|0","l31|0","l32|0", "l27|100", "l28|100", "l29|100"),_
Array("l110|100","l27|0","l28|0","l29|0", "l24|100", "l25|100", "l26|100"),_
Array("l24|0","l25|0","l26|0", "l106|0", "l110|0"))



Dim lSeqRainbow : Set lSeqRainbow = New LCSeq
lSeqRainbow.Name = "lSeqRainbow"
lSeqRainbow.Sequence = Array( Array("l101|100|0051B6","l102|100|1B593C","l103|100|0000C8","l104|100|C87900","l105|100|BF05A0","l106|100|7A29B9","l107|100|C70001","l108|100|C87900","l109|100|1B593C","l110|100|BF05A0","l111|100|C4023B","l115|100|0051B6"), _
Array("l102|100|1F5A3A","l103|100|0051B6","l104|100|C8401F","l105|100|8424B5","l106|100|7A2AB9","l107|100|BF05A0","l110|100|7A2AB9","l111|100|BF05A0","l115|100|1B593C"), _
Array("l101|100|1B593C","l102|100|C87900","l105|100|7A2AB9"), _
Array("l109|100|C87900"), _
Array("l103|100|1B593C","l104|100|C7170B","l106|100|0000C8","l107|100|7A2AB9","l108|100|C8401F","l111|100|7A2AB9"), _
Array("l101|100|C87900","l102|100|C8401F","l104|100|C70001","l105|100|0000C8","l110|100|0000C8","l115|100|C87900"), _
Array(), _
Array("l106|100|0051B6","l107|100|0000C8","l108|100|C70001","l109|100|C8401F","l111|100|0000C8"), _
Array("l102|100|C70001","l103|100|C87900","l104|100|BF05A0","l105|100|0051B6","l110|100|0051B6","l115|100|C8401F"), _
Array("l101|100|C8401F"), _
Array("l106|100|085393","l108|100|BF05A0","l109|100|C70001","l111|100|0050B6"), _
Array("l103|100|C8401F","l104|100|7A2AB9","l105|100|085392","l106|100|1B593C","l107|100|0051B6","l110|100|1B593C","l111|100|0051B6","l115|100|C70001"), _
Array("l101|100|C70001","l102|100|BF05A0","l105|100|1B593C"), _
Array("l109|100|BF05A0"), _
Array("l104|100|0000C8","l106|100|C87900","l107|100|1B593C","l108|100|7A2AB9","l111|100|1B593C"), _
Array("l101|100|BF05A0","l102|100|7A2AB9","l103|100|C70001","l105|100|C87900","l110|100|C87900","l115|100|BF05A0"), _
Array(), _
Array("l106|100|C8441C","l107|100|2F5D35","l108|100|0000C8","l109|100|7A2AB9","l111|100|C87900"), _
Array("l101|100|7C29B8","l102|100|0000C8","l103|100|BF05A0","l104|100|0051B6","l106|100|C8401F","l107|100|C87900","l110|100|C8401F","l115|100|7A2AB9"), _
Array("l101|100|7A2AB9","l105|100|C8401F"), _
Array("l108|100|0051B6","l109|100|0000C8"), _
Array("l102|100|000AC6","l103|100|7A2AB9","l104|100|1B593C","l106|100|C70001","l107|100|C8401F","l111|100|C8401F","l115|100|0000C8"), _
Array("l101|100|0000C8","l102|100|0051B6","l105|100|C70001","l110|100|C70001"), _
Array("l109|100|0051B6"), _
Array("l103|100|571EBD","l106|100|BF05A0","l107|100|C70001","l108|100|1B593C","l111|100|C70001"), _
Array("l101|100|0051B6","l102|100|1B593C","l103|100|0000C8","l104|100|C87900","l105|100|BF05A0","l110|100|BF05A0","l115|100|0051B6"), _
Array(), _
Array("l106|100|7A2AB9","l108|100|C87900","l109|100|1B593C","l111|100|BF05A0"), _
Array("l102|100|C87900","l103|100|0051B6","l104|100|C8401F","l105|100|7A2AB9","l107|100|BF05A0","l110|100|7A2AB9","l115|100|1B593C"), _
Array("l101|100|1B593C"), _
Array("l109|100|C87900"), _
Array("l103|100|1B593C","l104|100|C70001","l106|100|0000C8","l107|100|7A2AB9","l108|100|C8401F","l111|100|7A2AB9","l115|100|C87900"), _
Array("l101|100|C87900","l102|100|C8401F","l105|100|0000C8","l110|100|0000C8"), _
Array(), _
Array("l106|100|0051B6","l107|100|0000C8","l108|100|C70001","l109|100|C8401F","l111|100|0000C8"), _
Array("l101|100|C8401F","l102|100|C70001","l103|100|C87900","l104|100|BF05A0","l105|100|0051B6","l110|100|0051B6","l115|100|C8401F"), _
Array(), _
Array("l106|100|1B593C","l107|100|0023C0","l108|100|BF05A0","l109|100|C70001","l111|100|0051B6"), _
Array("l102|100|C40345","l103|100|C8401F","l104|100|7A2AB9","l105|100|1B593C","l107|100|0051B6","l110|100|1B593C","l115|100|C70001"), _
Array("l101|100|C70001","l102|100|BF05A0"), _
Array("l108|100|7A2AB9","l109|100|BF05A0"), _
Array("l104|100|0000C8","l106|100|C87900","l107|100|1B593C","l111|100|1B593C"), _
Array("l101|100|BF05A0","l102|100|7A2AB9","l103|100|C70001","l105|100|C87900","l110|100|C87900","l115|100|BF05A0"), _
Array("l109|100|B20DA5"), _
Array("l104|100|0049B8","l106|100|C8401F","l107|100|C87900","l108|100|0000C8","l109|100|7A2AB9","l111|100|C87900"), _
Array("l101|100|7A2AB9","l102|100|0000C8","l103|100|BF05A0","l104|100|0051B6","l105|100|C85D0F","l110|100|C8401F","l115|100|7A2AB9"))
lSeqRainbow.UpdateInterval = 20
lSeqRainbow.Color = Null
lSeqRainbow.Repeat = False



lSeqModeSelect.UpdateInterval = 100
lSeqModeSelect.color = Null
lSeqModeSelect.Repeat = False

lightCtrl.CreateSeqRunner "lSeqRunnerModeSelect"


'Hack for spongebob as blender export doesn't match control light names.
l103.UserValue = "GIString3"
l104.UserValue = "GIString4"
l105.UserValue = "GIString5"
l106.UserValue = "GIString6"
l107.UserValue = "GIString7"
l108.UserValue = "GIString8"
l111.UserValue = "GIString11"


l101.UserValue = "GIString1"
l102.UserValue = "GIString2"
l115.UserValue = "Spotlights"

'***********************************************************************************************************************
' Lights State Controller - 8.0.0
'
' A light state controller for original vpx tables.
'
' Documentation: https://github.com/mpcarr/vpx-light-controller
'
'***********************************************************************************************************************

Class LStateController

    Private m_currentFrameState, m_on, m_off, m_seqRunners, m_lights, m_seqs, m_vpxLightSyncRunning, m_vpxLightSyncClear, m_vpxLightSyncCollection, m_tableSeqColor, m_tableSeqFadeUp, m_tableSeqFadeDown, m_frametime, m_initFrameTime, m_pulse, m_pulseInterval, useVpxLights, m_lightmaps, m_seqOverrideRunners

    Private Sub Class_Initialize()
        Set m_lights = CreateObject("Scripting.Dictionary")
        Set m_on = CreateObject("Scripting.Dictionary")
        Set m_off = CreateObject("Scripting.Dictionary")
        Set m_seqRunners = CreateObject("Scripting.Dictionary")
        Set m_seqOverrideRunners = CreateObject("Scripting.Dictionary")
        Set m_currentFrameState = CreateObject("Scripting.Dictionary")
        Set m_seqs = CreateObject("Scripting.Dictionary")
        Set m_pulse = CreateObject("Scripting.Dictionary")
        Set m_on = CreateObject("Scripting.Dictionary")
        m_vpxLightSyncRunning = False
        m_vpxLightSyncCollection = Null
    m_initFrameTime = 0
        m_frameTime = 0
        m_pulseInterval = 26
        m_vpxLightSyncClear = False
        m_tableSeqColor = Null
        m_tableSeqFadeUp = Null
        m_tableSeqFadeDown = Null
        useVpxLights = False
        Set m_lightmaps = CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub AssignStateForFrame(key, state)
        If m_currentFrameState.Exists(key) Then
            m_currentFrameState.Remove key
        End If
        m_currentFrameState.Add key, state
    End Sub

    Public Sub LoadLightShows()
        Dim oFile
        Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
        Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/lights-out.txt",2,true)
        For Each oFile In oFSO.GetFolder(cGameName & "_LightShows").Files
            If LCase(oFSO.GetExtensionName(oFile.Name)) = "yaml" And Not Left(oFile.Name,6) = "lights" Then
                Dim textStream : Set textStream = oFSO.OpenTextFile(oFile.Path, 1)
                Dim show : show = textStream.ReadAll
                Dim fileName : fileName = "lSeq" & Replace(oFSO.GetFileName(oFile.Name), "."&oFSO.GetExtensionName(oFile.Name), "")
                Dim lcSeq : lcSeq = "Dim " & fileName & " : Set " & fileName & " = New LCSeq"&vbCrLf
                lcSeq = lcSeq + fileName & ".Name = """&fileName&""""&vbCrLf
                Dim seq : seq = ""
                Dim re : Set re = New RegExp
                With re
                    .Pattern    = "- time:.*?\n"
                    .IgnoreCase = False
                    .Global     = True
                End With
                Dim matches : Set matches = re.execute(show)
                Dim steps : steps = matches.Count
                Dim match, nextMatchIndex, uniqueLights
                Set uniqueLights = CreateObject("Scripting.Dictionary")
                nextMatchIndex = 1
                For Each match in matches
                    Dim lightStep
                    If Not nextMatchIndex < steps Then
                        lightStep = Mid(show, match.FirstIndex, Len(show))
                    Else
                        lightStep = Mid(show, match.FirstIndex, matches(nextMatchIndex).FirstIndex - match.FirstIndex)
                        nextMatchIndex = nextMatchIndex + 1
                    End If

                    Dim re1 : Set re1 = New RegExp
                    With re1
                        .Pattern        = ".*:?: '([A-Fa-f0-9]{6})'"
                        .IgnoreCase     = True
                        .Global         = True
                    End With

                    Dim lightMatches : Set lightMatches = re1.execute(lightStep)
                    If lightMatches.Count > 0 Then
                        Dim lightMatch, lightStr, lightSplit
                        lightStr = "Array("
                        lightSplit = 0
                        For Each lightMatch in lightMatches
                            Dim sParts : sParts = Split(lightMatch.Value, ":")
                            Dim lightName : lightName = Trim(sParts(0))
                            Dim color : color = Trim(Replace(sParts(1),"'", ""))
                            If color = "000000" Then
                                lightStr = lightStr + """"&lightName&"|0|000000"","
                            Else
                                lightStr = lightStr + """"&lightName&"|100|"&color&""","
                            End If

                            If Len(lightStr)+20 > 2000 And lightSplit = 0 Then
                                lightSplit = Len(lightStr)
                            End If

                            uniqueLights(lightname) = 0
                        Next
                        lightStr = Left(lightStr, Len(lightStr) - 1)
                        lightStr = lightStr & ")"

                        If lightSplit > 0 Then
                            lightStr = Left(lightStr, lightSplit) & " _ " & vbCrLF & Right(lightStr, Len(lightStr)-lightSplit)
                        End If

                        seq = seq + lightStr & ", _"&vbCrLf
                    Else
                        seq = seq + "Array(), _"&vbCrLf
                    End If


                    Set re1 = Nothing
                Next

                lcSeq = lcSeq + filename & ".Sequence = Array( " & Left(seq, Len(seq) - 5) & ")"&vbCrLf
                'lcSeq = lcSeq + seq & vbCrLf
                lcSeq = lcSeq + fileName & ".UpdateInterval = 20"&vbCrLf
                lcSeq = lcSeq + fileName & ".Color = Null"&vbCrLf
                lcSeq = lcSeq + fileName & ".Repeat = False"&vbCrLf

                'MsgBox(lcSeq)
                objFileToWrite.WriteLine(lcSeq)
                ExecuteGlobal lcSeq
                Set re = Nothing

                textStream.Close
            End if
        Next
        'Clean up
        objFileToWrite.Close
        Set objFileToWrite = Nothing
        Set oFile = Nothing
        Set oFSO = Nothing
    End Sub

    Public Sub CompileLights(collection, name)
        Dim light
        Dim lights : lights = "light:" & vbCrLf
        For Each light in collection
            lights = lights + light.name & ":"&vbCrLf
            lights = lights + "   x: "& light.x/tablewidth & vbCrLf
            lights = lights + "   y: "& light.y/tableheight & vbCrLf
        Next
        Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/lights-"&name&".yaml",2,true)
      objFileToWrite.WriteLine(lights)
      objFileToWrite.Close
      Set objFileToWrite = Nothing
 '       Debug.print("Lights YAML File saved to: " & cGameName & "LightShows/lights-"&name&".yaml")
    End Sub

    Public Sub RegisterLights(mode)

        Dim idx,tmp,vpxLight,lcItem
        If mode = "Lampz" Then

            For idx = 0 to UBound(Lampz.obj)
                If Lampz.IsLight(idx) Then
                    Set lcItem = new LCItem
                    If IsArray(Lampz.obj(idx)) Then
                        tmp = Lampz.obj(idx)
                        Set vpxLight = tmp(0)
                    Else
                        Set vpxLight = Lampz.obj(idx)

                    End If
                    Lampz.Modulate(idx) = 1/100
                    Lampz.FadeSpeedUp(idx) = 100/30 : Lampz.FadeSpeedDown(idx) = 100/120
                    lcItem.Init idx, vpxLight.BlinkInterval, Array(vpxLight.color, vpxLight.colorFull), vpxLight.name, vpxLight.x, vpxLight.y

                    m_lights.Add vpxLight.Name, lcItem
                    m_seqRunners.Add "lSeqRunner" & CStr(vpxLight.name), new LCSeqRunner
                End If
            Next
        ElseIf mode = "VPX" Then
            useVpxLights = True


            For idx = 0 to UBound(Lights)
                vpxLight = Null
                Set lcItem = new LCItem
                If IsArray(Lights(idx)) Then
                    tmp = Lights(idx)
                    Set vpxLight = tmp(0)
                ElseIf IsObject(Lights(idx)) Then
                    Set vpxLight = Lights(idx)
                End If
        If Not IsNull(vpxLight) Then
                    Dim e, lmStr: lmStr = "lmArr = Array("
                    For Each e in GetElements()
                        If Right(LCase(e.Name), Len(vpxLight.Name)+1) = "_" & LCase(vpxLight.Name) Or Right(LCase(e.Name), Len(vpxLight.UserValue)+1) = "_" & LCase(vpxLight.UserValue) Then
  '                          Debug.Print(e.Name)
                            lmStr = lmStr & e.Name & ","
                        End If

            If InStr(e.Name, "_" & vpxLight.Name & "_") Or InStr(e.Name, "_" & vpxLight.UserValue & "_") Then
   '                         Debug.Print(e.Name)
                            lmStr = lmStr & e.Name & ","
                        End If

                    Next
                    lmStr = lmStr & "Null)"
                    lmStr = Replace(lmStr, ",Null)", ")")
              ExecuteGlobal "Dim lmArr : "&lmStr
                    m_lightmaps.Add vpxLight.Name, lmArr
      '              Debug.print("Registering Light: "& vpxLight.name)
                    lcItem.Init idx, vpxLight.BlinkInterval, Array(vpxLight.color, vpxLight.colorFull), vpxLight.name, vpxLight.x, vpxLight.y
                    m_lights.Add vpxLight.Name, lcItem
                    m_seqRunners.Add "lSeqRunner" & CStr(vpxLight.name), new LCSeqRunner
                End If
            Next
      SyncLightMapColors()
        End If
    End Sub

    Private Function ColtoArray(aDict)  'converts a collection to an indexed array. Indexes will come out random probably.
        redim a(999)
        dim count : count = 0
        dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
        redim preserve a(count-1) : ColtoArray = a
    End Function

  Public Sub AddLight(light, idx)
        If m_lights.Exists(light.name) Then
            Exit Sub
        End If
        Dim lcItem : Set lcItem = new LCItem
        lcItem.Init idx, light.BlinkInterval, Array(light.color, light.colorFull), light.name, light.x, light.y
        m_lights.Add light.Name, lcItem
        m_seqRunners.Add "lSeqRunner" & CStr(light.name), new LCSeqRunner
    End Sub

    Public Sub LightState(light, state)
        m_lightOff(light.name)
        If state = 1 Then
            m_lightOn(light.name)
        ElseIF state = 2 Then
            Blink(light)
        End If
    End Sub

    Public Sub LightOn(light)
        m_LightOn(light.name)
    End Sub

    Public Sub LightOnWithColor(light, color)
        m_LightOnWithColor light.name, color
    End Sub

    Public Sub FlickerOn(light)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            m_lightOn(name)

            If m_pulse.Exists(name) Then
                Exit Sub
            End If
            m_pulse.Add name, (new PulseState)(m_lights(name), Array(37,100,24,0,70), 0, m_pulseInterval, 1)
        End If
    End Sub

    Public Sub LightColor(light, color)
        If m_lights.Exists(light.name) Then
            m_lights(light.name).Color = color
            'Update internal blink seq for light
            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").Color = color
            End If

        End If
    End Sub

    Private Sub m_LightOn(name)
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then
                m_off.Remove(name)
            End If
            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If
            If m_on.Exists(name) Then
                Exit Sub
            End If
            m_on.Add name, m_lights(name)
        End If
    End Sub

    Private Sub m_LightOnWithColor(name, color)
        If m_lights.Exists(name) Then
            m_lights(name).Color = color
            If m_off.Exists(name) Then
                m_off.Remove(name)
            End If

            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If

            If m_on.Exists(name) Then
                Exit Sub
            End If
            m_on.Add name, m_lights(name)
        End If
    End Sub

    Public Sub LightOff(light)
        m_lightOff(light.name)
    End Sub

    Private Sub m_lightOff(name)
        If m_lights.Exists(name) Then
            If m_on.Exists(name) Then
                m_on.Remove(name)
            End If

            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If

            If m_off.Exists(name) Then
                Exit Sub
            End If
            m_off.Add name, m_lights(name)
        End If
    End Sub

    Public Sub UpdateBlinkInterval(light, interval)
        If m_lights.Exists(light.name) Then
            light.BlinkInterval = interval
            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs.Item(light.name & "Blink").UpdateInterval = interval
            End If
        End If
    End Sub


    Public Sub Pulse(light, repeatCount)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then
                m_off.Remove(name)
            End If
            If m_pulse.Exists(name) Then
                Exit Sub
            End If
            'Array(100,94,32,13,6,3,0)
            m_pulse.Add name, (new PulseState)(m_lights(name), Array(37,100,24,0,70,100,12,0), 0, m_pulseInterval, repeatCount)
        End If
    End Sub

    Public Sub PulseWithProfile(light, profile, repeatCount)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then
                m_off.Remove(name)
            End If
            If m_pulse.Exists(name) Then
                Exit Sub
            End If
            m_pulse.Add name, (new PulseState)(m_lights(name), profile, 0, m_pulseInterval, repeatCount)
        End If
    End Sub

    Public Sub LightLevel(light, lvl)
        If m_lights.Exists(light.name) Then
            m_lights(light.name).Level = lvl

            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").Sequence = m_buildBlinkSeq(light)
            End If
        End If
    End Sub


    Public Sub AddShot(name, light, color)
        If m_lights.Exists(light.name) Then
            If m_seqs.Exists(name & light.name) Then
                m_seqs(name & light.name).Color = color
                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem m_seqs(name & light.name)
            Else
                Dim stateOn : stateOn = light.name&"|100"
                Dim stateOff : stateOff = light.name&"|0"
                Dim seq : Set seq = new LCSeq
                seq.Name = name
                seq.Sequence = Array(stateOn, stateOff,stateOn, stateOff)
                seq.Color = color
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem seq
                m_seqs.Add name & light.name, seq
            End If
            If m_on.Exists(light.name) Then
                m_on.Remove light.name
            End If
        End If
    End Sub

    Public Sub RemoveShot(name, light)
        If m_lights.Exists(light.name) And m_seqs.Exists(name & light.name) Then
            m_seqRunners("lSeqRunner"&CStr(light.name)).RemoveItem m_seqs(name & light.name)
            If IsNUll(m_seqRunners("lSeqRunner"&CStr(light.name)).CurrentItem) Then
               LightOff(light)
            End If
        End If
    End Sub

    Public Sub RemoveAllShots()
        Dim light
        For Each light in m_lights.Keys()
            m_seqRunners("lSeqRunner"&CStr(light)).RemoveAll
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
        Next
    End Sub

    Public Sub RemoveShotsFromLight(light)
        If m_lights.Exists(light.name) Then
            m_seqRunners("lSeqRunner"&CStr(light.name)).RemoveAll
            m_lightOff(light.name)
        End If
    End Sub

    Public Sub Blink(light)
        If m_lights.Exists(light.name) Then

            If m_seqs.Exists(light.name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem m_seqs(light.name & "Blink")
            Else
                Dim seq : Set seq = new LCSeq
                seq.Name = light.name & "Blink"
                seq.Sequence = m_buildBlinkSeq(light)
                seq.Color = Null
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem seq
                m_seqs.Add light.name & "Blink", seq
            End If
            If m_on.Exists(light.name) Then
                m_on.Remove light.name
            End If
        End If
    End Sub

    Public Function GetLightState(light)
        GetLightState = 0
        If(m_lights.Exists(light.name)) Then
            If m_on.Exists(light.name) Then
                GetLightState = 1
            Else
                If m_seqs.Exists(light.name & "Blink") Then
                    GetLightState = 2
                End If
            End If
        End If
    End Function

    Public Function IsShotLit(name, light)
        IsShotLit = False
        If(m_lights.Exists(light.name)) Then
            If m_seqRunners("lSeqRunner"&CStr(light.name)).HasSeq(name) Then
                IsShotLit = True
            End If
        End If
    End Function

    Public Sub CreateSeqRunner(name)
        If m_seqRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        m_seqRunners.Add name, seqRunner
    End Sub

    Private Sub CreateOverrideSeqRunner(name)
        If m_seqOverrideRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        m_seqOverrideRunners.Add name, seqRunner
    End Sub

    Public Sub AddLightSeq(lcSeqRunner, lcSeq)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If

        m_seqRunners(lcSeqRunner).AddItem lcSeq
    End Sub

    Public Sub RemoveLightSeq(lcSeqRunner, lcSeq)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If

        Dim light
        For Each light in lcSeq.LightsInSeq
            If(m_lights.Exists(light)) Then
                    AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            End If
        Next

        m_seqRunners(lcSeqRunner).RemoveItem lcSeq
    End Sub

    Public Sub RemoveAllLightSeq(lcSeqRunner)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If
        Dim lcSeqKey, light, seqs, lcSeq
        Set seqs = m_seqRunners(lcSeqRunner).Items()
        For Each lcSeqKey in seqs.Keys()
      Set lcSeq = seqs(lcSeqKey)
            For Each light in lcSeq.LightsInSeq
                If(m_lights.Exists(light)) Then
                    AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
                End If
            Next
        Next

        m_seqRunners(lcSeqRunner).RemoveAll
    End Sub

    Public Sub AddTableLightSeq(name, lcSeq)
        CreateOverrideSeqRunner(name)

        Dim seqOverride, hasOverride
        hasOverride = False
        For Each seqOverride In m_seqOverrideRunners.Keys()
            If Not IsNull(m_seqOverrideRunners(seqOverride).CurrentItem) Then
                hasOverride = True
            End If
        Next
        If hasOverride = False Then
            Dim light
            For Each light in m_lights.Keys()
                AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            Next
        End If
        m_seqOverrideRunners(name).AddItem lcSeq
    End Sub

    Public Sub RemoveTableLightSeq(name, lcSeq)
        m_seqOverrideRunners(name).RemoveItem lcSeq
        Dim seqOverride, hasOverride
        hasOverride = False
        For Each seqOverride In m_seqOverrideRunners.Keys()
            If Not IsNull(m_seqOverrideRunners(seqOverride).CurrentItem) Then
                hasOverride = True
            End If
        Next
        If hasOverride = False Then
            Dim light
            For Each light in m_lights.Keys()
                AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            Next
        End If
    End Sub

    Public Sub RemoveAllTableLightSeqs()
        Dim light, runner
        For Each runner in m_seqOverrideRunners.Keys()
            m_seqOverrideRunners(runner).RemoveAll()
        Next
    For Each light in m_lights.Keys()
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
        Next
    End Sub

   Public Sub SyncLightMapColors()
        dim light,lm
        For Each light in m_lights.Keys()
            If m_lightmaps.Exists(light) Then
                For Each lm in m_lightmaps(light)
                    dim color : color = m_lights(light).Color
                    If not IsNull(lm) Then
            lm.Color = color(0)
          End If
                Next
            End If
        Next
    End Sub

    Public Sub SyncWithVpxLights(lightSeq)
        m_vpxLightSyncCollection = ColToArray(eval(lightSeq.collection))
        m_vpxLightSyncRunning = True
    End Sub

    Public Sub StopSyncWithVpxLights()
        m_vpxLightSyncRunning = False
        m_vpxLightSyncClear = True
    m_tableSeqColor = Null
        m_tableSeqFadeUp = Null
        m_tableSeqFadeDown = Null
    End Sub

  Public Sub SetVpxSyncLightColor(color)
    m_tableSeqColor = color
  End Sub

    Public Sub SetTableSequenceFade(fadeUp, fadeDown)
    m_tableSeqFadeUp = fadeUp
        m_tableSeqFadeDown = fadeDown
  End Sub

    Public Sub UseToolkitColoredLightMaps()
        If useVpxLights = True Then
            Exit Sub
        End If

        Dim sUpdateLightMap
        sUpdateLightMap = "Sub UpdateLightMap(idx, lightmap, intensity, ByVal aLvl)" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   if Lampz.UseFunc then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   lightmap.Opacity = aLvl * intensity" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   If IsArray(Lampz.obj(idx) ) Then" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "       lightmap.Color = Lampz.obj(idx)(0).color" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   Else" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "       lightmap.color = Lampz.obj(idx).color" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   End If" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "End Sub" + vbCrLf

        ExecuteGlobal sUpdateLightMap

        Dim x
        For x=0 to Ubound(Lampz.cCallback)
            Lampz.cCallback(x) = Replace(Lampz.cCallback(x), "UpdateLightMap ", "UpdateLightMap " & x & ",")
            Lampz.Callback(x) = "" 'Force Callback Sub to be build
        Next
    End Sub

    Private Function m_buildBlinkSeq(light)
        Dim i, buff : buff = Array()
        ReDim buff(Len(light.BlinkPattern)-1)
        For i = 0 To Len(light.BlinkPattern)-1

            If Mid(light.BlinkPattern, i+1, 1) = 1 Then
                buff(i) = light.name & "|100"
            Else
                buff(i) = light.name & "|0"
            End If
        Next
        m_buildBlinkSeq=buff
    End Function

    Private Function GetTmpLight(idx)
        If useVpxLights = True Then
          If IsArray(Lights(idx) ) Then 'if array
                Set GetTmpLight = Lights(idx)(0)
            Else
                Set GetTmpLight = Lights(idx)
            End If
        Else
            If IsArray(Lampz.obj(idx) ) Then  'if array
                Set GetTmpLight = Lampz.obj(idx)(0)
            Else
                Set GetTmpLight = Lampz.obj(idx)
            End If
        End If

    End Function

    Public Sub ResetLights()
        Dim light
        For Each light in m_lights.Keys()
            m_seqRunners("lSeqRunner"&CStr(light)).RemoveAll
            m_lightOff(light)
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
        Next
        RemoveAllTableLightSeqs()
        Dim k
        For Each k in m_seqRunners.Keys()
            Dim lsRunner: Set lsRunner = m_seqRunners(k)
            lsRunner.RemoveAll
        Next

    End Sub

    Public Sub Update()

    m_frameTime = gametime - m_initFrameTime : m_initFrameTime = gametime
    Dim x
        Dim lk
        dim color
        Dim lightKey
        Dim lcItem
        Dim tmpLight
        Dim seqOverride, hasOverride
        hasOverride = False
        For Each seqOverride In m_seqOverrideRunners.Keys()
            If Not IsNull(m_seqOverrideRunners(seqOverride).CurrentItem) Then
                RunLightSeq m_seqOverrideRunners(seqOverride)
                hasOverride = True
            End If
        Next
        If hasOverride = False Then

            If HasKeys(m_on) Then
                For Each lightKey in m_on.Keys()
                    Set lcItem = m_on(lightKey)
                    AssignStateForFrame lightKey, (new FrameState)(lcItem.level, m_on(lightKey).Color, m_on(lightKey).Idx)
                Next
            End If

            If HasKeys(m_pulse) Then
                For Each lightKey in m_pulse.Keys()
                    AssignStateForFrame lightKey, (new FrameState)(m_pulse(lightKey).PulseAt(m_pulse(lightKey).idx), m_pulse(lightKey).light.Color, m_pulse(lightKey).light.Idx)
                    Dim pulseUpdateInt : pulseUpdateInt = m_pulse(lightKey).interval - m_frameTime
                    Dim pulseIdx : pulseIdx = m_pulse(lightKey).idx
                    If pulseUpdateInt <= 0 Then
                        pulseUpdateInt = m_pulseInterval
                        pulseIdx = pulseIdx + 1
                    End If

                    Dim pulses : pulses = m_pulse(lightKey).pulses
          Dim pulseCount : pulseCount = m_pulse(lightKey).Cnt
                    If pulseIdx > UBound(m_pulse(lightKey).pulses) Then
            m_pulse.Remove lightKey
            If pulseCount > 0 Then
                            pulseCount = pulseCount - 1
                            pulseIdx = 0
                            m_pulse.Add lightKey, (new PulseState)(m_lights(lightKey),pulses, pulseIdx, pulseUpdateInt, pulseCount)
                        End If
                    Else
            m_pulse.Remove lightKey
                        m_pulse.Add lightKey, (new PulseState)(m_lights(lightKey),pulses, pulseIdx, pulseUpdateInt, pulseCount)
                    End If
                Next
            End If

            If HasKeys(m_off) Then
                For Each lightKey in m_off.Keys()
                    Set lcItem = m_off(lightKey)
                    AssignStateForFrame lightKey, (new FrameState)(0, Null, lcItem.Idx)
                Next
            End If

            If HasKeys(m_seqRunners) Then
                Dim k
                For Each k in m_seqRunners.Keys()
                    Dim lsRunner: Set lsRunner = m_seqRunners(k)
                    If Not IsNull(lsRunner.CurrentItem) Then
                            RunLightSeq lsRunner
                    End If
                Next
            End If

            If m_vpxLightSyncRunning = True Then
                Dim lx
                If Not IsNull(m_vpxLightSyncCollection) Then
                    For Each lx in m_vpxLightSyncCollection
                        'sync each light being ran by the vpx LS
                        dim syncLight : syncLight = Null
                        If m_lights.Exists(lx.name) Then
                            'found a light
                            Set syncLight = m_lights(lx.name)
                        End If
                        If Not IsNull(syncLight) Then
                            'Found a light to sync.
                            Dim lightState

                            If IsNull(m_tableSeqColor) Then
                                color = syncLight.Color
                            Else
                                If Not IsArray(m_tableSeqColor) Then
                                    color = Array(m_TableSeqColor, Null)
                                Else
                                    color = m_tableSeqColor
                                End If
                            End If

                            'TODO - Fix VPX Fade
                            If Not useVpxLights = True Then
                                If Not IsNull(m_tableSeqFadeUp) Then
                                    Lampz.FadeSpeedUp(syncLight.Idx) = m_tableSeqFadeUp
                                End If
                                If Not IsNull(m_tableSeqFadeDown) Then
                                    Lampz.FadeSpeedDown(syncLight.Idx) = m_tableSeqFadeDown
                                End If
                            End If

                            AssignStateForFrame syncLight.name, (new FrameState)(lx.GetInPlayState*100,color, syncLight.Idx)
                        End If
                    Next
            End If
            End If

            If m_vpxLightSyncClear = True Then
                If Not IsNull(m_vpxLightSyncCollection) Then
                    For Each lk in m_vpxLightSyncCollection
                        'sync each light being ran by the vpx LS
                        dim syncClearLight : syncClearLight = Null
                        If m_lights.Exists(lk.name) Then
                            'found a light
                            Set syncClearLight = m_lights(lk.name)
                        End If
                        If Not IsNull(syncClearLight) Then
                            AssignStateForFrame syncClearLight.name, (new FrameState)(0, Null, syncClearLight.idx)
                            'TODO - Only do fade speed for lampz
                            If Not useVpxLights = True Then
                                Lampz.FadeSpeedUp(syncClearLight.Idx) = 100/30
                                Lampz.FadeSpeedDown(syncClearLight.Idx) = 100/120
                            End If
                        End If
                    Next
                End If

                m_vpxLightSyncClear = False
            End If
        End If


        If HasKeys(m_currentFrameState) Then

            Dim frameStateKey
            For Each frameStateKey in m_currentFrameState.Keys()
                Dim idx : idx = m_currentFrameState(frameStateKey).idx

                Dim newColor : newColor = m_currentFrameState(frameStateKey).colors
                Dim bUpdate

                If Not IsNull(newColor) Then
                    'Check current color is the new color coming in, if not, set the new color.

                    Set tmpLight = GetTmpLight(idx)

          Dim c, cf
          c = newColor(0)
          cf= newColor(1)

          If Not IsNull(c) Then
            If Not CStr(tmpLight.Color) = CStr(c) Then
              bUpdate = True
            End If
          End If

          If Not IsNull(cf) Then
            If Not CStr(tmpLight.ColorFull) = CStr(cf) Then
              bUpdate = True
            End If
          End If
              End If

                If useVpxLights = False Then
                    If bUpdate Then
                        'Update lamp color
                        If IsArray(Lampz.obj(idx)) Then
                            for each x in Lampz.obj(idx)
                                If Not IsNull(c) Then
                                    x.color = c
                                End If
                                If Not IsNull(cf) Then
                                    x.colorFull = cf
                                End If
                            Next
                        Else
                            If Not IsNull(c) Then
                                Lampz.obj(idx).color = c
                            End If
                            If Not IsNull(cf) Then
                                Lampz.obj(idx).colorFull = cf
                            End If
                        End If
                        If Lampz.UseCallBack(idx) then Proc Lampz.name & idx,Lampz.Lvl(idx)*Lampz.Modulate(idx) 'Force Callbacks Proc
                    End If
                    Lampz.state(idx) = CInt(m_currentFrameState(frameStateKey).level) 'Lampz will handle redundant updates
                Else
                    Dim lm
                    If IsArray(Lights(idx)) Then
                        For Each x in Lights(idx)
                            If bUpdate Then
                                If Not IsNull(c) Then
                                    x.color = c
                                End If
                                If Not IsNull(cf) Then
                                    x.colorFull = cf
                                End If
                                If m_lightmaps.Exists(x.Name) Then
                                    For Each lm in m_lightmaps(x.Name)
                                        lm.Color = c
                                    Next
                                End If
                            End If
                            x.State = m_currentFrameState(frameStateKey).level/100
                        Next
                    Else
                        If bUpdate Then
                            If Not IsNull(c) Then
                                Lights(idx).color = c
                            End If
                            If Not IsNull(cf) Then
                                Lights(idx).colorFull = cf
                            End If
                            If m_lightmaps.Exists(Lights(idx).Name) Then
                                For Each lm in m_lightmaps(Lights(idx).Name)
                                    If Not IsNull(lm) Then
                                        lm.Color = c
                                    End If
                                Next
                            End If
                        End If
                        Lights(idx).State = m_currentFrameState(frameStateKey).level/100
                    End If
                End If





            Next
        End If
        m_currentFrameState.RemoveAll
        m_off.RemoveAll

    End Sub

    Private Function HexToInt(hex)
        HexToInt = CInt("&H" & hex)
    End Function

    Private Function HasKeys(o)
        Dim Success
        Success = False

        On Error Resume Next
            o.Keys()
            Success = (Err.Number = 0)
        On Error Goto 0
        HasKeys = Success
    End Function

    Private Sub RunLightSeq(seqRunner)

        Dim lcSeq: Set lcSeq = seqRunner.CurrentItem
        dim lsName, isSeqEnd
        If UBound(lcSeq.Sequence)<lcSeq.CurrentIdx Then
            isSeqEnd = True
        Else
            isSeqEnd = False
        End If

        dim lightInSeq
        For each lightInSeq in lcSeq.LightsInSeq

            If isSeqEnd Then



            'Needs a guard here for something, but i've forgotten.
            'I remember: Only reset the light if there isn't frame data for the light.
            'e.g. a previous seq has affected the light, we don't want to clear that here on this frame
                If m_lights.Exists(lightInSeq) = True AND NOT m_currentFrameState.Exists(lightInSeq) Then


                    If lcSeq.Name = "lSeqRgbRandomRed" AND lightInSeq = "l46" Then
      '               Debug.print("Reseting l46")
                    End If
                   AssignStateForFrame lightInSeq, (new FrameState)(0, Null, m_lights(lightInSeq).Idx)
                End If
            Else



                If m_currentFrameState.Exists(lightInSeq) Then


                    'already frame data for this light.
                    'replace with the last known state from this seq
                    If Not IsNull(lcSeq.LastLightState(lightInSeq)) Then
           '             If lcSeq.Name = "lSeqRgbRandomRed" AND lightInSeq = "l46" Then
           '                Debug.print("Assigning Previous State for l46")
            '            End If
            AssignStateForFrame lightInSeq, lcSeq.LastLightState(lightInSeq)
                    End If
                End If

            End If
        Next

        If isSeqEnd Then
            lcSeq.CurrentIdx = 0
            seqRunner.NextItem()
        End If

        If Not IsNull(seqRunner.CurrentItem) Then
            Dim framesRemaining, seq, color
            Set lcSeq = seqRunner.CurrentItem
            seq = lcSeq.Sequence


            Dim name
            Dim ls, x
            If IsArray(seq(lcSeq.CurrentIdx)) Then
                For x = 0 To UBound(seq(lcSeq.CurrentIdx))
                    lsName = Split(seq(lcSeq.CurrentIdx)(x),"|")
                    name = lsName(0)
                    If m_lights.Exists(name) Then
                        Set ls = m_lights(name)

            color = lcSeq.Color

                        If IsNull(color) Then
              color = ls.Color
                        End If

                        If Ubound(lsName) = 2 Then
              If lsName(2) = "FFFFFF" Then
                                AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                            Else
                                AssignStateForFrame name, (new FrameState)(lsName(1), Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), RGB(0,0,0)), ls.Idx)
                            End If
                        Else
                            AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                        End If
                        lcSeq.SetLastLightState name, m_currentFrameState(name)
                    End If
                Next
            Else
                lsName = Split(seq(lcSeq.CurrentIdx),"|")
                name = lsName(0)
                If m_lights.Exists(name) Then
                    Set ls = m_lights(name)

          color = lcSeq.Color
                    If IsNull(color) Then
                        color = ls.Color
                    End If
                    If Ubound(lsName) = 2 Then
                        If lsName(2) = "FFFFFF" Then
                            AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                        Else
                            AssignStateForFrame name, (new FrameState)(lsName(1), Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), RGB(0,0,0)), ls.Idx)
                        End If
                    Else
                        AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                    End If
                    lcSeq.SetLastLightState name, m_currentFrameState(name)
                End If
            End If

            framesRemaining = lcSeq.Update(m_frameTime)
            If framesRemaining < 0 Then
                lcSeq.ResetInterval()
                lcSeq.NextFrame()
            End If

        End If
    End Sub

End Class

Class FrameState
    Private m_level, m_colors, m_idx

    Public Property Get Level(): Level = m_level: End Property
    Public Property Let Level(input): m_level = input: End Property

    Public Property Get Colors(): Colors = m_colors: End Property
    Public Property Let Colors(input): m_colors = input: End Property

    Public Property Get Idx(): Idx = m_idx: End Property
    Public Property Let Idx(input): m_idx = input: End Property

    Public default function init(level, colors, idx)
    m_level = level
    m_colors = colors
    m_idx = idx

    Set Init = Me
    End Function

    Public Function ColorAt(idx)
        ColorAt = m_colors(idx)
    End Function
End Class

Class PulseState
    Private m_light, m_pulses, m_idx, m_interval, m_cnt

    Public Property Get Light(): Set Light = m_light: End Property
    Public Property Let Light(input): Set m_light = input: End Property

    Public Property Get Pulses(): Pulses = m_pulses: End Property
    Public Property Let Pulses(input): m_pulses = input: End Property

    Public Property Get Idx(): Idx = m_idx: End Property
    Public Property Let Idx(input): m_idx = input: End Property

    Public Property Get Interval(): Interval = m_interval: End Property
    Public Property Let Interval(input): m_interval = input: End Property

    Public Property Get Cnt(): Cnt = m_cnt: End Property
    Public Property Let Cnt(input): m_cnt = input: End Property

    Public default function init(light, pulses, idx, interval, cnt)
    Set m_light = light
    m_pulses = pulses
    'debug.Print(Join(Pulses))
    m_idx = idx
    m_interval = interval
    m_cnt = cnt

    Set Init = Me
    End Function

    Public Function PulseAt(idx)
        PulseAt = m_pulses(idx)
    End Function
End Class

Class LCItem

  Private m_Idx, m_State, m_blinkSeq, m_color, m_name, m_level, m_x, m_y

        Public Property Get Idx()
            Idx=m_Idx
        End Property

        Public Property Get Color()
            Color=m_color
        End Property

        Public Property Let Color(input)
            If IsNull(input) Then
        m_Color = Null
      Else
        If Not IsArray(input) Then
          input = Array(input, null)
        End If
        m_Color = input
      End If
      End Property

        Public Property Let Level(input)
            m_level = input
      End Property

        Public Property Get Level()
            Level=m_level
        End Property

        Public Property Get Name()
            Name=m_name
        End Property

        Public Property Get X()
            X=m_x
        End Property

        Public Property Get Y()
            Y=m_y
        End Property

        Public Sub Init(idx, intervalMs, color, name, x, y)
            m_Idx = idx
            If Not IsArray(color) Then
                m_color = Array(color, null)
            Else
                m_color = color
            End If
            m_name = name
            m_level = 100
            m_x = x
            m_y = y
      End Sub

End Class

Class LCSeq

  Private m_currentIdx, m_sequence, m_name, m_image, m_color, m_updateInterval, m_Frames, m_repeat, m_lightsInSeq, m_lastLightStates

    Public Property Get CurrentIdx()
        CurrentIdx=m_currentIdx
    End Property

    Public Property Let CurrentIdx(input)
    m_lastLightStates.RemoveAll()
        m_currentIdx = input
    End Property

    Public Property Get LightsInSeq()
        LightsInSeq=m_lightsInSeq.Keys()
    End Property

    Public Property Get Sequence()
        Sequence=m_sequence
    End Property

  Public Property Let Sequence(input)
    m_sequence = input
        dim item, light, lightItem
        for each item in input
            If IsArray(item) Then
                for each light in item
                    lightItem = Split(light,"|")
                    If Not m_lightsInSeq.Exists(lightItem(0)) Then
                        m_lightsInSeq.Add lightItem(0), True
                    End If
                next
            Else
                lightItem = Split(item,"|")
                If Not m_lightsInSeq.Exists(lightItem(0)) Then
                    m_lightsInSeq.Add lightItem(0), True
                End If
            End If
        next
  End Property

    Public Property Get LastLightState(light)
    If m_lastLightStates.Exists(light) Then
      dim c : Set c = m_lastLightStates(light)
      Set LastLightState = c
    Else
      LastLightState = Null
    End If
    End Property

    Public Property Let LastLightState(light, input)
        If m_lastLightStates.Exists(light) Then
            m_lastLightStates.Remove light
        End If
    If input.level > 0 Then
      m_lastLightStates.Add light, input
    End If
    End Property

    Public Sub SetLastLightState(light, input)
        If m_lastLightStates.Exists(light) Then
            m_lastLightStates.Remove light
        End If
        If input.level > 0 Then
                m_lastLightStates.Add light, input
        End If
    End Sub

    Public Property Get Color()
        Color=m_color
    End Property

  Public Property Let Color(input)
    If IsNull(input) Then
      m_Color = Null
    Else
      If Not IsArray(input) Then
        input = Array(input, null)
      End If
      m_Color = input
    End If
  End Property

    Public Property Get Name()
        Name=m_name
    End Property

  Public Property Let Name(input)
    m_name = input
  End Property

    Public Property Get UpdateInterval()
        UpdateInterval=m_updateInterval
    End Property

    Public Property Let UpdateInterval(input)
        m_updateInterval = input
        m_Frames = input
    End Property

    Public Property Get Repeat()
        Repeat=m_repeat
    End Property

    Public Property Let Repeat(input)
        m_repeat = input
    End Property

    Private Sub Class_Initialize()
        m_currentIdx = 0
        m_color = Array(Null, Null)
        m_updateInterval = 180
        m_repeat = False
        m_Frames = 180
        Set m_lightsInSeq = CreateObject("Scripting.Dictionary")
        Set m_lastLightStates = CreateObject("Scripting.Dictionary")
    End Sub

    Public Property Get Update(framesPassed)
        m_Frames = m_Frames - framesPassed
        Update = m_Frames
    End Property

    Public Sub NextFrame()
        m_currentIdx = m_currentIdx + 1
    End Sub

    Public Sub ResetInterval()

        m_Frames = m_updateInterval
        Exit Sub

        If Not IsNull(m_sequence) And UBound(m_sequence) > 1 Then

        'For i = 0 To totalSteps - 1
        '    currentStep = i
        '    duration = 20 ' Base duration of 20ms
            'Debug.print("TotalSteps: " & UBound(m_sequence)-1)
            Dim easeAmount : easeAmount = Round(m_currentIdx / UBound(m_sequence), 2) ' Normalize current step
            if easeAmount < 0 then
                easeAmount = 0
            elseif easeAmount > 1 then
                easeAmount = 1
            end if
            'Debug.print("Step: " & m_currentIdx)
            'Debug.print("Ease Amount: "& easeAmount)
            Dim newDuration : newDuration = 100 - Lerp(20, 80, EaseIn(easeAmount) )' Apply EaseInOut to duration
            'Debug.print("Duration: "& Round(newDuration))
            'Dim newDuration : newDuration = 100- Lerp(20, 80, Spike(easeAmount) )' Apply EaseInOut to duration

            m_frames = newDuration
        Else
            m_Frames = m_updateInterval
        End If
    End Sub

End Class

Class LCSeqRunner

  Private m_name, m_items,m_currentItemIdx

    Public Property Get Name()
        Name=m_name
    End Property

  Public Property Let Name(input)
    m_name = input
  End Property

    Public Property Get Items()
    Set Items = m_items
  End Property

    Public Property Get CurrentItem()
        Dim items: items = m_items.Items()
        If m_currentItemIdx > UBound(items) Then
            m_currentItemIdx = 0
        End If
        If UBound(items) = -1 Then
            CurrentItem  = Null
        Else
            Set CurrentItem = items(m_currentItemIdx)
        End If
    End Property

    Private Sub Class_Initialize()
        Set m_items = CreateObject("Scripting.Dictionary")
        m_currentItemIdx = 0
    End Sub

    Public Sub AddItem(item)
        If Not IsNull(item) Then
            If Not m_items.Exists(item.Name) Then
                m_items.Add item.Name, item
            End If
        End If
    End Sub

    Public Sub RemoveAll()
        Dim item
        For Each item in m_items.Keys()
            m_items(item).ResetInterval
            m_items(item).CurrentIdx = 0
            m_items.Remove item
        Next
    End Sub

    Public Sub RemoveItem(item)
        If Not IsNull(item) Then
            If m_items.Exists(item.Name) Then
                    item.ResetInterval
                    item.CurrentIdx = 0
                    m_items.Remove item.Name
            End If
        End If
    End Sub

    Public Sub NextItem()
        Dim items: items = m_items.Items
        If items(m_currentItemIdx).Repeat = False Then
            RemoveItem(items(m_currentItemIdx))
        Else
            m_currentItemIdx = m_currentItemIdx + 1
        End If

        If m_currentItemIdx > UBound(m_items.Items) Then
            m_currentItemIdx = 0
        End If
    End Sub

    Public Function HasSeq(name)
        If m_items.Exists(name) Then
            HasSeq = True
        Else
            HasSeq = False
        End If
    End Function

End Class


Function Lerp(startValue, endValue, amount)
    Lerp = startValue + (endValue - startValue) * amount
End Function

Function Flip(x)
    Flip = 1 - x
End Function

Function EaseIn(amount)
    EaseIn = amount * amount
End Function

Function EaseOut(amount)
    EaseOut = Flip(Sqr(Flip(amount)))
End Function

Function EaseInOut(amount)
    EaseInOut = Lerp(EaseIn(amount), EaseOut(amount), amount)
End Function

Function Spike(t)
    If t <= 0.5 Then
        Spike = EaseIn(t / 0.5)
    Else
        Spike = EaseIn(Flip(t)/0.5)
    End If
End Function

















'VR Room Code Below - Rawd *************************

Randomize
Dim TVPos: TVPos = 1
Dim Stuff
Dim BobAnimation: BobAnimation = 0
Dim BobSwaySpeed: BobSwaySpeed = 0.1
Dim VRPatrickSpeed: VRPatrickSpeed = 1
Dim GaryMove: GaryMove = 1
Dim StartButtonBubbleOn: StartButtonBubbleOn = False
Dim LaunchButtonBubbleOn: LaunchButtonBubbleOn = False
Dim LeftFlipperBubbleOn: LeftFlipperBubbleOn = False
Dim RightFlipperBubbleOn: RightFlipperBubbleOn = False

VRRoomInit

Sub VRRoomInit
If VRRoom > 0 then
  FlexOnFlasher = 1
  'Hide unwanted VLM Visuals objects...
  PincabRails_BM_Lit_Room.visible = false
  PincabRails_LM_Flashers_f141.visible = false
  PincabRails_LM_Flashers_f142.visible = false
' PincabRails_LM_GIString8_L108.visible = false
  SetBackglass
  If VR_Backglass = 0 Then
    BGDark.imageA = "spongebg3scrn-rev1"
    BGBright.imageA = "spongebg3scrn-rev1"
    BGBright.imageB = "spongebg3scrn-rev1"
  Else
    BGDark.imageA = "Backglass_ALT_1"
    BGBright.imageA = "Backglass_ALT_1"
    BGBright.imageB = "Backglass_ALT_1"
  End If

  If VRRoom = 1 then
    For each Stuff in VRMax: Stuff.visible = true: Next
    For each Stuff in VRCab: Stuff.visible = true: Next
    BobWalking.visible = True
    ClothesWalking.visible = true
    VRAnimationTimer.enabled = True
    TVTimer.enabled = True
    VRButtonTimer.enabled = True
    If Scratches = 1 then GlassImpurities.visible = True
  End If

  If VRRoom = 2 then
    For each Stuff in VRMin: Stuff.visible = true: Next
    For each Stuff in VRCab: Stuff.visible = true: Next
    If Scratches = 1 then GlassImpurities.visible = True
  End If

  If VRRoom = 3 then
    ' show just basic cab components (rails, lockbar and backbox/backglass
    PinCab_Backbox.visible = True
'   VRBackglassFlasher.visible = True
    VRDMD5.visible = True
    DMDHousing.visible = True
    PinCab_Rails.visible = True
  End If
End If
End Sub

Sub TVTimer_Timer
TVPos = TVPos +1
If TVPos = 6 then TVPos = 1
If TVPos = 1 then VRTV.ImageA = "TV1"
If TVPos = 2 then VRTV.ImageA = "TV2"
If TVPos = 3 then VRTV.ImageA = "TV3"
If TVPos = 4 then VRTV.ImageA = "TV4"
If TVPos = 5 then VRTV.ImageA = "TV5"
End Sub


Sub VRButtonTimer_timer  ' runs the VR Start button light and Plunger light wben needed.
if BallInLane = 1 then
  If VRLaunchButton.Blenddisablelighting = 0 then
    VRLaunchButton.Blenddisablelighting = 3
  Else
    VRLaunchButton.Blenddisablelighting = 0
  end If
Else
VRLaunchButton.Blenddisablelighting = 0
end if
if StartGame =0 Then
  If VRStartButton.Blenddisablelighting = 0 then
    VRStartButton.Blenddisablelighting = 3
  Else
    VRStartButton.Blenddisablelighting = 0
  end If
Else
VRStartButton.Blenddisablelighting = 0
end If
End Sub

Dim EndBOBAnim : EndBOBAnim = 0
Sub VRAnimationTimer_timer  'Main Animation timer for everything

If BobAnimation = 1 then
  For Each Stuff in VRBob: Stuff.z = Stuff.z + 1: next
  if BobWalking.z =>-2770 then BobAnimation = 2:ClothesWalking.image = "Eyes3"
end if

If BobAnimation = 2 then
  if VRSingleDoor1.roty > -90 then
    VRSingleDoor1.roty = VRSingleDoor1.roty - 1
    VRSingleDoorHousing1.roty = VRSingleDoorHousing1.roty -1
    if VRSingleDoor1.roty <= -90 then
      BobAnimation = 3
      ClothesWalking.image = "EyesClosed2"
      BobWalking.PlayAnimEndless(0.09)
      ClothesWalking.PlayAnimEndless(0.09)
    End if
  end If
end if

If BobAnimation = 3 then
For Each Stuff in VRBob: Stuff.y = Stuff.y + 15: next
For Each Stuff in VRBob: Stuff.z = Stuff.z + 1.5: next
For Each Stuff in VRBob: Stuff.x = Stuff.x - 5.2: next ' move forward/right

' Side to Side Sway..
For Each Stuff in VRBob: Stuff.Rotz = Stuff.Rotz + BobSwaySpeed: next
if BobWalking.Rotz =>1 then BobSwaySpeed = -0.1
if BobWalking.Rotz =<-1 then BobSwaySpeed = 0.1

  if BobWalking.y > -5500 then 'close the door behind him
    if VRSingleDoor1.roty < 0 then
      VRSingleDoor1.roty = VRSingleDoor1.roty + 1
      VRSingleDoorHousing1.roty = VRSingleDoorHousing1.roty +1
    end If
  end if

if BobWalking.y > -300 then BobAnimation = 4: 'stop animation here - Animation 4 will be him jumping to the stool. jump up and down?  arms raised?
End If

If BobAnimation = 4 then
'jump to stool..
if BobWalking.y < 400 then
  For Each Stuff in VRBob: Stuff.y = Stuff.y + 22: next
  For Each Stuff in VRBob: Stuff.z = Stuff.z + 16.5: next
end If

if BobWalking.y >400 and  BobWalking.y <1000  then
  For Each Stuff in VRBob: Stuff.y = Stuff.y + 24: next
  For Each Stuff in VRBob: Stuff.z = Stuff.z - 5.5: next
  For Each Stuff in VRBob: Stuff.roty = Stuff.roty + 3: next ' turn in the air?
end If

if BobWalking.y =>1000 Then
  BobAnimation = 5
  BobWalking.StopAnim()
  ClothesWalking.StopAnim()  'stop for now
  BobJumping.PlayAnimEndless(0.09)
  ClothesJumping.PlayAnimEndless(0.09)
  ClothesWalking.image = "Eyes3"
end If

end If

if BobAnimation = 5 then
  BobWalking.visible = false
  ClothesWalking.visible = false
  BobJumping.visible = true
  ClothesJumping.visible = true
  'He stays like this until Multiball end...
end if

If BobAnimation = 6 then
  'jump back and rotate and land...
  For Each Stuff in VRBob: Stuff.y = Stuff.y - 15: next
  For Each Stuff in VRBob: Stuff.z = Stuff.z + 6.5: next
  For Each Stuff in VRBob: Stuff.x = Stuff.x + 16: next
  For Each Stuff in VRBob: Stuff.Roty = Stuff.Roty - 3: next
  if BobWalking.y =<600 Then BobAnimation = 7  ' Hes jumped back, now stop him.. testing for landing position
End If

If BobAnimation = 7 then  ' fall back to ground.
  For Each Stuff in VRBob: Stuff.y = Stuff.y - 5: next
  For Each Stuff in VRBob: Stuff.z = Stuff.z - 8.5: next
  For Each Stuff in VRBob: Stuff.x = Stuff.x + 9: next
    if BobWalking.y =<300 Then BobAnimation = 8 :ClothesWalking.image = "EyesClosed2" ' Hes jumped back, now stop him.. testing for landing position
End If

If BobAnimation = 8 then
  'move him towards the back door now...
  For Each Stuff in VRBob: Stuff.y = Stuff.y + 15: next
  For Each Stuff in VRBob: Stuff.z = Stuff.z + 1.5: next
  For Each Stuff in VRBob: Stuff.x = Stuff.x + 5: next

  ' Side to Side Sway..
  For Each Stuff in VRBob: Stuff.Rotz = Stuff.Rotz + BobSwaySpeed: next
  if BobWalking.Rotz =>1 then BobSwaySpeed = -0.1
  if BobWalking.Rotz =<-1 then BobSwaySpeed = 0.1

  if BobWalking.y > 7000 and BobWalking.y <9000  then 'Open back door
    if VRSingleDoor3.roty < 90 then
      VRSingleDoor3.roty = VRSingleDoor3.roty + 1
      VRSingleDoorHousing3.roty = VRSingleDoorHousing3.roty +1
    end If
  end if

  if BobWalking.y > 9800 then 'Close back door
    if VRSingleDoor3.roty > 0 then
      VRSingleDoor3.roty = VRSingleDoor3.roty - 1
      VRSingleDoorHousing3.roty = VRSingleDoorHousing3.roty -1
    end If
  end if

'move Bob back to starting position
  if BobWalking.y > 12800 then
  ' Debug.print "bobanim  over"
    BobAnimation = 0
    BobWalking.StopAnim()
    ClothesWalking.StopAnim()
    For Each Stuff in VRBob: Stuff.y = -10599.35: Next
    For Each Stuff in VRBob: Stuff.x = 5533.112: Next
    For Each Stuff in VRBob: Stuff.z = -3500: Next
    For Each Stuff in VRBob: Stuff.RotY = 0: Next
  end if
End If ' End Bob Animation sequence

' gary
If GaryMove = 1 then
  For each Stuff in VRGary: Stuff.y = Stuff.y + 1: Next
  For each Stuff in VRGary: Stuff.z = Stuff.z + .105: Next
  If VRGaryEyes.y =>5380 then GaryMove = 2
End if

If GaryMove = 2 then
  For each Stuff in VRGary: Stuff.Roty = Stuff.Roty + .01: Next  'slow movement at first.
  If VRGaryEyes.roty => 95 then GaryMove = 3
end If

If GaryMove = 3 then
  For each Stuff in VRGary: Stuff.Roty = Stuff.Roty + .1: Next  ' turn faster now.
  If VRGaryEyes.roty => 270 then GaryMove = 4
end if

If GaryMove = 4 then
  For each Stuff in VRGary: Stuff.y = Stuff.y - 1: Next
  For each Stuff in VRGary: Stuff.z = Stuff.z - .105: Next
  if VRGaryEyes.y =< -6702 then   GaryMove = 5 :For each Stuff in VRGary: Stuff.Roty = -90:next  'changing value from 270 to -90..  sheesh..
end if

If GaryMove = 5 then
  For each Stuff in VRGary: Stuff.Roty = Stuff.Roty + .1: Next
  If VRGaryEyes.roty => 90 then
    For each Stuff in VRGary: Stuff.Roty = 90:next ' reset angle value
    GaryMove = 1
  end If
End If

  'Move WindowGuys on this timer also..
  For each Stuff in VRWindowGuys: Stuff.Z = Stuff.Z + VRPatrickSpeed:next  'VRPatrick.z = VRPatrick.z + VRPatrickSpeed
  if VRPatrick.z > 1000 then VRPatrickSpeed = -1

'Patrick
  if VRPatrick.z < -1200 then
    VRPatrickSpeed = 1
    If VRPatrick.Visible = true then
      VRPatrick.Visible = False
      VRSquid.Visible = true
      Else
      VRPatrick.Visible = true
      VRSquid.Visible = false

    End If
  end If

'Bubbles in Window constantly run...
VRBubble1.height = VRBubble1.height + 4
VRBubble2.height = VRBubble2.height + 5
VRBubble3.height = VRBubble3.height + 7
If VRBubble1.height > 3000 then VRBubble1.height = 300: VRBubble1.y = -300 + rnd(1)*1100   ' -500y to +500Y
If VRBubble2.height > 4000 then VRBubble2.height = 300: VRBubble2.y = -300 + rnd(1)*1100
If VRBubble3.height > 4500 then VRBubble3.height = 200: VRBubble3.y = -300 + rnd(1)*1100

'Cabinet button Bubbles...
if StartButtonBubbleOn = true then VRBubbleStartButton.z = VRBubbleStartButton.z + 9
if LaunchButtonBubbleOn = true then VRBubbleLaunchButton.z = VRBubbleLaunchButton.z + 9
if LeftFlipperBubbleOn = true then VRBubbleLeftFlipper.z = VRBubbleLeftFlipper.z + 9
if RightFlipperBubbleOn = true then VRBubbleRightFlipper.z = VRBubbleRightFlipper.z + 9

If VRBubbleStartButton.z =>6500 then StartButtonBubbleOn = false: VRBubbleStartButton.visible = false : VRBubbleStartButton.z = -580
If VRBubbleLaunchButton.z =>6500 then LaunchButtonBubbleOn = false: VRBubbleLaunchButton.visible = false : VRBubbleLaunchButton.z = -580
If VRBubbleLeftFlipper.z =>6500 then LeftFlipperBubbleOn = false: VRBubbleLeftFlipper.visible = false : VRBubbleLeftFlipper.z = -480
If VRBubbleRightFlipper.z =>6500 then RightFlipperBubbleOn = false: VRBubbleRightFlipper.visible = false : VRBubbleRightFlipper.z = -480
End Sub


'**********************************************************
'*******Set Up VR Backglass and Backglass Flashers  *******
'**********************************************************

Sub SetBackglass()

BGDark.visible = True
BGBright.visible = True

Dim VRobj

  For Each VRobj In VRBackglass
    VRobj.x = VRobj.x
    VRobj.height = - VRobj.y + 1360
    VRobj.y = 60 'adjusts the distance from the backglass towards the user
    VRobj.rotx = -87
  Next

End Sub

' ******************
' **** Flashers ****
' ******************

'Spongebob
dim VRBGFLSplvl
sub VRBGFLSponge
  If VRRoom > 0 Then
    VRBGFLSplvl = 1
    VRBGFLSP_1_timer
  End If
end sub

sub VRBGFLSP_1_timer()
  if Not VRBGFLSP_1.TimerEnabled then
    If VR_Backglass = 1 Then
      VRBGFLSP_1.visible = true
      VRBGFLSP_2.visible = true
      VRBGFLSP_3.visible = true
      VRBGFLSP_4.visible = true
      VRBGFLSP_5.visible = true
    Else
      VRBGFLSP_001.visible = true
      VRBGFLSP_002.visible = true
      VRBGFLSP_003.visible = true
      VRBGFLSP_004.visible = true
    End If
    VRBGFLSP_1.TimerEnabled = true
  end if
  If VR_Backglass = 1 Then
    VRBGFLSP_1.opacity = 70 * VRBGFLSplvl^1.5
    VRBGFLSP_2.opacity = 70 * VRBGFLSplvl^1.5
    VRBGFLSP_3.opacity = 70 * VRBGFLSplvl^1.5
    VRBGFLSP_4.opacity = 125 * VRBGFLSplvl^2
    VRBGFLSP_5.opacity = 125 * VRBGFLSplvl^2.5
  Else
    VRBGFLSP_001.opacity = 35 * VRBGFLSplvl^1.5
    VRBGFLSP_002.opacity = 35 * VRBGFLSplvl^1.5
    VRBGFLSP_003.opacity = 35 * VRBGFLSplvl^1.5
    VRBGFLSP_004.opacity = 165 * VRBGFLSplvl^2
  End If
  VRBGFLSplvl = 0.80 * VRBGFLSplvl - 0.01
  if VRBGFLSplvl < 0 then VRBGFLSplvl = 0

  if VRBGFLSplvl =< 0 Then
    If VR_Backglass = 1 Then
      VRBGFLSP_1.visible = false
      VRBGFLSP_2.visible = false
      VRBGFLSP_3.visible = false
      VRBGFLSP_4.visible = false
      VRBGFLSP_5.visible = false
    Else
      VRBGFLSP_001.visible = false
      VRBGFLSP_002.visible = false
      VRBGFLSP_003.visible = false
      VRBGFLSP_004.visible = false
    End If
    VRBGFLSP_1.TimerEnabled = false
  end if
end sub

'Patrick
dim VRBGFLPAlvl
sub VRBGFLPat
  If VRRoom > 0 Then
    VRBGFLPAlvl = 1
    VRBGFLPA_1_timer
  End If
end sub

sub VRBGFLPA_1_timer()
  if Not VRBGFLPA_1.TimerEnabled then
    If VR_Backglass = 1 Then
      VRBGFLPA_1.visible = true
      VRBGFLPA_2.visible = true
      VRBGFLPA_3.visible = true
      VRBGFLPA_4.visible = true
      VRBGFLPA_5.visible = true
    Else
      VRBGFLPA_001.visible = true
      VRBGFLPA_002.visible = true
      VRBGFLPA_003.visible = true
      VRBGFLPA_004.visible = true
    End If
    VRBGFLPA_1.TimerEnabled = true
  end if
  If VR_Backglass = 1 Then
    VRBGFLPA_1.opacity = 70 * VRBGFLPAlvl^1.5
    VRBGFLPA_2.opacity = 70 * VRBGFLPAlvl^1.5
    VRBGFLPA_3.opacity = 70 * VRBGFLPAlvl^1.5
    VRBGFLPA_4.opacity = 125 * VRBGFLPAlvl^2
    VRBGFLPA_5.opacity = 125 * VRBGFLPAlvl^2.5
  Else
    VRBGFLPA_001.opacity = 30 * VRBGFLPAlvl^1.5
    VRBGFLPA_002.opacity = 30 * VRBGFLPAlvl^1.5
    VRBGFLPA_003.opacity = 30 * VRBGFLPAlvl^1.5
    VRBGFLPA_004.opacity = 75 * VRBGFLPAlvl^2
  End If
  VRBGFLPAlvl = 0.80 * VRBGFLPAlvl - 0.01
  if VRBGFLPAlvl < 0 then VRBGFLPAlvl = 0
  if VRBGFLPAlvl =< 0 Then
    If VR_Backglass = 1 Then
      VRBGFLPA_1.visible = false
      VRBGFLPA_2.visible = false
      VRBGFLPA_3.visible = false
      VRBGFLPA_4.visible = false
      VRBGFLPA_5.visible = false
    Else
      VRBGFLPA_001.visible = false
      VRBGFLPA_002.visible = false
      VRBGFLPA_003.visible = false
      VRBGFLPA_004.visible = false
    End If
    VRBGFLPA_1.TimerEnabled = false
  end if
end sub

'Bumper 1
dim VRBGFLBump1lvl
sub VRBGFLBump1
  If VRRoom > 0 Then
    VRBGFLBump1lvl = 1
    VRBGFLBump1_1_timer
  End If
end sub

sub VRBGFLBump1_1_timer()
  if Not VRBGFLBump1_1.TimerEnabled then
    If VR_Backglass = 1 Then
      VRBGFLBump1_1.visible = true
      VRBGFLBump1_2.visible = true
      VRBGFLBump1_3.visible = true
      VRBGFLBump1_4.visible = true
    Else
      VRBGFLBump1_001.visible = true
      VRBGFLBump1_002.visible = true
      VRBGFLBump1_003.visible = true
      VRBGFLBump1_004.visible = true
    End If
    VRBGFLBump1_1.TimerEnabled = true
  end if
  If VR_Backglass = 1 Then
    VRBGFLBump1_1.opacity = 50 * VRBGFLBump1lvl^1.5
    VRBGFLBump1_2.opacity = 50 * VRBGFLBump1lvl^1.5
    VRBGFLBump1_3.opacity = 50 * VRBGFLBump1lvl^1.5
    VRBGFLBump1_4.opacity = 100 * VRBGFLBump1lvl^2
  Else
    VRBGFLBump1_001.opacity = 40 * VRBGFLBump1lvl^1.5
    VRBGFLBump1_002.opacity = 40 * VRBGFLBump1lvl^1.5
    VRBGFLBump1_003.opacity = 40 * VRBGFLBump1lvl^1.5
    VRBGFLBump1_004.opacity = 120 * VRBGFLBump1lvl^2
  End If
  VRBGFLBump1lvl = 0.80 * VRBGFLBump1lvl - 0.01
  if VRBGFLBump1lvl < 0 then VRBGFLBump1lvl = 0
  if VRBGFLBump1lvl =< 0 Then
    If VR_Backglass = 1 Then
      VRBGFLBump1_1.visible = false
      VRBGFLBump1_2.visible = false
      VRBGFLBump1_3.visible = false
      VRBGFLBump1_4.visible = false
    Else
      VRBGFLBump1_001.visible = false
      VRBGFLBump1_002.visible = false
      VRBGFLBump1_003.visible = false
      VRBGFLBump1_004.visible = false
    End If
    VRBGFLBump1_1.TimerEnabled = false
  end if
end sub

'Bumper 2
dim VRBGFLBump2lvl
sub VRBGFLBump2
  If VRRoom > 0 Then
    VRBGFLBump2lvl = 1
    VRBGFLBump2_1_timer
  End If
end sub

sub VRBGFLBump2_1_timer()
  if Not VRBGFLBump2_1.TimerEnabled then
    If VR_Backglass = 1 Then
      VRBGFLBump2_1.visible = true
      VRBGFLBump2_2.visible = true
      VRBGFLBump2_3.visible = true
      VRBGFLBump2_4.visible = true
    Else
      VRBGFLBump2_001.visible = true
      VRBGFLBump2_002.visible = true
      VRBGFLBump2_003.visible = true
      VRBGFLBump2_004.visible = true
    End If
    VRBGFLBump2_1.TimerEnabled = true
  end if
  If VR_Backglass = 1 Then
    VRBGFLBump2_1.opacity = 50 * VRBGFLBump2lvl^1.5
    VRBGFLBump2_2.opacity = 50 * VRBGFLBump2lvl^1.5
    VRBGFLBump2_3.opacity = 50 * VRBGFLBump2lvl^1.5
    VRBGFLBump2_4.opacity = 100 * VRBGFLBump2lvl^2
  Else
    VRBGFLBump2_001.opacity = 40 * VRBGFLBump2lvl^1.5
    VRBGFLBump2_002.opacity = 40 * VRBGFLBump2lvl^1.5
    VRBGFLBump2_003.opacity = 40 * VRBGFLBump2lvl^1.5
    VRBGFLBump2_004.opacity = 120 * VRBGFLBump2lvl^2
  End If
  VRBGFLBump2lvl = 0.80 * VRBGFLBump2lvl - 0.01
  if VRBGFLBump2lvl < 0 then VRBGFLBump2lvl = 0
  if VRBGFLBump2lvl =< 0 Then
    If VR_Backglass = 1 Then
      VRBGFLBump2_1.visible = false
      VRBGFLBump2_2.visible = false
      VRBGFLBump2_3.visible = false
      VRBGFLBump2_4.visible = false
    Else
      VRBGFLBump2_001.visible = false
      VRBGFLBump2_002.visible = false
      VRBGFLBump2_003.visible = false
      VRBGFLBump2_004.visible = false
    End If
    VRBGFLBump2_1.TimerEnabled = false
  end if
end sub

' End VR code *************************************




'******************************************************
'   BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1.0         'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)
Dim GIfalloff : GIfalloff   = 250
Const PLOffset = 0.5
Dim PLGain: PLGain = (1-PLOffset)/(1260-2100)

Sub UpdateBallBrightness
  Dim BOT, s, b_base, b_r, b_g, b_b, d_w, l, LSd, gilvlavg
  BOT = getballs

  gilvlavg = Playfield_LM_GIString1_gi005.Opacity / 100
  if gilvlavg > 1 then gilvlavg = 1
  b_base = 70 * BallBrightness + 100 * gilvlavg  + 55 * RoomBrightness/100 + 30

  For s = 0 To UBound(BOT)
    ' Handle plunger lane
    If InRect(BOT(s).x,BOT(s).y,870,2100,870,1260,930,1260,930,2100) Then
      d_w = d_w*(PLOffset+PLGain*(BOT(s).y-2100))
    End If

'   ' GI light distance
'   d_w = GIfalloff
'   For Each l in DynamicSources
'     LSd = Distance(BOT(s).x, BOT(s).y, l.x, l.y) 'Calculating the Linear distance to the Source
'     If LSd < d_w then d_w = LSd
'   Next
'   d_w = b_base + 70 * (1 - d_w / GIfalloff) * gilvlavg

    ' Handle z direction
    d_w = b_base*(1 - (BOT(s).z-25)/500)
    If d_w < 30 Then d_w = 30

    ' Assign color
    b_r = Int(d_w)
    b_g = Int(d_w)
    b_b = Int(d_w)
    If b_r > 255 Then b_r = 255
    If b_g > 255 Then b_g = 255
    If b_b > 255 Then b_b = 255
    BOT(s).color = b_r + (b_g * 256) + (b_b * 256 * 256)
    'debug.print "--- ball.color level="&b_r
  Next
End Sub



'****************************
' Room Brightness
'****************************

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","_noXtraShading","_noXtraShadingVR","_noXtraShadingVRDark", _
  "Metal","Metal_Blue_Powdercoat1","Metal_Black_Powdercoat","Metal Dark","Metal Black","Rubber Black", _
  "Plastic with an image","Plastic with an imageVR","Plastic with an image opacity ","Plastic Grey")
Dim SavedMtlColorArray:     SavedMtlColorArray     = Array(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)


Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0

  ' Lighting level
  Dim v: v=(lvl * 220 + 35)/255

  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    ModulateMaterialBaseColor RoomBrightnessMtlArray(i), i, v
  Next

End Sub

SaveMtlColors
Sub SaveMtlColors
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

'UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
'               OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
'               float elasticity, float elasticityFalloff, float friction, float scatterAngle)


'***********  big qrcode

'flux: InitFlexScorbitClaimDMD - to show qr
'flux: CloseFlexScorbitClaimDMD - close it

Dim QRClaim_Time : QRClaim_Time = 0
Dim inQRClaimDMD : inQRClaimDMD = False
Dim QRFont1, QRFont2
Sub StartQRclaim
  Dim a, scene

  Set QRFont1 = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
  Set QRFont2 = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", rgb(100,100,100), vbWhite, 0)


  Set scene = FlexDMD.NewGroup("ClaimQR")
  scene.AddActor FlexDMD.NewImage("black", "black.png")


  Set a = FlexDMD.NewLabel("TopLine", QRFont2, "OTHER MAGNA FOR OPTION MENU")   : a.SetAlignedPosition 64,  7, FlexDMD_Align_Center       : scene.AddActor a
  If ScorbitActive <> 1 or isnull(scorbit) then
    Set a = FlexDMD.NewLabel("MidLine", QRFont1, "NEED PAIRING IN OPTIONS")   : a.SetAlignedPosition 64, 16, FlexDMD_Align_Center       : scene.AddActor a
  Elseif CurrentBall <> 3 or StartGame = 0 Then
    Set a = FlexDMD.NewLabel("MidLine", QRFont1, "QRCLAIM ONLY ON BALL 1")    : a.SetAlignedPosition 64, 16, FlexDMD_Align_Center       : scene.AddActor a
  Else
    Set a = FlexDMD.NewLabel("MidLine", QRFont1, "HOLD FOR QRCLAIM 0.0/4 SEC")    : a.SetAlignedPosition 64, 16, FlexDMD_Align_Center       : scene.AddActor a
  End If
  Set a = FlexDMD.NewLabel("BotLine", QRFont2, "RELEASE TO EXIT")         : a.SetAlignedPosition 64, 25, FlexDMD_Align_Center       : scene.AddActor a

  FlexDMD.LockRenderThread
  FlexDMD.Stage.AddActor scene
  FlexDMD.UnlockRenderThread
  inQRClaimDMD = True
  QRClaim_Time = 1
End Sub

Sub StopQRClaim


  QRClaim_Time = 0
  inQRClaimDMD = False
  FlexDMD.LockRenderThread
  FlexDMD.stage.GetImage("black").remove
  FlexDMD.stage.GetLabel("TopLine").remove
  FlexDMD.stage.GetLabel("MidLine").remove
  FlexDMD.stage.GetLabel("BotLine").remove
  FlexDMD.UnlockRenderThread
  CloseFlexScorbitClaimDMD
End Sub




'***********************************************************************
'* TABLE OPTIONS *******************************************************
'***********************************************************************

Dim BumperLightIntensity : BumperLightIntensity = 100 'Level of bumper light intensity (0 to 100), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 3
Dim ScorbitActive : ScorbitActive = 0
Dim OutPostMod : OutPostMod = 1   'Difficulty : 0 = Easy, 1 = Medium, 2 = Hard
Dim Easypost : Easypost = 0     'Extra Left Post : 0 = Off, 1 = On
Dim StagedFlipperMod : StagedFlipperMod = 0
'Dim MechVolume : MechVolume = 0.8      'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
'Dim BallRollVolume : BallRollVolume = 0.5  'Level of ball rolling volume. Value between 0 and 1
'Dim RampRollVolume : RampRollVolume = 0.5  'Level of ramp rolling volume. Value between 0 and 1

'Dim Cabinetmode  : Cabinetmode = 0     '0 - Siderails On, 1 - Siderails Off

'Dim DynamicBallShadowsOn : DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
'Dim AmbientBallShadowOn : AmbientBallShadowOn = 1    '0 = Static shadow under ball ("flasher" image, like JP's)
                                    '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
                                    '2 = flasher image shadow, but it moves like ninuzzu's

'Dim VRRoomChoice : VRRoomChoice = 0        '0 - Minimal Room

Dim UseCredits

' Base options
Const Opt_Info_1 = 0
Const Opt_Light = 1
Const Opt_LUT = 2

Const Opt_Scorbit = 3

Const Opt_RomVolume = 4
Const Opt_MechVolume = 5
Const Opt_Volume_Ramp = 6
Const Opt_Volume_Ball = 7
Const Opt_Staged_Flipper = 8
Const Opt_FreePlay = 9

'Const Opt_DynBallShadow = 8
'Const Opt_AmbientBallShadow = 9
Const Opt_Info_2 = 10
Const Opt_Reset = 11

Const NOptions = 12


Dim OptionDMD: Set OptionDMD = Nothing
Dim bOptionsMagna, bInOptions : bOptionsMagna = False
Dim OptPos, OptSelected, OptN, OptTop, OptBot, OptSel
Dim OptFontHi, OptFontLo, OptFontArrow
Dim ResetOption : Resetoption = "NO"

Sub Options_Open
  bOptionsMagna = False
  PlaySound "Sfx_Timer",0,RomSoundVolume


' On Error Resume Next
' Set OptionDMD = CreateObject("FlexDMD.FlexDMD")
' On Error Goto 0
' If OptionDMD is Nothing Then
'   Debug.Print "FlexDMD is not installed"
'   Debug.Print "Option UI can not be opened"
'   MsgBox "You need to install FlexDMD to access table options"
'   Exit Sub
' End If
' Debug.Print "Option UI opened"
' If ShowDT Then OptionDMDFlasher.RotX = -(Table1.Inclination + Table1.Layback)


  bInOptions = True
  Dampenmusic2 = 0.2

  OptPos = Opt_Info_1
  OptSelected = False
' OptionDMD.Show = False
' OptionDMD.RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
' OptionDMD.Width = 128
' OptionDMD.Height = 32
' OptionDMD.Clear = True
' OptionDMD.Run = True

  Dim a, scene, font
  Set scene = FlexDMD.NewGroup("Options")

  scene.AddActor FlexDMD.NewImage("blackbg", "black.png")

  Set OptFontHi = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
  Set OptFontLo = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(100, 100, 100), RGB(100, 100, 100), 0)
  Set OptFontArrow = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(150, 150, 0), RGB(150, 150, 0), 0)

  Set OptSel = FlexDMD.NewGroup("Sel")
  Set a = FlexDMD.NewLabel(">", OptFontArrow, ">>>                <<<")
  a.SetAlignedPosition 1, 16, FlexDMD_Align_Left
  OptSel.AddActor a
' Set a = FlexDMD.NewLabel("<", OptFontLo, "<<<")
' a.SetAlignedPosition 127, 16, FlexDMD_Align_Right
' OptSel.AddActor a
  scene.AddActor OptSel
  OptSel.SetBounds 0, 0, 128, 32
  OptSel.Visible = False

  Set a = FlexDMD.NewLabel("Info1", OptFontLo, "MAGNA EXIT/ENTER")
  a.SetAlignedPosition 1, 32, FlexDMD_Align_BottomLeft
  scene.AddActor a
  Set a = FlexDMD.NewLabel("Info2", OptFontLo, "FLIPPER SELECT")
  a.SetAlignedPosition 127, 32, FlexDMD_Align_BottomRight
  scene.AddActor a
  Set OptN = FlexDMD.NewLabel("Pos", OptFontLo, "OPTION")
  Set OptTop = FlexDMD.NewLabel("Top", OptFontLo, "LINE 1")
  Set OptBot = FlexDMD.NewLabel("Bottom", OptFontLo, "LINE 2")
  scene.AddActor OptN
  scene.AddActor OptTop
  scene.AddActor OptBot
  Options_OnOptChg
  FlexDMD.LockRenderThread
  FlexDMD.Stage.AddActor scene
  FlexDMD.UnlockRenderThread
' OptionDMDFlasher.Visible = True
  FlexDMD.Stage.GetLabel("Pos").font = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(255, 200, 0), rgb(0,0,0),0)
  FlexDMD.Stage.GetLabel("Pos").SetAlignedPosition 1, 1, FlexDMD_Align_TopLeft
End Sub


Dim OptionKeysEnabled : OptionKeysEnabled = False
Dim OptionBlink : OptionBlink = 0
Dim SelectAnim : SelectAnim = 0
Sub Options_UpdateDMD

  If OptionBlink > 0 Then

    OptionKeysEnabled = False

    If frame mod 4 = 1 Then
      OptionBlink = OptionBlink - 1
      Select Case OptionBlink
        Case 9,7,5,3,1 :    'Off
          OptTop.Visible = False
          OptBot.Visible = False
        case 8,6,4,2,0 :   ' on
          OptTop.Visible = True
          OptBot.Visible = True
      End Select

    End If
  Else
    OptionKeysEnabled = True
  End If

  If frame mod 2 = 1 Then
    Select Case SelectAnim
      Case 0 : SelectAnim = 1 : FlexDMD.Stage.GetLabel(">").text = "   >                        <   "
      Case 1 : SelectAnim = 2 : FlexDMD.Stage.GetLabel(">").text = ">                              <"
      Case 2 : SelectAnim = 3 : FlexDMD.Stage.GetLabel(">").text = " >                            < "
      Case 3 : SelectAnim = 0 : FlexDMD.Stage.GetLabel(">").text = "  >                          <  "
    End Select
  End If
  If frame mod 4 = 1 Then
    If Rightmagnablink = 1 Then
      FlexDMD.stage.getLabel("Info1").Font = OptFontHi
      FlexDMD.stage.getLabel("Info1").text = "           ENTER"
      Rightmagnablink = 0
    Elseif Leftmagnablink = 1 Then
      FlexDMD.stage.getLabel("Info1").Font = OptFontHi
      FlexDMD.stage.getLabel("Info1").text = "      EXIT      "
      Leftmagnablink = 0
    Else
      FlexDMD.stage.getLabel("Info1").Font = OptFontLo
      FlexDMD.stage.getLabel("Info1").text = "MAGNA EXIT/ENTER"
    End If
    If Flipperoption = 1 Then
      Flipperoption = 0
      FlexDMD.stage.getLabel("Info2").Font = OptFontHi
    Else
      FlexDMD.stage.getLabel("Info2").Font = OptFontLo
    End If
  End If


' Set a = FlexDMD.NewLabel("Info1", OptFontLo, "MAGNA EXIT/ENTER")
' a.SetAlignedPosition 1, 32, FlexDMD_Align_BottomLeft
' scene.AddActor a
' Set a = FlexDMD.NewLabel("Info2", OptFontLo, "FLIPPER SELECT")


' If OptionDMD is Nothing Then Exit Sub
' Dim DMDp: DMDp = OptionDMD.DmdPixels
' If Not IsEmpty(DMDp) Then
'   OptionDMDFlasher.DMDWidth = OptionDMD.Width
'   OptionDMDFlasher.DMDHeight = OptionDMD.Height
'   OptionDMDFlasher.DMDPixels = DMDp
' End If
End Sub

Sub Options_Close
  bInOptions = False
  Dampenmusic2 = 1
  PlaySound "Sfx_ModeReady",0,RomSoundVolume

  DMDScorebit.Visible = False
  DMDScorebit.TimerEnabled = False
  If Not IsNull(Scorbit) Then
    If Scorbit.bNeedsPairing = True Then
      Scorbit = Null
      tmrScorbit.Enabled=False
      ScorbitActive = 0
      If Not IsNull(FlexDMDScorbit) Then
        FlexDMDScorbit.Show = False
        FlexDMDScorbit.Run = False
        FlexDMDScorbit = NULL
      End If
      CloseFlexScorbitClaimDMD()
    End If
  End If

' OptionDMDFlasher.Visible = False
' If OptionDMD is Nothing Then Exit Sub
' OptionDMD.Run = False
' Set OptionDMD = Nothing

' fixing remove all option things here
      FlexDMD.stage.GetImage("blackbg").remove
      FlexDMD.stage.GetLabel(">").remove
      FlexDMD.stage.GetLabel("Info1").remove
      FlexDMD.stage.GetLabel("Info2").remove
      FlexDMD.stage.GetLabel("Pos").remove
      FlexDMD.stage.GetLabel("Top").remove
      FlexDMD.stage.GetLabel("Bottom").remove

  Options_Save

End Sub

Function Options_OnOffText(opt)
  If opt Then
    Options_OnOffText = "ON"
  Else
    Options_OnOffText = "OFF"
  End If
End Function

Sub Options_OnOptChg
  If FlexDMD is Nothing Then Exit Sub
' FlexDMD.LockRenderThread
' If RenderingMode <> 2 Then
'   If OptPos < Opt_VRRoomChoice Then
'     OptN.Text = (OptPos+1) & "/" & (NOptions - 4)
'   Else
'     OptN.Text = (OptPos+1 - 4) & "/" & (NOptions - 4)
'   End If
' Else
'   OptN.Text = (OptPos+1) & "/" & NOptions
' End If
    ' Highlight option or current value based on whether it is selected or not

  If Not OptPos=3 And Not IsNull(FlexDMDScorbit) Then
    FlexDMDScorbit.Show = False
    FlexDMDScorbit.Run = False
    FlexDMDScorbit = NULL
    DMDScorebit.Visible = False
    DMDScorebit.TimerEnabled = False
  End If

  If OptSelected Then
    OptTop.Font = OptFontLo
    OptBot.Font = OptFontHi
    OptSel.Visible = True
  Else
    OptTop.Font = OptFontHi
    OptBot.Font = OptFontLo
    OptSel.Visible = False
  End If

    ' Display option value based on which option (pos) we are currently at
  If OptPos = Opt_Light Then
      OptTop.Text = "LIGHT LEVEL"
    If VRRoom > 0 Then
    Else
      OptTop.Text = "LIGHT LEVEL (Req Restart)"
    End If
    OptBot.Text = "LEVEL " & RoomBrightness
  ElseIf OptPos = Opt_LUT Then
    OptTop.Text = "COLOR SATURATION"
'   OptBot.Text = "LUT " & CInt(ColorLUT)
    if ColorLUT = 1 Then OptBot.text = "DISABLED"
    if ColorLUT = 2 Then OptBot.text = "DESATURATED -10%"
    if ColorLUT = 3 Then OptBot.text = "DESATURATED -20%"
    if ColorLUT = 4 Then OptBot.text = "DESATURATED -30%"
    if ColorLUT = 5 Then OptBot.text = "DESATURATED -40%"
    if ColorLUT = 6 Then OptBot.text = "DESATURATED -50%"
    if ColorLUT = 7 Then OptBot.text = "DESATURATED -60%"
    if ColorLUT = 8 Then OptBot.text = "DESATURATED -70%"
    if ColorLUT = 9 Then OptBot.text = "DESATURATED -80%"
    if ColorLUT = 10 Then OptBot.text = "DESATURATED -90%"
    if ColorLUT = 11 Then OptBot.text = "BLACK'N WHITE"
  ElseIf OptPos = Opt_Scorbit Then
    OptTop.Text = "SCORBIT"
    if ScorbitActive = 0 Then OptBot.text = "OFF"
    if ScorbitActive = 1 Then OptBot.text = "ACTIVE" : InitFlexScorbitDMD
    SaveValue cGameName, "SCORBIT", ScorbitActive
  ElseIf OptPos = Opt_RomVolume Then
        OptTop.Text = "ROM VOLUME"
        OptBot.Text = "LEVEL " & CInt(RomSoundVolume * 100)
    ElseIf OptPos = Opt_MechVolume Then
    OptTop.Text = "MECH VOLUME"
    OptBot.Text = "LEVEL " & CInt(MechVolume * 100)
  ElseIf OptPos = Opt_Volume_Ramp Then
    OptTop.Text = "RAMP VOLUME"
    OptBot.Text = "LEVEL " & CInt(RampRollVolume * 100)
  ElseIf OptPos = Opt_Volume_Ball Then
    OptTop.Text = "BALL VOLUME"
    OptBot.Text = "LEVEL " & CInt(BallRollVolume * 100)
  ElseIf OptPos = Opt_Staged_Flipper Then
    OptTop.Text = "STAGED FLIPPER"
    OptBot.Text = Options_OnOffText(StagedFlipperMod)
' ElseIf OptPos = Opt_Cabinet Then
'   OptTop.Text = "CABINET MODE"
'   OptBot.Text = Options_OnOffText(CabinetMode)
' ElseIf OptPos = Opt_DynBallShadow Then
'   OptTop.Text = "DYN. BALL SHADOWS"
'   OptBot.Text = Options_OnOffText(DynamicBallShadowsOn)
' ElseIf OptPos = Opt_AmbientBallShadow Then
'   OptTop.Text = "AMB. BALL SHADOWS"
'   If AmbientBallShadowOn = 0 Then OptBot.Text = "STATIC"
'   If AmbientBallShadowOn = 1 Then OptBot.Text = "MOVING"
'   If AmbientBallShadowOn = 2 Then OptBot.Text = "FLASHER"
' ElseIf OptPos = Opt_VRRoomChoice Then
'   OptTop.Text = "VR ROOM"
'   If VRRoomChoice = 0 Then OptBot.Text = "MINIMAL"
'   If VRRoomChoice = 1 Then OptBot.Text = "ULTRA"
  ElseIf OptPos = Opt_Info_1 Then
    OptTop.Text = "VPX " & VersionMajor & "." & VersionMinor & "." & VersionRevision
    OptBot.Text = "SPONGEBOB " & TableVersion
  ElseIf OptPos = Opt_Info_2 Then
    OptTop.Text = "RENDER MODE"
    If RenderingMode = 0 Then OptBot.Text = "DEFAULT"
    If RenderingMode = 1 Then OptBot.Text = "STEREO 3D"
    If RenderingMode = 2 Then OptBot.Text = "VR"
  Elseif OptPos = Opt_Reset Then
    OptTop.Text = "RESET HIGHSCORES"
    OptBot.text = Resetoption
  Elseif OptPos = Opt_FreePlay Then
    OptTop.Text = "FREE PLAY"
    If UseCredits Then OptBot.text = "USE CREDITS" Else OptBot.text = "ENABLED"
  End If
  OptTop.Pack
  OptTop.SetAlignedPosition 127, 1, FlexDMD_Align_TopRight 'bug? not aligning right
  OptBot.SetAlignedPosition 64, 16, FlexDMD_Align_Center
' FlexDMD.UnlockRenderThread
  UpdateMods
End Sub


Sub Options_Toggle(amount)
  If FlexDMD is Nothing Then Exit Sub
  If OptPos = Opt_Light Then
    RoomBrightness = RoomBrightness + amount * 10
    If RoomBrightness < 0 Then RoomBrightness = 100
    If RoomBrightness > 100 Then RoomBrightness = 0
  ElseIf OptPos = Opt_LUT Then
    ColorLUT = ColorLUT + amount * 1
    If ColorLUT < 1 Then ColorLUT = 11
    If ColorLUT > 11 Then ColorLUT = 1
  ElseIf OptPos = Opt_Scorbit Then
    If ScorbitActive = 1 Then
      ScorbitActive = 0
      CloseFlexScorbitClaimDMD
    Else
      ScorbitActive = 1
      InitFlexScorbitDMD
    End If
    ElseIf OptPos = Opt_RomVolume Then
        RomSoundVolume = RomSoundVolume + amount * 0.1
    If RomSoundVolume < 0 Then RomSoundVolume = 1
    If RomSoundVolume > 1 Then RomSoundVolume = 0
  ElseIf OptPos = Opt_MechVolume Then
    MechVolume = MechVolume + amount * 0.1
    If MechVolume < 0 Then MechVolume = 1
    If MechVolume > 1 Then MechVolume = 0
  ElseIf OptPos = Opt_Volume_Ramp Then
    RampRollVolume = RampRollVolume + amount * 0.1
    If RampRollVolume < 0 Then RampRollVolume = 1
    If RampRollVolume > 1 Then RampRollVolume = 0
  ElseIf OptPos = Opt_Volume_Ball Then
    BallRollVolume = BallRollVolume + amount * 0.1
    If BallRollVolume < 0 Then BallRollVolume = 1
    If BallRollVolume > 1 Then BallRollVolume = 0
  ElseIf OptPos = Opt_Staged_Flipper Then
    StagedFlipperMod = 1 - StagedFlipperMod
' ElseIf OptPos = Opt_DynBallShadow Then
'   DynamicBallShadowsOn = 1 - DynamicBallShadowsOn
' ElseIf OptPos = Opt_AmbientBallShadow Then
'   AmbientBallShadowOn = AmbientBallShadowOn + 1
'   If AmbientBallShadowOn > 2 Then AmbientBallShadowOn = 0
  Elseif OptPos = Opt_Reset Then
    If Resetoption = "NO" Then Resetoption = "YES (CAREFUL)" Else Resetoption = "NO"
  Elseif OptTop.Text = "FREE PLAY" Then
    UseCredits = Not UseCredits
    If UseCredits Then OptBot.text = "USE CREDITS" Else OptBot.text = "ENABLED"
  End If
End Sub


Dim leftmagnablink
Dim rightmagnablink
Dim Flipperoption
Sub Options_KeyDown(ByVal keycode)
  If Not OptionKeysEnabled Then Exit Sub

  SkillShotSfx

  If OptSelected Then
    If keycode = LeftMagnaSave Then ' Exit / Cancel
      leftmagnablink = 1
      OptSelected = False
      Resetoption = "NO"
    ElseIf keycode = RightMagnaSave Then ' Enter / Select
      rightmagnablink = 1
      If Resetoption ="YES (CAREFUL)" Then
        Resetoption = "ARE YOU SURE"
      Elseif  Resetoption ="ARE YOU SURE" Then
        ResetHighscores
        SaveHighScores
        Resetoption = "RESET"
        OptSelected = False
        OptionBlink = 10
      Else
        OptionBlink = 6
        OptSelected = False
      End If
    ElseIf keycode = LeftFlipperKey Then ' Next / +
      Flipperoption = 1
      Options_Toggle  -1
    ElseIf keycode = RightFlipperKey Then ' Prev / -
      Flipperoption = 1
      Options_Toggle  1
    End If
  Else
    If keycode = LeftMagnaSave Then ' Exit / Cancel
      leftmagnablink = 1
      Options_Close
    ElseIf keycode = RightMagnaSave Then ' Enter / Select
      rightmagnablink = 1
      If OptPos < NOptions And optpos > 0 Then OptSelected = True   ' 0 = spongeinfoscreen
    ElseIf keycode = LeftFlipperKey Then ' Next / +
      Flipperoption = 1
      OptPos = OptPos - 1
'     If OptPos = Opt_VRTopperOn And RenderingMode <> 2 Then OptPos = OptPos - 1 ' Skip VR option in non VR mode
'     If OptPos = Opt_VRSideBlades And RenderingMode <> 2 Then OptPos = OptPos - 1 ' Skip VR option in non VR mode
'     If OptPos = Opt_VRBackglassGI And RenderingMode <> 2 Then OptPos = OptPos - 1 ' Skip VR option in non VR mode
'     If OptPos = Opt_VRRoomChoice And RenderingMode <> 2 Then OptPos = OptPos - 1 ' Skip VR option in non VR mode
      If OptPos < 0 Then OptPos = NOptions - 1
    ElseIf keycode = RightFlipperKey Then ' Prev / -
      Flipperoption = 1
      OptPos = OptPos + 1
'     If OptPos = Opt_VRRoomChoice And RenderingMode <> 2 Then OptPos = OptPos + 1 ' Skip VR option in non VR mode
'     If OptPos = Opt_VRBackglassGI And RenderingMode <> 2 Then OptPos = OptPos + 1 ' Skip VR option in non VR mode
'     If OptPos = Opt_VRSideBlades And RenderingMode <> 2 Then OptPos = OptPos + 1 ' Skip VR option in non VR mode
'     If OptPos = Opt_VRTopperOn And RenderingMode <> 2 Then OptPos = OptPos + 1 ' Skip VR option in non VR mode
      If OptPos >= NOPtions Then OptPos = 0
    End If
  End If
  Options_OnOptChg
End Sub

Sub Options_Reset
  UseCredits = False
  Credits = 0
  RoomBrightness = 30
  ColorLUT = 1
  RomSoundVolume = 0.1
  MechVolume = 0.8
  RampRollVolume = 0.5
  BallRollVolume = 0.5
  StagedFlipperMod = 0
  DynamicBallShadowsOn = 0
  AmbientBallShadowOn = 1

    UpdateMods
End Sub

Sub Options_Save
    Dim objFSO, SettingsFile

    ' Create a FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")

  If Not objFSO.FolderExists(UserDirectory) then
        Options_Reset
    Exit Sub
  End if

    ' Create or open the settings file
    Set SettingsFile = objFSO.CreateTextFile(UserDirectory & cGameName & "_opts2.txt", True)

    ' Write settings to file
    SettingsFile.WriteLine "SpongebobOptions:B"
  If UseCredits = False Then SettingsFile.WriteLine "FREEPLAY:" & "YES"
  If UseCredits = True Then SettingsFile.WriteLine "FREEPLAY:" & "NO"
    SettingsFile.WriteLine "CREDITS:" & Credits
    SettingsFile.WriteLine "LIGHT LEVEL:" & RoomBrightness
    SettingsFile.WriteLine "COLOR SATURATION:" & ColorLUT
    SettingsFile.WriteLine "SCORBIT:" & ScorbitActive
    SettingsFile.WriteLine "ROM VOLUME:" & RomSoundVolume * 100
    SettingsFile.WriteLine "MECH VOLUME:" & MechVolume * 100
    SettingsFile.WriteLine "RAMP VOLUME:" & RampRollVolume * 100
    SettingsFile.WriteLine "BALL VOLUME:" & BallRollVolume * 100
    SettingsFile.WriteLine "STAGED FLIPPER:" & StagedFlipperMod
    SettingsFile.WriteLine "DYN BALL SHADOWS:" & DynamicBallShadowsOn
    SettingsFile.WriteLine "AMB BALL SHADOWS:" & AmbientBallShadowOn



    ' Close the file
    SettingsFile.Close

    ' Clean up
    Set SettingsFile = Nothing
    Set objFSO = Nothing
End Sub

Sub Options_Load
    Dim objFSO, SettingsFile, lineData, keyValue

'Set resetvalues incase something dont load
  UseCredits = False
  RoomBrightness = 30
  ColorLUT = 1
  RomSoundVolume = 0.1
  MechVolume = 0.8
  RampRollVolume = 0.5
  BallRollVolume = 0.5
  StagedFlipperMod = 0
  DynamicBallShadowsOn = 0
  AmbientBallShadowOn = 1


    ' Create a FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")

  If Not objFSO.FolderExists(UserDirectory) then
        Options_Reset
    Exit Sub
  End if

    If objFSO.FileExists(UserDirectory & cGameName & "_opts2.txt") Then
        ' Open the settings file for reading
        Set SettingsFile = objFSO.OpenTextFile(UserDirectory & cGameName & "_opts2.txt", 1)

        ' Read settings from file
        Do Until SettingsFile.AtEndOfStream
            lineData = SettingsFile.ReadLine
            keyValue = Split(lineData, ":")

            Select Case Trim(keyValue(0))
                Case "LIGHT LEVEL"
                    RoomBrightness = Trim(keyValue(1))
                Case "COLOR SATURATION"
                    ColorLUT = Trim(keyValue(1))
        Case "SCORBIT"
                    ScorbitActive = Trim(keyValue(1))
        Case "ROM VOLUME"
          RomSoundVolume = Trim(keyValue(1)) / 100
                Case "MECH VOLUME"
          MechVolume = Trim(keyValue(1)) / 100
                Case "RAMP VOLUME"
          RampRollVolume = Trim(keyValue(1)) / 100
                Case "BALL VOLUME"
                    BallRollVolume = Trim(keyValue(1)) / 100
        Case "STAGED FLIPPER"
                    StagedFlipperMod = Trim(keyValue(1))
                Case "DYN BALL SHADOWS"
                    DynamicBallShadowsOn = 0 'Trim(keyValue(1))
                Case "AMB BALL SHADOWS"
                    AmbientBallShadowOn = 1 'Trim(keyValue(1))
        Case "FREEPLAY"
          If Trim(keyValue(1)) = "NO" Then UseCredits = True Else UseCredits = False
            End Select
        Loop
        ' Close the file
        SettingsFile.Close
    Else
        ' File does not exist, set default values
        Options_Reset
        Options_Save
    End If

    ' Clean up
    Set SettingsFile = Nothing
    Set objFSO = Nothing

  UpdateMods
End Sub

Sub UpdateMods
  Dim BL, LM, x, y, c, enabled

  '*********************
  'VR Room
  '*********************

  'If RenderingMode = 2 Then VRRoom = VRRoomChoice + 1 Else VRRoom = 0
  'SetupRoom


' '*********************
' 'Cabinet Mode
' '*********************
'
' If CabinetMode = 1 and VRRoom < 1 then
'   PinCab_Rails.visible=0
' Else
'   PinCab_Rails.visible=1
' end If

  '*********************
  'Room light level
  '*********************

  If VRRoom > 0 Then SetRoomBrightness RoomBrightness/100

  '*********************
  'Color LUT
  '*********************

  if ColorLUT = 1 Then Table1.ColorGradeImage = ""
  if ColorLUT = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-30"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "colorgradelut256x16-40"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "colorgradelut256x16-50"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "colorgradelut256x16-60"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "colorgradelut256x16-70"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "colorgradelut256x16-80"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "colorgradelut256x16-90"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "colorgradelut256x16-100"

  If ScorbitActive = 1 Then
    StartScorbit
  Else
    DMDScorebit.Visible = False
    DMDScorebit.TimerEnabled = False
    Scorbit = Null
    tmrScorbit.Enabled=False
    If Not IsNull(FlexDMDScorbit) Then
      FlexDMDScorbit.Show = False
      FlexDMDScorbit.Run = False
      FlexDMDScorbit = NULL
    End If
    CloseFlexScorbitClaimDMD()
  End IF

End Sub

Sub Wall106_Hit()
End Sub











Sub StartScorbit
  If IsNull(Scorbit) Then
    If ScorbitActive = 1 then
      ScorbitExesCheck ' Check the exe's are in the tables folder.
      If ScorbitActive = 1 then ' check again as the check exes may have disabled scorbit
        Set Scorbit = New ScorbitIF
        If Scorbit.DoInit(4156, "SPONGEQR", TableVersion, "spongebob-vpin") then  ' Prod
          tmrScorbit.Interval=2000
          tmrScorbit.UserValue = 0
          tmrScorbit.Enabled=True
          Scorbit.UploadLog = 1
        End If
      End If
    End If
  End If
End Sub

Sub InitFlexScorbitDMD()
  If IsNull(FlexDMDScorbit) Then
    Set FlexDMDScorbit = CreateObject("FlexDMD.FlexDMD")
    If FlexDMDScorbit is Nothing Then
      MsgBox "No FlexDMD found. This table will NOT run without it."
      Exit Sub
    End If
    With FlexDMDScorbit
      .ProjectFolder = TablesDir & "\SPONGEQR"
    End With
    FlexDMDScorbit.RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
    FlexDMDScorbit.Width = 300
    FlexDMDScorbit.Height = 300
    FlexDMDScorbit.Clear = True
    FlexDMDScorbit.Show = False
    FlexDMDScorbit.Run = False

    dim scene
    Set scene = FlexDMDScorbit.NewGroup("scene")
    dim qrImage, qrLabel
    Set qrLabel = FlexDMDScorbit.NewLabel("Loading", FlexDMDScorbit.NewFont("FlexDMD.Resources.bm_army-12.fnt", vbWhite, RGB(0, 0, 0), 0), "Loading Scorbit") : qrLabel.Visible = False : scene.AddActor qrLabel
    qrLabel.SetAlignedPosition 150, 150, FlexDMD_Align_Center
    FlexDMDScorbit.LockRenderThread
    FlexDMDScorbit.RenderMode =  FlexDMD_RenderMode_DMD_RGB
    FlexDMDScorbit.Stage.RemoveAll
    FlexDMDScorbit.Stage.AddActor scene
    FlexDMDScorbit.Show = False
    FlexDMDScorbit.Run = True
    FlexDMDScorbit.UnlockRenderThread


    CreateScorebitLoadingDMD()
    'Flasher
    DMDScorebit.Visible = True
    DMDScorebit.TimerEnabled = True
  End If

End Sub


Dim noQRimage :
Sub InitFlexScorbitClaimDMD()
  noQRimage = False

  If IsNull(FlexDMDScorbitClaim) Then
    Set FlexDMDScorbitClaim = CreateObject("FlexDMD.FlexDMD")
    If FlexDMDScorbitClaim is Nothing Then
      MsgBox "No FlexDMD found. This table will NOT run without it."
      Exit Sub
    End If
    With FlexDMDScorbitClaim
      .ProjectFolder = TablesDir & "\SPONGEQR"
    End With
    FlexDMDScorbitClaim.RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
    FlexDMDScorbitClaim.Width = 300
    FlexDMDScorbitClaim.Height = 300
    FlexDMDScorbitClaim.Clear = True
    FlexDMDScorbitClaim.Show = False
    FlexDMDScorbitClaim.Run = False

    dim scene
    Set scene = FlexDMDScorbitClaim.NewGroup("scene")
    dim qrImage
    DMDScorebitClaim.Visible = True
  On Error Resume Next
    Set qrImage = FlexDMDScorbitClaim.NewImage("QRClaim",   "QRClaim.png")  : qrImage.SetBounds 0, 0, 300, 300 : qrImage.Visible = True : scene.AddActor qrImage
    If err Then FlexDMDScorbitClaim=nothing : noQRimage = True : exit sub
    err clr

    FlexDMDScorbitClaim.LockRenderThread
    FlexDMDScorbitClaim.RenderMode =  FlexDMD_RenderMode_DMD_RGB
    FlexDMDScorbitClaim.Stage.RemoveAll
    FlexDMDScorbitClaim.Stage.AddActor scene
    FlexDMDScorbitClaim.Show = False
    FlexDMDScorbitClaim.Run = True
    FlexDMDScorbitClaim.UnlockRenderThread
    'Flasher
    DMDScorebitClaim.TimerEnabled = True

  End If
End Sub

Sub CloseFlexScorbitClaimDMD()

  If Not IsNull(FlexDMDScorbitClaim) Then
    FlexDMDScorbitClaim.Show = False
    FlexDMDScorbitClaim.Run = False
    FlexDMDScorbitClaim = NULL
    DMDScorebitClaim.TimerEnabled = False
    DMDScorebitClaim.Visible = False
    End If

End Sub

Sub DMDScorebit_Timer()
  Dim DMDScoreBitp
  DMDScoreBitp = FlexDMDScorbit.DmdColoredPixels
  If Not IsEmpty(DMDScoreBitp) Then
    DMDScorebit.DMDWidth = FlexDMDScorbit.Width
    DMDScorebit.DMDHeight = FlexDMDScorbit.Height
    DMDScorebit.DMDColoredPixels = DMDScoreBitp
  End If
End Sub

Sub DMDScorebitClaim_Timer()
  Dim DMDScoreBitp
  DMDScoreBitp = FlexDMDScorbitClaim.DmdColoredPixels
  If Not IsEmpty(DMDScoreBitp) Then
    DMDScorebitClaim.DMDWidth = FlexDMDScorbitClaim.Width
    DMDScorebitClaim.DMDHeight = FlexDMDScorbitClaim.Height
    DMDScorebitClaim.DMDColoredPixels = DMDScoreBitp
  End If
End Sub


Sub CreateScorebitLoadingDMD

  FlexDMDScorbit.LockRenderThread
  dim scene
  Set scene = FlexDMDScorbit.Stage.GetGroup("scene")
  FlexDMDScorbit.Stage.GetLabel("Loading").Visible = True
  FlexDMDScorbit.Stage.GetLabel("Loading").Text = "LOADING SCORBIT"
  If Not IsNull(Scorbit) Then
    If Scorbit.bNeedsPairing = False Then
      FlexDMDScorbit.Stage.GetLabel("Loading").Text = "PAIRED"
    End If
  End If

  FlexDMDScorbit.Stage.GetLabel("Loading").SetAlignedPosition 150, 150, FlexDMD_Align_Center
  If scene.HasChild("QRPairing") = True Then
    FlexDMDScorbit.Stage.GetImage("QRPairing").Visible = False
  End If
  If scene.HasChild("QRClaim") = True Then
    FlexDMDScorbit.Stage.GetImage("QRClaim").Visible = False
  End If
  FlexDMDScorbit.Show = False
  FlexDMDScorbit.Run = True
  FlexDMDScorbit.UnlockRenderThread

End Sub


Sub CreateScorebitPairingDMD

  FlexDMDScorbit.LockRenderThread
  dim scene, qrImage
  Set scene = FlexDMDScorbit.Stage.GetGroup("scene")
  If scene.HasChild("QRPairing") = False Then
    Set qrImage = FlexDMDScorbit.NewImage("QRPairing",    "QRCode.png") : qrImage.SetBounds 50, 0, 200, 200 : qrImage.Visible = False : scene.AddActor qrImage
  End If
  If Scorbit.bNeedsPairing = True Then
    FlexDMDScorbit.Stage.GetLabel("Loading").Text = "MACHINE NEEDS PAIRING"
  Else
    FlexDMDScorbit.Stage.GetLabel("Loading").Text = "PAIRED"
  End If
  FlexDMDScorbit.Stage.GetLabel("Loading").SetAlignedPosition 150, 250, FlexDMD_Align_Center
  FlexDMDScorbit.Stage.GetLabel("Loading").Visible = True
  If scene.HasChild("QRPairing") = True Then
    FlexDMDScorbit.Stage.GetImage("QRPairing").Visible = True
  End If
  If scene.HasChild("QRClaim") Then
    FlexDMDScorbit.Stage.GetImage("QRClaim").Visible = False
  End If
  FlexDMDScorbit.Show = False
  FlexDMDScorbit.Run = True
  FlexDMDScorbit.UnlockRenderThread

End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  SCORBIT Interface
' To Use:
' 1) Define a timer tmrScorbit
' 2) Call DoInit at the end of PupInit or in Table Init if you are nto using pup with the appropriate parameters
'   if Scorbit.DoInit(389, "PupOverlays", "1.0.0", "GRWvz-MP37P") then
'     tmrScorbit.Interval=2000
'     tmrScorbit.UserValue = 0
'     tmrScorbit.Enabled=True
'   End if
' 3) Customize helper functions below for different events if you want or make your own
' 4) Call
'   StartSession - When a game starts
'   StopSession - When the game is over
'   SendUpdate - When Score Changes
'   SetGameMode - When different game events happen like starting a mode, MB etc.  (ScorbitBuildGameModes helper function shows you how)
' 5) Drop the binaries sQRCode.exe and sToken.exe in your Pup Root so we can create session kets and QRCodes.
' 6) Callbacks
'   Scorbit_Paired      - Called when machine is successfully paired.  Hide QRCode and play a sound
'   Scorbit_PlayerClaimed - Called when player is claimed.  Hide QRCode, play a sound and display name
'
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' TABLE CUSTOMIZATION START HERE


Sub Scorbit_NeedsPairing()                ' Scorbit callback when new machine needs pairing

  If bInOptions = False Or Not OptPos=3 Then

    DMDScorebit.Visible = False
    DMDScorebit.TimerEnabled = False
    Scorbit = Null
    tmrScorbit.Enabled=False
    ScorbitActive = 0
    If Not IsNull(FlexDMDScorbit) Then
      FlexDMDScorbit.Show = False
      FlexDMDScorbit.Run = False
      FlexDMDScorbit = NULL
    End If

  Else
    CreateScorebitPairingDMD()
  End If
End Sub

Sub Scorbit_Paired()                ' Scorbit callback when new machine is paired
  If bInOptions = False Or Not OptPos=3 Then
    DMDScorebit.Visible = False
    DMDScorebit.TimerEnabled = False
    If Not IsNull(FlexDMDScorbit) Then
      FlexDMDScorbit.Show = False
      FlexDMDScorbit.Run = False
      FlexDMDScorbit = NULL
    End If
  Else
    CreateScorebitLoadingDMD()
  End If

  If startgame = 1 Then
    If CurrentBall = 1 Then
      If Not IsNull(Scorbit) Then
        If ScorbitActive = 1 And Scorbit.bNeedsPairing = False Then
          Scorbit.StartSession()
        End If
      End If
    End If
  End If


End Sub

Sub Scorbit_PlayerClaimed(PlayerNum, PlayerName)  ' Scorbit callback when QR Is Claimed


End Sub


Sub ScorbitClaimQR(bShow)           '  Show QRCode on first ball for users to claim this position

End Sub

Sub ScorbitBuildGameModes()   ' Custom function to build the game modes for better stats
  dim GameModeStr

  Scorbit.SetGameMode(GameModeStr)

End Sub

Sub Scorbit_LOGUpload(state)  ' Callback during the log creation process.  0=Creating Log, 1=Uploading Log, 2=Done
  Select Case state
    case 0:
      'Debug.print "CREATING LOG"
    case 1:
      'Debug.print "Uploading LOG"
    case 2:
      'Debug.print "LOG Complete"
  End Select
End Sub
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' TABLE CUSTOMIZATION END HERE - NO NEED TO EDIT BELOW THIS LINE


dim Scorbit : Scorbit = Null
' Workaround - Call get a reference to Member Function
Sub tmrScorbit_Timer()                ' Timer to send heartbeat
  Scorbit.DoTimer(tmrScorbit.UserValue)
  tmrScorbit.UserValue=tmrScorbit.UserValue+1
  if tmrScorbit.UserValue>5 then tmrScorbit.UserValue=0
End Sub
Function ScorbitIF_Callback()
  Scorbit.Callback()
End Function
Class ScorbitIF

  Public bSessionActive
  Public bNeedsPairing
  Private bUploadLog
  Private bActive
  Private LOGFILE(10000000)
  Private LogIdx

  Private bProduction

  Private TypeLib
  Private MyMac
  Private Serial
  Private MyUUID
  Private TableVersion

  Private SessionUUID
  Private SessionSeq
  Private SessionTimeStart
  Private bRunAsynch
  Private bWaitResp
  Private GameMode
  Private GameModeOrig    ' Non escaped version for log
  Private VenueMachineID
  Private CachedPlayerNames(4)
  Private SaveCurrentPlayer

  Public bEnabled
  Private sToken
  Private machineID
  Private dirQRCode
  Private opdbID
  Private wsh

  Private objXmlHttpMain
  Private objXmlHttpMainAsync
  Private fso
  Private Domain

  Public Sub Class_Initialize()
    bActive="false"
    bSessionActive=False
    bEnabled=False
  End Sub

  Property Let UploadLog(bValue)
    bUploadLog = bValue
  End Property

  Sub DoTimer(bInterval)  ' 2 second interval
    dim holdScores(4)
    dim i

    if bInterval=0 then
      SendHeartbeat()
    elseif bRunAsynch then ' Game in play
      'debug.Print("Scorbit.SendUpdate " & TotalScore(1) & ", " & TotalScore(2) & ", " & TotalScore(3) & ", " & TotalScore(4) & ", " & 4-CurrentBall & ", " & CurrentPlayer & ", " & Players)

      If currentball < 1 Then
        Scorbit.SendUpdate TotalScore(1), TotalScore(2), TotalScore(3), TotalScore(4), 0 , CurrentPlayer, Players
      Else
        Scorbit.SendUpdate TotalScore(1), TotalScore(2), TotalScore(3), TotalScore(4), 4-CurrentBall, CurrentPlayer, Players
      End If

    End if
  End Sub

  Function GetName(PlayerNum) ' Return Parsed Players name
    if PlayerNum<1 or PlayerNum>4 then
      GetName=""
    else
      GetName=CachedPlayerNames(PlayerNum)
    End if
  End Function

  Function DoInit(MyMachineID, Directory_PupQRCode, Version, opdb)
    dim Nad
    Dim EndPoint
    Dim resultStr
    Dim UUIDParts
    Dim UUIDFile

    bProduction=1
'   bProduction=0
    SaveCurrentPlayer=0
    VenueMachineID=""
    bWaitResp=False
    bRunAsynch=False
    DoInit=False
    opdbID=opdb
    dirQrCode=Directory_PupQRCode
    MachineID=MyMachineID
    TableVersion=version
    bNeedsPairing=False
    if bProduction then
      domain = "api.scorbit.io"
    else
      domain = "staging.scorbit.io"
      domain = "scorbit-api-staging.herokuapp.com"
    End if
    'msgbox "production: " & bProduction
    Set fso = CreateObject("Scripting.FileSystemObject")
    dim objLocator:Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Dim objService:Set objService = objLocator.ConnectServer(".", "root\cimv2")
    Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP")
    Set objXmlHttpMainAsync = CreateObject("Microsoft.XMLHTTP")
    objXmlHttpMain.onreadystatechange = GetRef("ScorbitIF_Callback")
    Set wsh = CreateObject("WScript.Shell")

    ' Get Mac for Serial Number
    dim Nads: set Nads = objService.ExecQuery("Select * from Win32_NetworkAdapter where physicaladapter=true")
    for each Nad in Nads
      if not isnull(Nad.MACAddress) then
        'msgbox "Using MAC Addresses:" & Nad.MACAddress & " From Adapter:" & Nad.description
        MyMac=replace(Nad.MACAddress, ":", "")
        Exit For
      End if
    Next
    Serial=eval("&H" & mid(MyMac, 5))
'   Serial=123456
  ' debug.print "Serial: " & Serial
    serial = serial + MachineID
  ' debug.print "New Serial with machine id: " & Serial

    ' Get System UUID
    set Nads = objService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
    for each Nad in Nads
      'msgbox "Using UUID:" & Nad.UUID
      MyUUID=Nad.UUID
      Exit For
    Next

    if MyUUID="" then
      MsgBox "SCORBIT - Can get UUID, Disabling."
      Exit Function
    elseif MyUUID="03000200-0400-0500-0006-000700080009" or ScorbitAlternateUUID then
      If fso.FolderExists(UserDirectory) then
        If fso.FileExists(UserDirectory & "ScorbitUUID.dat") then
          Set UUIDFile = fso.OpenTextFile(UserDirectory & "ScorbitUUID.dat",1)
          MyUUID = UUIDFile.ReadLine()
          UUIDFile.Close
          Set UUIDFile = Nothing
        Else
          MyUUID=GUID()
          Set UUIDFile=fso.CreateTextFile(UserDirectory & "ScorbitUUID.dat",True)
          UUIDFile.WriteLine MyUUID
          UUIDFile.Close
          Set UUIDFile=Nothing
    End if
      End if
    End if

    ' Clean UUID
    UUIDParts=split(MyUUID, "-")
    'msgbox UUIDParts(0)
    MyUUID=LCASE(Hex(eval("&h" & UUIDParts(0))+MyMachineID)) & UUIDParts(1) &  UUIDParts(2) &  UUIDParts(3) & UUIDParts(4)     ' Add MachineID to UUID
    'msgbox UUIDParts(0)
    MyUUID=LPad(MyUUID, 32, "0")
'   MyUUID=Replace(MyUUID, "-",  "")
  ' Debug.print "MyUUID:" & MyUUID


' Debug
'   myUUID="adc12b19a3504453a7414e722f58737f"
'   Serial="123456778"

    'create own folder for table QRCodes TablesDir & "\" & dirQrCode
    If Not fso.FolderExists(TablesDir & "\" & dirQrCode) then
      fso.CreateFolder(TablesDir & "\" & dirQrCode)
    end if

    ' Authenticate and get our token
    if getStoken() then
      bEnabled=True
'     SendHeartbeat
      DoInit=True
    End if
  End Function



  Sub Callback()
    Dim ResponseStr
    Dim i
    Dim Parts
    Dim Parts2
    Dim Parts3
    if bEnabled=False then Exit Sub

    if bWaitResp and objXmlHttpMain.readystate=4 then
      'Debug.print "CALLBACK: " & objXmlHttpMain.Status & " " & objXmlHttpMain.readystate
      if objXmlHttpMain.Status=200 and objXmlHttpMain.readystate = 4 then
        ResponseStr=objXmlHttpMain.responseText
        'Debug.print "RESPONSE: " & ResponseStr

        ' Parse Name
        if CachedPlayerNames(SaveCurrentPlayer-1)="" then  ' Player doesnt have a name
          if instr(1, ResponseStr, "cached_display_name") <> 0 Then ' There are names in the result
            Parts=Split(ResponseStr,",{")             ' split it
            if ubound(Parts)>=SaveCurrentPlayer-1 then        ' Make sure they are enough avail
              if instr(1, Parts(SaveCurrentPlayer-1), "cached_display_name")<>0 then  ' See if mine has a name
                CachedPlayerNames(SaveCurrentPlayer-1)=GetJSONValue(Parts(SaveCurrentPlayer-1), "cached_display_name")    ' Get my name
                CachedPlayerNames(SaveCurrentPlayer-1)=Replace(CachedPlayerNames(SaveCurrentPlayer-1), """", "")
                Scorbit_PlayerClaimed SaveCurrentPlayer, CachedPlayerNames(SaveCurrentPlayer-1)
                'Debug.print "Player Claim:" & SaveCurrentPlayer & " " & CachedPlayerNames(SaveCurrentPlayer-1)
              End if
            End if
          End if
        else                            ' Check for unclaim
          if instr(1, ResponseStr, """player"":null")<>0 Then ' Someone doesnt have a name
            Parts=Split(ResponseStr,"[")            ' split it
'Debug.print "Parts:" & Parts(1)
            Parts2=Split(Parts(1),"}")              ' split it
            for i = 0 to Ubound(Parts2)
'Debug.print "Parts2:" & Parts2(i)
            if instr(1, Parts2(i), """player"":null")<>0 Then
                CachedPlayerNames(i)=""
              End if
            Next
          End if
        End if
      End if
      bWaitResp=False
    End if
  End Sub



  Public Sub StartSession()
    if bEnabled=False then Exit Sub
'msgbox  "Scorbit Start Session"
    CachedPlayerNames(0)=""
    CachedPlayerNames(1)=""
    CachedPlayerNames(2)=""
    CachedPlayerNames(3)=""
    bRunAsynch=True
    bActive="true"
    bSessionActive=True
    SessionSeq=0
    SessionUUID=GUID()
    SessionTimeStart=GameTime
    LogIdx=0
    SendUpdate 0, 0, 0, 0, 1, 1, 1
  End Sub

  Public Sub StopSession(P1Score, P2Score, P3Score, P4Score, NumberPlayers)
    StopSession2 P1Score, P2Score, P3Score, P4Score, NumberPlayers, False
  End Sub

  Public Sub StopSession2(P1Score, P2Score, P3Score, P4Score, NumberPlayers, bCancel)
    Dim i
    dim objFile
    if bEnabled=False then Exit Sub
    if bSessionActive=False then Exit Sub
'Debug.print "Scorbit Stop Session"

    bRunAsynch=False
    bActive="false"
    bSessionActive=False
    SendUpdate P1Score, P2Score, P3Score, P4Score, -1, -1, NumberPlayers
'   SendHeartbeat

    if bUploadLog and LogIdx<>0 and bCancel=False then
  '   Debug.print "Creating Scorbit Log: Size" & LogIdx
      Scorbit_LOGUpload(0)
'     Set objFile = fso.CreateTextFile(puplayer.getroot&"\" & cGameName & "\sGameLog.csv")
      Set objFile = fso.CreateTextFile(TablesDir & "\SPONGEQR\sGameLog.csv")
      For i = 0 to LogIdx-1
        objFile.Writeline LOGFILE(i)
      Next
      objFile.Close
      LogIdx=0
      Scorbit_LOGUpload(1)
'     pvPostFile "https://" & domain & "/api/session_log/", puplayer.getroot&"\" & cGameName & "\sGameLog.csv", False
      pvPostFile "https://" & domain & "/api/session_log/", TablesDir & "\SPONGEQR\sGameLog.csv", False
      Scorbit_LOGUpload(2)
      on error resume next
'     fso.DeleteFile(puplayer.getroot&"\" & cGameName & "\sGameLog.csv")
      fso.DeleteFile(TablesDir & "\SPONGEQR\sGameLog.csv")
      on error goto 0
    End if

  End Sub

  Public Sub SetGameMode(GameModeStr)
    GameModeOrig=GameModeStr
    GameMode=GameModeStr
    GameMode=Replace(GameMode, ":", "%3a")
    GameMode=Replace(GameMode, ";", "%3b")
    GameMode=Replace(GameMode, " ", "%20")
    GameMode=Replace(GameMode, "{", "%7B")
    GameMode=Replace(GameMode, "}", "%7D")
  End sub

  Public Sub SendUpdate(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers)
    SendUpdateAsynch P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers, bRunAsynch
  End Sub

  Public Sub SendUpdateAsynch(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers, bAsynch)
    dim i
    Dim PostData
    Dim resultStr
    dim LogScores(4)

    if bUploadLog then
      if NumberPlayers>=1 then LogScores(0)=P1Score
      if NumberPlayers>=2 then LogScores(1)=P2Score
      if NumberPlayers>=3 then LogScores(2)=P3Score
      if NumberPlayers>=4 then LogScores(3)=P4Score
      LOGFILE(LogIdx)=DateDiff("S", "1/1/1970", Now()) & "," & LogScores(0) & "," & LogScores(1) & "," & LogScores(2) & "," & LogScores(3) & ",,," &  CurrentPlayer & "," & CurrentBall & ",""" & GameModeOrig & """"
      LogIdx=LogIdx+1
    End if

    if bEnabled=False then Exit Sub
    if bWaitResp then exit sub ' Drop message until we get our next response
'   debug.print "Current players: " & CurrentPlayer
    SaveCurrentPlayer=CurrentPlayer
'   PostData = "session_uuid=" & SessionUUID & "&session_time=" & DateDiff("S", "1/1/1970", Now()) & _
'         "&session_sequence=" & SessionSeq & "&active=" & bActive
    PostData = "session_uuid=" & SessionUUID & "&session_time=" & GameTime-SessionTimeStart+1 & _
          "&session_sequence=" & SessionSeq & "&active=" & bActive

    SessionSeq=SessionSeq+1
    if NumberPlayers > 0 then
      for i = 0 to NumberPlayers-1
        PostData = PostData & "&current_p" & i+1 & "_score="
        if i <= NumberPlayers-1 then
                    if i = 0 then PostData = PostData & P1Score
                    if i = 1 then PostData = PostData & P2Score
                    if i = 2 then PostData = PostData & P3Score
                    if i = 3 then PostData = PostData & P4Score
        else
          PostData = PostData & "-1"
        End if
      Next

      PostData = PostData & "&current_ball=" & CurrentBall & "&current_player=" & CurrentPlayer
      if GameMode<>"" then PostData=PostData & "&game_modes=" & GameMode
    End if
    resultStr = PostMsg("https://" & domain, "/api/entry/", PostData, bAsynch)
'   if resultStr<>"" then Debug.print "SendUpdate Resp:" & resultStr
  End Sub

' PRIVATE BELOW
  Private Function LPad(StringToPad, Length, CharacterToPad)
    Dim x : x = 0
    If Length > Len(StringToPad) Then x = Length - len(StringToPad)
    LPad = String(x, CharacterToPad) & StringToPad
  End Function

  Private Function GUID()
    Dim TypeLib
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    GUID = Mid(TypeLib.Guid, 2, 36)

'   Set wsh = CreateObject("WScript.Shell")
'   Set fso = CreateObject("Scripting.FileSystemObject")
'
'   dim rc
'   dim result
'   dim objFileToRead
'   Dim sessionID:sessionID=puplayer.getroot&"\" & cGameName & "\sessionID.txt"
'
'   on error resume next
'   fso.DeleteFile(sessionID)
'   On error goto 0
'
'   rc = wsh.Run("powershell -Command ""(New-Guid).Guid"" | out-file -encoding ascii " & sessionID, 0, True)
'   if FileExists(sessionID) and rc=0 then
'     Set objFileToRead = fso.OpenTextFile(sessionID,1)
'     result = objFileToRead.ReadLine()
'     objFileToRead.Close
'     GUID=result
'   else
'     MsgBox "Cant Create SessionUUID through powershell. Disabling Scorbit"
'     bEnabled=False
'   End if

  End Function

  Private Function GetJSONValue(JSONStr, key)
    dim i
    Dim tmpStrs,tmpStrs2
    if Instr(1, JSONStr, key)<>0 then
      tmpStrs=split(JSONStr,",")
      for i = 0 to ubound(tmpStrs)
        if instr(1, tmpStrs(i), key)<>0 then
          tmpStrs2=split(tmpStrs(i),":")
          GetJSONValue=tmpStrs2(1)
          exit for
        End if
      Next
    End if
  End Function

  Private Sub SendHeartbeat()
    Dim resultStr
    dim TmpStr
    Dim Command
    Dim rc
'   Dim QRFile:QRFile=puplayer.getroot&"\" & cGameName & "\" & dirQrCode
    Dim QRFile:QRFile=TablesDir & "\" & dirQrCode
    if bEnabled=False then Exit Sub
    resultStr = GetMsgHdr("https://" & domain, "/api/heartbeat/", "Authorization", "SToken " & sToken)
'Debug.print "Heartbeat Resp:" & resultStr
    If VenueMachineID="" then

      if resultStr<>"" and Instr(resultStr, """unpaired"":true")=0 then   ' We Paired
        bNeedsPairing=False
    '   debug.print "Paired"
        Scorbit_Paired()
      else
    '   debug.print "Needs Pairing"
        bNeedsPairing=True
        Scorbit_NeedsPairing()
'       if not FScorbitQRIcon.visible then showQRPairImage
      End if

      TmpStr=GetJSONValue(resultStr, "venuemachine_id")
      if TmpStr<>"" then
        VenueMachineID=TmpStr
'Debug.print "VenueMachineID=" & VenueMachineID
'       Command = puplayer.getroot&"\" & cGameName & "\sQRCode.exe " & VenueMachineID & " " & opdbID & " " & QRFile
      ' debug.print "RUN sqrcode"
        Command = """" & TablesDir & "\sQRCode.exe"" " & VenueMachineID & " " & opdbID & " """ & QRFile & """"
'       msgbox Command
        rc = wsh.Run(Command, 0, False)
      End if
    End if
  End Sub

  Private Function getStoken()
    Dim result
    Dim results
'   dim wsh
    Dim tmpUUID:tmpUUID="adc12b19a3504453a7414e722f58736b"
    Dim tmpVendor:tmpVendor="vscorbitron"
    Dim tmpSerial:tmpSerial="999990104"
'   Dim QRFile:QRFile=puplayer.getroot&"\" & cGameName & "\" & dirQrCode
    Dim QRFile:QRFile=TablesDir & "\" & dirQrCode
'   Dim sTokenFile:sTokenFile=puplayer.getroot&"\" & cGameName & "\sToken.dat"
    Dim sTokenFile:sTokenFile=TablesDir & "\sToken.dat"

    ' Set everything up
    tmpUUID=MyUUID
    tmpVendor="vpin"
    tmpSerial=Serial

    on error resume next
    fso.DeleteFile("""" & sTokenFile & """")
    On error goto 0

    ' get sToken and generate QRCode
'   Set wsh = CreateObject("WScript.Shell")
    Dim waitOnReturn: waitOnReturn = True
    Dim windowStyle: windowStyle = 0
    Dim Command
    Dim rc
    Dim objFileToRead

'   msgbox """" & " 55"

'   Command = puplayer.getroot&"\" & cGameName & "\sToken.exe " & tmpUUID & " " & tmpVendor & " " &  tmpSerial & " " & MachineID & " " & QRFile & " " & sTokenFile & " " & domain
  ' debug.print "RUN sToken"
    Command = """" & TablesDir & "\sToken.exe"" " & tmpUUID & " " & tmpVendor & " " &  tmpSerial & " " & MachineID & " """ & QRFile & """ """ & sTokenFile & """ " & domain
'msgbox "RUNNING:" & Command
    rc = wsh.Run(Command, windowStyle, waitOnReturn)
'msgbox "Return:" & rc
'   if FileExists(puplayer.getroot&"\" & cGameName & "\sToken.dat") and rc=0 then
'   msgbox """" & TablesDir & "\sToken.dat"""
    if FileExists(TablesDir & "\sToken.dat") and rc=0 then
'     Set objFileToRead = fso.OpenTextFile(puplayer.getroot&"\" & cGameName & "\sToken.dat",1)
'     msgbox """" & TablesDir & "\sToken.dat"""
      Set objFileToRead = fso.OpenTextFile(TablesDir & "\sToken.dat",1)
      result = objFileToRead.ReadLine()
      objFileToRead.Close
      Set objFileToRead = Nothing
'msgbox result

      if Instr(1, result, "Invalid timestamp")<> 0 then
        MsgBox "Scorbit Timestamp Error: Please make sure the time on your system is exact"
        getStoken=False
      elseif Instr(1, result, "Internal Server error")<> 0 then
        MsgBox "Scorbit Internal Server error ??"
        getStoken=False
      elseif Instr(1, result, ":")<>0 then
        results=split(result, ":")
        sToken=results(1)
        sToken=mid(sToken, 3, len(sToken)-4)
'Debug.print "Got TOKEN:" & sToken
        getStoken=True
      Else
'Debug.print "ERROR:" & result
        getStoken=False
      End if
    else
'msgbox "ERROR No File:" & rc
    End if

  End Function

  private Function FileExists(FilePath)
    If fso.FileExists(FilePath) Then
      FileExists=CBool(1)
    Else
      FileExists=CBool(0)
    End If
  End Function

  Private Function GetMsg(URLBase, endpoint)
    GetMsg = GetMsgHdr(URLBase, endpoint, "", "")
  End Function

  Private Function GetMsgHdr(URLBase, endpoint, Hdr1, Hdr1Val)
    Dim Url
    Url = URLBase + endpoint & "?session_active=" & bActive
'Debug.print "Url:" & Url  & "  Async=" & bRunAsynch
    objXmlHttpMain.open "GET", Url, bRunAsynch
'   objXmlHttpMain.setRequestHeader "Content-Type", "text/xml"
    objXmlHttpMain.setRequestHeader "Cache-Control", "no-cache"
    if Hdr1<> "" then objXmlHttpMain.setRequestHeader Hdr1, Hdr1Val

'   on error resume next
      err.clear
      objXmlHttpMain.send ""
      if err.number=-2147012867 then
        MsgBox "Multiplayer Server is down (" & err.number & ") " & Err.Description
        bEnabled=False
      elseif err.number <> 0 then
        debug.print "Server error: (" & err.number & ") " & Err.Description
      End if
      if bRunAsynch=False then
          Debug.print "Status: " & objXmlHttpMain.status
          If objXmlHttpMain.status = 200 Then
            GetMsgHdr = objXmlHttpMain.responseText
        Else
            GetMsgHdr=""
          End if
      Else
        bWaitResp=True
        GetMsgHdr=""
      End if
'   On error goto 0

  End Function

  Private Function PostMsg(URLBase, endpoint, PostData, bAsynch)
    Dim Url

    Url = URLBase + endpoint
'debug.print "PostMSg:" & Url & " " & PostData

    objXmlHttpMain.open "POST",Url, bAsynch
    objXmlHttpMain.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objXmlHttpMain.setRequestHeader "Content-Length", Len(PostData)
    objXmlHttpMain.setRequestHeader "Cache-Control", "no-cache"
    objXmlHttpMain.setRequestHeader "Authorization", "SToken " & sToken
    if bAsynch then bWaitResp=True

    on error resume next
      objXmlHttpMain.send PostData
      if err.number=-2147012867 then
        MsgBox "Multiplayer Server is down (" & err.number & ") " & Err.Description
        bEnabled=False
      elseif err.number <> 0 then
        'debug.print "Multiplayer Server error (" & err.number & ") " & Err.Description
      End if
      If objXmlHttpMain.status = 200 Then
        PostMsg = objXmlHttpMain.responseText
      else
        PostMsg="ERROR: " & objXmlHttpMain.status & " >" & objXmlHttpMain.responseText & "<"
      End if
    On error goto 0
  End Function

  Private Function pvPostFile(sUrl, sFileName, bAsync)
'Debug.print "Posting File " & sUrl & " " & sFileName & " " & bAsync & " File:" & Mid(sFileName, InStrRev(sFileName, "\") + 1)
    Dim STR_BOUNDARY:STR_BOUNDARY  = GUID()
    Dim nFile
    Dim baBuffer()
    Dim sPostData
    Dim Response

    '--- read file
    Set nFile = fso.GetFile(sFileName)
    With nFile.OpenAsTextStream()
      sPostData = .Read(nFile.Size)
      .Close
    End With
'   fso.Open sFileName For Binary Access Read As nFile
'   If LOF(nFile) > 0 Then
'     ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
'     Get nFile, , baBuffer
'     sPostData = StrConv(baBuffer, vbUnicode)
'   End If
'   Close nFile

    '--- prepare body
    sPostData = "--" & STR_BOUNDARY & vbCrLf & _
      "Content-Disposition: form-data; name=""uuid""" & vbCrLf & vbCrLf & _
      SessionUUID & vbcrlf & _
      "--" & STR_BOUNDARY & vbCrLf & _
      "Content-Disposition: form-data; name=""log_file""; filename=""" & SessionUUID & ".csv""" & vbCrLf & _
      "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
      sPostData & vbCrLf & _
      "--" & STR_BOUNDARY & "--"

'Debug.print "POSTDATA: " & sPostData & vbcrlf

    '--- post
    With objXmlHttpMain
      .Open "POST", sUrl, bAsync
      .SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
      .SetRequestHeader "Authorization", "SToken " & sToken
      .Send sPostData ' pvToByteArray(sPostData)
      If Not bAsync Then
        Response= .ResponseText
        pvPostFile = Response
'Debug.print "Upload Response: " & Response
      End If
    End With

  End Function

  Private Function pvToByteArray(sText)
    pvToByteArray = StrConv(sText, 128)   ' vbFromUnicode
  End Function

End Class
'  END SCORBIT
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


''QRView support by iaakki
Function GetTablesFolder()
    Dim GTF
    Set GTF = CreateObject("Scripting.FileSystemObject")
    GetTablesFolder= GTF.GetParentFolderName(userdirectory) & "\Tables"
    set GTF = nothing
End Function

' Checks that all needed binaries are available in correct place..
Sub ScorbitExesCheck
  dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")

  If fso.FileExists(TablesDir & "\sToken.exe") then
'   msgbox "Stoken.exe found at: " & TablesDir & "\sToken.exe"
  else
    msgbox "Stoken.exe NOT found at: " & TablesDir & "\sToken.exe Disabling Scorbit for now."
    Scorbitactive = 0
    SaveValue cGameName, "SCORBIT", ScorbitActive
  end if

  If fso.FileExists(TablesDir & "\sQRCode.exe") then
'   msgbox "sQRCode.exe found at: " & TablesDir & "\sQRCode.exe"
  else
    msgbox "sQRCode.exe NOT found at: " & TablesDir & "\sQRCode.exe Disabling Scorbit for now."
    Scorbitactive = 0
    SaveValue cGameName, "SCORBIT", ScorbitActive
  end if
end sub






' CHANGE LOG
' 002 - apophis - Added Lampz. Added temporary visible insert lights
' 004 - apophis - Added VLM helper code. Added RoomBrightness option. Added Roth target code.
' 006 - apophis - Udpated VLM code. Added flasher subs.
' 007 - flux    - Added light state controller, modulated flashers.
' 009 - apophis - Fixed gates visiblity. Fixed center target bank height. Updated desktop background
' 010 - Gedankekojote97 - new render with split gi, spots & clear blubs
' 011 - flux - hooked up GI Strings, added GI animations and coloured lightmaps
' 012 - flux - added updatelightmapwithcolor for colored gi and flashers
' 017 - apophis - Fixes to to gates and rollovers visibilities. Animate diverter and red drop target under pineapple. Reworked target and toy shake code. Ball brightness code added. Other random fixes.
' 018 -
' 019 -
' 020 - friscopinball - rendered, normalized, and cleaned all backglass audio sounds/effects
' 020a - Rawd - Added VR room
' 021 - apophis - Fixed gate animations. Fixed rollover wires. Tried to fix ST target animations. Optimized animations (some dont need to always update). Set up RoomBrightness option. Made 3 target bank easier to drop (similar to AFM logic).
'.022 - Rawd - Finalized VR room and code
' 023 - flux - sling flasher improvements
' 024 - apophis - Soften the sideblade light maps. Made physical scoop a little wider. Increased small flipper strength.
' 025 - apophis - Fixed/added tilt functionality for desktop and mechanical tilt bob. Fixed stuck ball issue at top of PF. Made it so flippers dont die in mode 2.
' 025 - Gedankekojote97 - Fixed B2S Code and implemented triggers. Cleaned up settings
' 032 - Leojreimroc - VR Backglass Lighting. Slightly adjusted DMD in VR.
' 036 - apophis - Updated with new bake and VLM scripts. Fixed some scripting to work with moviable objects corrected orientations.
' 037 - apophis - Updated with new bake and VLM scripts
' 038 - apophis - Updated with new toolkit import and VLM scripts. The new toolkit import includes fewer tri compared to previous. Automated CabMode.
' 039 - apophis - Updated with new toolkit import and VLM scripts. The new toolkit import includes Flupper's optimized figure models (Thanks!). Merged in jsm174's script updates (Thanks!). Light state controller still broken.
' 040 - flux - update light state controller for standalone vpx.
' 041 - flux - added jsm174 fixes, coded shooter lane and mode ready flashers.
' 042 - apophis - Updated with new toolkit import and VLM scripts.
' 043 - friscopinball - fixed some sounds
' 044 - apophis - Added flux's script patch. Prevented ball reflections due to contrl lights. Updated some walls to make ball stucks less likely. Darkened pf image a bit. Fixed Sfx_Greenhorn sound.
' 045 - flux - fix GI coloured lights
' 046 - apophis - Update logic to make 3-bank and droptarget easier to drop to start MB.  Inlane ball slowdown code added. PF friction increased. Main flipper strength decreased slightly. Randomize drainshooter angle in RandomKicker sub.
' 047 - flux - fix getvideo error for vidattr
' 048 - flux - added haschild check for vidattr, added options menu, reduced sling strength
' 049 - Gedankekojote97 - New optimized toolkit import.
' 050 - apophis - Toolkit script updates. Light object updates. Removed duplicate sound effects. Added RotateLaneLights on flipper press. Updated SetRoomBrightness code. Updated ball brightness code. Updated ball images. Removed DrainShooter2. Fixed option DMD text allignment issue.
' 051 - flux - Removed Lampz, Updated Light Controller to support vpmMapLights, Removed duplicate lights for reflections ( use main control instead)
' 052 - 055 oqqsan - dmd, fixed skillshot triggered on ballsave, added some animations on dmd
' 056 - oqqsan  adjusted some dmd calls, more time and moved a few text positions
' 058 - oqqsan - gameover sequence longer and working :) added gif on all modestarts
' 065 - oqqsan - enter highscore made - ballsave tweeked to blink faster and has a grace period of 2 seconds ( same length just the light goes off early )
' 069 - oqqsan - addons on dmd, smal fixes for standalone, fixed the looping music finaly
' 074 - oqqsan - fixes, addons made for attractmode , new inserts scripted but not Enabled
' 084 - Sixtoe - Tidied up layout and ramp layers, made GI lamps hidden, changed interval to blinkinterval
' 085 - oqqsan - adding more stuff :P
' 092 - Gedankekojote97 - Added new 4k bake
' 093 - apophis - Updated with new bake VLM scripts. Added PF reflections. Aded PF image for ball reflections. Made parts and PF static. Added phys rubber on left of scoop. Updated animations: flippers, switches, gates, spinner
' 094 - Gedankekojote97 - Added new 4k bake. Fixed flippers and other stuff.
' 095 - apophis - Updated with new bake and VLM scripts. Fixed left bottom flipper (had -1 scale). Readded the PF reflections. Added refractions on left ramp. Fixed some insert reflection colors.
' 096 - apophis - Fixed (?) lower left flipper meshes by mirroring
' 097 - flux - Fixed lightmap colors
' 098 - oqqsan -
' 099 - Gedankekojote97 - Added new 4k bake.
' 100 - apophis - Updated with new bake VLM scripts. Fixed post passes. Added new FlipperCradleCollision trick. Re-fixed: PF reflections, PF-ball reflections, ramp refractions.
' 103 - flux - actually fix lightmap colors, hook up new light controller
' 109 - oqqsan many unlisted dmd and game changes that didnt get put into this list :P lazy
' 110 - oqqsan added spongepirate movment   If B2son or B2SonDesktop = 1 or CabMode = 1 then startcontroller  ( added cabmode there )
' 116 - apophis - Re-fixed: Playfield reflections, ramp refractions. VLM arrays.
' 131+- oqqsan - adding features   WizardShake. still lazy so dont fill in the list :P
' 134 - flux - made vlm parts static rendering, put light controller on it's own timer at 20ms, no need for -1 timer
' 134b - MrGrynch - Removed pauses in looping music
' 135 - oqqsan - fixed bugs - added more dmd calls for comboEB - ball2+ restart hold startbutton 2 sec for restart. adjusted the ramp openings+ top of rightramp and outlane opening a tiny bit bigger
' 136-7 - oqqsan - updates and fixes + restart mb if no SJPgotten
' 138 - oqqsan - fixed stuff + added normalized mode tracks+intro+endings from MrGrynch
' 139 - MrGrynch - new code for saving/loading settings and high scores to files.  No more registry.  Files in VisualPinball/User folder.  Simplified high score check
' 140 - MrGrynch - Added ROM sound volume to options.  Removed VPReg storage.  Code cleanup
' 143 - oqqsan - added the last sfx normalizes sounds from mrgrynch :)  some fixes on scoop kicker, added eob bonus on finishing<- wizards - options work better now thanks mrgrynch :) added him to attract mode
' 148 - oqqsan - added optionDMD - reset highscore . fixed on load/save highscore ( the " " made trouble )  wiz2 test and bug fixws
' 149 - oqqsan - optionDMD now shown on the regular DMD !   :)
' 150 - oqqsan - optionDMD added som sparkle o.0
' 153 - flux - scorbit integration
' 154 - oqqsan - fixed angles behind pinapple so balls dont get so stuck :P
' 163 - forgot to write again many smal fixes tho :)    more sparkle to Options on the dmd :)
' 166 - apophis - Fixed MechVolume. Fix room brightness initiziation. Hooked up VR room to brightness code. Updated delay logic in table1_init code. Removed old LUT code. Merged table version variables. Set Light002 DB to -2000.
' 167 - flux - Testing scorbit qr scanning on desktop
' 168 - oqqsan - triggerwall and the 2 diverters animated . load save options without decimals. if vrooom>0 then controller stop ... BOB animation fixes ? mby this time it gets right ???
' 175 - oqqsan -  New playquotes from embee, mrgrych normalized them all incl a few i found
' 176 - apophis/Gedankekojote97 - New 4k batch imported. Ray traced shadows setup. Disabled old dyn BS and removed BS options from flex menu. Fixes for standalone. Changed Light002 name to f148. Re-fixed: Playfield reflections, PF-ball reflections, ramp refractions. VLM arrays.
' 177 - apophis - moved flippers vertically up by 5 vp units. Fixed flipper shadows.
 '179 - Gedankekojote97 - Added new 4k bake.
' 180 - apophis -  Hopefully fixed white flashes on PF. Re-fixed: Parts and PF set to static, Playfield reflections, PF-ball reflections, ramp refractions. VLM arrays. Flipper vertical positions. Deleted warmup primitives. Disable lightmap reflections.
' 181 - flux - fix GI assignments, added GI split gi lights into alllights
' 182 - Gedankekojote97
' 183 - oqqsan -
' 186 - flux - move scorbit qr codes
'Release 2.0
' 2.0.1 - Sixtoe - adjusted right ramp geometry to prevent stuck ball
' 2.0.2 - rothbauerw - Added ramp stuck ball fix
' 2.0.3 - Primetime5k - Staged Flipper support, menu option added
' 2.0.4 - apophis - Fixed Staged Flipper implementation and menu option. Scorbit.UploadLog=1 (thanks Manu). Fixed 15M point skillshot (thanks JasonDH). Added LoadEM to table1_init and fixed target bank vert position (thanks rothbauerw). Commented out startcontroller.
' 2.0.5 - oqqsan - frisco spelled correct :)  adjusted wizardscoring ( abit lower now )  wizard1done variable now reset every gamestart
' 2.0.5b- oqqsan - MB JP 20->50m .  Modes need 1 JP or timer runs out on modes before you can start new mode !
' 2.0.6 - oqqsan - hurryup jplvl+2 25m score . JP count in mode: 5=jplvl+1 and scoring   NO scoring in 3 seconds = timers Stop
' 2.0.7 - oqqsan - added more DOF calls , fixed eob bonus skip at tilt again (sneeky "Or" should be "And") messed up when added restartgame option way back when
' 2.0.8 - oqqsan - dof shaker call on plankton + wizjp instead of nudge : dmdcalls on timestopp : megaJP now on 3'rd jp on each mode : need 1jp before wiz mode avail : some other fixes aswell :)
' 2.0.9 - apophis - added extra dof calls
' 2.0.10 - oqqsan - bugfix for MB ( nasty one )  added credits+extras, option menu for freeplay, pineapple color on skillshot, highscore reset !
' 2.0.12 - flux - added scorbit mode updates
' 2.0.13 - oqqsan - moved the scorebit mode to another timer
' 2.0.14 - oqqsan - smal fixes
' 2.0.15 - oqqsan - 2 super smal changes
'Release 2.1

'DOF list almost at top of script
