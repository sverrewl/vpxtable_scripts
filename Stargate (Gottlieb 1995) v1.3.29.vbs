' Stargate (Gottlieb 1995)

' 1.3.1 - mcarter78 - Initial physics test (all visual objects hidden, no targets) -- Full physics overhaul, new physical trough & kickers, new playfield mesh, VUK ramps, new posts & rubbers, VPM mapped lights, roth targets added but not hooked up yet, new SSF (will add back in custom Stargate sounds where applicable)
' 1.3.6 - Sixtoe - Moved position of pop bumpers, rebuilt inlanes, rebuild left ramp, rebuilt right ramp, rebuilt right upper flipper area, reduced playfield friction, turned slope down to 6, messed around with flipper physics, touched most of the physics on the table. WIP.
' 1.3.7 - mcarter78 - Fix flashers to animate with L120-L125, clean up unused primitives/walls etc, add placeholder bg image, set initial state for slingshot animations, some code cleanup
' 1.3.8 - mcarter78 - Add options for Lighting colors, LUT, Art Blades,  & Cab Mode
' 1.3.9 - mcarter78 - Restore all solenoid sounds recorded from Stargate, animate rollovers & bumpers (WIP)
' 1.3.10 - mcarter78 - Add ramp triggers for rolling/drop sounds, add inlane speed limit code, tweak pyramid drop animation, slight flipper size adjustment to match visuals
' 1.3.12 - Sixtoe - Rebuilt subway, rebuilt playfield mesh and drop hole areas and deflectors / rubbers, added roof to ramps, turned down all kickers, adjusted sleeve materials, deleted old auto plunger, turned down impulse plunger, tweaked numerous rubbers, unified playfield walls.
' 1.3.13 - mcarter78 - Reduce flipper end angles & reshape triggers, animate autoplunger, make some plastics collidable, slightly increase kicking targets resistance spring strength, add walls at left ramp to prevent airballs
' 1.3.14 - apophis - Removed a bunch of old images. Removed old plastic primitives. Fixed collections for a few collidable posts. Adjusted URF physics. Removed flipper corrections from URF. Commented out upper flipper sfx for now. Refactored AnimateFlippers. Minor update to playfield_mesh. Minor change to SetInsertsColor. Adjusted bumper size and strength. Added new ball images. Glider sfx volume reduced in sound mgr. Fixed some ramp sfx. Fixed SarcArm animation. Fixed sling animations.
' 1.3.15 - Mecha_Enron - Updated render for GI split, fixed drop target UV to correct harsh line in middle, changed env tex to hospital room, added point lights above inserts
' 1.3.16 - mcarter78 - Add GI light color options, adjust upper flipper angle, add flipper for autoplunger animation, set slope to 8, adjustable down to 6, slightly increase vuk strength for guardian orbit vuks
' 1.3.17 - apophis - Updated vuk kickball code to be more robust. Fixed slingshot wall heights. Fixed GI reflection colors. Set bottom 6 GI lights to cast shadows (in anticipation for next bake).
' 1.3.18 - apophis - Fixed KTKick()
' 1.3.19 - apophis - Fixed KTHit()
' 1.3.20 - Sixtoe - Rebuilt slings (just in Case), tweaked flipper positions and triggers,
' 1.3.21 - tomate - Ball trap on the left wireramp fixed
' 1.3.22 - rothbauerw - Fixed kicker target code.
' 1.3.23 - apophis - Using star triggers for sw90, sw100, and sw110 now. Adjusted flipper polarity curves.
' 1.3.24 - Mecha_enron - Fixed pyramid animation
' 1.3.25 - Mecha_enron - Updated render to allow for day/night on ramp decals and correct transparancy of shooter ramp
' 1.3.26 - Mecha_enron - Added VLM.Bake.Pyramid to RoomBrightness array, fixed GI code, coded Cabinet Mode Option
' 1.3.27 - apophis - Fixed GI_002 raytraced ball shadows.
' 1.3.28 - mcarter78 - Add physical walls for Guardians, new DT backdrop
' 1.3.29 - mcarter78 - Fix autoplunger animation

'********************************************************************************************
'*  STARGATE VPX  from 32Assassin and JLouLoulou                    *
'*  All made from real Stargate scan, redraw etc etc                    *
'*                                              *
'*  With help from:                                     *
'*  Thalamus-JPJ-Fleep-Rothbaeur for SSF                          *
'*  DjRobx for Glidercraft script                             *
'*  Arngrim for DOF                                     *
'*  Thanks to them                                      *
'*                                              *
'*  Specials thanks to Fuzzel, Rascal, Chucky, Mat'for their advise, testing and feedback.  *
'*                                              *
'*  Thanks to Rothbauer and Fleep for TOM soundfiles and SSF script parts         *
'*                                              *
'*  Thanks to Chucky for making b2s backglass and long long testing             *
'*                                              *
'*  And thanks to 32Assassin to be patient with me, its the first time I work on VPX.   *
'*                                              *
'*  VR Conversion by Sixtoe                                 *
'*                                              *
'********************************************************************************************


Option Explicit
Randomize
SetLocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="stargat5"

'Map Switch Numbers
'Const swStartButton  =4
'Const swTournament =5
'Const swFrontDoor  =6
'Const BuyInButton  =3

Const BallSize = 50   'Ball size must be 50
Const BallMass = 1    'Ball mass must be 1
Const tnob = 4      'Total number of balls on the playfield including captive balls.
Const lob = 0     'Total number of locked balls

'  Standard definitions
Const UseVPMModSol = 2
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""      'Sound sample used for this, obsolete.
Const SSolenoidOff = ""     ' ^
Const SFlipperOn = ""     ' ^
Const SFlipperOff = ""      ' ^
Const SCoin = ""        '
Const UseGI = 1
Const swLCoin = 0
Const swRCoin = 1
Const swCCoin = 2
Const swCoinShuttle = 3
Const swStartButton = 4
Const swTournament = 5
Const swFrontDoor = 6

'Internal DMD in Desktop Mode, using a textbox (must be called before LoadVPM)
Dim UseVPMDMD, VRRoom, DesktopMode, VarHidden, UseVPMColoredDMD
DesktopMode = Table1.ShowDT
If RenderingMode = 2 OR TestVRonDT=True Then UseVPMDMD = True Else UseVPMDMD = DesktopMode
If DesktopMode = true then
    UseVPMColoredDMD = true
    VarHidden = 1
Else
    UseVPMColoredDMD = False
    VarHidden = 0
End If

LoadVPM "03060000", "GTS3.VBS", 3.10

' Globals
Dim TestVRonDT : TestVRonDT = False
Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim PlungerIM

'*************************************************************
'Script Options
'*************************************************************


TournamentTimer.Interval = 3000   'If you have "switch short return 5 / 6" message at ROM startup,
                'set it to 3000 and Increase it in steps of 100 untul it desapear
                'between each restart of VPX

' Thalamus 2020 April : Improved directional sounds

'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 1    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolTarg   = 1    ' Targets volume.
Const VolFlip   = 1    ' Flipper volume.
Const VolRH     = 1    ' Rubber hit volume.

Const RolVol    = 1    ' Rolling volume.
Const MroVol    = 3    ' Metal rolling volume.
Const MroBVol   = 20    ' Metal rolling non wires volume.
Const MModVol   = 40    ' Metal Modulation volume.
Const ProVol    = 5    ' Plastic rolling volume.
Const DroVol    = 1    ' Drop ball volume.



'******************************************************
'  ZTIM: Timers
'******************************************************

'The FrameTimer interval should be -1, so executes at the display frame rate
'The frame timer should be used to update anything visual, like some animations, shadows, etc.
'However, a lot of animations will be handled in their respective _animate subroutines.

Dim FrameTime, InitFrameTime
InitFrameTime = 0

FrameTimer.Interval = -1
FrameTimer.Enabled = true
Sub FrameTimer_Timer()
  FrameTime = gametime - InitFrameTime 'Calculate FrameTime as some animuations could use this
  InitFrameTime = gametime  'Count frametime
  'Add animation stuff here
  'BSUpdate
  UpdateBallBrightness
  RollingUpdate
  DoDTAnim
  UpdateDropTargets
  DoSTAnim
  UpdateStandupTargets
  DoKTAnim
  UpdateKickingTargets
  'ChangedSwitches 'debug code
  ' TimerPlunger
  ' TimerPlunger2
  AnimateAutoPlunger
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
CorTimer.Enabled = true
Sub CorTimer_Timer(): Cor.Update: End Sub

Sub L120_animate
  f0.state = L120.state
End Sub

Sub L121_animate
  f1.state = L121.state
End Sub

Sub L122_animate
  f2.state = L122.state
End Sub

Sub L123_animate
  f3.state = L123.state
End Sub

Sub L124_animate
  f4.state = L124.state
End Sub

Sub L125_animate
  f5.state = L125.state
End Sub

Sub AnimateAutoPlunger
  Dim BP
  For Each BP in BP_AutoPlunger: BP.transy = (APFlipper.startangle - APFlipper.currentangle) * 2 : Next
End Sub


'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************

Dim bsTrough, bsSaucer, dtT, dtL, dtR
Dim SBall1, SBall2, SBall3, SBall4, gBOT

Sub Table1_Init
  vpmInit Me

  vpmMapLights AllLamps

  With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Stargate" & vbNewLine & "VPW"
        '.Games(cGameName).Settings.Value("rol") = 0
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        '.Hidden = VarHidden
    .Hidden = 0
        .Dip(0) = (0 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 0 * 32 + 1 * 64 + 1 * 128) '01-08
        .Dip(1) = (0 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 1 * 32 + 1 * 64 + 1 * 128) '09-16
        .Dip(2) = (0 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 1 * 32 + 1 * 64 + 1 * 128) '17-24
        .Dip(3) = (1 * 1 + 1 * 2 + 1 * 4 + 0 * 8 + 1 * 16 + 0 * 32 + 1 * 64 + 1 * 128) '25-32
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch = swTilt
  vpmNudge.Sensitivity = 2
  vpmNudge.TiltObj = Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot)

  'Trough
  Set SBall1 = swTrough1.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SBall2 = swTrough2.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SBall3 = swTrough3.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SBall4 = swTrough4.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Controller.Switch(34) = 1
  Controller.Switch(24) = 1
  gBOT = Array(SBall1,SBall2,SBall3,SBall4)


  Const IMPowerSetting = 50
  Const IMTime = 0.6
  Set plungerIM = New cvpmImpulseP
  With plungerIM
    .InitImpulseP swPlunger, IMPowerSetting, IMTime
    .Random 0.3
    .CreateEvents "plungerIM"
  End With

  'Set autoplunger Position
  Dim BP
  For Each BP in BP_Autoplunger : BP.Y = 1948 : Next

  'GION = 1 ' Turn GI on at table load
  GIState False

  SetupRoom
  InitSlings

End Sub


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim VRRoomChoice: VRRoomChoice = 2
Dim ColorLUT: ColorLUT=0
Dim BackVis: BackVis=1

Dim VRPreview : VRPreview = False

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If
  Dim v
  Dim BP

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 0, 11, 1, 0, 0, _
    Array("Normal", "Vibrant", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 0 Then Table1.ColorGradeImage = "colorgradelut256x16_1to1"
  if ColorLUT = 1 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat100"
  if ColorLUT = 2 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat90"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat80"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat70"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat60"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat50"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat40"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat30"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat20"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat10"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat00"

  'GI Lamps Color
  v = Table1.Option("GI Lamps Color", 0, 7, 1, 0, 0, Array("Incandescent (warm white)", "LED (cool white)", "Purple", "Blue", "Red", "Green", "Yellow", "JLou Style"))
  SetGiColor v

  'Insert Lamps Color
  v = Table1.Option("Insert Lamps Color", 0, 1, 1, 0, 0, Array("Incandescent (warm white)", "LED (cool white)"))
  SetInsertsColor v

  'Flasher Lamps Color
  v = Table1.Option("Flasher Lamps Color", 0, 6, 1, 0, 0, Array("Incandescent (warm white)", "LED (cool white)", "Purple", "Blue", "Red", "Green", "Yellow"))
  SetFlashersColor v

  ' Art Blades
  v = Table1.Option("Art Blades", 0, 1, 1, 1, 0, Array("Off", "On"))
  For Each BP in BP_SideBlades_Art : BP.Visible = v: Next

    ' Toggle Rails
  If RenderingMode <> 2 And Not VRPreview Then
    v = Table1.Option("Cabinet Mode", 0, 1, 1, 1, 0, Array("On", "Off"))
    For Each BP in BP_Wood_guide___Painted_Blaxk_1 : BP.Visible = v: Next
    For Each BP in BP_PinCab_Rails_021 : BP.Visible = v: Next
    For Each BP in BP_PinCab_Rails_022 : BP.Visible = v: Next
    For Each BP in BP_PinCab_Rails_023 : BP.Visible = v: Next
  End If

  'VR Room
  'VRRoomChoice = Table1.Option("VR Room", 0, 2, 1, 2, 0, Array("Ultra Minimal", "Minimal", "Old Arcade"))
  'SetupRoom

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

'****************************
' ZGIC: GI Colors
'****************************

''''''''' GI Color Change options
' reference: https://andi-siess.de/rgb-to-color-temperature/

Dim c2700k: c2700k = rgb(255, 169, 87)
Dim c3000k: c3000k = rgb(255, 180, 107)
Dim c5000k: c5000k = rgb(255, 244, 226)

Dim cRedFull: cRedFull = rgb(255,0,0)
Dim cRed: cRed= rgb(255,5,5)
Dim cPinkFull: cPinkFull = rgb(255,0,225)
Dim cPink: cPink = rgb(255,5,255)
Dim cWhiteFull: cWhiteFull = rgb(255,255,128)
Dim cWhite: cWhite = rgb(255,255,255)
Dim cBlueFull: cBlueFull= rgb(0,0,255)
Dim cBlue: cBlue = rgb(5,5,255)
Dim cCyanFull: cCyanFull= rgb(0,255,255)
Dim cCyan : cCyan = rgb(5,128,255)
Dim cYellowFull:cYellowFull  = rgb(255,255,128)
Dim cYellow: cYellow = rgb(255,255,0)
Dim cOrangeFull: cOrangeFull = rgb(255,128,0)
Dim cOrange: cOrange = rgb(255,70,5)
Dim cGreenFull: cGreenFull = rgb(0,255,0)
Dim cGreen: cGreen = rgb(5,255,5)
Dim cPurpleFull: cPurpleFull = rgb(128,0,255)
Dim cPurple: cPurple = rgb(60,5,255)
Dim cAmberFull: cAmberFull = rgb(255,197,143)
Dim cAmber: cAmber = rgb(255,197,143)

Dim cArray
cArray = Array(c2700k,c5000k,cPurple,cBlue,cRed,cGreen,cYellow)
Dim jArray
jArray = Array(cRed,cBlue,cGreen,cYellow,cPurple)

sub SetGiColor(c)
  Dim xx, BL
  If c = 7 Then
    For each xx in GIR
      xx.color = cArray(4)
      xx.colorfull = cArray(4)
    Next
    For each xx in GIB
      xx.color = cArray(3)
      xx.colorfull = cArray(3)
    Next
    For each xx in GIG
      xx.color = cArray(5)
      xx.colorfull = cArray(5)
    Next
    For each xx in GIY
      xx.color = cArray(6)
      xx.colorfull = cArray(6)
    Next
    For each xx in GIP
      xx.color = cArray(2)
      xx.colorfull = cArray(2)
    Next

    For each BL in BL_GIR: BL.color = cArray(4): Next
    For each BL in BL_GIB: BL.color = cArray(3): Next
    For each BL in BL_GIG: BL.color = cArray(5): Next
    For each BL in BL_GIY: BL.color = cArray(6): Next
    For each BL in BL_GIP_GI_001: Bl.color = cArray(2): Next
    For each BL in BL_GIP_GI_002: Bl.color = cArray(2): Next
    For each BL in BL_GIP_GI_006: Bl.color = cArray(2): Next
    For each BL in BL_GIP_GI_007: Bl.color = cArray(2): Next
  Else
    For each xx in GI
      xx.color = cArray(c)
      xx.colorfull = cArray(c)
    Next
    For each BL in BL_GIR: BL.color = cArray(c): Next
    For each BL in BL_GIB: BL.color = cArray(c): Next
    For each BL in BL_GIG: BL.color = cArray(c): Next
    For each BL in BL_GIY: BL.color = cArray(c): Next
    For each BL in BL_GIP_GI_001: Bl.color = cArray(c): Next
    For each BL in BL_GIP_GI_002: Bl.color = cArray(c): Next
    For each BL in BL_GIP_GI_006: Bl.color = cArray(c): Next
    For each BL in BL_GIP_GI_007: Bl.color = cArray(c): Next
  End If
end Sub

sub SetInsertsColor(c)
  Dim xx, BL
' For each xx in AllLamps   'commenting this out so reflection on ball looks correct
'   xx.color = cArray(c)
'   xx.colorfull = cArray(c)
'   'Select Case c
'   ' Case 0: xx.fadespeedup = 120:: xx.fadespeeddown = 120
'   ' Case Else: xx.fadespeedup = 0: xx.fadespeeddown = 0
'   'End Select
' Next
  'gim001.color = cArray(0): gim001.colorfull = cArray(0)
  For each BL in BL_L_l5: BL.color = cArray(c): Next
  For each BL in BL_L_l0: BL.color = cArray(c): Next
  For each BL in BL_L_l11: BL.color = cArray(c): Next
  For each BL in BL_L_l12: BL.color = cArray(c): Next
  For each BL in BL_L_l13: BL.color = cArray(c): Next
  For each BL in BL_L_l14: BL.color = cArray(c): Next
  For each BL in BL_L_l16: BL.color = cArray(c): Next
  For each BL in BL_L_l17: BL.color = cArray(c): Next
  For each BL in BL_L_l21: BL.color = cArray(c): Next
  For each BL in BL_L_l22: BL.color = cArray(c): Next
  For each BL in BL_L_l23: BL.color = cArray(c): Next
  For each BL in BL_L_l24: BL.color = cArray(c): Next
  For each BL in BL_L_l25: BL.color = cArray(c): Next
  For each BL in BL_L_l26: BL.color = cArray(c): Next
  For each BL in BL_L_l27: BL.color = cArray(c): Next
  For each BL in BL_L_l32: BL.color = cArray(c): Next
  For each BL in BL_L_l33: BL.color = cArray(c): Next
  For each BL in BL_L_l34: BL.color = cArray(c): Next
  For each BL in BL_L_l35: BL.color = cArray(c): Next
  For each BL in BL_L_l36: BL.color = cArray(c): Next
  For each BL in BL_L_l37: BL.color = cArray(c): Next
  For each BL in BL_L_l41: BL.color = cArray(c): Next
  For each BL in BL_L_l42: BL.color = cArray(c): Next
  For each BL in BL_L_l43: BL.color = cArray(c): Next
  For each BL in BL_L_l44: BL.color = cArray(c): Next
  For each BL in BL_L_l45: BL.color = cArray(c): Next
  For each BL in BL_L_l46: BL.color = cArray(c): Next
  For each BL in BL_L_l47: BL.color = cArray(c): Next
  For each BL in BL_L_l55: BL.color = cArray(c): Next
  For each BL in BL_L_l56: BL.color = cArray(c): Next
  For each BL in BL_L_l57: BL.color = cArray(c): Next
  For each BL in BL_L_l6: BL.color = cArray(c): Next
  For each BL in BL_L_l65: BL.color = cArray(c): Next
  For each BL in BL_L_l66: BL.color = cArray(c): Next
  For each BL in BL_L_l67: BL.color = cArray(c): Next
  For each BL in BL_L_l7: BL.color = cArray(c): Next
  For each BL in BL_L_l71: BL.color = cArray(c): Next
  For each BL in BL_L_l72: BL.color = cArray(c): Next
  For each BL in BL_L_l73: BL.color = cArray(c): Next
  For each BL in BL_L_l74: BL.color = cArray(c): Next
  For each BL in BL_L_l75: BL.color = cArray(c): Next
  For each BL in BL_L_l76: BL.color = cArray(c): Next
  For each BL in BL_L_l77: BL.color = cArray(c): Next
  For each BL in BL_L_l81: BL.color = cArray(c): Next
  For each BL in BL_L_l82: BL.color = cArray(c): Next
  For each BL in BL_L_l83: BL.color = cArray(c): Next
  For each BL in BL_L_l84: BL.color = cArray(c): Next
  For each BL in BL_L_l85: BL.color = cArray(c): Next
  For each BL in BL_L_l86: BL.color = cArray(c): Next
  For each BL in BL_L_l87: BL.color = cArray(c): Next
end Sub

sub SetFlashersColor(c)
  Dim xx, BL
  For each xx in Flashers
    xx.color = cArray(c)
    xx.colorfull = cArray(c)
  Next
  For each BL in BL_F_f0: BL.color = cArray(c): Next
  For each BL in BL_F_f1: BL.color = cArray(c): Next
  For each BL in BL_F_f2: BL.color = cArray(c): Next
  For each BL in BL_F_f3: BL.color = cArray(c): Next
  For each BL in BL_F_f4: BL.color = cArray(c): Next
  For each BL in BL_F_f5: BL.color = cArray(c): Next
end Sub



'******************************************************
' ZKEY: Key Press Handling
'******************************************************

Sub table1_KeyDown(ByVal Keycode)
  If Keycode = LeftFlipperKey Then
    'MovButtonLeft.TransX = 8
    Controller.Switch(81) = 1
  End If
    If keycode = RightFlipperkey then
    'MovButtonRight.TransX = -8
    Controller.Switch(82) = 1
  End If
    'If LowerFlipper Then
    '    If keycode = RightFlipperkey then
    '        RightFlipper1.RotateToEnd
  '   UpdateTopper
    '    End If
    'End If
    If keycode = PlungerKey Then
    SoundPlungerPull
    Plunger.Pullback
    'TimerVRPlunger.Enabled = True
    'TimerVRPlunger1.Enabled = False
    'MovPlunger.TransY = 0
  End If
    If keycode = RightMagnaSave Then Controller.Switch(5) = 1  'Tournament Mode
    If keycode = LeftTiltKey Then Nudge 90, 2: SoundNudgeLeft
    If keycode = RightTiltKey Then Nudge 270, 2: SoundNudgeRight
    If keycode = CenterTiltKey Then Nudge 0, 2: SoundNudgeCenter
  If keycode = StartGameKey Then
    SoundStartButton
    'MovStart.TransY = -8
    'MovStart1.TransY = -8
  End If
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then SoundCoinIn
    If vpmKeyDown(keycode) Then Exit Sub
End Sub



Sub table1_KeyUp(ByVal Keycode)
  If Keycode = LeftFlipperKey Then
    Controller.Switch(81) = 0
    'MovButtonLeft.TransX = 0
  End If
  If Keycode = RightFlipperKey Then
    Controller.Switch(82) = 0
    RightFlipper1.RotatetoStart
    'MovButtonRight.TransX = 0
  End If
    If keycode = PlungerKey Then
    If Controller.Switch(31) = 1 Then
      SoundPlungerReleaseBall
    Else
      SoundPlungerReleaseNoBall
    End If
    Plunger.Fire
    'TimerVRPlunger.Enabled = False
    'TimerVRPlunger1.Enabled = True
  End If
    If keycode = RightMagnaSave Then Controller.Switch(5) = 0
  If keycode = StartGameKey Then
    'MovStart.TransY = 0
    'MovStart1.TransY = 0
  End If
    If vpmKeyUp(keycode) Then Exit Sub
End Sub


'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub TimerPlunger
  If VR_Primary_plunger.Y < 110 then
      VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2
 'debug.print plunger.position
  VR_Primary_plunger.Y = -25 + (5* Plunger.Position) -20
End Sub

'**********************************************************************************************************




'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
'PinMame has no Zero ID support add 1 from service manual

'SolCallback(1)=""  'Bumper 1
'SolCallback(2)=""  'Bumper 2
'SolCallback(3)=""  'Left SlingShot
'SolCallback(4)=""  'Right SlingShot
'SolCallback(5)=""  'sw14 kickback
'SolCallback(6)=""  'sw15 kickback
'SolCallback(7)=""  'sw16 kickback
SolCallback(8) = "SolLeftPlunge"
SolCallback(9) = "SolAutoFire"
SolCallback(10) = "LeftPop"
SolCallback(11) = "BotPop"
SolCallback(12) = "VukTopPop"
SolCallback(13) = "SolDiv"
SolCallback(14) = "SolPivL" 'Left Pivot Target
SolCallback(15) = "SolPivR" 'Right Pivot Target
SolCallback(16) = "SolPyramid" 'Pyramid Unit
SolCallback(17) = "LeftDropUp"
SolCallback(18) = "TopDropUp"
SolCallback(19) = "RightDropUp"
SolCallback(20) = "RightDropTrip"
'SolCallback(21)=""  'Not Used
'SolCallback(22)=""  'Rope Lights Backglass
SolCallback(23) = "SolGlid1" 'Left and Right Glide Motor
SolCallback(24) = "SolGlid2" 'Forward Glide motor
'SolCallback(25)=""  'Not Used
'SolCallback(26)="PFGI" 'BackBox GI used as PF GI
'SolCallback(27) = "" 'Ticket dispenser
SolCallback(28) = "SolRelease"
SolCallback(29) = "SolTrough"
SolCallback(30) = "SolKnocker"
SolCallback(31) = "GIState" 'Tilt Relay and PF GI
'SolCallback(32)=""  'Game Over Relay

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************


Sub SolDiv(Enabled)
  If Enabled Then
    Flipper1.RotateToEnd
    Controller.Switch(101)=1
    PlaySoundAtVol SoundFX("S13_LowerLeftBallGateOpen",DOFContactors), Flipper1, 1
  Else
    Flipper1.RotateToStart
    Controller.Switch(101)=0
    PlaySoundAtVol SoundFX("S13_LowerLeftBallGateClose",DOFContactors), Flipper1, 1
  End If
End Sub

Sub SolKnocker(Enabled)
  If Enabled Then
    PlaysoundAtVol "S30_Knocker", BackglassRelay, VolTarg
  End If
End Sub



'******************************************************
'  ZGIU:  GI Control
'******************************************************

dim gilvl

Sub GIState(Enabled)
  If Not Enabled Then
    gilvl = 1
    Sound_Flash_Relay 1 , Bumper1
  Else
    gilvl = 0
    Sound_Flash_Relay 0 , Bumper1
  End If
  Dim bulb: For Each bulb in GI: bulb.State = gilvl: Next
End Sub


Sub SolBG(Enabled)
If Not Enabled Then
  gibg.state = 1
  Sound_Gi_Relay 1 , BackglassRelay
Else
  gibg.state = 0
  Sound_Gi_Relay 0 , BackglassRelay
End If
End Sub


Dim GION



'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************

Sub swTrough1_UnHit : UpdateTrough : End Sub
Sub swTrough1_Hit   : UpdateTrough : End Sub
Sub swTrough2_UnHit : Controller.Switch(34) = 0 : UpdateTrough : End Sub
Sub swTrough2_Hit   : Controller.Switch(34) = 1 : UpdateTrough : End Sub
Sub swTrough3_UnHit : UpdateTrough : End Sub
Sub swTrough3_Hit   : UpdateTrough : End Sub
Sub swTrough4_UnHit : Controller.Switch(24) = 0 : UpdateTrough : CheckDrain : End Sub
Sub swTrough4_Hit   : UpdateTrough : End Sub
Sub Drain_UnHit  : UpdateTrough : End Sub
Sub Drain_Hit    : Controller.Switch(24) = 1 : UpdateTrough : PlaysoundAtVol "S29_Outhole", Drain, VolTarg : End Sub

Sub CheckDrain  'Hacky fix for trough issue
  if Drain.BallCntOver > 0 Then Controller.Switch(24) = 1
End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 100
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If swTrough1.BallCntOver = 0 Then swTrough2.kick 57, 10
  If swTrough2.BallCntOver = 0 Then swTrough3.kick 57, 10
  If swTrough3.BallCntOver = 0 Then swTrough4.kick 57, 10
  UpdateTroughTimer.Enabled = 0
End Sub

'*****************  DRAIN & RELEASE  ******************

Sub SolTrough(enabled)
  If enabled Then
    Drain.kick 57, 20
    SoundSaucerKick 1, Drain
  End If
End Sub

Sub SolRelease(enabled)
  If enabled Then
    swTrough1.kick 57, 10
    PlaysoundAtVol "S28_BallRelease", swTrough1, VolTarg
  End If
End Sub

Sub SolLeftPlunge(enabled)
  If enabled Then
    sw25.Kick 0, 40
    PlaysoundAtVol "S08_LowerLeftKicker", sw25, VolTarg
  End If
End Sub

 ' Drain hole and kickers
Sub sw25_Hit: Controller.Switch(25) = 1: RandomSoundMetal : End Sub
Sub sw25_UnHit: Controller.Switch(25) = 0 : End Sub



'************************************************************
' ZVUK: VUKs and Kickers
'************************************************************


'******************* VUKs **********************

Dim KickerBall33, KickerBall23, KickerBall80    'Each VUK needs its own "kickerball"

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)  'Defines how KickBall works
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

'Right VUK
Sub sw33_Hit                    'Switch associated with Kicker
    set KickerBall33 = activeball
    Controller.Switch(33) = 1
    SoundSaucerLock
  WireRampOn False
End Sub

Sub sw33_UnHit                    'Switch associated with Kicker
    KickerBall33 = Empty
    Controller.Switch(33) = 0
  WireRampOn False
End Sub

Sub VukTopPop(Enable)             'Solonoid name associated with kicker.
    If Enable then
    If Not IsEmpty(KickerBall33) Then
      KickBall KickerBall33, 0, 0, 40, 0
      SoundSaucerKick 1, sw33
    Else
      SoundSaucerKick 0, sw33
    End If
  End If
End Sub

'Center VUK
Sub sw23_Hit                    'Switch associated with Kicker
    set KickerBall23 = activeball
    Controller.Switch(23) = 1
    SoundSaucerLock
  WireRampOn False
End Sub

Sub sw23_UnHit                    'Switch associated with Kicker
    KickerBall23 = Empty
  Controller.Switch(23) = 0
  WireRampOn False
End Sub

Sub BotPop(Enable)              'Solonoid name associated with kicker.
    If Enable then
    If Not IsEmpty(KickerBall23) Then
      KickBall KickerBall23, 0, 0, 60, 0
      SoundSaucerKick 1, sw23
    Else
      SoundSaucerKick 0, sw23
    End If
  End If
End Sub

'Left VUK
Sub sw80_Hit
    set KickerBall80 = activeball
    Controller.Switch(80) = 1
    SoundSaucerLock
  WireRampOn False
End Sub

Sub sw80_UnHit
    KickerBall80 = Empty
    Controller.Switch(80) = 0
  WireRampOn False
End Sub

Sub LeftPop(Enable)
    If Enable then
    If Not IsEmpty(KickerBall80) Then
      KickBall KickerBall80, 0, 0, 60, 0
      SoundSaucerKick 1, sw80
    Else
      SoundSaucerKick 0, sw80
    End If
  End If
End Sub

'Impulse Plunger
Sub SolAutofire(Enabled)
  If Enabled Then
    APFlipper.RotateToEnd
    PlungerIM.AutoFire
    Plunger.timerenabled = 1
    PlaySoundAtVol SoundFX("S09_ShooterLaneKicker",DOFContactors), swTrough1, 1
  End If
End Sub

Sub Plunger_timer
  APFlipper.RotateToStart
  Plunger.timerenabled = 0
End Sub

'Bumpers
Sub Bumper1_Hit
  vpmTimer.PulseSw(10)
  PlaySoundAtVol SoundFX("S02_BumperBottom",DOFContactors), ActiveBall, VolBump
    me.timerenabled = 1

  Dim BP
  For Each BP in BP_Bumper1_Skirt
    BP.roty=skirtAY(me,Activeball)
    BP.rotx=skirtAX(me,Activeball)
  Next

  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub Bumper1_timer
  Dim BP
  For Each BP in BP_Bumper1_Skirt
    BP.roty=0
    BP.rotx=0
  Next
  me.timerenabled=0
end sub

Sub Bumper1_Animate
  Dim z: z = Bumper1.CurrentRingOffset

  Dim BP
  For Each BP in BP_Bumper1_Ring : BP.transz = z : Next
End Sub

Sub Bumper2_Hit
  vpmTimer.PulseSw(11)
  PlaySoundAtVol SoundFX("S01_Bumper_Top",DOFContactors), ActiveBall, VolBump
    me.timerenabled = 1

  Dim BP
  For Each BP in BP_Bumper2_Skirt
    BP.roty=skirtAY(me,Activeball)
    BP.rotx=skirtAX(me,Activeball)
  Next

  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub Bumper2_timer
  Dim BP
  For Each BP in BP_Bumper2_Skirt
    BP.roty=0
    BP.rotx=0
  Next
  me.timerenabled=0
end sub

Sub Bumper2_Animate
  Dim z: z = Bumper2.CurrentRingOffset

  Dim BP
  For Each BP in BP_Bumper2_Ring : BP.transz = z : Next
End Sub



'******************************************************
'     SKIRT ANIMATION FUNCTIONS
'******************************************************
' NOTE: set bumper object timer to around 150-175 in order to be able
' to actually see the animaation, adjust to your liking

'Const PI = 3.1415926
Const SkirtTilt=5   'angle of skirt tilting in degrees

Function SkirtAX(bumper, bumperball)
  skirtAX=cos(skirtA(bumper,bumperball))*(SkirtTilt)    'x component of angle
  if (bumper.y<bumperball.y) then skirtAX=skirtAX*-1    'adjust for ball hit bottom half
End Function

Function SkirtAY(bumper, bumperball)
  skirtAY=sin(skirtA(bumper,bumperball))*(SkirtTilt)    'y component of angle
  if (bumper.x>bumperball.x) then skirtAY=skirtAY*-1    'adjust for ball hit left half
End Function

Function SkirtA(bumper, bumperball)
  dim hitx, hity, dx, dy
  hitx=bumperball.x
  hity=bumperball.y

  dy=Abs(hity-bumper.y)         'y offset ball at hit to center of bumper
  if dy=0 then dy=0.0000001
  dx=Abs(hitx-bumper.x)         'x offset ball at hit to center of bumper
  skirtA=(atn(dx/dy)) '/(PI/180)      'angle in radians to ball from center of Bumper1
End Function



'StandUp Target
Sub sw22_Hit:vpmTimer.PulseSw 22:sw32p.TransX = -10  :Me.Timerenabled = 1 : PlaySoundAtVol SoundFX("Target",DOFTargets) , ActiveBall, VolTarg: End Sub
Sub sw22_Timer: sw32p.TransX = 0 : Me.Timerenabled = 0 : End Sub
'Sub sw32_Hit:vpmTimer.PulseSw 32:sw32p.TransX = -10  :Me.Timerenabled = 1 : PlaySoundAtVol SoundFX("Target",DOFTargets) , ActiveBall, VolTarg: End Sub
'Sub sw32_Timer: sw32p.TransX = 0 : Me.Timerenabled = 0 : End Sub

'Wire Triggers
Sub SW31_Hit:Controller.Switch(31)=1:AnimateWire BP_sw31, 1:End Sub
Sub SW31_unHit:Controller.Switch(31)=0:AnimateWire BP_sw31, 0:End Sub
Sub SW111_Hit:Controller.Switch(111)=1:AnimateWire BP_sw111, 1:End Sub
Sub SW111_unHit:Controller.Switch(111)=0:AnimateWire BP_sw111, 0:End Sub
Sub SW112_Hit:Controller.Switch(112)=1:AnimateWire BP_sw112, 1:leftInlaneSpeedLimit:End Sub
Sub SW112_unHit:Controller.Switch(112)=0:AnimateWire BP_sw112, 0:End Sub
Sub SW113_Hit:Controller.Switch(113)=1:AnimateWire BP_sw113, 1:rightInlaneSpeedLimit:End Sub
Sub SW113_unHit:Controller.Switch(113)=0:AnimateWire BP_sw113, 0:End Sub
Sub SW114_Hit:Controller.Switch(114)=1:AnimateWire BP_sw114, 1:End Sub
Sub SW114_unHit:Controller.Switch(114)=0:AnimateWire BP_sw114, 0:End Sub
Sub SW115_Hit:Controller.Switch(115)=1:AnimateWire BP_sw115, 1:End Sub
Sub SW115_unHit:Controller.Switch(115)=0:AnimateWire BP_sw115, 0:End Sub

Sub AnimateWire(group, action) ' Action = 1 - to drop, 0 to raise)
  Dim BP
  If action = 1 Then
    For Each BP in group : BP.transz = -13 : Next
  Else
    For Each BP in group : BP.transz = 0 : Next
  End If
End Sub


'****************************
' ZINL: Inlane speed limit code
'****************************

Sub leftInlaneSpeedLimit
  'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
  If activeball.vely < 0 Then Exit Sub              'don't affect upwards movement
    activeball.AngMomZ = -abs(activeball.AngMomZ) * RndNum(3,6)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

    If abs(activeball.vely) > 5 Then activeball.vely = 0.8 * activeball.vely
    If abs(activeball.vely) > 10 Then activeball.vely = 0.8 * activeball.vely
    If abs(activeball.vely) > 15 Then activeball.vely = 0.8 * activeball.vely
    If activeball.vely > 16 Then activeball.vely = RndNum(14,16)
    If activeball.vely < -16 Then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub


Sub rightInlaneSpeedLimit
  'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
  If activeball.vely < 0 Then Exit Sub              'don't affect upwards movement

    activeball.AngMomZ = abs(activeball.AngMomZ) * RndNum(2,4)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

  If abs(activeball.vely) > 5 Then activeball.vely = 0.8 * activeball.vely
    If abs(activeball.vely) > 10 Then activeball.vely = 0.8 * activeball.vely
    If abs(activeball.vely) > 15 Then activeball.vely = 0.8 * activeball.vely
    If activeball.vely > 16 Then activeball.vely = RndNum(14,16)
    If activeball.vely < -16 Then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub


'Ramp Triggers
Sub sw90_Hit :vpmTimer.PulseSw 90:End Sub
Sub sw100_Hit :vpmTimer.PulseSw 100:End Sub
Sub sw110_Hit :vpmTimer.PulseSw 110:End Sub

Sub RampTrigger001_Hit
  WireRampOff
  RandomSoundRampStop RampTrigger001
  RandomSoundTargetHitStrong
  WireRampOn True
End Sub

Sub RampTrigger002_Hit
  WireRampOff
  RandomSoundRubberWeak
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub RampTrigger003_Hit
  WireRampOff
  WireRampOn False
End Sub

Sub RampTrigger004_Hit
  RandomSoundMetal
  WireRampOn True
End Sub

Sub RampTrigger005_Hit
  WireRampOff
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub RampTrigger006_Hit
  WireRampOff
  RandomSoundRubberWeak
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub RampTrigger007_Hit
  WireRampOff
  RandomSoundRubberWeak
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub RampTrigger008_Hit
  RandomSoundMetal
  WireRampOn False
End Sub

Sub RampTrigger009_Hit
  RandomSoundMetal
  WireRampOn False
End Sub

Sub RampTrigger010_Hit
  WireRampOff
  WireRampOn True
End Sub

Sub RampTrigger011_Hit
  WireRampOff
  WireRampOn False
End Sub



'Pyramid switch subway
Sub sw91_Hit
  vpmTimer.PulseSw 91
  PlaySoundAtVol"popperball_pyramide" , ActiveBall, 1
  WireRampOn False
End Sub


'***************************************************
'      Left Guardian
'***************************************************

'Sub sw116_Hit : vpmTimer.PulseSw 116 : PlaySoundAtVol SoundFX("Target",DOFTargets) , ActiveBall, VolTarg: End Sub

Sub SolPivL(Enabled)
  if enabled then
    sw116.collidable=false
    sw116a.collidable=false
        LeftFlipper2.RotateToEnd
    PlaySoundAtVol SoundFX("S14_LeftPivotTargetOpen",DOFContactors), BM_GuardianL, 1
  else
    sw116.collidable=true
    sw116a.collidable=true
        LeftFlipper2.RotateToStart
    PlaySoundAtVol SoundFX("S14_LeftPivotTargetClose",DOFContactors), BM_GuardianL, 1
  end if
End Sub

'***************************************************
'      Right Guardian
'***************************************************

'Sub sw117_Hit : vpmTimer.PulseSw 117 : PlaySoundAtVol SoundFX("Target",DOFTargets) , ActiveBall, 1: End Sub

Sub SolPivR(Enabled)
  if enabled then
    sw117.collidable=false
    sw117a.collidable=false
        RightFlipper2.RotateToEnd
    PlaySoundAtVol SoundFX("S15_RightPivotTargetOpen",DOFContactors), BM_GuardianR, 1
  else
    sw117.collidable=true
    sw117a.collidable=true
        RightFlipper2.RotateToStart
    PlaySoundAtVol SoundFX("S15_RightPivotTargetClose",DOFContactors), BM_GuardianR, 1
  end if
End Sub

'***************************************************
'      Moving Pyramid
'***************************************************

Dim PyramidOpen:PyramidOpen = False

Sub SolPyramid(Enabled)
     If Enabled Then
        LeftFlipper1.RotateToEnd
    PlaySoundAtVol SoundFX("S16_TopPyramidOpen",DOFContactors), BM_Glider_1, 1
    PyramidOpen = -1
     Else
        LeftFlipper1.RotateToStart
    PlaySoundAtVol SoundFX("S16_TopPyramidClose",DOFContactors), BM_Glider_1, 1
    PyramidOpen = 0
     End If
  vpmTimer.AddTimer 500, "Controller.Switch(102)=" & PyramidOpen & "'"
  'Controller.Switch(102)=PyramidOpen
  'Debug.Print "Pyramid " & PyramidOpen
  End Sub


'***************************************************
'      Glider
'***************************************************


Dim GliderForwardSpeed:GliderForwardSpeed = 3 / GliderTimer.Interval
Dim GliderLRMotorSpeed:GliderLRMotorSpeed = 3 / GliderTimer.Interval
Dim GliderForwardDir:GliderForwardDir = 1
Dim GliderLROn:GliderLROn = False
Dim GliderFROn:GliderFROn = False

Dim GliderYMin:GliderYMin = 0
Dim GliderYMax:GliderYMax = 138
Dim GliderDist:GliderDist = GliderYMax - GliderYMin
Dim GliderLRMotor:GliderLRMotor = 0
Dim GliderRotMax:GliderRotMax = 20
Dim GliderY:GliderY = GliderYMin
Dim GliderSoundFR:GliderSoundFR = False
Dim GliderSoundLR:GliderSoundLR = False

Sub GliderTimer_Timer
  if GliderFROn then
    if not GliderSoundFR then
      PlaySound SoundFX("Glidercraft_retract",DOFGear), -1, 1, AudioPan(sw91), 0, 0, 1, 0, AudioFade(sw91)
      GliderSoundFR = true
    end if
    if GliderForwardDir > 0 then
      if GliderY < GliderYMax then
        GliderY = GliderY + GliderForwardSpeed
      Else
        GliderForwardDir = -1
      end if
    elseif GliderForwardDir < 0 then
      if GliderY > GliderYMin then
        GliderY = GliderY - GliderForwardSpeed
      Else
        GliderForwardDir = 1
      end if
    end if
  Else
    if GliderSoundFR then
      StopSound("Glidercraft_retract")
      GliderSoundFR = False
    end if
  end if
  if GliderLROn Then
    if not GliderSoundLR then
      PlaySound SoundFX("Glidercraft",DOFGear), -1, 1, AudioPan(sw91), 0, 0, 1, 0, AudioFade(sw91)
      GliderSoundLR = True
    end if
    GliderLRMotor = GliderLRMotor + GliderLRMotorSpeed
    if GliderLRMotor > 360 then GliderLRMotor = GliderLRMotor - 360
  else
    if GliderSoundLR then
      StopSound("Glidercraft")
      GliderSoundLR = false
    end if
  end if

  Dim GliderRot:GliderRot = Sin(GliderLRMotor * 3.14159 / 180) * GliderRotMax

  if GliderY <= GliderYMin + 2 then
    Controller.Switch(21) = 1
  Else
    Controller.Switch(21) = 0
  end if

  if GliderRot > GliderRotMax - 1 then
    Controller.Switch(30) = 1
  Else
    Controller.Switch(30) = 0
  end if

  Dim BP
  For each BP in BP_Glider_1
        BP.Transy = GliderY
        BP.RotZ = -GliderRot
    BP.visible = (GliderY < 30)
    Next
  For each BP in BP_Glider_2
        BP.Transy = GliderY
        BP.RotZ = -GliderRot
    BP.visible = (GliderY < 45)
    Next

  For each BP in BP_Glider_1_001
        BP.Transy = -48 + GliderY
        BP.RotZ = -GliderRot-2
    BP.visible = (GliderY < 60 And GliderY >= 30)
    Next
  For each BP in BP_Glider_2_001
        BP.Transy = -100 + GliderY
        BP.RotZ = -GliderRot
    BP.visible = (GliderY >= 45)
    Next
  For each BP in BP_Glider_1_002
        BP.Transy = -100 + GliderY
        BP.RotZ = -GliderRot-2
    BP.visible = (GliderY >= 60)
    Next

  if not GliderLROn and not GliderFROn then me.Enabled = 0
End Sub

'sw21 Glider Motor Stop
'sw30 Glider Right Motor stop

Sub SolGlid1(Enabled) 'Left and Right Glide Motor
  GliderLROn = Enabled
  GliderTimer.Enabled = 1
End Sub

Sub SolGlid2(Enabled) 'Forward Glide motor
  GliderFROn = Enabled
  GliderTimer.Enabled = 1
End Sub


'************************************************************
' ZSWI: SWITCHES
'************************************************************

'********************** Drop Targets ************************
Sub sw17_Hit: DTHit 17: TargetBouncer Activeball, 0.5: End Sub
Sub sw27_Hit: DTHit 27: TargetBouncer Activeball, 0.5: End Sub
Sub sw37_Hit: DTHit 37: TargetBouncer Activeball, 0.5: End Sub


Sub sw26_Hit: DTHit 26: TargetBouncer Activeball, 0.5: End Sub
Sub sw36_Hit: DTHit 36: TargetBouncer Activeball, 0.5: End Sub

Sub sw35_Hit: DTHit 35: TargetBouncer Activeball, 0.5: End Sub

Sub sw32_Hit: STHit 32: TargetBouncer Activeball, 0.5: End Sub
Sub sw116_Hit: STHit 116: TargetBouncer Activeball, 0.5: End Sub
Sub sw117_Hit: STHit 117: TargetBouncer Activeball, 0.5: End Sub


Sub LeftDropUp(enabled)
  'debug.print "SolResetDropsL "&enabled
  if enabled then
    PlaysoundAtVol "S17_3BankDropTargetReset", BM_DT_sw17, VolTarg
    DTRaise 17
    DTRaise 27
    DTRaise 37
  end if
End Sub

Sub TopDropUp(enabled)
  'debug.print "SolResetDropsL "&enabled
  if enabled then
    PlaysoundAtVol "S18_2BankDropTargetReset", BM_DT_sw26, VolTarg
    DTRaise 26
    DTRaise 36
  end if
End Sub

Sub RightDropUp(enabled)
  'debug.print "SolResetDropsL "&enabled
  if enabled then
    PlaysoundAtVol "S19_RollOverTargetReset", BM_DT_sw35, VolTarg
    DTRaise 35
  end if
End Sub

Sub RightDropTrip(enabled)
  'debug.print "SolResetDropsL "&enabled
  if enabled then
    PlaysoundAtVol "S20_RollOverTargetTrip", BM_DT_sw35, VolTarg
    DTDrop 35
  end if
End Sub

'******************************************************
'  ZANI: Misc Animations
'******************************************************

'--------flippers-----------
Sub LeftFlipper_Animate
  Dim va, BP
    Dim a : a = LeftFlipper.CurrentAngle
    FlipperLSh.RotZ = a

    va = 255.0 * (123.5 - LeftFlipper.CurrentAngle) / (123.5 -  67)

    For each BP in BP_LF
        BP.Rotz = a
        BP.visible = va < 128.0
    Next
    For each BP in BP_LFU
        BP.Rotz = a
        BP.visible = va >= 128.0
    Next
End Sub

Sub RightFlipper_Animate
  Dim vb, BP
  Dim b : b = RightFlipper.CurrentAngle
    FlipperRSh.RotZ = b

    vb = 255.0 * (-123.5 - RightFlipper.CurrentAngle) / (-123.5 +  67)

    For each BP in BP_RF
        BP.Rotz = b
        BP.visible = vb < 128.0
    Next
    For each BP in BP_RFU
        BP.Rotz = b
        BP.visible = vb >= 128.0
    Next
End Sub

Sub RightFlipper1_Animate
  Dim vc, BP
  Dim c : c = RightFlipper1.CurrentAngle
    FlipperRSh1.RotZ = c

    vc = 255.0 * (212 - RightFlipper1.CurrentAngle) / (212 +  -268.5)

    For each BP in BP_RF1
        BP.Rotz = c
        BP.visible = vc < 128.0
    Next
    For each BP in BP_RF1U
        BP.Rotz = c
        BP.visible = vc >= 128.0
    Next
End Sub

'--------pyramid-----------
Sub LeftFlipper1_Animate
    Dim a : a = LeftFlipper1.CurrentAngle

    Dim v, BP
    v = 255.0 * (120 - LeftFlipper1.CurrentAngle) / (120 + 0)

    For each BP in BP_Pyramid1
        BP.Transz = a
        BP.visible = a < 75
    Next
    For each BP in BP_Pyramid1_001
        BP.Transz = a - 120
        BP.visible = a >= 75
    Next

  SetPyrBrightness

End Sub

Dim PyramidMtlArray: PyramidMtlArray = Array("VLM.Bake.Pyramid")

Sub SetPyrBrightness

  ' Lighting level
  Dim v: v=((-1/90) * LeftFlipper1.CurrentAngle + 1) * (NightDay/100)
  'debug.print v

  Dim i: For i = 0 to UBound(PyramidMtlArray)
    ModulateMaterialBaseColorPyr PyramidMtlArray(i), i, v
  Next
End Sub

Dim SavedMtlColorArrayPyr
SaveMtlColorsPyr
Sub SaveMtlColorsPyr
  ReDim SavedMtlColorArrayPyr(UBound(PyramidMtlArray))
  Dim i: For i = 0 to UBound(PyramidMtlArray)
    SaveMaterialBaseColorPyr PyramidMtlArray(i), i
  Next
End Sub

Sub SaveMaterialBaseColorPyr(name, idx)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  SavedMtlColorArrayPyr(idx) = round(base,0)
End Sub


Sub ModulateMaterialBaseColorPyr(name, idx, val)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  Dim red, green, blue, saved_base, new_base

  'First get the existing material properties
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

  'Get saved color
  saved_base = SavedMtlColorArrayPyr(idx)

  'Next extract the r,g,b values from the base color
  red = saved_base And &HFF
  green = (saved_base \ &H100) And &HFF
  blue = (saved_base \ &H10000) And &HFF
  'msgbox red & " " & green & " " & blue

  'Create new color scaled down by 'val', and update the material
  new_base = RGB(red*val, green*val, blue*val)
  UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, new_base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub

'--------ball gate-----------
Sub Flipper1_Animate
    Dim a : a = Flipper1.CurrentAngle

    Dim v, BP
    v = 255.0 * (-55 - Flipper1.CurrentAngle) / (-55 + 14)

    For each BP in BP_SarcArm
        BP.Rotz = a
        BP.visible = v < 128
    Next
    For each BP in BP_SarcArmU
        BP.Rotz = a
        BP.visible = v >= 128
    Next
End Sub

'--------guardians-----------
Sub LeftFlipper2_Animate
    Dim a : a = LeftFlipper2.CurrentAngle

    Dim v, BP
    v = 255.0 * (60 - LeftFlipper2.CurrentAngle) / (0 +  60)

    For each BP in BP_GuardianL
        BP.Rotx = a+60
        BP.Visible = (a < -30)
    Next
    For each BP in BP_GuardianL_001
        BP.Rotx = a
        BP.Visible = (a >= -30)
    Next
End Sub

Sub RightFlipper2_Animate
    Dim a : a = RightFlipper2.CurrentAngle

    Dim v, BP
    v = 255.0 * (60 - RightFlipper2.CurrentAngle) / (0 +  60)

    For each BP in BP_GuardianR
        BP.Rotx = a+60
        BP.Visible = (a < -30)
    Next
    For each BP in BP_GuardianR_001
        BP.Rotx = a
        BP.Visible = (a >= -30)
    Next
End Sub


'******************************************************
' ZFLP: FLIPPERS
'******************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sURFlipper) = "SolURFlipper"

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    FlipperActivate RightFlipper, RFPress
    RF.Fire
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolURFlipper(Enabled)
  If Enabled Then
    FlipperActivate RightFlipper1, URFPress
    RightFlipper1.RotateToEnd
'   If rightflipper1.currentangle < rightflipper1.endangle + ReflipAngle Then
'     RandomSoundReflipUpRight RightFlipper1
'   Else
'     SoundFlipperUpAttackRight RightFlipper1
'     RandomSoundFlipperUpRight RightFlipper1
'   End If
  Else
    FlipperDeActivate RightFlipper1, URFPress
    RightFlipper1.RotateToStart
'   If RightFlipper1.currentangle < RightFlipper1.startAngle - 5 Then
'     RandomSoundFlipperDownRight RightFlipper1
'   End If
'   FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub


'Flipper collide subs
Sub RightFlipper1_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper1, URFCount, parm
  RightFlipperCollide parm
End Sub

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


'************************************************************
' ZSLG: Slingshot Animations
'************************************************************

Sub InitSlings
  Dim BP
  For Each BP in BP_RSling1 : BP.Visible = 0: Next
  For Each BP in BP_RSling2 : BP.Visible = 0: Next
  For Each BP in BP_LSling1 : BP.Visible = 0: Next
  For Each BP in BP_LSling2 : BP.Visible = 0: Next
End Sub


Dim LStep : LStep = 0 : LeftSlingShot.TimerEnabled = 1
Dim RStep : RStep = 0 : RightSlingShot.TimerEnabled = 1

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(13)            'Sling Switch Number
  RStep = 0
  RightSlingShot_Timer
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 17
  PlaysoundAtVol "S04_SlingshotRight", BM_SlingR, VolTarg
End Sub

Sub RightSlingShot_Timer
  Dim BL
  Dim x1, x2, y: x1 = True:x2 = False:y = -20
    Select Case RStep
        Case 2:x1 = False:x2 = True:y = -10
        Case 3:x1 = False:x2 = False:y = 0:RightSlingShot.TimerEnabled = 0
    End Select

  For Each BL in BP_RSling1 : BL.Visible = x1: Next
  For Each BL in BP_RSling2 : BL.Visible = x2: Next
  For Each BL in BP_SlingR : BL.TransX = y: Next

    RStep = RStep + 1
End Sub


Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(12)            'Sling Switch Number
  LStep = 0
  LeftSlingShot_Timer
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 17
  PlaysoundAtVol "S03_SlingshotLeft", BM_SlingL, VolTarg
End Sub

Sub LeftSlingShot_Timer
  Dim BL
  Dim x1, x2, y: x1 = True:x2 = False:y = -20
    Select Case LStep
        Case 3:x1 = False:x2 = True:y = -10
        Case 4:x1 = False:x2 = False:y = 0:LeftSlingShot.TimerEnabled = 0
    End Select

  For Each BL in BP_LSling1 : BL.Visible = x1: Next
  For Each BL in BP_LSling2 : BL.Visible = x2: Next
  For Each BL in BP_SlingL : BL.TransX = -y: Next

    LStep = LStep + 1
End Sub





'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS: Set LS = New SlingshotCorrection
Dim RS: Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00, - 5
  AddSlingsPt 1, 0.45, - 7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  5
End Sub

Sub AddSlingsPt(idx, aX, aY)    'debugger wrapper for adjusting flipper script In-game
  Dim a
  a = Array(LS, RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub

Class SlingshotCorrection
  Public DebugOn, Enabled
  Private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut

  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
    Enabled = True
  End Sub

  Public Property Let Object(aInput)
    Set Slingshot = aInput
  End Property

  Public Property Let EndPoint1(aInput)
    SlingX1 = aInput.x
    SlingY1 = aInput.y
  End Property

  Public Property Let EndPoint2(aInput)
    SlingX2 = aInput.x
    SlingY2 = aInput.y
  End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
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


  Public Sub VelocityCorrect(aBall)
    Dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1
      YL = SlingY1
      XR = SlingX2
      YR = SlingY2
    Else
      XL = SlingX2
      YL = SlingY2
      XR = SlingX1
      YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If Abs(XR - XL) > Abs(YR - YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If Not IsEmpty(ModIn(0) ) Then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      '   debug.print " BallPos=" & BallPos &" Angle=" & Angle
      '   debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled Then aBall.Velx = RotVxVy(0)
      If Enabled Then aBall.Vely = RotVxVy(1)
      '   debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      '   debug.print " "
    End If
  End Sub
End Class


'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity


'*******************************************
' Early 90's and after

Sub InitPolarity()
  Dim x, a
  a = Array(LF, RF)
  For Each x In a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 5.5
    x.AddPt "Polarity", 2, 0.16, - 5.5
    x.AddPt "Polarity", 3, 0.20, - 0.75
    x.AddPt "Polarity", 4, 0.25, - 1.25
    x.AddPt "Polarity", 5, 0.3, - 1.75
    x.AddPt "Polarity", 6, 0.4, - 3.5
    x.AddPt "Polarity", 7, 0.5, - 5.25
    x.AddPt "Polarity", 8, 0.7, - 4.0
    x.AddPt "Polarity", 9, 0.75, - 3.5
    x.AddPt "Polarity", 10, 0.8, - 3.0
    x.AddPt "Polarity", 11, 0.85, - 2.5
    x.AddPt "Polarity", 12, 0.9, - 2.0
    x.AddPt "Polarity", 13, 0.95, - 1.5
    x.AddPt "Polarity", 14, 1, - 1.0
    x.AddPt "Polarity", 15, 1.05, -0.5
    x.AddPt "Polarity", 16, 1.1, 0
    x.AddPt "Polarity", 17, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.23, 0.85
    x.AddPt "Velocity", 2, 0.35, 0.88
    x.AddPt "Velocity", 3, 0.45, 1
    x.AddPt "Velocity", 4, 0.6, 1 '0.982
    x.AddPt "Velocity", 5, 0.62, 1.0
    x.AddPt "Velocity", 6, 0.702, 0.968
    x.AddPt "Velocity", 7, 0.95,  0.968
    x.AddPt "Velocity", 8, 1.03,  0.945
    x.AddPt "Velocity", 9, 1.5,  0.945

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
  FlipperTricks RightFlipper1, URFPress, URFCount, URFEndAngle, URFState
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

Function Distance2Obj(obj1, obj2)
  Distance2Obj = SQR((obj1.x - obj2.x)^2 + (obj1.y - obj2.y)^2)
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

Dim LFPress, RFPress, LFCount, RFCount, URFPress, URFCount
Dim LFState, RFState, URFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle, URFEndAngle


Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

URFState = 1
LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
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
'Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
Const EOSReturn = 0.035  'mid 80's to early 90's
'Const EOSReturn = 0.025  'mid 90's and later

URFEndAngle = Rightflipper1.endangle
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
Sub NoTargetBounce_Hit
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


'******************************************************
' ZRST: STAND-UP TARGET INITIALIZATION
'******************************************************


Class StandupTarget
  Private m_primary, m_prim, m_sw, m_animate

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public default Function init(primary, prim, sw, animate)
    Set m_primary = primary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate

    Set Init = Me
  End Function
End Class


'Define a variable for each stand-up target
Dim ST32, ST116, ST117

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:  vp target to determine target hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          transy must be used to offset the target animation
'   switch:  ROM switch number
'   animate:  Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST32 = (new StandupTarget)(sw32, BM_ST_sw32, 32, 0)
Set ST116 = (new StandupTarget)(sw116, BM_GuardianL, 116, 0)
Set ST117 = (new StandupTarget)(sw117, BM_GuardianR, 117, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST32, ST116, ST117)

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.2    'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
  Dim i
  i = STArrayID(switch)

  PlayTargetSound
  STArray(i).animate = STCheckHit(ActiveBall,STArray(i).primary)

  If STArray(i).animate <> 0 Then
    DTBallPhysics ActiveBall, STArray(i).primary.orientation, STMass
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
    primary.uservalue = GameTime
  End If

  animtime = GameTime - primary.uservalue

  If animate = 1 Then
    primary.collidable = 0
    prim.transy =  - STMaxOffset
    vpmTimer.PulseSw switch mod 100
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


Sub UpdateStandupTargets
  dim BP, ty

    ty = BM_ST_sw32.transy
  For each BP in BP_ST_sw32 : BP.transy = ty: Next

    ty = BM_GuardianL.transy
  For each BP in BP_GuardianL : BP.transy = ty: Next

    ty = BM_GuardianR.transy
  For each BP in BP_GuardianR : BP.transy = ty: Next


End Sub

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
Dim DT17, DT27, DT37, DT26, DT36, DT35

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

Set DT17 = (new DropTarget)(sw17, sw17a, BM_DT_sw17, 17, 0, false)
Set DT27 = (new DropTarget)(sw27, sw27a, BM_DT_sw27, 27, 0, false)
Set DT37 = (new DropTarget)(sw37, sw37a, BM_DT_sw37, 37, 0, false)
Set DT26 = (new DropTarget)(sw26, sw26a, BM_DT_sw26, 26, 0, false)
Set DT36 = (new DropTarget)(sw36, sw36a, BM_DT_sw36, 36, 0, false)
Set DT35 = (new DropTarget)(sw35, sw35a, BM_DT_sw35, 35, 0, false)


Dim DTArray
DTArray = Array(DT17, DT27, DT37, DT26, DT36, DT35)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 15 'VP units primitive raises above the up position on drops up
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
    If switch = 35 Then
      PlaysoundAtVol "S20_RollOverTargetTrip", prim, VolTarg
    Else
      SoundDropTargetDrop prim
    End If
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
      'Dim BOT
      'BOT = GetBalls

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

Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_DT_sw17.transz
  rx = BM_DT_sw17.rotx
  ry = BM_DT_sw17.roty
  For each BP in BP_DT_sw17: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT_sw27.transz
  rx = BM_DT_sw27.rotx
  ry = BM_DT_sw27.roty
  For each BP in BP_DT_sw27: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT_sw37.transz
  rx = BM_DT_sw37.rotx
  ry = BM_DT_sw37.roty
  For each BP in BP_DT_sw37: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT_sw26.transz
  rx = BM_DT_sw26.rotx
  ry = BM_DT_sw26.roty
  For each BP in BP_DT_sw26: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT_sw36.transz
  rx = BM_DT_sw36.rotx
  ry = BM_DT_sw36.roty
  For each BP in BP_DT_sw36: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT_sw35.transz
  rx = BM_DT_sw35.rotx
  ry = BM_DT_sw35.roty
  For each BP in BP_DT_sw35: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next
End Sub


'******************************************************
'****  END DROP TARGETS
'******************************************************

'*******************************************
'  ZKTA: KICKING TARGETS
'*******************************************

 'activeball needed?

Sub sw14_Hit: KTHit 14: TargetBouncer Activeball, 0.5: End Sub
Sub sw15_Hit: KTHit 15: TargetBouncer Activeball, 0.5: End Sub
Sub sw16_Hit: KTHit 16: TargetBouncer Activeball, 0.5: End Sub

'Kick Objects
Sub sw14col_hit: KTKick 14: End Sub
Sub sw15col_hit: KTKick 15: End Sub
Sub sw16col_hit: KTKick 16: End Sub

Class KickingTarget
  Private m_primary, m_prim, m_sw, m_animate
  Public ball

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public default Function init(primary, prim, sw, animate)
    Set m_primary = primary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate

    Set Init = Me
  End Function
End Class

'Define a variable for each kicking target
Dim KT14, KT15, KT16  'Kicking

'Set array with kicking target objects
'
'KickingTargetvar = Array(primary, prim, swtich)
'   primary:  vp target to determine target hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          transy must be used to offset the target animation
'   switch:  ROM switch number
'   animate:  Arrary slot for handling the animation instrucitons, set to 0
'

Set KT14 = (new KickingTarget)(sw14, BM_KT_sw14, 14, 0)
Set KT15 = (new KickingTarget)(sw15, BM_KT_sw15, 15, 0)
Set KT16 = (new KickingTarget)(sw16, BM_KT_sw16, 16, 0)

'Add all the Kicking Target Arrays to Kicking Target Animation Array
'   KTAnimationArray = Array(KT1, KT2, ....)
Dim KTArray
KTArray = Array(KT14, KT15, KT16)

' kSpring   - strength of the target spring (non-kicking)
' kKickVel  - velocity imparted on the ball with the target solenoid fires
' kDist   - distance the target must be displaced before the switch registers and the solenoid fires
' kWidth  - total width of the target
Dim kSpring, kKickVel, kDist, kWidth
kSpring = 25
kKickvel = 30
kDist = 30
kWidth = 66

'''''' KICKING TARGETS FUNCTIONS

Sub KTHit(switch)
  Dim i
  i = KTArrayID(switch)

  PlayTargetSound
  KTArray(i).animate = STCheckHit(Activeball,KTArray(i).primary)

  Set KTArray(i).ball = Activeball

  Activeball.velx = cor.BallVelx(Activeball.id)
  Activeball.vely = cor.BallVely(Activeball.id)

  DoKTAnim
End Sub

Sub KTKick(switch)
  Dim i
  i = KTArrayID(switch)

  KTArray(i).animate = 2

  Activeball.velx = cor.BallVelx(Activeball.id)
  Activeball.vely = cor.BallVely(Activeball.id)

  DoKTAnim
End Sub


Function KTArrayID(switch)
  Dim i
  For i = 0 To UBound(KTArray)
    If KTArray(i).sw = switch Then
      KTArrayID = i
      Exit Function
    End If
  Next
End Function

Sub KTAnim_Timer
  DoKTAnim
  dim bangle
End Sub

Sub DoKTAnim()
  Dim i
  For i = 0 To UBound(KTArray)
    KTArray(i).animate = KTAnimate(KTArray(i)) '(KTArray(i).primary,KTArray(i).prim,KTArray(i).sw,KTArray(i).animate,KTArray(i).ball)
  Next
End Sub

Function KTAnimate(arr) '(primary, prim, switch,  animate, aball)
  KTAnimate = arr.animate

  If arr.animate = 0  Then
    arr.primary.uservalue = 0
    KTAnimate = 0
    arr.primary.collidable = 1
    Exit Function
  ElseIf arr.primary.uservalue = 0 Then
    arr.primary.uservalue = GameTime
  End If

  Dim animtime, btdist, btwidth, btangle, bangle, angle, nposx, nposy, tdist, kparavel, kballvel
  tdist = 31.31

  angle = arr.primary.orientation

  animtime = GameTime - arr.primary.uservalue
  arr.primary.uservalue = GameTime

  nposx = arr.primary.x + arr.prim.transy * dCos(angle + 90)
  nposy = arr.primary.y + arr.prim.transy * dSin(angle + 90)

  btwidth = DistancePL(arr.ball.x,arr.ball.y,arr.primary.x,arr.primary.y,arr.primary.x+dcos(angle+90),arr.primary.y+dsin(angle+90))
  btdist = DistancePL(arr.ball.x,arr.ball.y,nposx,nposy,nposx+dcos(angle),nposy+dsin(angle))

  kballvel = sqr(arr.ball.velx^2 + arr.ball.vely^2)

  if kballvel <> 0 Then
    bangle = arcCos(arr.ball.velx/kballvel)*180/PI
  Else
    bangle = 0
  End If

  If arr.ball.vely < 0 Then
    bangle = bangle * -1
  End If

  btangle = bangle - angle

  arr.prim.transy = arr.prim.transy + btdist - tdist

  'debug.print btangle & " " & btwidth & " " & btdist &  " " & prim.transy

  If arr.animate = 1 Then
    arr.primary.collidable = 0

    If btdist < tdist and btwidth < kWidth/2 + 25 Then
      if abs(arr.prim.transy) >= kDist Then
        vpmTimer.PulseSw arr.sw
        'fire Solenoid
        Select Case arr.sw:
          Case 14:
            PlaysoundAtVol "S05_LeftKickingTarget", BM_KT_sw14, VolTarg
          Case 15:
            PlaysoundAtVol "S06_CenterKickingTarget", BM_KT_sw15, VolTarg
          Case 16:
            PlaysoundAtVol "S07_RightKickingTarget", BM_KT_sw16, VolTarg
        End Select
        kparavel = Sqr(arr.ball.velx^2 + arr.ball.vely^2)*dCos(btangle)
        arr.ball.velx = dcos(angle)*kparavel - (kKickVel * dsin(angle))
        arr.ball.vely = dsin(angle)*kparavel + (kKickVel * dcos(angle))
        KTAnimate = 3
        'debug.print HHBall.velx & " fire " & HHBall.vely & " kpara " & kparavel
      Else
        arr.ball.velx = arr.ball.velx - (kSpring * dsin(angle) * abs(arr.prim.transy) * animtime/1000)
        arr.ball.vely = arr.ball.vely + (kSpring * dcos(angle) * abs(arr.prim.transy) * animtime/1000)
        'debug.print HHBall.velx & " nofire " & HHBall.vely
        'debug.print - (kSpring * dsin(angle) * abs(prim.transy) * animtime/1000) & " delta " & + (kSpring * dcos(angle) * abs(prim.transy) * animtime/1000)
      End If
    Elseif btdist > tdist then
      arr.ball.velx = arr.ball.velx - (kSpring * dsin(angle) * abs(arr.prim.transy) * animtime/1000)
      arr.ball.vely = arr.ball.vely + (kSpring * dcos(angle) * abs(arr.prim.transy) * animtime/1000)
      'debug.print HHBall.velx & " nofire " & HHBall.vely
      'debug.print - (kSpring * dsin(angle) * abs(prim.transy) * animtime/1000) & " delta " & + (kSpring * dcos(angle) * abs(prim.transy) * animtime/1000)

      if arr.prim.transy >= 0 Then
        arr.prim.transy = 0
        arr.ball = Empty
        KTAnimate = 0
        Exit Function
      end If
    Else
      arr.prim.transy = arr.prim.transy + 1
      if arr.prim.transy >= 0 Then
        arr.prim.transy = 0
        arr.ball = Empty
        KTAnimate = 0
        Exit Function
      end If
    End If
  Elseif arr.animate = 2 Then
    vpmTimer.PulseSw arr.sw
    'fire Solenoid
    RandomSoundSlingshotLeft arr.primary
    kparavel = Sqr(arr.ball.velx^2 + arr.ball.vely^2)*dCos(btangle)
    arr.ball.velx = dcos(angle)*kparavel - (kKickVel * dsin(angle))
    arr.ball.vely = dsin(angle)*kparavel + (kKickVel * dcos(angle))
    KTAnimate = 3
    'debug.print HHBall.velx & " fire2 " & HHBall.vely & " kpara " & kparavel
  Elseif arr.animate = 3 Then
    'debug.print HHBall.velx & " fire3 " & HHBall.vely & " tang " & dCos(btangle) * kballvel

    if arr.prim.transy >= 0 Then
      arr.prim.transy = 0
      arr.ball = Empty
      KTAnimate = 0
      Exit Function
    end If
  End If
End Function

Sub UpdateKickingTargets
  dim BP, ty

    ty = BM_KT_sw14.transy
  For each BP in BP_KT_sw14 : BP.transy = ty: Next

    ty = BM_KT_sw15.transy
  For each BP in BP_KT_sw15 : BP.transy = ty: Next

    ty = BM_KT_sw16.transy
  For each BP in BP_KT_sw16 : BP.transy = ty: Next

End Sub

'******************************************************
'*   END KICKING TARGETS
'******************************************************





'******************************************
'  JLouLou SYS3 Freeplay & Tournament MOD
'******************************************

sub TournamentTimer_Timer()
  Controller.Switch(5)=1
  Controller.Switch(6)=1
  TournamentTimer.Enabled = False
End Sub


'***************************************************************
' ZVRR: VR Room
'***************************************************************

Sub SetupRoom
  Dim VRThing, BP
  If RenderingMode = 2 OR TestVRonDT = True Then
    For Each VRThing in VR: VRThing.visible = 1: Next
  Else
    For Each VRThing in VR: VRThing.visible = 0: Next
  End If
End Sub



'******************************************************
' ZBRL: BALL ROLLING AND DROP SOUNDS
'******************************************************


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
  Dim b, gBOT
  gBOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ < -1 and gBOT(b).z < 55 and gBOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz > -7 Then
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
dim RampBalls(4,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(4)

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
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
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

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub


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

  If tmp > 0 Then
    AudioFade = CSng(tmp ^ 10)
  Else
    AudioFade = CSng( - (( - tmp) ^ 10) )
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

  If tmp > 0 Then
    AudioPan = CSng(tmp ^ 10)
  Else
    AudioPan = CSng( - (( - tmp) ^ 10) )
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
  PlaySoundAtLevelStatic SoundFX("Knocker",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 4) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 7) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 7) + 1,DOFContactors), SlingshotSoundLevel, Sling
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

Sub SoundScoopFall
      PlaySoundAtBallVol "z_RCT_Scoop_Fall_Skillshot", VolumeDial*0.1
End Sub

Sub SoundScoopKick
      PlaySoundAtBallVol "z_RCT_Scoop", VolumeDial*0.3
End Sub


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

Sub RandomSoundMetal2(ball)
  PlaySoundAtLevelStatic ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ball) * MetalImpactSoundFactor, ball
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
  Cor.Update
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
Const RelayGISoundLevel = 1.08    'volume level; range [0, 1];

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


'///////////////////////////  COIN SOUNDS  ///////////////////////////

Sub SoundCoinIn
  Select Case Int(rnd*3)
    Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
    Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
    Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////


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

Function dArcSin(x)
  If X = 1 Then
    dArcSin = 90
  ElseIf x = -1 Then
    dArcSin = -90
  Else
    dArcSin = Atn(X / Sqr(-X * X + 1))*180/PI
  End If
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


'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 860          'X position of punger lane left
Const PLRight = 920       'X position of punger lane right
Const PLTop = 1225        'Y position of punger lane top
Const PLBottom = 1900       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135 ' orig was 120 and 70

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    If gBOT(s).z >= 24 Then ' above PF
      d_w = b_base*(1 - (gBOT(s).z-25)/500)
    Else  ' below PF
      d_w = b_base*(1 + (gBOT(s).z-25)/100)
    End If
    If d_w < 30 Then d_w = 30
    ' Handle plunger lane
    If InRect(gBOT(s).x,gBOT(s).y,PLLeft,PLBottom,PLLeft,PLTop,PLRight,PLTop,PLRight,PLBottom) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-PLBottom))
    End If
    ' Assign color
    b_r = Int(d_w)
    b_g = Int(d_w)
    b_b = Int(d_w)
    If b_r > 255 Then b_r = 255
    If b_g > 255 Then b_g = 255
    If b_b > 255 Then b_b = 255
    gBOT(s).color = b_r + (b_g * 256) + (b_b * 256 * 256)
    'debug.print "--- ball.color level="&b_r
  Next
End Sub

' this is for updating the lower playfield ball brightness
Sub L02_animate
  dim c: c = 34 + L02.state*180
  c = Int(c)
  If c > 255 Then c = 255
  'debug.print "LPF ball c="&c
  CarBall.color = c + (c * 256) + (c * 256 * 256)
End Sub


'****************************
'   ZRBR: Room Brightness
'****************************

'This code only applies to lightmapped tables. It is here for reference.
'NOTE: Objects bightness will be affected by the Day/Night slider only if their blenddisablelighting property is less than 1.
'      Lightmapped table primitives have their blenddisablelighting equal to 1, therefore we need this SetRoomBrightness sub
'      to handle updating their effective ambient brighness.

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid", "VLM.Bake.Pyramid")

Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0

  ' Lighting level
  Dim v: v=(lvl * 245 + 10)/255

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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub




' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Autoplunger: BP_Autoplunger=Array(BM_Autoplunger)
Dim BP_Bumper1_Ring: BP_Bumper1_Ring=Array(BM_Bumper1_Ring, LM_GIP_GI_002_Bumper1_Ring, LM_GIG_Bumper1_Ring, LM_GIP_GI_007_Bumper1_Ring, LM_L_L41_Bumper1_Ring, LM_L_L43_Bumper1_Ring, LM_F_f0_Bumper1_Ring)
Dim BP_Bumper1_Skirt: BP_Bumper1_Skirt=Array(BM_Bumper1_Skirt, LM_GIP_GI_007_Bumper1_Skirt, LM_GIB_Bumper1_Skirt, LM_L_L15_Bumper1_Skirt, LM_L_L17_Bumper1_Skirt, LM_L_L41_Bumper1_Skirt, LM_L_L43_Bumper1_Skirt, LM_F_f0_Bumper1_Skirt)
Dim BP_Bumper2_Ring: BP_Bumper2_Ring=Array(BM_Bumper2_Ring, LM_GIP_GI_007_Bumper2_Ring, LM_GIB_Bumper2_Ring, LM_L_L42_Bumper2_Ring, LM_L_L44_Bumper2_Ring, LM_F_f0_Bumper2_Ring, LM_F_f5_Bumper2_Ring)
Dim BP_Bumper2_Skirt: BP_Bumper2_Skirt=Array(BM_Bumper2_Skirt, LM_GIP_GI_007_Bumper2_Skirt, LM_GIB_Bumper2_Skirt, LM_L_L42_Bumper2_Skirt, LM_L_L44_Bumper2_Skirt, LM_F_f0_Bumper2_Skirt)
Dim BP_DT_sw17: BP_DT_sw17=Array(BM_DT_sw17, LM_GIP_GI_001_DT_sw17, LM_GIP_GI_002_DT_sw17, LM_GIG_DT_sw17, LM_GIP_GI_006_DT_sw17, LM_GIP_GI_007_DT_sw17, LM_F_f5_DT_sw17)
Dim BP_DT_sw26: BP_DT_sw26=Array(BM_DT_sw26, LM_GIG_DT_sw26, LM_GIP_GI_007_DT_sw26, LM_GIB_DT_sw26, LM_L_L17_DT_sw26, LM_L_L33_DT_sw26, LM_L_L34_DT_sw26, LM_L_L45_DT_sw26, LM_F_f4_DT_sw26, LM_F_f5_DT_sw26)
Dim BP_DT_sw27: BP_DT_sw27=Array(BM_DT_sw27, LM_GIP_GI_001_DT_sw27, LM_GIG_DT_sw27, LM_GIP_GI_006_DT_sw27, LM_GIP_GI_007_DT_sw27, LM_F_f5_DT_sw27)
Dim BP_DT_sw35: BP_DT_sw35=Array(BM_DT_sw35, LM_GIR_DT_sw35, LM_L_L27_DT_sw35, LM_F_f5_DT_sw35)
Dim BP_DT_sw36: BP_DT_sw36=Array(BM_DT_sw36, LM_GIG_DT_sw36, LM_GIP_GI_007_DT_sw36, LM_GIB_DT_sw36, LM_L_L34_DT_sw36, LM_L_L45_DT_sw36, LM_L_L46_DT_sw36, LM_F_f1_DT_sw36, LM_F_f4_DT_sw36, LM_F_f5_DT_sw36)
Dim BP_DT_sw37: BP_DT_sw37=Array(BM_DT_sw37, LM_GIP_GI_002_DT_sw37, LM_GIG_DT_sw37, LM_GIP_GI_006_DT_sw37, LM_GIP_GI_007_DT_sw37, LM_GIB_DT_sw37, LM_L_L15_DT_sw37, LM_F_f5_DT_sw37)
Dim BP_Glider_1: BP_Glider_1=Array(BM_Glider_1, LM_GIG_Glider_1, LM_GIY_Glider_1, LM_F_f1_Glider_1, LM_F_f2_Glider_1, LM_F_f3_Glider_1, LM_F_f4_Glider_1)
Dim BP_Glider_1_001: BP_Glider_1_001=Array(BM_Glider_1_001, LM_GIY_Glider_1_001, LM_F_f1_Glider_1_001, LM_F_f2_Glider_1_001, LM_F_f3_Glider_1_001, LM_F_f4_Glider_1_001)
Dim BP_Glider_1_002: BP_Glider_1_002=Array(BM_Glider_1_002, LM_GIG_Glider_1_002, LM_GIY_Glider_1_002, LM_F_f2_Glider_1_002, LM_F_f3_Glider_1_002, LM_F_f4_Glider_1_002)
Dim BP_Glider_2: BP_Glider_2=Array(BM_Glider_2, LM_GIY_Glider_2, LM_F_f1_Glider_2, LM_F_f2_Glider_2, LM_F_f3_Glider_2, LM_F_f4_Glider_2)
Dim BP_Glider_2_001: BP_Glider_2_001=Array(BM_Glider_2_001, LM_F_f2_Glider_2_001, LM_F_f3_Glider_2_001, LM_F_f4_Glider_2_001)
Dim BP_GuardianL: BP_GuardianL=Array(BM_GuardianL, LM_GIB_GuardianL, LM_L_L44_GuardianL, LM_F_f4_GuardianL)
Dim BP_GuardianL_001: BP_GuardianL_001=Array(BM_GuardianL_001, LM_GIG_GuardianL_001, LM_GIB_GuardianL_001, LM_L_L44_GuardianL_001, LM_F_f1_GuardianL_001, LM_F_f4_GuardianL_001)
Dim BP_GuardianR: BP_GuardianR=Array(BM_GuardianR, LM_GIG_GuardianR, LM_GIR_GuardianR, LM_GIY_GuardianR, LM_L_L65_GuardianR, LM_L_L66_GuardianR, LM_L_L67_GuardianR, LM_F_f1_GuardianR, LM_F_f2_GuardianR, LM_F_f3_GuardianR, LM_F_f4_GuardianR)
Dim BP_GuardianR_001: BP_GuardianR_001=Array(BM_GuardianR_001, LM_GIG_GuardianR_001, LM_GIR_GuardianR_001, LM_L_L67_GuardianR_001, LM_F_f2_GuardianR_001, LM_F_f3_GuardianR_001, LM_F_f4_GuardianR_001)
Dim BP_KT_sw14: BP_KT_sw14=Array(BM_KT_sw14, LM_GIB_KT_sw14, LM_L_L44_KT_sw14, LM_F_f0_KT_sw14)
Dim BP_KT_sw15: BP_KT_sw15=Array(BM_KT_sw15, LM_GIG_KT_sw15, LM_GIR_KT_sw15, LM_GIB_KT_sw15, LM_GIY_KT_sw15, LM_L_L57_KT_sw15, LM_F_f1_KT_sw15, LM_F_f2_KT_sw15, LM_F_f4_KT_sw15)
Dim BP_KT_sw16: BP_KT_sw16=Array(BM_KT_sw16, LM_GIP_GI_006_KT_sw16, LM_GIR_KT_sw16, LM_L_L76_KT_sw16, LM_F_f5_KT_sw16)
Dim BP_LF: BP_LF=Array(BM_LF, LM_GIP_GI_001_LF, LM_GIG_LF, LM_L_L0_LF, LM_L_L82_LF, LM_L_L83_LF)
Dim BP_LFU: BP_LFU=Array(BM_LFU, LM_GIP_GI_001_LFU, LM_GIP_GI_002_LFU, LM_GIG_LFU, LM_GIP_GI_007_LFU, LM_L_L82_LFU, LM_L_L83_LFU, LM_F_f5_LFU)
Dim BP_LSling: BP_LSling=Array(BM_LSling, LM_GIP_GI_001_LSling, LM_GIP_GI_002_LSling, LM_GIG_LSling, LM_GIP_GI_007_LSling, LM_L_L12_LSling, LM_L_L82_LSling, LM_F_f5_LSling)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GIP_GI_001_LSling1, LM_GIP_GI_002_LSling1, LM_GIG_LSling1, LM_GIP_GI_007_LSling1, LM_L_L12_LSling1, LM_L_L81_LSling1, LM_L_L82_LSling1, LM_F_f4_LSling1, LM_F_f5_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GIP_GI_001_LSling2, LM_GIP_GI_002_LSling2, LM_GIG_LSling2, LM_GIP_GI_007_LSling2, LM_L_L12_LSling2, LM_L_L82_LSling2, LM_F_f4_LSling2, LM_F_f5_LSling2)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_GIP_GI_002_Layer1, LM_GIG_Layer1, LM_GIP_GI_006_Layer1, LM_GIR_Layer1, LM_GIB_Layer1, LM_L_L25_Layer1, LM_L_L26_Layer1, LM_L_L27_Layer1, LM_F_f1_Layer1, LM_F_f2_Layer1, LM_F_f4_Layer1, LM_F_f5_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_GIP_GI_002_Layer2, LM_GIG_Layer2, LM_GIP_GI_007_Layer2, LM_L_L11_Layer2, LM_L_L12_Layer2, LM_L_L43_Layer2, LM_L_L45_Layer2, LM_F_f1_Layer2, LM_F_f3_Layer2, LM_F_f5_Layer2)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_GIP_GI_001_Parts, LM_GIP_GI_002_Parts, LM_GIG_Parts, LM_GIP_GI_006_Parts, LM_GIP_GI_007_Parts, LM_GIR_Parts, LM_GIB_Parts, LM_GIY_Parts, LM_L_L0_Parts, LM_L_L11_Parts, LM_L_L12_Parts, LM_L_L13_Parts, LM_L_L14_Parts, LM_L_L15_Parts, LM_L_L16_Parts, LM_L_L17_Parts, LM_L_L21_Parts, LM_L_L22_Parts, LM_L_L23_Parts, LM_L_L24_Parts, LM_L_L25_Parts, LM_L_L26_Parts, LM_L_L27_Parts, LM_L_L32_Parts, LM_L_L33_Parts, LM_L_L34_Parts, LM_L_L35_Parts, LM_L_L36_Parts, LM_L_L37_Parts, LM_L_L41_Parts, LM_L_L42_Parts, LM_L_L43_Parts, LM_L_L44_Parts, LM_L_L45_Parts, LM_L_L46_Parts, LM_L_L47_Parts, LM_L_L55_Parts, LM_L_L56_Parts, LM_L_L57_Parts, LM_L_L5_Parts, LM_L_L65_Parts, LM_L_L66_Parts, LM_L_L67_Parts, LM_L_L6_Parts, LM_L_L71_Parts, LM_L_L72_Parts, LM_L_L73_Parts, LM_L_L74_Parts, LM_L_L75_Parts, LM_L_L76_Parts, LM_L_L77_Parts, LM_L_L7_Parts, LM_L_L81_Parts, LM_L_L82_Parts, LM_L_L83_Parts, LM_L_L84_Parts, LM_L_L85_Parts, LM_L_L86_Parts, LM_L_L87_Parts, LM_F_f0_Parts, LM_F_f1_Parts, _
  LM_F_f2_Parts, LM_F_f3_Parts, LM_F_f4_Parts, LM_F_f5_Parts)
Dim BP_PinCab_Rails_021: BP_PinCab_Rails_021=Array(BM_PinCab_Rails_021)
Dim BP_PinCab_Rails_022: BP_PinCab_Rails_022=Array(BM_PinCab_Rails_022)
Dim BP_PinCab_Rails_023: BP_PinCab_Rails_023=Array(BM_PinCab_Rails_023)
Dim BP_Plastics: BP_Plastics=Array(BM_Plastics, LM_GIP_GI_001_Plastics, LM_GIG_Plastics, LM_GIP_GI_006_Plastics, LM_GIP_GI_007_Plastics, LM_GIR_Plastics, LM_GIB_Plastics, LM_GIY_Plastics, LM_L_L11_Plastics, LM_L_L42_Plastics, LM_F_f0_Plastics, LM_F_f1_Plastics, LM_F_f2_Plastics, LM_F_f3_Plastics, LM_F_f4_Plastics, LM_F_f5_Plastics)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GIP_GI_001_Playfield, LM_GIP_GI_002_Playfield, LM_GIG_Playfield, LM_GIP_GI_006_Playfield, LM_GIP_GI_007_Playfield, LM_GIR_Playfield, LM_GIB_Playfield, LM_GIY_Playfield, LM_L_L0_Playfield, LM_L_L11_Playfield, LM_L_L12_Playfield, LM_L_L13_Playfield, LM_L_L14_Playfield, LM_L_L15_Playfield, LM_L_L16_Playfield, LM_L_L17_Playfield, LM_L_L21_Playfield, LM_L_L22_Playfield, LM_L_L23_Playfield, LM_L_L24_Playfield, LM_L_L25_Playfield, LM_L_L26_Playfield, LM_L_L27_Playfield, LM_L_L32_Playfield, LM_L_L33_Playfield, LM_L_L34_Playfield, LM_L_L35_Playfield, LM_L_L36_Playfield, LM_L_L37_Playfield, LM_L_L43_Playfield, LM_L_L44_Playfield, LM_L_L45_Playfield, LM_L_L46_Playfield, LM_L_L47_Playfield, LM_L_L55_Playfield, LM_L_L56_Playfield, LM_L_L57_Playfield, LM_L_L5_Playfield, LM_L_L65_Playfield, LM_L_L66_Playfield, LM_L_L67_Playfield, LM_L_L6_Playfield, LM_L_L71_Playfield, LM_L_L72_Playfield, LM_L_L73_Playfield, LM_L_L74_Playfield, LM_L_L75_Playfield, LM_L_L76_Playfield, _
  LM_L_L77_Playfield, LM_L_L7_Playfield, LM_L_L81_Playfield, LM_L_L82_Playfield, LM_L_L83_Playfield, LM_L_L84_Playfield, LM_L_L85_Playfield, LM_L_L86_Playfield, LM_L_L87_Playfield, LM_F_f0_Playfield, LM_F_f1_Playfield, LM_F_f2_Playfield, LM_F_f3_Playfield, LM_F_f4_Playfield, LM_F_f5_Playfield)
Dim BP_Pyramid1: BP_Pyramid1=Array(BM_Pyramid1, LM_GIG_Pyramid1, LM_GIB_Pyramid1, LM_F_f1_Pyramid1, LM_F_f2_Pyramid1, LM_F_f3_Pyramid1, LM_F_f4_Pyramid1)
Dim BP_Pyramid1_001: BP_Pyramid1_001=Array(BM_Pyramid1_001, LM_F_f2_Pyramid1_001, LM_F_f3_Pyramid1_001)
Dim BP_RF: BP_RF=Array(BM_RF, LM_GIP_GI_001_RF, LM_GIP_GI_002_RF, LM_GIG_RF, LM_L_L0_RF, LM_L_L84_RF, LM_L_L85_RF, LM_F_f5_RF)
Dim BP_RF1: BP_RF1=Array(BM_RF1, LM_GIP_GI_006_RF1, LM_GIR_RF1, LM_L_L25_RF1, LM_F_f4_RF1, LM_F_f5_RF1)
Dim BP_RF1U: BP_RF1U=Array(BM_RF1U, LM_GIG_RF1U, LM_GIP_GI_006_RF1U, LM_GIR_RF1U, LM_L_L25_RF1U, LM_L_L26_RF1U, LM_L_L35_RF1U, LM_F_f1_RF1U, LM_F_f4_RF1U, LM_F_f5_RF1U)
Dim BP_RFU: BP_RFU=Array(BM_RFU, LM_GIP_GI_001_RFU, LM_GIP_GI_002_RFU, LM_GIG_RFU, LM_L_L84_RFU, LM_L_L85_RFU, LM_F_f5_RFU)
Dim BP_RSling: BP_RSling=Array(BM_RSling, LM_GIP_GI_001_RSling, LM_GIP_GI_002_RSling, LM_GIG_RSling, LM_GIP_GI_006_RSling, LM_L_L13_RSling, LM_L_L6_RSling, LM_L_L7_RSling, LM_L_L85_RSling, LM_F_f5_RSling)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GIP_GI_001_RSling1, LM_GIP_GI_002_RSling1, LM_GIG_RSling1, LM_GIP_GI_006_RSling1, LM_L_L13_RSling1, LM_L_L5_RSling1, LM_L_L7_RSling1, LM_L_L85_RSling1, LM_L_L86_RSling1, LM_F_f5_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GIP_GI_001_RSling2, LM_GIP_GI_002_RSling2, LM_GIG_RSling2, LM_GIP_GI_006_RSling2, LM_L_L13_RSling2, LM_L_L6_RSling2, LM_L_L7_RSling2, LM_L_L85_RSling2, LM_L_L86_RSling2, LM_F_f5_RSling2)
Dim BP_ST_sw22: BP_ST_sw22=Array(BM_ST_sw22, LM_GIG_ST_sw22, LM_GIY_ST_sw22, LM_F_f2_ST_sw22, LM_F_f4_ST_sw22)
Dim BP_ST_sw32: BP_ST_sw32=Array(BM_ST_sw32, LM_GIG_ST_sw32, LM_GIY_ST_sw32, LM_L_L46_ST_sw32, LM_L_L47_ST_sw32, LM_L_L55_ST_sw32, LM_F_f1_ST_sw32, LM_F_f4_ST_sw32)
Dim BP_SarcArm: BP_SarcArm=Array(BM_SarcArm, LM_GIP_GI_002_SarcArm, LM_F_f5_SarcArm)
Dim BP_SarcArmU: BP_SarcArmU=Array(BM_SarcArmU, LM_GIP_GI_002_SarcArmU, LM_GIG_SarcArmU)
Dim BP_SideBlades_Art: BP_SideBlades_Art=Array(BM_SideBlades_Art, LM_GIP_GI_001_SideBlades_Art, LM_GIP_GI_002_SideBlades_Art, LM_GIG_SideBlades_Art, LM_GIP_GI_006_SideBlades_Art, LM_GIP_GI_007_SideBlades_Art, LM_GIR_SideBlades_Art, LM_GIB_SideBlades_Art, LM_GIY_SideBlades_Art, LM_L_L43_SideBlades_Art, LM_F_f0_SideBlades_Art, LM_F_f1_SideBlades_Art, LM_F_f2_SideBlades_Art, LM_F_f3_SideBlades_Art, LM_F_f5_SideBlades_Art)
Dim BP_SlingL: BP_SlingL=Array(BM_SlingL, LM_GIP_GI_002_SlingL, LM_GIG_SlingL, LM_F_f5_SlingL)
Dim BP_SlingR: BP_SlingR=Array(BM_SlingR, LM_GIP_GI_001_SlingR, LM_GIG_SlingR)
Dim BP_UnderPF: BP_UnderPF=Array(BM_UnderPF, LM_GIP_GI_001_UnderPF, LM_GIP_GI_002_UnderPF, LM_GIG_UnderPF, LM_GIP_GI_006_UnderPF, LM_GIP_GI_007_UnderPF, LM_GIR_UnderPF, LM_GIB_UnderPF, LM_GIY_UnderPF, LM_L_L0_UnderPF, LM_L_L12_UnderPF, LM_L_L13_UnderPF, LM_L_L14_UnderPF, LM_L_L15_UnderPF, LM_L_L16_UnderPF, LM_L_L17_UnderPF, LM_L_L21_UnderPF, LM_L_L22_UnderPF, LM_L_L23_UnderPF, LM_L_L24_UnderPF, LM_L_L25_UnderPF, LM_L_L26_UnderPF, LM_L_L27_UnderPF, LM_L_L32_UnderPF, LM_L_L33_UnderPF, LM_L_L34_UnderPF, LM_L_L36_UnderPF, LM_L_L37_UnderPF, LM_L_L42_UnderPF, LM_L_L43_UnderPF, LM_L_L44_UnderPF, LM_L_L45_UnderPF, LM_L_L46_UnderPF, LM_L_L47_UnderPF, LM_L_L55_UnderPF, LM_L_L56_UnderPF, LM_L_L57_UnderPF, LM_L_L65_UnderPF, LM_L_L66_UnderPF, LM_L_L67_UnderPF, LM_L_L6_UnderPF, LM_L_L71_UnderPF, LM_L_L72_UnderPF, LM_L_L73_UnderPF, LM_L_L74_UnderPF, LM_L_L75_UnderPF, LM_L_L76_UnderPF, LM_L_L77_UnderPF, LM_L_L7_UnderPF, LM_L_L81_UnderPF, LM_L_L82_UnderPF, LM_L_L83_UnderPF, LM_L_L84_UnderPF, LM_L_L85_UnderPF, _
  LM_L_L86_UnderPF, LM_L_L87_UnderPF, LM_F_f0_UnderPF, LM_F_f1_UnderPF, LM_F_f2_UnderPF, LM_F_f3_UnderPF, LM_F_f4_UnderPF, LM_F_f5_UnderPF)
Dim BP_Wood_guide___Painted_Blaxk_1: BP_Wood_guide___Painted_Blaxk_1=Array(BM_Wood_guide___Painted_Blaxk_1, LM_F_f2_Wood_guide___Painted_Bl, LM_F_f3_Wood_guide___Painted_Bl)
Dim BP_sw111: BP_sw111=Array(BM_sw111, LM_GIP_GI_002_sw111, LM_GIG_sw111)
Dim BP_sw112: BP_sw112=Array(BM_sw112, LM_GIP_GI_002_sw112, LM_GIG_sw112, LM_F_f4_sw112)
Dim BP_sw113: BP_sw113=Array(BM_sw113, LM_GIG_sw113, LM_F_f5_sw113)
Dim BP_sw114: BP_sw114=Array(BM_sw114, LM_GIG_sw114)
Dim BP_sw115: BP_sw115=Array(BM_sw115, LM_GIP_GI_007_sw115)
Dim BP_sw31: BP_sw31=Array(BM_sw31)
' Arrays per lighting scenario
Dim BL_F_f0: BL_F_f0=Array(LM_F_f0_Bumper1_Ring, LM_F_f0_Bumper1_Skirt, LM_F_f0_Bumper2_Ring, LM_F_f0_Bumper2_Skirt, LM_F_f0_KT_sw14, LM_F_f0_Parts, LM_F_f0_Plastics, LM_F_f0_Playfield, LM_F_f0_SideBlades_Art, LM_F_f0_UnderPF)
Dim BL_F_f1: BL_F_f1=Array(LM_F_f1_DT_sw36, LM_F_f1_Glider_1, LM_F_f1_Glider_1_001, LM_F_f1_Glider_2, LM_F_f1_GuardianL_001, LM_F_f1_GuardianR, LM_F_f1_KT_sw15, LM_F_f1_Layer1, LM_F_f1_Layer2, LM_F_f1_Parts, LM_F_f1_Plastics, LM_F_f1_Playfield, LM_F_f1_Pyramid1, LM_F_f1_RF1U, LM_F_f1_ST_sw32, LM_F_f1_SideBlades_Art, LM_F_f1_UnderPF)
Dim BL_F_f2: BL_F_f2=Array(LM_F_f2_Glider_1, LM_F_f2_Glider_1_001, LM_F_f2_Glider_1_002, LM_F_f2_Glider_2, LM_F_f2_Glider_2_001, LM_F_f2_GuardianR, LM_F_f2_GuardianR_001, LM_F_f2_KT_sw15, LM_F_f2_Layer1, LM_F_f2_Parts, LM_F_f2_Plastics, LM_F_f2_Playfield, LM_F_f2_Pyramid1, LM_F_f2_Pyramid1_001, LM_F_f2_ST_sw22, LM_F_f2_SideBlades_Art, LM_F_f2_UnderPF, LM_F_f2_Wood_guide___Painted_Bl)
Dim BL_F_f3: BL_F_f3=Array(LM_F_f3_Glider_1, LM_F_f3_Glider_1_001, LM_F_f3_Glider_1_002, LM_F_f3_Glider_2, LM_F_f3_Glider_2_001, LM_F_f3_GuardianR, LM_F_f3_GuardianR_001, LM_F_f3_Layer2, LM_F_f3_Parts, LM_F_f3_Plastics, LM_F_f3_Playfield, LM_F_f3_Pyramid1, LM_F_f3_Pyramid1_001, LM_F_f3_SideBlades_Art, LM_F_f3_UnderPF, LM_F_f3_Wood_guide___Painted_Bl)
Dim BL_F_f4: BL_F_f4=Array(LM_F_f4_DT_sw26, LM_F_f4_DT_sw36, LM_F_f4_Glider_1, LM_F_f4_Glider_1_001, LM_F_f4_Glider_1_002, LM_F_f4_Glider_2, LM_F_f4_Glider_2_001, LM_F_f4_GuardianL, LM_F_f4_GuardianL_001, LM_F_f4_GuardianR, LM_F_f4_GuardianR_001, LM_F_f4_KT_sw15, LM_F_f4_LSling1, LM_F_f4_LSling2, LM_F_f4_Layer1, LM_F_f4_Parts, LM_F_f4_Plastics, LM_F_f4_Playfield, LM_F_f4_Pyramid1, LM_F_f4_RF1, LM_F_f4_RF1U, LM_F_f4_ST_sw22, LM_F_f4_ST_sw32, LM_F_f4_UnderPF, LM_F_f4_sw112)
Dim BL_F_f5: BL_F_f5=Array(LM_F_f5_Bumper2_Ring, LM_F_f5_DT_sw17, LM_F_f5_DT_sw26, LM_F_f5_DT_sw27, LM_F_f5_DT_sw35, LM_F_f5_DT_sw36, LM_F_f5_DT_sw37, LM_F_f5_KT_sw16, LM_F_f5_LFU, LM_F_f5_LSling, LM_F_f5_LSling1, LM_F_f5_LSling2, LM_F_f5_Layer1, LM_F_f5_Layer2, LM_F_f5_Parts, LM_F_f5_Plastics, LM_F_f5_Playfield, LM_F_f5_RF, LM_F_f5_RF1, LM_F_f5_RF1U, LM_F_f5_RFU, LM_F_f5_RSling, LM_F_f5_RSling1, LM_F_f5_RSling2, LM_F_f5_SarcArm, LM_F_f5_SideBlades_Art, LM_F_f5_SlingL, LM_F_f5_UnderPF, LM_F_f5_sw113)
Dim BL_GIB: BL_GIB=Array(LM_GIB_Bumper1_Skirt, LM_GIB_Bumper2_Ring, LM_GIB_Bumper2_Skirt, LM_GIB_DT_sw26, LM_GIB_DT_sw36, LM_GIB_DT_sw37, LM_GIB_GuardianL, LM_GIB_GuardianL_001, LM_GIB_KT_sw14, LM_GIB_KT_sw15, LM_GIB_Layer1, LM_GIB_Parts, LM_GIB_Plastics, LM_GIB_Playfield, LM_GIB_Pyramid1, LM_GIB_SideBlades_Art, LM_GIB_UnderPF)
Dim BL_GIG: BL_GIG=Array(LM_GIG_Bumper1_Ring, LM_GIG_DT_sw17, LM_GIG_DT_sw26, LM_GIG_DT_sw27, LM_GIG_DT_sw36, LM_GIG_DT_sw37, LM_GIG_Glider_1, LM_GIG_Glider_1_002, LM_GIG_GuardianL_001, LM_GIG_GuardianR, LM_GIG_GuardianR_001, LM_GIG_KT_sw15, LM_GIG_LF, LM_GIG_LFU, LM_GIG_LSling, LM_GIG_LSling1, LM_GIG_LSling2, LM_GIG_Layer1, LM_GIG_Layer2, LM_GIG_Parts, LM_GIG_Plastics, LM_GIG_Playfield, LM_GIG_Pyramid1, LM_GIG_RF, LM_GIG_RF1U, LM_GIG_RFU, LM_GIG_RSling, LM_GIG_RSling1, LM_GIG_RSling2, LM_GIG_ST_sw22, LM_GIG_ST_sw32, LM_GIG_SarcArmU, LM_GIG_SideBlades_Art, LM_GIG_SlingL, LM_GIG_SlingR, LM_GIG_UnderPF, LM_GIG_sw111, LM_GIG_sw112, LM_GIG_sw113, LM_GIG_sw114)
Dim BL_GIP_GI_001: BL_GIP_GI_001=Array(LM_GIP_GI_001_DT_sw17, LM_GIP_GI_001_DT_sw27, LM_GIP_GI_001_LF, LM_GIP_GI_001_LFU, LM_GIP_GI_001_LSling, LM_GIP_GI_001_LSling1, LM_GIP_GI_001_LSling2, LM_GIP_GI_001_Parts, LM_GIP_GI_001_Plastics, LM_GIP_GI_001_Playfield, LM_GIP_GI_001_RF, LM_GIP_GI_001_RFU, LM_GIP_GI_001_RSling, LM_GIP_GI_001_RSling1, LM_GIP_GI_001_RSling2, LM_GIP_GI_001_SideBlades_Art, LM_GIP_GI_001_SlingR, LM_GIP_GI_001_UnderPF)
Dim BL_GIP_GI_002: BL_GIP_GI_002=Array(LM_GIP_GI_002_Bumper1_Ring, LM_GIP_GI_002_DT_sw17, LM_GIP_GI_002_DT_sw37, LM_GIP_GI_002_LFU, LM_GIP_GI_002_LSling, LM_GIP_GI_002_LSling1, LM_GIP_GI_002_LSling2, LM_GIP_GI_002_Layer1, LM_GIP_GI_002_Layer2, LM_GIP_GI_002_Parts, LM_GIP_GI_002_Playfield, LM_GIP_GI_002_RF, LM_GIP_GI_002_RFU, LM_GIP_GI_002_RSling, LM_GIP_GI_002_RSling1, LM_GIP_GI_002_RSling2, LM_GIP_GI_002_SarcArm, LM_GIP_GI_002_SarcArmU, LM_GIP_GI_002_SideBlades_Art, LM_GIP_GI_002_SlingL, LM_GIP_GI_002_UnderPF, LM_GIP_GI_002_sw111, LM_GIP_GI_002_sw112)
Dim BL_GIP_GI_006: BL_GIP_GI_006=Array(LM_GIP_GI_006_DT_sw17, LM_GIP_GI_006_DT_sw27, LM_GIP_GI_006_DT_sw37, LM_GIP_GI_006_KT_sw16, LM_GIP_GI_006_Layer1, LM_GIP_GI_006_Parts, LM_GIP_GI_006_Plastics, LM_GIP_GI_006_Playfield, LM_GIP_GI_006_RF1, LM_GIP_GI_006_RF1U, LM_GIP_GI_006_RSling, LM_GIP_GI_006_RSling1, LM_GIP_GI_006_RSling2, LM_GIP_GI_006_SideBlades_Art, LM_GIP_GI_006_UnderPF)
Dim BL_GIP_GI_007: BL_GIP_GI_007=Array(LM_GIP_GI_007_Bumper1_Ring, LM_GIP_GI_007_Bumper1_Skirt, LM_GIP_GI_007_Bumper2_Ring, LM_GIP_GI_007_Bumper2_Skirt, LM_GIP_GI_007_DT_sw17, LM_GIP_GI_007_DT_sw26, LM_GIP_GI_007_DT_sw27, LM_GIP_GI_007_DT_sw36, LM_GIP_GI_007_DT_sw37, LM_GIP_GI_007_LFU, LM_GIP_GI_007_LSling, LM_GIP_GI_007_LSling1, LM_GIP_GI_007_LSling2, LM_GIP_GI_007_Layer2, LM_GIP_GI_007_Parts, LM_GIP_GI_007_Plastics, LM_GIP_GI_007_Playfield, LM_GIP_GI_007_SideBlades_Art, LM_GIP_GI_007_UnderPF, LM_GIP_GI_007_sw115)
Dim BL_GIR: BL_GIR=Array(LM_GIR_DT_sw35, LM_GIR_GuardianR, LM_GIR_GuardianR_001, LM_GIR_KT_sw15, LM_GIR_KT_sw16, LM_GIR_Layer1, LM_GIR_Parts, LM_GIR_Plastics, LM_GIR_Playfield, LM_GIR_RF1, LM_GIR_RF1U, LM_GIR_SideBlades_Art, LM_GIR_UnderPF)
Dim BL_GIY: BL_GIY=Array(LM_GIY_Glider_1, LM_GIY_Glider_1_001, LM_GIY_Glider_1_002, LM_GIY_Glider_2, LM_GIY_GuardianR, LM_GIY_KT_sw15, LM_GIY_Parts, LM_GIY_Plastics, LM_GIY_Playfield, LM_GIY_ST_sw22, LM_GIY_ST_sw32, LM_GIY_SideBlades_Art, LM_GIY_UnderPF)
Dim BL_L_L0: BL_L_L0=Array(LM_L_L0_LF, LM_L_L0_Parts, LM_L_L0_Playfield, LM_L_L0_RF, LM_L_L0_UnderPF)
Dim BL_L_L11: BL_L_L11=Array(LM_L_L11_Layer2, LM_L_L11_Parts, LM_L_L11_Plastics, LM_L_L11_Playfield)
Dim BL_L_L12: BL_L_L12=Array(LM_L_L12_LSling, LM_L_L12_LSling1, LM_L_L12_LSling2, LM_L_L12_Layer2, LM_L_L12_Parts, LM_L_L12_Playfield, LM_L_L12_UnderPF)
Dim BL_L_L13: BL_L_L13=Array(LM_L_L13_Parts, LM_L_L13_Playfield, LM_L_L13_RSling, LM_L_L13_RSling1, LM_L_L13_RSling2, LM_L_L13_UnderPF)
Dim BL_L_L14: BL_L_L14=Array(LM_L_L14_Parts, LM_L_L14_Playfield, LM_L_L14_UnderPF)
Dim BL_L_L15: BL_L_L15=Array(LM_L_L15_Bumper1_Skirt, LM_L_L15_DT_sw37, LM_L_L15_Parts, LM_L_L15_Playfield, LM_L_L15_UnderPF)
Dim BL_L_L16: BL_L_L16=Array(LM_L_L16_Parts, LM_L_L16_Playfield, LM_L_L16_UnderPF)
Dim BL_L_L17: BL_L_L17=Array(LM_L_L17_Bumper1_Skirt, LM_L_L17_DT_sw26, LM_L_L17_Parts, LM_L_L17_Playfield, LM_L_L17_UnderPF)
Dim BL_L_L21: BL_L_L21=Array(LM_L_L21_Parts, LM_L_L21_Playfield, LM_L_L21_UnderPF)
Dim BL_L_L22: BL_L_L22=Array(LM_L_L22_Parts, LM_L_L22_Playfield, LM_L_L22_UnderPF)
Dim BL_L_L23: BL_L_L23=Array(LM_L_L23_Parts, LM_L_L23_Playfield, LM_L_L23_UnderPF)
Dim BL_L_L24: BL_L_L24=Array(LM_L_L24_Parts, LM_L_L24_Playfield, LM_L_L24_UnderPF)
Dim BL_L_L25: BL_L_L25=Array(LM_L_L25_Layer1, LM_L_L25_Parts, LM_L_L25_Playfield, LM_L_L25_RF1, LM_L_L25_RF1U, LM_L_L25_UnderPF)
Dim BL_L_L26: BL_L_L26=Array(LM_L_L26_Layer1, LM_L_L26_Parts, LM_L_L26_Playfield, LM_L_L26_RF1U, LM_L_L26_UnderPF)
Dim BL_L_L27: BL_L_L27=Array(LM_L_L27_DT_sw35, LM_L_L27_Layer1, LM_L_L27_Parts, LM_L_L27_Playfield, LM_L_L27_UnderPF)
Dim BL_L_L32: BL_L_L32=Array(LM_L_L32_Parts, LM_L_L32_Playfield, LM_L_L32_UnderPF)
Dim BL_L_L33: BL_L_L33=Array(LM_L_L33_DT_sw26, LM_L_L33_Parts, LM_L_L33_Playfield, LM_L_L33_UnderPF)
Dim BL_L_L34: BL_L_L34=Array(LM_L_L34_DT_sw26, LM_L_L34_DT_sw36, LM_L_L34_Parts, LM_L_L34_Playfield, LM_L_L34_UnderPF)
Dim BL_L_L35: BL_L_L35=Array(LM_L_L35_Parts, LM_L_L35_Playfield, LM_L_L35_RF1U)
Dim BL_L_L36: BL_L_L36=Array(LM_L_L36_Parts, LM_L_L36_Playfield, LM_L_L36_UnderPF)
Dim BL_L_L37: BL_L_L37=Array(LM_L_L37_Parts, LM_L_L37_Playfield, LM_L_L37_UnderPF)
Dim BL_L_L41: BL_L_L41=Array(LM_L_L41_Bumper1_Ring, LM_L_L41_Bumper1_Skirt, LM_L_L41_Parts)
Dim BL_L_L42: BL_L_L42=Array(LM_L_L42_Bumper2_Ring, LM_L_L42_Bumper2_Skirt, LM_L_L42_Parts, LM_L_L42_Plastics, LM_L_L42_UnderPF)
Dim BL_L_L43: BL_L_L43=Array(LM_L_L43_Bumper1_Ring, LM_L_L43_Bumper1_Skirt, LM_L_L43_Layer2, LM_L_L43_Parts, LM_L_L43_Playfield, LM_L_L43_SideBlades_Art, LM_L_L43_UnderPF)
Dim BL_L_L44: BL_L_L44=Array(LM_L_L44_Bumper2_Ring, LM_L_L44_Bumper2_Skirt, LM_L_L44_GuardianL, LM_L_L44_GuardianL_001, LM_L_L44_KT_sw14, LM_L_L44_Parts, LM_L_L44_Playfield, LM_L_L44_UnderPF)
Dim BL_L_L45: BL_L_L45=Array(LM_L_L45_DT_sw26, LM_L_L45_DT_sw36, LM_L_L45_Layer2, LM_L_L45_Parts, LM_L_L45_Playfield, LM_L_L45_UnderPF)
Dim BL_L_L46: BL_L_L46=Array(LM_L_L46_DT_sw36, LM_L_L46_Parts, LM_L_L46_Playfield, LM_L_L46_ST_sw32, LM_L_L46_UnderPF)
Dim BL_L_L47: BL_L_L47=Array(LM_L_L47_Parts, LM_L_L47_Playfield, LM_L_L47_ST_sw32, LM_L_L47_UnderPF)
Dim BL_L_L5: BL_L_L5=Array(LM_L_L5_Parts, LM_L_L5_Playfield, LM_L_L5_RSling1)
Dim BL_L_L55: BL_L_L55=Array(LM_L_L55_Parts, LM_L_L55_Playfield, LM_L_L55_ST_sw32, LM_L_L55_UnderPF)
Dim BL_L_L56: BL_L_L56=Array(LM_L_L56_Parts, LM_L_L56_Playfield, LM_L_L56_UnderPF)
Dim BL_L_L57: BL_L_L57=Array(LM_L_L57_KT_sw15, LM_L_L57_Parts, LM_L_L57_Playfield, LM_L_L57_UnderPF)
Dim BL_L_L6: BL_L_L6=Array(LM_L_L6_Parts, LM_L_L6_Playfield, LM_L_L6_RSling, LM_L_L6_RSling2, LM_L_L6_UnderPF)
Dim BL_L_L65: BL_L_L65=Array(LM_L_L65_GuardianR, LM_L_L65_Parts, LM_L_L65_Playfield, LM_L_L65_UnderPF)
Dim BL_L_L66: BL_L_L66=Array(LM_L_L66_GuardianR, LM_L_L66_Parts, LM_L_L66_Playfield, LM_L_L66_UnderPF)
Dim BL_L_L67: BL_L_L67=Array(LM_L_L67_GuardianR, LM_L_L67_GuardianR_001, LM_L_L67_Parts, LM_L_L67_Playfield, LM_L_L67_UnderPF)
Dim BL_L_L7: BL_L_L7=Array(LM_L_L7_Parts, LM_L_L7_Playfield, LM_L_L7_RSling, LM_L_L7_RSling1, LM_L_L7_RSling2, LM_L_L7_UnderPF)
Dim BL_L_L71: BL_L_L71=Array(LM_L_L71_Parts, LM_L_L71_Playfield, LM_L_L71_UnderPF)
Dim BL_L_L72: BL_L_L72=Array(LM_L_L72_Parts, LM_L_L72_Playfield, LM_L_L72_UnderPF)
Dim BL_L_L73: BL_L_L73=Array(LM_L_L73_Parts, LM_L_L73_Playfield, LM_L_L73_UnderPF)
Dim BL_L_L74: BL_L_L74=Array(LM_L_L74_Parts, LM_L_L74_Playfield, LM_L_L74_UnderPF)
Dim BL_L_L75: BL_L_L75=Array(LM_L_L75_Parts, LM_L_L75_Playfield, LM_L_L75_UnderPF)
Dim BL_L_L76: BL_L_L76=Array(LM_L_L76_KT_sw16, LM_L_L76_Parts, LM_L_L76_Playfield, LM_L_L76_UnderPF)
Dim BL_L_L77: BL_L_L77=Array(LM_L_L77_Parts, LM_L_L77_Playfield, LM_L_L77_UnderPF)
Dim BL_L_L81: BL_L_L81=Array(LM_L_L81_LSling1, LM_L_L81_Parts, LM_L_L81_Playfield, LM_L_L81_UnderPF)
Dim BL_L_L82: BL_L_L82=Array(LM_L_L82_LF, LM_L_L82_LFU, LM_L_L82_LSling, LM_L_L82_LSling1, LM_L_L82_LSling2, LM_L_L82_Parts, LM_L_L82_Playfield, LM_L_L82_UnderPF)
Dim BL_L_L83: BL_L_L83=Array(LM_L_L83_LF, LM_L_L83_LFU, LM_L_L83_Parts, LM_L_L83_Playfield, LM_L_L83_UnderPF)
Dim BL_L_L84: BL_L_L84=Array(LM_L_L84_Parts, LM_L_L84_Playfield, LM_L_L84_RF, LM_L_L84_RFU, LM_L_L84_UnderPF)
Dim BL_L_L85: BL_L_L85=Array(LM_L_L85_Parts, LM_L_L85_Playfield, LM_L_L85_RF, LM_L_L85_RFU, LM_L_L85_RSling, LM_L_L85_RSling1, LM_L_L85_RSling2, LM_L_L85_UnderPF)
Dim BL_L_L86: BL_L_L86=Array(LM_L_L86_Parts, LM_L_L86_Playfield, LM_L_L86_RSling1, LM_L_L86_RSling2, LM_L_L86_UnderPF)
Dim BL_L_L87: BL_L_L87=Array(LM_L_L87_Parts, LM_L_L87_Playfield, LM_L_L87_UnderPF)
Dim BL_World: BL_World=Array(BM_Autoplunger, BM_Bumper1_Ring, BM_Bumper1_Skirt, BM_Bumper2_Ring, BM_Bumper2_Skirt, BM_DT_sw17, BM_DT_sw26, BM_DT_sw27, BM_DT_sw35, BM_DT_sw36, BM_DT_sw37, BM_Glider_1, BM_Glider_1_001, BM_Glider_1_002, BM_Glider_2, BM_Glider_2_001, BM_GuardianL, BM_GuardianL_001, BM_GuardianR, BM_GuardianR_001, BM_KT_sw14, BM_KT_sw15, BM_KT_sw16, BM_LF, BM_LFU, BM_LSling, BM_LSling1, BM_LSling2, BM_Layer1, BM_Layer2, BM_Parts, BM_PinCab_Rails_021, BM_PinCab_Rails_022, BM_PinCab_Rails_023, BM_Plastics, BM_Playfield, BM_Pyramid1, BM_Pyramid1_001, BM_RF, BM_RF1, BM_RF1U, BM_RFU, BM_RSling, BM_RSling1, BM_RSling2, BM_ST_sw22, BM_ST_sw32, BM_SarcArm, BM_SarcArmU, BM_SideBlades_Art, BM_SlingL, BM_SlingR, BM_UnderPF, BM_Wood_guide___Painted_Blaxk_1, BM_sw111, BM_sw112, BM_sw113, BM_sw114, BM_sw115, BM_sw31)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Autoplunger, BM_Bumper1_Ring, BM_Bumper1_Skirt, BM_Bumper2_Ring, BM_Bumper2_Skirt, BM_DT_sw17, BM_DT_sw26, BM_DT_sw27, BM_DT_sw35, BM_DT_sw36, BM_DT_sw37, BM_Glider_1, BM_Glider_1_001, BM_Glider_1_002, BM_Glider_2, BM_Glider_2_001, BM_GuardianL, BM_GuardianL_001, BM_GuardianR, BM_GuardianR_001, BM_KT_sw14, BM_KT_sw15, BM_KT_sw16, BM_LF, BM_LFU, BM_LSling, BM_LSling1, BM_LSling2, BM_Layer1, BM_Layer2, BM_Parts, BM_PinCab_Rails_021, BM_PinCab_Rails_022, BM_PinCab_Rails_023, BM_Plastics, BM_Playfield, BM_Pyramid1, BM_Pyramid1_001, BM_RF, BM_RF1, BM_RF1U, BM_RFU, BM_RSling, BM_RSling1, BM_RSling2, BM_ST_sw22, BM_ST_sw32, BM_SarcArm, BM_SarcArmU, BM_SideBlades_Art, BM_SlingL, BM_SlingR, BM_UnderPF, BM_Wood_guide___Painted_Blaxk_1, BM_sw111, BM_sw112, BM_sw113, BM_sw114, BM_sw115, BM_sw31)
Dim BG_Lightmap: BG_Lightmap=Array(LM_F_f0_Bumper1_Ring, LM_F_f0_Bumper1_Skirt, LM_F_f0_Bumper2_Ring, LM_F_f0_Bumper2_Skirt, LM_F_f0_KT_sw14, LM_F_f0_Parts, LM_F_f0_Plastics, LM_F_f0_Playfield, LM_F_f0_SideBlades_Art, LM_F_f0_UnderPF, LM_F_f1_DT_sw36, LM_F_f1_Glider_1, LM_F_f1_Glider_1_001, LM_F_f1_Glider_2, LM_F_f1_GuardianL_001, LM_F_f1_GuardianR, LM_F_f1_KT_sw15, LM_F_f1_Layer1, LM_F_f1_Layer2, LM_F_f1_Parts, LM_F_f1_Plastics, LM_F_f1_Playfield, LM_F_f1_Pyramid1, LM_F_f1_RF1U, LM_F_f1_ST_sw32, LM_F_f1_SideBlades_Art, LM_F_f1_UnderPF, LM_F_f2_Glider_1, LM_F_f2_Glider_1_001, LM_F_f2_Glider_1_002, LM_F_f2_Glider_2, LM_F_f2_Glider_2_001, LM_F_f2_GuardianR, LM_F_f2_GuardianR_001, LM_F_f2_KT_sw15, LM_F_f2_Layer1, LM_F_f2_Parts, LM_F_f2_Plastics, LM_F_f2_Playfield, LM_F_f2_Pyramid1, LM_F_f2_Pyramid1_001, LM_F_f2_ST_sw22, LM_F_f2_SideBlades_Art, LM_F_f2_UnderPF, LM_F_f2_Wood_guide___Painted_Bl, LM_F_f3_Glider_1, LM_F_f3_Glider_1_001, LM_F_f3_Glider_1_002, LM_F_f3_Glider_2, LM_F_f3_Glider_2_001, LM_F_f3_GuardianR, _
  LM_F_f3_GuardianR_001, LM_F_f3_Layer2, LM_F_f3_Parts, LM_F_f3_Plastics, LM_F_f3_Playfield, LM_F_f3_Pyramid1, LM_F_f3_Pyramid1_001, LM_F_f3_SideBlades_Art, LM_F_f3_UnderPF, LM_F_f3_Wood_guide___Painted_Bl, LM_F_f4_DT_sw26, LM_F_f4_DT_sw36, LM_F_f4_Glider_1, LM_F_f4_Glider_1_001, LM_F_f4_Glider_1_002, LM_F_f4_Glider_2, LM_F_f4_Glider_2_001, LM_F_f4_GuardianL, LM_F_f4_GuardianL_001, LM_F_f4_GuardianR, LM_F_f4_GuardianR_001, LM_F_f4_KT_sw15, LM_F_f4_LSling1, LM_F_f4_LSling2, LM_F_f4_Layer1, LM_F_f4_Parts, LM_F_f4_Plastics, LM_F_f4_Playfield, LM_F_f4_Pyramid1, LM_F_f4_RF1, LM_F_f4_RF1U, LM_F_f4_ST_sw22, LM_F_f4_ST_sw32, LM_F_f4_UnderPF, LM_F_f4_sw112, LM_F_f5_Bumper2_Ring, LM_F_f5_DT_sw17, LM_F_f5_DT_sw26, LM_F_f5_DT_sw27, LM_F_f5_DT_sw35, LM_F_f5_DT_sw36, LM_F_f5_DT_sw37, LM_F_f5_KT_sw16, LM_F_f5_LFU, LM_F_f5_LSling, LM_F_f5_LSling1, LM_F_f5_LSling2, LM_F_f5_Layer1, LM_F_f5_Layer2, LM_F_f5_Parts, LM_F_f5_Plastics, LM_F_f5_Playfield, LM_F_f5_RF, LM_F_f5_RF1, LM_F_f5_RF1U, LM_F_f5_RFU, LM_F_f5_RSling, _
  LM_F_f5_RSling1, LM_F_f5_RSling2, LM_F_f5_SarcArm, LM_F_f5_SideBlades_Art, LM_F_f5_SlingL, LM_F_f5_UnderPF, LM_F_f5_sw113, LM_GIB_Bumper1_Skirt, LM_GIB_Bumper2_Ring, LM_GIB_Bumper2_Skirt, LM_GIB_DT_sw26, LM_GIB_DT_sw36, LM_GIB_DT_sw37, LM_GIB_GuardianL, LM_GIB_GuardianL_001, LM_GIB_KT_sw14, LM_GIB_KT_sw15, LM_GIB_Layer1, LM_GIB_Parts, LM_GIB_Plastics, LM_GIB_Playfield, LM_GIB_Pyramid1, LM_GIB_SideBlades_Art, LM_GIB_UnderPF, LM_GIG_Bumper1_Ring, LM_GIG_DT_sw17, LM_GIG_DT_sw26, LM_GIG_DT_sw27, LM_GIG_DT_sw36, LM_GIG_DT_sw37, LM_GIG_Glider_1, LM_GIG_Glider_1_002, LM_GIG_GuardianL_001, LM_GIG_GuardianR, LM_GIG_GuardianR_001, LM_GIG_KT_sw15, LM_GIG_LF, LM_GIG_LFU, LM_GIG_LSling, LM_GIG_LSling1, LM_GIG_LSling2, LM_GIG_Layer1, LM_GIG_Layer2, LM_GIG_Parts, LM_GIG_Plastics, LM_GIG_Playfield, LM_GIG_Pyramid1, LM_GIG_RF, LM_GIG_RF1U, LM_GIG_RFU, LM_GIG_RSling, LM_GIG_RSling1, LM_GIG_RSling2, LM_GIG_ST_sw22, LM_GIG_ST_sw32, LM_GIG_SarcArmU, LM_GIG_SideBlades_Art, LM_GIG_SlingL, LM_GIG_SlingR, LM_GIG_UnderPF, _
  LM_GIG_sw111, LM_GIG_sw112, LM_GIG_sw113, LM_GIG_sw114, LM_GIP_GI_001_DT_sw17, LM_GIP_GI_001_DT_sw27, LM_GIP_GI_001_LF, LM_GIP_GI_001_LFU, LM_GIP_GI_001_LSling, LM_GIP_GI_001_LSling1, LM_GIP_GI_001_LSling2, LM_GIP_GI_001_Parts, LM_GIP_GI_001_Plastics, LM_GIP_GI_001_Playfield, LM_GIP_GI_001_RF, LM_GIP_GI_001_RFU, LM_GIP_GI_001_RSling, LM_GIP_GI_001_RSling1, LM_GIP_GI_001_RSling2, LM_GIP_GI_001_SideBlades_Art, LM_GIP_GI_001_SlingR, LM_GIP_GI_001_UnderPF, LM_GIP_GI_002_Bumper1_Ring, LM_GIP_GI_002_DT_sw17, LM_GIP_GI_002_DT_sw37, LM_GIP_GI_002_LFU, LM_GIP_GI_002_LSling, LM_GIP_GI_002_LSling1, LM_GIP_GI_002_LSling2, LM_GIP_GI_002_Layer1, LM_GIP_GI_002_Layer2, LM_GIP_GI_002_Parts, LM_GIP_GI_002_Playfield, LM_GIP_GI_002_RF, LM_GIP_GI_002_RFU, LM_GIP_GI_002_RSling, LM_GIP_GI_002_RSling1, LM_GIP_GI_002_RSling2, LM_GIP_GI_002_SarcArm, LM_GIP_GI_002_SarcArmU, LM_GIP_GI_002_SideBlades_Art, LM_GIP_GI_002_SlingL, LM_GIP_GI_002_UnderPF, LM_GIP_GI_002_sw111, LM_GIP_GI_002_sw112, LM_GIP_GI_006_DT_sw17, LM_GIP_GI_006_DT_sw27, _
  LM_GIP_GI_006_DT_sw37, LM_GIP_GI_006_KT_sw16, LM_GIP_GI_006_Layer1, LM_GIP_GI_006_Parts, LM_GIP_GI_006_Plastics, LM_GIP_GI_006_Playfield, LM_GIP_GI_006_RF1, LM_GIP_GI_006_RF1U, LM_GIP_GI_006_RSling, LM_GIP_GI_006_RSling1, LM_GIP_GI_006_RSling2, LM_GIP_GI_006_SideBlades_Art, LM_GIP_GI_006_UnderPF, LM_GIP_GI_007_Bumper1_Ring, LM_GIP_GI_007_Bumper1_Skirt, LM_GIP_GI_007_Bumper2_Ring, LM_GIP_GI_007_Bumper2_Skirt, LM_GIP_GI_007_DT_sw17, LM_GIP_GI_007_DT_sw26, LM_GIP_GI_007_DT_sw27, LM_GIP_GI_007_DT_sw36, LM_GIP_GI_007_DT_sw37, LM_GIP_GI_007_LFU, LM_GIP_GI_007_LSling, LM_GIP_GI_007_LSling1, LM_GIP_GI_007_LSling2, LM_GIP_GI_007_Layer2, LM_GIP_GI_007_Parts, LM_GIP_GI_007_Plastics, LM_GIP_GI_007_Playfield, LM_GIP_GI_007_SideBlades_Art, LM_GIP_GI_007_UnderPF, LM_GIP_GI_007_sw115, LM_GIR_DT_sw35, LM_GIR_GuardianR, LM_GIR_GuardianR_001, LM_GIR_KT_sw15, LM_GIR_KT_sw16, LM_GIR_Layer1, LM_GIR_Parts, LM_GIR_Plastics, LM_GIR_Playfield, LM_GIR_RF1, LM_GIR_RF1U, LM_GIR_SideBlades_Art, LM_GIR_UnderPF, LM_GIY_Glider_1, _
  LM_GIY_Glider_1_001, LM_GIY_Glider_1_002, LM_GIY_Glider_2, LM_GIY_GuardianR, LM_GIY_KT_sw15, LM_GIY_Parts, LM_GIY_Plastics, LM_GIY_Playfield, LM_GIY_ST_sw22, LM_GIY_ST_sw32, LM_GIY_SideBlades_Art, LM_GIY_UnderPF, LM_L_L0_LF, LM_L_L0_Parts, LM_L_L0_Playfield, LM_L_L0_RF, LM_L_L0_UnderPF, LM_L_L11_Layer2, LM_L_L11_Parts, LM_L_L11_Plastics, LM_L_L11_Playfield, LM_L_L12_LSling, LM_L_L12_LSling1, LM_L_L12_LSling2, LM_L_L12_Layer2, LM_L_L12_Parts, LM_L_L12_Playfield, LM_L_L12_UnderPF, LM_L_L13_Parts, LM_L_L13_Playfield, LM_L_L13_RSling, LM_L_L13_RSling1, LM_L_L13_RSling2, LM_L_L13_UnderPF, LM_L_L14_Parts, LM_L_L14_Playfield, LM_L_L14_UnderPF, LM_L_L15_Bumper1_Skirt, LM_L_L15_DT_sw37, LM_L_L15_Parts, LM_L_L15_Playfield, LM_L_L15_UnderPF, LM_L_L16_Parts, LM_L_L16_Playfield, LM_L_L16_UnderPF, LM_L_L17_Bumper1_Skirt, LM_L_L17_DT_sw26, LM_L_L17_Parts, LM_L_L17_Playfield, LM_L_L17_UnderPF, LM_L_L21_Parts, LM_L_L21_Playfield, LM_L_L21_UnderPF, LM_L_L22_Parts, LM_L_L22_Playfield, LM_L_L22_UnderPF, LM_L_L23_Parts, _
  LM_L_L23_Playfield, LM_L_L23_UnderPF, LM_L_L24_Parts, LM_L_L24_Playfield, LM_L_L24_UnderPF, LM_L_L25_Layer1, LM_L_L25_Parts, LM_L_L25_Playfield, LM_L_L25_RF1, LM_L_L25_RF1U, LM_L_L25_UnderPF, LM_L_L26_Layer1, LM_L_L26_Parts, LM_L_L26_Playfield, LM_L_L26_RF1U, LM_L_L26_UnderPF, LM_L_L27_DT_sw35, LM_L_L27_Layer1, LM_L_L27_Parts, LM_L_L27_Playfield, LM_L_L27_UnderPF, LM_L_L32_Parts, LM_L_L32_Playfield, LM_L_L32_UnderPF, LM_L_L33_DT_sw26, LM_L_L33_Parts, LM_L_L33_Playfield, LM_L_L33_UnderPF, LM_L_L34_DT_sw26, LM_L_L34_DT_sw36, LM_L_L34_Parts, LM_L_L34_Playfield, LM_L_L34_UnderPF, LM_L_L35_Parts, LM_L_L35_Playfield, LM_L_L35_RF1U, LM_L_L36_Parts, LM_L_L36_Playfield, LM_L_L36_UnderPF, LM_L_L37_Parts, LM_L_L37_Playfield, LM_L_L37_UnderPF, LM_L_L41_Bumper1_Ring, LM_L_L41_Bumper1_Skirt, LM_L_L41_Parts, LM_L_L42_Bumper2_Ring, LM_L_L42_Bumper2_Skirt, LM_L_L42_Parts, LM_L_L42_Plastics, LM_L_L42_UnderPF, LM_L_L43_Bumper1_Ring, LM_L_L43_Bumper1_Skirt, LM_L_L43_Layer2, LM_L_L43_Parts, LM_L_L43_Playfield, _
  LM_L_L43_SideBlades_Art, LM_L_L43_UnderPF, LM_L_L44_Bumper2_Ring, LM_L_L44_Bumper2_Skirt, LM_L_L44_GuardianL, LM_L_L44_GuardianL_001, LM_L_L44_KT_sw14, LM_L_L44_Parts, LM_L_L44_Playfield, LM_L_L44_UnderPF, LM_L_L45_DT_sw26, LM_L_L45_DT_sw36, LM_L_L45_Layer2, LM_L_L45_Parts, LM_L_L45_Playfield, LM_L_L45_UnderPF, LM_L_L46_DT_sw36, LM_L_L46_Parts, LM_L_L46_Playfield, LM_L_L46_ST_sw32, LM_L_L46_UnderPF, LM_L_L47_Parts, LM_L_L47_Playfield, LM_L_L47_ST_sw32, LM_L_L47_UnderPF, LM_L_L5_Parts, LM_L_L5_Playfield, LM_L_L5_RSling1, LM_L_L55_Parts, LM_L_L55_Playfield, LM_L_L55_ST_sw32, LM_L_L55_UnderPF, LM_L_L56_Parts, LM_L_L56_Playfield, LM_L_L56_UnderPF, LM_L_L57_KT_sw15, LM_L_L57_Parts, LM_L_L57_Playfield, LM_L_L57_UnderPF, LM_L_L6_Parts, LM_L_L6_Playfield, LM_L_L6_RSling, LM_L_L6_RSling2, LM_L_L6_UnderPF, LM_L_L65_GuardianR, LM_L_L65_Parts, LM_L_L65_Playfield, LM_L_L65_UnderPF, LM_L_L66_GuardianR, LM_L_L66_Parts, LM_L_L66_Playfield, LM_L_L66_UnderPF, LM_L_L67_GuardianR, LM_L_L67_GuardianR_001, LM_L_L67_Parts, _
  LM_L_L67_Playfield, LM_L_L67_UnderPF, LM_L_L7_Parts, LM_L_L7_Playfield, LM_L_L7_RSling, LM_L_L7_RSling1, LM_L_L7_RSling2, LM_L_L7_UnderPF, LM_L_L71_Parts, LM_L_L71_Playfield, LM_L_L71_UnderPF, LM_L_L72_Parts, LM_L_L72_Playfield, LM_L_L72_UnderPF, LM_L_L73_Parts, LM_L_L73_Playfield, LM_L_L73_UnderPF, LM_L_L74_Parts, LM_L_L74_Playfield, LM_L_L74_UnderPF, LM_L_L75_Parts, LM_L_L75_Playfield, LM_L_L75_UnderPF, LM_L_L76_KT_sw16, LM_L_L76_Parts, LM_L_L76_Playfield, LM_L_L76_UnderPF, LM_L_L77_Parts, LM_L_L77_Playfield, LM_L_L77_UnderPF, LM_L_L81_LSling1, LM_L_L81_Parts, LM_L_L81_Playfield, LM_L_L81_UnderPF, LM_L_L82_LF, LM_L_L82_LFU, LM_L_L82_LSling, LM_L_L82_LSling1, LM_L_L82_LSling2, LM_L_L82_Parts, LM_L_L82_Playfield, LM_L_L82_UnderPF, LM_L_L83_LF, LM_L_L83_LFU, LM_L_L83_Parts, LM_L_L83_Playfield, LM_L_L83_UnderPF, LM_L_L84_Parts, LM_L_L84_Playfield, LM_L_L84_RF, LM_L_L84_RFU, LM_L_L84_UnderPF, LM_L_L85_Parts, LM_L_L85_Playfield, LM_L_L85_RF, LM_L_L85_RFU, LM_L_L85_RSling, LM_L_L85_RSling1, LM_L_L85_RSling2, _
  LM_L_L85_UnderPF, LM_L_L86_Parts, LM_L_L86_Playfield, LM_L_L86_RSling1, LM_L_L86_RSling2, LM_L_L86_UnderPF, LM_L_L87_Parts, LM_L_L87_Playfield, LM_L_L87_UnderPF)
Dim BG_All: BG_All=Array(BM_Autoplunger, BM_Bumper1_Ring, BM_Bumper1_Skirt, BM_Bumper2_Ring, BM_Bumper2_Skirt, BM_DT_sw17, BM_DT_sw26, BM_DT_sw27, BM_DT_sw35, BM_DT_sw36, BM_DT_sw37, BM_Glider_1, BM_Glider_1_001, BM_Glider_1_002, BM_Glider_2, BM_Glider_2_001, BM_GuardianL, BM_GuardianL_001, BM_GuardianR, BM_GuardianR_001, BM_KT_sw14, BM_KT_sw15, BM_KT_sw16, BM_LF, BM_LFU, BM_LSling, BM_LSling1, BM_LSling2, BM_Layer1, BM_Layer2, BM_Parts, BM_PinCab_Rails_021, BM_PinCab_Rails_022, BM_PinCab_Rails_023, BM_Plastics, BM_Playfield, BM_Pyramid1, BM_Pyramid1_001, BM_RF, BM_RF1, BM_RF1U, BM_RFU, BM_RSling, BM_RSling1, BM_RSling2, BM_ST_sw22, BM_ST_sw32, BM_SarcArm, BM_SarcArmU, BM_SideBlades_Art, BM_SlingL, BM_SlingR, BM_UnderPF, BM_Wood_guide___Painted_Blaxk_1, BM_sw111, BM_sw112, BM_sw113, BM_sw114, BM_sw115, BM_sw31, LM_F_f0_Bumper1_Ring, LM_F_f0_Bumper1_Skirt, LM_F_f0_Bumper2_Ring, LM_F_f0_Bumper2_Skirt, LM_F_f0_KT_sw14, LM_F_f0_Parts, LM_F_f0_Plastics, LM_F_f0_Playfield, LM_F_f0_SideBlades_Art, LM_F_f0_UnderPF, _
  LM_F_f1_DT_sw36, LM_F_f1_Glider_1, LM_F_f1_Glider_1_001, LM_F_f1_Glider_2, LM_F_f1_GuardianL_001, LM_F_f1_GuardianR, LM_F_f1_KT_sw15, LM_F_f1_Layer1, LM_F_f1_Layer2, LM_F_f1_Parts, LM_F_f1_Plastics, LM_F_f1_Playfield, LM_F_f1_Pyramid1, LM_F_f1_RF1U, LM_F_f1_ST_sw32, LM_F_f1_SideBlades_Art, LM_F_f1_UnderPF, LM_F_f2_Glider_1, LM_F_f2_Glider_1_001, LM_F_f2_Glider_1_002, LM_F_f2_Glider_2, LM_F_f2_Glider_2_001, LM_F_f2_GuardianR, LM_F_f2_GuardianR_001, LM_F_f2_KT_sw15, LM_F_f2_Layer1, LM_F_f2_Parts, LM_F_f2_Plastics, LM_F_f2_Playfield, LM_F_f2_Pyramid1, LM_F_f2_Pyramid1_001, LM_F_f2_ST_sw22, LM_F_f2_SideBlades_Art, LM_F_f2_UnderPF, LM_F_f2_Wood_guide___Painted_Bl, LM_F_f3_Glider_1, LM_F_f3_Glider_1_001, LM_F_f3_Glider_1_002, LM_F_f3_Glider_2, LM_F_f3_Glider_2_001, LM_F_f3_GuardianR, LM_F_f3_GuardianR_001, LM_F_f3_Layer2, LM_F_f3_Parts, LM_F_f3_Plastics, LM_F_f3_Playfield, LM_F_f3_Pyramid1, LM_F_f3_Pyramid1_001, LM_F_f3_SideBlades_Art, LM_F_f3_UnderPF, LM_F_f3_Wood_guide___Painted_Bl, LM_F_f4_DT_sw26, _
  LM_F_f4_DT_sw36, LM_F_f4_Glider_1, LM_F_f4_Glider_1_001, LM_F_f4_Glider_1_002, LM_F_f4_Glider_2, LM_F_f4_Glider_2_001, LM_F_f4_GuardianL, LM_F_f4_GuardianL_001, LM_F_f4_GuardianR, LM_F_f4_GuardianR_001, LM_F_f4_KT_sw15, LM_F_f4_LSling1, LM_F_f4_LSling2, LM_F_f4_Layer1, LM_F_f4_Parts, LM_F_f4_Plastics, LM_F_f4_Playfield, LM_F_f4_Pyramid1, LM_F_f4_RF1, LM_F_f4_RF1U, LM_F_f4_ST_sw22, LM_F_f4_ST_sw32, LM_F_f4_UnderPF, LM_F_f4_sw112, LM_F_f5_Bumper2_Ring, LM_F_f5_DT_sw17, LM_F_f5_DT_sw26, LM_F_f5_DT_sw27, LM_F_f5_DT_sw35, LM_F_f5_DT_sw36, LM_F_f5_DT_sw37, LM_F_f5_KT_sw16, LM_F_f5_LFU, LM_F_f5_LSling, LM_F_f5_LSling1, LM_F_f5_LSling2, LM_F_f5_Layer1, LM_F_f5_Layer2, LM_F_f5_Parts, LM_F_f5_Plastics, LM_F_f5_Playfield, LM_F_f5_RF, LM_F_f5_RF1, LM_F_f5_RF1U, LM_F_f5_RFU, LM_F_f5_RSling, LM_F_f5_RSling1, LM_F_f5_RSling2, LM_F_f5_SarcArm, LM_F_f5_SideBlades_Art, LM_F_f5_SlingL, LM_F_f5_UnderPF, LM_F_f5_sw113, LM_GIB_Bumper1_Skirt, LM_GIB_Bumper2_Ring, LM_GIB_Bumper2_Skirt, LM_GIB_DT_sw26, LM_GIB_DT_sw36, LM_GIB_DT_sw37, _
  LM_GIB_GuardianL, LM_GIB_GuardianL_001, LM_GIB_KT_sw14, LM_GIB_KT_sw15, LM_GIB_Layer1, LM_GIB_Parts, LM_GIB_Plastics, LM_GIB_Playfield, LM_GIB_Pyramid1, LM_GIB_SideBlades_Art, LM_GIB_UnderPF, LM_GIG_Bumper1_Ring, LM_GIG_DT_sw17, LM_GIG_DT_sw26, LM_GIG_DT_sw27, LM_GIG_DT_sw36, LM_GIG_DT_sw37, LM_GIG_Glider_1, LM_GIG_Glider_1_002, LM_GIG_GuardianL_001, LM_GIG_GuardianR, LM_GIG_GuardianR_001, LM_GIG_KT_sw15, LM_GIG_LF, LM_GIG_LFU, LM_GIG_LSling, LM_GIG_LSling1, LM_GIG_LSling2, LM_GIG_Layer1, LM_GIG_Layer2, LM_GIG_Parts, LM_GIG_Plastics, LM_GIG_Playfield, LM_GIG_Pyramid1, LM_GIG_RF, LM_GIG_RF1U, LM_GIG_RFU, LM_GIG_RSling, LM_GIG_RSling1, LM_GIG_RSling2, LM_GIG_ST_sw22, LM_GIG_ST_sw32, LM_GIG_SarcArmU, LM_GIG_SideBlades_Art, LM_GIG_SlingL, LM_GIG_SlingR, LM_GIG_UnderPF, LM_GIG_sw111, LM_GIG_sw112, LM_GIG_sw113, LM_GIG_sw114, LM_GIP_GI_001_DT_sw17, LM_GIP_GI_001_DT_sw27, LM_GIP_GI_001_LF, LM_GIP_GI_001_LFU, LM_GIP_GI_001_LSling, LM_GIP_GI_001_LSling1, LM_GIP_GI_001_LSling2, LM_GIP_GI_001_Parts, _
  LM_GIP_GI_001_Plastics, LM_GIP_GI_001_Playfield, LM_GIP_GI_001_RF, LM_GIP_GI_001_RFU, LM_GIP_GI_001_RSling, LM_GIP_GI_001_RSling1, LM_GIP_GI_001_RSling2, LM_GIP_GI_001_SideBlades_Art, LM_GIP_GI_001_SlingR, LM_GIP_GI_001_UnderPF, LM_GIP_GI_002_Bumper1_Ring, LM_GIP_GI_002_DT_sw17, LM_GIP_GI_002_DT_sw37, LM_GIP_GI_002_LFU, LM_GIP_GI_002_LSling, LM_GIP_GI_002_LSling1, LM_GIP_GI_002_LSling2, LM_GIP_GI_002_Layer1, LM_GIP_GI_002_Layer2, LM_GIP_GI_002_Parts, LM_GIP_GI_002_Playfield, LM_GIP_GI_002_RF, LM_GIP_GI_002_RFU, LM_GIP_GI_002_RSling, LM_GIP_GI_002_RSling1, LM_GIP_GI_002_RSling2, LM_GIP_GI_002_SarcArm, LM_GIP_GI_002_SarcArmU, LM_GIP_GI_002_SideBlades_Art, LM_GIP_GI_002_SlingL, LM_GIP_GI_002_UnderPF, LM_GIP_GI_002_sw111, LM_GIP_GI_002_sw112, LM_GIP_GI_006_DT_sw17, LM_GIP_GI_006_DT_sw27, LM_GIP_GI_006_DT_sw37, LM_GIP_GI_006_KT_sw16, LM_GIP_GI_006_Layer1, LM_GIP_GI_006_Parts, LM_GIP_GI_006_Plastics, LM_GIP_GI_006_Playfield, LM_GIP_GI_006_RF1, LM_GIP_GI_006_RF1U, LM_GIP_GI_006_RSling, LM_GIP_GI_006_RSling1, _
  LM_GIP_GI_006_RSling2, LM_GIP_GI_006_SideBlades_Art, LM_GIP_GI_006_UnderPF, LM_GIP_GI_007_Bumper1_Ring, LM_GIP_GI_007_Bumper1_Skirt, LM_GIP_GI_007_Bumper2_Ring, LM_GIP_GI_007_Bumper2_Skirt, LM_GIP_GI_007_DT_sw17, LM_GIP_GI_007_DT_sw26, LM_GIP_GI_007_DT_sw27, LM_GIP_GI_007_DT_sw36, LM_GIP_GI_007_DT_sw37, LM_GIP_GI_007_LFU, LM_GIP_GI_007_LSling, LM_GIP_GI_007_LSling1, LM_GIP_GI_007_LSling2, LM_GIP_GI_007_Layer2, LM_GIP_GI_007_Parts, LM_GIP_GI_007_Plastics, LM_GIP_GI_007_Playfield, LM_GIP_GI_007_SideBlades_Art, LM_GIP_GI_007_UnderPF, LM_GIP_GI_007_sw115, LM_GIR_DT_sw35, LM_GIR_GuardianR, LM_GIR_GuardianR_001, LM_GIR_KT_sw15, LM_GIR_KT_sw16, LM_GIR_Layer1, LM_GIR_Parts, LM_GIR_Plastics, LM_GIR_Playfield, LM_GIR_RF1, LM_GIR_RF1U, LM_GIR_SideBlades_Art, LM_GIR_UnderPF, LM_GIY_Glider_1, LM_GIY_Glider_1_001, LM_GIY_Glider_1_002, LM_GIY_Glider_2, LM_GIY_GuardianR, LM_GIY_KT_sw15, LM_GIY_Parts, LM_GIY_Plastics, LM_GIY_Playfield, LM_GIY_ST_sw22, LM_GIY_ST_sw32, LM_GIY_SideBlades_Art, LM_GIY_UnderPF, LM_L_L0_LF, _
  LM_L_L0_Parts, LM_L_L0_Playfield, LM_L_L0_RF, LM_L_L0_UnderPF, LM_L_L11_Layer2, LM_L_L11_Parts, LM_L_L11_Plastics, LM_L_L11_Playfield, LM_L_L12_LSling, LM_L_L12_LSling1, LM_L_L12_LSling2, LM_L_L12_Layer2, LM_L_L12_Parts, LM_L_L12_Playfield, LM_L_L12_UnderPF, LM_L_L13_Parts, LM_L_L13_Playfield, LM_L_L13_RSling, LM_L_L13_RSling1, LM_L_L13_RSling2, LM_L_L13_UnderPF, LM_L_L14_Parts, LM_L_L14_Playfield, LM_L_L14_UnderPF, LM_L_L15_Bumper1_Skirt, LM_L_L15_DT_sw37, LM_L_L15_Parts, LM_L_L15_Playfield, LM_L_L15_UnderPF, LM_L_L16_Parts, LM_L_L16_Playfield, LM_L_L16_UnderPF, LM_L_L17_Bumper1_Skirt, LM_L_L17_DT_sw26, LM_L_L17_Parts, LM_L_L17_Playfield, LM_L_L17_UnderPF, LM_L_L21_Parts, LM_L_L21_Playfield, LM_L_L21_UnderPF, LM_L_L22_Parts, LM_L_L22_Playfield, LM_L_L22_UnderPF, LM_L_L23_Parts, LM_L_L23_Playfield, LM_L_L23_UnderPF, LM_L_L24_Parts, LM_L_L24_Playfield, LM_L_L24_UnderPF, LM_L_L25_Layer1, LM_L_L25_Parts, LM_L_L25_Playfield, LM_L_L25_RF1, LM_L_L25_RF1U, LM_L_L25_UnderPF, LM_L_L26_Layer1, LM_L_L26_Parts, _
  LM_L_L26_Playfield, LM_L_L26_RF1U, LM_L_L26_UnderPF, LM_L_L27_DT_sw35, LM_L_L27_Layer1, LM_L_L27_Parts, LM_L_L27_Playfield, LM_L_L27_UnderPF, LM_L_L32_Parts, LM_L_L32_Playfield, LM_L_L32_UnderPF, LM_L_L33_DT_sw26, LM_L_L33_Parts, LM_L_L33_Playfield, LM_L_L33_UnderPF, LM_L_L34_DT_sw26, LM_L_L34_DT_sw36, LM_L_L34_Parts, LM_L_L34_Playfield, LM_L_L34_UnderPF, LM_L_L35_Parts, LM_L_L35_Playfield, LM_L_L35_RF1U, LM_L_L36_Parts, LM_L_L36_Playfield, LM_L_L36_UnderPF, LM_L_L37_Parts, LM_L_L37_Playfield, LM_L_L37_UnderPF, LM_L_L41_Bumper1_Ring, LM_L_L41_Bumper1_Skirt, LM_L_L41_Parts, LM_L_L42_Bumper2_Ring, LM_L_L42_Bumper2_Skirt, LM_L_L42_Parts, LM_L_L42_Plastics, LM_L_L42_UnderPF, LM_L_L43_Bumper1_Ring, LM_L_L43_Bumper1_Skirt, LM_L_L43_Layer2, LM_L_L43_Parts, LM_L_L43_Playfield, LM_L_L43_SideBlades_Art, LM_L_L43_UnderPF, LM_L_L44_Bumper2_Ring, LM_L_L44_Bumper2_Skirt, LM_L_L44_GuardianL, LM_L_L44_GuardianL_001, LM_L_L44_KT_sw14, LM_L_L44_Parts, LM_L_L44_Playfield, LM_L_L44_UnderPF, LM_L_L45_DT_sw26, LM_L_L45_DT_sw36, _
  LM_L_L45_Layer2, LM_L_L45_Parts, LM_L_L45_Playfield, LM_L_L45_UnderPF, LM_L_L46_DT_sw36, LM_L_L46_Parts, LM_L_L46_Playfield, LM_L_L46_ST_sw32, LM_L_L46_UnderPF, LM_L_L47_Parts, LM_L_L47_Playfield, LM_L_L47_ST_sw32, LM_L_L47_UnderPF, LM_L_L5_Parts, LM_L_L5_Playfield, LM_L_L5_RSling1, LM_L_L55_Parts, LM_L_L55_Playfield, LM_L_L55_ST_sw32, LM_L_L55_UnderPF, LM_L_L56_Parts, LM_L_L56_Playfield, LM_L_L56_UnderPF, LM_L_L57_KT_sw15, LM_L_L57_Parts, LM_L_L57_Playfield, LM_L_L57_UnderPF, LM_L_L6_Parts, LM_L_L6_Playfield, LM_L_L6_RSling, LM_L_L6_RSling2, LM_L_L6_UnderPF, LM_L_L65_GuardianR, LM_L_L65_Parts, LM_L_L65_Playfield, LM_L_L65_UnderPF, LM_L_L66_GuardianR, LM_L_L66_Parts, LM_L_L66_Playfield, LM_L_L66_UnderPF, LM_L_L67_GuardianR, LM_L_L67_GuardianR_001, LM_L_L67_Parts, LM_L_L67_Playfield, LM_L_L67_UnderPF, LM_L_L7_Parts, LM_L_L7_Playfield, LM_L_L7_RSling, LM_L_L7_RSling1, LM_L_L7_RSling2, LM_L_L7_UnderPF, LM_L_L71_Parts, LM_L_L71_Playfield, LM_L_L71_UnderPF, LM_L_L72_Parts, LM_L_L72_Playfield, LM_L_L72_UnderPF, _
  LM_L_L73_Parts, LM_L_L73_Playfield, LM_L_L73_UnderPF, LM_L_L74_Parts, LM_L_L74_Playfield, LM_L_L74_UnderPF, LM_L_L75_Parts, LM_L_L75_Playfield, LM_L_L75_UnderPF, LM_L_L76_KT_sw16, LM_L_L76_Parts, LM_L_L76_Playfield, LM_L_L76_UnderPF, LM_L_L77_Parts, LM_L_L77_Playfield, LM_L_L77_UnderPF, LM_L_L81_LSling1, LM_L_L81_Parts, LM_L_L81_Playfield, LM_L_L81_UnderPF, LM_L_L82_LF, LM_L_L82_LFU, LM_L_L82_LSling, LM_L_L82_LSling1, LM_L_L82_LSling2, LM_L_L82_Parts, LM_L_L82_Playfield, LM_L_L82_UnderPF, LM_L_L83_LF, LM_L_L83_LFU, LM_L_L83_Parts, LM_L_L83_Playfield, LM_L_L83_UnderPF, LM_L_L84_Parts, LM_L_L84_Playfield, LM_L_L84_RF, LM_L_L84_RFU, LM_L_L84_UnderPF, LM_L_L85_Parts, LM_L_L85_Playfield, LM_L_L85_RF, LM_L_L85_RFU, LM_L_L85_RSling, LM_L_L85_RSling1, LM_L_L85_RSling2, LM_L_L85_UnderPF, LM_L_L86_Parts, LM_L_L86_Playfield, LM_L_L86_RSling1, LM_L_L86_RSling2, LM_L_L86_UnderPF, LM_L_L87_Parts, LM_L_L87_Playfield, LM_L_L87_UnderPF)
' VLM  Arrays - End
