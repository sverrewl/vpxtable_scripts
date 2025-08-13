' Based on Williams's Cyclone / IPD No. 617 / 1988 / 4 Players

' VP9 - Original version by melon - 2009.
' VPX - Original version by Sliderpoint - 2017.
' VR-Hybrid version by ClarkKent, RobbyKingPin, DarthVito, Flupper, TastyWasps, hauntfreaks, DGrimmReaper, apophis, chokeee & LoadedWeapon -2024/2025.
' Represented in chronological order of people who entered on this project.
'
' This mod contains the following improvements:
'
' ClarkKent (The Operator and Project Starter):
' -Added new playfield texture with the proper scans, rebuild the entire table to match
' -Added new plastic textures and made small changes to them in the editor
' -Giving assistance to ensure everything is pixel perfect as usual
' -Providing acces to a full library of images with the higest details I have seen yet so far
' -Corrected the ramps and ramptextures and all corresponding rivets/screws and added more screws on the backpanel. Also corrected the ramp protectors to look a little bit more exact
' -Removed the decal from the jackpot texture on the ramp and added a new one separately. Made the Cometramp primitive collidable
' -Constructed nice scoop primitives for the spookhouse and the boomerang
'
' RobbyKingPin (Ride Mechanic and Project Leader):
' -nFozzy/Rothbauerw physics and Fleep 2.0 sound
' -Created new backdrops for desktop and added the alphanumeric displays from TastyWasps' Banzai Run and changed the color settings to recommendations from hauntfreaks
' -Added new mysterywheel and jackpot lights on the backdrop, based on the mechcodes from JPSalas
' -Added Flupper's flipperbats from White Water and changed the texture, also changed light settings to match Flupper's settings
' -Changed the through, made it simular to VPW's Bad Cats
' -Improved all lights and normal flashers, copied insert lights and GI lights from JPSalas and changed the settings on the GI lights to match VPW's settings and added some missing lights
' -Fixed the VR mysterywheel wich was a big pain in the a** for the entire VR A-team on this project and repositioned the VR Digits of the bottom displays
' -Created new textures for the comet lights and a texture for the siderails and lockdownbar and put a skin on the primitive. New instruction cards made. New siderails images provided by Cliffy AKA passion4pins
' -Repositioned and changed the height on the biff bars and also repositioned the rubber peg in between the flippers
' -Fixed some missing switches
' -Added Flupper Domes with edited primitives & Flupper Bumpers, converted 2 Flupper Inserts to work with the Flupper Domes code and created a brand new Flupper Flasherlamps code for the primitives underneath the plastics
' -Edited the ferriswheel, Pincab_cabinet, backbox, rails and coindoor primitives to look like a System 11 cabinet and created a new speakerpanel with glass.
' -Added Rothbauerw standup and droptarget codes with help from Rothauerw  & Apophis to get everything working as it should.
' -Created a primitive for the Skillshot ramp and made the primitive collidable, all new plastics and droptarget created
' -Updated Rubber Dampeners (v.2.2)
' -VPW Ambient Ball Shadow added to replace the old VPW Dynamic Ball Shadow (v.2.2)
' -Added new custom sideblades and edited the primitive (v.2.2)
' -VPW Tweak Menu added to replace the old VPW Options UI (v.2.2)
' -Resized flippers matching settings provided by apophis, reworked flipper triggers (v.2.2)
' -Added more options for reflections and Mixed Reality. (v.2.3)
' -Added new VR backglass using hauntfreaks assets (v.2.3)
'
' DarthVito (Cotton Candy and VR Wizard):
' -Added VR models, used assets from Senseless
' -Added a nice VR Game room with arcade machines and the new minimal room provided by hauntfreaks
' -Added new VR background with elite looking framed posters
'
' Flupper (Track Master and Light Settings Expert):
' -Created a more blurred environment image, more friendly (less jaggies) for desktop
' -Lowered the "amount" in the "shadows" material
' -Changed the ramp textures and ramp materials a bit and added refraction, changed the blenddisablelightings for the ramps and set the refraction to 8
' -Made a primitive and a texture for the shooter wireramp and biff bars
'
' TastyWasps (Fireworks Lead and VR Grandmaster):
' -Added VR backglass functionality from Mike Da Spike's model and lots of flasher placements
' -Backbox changed to a proper flat face so that the UV art is not distorted from 3D effects
' -Animated the plunger and flipper buttons
' -VR Cab graphics from hauntfreaks installed
' -On/Off VR Backglass from HauntFreaks installed
' -VR flashers / digits repositioned for new VR assets
'
' hauntfreaks (Imagineer and VR Cabinet Textures and Backglass Specialist):
' -Reworked all VR cabinet textures
' -Made a new set for on/off VR Backglass
' -Manually moved all the flasher on the VR BG to their correct positions
' -Help with making the Flupper Bumpers look like red instead of orange
' -Provided a second backdrop with some cool wallpaper
'
' DGrimmReaper (Did Something Amazing, The VR Mechanic):
' -Added PinCab_Backbox_Neck primitive
' -Fixed arcade machines and removed primitive4 which was a moon in senseless room
' -Fixed Backbox and Backglass alignment
' -Fixed GameRoom Floor and Arcade heights
'
' apophis (Some dude, but in reality a Lifesaver on this project!):
' -Added codes for PWM on the GI, flashers and inserts, without a doubt the hardest part of this project!
'
' chokeee (Confections and Special Effects Enhancer):
' -Added Flupper Inserts controlled by nFozzy Lampz
' -Added inserts decals, tweaked inserts, changed rubbers and white screws material, changed intensity of some GI lights.file
' -Added extra flashers for wall reflections and changed some metal materials
'
' LoadedWeapon (Assistant Mechanic):
' -Created a custom sideblades primitive to fit on Legends Cabinets and provided a secondary ini file for the POV to match
'
' Special thanks to everybody involved during development and testing: Smaug, PinStratsDan, Primetime5K, Apophis, Sixtoe, Bord, MikeDASpike, Tomate, wrd1972, Studlygoorite, Sinizin
' Shoutout to:
' -apophis for saving my b*tt big time once again, Smaug, Cliffy & Rothbauerw
' -melon, Sliderpoint, Flupper, Wildman, Westworld & hauntfreaks for their work on previous versions
'
' If I didn't mention anybody else in this development cycle I am unaware of this and just know that I still think you are awesome and also want to say you rock!
' And basically everybody from VPW I didn't mention just for having me as part of the family, you all absolutely rock!
'
Option Explicit
Randomize
SetLocale 1033

' Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SCoin = ""
Const BallSize = 50        'Ball diameter in VPX units; must be 50
Const BallMass = 1          'Ball mass must be 1
Const tnob = 1            'Total number of balls the table can hold
Const lob = 0

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

'//  PositionalSoundPlaybackConfiguration:
'//  Specifies the sound playback configuration. Options:
'//  1 = Mono
'//  2 = Stereo (Only L+R channels)
'//  3 = Surround (Full surround sound support for SSF - Front L + Front R + Rear L + Rear R channels)
Const PositionalSoundPlaybackConfiguration = 3

'//  RelaysPosition:
'//  1 = Relays positioned with power board (Provides sound spread from the left and right back channels)
'//  2 = Relays positioned with GI strings (Provides sound spread from left/right/front/rear surround channels)
Const RelaysPosition = 2

'************************************************************************
Const DebugFlashers = False
Const DebugGI = False

Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height
Dim BIP             ' Balls in play
BIP = 0
Dim BIPL            ' Ball in plunger lane
BIPL = False
Dim plungerpress

'**************************************************************
'                    USER OPTIONS
'**************************************************************

'----- VR Options -----
Dim VRRoomChoice : VRRoomChoice = 1       ' 1 = Game Room 2 = Minimal Room
Const VRTest = 0        ' 1 = Testing VR in Live View, 0 = Do not force VR mode.
Const MysteryWheelDT = 0 'Set to 1 to make the Mystery wheel visible on desktop. Don't use a b2s file when enabled

Const cGameName = "cycln_l5"

Dim ModSol : ModSol = 2
Dim UseVPMModSol

Dim op: op = LoadValue(cGameName, "MODSOL")
If op <> "" Then ModSol = CInt(op):  Else ModSol = 2: End If

If ModSol = 1 Then
  UseVPMModSol = 0
  BumperTimer.Enabled = True
Else
  UseVPMModSol = 2
  BumperTimer.Enabled = False
End If

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 1
Const UseSync = 0
Const HandleMech = 0

LoadVPM "03060000", "S11.vbs", 3.26

Dim CyclnBall1, gBOT

Dim DesktopMode: DesktopMode = Table1.ShowDT

Dim VarHidden, x

If Table1.ShowDT = True And renderingmode <> 2 Then
  For each x in aReels
    x.Visible = 1
  Next
  VarHidden = 1
  PinCab_Rails.Visible = 1
  PinCab_Blades.Visible = 1
  SideBladesLegends.Visible = 0
Else
    For each x in aReels
        x.Visible = 0
    Next
    VarHidden = 0
  PinCab_Rails.Visible = 0
  If Table1.ShowFSS = False And LegendsCabMod = 1 Then
    PinCab_Blades.Visible = 0
    SideBladesLegends.Visible = 1
    Flasherbloom1.Height = 500
    Flasherbloom2.Height = 500
    Flasherbloom3.Height = 500
    Flasherbloom4.Height = 500
    Flasherbloom5.Height = 500
    Flasherbloom6.Height = 500
    Flasherbloom7.Height = 500
    Flasherlampbloom1. Height = 500
    Flasherlampbloom2. Height = 500
    Flasherlampbloom3. Height = 500
    Flasherlampbloom4. Height = 500
    If SideBladeMod = 0 Then
      SideBladesLegends.image = "Sidewalls_C_Original"
    Else
      SideBladesLegends.image = "Sidewalls_C"
    End If
  End If
End If

if B2SOn = true then VarHidden = 1

'**********************************************************************************************************
'VR Room Animations
'**********************************************************************************************************

Dim VRCounter
VRCounter = 1

Dim VR2Counter
VR2Counter = 1

Dim VR3Counter
VR3Counter = 1

Sub TimerAnimateCard1_Timer()
  Arcade_screen.Image = "EA-" & VRCounter
  VRCounter = VRCounter + 1
  If VRCounter > 7 Then
    VRCounter = 1
  End If
End Sub

Sub TimerAnimateCard2_Timer()
  Arcade2_screen.Image = "DK " & VR2Counter
  VR2Counter = VR2Counter + 1
  If VR2Counter > 21 Then
    VR2Counter = 1
  End If
End Sub

Sub TimerAnimateCard3_Timer()
  Arcade3_screen.Image = "MSP-" & VR3Counter
  VR3Counter = VR3Counter + 1
  If VR3Counter > 5 Then
    VR3Counter = 1
  End If
End Sub

'----- VR Room Auto-Detect -----
Dim VR_Obj, VR_Room, VRMode

If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
  VRMode = True
  Topper_Backlight.Visible = 1
  PinCab_Rails.Visible = 1
  For Each x in aReels : x.Visible = 0 : Next
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in VRLedDigits : VR_Obj.Visible = 1 : Next
  If VRRoomChoice = 1 Then
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRGameRoom : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in DTMysteryWheel : VR_Obj.Visible = 0 : Next
    Room360.Visible = 0
  End If
  If VRRoomChoice = 2 Then
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in VRGameRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in DTMysteryWheel : VR_Obj.Visible = 0 : Next
    Room360.Visible = 0
  End If
  If VRRoomChoice = 3 Then
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRGameRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in DTMysteryWheel : VR_Obj.Visible = 0 : Next
    Room360.Visible = 1
  End If
Else
  VRMode = False
  Topper_Backlight.Visible = 0
  Room360.Visible = 0
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRGameRoom : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRLedDigits : VR_Obj.Visible = 0 : Next
  If MysteryWheelDT = 1 Then
    For Each VR_Obj in DTMysteryWheel : VR_Obj.Visible = 1 : Next
  Else
    For Each VR_Obj in DTMysteryWheel : VR_Obj.Visible = 0 : Next
  End If
End If

'*******************************************
' ZTIM:Timers
'*******************************************

Dim FrameTime, InitFrameTime
InitFrameTime = 0

Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  FrameTime = GameTime - InitFrameTime
  InitFrameTime = GameTime  'Count frametime
  RollingUpdate     ' Update rolling sounds
  DoDTAnim    'handle drop target animations
  DoSTAnim    'handle stand up target animations
  If AmbientBallShadowOn = 1 Then
    BSUpdate
  End If
  UpdateBallBrightness
  UpdatePlunger         'VR plunger updates
  Flasherflash1a.opacity = Flasherflash1.opacity
  Flasherflash1b.opacity = Flasherflash1.opacity
  Flasherflash1c.opacity = Flasherflash1.opacity
  Flasherflash1d.opacity = Flasherflash1.opacity
' Flasherflash1e.opacity = Flasherflash1.opacity
' Flasherflash1f.opacity = Flasherflash1.opacity
  Flasherflash2a.opacity = Flasherflash2.opacity
  Flasherflash2b.opacity = Flasherflash2.opacity
  Flasherflash3a.opacity = Flasherflash3.opacity
  Flasherflash3b.opacity = Flasherflash3.opacity
  Flasherflash4a.opacity = Flasherflash4.opacity
  Flasherflash4b.opacity = Flasherflash4.opacity
  Flasherflash4c.opacity = Flasherflash4.opacity
  Flasherflash4d.opacity = Flasherflash4.opacity
  Flasherflash5a.opacity = Flasherflash5.opacity
  Flasherflash6a.opacity = Flasherflash6.opacity
  Flasherflash7a.opacity = Flasherflash7.opacity
  Flasherlampflash1a.opacity = Flasherlampflash1.opacity
  Flasherlampflash1b.opacity = Flasherlampflash1.opacity
  Flasherlampflash1c.opacity = Flasherlampflash1.opacity
  Flasherlampflash4a.opacity = Flasherlampflash4.opacity

  Flasherflash1a.visible = Flasherflash1.visible
  Flasherflash1b.visible = Flasherflash1.visible
  Flasherflash1c.visible = Flasherflash1.visible
  Flasherflash1d.visible = Flasherflash1.visible
' Flasherflash1e.visible = Flasherflash1.visible
' Flasherflash1f.visible = Flasherflash1.visible
  Flasherflash2a.visible = Flasherflash2.visible
  Flasherflash2b.visible = Flasherflash2.visible
  Flasherflash3a.visible = Flasherflash3.visible
  Flasherflash3b.visible = Flasherflash3.visible
  Flasherflash4a.visible = Flasherflash4.visible
  Flasherflash4b.visible = Flasherflash4.visible
  Flasherflash4c.visible = Flasherflash4.visible
  Flasherflash4d.visible = Flasherflash4.visible
  Flasherflash5a.visible = Flasherflash5.visible
  Flasherflash6a.visible = Flasherflash6.visible
  Flasherflash7a.visible = Flasherflash7.visible
  Flasherlampflash1a.visible = Flasherlampflash1.visible
  Flasherlampflash1b.visible = Flasherlampflash1.visible
  Flasherlampflash1c.visible = Flasherlampflash1.visible
  Flasherlampflash4a.visible = Flasherlampflash4.visible
  If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
    VRDisplayTimer
    Pincab_Backglass.Visible = 1
    If Pincab_BackglassL33.BlendDisableLighting  < 0.05 Then : Pincab_BackglassL33.Visible = 0 : Else : Pincab_BackglassL33.Visible = 1
    If Pincab_BackglassL34.BlendDisableLighting  < 0.05 Then : Pincab_BackglassL34.Visible = 0 : Else : Pincab_BackglassL34.Visible = 1
    If Pincab_BackglassL35.BlendDisableLighting  < 0.05 Then : Pincab_BackglassL35.Visible = 0 : Else : Pincab_BackglassL35.Visible = 1
    If Pincab_BackglassL36.BlendDisableLighting < 0.05 Then : Pincab_BackglassL36.Visible = 0 : Else : Pincab_BackglassL36.Visible = 1
    If Pincab_BackglassL37.BlendDisableLighting < 0.05 Then : Pincab_BackglassL37.Visible = 0 : Else : Pincab_BackglassL37.Visible = 1
    If Pincab_BackglassL38.BlendDisableLighting < 0.05 Then : Pincab_BackglassL38.Visible = 0 : Else : Pincab_BackglassL38.Visible = 1
    If Pincab_BackglassL39.BlendDisableLighting < 0.05 Then : Pincab_BackglassL39.Visible = 0 : Else : Pincab_BackglassL39.Visible = 1
    If Pincab_BackglassL40.BlendDisableLighting < 0.05 Then : Pincab_BackglassL40.Visible = 0 : Else : Pincab_BackglassL40.Visible = 1
    If Pincab_BackglassL41.BlendDisableLighting < 0.05 Then : Pincab_BackglassL41.Visible = 0 : Else : Pincab_BackglassL41.Visible = 1
    If Pincab_BackglassL42.BlendDisableLighting < 0.05 Then : Pincab_BackglassL42.Visible = 0 : Else : Pincab_BackglassL42.Visible = 1
    If Pincab_BackglassL43.BlendDisableLighting < 0.05 Then : Pincab_BackglassL43.Visible = 0 : Else : Pincab_BackglassL43.Visible = 1
    If Pincab_BackglassL44.BlendDisableLighting < 0.05 Then : Pincab_BackglassL44.Visible = 0 : Else : Pincab_BackglassL44.Visible = 1
    If Pincab_BackglassL45.BlendDisableLighting < 0.05 Then : Pincab_BackglassL45.Visible = 0 : Else : Pincab_BackglassL45.Visible = 1
    If Pincab_BackglassL46.BlendDisableLighting < 0.05 Then : Pincab_BackglassL46.Visible = 0 : Else : Pincab_BackglassL46.Visible = 1
    If Pincab_BackglassL47.BlendDisableLighting < 0.05 Then : Pincab_BackglassL47.Visible = 0 : Else : Pincab_BackglassL47.Visible = 1
    If Pincab_BackglassL48.BlendDisableLighting < 0.05 Then : Pincab_BackglassL48.Visible = 0 : Else : Pincab_BackglassL48.Visible = 1
    If Pincab_BackglassL55.BlendDisableLighting < 0.05 Then : Pincab_BackglassL55.Visible = 0 : Else : Pincab_BackglassL55.Visible = 1
    If Pincab_BackglassL56.BlendDisableLighting < 0.05 Then : Pincab_BackglassL56.Visible = 0 : Else : Pincab_BackglassL56.Visible = 1
    If Pincab_BackglassL57.BlendDisableLighting < 0.05 Then : Pincab_BackglassL57.Visible = 0 : Else : Pincab_BackglassL57.Visible = 1
    If Pincab_BackglassL58.BlendDisableLighting < 0.05 Then : Pincab_BackglassL58.Visible = 0 : Else : Pincab_BackglassL58.Visible = 1
    If Pincab_BackglassL59.BlendDisableLighting < 0.05 Then : Pincab_BackglassL59.Visible = 0 : Else : Pincab_BackglassL59.Visible = 1
    If Pincab_BackglassL60.BlendDisableLighting < 0.05 Then : Pincab_BackglassL60.Visible = 0 : Else : Pincab_BackglassL60.Visible = 1
    If Pincab_BackglassL61.BlendDisableLighting < 0.05 Then : Pincab_BackglassL61.Visible = 0 : Else : Pincab_BackglassL61.Visible = 1
    If Pincab_BackglassL62.BlendDisableLighting < 0.05 Then : Pincab_BackglassL62.Visible = 0 : Else : Pincab_BackglassL62.Visible = 1
    If Pincab_BackglassL63.BlendDisableLighting < 0.05 Then : Pincab_BackglassL63.Visible = 0 : Else : Pincab_BackglassL63.Visible = 1
    If Pincab_BackglassL64.BlendDisableLighting < 0.05 Then : Pincab_BackglassL64.Visible = 0 : Else : Pincab_BackglassL64.Visible = 1
    If Controller.Solenoid(25) = 0 Then : Pincab_BackglassS25.Visible = 0 : Else : Pincab_BackglassS25.Visible = 1 ' Jackpot
    If Controller.Solenoid(26) = 0 Then : Pincab_BackglassS26.Visible = 0 : Else : Pincab_BackglassS26.Visible = 1 ' Mouth
    If Controller.Solenoid(27) = 0 Then : Pincab_BackglassS27.Visible = 0 : Else : Pincab_BackglassS27.Visible = 1 ' Left Eye
    If Controller.Solenoid(28) = 0 Then : Pincab_BackglassS28.Visible = 0 : Else : Pincab_BackglassS28.Visible = 1 ' Firework 1
    If Controller.Solenoid(29) = 0 Then : Pincab_BackglassS29.Visible = 0 : Else : Pincab_BackglassS29.Visible = 1 ' Firework 2
    If Controller.Solenoid(31) = 0 Then : Pincab_BackglassS31.Visible = 0 : Else : Pincab_BackglassS31.Visible = 1 ' Right Eye
    If Controller.Solenoid(32) = 0 Then : Pincab_BackglassS32.Visible = 0 : Else : Pincab_BackglassS32.Visible = 1 ' Firework 3
    If Controller.Solenoid(11) = 0 Then ' GI for backglass on or off
      Pincab_BackglassS11.Visible = 0
    Else
      Pincab_BackglassS11.Visible = 1
    End If
  Else
    If BackglassReflOpt = 0 Then
      Pincab_Backglass.Visible = 0
      Pincab_BackglassS11.Visible = 0
      Pincab_BackglassS25.Visible = 0
      Pincab_BackglassS26.Visible = 0
      Pincab_BackglassS27.Visible = 0
      Pincab_BackglassS28.Visible = 0
      Pincab_BackglassS29.Visible = 0
      Pincab_BackglassS31.Visible = 0
      Pincab_BackglassS32.Visible = 0
      Pincab_BackglassL33.Visible = 0
      Pincab_BackglassL34.Visible = 0
      Pincab_BackglassL35.Visible = 0
      Pincab_BackglassL36.Visible = 0
      Pincab_BackglassL37.Visible = 0
      Pincab_BackglassL38.Visible = 0
      Pincab_BackglassL39.Visible = 0
      Pincab_BackglassL40.Visible = 0
      Pincab_BackglassL41.Visible = 0
      Pincab_BackglassL42.Visible = 0
      Pincab_BackglassL43.Visible = 0
      Pincab_BackglassL44.Visible = 0
      Pincab_BackglassL45.Visible = 0
      Pincab_BackglassL46.Visible = 0
      Pincab_BackglassL47.Visible = 0
      Pincab_BackglassL48.Visible = 0
      Pincab_BackglassL55.Visible = 0
      Pincab_BackglassL56.Visible = 0
      Pincab_BackglassL57.Visible = 0
      Pincab_BackglassL58.Visible = 0
      Pincab_BackglassL59.Visible = 0
      Pincab_BackglassL60.Visible = 0
      Pincab_BackglassL61.Visible = 0
      Pincab_BackglassL62.Visible = 0
      Pincab_BackglassL63.Visible = 0
      Pincab_BackglassL64.Visible = 0
    Else
      Pincab_Backglass.visible = 1
      If Pincab_BackglassL33.BlendDisableLighting < 0.05 Then : Pincab_BackglassL33.Visible = 0 : Else : Pincab_BackglassL33.Visible = 1
      If Pincab_BackglassL34.BlendDisableLighting  < 0.05 Then : Pincab_BackglassL34.Visible = 0 : Else : Pincab_BackglassL34.Visible = 1
      If Pincab_BackglassL35.BlendDisableLighting  < 0.05 Then : Pincab_BackglassL35.Visible = 0 : Else : Pincab_BackglassL35.Visible = 1
      If Pincab_BackglassL36.BlendDisableLighting < 0.05 Then : Pincab_BackglassL36.Visible = 0 : Else : Pincab_BackglassL36.Visible = 1
      If Pincab_BackglassL37.BlendDisableLighting < 0.05 Then : Pincab_BackglassL37.Visible = 0 : Else : Pincab_BackglassL37.Visible = 1
      If Pincab_BackglassL38.BlendDisableLighting < 0.05 Then : Pincab_BackglassL38.Visible = 0 : Else : Pincab_BackglassL38.Visible = 1
      If Pincab_BackglassL39.BlendDisableLighting < 0.05 Then : Pincab_BackglassL39.Visible = 0 : Else : Pincab_BackglassL39.Visible = 1
      If Pincab_BackglassL40.BlendDisableLighting < 0.05 Then : Pincab_BackglassL40.Visible = 0 : Else : Pincab_BackglassL40.Visible = 1
      If Pincab_BackglassL41.BlendDisableLighting < 0.05 Then : Pincab_BackglassL41.Visible = 0 : Else : Pincab_BackglassL41.Visible = 1
      If Pincab_BackglassL42.BlendDisableLighting < 0.05 Then : Pincab_BackglassL42.Visible = 0 : Else : Pincab_BackglassL42.Visible = 1
      If Pincab_BackglassL43.BlendDisableLighting < 0.05 Then : Pincab_BackglassL43.Visible = 0 : Else : Pincab_BackglassL43.Visible = 1
      If Pincab_BackglassL44.BlendDisableLighting < 0.05 Then : Pincab_BackglassL44.Visible = 0 : Else : Pincab_BackglassL44.Visible = 1
      If Pincab_BackglassL45.BlendDisableLighting < 0.05 Then : Pincab_BackglassL45.Visible = 0 : Else : Pincab_BackglassL45.Visible = 1
      If Pincab_BackglassL46.BlendDisableLighting < 0.05 Then : Pincab_BackglassL46.Visible = 0 : Else : Pincab_BackglassL46.Visible = 1
      If Pincab_BackglassL47.BlendDisableLighting < 0.05 Then : Pincab_BackglassL47.Visible = 0 : Else : Pincab_BackglassL47.Visible = 1
      If Pincab_BackglassL48.BlendDisableLighting < 0.05 Then : Pincab_BackglassL48.Visible = 0 : Else : Pincab_BackglassL48.Visible = 1
      If Pincab_BackglassL55.BlendDisableLighting < 0.05 Then : Pincab_BackglassL55.Visible = 0 : Else : Pincab_BackglassL55.Visible = 1
      If Pincab_BackglassL56.BlendDisableLighting < 0.05 Then : Pincab_BackglassL56.Visible = 0 : Else : Pincab_BackglassL56.Visible = 1
      If Pincab_BackglassL57.BlendDisableLighting < 0.05 Then : Pincab_BackglassL57.Visible = 0 : Else : Pincab_BackglassL57.Visible = 1
      If Pincab_BackglassL58.BlendDisableLighting < 0.05 Then : Pincab_BackglassL58.Visible = 0 : Else : Pincab_BackglassL58.Visible = 1
      If Pincab_BackglassL59.BlendDisableLighting < 0.05 Then : Pincab_BackglassL59.Visible = 0 : Else : Pincab_BackglassL59.Visible = 1
      If Pincab_BackglassL60.BlendDisableLighting < 0.05 Then : Pincab_BackglassL60.Visible = 0 : Else : Pincab_BackglassL60.Visible = 1
      If Pincab_BackglassL61.BlendDisableLighting < 0.05 Then : Pincab_BackglassL61.Visible = 0 : Else : Pincab_BackglassL61.Visible = 1
      If Pincab_BackglassL62.BlendDisableLighting < 0.05 Then : Pincab_BackglassL62.Visible = 0 : Else : Pincab_BackglassL62.Visible = 1
      If Pincab_BackglassL63.BlendDisableLighting < 0.05 Then : Pincab_BackglassL63.Visible = 0 : Else : Pincab_BackglassL63.Visible = 1
      If Pincab_BackglassL64.BlendDisableLighting < 0.05 Then : Pincab_BackglassL64.Visible = 0 : Else : Pincab_BackglassL64.Visible = 1
      If Controller.Solenoid(25) = 0 Then : Pincab_BackglassS25.Visible = 0 : Else : Pincab_BackglassS25.Visible = 1 ' Jackpot
      If Controller.Solenoid(26) = 0 Then : Pincab_BackglassS26.Visible = 0 : Else : Pincab_BackglassS26.Visible = 1 ' Mouth
      If Controller.Solenoid(27) = 0 Then : Pincab_BackglassS27.Visible = 0 : Else : Pincab_BackglassS27.Visible = 1 ' Left Eye
      If Controller.Solenoid(28) = 0 Then : Pincab_BackglassS28.Visible = 0 : Else : Pincab_BackglassS28.Visible = 1  ' Firework 1
      If Controller.Solenoid(29) = 0 Then : Pincab_BackglassS29.Visible = 0 : Else : Pincab_BackglassS29.Visible = 1 ' Firework 2
      If Controller.Solenoid(31) = 0 Then : Pincab_BackglassS31.Visible = 0 : Else : Pincab_BackglassS31.Visible = 1 ' Right Eye
      If Controller.Solenoid(32) = 0 Then : Pincab_BackglassS32.Visible = 0 : Else : Pincab_BackglassS32.Visible = 1 ' Firework 3
      If Controller.Solenoid(11) = 0 Then ' GI for backglass on or off
        Pincab_BackglassS11.Visible = 0
      Else
        Pincab_BackglassS11.Visible = 1
      End If
    End If
  End If
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

Sub UpdatePlunger()
  If plungerpress = 1 then
    If Pincab_Shooter.Y < -25.5007 then
      Pincab_Shooter.Y = Pincab_Shooter.Y + 2.5
    End If
  Else
    Pincab_Shooter.Y = -125.5007 + (5 * Plunger.Position) - 20
  End If
End Sub

Sub Table1_Init()
  vpmInit me

  With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Cyclone (Williams 1988)"
        .Games(cGameName).Settings.Value("rol") = 0 '1 = rotated display, 0 = normal
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
    .Hidden = DesktopMode
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
  Controller.Run

    ' Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

  Set CyclnBall1 = Outhole.CreateSizedballWithMass(Ballsize/2,Ballmass)

  gBOT = Array(CyclnBall1)

    Controller.Switch(10) = 1

    Dim mWheelMech
    Set mWheelMech = New cvpmMech
  If MysteryWheelDT = 0 Then
    With mWheelMech
      .MType = vpmMechStepSol + vpmMechCircle + vpmMechLinear + vpmMechFast
      .Sol1 = 14
      .Sol2 = 13
      .Length = 200
      .Steps = 360
      .AddSw 41, 0, 180
      .Callback = GetRef("SpinWheel")
      .Start
    End With
  Else
    With mWheelMech
      .MType = vpmMechStepSol + vpmMechCircle + vpmMechLinear + vpmMechFast
      .Sol1 = 14
      .Sol2 = 13
      .Length = 200
      .Steps = 16
      .AddSw 41, 0, 8
      .Callback = GetRef("SpinWheel")
      .Start
    End With
  End If

  'Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

  If ModSol = 0 Then
    Dim bub
    For each bub in GI:bub.State = 1:Next
  End If

  'Initialize slings
  RStep = 0 : RightSlingShot.Timerenabled = True
  LStep = 0 : LeftSlingShot.Timerenabled = True

  If VRMode = True Then
    setup_backglass()
    InitDigits
  End If

  FlFadeBumper 1,1
  FlFadeBumper 2,1
  FlFadeBumper 3,1

End Sub

Sub Table1_KeyDown(keycode)
  If Keycode = StartGameKey Then
        SoundStartButton
    'StartButton.y = StartButton.y -5
    'StartButton2.y = StartButton2.y -5
  End If
    If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft
    If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight
    If keycode = CenterTiltKey Then Nudge 0, 1:SoundNudgeCenter
  If keycode = LeftFlipperKey then Pincab_LeftFlipperButton.X = Pincab_LeftFlipperButton.X + 8
  If keycode = RightFlipperKey then Pincab_RightFlipperButton.X = Pincab_RightFlipperButton.X - 8
    If keycode = PlungerKey Then
    SoundPlungerPull
    Plunger.Pullback
    plungerpress = 1
  End If
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then 'Use this for ROM based games
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If KeyCode = PlungerKey Then
    Plunger.Fire
    plungerpress = 0
    If BIPL = 1 Then
      SoundPlungerReleaseBall()   'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall() 'Plunger release sound when there is no ball in shooter lane
    End If
  End If
  If keycode = LeftFlipperKey then Pincab_LeftFlipperButton.X = Pincab_LeftFlipperButton.X - 8
  If keycode = RightFlipperKey then Pincab_RightFlipperButton.X = Pincab_RightFlipperButton.X + 8
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

Function TroughIsFull
  TroughIsFull =  Controller.Switch(10)
End Function

Sub Table1_Paused: Controller.Pause = 1: End Sub
Sub Table1_unPaused: Controller.Pause = 0: End Sub
Sub Table1_Exit: Controller.Stop: End Sub

Dim GiIntensity
GiIntensity = 1   'used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 1000       'X position of punger lane left
Const PLRight = 1060      'X position of punger lane right
Const PLTop = 1225        'Y position of punger lane top
Const PLBottom = 1900       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135*gilvl ' orig was 120 and 70

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    d_w = b_base*(1 - (gBOT(s).z-25)/500)
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

'**************************************************************
'               SOLENOIDS
'**************************************************************
SolCallback(1) = "SolBallRelease"
SolCallback(3) = "EMKicker"
SolCallback(4) = "solBoomerangKickout"
SolCallback(5) = "SpookHouseDropUp"
SolCallback(7) = "KnockerSolenoid"
SolCallback(11) = "SolBackGenIllumin" 'BG
SolCallback(16) = "SolWheelDrive"

If UseVPMModSol Then
  SolModCallback(10) = "SolModGI"   'PF GI
  SolModCallback(15) = "FlashMod15" 'boomerang flashers
  SolModCallback(25) = "FlashMod25" 'pops left
  SolModCallback(26) = "FlashMod26" 'pops right
  SolModCallback(27) = "FlashMod27" 'backmiddle
  SolModCallback(28) = "FlashMod28" 'cats
  SolModCallback(29) = "FlashMod29" 'ducks
  SolModCallback(30) = "FlashMod30" 'backleft/Ferris
  SolModCallback(31) = "FlashMod31" 'Cyclone Flasher
  SolModCallback(32) = "FlashMod32" 'SpookHouse
Else
  SolCallback(10) = "SolGI"   'PF GI
  SolCallback(15) = "Flash15" 'boomerang flashers
  SolCallback(25) = "Flash17" 'popbumper left
  SolCallback(25) = "Flash25" 'pops left
  SolCallback(26) = "Flash26" 'pops right
  SolCallback(27) = "Flash27" 'backmiddle
  SolCallback(28) = "Flash28" 'cats
  SolCallback(29) = "Flash29" 'ducks
  SolCallback(30) = "Flash30" 'backleft/Ferris
  SolCallback(31) = "Flash31" 'Cyclone Flasher
  SolCallback(32) = "Flash32" 'SpookHouse
End If

Sub SolBallRelease(enabled)
    If enabled Then
        BIP = BIP + 1
    Outhole.kick 90, 45
    RandomSoundShooterFeeder
    End If
End Sub

Sub Outhole_Hit
  PlaySound "drain", 0, MechVolume, AudioPan(Outhole), 0,0,0, 1, AudioFade(Outhole):controller.switch(10) = 1
  BIP = BIP - 1
End Sub

Sub Outhole_unHit:Controller.Switch(10) = 0: End Sub

Dim empos, kForce
Sub EMKicker(Enabled)
    If Enabled Then
  kForce = 46.5 + (Rnd*(rnd*9.3))
    KSalida.Kick 0,kForce
    Playsound SoundFX("solenoid",DOFContactors), 0, MechVolume, AudioPan(KSalida), 0,0,0, 1, AudioFade(KSalida)
        empos = 0
        emkickertimer.Enabled = 1
    End If
End Sub

Sub KSalida_Hit
    Playsound "MetalHit", 0, MechVolume, AudioPan(KSalida), 0,0,0, 1, AudioFade(KSalida)
End Sub

Sub emkickertimer_Timer
    Select Case empos
        Case 0:emk.TransY = -10
        Case 1:emk.TransY = -20
        Case 2:emk.TransY = -30
        Case 3:emk.TransY = -40
        Case 4:emk.TransY = -30
        Case 5:emk.TransY = -20
        Case 6:emk.TransY = -10
        Case 7:emk.TransY = 0
        Case else emkickertimer.Enabled = 0
    End Select
    empos = empos + 1
End Sub

Sub SolBackGenIllumin(Enabled)
  If Enabled Then
    'PlaySound"tickon" 'plays in BG sounds
  Else
    'PlaySound"tickoff" 'Plays in BG sounds
  End If
End Sub


'*******************************************
' ZFLP: Flippers
'*******************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const ReflipAngle = 20
Const QuickFlipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire  'leftflipper.rotatetoend

    If LeftFlipper.currentangle < LeftFlipper.endangle + ReflipAngle Then
      'Play partial flip sound and stop any flip down sound
      StopAnyFlipperLowerLeftDown()
      RandomSoundFlipperLowerLeftReflip LeftFlipper
    Else
      'Play full flip sound
      If BallNearLF = 0 Then
        RandomSoundFlipperLowerLeftUpFullStroke LeftFlipper
      End If
      If BallNearLF = 1 Then
        Select Case Int(Rnd*2)+1
          Case 1 : RandomSoundFlipperLowerLeftUpDampenedStroke LeftFlipper
          Case 2 : RandomSoundFlipperLowerLeftUpFullStroke LeftFlipper
        End Select
      End If
    End If
  Else
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      'Play flip down sound
      RandomSoundFlipperLowerLeftDown LeftFlipper
    End If
    If LeftFlipper.currentangle < LeftFlipper.startAngle + QuickFlipAngle and LeftFlipper.currentangle <> LeftFlipper.endangle Then
      'Play quick flip sound and stop any flip up sound
      StopAnyFlipperLowerLeftUp()
      RandomSoundLowerLeftQuickFlipUp()
    Else
      FlipperLeftLowerHitParm = FlipperUpSoundLevel
    End If
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    FlipperActivate RightFlipper, RFPress
    RF.Fire 'rightflipper.rotatetoend

    If RightFlipper.currentangle > RightFlipper.endangle - ReflipAngle Then
      'Play partial flip sound and stop any flip down sound
      StopAnyFlipperLowerRightDown()
      RandomSoundFlipperLowerRightReflip RightFlipper
    Else
      'Play full flip sound
      If BallNearRF = 0 Then
        RandomSoundFlipperLowerRightUpFullStroke RightFlipper
      End If

      If BallNearRF = 1 Then
        Select Case Int(Rnd*2)+1
          Case 1 : RandomSoundFlipperLowerRightUpDampenedStroke RightFlipper
          Case 2 : RandomSoundFlipperLowerRightUpFullStroke RightFlipper
        End Select
      End If
    End If
  Else
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      'Play flip down sound
      RandomSoundFlipperLowerRightDown RightFlipper
    End If
    If RightFlipper.currentangle < RightFlipper.startAngle + QuickFlipAngle and RightFlipper.currentangle <> RightFlipper.endangle Then
      'Play quick flip sound and stop any flip up sound
      StopAnyFlipperLowerRightUp()
      RandomSoundLowerRightQuickFlipUp()
    Else
      FlipperRightLowerHitParm = FlipperUpSoundLevel
    End If
  End If
End Sub

'******************************************************
'  ZANI: Misc Animations
'******************************************************

Sub LeftFlipper_Animate
  dim a: a = LeftFlipper.CurrentAngle
  FlipperShadowL.RotZ = a
  batleft.objrotz = a
End Sub

Sub RightFlipper_Animate
  dim a: a = RightFlipper.CurrentAngle
  FlipperShadowR.RotZ = a
  batright.objrotz = a
End Sub

'**************************************************************
'          BUMPERS
'**************************************************************

Sub Bumper1_Hit
  vpmTimer.pulsesw 60
  RandomSoundBumperLeft Bumper1
End Sub

Sub Bumper2_Hit
  vpmTimer.pulsesw 61
  RandomSoundBumperUp Bumper2
End Sub

Sub Bumper3_Hit
  vpmTimer.pulsesw 62
  RandomSoundBumperLow Bumper3
End Sub

'****************************************************************
' ZSLG: Slingshots
'****************************************************************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
    vpmTimer.PulseSw 63
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 10
  RandomSoundSlingshotLeft
End Sub

Sub LeftSlingShot_Timer ' animation of the rubber
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
    vpmTimer.PulseSw 64
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 10
  RandomSoundSlingshotRight
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'**************************************************************
'          Targets
'**************************************************************
'Ducks
Sub sw19_Hit() : STHit 19 : End sub
Sub sw20_Hit() : STHit 20 : End sub
Sub sw21_Hit() : STHit 21 : End sub
'Cats
Sub sw22_Hit() : STHit 22 : End sub
Sub sw23_Hit() : STHit 23 : End sub
Sub sw24_Hit() : STHit 24 : End sub

Sub sw59_Hit() : STHit 59 : End sub

'Boomerang Kicker
dim bola, bola2, altura, altura2
Sub sw13_Hit:Controller.Switch(13) = 1:PlaySound "kicker2-Boomerang-Enter-Kicker", 0, MechVolume, AudioPan(sw13), 0, Pitch(ActiveBall), 1, 0, AudioFade(sw13):End Sub
Sub sw13_unHit: Controller.Switch(13) = 0:end Sub
Sub sw11_Hit:vpmtimer.pulsesw 11:End Sub

'********************** Drop Target ************************
Sub sw18_Hit(): DTHit 18 : TargetBouncer Activeball, 1.5 : End Sub

Sub SpookHouseDropUp(Enabled)
  If Enabled then
    PlaySoundAt SoundFX(DTResetSound,DOFContactors),sw18p
    DTRaise 18
  end if
End Sub

Sub SolBoomerangKickOut(Enabled)
  Kicker2.Enabled = 0
  Sw13.kick 325, 36
  Kicker2.timerenabled = 1
  Playsound SoundFX("boomerang_kick",DOFContactors), 0, MechVolume, AudioPan(Sw13), 0, 0, 1, 0, AudioFade(Sw13)
End Sub

Sub Kicker2_Hit:Playsound "Scoop_Enter", 0, MechVolume, AudioPan(Kicker2), 0, 0, 1, 0, AudioFade(Kicker2):End Sub
Sub Kicker1_Hit:Playsound "Kicker2", 0, MechVolume, AudioPan(Kicker1), 0, 0, 1, 0, AudioFade(Kicker1):End Sub

Sub Kicker2_Timer
  Kicker2.Enabled = 1
  Kicker2.timerenabled = 0
End Sub

'**************************************************************
'             SWITCHES
'**************************************************************
Sub ShooterLane_Hit:Controller.Switch(25) = 1 : BIPL = True : End Sub
Sub ShooterLane_Unhit:Controller.Switch(25) = 0 : BIPL = False:End Sub

'Pasillos
Sub sw53_Hit:Controller.Switch(53) = 1:End Sub
Sub sw53_Unhit:Controller.Switch(53) = 0:End Sub
Sub sw54_Hit:Controller.Switch(54) = 1:End Sub
Sub sw54_Unhit:Controller.Switch(54) = 0:End Sub
Sub sw55_Hit:Controller.Switch(55) = 1:End Sub
Sub sw55_Unhit:Controller.Switch(55) = 0:End Sub
Sub sw56_Hit:Controller.Switch(56) = 1:End Sub
Sub sw56_Unhit:Controller.Switch(56) = 0:End Sub
Sub sw14_Hit:Controller.Switch(14) = 1:End Sub
Sub sw14_Unhit:Controller.Switch(14) = 0:End Sub
Sub sw15_Hit:Controller.Switch(15) = 1:End Sub
Sub sw15_Unhit:Controller.Switch(15) = 0:End Sub
Sub sw16_Hit:Controller.Switch(16) = 1:End Sub
Sub sw16_Unhit:Controller.Switch(16) = 0:End Sub
Sub Sw17_Hit:Controller.Switch(17) = 1:End Sub 'Entrance ferris wheel
Sub Sw17_UnHit:Controller.Switch(17) = 0:End Sub
Sub Sw31_Hit:Controller.Switch(31) = 1:PlaySound"comet_spinner", 0, MechVolume, AudioPan(sw31), 0, 0, 1, 0, AudioFade(sw31):End Sub'Entrance rampa comet
Sub Sw31_UnHit
  Controller.Switch(31) = 0
  If ActiveBall.VelY < 0 Then
    PlaySound"comet_ramp", 0, MechVolume, AudioPan(sw31), 0, 0, 1, 0, AudioFade(sw31)
  Else
    StopSound"comet_ramp"
  End If
  If (ActiveBall.VelY > 0) Then
    RandomSoundRampFlapDown()
  ElseIf (ActiveBall.VelY < 0) Then
    RandomSoundRampFlapUp()
  End If
End Sub
Sub Sw32_Hit:Controller.Switch(32) =1:SwitchWire2.objRotZ = -15:End Sub                          'Score ramp comet
Sub Sw32_UnHit:Controller.Switch(32) =0:SwitchWire2.objRotZ = 0:End Sub
Sub Sw34_Hit:Controller.Switch(34) =1:SwitchWire1.objRotZ = -15:End Sub                          'Score ramp cyclone
Sub Sw34_UnHit:Controller.Switch(34) =0:SwitchWire1.objRotZ = 0:End Sub

Sub sw35_Hit:VPMTimer.PulseSw 35:End Sub
Sub sw36_Hit:VPMTimer.PulseSw 36:End Sub
Sub sw37_Hit:VPMTimer.PulseSw 37:End Sub

'SkillShot
Sub sw26_Hit():VPMTimer.PulseSw 26:End Sub
Sub sw27_Hit():VPMTimer.PulseSw 27:End Sub
Sub sw28_Hit():VPMTimer.PulseSw 28:End Sub
Sub sw29_Hit():VPMTimer.PulseSw 29:End Sub
Sub sw30_Hit():VPMTimer.PulseSw 30:End Sub

'**************************************************************
'             MYSTERY WHEEL
'**************************************************************

Sub SpinWheel(aNewPos, aSpeed, aLastPos)
    If aNewPos <> aLastPos then
        WheelR.SetValue aNewPos
    VRMysteryWheelReel.RotY = aNewPos
    End If
End Sub

'**************************************************************
'   NEW FERRIS WHEEL  - Ball movement based on Rothbauerw script - MWR
'**************************************************************

Dim Tball,tcount,xstart,ystart,zstart,xbstart,ybstart,zbstart,xphstart,yphstart,zphstart
Dim dradius, xyangle, zangle, dzangle

Function PI()
  PI = 4*Atn(1)
End Function

Function Radians(angle)
  Radians = PI * angle / 180
End Function

xphstart = 230
yphstart = 243
zphstart = -18

dradius= 80
xyangle = 290
zangle = 475

xstart = dradius*cos(radians(xyangle))*Sin(Radians(zangle))
ystart = dradius*sin(radians(xyangle))*Sin(Radians(zangle))
zstart = dradius*cos(radians(zangle))

Sub Trigger2_Hit
  PlaySound"ferris_hit1", 0, MechVolume, AudioPan(Trigger2), 0, 0, 1, 0, AudioFade(Trigger2)
  activeball.vely=0
  activeball.velz=0
  activeball.x=xphstart
  activeball.y=yphstart
  activeball.z=zphstart
  Set Tball = Activeball
  tcount = 0
  WheelHit.enabled = 1
  Trigger2.enabled = 0
End Sub

Sub Trigger2_timer
    WheelHit.Enabled = 0
  Dim xd, yd, zd
  If tcount =  1 Then
    xbstart = Tball.x
    ybstart = Tball.y
    zbstart = Tball.z
    dzangle = zangle
  End If

  If tcount > 2 Then
    dzangle = dzangle - .4
    xd = xstart - dradius*cos(radians(xyangle))*Sin(Radians(dzangle))
    yd = ystart - dradius*sin(radians(xyangle))*Sin(Radians(dzangle))
    zd = zstart - dradius*cos(radians(dzangle))
    Tball.x = xbstart - xd
    Tball.y = ybstart - yd
    Tball.z = zbstart - zd
    Tball.velx = 0
    Tball.vely = 0
    Tball.velz = 0
    if dzangle < 350 Then
      dzangle = 0
      Trigger2.TimerEnabled = false
      Trigger2.Enabled = 1
    end if
  End If
  tcount = tcount + 1
End Sub

Sub SolWheelDrive(enabled)
    if enabled then
        FWTimer.Enabled = 1
    PlaySound"motor", -1, MechVolume/5, AudioPan(Trigger2), 0, 0, 1, 0, AudioFade(Trigger2)
    else
        FWTimer.Enabled = 0
    StopSound"motor"
    end if
End Sub

Sub FWTimer_Timer
    RedWheel.RotX = (RedWheel.RotX + 1) MOD 360

  If  RedWheel.RotX >15 and RedWheel.RotX <65 Then
    FerrisBlockedWall.isdropped = 1
  elseif RedWheel.RotX >105 and RedWheel.RotX <150 Then
    FerrisBlockedWall.isdropped = 1
  elseif RedWheel.RotX >200 and RedWheel.RotX <240 Then
    FerrisBlockedWall.isdropped = 1
  elseif RedWheel.RotX >285 and RedWheel.RotX <330 Then
    FerrisBlockedWall.isdropped = 1
  else
    FerrisBlockedWall.isdropped = 0
  End If
End Sub

Sub WheelHit_Timer
  If (RedWheel.Rotx > 49 and RedWheel.Rotx < 52) or (RedWheel.Rotx > 139 and RedWheel.Rotx < 142) or (RedWheel.Rotx > 229 AND RedWheel.Rotx < 232) or (RedWheel.Rotx > 319 and RedWheel.Rotx < 322) Then
' If RedWheel.Rotx = 50  or RedWheel.Rotx = 140 or RedWheel.Rotx = 230 or RedWheel.RotX = 320 Then
    Trigger2.TimerEnabled = True
  End If
End Sub

Sub FerrisBlockedWall_Hit
  PlaySound "MetalHit", 0, (Vol(ActiveBall)*5), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

'**************************************************************
'   END FERRIS WHEEL
'**************************************************************

' Ferris Wheel RAMP Sounds
Sub FerrisSE_Hit: PlaySound"ferris_Exit", 0, MechVolume, AudioPan(FerrisSE), 0, 0, 1, 0, AudioFade(FerrisSE):End Sub
Sub FerrisS1_Hit: PlaySound"ferris_hit1", 0, MechVolume, AudioPan(FerrisS1), 0, 0, 1, 0, AudioFade(FerrisS1):End Sub
Sub FerrisS2_Hit: PlaySound"ferris_hit2", 0, MechVolume, AudioPan(FerrisS2), 0, 0, 1, 0, AudioFade(FerrisS2):End Sub
Sub FerrisS3_Hit: PlaySound"ferris_hit1", 0, MechVolume, AudioPan(FerrisS3), 0, 0, 1, 0, AudioFade(FerrisS3):End Sub
Sub FerrisS4_Hit: PlaySound"ferris_hit2", 0, MechVolume, AudioPan(FerrisS4), 0, 0, 1, 0, AudioFade(FerrisS4):End Sub
Sub FerrisS5_Hit: PlaySound"ferris_hit1", 0, MechVolume, AudioPan(FerrisS5), 0, 0, 1, 0, AudioFade(FerrisS5):End Sub
Sub FerrisS6_Hit: PlaySound"ferris_hit2", 0, MechVolume, AudioPan(FerrisS6), 0, 0, 1, 0, AudioFade(FerrisS6):End Sub
Sub FerrisS7_Hit: PlaySound"ferris_hit1", 0, MechVolume, AudioPan(FerrisS7), 0, 0, 1, 0, AudioFade(FerrisS7):End Sub
Sub FerrisS8_Hit: PlaySound"ferris_hit2", 0, MechVolume, AudioPan(FerrisS8), 0, 0, 1, 0, AudioFade(FerrisS8):End Sub
'Sub FerrisS9_Hit: bsRampOn: PlaySound"top_lane", 0, 1, AudioPan(FerrisS9), 0, 0, 1, 0, AudioFade(FerrisS9):End Sub

Sub FinalRampaComet_Hit
    Playsound "comet_exit1", 0, MechVolume, AudioPan(FinalRampaComet), 0, 0, 1, 0, AudioFade(FinalRampaComet)
End Sub

Sub MitadRampaComet_Hit
    me.timerinterval = 180
    PlaySound"comet_exit3", 0, MechVolume, AudioPan(MitadRampaComet), 0, 0, 1, 0, AudioFade(MitadRampaComet)
End Sub

Sub FinalRampaCyclone_Hit
    me.timerinterval = 150
    PlaySound"comet_exit2", 0, MechVolume, AudioPan(FinalRampaCyclone), 0, 0, 1, 0, AudioFade(FinalRampaCyclone)
End Sub

Sub FinalRampaFerris_Hit
    me.timerinterval = 150
  PlaySound"ferris_ball_drop", 0, MechVolume, AudioPan(FinalRampaFerris), 0, 0, 1, 0, AudioFade(FinalRampaFerris)
End Sub

'Comet /Cyclone Ramp sounds
Sub CometS1_Hit
  If ActiveBall.VelY < 0 then
    PlaySound"cyclone_ramp_enter", 0, MechVolume, AudioPan(CometS1), 0, 0, 1, 0, AudioFade(CometS1)
  End If
  If (ActiveBall.VelY > 0) Then
    RandomSoundRampFlapDown()
  ElseIf (ActiveBall.VelY < 0) Then
    RandomSoundRampFlapUp()
  End If
End Sub
Sub CometS1_UnHit
  If ActiveBall.VelY > 0 Then
    StopSound"cyclone_ramp_enter"
  End If
End Sub

Sub CometS2_Hit: PlaySound"comet_hit1", 0, MechVolume, AudioPan(CometS2), 0, 0, 1, 0, AudioFade(CometS2):End Sub
Sub CometS3_Hit: PlaySound"comet_hit2", 0, MechVolume, AudioPan(CometS3), 0, 0, 1, 0, AudioFade(CometS3):End Sub
Sub CometS4_Hit: PlaySound"comet_hit1", 0, MechVolume, AudioPan(CometS4), 0, 0, 1, 0, AudioFade(CometS4):End Sub
Sub CometS5_Hit: PlaySound"comet_hit2", 0, MechVolume, AudioPan(CometS5), 0, 0, 1, 0, AudioFade(CometS5):End Sub
Sub CometS6_Hit: PlaySound"comet_hit1", 0, MechVolume, AudioPan(CometS6), 0, 0, 1, 0, AudioFade(CometS6):End Sub
Sub CometS7_Hit: PlaySound"comet_hit2", 0, MechVolume, AudioPan(CometS7), 0, 0, 1, 0, AudioFade(CometS7):End Sub
Sub CometS8_Hit: PlaySound"comet_hit1", 0, MechVolume, AudioPan(CometS8), 0, 0, 1, 0, AudioFade(CometS8):End Sub
Sub CometS9_Hit: PlaySound"comet_hit2", 0, MechVolume, AudioPan(CometS9), 0, 0, 1, 0, AudioFade(CometS9):End Sub

Sub SkillTrigger1_hit()
  WireRampOn False
  If (ActiveBall.VelY > 0) Then
    RandomSoundRampFlapDown()
  ElseIf (ActiveBall.VelY < 0) Then
    RandomSoundRampFlapUp()
  End If
End Sub

Sub SkillTrigger1_UnHit()
  If ActiveBall.VelY > 0 Then
    WireRampOff
  End If
End Sub

Sub SkillTrigger2_hit()
  WireRampOff
End Sub

Sub SkillTrigger2_UnHit()
  WireRampOn True
End Sub

'SkillShot drop sounds
Sub skillS1_Hit:WireRampOff:PlaySound"ball_drop", 0, MechVolume, AudioPan(skillS1), 0, 0, 1, 0, AudioFade(skillS1):End Sub
Sub skillS2_Hit:WireRampOff:PlaySound"ball_drop", 0, MechVolume, AudioPan(skillS2), 0, 0, 1, 0, AudioFade(skillS2):End Sub
Sub skillS3_Hit:WireRampOff:PlaySound"ball_drop", 0, MechVolume, AudioPan(skillS3), 0, 0, 1, 0, AudioFade(skillS3):End Sub
Sub skillS4_Hit:WireRampOff:PlaySound"ball_drop", 0, MechVolume, AudioPan(skillS4), 0, 0, 1, 0, AudioFade(skillS4):End Sub
Sub skillS5_Hit:WireRampOff:PlaySound"ball_drop", 0, MechVolume, AudioPan(skillS5), 0, 0, 1, 0, AudioFade(skillS5):End Sub
Sub skillS6_Hit:WireRampOff:PlaySound"ball_drop", 0, MechVolume, AudioPan(skillS6), 0, 0, 1, 0, AudioFade(skillS6):End Sub

'************************************
'          LEDs Display
'************************************

'Desktop
Dim Digits(28)
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
Digits(14) = Array(b00, b02, b05, b06, b04, b01, b03, b07)
Digits(15) = Array(b10, b12, b15, b16, b14, b11, b13)
Digits(16) = Array(b20, b22, b25, b26, b24, b21, b23)
Digits(17) = Array(b30, b32, b35, b36, b34, b31, b33, b37)
Digits(18) = Array(b40, b42, b45, b46, b44, b41, b43)
Digits(19) = Array(b50, b52, b55, b56, b54, b51, b53)
Digits(20) = Array(b60, b62, b65, b66, b64, b61, b63)
Digits(21) = Array(b70, b72, b75, b76, b74, b71, b73, b77)
Digits(22) = Array(b80, b82, b85, b86, b84, b81, b83)
Digits(23) = Array(b90, b92, b95, b96, b94, b91, b93)
Digits(24) = Array(ba0, ba2, ba5, ba6, ba4, ba1, ba3, ba7)
Digits(25) = Array(bb0, bb2, bb5, bb6, bb4, bb1, bb3)
Digits(26) = Array(bc0, bc2, bc5, bc6, bc4, bc1, bc3)
Digits(27) = Array(bd0, bd2, bd5, bd6, bd4, bd1, bd3)

'***************************************************************************
' VR BG Display
'***************************************************************************
Dim DisplayColor : DisplayColor = RGB(255,88,32)  ' Color of VR digits

Dim VRDigits(27)
VRDigits(0)=Array(ax00, ax05, ax0c, ax0d, ax08, ax01, ax06, ax0f, ax02, ax03, ax04, ax07, ax0b, ax0a, ax09, ax0e)
VRDigits(1)=Array(ax10, ax15, ax1c, ax1d, ax18, ax11, ax16, ax1f, ax12, ax13, ax14, ax17, ax1b, ax1a, ax19, ax1e)
VRDigits(2)=Array(ax20, ax25, ax2c, ax2d, ax28, ax21, ax26, ax2f, ax22, ax23, ax24, ax27, ax2b, ax2a, ax29, ax2e)
VRDigits(3)=Array(ax30, ax35, ax3c, ax3d, ax38, ax31, ax36, ax3f, ax32, ax33, ax34, ax37, ax3b, ax3a, ax39, ax3e)
VRDigits(4)=Array(ax40, ax45, ax4c, ax4d, ax48, ax41, ax46, ax4f, ax42, ax43, ax44, ax47, ax4b, ax4a, ax49, ax4e)
VRDigits(5)=Array(ax50, ax55, ax5c, ax5d, ax58, ax51, ax56, ax5f, ax52, ax53, ax54, ax57, ax5b, ax5a, ax59, ax5e)
VRDigits(6)=Array(ax60, ax65, ax6c, ax6d, ax68, ax61, ax66, ax6f, ax62, ax63, ax64, ax67, ax6b, ax6a, ax69, ax6e)
VRDigits(7)=Array(ax70, ax75, ax7c, ax7d, ax78, ax71, ax76, ax7f, ax72, ax73, ax74, ax77, ax7b, ax7a, ax79, ax7e)
VRDigits(8)=Array(ax80, ax85, ax8c, ax8d, ax88, ax81, ax86, ax8f, ax82, ax83, ax84, ax87, ax8b, ax8a, ax89, ax8e)
VRDigits(9)=Array(ax90, ax95, ax9c, ax9d, ax98, ax91, ax96, ax9f, ax92, ax93, ax94, ax97, ax9b, ax9a, ax99, ax9e)
VRDigits(10)=Array(axa0, axa5, axac, axad, axa8, axa1, axa6, axaf, axa2, axa3, axa4, axa7, axab, axaa, axa9, axae)
VRDigits(11)=Array(axb0, axb5, axbc, axbd, axb8, axb1, axb6, axbf, axb2, axb3, axb4, axb7, axbb, axba, axb9, axbe)
VRDigits(12)=Array(axc0, axc5, axcc, axcd, axc8, axc1, axc6, axcf, axc2, axc3, axc4, axc7, axcb, axca, axc9, axce)
VRDigits(13)=Array(axd0, axd5, axdc, axdd, axd8, axd1, axd6, axdf, axd2, axd3, axd4, axd7, axdb, axda, axd9, axde)

VRDigits(14) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6)
VRDigits(15) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)
VRDigits(16) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)
VRDigits(17) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6)
VRDigits(18) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)
VRDigits(19) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)
VRDigits(20) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6)
VRDigits(21) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6)
VRDigits(22) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)
VRDigits(23) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)
VRDigits(24) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6)
VRDigits(25) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)
VRDigits(26) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)
VRDigits(27) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6)

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen, ix, xx, yy, xobj

Sub setup_backglass()
  xoff = -20
  yoff = -66 ' 64
  zoff = 822
  xrot = -90
  zscale = 0.0000001
  xcen = 0  '(130 /2) - (92 / 2)
  ycen = (780 /2 ) + (203 /2)
  for ix = 0 to Ubound(VRDigits)
    For Each xobj In VRDigits(ix)
      xx = xobj.x
      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff
      If (yy < 0.) then
        yy = yy * -1
      end if
      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next
End Sub

Sub VRDisplayTimer
  Dim ii, jj, obj, b, x
  Dim ChgLED,num, chg, stat
  ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        For Each obj In VRDigits(num)
 '                  If chg And 1 Then obj.visible=stat And 1    'if you use the object color for off; turn the display object visible to not visible on the playfield, and uncomment this line out.
          If chg And 1 Then FadeDisplay obj, stat And 1
          chg=chg\2 : stat=stat\2
        Next
      Next
    End If
End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
    Object.Opacity = 50
  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 1
  End If
End Sub

Sub InitDigits()
  dim tmp, x, obj
  for x = 0 to uBound(VRDigits)
    if IsArray(VRDigits(x) ) then
      For each obj in VRDigits(x)
        obj.height = obj.height + 0
        FadeDisplay obj, 0
      next
    end If
  Next
End Sub

Sub DisplayTimer_Timer()
  If VRMode = False Then
    UpdateLeds
  End If
End Sub

Sub UpdateLeds
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
      For Each obj In Digits(num)
        If chg And 1 Then obj.State = stat And 1
        chg = chg \ 2:stat = stat \ 2
      Next
            'num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
      'For Each obj In Digits2(num)
      ' If chg And 1 Then obj.visible = stat And 1
      ' chg = chg\2 : stat = stat\2
      'Next
        Next
    End If
End Sub

'******************************************************
'****  BALL ROLLING AND DROP SOUNDS
'******************************************************
'
' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

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
  Dim b

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
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * MechVolume, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

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

' '******************************************************
' '****  END BALL ROLLING AND DROP SOUNDS
' '******************************************************

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
Dim objBallShadow(1)

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
'
'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

InitPolarity

''*******************************************
''  Late 80's early 90's
'
Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 5
    x.AddPt "Polarity", 2, 0.4, - 5
    x.AddPt "Polarity", 3, 0.6, - 4.5
    x.AddPt "Polarity", 4, 0.65, - 4.0
    x.AddPt "Polarity", 5, 0.7, - 3.5
    x.AddPt "Polarity", 6, 0.75, - 3.0
    x.AddPt "Polarity", 7, 0.8, - 2.5
    x.AddPt "Polarity", 8, 0.85, - 2.0
    x.AddPt "Polarity", 9, 0.9, - 1.5
    x.AddPt "Polarity", 10, 0.95, - 1.0
    x.AddPt "Polarity", 11, 1, - 0.5
    x.AddPt "Polarity", 12, 1.1, 0
    x.AddPt "Polarity", 13, 1.3, 0

    x.AddPt "Velocity", 0, 0, 1
    x.AddPt "Velocity", 1, 0.16, 1.06
    x.AddPt "Velocity", 2, 0.41, 1.05
    x.AddPt "Velocity", 3, 0.53, 1 '0.982
    x.AddPt "Velocity", 4, 0.702, 0.968
    x.AddPt "Velocity", 5, 0.95,  0.968
    x.AddPt "Velocity", 6, 1.03,  0.945
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
'   Otherwise it should function exactly the same as before

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
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
      If Not IsEmpty(balls(x) ) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x) ) Then
        balldata(x).Data = balls(x)
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef
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
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
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
     'Dim BOT
     'BOT = GetBalls

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

'*****************
' Maths
'*****************

'Dim PI
'PI = 4 * Atn(1)

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

Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
    RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
  DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

'Function Radians(Degrees)
' Radians = Degrees * PI / 180
'End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
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

Dim LFPress, RFPress, LFCount, RFCount
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
Const EOSReturn = 0.035  'mid 80's to early 90's

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
        'BOT = GetBalls

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

Const LiveDistanceMin = 30  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle)  '-1 for Right Flipper
  Dim LiveCatchBounce                                                           'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime
  CatchTime = GameTime - FCount

  If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
    If CatchTime <= LiveCatch * 0.5 Then                        'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    Else
      LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)    'Partial catch when catch happens a bit late
    End If

    If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx = 0
    ball.angmomy = 0
    ball.angmomz = 0
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
'   ZRDT: DROP TARGETS by Rothbauerw
'******************************************************

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
Dim DT18

Set DT18 = (new DropTarget)(sw18, sw18a, sw18p, 18, 0, false)

Dim DTArray
DTArray = Array(DT18)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 110 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 45 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "DropTarget" 'Drop Target Hit sound
Const DTDropSound = "DropTarget" 'Drop Target Drop sound
Const DTResetSound = "resetdrop" 'Drop Target reset sound

Const DTMass = 1 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
  Dim i
  i = DTArrayID(switch)

  PlayTargetSound
  DTArray(i).animate =  DTCheckBrick(Activeball,DTArray(i).prim)
  If DTArray(i).animate = 1 or DTArray(i).animate = 3 or DTArray(i).animate = 4 Then
    DTBallPhysics Activeball, DTArray(i).prim.rotz, DTMass
  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate = -1
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
  For i = 0 to uBound(DTArray)
    If DTArray(i).sw = switch Then DTArrayID = i:Exit Function
  Next
End Function

sub DTBallPhysics(aBall, angle, mass)
  dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
  calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

  aBall.velx = calc1 * cos(rangle) + calc2
  aBall.vely = calc1 * sin(rangle) + calc3
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - tan(bangle) * aball.x + tan(rangle2) * dtprim.x) / (tan(rangle2) - tan(bangle))
  Yintersect = tan(rangle2) * Xintersect + (dtprim.y - tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

  If perpvel > 0 and  perpvelafter <= 0 Then
    If DTEnableBrick = 1 and  perpvel > DTBrickVel and DTBrickVel <> 0 and cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
    DTCheckBrick = 4
  Else
    DTCheckBrick = 0
  End If
End Function

Sub DoDTAnim()
  Dim i
  For i=0 to Ubound(DTArray)
    DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
  Next
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
  dim transz, switchid
  Dim animtime, rangle

  switchid = switch

  Dim ind
  ind = DTArrayID(switchid)

  rangle = prim.rotz * PI / 180

  DTAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If (animate = 1 or animate = 4) and animtime < DTDropDelay Then
    primary.collidable = 0
  If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    DTAnimate = animate
    Exit Function
    elseif (animate = 1 or animate = 4) and animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
  End If

  if animate = 2 Then

    transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
    if prim.transz > -DTDropUnits  Then
      prim.transz = transz
    end if

    prim.rotx = DTMaxBend * cos(rangle)/2
    prim.roty = DTMaxBend * sin(rangle)/2

    if prim.transz <= -DTDropUnits Then
      prim.transz = -DTDropUnits
      secondary.collidable = 0
      DTArray(ind).isDropped = true 'Mark target as dropped
      controller.Switch(Switchid) = 1
      primary.uservalue = 0
      DTAnimate = 0
      Exit Function
    Else
      DTAnimate = 2
      Exit Function
    end If
  End If

  If animate = 3 and animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
  elseif animate = 3 and animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  if animate = -1 Then
    transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1

    If prim.transz = -DTDropUnits Then
      Dim b', BOT
'     BOT = GetBalls

      For b = 0 to UBound(gBOT)
        If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, -25,-10,25,-10,25,25,-25,25) and gBOT(b).z < prim.z+DTDropUnits+25 Then
          gBOT(b).velz = 20
        End If
      Next
    End If

    if prim.transz < 0 Then
      prim.transz = transz
    elseif transz > 0 then
      prim.transz = transz
    end if

    if prim.transz > DTDropUpUnits then
      DTAnimate = -2
      prim.transz = DTDropUpUnits
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = gametime
    end if
    primary.collidable = 0
    secondary.collidable = 1
    DTArray(ind).isDropped = false 'Mark target as not dropped
    controller.Switch(Switchid) = 0

  End If

  if animate = -2 and animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits
    if prim.transz < 0 then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    end If
  End If
End Function

'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS
'******************************************************

' Used for drop targets
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

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
    Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
    Dim rotxy
    rotxy = RotPoint(ax,ay,angle)
    rax = rotxy(0)+px : ray = rotxy(1)+py
    rotxy = RotPoint(bx,by,angle)
    rbx = rotxy(0)+px : rby = rotxy(1)+py
    rotxy = RotPoint(cx,cy,angle)
    rcx = rotxy(0)+px : rcy = rotxy(1)+py
    rotxy = RotPoint(dx,dy,angle)
    rdx = rotxy(0)+px : rdy = rotxy(1)+py

    InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

'******************************************************
'****  END DROP TARGETS
'******************************************************

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
Dim ST19, ST20, ST21, ST22, ST23, ST24, ST59

Set ST19 = (new StandupTarget)(sw19, psw19, 19, 0)
Set ST20 = (new StandupTarget)(sw20, psw20, 20, 0)
Set ST21 = (new StandupTarget)(sw21, psw21, 21, 0)

Set ST22 = (new StandupTarget)(sw22, psw22, 22, 0)
Set ST23 = (new StandupTarget)(sw23, psw23, 23, 0)
Set ST24 = (new StandupTarget)(sw24, psw24, 24, 0)

Set ST59 = (new StandupTarget)(sw59, psw59, 59, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST19, ST20, ST21, ST22, ST23, ST24, ST59)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5     'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 1        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
  Dim i
  i = STArrayID(switch)

  PlayTargetSound
  STArray(i).animate =  STCheckHit(Activeball,STArray(i).primary)

  If STArray(i).animate <> 0 Then
    DTBallPhysics Activeball, STArray(i).primary.orientation, STMass
  End If
  DoSTAnim
End Sub

Function STArrayID(switch)
  Dim i
  For i = 0 to uBound(STArray)
    If STArray(i).sw = switch Then STArrayID = i:Exit Function
  Next
End Function

'Check if target is hit on it's face
Function STCheckHit(aBall, target)
  dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
  rangle = (target.orientation - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

  If perpvel > 0 and  perpvelafter <= 0 Then
    STCheckHit = 1
  ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
    STCheckHit = 1
  Else
    STCheckHit = 0
  End If
End Function

Sub DoSTAnim()
  Dim i
  For i=0 to Ubound(STArray)
    STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prim,STArray(i).sw,STArray(i).animate)
  Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
  Dim animtime

  STAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    STAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If animate = 1 Then
    primary.collidable = 0
    prim.transx = -STMaxOffset
    vpmTimer.PulseSw switch
    STAnimate = 2
    Exit Function
  elseif animate = 2 Then
    prim.transx = prim.transx + STAnimStep
    If prim.transx >= 0 Then
      prim.transx = 0
      primary.collidable = 1
      STAnimate = 0
      Exit Function
    Else
      STAnimate = 2
    End If
  End If
End Function

'******************************************************
'***  END STAND-UP TARGETS
'******************************************************

'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7  'Level of bounces. Recommmended value of 0.7

Sub TargetBouncer(aBall,defvalue)
  Dim zMultiplier, vel, vratio
  If TargetBouncerEnabled = 1 And aball.z < 30 Then
    '   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    vel = BallSpeed(aBall)
    If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
    Select Case Int(Rnd * 6) + 1
      Case 1
        zMultiplier = 0.2 * defvalue
      Case 2
        zMultiplier = 0.25 * defvalue
      Case 3
        zMultiplier = 0.3 * defvalue
      Case 4
        zMultiplier = 0.4 * defvalue
      Case 5
        zMultiplier = 0.45 * defvalue
      Case 6
        zMultiplier = 0.5 * defvalue
    End Select
    aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
    aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
    aBall.vely = aBall.velx * vratio
    '   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    '   debug.print "conservation check: " & BallSpeed(aBall)/vel
  End If
End Sub

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer ActiveBall, 1
End Sub



'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************

Dim LS
Set LS = New SlingshotCorrection
Dim RS
Set RS = New SlingshotCorrection

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
  AddSlingsPt 0, 0.00, - 4
  AddSlingsPt 1, 0.45, - 7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4
End Sub

Sub AddSlingsPt(idx, aX, aY)    'debugger wrapper for adjusting flipper script In-game
  Dim a
  a = Array(LS, RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub

Function RotPoint(x,y,angle)
  dim rx, ry
  rx = x*dCos(angle) - y*dSin(angle)
  ry = x*dSin(angle) + y*dCos(angle)
  RotPoint = Array(rx,ry)
End Function

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
'**** RAMP ROLLING SFX
'******************************************************

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

'/////////////////////////////  PLASTIC RAMPS FLAPS - SOUNDS  ////////////////////////////
Sub RandomSoundRampFlapUp()
' debug.print "flap up"
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_1"), FlapSoundLevel
    Case 2 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_2"), FlapSoundLevel
    Case 3 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_3"), FlapSoundLevel
  End Select
End Sub

Sub RandomSoundRampFlapDown()
' debug.print "flap down"
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_1"), FlapSoundLevel
    Case 2 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_2"), FlapSoundLevel
    Case 3 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_3"), FlapSoundLevel
  End Select
End Sub


'/////////////////////////////  RAMP COLLISIONS  ////////////////////////////

dim LRHit1_volume, LRHit2_volume, LRHit3_volume
dim RRHit1_volume, RRHit2_volume, RRHit3_volume

Dim LeftRampSoundLevel
Dim RightRampSoundLevel
Dim RampFallbackSoundLevel

Dim FlapSoundLevel

'///////////////////////-----Ramps-----///////////////////////
'///////////////////////-----Plastic Ramps-----///////////////////////
LeftRampSoundLevel = 0.1                        'volume level; range [0, 1]
RightRampSoundLevel = 0.1                       'volume level; range [0, 1]
RampFallbackSoundLevel = 0.2                      'volume level; range [0, 1]

'///////////////////////-----Ramp Flaps-----///////////////////////
FlapSoundLevel = 0.8                          'volume level; range [0, 1]

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************

'////////////////////////////////////////////////////////////////////////////////
'////          Mechanical Sounds, by Fleep                                   ////
'////                     Last Updated: January, 2022                        ////
'////////////////////////////////////////////////////////////////////////////////
'
'/////////////////////////////////  CARTRIDGES  /////////////////////////////////
'
'//  General Mechanical Sounds Cartridges:
Const Cartridge_Bumpers         = "WS_PBT_REV01" 'Williams Pinbot Cartridge REV01
Const Cartridge_Slingshots        = "WS_PBT_REV01" 'Williams Pinbot Cartridge REV01
Const Cartridge_Flippers        = "WS_PBT_REV01" 'Williams Pinbot Cartridge REV01
Const Cartridge_Kickers         = "WS_WHD_REV01"
Const Cartridge_Diverters       = "WS_DNR_REV01" 'Williams Diner Cartridge REV01
Const Cartridge_Knocker         = "WS_WHD_REV02" 'Williams Whirlwind Cartridge REV02
Const Cartridge_Relays          = "WS_WHD_REV01"
Const Cartridge_Trough          = "WS_WHD_REV01"
Const Cartridge_Rollovers       = "WS_WHD_REV01"
Const Cartridge_Targets         = "WS_WHD_REV01"
Const Cartridge_Gates         = "WS_WHD_REV01"
Const Cartridge_Spinner         = "SY_TNA_REV01" 'Spooky Total Nuclear Annihilation Cartridge REV01
Const Cartridge_Rubber_Hits       = "WS_WHD_REV01"
Const Cartridge_Metal_Hits        = "WS_WHD_REV01"
Const Cartridge_Plastic_Hits      = "WS_WHD_REV01"
Const Cartridge_Wood_Hits       = "WS_WHD_REV01"
Const Cartridge_Cabinet_Sounds      = "WS_WHD_REV01"
Const Cartridge_Drain         = "WS_WHD_REV01"
Const Cartridge_Apron         = "WS_WHD_REV01"
Const Cartridge_Ball_Roll       = "BY_TOM_REV01" 'Bally Theatre of Magic Cartridge REV01
Const Cartridge_BallBallCollision   = "BY_WDT_REV01" 'Bally WHO Dunnit Cartridge REV01
Const Cartridge_Ball_Drop_Bump      = "WS_WHD_REV01"
Const Cartridge_Plastic_Ramps     = "WS_WHD_REV01"
Const Cartridge_Metal_Ramps       = "WS_WHD_REV01"
Const Cartridge_Ball_Guides       = "WS_WHD_REV01"
Const Cartridge_Table_Specifics     = "WS_WHD_REV01"

'////////////////////////////  SOUND SOURCE CREDITS  ////////////////////////////
'//  Special thanks go to the following contributors who have provided audio
'//  footage recordings:
'//
'//  Williams Whirlwind - Blackmoor, wrd1972
'//  Williams Diner - Nick Rusis
'//  Spooky Total Nuclear Annihilation - WildDogArcade, Ed and Gary
'//  Bally Theatre of Magic - CalleV, nickbuol
'//  Bally WHO Dunnit - Amazaley1
'//  Williams Pinbot - major_drain_pinball

'///////////////////////////  SOLENOIDS (COILS) CONFIG  /////////////////////////

'//  FLIPPER COILS:
'//  Flippers in this table: Lower Left Flipper, Lower Right Flipper, Upper Right Fliiper
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel
Dim FlipperLeftLowerHitParm, FlipperRightUpperHitParm, FlipperRightLowerHitParm

'//  Flipper Up Attacks initialize during playsound subs
Dim FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel

FlipperUpSoundLevel = 1
FlipperDownSoundLevel = 0.65
FlipperUpAttackMinimumSoundLevel = 0.010
FlipperUpAttackMaximumSoundLevel = 0.435

'//  Flipper Hit Param initialize with FlipperUpSoundLevel
'//  and dynamically modified calculated by ball flipper collision
FlipperLeftLowerHitParm = FlipperUpSoundLevel
FlipperRightLowerHitParm = FlipperUpSoundLevel
FlipperRightUpperHitParm = FlipperUpSoundLevel

Dim Solenoid_OutholeKicker_SoundLevel, Solenoid_ShooterFeeder_SoundLevel
Dim Solenoid_RightRampLifter_SoundLevel, Solenoid_LeftLockingKickback_SoundLevel
Dim Solenoid_TopEject_SoundLevel, Solenoid_Knocker_SoundLevel, Solenoid_DropTargetReset_SoundLevel
Dim Solenoid_Diverter_Enabled_SoundLevel, Solenoid_Diverter_Hold_SoundLevel, Solenoid_Diverter_Disabled_SoundLevel
Dim Solenoid_UnderPlayfieldKickbig_SoundLevel, Solenoid_Bumper_SoundMultiplier
Dim Solenoid_Slingshot_SoundLevel, Solenoid_RightRampDown_SoundLevel, AutoPlungerSoundLevel

AutoPlungerSoundLevel = 1                       'volume level; range [0, 1]
Solenoid_OutholeKicker_SoundLevel = 1
Solenoid_ShooterFeeder_SoundLevel = 1
Solenoid_RightRampLifter_SoundLevel = 0.3
Solenoid_RightRampDown_SoundLevel = 0.3
Solenoid_LeftLockingKickback_SoundLevel = 1
Solenoid_TopEject_SoundLevel = 1
Solenoid_Knocker_SoundLevel = 1
Solenoid_DropTargetReset_SoundLevel = 1
Solenoid_Diverter_Enabled_SoundLevel = 1
Solenoid_Diverter_Hold_SoundLevel = 0.7
Solenoid_Diverter_Disabled_SoundLevel = 0.4
Solenoid_UnderPlayfieldKickbig_SoundLevel = 1
Solenoid_Bumper_SoundMultiplier = 0.004 '8
Solenoid_Slingshot_SoundLevel = 1

Dim RelayLowerGISoundLevel, RelayUpperGISoundLevel, RelaySolenoidACSelectSoundLevel, RelayFlasherSoundLevel
RelayLowerGISoundLevel = 0.45
RelayUpperGISoundLevel = 0.45
RelaySolenoidACSelectSoundLevel = 0.3
RelayFlasherSoundLevel = 0.015

Dim Solenoid_BlowerMotor_SoundLevel, Solenoid_SpinWheelsMotor_SoundLevel
Solenoid_BlowerMotor_SoundLevel = 0.2
Solenoid_SpinWheelsMotor_SoundLevel = 0.2

'////////////////////////////  SWITCHES SOUND CONFIG  ///////////////////////////
Dim Switch_Gate_SoundLevel, SpinnerSoundLevel, RolloverSoundLevel, OutLaneRolloverSoundLevel, TargetSoundFactor

Switch_Gate_SoundLevel = 1
SpinnerSoundLevel = 0.1
RolloverSoundLevel = 0.55
OutLaneRolloverSoundLevel = 0.8
TargetSoundFactor = 0.8

'////////////////////  BALL HITS, BUMPS, DROPS SOUND CONFIG  ////////////////////
Dim BallWithBallCollisionSoundFactor, BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor
Dim WallImpactSoundFactor, MetalImpactSoundFactor, WireformAntiRebountRailSoundFactor
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor
Dim BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor, OutlaneWallsSoundFactor
Dim EjectBallBumpSoundLevel, HeadSaucerSoundLevel, EjectHoleEnterSoundLevel
Dim RightRampMetalWireDropToPlayfieldSoundLevel, LeftPlasticRampDropToLockSoundLevel, LeftPlasticRampDropToPlayfieldSoundLevel
Dim CellarLeftEnterSoundLevel, CellarRightEnterSoundLevel, CellerKickouBallDroptSoundLevel

BallWithBallCollisionSoundFactor = 3.2
BallBouncePlayfieldSoftFactor = 0.0015
BallBouncePlayfieldHardFactor = 0.0075
WallImpactSoundFactor = 0.075
MetalImpactSoundFactor = 0.075
RubberStrongSoundFactor = 0.045
RubberWeakSoundFactor = 0.055
RubberFlipperSoundFactor = 0.65
BottomArchBallGuideSoundFactor = 0.2
FlipperBallGuideSoundFactor = 0.015
WireformAntiRebountRailSoundFactor = 0.04
OutlaneWallsSoundFactor = 1
EjectBallBumpSoundLevel = 1
RightRampMetalWireDropToPlayfieldSoundLevel = 1
LeftPlasticRampDropToLockSoundLevel = 1
LeftPlasticRampDropToPlayfieldSoundLevel = 1
EjectHoleEnterSoundLevel = 0.75
HeadSaucerSoundLevel = 0.15
CellerKickouBallDroptSoundLevel = 1
CellarLeftEnterSoundLevel = 0.85
CellarRightEnterSoundLevel = 0.85

'///////////////////////  OTHER PLAYFIELD ELEMENTS CONFIG  //////////////////////
Dim RollingSoundFactor, RollingOnDiscSoundFactor, BallReleaseShooterLaneSoundLevel
Dim LeftPlasticRampEnteranceSoundLevel, RightPlasticRampEnteranceSoundLevel
Dim LeftPlasticRampRollSoundFactor, RightPlasticRampRollSoundFactor
Dim LeftMetalWireRampRollSoundFactor, RightPlasticRampHitsSoundLevel, LeftPlasticRampHitsSoundLevel
Dim SpinningDiscRolloverSoundFactor, SpinningDiscRolloverBumpSoundLevel
Dim LaneSoundFactor, LaneEnterSoundFactor, InnerLaneSoundFactor
Dim LaneLoudImpactMinimumSoundLevel, LaneLoudImpactMaximumSoundLevel
Dim GateSoundLevel

GateSoundLevel = 0.5 / 5      'volume level; range [0, 1]
RollingSoundFactor = 50
RollingOnDiscSoundFactor = 1.5
BallReleaseShooterLaneSoundLevel = 1
LeftPlasticRampEnteranceSoundLevel = 0.1
RightPlasticRampEnteranceSoundLevel = 0.1
LeftPlasticRampRollSoundFactor = 0.2
RightPlasticRampRollSoundFactor = 0.2
LeftMetalWireRampRollSoundFactor = 1
RightPlasticRampHitsSoundLevel = 1
LeftPlasticRampHitsSoundLevel = 1
SpinningDiscRolloverSoundFactor = 0.05
SpinningDiscRolloverBumpSoundLevel = 0.3
LaneEnterSoundFactor = 0.9
InnerLaneSoundFactor = 0.0005
LaneSoundFactor = 0.0004
LaneLoudImpactMinimumSoundLevel = 0
LaneLoudImpactMaximumSoundLevel = 0.4


'///////////////////////////  CABINET SOUND PARAMETERS  /////////////////////////
Dim NudgeLeftSoundLevel, NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel
Dim PlungerReleaseSoundLevel, PlungerPullSoundLevel, CoinSoundLevel

NudgeLeftSoundLevel = 1
NudgeRightSoundLevel = 1
NudgeCenterSoundLevel = 1
StartButtonSoundLevel = 0.1
PlungerReleaseSoundLevel = 1
PlungerPullSoundLevel = 1
CoinSoundLevel = 1


'///////////////////////////  MISC SOUND PARAMETERS  ////////////////////////////
Dim LutToggleSoundLevel :
LutToggleSoundLevel = 0.5


'////////////////////////////////  SOUND HELPERS  ///////////////////////////////
Dim SoundOn : SoundOn = 1
Dim SoundOff : SoundOff = 0
Dim Up : Up = 0
Dim Down : Down = 1
Dim RampUp : RampUp = 1
Dim RampDown : RampDown = 0
Dim RampDownSlow : RampDownSlow = 1
Dim RampDownFast : RampDownFast = 2
Dim CircuitA : CircuitA = 0
Dim CircuitC : CircuitC = 1

'//  Helper for Main (Lower) flippers dampened stroke
Dim BallNearLF : BallNearLF = 0
Dim BallNearRF : BallNearRF = 0

Sub TriggerBallNearLF_Hit()
  'Debug.Print "BallNearLF = 1"
  BallNearLF = 1
End Sub

Sub TriggerBallNearLF_UnHit()
  'Debug.Print "BallNearLF = 0"
  BallNearLF = 0
End Sub

Sub TriggerBallNearRF_Hit()
  'Debug.Print "BallNearRF = 1"
  BallNearRF = 1
End Sub

Sub TriggerBallNearRF_UnHit()
  'Debug.Print "BallNearLF = 0"
  BallNearRF = 0
End Sub


'///////////////////////  SOUND PLAYBACK SUBS / FUNCTIONS  //////////////////////
'//////////////////////  POSITIONAL SOUND PLAYBACK METHODS  /////////////////////

Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, min(1,aVol) * MechVolume, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, min(1,aVol) * MechVolume, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
    PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
    PlaySound playsoundparams, -1, min(1,aVol) * MechVolume, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub


'//////////////////////  SUPPORTING BALL & SOUND FUNCTIONS  /////////////////////

Function AudioFade(tableobj)
  Dim tmp
  Select Case PositionalSoundPlaybackConfiguration
    Case 1
      AudioFade = 0
    Case 2
      AudioFade = 0
    Case 3
      tmp = tableobj.y * 2 / tableheight-1
      If tmp > 0 Then
        AudioFade = Csng(tmp ^5) 'was 10
      Else
        AudioFade = Csng(-((- tmp) ^5) ) ' was 10
      End If
  End Select
End Function

'//  Calculates the pan for a tableobj based on the X position on the table.
Function AudioPan(tableobj)
  Dim tmp
  Select Case PositionalSoundPlaybackConfiguration
    Case 1
      AudioPan = 0
    Case 2
      tmp = tableobj.x * 2 / tablewidth-1
      If tmp > 0 Then
        AudioPan = Csng(tmp ^5) ' was 10
      Else
        AudioPan = Csng(-((- tmp) ^5) ) ' was 10
      End If
    Case 3
      tmp = tableobj.x * 2 / tablewidth-1

      If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
      Else
        AudioPan = Csng(-((- tmp) ^10) )
      End If
  End Select
End Function

'//  Calculates the volume of the sound based on the ball speed
Function Vol(ball)
  Vol = Csng(BallVel(ball) ^2)
End Function

'//  Calculates the pitch of the sound based on the ball speed
Function Pitch(ball)
    Pitch = BallVel(ball) * 20
End Function

'//  Calculates the ball speed
Function BallVel(ball)
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'//  Calculates the roll volume of the sound based on the ball speed
Dim TempBallVel
Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  TempBallVel = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/RollingSoundFactor)
  If TempBallVel = 1 Then TempBallVel = 0.999
  If TempBallVel = 0 Then TempBallVel = 0.001
  'debug.print TempBallVel
  TempBallVel = Csng(1/(1+(0.275*(((0.75*TempBallVel)/(1-TempBallVel))^(-2)))))
  VolPlayfieldRoll = TempBallVel
End Function

'//  Calculates the roll volume of the sound based on the ball speed
Function VolSpinningDiscRoll(ball)
  VolSpinningDiscRoll = RollingOnDiscSoundFactor * 0.1 * Csng(BallVel(ball) ^3)
End Function

'//  Calculates the roll volume of the sound based on the ball speed
Dim TempBallVelPlastic
Function VolPlasticMetalRampRoll(ball)
  'VolPlasticMetalRampRoll = RollingOnDiscSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
  TempBallVelPlastic = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/RollingSoundFactor)
  If TempBallVelPlastic = 1 Then TempBallVelPlastic = 0.999
  If TempBallVelPlastic = 0 Then TempBallVelPlastic = 0.001
  'debug.print TempBallVel
  TempBallVelPlastic = Csng(1/(1+(0.275*(((0.75*TempBallVelPlastic)/(1-TempBallVelPlastic))^(-2)))))
  VolPlasticMetalRampRoll = TempBallVelPlastic
End Function

'//  Calculates the roll pitch of the sound based on the ball speed
Dim TempPitchBallVel
Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  'PitchPlayfieldRoll = BallVel(ball) ^2 * 15
  'PitchPlayfieldRoll = Csng(BallVel(ball))/50 * 10000
  'PitchPlayfieldRoll = (1-((Csng(BallVel(ball))/50)^0.2)) * 20000

  'PitchPlayfieldRoll = (2*((Csng(BallVel(ball)))^0.7))/(2+(Csng(BallVel(ball)))) * 16000
  TempPitchBallVel = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/50)
  If TempPitchBallVel = 1 Then TempPitchBallVel = 0.999
  If TempPitchBallVel = 0 Then TempPitchBallVel = 0.001
  TempPitchBallVel = Csng(1/(1+(0.275*(((0.75*TempPitchBallVel)/(1-TempPitchBallVel))^(-2))))) * 10000
  PitchPlayfieldRoll = TempPitchBallVel
End Function

'//  Calculates the pitch of the sound based on the ball speed.
'//  Used for plastic ramps roll sound
Function PitchPlasticRamp(ball)
    PitchPlasticRamp = BallVel(ball) * 20
End Function

'//  Determines if a point (px,py) in inside a circle with a center of
'//  (cx,cy) coordinates and circleradius
Function InCircle(px,py,cx,cy,circleradius)
  Dim distance
  distance = SQR(((px-cx)^2) + ((py-cy)^2))

  If (distance < circleradius) Then
    InCircle = True
  Else
    InCircle = False
  End If
End Function

'///////////////////////////  PLAY SOUNDS SUBROUTINES  //////////////////////////
'//
'//  These Subroutines implement all mechanical playsounds including timers
'//
'//////////////////////////  GENERAL SOUND SUBROUTINES  /////////////////////////
Sub SoundStartButton()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Start_Button"), StartButtonSoundLevel, StartButtonPosition
End Sub

Sub SoundPlungerPull()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Pull_Slow"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerPullStop()
  StopSound Cartridge_Cabinet_Sounds & "_Plunger_Pull_Slow"
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Release_Ball_" & Int(Rnd*3)+1), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Release_Empty"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeLeftSoundLevel * MechVolume, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeRightSoundLevel * MechVolume, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySoundAtLevelStatic ("Nudge_" & Int(Rnd*3)+1), NudgeCenterSoundLevel * MechVolume, Outhole
End Sub

'///////////////////////  JP'S VP10 BALL COLLISION SOUND  ///////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  if abs(ball1.vely) < 1 And abs(ball2.vely) < 1 And InRect(ball1.x, ball1.y, 360,730,420,730,420,875,360,875) then
    exit sub  'don't rattle the locked balls
  end if
  PlaySound (Cartridge_BallBallCollision & "_BallBall_Collide_" & Int(Rnd*7)+1), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * MechVolume, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  FlipperCradleCollision ball1, ball2, velocity
End Sub

'///////////////////////////  OUTLANE ROLLOVER SOUNDS  //////////////////////////
Sub RandomSoundOutlaneRollover()
  PlaySoundAtLevelActiveBall (Cartridge_Rollovers & "_Rollover_Outlane_" & Int(Rnd*4)+1), OutLaneRolloverSoundLevel
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

'//////////////////////////  STADNING TARGET HIT SOUNDS  ////////////////////////
Sub RandomSoundTargetHitStrong()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_5",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_6",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_7",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_8",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
  End Select
End Sub

Sub RandomSoundTargetHitWeak()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_1",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_2",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_3",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_4",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
  End Select
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  /////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_2"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_12"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 3 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_14"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 4 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_18"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 5 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_19"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_20"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 7 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_21"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  Select Case Int(Rnd*12)+1
    Case 1 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_1"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 2 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_3"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 3 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_7"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 4 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_8"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 5 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_9"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 6 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_11"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 7 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_13"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 8 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_15"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 9 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_16"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 10 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_17"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 11 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_22"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 12 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_23"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
  End Select
End Sub

'////////////////////////////////////  DRAIN  ///////////////////////////////////
'///////////////////////////////  OUTHOLE SOUNDS  ///////////////////////////////
Sub RandomSoundOutholeHit(sw)
  PlaySoundAtLevelStatic (Cartridge_Trough & "_Outhole_Drain_Hit_" & Int(Rnd*4)+1), Solenoid_OutholeKicker_SoundLevel, sw
End Sub

Sub RandomSoundOutholeKicker()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Trough & "_Outhole_Kicker_" & Int(Rnd*4)+1,DOFContactors), Solenoid_OutholeKicker_SoundLevel, Outhole
End Sub

'/////////////////////  BALL SHOOTER FEEDER SOLENOID SOUNDS  ////////////////////
Sub RandomSoundShooterFeeder()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Trough & "_Shooter_Feeder_" & Int(Rnd*6)+1,DOFContactors), Solenoid_ShooterFeeder_SoundLevel, Outhole
End Sub

'///////  SHOOTER LANE - BALL RELEASE ROLL IN SHOOTER LANE SOUND - SOUND  ///////
Sub SoundBallReleaseShooterLane(toggle)
  Select Case toggle
    Case SoundOn
      PlaySoundAtLevelActiveBall (Cartridge_Table_Specifics & "_Ball_Launch_from_Shooter_Lane"), BallReleaseShooterLaneSoundLevel
    Case SoundOff
      StopSound Cartridge_Table_Specifics & "_Ball_Launch_from_Shooter_Lane"
  End Select
End Sub

'//////////////////////////////  KNOCKER SOLENOID  //////////////////////////////
Sub KnockerSolenoid(enabled)
  'PlaySoundAtLevelStatic SoundFX(Cartridge_Knocker & "_Knocker_Coil",DOFKnocker), Solenoid_Knocker_SoundLevel, KnockerPosition
  if Enabled then
    PlaySound SoundFX(Cartridge_Knocker & "_Knocker_Coil",DOFKnocker), 0, Solenoid_Knocker_SoundLevel
  end if
End Sub

'//////////////////////////  SLINGSHOT SOLENOID SOUNDS  /////////////////////////
Sub RandomSoundSlingshotLeft()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Slingshots & "_Slingshot_Left_" & Int(Rnd*26)+1,DOFContactors), Solenoid_Slingshot_SoundLevel, LeftSlingshotPosition
End Sub

Sub RandomSoundSlingshotRight()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Slingshots & "_Slingshot_Right_" & Int(Rnd*25)+1,DOFContactors), Solenoid_Slingshot_SoundLevel, RightSlingshotPosition
End Sub

'///////////////////////////  BUMPER SOLENOID SOUNDS  ///////////////////////////
'////////////////////////////////  BUMPERS - TOP  ///////////////////////////////
Sub RandomSoundBumperLeft(Bump)
' Debug.Print Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Left_" & Int(Rnd*22)+1,DOFContactors), Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier, Bump
End Sub

Sub RandomSoundBumperUp(Bump)
  'Debug.Print Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Up_" & Int(Rnd*25)+1,DOFContactors), Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier, Bump
End Sub

Sub RandomSoundBumperLow(Bump)
  'Debug.Print Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Low_" & Int(Rnd*28)+1,DOFContactors), Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier, Bump
End Sub



'///////////////////////  FLIPPER BATS SOUND SUBROUTINES  ///////////////////////
'//////////////////////  FLIPPER BATS SOLENOID CORE SOUND  //////////////////////
Sub RandomSoundFlipperLowerLeftUpFullStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_" & RndInt(1,10),DOFFlippers), FlipperLeftLowerHitParm, Flipper
End Sub

Sub RandomSoundFlipperLowerLeftUpDampenedStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Dampened_Stroke_" & RndInt(1,23),DOFFlippers), FlipperLeftLowerHitParm * 1.2, Flipper
End Sub

Sub RandomSoundFlipperLowerRightUpFullStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_" & RndInt(1,11),DOFFlippers), FlipperRightLowerHitParm, Flipper
End Sub

Sub RandomSoundFlipperLowerRightUpDampenedStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Dampened_Stroke_" & RndInt(1,23),DOFFlippers), FlipperLeftLowerHitParm * 1.2, Flipper
End Sub

Sub RandomSoundFlipperLowerLeftReflip(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Reflip_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperLowerRightReflip(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Reflip_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperLowerLeftDown(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Down_" & RndInt(1,10),DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperLowerRightDown(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Down_" & RndInt(1,11),DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundLowerLeftQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_QuickFlip_Up_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, LeftFlipper
End Sub

Sub RandomSoundLowerRightQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_QuickFlip_Up_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, RightFlipper
End Sub

Sub StopAnyFlipperLowerLeftUp()
  Dim anyFullStrokeSound
  Dim anyDampenedStrokeSound
  For anyFullStrokeSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_" & anyFullStrokeSound)
  Next
  For anyDampenedStrokeSound = 1 to 23
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Dampened_Stroke_" & anyDampenedStrokeSound)
  Next
End Sub

Sub StopAnyFlipperLowerRightUp()
  Dim anyFullStrokeSound
  Dim anyDampenedStrokeSound
  For anyFullStrokeSound = 1 to 11
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_" & anyFullStrokeSound)
  Next
  For anyDampenedStrokeSound = 1 to 23
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Dampened_Stroke_" & anyDampenedStrokeSound)
  Next
End Sub

Sub StopAnyFlipperLowerLeftDown()
  Dim anyFullDownSound
  For anyFullDownSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Down_" & anyFullDownSound)
  Next
End Sub

Sub StopAnyFlipperLowerRightDown()
  Dim anyFullDownSound
  For anyFullDownSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Down_" & anyFullDownSound)
  Next
End Sub

Sub FlipperHoldCoilLeft(toggle, flipper)
  Select Case toggle
    Case SoundOn
      PlaySoundAtLevelExistingStaticLoop (Cartridge_Flippers & "_Flipper_Hold_Coil_Low_Frequencies_Loop_Lower_Left"), FlipperHoldSoundLevel, flipper
      PlaySoundAtLevelExistingStaticLoop (Cartridge_Flippers & "_Flipper_Hold_Coil_Med_Frequencies_Loop_Lower_Left"), FlipperHoldSoundLevel, flipper
    Case SoundOff
      StopSound Cartridge_Flippers & "_Flipper_Hold_Coil_Low_Frequencies_Loop_Lower_Left"
      StopSound Cartridge_Flippers & "_Flipper_Hold_Coil_Med_Frequencies_Loop_Lower_Left"
  End Select
End Sub

Sub FlipperHoldCoilRight(toggle, flipper)
  Select Case toggle
    Case SoundOn
      PlaySoundAtLevelExistingStaticLoop (Cartridge_Flippers & "_Flipper_Hold_Coil_Low_Frequencies_Loop_Lower_Right"), FlipperHoldSoundLevel, flipper
      PlaySoundAtLevelExistingStaticLoop (Cartridge_Flippers & "_Flipper_Hold_Coil_Med_Frequencies_Loop_Lower_Right"), FlipperHoldSoundLevel, flipper
    Case SoundOff
      StopSound Cartridge_Flippers & "_Flipper_Hold_Coil_Low_Frequencies_Loop_Lower_Right"
      StopSound Cartridge_Flippers & "_Flipper_Hold_Coil_Med_Frequencies_Loop_Lower_Right"
  End Select
End Sub

'///////////////////////  FLIPPER BATS BALL COLLIDE SOUND  //////////////////////
dim angdamp, veldamp
angdamp = 0.2
veldamp = 0.8

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm

  If parm => 22 Then
    ' Strong hit safe values boundary
    ' Flipper stroke dampened
    FlipperLeftLowerHitParm = FlipperUpSoundLevel * 0.1
  Else
    If parm =< 1 Then
      ' Weak hit safe values boundary
      ' Flipper stroke full
      FlipperLeftLowerHitParm = FlipperUpSoundLevel
    Else
      ' Fully modulated hit
      FlipperLeftLowerHitParm = FlipperUpSoundLevel * (1-(parm/25))
    End If
  End If

  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm

  If parm => 22 Then
    ' Strong hit safe values boundary
    ' Flipper stroke dampened
    FlipperRightLowerHitParm = FlipperUpSoundLevel * 0.1
  Else
    If parm =< 1 Then
      ' Weak hit safe values boundary
      ' Flipper stroke full
      FlipperRightLowerHitParm = FlipperUpSoundLevel
    Else
      ' Fully modulated hit
      FlipperRightLowerHitParm = FlipperUpSoundLevel * (1-(parm/25))
    End If
  End If

  RandomSoundRubberFlipper(parm)
End Sub


Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Flipper_Hit_" & Int(Rnd*7)+1), parm / 25 * RubberFlipperSoundFactor
End Sub

'//////////////////////////  SOLENOID A/C SELECT RELAY  /////////////////////////
Sub Sound_Solenoid_AC(toggle)
  Select Case toggle
    Case CircuitA
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_A"), RelaySolenoidACSelectSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_A"), RelaySolenoidACSelectSoundLevel, ACSelectPosition
    Case CircuitC
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_C"), RelaySolenoidACSelectSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_C"), RelaySolenoidACSelectSoundLevel, ACSelectPosition
  End Select
End Sub

'//////////////////////////  GENERAL ILLUMINATION RELAYS  ///////////////////////
Sub Sound_LowerGI_Relay(toggle)
  Select Case toggle
    Case SoundOn
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_On"), RelayLowerGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_On"), RelayLowerGISoundLevel, GIUpperPosition
    Case SoundOff
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_Off"), RelayLowerGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_Off"), RelayLowerGISoundLevel, GIUpperPosition
  End Select
End Sub

Sub Sound_UpperGI_Relay(toggle)
  Select Case toggle
    Case SoundOn
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_On"), RelayUpperGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_On"), RelayUpperGISoundLevel, GILowerPosition
    Case SoundOff
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_Off"), RelayUpperGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_Off"), RelayUpperGISoundLevel, GILowerPosition
  End Select
End Sub

'///////////////////////////////  FLASHERS RELAY  ///////////////////////////////
Sub Sound_Flasher_Relay(toggle, tableobj)
  Select Case toggle
    Case SoundOn
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_On"), RelayFlasherSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_On"), RelayFlasherSoundLevel, tableobj
    Case SoundOff
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_Off"), RelayFlasherSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_Off"), RelayFlasherSoundLevel, tableobj
  End Select
End Sub

'////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  /////////////////////
'/////////////////////////////  RUBBERS AND POSTS  //////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ///////////////////////////////
Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundRubberStrong()
  End if
  If finalspeed <= 10 then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  /////////////////////
Sub RandomSoundRubberStrong()
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 10 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_10"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub

'///////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  /////////////////////
Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_" & Int(Rnd*8)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'///////////////////////////////  WALL IMPACTS  /////////////////////////////////
Sub RandomSoundWall()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor * 0.05
      Case 4 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  WALL IMPACTS EVENTS  ////////////////////////////

Sub Wall46_Hit()
  RandomSoundMetal()
End Sub

sub HitsMetal_Hit(IDX)
' debug.print "metal hit"
  RandomSoundMetal()
end sub

'RandomSoundBottomArchBallGuideSoftHit - Soft Bounces
Sub Wall78_Hit() : RandomSoundBottomArchBallGuideSoftHit() : End Sub
Sub Wall26_Hit() : RandomSoundBottomArchBallGuideSoftHit() : End Sub

'RandomSoundBottomArchBallGuideHardHit - Hard Hit
Sub Wall062_Hit() : RandomSoundBottomArchBallGuideHardHit() : End Sub
Sub Wall063_Hit() : RandomSoundBottomArchBallGuideHardHit() : End Sub

'RandomSoundFlipperBallGuide
Sub Wall49_Hit() : RandomSoundFlipperBallGuide() : End Sub
Sub Wall48_Hit() : RandomSoundFlipperBallGuide() : End Sub

'Outlane - Walls & Primitives
Sub Wall70_Hit() : RandomSoundOutlaneWalls() : End Sub
Sub Wall5_Hit() : RandomSoundOutlaneWalls() : End Sub
Sub Wall73_Hit() : RandomSoundOutlaneWalls() : End Sub
Sub Wall24_Hit() : RandomSoundOutlaneWalls() : End Sub

'////////////////////////////  INNER LEFT LANE WALLS  ///////////////////////////
Sub Wall226_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

Sub Wall225_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub


'right arch
Sub Wall141_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

Sub Wall132_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

Sub Wall224_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub


'/////////////////////////////  METAL TOUCH SOUNDS  /////////////////////////////
Sub RandomSoundMetal()
  Select Case Int(Rnd*20)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_1"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_2"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_3"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_4"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_5"), Vol(ActiveBall) * 0.02 * MetalImpactSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_6"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_7"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_8"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_9"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 10 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_10"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 11 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_11"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 12 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_12"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 13 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_13"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 14 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_14"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 15 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_15"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 16 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_16"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 17 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_17"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 18 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_18"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 19 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_19"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 20 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_20"), Vol(ActiveBall) * MetalImpactSoundFactor
  End Select
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub HitsWoods_Hit(idx)
' debug.print "wood hit"
  RandomSoundWood()
End Sub

Sub RandomSoundWood()
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

'///////////////////////////////  OUTLANES WALLS  ///////////////////////////////
Sub RandomSoundOutlaneWalls()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Outlane_Wall_" & Int(Rnd*9)+1), OutlaneWallsSoundFactor
End Sub

'///////////////////////////  BOTTOM ARCH BALL GUIDE  ///////////////////////////
'///////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////
Sub RandomSoundBottomArchBallGuideSoftHit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Arch_Ball_Guide_Hit_Soft_" & Int(Rnd*4)+1), BottomArchBallGuideSoundFactor
End Sub


'//////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Arch_Ball_Guide_Hit_Hard_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 3
End Sub

'//////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End if
End Sub

'/////////////////////////  WIREFORM ANTI-REBOUNT RAILS  ////////////////////////
Sub RandomSoundWireformAntiRebountRail()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed >= 10 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_3"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_4"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_5"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_6"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_7"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
    End Select
  End if
  If finalspeed < 10 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_1"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_2"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_8"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
    End Select
  End if
End Sub

'////////////////////////////  LANES AND INNER LOOPS  ///////////////////////////
'////////////////////  INNER LOOPS - LEFT ENTRANCE - EVENTS  ////////////////////
Sub LeftInnerLaneTriggerUp_Hit()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    If ActiveBall.VelY < 0 Then RandomSoundInnerLaneEnter()
  End If
End Sub

Sub LeftInnerLaneTriggerDown_Hit()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 7 then
    If ActiveBall.VelY > 0 Then RandomSoundInnerLaneEnter()
  End If
End Sub

'/////////////////////  INNER LOOPS - LEFT ENTRANCE - SOUNDS  ///////////////////
Sub RandomSoundInnerLaneEnter()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Ball_Guide_Hit_" & Int(Rnd*20)+1), Vol(ActiveBall) * InnerLaneSoundFactor
End Sub

'//////////////////////////////  LEFT LANE ENTRANCE  ////////////////////////////
'/////////////////////////  LEFT LANE ENTRANCE - EVENTS  ////////////////////////
Sub LeftLaneTrigger_Hit()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    If ActiveBall.VelY < 0 Then PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Lane_Left_Ball_Enter_Hit"), Vol(ActiveBall) * LaneEnterSoundFactor : RandomSoundLaneLeftEnter()
  End If
End Sub

'/////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////
Sub RandomSoundLaneLeftEnter()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Lane_Left_Ball_Roll_" & Int(Rnd*2)+1), Vol(ActiveBall) * LaneSoundFactor
End Sub

'/////////////////////////////  RIGHT LANE ENTRANCE  ////////////////////////////
'////////////////////////  RIGHT LANE ENTRANCE - EVENTS  ////////////////////////
Sub RightLaneTrigger_Hit()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    If ActiveBall.VelY < 0 Then PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Lane_Right_Ball_Enter_Hit"), Vol(ActiveBall) * LaneEnterSoundFactor : RandomSoundLaneRightEnter()
  End If
End Sub

'/////////////////  RIGHT LANE ENTRANCE (RIGHT ORBIT) - SOUNDS  /////////////////
Sub RandomSoundLaneRightEnter()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Lane_Right_Ball_Roll_" & Int(Rnd*3)+1), Vol(ActiveBall) * LaneSoundFactor
End Sub

'/////////  PLASTIC LEFT RAMP - RIGHT EXIT HOLE - TO PLAYFIELD - EVENT  /////////

'sub BallDrop1_hit  'sometimes drop sound cannot be heard here.
''  debug.print activeball.velz
' RandomSoundRightRampRightExitDropToPlayfield(BallDrop1)
'end sub


Sub BallDrop2_Hit()
  if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 70
  activeball.AngMomZ = abs(activeball.AngMomZ) * 3
  RandomSoundRightRampRightExitDropToPlayfield(BallDrop2)
' Call SoundRightPlasticRampPart2(SoundOff, ballvariablePlasticRampTimer1)
End Sub

'ss ramp
sub BallDrop3_hit
' debug.print "SS drop"
  RandomSoundRightRampRightExitDropToPlayfield(activeball)
end sub

'/////////  METAL WIRE RIGHT RAMP - EXIT HOLE - TO PLAYFIELD - EVENT  //////////
Sub RHelper3_Hit()
  if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 70
  activeball.AngMomZ = -abs(activeball.AngMomZ) * 3
  RandomSoundLeftRampDropToPlayfield()
End Sub

'
''***************************************************************
''Table MISC VP sounds
''***************************************************************

dim FaceGuideHitsSoundLevel : FaceGuideHitsSoundLevel = 0.002 * RightPlasticRampHitsSoundLevel
Sub ColFaceGuideL1_Hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_2"), FaceGuideHitsSoundLevel, activeball : End Sub
Sub F1Guide_Hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_1"), FaceGuideHitsSoundLevel, activeball : End Sub
Sub F1Guide2_Hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_1"), FaceGuideHitsSoundLevel, activeball : End Sub
Sub ColFaceGuideL2_Hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_2"), FaceGuideHitsSoundLevel, activeball : End Sub
Sub ColFaceGuideL3_Hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_3"), FaceGuideHitsSoundLevel, activeball : End Sub

sub pBlockBackhand_hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_4"), FaceGuideHitsSoundLevel, activeball : End Sub
sub pBlockBackhand001_hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_5"), FaceGuideHitsSoundLevel, activeball : End Sub

                          'volume level; range [0, 1]

' '///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
' Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim DelayedBallDropOnPlayfieldSoundLevel
DelayedBallDropOnPlayfieldSoundLevel = 0.8                  'volume level; range [0, 1]

' '/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
' '/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * MechVolume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, aVol * MechVolume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

' '/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
  RandomSoundWood()
End Sub

' '/////////////////////////////  METAL - EVENTS  ////////////////////////////
Sub Metals_Hit (idx)
  RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
  RandomSoundMetal
End Sub

' '/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
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

'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("resetdrop" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
  'PlaySoundAtLevelStatic SoundFX("droptarget" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

' '/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

'****************************************************************
' ZGII: GI
'****************************************************************

Dim gilvl   'General Illumination light state tracked for Dynamic Ball Shadows
gilvl = 1

Sub SolGI(Enabled)
  If Enabled Then
    'PlaySound"tickon" 'plays in BG sounds
    Sound_UpperGI_Relay SoundOff
    'Table1.ColorGradeImage = "ColorGrade_0"
    For each xx in GI:xx.State = 0:Next
    PFShadowsGION.visible = 0
    gilvl = 0
    FlBumperFadeTarget(1) = 0
    FlBumperFadeTarget(2) = 0
    FlBumperFadeTarget(3) = 0
    If DebugGI = True Then debug.print "GI OFF"
  Else
    For each xx in GI:xx.State = 1:Next
    'PlaySound"tickoff" 'plays in BG sounds
    Sound_UpperGI_Relay SoundOn
    'Table1.ColorGradeImage = "ColorGrade_1"
    PFShadowsGION.visible = 1
    gilvl = 1
    FlBumperFadeTarget(1) = 1
    FlBumperFadeTarget(2) = 1
    FlBumperFadeTarget(3) = 1
    If DebugGI = True Then debug.print "GI ON"
  End If
End Sub

'****************************************************************
' ZGII: GI PWM
'****************************************************************

dim GIflag: GIflag = True
Sub SolModGI(level)
  dim lvl: lvl = level

  If DebugGI = True Then debug.print "SolModGI level="&level&" lvl="&lvl
  For each xx in GI:xx.State = lvl:Next
  FlFadeBumper 1,lvl
  FlFadeBumper 2,lvl
  FlFadeBumper 3,lvl
  If lvl < 0.05 and GIflag = True Then
    GIflag = False
    Sound_UpperGI_Relay SoundOff
    'Table1.ColorGradeImage = "ColorGrade_0"
    PFShadowsGION.visible = 0
    gilvl = 0
    If DebugGI = True Then debug.print "GI OFF"
  Elseif lvl > 0.05 and GIflag = False Then
    GIflag = True
    Sound_UpperGI_Relay SoundOn
    'Table1.ColorGradeImage = "ColorGrade_1"
    PFShadowsGION.visible = 1
    gilvl = 1
    If DebugGI = True Then debug.print "GI ON"
  End If
End Sub

'******************************************************
'   ZFLD:  FLUPPER DOMES
'******************************************************
' Based on FlupperDomes2.2

' What you need in your table to use these flashers:
' Open this table and your table both in VPX
' Export all the materials domebasemat, Flashermaterial0 - 20 and import them in your table
' Export all textures (images) starting with the name "dome" and "ronddome" and import them into your table with the same names
' Export all textures (images) starting with the name "flasherbloom" and import them into your table with the same names
' Copy a set of 4 objects flasherbase, flasherlit, flasherlight and flasherflash from layer 7 to your table
' If you duplicate the four objects for a new flasher dome, be sure that they all end with the same number (in the 0-20 range)
' Copy the flasherbloom flashers from layer 10 to your table. you will need to make one per flasher dome that you plan to make
' Select the correct flasherbloom texture for each flasherbloom flasher, per flasher dome
' Copy the script below

' Place your flasher base primitive where you want the flasher located on your Table
' Then run InitFlasher in the script with the number of your flasher objects and the color of the flasher.  This will align the flasher object, light object, and
' flasher lit primitive.  It will also assign the appropriate flasher bloom images to the flasher bloom object.
'
' Example: InitFlasher 1, "green"
'
' Color Options: "blue", "green", "red", "purple", "yellow", "white", and "orange"

' You can use the RotateFlasher call to align the Rotz/ObjRotz of the flasher primitives with "handles".  Don't set those values in the editor,
' call the RotateFlasher sub instead (this call will likely crash VP if it's call for the flasher primitives without "handles")
'
' Example: RotateFlasher 1, 180    'where 1 is the flasher number and 180 is the angle of Z rotation

' For flashing the flasher use in the script: "ObjLevel(1) = 1 : FlasherFlash1_Timer"
' This should also work for flashers with variable flash levels from the rom, just use ObjLevel(1) = xx from the rom (in the range 0-1)
'
' Notes (please read!!):
' - Setting TestFlashers = 1 (below in the ScriptsDirectory) will allow you to see how the flasher objects are aligned (need the targetflasher image imported to your table)
' - The rotation of the primitives with "handles" is done with a script command, not on the primitive itself (see RotateFlasher below)
' - Color of the objects are set in the script, not on the primitive itself
' - Screws are optional to copy and position manually
' - If your table is not named "Table1" then change the name below in the script
' - Every flasher uses its own material (Flashermaterialxx), do not use it for anything else
' - Lighting > Bloom Strength affects how the flashers look, do not set it too high
' - Change RotY and RotX of flasherbase only when having a flasher something other then parallel to the playfield
' - Leave RotX of the flasherflash object to -45; this makes sure that the flash effect is visible in FS and DT
' - If you want to resize a flasher, be sure to resize flasherbase, flasherlit and flasherflash with the same percentage
' - If you think that the flasher effects are too bright, change flasherlightintensity and/or flasherflareintensity below

' Some more notes for users of the v1 flashers and/or JP's fading lights routines:
' - Delete all textures/primitives/script/materials in your table from the v1 flashers and scripts before you start; they don't mix well with v2
' - Remove flupperflash(m) routines if you have them; they do not work with this new script
' - Do not try to mix this v2 script with the JP fading light routine (that is making it too complicated), just use the example script below

' example script for rom based tables (non modulated):

' SolCallback(25)="FlashRed"
'
' Sub FlashRed(flstate)
' If Flstate Then
'   ObjTargetLevel(1) = 1
' Else
'   ObjTargetLevel(1) = 0
' End If
'   FlasherFlash1_Timer
' End Sub

' example script for rom based tables (modulated):

' SolModCallback(25)="FlashRed"
'
' Sub FlashRed(level)
' ObjTargetLevel(1) = level/255 : FlasherFlash1_Timer
' End Sub

Sub Flash15(level)
  If DebugFlashers Then Debug.print "Flash15 level=" & level
  Flash15a.State = level
  Flash15b.State = level
  Flash15c.State = level
End Sub

Sub FlashMod15(pwm)
  If DebugFlashers Then Debug.print "FlashMod15 level=" & pwm
  Flash15a.State = pwm
  Flash15b.State = pwm
  Flash15c.State = pwm
End Sub

Sub Flash25(Enabled)
  If DebugFlashers = True Then debug.print "Flash25 "& Enabled
  If Enabled Then
    ObjTargetLevel(6) = 1
    F25a.State = 1
    Sound_Flasher_Relay SoundOn, Flasherbase6
  Else
    ObjTargetLevel(6) = 0
    F25a.State = 0
    Sound_Flasher_Relay SoundOff, Flasherbase6
  End If
  Flasherflash6_Timer
End Sub

Sub FlashMod25(pwm)
  Dim FlasherRelay : FlasherRelay = 0
  If DebugFlashers = True Then debug.print "FlashMod25 "& pwm
  ModFlashFlasher 6,pwm
  F25a.State = pwm
  If pwm >= 0.99 and FlasherRelay = 0 Then
    Sound_Flasher_Relay SoundOn, Flasherbase6
    FlasherRelay = 1
  ElseIf pwm <= 0.01 and FlasherRelay = 1 Then
    Sound_Flasher_Relay SoundOff, Flasherbase6
    FlasherRelay = 0
  End If
End Sub

Sub Flash26(Enabled)
  If DebugFlashers = True Then debug.print "Flash26 "& Enabled
  If Enabled Then
    ObjTargetLevel(7) = 1
    Sound_Flasher_Relay SoundOn, Flasherbase7
  Else
    ObjTargetLevel(7) = 0
    Sound_Flasher_Relay SoundOff, Flasherbase7
  End If
  Flasherflash7_Timer
End Sub

Sub FlashMod26(pwm)
  Dim FlasherRelay : FlasherRelay = 0
  If DebugFlashers = True Then debug.print "FlashMod26 "& pwm
  ModFlashFlasher 7,pwm
  If pwm >= 0.99 and FlasherRelay = 0 Then
    Sound_Flasher_Relay SoundOn, Flasherbase7
    FlasherRelay = 1
  ElseIf pwm <= 0.01 and FlasherRelay = 1 Then
    Sound_Flasher_Relay SoundOff, Flasherbase7
    FlasherRelay = 0
  End If
End Sub

Sub Flash27(Enabled)
  If DebugFlashers = True Then debug.print "Flash27 "& Enabled
  If Enabled Then
    ObjTargetLevel(5) = 1
    Sound_Flasher_Relay SoundOn, Flasherbase5
  Else
    ObjTargetLevel(5) = 0
    Sound_Flasher_Relay SoundOff, Flasherbase5
  End If
  Flasherflash5_Timer
End Sub

Sub FlashMod27(pwm)
  Dim FlasherRelay : FlasherRelay = 0
  If DebugFlashers = True Then debug.print "FlashMod27 "& pwm
  ModFlashFlasher 5,pwm
  If pwm >= 0.99 and FlasherRelay = 0 Then
    Sound_Flasher_Relay SoundOn, Flasherbase5
    FlasherRelay = 1
  ElseIf pwm <= 0.01 and FlasherRelay = 1 Then
    Sound_Flasher_Relay SoundOff, Flasherbase5
    FlasherRelay = 0
  End If
End Sub

Sub Flash30(Enabled)
  If DebugFlashers = True Then debug.print "Flash30 "& Enabled
  If Enabled Then
    ObjTargetLevel(3) = 1
    ObjTargetLevel(4) = 1
    Sound_Flasher_Relay SoundOn, Flasherbase3
    Sound_Flasher_Relay SoundOn, Flasherbase4
  Else
    ObjTargetLevel(3) = 0
    ObjTargetLevel(4) = 0
    Sound_Flasher_Relay SoundOff, Flasherbase3
    Sound_Flasher_Relay SoundOff, Flasherbase4
  End If
  FlasherFlash3_Timer
  FlasherFlash4_Timer
End Sub

Sub FlashMod30(pwm)
  Dim FlasherRelay : FlasherRelay = 0
  If DebugFlashers = True Then debug.print "FlashMod30 "& pwm
  ModFlashFlasher 3,pwm
  ModFlashFlasher 4,pwm
  If pwm >= 0.99 and FlasherRelay = 0 Then
    Sound_Flasher_Relay SoundOn, Flasherbase3
    Sound_Flasher_Relay SoundOn, Flasherbase4
  ElseIf pwm <= 0.01 and FlasherRelay = 1 Then
    Sound_Flasher_Relay SoundOff, Flasherbase3
    Sound_Flasher_Relay SoundOff, Flasherbase4
  End If
End Sub

Sub Flash31(Enabled)
  If DebugFlashers = True Then debug.print "Flash31 "& Enabled
  If Enabled Then
    ObjTargetLevel(1) = 1
    Sound_Flasher_Relay SoundOn, Flasherbase1
  Else
    ObjTargetLevel(1) = 0
    Sound_Flasher_Relay SoundOff, Flasherbase1
  End If
  FlasherFlash1_Timer
End Sub

Sub FlashMod31(pwm)
  Dim FlasherRelay : FlasherRelay = 0
  If DebugFlashers = True Then debug.print "FlashMod31 "& pwm
  ModFlashFlasher 1,pwm
  If pwm >= 0.99 and FlasherRelay = 0 Then
    Sound_Flasher_Relay SoundOn, Flasherbase1
    FlasherRelay = 1
  ElseIf pwm <= 0.01 and FlasherRelay = 1 Then
    Sound_Flasher_Relay SoundOff, Flasherbase1
    FlasherRelay = 0
  End If
End Sub

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object      ***
Set TableRef = Table1      ' *** change this, if your table has another name           ***
FlasherLightIntensity = 0.1  ' *** lower this, if the VPX lights are too bright (i.e. 0.1)     ***
FlasherFlareIntensity = 0.1  ' *** lower this, if the flares are too bright (i.e. 0.1)       ***
FlasherBloomIntensity = 0.2  ' *** lower this, if the blooms are too bright (i.e. 0.1)       ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height

'initialise the flasher color, you can only choose from "blue", "green", "red", "purple", "yellow", "white", and "orange"
InitFlasher 1, "blue"
InitFlasher 2, "orange"
InitFlasher 3, "orange"
InitFlasher 4, "orange"
InitFlasher 5, "orange"
InitFlasher 6, "white"
InitFlasher 7, "white"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'   RotateFlasher 1,17
'   RotateFlasher 2,0
'   RotateFlasher 3,90
'   RotateFlasher 4,90

Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr)
  Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr)
  Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)

  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objbase(nr).x) / (objbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ
    objflasher(nr).height = objbase(nr).z + 40
  End If

  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0
  objlit(nr).visible = 0
  objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX
  objlit(nr).RotY = objbase(nr).RotY
  objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX
  objlit(nr).ObjRotY = objbase(nr).ObjRotY
  objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x
  objlit(nr).y = objbase(nr).y
  objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objbase(nr).roty > 135 Then
    objflasher(nr).y = objbase(nr).y + 50
    objflasher(nr).height = objbase(nr).z + 20
  Else
    objflasher(nr).y = objbase(nr).y + 20
    objflasher(nr).height = objbase(nr).z + 50
  End If
  objflasher(nr).x = objbase(nr).x

  'rothbauerw
  'Adjust the position of the light object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  objlight(nr).x = objbase(nr).x
  objlight(nr).y = objbase(nr).y
  objlight(nr).bulbhaloheight = objbase(nr).z - 10

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  Dim xthird, ythird
  xthird = tablewidth / 3
  ythird = tableheight / 3
  If objbase(nr).x >= xthird And objbase(nr).x <= xthird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenter"
    objbloom(nr).imageB = "flasherbloomCenter"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird Then
    objbloom(nr).imageA = "flasherbloomUpperLeft"
    objbloom(nr).imageB = "flasherbloomUpperLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird Then
    objbloom(nr).imageA = "flasherbloomUpperRight"
    objbloom(nr).imageB = "flasherbloomUpperRight"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenterLeft"
    objbloom(nr).imageB = "flasherbloomCenterLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenterRight"
    objbloom(nr).imageB = "flasherbloomCenterRight"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 3 Then
    objbloom(nr).imageA = "flasherbloomLowerLeft"
    objbloom(nr).imageB = "flasherbloomLowerLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 3 Then
    objbloom(nr).imageA = "flasherbloomLowerRight"
    objbloom(nr).imageB = "flasherbloomLowerRight"
  End If

  ' set the texture and color of all objects
  Select Case objbase(nr).image
    Case "dome2basewhite"
      objbase(nr).image = "dome2base" & col
      objlit(nr).image = "dome2lit" & col

    Case "ronddomebasewhite"
      objbase(nr).image = "ronddomebase" & col
      objlit(nr).image = "ronddomelit" & col

    Case "domeearbasewhite"
      objbase(nr).image = "domeearbase" & col
      objlit(nr).image = "domeearlit" & col
  End Select
  If TestFlashers = 0 Then
    objflasher(nr).imageA = "domeflashwhite"
    objflasher(nr).visible = 0
  End If
  Select Case col
    Case "blue"
      objlight(nr).color = RGB(4,120,255)
      objflasher(nr).color = RGB(200,255,255)
      objbloom(nr).color = RGB(4,120,255)
      objlight(nr).intensity = 5000

    Case "green"
      objlight(nr).color = RGB(12,255,4)
      objflasher(nr).color = RGB(12,255,4)
      objbloom(nr).color = RGB(12,255,4)

    Case "red"
      objlight(nr).color = RGB(255,32,4)
      objflasher(nr).color = RGB(255,32,4)
      objbloom(nr).color = RGB(255,32,4)

    Case "purple"
      objlight(nr).color = RGB(230,49,255)
      objflasher(nr).color = RGB(255,64,255)
      objbloom(nr).color = RGB(230,49,255)

    Case "yellow"
      objlight(nr).color = RGB(200,173,25)
      objflasher(nr).color = RGB(255,200,50)
      objbloom(nr).color = RGB(200,173,25)

    Case "white"
      objlight(nr).color = RGB(255,240,150)
      objflasher(nr).color = RGB(100,86,59)
      objbloom(nr).color = RGB(255,240,150)

    Case "orange"
      objlight(nr).color = RGB(255,70,0)
      objflasher(nr).color = RGB(255,70,0)
      objbloom(nr).color = RGB(255,70,0)
  End Select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT And ObjFlasher(nr).RotX =  - 45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle)
  angle = ((angle + 360 - objbase(nr).ObjRotZ) Mod 180) / 30
  objbase(nr).showframe(angle)
  objlit(nr).showframe(angle)
End Sub

Sub FlashFlasher(nr)
  If Not objflasher(nr).TimerEnabled Then
    objflasher(nr).TimerEnabled = True
    objflasher(nr).visible = 1
    objbloom(nr).visible = 1
    objlit(nr).visible = 1
  End If
  objflasher(nr).opacity = 1000 * FlasherFlareIntensity * ObjLevel(nr) ^ 2.5
  objbloom(nr).opacity = 100 * FlasherBloomIntensity * ObjLevel(nr) ^ 2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr) ^ 3
  objbase(nr).BlendDisableLighting = FlasherOffBrightness + 10 * ObjLevel(nr) ^ 3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr) ^ 2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  If Round(ObjTargetLevel(nr),1) > Round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) + 0.3
    If ObjLevel(nr) > 1 Then ObjLevel(nr) = 1
  ElseIf Round(ObjTargetLevel(nr),1) < Round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) * 0.85 - 0.01
    If ObjLevel(nr) < 0 Then ObjLevel(nr) = 0
  Else
    ObjLevel(nr) = Round(ObjTargetLevel(nr),1)
    objflasher(nr).TimerEnabled = False
  End If
  '   ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then
    objflasher(nr).TimerEnabled = False
    objflasher(nr).visible = 0
    objbloom(nr).visible = 0
    objlit(nr).visible = 0
  End If
End Sub

Sub ModFlashFlasher(nr, aValue)
    if aValue > 0 then
        objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1
    else
        objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0
    end if
    objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * aValue
    objbloom(nr).opacity = 100 *  FlasherBloomIntensity * aValue
    objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * aValue
    objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * aValue
    objlit(nr).BlendDisableLighting = 10 * aValue
    UpdateMaterial "Flashermaterial" & nr, 0,0,0,0,0,0,Round(aValue,1),RGB(255,255,255),0,0,False,True,0,0,0,0
End Sub

Sub FlasherFlash1_Timer()
  FlashFlasher(1)
End Sub
Sub FlasherFlash2_Timer()
  FlashFlasher(2)
End Sub
Sub FlasherFlash3_Timer()
  FlashFlasher(3)
End Sub
Sub FlasherFlash4_Timer()
  FlashFlasher(4)
End Sub
Sub FlasherFlash5_Timer()
  FlashFlasher(5)
End Sub
Sub FlasherFlash6_Timer()
  FlashFlasher(6)
End Sub
Sub FlasherFlash7_Timer()
  FlashFlasher(7)
End Sub

'******************************************************
'******  END FLUPPER DOMES
'******************************************************

'******************************************************
'   ZFLF:  FLUPPER FLASHERLAMPS
'******************************************************
' What you need in your table to use these flashers:
' Open this table and your table both in VPX
' Export all the materials flasherlampbasemat, Flasherlampmaterial0 - 20, Flasherbulbmaterial0 - 20 and import them in your table
' Export all textures (images) starting with the name "flasherbulb" and import them into your table with the same
' Export all textures (images) starting with the name "flasherlamp" and import them into your table with the same names
' Export all textures (images) starting with the name "flasherbloom" and import them into your table with the same names (only needed if you don't have Flupper Domes)
' Copy a set of 7 objects flasherlampbase, flasherlamplit, flasherbulbbase, flasherbulblit, flasherlampfilament, flasherlamplight and flasherlampflash from layer 7 to your table
' If you duplicate the four objects for a new flasherlamp, be sure that they all end with the same number (in the 0-20 range)
' Copy the flasherlampbloom flashers from layer 10 to your table. you will need to make one per flasher lamp that you plan to make
' Select the correct flasherbloom texture for each flasherlampbloom flasher, per flasher lamp
' Copy the script below

' Place your flasher base primitive where you want the flasher located on your Table
' Then run InitFlasherLamp in the script with the number of your flasher objects and the color of the flasher.  This will align the flasher object, light object, and
' flasher lit primitive.  It will also assign the appropriate flasher bloom images to the flasher bloom object.
'
' Example: InitLampFlasher 1
'
' You can use the RotateFlasherLamp call to align the Rotz/ObjRotz of the flasher primitives with "handles".  Don't set those values in the editor,
' call the RotateFlasherLamp sub instead (this call will likely crash VP if it's call for the flasher primitives without "handles")
'
' Example: RotateFlasherLamp 1, 180    'where 1 is the flasher number and 180 is the angle of Z rotation

' For flashing the flasher use in the script: "ObjLampTargetLevel(1) = 1 : FlasherFlashLamp1_Timer"
' This should also work for flashers with variable flash levels from the rom, just use ObjLampLevel(1) = xx from the rom (in the range 0-1)
'
' Notes (please read!!):
' - Setting TestFlashers = 1 (below in the ScriptsDirectory) will allow you to see how the flasher objects are aligned (need the targetflasher image imported to your table)
' - The rotation of the primitives with "handles" is done with a script command, not on the primitive itself (see RotateFlasherLamp below)
' - If your table is not named "Table1" then change the name below in the script
' - Every flasher uses its own material (Flasherlampmaterialxx), do not use it for anything else
' - Lighting > Bloom Strength affects how the flashers look, do not set it too high
' - Change RotY and RotX of flasherlampbase only when having a flasher something other then parallel to the playfield
' - Leave RotX of the flasherlampflash object to -45; this makes sure that the flash effect is visible in FS and DT
' - If you want to resize a flasher, be sure to resize flasherbase, flasherlamplit and flasherlampflash with the same percentage
' - If you think that the flasher effects are too bright, change flasherlamplightintensity and/or flasherlampflareintensity below

' Some more notes for users of the v1 flashers and/or JP's fading lights routines:
' - Delete all textures/primitives/script/materials in your table from the v1 flashers and scripts before you start; they don't mix well with v2
' - Remove flupperflash(m) routines if you have them; they do not work with this new script
' - Do not try to mix this v2 script with the JP fading light routine (that is making it too complicated), just use the example script below

' example script for rom based tables (non modulated):

' SolCallback(25)="FlashLamp"
'
' Sub FlashLamp(flstate)
' If Flstate Then
'   ObjLampTargetLevel(1) = 1
' Else
'   ObjLampTargetLevel(1) = 0
' End If
'   Flasherlampflash1_Timer
' End Sub

' example script for rom based tables (modulated):

' SolModCallback(25)="FlashLamp"
'
' Sub FlashLamp(level)
' ObjLampTargetLevel(1) = level/255 : Flasherlampflash1_Timer
' End Sub

Sub Flash28(Enabled)
  If DebugFlashers Then Debug.print "Flash28 " & level
  If Enabled Then
    ObjLampTargetLevel(1) = 1
    ObjLampTargetLevel(2) = 1
    Sound_Flasher_Relay SoundOn, Flasherlampbase1
    Sound_Flasher_Relay SoundOn, Flasherlampbase2
  Else
    ObjLampTargetLevel(1) = 0
    ObjLampTargetLevel(2) = 0
    Sound_Flasher_Relay SoundOff, Flasherlampbase1
    Sound_Flasher_Relay SoundOff, Flasherlampbase2
  End If
  Flasherlampflash1_Timer
  Flasherlampflash2_Timer
End Sub

Sub FlashMod28(pwm)
  Dim FlasherRelay : FlasherRelay = 0
  If DebugFlashers = True Then debug.print "FlashMod28 "& pwm
  ModFlashLampFlasher 1,pwm
  ModFlashLampFlasher 2,pwm
  If pwm >= 0.99 and FlasherRelay = 0 Then
    Sound_Flasher_Relay SoundOn, Flasherlampbase1
    Sound_Flasher_Relay SoundOn, Flasherlampbase2
    FlasherRelay = 1
  ElseIf pwm <= 0.01 and FlasherRelay = 1 Then
    Sound_Flasher_Relay SoundOff, Flasherlampbase1
    Sound_Flasher_Relay SoundOff, Flasherlampbase2
    FlasherRelay = 0
  End If
End Sub

Sub Flash29(Enabled)
  If DebugFlashers = True Then debug.print "Flash29 "& Enabled
  If Enabled Then
    ObjTargetLevel(2) = 1
    ObjLampTargetLevel(3) = 1
    Sound_Flasher_Relay SoundOn, Flasherbase2
    Sound_Flasher_Relay SoundOn, Flasherlampbase3
  Else
    ObjTargetLevel(2) = 0
    ObjLampTargetLevel(3) = 0
    Sound_Flasher_Relay SoundOff, Flasherbase2
    Sound_Flasher_Relay SoundOff, Flasherlampbase3
  End If
  FlasherFlash2_Timer
  FlasherLampFlash3_Timer
End Sub

Sub FlashMod29(pwm)
  Dim FlasherRelay : FlasherRelay = 0
  If DebugFlashers = True Then debug.print "FlashMod29 "& pwm
  ModFlashFlasher 2,pwm
  ModFlashLampFlasher 3,pwm
  If pwm >= 0.99 and FlasherRelay = 0 Then
    Sound_Flasher_Relay SoundOn, Flasherbase2
    Sound_Flasher_Relay SoundOn, Flasherlampbase3
    FlasherRelay = 1
  ElseIf pwm <= 0.01 and FlasherRelay = 1 Then
    Sound_Flasher_Relay SoundOff, Flasherbase2
    Sound_Flasher_Relay SoundOff, Flasherlampbase3
    FlasherRelay = 0
  End If
End Sub

Sub Flash32(Enabled)
  If DebugFlashers Then Debug.print "Flash32 " & Enabled
  If Enabled Then
    ObjLampTargetLevel(4) = 1
    Sound_Flasher_Relay SoundOn, Flasherlampbase4
  Else
    ObjLampTargetLevel(4) = 0
    Sound_Flasher_Relay SoundOff, Flasherlampbase4
  End If
  FlasherFlash7_Timer
End Sub

Sub FlashMod32(pwm)
  Dim FlasherRelay : FlasherRelay = 0
  If DebugFlashers Then Debug.print "FlashMod32 level=" & pwm
  ModFlashLampFlasher 4,pwm
  If pwm >= 0.99 and FlasherRelay = 0 Then
    Sound_Flasher_Relay SoundOn, Flasherlampbase4
    FlasherRelay = 1
  ElseIf pwm <= 0.01 and FlasherRelay = 1 Then
    Sound_Flasher_Relay SoundOff, Flasherlampbase4
    FlasherRelay = 0
  End If
End Sub

Dim TestLampFlashers, TableLampRef, FlasherLampLightIntensity, FlasherLampFlareIntensity, FlasherLampBloomIntensity, FlasherLampOffBrightness, FlasherBulbOffBrightness
' *********************************************************************
TestLampFlashers = 0        ' *** set this to 1 to check position of flasher object      ***
Set TableLampRef = Table1      ' *** change this, if your table has another name           ***
FlasherLampLightIntensity = 0.05   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)     ***
FlasherLampFlareIntensity = 0.05   ' *** lower this, if the flares are too bright (i.e. 0.1)       ***
FlasherLampBloomIntensity = 0.2  ' *** lower this, if the blooms are too bright (i.e. 0.1)       ***
FlasherLampOffBrightness = 0.5    ' *** brightness of the flasher lamp when switched off (range 0-2)  ***
FlasherBulbOffBrightness = 0.5    ' *** brightness of the flasher bulb when switched off (range 0-2)  ***
' *********************************************************************

Dim ObjLampLevel(20), objlampbase(20), objlamplit(20), objbulbbase(20), objbulblit(20), objlampflasher(20), objlampbloom(20), objlamplight(20), ObjLampTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height

InitLampFlasher 1
InitLampFlasher 2
InitLampFlasher 3
InitLampFlasher 4

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'   RotateLampFlasher 1,17
'   RotateLampFlasher 2,0
'   RotateLampFlasher 3,90
'   RotateLampFlasher 4,90

Sub InitLampFlasher(nr)
  ' store all objects in an array for use in FlashLampFlasher subroutine
  Set objlampbase(nr) = Eval("Flasherlampbase" & nr)
  Set objlamplit(nr) = Eval("Flasherlamplit" & nr)
  Set objbulbbase(nr) = Eval("Flasherbulbbase" & nr)
  Set objbulblit(nr) = Eval("Flasherbulblit" & nr)
  Set objlampflasher(nr) = Eval("Flasherlampflash" & nr)
  Set objlamplight(nr) = Eval("Flasherlamplight" & nr)
  Set objlampbloom(nr) = Eval("Flasherlampbloom" & nr)

  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objlampbase(nr).RotY = 0 Then
    objlampbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objlampbase(nr).x) / (objlampbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
  End If

  If objbulbbase(nr).RotY = 0 Then
    objbulbbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objbulbbase(nr).x) / (objbulbbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
    objlampflasher(nr).RotZ = objbulbbase(nr).ObjRotZ
    objlampflasher(nr).height = objbulbbase(nr).z + 40
  End If

  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlamplight(nr).IntensityScale = 0
  objlamplit(nr).visible = 0
  'objlamplit(nr).material = "Flasherlampmaterial" & nr
  objlamplit(nr).RotX = objlampbase(nr).RotX
  objlamplit(nr).RotY = objlampbase(nr).RotY
  objlamplit(nr).RotZ = objlampbase(nr).RotZ
  objlamplit(nr).ObjRotX = objlampbase(nr).ObjRotX
  objlamplit(nr).ObjRotY = objlampbase(nr).ObjRotY
  objlamplit(nr).ObjRotZ = objlampbase(nr).ObjRotZ
  objlamplit(nr).x = objlampbase(nr).x
  objlamplit(nr).y = objlampbase(nr).y
  objlamplit(nr).z = objlampbase(nr).z
  objlampbase(nr).BlendDisableLighting = FlasherLampOffBrightness

  objbulblit(nr).visible = 0
  objbulblit(nr).material = "Flasherbulbmaterial" & nr
  objbulblit(nr).RotX = objbulbbase(nr).RotX
  objbulblit(nr).RotY = objbulbbase(nr).RotY
  objbulblit(nr).RotZ = objbulbbase(nr).RotZ
  objbulblit(nr).ObjRotX = objbulbbase(nr).ObjRotX
  objbulblit(nr).ObjRotY = objbulbbase(nr).ObjRotY
  objbulblit(nr).ObjRotZ = objbulbbase(nr).ObjRotZ
  objbulblit(nr).x = objbulbbase(nr).x
  objbulblit(nr).y = objbulbbase(nr).y
  objbulblit(nr).z = objbulbbase(nr).z
  objbulbbase(nr).BlendDisableLighting = FlasherBulbOffBrightness

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objlampbase(nr).roty > 135 Then
    objlampflasher(nr).y = objlampbase(nr).y + 50
    objlampflasher(nr).height = objlampbase(nr).z + 20
  Else
    objlampflasher(nr).y = objlampbase(nr).y + 20
    objlampflasher(nr).height = objlampbase(nr).z + 50
  End If
  objlampflasher(nr).x = objlampbase(nr).x

  'rothbauerw
  'Adjust the position of the light object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  objlamplight(nr).x = objlampbase(nr).x
  objlamplight(nr).y = objlampbase(nr).y
  objlamplight(nr).bulbhaloheight = objlampbase(nr).z - 10

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  Dim xthird, ythird
  xthird = tablewidth / 3
  ythird = tableheight / 3
  If objlampbase(nr).x >= xthird And objlampbase(nr).x <= xthird * 2 Then
    objlampbloom(nr).imageA = "flasherbloomCenter"
    objlampbloom(nr).imageB = "flasherbloomCenter"
  ElseIf objlampbase(nr).x < xthird And objlampbase(nr).y < ythird Then
    objlampbloom(nr).imageA = "flasherbloomUpperLeft"
    objlampbloom(nr).imageB = "flasherbloomUpperLeft"
  ElseIf  objlampbase(nr).x > xthird * 2 And objlampbase(nr).y < ythird Then
    objlampbloom(nr).imageA = "flasherbloomUpperRight"
    objlampbloom(nr).imageB = "flasherbloomUpperRight"
  ElseIf objlampbase(nr).x < xthird And objlampbase(nr).y < ythird * 2 Then
    objlampbloom(nr).imageA = "flasherbloomCenterLeft"
    objlampbloom(nr).imageB = "flasherbloomCenterLeft"
  ElseIf  objlampbase(nr).x > xthird * 2 And objlampbase(nr).y < ythird * 2 Then
    objlampbloom(nr).imageA = "flasherbloomCenterRight"
    objlampbloom(nr).imageB = "flasherbloomCenterRight"
  ElseIf objlampbase(nr).x < xthird And objlampbase(nr).y < ythird * 3 Then
    objlampbloom(nr).imageA = "flasherbloomLowerLeft"
    objlampbloom(nr).imageB = "flasherbloomLowerLeft"
  ElseIf  objlampbase(nr).x > xthird * 2 And objlampbase(nr).y < ythird * 3 Then
    objlampbloom(nr).imageA = "flasherbloomLowerRight"
    objlampbloom(nr).imageB = "flasherbloomLowerRight"
  End If

  ' set the texture and color of all objects
  'objlampbase(nr).image = "lamp2base"
  'objlamplit(nr).image = "lamp2lit"
  objbulbbase(nr).image = "flasherbulb"
  objbulblit(nr).image = "flasherbulblit"
  If TestLampFlashers = 0 Then
    objlampflasher(nr).imageA = "lampflashwhite"
    objlampflasher(nr).visible = 0
  End If

  objlamplight(nr).color = RGB(255,240,150)
  objlampflasher(nr).color = RGB(100,86,59)
  objlampbloom(nr).color = RGB(255,240,150)

  objlamplight(nr).colorfull = objlamplight(nr).color
  If TableRef.ShowDT And ObjFlasher(nr).RotX =  - 45 Then
    objlampflasher(nr).height = objlampflasher(nr).height - 20 * objlampflasher(nr).y / tableheight
    objlampflasher(nr).y = objlampflasher(nr).y + 10
  End If
End Sub

Sub RotateLampFlasher(nr, angle)
  angle = ((angle + 360 - objlampbase(nr).ObjRotZ) Mod 180) / 30
  objlampbase(nr).showframe(angle)
  objlamplit(nr).showframe(angle)
  objbulbbase(nr).showframe(angle)
  objbulblit(nr).showframe(angle)
End Sub

Sub FlashLampFlasher(nr)
  If Not objlampflasher(nr).TimerEnabled Then
    objlampflasher(nr).TimerEnabled = True
    objlampflasher(nr).visible = 1
    objlampbloom(nr).visible = 1
    objlamplit(nr).visible = 1
    objbulblit(nr).visible = 1
  End If
  objlampflasher(nr).opacity = 1000 * FlasherLampFlareIntensity * ObjLampLevel(nr) ^ 2.5
  objlampbloom(nr).opacity = 100 * FlasherLampBloomIntensity * ObjLampLevel(nr) ^ 2.5
  objlamplight(nr).IntensityScale = 0.5 * FlasherLampLightIntensity * ObjLampLevel(nr) ^ 3
  objlampbase(nr).BlendDisableLighting = FlasherLampOffBrightness + 10 * ObjLampLevel(nr) ^ 3
  objlamplit(nr).BlendDisableLighting = 10 * ObjLampLevel(nr) ^ 2
  objbulbbase(nr).BlendDisableLighting = FlasherLampOffBrightness + 10 * ObjLampLevel(nr) ^ 3
  objbulblit(nr).BlendDisableLighting = 10 * ObjLampLevel(nr) ^ 2
  'UpdateMaterial "Flasherlampmaterial" & nr,0,0,0,0,0,0,ObjLampLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  objbulbbase(nr).BlendDisableLighting = FlasherLampOffBrightness + 10 * ObjLampLevel(nr) ^ 3
  objbulblit(nr).BlendDisableLighting = 10 * ObjLampLevel(nr) ^ 2
  UpdateMaterial "Flasherbulbmaterial" & nr,0,0,0,0,0,0,ObjLampLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  If Round(ObjLampTargetLevel(nr),1) > Round(ObjLampLevel(nr),1) Then
    ObjLampLevel(nr) = ObjLampLevel(nr) + 0.3
    If ObjLampLevel(nr) > 1 Then ObjLampLevel(nr) = 1
  ElseIf Round(ObjLampTargetLevel(nr),1) < Round(ObjLampLevel(nr),1) Then
    ObjLampLevel(nr) = ObjLampLevel(nr) * 0.85 - 0.01
    If ObjLampLevel(nr) < 0 Then ObjLampLevel(nr) = 0
  Else
    ObjLampLevel(nr) = Round(ObjLampTargetLevel(nr),1)
    objlampflasher(nr).TimerEnabled = False
  End If
  '   ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLampLevel(nr) < 0 Then
    objlampflasher(nr).TimerEnabled = False
    objlampflasher(nr).visible = 0
    objlampbloom(nr).visible = 0
    objlamplit(nr).visible = 0
    objbulblit(nr).visible = 0
  End If
End Sub

Sub ModFlashLampFlasher(nr, aValue)
    if aValue > 0 then
        objlampflasher(nr).visible = 1 : objlampbloom(nr).visible = 1 : objlamplit(nr).visible = 1 : objbulblit(nr).visible = 1
    else
        objlampflasher(nr).visible = 0 : objlampbloom(nr).visible = 0 : objlamplit(nr).visible = 0 : objbulblit(nr).visible = 0
    end if
    objlampflasher(nr).opacity = 1000 *  FlasherLampFlareIntensity * aValue
    objlampbloom(nr).opacity = 100 *  FlasherLampBloomIntensity * aValue
    objlamplight(nr).IntensityScale = 0.5 * FlasherLampLightIntensity * aValue
    objlampbase(nr).BlendDisableLighting =  FlasherLampOffBrightness + 10 * aValue
    objlamplit(nr).BlendDisableLighting = 10 * aValue
    'UpdateMaterial "Flasherlampmaterial" & nr, 0,0,0,0,0,0,Round(aValue,1),RGB(255,255,255),0,0,False,True,0,0,0,0
  objbulbbase(nr).BlendDisableLighting =  FlasherLampOffBrightness + 10 * aValue
    objbulblit(nr).BlendDisableLighting = 10 * aValue
    UpdateMaterial "Flasherbulbmaterial" & nr, 0,0,0,0,0,0,Round(aValue,1),RGB(255,255,255),0,0,False,True,0,0,0,0
End Sub

Sub FlasherLampFlash1_Timer()
  FlashLampFlasher(1)
End Sub
Sub FlasherLampFlash2_Timer()
  FlashLampFlasher(2)
End Sub
Sub FlasherLampFlash3_Timer()
  FlashLampFlasher(3)
End Sub
Sub FlasherLampFlash4_Timer()
  FlashLampFlasher(4)
End Sub

'******************************************************
'   ZFLB:  FLUPPER BUMPERS
'******************************************************
' Based on FlupperBumpers 0.145 final

' Explanation of how these bumpers work:
' There are 10 elements involved per bumper:
' - the shadow of the bumper ( a vpx flasher object)
' - the bumper skirt (primitive)
' - the bumperbase (primitive)
' - a vpx light which colors everything you can see through the bumpertop
' - the bulb (primitive)
' - another vpx light which lights up everything around the bumper
' - the bumpertop (primitive)
' - the VPX bumper object
' - the bumper screws (primitive)
' - the bulb highlight VPX flasher object
' All elements have a special name with the number of the bumper at the end, this is necessary for the fading routine and the initialisation.
' For the bulb and the bumpertop there is a unique material as well per bumpertop.
' To use these bumpers you have to first copy all 10 elements to your table.
' Also export the textures (images) with names that start with "Flbumper" and "Flhighlight" and materials with names that start with "bumper".
' Make sure that all the ten objects are aligned on center, if possible with the exact same x,y coordinates
' After that copy the script (below); also copy the BumperTimer vpx object to your table
' Every bumper needs to be initialised with the FlInitBumper command, see example below;
' Colors available are red, white, blue, orange, yellow, green, purple and blacklight.
' In a GI subroutine you can then call set the bumperlight intensity with the "FlBumperFadeTarget(nr) = value" command
' where nr is the number of the bumper, value is between 0 (off) and 1 (full on) (so you can also use 0.3 0.4 etc).

' Notes:
' - There is only one color for the disk; you can photoshop it to a different color
' - The bumpertops are angle independent up to a degree; my estimate is -45 to + 45 degrees horizontally, 0 (topview) to 70-80 degrees (frontview)
' - I built in correction for the day-night slider; this might not work perfectly, depending on your table lighting
' - These elements, textures and materials do NOT integrate with any of the lighting routines I have seen in use in many VPX tables
'   (just find the GI handling routine and insert the FlBumperFadeTarget statement)
' - If you want to use VPX native bumperdisks just copy my bumperdisk but make it invisible

' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
  DNA30 = 0
  DNA45 = (NightDay - 10) / 20
  DNA90 = 0
  DayNightAdjust = 0.4
Else
  DNA30 = (NightDay - 10) / 30
  DNA45 = (NightDay - 10) / 45
  DNA90 = (NightDay - 10) / 90
  DayNightAdjust = NightDay / 25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt
For cnt = 1 To 6
  FlBumperActive(cnt) = False
Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight
FlInitBumper 1, "red"
FlInitBumper 2, "red"
FlInitBumper 3, "red"

' ### uncomment the statement below to change the color for all bumpers ###
'Dim ind
'For ind = 1 To 3
' FlInitBumper ind, "red"
'Next

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True

  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1
  FlBumperFadeTarget(nr) = 1
  FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr)
  FlBumperTop(nr).material = "bumpertopmat" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr)
  Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr)
  Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr)
  FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr)
  FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)

  ' set the color for the two VPX lights
  Select Case col
    Case "red"
      FlBumperSmallLight(nr).color = RGB(255,4,0)
      FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
      FlBumperBigLight(nr).color = RGB(255,32,0)
      FlBumperBigLight(nr).colorfull = RGB(255,32,0)
      FlBumperHighlight(nr).color = RGB(255,64,0)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
      FlBumperSmallLight(nr).TransmissionScale = 0

    Case "blue"
      FlBumperBigLight(nr).color = RGB(32,80,255)
      FlBumperBigLight(nr).colorfull = RGB(32,80,255)
      FlBumperSmallLight(nr).color = RGB(0,80,255)
      FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
      FlBumperSmallLight(nr).TransmissionScale = 0
      MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
      FlBumperHighlight(nr).color = RGB(255,16,8)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "green"
      FlBumperSmallLight(nr).color = RGB(8,255,8)
      FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
      FlBumperBigLight(nr).color = RGB(32,255,32)
      FlBumperBigLight(nr).colorfull = RGB(32,255,32)
      FlBumperHighlight(nr).color = RGB(255,32,255)
      MaterialColor "bumpertopmat" & nr, RGB(16,255,16)
      FlBumperSmallLight(nr).TransmissionScale = 0.005
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "orange"
      FlBumperHighlight(nr).color = RGB(255,130,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).color = RGB(255,130,0)
      FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
      FlBumperBigLight(nr).color = RGB(255,190,8)
      FlBumperBigLight(nr).colorfull = RGB(255,190,8)

    Case "white"
      FlBumperBigLight(nr).color = RGB(255,230,190)
      FlBumperBigLight(nr).colorfull = RGB(255,230,190)
      FlBumperHighlight(nr).color = RGB(255,180,100)
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99

    Case "blacklight"
      FlBumperBigLight(nr).color = RGB(32,32,255)
      FlBumperBigLight(nr).colorfull = RGB(32,32,255)
      FlBumperHighlight(nr).color = RGB(48,8,255)
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "yellow"
      FlBumperSmallLight(nr).color = RGB(255,230,4)
      FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
      FlBumperBigLight(nr).color = RGB(255,240,50)
      FlBumperBigLight(nr).colorfull = RGB(255,240,50)
      FlBumperHighlight(nr).color = RGB(255,255,220)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0

    Case "purple"
      FlBumperBigLight(nr).color = RGB(80,32,255)
      FlBumperBigLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).color = RGB(80,32,255)
      FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperHighlight(nr).color = RGB(32,64,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
  End Select
End Sub

Sub FlFadeBumper(nr, Z)
  FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
  '   UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
  '        OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
  '        float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
  FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 ) * DayNightAdjust

  Select Case FlBumperColor(nr)
    Case "blue"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(38 - 24 * Z,130 - 98 * Z,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 500 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z + 0.97 * Z ^ 3)
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 10000 * (Z ^ 3) / (0.5 + DNA90)

    Case "green"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(16 + 16 * Sin(Z * 3.14),255,16 + 16 * Sin(Z * 3.14)), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
      FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 6000 * (Z ^ 3) / (1 + DNA90)

    Case "red"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(128, 16 - 11 * Z + 16 * Sin(Z * 3.14),0), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30 ^ 2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 30 * DayNightAdjust + 9000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 5 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z * 4,8 - Z * 8)

    Case "orange"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 100 - 22 * z + 16 * Sin(Z * 3.14),Z * 32), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30 ^ 2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,100 + Z * 50, 0)

    Case "white"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
      FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 8 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255,255 - 20 * Z,255 - 65 * Z)
      FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20 * Z,255 - 65 * Z)
      MaterialColor "bumpertopmat" & nr, RGB(255,235 - z * 36,220 - Z * 90)

    Case "blacklight"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 1, RGB(30 - 27 * Z ^ 0.03,30 - 28 * Z ^ 0.01, 255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z ^ 3
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255 - 240 * (Z ^ 0.1),255 - 240 * (Z ^ 0.1),255)
      FlBumperSmallLight(nr).colorfull = RGB(255 - 200 * z,255 - 200 * Z,255)
      MaterialColor "bumpertopmat" & nr, RGB(255 - 190 * Z,235 - z * 180,220 + 35 * Z)

    Case "yellow"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 180 + 40 * z, 48 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30 ^ 2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)

    Case "purple"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(128 - 118 * Z - 32 * Sin(Z * 3.14), 32 - 26 * Z ,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 15 + 200 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z + 0.97 * Z ^ 3)
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (0.5 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(128 - 60 * Z,32,255)
  End Select
End Sub

Sub BumperTimer_Timer
  Dim nr
  For nr = 1 To 6
    If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
      If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
    If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
      If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
  Next
End Sub

'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************

'******************************************************
'****  LAMPZ by nFozzy
'******************************************************


Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz
If UseVPMModSol = 0 Then
  Set Lampz = New LampFader
Else
  Set Lampz = New VPMLampUpdater  'PWM inserts
End If

InitLampsNF              ' Setup lamp assignments

LampTimer.Interval = -1   ' Using fixed value so the fading speed is same for every fps
LampTimer.Enabled = 1

Sub LampTimer_Timer()
  dim x, chglamp
  chglamp = Controller.ChangedLamps
  If Not IsEmpty(chglamp) Then
    For x = 0 To UBound(chglamp)      'nmbr = chglamp(x, 0), state = chglamp(x, 1)
      If UseVPMModSol = 0 Then
        Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
      Else
        Lampz.state(chglamp(x, 0)) = chglamp(x, 1)/255.0/0.7   'PWM inserts (note the /0.7 is because the ROM does not command full brightness usually. This is likely table specific)
      End If
      'if chglamp(x, 0)=18 then debug.print "L18.state = "&Lampz.state(chglamp(x, 0))  'used for debugging inserts
      'if chglamp(x, 0)=27 then debug.print "L27.state = "&Lampz.state(chglamp(x, 0))  'used for debugging inserts
    next
  End If
  If UseVPMModSol = 0 Then Lampz.Update2  'update (fading logic only)
End Sub


Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  If UseVPMModSol = 0 Then
    If Lampz.UseFunction Then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  End If
  pri.blenddisablelighting = aLvl * DLintensity
End Sub

Sub SetModLamp(id, val)
  Lampz.state(id) = val
  if id=118 then debug.print "SetModLamp " & val
End Sub


Sub InitLampsNF()
  Dim x, xx

  'Filtering (comment out to disable)
  If UseVPMModSol = 0 Then
    Lampz.Filter = "LampFilter" 'Puts all lamp intensityscale output (no callbacks) through this function before updating

    'Adjust fading speeds (max level / full MS fading time). The Modulate property must be set to 1 / max level if lamp is modulated.
    for x = 0 to 150 : Lampz.FadeSpeedUp(x) = 1/40 : Lampz.FadeSpeedDown(x) = 1/120 : next

  End If

  'Lampz Assignments
  '  In a ROM based table, the lamp ID is used to set the state of the Lampz objects

  'MassAssign is an optional way to do assignments. It'll create arrays automatically / append objects to existing arrays

  Lampz.MassAssign(1)= L1
  Lampz.MassAssign(1)= L1a
  Lampz.Callback(1) = "DisableLighting p1, 200,"

  Lampz.MassAssign(2)= L2
  Lampz.MassAssign(2)= L2a
  Lampz.Callback(2) = "DisableLighting p2, 200,"

  Lampz.MassAssign(3)= L3
  Lampz.MassAssign(3)= L3a
  Lampz.Callback(3) = "DisableLighting p3, 200,"

  Lampz.MassAssign(4)= L4
  Lampz.MassAssign(4)= L4a
  Lampz.Callback(4) = "DisableLighting p4, 200,"

  Lampz.MassAssign(5)= L5
  Lampz.MassAssign(5)= L5a
  Lampz.Callback(5) = "DisableLighting p5, 200,"

  Lampz.MassAssign(6)= L6
  Lampz.MassAssign(6)= L6a
  Lampz.Callback(6) = "DisableLighting p6, 200,"

  Lampz.MassAssign(7)= L7
  Lampz.MassAssign(7)= L7a
  Lampz.Callback(7) = "DisableLighting p7, 200,"

  Lampz.MassAssign(8)= L8
  Lampz.MassAssign(8)= L8a
  Lampz.Callback(8) = "DisableLighting p8, 200,"

  Lampz.MassAssign(9)= L9
  Lampz.MassAssign(9)= L9a
  Lampz.Callback(9) = "DisableLighting p9, 200,"

  Lampz.MassAssign(10)= L10
  Lampz.MassAssign(10)= L10a
  Lampz.Callback(10) = "DisableLighting p10, 200,"

  Lampz.MassAssign(11)= L11
  Lampz.MassAssign(11)= L11a
  Lampz.Callback(11) = "DisableLighting p11, 200,"

  Lampz.MassAssign(12)= L12
  Lampz.MassAssign(12)= L12a
  Lampz.Callback(12) = "DisableLighting p12, 200,"

  Lampz.MassAssign(13)= L13
  Lampz.MassAssign(13)= L13a
  Lampz.Callback(13) = "DisableLighting p13, 200,"

  Lampz.MassAssign(14)= L14
  Lampz.MassAssign(14)= L14a
  Lampz.MassAssign(14)= L14b
  Lampz.MassAssign(14)= L14c
  Lampz.Callback(14) = "DisableLighting p14, 200,"

  Lampz.MassAssign(15)= L15
  Lampz.MassAssign(15)= L15a
  Lampz.MassAssign(15)= L15b
  Lampz.MassAssign(15)= L15c
  Lampz.Callback(15) = "DisableLighting p15, 200,"

  Lampz.MassAssign(16)= L16
  Lampz.MassAssign(16)= L16a
  Lampz.MassAssign(16)= L16b
  Lampz.MassAssign(16)= L16c
  Lampz.Callback(16) = "DisableLighting p16, 200,"

  Lampz.MassAssign(17)= L17
  Lampz.MassAssign(17)= L17a
  Lampz.Callback(17) = "DisableLighting p17, 200,"

  Lampz.MassAssign(18)= L18
  Lampz.MassAssign(18)= L18a
  Lampz.Callback(18) = "DisableLighting p18, 200,"

  Lampz.MassAssign(19)= L19
  Lampz.MassAssign(19)= L19a
  Lampz.Callback(19) = "DisableLighting p19, 200,"

  Lampz.MassAssign(20)= L20
  Lampz.MassAssign(20)= L20a
  Lampz.Callback(20) = "DisableLighting p20, 200,"

  Lampz.MassAssign(21)= L21
  Lampz.MassAssign(21)= L21a
  Lampz.Callback(21) = "DisableLighting p21, 200,"

  Lampz.MassAssign(22)= L22
  Lampz.MassAssign(22)= L22a
  Lampz.Callback(22) = "DisableLighting p22, 200,"

  Lampz.MassAssign(23)= L23
  Lampz.MassAssign(23)= L23a
  Lampz.Callback(23) = "DisableLighting p23, 200,"

  Lampz.MassAssign(24)= L24
  Lampz.MassAssign(24)= L24a
  Lampz.Callback(24) = "DisableLighting p24, 200,"

  Lampz.MassAssign(25)= L25
  Lampz.MassAssign(25)= L25a
  Lampz.Callback(25) = "DisableLighting p25, 200,"

  Lampz.MassAssign(26)= L26
  Lampz.MassAssign(26)= L26a
  Lampz.Callback(26) = "DisableLighting p26, 200,"

  Lampz.MassAssign(27)= L27
  Lampz.MassAssign(27)= L27b

  Lampz.MassAssign(28)= L28
  Lampz.MassAssign(28)= L28b

  Lampz.MassAssign(29)= L29
  Lampz.MassAssign(29)= L29b

  Lampz.MassAssign(30)= L30
  Lampz.MassAssign(30)= L30b

  Lampz.MassAssign(31)= L31
  Lampz.MassAssign(31)= L31b

  Lampz.MassAssign(32)= L32
  Lampz.MassAssign(32)= L32b

  Lampz.MassAssign(33)= L33
  Lampz.MassAssign(34)= L34
  Lampz.MassAssign(35)= L35
  Lampz.MassAssign(36)= L36
  Lampz.MassAssign(37)= L37
  Lampz.MassAssign(38)= L38
  Lampz.MassAssign(39)= L39
  Lampz.MassAssign(40)= L40
  Lampz.MassAssign(41)= L41
  Lampz.MassAssign(42)= L42
  Lampz.MassAssign(43)= L43
  Lampz.MassAssign(44)= L44
  Lampz.MassAssign(45)= L45
  Lampz.MassAssign(46)= L46
  Lampz.MassAssign(47)= L47
  Lampz.MassAssign(48)= L48
  Lampz.MassAssign(55)= L55
  Lampz.MassAssign(56)= L56
  Lampz.Callback(33) = "DisableLighting Pincab_BackglassL33, 1,"
  Lampz.Callback(34) = "DisableLighting Pincab_BackglassL34, 1,"
  Lampz.Callback(35) = "DisableLighting Pincab_BackglassL35, 1,"
  Lampz.Callback(36) = "DisableLighting Pincab_BackglassL36, 1,"
  Lampz.Callback(37) = "DisableLighting Pincab_BackglassL37, 1,"
  Lampz.Callback(38) = "DisableLighting Pincab_BackglassL38, 1,"
  Lampz.Callback(39) = "DisableLighting Pincab_BackglassL39, 1,"
  Lampz.Callback(40) = "DisableLighting Pincab_BackglassL40, 1,"
  Lampz.Callback(41) = "DisableLighting Pincab_BackglassL41, 1,"
  Lampz.Callback(42) = "DisableLighting Pincab_BackglassL42, 1,"
  Lampz.Callback(43) = "DisableLighting Pincab_BackglassL43, 1,"
  Lampz.Callback(44) = "DisableLighting Pincab_BackglassL44, 1,"
  Lampz.Callback(45) = "DisableLighting Pincab_BackglassL45, 1,"
  Lampz.Callback(46) = "DisableLighting Pincab_BackglassL46, 1,"
  Lampz.Callback(47) = "DisableLighting Pincab_BackglassL47, 1,"
  Lampz.Callback(48) = "DisableLighting Pincab_BackglassL48, 1,"
  Lampz.Callback(55) = "DisableLighting Pincab_BackglassL55, 1,"
  Lampz.Callback(56) = "DisableLighting Pincab_BackglassL56, 1,"
  Lampz.Callback(57) = "DisableLighting Pincab_BackglassL57, 1,"
  Lampz.Callback(58) = "DisableLighting Pincab_BackglassL58, 1,"
  Lampz.Callback(59) = "DisableLighting Pincab_BackglassL59, 1,"
  Lampz.Callback(60) = "DisableLighting Pincab_BackglassL60, 1,"
  Lampz.Callback(61) = "DisableLighting Pincab_BackglassL61, 1,"
  Lampz.Callback(62) = "DisableLighting Pincab_BackglassL62, 1,"
  Lampz.Callback(63) = "DisableLighting Pincab_BackglassL63, 1,"
  Lampz.Callback(64) = "DisableLighting Pincab_BackglassL64, 1,"

  Lampz.MassAssign(49)= L49
  Lampz.MassAssign(49)= L49a
  Lampz.Callback(49) = "DisableLighting p49, 200,"

  Lampz.MassAssign(50)= L50
  Lampz.MassAssign(50)= L50a
  Lampz.Callback(50) = "DisableLighting p50, 200,"

  Lampz.MassAssign(51)= L51
  Lampz.MassAssign(51)= L51a
  Lampz.Callback(51) = "DisableLighting p51, 200,"

  Lampz.MassAssign(52)= L52
  Lampz.MassAssign(52)= L52a
  Lampz.Callback(52) = "DisableLighting p52, 200,"

  Lampz.MassAssign(53)= L53
  Lampz.MassAssign(53)= L53a
  Lampz.Callback(53) = "DisableLighting p53, 200,"

  Lampz.MassAssign(54)= L54
  Lampz.MassAssign(54)= L54a
  Lampz.Callback(54) = "DisableLighting p54, 200,"

  Lampz.MassAssign(57)= L57
  Lampz.MassAssign(57)= L57b

  Lampz.MassAssign(58)= L58
  Lampz.MassAssign(58)= L58b

  Lampz.MassAssign(58)= L59
  Lampz.MassAssign(58)= L59b

  Lampz.MassAssign(60)= L60
  Lampz.MassAssign(60)= L60b

  Lampz.MassAssign(61)= L61
  Lampz.MassAssign(61)= L61b

  Lampz.MassAssign(62)= L62
  Lampz.MassAssign(62)= L62b

  Lampz.MassAssign(63)= L63
  Lampz.MassAssign(63)= L63b

  Lampz.MassAssign(64)= L64
  Lampz.MassAssign(64)= L64b

  If UseVPMModSol = 0 Then
    'Turn off all lamps on startup
    Lampz.Init  'This just turns state of any lamps to 1
    'Immediate update to turn on GI, turn off lamps
    Lampz.Update
  Else
    For x = 0 to 150: Lampz.State(x) = 0: Next
  End If
End Sub


'====================
'Class jungle nf
'====================


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
'Version 0.14 - Updated to support modulated signals - Niwak

Class LampFader
  Public FadeSpeedDown(150), FadeSpeedUp(150)
  Private Lock(150), Loaded(150), OnOff(150)
  Public UseFunction
  Private cFilter
  Public UseCallback(150), cCallback(150)
  Public Lvl(150), Obj(150)
  Private Mult(150)
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
      OnOff(x) = 0
      Lock(x) = True : Loaded(x) = False
      Mult(x) = 1
    Next
    Name = "LampFaderNF" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    for x = 0 to uBound(OnOff)    'clear out empty obj
      if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
    Next
  End Sub

  Public Property Get Locked(idx) : Locked = Lock(idx) : End Property   ''debug.print Lampz.Locked(100) 'debug
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
    'msgbox "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    dim out : out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    ExecuteGlobal Out

  End Property

  Public Property Let state(ByVal idx, input) 'Major update path
    if TypeName(input) <> "Double" and typename(input) <> "Integer"  and typename(input) <> "Long" then
      If input Then
        input = 1
      Else
        input = 0
      End If
    End If
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
    ''debug.print debugstr
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
        if OnOff(x) > 0 then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x)
          if Lvl(x) >= OnOff(x) then Lvl(x) = OnOff(x) : Lock(x) = True
        else 'fade down
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
        if OnOff(x) > 0 then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= OnOff(x) then Lvl(x) = OnOff(x) : Lock(x) = True
        else 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
    Update
  End Sub

  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    dim x,xx, aLvl : for x = 0 to uBound(OnOff)
      if not Loaded(x) then
        aLvl = Lvl(x)*Mult(x)
        if IsArray(obj(x) ) Then  'if array
          If UseFunction then
            for each xx in obj(x) : xx.IntensityScale = cFilter(aLvl) : Next
          Else
            for each xx in obj(x) : xx.IntensityScale = aLvl : Next
          End If

        else            'if single lamp or flasher
          If UseFunction then
            obj(x).Intensityscale = cFilter(aLvl)
          Else
            obj(x).Intensityscale = aLvl
          End If
        end if
        'if TypeName(lvl(x)) <> "Double" and typename(lvl(x)) <> "Integer" and typename(lvl(x)) <> "Long" then msgbox "uhh " & 2 & " = " & lvl(x)
        'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
        If UseCallBack(x) then Proc name & x,aLvl 'Proc
        If Lock(x) Then
          if Lvl(x) = OnOff(x) or Lvl(x) = 0 then Loaded(x) = True  'finished fading
        end if
      end if
    Next
  End Sub
End Class

Class VPMLampUpdater
  Public Name
  Public Obj(150), OnOff(150)
  Private UseCallback(150), cCallback(150)

  Sub Class_Initialize()
    Name = "VPMLampUpdater" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    Dim x : For x = 0 to uBound(OnOff)
        OnOff(x) = 0
      Set Obj(x) = NullFader
    Next
  End Sub

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
  End Property

  Sub SetLamp(aIdx, aOn) : state(aIdx) = aOn : End Sub

  Public Property Let state(ByVal x, input)
    Dim xx
    OnOff(x) = input
    If IsArray(obj(x)) Then
      For Each xx In obj(x)
        xx.IntensityScale = input
        'debug.print x&"  obj.Intensityscale = " & input
      Next
    Else
      obj(x).Intensityscale = input
      'debug.print "obj("&x&").Intensityscale = " & input
    End if
    'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
    If UseCallBack(x) then Proc name & x,input
  End Property

  Public Property Get state(idx) : state = OnOff(idx) : end Property

End Class


'Lamp Filter
Function LampFilter(aLvl)
  LampFilter = aLvl^1.6 'exponential curve?
End Function

'Helper functions
Sub Proc(String, Callback)  'proc using a string and one argument
  'On Error Resume Next
  Dim p
  Set P = GetRef(String)
  P Callback
  If err.number = 13 Then  MsgBox "Proc error! No such procedure: " & vbNewLine & String
  If err.number = 424 Then MsgBox "Proc error! No such Object"
End Sub

Function AppendArray(ByVal aArray, aInput)  'append one value, object, or Array onto the end of a 1 dimensional array
  If IsArray(aInput) Then 'Input is an array...
    Dim tmp
    tmp = aArray
    If Not IsArray(aArray) Then 'if not array, create an array
      tmp = aInput
    Else          'Append existing array with aInput array
      ReDim Preserve tmp(UBound(aArray) + UBound(aInput) + 1) 'If existing array, increase bounds by uBound of incoming array
      Dim x
      For x = 0 To UBound(aInput)
        If IsObject(aInput(x)) Then
          Set tmp(x + UBound(aArray) + 1 ) = aInput(x)
        Else
          tmp(x + UBound(aArray) + 1 ) = aInput(x)
        End If
      Next
      AppendArray = tmp  'return new array
    End If
  Else 'Input is NOT an array...
    If Not IsArray(aArray) Then 'if not array, create an array
      aArray = Array(aArray, aInput)
    Else
      ReDim Preserve aArray(UBound(aArray) + 1) 'If array, increase bounds by 1
      If IsObject(aInput) Then
        Set aArray(UBound(aArray)) = aInput
      Else
        aArray(UBound(aArray)) = aInput
      End If
    End If
    AppendArray = aArray 'return new array
  End If
End Function

'***********************class jungle**************

'******************************************************
'****  END LAMPZ
'******************************************************

'******************************************************
'******  VPW TWEAK MENU
'******************************************************

Dim ColorLUT : ColorLUT = 1                 ' Color desaturation LUTs
Dim Anaglyph : Anaglyph = 0                           ' When set to 1 in the menu this will overwrite the default LUT slot to use the VPW Original 1 on 1
Dim MechVolume : MechVolume = 0.8                   ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5           ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5           ' Level of ramp rolling volume. Value between 0 and 1
Dim StagedFlippers : StagedFlippers = 0               ' Staged Flippers. 0 = Disabled, 1 = Enabled
Dim AmbientBallShadowOn : AmbientBallShadowon = 1     ' Ambient Ball Shadows. 0 = Disabled, 1 = Enabled
Dim RenderProbeOpt : RenderProbeOpt = 1           ' 0 = No Refraction Probes (best performance), 1 = Full Refraction Probes (best visual)
Dim PlayfieldReflections : PlayfieldReflections = 100 ' Defines the reflection strength of the (dynamic) table elements on the playfield (0-100) / 5 = 20 Max
Dim ReflOpt : ReflOpt = 1                 ' 0 = Reflections off, 1 = Reflections on
Dim LightReflOpt : LightReflOpt = 1             ' 0 = Light Reflections off, 1 = Light Reflections on
Dim InsertReflOpt : InsertReflOpt = 1           ' 0 = Insert Reflections off, 1 = Insert Reflections on
Dim BackglassReflOpt : BackglassReflOpt = 1           ' 0 = Backglass Reflections off, 1 = Backglass Reflections on

Dim InstMod, SideBladeMod, LegendsCabMod, GlassMod, DirtyClownMod, BackglassMod, LogoPosterMod

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
  'Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 27, 1, 1, 0, _
    Array("Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White", _
    "Fleep Natural Dark 1", "Fleep Natural Dark 2", "Fleep Warm Dark", "Fleep Warm Bright", "Fleep Warm Vivid Soft", "Fleep Warm Vivid Hard", "Skitso Natural & Balance", "Skitso Natural High Contrast", _
    "3rdaxis THX Standard", "Callev Brightness & Contrast", "Hauntfreaks Desaturated", "Tomate Washed out", "VPW Original 1 on 1", "Bassgeige", "Blacklight", "B&W Comic Book"))
  if ColorLUT = 1 And Anaglyph = 0 Then Table1.ColorGradeImage = "" '"Normal"
  if ColorLUT = 1 And Anaglyph = 1 Then Table1.ColorGradeImage = "colorgradelut256x16-vpw-original-1-on-1" '"VPW Original 1 on 1"
  if ColorLUT = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10" '"Desaturated 10%"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20" '"Desaturated 20%"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-30" '"Desaturated 30%"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "colorgradelut256x16-40" '"Desaturated 40%"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "colorgradelut256x16-50" '"Desaturated 50%"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "colorgradelut256x16-60" '"Desaturated 60%"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "colorgradelut256x16-70" '"Desaturated 70%"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "colorgradelut256x16-80" '"Desaturated 80%"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "colorgradelut256x16-90" '"Desaturated 90%"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "colorgradelut256x16-100" '"Black 'n White"
  if ColorLUT = 12 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-natural-dark-1" '"Fleep Natural Dark 1"
  if ColorLUT = 13 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-natural-dark-2" '"Fleep Natural Dark 2"
  if ColorLUT = 14 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-warm-dark" '"Fleep Warm Dark "
  if ColorLUT = 15 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-warm-bright" '"Fleep Warm Bright"
  if ColorLUT = 16 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-warm-vivid-soft" '"Fleep Warm Vivid Soft"
  if ColorLUT = 17 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-warm-vivid-hard" '"Fleep Warm Vivid Hard"
  if ColorLUT = 18 Then Table1.ColorGradeImage = "colorgradelut256x16-skitso-natural-and-balance" '"Skitso Natural & Balance"
  if ColorLUT = 19 Then Table1.ColorGradeImage = "colorgradelut256x16-skitso-natural-high-contrast" '"Skitso Natural High Contrast"
  if ColorLUT = 20 Then Table1.ColorGradeImage = "colorgradelut256x16-3rdaxis-thx-standard" '"3rdaxis THX Standard"
  if ColorLUT = 21 Then Table1.ColorGradeImage = "colorgradelut256x16-callev-brightness-and-contrast" '"Callev Brightness & Contrast"
  if ColorLUT = 22 Then Table1.ColorGradeImage = "colorgradelut256x16-hauntfreaks-desaturated" '"Hauntfreaks Desaturated"
  if ColorLUT = 23 Then Table1.ColorGradeImage = "colorgradelut256x16-tomate-washed-out" '"Tomate Washed out"
  if ColorLUT = 24 Then Table1.ColorGradeImage = "colorgradelut256x16-vpw-original-1-on-1" '"VPW Original 1 on 1"
  if ColorLUT = 25 Then Table1.ColorGradeImage = "colorgradelut256x16-bassgeige" '"Bassgeige"
  if ColorLUT = 26 Then Table1.ColorGradeImage = "colorgradelut256x16-blacklight" '"Blacklight"
  if ColorLUT = 27 Then Table1.ColorGradeImage = "colorgradelut256x16-b-and-w-comic-book" '"B&W Comic Book"
  'Anaglyph
  Anaglyph = Table1.Option("Anaglyph 3D/Headtracking Fix", 0, 1, 1, 0, 0, Array("Off", "On"))
  SetupAnaglyph
    'Sound volumes
    MechVolume = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)
  'Modulated Lights
  ModSol = Table1.Option("Modulated Lights (Needs Restart)", 1, 2, 1, 2, 0, Array("Off", "PWM"))
  SaveModSol
  'Ambient BallShadow
  AmbientBallShadowOn = Table1.Option("Ambient Ballshadow", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetupAmbientBallshadow
  'Reflections
  ReflOpt = Table1.Option("Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetRefl ReflOpt
  'Playfield Reflections
  PlayfieldReflections = Table1.Option("Playfield Reflections", 0, 1, 0.05, 1, 1)
  Table1.PlayfieldReflectionStrength = (PlayfieldReflections * 20)
  'Light Reflections
  LightReflOpt = Table1.Option("Light Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetLightRefl LightReflOpt
  'Insert Reflections
  InsertReflOpt = Table1.Option("Insert Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetInsertRefl InsertReflOpt
  'Backglass Reflections
  BackglassReflOpt = Table1.Option("Backglass Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetBackglassRefl BackglassReflOpt
  'Refraction Probe Setting
  RenderProbeOpt = Table1.Option("Refraction Probe Setting", 0, 1, 1, 1, 0, Array("No Refraction (best performance)", "Full Refraction (best visuals)"))
  SetRenderProbes RenderProbeOpt
  'Instruction Cards
  InstMod = Table1.Option("Instruction Cards", 0, 2, 1, 0, 0, Array("Default Instruction Cards", "Custom Instruction Cards 1", "Custom Instruction Cards 2"))
  SetupInstructionCards
  'Sideblades
  SideBladeMod = Table1.Option("Side Blades", 0, 1, 1, 0, 0, Array("Original", "Retrorefurbs"))
  SetupBlades
  'LegendsCab
  LegendsCabMod = Table1.Option("Legends Cabinet", 0, 1, 1, 0, 0, Array("Off", "On"))
  SetupLegendsCab
  'Scratched Glass
  GlassMod = Table1.Option("Scratched Glass", 0, 1, 1, 0, 0, Array("Off", "On"))
  SetupGlass
  'VR Room
    VRRoomChoice = Table1.Option("VR Room (VR-FSS)", 1, 3, 1, 1, 0, Array("Mega Room", "Minimal Room", "Mixed Reality"))
  SetupRoom
  'VR Gameroom Dirty Clown
  DirtyClownMod = Table1.Option("VR Gameroom Dirty Clown (VR-FSS)", 0, 1, 1, 0, 0, Array("Off", "On"))
  SetupClown
  'VR Minimal Room Logo & Poster (VR-FSS)
  LogoPosterMod = Table1.Option("VR Minimal Room Logo & Poster (VR-FSS)", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetupLogoPoster
  If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Sub SetupAnaglyph
  if ColorLUT = 1 And Anaglyph = 0 Then Table1.ColorGradeImage = "" '"Normal"
  if ColorLUT = 1 And Anaglyph = 1 Then Table1.ColorGradeImage = "colorgradelut256x16-vpw-original-1-on-1" '"VPW Original 1 on 1"
End Sub

Sub SaveModSol
  SaveValue cGameName, "MODSOL", ModSol
  If ModSol = 2 Then
    For each xx in GI : xx.Fader = 0 : Next
    For each xx in AllLamps : xx.Fader = 0 : Next
    For each xx in aReels : xx.Fader = 0 : Next
    For each xx in BumperLights : xx.Fader = 0 : Next
    For each xx in FlasherLights : xx.Fader = 0 : Next
  Else
    For each xx in GI : xx.Fader = 2 : Next
    For each xx in AllLamps : xx.Fader = 2 : Next
    For each xx in aReels : xx.Fader = 2 : Next
    For each xx in BumperLights : xx.Fader = 2 : Next
    For each xx in FlasherLights : xx.Fader = 2 : Next
  End If
End Sub

Sub SetupAmbientBallshadow
  If AmbientBallShadowOn = 1 Then
    For each xx in GI : xx.Shadows = 1 : Next
  Else
    For each xx in GI : xx.Shadows = 0 : Next
  End If
End Sub

Sub SetRefl(Opt)
  Dim xx
  Select Case Opt
    Case 0:
      For each xx in Reflections: xx.ReflectionEnabled = False : Next
    Case 1:
      For each xx in Reflections: xx.ReflectionEnabled = True : Next
  End Select
End Sub

Sub SetLightRefl(Opt)
  Dim xx
  Select Case Opt
    Case 0:
      For each xx in LightReflections: xx.ShowReflectionOnBall = False : Next
    Case 1:
      For each xx in LightReflections: xx.ShowReflectionOnBall = True : Next
  End Select
End Sub

Sub SetInsertRefl(Opt)
  Dim xx
  Select Case Opt
    Case 0:
      For each xx in InsertReflections: xx.ShowReflectionOnBall = False : Next
    Case 1:
      For each xx in InsertReflections: xx.ShowReflectionOnBall = True : Next
  End Select
End Sub

Sub SetBackglassRefl(Opt)
  Dim xx
  Select Case Opt
    Case 0:
      For each xx in BackglassReflections: xx.ReflectionEnabled = False : Next
      Pincab_Backglass.ReflectionEnabled = False
      Pincab_BackglassS11.ReflectionEnabled = False
      Pincab_BackglassS25.ReflectionEnabled = False
      Pincab_BackglassS26.ReflectionEnabled = False
      Pincab_BackglassS27.ReflectionEnabled = False
      Pincab_BackglassS28.ReflectionEnabled = False
      Pincab_BackglassS29.ReflectionEnabled = False
      Pincab_BackglassS31.ReflectionEnabled = False
      Pincab_BackglassS32.ReflectionEnabled = False
      Pincab_BackglassL33.ReflectionEnabled = False
      Pincab_BackglassL34.ReflectionEnabled = False
      Pincab_BackglassL35.ReflectionEnabled = False
      Pincab_BackglassL36.ReflectionEnabled = False
      Pincab_BackglassL37.ReflectionEnabled = False
      Pincab_BackglassL38.ReflectionEnabled = False
      Pincab_BackglassL39.ReflectionEnabled = False
      Pincab_BackglassL40.ReflectionEnabled = False
      Pincab_BackglassL41.ReflectionEnabled = False
      Pincab_BackglassL42.ReflectionEnabled = False
      Pincab_BackglassL43.ReflectionEnabled = False
      Pincab_BackglassL44.ReflectionEnabled = False
      Pincab_BackglassL45.ReflectionEnabled = False
      Pincab_BackglassL46.ReflectionEnabled = False
      Pincab_BackglassL47.ReflectionEnabled = False
      Pincab_BackglassL48.ReflectionEnabled = False
      Pincab_BackglassL55.ReflectionEnabled = False
      Pincab_BackglassL56.ReflectionEnabled = False
      Pincab_BackglassL57.ReflectionEnabled = False
      Pincab_BackglassL58.ReflectionEnabled = False
      Pincab_BackglassL59.ReflectionEnabled = False
      Pincab_BackglassL60.ReflectionEnabled = False
      Pincab_BackglassL61.ReflectionEnabled = False
      Pincab_BackglassL62.ReflectionEnabled = False
      Pincab_BackglassL63.ReflectionEnabled = False
      Pincab_BackglassL64.ReflectionEnabled = False
      If VRMode = False Then
        For Each VR_Obj in BackglassReflections : VR_Obj.Visible = 0 : Next
        Pincab_Backglass.Visible = 0
        Pincab_BackglassS11.Visible = 0
        Pincab_BackglassS25.Visible = 0
        Pincab_BackglassS26.Visible = 0
        Pincab_BackglassS27.Visible = 0
        Pincab_BackglassS28.Visible = 0
        Pincab_BackglassS29.Visible = 0
        Pincab_BackglassS31.Visible = 0
        Pincab_BackglassS32.Visible = 0
        Pincab_BackglassL33.Visible = 0
        Pincab_BackglassL34.Visible = 0
        Pincab_BackglassL35.Visible = 0
        Pincab_BackglassL36.Visible = 0
        Pincab_BackglassL37.Visible = 0
        Pincab_BackglassL38.Visible = 0
        Pincab_BackglassL39.Visible = 0
        Pincab_BackglassL40.Visible = 0
        Pincab_BackglassL41.Visible = 0
        Pincab_BackglassL42.Visible = 0
        Pincab_BackglassL43.Visible = 0
        Pincab_BackglassL44.Visible = 0
        Pincab_BackglassL45.Visible = 0
        Pincab_BackglassL46.Visible = 0
        Pincab_BackglassL47.Visible = 0
        Pincab_BackglassL48.Visible = 0
        Pincab_BackglassL55.Visible = 0
        Pincab_BackglassL56.Visible = 0
        Pincab_BackglassL57.Visible = 0
        Pincab_BackglassL58.Visible = 0
        Pincab_BackglassL59.Visible = 0
        Pincab_BackglassL60.Visible = 0
        Pincab_BackglassL61.Visible = 0
        Pincab_BackglassL62.Visible = 0
        Pincab_BackglassL63.Visible = 0
        Pincab_BackglassL64.Visible = 0
        Pincab_BackglassS11.Visible = 0
        Pincab_BackglassS25.Visible = 0
        Pincab_BackglassS26.Visible = 0
        Pincab_BackglassS27.Visible = 0
        Pincab_BackglassS28.Visible = 0
        Pincab_BackglassS29.Visible = 0
        Pincab_BackglassS31.Visible = 0
        Pincab_BackglassS32.Visible = 0
      End If
    Case 1:
      For Each xx in BackglassReflections: xx.ReflectionEnabled = True : Next
      For Each VR_Obj in BackglassReflections : VR_Obj.Visible = 1 : Next
      Pincab_Backglass.ReflectionEnabled = True
      Pincab_BackglassS11.ReflectionEnabled = True
      Pincab_BackglassS25.ReflectionEnabled = True
      Pincab_BackglassS26.ReflectionEnabled = True
      Pincab_BackglassS27.ReflectionEnabled = True
      Pincab_BackglassS28.ReflectionEnabled = True
      Pincab_BackglassS29.ReflectionEnabled = True
      Pincab_BackglassS31.ReflectionEnabled = True
      Pincab_BackglassS32.ReflectionEnabled = True
      Pincab_BackglassL33.ReflectionEnabled = True
      Pincab_BackglassL34.ReflectionEnabled = True
      Pincab_BackglassL35.ReflectionEnabled = True
      Pincab_BackglassL36.ReflectionEnabled = True
      Pincab_BackglassL37.ReflectionEnabled = True
      Pincab_BackglassL38.ReflectionEnabled = True
      Pincab_BackglassL39.ReflectionEnabled = True
      Pincab_BackglassL40.ReflectionEnabled = True
      Pincab_BackglassL41.ReflectionEnabled = True
      Pincab_BackglassL42.ReflectionEnabled = True
      Pincab_BackglassL43.ReflectionEnabled = True
      Pincab_BackglassL44.ReflectionEnabled = True
      Pincab_BackglassL45.ReflectionEnabled = True
      Pincab_BackglassL46.ReflectionEnabled = True
      Pincab_BackglassL47.ReflectionEnabled = True
      Pincab_BackglassL48.ReflectionEnabled = True
      Pincab_BackglassL55.ReflectionEnabled = True
      Pincab_BackglassL56.ReflectionEnabled = True
      Pincab_BackglassL57.ReflectionEnabled = True
      Pincab_BackglassL58.ReflectionEnabled = True
      Pincab_BackglassL59.ReflectionEnabled = True
      Pincab_BackglassL60.ReflectionEnabled = True
      Pincab_BackglassL61.ReflectionEnabled = True
      Pincab_BackglassL62.ReflectionEnabled = True
      Pincab_BackglassL63.ReflectionEnabled = True
      Pincab_BackglassL64.ReflectionEnabled = True
      Pincab_Backglass.Visible = 1
      Pincab_BackglassS11.Visible = 1
      Pincab_BackglassS25.Visible = 1
      Pincab_BackglassS26.Visible = 1
      Pincab_BackglassS27.Visible = 1
      Pincab_BackglassS28.Visible = 1
      Pincab_BackglassS29.Visible = 1
      Pincab_BackglassS31.Visible = 1
      Pincab_BackglassS32.Visible = 1
      Pincab_BackglassL33.Visible = 1
      Pincab_BackglassL34.Visible = 1
      Pincab_BackglassL35.Visible = 1
      Pincab_BackglassL36.Visible = 1
      Pincab_BackglassL37.Visible = 1
      Pincab_BackglassL38.Visible = 1
      Pincab_BackglassL39.Visible = 1
      Pincab_BackglassL40.Visible = 1
      Pincab_BackglassL41.Visible = 1
      Pincab_BackglassL42.Visible = 1
      Pincab_BackglassL43.Visible = 1
      Pincab_BackglassL44.Visible = 1
      Pincab_BackglassL45.Visible = 1
      Pincab_BackglassL46.Visible = 1
      Pincab_BackglassL47.Visible = 1
      Pincab_BackglassL48.Visible = 1
      Pincab_BackglassL55.Visible = 1
      Pincab_BackglassL56.Visible = 1
      Pincab_BackglassL57.Visible = 1
      Pincab_BackglassL58.Visible = 1
      Pincab_BackglassL59.Visible = 1
      Pincab_BackglassL60.Visible = 1
      Pincab_BackglassL61.Visible = 1
      Pincab_BackglassL62.Visible = 1
      Pincab_BackglassL63.Visible = 1
      Pincab_BackglassL64.Visible = 1
      Pincab_BackglassS11.Visible = 1
      Pincab_BackglassS25.Visible = 1
      Pincab_BackglassS26.Visible = 1
      Pincab_BackglassS27.Visible = 1
      Pincab_BackglassS28.Visible = 1
      Pincab_BackglassS29.Visible = 1
      Pincab_BackglassS31.Visible = 1
      Pincab_BackglassS32.Visible = 1
  End Select
End Sub

Sub SetRenderProbes(Opt)
  On Error Resume Next
    Select Case Opt
      Case 0:
        Primitive25.RefractionProbe = ""
        Primitive26.RefractionProbe = ""
        Primitive24.RefractionProbe = ""
      Case 1:
        Primitive25.RefractionProbe = "cometramp"
        Primitive26.RefractionProbe = "wheelramp"
        Primitive24.RefractionProbe = "wheelramp"
    End Select
  On Error Goto 0
End Sub

Sub SetupInstructionCards
  If InstMod = 0 Then
    Card1.imageA = "InstCard1"
    Card2.imageA = "InstCard2"
  End If
  If InstMod = 1 Then
    Card1.imageA = "InstCard_Custom1"
    Card2.imageA = "InstCard_Custom2"
  End If
  If InstMod = 2 Then
    Card1.imageA = "InstCard_Custom2_1"
    Card2.imageA = "InstCard_Custom2_2"
  End If
End Sub

Sub SetupBlades
  If SideBladeMod = 0 Then
    PinCab_Blades.image = "PinCab_Blades"
    SideBladesLegends.image = "Sidewalls_C_Original"
  End If
  If SideBladeMod = 1 Then
    PinCab_Blades.image = "PinCab_Blades_Custom1"
    SideBladesLegends.image = "Sidewalls_C"
  End if
End Sub

Sub SetupLegendsCab
  If Table1.ShowDT = False And Table1.ShowFSS = False And RenderingMode <> 2 And LegendsCabMod = 1 Then
    PinCab_Blades.Visible = 0
    SideBladesLegends.Visible = 1
    Flasherbloom1.Height = 500
    Flasherbloom2.Height = 500
    Flasherbloom3.Height = 500
    Flasherbloom4.Height = 500
    Flasherbloom5.Height = 500
    Flasherbloom6.Height = 500
    Flasherbloom7.Height = 500
    Flasherlampbloom1. Height = 500
    Flasherlampbloom2. Height = 500
    Flasherlampbloom3. Height = 500
    Flasherlampbloom4. Height = 500
  Else
    PinCab_Blades.Visible = 1
    SideBladesLegends.Visible = 0
    Flasherbloom1.Height = 250
    Flasherbloom2.Height = 250
    Flasherbloom3.Height = 250
    Flasherbloom4.Height = 250
    Flasherbloom5.Height = 250
    Flasherbloom6.Height = 250
    Flasherbloom7.Height = 250
    Flasherlampbloom1. Height = 250
    Flasherlampbloom2. Height = 250
    Flasherlampbloom3. Height = 250
    Flasherlampbloom4. Height = 250
  End If
End Sub

Sub SetupGlass
  If GlassMod = 1 Then
    PinCab_Glass_scratches.Visible = True
  Else
    PinCab_Glass_scratches.Visible = False
  End If
End Sub

Sub SetupRoom
  If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
    VRMode = True
    PinCab_Rails.Visible = 1
    For Each VR_Obj in VRCabinet:VR_Obj.Visible = 1:Next
    If VRRoomChoice = 1 Then
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
      For Each VR_Obj in VRGameRoom : VR_Obj.Visible = 1 : Next
      Room360.Visible = 0
      VR_Room_Logo.Visible = False
      Poster1.Visible = False
      Poster2.Visible = False
      Poster3.Visible = False
    End If
    If VRRoomChoice = 2 Then
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
      For Each VR_Obj in VRGameRoom : VR_Obj.Visible = 0 : Next
      Room360.Visible = 0
      If LogoPosterMod = 1 Then
        VR_Room_Logo.Visible = True
        Poster1.Visible = True
        Poster2.Visible = True
        Poster3.Visible = True
      Else
        VR_Room_Logo.Visible = False
        Poster1.Visible = False
        Poster2.Visible = False
        Poster3.Visible = False
      End If
    End If
    If VRRoomChoice = 3 Then
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
      For Each VR_Obj in VRGameRoom : VR_Obj.Visible = 0 : Next
      Room360.Visible = 1
      VR_Room_Logo.Visible = False
      Poster1.Visible = False
      Poster2.Visible = False
      Poster3.Visible = False
    End If
  End If
End Sub

Sub SetupClown
  If DirtyClownMod = 0 Then
    VRFrontwall.image = "BackWall"
  End If
  If DirtyClownMod = 1 Then
    VRFrontwall.image = "BackWallDirty"
  End If
End Sub

Sub SetupLogoPoster
  If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
    If VRRoomChoice = 2 Then
      If LogoPosterMod = 1 Then
        VR_Room_Logo.Visible = True
        Poster1.Visible = True
        Poster2.Visible = True
        Poster3.Visible = True
      Else
        VR_Room_Logo.Visible = False
        Poster1.Visible = False
        Poster2.Visible = False
        Poster3.Visible = False
      End If
    Else
      VR_Room_Logo.Visible = False
      Poster1.Visible = False
      Poster2.Visible = False
      Poster3.Visible = False
    End If
  End If
End Sub
