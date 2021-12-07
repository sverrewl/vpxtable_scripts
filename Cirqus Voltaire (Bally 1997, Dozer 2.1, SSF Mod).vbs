' Cirqus Voltaire / IPD No. 4059 / October, 1997 / 4 Players
' VP913 1.2 by JPSalas 2012
' Thanks to all the authors (Pinball Ken, Scapino, Pacdude, Fuseball, Wpcmame) who made this table before me.
' Since I have never played or seen the real table, this table is based on their tables.
' Thanks to Strangeleo for asking this table and his help with some graphics and testing the table.
' Parts of the script from the older tables.

' VPX version by Dozer and ninuzzu

' Thalamus 2018-07-24
' Table has already "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-12-18 : Added FFv2
' No special SSF tweaks yet.

Option Explicit
Randomize

'  .5 = lower volume
' 1.5 = higher volume

Const VolFlip   = 1    ' Flipper volume.

'-------------------------------------------------

Const Dozer_Cab = 0 ' LEAVE THIS TURNED OFF

'-------------------------------------------------

dim B2SOn:B2SOn = 1 'Activate The DirectB2S / DOF system.
Const DOFs = False 'Activate this to mute mech sounds for DOF
'-------------------------------------------------

'##### Choose the ROM, uncomment to use

'Const cGameName = "cv_20h" 'home rom without credits
Const cGameName = "cv_20hc" 'arcade rom with credits
'Const cGameName = "cv_14" 'arcade rom with credits

'-------------------------------------------------

'******** DMD OPTIONS *************
'
'The 3 options below decide how the DMD is going to be rendered.
'
'Setting this option to 0 will use the old style DMD which sits in front
'of the ramps and is not integrated into the table.
'
'Setting this option to 1 will use the new flasher based integrated DMD.
'By also setting the variable "dmdcolor" below you can change the color
'of the DMD to one of the 8 colors listed.

'Setting this option to 2 will use the new flasher based integrated DMD
'and will also allow you to set the DMD to a special 4 color mode which
'must be enabled via the VPM setup program.  If you use this option the "dmdcolor"
'variable below will have no effect as the DMD must be set to a base color to render
'the 4 color option properly.
'

Const IntegratedDMDType = 1 '0 = Off / 1 Normal / 2 Color.

'-------------------------------------------------

Const DMD_Plastic = 0 'Render a plastic cover over the DMD display.

'-------------------------------------------------

Const BackGlassDMD = 0 ' Place a static image on the playfield DMD if you use a 3 screen cab and don't want to use either
                       ' the integrated DMD or the normal DMD positioned on the playfield.
'-------------------------------------------------

Const dmdcolor = 2 'Set the color of the NORMAL Integrated DMD - '0 = Orange / 1 = Green / 2 = Red / 3 = Pink / 4 = Blue / 5 Cyan / 6 Purple / 7 Yellow

'-------------------------------------------------

Const DMD_Rot = 0 'Rotate the Standard DMD - 1 = Rotate/ 0 = Normal

'-------------------------------------------------




'********* GRAPHICS & TABLE OPTIONS ************


Const Ringmaster_Motor = 1 'Play a motor sound when the Ringmaster raises / lowers.
Const Ringmaster_Speed = 120

'-------------------------------------------------

Const Ramp_End_Bounce = 1 'Simulate the ball boucing off the ramp ends before dropping.

'-------------------------------------------------

'##### Render Spiral Flipper logos instead of Williams

Const Spiral_Flipper_Logo = 0

'-------------------------------------------------

'##### Render Rubber Posts with Yellow Texture instead of Black.

Const Yellow_Prim_Posts = 1

'-------------------------------------------------

'##### Turn on Ramp Ambient Flasher effect.

'1 = On / 0 = Off

Const Prim_Reflect = 1

'-------------------------------------------------

'##### Turn The OverBright Neons Tube lights on or off.

'1 = On / 0 = Off

Const BrightTube = 1

'-------------------------------------------------

'#### Is the aftermarket Plamsa Mod Orb Visible?

'1 = On / 0 =  Off

Const Plasmamod = 0


'##########################################
'THE NEXT TWO VARIABLES CONTROL THE DEFAULT COLORS
'OF THE NEON TUBE AND CAPTIVE BALL WHEN THE TABLE
'FIRST STARTS.  THEY CAN BOTH BE CHANGED IN GAME BY
'PRESSING THE EXTRA BALL KEY.
'##########################################

'#### Set the Color of the Neon Tube here.

'1 = Green / 2 = Red / 3 = Pink / 4 = Blue / 5 Cyan / 6 Purple / 7 Yellow / 8 Orange

Const NeonColor = 3

'##########################################


'#### Set the Color of the Menagerie / Cue / Wildball etc. etc. here.

'1 = Green / 2 = Red / 3 = Pink / 4 = Blue / 5 Cyan / 6 Purple / 7 Yellow / 8 Orange

Const Wildballcolor = 3

'##########################################

Dim RampMagBall



'######### END OF TABLE OPTIONS ###########

If IntegratedDMDType = 1 then
If dmdcolor = 0 then
dmd.color = RGB(255,128,0)
dmd1.color = RGB(255,128,0)
dmd.opacity = 170
dmd1.opacity = 170
Ramp1038.image = "backwall999_orange"
End If
If dmdcolor = 1 then
dmd.color = RGB(0,255,0)
dmd1.color = RGB(0,255,0)
dmd.opacity = 100
dmd1.opacity = 100
Ramp1038.image = "backwall999_green"
End If
If dmdcolor = 2 then
dmd.color = RGB(255,0,0)
dmd1.color = RGB(255,0,0)
dmd.opacity = 100
dmd1.opacity = 100
Ramp1038.image = "backwall999_red"
End If
If dmdcolor = 3 then
dmd.color = RGB(255,0,255)
dmd1.color = RGB(255,0,255)
dmd.opacity = 100
dmd1.opacity = 100
Ramp1038.image = "backwall999_purple"
End If
If dmdcolor = 4 then
dmd.color = RGB(0,0,255)
dmd1.color = RGB(0,0,255)
dmd.opacity = 100
dmd1.opacity = 100
Ramp1038.image = "backwall999_blue"
End If
If dmdcolor = 5 then
dmd.color = RGB(0,255,255)
dmd1.color = RGB(0,255,255)
dmd.opacity = 100
dmd1.opacity = 100
Ramp1038.image = "backwall999_cyan"
End If
If dmdcolor = 6 then
dmd.color = RGB(204,153,255)
dmd1.color = RGB(204,153,255)
dmd.opacity = 100
dmd1.opacity = 100
Ramp1038.image = "backwall999_pink"
End If
If dmdcolor = 7 then
dmd.color = RGB(255,255,0)
dmd1.color = RGB(255,255,0)
dmd.opacity = 100
dmd1.opacity = 100
Ramp1038.image = "backwall999_yellow"
End If
End If

If dmd_plastic = 1 Then
DMD2.visible = 1
End If

If ramp_end_bounce = 1 Then
LRSK.enabled = 0
RRSK.enabled = 0
RRS.collidable = 0
LRS.collidable = 0
End If

'If IntegratedDMDType = 0 Then
'Const UseVPMDMD = 0
'End If
If IntegratedDMDType = 1 Then
Const UseVPMDMD = 1
End If
If IntegratedDMDType = 2 Then
Const UseVPMColoredDMD = 1
End If

Dim DesktopMode:DesktopMode = Table1.ShowDT

If DesktopMode = True Then
  Cover1.TopMaterial = "Plastic Cover Clear1"
  Cover2.TopMaterial = "Plastic Cover Clear1"
  Cover3.TopMaterial = "Plastic Cover Clear1"
  Cover4.material = "Plastic Cover Clear1"
  Cover5.material = "Plastic Cover Clear1"
  Cover6.material = "Plastic Cover Clear1"
  Cover7.material = "Plastic Cover Clear1"
  Cover8.material = "Plastic Cover Clear1"
  Plastic_SlingL.material = "Plastic Cover Clear1"
  Plastic_SlingR.material = "Plastic Cover Clear1"
  Plastic_RM.material = "Plastic Cover Clear1"
'Ramp15.image = "Right_rail_brighter_D"
Can1.visible = 1
Can2.visible = 1
Can4.visible = 1
Wheel.visible = 1
Wheel1.visible = 1
F21c.height = 200
F24c.height = 200
F25c.height = 200
F26c.height = 170
F21c.Y = F21c.Y - 150
F24c.Y = F24c.Y - 150
F25c.Y = F25c.Y - 150
F26c.Y = F26c.Y - 150
F26c.X = F26c.X + 20
F21c.RotY = 90
F24c.RotY = 90
F25c.RotY = 90
F26c.RotY = 90
F37c.Height = 270
F37c.RotY = 90
DiscP.ObjRotX = -10
F1.Intensity = 12
'Light127.X = Light127.X -5
'Light7.X = Light7.X - 5
'LH1.visible = 1
'LH2.visible = 1
Prim_CenterRamp.material = "Ramps-Plastic_DT"
Prim_NeonRamp.material = "Ramps-Plastic_DT"
DMD1.visible = 1:DMD.visible = 0
DMD2.visible = 0
If DMD_Plastic = 1 Then
DMD3.visible = 1
Prim_Cabside.size_Z = 1
End If
else
Prim_Cabside.size_Z = 3
F1.Intensity = 0
Can1.visible = 0
Can2.visible = 0
Can4.visible = 0
Wheel.visible = 0
Wheel1.visible = 0
Ramp15.visible = 0
Ramp16.visible = 0
'LH1.visible = 1
'LH2.visible = 1
'LH1.Intensity = 1
'LH2.Intensity = 1
Prim_CenterRamp.material = "Ramps-Plastic"
Prim_NeonRamp.material = "Ramps-Plastic"
DMD1.visible = 0:DMD.visible = 1
End If

'###########################################

Dim Sound1, Sound2, Sound3, Sound4, Sound5, B2SController

 LoadVPM "01560000", "WPC.VBS", 3.46

' Thalamus - for Fast Flip v2
' NoUpperRightFlipper
NoUpperLeftFlipper

 Sub LoadVPM(VPMver, VBSfile, VBSver) 'Add new call to InitializeOptions to allow selection of controller through F6 menu
  On Error Resume Next
    If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
      ExecuteGlobal GetTextFile(VBSfile)
    If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description

    'InitializeOptions 'Enables New Controller change through F6 menu, so it needs to be placed before Controller selection

    'Select Case cController
      'Case 1:
        Set Controller = CreateObject("VPinMAME.Controller")
        If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
        If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
        If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
      'Case 2:
        'Set Controller = CreateObject("UltraVP.BackglassServ")
      'Case 3:
        'Set Controller = CreateObject("B2S.Server")
    'End Select
    If Err then
      msgbox "Invalid controller selected, defaulting to VPinMame"
      Set controller = CreateObject("VPinMAME.Controller")
    End If
      If B2SOn = 1 Then
      Set B2SController=CreateObject("B2S.Server")
      B2SController.B2SName = "cv_dmd"
      B2SController.Run
      if not IsObject(B2SController) then B2SOn=0
      End If
  On Error Goto 0
 End Sub

'==============================
Dim xxpp
If Yellow_Prim_Posts = 1 Then
For each xxpp in Rposts:xxpp.Image = "rubber-post-t1-yellow":next
else
For each xxpp in Rposts:xxpp.Image = "rubber-post-t1-black":next
End If

If Spiral_Flipper_Logo = 1 Then
LogoL.visible = 1
LogoR.visible = 1
else
LogoL.visible = 0
LogoR.visible = 0
End If

If BackGlassDMD = 1 AND IntegratedDMDType = 0 Then
Ramp1038.image = "backwall999_cv"
'DMD2.visible = 0
'else
'Ramp1038.image = "backwall999"
'DMD2.visible = 1
End If

If BrightTube = 1 Then
Light44.Intensity = 2
else
Light44.Intensity = 0
End If

If Plasmamod = 1 Then
DiscP.visible = 1
HexPost9.visible = 1
Bolt55.visible=0
Light118.Intensity = 5
else
DiscP.visible = 0
HexPost9.visible = 0
Bolt55.visible=1
Light118.Intensity = 0
End If

Dim wbcolor
If wildballcolor = 1 then
wbcolor = RGB(0,255,0)
End If
If wildballcolor = 2 then
wbcolor = RGB(255,0,0)
End If
If wildballcolor = 3 then
wbcolor = RGB(255,0,255)
End If
If wildballcolor = 4 then
wbcolor = RGB(0,0,255)
End If
If wildballcolor = 5 then
wbcolor = RGB(0,255,255)
End If
If wildballcolor = 6 then
wbcolor = RGB(204,153,255)
End If
If wildballcolor = 7 then
wbcolor = RGB(255,255,0)
End If
If wildballcolor = 8 then
wbcolor = RGB(255,128,0)
End If

'Colors
Dim bulb

If NeonColor = 1 Then
Neon_On.Color = RGB(0,255,0)
Neon_On.ColorFull = RGB(0,255,0)
Neon_Off.Color = RGB(0,255,0)
Neon_Off.ColorFull = RGB(0,255,0)
for each bulb in aNeons
bulb.Color=RGB(0,255,0)
next
for each bulb in bNeons
bulb.Color=RGB(0,255,0)
next
F37c.imageA = "aGreen"
F37c.imageB = "aGreen"
Prim_Neon_ON.material = "green_tube"
If prim_reflect = 1 Then
'Prim_LockRamp.material = "metal_green"
End If
CNeons = 1
'dmd.opacity = 100
'dmd1.opacity = 100
'Ramp1038.image = "backwall999_green"
End If

If NeonColor = 2 Then
Neon_On.Color = RGB(255,0,0)
Neon_On.ColorFull = RGB(255,0,0)
Neon_Off.Color = RGB(255,0,0)
Neon_Off.ColorFull = RGB(255,0,0)
for each bulb in aNeons
bulb.Color=RGB(255,0,0)
next
for each bulb in bNeons
bulb.Color=RGB(255,0,0)
next
F37c.imageA = "aRed"
F37c.imageB = "aRed"
Prim_Neon_ON.material = "red_tube"
If prim_reflect = 1 Then
'Prim_LockRamp.material = "metal_red"
End If
CNeons = 2
'dmd.opacity = 100
'dmd1.opacity = 100
'Ramp1038.image = "backwall999_red"
End If

If NeonColor = 3 Then
Neon_On.Color = RGB(255,0,255)
Neon_On.ColorFull = RGB(255,0,255)
Neon_Off.Color = RGB(255,0,255)
Neon_Off.ColorFull = RGB(255,0,255)
for each bulb in aNeons
bulb.Color=RGB(255,0,255)
next
for each bulb in bNeons
bulb.Color=RGB(255,0,255)
next
F37c.imageA = "aPurple"
F37c.imageB = "aPurple"
Prim_Neon_ON.material = "purple_tube"
If prim_reflect = 1 Then
'Prim_LockRamp.material = "metal_purple"
End If
CNeons = 3
'dmd.opacity = 100
'dmd1.opacity = 100
'Ramp1038.image = "backwall999_purple"
End If

If NeonColor = 4 Then
Neon_On.Color = RGB(0,0,255)
Neon_On.ColorFull = RGB(0,0,255)
Neon_Off.Color = RGB(0,0,255)
Neon_Off.ColorFull = RGB(0,0,255)
for each bulb in aNeons
bulb.Color=RGB(0,0,255)
next
for each bulb in bNeons
bulb.Color=RGB(0,0,255)
next
F37c.imageA = "aBlue"
F37c.imageB = "aBlue"
Prim_Neon_ON.material = "blue_tube"
If prim_reflect = 1 Then
'Prim_LockRamp.material = "metal_blue"
End If
CNeons = 4
'dmd.opacity = 100
'dmd1.opacity = 100
'Ramp1038.image = "backwall999_blue"
End If

If NeonColor = 5 Then
Neon_On.Color = RGB(0,255,255)
Neon_On.ColorFull = RGB(0,255,255)
Neon_Off.Color = RGB(0,255,255)
Neon_Off.ColorFull = RGB(0,255,255)
for each bulb in aNeons
bulb.Color=RGB(0,255,255)
next
for each bulb in bNeons
bulb.Color=RGB(0,255,255)
next
F37c.imageA = "aCyan"
F37c.imageB = "aCyan"
Prim_Neon_ON.material = "cyan_tube"
If prim_reflect = 1 Then
'Prim_LockRamp.material = "metal_cyan"
End If
CNeons = 5
'dmd.opacity = 100
'dmd1.opacity = 100
'Ramp1038.image = "backwall999_cyan"
End If

If NeonColor = 6 Then
Neon_On.Color = RGB(204,153,255)
Neon_On.ColorFull = RGB(204,153,255)
Neon_Off.Color = RGB(204,153,255)
Neon_Off.ColorFull = RGB(204,153,255)
for each bulb in aNeons
bulb.Color=RGB(204,153,255)
next
for each bulb in bNeons
bulb.Color=RGB(204,153,255)
next
F37c.imageA = "aPink"
F37c.imageB = "aPink"
Prim_Neon_ON.material = "pink_tube"
If prim_reflect = 1 Then
'Prim_LockRamp.material = "metal_pink"
End If
cNeons = 6
'dmd.opacity = 100
'dmd1.opacity = 100
'Ramp1038.image = "backwall999_pink"
End If

If NeonColor = 7 Then
Neon_On.Color = RGB(255,255,0)
Neon_On.ColorFull = RGB(255,255,0)
Neon_Off.Color = RGB(255,255,0)
Neon_Off.ColorFull = RGB(255,255,0)
for each bulb in aNeons
bulb.Color=RGB(255,255,0)
next
for each bulb in bNeons
bulb.Color=RGB(255,255,0)
next
F37c.imageA = "aYellow"
F37c.imageB = "aYellow"
Prim_Neon_ON.material = "yellow_tube"
If prim_reflect = 1 Then
'Prim_LockRamp.material = "metal_yellow"
End If
CNeons = 7
'dmd.opacity = 100
'dmd1.opacity = 100
'Ramp1038.image = "backwall999_yellow"
End If

If NeonColor = 8 Then
Neon_On.Color = RGB(255,128,0)
Neon_On.ColorFull = RGB(255,128,0)
Neon_Off.Color = RGB(255,128,0)
Neon_Off.ColorFull = RGB(255,128,0)
for each bulb in aNeons
bulb.Color=RGB(255,128,0)
next
for each bulb in bNeons
bulb.Color=RGB(255,128,0)
next
F37c.imageA = "aOrange"
F37c.imageB = "aOrange"
Prim_Neon_ON.material = "orange_tube"
If prim_reflect = 1 Then
'Prim_LockRamp.material = "metal_orange"
End If
CNeons = 8
'dmd.opacity = 170
'dmd1.opacity = 170
'Ramp1038.image = "backwall999_orange"
End If

' ----------------------------------------

Dim cneons

Sub Cycle_Neons

WBDes.enabled = 1

NewBCol.enabled = 1

If cneons > 7 Then
cneons = 0
End If

cneons = cneons + 1

If cneons = 8 Then
If IntegratedDMDType = 1 Then
dmd.opacity = 170
dmd1.opacity = 170
DOF 120,0
DOF 121,0
DOF 122,0
DOF 123,0
DOF 124,0
DOF 125,0
DOF 126,0
DOF 127,1
End If
Neon_On.Color = RGB(255,128,0)
Neon_On.ColorFull = RGB(255,128,0)
Neon_Off.Color = RGB(255,128,0)
Neon_Off.ColorFull = RGB(255,128,0)
for each bulb in aNeons
bulb.Color=RGB(255,128,0)
next
for each bulb in bNeons
bulb.Color=RGB(255,128,0)
next
wbcolor = RGB(255,128,0)
If IntegratedDMDType = 1 then
dmd.color = RGB(255,128,0)
dmd1.color = RGB(255,128,0)
End If
F37c.imageA = "aOrange"
F37c.imageB = "aOrange"
''NeonColor = 1
Prim_Neon_ON.material = "orange_tube"
If prim_reflect = 1 Then
Prim_LockRamp.material = "metal_orange"
End If
Prim_CenterRamp.material = "Ramps-Plastic_O"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_ORANGE"
DMD2.imageB = "Plastic_Ramp_2012_ORANGE"
DMD3.imageA = "Plastic_Ramp_2012_ORANGE"
DMD3.imageB = "Plastic_Ramp_2012_ORANGE"
End If
Ramp1038.image = "backwall999_orange"
End If


If cneons = 1 Then
dmd.opacity = 100
dmd1.opacity = 100
DOF 120,1 ' Green
DOF 121,0
DOF 122,0
DOF 123,0
DOF 124,0
DOF 125,0
DOF 126,0
DOF 127,0
Neon_On.Color = RGB(0,255,0)
Neon_On.ColorFull = RGB(0,255,0)
Neon_Off.Color = RGB(0,255,0)
Neon_Off.ColorFull = RGB(0,255,0)
for each bulb in aNeons
bulb.Color=RGB(0,255,0)
next
for each bulb in bNeons
bulb.Color=RGB(0,255,0)
next
wbcolor = RGB(0,255,0)
If IntegratedDMDType = 1 then
dmd.color = RGB(0,255,0)
dmd1.color = RGB(0,255,0)
End If
F37c.imageA = "aGreen"
F37c.imageB = "aGreen"
'NeonColor = 1
Prim_Neon_ON.material = "green_tube"
If prim_reflect = 1 Then
Prim_LockRamp.material = "metal_green"
End If
Prim_CenterRamp.material = "Ramps-Plastic_G"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_GREEN"
DMD2.imageB = "Plastic_Ramp_2012_GREEN"
DMD3.imageA = "Plastic_Ramp_2012_GREEN"
DMD3.imageB = "Plastic_Ramp_2012_GREEN"
End If
Ramp1038.image = "backwall999_green"
End If

If cneons = 2 Then
dmd.opacity = 100
dmd1.opacity = 100
DOF 120,0
DOF 121,1
DOF 122,0
DOF 123,0
DOF 124,0
DOF 125,0
DOF 126,0
DOF 127,0
Neon_On.Color = RGB(255,0,0)
Neon_On.ColorFull = RGB(255,0,0)
Neon_Off.Color = RGB(255,0,0)
Neon_Off.ColorFull = RGB(255,0,0)
for each bulb in aNeons
bulb.Color=RGB(255,0,0)
next
for each bulb in bNeons
bulb.Color=RGB(255,0,0)
next
wbcolor = RGB(255,0,0)
If IntegratedDMDType = 1 then
dmd.color = RGB(255,0,0)
dmd1.color = RGB(255,0,0)
End If
F37c.imageA = "aRed"
F37c.imageB = "aRed"
Prim_Neon_ON.material = "red_tube"
'NeonColor = 2
If prim_reflect = 1 Then
Prim_LockRamp.material = "metal_red"
End If
Prim_CenterRamp.material = "Ramps-Plastic_R"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_RED"
DMD2.imageB = "Plastic_Ramp_2012_RED"
DMD3.imageA = "Plastic_Ramp_2012_RED"
DMD3.imageB = "Plastic_Ramp_2012_RED"
End If
Ramp1038.image = "backwall999_red"
End If

If cneons = 3 Then
dmd.opacity = 100
dmd1.opacity = 100
DOF 120,0
DOF 121,0
DOF 122,1
DOF 123,0
DOF 124,0
DOF 125,0
DOF 126,0
DOF 127,0
Neon_On.Color = RGB(255,0,255)
Neon_On.ColorFull = RGB(255,0,255)
Neon_Off.Color = RGB(255,0,255)
Neon_Off.ColorFull = RGB(255,0,255)
for each bulb in aNeons
bulb.Color=RGB(255,0,255)
next
for each bulb in bNeons
bulb.Color=RGB(255,0,255)
next
wbcolor = RGB(255,0,255)
If IntegratedDMDType = 1 then
dmd.color = RGB(255,0,255)
dmd1.color = RGB(255,0,255)
End If
F37c.imageA = "aPurple"
F37c.imageB = "aPurple"
'NeonColor = 3
If prim_reflect = 1 Then
Prim_LockRamp.material = "metal_purple"
End If
Prim_Neon_ON.material = "purple_tube"
Prim_CenterRamp.material = "Ramps-Plastic_Pu"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_PURPLE"
DMD2.imageB = "Plastic_Ramp_2012_PURPLE"
DMD3.imageA = "Plastic_Ramp_2012_PURPLE"
DMD3.imageB = "Plastic_Ramp_2012_PURPLE"
End If
Ramp1038.image = "backwall999_purple"
End If

If cneons = 4 Then
dmd.opacity = 100
dmd1.opacity = 100
DOF 120,0
DOF 121,0
DOF 122,0
DOF 123,1
DOF 124,0
DOF 125,0
DOF 126,0
DOF 127,0
Neon_On.Color = RGB(0,0,255)
Neon_On.ColorFull = RGB(0,0,255)
Neon_Off.Color = RGB(0,0,255)
Neon_Off.ColorFull = RGB(0,0,255)
for each bulb in aNeons
bulb.Color=RGB(0,0,255)
next
for each bulb in bNeons
bulb.Color=RGB(0,0,255)
next
wbcolor = RGB(0,0,255)
If IntegratedDMDType = 1 then
dmd.color = RGB(0,0,255)
dmd1.color = RGB(0,0,255)
End If
F37c.imageA = "aBlue"
F37c.imageB = "aBlue"
'NeonColor = 4
If prim_reflect = 1 Then
Prim_LockRamp.material = "metal_blue"
End If
Prim_Neon_ON.material = "blue_tube"
Prim_CenterRamp.material = "Ramps-Plastic_B"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_BLUE"
DMD2.imageB = "Plastic_Ramp_2012_BLUE"
DMD3.imageA = "Plastic_Ramp_2012_BLUE"
DMD3.imageB = "Plastic_Ramp_2012_BLUE"
End If
Ramp1038.image = "backwall999_blue"
End If

If cneons = 5 Then
dmd.opacity = 100
dmd1.opacity = 100
DOF 120,0
DOF 121,0
DOF 122,0
DOF 123,0
DOF 124,1
DOF 125,0
DOF 126,0
DOF 127,0
Neon_On.Color = RGB(0,255,255)
Neon_On.ColorFull = RGB(0,255,255)
Neon_Off.Color = RGB(0,255,255)
Neon_Off.ColorFull = RGB(0,255,255)
for each bulb in aNeons
bulb.Color=RGB(0,255,255)
next
for each bulb in bNeons
bulb.Color=RGB(0,255,255)
next
wbcolor = RGB(0,255,255)
If IntegratedDMDType = 1 then
dmd.color = RGB(0,255,255)
dmd1.color = RGB(0,255,255)
End If
F37c.imageA = "aCyan"
F37c.imageB = "aCyan"
'NeonColor = 5
If prim_reflect = 1 Then
Prim_LockRamp.material = "metal_cyan"
End If
Prim_Neon_ON.material = "cyan_tube"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_CYAN"
DMD2.imageB = "Plastic_Ramp_2012_CYAN"
DMD3.imageA = "Plastic_Ramp_2012_CYAN"
DMD3.imageB = "Plastic_Ramp_2012_CYAN"
End If
Ramp1038.image = "backwall999_cyan"
End If

If cneons = 6 Then
dmd.opacity = 100
dmd1.opacity = 100
DOF 120,0
DOF 121,0
DOF 122,0
DOF 123,0
DOF 124,0
DOF 125,1
DOF 126,0
DOF 127,0
Neon_On.Color = RGB(204,153,255)
Neon_On.ColorFull = RGB(204,153,255)
Neon_Off.Color = RGB(204,153,255)
Neon_Off.ColorFull = RGB(204,153,255)
for each bulb in aNeons
bulb.Color=RGB(204,153,255)
next
for each bulb in bNeons
bulb.Color=RGB(204,153,255)
next
wbcolor = RGB(204,153,255)
If IntegratedDMDType = 1 then
dmd.color = RGB(204,153,255)
dmd1.color = RGB(204,153,255)
End If
F37c.imageA = "aPink"
F37c.imageB = "aPink"
Prim_Neon_ON.material = "pink_tube"
'NeonColor = 6
If prim_reflect = 1 Then
Prim_LockRamp.material = "metal_pink"
End If
Prim_CenterRamp.material = "Ramps-Plastic_Pk"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_PINK"
DMD2.imageB = "Plastic_Ramp_2012_PINK"
DMD3.imageA = "Plastic_Ramp_2012_PINK"
DMD3.imageB = "Plastic_Ramp_2012_PINK"
End If
Ramp1038.image = "backwall999_pink"
End If

If cneons = 7 Then
dmd.opacity = 100
dmd1.opacity = 100
DOF 120,0
DOF 121,0
DOF 122,0
DOF 123,0
DOF 124,0
DOF 125,0
DOF 126,1
DOF 127,0
Neon_On.Color = RGB(255,255,0)
Neon_On.ColorFull = RGB(255,255,0)
Neon_Off.Color = RGB(255,255,0)
Neon_Off.ColorFull = RGB(255,255,0)
for each bulb in aNeons
bulb.Color=RGB(255,255,0)
next
for each bulb in bNeons
bulb.Color=RGB(255,255,0)
next
wbcolor = RGB(255,255,0)
If IntegratedDMDType = 1 then
dmd.color = RGB(255,255,0)
dmd1.color = RGB(255,255,0)
End If
F37c.imageA = "aYellow"
F37c.imageB = "aYellow"
'NeonColor = 7
If prim_reflect = 1 Then
Prim_LockRamp.material = "metal_yellow"
End If
Prim_Neon_ON.material = "yellow_tube"
Prim_CenterRamp.material = "Ramps-Plastic_Y"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_YELLOW"
DMD2.imageB = "Plastic_Ramp_2012_YELLOW"
DMD3.imageA = "Plastic_Ramp_2012_YELLOW"
DMD3.imageB = "Plastic_Ramp_2012_YELLOW"
End If
Ramp1038.image = "backwall999_yellow"
End If
End Sub


Dim nxx, non
Sub Neons(enabled)

If enabled Then
DOF 37,1
non = 1
Prim_Neon_Off.visible = 0
Prim_Neon_On.visible = 1

If cneons = 1 AND prim_reflect = 1 Then
Prim_LockRamp.material = "metal_green"
End If
If cneons = 2 AND prim_reflect = 1 Then
Prim_LockRamp.material = "metal_red"
End If
If cneons = 3 AND prim_reflect = 1 Then
Prim_LockRamp.material = "metal_purple"
End If
If cneons = 4 AND prim_reflect = 1 Then
Prim_LockRamp.material = "metal_blue"
End If
If cneons = 5 AND prim_reflect = 1 Then
Prim_LockRamp.material = "metal_cyan"
End If
If cneons = 6 AND prim_reflect = 1 Then
Prim_LockRamp.material = "metal_pink"
End If
If cneons = 7 AND prim_reflect = 1 Then
Prim_LockRamp.material = "metal_yellow"
End If
If cneons = 8 AND prim_reflect = 1 Then
Prim_LockRamp.material = "metal_orange"
End If

If cneons = 1  Then
Prim_CenterRamp.material = "Ramps-Plastic_G"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_GREEN"
DMD2.imageB = "Plastic_Ramp_2012_GREEN"
DMD3.imageA = "Plastic_Ramp_2012_GREEN"
DMD3.imageB = "Plastic_Ramp_2012_GREEN"
End If
End If
If cneons = 2  Then
Prim_CenterRamp.material = "Ramps-Plastic_R"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_RED"
DMD2.imageB = "Plastic_Ramp_2012_RED"
DMD3.imageA = "Plastic_Ramp_2012_RED"
DMD3.imageB = "Plastic_Ramp_2012_RED"
End If
End If
If cneons = 3  Then
Prim_CenterRamp.material = "Ramps-Plastic_Pu"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_PURPLE"
DMD2.imageB = "Plastic_Ramp_2012_PURPLE"
DMD3.imageA = "Plastic_Ramp_2012_PURPLE"
DMD3.imageB = "Plastic_Ramp_2012_PURPLE"
End If
End If
If cneons = 4  Then
Prim_CenterRamp.material = "Ramps-Plastic_B"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_BLUE"
DMD2.imageB = "Plastic_Ramp_2012_BLUE"
DMD3.imageA = "Plastic_Ramp_2012_BLUE"
DMD3.imageB = "Plastic_Ramp_2012_BLUE"
End If
End If
If cneons = 5  Then
Prim_CenterRamp.material = "Ramps-Plastic_C"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_CYAN"
DMD2.imageB = "Plastic_Ramp_2012_CYAN"
DMD3.imageA = "Plastic_Ramp_2012_CYAN"
DMD3.imageB = "Plastic_Ramp_2012_CYAN"
End If
End If
If cneons = 6  Then
Prim_CenterRamp.material = "Ramps-Plastic_Pk"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_PINK"
DMD2.imageB = "Plastic_Ramp_2012_PINK"
DMD3.imageA = "Plastic_Ramp_2012_PINK"
DMD3.imageB = "Plastic_Ramp_2012_PINK"
End If
End If
If cneons = 7  Then
Prim_CenterRamp.material = "Ramps-Plastic_Y"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_YELLOW"
DMD2.imageB = "Plastic_Ramp_2012_YELLOW"
DMD3.imageA = "Plastic_Ramp_2012_YELLOW"
DMD3.imageB = "Plastic_Ramp_2012_YELLOW"
End If
End If
If cneons = 8  Then
Prim_CenterRamp.material = "Ramps-Plastic_O"
If DMD_Plastic = 1 Then
DMD2.imageA = "Plastic_Ramp_2012_ORANGE"
DMD2.imageB = "Plastic_Ramp_2012_ORANGE"
DMD3.imageA = "Plastic_Ramp_2012_ORANGE"
DMD3.imageB = "Plastic_Ramp_2012_ORANGE"
End If
End If

For each nxx in aNeons:nxx.State = 1:Next
For each nxx in bNeons:nxx.State = 1:Next
Neon_Off.State = 0
Neon_On.State = 1
F37F.enabled = 1:f37gstate = 1

'If Dozer_Cab = 1 Then
If cneons = 1 Then
DOF 120,1 ' Green
DOF 121,0
DOF 122,0
DOF 123,0
DOF 124,0
DOF 125,0
DOF 126,0
DOF 127,0
End If
If cneons = 2 Then
DOF 121,1 ' Red
DOF 120,0
DOF 122,0
DOF 123,0
DOF 124,0
DOF 125,0
DOF 126,0
DOF 127,0
End If
If cneons = 3 Then
DOF 122,1 ' Purple
DOF 121,0
DOF 120,0
DOF 123,0
DOF 124,0
DOF 125,0
DOF 126,0
DOF 127,0
End If
If cneons = 4 Then
DOF 123,1 ' Blue
DOF 121,0
DOF 122,0
DOF 120,0
DOF 124,0
DOF 125,0
DOF 126,0
DOF 127,0
End If
If cneons = 5 Then
DOF 124,1 ' Cyan
DOF 121,0
DOF 122,0
DOF 123,0
DOF 120,0
DOF 125,0
DOF 126,0
DOF 127,0
End If
If cneons = 6 Then
DOF 125,1 ' Pink
DOF 121,0
DOF 122,0
DOF 123,0
DOF 124,0
DOF 120,0
DOF 126,0
DOF 127,0
End If
If cneons = 7 Then
DOF 126,1 ' Yellow
DOF 121,0
DOF 122,0
DOF 123,0
DOF 124,0
DOF 125,0
DOF 120,0
DOF 127,0
End If
If cneons = 8 Then
DOF 127,1 ' Orange
DOF 121,0
DOF 122,0
DOF 123,0
DOF 124,0
DOF 125,0
DOF 120,0
DOF 126,0
End If
'End If

Else

DOF 37,0
non = 0

DMD2.imageA = "Plastic_Ramp_2012"
DMD2.imageB = "Plastic_Ramp_2012"
DMD3.imageA = "Plastic_Ramp_2012"
DMD3.imageB = "Plastic_Ramp_2012"

For each nxx in aNeons:nxx.State = 0:Next
For each nxx in bNeons:nxx.State = 0:Next
Prim_CenterRamp.material = "Ramps-Plastic"
Prim_Neon_Off.visible = 1
Prim_Neon_On.visible = 0
Prim_LockRamp.material = "Metal"
Neon_Off.State = 1
Neon_On.State = 0
F37F.enabled = 1:f37gstate = 2

'If Dozer_Cab = 1 Then
If NeonColor = 1 OR cneons = 1 Then
DOF 120,0 ' Green
End If
If NeonColor = 2 OR cneons = 2 Then
DOF 121,0 ' Red
End If
If NeonColor = 3 OR cneons = 3 Then
DOF 122,0 ' Purple
End If
If NeonColor = 4 OR cneons = 4 Then
DOF 123,0 ' Blue
End If
If NeonColor = 5 OR cneons = 5 Then
DOF 124,0 ' Cyan
End If
If NeonColor = 6 OR cneons = 6 Then
DOF 125,0 ' Pink
End If
If NeonColor = 7 or cneons = 7 Then
DOF 126,0 ' Yellow
End If
If NeonColor = 8 or cneons = 8 Then
DOF 127,0 ' Orange
End If
End If
'End If
End Sub

'End Colors


'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 1
Const UseSync = 0
Const HandleMech = 1

' Thal : Added because of useSolenoid=2
' Const cSingleLFlip = 0
' Const cSingleRFlip = 0

' Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SCoin = "fx_coin"

Set GiCallback2 = GetRef("UpdateGI")

Dim bsTrough, bsPopper, bsLeftSaucer, bsRightSaucer, vlLock, mechRM, mRingmasterMagnet, mLockMagnet, mJugglerMagnet, mspinmagnet, x, plungerIM

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps()

DOF 88,L88.State

DOF 22,L22.State
DOF 221,L21.State
DOF 11,L11.State
DOF 13,L13.State
DOF 14,L14.State
DOF 16,L16.State

If L83.State = 1 Then
Boom_Cap.image = "boom_cap_on"
else
Boom_Cap.image = "boom_cap_off"
End If

L84a.State = L84.State
BC2A.state = L84.state

If L84.State = 1 Then
BD2.image = "bumper_on"
else
BD2.image = "bumper_off"
End If

If L82.State = 1 Then
BD1.image = "bumper_on"
else
BD1.image = "bumper_off"
End If

L82a.State = L82.State
BC1A.state = L82.state

If L86.State = 1 Then
Disc1.image = "Disc1D"
else
Disc1.image = "Disc1D_Dark"
End If

If L77.State = 1 Then
Disc2.image = "Disc1D"
else
Disc2.image = "Disc1D_Dark"
End If

If L87.State = 1 Then
Disc3.image = "Disc1D"
else
Disc3.image = "Disc1D_Dark"
End If

If L85.State = 1 Then
Disc4.image = "Disc1D"
else
Disc4.image = "Disc1D_Dark"
End If

End Sub


' -------------------------------------
' Lights Array
' -------------------------------------
' lights, map all lights on the playfield to the Lights array
On Error Resume Next
Dim i
For i=0 To 127
  Execute "Set Lights(" & i & ")  = L" & i
Next

'************
' Table init.
'************

Sub Table1_Init
    vpmInit Me
    With Controller
    .GameName = cGameName
    .SplashInfoLine = "Cirqus Voltaire"
    .Games(cGameName).Settings.Value("rol") = DMD_Rot
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
     If IntegratedDMDType = 1 OR 2 Then
        .Hidden=1
         End If
         If IntegratedDMDType = 0 Then
        .Hidden=0
         End If
        .SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then
      MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description
      msgbox "After table loads, use F6 to choose a different Rom and reload the table."
      Exit Sub
    End If
    On Error Goto 0
    End With

    Sound1 = "ballrelease"
    Sound2 = "fx_Solenoid"
    Sound3 = "popper_ball"
        Sound4 = "KickerEnter"

    ' Nudging
    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(LBumper1, Bumper1, Bumper2, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 32, 33, 34, 35, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX(Sound1), SoundFX(Sound2)
        .InitEntrySnd Sound2, Sound2
        .Balls = 4
        .IsTrough = 1
    End With

   'Left saucer
    Set bsLeftSaucer = New cvpmBallStack
    With bsLeftSaucer
        .InitSaucer sw71, 71, 45, 10
        .InitExitSnd SoundFX(Sound3), SoundFX(Sound3)
        .CreateEvents "bsLeftSaucer", 0
    End With

    'Right saucer
    Set bsRightSaucer = New cvpmBallStack
    With bsRightSaucer
        .InitSaucer sw72, 72, 140, 6
        .InitExitSnd SoundFX(Sound3), SoundFX(Sound3)
        .CreateEvents "bsRightSaucer", 0
    End With

    ' Ball popper
    Set bsPopper = New cvpmBallStack
    With bsPopper
        .InitSw 0, 36, 0, 0, 0, 0, 0, 0
        .InitKick sw36, 290, 20
        .InitExitSnd SoundFX(Sound3), SoundFX(Sound3)
        .KickForceVar = 3
        .KickAngleVar = 15
        .KickBalls = 2
    End With

    ' Ring Master Motor
    Set mechRM = New cvpmMech
    With mechRM
        .MType = vpmMechLinear + vpmMechReverse + vpmMechOneDirSol + vpmMechLengthSw
        .Sol1 = 22
        .Sol2 = 39
        .Length = Ringmaster_Speed  '148
        .Steps = Ringmaster_Speed    ' 148
        .AddSw 44, 0, 2     '0
        .AddSw 43, Ringmaster_Speed * 73 \ 100, Ringmaster_Speed * 75 \ 100'109, 111    '200
        .AddSw 42, Ringmaster_Speed * 98 \ 100, Ringmaster_Speed        '265.5
    'Motor fix later down in the script
    End With

    ' Visible Lock
    Set vlLock = New cvpmVLock
    With vlLock
        .InitVLock Array(sw66, sw67, sw68), Array(sw66k, sw67k, sw68k), Array(66, 67, 68)
        .InitSnd SoundFX(Sound2), SoundFX(Sound2)
        .CreateEvents "vlLock"
    End With

    ' Ring Master Magnet
    Set mRingmasterMagnet = New cvpmMagnet
    With mRingmasterMagnet
        .InitMagnet RingmasterMagnet, 35
        '.Solenoid = 35 'own solenoid sub
        .GrabCenter = 0
        .Size = 100
        .CreateEvents "mRingmasterMagnet"
    End With

 ' Spin Magnet
    Set mSpinMagnet = New cvpmMagnet
    With mSpinMagnet
        .InitMagnet SpinMagnet, 5
        '.Solenoid = 35 'own solenoid sub
        .GrabCenter = 0
        .Size = 100
        .CreateEvents "mSpinMagnet"
    End With

    ' Lock Magnet
    Set mLockMagnet = New cvpmMagnet
    With mLockMagnet
        .InitMagnet LockMagnet, 50
        .Solenoid = 5
        .GrabCenter = 1
        .Size = 300
        .CreateEvents "mLockMagnet"
    End With

    ' Juggler Magnet
    ' Set mJugglerMagnet = New cvpmMagnet
    ' With mJugglerMagnet
    '    .InitMagnet JugglerMagnet, 70
    '    .Solenoid = 3
    '    .GrabCenter = 1
    '    .Size = 300
    '    .CreateEvents "mJugglerMagnet"
    ' End With

    Set mJugglerMagnet = New cvpmMagnet
        mJugglerMagnet.InitMagnet JugglerMagnet,70
        mJugglerMagnet.GrabCenter=True

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Init other dropwalls - animations
    Jackpot.IsDropped = 1
    vpmDoSolCallback 34, 0
    Boom_Down.enabled = 1   'not needed since the rom handles it
    CreateWildBall

    ' Fix Motor
    Plunger1.PullBack
    Controller.Switch(22) = 0
    vpmTimer.AddTimer 100, "StartMotor"

    vpmTimer.AddTimer 10, "GIUpdate"
End Sub

Sub GIUpdate(dummy):updategi 5,0:End Sub

Sub StartMotor(dummy)
    Controller.Switch(22) = 1 ' Coin door closed - fixes dead motor bug
    With mechRM
        .CallBack = GetRef("UpdateRM")
        .Start
    End With
End Sub

If IntegratedDMDType = 1 AND dmdcolor = 0 Then
dmd.opacity = 170
dmd1.opacity = 170
End If

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then Plunger2.Pullback:PlaySoundAt "PlungerPull",Plunger1
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge_left")
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge_right")
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge_forward")
    If keycode = 3 Then Cycle_Neons
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
  If keycode = PlungerKey Then Plunger2.Fire:PlaySoundat "plunger",Plunger2
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub JugglerMagnet_Hit()
MJugglerMagnet.AddBall ActiveBall
End Sub

Sub JugglerMagnet_UnHit()
MJugglerMagnet.RemoveBall ActiveBall
End Sub

'************************************************
'************Slingshots Animation****************
'************************************************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
  PlaySoundAt SoundFX("Left_slingshot"),SLING2
  vpmTimer.PulseSw 51
  LSling.Visible = 0
  LSling1.Visible = 1
  sling1.TransZ = -24
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -12
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


Sub RightSlingShot_Slingshot
  PlaySoundAt SoundFX("Right_Slingshot"),SLING1
  vpmTimer.PulseSw 52
  RSling.Visible = 0
  RSling1.Visible = 1
  sling2.TransZ = -24
  RStep = 0
  RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -12
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

Sub Bumper1_Hit:vpmTimer.PulseSw 53:PlaySoundAtBumperVol SoundFX("fx_bumper1"),Bumper1,1:Me.TimerEnabled = 1:Me.timerinterval = 10:End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 55:PlaySoundAtBumperVol SoundFX("fx_bumper2"),Bumper2,1:Me.TimerEnabled = 1:Me.timerinterval = 10:End Sub

Sub Lbumper1_Hit():vpmTimer.PulseSw 54:PlaySoundAtBumperVol SoundFX("fx_bumper3"),LBumper1,1: Me.TimerEnabled = 1:Me.timerinterval = 10:boom2_shake.enabled = 1:End Sub

Sub Bumper1_timer()
  BR1.Z = BR1.Z + (5 * dirRing1)
  If BR1.Z <= 0 Then dirRing1 = 1
  If BR1.Z >= 30 Then
    dirRing1 = -1
    BR1.Z = 30
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper2_timer()
  BR2.Z = BR2.Z + (5 * dirRing2)
  If BR2.Z <= 0 Then dirRing2 = 1
  If BR2.Z >= 30 Then
    dirRing2 = -1
    BR2.Z = 30
    Me.TimerEnabled = 0
  End If
End Sub

Sub LBumper1_Timer()

  Boom_Ring.Z = Boom_Ring.Z + (5 * dirRing3)
  If Boom_Ring.Z <= 0 Then dirRing3 = 1
  If Boom_Ring.Z >= 35 Then
    dirRing3 = -1
    Boom_Ring.Z = 35
    Me.TimerEnabled = 0
  End If

End Sub

' Drain holes, vuks & saucers
Sub Drain_Hit
    PlaySoundAtVol "drain",Drain,.8
    bsTrough.AddBall Me
End Sub

' Trap Door Holes with animation
Dim aBall, aZpos

Sub TDHole1_Hit
    Set aBall = ActiveBall
    'ClearballID
    PlaySound "fx_kicker_enter",0,1,0,0,0,0,0,.8
    aZpos = 35
    Me.TimerInterval = 2
    Me.TimerEnabled = 1
End Sub

Sub TDHole1_Timer
    aBall.Z = aZpos
    aZpos = aZpos-4
    If aZpos < -30 Then
        Me.TimerEnabled = 0
        Me.DestroyBall
        bsTrapDoor.AddBall Me
    End If
End Sub

'Spinner
Sub sw117_Spin():vpmTimer.PulseSw 117:PlaySound "fx_spinner",0,1,Pan(sw117),0,0,0,0,AudioFade(sw117):End Sub
Sub sw115_Spin():vpmTimer.PulseSw 115:PlaySound "fx_spinner",0,1,Pan(sw115),0,0,0,0,AudioFade(sw115):End Sub

' Rollovers & Ramp Switches
Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundAt "fx_sensor",sw27:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:PlaySoundAt "fx_sensor",sw57:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw48_Hit
'vpmTimer.PulseSw 48
Controller.Switch(48) = 1
pstep = 1
cplas.enabled = 1
PlaySoundAt "fx_sensor",sw48
If ActiveBall.VelY <= -35 Then
ActiveBall.VelY = -30
End If
End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor",sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor",sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAt "fx_sensor",sw15:End Sub
Sub sw15_Unhit:Controller.Switch(15) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAt "fx_sensor",sw23:End Sub
Sub sw23_Unhit:Controller.Switch(23) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:pstep = 1:cplas.enabled = 1:d2dir = 1:d2timer.enabled = 1:PlaySoundAt "fx_sensor",sw26:End Sub
Sub sw26_Unhit:Controller.Switch(26) = 0:d2dir = 2:d2timer.enabled = 1:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:pstep = 1:cplas.enabled = 1:d4dir = 1:d4timer.enabled = 1:PlaySoundAt "fx_sensor",sw17:End Sub
Sub sw17_Unhit:Controller.Switch(17) = 0:d4dir = 2:d4timer.enabled = 1:End Sub

Sub BallLockCollide_Hit
  me.TimerInterval = (2-vlLock.Balls) * ((1.0-(BallVel(ActiveBall)/20)) * 160)
  me.TimerEnabled = 1
End Sub


Sub BallLockCollide_Timer
  PlaySoundAt "metalhit_thin",sw68
  me.TimerEnabled = false
End Sub

Sub sw75_Hit:Controller.Switch(75) = 1:pstep = 1:cplas.enabled = 1:d3dir = 1:d3timer.enabled = 3:PlaySoundAt "fx_sensor",sw75:End Sub
Sub sw75_Unhit:Controller.Switch(75) = 0:d3dir = 2:d3timer.enabled = 1:End Sub

Sub sw76_Hit:Controller.Switch(76) = 1:pstep = 1:cplas.enabled = 1:d1dir = 1:d1timer.enabled = 1:PlaySoundAt "fx_sensor",sw76:End Sub
Sub sw76_Unhit:Controller.Switch(76) = 0:d1dir = 2:d1timer.enabled = 1:End Sub

Dim d1dir
Sub D1timer_Timer()
Select Case d1dir
Case 1:
If Disc1.TransZ <= 0 Then
me.enabled = 0
End If
Disc1.TransZ = Disc1.TransZ - 0.2
Case 2:
If Disc1.TransZ => 2.5 Then
me.enabled = 0
End If
Disc1.TransZ = Disc1.TransZ + 0.2
End Select
End Sub

Dim d2dir
Sub D2timer_Timer()
Select Case d2dir
Case 1:
If Disc2.TransZ <= 0 Then
me.enabled = 0
End If
Disc2.TransZ = Disc2.TransZ - 0.2
Case 2:
If Disc2.TransZ => 2.5 Then
me.enabled = 0
End If
Disc2.TransZ = Disc2.TransZ + 0.2
End Select
End Sub

Dim d3dir
Sub D3timer_Timer()
Select Case d3dir
Case 1:
If Disc3.TransZ <= 0 Then
me.enabled = 0
End If
Disc3.TransZ = Disc3.TransZ - 0.2
Case 2:
If Disc3.TransZ => 2.5 Then
me.enabled = 0
End If
Disc3.TransZ = Disc3.TransZ + 0.2
End Select
End Sub

Dim d4dir
Sub D4timer_Timer()
Select Case d4dir
Case 1:
If Disc4.TransZ <= 0 Then
me.enabled = 0
End If
Disc4.TransZ = Disc4.TransZ - 0.2
Case 2:
If Disc4.TransZ => 2.5 Then
me.enabled = 0
End If
Disc4.TransZ = Disc4.TransZ + 0.2
End Select
End Sub


Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "subway2",sw45:DOF 245, 2:End Sub
Sub sw45_Unhit:Controller.Switch(45) = 0:End Sub

' Ramps Switches

Sub sw64_Hit
    pstep = 1:cplas.enabled = 1
    Controller.Switch(64) = 1
    'MsgBox ActiveBall.VelX
    If ActiveBall.VelX > 20 then ActiveBall.VelX = 15
  Flipper1.rotatetoend
End Sub

Sub sw64_Unhit:Controller.Switch(64) = 0:Flipper1.rotatetostart:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:PlaySoundAt "subway2",sw65:Flipper2.rotatetoend:End Sub
Sub sw65_Unhit:Controller.Switch(65) = 0:Flipper2.rotatetostart:End Sub



Sub sw12_Hit
Controller.Switch(12) = 1
If ActiveBall.VelX <= -15 Then
'PlaySound "fx_metalrolling",0,1,Pan(ActiveBall)
End If
End Sub

Sub sw12_Unhit:Controller.Switch(12) = 0:End Sub


Sub sw16_Hit()
  Set RMBall = ActiveBall
  PlaySoundAt "fx_woodhit",sw16
  If RMCurrPos > Ringmaster_Speed * 72 \ 100  Then
    RMShake
        'If Dozer_Cab = 1 Then
    DOF 269,2
    'End If
  End If
        Controller.Switch(16)=1
        Eddy.enabled = 1
 End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:DOF 218, 1:End Sub
Sub sw18_Unhit:Controller.Switch(18) = 0:DOF 218, 0:End Sub

Sub sw63_Hit
Controller.Switch(63) = 1
If ActiveBall.VelY <= -30 Then
PlaySound "subway2",0,1,Pan(ActiveBall),0,0,0,0,AudioFade(ActiveBall)
End If
End Sub

Sub sw63_Unhit:Controller.Switch(63) = 0:End Sub


Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAt SoundFX("target"),sw41:DOF 241, 2:End Sub
Sub sw61_Hit:vpmTimer.PulseSw 61:PlaySoundAt SoundFX("target"),sw61:DOF 281, 2:End Sub
Sub sw62_Hit:vpmTimer.PulseSw 62:PlaySoundAt SoundFX("target"),sw62:DOF 282, 2:End Sub
Sub sw56_Hit:vpmTimer.PulseSw 56:PlaySoundAt SoundFX("target"),sw56:DOF 284, 2:End Sub
Sub sw58_Hit:vpmTimer.PulseSw 58:PlaySoundAt SoundFX("target"),sw58:DOF 283, 2:End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:PlaySoundAt SoundFX("target"),sw38:DOF 285, 2:End Sub
Sub sw38a_Hit:vpmTimer.PulseSw 38:PlaySoundAt SoundFX("target"),sw38:DOF 286, 2:End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:PlaySoundAt SoundFX("target"),sw37:DOF 287, 2:End Sub
Sub sw37b_Hit:vpmTimer.PulseSw 37:PlaySoundAt SoundFX("target"),sw37:DOF 288, 2:End Sub
Sub sw37d_Hit:vpmTimer.PulseSw 37:PlaySoundAt SoundFX("target"),sw37:DOF 289, 2:End Sub

Sub sw74_Hit:vpmTimer.PulseSw 74:PlaySoundAt SoundFX("target"),Kicker1:DOF 274, 2:End Sub

' Ramps helpers
Sub RHelp1_Hit()
    'StopSound "fx_metalrolling"
    vpmTimer.AddTimer 100, "BallSound1"
End Sub

Sub LRSK_Hit()
    StopSound "fx_metalrolling"
    vpmTimer.AddTimer 100, "BallSound1"
End Sub

Sub RRSK_Hit()
    StopSound "subway2"
    vpmTimer.AddTimer 100, "BallSound2"
End Sub

Sub RHelp2_Hit()
    'StopSound "fx_metalrolling"
    vpmTimer.AddTimer 100, "BallSound2"
End Sub

Sub RLW_Hit()
StopSound "fx_metalrolling"
PlaySoundAt "metalhit_thin",LRSK
End Sub

Sub RRW_Hit()
StopSound "subway2"
PlaySoundAt "metalhit_thin",RRSK
End Sub

Sub TLPOP_Hit()
PlaySoundAt "KickerEnter",TLPOP
End Sub

Sub TRPOP_Hit()
PlaySoundat "KickerEnter",TRPOP
End Sub

'Sub RHelp3_Hit:vpmTimer.AddTimer 150, "BallSound":End Sub
'
'Sub RHelp4_Hit:vpmTimer.AddTimer 150, "BallSound":End Sub

Sub BallSound(dummy):StopBumpSounds:PlaySound "Ball_Bounce",0,1,-.4,0,0,0,0,-.8:End Sub
Sub BallSound1(dummy):StopBumpSounds:PlaySound "Ball_Bounce",0,1,-.4,0,0,0,0,-.8:End Sub
Sub BallSound2(dummy):StopBumpSounds:PlaySound "Ball_Bounce",0,1,.4,0,0,0,0,-.8:End Sub
Sub BallSound3(dummy):StopBumpSounds:PlaySound "Ball_Bounce",0,1,.2,0,0,0,0,.8:End Sub
Sub BallSound4(dummy):StopBumpSounds:PlaySound "Ball_Bounce",0,1,.3,0,0,0,0,-.6:End Sub



'*********
'Solenoids
'*********

SolCallback(1) = "Auto_Plunger"
SolCallBack(2) = "SolCannon"
SolCallback(3) = "LLoop"
SolCallback(4) = "Middle_Jet"
SolCallback(5) = "Ramp_Magnet"
SolCallBack(7) = "BoomUp"
SolCallBack(8) = "BoomDown"
SolCallback(9) = "SolRelease"
SolCallback(10) = "XLSling"
SolCallback(11) = "XRSling"
SolCallback(12) = "Upper_Jet"
SolCallback(13) = "Lower_Jet"
SolCallBack(14) = "Sol14"
SolCallBack(15) = "Sol15"
'SolCallback(16) = "vlLock.SolExit"
SolCallback(16) = "Sol16"
SolCallBack(33) = "Sol33"
'SolCallback(34) = "vpmSolToggleWall LockDiverterOff,LockDiverterOn,""FX_Solenoid"","
SolCallback(34) = "Sol34"
SolCallBack(35) = "SolRingmasterMagnet"
SolCallBack(36) = "Sol36"

'Flashers
SolCallBack(17) = "Sol17"
SolCallBack(18) = "Sol18"
SolCallBack(19) = "Sol19"
SolCallBack(20) = "Sol20"
SolCallBack(21) = "Sol21"
SolCallback(22) = "RM_Motor_E"
SolCallBack(23) = "Sol23"
SolCallBack(24) = "Sol24"
SolCallBack(25) = "Sol25"
SolCallBack(26) = "Sol26"
SolCallBack(27) = "RMFlasher"
SolCallBack(28) = "Sol28"
SolCallBack(37) = "Neons"
SolCallback(39) = "RM_Motor"

'********************
' Special JP Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub LLoop(enabled)
If enabled Then
mJugglerMagnet.Magneton = 1
DOF 64,2
else
mJugglerMagnet.Magneton = 0
End If
End Sub

Sub Ramp_Magnet(enabled)
If enabled Then
DOF 64,2
    'on error goto 0
    StopBumpSounds
End If

End Sub


Sub RM_Motor(enabled)
If enabled Then
'PlaySound "bridge"
DOF 268,1
Else
'StopSound "bridge"
DOF 268,0
End If
End Sub

Sub RM_Motor_E(enabled)
If enabled Then
'PlaySound "bridge"
DOF 268,1
Else
'StopSound "bridge"
DOF 268,0
End If
End Sub

Sub Middle_Jet(enabled)
If enabled Then
DOF 265,1
Else
DOF 265,0
End If
End Sub

Sub Upper_Jet(enabled)
If enabled Then
DOF 266,1
Else
DOF 266,0
End If
End Sub

Sub Lower_Jet(enabled)
If enabled Then
DOF 267,1
Else
DOF 267,0
End If
End Sub

'===

Sub XLSling(enabled)
If enabled Then
DOF 263,1
Else
DOF 263,0
End If
End Sub

Sub XRSling(enabled)
If enabled Then
DOF 264,1
Else
DOF 264,0
End If
End Sub

Sub Sol14(enabled)
If enabled Then
bsLeftSaucer.ExitSol_On
DOF 259,2
Else
End If
End Sub

Sub Sol15(enabled)
If enabled Then
bsRightSaucer.ExitSol_On
DOF 260,2
Else
End If
End Sub

Sub Sol33(enabled)
If enabled Then
bsPopper.ExitSol_On
DOF 261,2
End If
End Sub

Sub Sol36(enabled)
If enabled Then
DOF 262,1
UpperPost.isdropped = 0
Else
DOF 262,0
UpperPost.isdropped = 1
End If
End Sub

Sub Sol16(enabled)
If enabled then
vlLock.SolExit true
'LockPost.Isdropped=true
lpdir = 1:lp.enabled = 1
DOF 256,1
else
vlLock.SolExit False
lpdir = 2:lp.enabled = 1
DOF 256,0
End If
End Sub

Dim lpdir
Sub LP_Timer()
Select Case lpdir
Case 1: ' Down
If LockPostP.Z <= 0 Then
me.enabled = 0
End If
LockPostP.Z = LockPostP.Z - 2
Case 2: ' Up
If LockPostP.Z => 55 Then
me.enabled = 0
LockPostP.Z = 55
End If
LockPostP.Z = LockPostP.Z + 2
End Select
End Sub

Sub Sol34(enabled)
If enabled Then
LockDiverterOn.isdropped = 0
LockDiverterOff.isdropped = 1
DBD.isdropped = 0
DOF 258,1
else
LockDiverterOn.isdropped = 1
LockDiverterOff.isdropped = 0
DBD.isdropped = 1
DOF 258,0
End If
End Sub

Sub Sol17(enabled)
If enabled Then
DOF 17, 1
F17.State = 1:F17a.State = 1
Else
DOF 17, 0
F17.State = 0:F17a.State = 0
End If
End Sub

Sub Sol18(enabled)
If enabled Then
DOF 18, 1
F18.State = 1:F18a.State = 1
Else
DOF 18, 0
F18.State = 0:F18a.State = 0
End If
End Sub

Sub Sol19(enabled)
If enabled Then
DOF 19, 1
F19.State = 1:F19a.State = 1
Else
DOF 19, 0
F19.State = 0:F19a.State = 0
End If
End Sub

Sub Sol20(enabled)
If enabled Then
DOF 20, 1
F20.State = 1:F20a.State = 1
Else
DOF 20, 0
F20.State = 0:F20a.State = 0
End If
End Sub

Sub Sol21(enabled)
If enabled Then
Prim_NeonRamp.material = "Ramps-Plastic_Y"
DOF 21,1
'for each bulb in aNeons
'bulb.Color=RGB(255,255,0)
'next
'For each nxx in aNeons:nxx.State = 1:Next
F21.State = 1:F21a.State = 1:F21b.State = 1:F21a1.state = 1':F21R.State = 1
If masterf = 0 Then
Ringmaster.image = "rm_off_ryel"
else
DOF 21,0
Ringmaster.image = "rm_on_ryel"
End If
F21F.enabled = 1:f21gstate = 1
If ShowDT = True Then
'BGF21.visible = 1
End If

Else
DOF 21,0
'For each nxx in aNeons:nxx.State = 0:Next
Prim_NeonRamp.material = "Ramps-Plastic"
F21.State = 0:F21a.State = 0:F21b.State = 0:F21a1.state = 0':F21R.State = 0
'If masterf = 0 Then
Ringmaster.image = "rm_off"
'End If
F21F.enabled = 1:f21gstate = 2
If ShowDT = True Then
'BGF21.visible = 0
End If
End If
End Sub

Sub Sol23(enabled)
If enabled Then
DOF 23,1
F23.State = 1:F23a.State = 1
Prim_NeonRamp.material = "Ramps-Plastic_R"
Else
DOF 23,0
F23.State = 0:F23a.State = 0
Prim_NeonRamp.material = "Ramps-Plastic"
End If
End Sub

Dim xxtf,rms
Sub Sol24(enabled)
If enabled Then
DOF 24,1
Prim_CenterRamp.material = "Ramps-Plastic_Y"
For each xxtf in F24ers:xxtf.State = 1:next
Prim_LeftWireRamp.material = "metal_yellow"
Prim_MenagerieGuide.material = "metal_yellow"
If masterf = 0 Then
ringmaster.image = "rm_off_lyel"
else
ringmaster.image = "rm_on_lyel"
End If
F24F.enabled = 1:f24gstate = 1
'If ramp_flash = 1 Then
'For each rms in Ramp_Tops:rms.image = "Aluminumy1":next
'End If
Else
DOF 24,0
For each xxtf in F24ers:xxtf.State = 0:next
If Prim_Reflect = 1 Then
Prim_LeftWireRamp.material = "Metal"
Prim_MenagerieGuide.material = "Metal"
End If
'If masterf = 0 Then
ringmaster.image = "rm_off"
'End If
F24F.enabled = 1:f24gstate = 2
'If ramp_flash = 1 Then
'For each rms in Ramp_Tops:rms.image = "Metal3":next
'End If
Prim_CenterRamp.material = "Ramps-Plastic"
End If
End Sub

Dim xxtfi
Sub Sol25(enabled)
If enabled Then
DMD2.imageA = "Plastic_Ramp_2012_LY"
DMD2.imageB = "Plastic_Ramp_2012_LY"
DOF 25,1
For each xxtfi in F25ers:xxtfi.State = 1:next
F25F.enabled = 1:f25gstate = 1
Else
DMD2.imageA = "Plastic_Ramp_2012"
DMD2.imageB = "Plastic_Ramp_2012"
DOF 25,0
For each xxtfi in F25ers:xxtfi.State = 0:next
F25F.enabled = 1:f25gstate = 2
End If
End Sub

Dim xxts
Sub Sol26(enabled)
If enabled Then
DMD2.imageA = "Plastic_Ramp_2012_RY"
DMD2.imageB = "Plastic_Ramp_2012_RY"
DOF 26,1
For each xxte in F26ers:xxte.State = 1:next
Else
DMD2.imageA = "Plastic_Ramp_2012"
DMD2.imageB = "Plastic_Ramp_2012"
DOF 26,0
For each xxte in F26ers:xxte.State = 0:next
End If
End Sub

Dim xxte
Sub Sol28(enabled)
If enabled Then
DOF 28,1
For each xxte in F28ers:xxte.State = 1:next
Else
DOF 28,0
For each xxte in F28ers:xxte.State = 0:next
End If
End Sub

'******************************************
' Use FlipperTimers to call div subs
'******************************************

Dim LFTCount:LFTCount=1
Dim RFTCount:RFTCount=1

Sub SolLFlipper(Enabled)

     If Enabled Then
     PlaySoundAtVol SoundFX("fx_flipperup"),LeftFlipper,VolFlip
     LeftFlipper.RotateToEnd
         DOF 250,1
     Else
     LFTCount=1
     PlaySoundAtVol SoundFX("fx_flipperdown"),LeftFlipper,VolFlip
     LeftFlipper.RotateToStart
         DOF 250,0
     End If
 End Sub

Sub SolRFlipper(Enabled)

     If Enabled Then
     PlaySoundAtVol SoundFX("fx_flipperup"),RightFlipper,VolFlip
     RightFlipper.RotateToEnd
         DOF 251,1
     Else
     RFTCount=1
     PlaySoundAtVol SoundFX("fx_flipperdown"),RightFlipper,VolFlip
     RightFlipper.RotateToStart
         DOF 251,0
     End If
 End Sub


'**************
' Solenoid Subs
'**************

Sub SolRelease(Enabled)
    If Enabled And bsTrough.Balls > 0 Then
        vpmTimer.PulseSw 31
        bsTrough.ExitSol_On
        DOF 252,2
    End If
End Sub

Sub Auto_Plunger(Enabled)
    If Enabled Then
        Plunger1.Fire
        PlaySoundAt SoundFX("fx_Solenoid"),Plunger1
        DOF 253,2
Else
        Plunger1.PullBack
    End If
End Sub



Sub SolCannon(Enabled)
If enabled Then
PlaySound SoundFX("fx_Solenoid"),0,2,-.1,0,0,1,0,-1
bgcan = 1:dtcan.enabled = 1
DOF 45,1
DOF 46,1
DOF 254,2
'Light93.State = 1
vpmTimer.AddTimer 400, "EndCannon"
vpmTimer.AddTimer 150, "EndCannonDOF"
End If
End Sub

Sub EndCannon(dummy)
vpmTimer.PulseSw 11
PlaySound "fx_Bell10",0,2,.2,0,0,1,0,-1
End Sub

Sub EndCannonDOF(dummy)
DOF 45,0
DOF 46,0
'Light93.State = 0
End Sub

' Dim FireCannon,CannonFlag
' CannonFlag=0
' Sub SolCannon(Enabled)
'   If CannonFlag=0 Then
'     If Enabled Then
'         PlaySound "fx_Solenoid"
'         FireCannon=FireCannon+1:If FireCannon>255 Then FireCannon=0
'         Cannon.Enabled=1
'         CannonFlag=1
'         DOF 254,2
'         DOF 15,1
'         DOF 16,1
'         Light93.State = 1
'    bgcan = 1:dtcan.enabled = 1
'     End If
'   End If
' End Sub

'Sub Cannon_Timer()
'DOF 15,0
'DOF 16,0
'Light93.State =0
'     CannonFlag=CannonFlag+1
'     Select Case CannonFlag
'        Case 2
'           vpmTimer.PulseSw 11
'           PlaySound "fx_Bell10"
'       Case 3
'           Cannon.Enabled=0
'           CannonFlag=0
'     End Select
'End Sub

Sub BoomUp(Enabled)
    If Enabled Then
    PlaySoundAt SoundFX("fx_solenoid"),Boom_Body
        Boom_Up.enabled = 1
    BoomTrigger.enabled = 1
        DOF 255,2
    End If
End Sub

Sub BoomDown(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_solenoid"),Boom_Body
    BoomTrigger.enabled = 0
        LBumper1.Collidable = 0
        Boom_Down.enabled = 1
        DOF 255,2
    End If
End Sub

Sub BoomTrigger_Hit
  activeball.z = activeball.z + 70
End Sub

Dim pstep, prun
pstep = 1

Sub cplas_timer()
prun = prun + 1
If prun = 2500 Then
cstub.enabled = 1
prun = 0
me.enabled = 0
End If
Select Case pstep
Case 1:DiscP.image = "plasma1":Light118.state = 1:pstep = 2
Case 2:DiscP.image = "plasma2":Light118.state = 0:pstep = 3
Case 3:DiscP.image = "plasma3":Light118.state = 1:pstep = 4
Case 4:DiscP.image = "plasma4":Light118.state = 0:pstep = 5
Case 5:DiscP.image = "plasma5":Light118.state = 1:pstep = 6
Case 6:DiscP.image = "plasma6":Light118.state = 0:pstep = 7
Case 7:DiscP.image = "plasma7":Light118.state = 1:pstep = 8
Case 8:DiscP.image = "plasma8":Light118.state = 0:pstep = 9
Case 9:DiscP.image = "plasma9":Light118.state = 1:pstep = 10
Case 10:DiscP.image = "plasma10":Light118.state = 0:pstep = 1
End Select
End Sub

Sub cstub_timer()
DiscP.image = "plasma_off_dark"
Light118.State = 0
me.enabled = 0
End Sub

' *********************
' Ring Master subs
' *********************

 Dim KickAngle
 Sub SolRingmasterMagnet(Enabled)
  mRingmasterMagnet.MagnetOn = Enabled
    DOF 270,Enabled
  RMMagnetkicker.Enabled = Enabled
    If Not Enabled And RMBallInMagnet Then
    mRingmasterMagnet.RemoveBall RMMagBall
    KickAngle = 135 + Rnd * 180
    cball.vely = cball.vely + dcos(KickAngle)*2
    cball.velx = cball.velx - dsin(KickAngle)*2
        PlaySound "fx_spring", 0, (cball.vely + cball.velx)/4, 0, 0.15 'Hanibal Spring Sound
    vpmTimer.AddTimer 400, "RMKick"
    vpmTimer.AddTimer 200, "BallSound3"
        DOF 270,0
    End If
 End Sub

 Sub RMMagnetkicker_Hit
    Set RMMagBall = ActiveBall
    RMBallInMagnet = 1
    'DOF 270,1
 End Sub

 Sub RMKick(dummy)
  RMMagnetkicker.kick KickAngle, rnd * 7 + 7
  RMBallInMagnet = 0
  RMMagBall = Empty
 End Sub

Dim RMBallInMagnet, RMMagBall, RMBall, RMFlashOn, RMCurrPos
RMBallInMagnet = 0:RMFlashOn = 0:RMCurrPos = 0
RMMagBall = Empty

 Sub sw38r_Hit
  Set RMBall = ActiveBall
  PlaySoundAt "fx_woodhit",sw38
  vpmTimer.PulseSw 38
  If RMCurrPos > Ringmaster_Speed * 72 \ 100  Then
    RMShake
  End If
 End Sub

 Sub UpdateRM(aCurrPos, aSpeed, aLastPos)
    If Ringmaster_Motor = 1 Then
'    PlaySound SoundFX("bridge2"), , 0.02
    PlaySound SoundFX("bridge2"),0,1,.1,0,0,0,0,-.9
    End If
    RMCurrPos = aCurrPos
  Ringmaster_Down.visible = NOT cbool(RMCurrPos>4)
  Ringmaster_Down2.visible = NOT cbool(RMCurrPos>4)
  Ringmaster.z = (RMCurrPos-4) * 265.5/(Ringmaster_Speed-4)- 265.5
    JackpotP.z = (RMCurrPos-4) * 265.5/(Ringmaster_Speed-4)- 265.5
  RMHitWall.IsDropped = NOT RMCurrPos > 4
     sw16.enabled = RMCurrPos > 4
  If (RMCurrPos >37 And RMCurrPos < 40)  AND RMCurrPos > aLastPos Then cball.vely = cball.vely - 1.5
  If RMBallInMagnet Then RMMagBall.z = Ringmaster.z + 50 + 265.5
    If aCurrPos >= Ringmaster_Speed * 97 \ 100 Then
        vpmSolToggleObj cRMHoles, Nothing, 0, 1
        RM_Make.enabled = 1
        RMHitWall.IsDropped = 1
        If RMBallInMagnet = 0 Then
        mSpinMagnet.MagnetOn = True
        End If
    sw16.enabled = 0
        Jackpot.Isdropped = 0
  Else
        vpmSolToggleObj cRMHoles, Nothing, 0, 0
        Jackpot.Isdropped = 1
        mSpinMagnet.MagnetOn = False
        RM_Make.enabled = 0
    End If
 End Sub

'*************  RM Shake Scripting  *************
 Dim mMagnet, cBall, pMod, rmmod

 Set mMagnet = new cvpmMagnet
 With mMagnet
  .InitMagnet WobbleMagnet, 1.5
  .Size = 100
  .CreateEvents mMagnet
  .MagnetOn = True
 End With
 WobbleInit

 Sub RMShake
  cball.velx = cball.velx + rmball.velx*pMod
  cball.vely = cball.vely + rmball.vely*pMod
 End Sub

Sub RMShake2
  cball.velx = cball.velx + activeball.velx*.05
  cball.vely = cball.vely + activeball.vely*.05
 End Sub

Sub RM_Make_Hit()
RMShake2
End Sub

'Includes stripped down version of my reverse slope scripting for a single ball
 Dim ngrav, ngravmod, pslope, nslope, slopemod
 Sub WobbleInit
  pslope = Table1.SlopeMin +((Table1.SlopeMax - Table1.SlopeMin) * Table1.GlobalDifficulty)
  nslope = pslope
  slopemod = pslope + nslope
  ngravmod = 60/aWobbleTimer.interval
  ngrav = slopemod * .0905 * Table1.Gravity / ngravmod
  pMod = .15          'percentage of hit power transfered to captive wobble ball
  Set cBall = ckicker.createball:cball.image = "blank":ckicker.Kick 0,0:mMagnet.addball cball
  aWobbleTimer.enabled = 1
 End Sub

 Sub aWobbleTimer_Timer
  BallShake.Enabled = RMBallInMagnet
  cBall.Vely = cBall.VelY-ngrav         'modifier for slope reversal/cancellation
  rmmod = (ringmaster.z+265.5)/265*.4       '.4 is a 40% modifier for ratio of ball movement to head movement
    If RMCurrPos >3 Then
    PlaySound "fx_spring", 0, (((ringmaster.rotx * 2)  + (ringmaster.roty * 2))/400), (ringmaster.rotx / 1000), ((ringmaster.rotx + ringmaster.roty)/2000) 'Hanibal Spring Sound
    End If
  ringmaster.rotx = (ckicker.y - cball.y)*rmmod
  ringmaster.roty = (cball.x - ckicker.x)*rmmod
 End Sub

 Sub BallShake_Timer
  If Not IsEmpty(RMMagBall) Then
    RMMagBall.y = RMMagnetkicker.y - dsin(ringmaster.rotx)*265.5
    RMMagBall.x = RMMagnetkicker.x + dsin(ringmaster.roty)*265.5
  End If
 End Sub

'*************  End Shake Scripting  ****************

' **************
' Subway Handler
' **************

' Side Show holes
Sub cSSHoles_Hit(idx):SubwayHandler Me(idx), 46:End Sub

' Ringmaster holes
Sub cRMHoles_Hit(idx):SubwayHandler Me(idx), 47:End Sub

Sub SubwayHandler(aKick, aSwNo)
    'ClearballID
    aKick.Destroyball:PlaySoundAt "RM_Lip",RM_Hole:PlaySound "subway2",0,1,Pan(ActiveBall),0,0,0,0,AudioFade(ActiveBall)
    vpmTimer.PulseSwitch aSwNo, 2000, "bsPopper.AddBall 0'"
End Sub

Sub Rolls_Hit(idx)
PlaySound "Rollover",0,1,Pan(ActiveBall),0,0,0,0,AudioFade(ActiveBall)
End Sub

' ********************************
'    Menagerie Ball (Wild Ball)
' ********************************

Dim WBall

Sub CreateWildBall()

Set WBall = kicker1.Createsizedball(51):WBall.color = (wbcolor):Wball.image = "powerball4":Wball.id = 666:kicker1.Kick 0, 0
WBall.mass = 1.5
End Sub

'***********
' Update GI
'***********
Dim xx
Dim gistep
gistep = 1/8

Sub UpdateGI(no, step)
Controller.Switch(22) = 1 'fix motor
If step = 0 OR step = 7 then exit sub

    Select Case no

        Case 0 'right

For each xx in GIRight:xx.IntensityScale = gistep * step:next

If Dozer_Cab = 1 Then
If step = 1 then B2SController.B2SSetData 111,0:B2SController.B2SSetData 112,0:B2SController.B2SSetData 113,0:B2SController.B2SSetData 114,0:B2SController.B2SSetData 115,0:B2SController.B2SSetData 116,0
If step = 2 then B2SController.B2SSetData 111,1:B2SController.B2SSetData 112,0:B2SController.B2SSetData 113,0:B2SController.B2SSetData 114,0:B2SController.B2SSetData 115,0:B2SController.B2SSetData 116,0
If step = 3 then B2SController.B2SSetData 112,1:B2SController.B2SSetData 111,0:B2SController.B2SSetData 113,0:B2SController.B2SSetData 114,0:B2SController.B2SSetData 115,0:B2SController.B2SSetData 116,0
If step = 4 then B2SController.B2SSetData 113,1:B2SController.B2SSetData 112,0:B2SController.B2SSetData 111,0:B2SController.B2SSetData 114,0:B2SController.B2SSetData 115,0:B2SController.B2SSetData 116,0
If step = 5 then B2SController.B2SSetData 114,1:B2SController.B2SSetData 112,0:B2SController.B2SSetData 113,0:B2SController.B2SSetData 111,0:B2SController.B2SSetData 115,0:B2SController.B2SSetData 116,0
If step = 6 then B2SController.B2SSetData 115,1:B2SController.B2SSetData 112,0:B2SController.B2SSetData 113,0:B2SController.B2SSetData 114,0:B2SController.B2SSetData 111,0:B2SController.B2SSetData 116,0
If step = 8 then B2SController.B2SSetData 116,1:B2SController.B2SSetData 112,0:B2SController.B2SSetData 113,0:B2SController.B2SSetData 114,0:B2SController.B2SSetData 115,0:B2SController.B2SSetData 111,0
End If
        Case 1 'middle

For each xx in GIMiddle:xx.IntensityScale = gistep * step:next

If step = 8 Then
For each xx in AllGI:xx.State = 1:Next
End If

If step > 4 Then
DOF 51,1
DOF 52,0
End If

If step < 4 Then
DOF 51,0
DOF 52,1
End If

        Case 2 'left

For each xx in GILeft:xx.IntensityScale = gistep * step:next

If Dozer_Cab = 1 Then
If step = 1 then B2SController.B2SSetData 101,0:B2SController.B2SSetData 102,0:B2SController.B2SSetData 103,0:B2SController.B2SSetData 104,0:B2SController.B2SSetData 105,0:B2SController.B2SSetData 106,0
If step = 2 then B2SController.B2SSetData 101,1:B2SController.B2SSetData 102,0:B2SController.B2SSetData 103,0:B2SController.B2SSetData 104,0:B2SController.B2SSetData 105,0:B2SController.B2SSetData 106,0
If step = 3 then B2SController.B2SSetData 102,1:B2SController.B2SSetData 101,0:B2SController.B2SSetData 103,0:B2SController.B2SSetData 104,0:B2SController.B2SSetData 105,0:B2SController.B2SSetData 106,0
If step = 4 then B2SController.B2SSetData 103,1:B2SController.B2SSetData 102,0:B2SController.B2SSetData 101,0:B2SController.B2SSetData 104,0:B2SController.B2SSetData 105,0:B2SController.B2SSetData 106,0
If step = 5 then B2SController.B2SSetData 104,1:B2SController.B2SSetData 102,0:B2SController.B2SSetData 103,0:B2SController.B2SSetData 101,0:B2SController.B2SSetData 105,0:B2SController.B2SSetData 106,0
If step = 6 then B2SController.B2SSetData 105,1:B2SController.B2SSetData 102,0:B2SController.B2SSetData 103,0:B2SController.B2SSetData 104,0:B2SController.B2SSetData 101,0:B2SController.B2SSetData 106,0
If step = 8 then B2SController.B2SSetData 106,1:B2SController.B2SSetData 102,0:B2SController.B2SSetData 103,0:B2SController.B2SSetData 104,0:B2SController.B2SSetData 105,0:B2SController.B2SSetData 101,0
End If

     End Select
End Sub

'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
    UpdateMechs
End Sub

 Sub UpdateMechs
  sw64p.rotY=Flipper1.currentangle
  sw65p.rotY=Flipper2.currentangle
  LogoR.rotz = RightFlipper.CurrentAngle
  LogoL.rotz = LeftFlipper.CurrentAngle
    LeftFlipperP.roty = LeftFlipper.CurrentAngle
    RightFlipperP.roty = RightFlipper.CurrentAngle
    If Boom_Cap.Z <= -3 Then
    Boom_Ring.Z = -35
    End If
 End Sub

Sub Dampen(dt,df,r)           'dt is threshold speed, df is dampen factor 0 to 1 (higher more dampening), r is randomness
  Dim dfRandomness
  r=cint(r)
  dfRandomness=INT(RND*(2*r+1))
  df=df+(r-dfRandomness)*.01
  If ABS(activeball.velx) > dt Then activeball.velx=activeball.velx*(1-df*(ABS(activeball.velx)/100))
  If ABS(activeball.vely) > dt Then activeball.vely=activeball.vely*(1-df*(ABS(activeball.vely)/100))
End Sub

Sub DampenXY (dtx,dfx,rx, dty, dfy, ry)  'dt is threshold speed, df is dampen factor 0 to 1 (higher more dampening), r is randomness
  Dim dfxRandomness
  Dim dfyRandomness
  rx=cint(rx)
  ry=cint(ry)
  dfxRandomness=INT(RND*(2*rx+1))
  dfyRandomness=INT(RND*(2*ry+1))
  dfx=dfx+(rx-dfxRandomness)*.01
  dfy=dfy+(ry-dfyRandomness)*.01
  If ABS(activeball.velx) > dtx Then activeball.velx=activeball.velx*(1-dfx*(ABS(activeball.velx)/100))
  If ABS(activeball.vely) > dty Then activeball.vely=activeball.vely*(1-dfy*(ABS(activeball.vely)/100))
End Sub

Sub FlipperDampener_Hit ()
  Dim YThreshold:YThreshold=-10 'Ball speed threshold to activate routine (-Y veloctiy for up table)
  Dim Level:Level=4       'Choose 1-10, 0 for off
  Dim XFactor:XFactor=2     'Straightening factor on X axis (higher value for straighter shots)
  Dim BallTriggerAngle, AngleMultiplier 'Angle calculated to apply less straightening on more lateral shots (further away from center line yields lower SIN result - used as coefficient)
  If ActiveBall.VelY < YThreshold AND Level <> 0 Then
    GetAngle ActiveBall.VelX, ActiveBall.VelY, BallTriggerAngle
    AngleMultiplier=Round(ABS(sin(BallTriggerAngle)),3)
    ActiveBall.VelY = ActiveBall.VelY*((1-(ABS(ActiveBall.VelY)/(200*10/Level-10*Level))))
    ActiveBall.VelX = ActiveBall.VelX*((1-XFactor*AngleMultiplier*(ABS(ActiveBall.VelY)/(200*10/Level-10*Level))))
  End If
End Sub

'=================================================
' <<<<<<<< GetAngle(X, Y, Anglename) >>>>>>>>
'=================================================
Dim Xin,Yin,rAngle,Radit,wAngle,Pi
Pi = csng(4*Atn(1))         '3.1415926535897932384626433832795

 Sub GetAngle(Xin, Yin, wAngle)
  If Sgn(Xin) = 0 Then
    If Sgn(Yin) = 1 Then rAngle = 3 * Pi/2 Else rAngle = Pi/2
    If Sgn(Yin) = 0 Then rAngle = 0
  Else
    rAngle = atn(-Yin/Xin)
  End If
  If sgn(Xin) = -1 Then Radit = Pi Else Radit = 0
  If sgn(Xin) = 1 and sgn(Yin) = 1 Then Radit = 2 * Pi
  wAngle = round((Radit + rAngle),4)
 End Sub

'** Extra math to make my life easier **
Function dCos(degrees)
  Dim Pi:Pi = CSng(4*Atn(1))
  dcos = cos(degrees * Pi/180)
  if ABS(dCos) < 0.000001 Then dCos = 0
  if ABS(dCos) > 0.999999 Then dCos = 1 * sgn(dCos)
End Function

Function dSin(degrees)
  Dim Pi:Pi = CSng(4*Atn(1))
  dsin = sin(degrees * Pi/180)
  if ABS(dSin) < 0.000001 Then dSin = 0
  if ABS(dSin) > 0.999999 Then dSin = 1 * sgn(dSin)
End Function

Function dAtn(x)
  Dim Pi:Pi = CSng(4*Atn(1))
  datn = atn(x) * 180 / Pi
End Function

Function dAtn2(X, Y)
  If X > 0 Then
    dAtn2 = dAtn(Y / X)
  ElseIf X < 0 Then
    dAtn2 = dAtn(Y / X) + 180 * Sgn(Y)
    If Y = 0 Then dAtn2 = dAtn2 + 180
    If Y < 0 Then dAtn2 = dAtn2 + 360
  Else
    dAtn2 = 90 * Sgn(Y)
  End If
  dAtn2 = dAtn2+90
End Function
'** End Extra math **

Sub Eddy_timer()
Controller.Switch(16) = 0
me.enabled = 0
End Sub

Dim masterf
Sub RMFlasher(enabled)
If enabled Then
RingMaster.Image = "rm_on"
f27.State = 1
f27a.State = 1
f27a1.State = 1
masterf = 1
DOF 27,1
Else
RingMaster.Image = "rm_off"
f27.State = 0
f27a.State = 0
f27a1.State = 0
masterf = 0
DOF 27,0
End If
End Sub

Dim rmfup

Sub rm_flash_up_timer()
Select Case rmfup
Case 1:Ringmaster.image = "rm_on_10":rmfup = 2
Case 2:Ringmaster.image = "rm_on_10":rmfup = 3
Case 3:Ringmaster.image = "rm_on_20":rmfup = 4
Case 4:Ringmaster.image = "rm_on_30":rmfup = 5
Case 5:Ringmaster.image = "rm_on_40":rmfup = 6
Case 6:Ringmaster.image = "rm_on_50":rmfup = 7
Case 7:Ringmaster.image = "rm_on_60":rmfup = 8
Case 8:Ringmaster.image = "rm_on_70":rmfup = 9
Case 9:Ringmaster.image = "rm_on_80":rmfup = 10
Case 10:Ringmaster.image = "rm_on_90":rmfup = 11
Case 11:Ringmaster.image = "rm_on":rmfup = 1
End Select
End Sub

Sub boom_up_timer()
If Boom_Cap.Z => 66 Then
me.enabled = 0
Boom_Shake.enabled = 1
LBumper1.Collidable = 1
Boom_PostColl.IsDropped = 0
End If
Boom_Cap.Z = Boom_Cap.Z + 1
Boom_Body.Z = Boom_Body.Z + 1
Boom_Bracket.Z = Boom_Bracket.Z + 1
Boom_Post.Z = Boom_Post.Z + 1
Boom_Ring.Z = Boom_Ring.Z + 1
Boom_Socket.Z = Boom_Socket.Z + 1
Boom_Rubber.Z = Boom_Rubber.Z + 1
End Sub

Sub boom_down_timer()
If Boom_Cap.Z <= -3 Then
me.enabled = 0
'LBumper1.Collidable = 0
Boom_PostColl.IsDropped = 1
End If
Boom_Cap.Z = Boom_Cap.Z - 1
Boom_Body.Z = Boom_Body.Z - 1
Boom_Bracket.Z = Boom_Bracket.Z - 1
Boom_Post.Z = Boom_Post.Z - 1
Boom_Ring.Z = Boom_Ring.Z - 1
Boom_Socket.Z = Boom_Socket.Z - 1
Boom_Rubber.Z = Boom_Rubber.Z - 1
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 200)
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

function AudioFade(ball)
    Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
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

Const tnob = 15 ' total number of balls
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
    dim IsRMMag: IsRMMag = False
    if not IsEmpty(RMMagBall) then if BOT(b) is RMMagBall then IsRMMag = true

        If BallVel(BOT(b) ) > 1 and not BOT(b) is cball and not IsRMMag Then
      rolling(b) = True
      if BOT(b).z < 30 Then ' Ball on playfield
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
      Else ' Ball on raised ramp
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+30000, 1, 0, AudioFade(BOT(b) )
        End If
    Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next

  on error resume next ' In case VP is too old..
    ' Kill ball spin
  if mLockMagnet.MagnetOn then
    dim rampballs:rampballs = mLockMagnet.Balls
    dim obj
    for each obj in rampballs
      obj.AngMomZ= 0
      obj.AngVelZ= 0
      obj.AngMomY= 0
      obj.AngVelY= 0
      obj.AngMomX= 0
      obj.AngVelX= 0
      obj.velx = 0
      obj.vely = 0
      obj.velz = 0
    next
  end if
  on error goto 0
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    If ball1.id = 666 OR ball2.id = 666 Then ' 666 is the Menagarie ball
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 1000, Pan(ball1), 0, Pitch(ball1) - 10000, 0, 0,AudioFade(ball1)
    else
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0,AudioFade(ball1)
    End If
End Sub

Sub Pins_Hit (idx)
  PlaySoundAtBall "pinhit_low"
End Sub

Sub Targets_Hit (idx)
  PlaySoundAtBall "target"
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySoundAtBall "metalhit_thin"
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySoundAtBall "metalhit_medium"
End Sub

Sub Metals2_Hit (idx)
  PlaySoundAtBall "metalhit2"
End Sub

Sub Gates_Hit (idx)
  PlaySoundAtBall "gate4"
End Sub

Sub Rollovers_Hit (idx)
  PlaySoundAtBall "rollover"
End Sub

Sub Rubbers_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, .5
End Sub

Sub Posts_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*4)+1, .5
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 3
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 3
End Sub

Sub WBDes_Hit()
me.destroyball
End Sub

Sub NewBCol_Timer()
WBDes.enabled = 0
Set WBall = kicker1.Createsizedball(51):WBall.color = (wbcolor):Wball.image = "powerball4":Wball.id = 666:kicker1.Kick 0, 0
WBall.mass = 1.5
me.enabled = 0
End Sub

Dim bgcan,atime

bgcan = 1

Sub dtcan_timer()
Select Case bgcan
Case 1:
If Can.Rotz => 21 Then
Can3.Visible = 1
Can.Visible = 1
End If
If Can.Rotz => 40 Then
Can3.Visible = 0
End If
If Can.Rotz =>185 Then
bgcan = 2
End If
Can.Rotz = Can.RotZ + 1
Case 2:
If Can.TransY <= -110 Then
Can.TransY = 0
Can.RotZ = 20
Can.Visible = 0
bgcan = 3
End If
Can.TransY = Can.TransY - 1
Case 3:
If atime => 600 Then
atime = 0
me.enabled = 0
End If
atime = atime + 1
Wheel.Rotz = Wheel.Rotz + 1
End Select
End Sub

Dim f24gscale,f24gstate

f24scale = 0
f24gstate = 1

Sub F24F_Timer()
Select Case f24gstate
Case 1:
If f24gscale => 100 Then
me.enabled = 0
End If
f24gscale = f24gscale + 1
f24c.intensityscale = f24gscale
Case 2:
If f24gscale <= 0 Then
me.enabled = 0
Exit Sub
End If
f24gscale = f24gscale - 1
f24c.intensityscale = f24gscale
End Select
End Sub

Dim f21gscale,f21gstate

f21scale = 0
f21gstate = 1

Sub F21F_Timer()
Select Case f21gstate
Case 1:
If f21gscale => 100 Then
me.enabled = 0
End If
f21gscale = f21gscale + 1
f21c.intensityscale = f21gscale
Case 2:
If f21gscale <= 0 Then
me.enabled = 0
Exit Sub
End If
f21gscale = f21gscale - 1
f21c.intensityscale = f21gscale
End Select
End Sub

Dim f25gscale,f25gstate

f25scale = 0
f25gstate = 1

Sub F25F_Timer()
Select Case f25gstate
Case 1:
If f25gscale => 100 Then
me.enabled = 0
End If
f25gscale = f25gscale + 1
f25c.intensityscale = f25gscale
Case 2:
If f25gscale <= 0 Then
me.enabled = 0
Exit Sub
End If
f25gscale = f25gscale - 1
f25c.intensityscale = f25gscale
End Select
End Sub

Dim f26gscale,f26gstate

f26scale = 0
f26gstate = 1

Sub F26F_Timer()
Select Case f26gstate
Case 1:
If f26gscale => 100 Then
me.enabled = 0
End If
f26gscale = f26gscale + 1
f26c.intensityscale = f26gscale
Case 2:
If f26gscale <= 0 Then
me.enabled = 0
Exit Sub
End If
f26gscale = f26gscale - 1
f26c.intensityscale = f26gscale
End Select
End Sub

Dim f37gscale,f37gstate

f37scale = 0
f37gstate = 1

Sub F37F_Timer()
Select Case f37gstate
Case 1:
If f37gscale => 100 Then
me.enabled = 0
End If
f37gscale = f37gscale + 1
f37c.intensityscale = f37gscale
Case 2:
If f37gscale <= 0 Then
me.enabled = 0
Exit Sub
End If
f37gscale = f37gscale - 1
f37c.intensityscale = f37gscale
End Select
End Sub

Sub SLE_Hit()
If ActiveBall.VelY <= 0 Then
vpmTimer.AddTimer 100, "ballsound4"
End If
End Sub

Sub centrampdrop_hit()
vpmTimer.AddTimer 100, "PlaySoundAt ""ball_bounce"",centrampdrop'"
End Sub

bbpos = 1
bbtw = 0
Dim bidx,bbpos,bbtw
Sub boom_shake_Timer()
bbtw = bbtw + 1
If bbtw => 100 Then
me.enabled = 0
bbtw = 0
For each bidx in Boom_Bits:bidx.TransX = 0:next
End If
Select Case bbpos
Case 1:
If Boom_Cap.Transx => 20 Then
bbpos = 2
End If
For each bidx in Boom_Bits:bidx.TransX = bidx.TransX + 1:next
Case 2:
If Boom_Cap.Transx <= 0 Then
bbpos = 1
End If
For each bidx in Boom_Bits:bidx.TransX = bidx.TransX - 1:next
End Select
End Sub

bbpos2 = 1
bbtw2 = 0
Dim bidx2,bbpos2,bbtw2
Sub boom2_shake_Timer()
bbtw2 = bbtw2 + 1
If bbtw2 => 100 Then
me.enabled = 0
bbtw2 = 0
For each bidx2 in Boom_Bits:bidx2.TransX = 0:next
End If
Select Case bbpos2
Case 1:
If Boom_Cap.Transx => 10 Then
bbpos2 = 2
End If
For each bidx2 in Boom_Bits:bidx2.TransX = bidx2.TransX + 1:next
Case 2:
If Boom_Cap.Transx <= 0 Then
bbpos2 = 1
End If
For each bidx2 in Boom_Bits:bidx2.TransX = bidx2.TransX - 1:next
End Select
End Sub

Sub DOF(dofevent, dofstate)
If B2SOn=1 Then
If dofstate = 2 Then
B2SController.B2SSetData dofevent, 1:B2SController.B2SSetData dofevent, 0
Else  B2SController.B2SSetData dofevent, dofstate
End If
End If
End Sub

Function SoundFX (Sound)
  If DOFs Then
    SoundFX = ""
  Else
    SoundFX = Sound
  End If
End Function

Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
    If b2sOn = 1 Then
  B2SController.Stop
  End If
End Sub


Dim NextOrbitHit:NextOrbitHit = 0

Sub WireRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump3 .5, Pitch(ActiveBall)+5
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .5 + (Rnd * .2)
  end if
End Sub

Sub WireLaunchRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump .5, Pitch(ActiveBall)+5
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .2 + (Rnd * .2)
  end if
End Sub

Sub PlasticRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump 1, -20000
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .1 + (Rnd * .2)
  end if
End Sub



' Requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
  dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

'' Requires metalguidebump1 to 2 in Sound Manager
'Sub RandomBump2(voladj, freq)
' dim BumpSnd:BumpSnd= "metalguidebump" & CStr(Int(Rnd*2)+1)
'   PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
'End Sub

' Requires WireRampBump1 to 5 in Sound Manager
Sub RandomBump3(voladj, freq)
  dim BumpSnd:BumpSnd= "WireRampBump" & CStr(Int(Rnd*5)+1)
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub



' Stop Bump Sounds
Sub StopBumpSounds()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:StopSound "PlasticRampBump" & i:next
NextOrbitHit = Timer + 1
  Debug.print "StopBumps"

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


'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
    PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub


'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
    PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub
