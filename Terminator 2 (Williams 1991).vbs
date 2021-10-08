Option Explicit
Randomize

'Terminator 2 (Williams 1991)
'Script, meshing, physics etc by nfozzy
'Derived from 'red mod' of T2 by 32Assassin
'PF redraw, ramp sticker and some plastics by nfozzy
' - Some plastic images may be derived from 32Assassin, Tipito or UncleReamus's redraws
'Low poly locknut derived from Dark's hipoly one
'Hunter killer ship by Lio
'sfx by knorr
'Scorecard by Inkochnito

'Special thanks to ClarkKent for the PF scan, arconovum for the FSS xml
'=============================================

'Press KeyRules (Default R) to view the rulecard
'left magnasave changes GI color. Holding the right magnasave changes strings independently

'You can also test custom colors in the debug window with 'GIcolor', which can take R,G,B info in an array.
'For example: GIcolor Array(255,15,100)


'changelog

'Version 1.12a - Lamps now calculate at -1 interval by default. Fixed a script error with DynamicLamps that was causing excessive updates

'Version 1.12 -
'Modified skull mesh + added animation
'New optional cannon speed - supercharged
'Optimizations:
'Playfield AO image resized from 8k to 4k
'Animations now calculate at -1 interval by default

'Version 1.11 - Turned floating text off by default (was supposed to be off by Default), tweaked flippers

'Version 1.1 -
'*New head mesh by nFozzy
'*New bumper cap texture
'*Reworked flasher fading speed and output filters to be more realistic
'*Fixed clamping on some very dark flasher images
'*Optional floating text script

'*Physics -
'*New (slower) cannon speed default (old cannon is available in options)
'*Rebuilt flippers (Old flippers are available in options)
'*Remade sweep shot metal with a better curve
'*Special nudge script now enabled by default

'*Known bugs: Day/Night slider is screwed up

'Version 1.02 - new screws and locknuts
'Version 1.01 - Save GI colors, Removed y-based flipper coef safety, few missing credits


'TODO
'fix day/night slider


if Version < 10400 then msgbox "This table requires Visual Pinball 10.4 beta or newer!" & vbnewline & "Your version: " & Version/1000
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0
SetLocale(1033)

'Options

'Cannon speeds
'1: Slow, Factory Default
'2: Moderate, Aftermarket Cannon
'3: Supercharged Cannon
Const GunType = 1

Const OldFlippers = False 'Old, weaker 1.0 flipper values. WILL OVERRIDE EDITOR FLIPPER STRENGTH! (Default: False)
Const FloatingText = False 'Floating popup scores. Only works for player 1! Currently broken with B2S (Default: False)
Const FastLamps = False 'Fast lamp calculation, more accurate at 60FPS. Turn off if lamps are moving in slo-mo. (default: False)
Const FastAnimations = False 'Fast animation calculation, more accurate at 60FPS. Turn off if animations are moving in slo-mo. (default: False)

Const KeyboardNudgescript = True 'Cool keyboard nudge script, but causes script errors in some locales (Default: True)

'Ambient Occlusion PF layer (Big shadow texture, ~100MB)
AOPass.Visible = True 'default True

DMD_FSS.visible = table1.ShowFSS 'Cabinet DMD
siderails.visible = table1.showdt
Lockdownbar.visible = table1.showdt

'GIsideL.imagea = "gisideLlow" 'alternative sidewall GI reflections, for a low camera perspective (uncomment to use)
'GIsideR.imagea = "gisideRlow"

dim ballsize : ballsize = 50
dim ballmass : ballmass = 1

dim DeskTopMode : DeskTopMode = table1.showdt
dim UseVPMColoredDMD : UseVPMColoredDMD = DMD_FSS.visible
dim UseVPMNVRAM : UseVPMNVRAM = cBool(FloatingText)
Const cGameName="t2_l8"
const uselamps = 0, UseSolenoids = 0 'don't mess with me
Const UseVPMModSol = True
Const SCoin = "fx_Coin"
LoadVPM "02700000","WPC.vbs",3.56
redim gicallback(0)
redim gicallback2(0)

'Something similar to this will go into core.vbs soon
dim vpmSolFlipsTEMP : set vpmSolFlipsTEMP = New cvFastFlipsZ : vpmSolFlipsTEMP.Name = "vpmSolFlipsTEMP"
vpmSolFlipsTEMP.DebugOn = False 'debug
DebugToggle 'there's a ton of these textboxes

'*************************************************

Sub InitSolFlipz()
	'if not UseSolFlips then exit sub
	'vpmSolFlips.Solenoid = GameOnSolenoid
	'if not IsEmpty(SolCallback(sLLFlipper)) then vpmSolFlips.CallBackL = SolCallback(sLLFlipper)
	'if not IsEmpty(SolCallback(sLRFlipper)) then vpmSolFlips.CallBackR = SolCallback(sLRFlipper)
	vpmSolFlipsTEMP.Solenoid = 31'GameOnSolenoid
	if not IsEmpty(SolCallback(sLLFlipper)) then vpmSolFlipsTEMP.CallBackL = SolCallback(sLLFlipper)
	if not IsEmpty(SolCallback(sLRFlipper)) then vpmSolFlipsTEMP.CallBackR = SolCallback(sLRFlipper)

End Sub

Sub Table1_Init
	vpmInit Me
	InitSolFlipz
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Terminator 2 (Williams 1991)"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        '.hidden = not table1.showdt
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

	LoadColors	'load last GI colors from VPReg.stg

	UpdateTroughSwitches

	vpmNudge.TiltSwitch=14
	vpmNudge.Sensitivity=3
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingShot,RightSlingShot)
	vpmnudge.solgameon False
	'Controller.Switch(24) = 1 'sw24 Test Position, Always Closed
	'Controller.Switch(22) = 1 'Coin door

	scorecard_DT.visible = DeskTopMode or table1.showFSS : Scorecard_FS.visible = Not DeskTopMode
End Sub

Sub Table1_Paused() : StopAllrolling : Controller.Pause = True: End Sub
Sub Table1_Exit()
	SaveColors	'Save GI colors to VPReg.stg
	Controller.Pause = False
	Controller.Stop
End Sub


dim catchinput(1)
Sub Table1_KeyDown(ByVal KeyCode)
	If keycode = PlungerKey Then controller.Switch(34) = 1 : Playsound "Trigger", 0, 0.1, 0.05	'sw34 Grip Trigger
	If keycode = LeftTiltKey Then if KeyboardNudgescript then nfNudge -1, 0.2
	If keycode = RightTiltKey Then if KeyboardNudgescript then nfNudge 1, 0.2
	If keycode = CenterTiltKey Then if KeyboardNudgescript then nfNudge 0, 0.2
	if keycode = keyrules then lampz.state(10) = 1
	if keycode = leftmagnasave then catchinput(0) = True : changeGI
	if keycode = Rightmagnasave then catchinput(1) = True
	'TEMP!!!
	if keycode = RightFlipperKey then vpmSolFlipsTEMP.FlipR True : controller.Switch(swLRFlip) = 1: Exit Sub
	if keycode = LeftFlipperKey then vpmSolFlipsTEMP.FlipL True	: controller.Switch(swLLFlip) = 1: Exit Sub

	If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If keycode = PlungerKey Then controller.Switch(34) = 0 'sw34 Grip Trigger
	if keycode = keyrules then lampz.state(10) = 0
	if keycode = leftmagnasave then catchinput(0) = False
	if keycode = Rightmagnasave then catchinput(1) = False
	'TEMP!!!
	if keycode = RightFlipperKey then vpmSolFlipsTEMP.FlipR False : controller.Switch(swLRFlip) = 0: Exit Sub
	if keycode = LeftFlipperKey then vpmSolFlipsTEMP.FlipL False : controller.Switch(swLLFlip) = 0: Exit Sub

	If KeyUpHandler(keycode) Then Exit Sub
End Sub

Sub nfNudge(aDir,Strength)
	dim steps: steps = 3
	nfDoNudge aDir, Strength, 1, steps
	dim str : str = "nfDoNudge " & aDir & ", "& strength/2 & ", " & "-1," & steps & " REM"
	vpmtimer.addtimer 250, str '	'Step 2: reverse direction with half strength
End Sub

sub nfDoNudge(aDir, Strength, Invert, aStep)
	dim b: if aDir = 0 then
		for each b in getballs : if not NudgeOutOfBounds(b) then b.Vely = b.Vely + strength * Invert end if : next
	else
		for each b in getballs : if not NudgeOutOfBounds(b) then b.Velx = b.VelX + strength * aDir * Invert end if : next
	end if
	dim str : str = "nfDoNudge " & aDir & ", "& strength & ", " & invert & "," & aStep-1 & " REM"
	'debug.print str
	if aStep > 0 then vpmtimer.addtimer 40, str
End Sub

Function NudgeOutOfBounds(byval aBall) 'this is annoying. Balls in kickers will accumulate fake velocity and start rolling sounds.
	if aBall.x > 730 and aball.y > 1900 or aball.ForceReflection then NudgeOutOfBounds = True
End Function

'Physics Debug Commands
Sub Destroyer_Hit():me.destroyball:end sub 'trough will automatically delete overflow
Sub DebugToggle() : dim x: for each x in TextBoxes : x.visible = not x.visible : next : End Sub

'=============
'Lamp Handling and object init
'=============

dim gilevels1,gilevels2
Sub InitGI()
	dim a : a = flashers.obj(2)	: redim gilevels1(uBound(a))
	dim x : for x = 0 to uBound(gilevels1) :
		if typename(a(x)) = "Flasher" then
			gilevels1(x) = a(x).opacity
		elseif typename(a(x) ) = "Light" then
			gilevels1(x) = a(x).Intensity
	end if : Next
	a = flashers.obj(4)	: redim gilevels2(uBound(a))
	for x = 0 to uBound(gilevels2) : if typename(a(x)) = "Flasher" then gilevels2(x) = a(x).opacity end if : Next
End Sub

dim Colorseq : Colorseq = 1
Sub changeGI()
	if flashers.lvl(4) * flashers.lvl(2) = 0 then exit sub
	playsound "Trigger"
	gicolor colorseq  : if Colorseq = 6 then Colorseq = 0 else Colorseq = Colorseq + 1 end if
End Sub

Function Luminance(aRgb)	'input: array, output: value between 0 and 255 '240?
	Luminance = (argb(0)*0.3 + argb(1)*0.59 + argb(2)*0.11)
End Function

Sub SaveColors()
	SaveValue "T2NF","GIcolorL", formatRGB(GIcolorL)
	SaveValue "T2NF","GIcolorR", formatRGB(GIcolorR)
	'tb.text = "saving:" & vbnewline & formatRGB(GiColorL) & vbnewline & FormatRGB(GiColorR)
End Sub

Sub LoadColors()
	if LoadValue("T2NF", "GIcolorL") = "" then exit sub
	dim x,x2 : x = LoadValue("T2NF", "GIcolorR") : 	x2 = LoadValue("T2NF", "GIcolorL")
	catchinput(1) = False : gicolor Array(mid(x, 1, 3),mid(x, 4, 3),mid(x, 7, 3))
	catchinput(1) = True : gicolor Array(mid(x2, 1, 3),mid(x2, 4, 3),mid(x2, 7, 3)) : CatchInput(1) = False

	'tb.text = mid(x2, 1, 3) & " " & mid(x2, 4, 3) & " " & mid(x2, 7, 3)
End Sub

Function FormatRGB(Byval aArray)
	dim x: for x = 0 to 2
		if aArray(x) < 10 and len(aArray(x)) = 1 then
			aArray(x) = "00" & aArray(x)
		elseif aArray(x) < 100 and Len(aArray(x)) = 2 then
			'debug.print "array" & x & "(" & aArray(x) & ") < 100, addding a 0 before it"
			aArray(x) = "0"  & aArray(x)
		end if
	Next
	FormatRGB = aArray(0) & aArray(1) & aArray(2)
End Function

function DeSat(ByVal aRGB, ByVal aSat)	'simple desaturation function (returns rgb)
	dim r, g, b, L
	L = 0.3*aRGB(0) + 0.59*aRGB(1) + 0.1*aRGB(2)
	r = aRGB(0) + aSat * (L - aRGB(0))
	g = aRGB(1) + aSat * (L - aRGB(1))
	b = aRGB(2) + aSat * (L - aRGB(2))
	'desat = array(r,g,b)	'return array
	desat = rgb(r,g,b)		'return rgb
End Function

Sub MatchColorWithLuminance(aFlashers, aRGB, aOpacity)
	dim lumincoef : lumincoef = 255/luminance(aRGB)/1
	dim x: for x = 0 to uBound(aFlashers)
		if typename(aFlashers(x) ) = "Flasher" then
			aFlashers(x).Opacity = aOpacity(x) * lumincoef
			'tb.text = "set" & aFlashers(x).name & "to " & aOpacity(x) & " * " & round(lumincoef,4) & vbnewline & round(Luminance(argb),3)
			aFlashers(x).color = rgb(argb(0), aRGB(1), aRGB(2))
			if aFlashers(x).name = "GIsideL" or aFlashers(x).name = "GIsideR" then
				if rgb(argb(0), aRGB(1), aRGB(2)) = 16777215 then   	'if white, slightly less boring reflection color
					aFlashers(x).color = rgb(65,127,255)
				End If
			End If
		elseif Typename(aFlashers(x) ) = "Light" Then
			aFlashers(x).Intensity = aOpacity(x) * lumincoef
			if aFlashers(x).colorfull then aFlashers(x).colorfull = rgb(argb(0), aRGB(1), aRGB(2))
			if aFlashers(x).color then aFlashers(x).color = rgb(argb(0), aRGB(1), aRGB(2))
		end If
	Next
	'tb.text = round(lumincoef,3)
End Sub

dim GiColorL, GiColorR 'Save GIcolor in memory
Sub GIcolor(input)	'can input rgb in array form
'	dim aL   : aL = Array(gil, gilp, gihand)	'cut image swap stuff
'	dim aL1 : aL1 = Array("gil", "gilp" ,"gihand")
'	dim aR   : aR = Array(gir, girp1, girp2, girp3, girp_corner, girp_corner1, gihand1)
'	dim aR1 : aR1 = Array("gir", "girp1", "girp2","girp3","girp_corner","girp_corner","gihand")
	dim x, c : c = Array(255,255,255)
	if isarray(input) then
		c = input
	else
		select case input
			'gicolor array(255, 15, 100)	'pink meh
			case 0	: c = Array(255, 162, 37)'(255, 127, 37)	'Warm White (LED)
			case 1  : c = Array(255,255,255)	'white
			case 2	: c = Array(255, 15, 3)		'Red flood
			'case 3	: c = Array(133, 11, 255)	'purple Flood
			case 3	: c = Array(85,13,255)		'95% sat Violet
			case 4	: c = Array(36, 54, 255)	'Blue
			case 5	: c = Array(20, 255, 12)	'Green
			case 6	: c = Array(5, 255, 127)	'aqua
		end Select	'gicolor array(255,45,0) 'orange
	end if
	MatchColorWithLuminance Flashers.obj(2), c, gilevels1 : GiColorL = c
	if not catchinput(1) then MatchColorWithLuminance Flashers.obj(4), c, gilevels2 : GiColorR = c
'	for x = 0 to ubound(aL)
'		if aL(x).imagea <> aL1(x) then aL(x).imagea = aL1(x)
'	Next
'	for x = 0 to ubound(aR)
'		if aR(x).imagea <> aR1(x) then aR(x).imagea = aR1(x)
'	Next
End Sub

Sub tweakopacity(aInput, aInput2)	'debug
	dim a : a = Array(gil, gir, gilp, girp1, girp2, girp3, girp_corner, girp_corner1, gihand, gihand1, gisideL, gisideR)
	'dim aa : aa = Array(gil, gir, gilp, girp1, girp2, girp3, girp_corner, girp_corner1, gihand, gihand1)
	'dim a2 : a2 = Array("gil", "gir", "gilp", "girp1", "girp2","girp3","girp_corner","girp_corner","gihand")
	dim x: for x = 2 to uBound(a) : a(x).opacity = aInput : Next
	for x = 6 to uBound(a) : a(x).opacity = aInput2 : Next
End Sub

'gicolor 3
'gil.opacity = 100000 : gir.opacity = gil.opacity

Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class	'todo do better
dim NullFader : set NullFader = new NullFadingObject
dim Lampz : Set Lampz = New LampFader			'Lamps
dim Flashers : set Flashers = New DynamicLamps	'Flashers & GI
InitLamps
Sub InitLamps()
	'setup insert masks
	dim x : for each x in Inserts : x.ImageB = x.name & "_BaseTexBaked" : Next
	dim a : a = Array(LeftRampCover,LeftRampSign,RIghtRampCover, RightRampSign)
	for each x in a : x.z = 0 : next

	'all primitives are on z = 0
	'by starting higher meshes > 0 then bringing them down in script can sometimes correct draw issues without using DepthBias
	Plastics_LVL2.z = 0
	BumperCaps.z = 0
	RightRamp.z = 0

	for each x in Inserts	 'Simple map lamps
		if TypeName(x) = "Flasher" or TypeName(x) = "Light" then
			On Error Resume Next
			'if Mid(x.Name,1,1) = "l" or Mid(x.Name,1,1) = "L" Then
				x.UserValue = cInt(mid(x.name,2,2))
				Set lampz.obj(x.UserValue) = x
				x.UserValue = Empty
			'end if
			On Error Goto 0
		End If
	Next
	lampz.obj(17) = Array(l17, l17a)
	lampz.obj(21) = Array(l21, l21t)
	lampz.obj(22) = Array(l22, l22a)
	lampz.obj(51) = Array(l51, l51a)
	lampz.Callback(10) = "animscorecard"
	lampz.FadeSpeedup(10) = 0.006
	lampz.FadeSpeedDown(10) = 0.008
	Lampz.Callback(52) = "L52Eyes"

	Lampz.Filter = "TestFunction"	'put lamp output through a function. Low-pass or High-pass filter or whatever. (only set lamps, no callbacks)
	flashers.Filter = "TestFunction2"	'

	'Flashers & GI

	for x = 5 to 28
		flashers.FadeSpeedUp(x) = 1/64
		flashers.FadeSpeedDown(x) = 1/64
	Next

	Set Flashers.obj(0) = GI_BackGlass '"GI 0 - Top Insert"
		'"GI 1 - Bottom Insert" '?
	Flashers.obj(3) = Array(Lgi1, Lgi2, Lgi3, Lgi4)	'cpu

	Flashers.obj(2) = Array(GIR, GIRP1,GIRP2,GIRP3,GIRP_Corner,GIrp_corner1,GIRT_1,GIRT_11,GIRT_12,GIRT_15,_
	GIRT_16,GIRT_17,GIRT_18,GIRT_2,GIRT_3,GIRT_4,GIRT_5,GIhand1, gibackwallR,GIsideR,GImtl,GIRcannonT, gilaneR)	'GI 2 - Right
	Flashers.Callback(2) = "GIright"
	Flashers.obj(4) = Array(GIL, GILP, Gilt_1, GILT_7, GILT_8,GILT_9,GIhand,GIbackwallL,GIsideL,gilaneL)	'GI 4 - Left
	Flashers.Callback(4) = "GIleft"
	'Flashers.Callback(2) = "GIJets"
	for x = 0 to 4 : flashers.fadespeedup(x) = 0.02 : flashers.fadespeeddown(x) = 0.015 : next

	'Set Flashers.obj(17) = F17A								'Hot Dog
	Flashers.Callback(17) = "HotDogSwap"					'Hot Dog (image, no Filter)
	Flashers.obj(18) = Array(F18, F18P, F18W, f18_bg, f18hand)		'Right Sling
	Flashers.obj(19) = Array(F19, F19P, F19W, F19_BG, f19hand)		'Left Sling
	Flashers.Callback(18) = "f18bulbUpdate"
	Flashers.Callback(19) = "f19bulbUpdate"

	Flashers.obj(20) = Array(F20, F20P, F20W, F20_BG)		'Left Lock
	Flashers.obj(21) = Array(F21, F21P, F21W, f21pl)				'Gun
	Flashers.Callback(21) = "GunBulbs"
	Flashers.obj(22) = Array(F22, F22BW, F22W, F22T)		'Right Ramp
	Flashers.obj(23) = Array(F23, F23BW, F23W, F23P, f23_BG)'Left Ramp
	Set Flashers.obj(24) = F24								'backglass flasher
	Flashers.obj(25) = Array(F25, F25w)						'Targets
	Flashers.obj(26) = Array(F26, F26w, F26BW, F26P)		'Popper, Left
	Flashers.obj(27) = Array(F27, F27Bw, F27P, f27_bg)		'Popper, Right


	dim daynightcoef : daynightcoef = LinearEnvelope(table1.NightDay, array(11,50,100), array(5,1,0.4))
	'tb.text = table1.nightday & vbnewline & daynightcoef
	modulateall daynightcoef


	'flashers.InitValues 'saves initial intensity and opacity
'

	Lampz.Update	'start all lamps off
	Flashers.Update	'start all flashers off

	Lampz.Init
	Flashers.Init

	InitGI 'saves GI opacity and other info
	gicolor 0

End Sub

dim DayNightCoef
Sub ModulateAll(aCoef)
	DayNightCoef = aCoef
	On Error Resume Next
	dim x : for x = 0 to 28 : flashers.modulate(x) = aCoef : Next
	for each x in Inserts
		x.Opacity = x.Opacity * aCoef
		x.Intensity = x.Intensity * aCoef
	next
End Sub

Function TestFunction(byVal a)	'you can put all lamps through this function, for a lowpass or something
	TESTFUNCTION = a^1.6
End Function
Function TestFunction2(byVal a)	'you can put all lamps through this function, for a lowpass or something
	TESTFUNCTION2 = a^1.8
End Function

Sub HotDogSwap(aLVL)
f17a.intensityscale = aLvl : f17.intensityscale = aLvl
End Sub 'no filter


Sub GunBulbs(aLVL) : F21bulb.BlendDisableLighting = Flashers.FilterOut(aLVL) * 17.4 : End Sub
Sub F18bulbUpdate(aLVL) : F18bulb.BlendDisableLighting = Flashers.FilterOut(aLVL) * 17.4 : End Sub
Sub F19bulbUpdate(aLVL) : f19bulb.BlendDisableLighting = Flashers.FilterOut(aLVL) * 17.4 : End Sub

dim GIscale : GIscale = 2.5
Sub GIleft(ByVal aLVL)
	gi_bulbL.BlendDisableLighting = aLvl *15 '1.03b	'didn't know this could go over 1!
	aLVL = Flashers.FilterOut(aLVL)
	'brighten the flashers when GI is off

	dim Offset
	Offset = (GiScale-1) * (ABS(aLvl-1 )  ) + 1	'invert

	'dim offset : offset = pSlope(aLVL, 1, 1, 0, GIscale)	'temp
	flashers.modulate(19) = pSlope(aLVL, 1, 1, 0, 4) * DayNightCoef
	flashers.modulate(20) = Offset * DayNightCoef
	flashers.modulate(23) = Offset * DayNightCoef
	flashers.modulate(25) = Offset * DayNightCoef
	flashers.modulate(26) = Offset * DayNightCoef
	'tb.text = "offset " & round(offset,3)  & vbnewline & f18.opacity
	'special, hotdog
	flashers.modulate(19) = pSlope(aLVL*flashers.lvl(2), 1, 1, 0, 2.5) 'todo
End Sub

Sub GIright(ByVal aLVL)
	gi_bulbR.BlendDisableLighting = aLvl *15 '1.03b	'didn't know this could go over 1!
	aLVL = Flashers.FilterOut(aLVL)

	dim Offset
	Offset = (GiScale-1) * (ABS(aLvl-1 )  ) + 1	'invert

	flashers.modulate(18) = pSlope(aLVL, 1, 1, 0, 5) * DayNightCoef
	flashers.modulate(21) = offset * DayNightCoef
	flashers.modulate(22) = offset * DayNightCoef
	flashers.modulate(27) = offset * DayNightCoef

	'jets
	dim x, a
	a = Array(GIjet_Transmit1, GIjet_Ambient1, GIJet_Flasher1, _
				GIjet_Transmit2, GIjet_Ambient2, GIJet_Flasher2, _
					GIjet_Transmit3, GIjet_Ambient3, GIJet_Flasher3)
	for each x in a : x.IntensityScale = aLvl : next
	a = Array(GIjet_BulbP1,GIjet_BulbP2,GIjet_BulbP3)
	for each x in a : x.BlendDisableLighting = aLvl*3 : next '*3 1.03b	'didn't know this could go over 1!
End Sub

Sub L52Eyes(ByVal aLVL) ': tb.text = "L52: " & aLVL :_

	aLVL = Lampz.FilterOut(aLVL)
	if aLVL = 0 then L52p.image = "L52P_off" else L52p.image = "L52P_on"
	L52p.BlendDisableLighting = aLvl*3	'*3 1.03b	'didn't know this could go over 1!
	l52a.intensityscale = aLvl : l52a1.intensityscale = aLvl
End Sub

Sub animScoreCard(alvl)
	if desktopmode then
		scorecard_DT.showframe alvl
	Else
		scorecard_fs.showframe alvl
	end If
End Sub


'Debug functions
Sub TestFlashers()	'cycle through flashers
	dim x : for x = 17 to 27
		nftimer.addtimer (x-16)*150, "setlamp " & x & ", 1"
		nftimer.addtimer ((x-16)*200)+220, "setlamp " & x & ", 0"
	Next
End Sub
Sub SetLamp(nr, input) 'legacy debug commands. Also supports arrays
	if IsArray(nr) then : dim x : for each x in nr : Flashers.state(x) = input : next : exit sub
	Flashers.state(nr) = input
End Sub
Sub SetLampz(nr, input) 'legacy debug commands. Also supports arrays
	if IsArray(nr) then : dim x : for each x in nr : lampz.state(x) = input : next : exit sub
	lampz.state(nr) = input
End Sub
Sub GIoff() : dim x: for x = 0 to 5 : Flashers.State(x) = 0 : next : end sub 'debug gi
Sub GIon() : dim x: for x = 0 to 5 : Flashers.State(x) = 1 : next : end sub 'debug gi
'---------------------


'Game Update Timers
'----------------------------
'Logic updates go on interval = 1 (ControllerUpdates)
'Object updates go here on interval = -1 (GameUpdates)
'----------------------------
Dim FrameTime, InitFrameTime
GameUpdates.Interval = -1 'don't touch
Sub GameUpdates_Timer()
	FrameTime = GameTime - InitFrameTime : InitFrameTime = GameTime
	'Lamp Game updates
	if FastLamps then
		Lampz.Update
		Flashers.Update
	Else
		Lampz.Update2
		Flashers.Update2
	end if

	'Animation Game Updates
	aSkullHit.Update2
	alutburst.update2
	if FastAnimations then
		aPopper.Update
		aBumper1.Update
		aBumper2.Update
		aBumper3.Update
		aLeftSlingArm.Update
		aRightSlingArm.Update
		aLeftSling.Update
		aRightSling.Update
		aLock1.Update
		aLock2.Update
		aGun.Update
		aGunPlunger.Update
		aDTup.Update
		aDTDown.Update
		aDThit.Update
	Else	'frametime calculated fading and updating. Otherwise, two updates on two timers: 'update1 / update'
		aPopper.Update2
		aBumper1.Update2
		aBumper2.Update2
		aBumper3.Update2
		aLeftSlingArm.Update2
		aRightSlingArm.Update2
		aLeftSling.Update2
		aRightSling.Update2
		aLock1.Update2
		aLock2.Update2
		aGun.Update2
		aGunPlunger.Update2
		aDTup.Update2
		aDTDown.Update2
		aDThit.Update2
	end if



	'Ball clones
	cLeftLock.Update : cPopper.Update: cTopLock.Update
	'Virtual Kickers
	LeftLock.Update : Popper.Update : TopLock.Update

	'Rolling updates
	roll.update
End Sub

'Controller/Logic updates
ControllerUpdates.Interval = 1 'don't touch
Sub ControllerUpdates_Timer()
	dim x, chglamp, chgsol, chgGI, tmp
	chglamp = Controller.ChangedLamps
	If Not IsEmpty(chglamp) Then
		For x = 0 To UBound(chglamp) 			'nmbr = chglamp(x, 0), state = chglamp(x, 1)
			Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
		next
	End If

	chgsol = Controller.ChangedSolenoids
	If Not IsEmpty(ChgSol) then
		for x = 0 to uBound(ChgSol)
			Select Case ChgSol(x, 0)
				Case 17, 18, 19, 20, 21, 22, 23, 24,25,26,27 : Flashers.State(ChgSol(x,0) ) = ChgSol(x,1)/255 'Flashers
				'Case 31 : if UseSolFlips then vpmSolFlips.TiltSol cbool(ChgSol(x, 1))
				Case 31 : vpmSolFlipsTEMP.TiltSol cbool(ChgSol(x, 1))	'TEMP
				Case Else
					'SolModCallBack
					tmp = SolModCallback(ChgSol(x,0) )
					If tmp <> "" Then Execute tmp & " " & ChgSol(x, 1)
					'SolCallBack
					if ChgSol(x,1) > 1 then ChgSol(x,1) = 1
					If ChgSol(x, 1) <> SolPrevState((ChgSol(x,0) )) Then
						SolPrevState((ChgSol(x,0) )) = ChgSol(x, 1)
						tmp = SolCallback((ChgSol(x,0) ))
						If tmp <> "" Then Execute tmp & vpmTrueFalse(ChgSol(x, 1)+1)
					end if
			end Select

		Next
	end If
	chgGI = Controller.ChangedGIStrings
	If Not IsEmpty(chgGI) then
		for x = 0 to uBound(chgGI)
			'GIStates(chgGI(x, 0) ) = chgGI(x, 1)	'temp debug box thing
			if chgGI(x, 1) = 8 then chgGI(x, 1) = 7	'something about 8 being unreliable
			Flashers.State(chgGI(x,0) ) = chgGI(x, 1)/7
		Next
	end If
	If FastLamps then
		Lampz.Update1
		Flashers.Update1 	'GI and Flashers both handled by 'Flashers' (DynamicLamps object)
	end if
	'Animation Updates `````````````````````
	if FastAnimations then
		'aTester.Update1
		aPopper.Update1
		aBumper1.Update1
		aBumper2.Update1
		aBumper3.Update1
		aLeftSlingArm.Update1
		aRightSlingArm.Update1
		aLeftSling.Update1
		aRightSling.Update1
		aLock1.Update1
		aLock2.Update1
		aGun.Update1
		aGunPlunger.Update1
		aDTup.Update1
		aDTDown.Update1
		aDThit.Update1
	end if

	'COR tracker
	cor.update	'room for optimization

	'NFtimer (only used for Debug)
	nftimer.update

End Sub


'===================
'KEYFRAME ANIMATIONS
'===================
dim aBumper1 : Set aBumper1 = New cAnimation
dim aBumper2 : Set aBumper2 = New cAnimation
dim aBumper3 : Set aBumper3 = New cAnimation
dim aLeftSlingArm : Set aLeftSlingArm = New cAnimation
dim aRightSlingArm : Set aRightSlingArm = New cAnimation
dim aLeftSling : Set aLeftSling = New cAnimation
dim aRightSling : Set aRightSling = New cAnimation
dim aLock1, aLock2 : Set aLock1 = New cAnimation : Set aLock2 = New cAnimation
dim aPopper : Set aPopper = New cAnimation
dim aGun : set aGun = New cAnimation
dim aGunPlunger : Set aGunPlunger = new cAnimation
dim aDTup : set aDTup = new cAnimation
dim aDTDown : set aDTDown = new cAnimation
dim aDThit : set aDThit = new cAnimation
dim aSkullHit : Set aSkullHit = New cAnimation


Sub animSkullHit(aLVL) : dim x : for each x in array(skull, l52p, L52case) : x.RotX = aLvl : next : End Sub

Sub SkullHit(aSpeed)
	if aSpeed > 15 then
		aSkullHit.ModPoint 1, 48, PSlope(aSpeed, 15, -6, 50, -12)
		askullhit.play
	end if
End Sub

InitAnimations
Sub InitAnimations()
	dim x

	with aSkullHit
		askullHit.addpoint 0, 0, 0
		askullHit.addpoint 1, 48, -7
		askullHit.addpoint 2, 96, 0
		.Callback = "animSkullHit"
	end With

	for each x in Array(aBumper1, aBumper2, aBumper3)
		x.AddPoint 0, 0, 0
		x.AddPoint 1, 21, -48	'5 down
		x.AddPoint 2, 67, -48	'11 hold
		x.AddPoint 3, 100, 2.9	'8 Up
		x.AddPoint 4, 121, 0
		x.AddPoint 5, 141, 1.6
		x.AddPoint 6, 161, 0
	next
	aBumper1.Callback = "animBumper1"
	aBumper2.Callback = "animBumper2"
	aBumper3.Callback = "animBumper3"

	for each x in array(aLeftSlingArm, aRightSlingArm)
		x.AddPoint 0, 0, -4.25
		x.AddPoint 1, 10, 0		'hit sling
		x.AddPoint 2, 31, 17.25		'5 down
		x.AddPoint 3, 77, 17.25		'11 hold
		x.AddPoint 4, 110, -4.25	'8 Up
	next
	aLeftSlingArm.Callback = "animLeftSlingArm"
	aRightSlingArm.Callback = "animRightSlingArm"

	for each x in Array(aLeftSling, aRightSling)
		x.AddPoint 0, 0, 0
		x.AddPoint 1, 10, 0	'wait for kicker
		x.AddPoint 2, 31, 1		'5 down
		x.AddPoint 3, 77, 1		'11 hold
		x.AddPoint 4, 110, 0	'8 Up
	next
	aLeftSling.Callback = "animLeftSling"
	aRightSling.Callback = "animRightSling"

	for each x in Array(aLock1, aLock2) 'Williams Kickers
		x.AddPoint 0, 0, 180
		x.AddPoint 1, 21, 155	'5 down
		x.AddPoint 2, 67, 155	'11 hold
		x.AddPoint 3, 100, 185	'8 Up
		x.AddPoint 4, 121, 180
	Next
	aLock1.Callback = "animLock1"
	aLock2.Callback = "animLock2"

	With aPopper				'Williams Popper
		.AddPoint 0, 0, 0
		.AddPoint 1, 21, 47	'5 down
		.AddPoint 2, 67, 47	'11 hold
		.AddPoint 3, 100, -10	'8 Up
		.AddPoint 4, 121, 0
		.AddPoint 5, 141, -3.5
		.AddPoint 6, 161, 0
	End With
	aPopper.Callback = "animPopper"

	with aGun
		Select Case GunType
			Case 3	'Case 3: Supercharged speed  (superjackpot = ???)
				.AddPoint 0, 0, 0				'home
				.AddPoint 1, 200, 0				'home
				.AddPoint 2, 1817, 1			'full rotation
				.AddPoint 3, 1983, 1			'full rotation
				.AddPoint 4, 3867, 2			'returned home
			Case 2	'Case 2: Faster Aftermarket cannon  (superjackpot = Middle Target)
				.AddPoint 0, 0, 0				'home
				.AddPoint 1, 425, 0				'home
				.AddPoint 2, 2900, 1			'full rotation
				.AddPoint 3, 3425, 1			'full rotation
				.AddPoint 4, 5425, 2			'returned home
			Case Else	'case 1: Slow cannon Factory Default (superjackpot = Top / Bottom Targets)
				.AddPoint 0, 0, 0				'home
				.AddPoint 1, 725, 0				'home '.AddPoint 1, 795, 0
				.AddPoint 2, 3892, 1			'full rotation
				.AddPoint 3, 4339, 1			'full rotation
				.AddPoint 4, 7650, 2			'returned home '.AddPoint 4, 7701, 2
		End Select
	End With
	aGun.debugOn = False
	aGun.Callback = "animGun"

	with aGunPlunger
		.AddPoint 0, 0, 0			'home
		.AddPoint 1, 21, 47			'fired
		.AddPoint 2, 67, 47			'hold
		.AddPoint 3, 100, 0			'return
	End With
	aGunPlunger.Callback = "animGunPlunger"

	'drop target
	with adTup
		.AddPoint 0, 0, -35	'rest -35	'mod me if in a different spot...
		.AddPoint 1, 21, -30 'up
		.AddPoint 2, 100, -30 'hold
		.AddPoint 3, 133, -35 '(rest)
	End With
	aDTup.Callback = "animDTup"

	with aDTdown
		.AddPoint 0, 0, -35	'rest -35
		.AddPoint 1, 33, -100 'Down
	End With
	aDTDown.Callback = "animDTdown"

	with aDTHit
		.AddPoint 0, 0, 0	'rest
		.AddPoint 1, 33, 5 	'rot back 'mod me based on ball speed !
		.AddPoint 2, 66, 0 	'rest
	End With
	aDTHit.Callback = "animDThit"
End Sub

'Animation Callbacks
Sub animBumper1(Value) : Bumper1Ring.Z = Value : End Sub
Sub animBumper2(Value) : Bumper2Ring.Z = Value : End Sub
Sub animBumper3(Value) : Bumper3Ring.Z = Value : End Sub
Sub animLeftSlingArm(Value) : SlingKicker1.RotX = Value : End Sub
Sub animRightSlingArm(Value) : SlingKicker2.RotX = Value : End Sub
Sub animLeftSling(Value) : LeftSlingBand.ShowFrame Value :  End Sub
Sub animRightSling(Value) : RightSlingBand.ShowFrame Value : End Sub
Sub animLock1(Value) : Kicker_LeftLock_Arm.Rotx = Value : End Sub
Sub animLock2(Value) : Kicker_RightLock_Arm.Rotx = Value : End Sub
Sub animPopper(Value) : Popper_Kicker.z = Value : End Sub
Sub animGunPlunger(Value): GunPlunger.TransY = Value : End Sub


'Drop Target Script
'DropTarget.z = -105	'full down
'DropTarget.z = -35  'normal

'droptarget.rotx = 5->8 or whatever memory hit

Sub animDTup(Value) : DropTarget.z = Value : End Sub
Sub animDTdown(Value) : DropTarget.z = Value : End Sub
Sub animDThit(Value) : DropTarget.RotX = Value : End Sub


SolCallback(28) = "SolDTUp" 'Drop Target	'dropup

'10.4 playsound args - name,loopcount,volume,pan,randompitch,pitch,UseExisting,Restart,Fade
Sub SolDTUp(aOn)
	if not aOn then Exit Sub
	playsound SoundFX("dropup",DOFDropTargets),0,1,Pan(DropTarget),0,0,0,0,Fade(DropTarget)
	aDTup.ModPoint 0, 0, DropTarget.z	'move start key of animation to wherever drop target currently is
	aDTup.Play
	DropTarget.Collidable = True
	'DropTargetu.Collidable = True
	Controller.Switch(77) = False
end Sub

'SolCallback(12) = "SolDTdown" '"knock down"
''This is a cut feature. Maybe one day a prototype rom will surface and this will do something.
'dim DropDownFeature : DropDownFeature = False	'enable cut solenoid12	(no point)
'Sub SolDTdown(aOn) 'this might not do anything. Cut feature...
'	tbsol12.text = "Sol12: " & aOn : tbsol12.TimerEnabled = 1 : tbSol12.TimerInterval = 1200
'	if not aOn or not DropDownFeature then Exit Sub
'	'playsound "dropdownsolenoid" 'todo
'	aDTdown.ModPoint 0, 0, DropTarget.z	'move start key of animation to wherever drop target currently is
'	aDTdown.Play
'	DropTarget.Collidable = False
'	'DropTargetu.Collidable = False
'	Controller.Switch(77) = True
'End Sub
'Sub tbSol12_Timer() : me.text = "" : me.TimerInterval = 0 : End Sub

Sub DropTarget_Hit() 'DT hit
	SkullHit BallSpeed(activeball) ' skull hit anim
	Set LastSwitch = DTpos 'Floating Text
	playsound SoundFX("dropdown",DOFDropTargets),0,1,Pan(DropTarget),0,0,0,0,Fade(DropTarget)
	aDThit.ModPoint 1, 33, PSlope(BallSpeed(activeball), 15, 2, 55, 10)
	aDTHit.Play
	aDTdown.ModPoint 0, 0, DropTarget.z	'move start key of animation to wherever drop target currently is
	aDTdown.Play
	DropTarget.Collidable = False
	'DropTargetu.Collidable = False
	'tb.text = "DT hit!"
	Controller.Switch(77) = True
End Sub

'Gun Script
'Messy, adapted from an older script

dim firepower : firepower = 850
Sub SolGunFire(aOn)
	Select Case aOn
		Case True	'todo firing anim
			aGunPlunger.Play
			mGunMagnet.Strength = -firepower	'fire power
			Playsound SoundFX("GunKick",DOFcontactors),0,1,Pan(GunMagnet),0,0,0,0,Fade(GunMagnet)
			if GunOn then Playsound "SolBallKick",0,1,0.1
		Case False	: 	mGunMagnet.Strength = 30 'hold ball power
	End Select
End Sub

Sub SolGun(aOn)
	Select Case aOn
		Case True : StopGun = False :  aGun.PlayLoop : if aGun.DebugOn then LGun3.state = 1
		Case False:	StopGun = True : if aGun.DebugOn then LGun3.state = 0
	End Select
End Sub

dim mGunMagnet : Set mGunMagnet = new cvpmMagnet
with mGunMagnet
	.InitMagnet GunMagnet, 30
	.GrabCenter = 1
	.MagnetOn = 1	'always on
	.Size = 125
End With
dim GunOn : GunOn = False
Sub GunMagnet_Hit()   : mGunMagnet.AddBall ActiveBall   : GunOn = True : Controller.Switch(31) = 1 : End Sub
Sub GunMagnet_UnHit() : mGunMagnet.Removeball ActiveBall : activeball.velz = 5: GunOn = False : Controller.Switch(31) = 0 : End Sub
agun.debugOn = False
'aGun.Play
dim GunSpeedCalc, StopGun
Sub animGun(Value)
	dim GunOutDebug : GunOutDebug = InterpolateGun(Value+1)
	if GunOutDebug = 3 then GunOutDebug = 1
	dim GunOut : GunOut = LinearEnvelope(GunOutDebug, array(1,2,3),array(0,78,0))	'0, 83, 0
	Gun.Rotz = GunOut
	GunPlunger.RotZ = GunOut
	FlipperL.StartAngle = (GunOut*-1)-14
	FlipperR.StartAngle = (GunOut*-1)+14
	'10.4 playsound args - name,loopcount,volume,pan,randompitch,pitch,UseExisting,Restart,Fade
	'SFX
	GunSpeedCalc = InterpolateGun(GunOutDebug) - GunSpeedCalc	'subtract from last position to find approx speed (for sfx pitch adjustment)
	Playsound SoundFX("RGunMotor",DOFgear), -1, 1, Pan(GunMagnet),0, PSlope(GunSpeedCalc*1000, 0,-500, 10,1000), 1, 0, Fade(GunMagnet)
	GunSpeedCalc = InterpolateGun(GunOutDebug)

	dim DebugSwitch : debugswitch = "??? " & GunOutDebug
	if GunOut < 2 then
		Controller.Switch(32) = False
		Controller.Switch(33) = True
		DebugSwitch = "Switch home"
	elseif GunOut < 47 then 	'<47	'Enables cannon gun
		Controller.Switch(32) = False
		Controller.Switch(33) = False
		DebugSwitch = "Departing..."
	elseif GunOut < 62 then 	'<62	'auto fires here if player doesn't (or tilt)
		Controller.Switch(32) = True
		Controller.Switch(33) = False
		DebugSwitch = "Switch midpoint"
	Else
		Controller.Switch(32) = False
		Controller.Switch(33) = False
		DebugSwitch = "Turning around"
	End If

	if StopGun then  StopSound SoundFX("RGunMotor",DOFgear): aGun.Pause' : Playsound SoundFX("T2GunMotorEnd",DOFgear) :

	if aGun.DebugOn then Lgun1.state = Controller.Switch(32) : Lgun2.state = Controller.Switch(33)
	if aGun.DebugOn then tbgun.text = "Val In " & Round(Value,2) & vbnewline & _
				"GunPos In " & Round(GunOutDebug,2) & vbnewline & _
				 DebugSwitch & vbnewline & _
				"GunOut " & Round(GunOut,2) & vbnewline & _
				 "gunpos interpolate: " & Round(GunOut,2) & vbnewline & _
				 "SpeedCalc " & Round(GunSpeedCalc*1000,2) & vbnewline & _
				" "
End Sub

Function InterpolateGun(input)	'Weird old part of script. This is all the 'acceleration' is, the linear keyframe is put through a polynomial
	dim x, xholdover : xholdover = 0
	x = input
	if x > 1 then
		xholdover = cint(x+0.5)-1
		x = x - cint(x+0.5)+1
	end if
'	InterpolateGun = -0.217469*x^3 + 1.10481*x^2 + 0.112656*x - 2.22045*10^-16	'Very low-end heavy
	'InterpolateGun = -1.94137*x^3 + 2.91206*x^2 + 0.0293147*x + 0	'Balanced but heavy	'1.02
	InterpolateGun = -1.97035*x^3 + 2.95552*x^2 + 0.0148269*x - 2.22045*(1E-16)	'Balanced heavier
	InterpolateGun = InterpolateGun + xholdover
	if InterpolateGun < 0 then InterpolateGun = 0
End Function


Sub FeedGun() 'debug
	kas -5, 15 : ka k, 327, 535
End Sub

'Other Solenoids

'Flippers
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(aOn)
	Select Case aOn'10.4 playsound args - name,loopcount,volume,pan,randompitch,pitch,UseExisting,Restart,Fade
		Case True :LF.fire : PlaySound SoundFX("FlipperR",DOFFlippers), 0, 1, Pan(LeftFlipper), 0.01, 0,0,0,Fade(LeftFlipper)
		Case Else :LeftFlipper.RotateToStart : PlaySound SoundFX("FlipperDown",DOFFlippers), 0, 0.01, Pan(LeftFlipper), 0, 0, 1, 0, Fade(LeftFlipper)
	End Select
End Sub

Sub SolRFlipper(aOn)
	Select Case aOn
		Case True : RF.Fire : PlaySound SoundFX("FlipperL",DOFFlippers), 0, 1, Pan(RightFlipper), 0.01, 0,0,0,Fade(RightFlipper)
		Case Else : RightFlipper.RotateToStart : PlaySound SoundFX("FlipperDown",DOFFlippers), 0, 0.01, Pan(RightFlipper), 0, 0, 1, 0, Fade(RightFlipper)
	End Select
End Sub



'Cloneballs and virtual kicker systems -

'LeftLock
dim cLeftLock : Set cLeftLock = new cCloneBall
With cLeftLock 			'Left Ball Cloner
	.Name = "cLeftLock"
	.Cache = Array(cbank1,cbank2,cbank3,cbank4)
	'.Destroyer = LeftLockKiller
	.Destroyer = BallKiller
	.DestroyerHeight = 0
	.ZOffset = -500
	.Debug = False
End With

Dim LeftLock : set LeftLock = New VirtualKicker

With LeftLock
	.Name = "LeftLock"	'important for callbacks
	'.Origin = Array(50.11972-500, 965.0734, 0-100)
	.Origin = Array(50.2, 960.85, 0-515)
	.SIze =40
	.InitExitBounds = Array(250,250,250)	'x y z auto exit bounds.
	.debug = False
End With

'Popper
dim cPopper : set cPopper = new cCloneBall
with cPopper
	.cache = array(cbank5, cbank6,cbank7,cbank8)
	.zOffset = -500
	.DestroyerHeight = 0
	.Destroyer = BallKiller
	.debug = False
	.name = "cPopper"
End With

dim Popper : set Popper = New VirtualKicker
With Popper
	.Name = "Popper"
	.Origin = Array(274.0037, 213.6675, -505)
	.KickOffset = True 'True by default - 'When regular kickers kick, they immediately move the ball up to the surface of the table first.
	.Size = 40
	.InitExitBounds = Array(150,170,120)	'x y z auto exit bounds.
	.Debug = False
End With

'TopLock
dim cTopLock : set cTopLock = New cCloneBall
with cTopLock
	.cache = Array(tBank1,tBank2,tBank3,tBank4)
	.zOffset = -500
	.DestroyerHeight = 0
	.Destroyer = BallKiller
	.debug = False
	.name = "cTopLock"
End With

dim TopLock : set TopLock = New VirtualKicker
With TopLock
	.Origin = Array(737.3464, 65.12592, -505)
	.Size = 37
	.InitExitBounds = Array(100,75,100)
	.Debug = False
	.Name = "TopLock"
End With



'Whole cloner trigger / event setup. Only one trigger is necessary if you use InitExitBounds
Sub CLeftTrigger_Hit()  cLeftLock.AddBall activeball : End Sub	'enter Cloner
Sub CLeftLock_Entry(aBall) : LeftLock.AddBall aBall: Set LastSwitch = LeftPopperPos : End Sub	'Cloner adds ball to virtual kicker (make sure trigger is within exitbounds!!!!)
Sub LeftLock_Exit(aBall) : cLeftLock.RemoveBall aBall : End Sub 'kicker Out of bounds, virtual kicker removes ball from cloner

Sub CPopperTrigger_Hit() : cPopper.Addball activeball : End Sub
Sub cPopper_Entry(aBall) : Popper.Addball aBall : Set LastSwitch = TopPopperPos :End Sub
Sub Popper_Exit(aBall) : cPopper.RemoveBall aBall : End Sub

Sub CTopTrigger_Hit() : cTopLock.AddBall activeball : End Sub
Sub cTopLock_Entry(aBall) : TopLock.AddBall aBall : Set LastSwitch = RightPopperPos : End Sub
Sub TopLock_Exit(aBall) : cTopLock.Removeball aBall : End Sub



'**************** Other sound functions ****************
Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallSpeed(ball) ^2 / 2000)
End Function

Function Pan(aObj) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp : tmp = aObj.x * 2 / Table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function PanX(x) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp : tmp = x * 2 / Table1.width-1
    If tmp > 0 Then
        Panx = Csng(tmp ^10)
    Else
        Panx = Csng(-((- tmp) ^10) )
    End If
End Function

Function Fade(aObj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp : tmp = aObj.y * 2 / table1.height-1
	If tmp > 0 Then
		Fade = Csng(tmp ^10)
	Else
		Fade = Csng(-((- tmp) ^10) )
	End If
End Function

Function FadeY(Y) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp : tmp = y * 2 / table1.height-1
	If tmp > 0 Then
		FadeY = Csng(tmp ^10)
	Else
		FadeY = Csng(-((- tmp) ^10) )
	End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = Ballspeed(ball) * 20
End Function
'new
Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = Pslope(BallSpeed(ball), 1, -1000, 60, 10000)
End Function
Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
    BallPitchV = Pslope(BallSpeed(ball), 1, -4000, 60, 7000)
End Function




'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR
dim RubbersD : Set RubbersD = new Dampener	'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False	'shows info in textbox "TBPout"
RubbersD.Print = False	'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935	'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96	'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967	'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64	'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener	'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False	'shows info in textbox "TBPout"
SleevesD.Print = False	'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'kas 354, 32.5 : ka k, 302.86, 1765.42



'dim TargetsD : Set TargetsD = new Dampener	'0.625 flat cor is pretty good by my track
'TargetsD.name = "Target"
'TargetsD.debugOn = True	'shows info in textbox "TBPout"
'TargetsD.Print = False	'debug, reports in debugger (in vel, out cor)
'
''for best results, try to match in-game velocity as closely as possible to the desired curve
'TargetsD.addpoint 0, 0, 1
'TargetsD.addpoint 1, 0.1, 1
'Targetsd.CopyCoef RubbersD, 0.8
'TargetsD.addpoint 1, 0.1, 1
'kas 40, 15 : ka k, 263.1, 1665
'kas 30, 15 : ka k, 192.73, 1116.5
'TargetsD.addpoint 1, 0.1, 1
'
'kas 5.5, 50 : ka k, 373.6 , 1759.9


Class Dampener
	Public Print, debugOn 'tbpOut.text
	public name, Threshold 	'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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

		aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
		if debugOn then TBPout.text = str
	End Sub

	Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
		dim x : for x = 0 to uBound(aObj.ModIn)
			addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
		Next
	End Sub


	Public Sub Report() 	'debug, reports all coords in tbPL.text
		if not debugOn then exit sub
		dim a1, a2 : a1 = ModIn : a2 = ModOut
		dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
		TBPout.text = str
	End Sub


End Class

'Tracks ball velocity for judging bounce calculations & angle
'apologies to JimmyFingers is this is what his script does. I know his tracks ball velocity too but idk how it works in particular
dim cor : set cor = New CoRTracker
cor.debugOn = False
'cor.update() - put this on a low interval timer
Class CoRTracker
	public DebugOn 'tbpIn.text
	public ballvel

	Private Sub Class_Initialize : redim ballvel(0) : End Sub
	'TODO this would be better if it didn't do the sorting every ms, but instead every time it's pulled for COR stuff
	Public Sub Update()	'tracks in-ball-velocity
		dim str, b, AllBalls, highestID : allBalls = getballs
		if uBound(allballs) < 0 then if DebugOn then str = "no balls" : TBPin.text = str : exit Sub else exit sub end if: end if
		for each b in allballs
			if b.id >= HighestID then highestID = b.id
		Next

		if uBound(ballvel) < highestID then redim ballvel(highestID)	'set bounds

		for each b in allballs
			ballvel(b.id) = BallSpeed(b)
			if DebugOn then
				dim s, bs 'debug spacer, ballspeed
				bs = round(BallSpeed(b),1)
				if bs < 10 then s = " " else s = "" end if
				str = str & b.id & ": " & s & bs & vbnewline
				'str = str & b.id & ": " & s & bs & "z:" & b.z & vbnewline
			end if
		Next
		if DebugOn then str = "ubound ballvels: " & ubound(ballvel) & vbnewline & str : if TBPin.text <> str then TBPin.text = str : end if
	End Sub
End Class



dim Roll : set roll = new RollingSounds	'a bit glitchy todo
'Rollingsounds object info
'.InitSFX 	-	String. Default sfx stem, required
'.DebugOn	-	Debug info on textbox 'tbroll'

'.Play		-	activeball, string. changes sfx stem to this string. Set to Empty to disable.rolling sounds.
'.Drop		-	Activeball, string (drop sound). Starts automatic drop sound handling.
'.Update	-	update on a -1 probs
With Roll
	.InitSFX = "tablerolling"
	.debugon = False
End With

Sub RightRampSound_Hit() : if activeball.vely < 0 then Roll.play activeball, "RampLoop" end if : End Sub	'up
Sub RightRampSound_UnHit() : if activeball.vely >= 0 then roll.play activeball, "tablerolling" end if : End Sub 'doiwn

Sub RightRampSound1_Hit() : playsound "WireRamp_Hit", 0, Vol(activeball), Pan(activeball),0.1,0,0,0,Fade(activeball) : roll.play activeball, "Wireloop" : End Sub

Sub RightRampDip_Hit() : playsound "WireRamp_Dip", 0, 0.25, Panx(870), 0.01, 0,1,0,Fade(activeball) : End Sub

Sub RightRampSound2_Hit()	'temp needs drop
	playsound "WireRamp_Stop", 0, Vol(activeball), Pan(activeball),0.1,0,0,0,Fade(activeball)
	Roll.Drop activeball, "BallDropTall"
End Sub
'
'Sub Trigger1_UnHit() : playsound "WireRamp_Hit", 0, 0.1, Pan(Trigger1), 0.01, 4000,1,1,Fade(activeball) : End Sub
'Sub Trigger2_UnHit() : playsound "WireRamp_Hit", 0, 0.1, Pan(Trigger2), 0.01, 7000,1,1,Fade(activeball) : End Sub


Sub LeftRampSound_Hit() : if activeball.vely < 0 then Roll.play activeball, "RampLoop" end if : End Sub	'up
Sub LeftRampSound_UnHit() : if activeball.vely >= 0 then roll.play activeball, "tablerolling" end if : End Sub 'doiwn
'10.4 playsound args - name,loopcount,volume,pan,randompitch,pitch,UseExisting,Restart,Fade
Sub LeftRampSound1_Hit() : playsound "WireRamp_Hit", 0, Vol(activeball), Pan(activeball),0.1,0,0,0,Fade(activeball)  : roll.play activeball, "Wireloop" : End Sub

'ub LeftRampDip_Hit() : playsound "WireRamp_Hit", 0, 0.3, Panx(121), 0.01, 4000,1,0,Fade(activeball) : End Sub

Sub LeftRampSound2_Hit()	'temp needs drop
	playsound "WireRamp_Stop", 0, Vol(activeball), Pan(activeball),0.1,0,0,0,Fade(activeball)
	Roll.Drop activeball, "BallDropTall"
	'Roll.play activeball, "BallDropTall"
End Sub

Sub VukRampSound1_Hit()
	Roll.play activeball, "WireRamp_Right"
End Sub
'Sub VukRampSound2_Hit() : Roll.play activeball, Empty :  Roll.Drop activeball, "BallDropTall" : End Sub
Sub VukRampSound2_Hit() : playsound "WireRamp_Stop", 0, Vol(activeball), Pan(activeball),0.1,0,0,0,Fade(activeball) : Roll.play activeball, Empty : End Sub

Sub StopAllRolling() 	'call this at table pause!!! TODO
	dim b : for b = 0 to 30
		StopSound("tablerolling" & b)
		StopSound("RampLoop" & b)
		StopSound("Wireloop" & b)
	next
end sub




'
'Playsound "tablerolling1", -1, 1, 0, 0, PSlope(35, 1, -1000, 60, 10000),1
'Playsound "RampLoop1", -1, 1, 0, 0, PSlope(35, 1, -4000, 60, 7000),1
'Playsound "RampLoop2", -1, 1, 0, 0, PSlope(60, 1, 18000, 60, 25000),1
'Playsound "WireLoop1", -1, 1, 0, 0, PSlope(60, 1, 4000, 60, 12000),1
'StopAllRolling

'
'Playsound "tablerolling1", -1, Vol(aBall)^0.8, Pan(aBall), 0, PSlope(ballspeed(aBall), 1, -1000, 60, 10000),1, 0,Fade(aBall)
'Playsound "RampLoop1", -1, Vol(aBall)^0.8, Pan(aBall), 0, PSlope(ballspeed(aBall), 1, -4000, 60, 7000),1, 0,Fade(aBall)
'Playsound "RampLoop2", -1, Vol(aBall)^0.8, Pan(aBall), 0, PSlope(ballspeed(aBall), 1, 18000, 60, 25000),1, 0,Fade(aBall)
'Playsound "WireLoop1", -1, Vol(aBall)^0.8, Pan(aBall), 0, PSlope(ballspeed(aBall), 1, 4000, 60, 12000),1, 0,Fade(aBall)
'
Function VolV(aball) : VolV = PSlope(BallSpeed(aBall), 0, 0, 60, 0.5) :  : End Function


Class RollingSounds	'a bit glitchy todo
	public DebugOn 'tbroll.text
	Private ballcount, LastBalls, SFX, Lock
	private DropFlag, DefaultSFX

	Private Sub Class_Initialize : ballcount = 0 : redim LastBalls(0) : redim SFX(30) : redim Lock(30): redim DropFlag(30):End Sub

	Public Sub PlaySoundGo(aBall, aIDX)	'mess with sounds here
		Select Case SFX(aIdx)	'this is extremely case sensitive
			Case "tablerolling" : Playsound(SFX(aIdx) & aIdx), -1, (VolV(aBall)), Pan(aBall), 0, PSlope(ballspeed(aBall), 1, -1000, 60, 10000),1, 0,Fade(aBall)
			Case "RampLoop"		: Playsound(SFX(aIdx) & aIdx), -1, Vol(aBall)^0.9, Pan(aBall), 0, PSlope(ballspeed(aBall), 1, -4000, 60, 7000),1, 0,Fade(aBall)
			Case "WireRamp_Right": playsound sfx(aIdx), 0, 1, Pan(aBall), 0.01, 0,1,0,Fade(aBall) ' : sfx(aIdx) = Empty
			'Case "RampLoop2"	: Playsound("RampLoop"& aIdx), -1, Vol(aBall)^0.9, Pan(aBall), 0, PSlope(ballspeed(aBall), 1, 18000, 60, 25000),1, 0,Fade(aBall)
			Case "Wireloop"		: Playsound(SFX(aIdx) & aIdx), -1, Vol(aBall)^0.9, Pan(aBall), 0, PSlope(ballspeed(aBall), 1, 4000, 60, 12000),1, 0,Fade(aBall)
			Case Empty : Exit Sub
			case else : msgbox "Rollingsounds PlaySoundGo sub: " & vbnewline & "no such sound??? :" & SFX(aIdx)
		End Select
	End Sub

	Public Sub Play(aBall, aStr)	'change ballrolling sfx string
		dim x : x = FindBallIdx(aBall)
		StopSound(SFX(x) & x)
		StopSound(SFX(x))
		SFX(x) = aStr : Lock(x) = False
	End Sub

	Private Function FindBallIdx(byval aBall)	'Returns 'GetBalls' index of arg ball
		dim gBalls : gBalls = GetBalls
		if uBound(gBalls) < ballcount then Update 'Exit Sub	'prevent glitching if this calls before update handles deleted balls
		dim x : for x = 0 to uBound(gBalls)
			if aBall.ID = gBalls(x).id then
				FindBallIdx = x
			End If
		Next
	End Function

	Public Property Let InitSFX(aStr) : DefaultSFX = aStr: End Property

	Public Sub Drop(aBall, aStr) :
		dim x : x = FindBallIdx(aBall)
		DropFlag(x) = aStr
	End Sub


	Public Sub Update()
		if uBound(getballs) < ballcount then 'catch deleted balls
			dim x : for x = uBound(getballs)+1 to ballcount
				stopsound(SFX(x) & x)
				'msgbox "stopsound" & x
			Next
		end If
		if uBound(getballs) > 0 then
			dim gballs :  gballs = getballs
			Ballcount = uBound(gballs)
			dim str : str = uBound(gballs) & " " & ballcount & vbnewline
			for x = 0 to uBound(gballs)
				if Not IsEmpty(DropFlag(x)) Then			'Drop handling
					'if gballs(x).velz < -3 then 'hm
					if gballs(x).z < ballsize+5 then 'hm
						playsound DropFlag(x), 0, 0.1, Pan(gballs(x)), 0, 0, 1, 0,Fade(gballs(x))	'adjust volume here
						StopSound(SFX(x) & x)
						DropFlag(x) = Empty : SFX(x) = Empty
						'msgbox gballs(x).velz
						str = str & "DropFlag " & x & " velz=" & gballs(x).velz & vbnewline
					end If
				End If

				if BallSpeed(gballs(x)) > 0.5 Then			'rolling sound handling
					if IsEmpty(SFX(x)) and gballs(x).z < ballsize + 5 then sfx(x) = DefaultSFX end if
					PlaySoundGo gballs(x), x : Lock(x) = False
					'Playsound (SFX(x) & x), -1, Vol(gballs(x))^0.8, Pan(gballs(x)), 0, BallPitch(gballs(x)), 1, 0,Fade(gballs(x)) : Lock(x) = False
				Else
					if not lock(x) Then Stopsound(SFX(x) & x) : Lock(x) = True
				end if
				str = str & x & ": " & SFX(x) & x & " lock:" & lock(x) & vbnewline
			Next
			if debugon then if tbroll.text <> str then tbroll.text = str  end if : end if
		end if
	End Sub

End Class







Sub Col_Posts_Hit(idx)
	RubbersD.dampen Activeball
	PlaySound RandomBand, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub
Sub Col_Pegs_Hit(idx)
	RubbersD.dampen Activeball
	PlaySound RandomPost, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub

Sub PostSleeves_Hit()
	SleevesD.Dampen activeball
	PlaySound RandomPost, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub
'10.4 playsound args - name,loopcount,volume,pan,randompitch,pitch,UseExisting,Restart,Fade
Sub Targets_Hit(idx)
	'TargetsD.dampen Activeball	'got 0.625 cor flat in tests
	PlaySound SoundFX("target",DOFTargets), 0, 0.2, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, Fade(Activeball)
	PlaySound SoundFX("targethit",0), 0, Vol(ActiveBall) , Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, Fade(Activeball)
End Sub

Sub Metals_lvl3_Hit() 'Inlanes
	playsound "MetalHit2", 0, Vol(activeball), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, Fade(Activeball)
End Sub

Sub MetalSFX_Hit(idx)
	playsound "metalhit_Medium", 0, Vol(activeball), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(Activeball)
End Sub

Sub LeftFlipper_Collide(parm)
	if ballspeed(activeball) > 1 then Playsound RandomFlipper, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(Activeball), 1, 0, Fade(Activeball)
end Sub

Sub RightFlipper_Collide(parm)
	if ballspeed(activeball) > 1 then Playsound RandomFlipper, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(Activeball), 1, 0, Fade(Activeball)
end Sub


'SFX string functions -SFX - Posts 1-5, Bands 1-4 and 11,22,33,44
Function RandomPost() : RandomPost = "Post" & rndnum(1,5) : End Function

Function RandomBand()
		dim x : x = rndnum(1,4)
		if BallSpeed(activeball) > 30 then
			RandomBand = "Rubber" & x & x	'ex. Playsound "Band44"
		else
			RandomBand = "Rubber" & x	'ex. Playsound "Band4"
		End If
End Function

'Flipper collide sound
Function RandomFlipper() : dim x : x = RndNum(1,3)  : RandomFlipper = "flip_hit_" & x : End Function

' Ball Collision Sound
Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0,Fade(ball1)
End Sub













'Physics debug helpers
'Physics debug helpers
'Physics debug helpers
'Physics debug helpers
'======================


'Sub PlungerTrigger_Hit() : if activeball.vely < 0 then TBPP.text = round(BallSpeed(activeball),2) : End If : End Sub

Sub TBPprint_Timer() : me.timerenabled = False : me.text = "" : end Sub
'www.shodor.org/interactivate/activities/SimplePlot/

dim tdir, tforce : tdir = 0 : tforce = 0
sub kas(aDir, aForce) : tdir = aDir : tforce = aForce : end Sub 'sets tdir and tforce

'kas 90, 0.1 : ka k, 426.8965, 1869 : coef=1
'kas 90, 0.15 : ka k, 426.8965, 1869 : coef=0.95	'shitty middle post slow
'kas 90, 0.15 : ka k, 309.4953, 502.0567 : coef=1	'drop test

'kas 270, 15 : ka k, 439.8656, 563.9827 : coef=1	'bumper kick
'kas 270, 15 : ka k, 460.7, 823.1 : coef=1	'1h4350

'Sub Bumpers_Hit(idx) : cortest activeball, "Bumper" : End Sub

sub ka(aObj, aX, aY) 'drop a ball from this specific position
	if not TypeName(aObj) = "Kicker" then exit sub
	aObj.createsizedballwithmass ballsize/2, ballmass
	dim b : set b = aObj.LastCapturedBall
	if aX <> 0 then b.x = aX end if : if aY <> 0 then b.y = aY end if
	b.ReflectionEnabled = False
	aObj.kick tDir, tForce
end Sub


sub kap(aDelay)	'flip after this MS delay, Left
	TBPF.text = "Ldelay " & aDelay
	k.createsizedballwithmass ballsize/2, ballmass
	dim b : set b = k.LastCapturedBall
	b.x = leftflipper.x+17 : b.y = 1740 : k.kick 0, 0
	SolLFlipper True

	nftimer.addtimer 2300, "SolLFlipper False'"
	nftimer.addtimer 2300+aDelay, "SolLFlipper True : exclamate TBPF :"
	nftimer.addtimer 2400+aDelay, "SolLFlipper False'"
End Sub

sub kapf(aDelay)	'flip after this MS delay, Right
	TBPF1.text = "Rdelay " & aDelay
	k.createsizedballwithmass ballsize/2, ballmass
	dim b : set b = k.LastCapturedBall
	b.x = RightFlipper.x-17 : b.y = 1740 : k.kick 0, 0
	SolRFlipper True

	nftimer.addtimer 2100, "SolRflipper False"
	nftimer.addtimer 2100+aDelay, "SolRFlipper True : exclamate TBPF1"
	nftimer.addtimer 2200+aDelay, "SolRFlipper False"
End Sub

Sub exclamate(aTB) : aTb.text = aTb.text & "!" : End Sub

'kapp 49 '=50 pretty random
'kapp 50.06 '=48
'kappf 71.5 'DT	vel 51?
'kapp 87.25 '% <50 (ramp) ill...
'kapp 94 '% =46 (late ramp)
'kapp 102.5 '% =41 (gate)

dim kappStart, kappPos
Sub Kapp(aPos)			'flip at this specific flipper position, left
	kappPos = aPos/100
	if IsEmpty(kappstart) then
		kappstart = gametime
		TBPF.text = "Lpos " & Round(aPos,3) & "%"
		k.createsizedballwithmass ballsize/2, ballmass
		dim b : set b = k.LastCapturedBall
		b.x = leftflipper.x+17 : b.y = 1740 : k.kick 0, 0
		SolLFlipper True
		nftimer.addtimer 2100, "SolLflipper False"
		kappt.Enabled = 1
	Else
		tbpf.text = "Lpos wait a minute"
	end If
End Sub
Sub kappt_timer()
	if gametime <= kappstart + 2000 Then
		tbpf.text = lf.pos
	Elseif gametime <= kappstart + 3000 Then
		if lf.pos >= kappPos then
			SolLflipper True
			nftimer.addtimer 100, "SolLflipper False"
			tbpf.text = round(lf.pos,2) & ">" & vbnewline & round(kapppos,2) & "!"
			kappstart = Empty : me.Enabled = 0
		end If
	Else
		kappstart = Empty : me.Enabled = 0
	end If
End Sub

dim kappfStart, kappfPos
Sub Kappf(aPos)	'put in a percentage flipper position, right
	kappfPos = aPos/100
	if IsEmpty(kappfstart) then
		kappfstart = gametime
		TBPF1.text = "Rpos " & round(aPos,3) & "%"
		k.createsizedballwithmass ballsize/2, ballmass
		dim b : set b = k.LastCapturedBall
		b.x = Rightflipper.x-17 : b.y = 1740 : k.kick 0, 0
		SolRFlipper True
		nftimer.addtimer 2100, "SolRflipper False"
		kapptf.Enabled = 1
	Else
		tbpf1.text = "Rpos wait a minute"
	end If
End Sub
Sub kapptf_timer()
	if gametime <= kappfstart + 2000 Then
		tbpf.text = lf.pos
	elseif gametime <= kappfstart + 4000 Then
		if rf.pos >= kappfPos then
			SolRflipper True
			nftimer.addtimer 100, "SolRflipper False"
			tbpf1.text = round(Rf.pos,2) & ">" & vbnewline & round(kappFpos,2) & "!"
			kappfstart = Empty : me.Enabled = 0
		end If
	Else
		kappfstart = Empty : me.Enabled = 0
	end If
End Sub




Class TimerEntry
	public command, starttime, endtime
	Public Sub Fire : execute command : Reset : end Sub
	public sub Reset : command = 0 : starttime = 0 : endtime = 0 : End Sub
End Class
dim NFtimer : set NFtimer = new cNFtimer
Class cNFtimer	'fast update timer object similar to cvpmtimer (used for debug only)
	Public PulseDuration	'pulseSw duration
	private queue(99), count
	private sub Class_Initialize : PulseDuration = 60 : count=0 : dim x : for x = 1 to uBound(queue) : set queue(x) = new TimerEntry : next : End Sub

	Public Sub AddTimer(aDelay, aStr)
		queue(Count+1).command = aStr : queue(count+1).StartTime = gametime : queue(count+1).EndTime = gametime + aDelay
		count = count + 1
	End Sub

	Public Sub PulseSw(aSw) : Controller.Switch(aSw) = 1 : AddTimer PulseDuration, "Controller.Switch(" & aSw & ") = 0" : End Sub

	Public Sub Update()
		if count then
			dim i : for i = 1 to 99
				if queue(i).EndTime > 0 and queue(i).EndTime <= gametime then queue(i).Fire  : count = count - 1
			next
		end If
		'debugger
	End Sub

	Private Sub Debugger
		dim str,x
		for x = 1 to uBound(queue)
			if queue(x).endtime <> 0 then str = str & x & ":" & queue(x).endtime - queue(x).starttime & " " & queue(x).command & vbnewline
		Next
		tbtimer.text = "count " & count & vbnewline & str	'resource heavy
	End Sub
End Class
TBTimer.Text = ""


'FlipperPolarity 0.10

'No longer includes overall speedhack because it was redundant with the velocity correction stuff.
'If you are experiencing flipper lag please rig up solenoids on a 1-interval timer, or use a flipper solenoid if possible!!


'Setup -
'Triggers tight to the flippers TriggerLF and TriggerRF. Timers as low as possible
'Debug box TBpl (for .debug = True)


dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity
InitPolarity
Sub InitPolarity()
	dim x, a : a = Array(LF, RF)
	for each x in a
		'safety coefficient (diminishes polarity correction only)
		'x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 0	'don't mess with these
		x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1	'disabled
		x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1

		x.enabled = True
		'x.DebugOn = True : stickL.visible = True : tbpl.visible = True : vpmSolFlipsTEMP.DebugOn = True
		x.TimeDelay = 44
	Next

'	'beta 1 quick
'	AddPt "Polarity", 0, 0.05, -3.77
'	AddPt "Polarity", 1, 0.12, -3.5
'	AddPt "Polarity", 2, 0.25, -3
'	AddPt "Polarity", 3, 0.65, -3
'	AddPt "Polarity", 4, 0.88, 0
'
'	addpt "Velocity", 0, 0, 1
'	addpt "Velocity", 1, 0.95, 0.968
'	addpt "Velocity", 2, 1.03, 0.945


	if OldFlippers then
		leftflipper.strength = 2500 : rightflipper.strength = leftflipper.strength
		'beta 2a1 - evened out by hand
		addpt "Velocity", 0, 0, 	1
		addpt "Velocity", 1, 0.327, 1.05
		addpt "Velocity", 2, 0.41, 	1.05
		addpt "Velocity", 3, 0.53, 	1'0.982
		addpt "Velocity", 4, 0.702, 0.968
		addpt "Velocity", 5, 0.95,  0.968
		addpt "Velocity", 6, 1.03, 	0.945


		AddPt "Polarity", 0, 0, 0
		AddPt "Polarity", 1, 0.16, -4
		AddPt "Polarity", 2, 0.326, -4.2
		AddPt "Polarity", 3, 0.36, -3.7
		AddPt "Polarity", 4, 0.41, -4.2
		AddPt "Polarity", 5, 0.65, -2.3
		AddPt "Polarity", 6, 0.71, -2
		AddPt "Polarity", 7, 0.785,-1.8
		AddPt "Polarity", 8, 0.88, 0
	Else
'		rf.report "Velocity"	'1.1
'		addpt "Velocity", 0, 0, 	1
'		addpt "Velocity", 1, 0.2, 	1.06
'		addpt "Velocity", 2, 0.34, 1.05
'		addpt "Velocity", 3, 0.41, 	1.05
'		addpt "Velocity", 4, 0.6, 	1.0'0.982
'		addpt "Velocity", 5, 0.702, 0.968
'		addpt "Velocity", 6, 0.95,  0.968
'		addpt "Velocity", 7, 1.03, 	0.945
'
'		rf.report "Polarity"
'		AddPt "Polarity", 0, 0, 0
'		AddPt "Polarity", 1, 0.16, -4.7
'		AddPt "Polarity", 2, 0.33, -5
'		AddPt "Polarity", 3, 0.37, -4.7	'4.2
'		AddPt "Polarity", 4, 0.41, -5
'		AddPt "Polarity", 5, 0.45, -4.7 '4.2
'		AddPt "Polarity", 6, 0.55,-4.7
'		AddPt "Polarity", 7, 0.66, -2.8'-2.1896
'		AddPt "Polarity", 8, 0.71, -2
'		AddPt "Polarity", 9, 0.81, -2
'		AddPt "Polarity", 10, 0.88, 0

		rf.report "Velocity"
		addpt "Velocity", 0, 0, 	1
		addpt "Velocity", 1, 0.2, 	1.07
		addpt "Velocity", 2, 0.41, 1.05
		addpt "Velocity", 3, 0.44, 1
		addpt "Velocity", 4, 0.65, 	1.0'0.982
		addpt "Velocity", 5, 0.702, 0.968
		addpt "Velocity", 6, 0.95,  0.968
		addpt "Velocity", 7, 1.03, 	0.945


		rf.report "Polarity"
		AddPt "Polarity", 0, 0, 0
		AddPt "Polarity", 1, 0.16, -4.7
		AddPt "Polarity", 2, 0.33, -5
		AddPt "Polarity", 3, 0.37, -4.7	'4.2
		AddPt "Polarity", 4, 0.41, -5.3
		AddPt "Polarity", 5, 0.45, -5 '4.2
		AddPt "Polarity", 6, 0.576,-4.7
		AddPt "Polarity", 7, 0.66, -2.8'-2.1896
		AddPt "Polarity", 8, 0.743, -1.5
		AddPt "Polarity", 9, 0.81, -1.5
		AddPt "Polarity", 10, 0.88, 0

	end if

	LF.Object = LeftFlipper
	LF.EndPoint = EndPointLp	'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
	RF.Object = RightFlipper
	RF.EndPoint = EndPointRp
End Sub

Sub AddPt(aStr, idx, aX, aY)	'debugger wrapper for adjusting flipper script in-game
	dim a : a = Array(LF, RF)
	dim x : for each x in a
		x.addpoint aStr, idx, aX, aY
	Next
End Sub


Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

















'Solenoids Part 3?

'Kicker Hit / Unhit Callbacks

'Sub Lamp1_Timer() : me.timerenabled = 0 : LeftLock.Kick 180, 20, 5 : End Sub
'Sub Lamp2_Timer() : me.timerenabled = 0 : Popper.Kick 0, 0, 52 : End Sub

SolCallback(1) = "SolPopper" 'Ball Popper

'Popper callbacks
'Sub Popper_Hit() : Controller.Switch(76) = 1 : Lamp2.visible = True	: End Sub 'debug
Sub Popper_Hit()
	Controller.Switch(76) = 1 	'10.4 playsound args - name,loopcount,volume,pan,randompitch,pitch,UseExisting,Restart,Fade
	playsound "safehousehit", 0, 1, Pan(zCOL_Gate4), 0.1, 0, 1,0, Fade(zCOL_Gate4)
End Sub
'Sub Popper_UnHit() : Controller.Switch(76) = 0 : Lamp2.Visible = False : End Sub	'debug
Sub Popper_UnHit() : Controller.Switch(76) = 0 : StopSound "safehousehit":  End Sub

Sub SolPopper(aOn)
	if not aOn then Exit Sub
	'Popper.Kick 0, 0, 52
	Popper.Kick 0, 0, 45
	aPopper.Play
	Select Case Popper.State	'switch state
		Case True
			playsound SoundFX("leftpopperoutnew",DOFcontactors), 0, 1, Pan(zCOL_Gate4), 0.1, 0, 1, 0, Fade(zCOL_Gate4)
		Case False:	playsound SoundFX("leftpopper",DOFcontactors), 0, 1, Pan(zCOL_Gate4), 0.1, 0, 1, 0, Fade(zCOL_Gate4)
	End Select
End Sub

SolCallback(10) = "SolLock2" 'Top Lock

Sub TopLock_Hit() '10.4 playsound args - name,loopcount,volume,pan,randompitch,pitch,UseExisting,Restart,Fade
	playsound "safehousehit", 0, 1, PanX(730), 0.1, 0, 1, 0, FadeY(70)
	controller.Switch(55) = 1' : lamp3.visible = true
End Sub
'Sub TopLock_UnHit() : controller.Switch(55) = 0 : Lamp3.Visible = False : End Sub
Sub TopLock_UnHit() : controller.Switch(55) = 0 : StopSound "safehousehit" :  End Sub

Sub SolLock2(aOn)
	if not aOn then Exit Sub
	aLock2.Play
	'TopLock.Kick 262, 20, 5
	TopLock.Kick 268, 15, 8
	Select Case False 'TopLock.State	'switch state	TODO
		Case True:	playsound SoundFX("Kicker_Release",DOFcontactors), 0, 1, PanX(700), 0.1, 0, 1, 0, FadeY(100)
		Case False:	playsound SoundFX("LeftEject",DOFcontactors), 0, 1, PanX(700), 0.1, 0, 1, 0, FadeY(100)
	End Select
End Sub

SolCallback(16) = "SolLock1" 'Left Lock

Sub LeftLock_Hit()
	Controller.Switch(51) = 1' : Lamp1.visible = True	'debug
	playsound "safehousehit", 0, 1, PanX(56), 0.1, 0, 1, 0, FadeY(963)
End Sub
Sub LeftLock_UnHit()
	Controller.Switch(51) = 0' : Lamp1.Visible = False	'debug
	StopSound "safehousehit"
End Sub

Sub SolLock1(aOn)
	if not aOn then Exit Sub
	aLock1.Play
	'LeftLock.Kick 180, 20, 5
	LeftLock.Kick 175, 15, 8
	Select Case LeftLock.State	'switch state
		Case True:	playsound SoundFX("Kicker_Release",DOFContactors), 0, 1, PanX(86), 0.1, 0, 1, 0, FadeY(1000)
		Case False:	playsound SoundFX("LeftEject",DOFContactors), 0, 1, PanX(86), 0.1, 0, 1, 0, FadeY(1000)
	End Select
End Sub


SolCallback(2) = "SolGunFire" 'Gun Kicker

SolCallback(5) = "SolRightSling" 'Left Sling
SolCallback(6) = "SolLeftSling" 'Right Sling

Sub LeftSlingShot_SlingShot() : vpmtimer.PulseSw 44 : if not SlingOn then aLeftSlingArm.play : aLeftSling.play : _
Playsound SoundFX("LeftSling",DOFcontactors),0,1,Panx(200),0.05,0,0,0,FadeY(1600) : Set lastswitch = LeftSlingPos : end if : End Sub
Sub RightSlingShot_SlingShot() : VpmTimer.PulseSw 45 :if not SlingOn then aRightSlingArm.play : aRightSling.play : _
Playsound SoundFX("LeftSling",DOFcontactors),0,1,Panx(650),0.05,0,0,0,FadeY(1600) : Set lastswitch = RightSlingPos : end if : End Sub

Sub SolLeftSling(aOn) : if aOn and SlingOn then aLeftSlingArm.play : aLeftSling.play : _
Playsound SoundFX("LeftSling",DOFcontactors),0,1,Panx(200),0.05,0,0,0,FadeY(1600) : end if : End Sub
Sub SolRightSling(aOn) : if aOn and SlingOn then aRightSlingArm.play : aRightSling.play : _
Playsound SoundFX("LeftSling",DOFcontactors),0,1,Panx(650),0.05,0,0,0,FadeY(1600) : end if : End Sub


Dim SlingOn : SlingOn = True
Sub SlingArea_Hit() : SlingOn = False : End Sub
Sub SlingArea_UnHit() : if me.BallCntOver < 1 then SlingOn = True end if : End Sub

SolCallback(7) = "SolKnocker" : Sub SolKnocker(aOn) : if aOn then Playsound SoundFX("Knocker",DOFknocker) End If: End Sub 'Knocker



Dim KickBackEmpty : KickBackEmpty = True
Sub KickBackArea_Hit() : KickBackEmpty = False : End Sub
Sub KickBackArea_UnHit() : if me.BallCntOver < 1 then KickBackEmpty = True end if : End Sub
SolCallback(8) = "SolKickBack" 'Kickback
Sub SolKickBack(aOn)
	if aOn then
		KickBack.Fire
		Playsound SoundFX("Kickback",DOFcontactors),0,1,Pan(KickBack),0.05,0,0,0,Fade(KickBack) : if not KickBackEmpty then Playsound "SolBallKick",0,1,-0.1
	Else : KickBack.PullBack
	End If
End Sub

Dim PlungerEmpty : PlungerEmpty = True
Sub PlungerArea_Hit() : PlungerEmpty = False : End Sub
Sub PlungerArea_UnHit() : if me.BallCntOver < 1 then PlungerEmpty = True end if : End Sub
SolCallback(9) = "SolPlunger" : 	 'Plunger
Sub SolPlunger(aOn)
	if aOn then
		Plunger.Fire
		Playsound SoundFX("AutoPlunge",DofContactors),0,1,Pan(Plunger),0.05,0,0,0,Fade(Plunger) : if not PlungerEmpty then Playsound "SolBallKick",0,1,0.1
	Else  : Plunger.PullBack
	End If
End Sub
Plunger.Pullback
Kickback.Pullback

SolCallback(11) = "SolGun" 'Gun Motor



SolCallback(13) = "SolBumper1"
SolCallback(14) = "SolBumper2" 'Right Jet
SolCallback(15) = "SolBumper3" 'Bottom Jet

Dim BumpOn : BumpOn = True	'Allow for controller-driven bumpers if the ball is out of the way
Sub BumperArea_Hit() : BumpOn = False : End Sub'10.4 playsound args - name,loopcount,volume,pan,randompitch,pitch,UseExisting,Restart,Fade
Sub BumperArea_UnHit() : if me.BallCntOver < 1 then BumpOn = True end if: End Sub
Sub SolBumper1(aOn) : if aOn and BumpOn then aBumper1.play : _
Playsound SoundFX("LeftJet (2)",DOFcontactors),0,1,Pan(Bumper1),0.05,0,0,0,Fade(Bumper1) : end if : End Sub
Sub SolBumper2(aOn) : if aOn and BumpOn then aBumper2.play : _
Playsound SoundFX("RightJet (2)",DOFcontactors),0,1,Pan(Bumper2),0.05,0,0,0,Fade(Bumper2) : end if : End Sub
Sub SolBumper3(aOn) : if aOn and BumpOn then aBumper3.play : _
Playsound SoundFX("BottomJet",DOFcontactors),0,1,Pan(Bumper3),0.05,0,0,0,Fade(Bumper3) : end if : End Sub

Sub Bumper1_Hit() : vpmTimer.PulseSw 41 : if not BumpOn then aBumper1.play : _
Playsound SoundFX("LeftJet (2)",DOFcontactors),0,1,Pan(Bumper1),0.05,0,0,0,Fade(Bumper1) : end if : End Sub
Sub Bumper2_Hit() : vpmTimer.PulseSw 42 : if not BumpOn then aBumper2.play : _
Playsound SoundFX("RightJet (2)",DOFcontactors),0,1,Pan(Bumper2),0.05,0,0,0,Fade(Bumper2) : end if : End Sub
Sub Bumper3_Hit() : vpmTimer.PulseSw 43 : if not BumpOn then aBumper3.play : _
Playsound SoundFX("BottomJet",DOFcontactors),0,1,Pan(Bumper3),0.05,0,0,0,Fade(Bumper3) : end if : End Sub



'Timerless Trough handling

'Switches
'sw11 Left Flipper
'sw12 Right Flipper
'sw13 Start Button
'sw14 tilt

SolCallback(3) = "TroughSolIn" 'Outhole	'troughin
SolCallback(4) = "TroughSolOut" 'Trough	'troughout


dim TKickers, TSwitches
TKickers = Array(Drain, sw15, sw16, sw17)  '4 ball Trough
TSwitches = Array(18, 15, 16, 17)

Sub TroughSFX_Hit(): Playsound "Trough2", 0, 0.2, 0: End Sub
Sub TroughSFX_UnHit(): StopSound "Trough2": End Sub

Sub TroughSolIn(aEnabled) :
	Select Case aEnabled
		Case True
			TGate.Collidable = False
			if TKickers(0).ballcntover then Playsound SoundFX("Trough1",DOFcontactors), 0, 0.5, 0.005 : Else playsound SoundFX("LeftEject",DOFcontactors), 0, 0.5, 0.005 : End If
			TKickers(0).Kick 60, 15 : UpdateTrough
		Case False: TGate.Collidable = True
	End Select
End Sub

Sub TroughSolOut(aEnabled)
	Select Case aEnabled
		Case True
			if TKickers(uBound(TKickers)).BallCntOver then
				playsound SoundFX("BallRelease",DOFcontactors), 0, 0.5, 0.03
			Else
				playsound SoundFX("LeftEject",DOFcontactors), 0, 0.5, 0.01
			End If
			TKickers(uBound(TKickers)).Kick 60, 5, 4
			Controller.switch(TSwitches(uBound(TSwitches))) = 0 : dEBUGtROUGHsWITCHES
		Case False : VPMTIMER.ADDTIMER 300, "UpdateTroughSwitches'"	'well almost timerless
	End Select
End Sub

'sw15, sw18, sw17, sw16

sw15.createsizedballwithmass ballsize/2, ballmass
sw16.createsizedballwithmass ballsize/2, ballmass
sw17.createsizedballwithmass ballsize/2, ballmass


Sub UpdateTrough()
'	if sw16.BallCntOver = 0 then sw17.kick 60,15
'	if sw17.BallCntOver = 0 then sw18.Kick 60,15
	if Tkickers(uBound(TKickers)).BallCntOver = 0 then Tkickers(uBound(TKickers)-1).kick 60,5
	if Tkickers(uBound(TKickers)-1).BallCntOver = 0 then Tkickers(uBound(TKickers)-2).kick 60,5
'	if sw17.BallCntOver = 0 then sw18.Kick 60,15
'	dim x : for x = 1 to uBound(TKickers) step -1
'		if Tkickers(x).BallCntOver = 0 then Tkickers(x-1).kick 60, 15
'	Next
End Sub
dim TroughFull : TroughFull = False
Sub UpdateTroughSwitches()
	dim count, x  : count = 0
	for x = 0 to uBound(TKickers)
		if TKickers(x).ballcntover then controller.Switch(TSwitches(x)) = 1 : count = count + 1: else : controller.Switch(TSwitches(x)) = 0 : end if
	next
	if count = uBound(TKickers)+1 then TKickers(0).destroyball : controller.Switch(TSwitches(0)) = 0
	UpdateTrough
	dEBUGtROUGHsWITCHES

End Sub

sUB dEBUGtROUGHsWITCHES()
	dim debugon
	'debugon = True : tbt.visible = 1 : apron.visible = 0	'debug
	dim str, count, x  : count = 0
	for x = 0 to uBound(TKickers)
		str = str & "Trough Slot " & x & ":" & TKickers(x).ballcntover & "sw:" & TSwitches(x) & ": " & controller.Switch(TSwitches(x) ) & vbnewline
	next
	if debugon then tbt.text = str & gametime
eND SUB

Sub drain_Hit()
	'Playsound "Drain", 0, 0.5
	Set lastswitch = TroughPos	'update floating text
	UpdateTroughSwitches
End Sub
Sub sw17_Hit() : UpdateTroughSwitches : End Sub
Sub sw16_Hit() : UpdateTroughSwitches : End Sub
Sub sw15_Hit() : UpdateTroughSwitches : activeball.mass = ballmass: End Sub






''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''
''''''''''Switches  ''''''''''''''''
''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''

'sw21 Slam Tilt
'sw22 Coin Door Closed
'sw23 Ticket Dispenser (??)
	'sw24 Test Position, Always Closed
	'sw25 Left Outlane
	'sw26 Left Return Lane (inlane Left)
	'sw27 Right Return Lane (inlane Right)
	'sw28 Right Outlane
Sub sw25_Hit():controller.Switch(25) = 1 : End Sub
Sub sw26_Hit():controller.Switch(26) = 1 : End Sub
Sub sw27_Hit():controller.Switch(27) = 1 : End Sub
Sub sw28_Hit():controller.Switch(28) = 1 : End Sub
Sub sw25_UnHit():controller.Switch(25) = 0 : End Sub
Sub sw26_UnHit():controller.Switch(26) = 0 : End Sub
Sub sw27_UnHit():controller.Switch(27) = 0 : End Sub
Sub sw28_UnHit():controller.Switch(28) = 0 : End Sub
	'sw31 Gun Loaded
	'sw32 Gun Mark
	'sw33 Gun Home
	'sw34 Grip Trigger
'sw35 (Not Used)
	'sw36 Mid Left Stand-up Target
	'sw37 Mid Center Stand-up Target
	'sw38 Mid Right Stand-up Target
Sub sw36_Hit() :  vpmTimer.PulseSw 36 : End Sub
Sub sw37_Hit() :  vpmTimer.PulseSw 37  : End Sub
Sub sw38_Hit() :  vpmTimer.PulseSw 38  : End Sub




	'sw41 Left Jet
	'sw42 Right Jet
	'sw43 Bottom Jet
	'sw44 Left Sling
	'sw45 Right Sling
	'sw46 Top Right Stand-up Target
	'sw47 Mid Right Stand-up Target
	'sw48 Bot Right Stand-up Target
Sub sw46_Hit() : vpmTimer.PulseSw 46 : End Sub
Sub sw47_Hit() : vpmTimer.PulseSw 47 : End Sub
Sub sw48_Hit() : vpmTimer.PulseSw 48 : End Sub

'sw51 Left Lock
'sw52 (Not Used)
	'sw53 Low Escape Route
	'sw54 High Escape Route
Sub sw53_Hit():controller.Switch(53) = 1 : End Sub
Sub sw54_Hit():controller.Switch(54) = 1 : End Sub
Sub sw53_UnHit():controller.Switch(53) = 0 : End Sub
Sub sw54_UnHit():controller.Switch(54) = 0 : End Sub
'sw55 Top Lock
	'sw56 Top Lane Left
	'sw57 Top Lane Center
	'sw58 Top Lane Right
Sub sw56_Hit():controller.Switch(56) = 1 : End Sub
Sub sw57_Hit():controller.Switch(57) = 1 : End Sub
Sub sw58_Hit():controller.Switch(58) = 1 : End Sub
Sub sw56_UnHit():controller.Switch(56) = 0 : End Sub
Sub sw57_UnHit():controller.Switch(57) = 0 : End Sub
Sub sw58_UnHit():controller.Switch(58) = 0 : End Sub


	'sw61 Left Ramp Entry
	'sw62 Left Ramp Made
	'sw63 Right Ramp Entry
	'sw64 Right Ramp Made

Sub Sw61_Hit() : VpmTimer.PulseSw 61 : End Sub
Sub Sw62_Hit() : VpmTimer.PulseSw 62 : End Sub
Sub Sw63_Hit() : VpmTimer.PulseSw 63 : End Sub
Sub Sw64_Hit() : VpmTimer.PulseSw 64 : End Sub

	'sw65 Low Chase Loop
	'sw66 High Chase Loop
Sub sw65_Hit():controller.Switch(65) = 1 : End Sub
Sub sw66_Hit():controller.Switch(66) = 1 : End Sub
Sub sw65_UnHit():controller.Switch(65) = 0 : End Sub
Sub sw66_UnHit():controller.Switch(66) = 0 : End Sub
'sw67 (Not Used)
'sw68 (Not Used)

	'sw71 Target 1 High
	'sw72 Target 2
	'sw73 Target 3
	'sw74 Target 4
	'sw75 Target 5 Low

Sub sw71_Hit() : SweepLeftTargets 0, Activeball : End Sub
Sub sw72_Hit() : SweepLeftTargets 1, Activeball : End Sub
Sub sw73_Hit() : SweepLeftTargets 2, Activeball : End Sub
Sub sw74_Hit() : SweepLeftTargets 3, Activeball : End Sub
Sub sw75_Hit() : SweepLeftTargets 4, Activeball : End Sub


Sub SweepLeftTargets(aIDX, aBall)
	dim aSw : aSw = array(71, 72, 73, 74, 75)
	dim aTg : aTg = array(sw71, sw72, sw73, sw74, Sw75)
	dim Tolerance : Tolerance = 10	'Tolerance, in VP units
	dim Midpoint, BallY : BallY = aBall.Y
	'dim str	'debug
	vpmTimer.PulseSw aSw(aIdx)
	if aIDX > LBound(atG) and BallY <= aTg(aIdx).Y then 'check higher (-) than hit target ...
		midpoint = ((aTg(aIdx).Y + aTg(aIdx-1).Y) / 2) - 0.2236 '0.2236 tangent correction
			'str = "lower than idx : " & aidx & vbnewline
		If BallY <= midpoint + Tolerance AND BallY >= midpoint - Tolerance Then
			'str = str & "sweeping " & aidx & " " & aIdx-1
			 vpmTimer.PulseSw aSw(aIdx-1)
		End If
	elseif aIDX < UBound(atG) and BallY >= aTg(aIdx).Y then 'check lower (+) than hit target ...
		midpoint = ((aTg(aIdx).Y + aTg(aIdx+1).Y) / 2) - 0.2236
		'str = "higher than idx : " & aidx
		If BallY <= midpoint + Tolerance AND BallY >= midpoint - Tolerance Then
			'str = str & "sweeping " & aidx & " " & aIdx+1
			 vpmTimer.PulseSw aSw(aIdx+1)
		End If
	Else
'		str = "error, idx " & aidx & vbnewline
'		dim a, b,c,d
'		a = aIdx > uBound(atg) : b = BallY <= aTg(aIdx).Y
'		c = aIdx < uBound(atg) : d = BallY >= aTg(aIdx).Y
'		str = str & a & " " & b & vbnewline & c & " " & d
	end if
	'tbsweep.text = str : tbsweep.timerenabled = 1
End Sub
Sub tbsweep_TImer() : me.text = Empty : End Sub	'just clearing debug box



'sw76 Ball Popper
	'sw77 Drop Target
	'sw78 Shooter
Sub sw78_Hit():controller.Switch(78) = 1 : End Sub
Sub sw78_UnHit():controller.Switch(78) = 0 : End Sub









'====================
'Class jungle nf
'=============

'No-op object instead of adding more conditionals to the main loop
'It also prevents errors if empty lamp numbers are called, and it's only one object
'should be g2g?

'Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class

'version 0.11 - Mass Assign, Changed modulate style
'version 0.12 - Update2 (single -1 timer update) update method for core.vbs
'Version 0.12a - Filter can now be accessed via 'FilterOut'

Class LampFader
	Public FadeSpeedDown(140), FadeSpeedUp(140)
	Private Lock(140), Loaded(140), OnOff(140)
	Public UseFunction
	Private cFilter
	Private UseCallback(140), cCallback(140)
	Public Lvl(140), Obj(140)
	Private Mult(140)
	'Public FrameTime
	Private InitFrame

	Sub Class_Initialize()
		InitFrame = 0
		dim x : for x = 0 to uBound(OnOff) 	'Set up fade speeds
			FadeSpeedDown(x) = 1/100	'fade speed down
			FadeSpeedUp(x) = 1/80		'Fade speed up
			UseFunction = False
			lvl(x) = 0
			OnOff(x) = False
			Lock(x) = True : Loaded(x) = False
			Mult(x) = 1
		Next

		for x = 0 to uBound(OnOff) 		'clear out empty obj
			if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
		Next
	End Sub

	Public Property Get Locked(idx) : Locked = Lock(idx) : End Property		'debug.print Lampz.Locked(100)	'debug
	Public Property Get Load(idx) : Load = Loaded(idx) : End Property		'debug.print Lampz.Load(100)	'debug
	Public Property Get state(idx) : state = OnOff(idx) : end Property
	Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
	Public Function FilterOut(aInput) : if UseFunction Then FilterOut = cFilter(aInput) Else FilterOut = aInput End If : End Function
	Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property 'execute
	'Public Property Let Callback(idx, String) : Set cCallback(idx) = GetRef(String) : UseCallBack(idx) = True : End Property

	Public Property Let state(ByVal idx, input) 'Major update path
		if Input <> OnOff(idx) then  'discard redundant updates
			OnOff(idx) = input
			Lock(idx) = False
			Loaded(idx) = False
		End If
	End Property

	'Mass assign, Builds arrays where necessary
	Sub MassAssign(aIdx, aInput)
		If typename(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
			if IsArray(aInput) then
				obj(aIdx) = aInput
			Else
				Set obj(aIdx) = aInput
			end if
		Else
			Obj(aIdx) = AppendArray(obj(aIdx), aInput)
		end if
	end Sub

	Sub SetLamp(aIdx, aOn) : state(aIdx) = aOn : End Sub	'Solenoid Handler

	Public Sub TurnOnStates()	'If obj contains any light objects, set their states to 1 (Fading is our job!)
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
	Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub	'turn state to 1

	Public Sub Init()	'Just runs TurnOnStates right now
		TurnOnStates
	End Sub

	Public Property Let Modulate(aIdx, aCoef) : Mult(aIdx) = aCoef : Lock(aIdx) = False : Loaded(aIdx) = False: End Property
	Public Property Get Modulate(aIdx) : Modulate = Mult(aIdx) : End Property

	Public Sub Update1()	 'Handle all boolean numeric fading. If done fading, Lock(x) = True. Update on a '1' interval Timer!
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

	Public Sub Update2()	 'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
		'FrameTime = gametime - InitFrame : InitFrame = GameTime	'Calculate frametime
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

	Public Sub Update()	'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
		dim x,xx : for x = 0 to uBound(OnOff)
			if not Loaded(x) then
				if IsArray(obj(x) ) Then	'if array
					If UseFunction then
						for each xx in obj(x) : xx.IntensityScale = cFilter(Lvl(x)*Mult(x)) : Next
					Else
						for each xx in obj(x) : xx.IntensityScale = Lvl(x)*Mult(x) : Next
					End If
				else						'if single lamp or flasher
					If UseFunction then
						obj(x).Intensityscale = cFilter(Lvl(x)*Mult(x))
					Else
						obj(x).Intensityscale = Lvl(x)
					End If
				end if
				If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x))	'Callback (execute)
				'If UseCallBack(x) then cCallback(x)(Lvl(x))	'Callback
				If Lock(x) Then
					if Lvl(x) = 1 or Lvl(x) = 0 then Loaded(x) = True	'finished fading
				end if
			end if
		Next
	End Sub
End Class


'version 0.11 - Mass Assign, Changed modulate style
'version 0.12 - Update2 (single -1 timer update) update method for core.vbs
'Version 0.12a - Filter can now be publicly accessed via 'FilterOut'

Class DynamicLamps 'Lamps that fade up and down. GI and Flasher handling
	Public Loaded(50), FadeSpeedDown(50), FadeSpeedUp(50)
	Public Lock(50), SolModValue(50)
	Private UseCallback(50), cCallback(50)
	Public Lvl(50)
	Public Obj(50)
	Private UseFunction, cFilter
	private Mult(50)

	'Public FrameTime
	Private InitFrame

	Private Sub Class_Initialize()
		InitFrame = 0
		dim x : for x = 0 to uBound(Obj)
			FadeSpeedup(x) = 0.01
			FadeSpeedDown(x) = 0.01
			lvl(x) = 0.0001 : SolModValue(x) = 0
			Lock(x) = True : Loaded(x) = False
			mult(x) = 1
			if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
		next
	End Sub

	Public Property Get Locked(idx) : Locked = Lock(idx) : End Property
	Public Property Get Load(idx) : Load = Loaded(idx) : End Property		'debug.print flashers.Load(100)	'debug
	Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property 'execute
	'Public Property Let Callback(idx, String) : Set cCallback(idx) = GetRef(String) : UseCallBack(idx) = True : End Property
	Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
	Public Function FilterOut(aInput) : if UseFunction Then FilterOut = cFilter(aInput) Else FilterOut = aInput End If : End Function

	Public Property Let State(idx,Value)
		'If Value = SolModValue(idx) Then Exit Property ' Discard redundant updates
		If Value <> SolModValue(idx) Then ' Discard redundant updates
			SolModValue(idx) = Value
			Lock(idx) = False : Loaded(idx) = False
		End If
	End Property
	Public Property Get state(idx) : state = SolModValue(idx) : end Property

	'Mass assign, Builds arrays where necessary
	Sub MassAssign(aIdx, aInput)
		If typename(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
			if IsArray(aInput) then
				obj(aIdx) = aInput
			Else
				Set obj(aIdx) = aInput
			end if
		Else
			Obj(aIdx) = AppendArray(obj(aIdx), aInput)
		end if
	end Sub

	'solcallback (solmodcallback) handler
	Sub SetLamp(aIdx, aInput) : state(aIdx) = aInput : End Sub	'0->1 Input
	Sub SetModLamp(aIdx, aInput) : state(aIdx) = aInput/255 : End Sub	'0->255 Input
	Sub SetGI(aIdx, ByVal aInput) : if aInput = 8 then aInput = 7 end if : state(aIdx) = aInput/7 : End Sub	'0->8 WPC GI input

	Public Sub TurnOnStates()	'If obj contains any light objects, set their states to 1 (Fading is our job!)
		dim debugstr
		dim idx : for idx = 0 to uBound(obj)
			if IsArray(obj(idx)) then
				'debugstr = debugstr & "array found at " & idx & "..."
				dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
				for x = 0 to uBound(tmp)
					if typename(tmp(x)) = "Light" then DisableState tmp(x) ': debugstr = debugstr & tmp(x).name & " state'd" & vbnewline

				Next
			Else
				if typename(obj(idx)) = "Light" then DisableState obj(idx) ': debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline

			end if
		Next
		'debug.print debugstr
	End Sub
	Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub	'turn state to 1

	Public Sub Init()	'just call turnonstates for now
		TurnOnStates
	End Sub

	Public Property Let Modulate(aIdx, aCoef) : Mult(aIdx) = aCoef : Lock(aIdx) = False : Loaded(aIdx) = False: End Property
	Public Property Get Modulate(aIdx) : Modulate = Mult(aIdx) : End Property

	Public Sub Update1()	 'Handle all numeric fading. If done fading, Lock(x) = True
		'dim stringer
		dim x : for x = 0 to uBound(Lvl)
			'stringer = "Locked @ " & SolModValue(x)
			if not Lock(x) then 'and not Loaded(x) then
				If lvl(x) < SolModValue(x) then '+
					'stringer = "Fading Up " & lvl(x) & " + " & FadeSpeedUp(x)
					Lvl(x) = Lvl(x) + FadeSpeedUp(x)
					if Lvl(x) >= SolModValue(x) then Lvl(x) = SolModValue(x) : Lock(x) = True
				ElseIf Lvl(x) > SolModValue(x) Then '-
					Lvl(x) = Lvl(x) - FadeSpeedDown(x)
					'stringer = "Fading Down " & lvl(x) & " - " & FadeSpeedDown(x)
					if Lvl(x) <= SolModValue(x) then Lvl(x) = SolModValue(x) : Lock(x) = True
				ElseIf Lvl(x) = 0 then
					Lock(x) = True
				End If
			end if
		Next
		'tbF.text = stringer
	End Sub

	Public Sub Update2()	 'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
		'FrameTime = gametime - InitFrame : InitFrame = GameTime	'Calculate frametime
		dim x : for x = 0 to uBound(Lvl)
			if not Lock(x) then 'and not Loaded(x) then
				If lvl(x) < SolModValue(x) then '+
					Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
					if Lvl(x) >= SolModValue(x) then Lvl(x) = SolModValue(x) : Lock(x) = True
				ElseIf Lvl(x) > SolModValue(x) Then '-
					Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
					if Lvl(x) <= SolModValue(x) then Lvl(x) = SolModValue(x) : Lock(x) = True
				ElseIf Lvl(x) = 0 then
					Lock(x) = True
				End If
			end if
		Next
		Update
	End Sub

	Public Sub Update()	'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
		dim x,xx
		for x = 0 to uBound(Lvl)
			if not Loaded(x) then
				if IsArray(obj(x) ) Then	'if array
					If UseFunction then
						for each xx in obj(x) : xx.IntensityScale = cFilter(Lvl(x)*mult(x)) : Next
					Else
						for each xx in obj(x) : xx.IntensityScale = Lvl(x)*mult(x) : Next
					End If
				else						'if single lamp or flasher
					If UseFunction then
						obj(x).Intensityscale = cFilter(Lvl(x)*mult(x))
					Else
						obj(x).Intensityscale = Lvl(x)*mult(x)
					End If
				end if
				If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)*mult(x))	'Callback (execute)
				'If UseCallBack(x) then cCallback(x)(Lvl(x))	'Callback	'Callback (execute)
				If Lock(x) Then
					Loaded(x) = True
				end if
			end if
		Next
	End Sub
End Class


'Helper function
Function AppendArray(ByVal aArray, aInput)	'append one value, object, or Array onto the end of a 1 dimensional array
	if IsArray(aInput) then 'Input is an array...
		dim tmp : tmp = aArray
		If not IsArray(aArray) Then	'if not array, create an array
			tmp = aInput
		Else						'Append existing array with aInput array
			Redim Preserve tmp(uBound(aArray) + uBound(aInput)+1)	'If existing array, increase bounds by uBound of incoming array
			dim x : for x = 0 to uBound(aInput)
				if isObject(aInput(x)) then
					Set tmp(x+uBound(aArray)+1 ) = aInput(x)
				Else
					tmp(x+uBound(aArray)+1 ) = aInput(x)
				End If
			Next
		AppendArray = tmp	 'return new array
		End If
	Else 'Input is NOT an array...
		If not IsArray(aArray) Then	'if not array, create an array
			aArray = Array(aArray, aInput)
		Else
			Redim Preserve aArray(uBound(aArray)+1)	'If array, increase bounds by 1
			if isObject(aInput) then
				Set aArray(uBound(aArray)) = aInput
			Else
				aArray(uBound(aArray)) = aInput
			End If
		End If
		AppendArray = aArray 'return new array
	End If
End Function











Sub TBF1_Timer()	'debugging dynamiclamps object
	dim x : x = 3
	me.text = "CPU input: " & Round(Flashers.State(x),2) & "(" & GIStates(x) & ")" & vbnewline & _
			  "Value: " & Flashers.SolModValue(x) & vbnewline & _
			  "Lvl: " & Flashers.Lvl(x) & vbnewline & _
			  "Locks: " & Flashers.Locked(x) & " " & Flashers.Loaded(x) & vbnewline & _
			  "Object: " & Lgi1.IntensityScale
End Sub



'dim aTester : Set aTester = New cAnimation
'
'with aTester
'	'.Debug = True	'debug has to go at the start, sry
'	.AddPoint 0, 0, 0	'Keyframe#, Time, Output
'	.AddPoint 1, 500, 35	'100ms, 35
'	.AddPoint 2, 1000, -455
'	.AddPoint 3, 2000, 1555
'	.AddPoint 4, 3000, 0
'	.AddPoint 5, 3100, -5
'	.AddPoint 6, 3200, 0
'	.AddPoint 7, 4200, 0
'	.AddPoint 8, 5000, -25
'	.AddPoint 9, 5500, 0
'	.Callback = "aTesterCallBack"
'End With
'Sub aTesterCallBack(Value) : Plastics_LVL1.z = Value : End Sub


'Keyframe Animation Class

'Setup
'.Update1 - update logic. Use 1 interval
'.Update - update objects. recommended -1 interval
'.Update2 - TODO alternative, updates both on -1 TODO

'Properties
'.State - returns if animation state (true or False)
'.Addpoint - Add keyframes. 3 argument sub : Keyframe#, Time Value, Output Value. Keep keyframes sequential, and timeline straight.
'.Modpoint - Modify an existing point
'.Debug - display debug animation (set before .addpoint to get full debug info)
'.Play - Play Animation
'.Pause - Pause mid-animation
'.Callback - string. Sub to call when animation is updated, with one argument sending the interpolated animation info

'Events
'.Callback(argument) - whatever you set callback to. Manually attach animation to this value - ie Showframe, Height, RotX, RotY, whatever...

'Keyframe Animation Class

'Setup
'.Update1 - update logic. Use 1 interval
'.Update - update objects. recommended -1 interval
'.Update2 - TODO alternative, updates both on -1 TODO

'Properties
'.State - returns if animation state (true or False)
'.Addpoint - Add keyframes. 3 argument sub : Keyframe#, Time Value, Output Value. Keep keyframes sequential, and timeline straight.
'.Modpoint - Modify an existing point
'.Debug - display debug animation (set before .addpoint to get full debug info)
'.Play - Play Animation
'.Pause - Pause mid-animation
'.Callback - string. Sub to call when animation is updated, with one argument sending the interpolated animation info

'Events
'.Callback(argument) - whatever you set callback to. Manually attach animation to this value - ie Showframe, Height, RotX, RotY, whatever...

Class cAnimation
	Public DebugOn
	Private KeyTemp(99,1)
	Private Lock, Loaded, StopAnim, UpdateSub
	Private ms, lvl, KeyStep, KeyLVL 'make these private later
	private LoopAnim

	Private Sub Class_Initialize : redim KeyStep(99) : redim KeyLVL(99) : Lock = True : Loaded = True : ms = 0: End Sub

	Public Property Get State : State = not Lock : End Property
	Public Property Let CallBack(String) : UpdateSub = String : End Property

	public Sub AddPoint(aKey, aMS, aLVL)
		KeyTemp(aKey, 0) = aMS : KeyTemp(aKey, 1) = aLVL
		Shuffle aKey
	End Sub

	'  v  v   v    keyframes IDX / (0)
	'	  .
	'	 / \lvl (1)
	'___/	 \___
	'-----MS--------->

	'in -> AddPoint(KeyFrame#, 0) = KeyFrame(Time)
	'in -> AddPoint(KeyFrame#, 1) = KeyFrame(LVL)
	'	(1d array conversion)
	'into -> KeyStep(99)
	'into -> KeyLvl(99)
	Private Sub Shuffle(aKey) 'shuffle down keyframe data into 1d arrays 'this sucks, it does't actually shuffle anything
		redim preserve KeyStep(99) : redim preserve KeyLvl(99)
		dim str : str = "shuffling @ " & akey & vbnewline
		dim x : for x = 0 to uBound(KeyTemp)
			if KeyTemp(x,0) <> "" Then
				KeyStep(x) = KeyTemp(x,0) : KeyLvl(x) = KeyTemp(x,1)
			Else
				if x = 0 then msgbox "cAnimation error: Please start at keyframe 0!" : exit Sub
				redim preserve KeyStep(x-1) : redim preserve KeyLvl(x-1) : Exit For
			end If
		Next
		str = str & "uBound step:" & uBound(keystep) & vbnewline & "uBound KeyLvl:" & uBound(KeyLvl) & vbnewline
		If DebugOn then TBanima.text = str & "printing steps:" & vbnewline & PrintArray(keystep) & vbnewline & "printing step values:" & vbnewline & PrintArray(keylvl)
	End Sub

	Private function PrintArray(aArray)	'debug
		dim str, x : for x = 0 to uBound(aArray) : str = str & x & ":" & aArray(x) & vbnewline : Next : printarray = str
	end Function

	Public Sub ModPoint(idx, aMs, aLvl) : KeyStep(idx) = aMs : KeyLVL(idx) = aLvl : End Sub  'modify a point after it's set

	Public Sub Play()	: StopAnim = False : Lock = False : Loaded = False : LoopAnim = False :  End Sub 'play animation
	Public Sub PlayLoop()	: StopAnim = False : Lock = False : Loaded = False : LoopAnim = True: End Sub 'play animation
	Public Sub Pause()	: StopAnim = True : end Sub	'pause animation


	Public Sub Update2()	 'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
		'FrameTime = gametime - InitFrame : InitFrame = GameTime	'Calculate frametime
		if not lock then
			if ms > keystep(uBound(keystep)) then
				If LoopAnim then ms = 0 else StopAnim = True : ms = 0	'No looping
			End If
			if StopAnim then Lock = True	'if stopped by script or by end of animation
			if Not Lock Then ms = ms + 1*FrameTime : lvl = LinearEnvelope(ms, KeyStep, KeyLVL)
		end if
		Update
	End Sub

	Public Sub Update1()	'update logic
		if not lock then
			if ms > keystep(uBound(keystep)) then
				If LoopAnim then ms = 0 else StopAnim = True : ms = 0	'No looping
			End If
			if StopAnim then Lock = True	'if stopped by script or by end of animation
			if Not Lock Then ms = ms + 1 : lvl = LinearEnvelope(ms, KeyStep, KeyLVL)
		end if
	End Sub

	Public Sub Update() 	'Update object
		if Not Loaded then
			if Lock then Loaded = True
			if DebugOn then dim str : str = "ms:" & ms & vbnewline & "lvl:" & lvl & vbnewline & _
									Lock & " " & loaded & vbnewline :	tbanim.text = str
			proc UpdateSub, lvl
		end if
	End Sub



End Class

'Variable Envelope. Infinite amount of points!
'Keep keyframes in order, bounds equal, and all that.
'	L1			L2		  L3	    L4		L5... etc...	  L(uBound(xKeyFrame)
'          .
'        /  \
'      /      \				   .
'    /          \           _/  \__
'  /              \       _/       \__
'/                  \ . /             \__
'0.........1..........2........3.........4..........etc........uBound(xKeyFrame)


'in animation's case, xInput = MS
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
	dim y 'Y output
	dim L 'Line
	dim ii : for ii = 1 to uBound(xKeyFrame)	'find active line
		if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
	Next
	if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)	'catch line overrun
	Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

	'Clamp if on the boundry lines
	'if L=1 and Y < yLvl(LBound(yLvl) ) then Y = yLvl(lBound(yLvl) )
	'if L=uBound(xKeyFrame) and Y > yLvl(uBound(yLvl) ) then Y = yLvl(uBound(yLvl) )
	'clamp 2.0
	if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) ) 	'Clamp lower
	if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )	'Clamp upper

	LinearEnvelope = Y
End Function

'debug arrays
'dim Keyz(5), Lvls(5)
'Keyz(0) = 0 : Lvls(0) = -10
'Keyz(1) = 1 : Lvls(1) = 21
'Keyz(2) = 2 : Lvls(2) = 32
'Keyz(3) = 3 : Lvls(3) = -43
'Keyz(4) = 4 : Lvls(4) = 54
'Keyz(5) = 5 : Lvls(5) = 65
'tb.text = LinearEnvelope(5, keyz,lvls)

'======================================





'cCloneBall Class v0.1

'Setup
'Trigger start (& end, if not paired with a VirtualKicker) calling .addball activeball and .removeball activeball. .Removeball trigger should be on your offset meshes.
'OFFSET CLONED PHYSICS MESHES. They should be exactly the same except offset by (whatever coordinate offsets are set).
'Kickers - one for each ball.
'Destroyer trigger- a star kicker on a platform somewhere to the side, or under the playfield.

'Properties
'	.Name 	'Set a name. Used for the events.
'	.Cache 	'Set up your array of kickers. One for each possible ball that can fit in the area. example: Array(cbank1,cbank2,cbank3,cbank4)
'	.Destroyer	'name of your destroyer trigger (use star type with radius and put it somewhere out of the way)
'	.DestroyerHeight 'Height of your destroyer trigger, if applicable (optional, default 0)
'	.XOffset 'X offset of your cloned physics meshes
'	.YOffset 'Y offset of your cloned physics meshes
'	.ZOffset 'Z offset of your cloned physics meshes
'	.Debug 	'Enables debug boxes (TBC and TBCa)
'
'Methods:
'.AddBall ball 	- add ball and start cloning
'.RemoveBall ball - Remove ball, stop cloning

'Events:
'(name)_Entry(Ball) - Called when a ball is added. Used to sync ball tracking w/ virtual kicker

Class cCloneBall
	Public xOffset, yOffset, zOffset, DestroyerHeight, Debug, name
	Private nob, x, idx, Killer
	dim Kickers, Ball, sBall, Lamps

	Private Sub Class_Initialize  : nob=0 : DestroyerHeight=0 : xOffset=0 : yOffset=0 : zOffset=0: end sub

	Public Property Let Cache(input)	'assign kickers
		if not IsArray(input) then msgbox "cCloneball error! .Cache must be an array" : Exit Property
		nob = uBound(input)
		Redim Kickers(uBound(input) )
		Redim Ball(uBound(input) )
		Redim sBall(uBound(input) )
		Redim Lamps(uBound(input) ) 'debug lamps
		for x = 0 to uBound(input) : Set Kickers(x) = input(x) : next
	End Property

	Public Property Let Destroyer(Object)
'		if not typename(Object) = "Kicker" then
'			msgbox "Cloneball error, Destroyer must be a kicker" : exit property
		if not typename(Object) = "Trigger" then
			msgbox "Cloneball error, Destroyer must be a Trigger" : exit property

		else
			Object.Enabled = True
			set killer = object
			executeglobal "sub " & killer.name & "_hit(): me.destroyball : End Sub'"
	end if
	end property


	Private Sub MatchVel(BallOut, BallIn)
		BallOut.VelX = BallIn.VelX : BallOut.VelY = BallIn.VelY : BallOut.VelZ = BallIn.VelZ
	End Sub
	Private Sub MatchPos(BallOut, BallIn, InOut)	'-1 = into table, 1 = out of table
		BallOut.X = BallIn.X + (xoffset * inout): BallOut.Y = BallIn.Y + (yoffset * inout)
		BallOut.Z = BallIn.Z + (zoffset * inout)
	End Sub

	Private Sub MatchPos2(BallOut, BallIn, InOut)	'for kickers!
		BallOut.X = BallIn.X + (xoffset * inout): BallOut.Y = BallIn.Y + (yoffset * inout)
		BallOut.Z = BallIn.Z + BallIn.Radius + (zoffset * inout)
	End Sub

	'Private Sub KillBall(aBall) : aBall.Visible = False : aBall.x = Killer.x : aBall.y = killer.y-35 : aBall.z = kickerdepth+37 : KillVel aBall: End Sub
	Private Sub KillBall(aBall) : killvel aBall : aBall.x = Killer.x : aBall.y = killer.y-killer.radius-ballsize : aBall.z = DestroyerHeight+ballsize : aBall.VelY = 50 : End Sub
	Private Sub KillVel(aBall) : aBall.Visible = False : aBall.VelX = 0 : aBall.VelY = 0 : aBall.VelZ = 0 : End Sub

	Public Sub Update()
		if IsEmpty(Ball) then exit sub
		for idx = 0 to nob
			if not IsEmpty(Ball(idx)) Then
				MatchPos2 Ball(idx), sBall(idx), -1	'match table ball to sBall
				'KillVel Ball(idx)
			end If
		next
		if Debug then DebugBoxes
	End Sub

	Private Sub DebugBoxes()
		'On Error Resume Next
		dim str
			if IsEmpty(ball(0)) then exit sub
		str = Str & "lock Ball0: " & IsEmpty(ball(0)) & " " & round(sBall(0).z) & " " & round(Ball(0).z) & vbnewline & _
					"lock Ball1: " & IsEmpty(ball(1)) & vbnewline & _
					"lock Ball2: " & IsEmpty(ball(2)) & vbnewline & _
					"lock Ball3: " & IsEmpty(ball(3)) & vbnewline & _
					" ... "
		'On Error Goto 0
		TBC.text = str
	End Sub

	Public Sub AddBall(aBall)  	'put ball ref into array
		dim str
		if nob > uBound(Ball) then msgbox "cCloneball error! too many balls (" & nob+1 & "/ " & uBound(Ball) & ")!" : Exit Sub
		str = "Addball Go!" & vbnewline

		for idx = 0 to nob
			if IsEmpty(Ball(idx) ) then

				'track ball pos, kill ball, replace with imposter
				dim tball : set tball = new SpoofBall	'just an ID corpse to prevent collision errors. Ball(x) is the imposter.
				tball.X = aBall.x : tball.y = aBall.y : tball.z = aBall.z :tball.Id = aBall.id
				tball.VelX = aBall.velx : tball.VelY = aBall.vely : tball.VelZ = aBall.velz
				tball.mass = aBall.mass : tball.radius = aBall.radius
				KillBall aBall  'kill ball



				Kickers(idx).CreateSizedBallWithMass tball.radius, tball.mass 'don't kick. Physics are disabled for this ball
				set Ball(idx) = Kickers(idx).LastCapturedBall : ball(idx).Visible = False : if Debug then ball(idx).Color = rgb(255,255,25)
				Ball(idx).id = tball.ID	'keep the balls 2018
				MatchPos2 Ball(idx), tball, 0 'replace with imposter
				ball(idx).Visible = True
				Ball(idx).ForceReflection = True

				'str = str & "Set ball " & x & " to aBall..." & vbnewline	'debug string
				'Position then Kick lower PF ball
				K.CreateSizedBallWithMass tball.radius, tball.mass : k.kick 0,0
				set sBall(idx) = k.LastCapturedBall : if Debug then sball(idx).Color = rgb(25,25,255)
				'Update	'match positions
				MatchPos sBall(idx), tball, 1
				MatchVel sBall(idx), tball	'match velocity of 1 to velocity of 2
				sBall(idx).id = tball.ID+10	'keep the balls 2018

				'Fire callback with ball references
				proc name & "_Entry", sBall(idx)

				'str = str & "Blue sBall" & x & ", match ball @ " & vbnewline & _
				'	  round(Ball(idx).x,0) & " / " & round(Ball(idx).y,0) & " / " & round(Ball(idx).z,0) & vbnewline
				'Update
				if Debug then tbca.text = name & " AddBall #" & idx & vbnewline & gametime

				exit sub
			else
				str = str & "IsEmpty " & x & "=" & isempty(Ball(idx)) & vbnewline
			end if

		Next
		'str = str & "IsEmpty?" & vbnewline & isempty(Ball(0)
		if Debug then tbca.text = str & "they're all full?"
	End Sub

	Public Sub RemoveBall(aBall)
		dim str : str = "Remove ball Go" & vbnewline
		'should this trigger be on the secret ball? w/e
		for idx = 0 to nob
			if not IsEmpty(Ball(idx)) Then
				if sBall(idx).id = aBall.Id Then
					Ball(idx).z = Ball(idx).z - Ball(idx).radius	'prevent hopping
					Kickers(idx).kick 0, 0
					MatchVel Ball(idx), sBall(idx)
					if Debug then Ball(idx).color = rgb(255,255,199)

					KillBall sBall(idx)
					Ball(idx).ForceReflection = False
					Ball(idx) = Empty : sBall(idx) = Empty
				end If
			end If
		next
	End Sub

End Class
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




'VirtualKicker Class v0.1
'A kicker stand-in that does not disable physics on the ball. Designed for proper, meshed-out physical kickers.
'Works via ball tracking.


'Setup
'Object.Update on a timer (100 interval is probably fine)
'In and Out points addball and removeball points
'(for this cloneball setup, addball will be called by clone trigger, and then removeball happens automatically via ExitBounds)

'Variables & Properties-
'Origin - X, Y, Z coordinates of the kicker. Set up in array - example, .Origin = Array(329.25, 102, -25)
'Name - Important for the events to work properly. Example Popper.Name = "Popper"
'Size - Size of the kicker. Default = 25
'State - State of the kicker switch. True or False.
'KickOffset - When kicking out, Will teleport the ball up by it's radius, as normal kickers do. (True by default)
'InitExitBounds - 3 coord array like Origin. An alternate way to automatically remove balls from tracking. If set, calls "_exit" event.
'Debug - Enable debug boxes. Uses two boxes named "tbk" & "tbka"

'Methods-
'.Kick - Direction, Speed, zSpeed	- works just like a normal kicker.

'Events-
'(name)_Hit() 'called once when kicker is hit
'(name)_UnHit() 'called once when kicker is unhit
'(name)_Exit(Ball) 'Special, called when ball leaves InitExitBounds. Used to sync balltracking w/ cCloneball

'Known Bugs & TODOs
'If there are multiple balls in the switch position, the lowest ball (by Z height) will be kicked. That's going to cause issues for level ball-troughs.

Class VirtualKicker
	private x, y, z, idx	'don't use x for helper var, it's the X coord of the virtual kicker origin
	private ExitB(2), AutoExit	'optional, automatic removeball stuff
	Public Size, State, Name, KickOffset, Debug
	Private Queue(3),Locked(3)	'temp bounds

	Private Sub Class_Initialize : Size = 25 : KickOffset = True: AutoExit = False: End Sub

	Public Property Let Origin(ByVal Object)
		if isArray(Object) Then
			if uBound(Object) <> 2 then msgbox "VirtualKicker.Origin array must be 3 values, IE array(Xvalue,Yvalue,Zvalue)!" : Exit Property : End If
			x = Object(0) : y = Object(1) : z = Object(2)
			'tbka.text = "Kicker Assigned @ " & vbnewline & round(x) & vbnewline & round(y) & vbnewline & round(z)
		Else
			msgbox "VirtualKicker.Origin = ??? please use an array with 3 coordinates"
		end If
	End Property

	Public Property Let InitExitBounds(aArray)
		if uBound(aArray) <> 2 then msgbox "VirtualKicker.InitExitBounds array must be 3 values, IE array(Xvalue,Yvalue,Zvalue)!" : Exit Property : End If
		AutoExit = True : ExitB(0) = aArray(0) : ExitB(1) = aArray(1) : ExitB(2) = aArray(2)
	End Property



	'ball tracking
	Public Sub AddBall(aBall)
		for idx = 0 to uBound(Queue)
			if IsEmpty(Queue(idx) ) Then
				Set queue(idx) = aBall
				Exit Sub
			end If
		Next
	End Sub

	Public Sub RemoveBall(aBall)	'handle automatically or manually or w/e
		for idx = 0 to uBound(Queue)
			if not IsEmpty(Queue(idx) ) Then
				if aBall.ID = Queue(idx).ID then
					Queue(idx) = Empty
					if Debug then tbka.text = "Removed Ball @ " & idx
					Exit Sub
				end If
			end If
		Next
		if Debug then tbka.text = "removeball error!" & vbnewline & "attempted to remove ballid=" & aBall.id & "but none was found!"
	End Sub

	Private Function CheckPos(aBall) 'check if ball is in kicker Position. Return largest range from origin
		dim xx, yy, zz
		xx = abs(abs(x) - abs(aBall.x))
		yy = abs(abs(y) - abs(aBall.y))
		zz = abs(abs(z) - abs(aBall.z))
		CheckPos = Max(array(xx,yy,zz) )
	End Function

	Private Function OutOfBounds(aBall)
		dim Bool : bool = False
		if abs(abs(x) - abs(aBall.x)) > exitb(0) then bool = True
		if abs(abs(y) - abs(aBall.y)) > exitb(1) then bool = True
		if abs(abs(z) - abs(aBall.z)) > exitb(2) then bool = True
		OutOfBounds = Bool
	End Function

	Public Sub Update() 'meat of this Object
		'Switch Handling
		dim str
		dim Ranger
		dim SwitchOn : SwitchOn = False 'if at least one ball is in locked position, this becomes true
		'if IsEmpty(queue) then tbka.text = "update queue empty" : Exit Sub
		for idx = 0 to uBound(queue)	'Go through all balls, check off ones that are in switch position
			if not IsEmpty(queue(idx) ) then
				Ranger = CheckPos(queue(idx) )	'Check if any ball is within Switch Range
				if Ranger < Size then
					Locked(idx) = True : SwitchOn = True
				Else
					Locked(idx) = False
				end if

				if AutoExit then 'automatically remove balls that are outside InitExitBounds, and execute _exit event
					if OutOfBounds(queue(idx) ) then
						proc name & "_Exit", queue(idx) : RemoveBall(queue(idx) )
						if Debug then tbka.text = name & " AutoExit removeball " & idx & vbnewline & gametime
					end if
				end If
			end if
		Next
		SetState SwitchOn

		if Debug then DebugBoxes
	End Sub

	Private Sub DebugBoxes()
		dim str
		if IsEmpty(queue(0)) then exit sub
		str = Str & "Switch state? " & State & vbnewline & _
					"lock Ball0: " & IsEmpty(queue(0)) & vbnewline & _
					"lock Ball1: " & IsEmpty(queue(1)) & vbnewline & _
					"lock Ball2: " & IsEmpty(queue(2)) & vbnewline & _
					"lock Ball3: " & IsEmpty(queue(3)) & vbnewline & _
					" ... "
		TBk.text = str
	End Sub

	Private Function ClosestBallToOrigin(aArray)
		redim Zheights(uBound(queue))
		dim tmpidx, f
		for idx = 0 to uBound(aArray)
			if IsObject(aArray(idx)) then zHeights(idx) = aArray(idx).Z
		Next
			f = MinIDX(zHeights, tmpidx)	'find lowest number in array, AND return index of that number (tmpidx)
			Set ClosestBallToOrigin = queue(tmpidx)
	End Function

	Public Sub Kick(aDir, aSpeed, aVelZ)
		dim KickBall ' find lowest ball that's on the trigger and kick It
		'redim ballpositions(uBound(queue))
		for idx = 0 to uBound(queue)
			if State then
				Set KickBall = ClosestBallToOrigin(queue)
				'ballpositions(idx) = queue(idx).z
			end if
		Next
		'play SFX here?
		if IsEmpty(KickBall) then Exit Sub

		If KickOffset Then KickBall.Z = KickBall.Z + KickBall.Radius 'When regular kickers kick, they teleport the ball up a bit first
		'Degrees->Radians conversion stuff. Might not be very accurate, sorry
		aDir = aDir*-1+180
		do while aDir > 360 : aDir = aDir - 360 : Loop
		do while aDir < 0 : aDir = aDir + 360 : Loop
		aDir = ((aDir)*3.1415) / 180
		KickBall.VelX = Sin(aDir) * aSpeed
		KickBall.Vely = Cos(aDir) * aSpeed
		KickBall.VelZ = aVelZ
	end Sub

	Private Sub SetState(aEnabled)
		If State = aEnabled then  Exit Sub 'Discard redundant updates
		if aEnabled Then
			'Lamp1.visible = True 'execute whatever TODO
			if isempty(name) then msgbox "VirtualKicker error: Please set a .name for this object for _hit to call!"
			execute name & "_Hit"
		Elseif Not aEnabled then
			'Lamp1.visible = False
			execute name & "_UnHit"
		End if
		State = aEnabled
	End Sub

End Class


'Common Functions

Sub Proc(string, Callback)	'proc using a string and one argument
	'On Error Resume Next
	dim p : Set P = GetRef(String)
	P Callback
	If err.number = 13 then  msgbox "Proc error! No such procedure: " & vbnewline & string
	if err.number = 424 then msgbox "Proc error! No such Object"
End Sub

Function Min(aArray)
	dim idx, MinItem', str
	for idx = 0 to uBound(aArray)
		if IsEmpty(MinItem) then
			if not IsEmpty(aArray(idx)) then
				MinItem = aArray(idx)
			end If
		end if
		if not IsEmpty(aArray(idx) ) then
			If aArray(idx) < Minitem then Minitem = aArray(idx)
		end If
	Next
	Min = Minitem
End Function

Function MinIDX(byval aArray, byref index)	'min, but also returns Index number of lowest
	dim idx, MinItem', str
	for idx = 0 to uBound(aArray)
		if IsEmpty(MinItem) then
			if not IsEmpty(aArray(idx)) then
				MinItem = aArray(idx)
				index = idx
			end If
		end if
		if not IsEmpty(aArray(idx) ) then
			If aArray(idx) < Minitem then Minitem = aArray(idx) : index = idx
		end If
	Next
	MinIDX = Minitem
End Function

Function Max(aArray)
	dim idx, MaxItem', str
	for idx = 0 to uBound(aArray)
		if IsEmpty(MaxItem) then
			if not IsEmpty(aArray(idx)) then
				MaxItem = aArray(idx)
			end If
		end if
		if not IsEmpty(aArray(idx) ) then
			If aArray(idx) > MaxItem then MaxItem = aArray(idx)
		end If
	Next
	Max = MaxItem
End Function


Function MaxIDX(byval aArray, byref index)	'max, but also returns Index number of highest
	dim idx, MaxItem', str
	for idx = 0 to uBound(aArray)
		if IsEmpty(MaxItem) then
			if not IsEmpty(aArray(idx)) then
				MaxItem = aArray(idx)
				index = idx
			end If
		end if
		if not IsEmpty(aArray(idx) ) then
			If aArray(idx) > MaxItem then MaxItem = aArray(idx) : index = idx
		end If
	Next
	Max = MaxItem
End Function







'Methods:
'.TimeDelay - Delay before trigger shuts off automatically. Default = 80 (ms)
'.AddPoint - "Polarity", "Velocity", "Ycoef" coordinate points. Use one of these 3 strings, keep coordinates sequential. x = %position on the flipper, y = output
'.Object - set to flipper reference. Optional.
'.StartPoint - set start point coord. Unnecessary, if .object is used.

'Called with flipper -
'ProcessBalls - catches ball data.
' - OR -
'.Fire - fires flipper.rotatetoend automatically + processballs. Requires .Object to be set to the flipper.


'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball

Class FlipperPolarity
	Public DebugOn, Enabled
	Private FlipAt	'Timer variable (IE 'flip at 723,530ms...)
	Public TimeDelay	'delay before trigger turns off and polarity is disabled TODO set time!
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

	Public Sub Report(aChooseArray) 	'debug, reports all coords in tbPL.text
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
				if DebugOn then StickL.visible = True : StickL.x = balldata(x).x		'debug TODO
			End If
		Next
		PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
		PartialFlipCoef = abs(PartialFlipCoef-1)
	End Sub
	Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function	'Timer shutoff for polaritycorrect

	Public Sub PolarityCorrect(aBall)
		if FlipperOn() then
			dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1
			dim teststr : teststr = "Cutoff"
			tmp = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
			if tmp < 0.1 then 'if real ball position is behind flipper, exit Sub to prevent stucks	'Disabled 1.03, I think it's the Mesh that's causing stucks, not this
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
					if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)				'find safety coefficient 'ycoef' data
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
		 : 			VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
					if Enabled then aBall.Velx = aBall.Velx*VelCoef
					if Enabled then aBall.Vely = aBall.Vely*VelCoef
				end if
			End If

			'Polarity Correction (optional now)
			if not IsEmpty(PolarityIn(0) ) then
				If StartPoint > EndPoint then LR = -1	'Reverse polarity if left flipper
				dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
				if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
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
	for x = 0 to uBound(aArray)	'Shuffle objects in a temp array
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
	redim aArray(aCount-1+offset)	'Resize original array
	for x = 0 to aCount-1		'set objects back into original array
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

Function PSlope(Input, X1, Y1, X2, Y2)	'Set up line via two points, no clamping. Input X, output Y
	dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
	Y = M*x+b
	PSlope = Y
End Function

Function NullFunctionZ(aEnabled):End Function	'1 argument null function placeholder	 TODO move me or replac eme
Class cvFastFlipsZ
	Public TiltObjects, DebugOn, Name, Delay
	Private SubL, SubUL, SubR, SubUR, FlippersEnabled,  LagCompensation, FlipState(3), Sol

	Private Sub Class_Initialize()
		Delay = 0 : FlippersEnabled = False : DebugOn = False : LagCompensation = False : Sol = 0 : TiltObjects = True
		Set SubL = GetRef("NullFunctionZ"): Set SubR = GetRef("NullFunctionZ") : Set SubUL = GetRef("NullFunctionZ"): Set SubUR = GetRef("NullFunctionZ")
	End Sub

	'set callbacks
	Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : Decouple sLLFlipper, aInput: End Property
	Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : Decouple sULFlipper, aInput: End Property
	Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : Decouple sLRFlipper, aInput: End Property
	Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : Decouple sURFlipper, aInput: End Property

	'Automatically decouple flipper solcallback script lines (only if both are pointing to the same sub) thanks gtxjoe
	Private Sub Decouple(aSolType, aInput)  : If StrComp(SolCallback(aSolType),aInput,1) = 0 then SolCallback(aSolType) = Empty End If : End Sub
	Public Property Let Solenoid(aInput) : if not IsEmpty(aInput) then Sol = aInput : end if : End Property	'set solenoid
	Public Property Get Solenoid : Solenoid = sol : End Property

	'call callbacks
	Public Sub FlipL(aEnabled)
		FlipState(0) = aEnabled	'track flipper button states: the game-on sol flips immediately if the button is held down (1.1)
		If not FlippersEnabled and not DebugOn then Exit Sub
		subL aEnabled
	End Sub

	Public Sub FlipR(aEnabled)
		FlipState(1) = aEnabled
		If not FlippersEnabled and not DebugOn then Exit Sub
		subR aEnabled
	End Sub

	Public Sub FlipUL(aEnabled)
		FlipState(2) = aEnabled
		If not FlippersEnabled and not DebugOn then Exit Sub
		subUL aEnabled
	End Sub

	Public Sub FlipUR(aEnabled)
		FlipState(3) = aEnabled
		If not FlippersEnabled and not DebugOn then Exit Sub
		subUR aEnabled
	End Sub

	Public Sub TiltSol(aEnabled)	'Handle solenoid / Delay (if delayinit)
		If delay > 0 and not aEnabled then 	'handle delay
			vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
			LagCompensation = True
		else
			If Delay > 0 then LagCompensation = False
			EnableFlippers(aEnabled)
		end If
	End Sub

	Sub FireDelay() : If LagCompensation then EnableFlippers False End If : End Sub

	Private Sub EnableFlippers(aEnabled)
		If aEnabled then SubL FlipState(0) : SubR FlipState(1) : subUL FlipState(2) : subUR FlipState(3)
		FlippersEnabled = aEnabled
		If TiltObjects then vpmnudge.solgameon aEnabled
		If Not aEnabled then
			subL False
			subR False
			If not IsEmpty(subUL) then subUL False
			If not IsEmpty(subUR) then subUR False
		End If
	End Sub
End Class

'
''Old Pinmame sol tracker debug script
'
''Solenoid / GI debug boxes
'Dim GIStates(10)
''dim SolStates(100)
'SolStatesInit
'Sub SolStatesInit() : dim x : for x = 0 to 100 : SolStates(x) = 0 : next : End Sub
'
'Sub MegaBoxSol_Timer()
'	dim x, e : e = SolStates	'set solstates to controller.changeSol in main loop
'	dim str : str = "Solenoid Table" & vbnewline
'	dim n0, n1, n2, n3, n4', n5
'	for x = 0 to 9
'		n0 = x : if n0 < 10 then n0 = "0" & n0 end if : n0 = " " & n0 & ":"
'		n1 = x+10 : n1 = " " & n1 & ":"
'		n2 = x+20 : n2 = " " & n2 & ":"
'		n3 = x+30 : n3 = " " & n3 & ":"
'		n4 = x+40 : n4 = " " & n4 & ":"
'		str = str & n0 & e(x) & n1 & e(x+10) & n2 & e(x+20) & n3 & e(x+30) & vbnewline
'	next
'	me.text = str' & str2
'End Sub
'
'Sub MegaBoxGI_Timer()
'	dim x, e : e = GIStates	'set GIstates to controller.changeGI in main loop
'	dim str : str = "GI Table" & vbnewline
'	dim n0, n1, n2, n3, n4', n5
'	for x = 0 to 4 : str = str & x & ":" & e(x) & vbnewline : next
'	me.text = str' & str2
'End Sub




'FLOATING TEXT
'*****************

' low point floating text objects
dim FtLow1 : Set ftLow1 = New FloatingTextFlasher
dim FtLow2 : Set ftLow2 = New FloatingTextFlasher
dim FtLow3 : Set ftLow3 = New FloatingTextFlasher
dim FtLow4 : Set ftLow4 = New FloatingTextFlasher
dim FtLow5 : Set ftLow5 = New FloatingTextFlasher
dim FTlowArray : FTlowArray = array(0, FTlow1,ftLow2,ftLow3,ftLow4,ftLow5)

' med point floating text objects
dim FTmed1 : Set FTmed1 = New FloatingTextFlasher
dim FTmed2 : Set FTmed2 = New FloatingTextFlasher
dim FTmed3 : Set FTmed3 = New FloatingTextFlasher
dim FTmed4 : Set FTmed4 = New FloatingTextFlasher
dim FTmed5 : Set FTmed5 = New FloatingTextFlasher
dim FTmedArray : FTmedArray = array(0, FTmed1,FTmed2,FTmed3,FTmed4,FTmed5)

' high point floating text objects
dim FThi1 : Set FThi1 = New FloatingTextFlasher
dim FThi2 : Set FThi2 = New FloatingTextFlasher
dim FThi3 : Set FThi3 = New FloatingTextFlasher
dim FThiArray : FThiArray = array(0, FThi1,FThi2,FThi3)
'FThi1.textat "2234567", 300, 1140
If FloatingText then InitFloatingText
Sub InitFloatingText()
	dim x, a
	ScoreBox.TimerEnabled = True
	'---end old

	FThi1.Sprites = Array(FThi1_1, FThi1_2, FThi1_3, FThi1_4, FThi1_5, FThi1_6)
	FThi2.Sprites = Array(FThi2_1, FThi2_2, FThi2_3, FThi2_4, FThi2_5, FThi2_6)
	FThi3.Sprites = Array(FThi3_1, FThi3_2, FThi3_3, FThi3_4, FThi3_5, FThi3_6)

	a = FThiArray
	For x = 1 to uBound(a)
		with a(x)
			.Prefix = "Font_"
			.Size = 29*2*2
			.FadeSpeedUp = 1/3500
			.RotX = -37
		end With
	Next

	FTmed1.Sprites = Array(FTmed1_1, FTmed1_2, FTmed1_3, FTmed1_4, FTmed1_5, FTmed1_6)
	FTmed2.Sprites = Array(FTmed2_1, FTmed2_2, FTmed2_3, FTmed2_4, FTmed2_5, FTmed2_6)
	FTmed3.Sprites = Array(FTmed3_1, FTmed3_2, FTmed3_3, FTmed3_4, FTmed3_5, FTmed3_6)
	FTmed4.Sprites = Array(FTmed4_1, FTmed4_2, FTmed4_3, FTmed4_4, FTmed4_5, FTmed4_6)
	FTmed5.Sprites = Array(FTmed5_1, FTmed5_2, FTmed5_3, FTmed5_4, FTmed5_5, FTmed5_6)

	a = FTmedArray
	For x = 1 to uBound(a)
		with a(x)
			.Prefix = "Font_"
			.Size = 29
			.FadeSpeedUp = 1/1500
			.RotX = -37
		end With
	Next


	FtLow1.Sprites = Array(FtLow1_1, FtLow1_2, FtLow1_3, FtLow1_4, FtLow1_5, FTlow1_6)
	FtLow2.Sprites = Array(FtLow2_1, FtLow2_2, FtLow2_3, FtLow2_4, FtLow2_5, FtLow2_6)
	FtLow3.Sprites = Array(FtLow3_1, FtLow3_2, FtLow3_3, FtLow3_4, FtLow3_5, FtLow3_6)
	FtLow4.Sprites = Array(FtLow4_1, FtLow4_2, FtLow4_3, FtLow4_4, FtLow4_5, FtLow4_6)
	FtLow5.Sprites = Array(FtLow5_1, FtLow5_2, FtLow5_3, FtLow5_4, FtLow5_5, FtLow5_6)


	a = FTlowArray
	For x = 1 to uBound(a)
		with a(x)
			.Prefix = "Font_"
			.Size = 29/2
			.FadeSpeedUp = 1/800
			.RotX = -37
		end With
	Next

'
'
'	for each x in aSwitches
'		if typename(x) <> "Primitive" and _
'			typename(x) <> "Bumper" and _
'			typename(x) <> "Trigger" and _
'			typename(x) <> "HitTarget" and _
'			typename(x) <> "Kicker" Then _
'			tb.text = tb.text & x.name
'	Next

End Sub

'ft1.textat "1234567", 300, 1140
'ft1.textat "2234567", 300, 1140

Sub PlaceFloatingTextHi(aArray, aPointGain, ByVal aX, ByVal aY)	'center text a bit for the big scores
	aX = (aX + (table1.width/2))/2
	aY = (aY + (table1.Height/2))/2
	PlaceFloatingText aArray, aPointGain, aX, aY
End Sub

Sub PlaceFloatingText(aArray, aPointGain, aX, aY)
	if gametime < 1000 then exit sub
	dim x, ovrflw, debugstr, gtg
	debugstr = "Overflow..."

	for ovrflw = 1 to 99 'overflow levels
		if gtg then Exit For
		for x = 1 to uBound(aArray)
			if aArray(x).Mask < ovrflw then
				aArray(x).textat aPointGain, aX, aY
				'debugstr = debugstr & " assigned to " & x & " " & ovrflw
				aArray(x).Mask = ovrflw
				gtg = True
				Exit For
			End If
		Next
	Next

End Sub


dim aLutBurst : Set aLutBurst = New cAnimation
dim DefaultLut : DefaultLut = Table1.ColorGradeImage

with aLutBurst
	.AddPoint 0, 0, 0
	.AddPoint 1, 130, 20 'up
	.AddPoint 2, 150, 14 'hold
	.AddPoint 3, 180, 20 'hold
	.AddPoint 4, 210, 14 'hold
	.AddPoint 5, 240, 20 'hold
	.AddPoint 6, 270, 14 'hold
	.AddPoint 7, 290, 20 'hold
	.AddPoint 8, 320, 14 'hold
	.AddPoint 9, 350, 20 'hold
	.AddPoint 10,350+100, 0 'down
	.Callback = "animLutBurst"
End With
Sub animLutBurst(aLVL)
	if aLvl = 0 then
		table1.ColorGradeImage = DefaultLut
	else
		table1.ColorGradeImage = "RedLut_" & Round(aLVL)
	end if
ENd Sub
animlutburst 0


Dim LeftSlingPos : Set LeftSlingPos = New WallSwitchPos
Dim RightSlingPos : Set RightSlingPos = New WallSwitchPos
Dim TroughPos : Set TroughPos = New WallSwitchPos
Dim LeftPopperPos : Set LeftPopperPos = New WallSwitchPos
Dim RightPopperPos : Set RightPopperPos = New WallSwitchPos
Dim TopPopperPos : Set TopPopperPos = New WallSwitchPos
Dim DTPos : Set DTPos = New WallSwitchPos

DTpos.Name = "DropTarget"
DTpos.X = 285
DTpos.Y = 440

'Set lastswitch = LeftSlingPos
TroughPos.Name = "Trough"
TroughPos.X = 425
TroughPos.Y = 1515

LeftSlingPos.Name = "LeftSlingShot"
LeftSlingPos.X = 160+25
LeftSlingPos.Y = 1480

RightSlingPos.Name = "RightSlingShot"
RightSlingPos.X = 700-25
RightSlingPos.Y = 1480

LeftPopperPos.Name = "LeftPopperPos"
LeftPopperPos.X = 180
LeftPopperPos.Y = 1050

RightPopperPos.Name = "RightPopperPos"
RightPopperPos.X = 730
RightPopperPos.Y = 200

TopPopperPos.Name = "TopPopperPos"
TopPopperPos.X = 300
TopPopperPos.Y = 420


dim LastSwitch : Set LastSwitch = TroughPos
dim LastScore : LastScore = 0


Sub ScoreBox_Timer()
	Dim NVRAM, x, overflow :Overflow = True
	NVRAM = Controller.NVRAM
	dim str : str = _
	ConvertBCD(NVRAM(CInt("&h1730"))) & _
	ConvertBCD(NVRAM(CInt("&h1731"))) & _
	ConvertBCD(NVRAM(CInt("&h1732"))) & _
	ConvertBCD(NVRAM(CInt("&h1733"))) & _
	ConvertBCD(NVRAM(CInt("&h1734")))' & _
	'ConvertBCD(NVRAM(CInt("&h1735")))' & _

	str = round(str)


	dim debugstr : debugstr = str
	dim PointGain
	PointGain = Str - LastScore
	LastScore = str

	if PointGain >= 30000000 Then	'hi point scores
		PlaceFloatingTextHi FThiArray, PointGain, LastSwitch.x, LastSwitch.y
		aLutBurst.Play
	elseif pointgain >= 1000000 then 'medium point scores
		PlaceFloatingText FTmedArray, PointGain, LastSwitch.x, LastSwitch.y

	elseif pointgain > 0 then	'low point scores   (FTlow1 -> FtLow5)
		PlaceFloatingText FTlowArray, PointGain, LastSwitch.x, LastSwitch.y

	end if
'	if debugstr <> "" then
'		if tb.text <> debugstr then tb.text = debugstr
'	end if
	'ft1.Update2
	'ft2.Update2
	for x = 1 to uBound(FTlowArray) : FTlowArray(x).Update2 : Next
	for x = 1 to uBound(FTmedArray) : FTmedArray(x).Update2 : Next
	for x = 1 to uBound(FThiArray) : FThiArray(x).Update2 : Next
End Sub

Function ConvertBCD(v)
	ConvertBCD = "" & ((v AND &hF0) / 16) & (v AND &hF)
End Function

Sub aSwitches_Hit(aIDX)
	Set LastSwitch = aSwitches(aIDX)
End Sub

'tbDB.timerenabled = 0
Sub tbDB_Timer()
	dim str,x
	for x = 1 to uBound(FTlowArray)
		str = str & "Load" & x & ": " & Ftlowarray(x).loaded & " " & _
			"mask" & Ftlowarray(x).mask & vbnewline


	Next
	if me.text <> str then me.text = str end if
End Sub

Class WallSwitchPos : Public x,y,name : End Class



Class FloatingTextFlasher
	public count, char, Prfx
	Public Size, Text, FadeSpeedUp, lock, loaded, lvl, z
	public mask 'for handling overflow

	Private Sub Class_Initialize : lvl = 0 : FadeSpeedUp = 1/1500 : loaded = 1 : Mask = 0 : end sub

	Public Property Let RotX(aInput) : dim x : for each x in Sprites : x.RotX = aInput : Next : End Property
	Public Property Let RotY(aInput) : dim x : for each x in Sprites : x.RotY = aInput : Next : End Property
	Public Property Let RotZ(aInput) : dim x : for each x in Sprites : x.RotZ = aInput : Next : End Property

	Public Property Let Sprites(aArray)
		if IsArray(aArray) Then
			char = aArray
			Count = uBound(aArray)
			z = aArray(0).height
		Else
			msgbox "FloatingText Error, 'Sprites' must be an array!"
		End If
	End Property
	Public Property Get Sprites : Sprites = char : End Property

	Public Property Let Prefix(aStr)
		Prfx = aStr
	End Property
	Public Property Get Prefix : Prefix = prfx : End Property

	'Position text
	Public Sub TextAt(aStr, aX, aY)
		Text = aStr
		dim idx
		'Update Position
		'0 '1 '2
		'a(0) = aX
		'a(1) = aX + Size * index
		for idx = 0 to Count
			char(idx).x = aX + (Size * idx) - (Len(Text)*Size)/2	'len part centers text
			char(idx).y = aY
		Next

		'Update Text
		dim tmp
		for idx = 0 to Count
			tmp = Mid(aStr, idx+1, 1)
			if tmp <> "" then
				char(idx).visible = True
				char(idx).ImageA = Prfx & tmp
				char(idx).ImageB = ""
			Else
				char(idx).visible = False
			end If
		Next
		FormatNumbers

		'start fading / floating up
		lock = False : Loaded = False : lvl = 0
	End Sub


	Private Sub FormatNumbers()
		If Len(Text) > 9 then
			Commalate len(Text)-9
		End If
		If Len(Text) > 6 then
			Commalate len(Text)-6
		End If
		If Len(Text) > 3 Then
			Commalate len(Text)-3
		End If
	End Sub

	Private Sub Commalate(aIDX)
		char(aIdx-1).ImageB = Prfx & "Comma"
	End Sub




	Public Sub Update2()	 'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
		'FrameTime = gametime - InitFrame : InitFrame = GameTime	'Calculate frametime
		if not Lock then
			Lvl = Lvl + FadeSpeedUp * frametime
			if Lvl >= 1 then Lvl = 1 : Lock = True
		end if
		Update
	End Sub

	Private Sub Update()	'Handle object updates
			if not Loaded then
				dim opacitycurve
				if lvl > 0.5 then
					opacitycurve = pSlope(lvl, 0, 1, 1, 0)
				Else
					opacitycurve = 1
				end If

				dim x
				for each x in char
					x.height = z + (lvl * 100)
					x.IntensityScale = opacitycurve
				Next
				if lvl >= 0.5 then Mask = 0
				If Lock Then
					'if Lvl = 1 or Lvl = 0 then Loaded = True	'finished fading
					if Lvl = 1 then Loaded = True 	'finished fading
				end if
			end if
			'dim str : str = "lock&load:" & lock & " " & loaded & vbnewline & char(0).opacity & vbnewline & char(0).height & vbnewline & char(0).visible & vbnewline & char(0).x & char(0).y
			'if tbFT.text <> str then tbft.text = str
	End Sub

End Class
