'*******************************************************************************************************
'
'					   	    	 Surf'n Safari Premier 1991 VPX v1.2.0
'								http://www.ipdb.org/machine.cgi?id=2461
'
'											Created by Kiwi
'
'*******************************************************************************************************

Option Explicit
Randomize

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

Const VolBump   = 2    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


'*******************************************************************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'********************************************** OPTIONS ************************************************

Const cGameName   = "surfnsaf"

'************ DMD Visible/Hidden : 0 visible , 1 hidden

'Const DMDHidden = 0

Dim DMDHidden

 If Table1.ShowFSS = True Then DMDHidden = 1

 If B2SOn = True Then DMDHidden = 1

'******************************************** Ball

Const BallSize    = 50

'Const BallMass    = 1.025		'Mass=(53.11^3)/125000 ,(BallSize^3)/125000

'******************************************** Flashers Level

Const Lumen    = 10

'************ Rails and rail lights Hidden/Visible in FS mode : 0 hidden , 1 visible

Const RailsLights = 0

'************ Glasses Color : 1=White ,2=Yellow ,3=Orange ,4=Red ,5=Violet ,6=Green ,7=Blue

Const GlassesColor = 1

'******************************************** OPTIONS END **********************************************

'******************************************** FSS Init

 If Table1.ShowFSS = False Then
	flb1.Visible = 0
	flb2.Visible = 0
	flb3.Visible = 0
	flb4.Visible = 0
	flb5.Visible = 0
	flb6.Visible = 0
	flb7.Visible = 0
	flb8.Visible = 0
	flb9.Visible = 0
	flb10.Visible = 0
	flb11.Visible = 0
	flb12.Visible = 0
	flb13.Visible = 0
	flb14.Visible = 0
	flb15.Visible = 0
	flb16.Visible = 0
	flb17.Visible = 0
	flb18.Visible = 0
	flb19.Visible = 0
	flb20.Visible = 0
	flb21.Visible = 0
	flb22.Visible = 0
	flb23.Visible = 0
	flb24.Visible = 0
	flb25.Visible = 0
	flb26.Visible = 0
	flb27.Visible = 0
	flb28.Visible = 0
	f131.Visible = 0
	f131a.Visible = 0
	f132.Visible = 0
	f132a.Visible = 0
	f133.Visible = 0
	f133a.Visible = 0
	f134.Visible = 0
	f134a.Visible = 0

	f60.Visible = 0
	f61.Visible = 0
	f62.Visible = 0
	f63.Visible = 0
	f64.Visible = 0
	f65.Visible = 0
	f77.Visible = 0
	f87.Visible = 0
	f97.Visible = 0
	f107.Visible = 0
	f117.Visible = 0

End If


LoadVPM "01560000", "gts3.VBS", 3.26

Dim bsTrough, cdtBank, ddtBank, bsUK, bsLK, mHole, mHole1, mHole2, PinPlay

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "coin3"


Const swLCoin = 0
Const swRCoin = 1
Const swCCoin = 2
Const swCoinShuttle = 3
Const swStartButton = 4
Const swTournament = 5
Const swFrontDoor = 6

'************
' Table init.
'************

Sub Table1_Init
	vpmInit me
	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "Surf'n Safari Premier (Gottlieb 1991)" & vbNewLine & "VPX table by Kiwi 1.2.0"
		.HandleMechanics = 0
		.HandleKeyboard = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.ShowTitle = 0
		.Hidden = DMDHidden
'		.DoubleSize = 1
'		.Games(cGameName).Settings.Value("dmd_pos_x")=0
'		.Games(cGameName).Settings.Value("dmd_pos_y")=0
'		.Games(cGameName).Settings.Value("dmd_width")=400
'		.Games(cGameName).Settings.Value("dmd_height")=92
'		.Games(cGameName).Settings.Value("rol") = 0
'		.Games(cGameName).Settings.Value("sound") = 1
'		.Games(cGameName).Settings.Value("ddraw") = 1
		.Games(cGameName).Settings.Value("dmd_red")=60
		.Games(cGameName).Settings.Value("dmd_green")=180
		.Games(cGameName).Settings.Value("dmd_blue")=255
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

	' Nudging
	vpmNudge.TiltSwitch = 151
	vpmNudge.Sensitivity = 2
	vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingShot, RightSlingShot, sw15)

	' Trough
	Set bsTrough = New cvpmBallStack
	With bsTrough
		.InitSw 5, 0, 0, 25, 0, 0, 0, 0
		.InitKick BallRelease, 68, 6
		.InitEntrySnd "Solenoid", "Solenoid"
		.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
		.Balls = 3
	End With

	' Top Upkicker
	Set bsUK = New cvpmBallStack
	With bsUK
		.InitSaucer sw21, 21, 0, 32
		.KickZ = 1.56
		.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
	End With

	' Left Upkicker
	Set bsLK = New cvpmBallStack
	With bsLK
		.InitSaucer sw35, 35, 0, 35
		.KickZ = 1.56
		.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
	End With

	' Drop targets
	set cdtBank = new cvpmdroptarget
	With cdtBank
		.initdrop array(sw16, sw26, sw36), array(16, 26, 36)
'		.initsnd SoundFX("DROPTARG",DOFContactors), SoundFX("DTResetB",DOFContactors)
	End With

	set ddtBank = new cvpmdroptarget
	With ddtBank
		.initdrop array(sw17, sw27, sw37), array(17, 27, 37)
'		.initsnd SoundFX("DROPTARG",DOFContactors), SoundFX("DTResetB",DOFContactors)
	End With

	' Low powered Magnet , Left VUK
    Set mHole = New cvpmMagnet
    With mHole
        .initMagnet MagTrigger, 3
        .X = 77
        .Y = 1004.75
        .Size = 37
        .GrabCenter = 0
        .MagnetOn = 1
        .CreateEvents "mHole"
    End With

	' Low powered Magnet1 , Skill Shot
    Set mHole1 = New cvpmMagnet
    With mHole1
        .initMagnet MagTrigger1, 4
        .X = 475
        .Y = 188
        .Size = 35
        .GrabCenter = 0
        .MagnetOn = 1
        .CreateEvents "mHole1"
    End With

	' Low powered Magnet2 , Whirlpool
    Set mHole2 = New cvpmMagnet
    With mHole2
        .initMagnet MagTrigger2, 1
        .X = 462
        .Y = 405
        .Size = 95
        .GrabCenter = 0
        .MagnetOn = 0
        .CreateEvents "mHole2"
    End With

' Main Timer init
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

' GI Delay Timer
	GIDelay.Enabled = 1

' Init Droptargets
	Arm.IsDropped=1
	Arm2.IsDropped=1

	If ShowDT=True Then
		RailSx.visible=1
		RailDx.visible=1
		Trim.visible=1
		TrimS1.visible=1
		TrimS2.visible=1
		TrimS3.visible=1
		TrimS4.visible=1

		fgit1L.visible=1
		fgit1R.visible=1
		f43a.visible=1
		f50a.visible=1
		f56a.visible=1
		f57a.visible=1
		Bulb10b.visible=1
	Else
		RailSx.visible=RailsLights
		RailDx.visible=RailsLights
		Trim.visible=RailsLights
		TrimS1.visible=RailsLights
		TrimS2.visible=RailsLights
		TrimS3.visible=RailsLights
		TrimS4.visible=RailsLights

		fgit1L.visible=RailsLights
		fgit1R.visible=RailsLights
		f43a.visible=RailsLights
		f50a.visible=RailsLights
		f56a.visible=RailsLights
		f57a.visible=RailsLights
		Bulb10b.visible=RailsLights
	End If

' Backbox

	f60.Y = 1:f61.Y = 1:f62.Y = 1:f63.Y = 1:f64.Y = 1:f65.Y = 1:f77.Y = 1:f87.Y = 1:f97.Y = 1:f107.Y = 1:f117.Y = 1
	f131.Y = 1:f131a.Y = 1:f132.Y = 1:f132a.Y = 1:f133.Y = 1:f133a.Y = 1:f134.Y = 1:f134a.Y = 1
	flb1.Y = 1:flb2.Y = 1:flb3.Y = 1:flb4.Y = 1:flb5.Y = 1:flb6.Y = 1:flb7.Y = 1:flb8.Y = 1:flb9.Y = 1
	flb10.Y = 1:flb11.Y = 1:flb12.Y = 1:flb13.Y = 1:flb14.Y = 1:flb15.Y = 1:flb16.Y = 1:flb17.Y = 1:flb18.Y = 1:flb19.Y = 1
	flb20.Y = 1:flb21.Y = 1:flb22.Y = 1:flb23.Y = 1:flb24.Y = 1:flb25.Y = 1:flb26.Y = 1:flb27.Y = 1:flb28.Y = 1

	a00.Y = 1:a01.Y = 1:a02.Y = 1:a03.Y = 1:a04.Y = 1:a05.Y = 1:a06.Y = 1:a07.Y = 1:a08.Y = 1:a09.Y = 1:a0a.Y = 1:a0b.Y = 1:a0c.Y = 1:a0d.Y = 1:a0e.Y = 1:a0f.Y = 1
	a10.Y = 1:a11.Y = 1:a12.Y = 1:a13.Y = 1:a14.Y = 1:a15.Y = 1:a16.Y = 1:a17.Y = 1:a18.Y = 1:a19.Y = 1:a1a.Y = 1:a1b.Y = 1:a1c.Y = 1:a1d.Y = 1:a1e.Y = 1:a1f.Y = 1
	a20.Y = 1:a21.Y = 1:a22.Y = 1:a23.Y = 1:a24.Y = 1:a25.Y = 1:a26.Y = 1:a27.Y = 1:a28.Y = 1:a29.Y = 1:a2a.Y = 1:a2b.Y = 1:a2c.Y = 1:a2d.Y = 1:a2e.Y = 1:a2f.Y = 1
	a30.Y = 1:a31.Y = 1:a32.Y = 1:a33.Y = 1:a34.Y = 1:a35.Y = 1:a36.Y = 1:a37.Y = 1:a38.Y = 1:a39.Y = 1:a3a.Y = 1:a3b.Y = 1:a3c.Y = 1:a3d.Y = 1:a3e.Y = 1:a3f.Y = 1
	a40.Y = 1:a41.Y = 1:a42.Y = 1:a43.Y = 1:a44.Y = 1:a45.Y = 1:a46.Y = 1:a47.Y = 1:a48.Y = 1:a49.Y = 1:a4a.Y = 1:a4b.Y = 1:a4c.Y = 1:a4d.Y = 1:a4e.Y = 1:a4f.Y = 1
	a50.Y = 1:a51.Y = 1:a52.Y = 1:a53.Y = 1:a54.Y = 1:a55.Y = 1:a56.Y = 1:a57.Y = 1:a58.Y = 1:a59.Y = 1:a5a.Y = 1:a5b.Y = 1:a5c.Y = 1:a5d.Y = 1:a5e.Y = 1:a5f.Y = 1
	a60.Y = 1:a61.Y = 1:a62.Y = 1:a63.Y = 1:a64.Y = 1:a65.Y = 1:a66.Y = 1:a67.Y = 1:a68.Y = 1:a69.Y = 1:a6a.Y = 1:a6b.Y = 1:a6c.Y = 1:a6d.Y = 1:a6e.Y = 1:a6f.Y = 1
	a70.Y = 1:a71.Y = 1:a72.Y = 1:a73.Y = 1:a74.Y = 1:a75.Y = 1:a76.Y = 1:a77.Y = 1:a78.Y = 1:a79.Y = 1:a7a.Y = 1:a7b.Y = 1:a7c.Y = 1:a7d.Y = 1:a7e.Y = 1:a7f.Y = 1
	a80.Y = 1:a81.Y = 1:a82.Y = 1:a83.Y = 1:a84.Y = 1:a85.Y = 1:a86.Y = 1:a87.Y = 1:a88.Y = 1:a89.Y = 1:a8a.Y = 1:a8b.Y = 1:a8c.Y = 1:a8d.Y = 1:a8e.Y = 1:a8f.Y = 1
	a90.Y = 1:a91.Y = 1:a92.Y = 1:a93.Y = 1:a94.Y = 1:a95.Y = 1:a96.Y = 1:a97.Y = 1:a98.Y = 1:a99.Y = 1:a9a.Y = 1:a9b.Y = 1:a9c.Y = 1:a9d.Y = 1:a9e.Y = 1:a9f.Y = 1
	aa0.Y = 1:aa1.Y = 1:aa2.Y = 1:aa3.Y = 1:aa4.Y = 1:aa5.Y = 1:aa6.Y = 1:aa7.Y = 1:aa8.Y = 1:aa9.Y = 1:aaa.Y = 1:aab.Y = 1:aac.Y = 1:aad.Y = 1:aae.Y = 1:aaf.Y = 1
	ab0.Y = 1:ab1.Y = 1:ab2.Y = 1:ab3.Y = 1:ab4.Y = 1:ab5.Y = 1:ab6.Y = 1:ab7.Y = 1:ab8.Y = 1:ab9.Y = 1:aba.Y = 1:abb.Y = 1:abc.Y = 1:abd.Y = 1:abe.Y = 1:abf.Y = 1
	ac0.Y = 1:ac1.Y = 1:ac2.Y = 1:ac3.Y = 1:ac4.Y = 1:ac5.Y = 1:ac6.Y = 1:ac7.Y = 1:ac8.Y = 1:ac9.Y = 1:aca.Y = 1:acb.Y = 1:acc.Y = 1:acd.Y = 1:ace.Y = 1:acf.Y = 1
	ad0.Y = 1:ad1.Y = 1:ad2.Y = 1:ad3.Y = 1:ad4.Y = 1:ad5.Y = 1:ad6.Y = 1:ad7.Y = 1:ad8.Y = 1:ad9.Y = 1:ada.Y = 1:adb.Y = 1:adc.Y = 1:add.Y = 1:ade.Y = 1:adf.Y = 1
	ae0.Y = 1:ae1.Y = 1:ae2.Y = 1:ae3.Y = 1:ae4.Y = 1:ae5.Y = 1:ae6.Y = 1:ae7.Y = 1:ae8.Y = 1:ae9.Y = 1:aea.Y = 1:aeb.Y = 1:aec.Y = 1:aed.Y = 1:aee.Y = 1:aef.Y = 1
	af0.Y = 1:af1.Y = 1:af2.Y = 1:af3.Y = 1:af4.Y = 1:af5.Y = 1:af6.Y = 1:af7.Y = 1:af8.Y = 1:af9.Y = 1:afa.Y = 1:afb.Y = 1:afc.Y = 1:afd.Y = 1:afe.Y = 1:aff.Y = 1
	b00.Y = 1:b01.Y = 1:b02.Y = 1:b03.Y = 1:b04.Y = 1:b05.Y = 1:b06.Y = 1:b07.Y = 1:b08.Y = 1:b09.Y = 1:b0a.Y = 1:b0b.Y = 1:b0c.Y = 1:b0d.Y = 1:b0e.Y = 1:b0f.Y = 1
	b10.Y = 1:b11.Y = 1:b12.Y = 1:b13.Y = 1:b14.Y = 1:b15.Y = 1:b16.Y = 1:b17.Y = 1:b18.Y = 1:b19.Y = 1:b1a.Y = 1:b1b.Y = 1:b1c.Y = 1:b1d.Y = 1:b1e.Y = 1:b1f.Y = 1
	b20.Y = 1:b21.Y = 1:b22.Y = 1:b23.Y = 1:b24.Y = 1:b25.Y = 1:b26.Y = 1:b27.Y = 1:b28.Y = 1:b29.Y = 1:b2a.Y = 1:b2b.Y = 1:b2c.Y = 1:b2d.Y = 1:b2e.Y = 1:b2f.Y = 1
	b30.Y = 1:b31.Y = 1:b32.Y = 1:b33.Y = 1:b34.Y = 1:b35.Y = 1:b36.Y = 1:b37.Y = 1:b38.Y = 1:b39.Y = 1:b3a.Y = 1:b3b.Y = 1:b3c.Y = 1:b3d.Y = 1:b3e.Y = 1:b3f.Y = 1

	b40.Y = 1:b41.Y = 1:b42.Y = 1:b43.Y = 1:b44.Y = 1:b45.Y = 1:b46.Y = 1:b47.Y = 1:b48.Y = 1:b49.Y = 1:b4a.Y = 1:b4b.Y = 1:b4c.Y = 1:b4d.Y = 1:b4e.Y = 1:b4f.Y = 1
	b50.Y = 1:b51.Y = 1:b52.Y = 1:b53.Y = 1:b54.Y = 1:b55.Y = 1:b56.Y = 1:b57.Y = 1:b58.Y = 1:b59.Y = 1:b5a.Y = 1:b5b.Y = 1:b5c.Y = 1:b5d.Y = 1:b5e.Y = 1:b5f.Y = 1
	b60.Y = 1:b61.Y = 1:b62.Y = 1:b63.Y = 1:b64.Y = 1:b65.Y = 1:b66.Y = 1:b67.Y = 1:b68.Y = 1:b69.Y = 1:b6a.Y = 1:b6b.Y = 1:b6c.Y = 1:b6d.Y = 1:b6e.Y = 1:b6f.Y = 1
	b70.Y = 1:b71.Y = 1:b72.Y = 1:b73.Y = 1:b74.Y = 1:b75.Y = 1:b76.Y = 1:b77.Y = 1:b78.Y = 1:b79.Y = 1:b7a.Y = 1:b7b.Y = 1:b7c.Y = 1:b7d.Y = 1:b7e.Y = 1:b7f.Y = 1
	b80.Y = 1:b81.Y = 1:b82.Y = 1:b83.Y = 1:b84.Y = 1:b85.Y = 1:b86.Y = 1:b87.Y = 1:b88.Y = 1:b89.Y = 1:b8a.Y = 1:b8b.Y = 1:b8c.Y = 1:b8d.Y = 1:b8e.Y = 1:b8f.Y = 1
	b90.Y = 1:b91.Y = 1:b92.Y = 1:b93.Y = 1:b94.Y = 1:b95.Y = 1:b96.Y = 1:b97.Y = 1:b98.Y = 1:b99.Y = 1:b9a.Y = 1:b9b.Y = 1:b9c.Y = 1:b9d.Y = 1:b9e.Y = 1:b9f.Y = 1
	ba0.Y = 1:ba1.Y = 1:ba2.Y = 1:ba3.Y = 1:ba4.Y = 1:ba5.Y = 1:ba6.Y = 1:ba7.Y = 1:ba8.Y = 1:ba9.Y = 1:baa.Y = 1:bab.Y = 1:bac.Y = 1:bad.Y = 1:bae.Y = 1:baf.Y = 1
	bb0.Y = 1:bb1.Y = 1:bb2.Y = 1:bb3.Y = 1:bb4.Y = 1:bb5.Y = 1:bb6.Y = 1:bb7.Y = 1:bb8.Y = 1:bb9.Y = 1:bba.Y = 1:bbb.Y = 1:bbc.Y = 1:bbd.Y = 1:bbe.Y = 1:bbf.Y = 1
	bc0.Y = 1:bc1.Y = 1:bc2.Y = 1:bc3.Y = 1:bc4.Y = 1:bc5.Y = 1:bc6.Y = 1:bc7.Y = 1:bc8.Y = 1:bc9.Y = 1:bca.Y = 1:bcb.Y = 1:bcc.Y = 1:bcd.Y = 1:bce.Y = 1:bcf.Y = 1
	bd0.Y = 1:bd1.Y = 1:bd2.Y = 1:bd3.Y = 1:bd4.Y = 1:bd5.Y = 1:bd6.Y = 1:bd7.Y = 1:bd8.Y = 1:bd9.Y = 1:bda.Y = 1:bdb.Y = 1:bdc.Y = 1:bdd.Y = 1:bde.Y = 1:bdf.Y = 1
	be0.Y = 1:be1.Y = 1:be2.Y = 1:be3.Y = 1:be4.Y = 1:be5.Y = 1:be6.Y = 1:be7.Y = 1:be8.Y = 1:be9.Y = 1:bea.Y = 1:beb.Y = 1:bec.Y = 1:bed.Y = 1:bee.Y = 1:bef.Y = 1
	bf0.Y = 1:bf1.Y = 1:bf2.Y = 1:bf3.Y = 1:bf4.Y = 1:bf5.Y = 1:bf6.Y = 1:bf7.Y = 1:bf8.Y = 1:bf9.Y = 1:bfa.Y = 1:bfb.Y = 1:bfc.Y = 1:bfd.Y = 1:bfe.Y = 1:bff.Y = 1
	c00.Y = 1:c01.Y = 1:c02.Y = 1:c03.Y = 1:c04.Y = 1:c05.Y = 1:c06.Y = 1:c07.Y = 1:c08.Y = 1:c09.Y = 1:c0a.Y = 1:c0b.Y = 1:c0c.Y = 1:c0d.Y = 1:c0e.Y = 1:c0f.Y = 1
	c10.Y = 1:c11.Y = 1:c12.Y = 1:c13.Y = 1:c14.Y = 1:c15.Y = 1:c16.Y = 1:c17.Y = 1:c18.Y = 1:c19.Y = 1:c1a.Y = 1:c1b.Y = 1:c1c.Y = 1:c1d.Y = 1:c1e.Y = 1:c1f.Y = 1
	c20.Y = 1:c21.Y = 1:c22.Y = 1:c23.Y = 1:c24.Y = 1:c25.Y = 1:c26.Y = 1:c27.Y = 1:c28.Y = 1:c29.Y = 1:c2a.Y = 1:c2b.Y = 1:c2c.Y = 1:c2d.Y = 1:c2e.Y = 1:c2f.Y = 1
	c30.Y = 1:c31.Y = 1:c32.Y = 1:c33.Y = 1:c34.Y = 1:c35.Y = 1:c36.Y = 1:c37.Y = 1:c38.Y = 1:c39.Y = 1:c3a.Y = 1:c3b.Y = 1:c3c.Y = 1:c3d.Y = 1:c3e.Y = 1:c3f.Y = 1
	c40.Y = 1:c41.Y = 1:c42.Y = 1:c43.Y = 1:c44.Y = 1:c45.Y = 1:c46.Y = 1:c47.Y = 1:c48.Y = 1:c49.Y = 1:c4a.Y = 1:c4b.Y = 1:c4c.Y = 1:c4d.Y = 1:c4e.Y = 1:c4f.Y = 1
	c50.Y = 1:c51.Y = 1:c52.Y = 1:c53.Y = 1:c54.Y = 1:c55.Y = 1:c56.Y = 1:c57.Y = 1:c58.Y = 1:c59.Y = 1:c5a.Y = 1:c5b.Y = 1:c5c.Y = 1:c5d.Y = 1:c5e.Y = 1:c5f.Y = 1
	c60.Y = 1:c61.Y = 1:c62.Y = 1:c63.Y = 1:c64.Y = 1:c65.Y = 1:c66.Y = 1:c67.Y = 1:c68.Y = 1:c69.Y = 1:c6a.Y = 1:c6b.Y = 1:c6c.Y = 1:c6d.Y = 1:c6e.Y = 1:c6f.Y = 1
	c70.Y = 1:c71.Y = 1:c72.Y = 1:c73.Y = 1:c74.Y = 1:c75.Y = 1:c76.Y = 1:c77.Y = 1:c78.Y = 1:c79.Y = 1:c7a.Y = 1:c7b.Y = 1:c7c.Y = 1:c7d.Y = 1:c7e.Y = 1:c7f.Y = 1

End Sub

Sub Table1_Exit():Controller.Stop:End Sub
Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

' GI Init
Sub GIDelay_Timer()
	SetLamp 160, 1
	SetLamp 126, 1
	GIDelay.Enabled = 0
End Sub

' Glasses Color
 If GlassesColor=1 Then:Glasses.Image="GlassWhite":End If
 If GlassesColor=2 Then:Glasses.Image="GlassYellow":End If
 If GlassesColor=3 Then:Glasses.Image="GlassOrange":End If
 If GlassesColor=4 Then:Glasses.Image="GlassRed":End If
 If GlassesColor=5 Then:Glasses.Image="GlassViolet":End If
 If GlassesColor=6 Then:Glasses.Image="GlassGreen":End If
 If GlassesColor=7 Then:Glasses.Image="GlassBlue":End If

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyCode = LeftFlipperKey Then Controller.Switch(6) = 1
	If KeyCode = RightFlipperKey Then  Controller.Switch(7) = 1
	If KeyCode = PlungerKey Then Plunger.Pullback
	If KeyCode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("nudge_left",0)
	If KeyCode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("nudge_right",0)
	If KeyCode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("nudge_forward",0)
	If vpmKeyDown(KeyCode) Then Exit Sub
    'debug key
    If KeyCode = "3" Then
        SetLamp 171, 1
        SetLamp 172, 1
        SetLamp 175, 1
        SetLamp 190, 1
        SetLamp 197, 1
        SetLamp 198, 1
        SetLamp 199, 1
        f151.State = 1
        f151a.State = 1
    End If
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyCode = LeftFlipperKey Then Controller.Switch(6) = 0
	If KeyCode = RightFlipperKey Then Controller.Switch(7) = 0
	If KeyCode = PlungerKey Then Plunger.Fire:PlaySoundAtVol "plunger2", Plunger, 1
	If vpmKeyUp(KeyCode) Then Exit Sub
    'debug key
    If KeyCode = "3" Then
        SetLamp 171, 0
        SetLamp 172, 0
        SetLamp 175, 0
        SetLamp 190, 0
        SetLamp 197, 0
        SetLamp 198, 0
        SetLamp 199, 0
        f151.State = 0
        f151a.State = 0
    End If
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingshot_Slingshot:If PinPlay=1 Then:vpmTimer.PulseSw 13:LeftSling.Visible=1:SxEmKickerT1.TransX=-28:LStep=0:Me.TimerEnabled=1:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), SxEmKickerT1, 1:End If:End Sub
Sub LeftSlingshot_Timer
	Select Case LStep
		Case 0:LeftSling.Visible = 1
		Case 1: 'pause
		Case 2:LeftSling.Visible = 0 :LeftSling1.Visible = 1:SxEmKickerT1.TransX=-23
		Case 3:LeftSling1.Visible = 0:LeftSling2.Visible = 1:SxEmKickerT1.TransX=-18.5
		Case 4:LeftSling2.Visible = 0:Me.TimerEnabled = 0:SxEmKickerT1.TransX=0
	End Select
	LStep = LStep + 1
End Sub

Sub RightSlingshot_Slingshot:If PinPlay=1 Then:vpmTimer.PulseSw 14:RightSling.Visible=1:DxEmKickerT1.TransX=-28:RStep=0:Me.TimerEnabled=1:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), DxEmKickerT1, 1:End If:End Sub
Sub RightSlingshot_Timer
	Select Case RStep
		Case 0:RightSling.Visible = 1
		Case 1: 'pause
		Case 2:RightSling.Visible = 0 :RightSling1.Visible = 1:DxEmKickerT1.TransX=-23
		Case 3:RightSling1.Visible = 0:RightSling2.Visible = 1:DxEmKickerT1.TransX=-18.5
		Case 4:RightSling2.Visible = 0:Me.TimerEnabled = 0:DxEmKickerT1.TransX=0
	End Select
	RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1_Hit
	If PinPlay=1 Then
	vpmTimer.PulseSw 10:Me.TimerEnabled=0:Ring1.Z = 20:Me.TimerEnabled=1:PlaySoundAtVol SoundFX("jet1",DOFContactors), ActiveBall, VolBump
End If
End Sub
Sub Bumper1_Timer()
	Ring1.Z = Ring1.Z +2
 If Ring1.Z = 50 Then:Me.TimerEnabled = 0
End Sub

Sub Bumper2_Hit
	If PinPlay=1 Then
	vpmTimer.PulseSw 11:Me.TimerEnabled=0:Ring2.Z = 20:Me.TimerEnabled=1:PlaySoundAtVol SoundFX("jet1",DOFContactors), ActiveBall, VolBump
End If
End Sub
Sub Bumper2_Timer()
	Ring2.Z = Ring2.Z +2
 If Ring2.Z = 50 Then:Me.TimerEnabled = 0
End Sub

Sub Bumper3_Hit
	If PinPlay=1 Then
	vpmTimer.PulseSw 12:Me.TimerEnabled=0:Ring3.Z = 20:Me.TimerEnabled=1:PlaySoundAtVol SoundFX("jet1",DOFContactors), ActiveBall, VolBump
End If
End Sub
Sub Bumper3_Timer()
	Ring3.Z = Ring3.Z +2
 If Ring3.Z = 50 Then:Me.TimerEnabled = 0
End Sub

' Spinner
Sub sw20_Spin:vpmTimer.PulseSw 20:PlaySoundAtVol "spinner", sw20, VolSpin:End Sub

' Eject holes
Sub Drain_Hit:bsTrough.AddBall Me:PlaySoundAtVol "drain1a", drain, 1:End Sub
Sub sw21_Hit:PlaySoundAtVol "fx_kicker_enter1", sw21, VolKick:bsUK.AddBall 0:End Sub
Sub sw35_Hit:PlaySoundAtVol "fx_kicker_enter1", sw35, VolKick:bsLK.AddBall 0:End Sub

' Rollovers
Sub sw23_Hit:  Controller.Switch(23) = 1:PlaySoundAtVol "sensor", ActiveBall, VolRol :Psw23.Z=8:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:Psw23.Z=23:End Sub
Sub sw24_Hit:  Controller.Switch(24) = 1:PlaySoundAtVol "sensor", ActiveBall, VolRol :Psw24.Z=8:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:Psw24.Z=23:End Sub
Sub sw30_Hit:  Controller.Switch(30) = 1:PlaySoundAtVol "sensor", ActiveBall, VolRol :Psw30.Z=8:End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:Psw30.Z=23:End Sub
Sub sw31_Hit:  Controller.Switch(31) = 1:PlaySoundAtVol "sensor", ActiveBall, VolRol :End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub
Sub sw33_Hit:  Controller.Switch(33) = 1:PlaySoundAtVol "sensor", ActiveBall, VolRol :Psw33.Z=8:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:Psw33.Z=23:End Sub
Sub sw34_Hit:  Controller.Switch(34) = 1:PlaySoundAtVol "sensor", ActiveBall, VolRol :Psw34.Z=8:End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:Psw34.Z=23:End Sub

'Ramp sensors
Sub sw50_Hit:  Controller.Switch(50) = 1:End Sub
Sub sw50_UnHit:Controller.Switch(50) = 0:End Sub
Sub sw51_Hit:  Controller.Switch(51) = 1:End Sub
Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub
Sub sw52_Hit:  Controller.Switch(52) = 1:End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub
Sub sw52a_Hit:  Controller.Switch(52) = 1:End Sub
Sub sw52a_UnHit:Controller.Switch(52) = 0:End Sub

' Targets
Sub sw32_Hit:vpmTimer.PulseSw 32:Psw32.TransX=-5:Me.TimerEnabled=1:PlaySoundAtVol SoundFX("target",DOFContactors), ActiveBall, VolTarg:End Sub
Sub sw32_Timer:Psw32.TransX=0:Me.TimerEnabled=0:End Sub

' Kicking Targets
Sub sw15_Slingshot:vpmTimer.PulseSw 15:KickingTsw15.RotY=8:Me.TimerEnabled=1:PlaySoundAtVol SoundFX("target",DOFContactors), ActiveBall, VolTarg:End Sub
Sub sw15_Timer:KickingTsw15.RotY=0:Me.TimerEnabled=0:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:KickingTsw22.RotY=8:Me.TimerEnabled=1:PlaySoundAtVol SoundFX("target",DOFContactors), ActiveBall, VolTarg:End Sub
Sub sw22_Timer:KickingTsw22.RotY=0:Me.TimerEnabled=0:End Sub

' Droptargets
Sub sw16_Hit():PlaySoundAtVol SoundFX("DROPTARG",DOFContactors), ActiveBall, 1:End Sub
Sub sw16_Dropped:cdtbank.Hit 1: End Sub

Sub sw26_Hit():PlaySoundAtVol SoundFX("DROPTARG",DOFContactors), ActiveBall, 1:End Sub
Sub sw26_Dropped:cdtbank.Hit 2:End Sub

Sub sw36_Hit():PlaySoundAtVol SoundFX("DROPTARG",DOFContactors), ActiveBall, 1:End Sub
Sub sw36_Dropped:cdtbank.Hit 3:End Sub

Sub sw17_Hit():PlaySoundAtVol SoundFX("DROPTARG",DOFContactors), ActiveBall, 1
 If Bulb12.State=1 Then
	Bulb12T17.State=1
Else
	Bulb12T17.State=0
End If
	Drop17T=1
End Sub
Sub sw17_Dropped:ddtbank.Hit 1:End Sub

Sub sw27_Hit():PlaySoundAtVol SoundFX("DROPTARG",DOFContactors), ActiveBall, 1
 If Bulb12.State=1 Then
	Bulb12T27.State=1
Else
	Bulb12T27.State=0
End If
	Drop27T=1
End Sub
Sub sw27_Dropped:ddtbank.Hit 2:End Sub

Sub sw37_Hit():PlaySoundAtVol SoundFX("DROPTARG",DOFContactors), ActiveBall, 1
 If Bulb11.State=1 Then
	Bulb11T37.State=1
Else
	Bulb11T37.State=0
End If
	Drop37T=1
End Sub
Sub sw37_Dropped:ddtbank.Hit 3:End Sub

' Gates
Sub Gate1_Hit:PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub

' Fx Sounds
Sub Trigger1_Hit:PlaySoundAtVol "WireRolling", ActiveBall, 1 :End Sub
Sub Trigger2_Hit:PlaySoundAtVol "WireRolling", ActiveBall, 1 :End Sub
Sub EndStop_Hit :PlaySoundAtVol "WireHit", ActiveBall, 1 :StopSound "WireRolling":End Sub
Sub EndStop1_Hit:PlaySoundAtVol "WireHit", ActiveBall, 1 :StopSound "WireRolling":End Sub

Sub TriggerFlapSx_Hit:PlaySoundAtVol "FlapHit2", ActiveBall, 1 :End Sub
Sub TriggerFlapDx_Hit:PlaySoundAtVol "FlapHit2", ActiveBall, 1 :End Sub
Sub TriggerOVUK_Hit:PlaySoundAtVol "WoodHitD225",ActiveBall, 1:End Sub
Sub Trigger4_Hit:PlaySoundAtVol "PlasticJump",ActiveBall, 1:End Sub
Sub StopVUK2_Hit:PlaySoundAtVol "WireHit",ActiveBall, 1:End Sub
Sub KickerW_Hit:PlaySoundAtVol "kicker_enter",ActiveBall, 1:End Sub
Sub TriggerOVUK1_Hit:PlaySoundAtVol "kicker_enter",ActiveBall, 1:End Sub
Sub TriggerORamp_Hit:PlaySoundAtVol "kicker_enter",ActiveBall, 1:End Sub
Sub CupFX1_Hit():ActiveBall.VelY=ActiveBall.VelY+0.1:PlaySoundAtVol "PlasticJump",ActiveBall, 1:End Sub
Sub CupFX2_Hit():ActiveBall.VelY=ActiveBall.VelY-0.1:PlaySoundAtVol "PlasticJump",ActiveBall, 1:End Sub

' add additional (optional) parameters to PlaySound to increase/decrease the frequency,
' apply all the settings to an already playing sample and choose if to restart this sample from the beginning or not
' PlaySound "name",loopcount,volume,pan,randompitch,pitch,useexisting,restart
' pan ranges from -1.0 (left) over 0.0 (both) to 1.0 (right)
' randompitch ranges from 0.0 (no randomization) to 1.0 (vary between half speed to double speed)
' pitch can be positive or negative and directly adds onto the standard sample frequency
' useexisting is 0 or 1 (if no existing/playing copy of the sound is found, then a new one is created)
' restart is 0 or 1 (only useful if useexisting is 1)

'*********
'Solenoids
'*********

SolCallback(7) = "bsLKBallRelease"
SolCallback(8) = "dtcbank"
SolCallback(9) = "dtdbank"
SolCallback(10) = "bsUK.SolOut"
SolCallback(11) = "setlamp 131,"
SolCallback(12) = "setlamp 132,"
SolCallback(13) = "setlamp 133,"
SolCallback(14) = "setlamp 134,"
SolCallback(15) = "setlamp 125,"
SolCallback(16) = "setlamp 150,"	'lamp
SolCallback(17) = "setlamp 197,"
SolCallback(18) = "setlamp 198,"
SolCallback(19) = "setlamp 199,"
SolCallback(20) = "setlamp 190,"
SolCallback(21) = "setlamp 171,"
SolCallback(22) = "setlamp 172,"
SolCallback(23) = "setlamp 122,"
SolCallback(24) = "setlamp 151,"
SolCallback(25) = "setlamp 175,"
SolCallback(26) = "Lightbox"	' Lightbox Insert Illum. Relay (A)
SolCallback(28) = "bsTrough.SolOut"
SolCallback(29) = "bsTrough.SolIn"
SolCallback(30) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(31) = "GIRelay"
Solcallback(32) = "SolRun"

Sub dtcbank(Enabled)
 If Enabled Then
	cdtbank.DropSol_On
End If
	PlaySound SoundFX("DTResetB",DOFContactors) ' TODO
End Sub

Sub dtdbank(Enabled)
 If Enabled Then
	ddtbank.DropSol_On
End If
	Drop17T=0
	Drop27T=0
	Drop37T=0
	Bulb12T17.TimerEnabled=1
	PlaySound SoundFX("DTResetB",DOFContactors),0,1,0.2
End Sub

Sub bsLKBallRelease(Enabled)
 If Enabled Then
	bsLK.ExitSol_On
	Arm.IsDropped=0
	Arm.TimerEnabled=1
	Arm2.IsDropped=0
End If
End Sub

Sub Arm_Timer
	Arm.IsDropped=1
	Arm.TimerEnabled=0
	Arm2.IsDropped=1
End Sub

Sub Bulb12T17_Timer()
	Bulb12T17.State=0
	Bulb12T27.State=0
	Bulb11T37.State=0
	Bulb12T17.TimerEnabled=0
End Sub

'**************
' GI
'**************

Dim Drop17T, Drop27T, Drop37T
Sub GIRelay(Enabled)
	Dim GIoffon
	GIoffon = ABS(ABS(Enabled) -1)
	SetLamp 160, GIoffon
 If Drop17T=1 Then
	Bulb12T17.State=GIoffon
End If
 If Drop27T=1 Then
	Bulb12T27.State=GIoffon
End If
 If Drop37T=1 Then
	Bulb11T37.State=GIoffon
End If
End Sub

Sub SolRun(Enabled)
	vpmNudge.SolGameOn Enabled
 If Enabled Then
	PinPlay=1
Else
	PinPlay=0
'	LeftFlipper.RotateToStart
'	RightFlipper.RotateToStart
End If
End Sub

Sub Lightbox(Enabled)
	Dim GIoffon
	GIoffon = ABS(ABS(Enabled) -1)
	SetLamp 126, GIoffon
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
	If Enabled Then
		PlaySoundAtVol SoundFX("flipperup_left",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToEnd
	Else
        PlaySoundAtVol SoundFX("flipperdown_left",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToStart
	End If
End Sub

Sub SolRFlipper(Enabled)
	If Enabled Then
		PlaySoundAtVol SoundFX("flipperup_right",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToEnd
	Else
        PlaySoundAtVol SoundFX("flipperdown_right",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToStart
	End If
End Sub

Sub LeftFlipper_Collide(parm)
	PlaySoundAtBallVol "rubber_flipper",1
End Sub

Sub RightFlipper_Collide(parm)
	PlaySoundAtBallVol "rubber_flipper",1
End Sub

'**********************************
'       JP's VP10 Fading Lamps & Flashers v2
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
    UpdateLamps
    UpdateLeds
End Sub

Sub UpdateLamps
'	FadeL 0, l0, "On", "F66", "F33", "Off"
	NFadeL 0, l0
	NFadeL 1, l1
	NFadeL 2, l2
	NFadeL 3, l3
	NFadeL 4, l4
	NFadeL 5, l5
	NFadeL 6, l6
	NFadeL 7, l7
	NFadeL 10, l10
	NFadeL 11, l11
	NFadeL 12, l12
	NFadeL 13, l13
	NFadeL 14, l14
	NFadeL 15, l15
	NFadeL 16, l16
	NFadeL 17, l17
	NFadeL 20, l20
	NFadeL 21, l21
	NFadeL 22, l22
	NFadeL 23, l23
	NFadeL 24, l24
	NFadeL 25, l25
	NFadeL 26, l26
	NFadeL 27, l27
	NFadeL 30, l30
	NFadeL 31, l31
	NFadeL 32, l32
	NFadeL 33, l33
	NFadeL 34, l34
	NFadeL 35, l35
	NFadeL 36, l36
	NFadeL 37, l37
	NFadeL 40, l40
	NFadeL 41, l41

	NFadeL 67, l67
	NFadeL 70, l70
	NFadeL 71, l71
	NFadeL 72, l72
	NFadeL 73, l73
	NFadeL 74, l74
	NFadeL 75, l75
	NFadeL 76, l76
	NFadeL 80, l80
	NFadeL 81, l81
	NFadeL 82, l82
	NFadeL 83, l83
	NFadeL 84, l84
	NFadeL 85, l85
	NFadeL 86, l86
	NFadeL 90, l90
	NFadeL 91, l91
	NFadeL 92, l92
	NFadeL 93, l93
	NFadeL 94, l94
	NFadeL 95, l95
	NFadeL 96, l96
	NFadeL 100, l100
	NFadeL 101, l101
	NFadeL 102, l102
	NFadeL 103, l103
	NFadeL 104, l104
	NFadeL 105, l105
	NFadeL 106, l106
	NFadeL 110, l110
	NFadeL 111, l111
	NFadeL 112, l112
	NFadeL 113, l113
	NFadeL 114, l114
	NFadeL 115, l115
	NFadeL 116, l116
	NFadeL 122, l122
	NFadeL 125, l125

	Flash 42, f42
	Flashm 43, f43
	Flash 43, f43a
	FastFlash 44, f44
	FastFlash 45, f45
	FastFlash 46, f46
	FastFlash 47, f47
	Flashm 50, f50
	Flash 50, f50a
	Flash 51, f51
	FastFlash 52, f52
	FastFlash 53, f53
	FastFlash 54, f54
	FastFlash 55, f55
	Flashm 56, f56
	Flash 56, f56a
	Flashm 57, f57
	Flash 57, f57a
	FastFlashm 66, f66
	FastFlash 66, f66a
	NFadeLm 150, fur1
	NFadeLm 150, fur2
	Flashm 150, fur1a
	Flash 150, fur2a
	NFadeLm 151, f151
	NFadeL 151, f151a

	NFadeLm 160, Bulb1
	NFadeLm 160, Bulb1a
	NFadeLm 160, Bulb2
	NFadeLm 160, Bulb2a
	NFadeLm 160, Bulb3
	NFadeLm 160, Bulb3a
	NFadeLm 160, Bulb4
	NFadeLm 160, Bulb4a
	NFadeLm 160, Bulb5
	NFadeLm 160, Bulb5a
	NFadeLm 160, Bulb6
	NFadeLm 160, Bulb6a
	NFadeLm 160, Bulb7
	NFadeLm 160, Bulb7a
	NFadeLm 160, Bulb8
	NFadeLm 160, Bulb8a
	NFadeLm 160, Bulb9
	NFadeLm 160, Bulb9a
	NFadeLm 160, Bulb10
	NFadeLm 160, Bulb10a
	NFadeLm 160, Bulb11
	NFadeLm 160, Bulb11a
	NFadeLm 160, Bulb12
	NFadeLm 160, Bulb12a
	NFadeLm 160, Bulb13
	NFadeLm 160, Bulb13a
	NFadeLm 160, Bulb14
	NFadeLm 160, Bulb14a

	Flashm 160, fgit1
	Flashm 160, fgit1L
	Flashm 160, fgit1R
	Flashm 160, fgit2
	Flashm 160, fgit3
	Flashm 160, fgit4
	Flashm 160, fgit5
	Flashm 160, fgit6
	Flashm 160, fgit7
	Flashm 160, fgit8
	Flashm 160, fgit9
	Flashm 160, fgit10
	Flashm 160, fgit11
	Flashm 160, fgit12
	Flashm 160, fgit13
	Flashm 160, fgit14
	Flash 160, Bulb10b
	Flash 171, f171
	Flash 172, f172
	Flash 175, f175
	Flash 190, f190
	Flash 197, f197
	Flash 198, f198
	NFadeLm 199, l199
	Flash 199, f199

' Backbox

	Flash 60, f60
	Flash 61, f61
	Flash 62, f62
	Flash 63, f63
	Flash 64, f64
	Flash 65, f65

	Flash 77, f77
	Flash 87, f87
	Flash 97, f97
	Flash 107, f107
	Flash 117, f117

	Flashm 131, f131
	Flash 131, f131a
	Flashm 132, f132
	Flash 132, f132a
	Flashm 133, f133
	Flash 133, f133a
	Flashm 134, f134
	Flash 134, f134a

	Flashm 126, flb1
	Flashm 126, flb2
	Flashm 126, flb3
	Flashm 126, flb4
	Flashm 126, flb5
	Flashm 126, flb6
	Flashm 126, flb7
	Flashm 126, flb8
	Flashm 126, flb9
	Flashm 126, flb10
	Flashm 126, flb11
	Flashm 126, flb12
	Flashm 126, flb13
	Flashm 126, flb14
	Flashm 126, flb15
	Flashm 126, flb16
	Flashm 126, flb17
	Flashm 126, flb18
	Flashm 126, flb19
	Flashm 126, flb20
	Flashm 126, flb21
	Flashm 126, flb22
	Flashm 126, flb23
	Flashm 126, flb24
	Flashm 126, flb25
	Flashm 126, flb26
	Flashm 126, flb27
	Flash 126, flb28

End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.5   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.07	' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

Sub SetModLamp(nr, level)
	FlashLevel(nr) = level /150 'lights & flashers
End Sub

' Lights: old method, using 4 images

Sub FadeL(nr, light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:light.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:light.image = a:light.State = 1:FadingLevel(nr) = 1   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1           'wait
        Case 9:light.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1        'wait
        Case 13:light.image = d:Light.State = 0:FadingLevel(nr) = 0  'Off
    End Select
End Sub

Sub FadeLm(nr, light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:light.image = b
        Case 5:light.image = a
        Case 9:light.image = c
        Case 13:light.image = d
    End Select
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself

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

Sub LightMod(nr, object) ' modulated lights used as flashers
    Object.IntensityScale = FlashLevel(nr)
	Object.State = 1
End Sub

'Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = Lumen*FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = Lumen*FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = Lumen*FlashLevel(nr)
End Sub

Sub FlashMod(nr, object) 'sets the flashlevel from the SolModCallback
    Object.IntensityScale = Lumen*FlashLevel(nr)
End Sub

Sub FastFlash(nr, object)
    Select Case FadingLevel(nr)
		Case 4:object.Visible = 0:FadingLevel(nr) = 0 'off
		Case 5:object.Visible = 1:FadingLevel(nr) = 1 'on
		Object.IntensityScale = Lumen*FlashMax(nr)
    End Select
End Sub

Sub FastFlashm(nr, object)
    Select Case FadingLevel(nr)
		Case 4:object.Visible = 0 'off
		Case 5:object.Visible = 1 'on
		Object.IntensityScale = Lumen*FlashMax(nr)
    End Select
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

' *********************************************************************
' 					Wall, rubber and metal hit sounds
' *********************************************************************

Sub Rubbers_Hit(idx):PlaySoundAtVol "rubber1", ActiveBall, 1 :End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

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

Const tnob = 4 ' total number of balls
Const lob = 0   'number of locked balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <87 Then		'Z Rolling limit
                ballpitch = Pitch(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 15000 'increase the pitch on a ramp or elevated surface
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
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

'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
	UpdateFlipperLogos
	RollingUpdate
	TrackSounds
End Sub

Dim PI
PI = Round(4 * Atn(1), 6) '3.1415926535897932384626433832795

Sub UpdateFlipperLogos
	FlipperSx.RotZ = LeftFlipper.CurrentAngle
	FlipperDx.RotZ = RightFlipper.CurrentAngle
	pSpinnerRod.TransY = sin( (sw20.CurrentAngle+180) * (2*PI/360)) * 8
	pSpinnerRod.TransZ = sin( (sw20.CurrentAngle- 90) * (2*PI/360)) * 8
	pSpinnerRod.RotX = sin( sw20.CurrentAngle * (2*PI/360)) * 6
End Sub

'*****************
' Leds Display
'*****************

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

Sub UpdateLeds
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
	If Table1.ShowFSS = True Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For Each obj In Digits(num)
                If chg And 1 Then obj.Visible = stat And 1
                chg = chg \ 2:stat = stat \ 2
            Next
        Next
    End If
    End If
End Sub

'********************************
' Sound Subs from Destruk's table
'********************************

Dim Playing
Playing = 0

'Music & Sound Stuff
Sub TrackSounds
    Dim NewSounds, ii, Snd
    NewSounds = Controller.NewSoundCommands
    If Not IsEmpty(NewSounds) Then
        For ii = 0 To UBound(NewSounds)
            Snd = NewSounds(ii, 0)
            If Snd = 58 Then ShootGently:LegsAnim
'            If Snd = 28 Then Heheheeee
        Next
    End If
End Sub

'Shoot Gently
Dim AnimStep
Sub ShootGently:AnimStep=1:TimerJaw.Enabled=1:End Sub
Sub TimerJaw_Timer()
	Select Case AnimStep
		Case 1:Jaw.RotX = Jaw.RotX - 1
			If Jaw.RotX = -8 Then:AnimStep=2:End If
		Case 2:Jaw.RotX = Jaw.RotX + 1
			If Jaw.RotX = 0 Then:AnimStep=3:End If
		Case 3:Jaw.RotX = Jaw.RotX - 1
			If Jaw.RotX = -8 Then:AnimStep=4:End If
		Case 4:Jaw.RotX = Jaw.RotX + 1
			If Jaw.RotX = 0 Then:AnimStep=5:End If
		Case 5:Jaw.RotX = Jaw.RotX - 1
			If Jaw.RotX = -8 Then:AnimStep=6:End If
		Case 6:Jaw.RotX = Jaw.RotX + 1
			If Jaw.RotX = 0 Then:TimerJaw.Enabled = 0:End If
	End Select
End Sub

Dim AnimStep1
Sub LegsAnim:AnimStep1=1:TimerLegs.Enabled=1:End Sub
Sub TimerLegs_Timer()
	Select Case AnimStep1
		Case 1:LegSx.RotX = LegSx.RotX - 1:LegDx.RotX = LegDx.RotX + 1
			If LegSx.RotX = -60 Then:AnimStep1=2:End If
		Case 2:LegSx.RotX = LegSx.RotX + 2:LegDx.RotX = LegDx.RotX - 2
			If LegSx.RotX = -40 Then:AnimStep1=3:End If
		Case 3:LegSx.RotX = LegSx.RotX - 2:LegDx.RotX = LegDx.RotX + 2
			If LegSx.RotX = -70 Then:AnimStep1=4:End If
		Case 4:LegSx.RotX = LegSx.RotX + 1:LegDx.RotX = LegDx.RotX - 1
			If LegSx.RotX = -0 Then:TimerLegs.Enabled = 0:End If
	End Select
End Sub

'He he heeee
'Dim AnimStep1
'Sub Heheheeee:AnimStep1=1:TimerJaw.Enabled=1:End Sub
'Sub TimerJaw_Timer
'	Select Case AnimStep1
'		Case 1:Jaw.RotX = -6:AnimStep1=2
'		Case 2:Jaw.RotX = 0 :AnimStep1=3
'		Case 3:Jaw.RotX = -6:AnimStep1=4
'		Case 4:Jaw.RotX = 0 :AnimStep1=5
'		Case 5:Jaw.RotX = -6:AnimStep1=6
'		Case 6:AnimStep1=7 'pause
'		Case 7:AnimStep1=8 'pause
'		Case 8:Jaw.RotX = 0:TimerJaw.Enabled = 0
'	End Select
'End Sub


'006  06-Excellent
'007  07-Extraball
'020  14-Million
'021  15-Advance souvenirs target for extraball
'023  17-The arcade target is lit
'024  18-Lit the birth
'025  19-You wont to pet my monkey
'026  1a-Ho ho ho ho ha ha ha ha
'027  1b-Le target a ?
'028  1c-Nice shoot man
'029  1d-Shoot the rapides to light the whirlpool miilion
'030  1e-The whirlpool million is lit
'033  21-You got absolutly nothing
'034  22-You advance the boomerang
'035  23-You advance the splash
'036  24-You advance the whirlpool
'037  25-You advance the rapides
'038  26-You advance the pipeline
'039  27-Shoot the ball
'040  28-He he heeee
'046  2e-Shoot all lights lit to light super jackpot
'047  2f-Shoot all lights lit to light jackpot
'048  30-Shoot the rapides for super jackpot
'049  31-Shoot the rapides for jackpot
'051  33-Shoot all lights lit for extraball
'052  34-Shoot the spinner
'053  35-Your time is up
'054  36-Hit the three million target
'056  38-Hey man shoot again
'057  39-You shoot to hard
'058  3a-Shoot gently
'059  3b-Shoot the pipeline for multiball
'060  3c-One more for the record
'061  3d-One more for the whirlpool
'062  3e-One more for the boomerang
'063  3f-One more for the splash
'065  41-One more for the rapides
'066  42-One more for the pipeline
'067  43-One more for double
'068  44-you shoot too sofly
'070  46-Extra special
'071  47-Super jackpot
'072  48-super score
'073  49-Three million
'074  4a-Super
'077  4d-Rodney
'078  4e-Two
'079  4f-Three
'080  50-Four
'081  51-Five
'085  55-Got all animals for special
'089  59-Oh noo
'105  69-Another coin please
'107  6b-Jackpot
'109  6d-Feel the power
