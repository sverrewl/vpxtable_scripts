Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=19 to 2
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

Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.



On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="empsback",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin",cCredits=""

LoadVPM "01120100","Hankin.vbs",3.02

Dim Starpic

'OPTION FOR PIC DISPLAY - PICS ARE ON BY DEFAULT
'*****************************************************************************************************
    Starpic = 1          'THIS TURNS PICS ON OR OFF    0 turns pics off   1 turns them on
'*****************************************************************************************************

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
  Ramp16.visible=1
  Ramp15.visible=1
  Primitive13.visible=1
  Flasherlight27.visible=1
  Primitive52.visible=1
  Flasherlight28.visible=0
  Flasherlight29.visible=0
Else
  Ramp16.visible=0
  Ramp15.visible=0
  Primitive13.visible=0
  Flasherlight27.visible=0
  Flasherlight28.visible=1
  Flasherlight29.visible=1
  Primitive52.visible=0
  Flasherflash28.height = 155:Flasherflash28.y = -50:Flasherflash28.x = 760
  Flasherflash29.height = 155:Flasherflash29.y = -65:Flasherflash29.x = 635
  Flasherflash30.height = 155:Flasherflash30.y = -35:Flasherflash30.x = 770
  Flasherflash31.height = 155:Flasherflash31.y = -55:Flasherflash31.x = 470
  Flasherflash32.height = 155:Flasherflash32.y = 31:Flasherflash32.x = 635
  Flasherflash33.height = 155:Flasherflash33.y = -55:Flasherflash33.x = 635
  Flasherflash34.height = 155:Flasherflash34.y = -20:Flasherflash34.x = 625
  Flasherflash35.height = 155:Flasherflash35.y = -60:Flasherflash35.x = 640
  Flasherflash36.height = 155:Flasherflash36.y = -50:Flasherflash36.x = 470
  Flasherflash37.height = 155:Flasherflash37.y = -52:Flasherflash37.x = 620
  Flasherflash38.height = 155:Flasherflash38.y = -52:Flasherflash38.x = 635
  Flasherflash39.height = 155:Flasherflash39.y = 0:Flasherflash39.x = 750
  Flasherflash40.height = 155:Flasherflash40.y = -52:Flasherflash40.x = 750
  Flasherflash41.height = 155:Flasherflash41.y = -52:Flasherflash41.x = 470
  Flasherflash42.height = 155:Flasherflash42.y = -67:Flasherflash42.x = 470
  Flasherflash43.height = 155:Flasherflash43.y = -45:Flasherflash43.x = 470
  Flasherflash44.height = 155:Flasherflash44.y = -30:Flasherflash44.x = 635
End if

'Fluppers Flasher Script
Dim FlashLevel1, FlashLevel2, FlashLevel3, FlashLevel4, FlashLevel5, FlashLevel6, FlashLevel7, FlashLevel8, FlashLevel9, FlashLevel10
Dim FlashLevel11, FlashLevel12, FlashLevel13, FlashLevel14, FlashLevel15, FlashLevel16, FlashLevel17, FlashLevel18, FlashLevel19, FlashLevel20
Dim FlashLevel21, FlashLevel22, FlashLevel23, FlashLevel24, FlashLevel25, FlashLevel26, FlashLevel27, FlashLevel28, FlashLevel29, FlashLevel30
Dim FlashLevel31, FlashLevel32, FlashLevel33, FlashLevel34, FlashLevel35, FlashLevel36
FlasherLight1.IntensityScale = 0
Flasherlight2.IntensityScale = 0
Flasherlight3.IntensityScale = 0
Flasherlight4.IntensityScale = 0
Flasherlight5.IntensityScale = 0
Flasherlight6.IntensityScale = 0
Flasherlight7.IntensityScale = 0
Flasherlight8.IntensityScale = 0
Flasherlight9.IntensityScale = 0
Flasherlight10.IntensityScale = 0
Flasherlight11.IntensityScale = 0
Flasherlight12.IntensityScale = 0
Flasherlight13.IntensityScale = 0
Flasherlight14.IntensityScale = 0
Flasherlight15.IntensityScale = 0
Flasherlight16.IntensityScale = 0
Flasherlight17.IntensityScale = 0
Flasherlight18.IntensityScale = 0
Flasherlight19.IntensityScale = 0
Flasherlight20.IntensityScale = 0
Flasherlight21.IntensityScale = 0
Flasherlight22.IntensityScale = 0
Flasherlight23.IntensityScale = 0
Flasherlight24.IntensityScale = 0
Flasherlight25.IntensityScale = 0
Flasherlight26.IntensityScale = 0
Flasherlight27.IntensityScale = 0
Flasherlight28.IntensityScale = 0
Flasherlight29.IntensityScale = 0
Flasherlight30.IntensityScale = 0
Flasherlight31.IntensityScale = 0
Flasherlight32.IntensityScale = 0
Flasherlight33.IntensityScale = 0
Flasherlight34.IntensityScale = 0
Flasherlight35.IntensityScale = 0
Flasherlight36.IntensityScale = 0



'*** white flasher ***
Sub FlasherFlash1_Timer()
	dim flashx3, matdim
	If not FlasherFlash1.TimerEnabled Then
		FlasherFlash1.TimerEnabled = True
		FlasherFlash1.visible = 1
		FlasherLit1.visible = 1
	End If
	flashx3 = FlashLevel1 * FlashLevel1 * FlashLevel1
	Flasherflash1.opacity = 1000 * flashx3
	FlasherLit1.BlendDisableLighting = 10 * flashx3
	Flasherbase1.BlendDisableLighting =  flashx3
	FlasherLight1.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel1)
	FlasherLit1.material = "domelit" & matdim
	FlashLevel1 = FlashLevel1 * 0.85 - 0.01
	If FlashLevel1 < 0.15 Then
		FlasherLit1.visible = 0
	Else
		FlasherLit1.visible = 1
	end If
	If FlashLevel1 < 0 Then
		FlasherFlash1.TimerEnabled = False
		FlasherFlash1.visible = 0
	End If
End Sub

'*** Red flasher ***
Sub FlasherFlash2_Timer()
	dim flashx3, matdim
	If not Flasherflash2.TimerEnabled Then
		Flasherflash2.TimerEnabled = True
		Flasherflash2.visible = 1
		Flasherlit2.visible = 1
	End If
	flashx3 = FlashLevel2 * FlashLevel2 * FlashLevel2
	Flasherflash2.opacity = 1500 * flashx3
	Flasherlit2.BlendDisableLighting = 10 * flashx3
	Flasherbase2.BlendDisableLighting =  flashx3
	Flasherlight2.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel2)
	Flasherlit2.material = "domelit" & matdim
	FlashLevel2 = FlashLevel2 * 0.9 - 0.01
	If FlashLevel2 < 0.15 Then
		Flasherlit2.visible = 0
	Else
		Flasherlit2.visible = 1
	end If
	If FlashLevel2 < 0 Then
		Flasherflash2.TimerEnabled = False
		Flasherflash2.visible = 0
	End If
End Sub

'*** blue flasher ***
Sub FlasherFlash3_Timer()
	dim flashx3, matdim
	If not Flasherflash3.TimerEnabled Then
		Flasherflash3.TimerEnabled = True
		Flasherflash3.visible = 1
		Flasherlit3.visible = 1
	End If
	flashx3 = FlashLevel3 * FlashLevel3 * FlashLevel3
	Flasherflash3.opacity = 8000 * flashx3
	Flasherlit3.BlendDisableLighting = 10 * flashx3
	Flasherbase3.BlendDisableLighting =  flashx3
	Flasherlight3.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel3)
	Flasherlit3.material = "domelit" & matdim
	FlashLevel3 = FlashLevel3 * 0.85 - 0.01
	If FlashLevel3 < 0.15 Then
		Flasherlit3.visible = 0
	Else
		Flasherlit3.visible = 1
	end If
	If FlashLevel3 < 0 Then
		Flasherflash3.TimerEnabled = False
		Flasherflash3.visible = 0
	End If
End Sub

Sub FlasherFlash4_Timer()
	dim flashx3, matdim
	If not FlasherFlash4.TimerEnabled Then
		FlasherFlash4.TimerEnabled = True
		FlasherFlash4.visible = 1
		FlasherLit4.visible = 1
	End If
	flashx3 = FlashLevel4 * FlashLevel4 * FlashLevel4
	Flasherflash4.opacity = 1000 * flashx3
	FlasherLit4.BlendDisableLighting = 10 * flashx3
	Flasherbase4.BlendDisableLighting =  flashx3
	FlasherLight4.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel4)
	FlasherLit4.material = "domelit" & matdim
	FlashLevel4 = FlashLevel4 * 0.85 - 0.01
	If FlashLevel4 < 0.15 Then
		FlasherLit4.visible = 0
	Else
		FlasherLit4.visible = 1
	end If
	If FlashLevel4 < 0 Then
		FlasherFlash4.TimerEnabled = False
		FlasherFlash4.visible = 0
	End If
End Sub

Sub FlasherFlash5_Timer()
	dim flashx3, matdim
	If not FlasherFlash5.TimerEnabled Then
		FlasherFlash5.TimerEnabled = True
		FlasherFlash5.visible = 1
		FlasherLit5.visible = 1
	End If
	flashx3 = FlashLevel5 * FlashLevel5 * FlashLevel5
	Flasherflash5.opacity = 1000 * flashx3
	FlasherLit5.BlendDisableLighting = 10 * flashx3
	Flasherbase5.BlendDisableLighting =  flashx3
	FlasherLight5.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel5)
	FlasherLit5.material = "domelit" & matdim
	FlashLevel5 = FlashLevel5 * 0.85 - 0.01
	If FlashLevel5 < 0.15 Then
		FlasherLit5.visible = 0
	Else
		FlasherLit5.visible = 1
	end If
	If FlashLevel5 < 0 Then
		FlasherFlash5.TimerEnabled = False
		FlasherFlash5.visible = 0
	End If
End Sub


Sub FlasherFlash6_Timer()
	dim flashx3, matdim
	If not Flasherflash6.TimerEnabled Then
		Flasherflash6.TimerEnabled = True
		Flasherflash6.visible = 1
		Flasherlit6.visible = 1
	End If
	flashx3 = FlashLevel6 * FlashLevel6 * FlashLevel6
	Flasherflash6.opacity = 1500 * flashx3
	Flasherlit6.BlendDisableLighting = 10 * flashx3
	Flasherbase6.BlendDisableLighting =  flashx3
	Flasherlight6.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel6)
	Flasherlit6.material = "domelit" & matdim
	FlashLevel6 = FlashLevel6 * 0.9 - 0.01
	If FlashLevel6 < 0.15 Then
		Flasherlit6.visible = 0
	Else
		Flasherlit6.visible = 1
	end If
	If FlashLevel6 < 0 Then
		Flasherflash6.TimerEnabled = False
		Flasherflash6.visible = 0
	End If
End Sub


Sub FlasherFlash7_Timer()
	dim flashx3, matdim
	If not Flasherflash7.TimerEnabled Then
		Flasherflash7.TimerEnabled = True
		Flasherflash7.visible = 1
		Flasherlit7.visible = 1
	End If
	flashx3 = FlashLevel7 * FlashLevel7 * FlashLevel7
	Flasherflash7.opacity = 1500 * flashx3
	Flasherlit7.BlendDisableLighting = 10 * flashx3
	Flasherbase7.BlendDisableLighting =  flashx3
	Flasherlight7.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel7)
	Flasherlit7.material = "domelit" & matdim
	FlashLevel7 = FlashLevel7 * 0.9 - 0.01
	If FlashLevel7 < 0.15 Then
		Flasherlit7.visible = 0
	Else
		Flasherlit7.visible = 1
	end If
	If FlashLevel7 < 0 Then
		Flasherflash7.TimerEnabled = False
		Flasherflash7.visible = 0
	End If
End Sub


Sub FlasherFlash8_Timer()
	dim flashx3, matdim
	If not Flasherflash8.TimerEnabled Then
		Flasherflash8.TimerEnabled = True
		Flasherflash8.visible = 1
		Flasherlit8.visible = 1
	End If
	flashx3 = FlashLevel8 * FlashLevel8 * FlashLevel8
	Flasherflash8.opacity = 1500 * flashx3
	Flasherlit8.BlendDisableLighting = 10 * flashx3
	Flasherbase8.BlendDisableLighting =  flashx3
	Flasherlight8.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel8)
	Flasherlit8.material = "domelit" & matdim
	FlashLevel8 = FlashLevel8 * 0.9 - 0.01
	If FlashLevel8 < 0.15 Then
		Flasherlit8.visible = 0
	Else
		Flasherlit8.visible = 1
	end If
	If FlashLevel8 < 0 Then
		Flasherflash8.TimerEnabled = False
		Flasherflash8.visible = 0
	End If
End Sub


Sub FlasherFlash9_Timer()
	dim flashx3, matdim
	If not Flasherflash9.TimerEnabled Then
		Flasherflash9.TimerEnabled = True
		Flasherflash9.visible = 1
		Flasherlit9.visible = 1
	End If
	flashx3 = FlashLevel9 * FlashLevel9 * FlashLevel9
	Flasherflash9.opacity = 1500 * flashx3
	Flasherlit9.BlendDisableLighting = 10 * flashx3
	Flasherbase9.BlendDisableLighting =  flashx3
	Flasherlight9.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel9)
	Flasherlit9.material = "domelit" & matdim
	FlashLevel9 = FlashLevel9 * 0.9 - 0.01
	If FlashLevel9 < 0.15 Then
		Flasherlit9.visible = 0
	Else
		Flasherlit9.visible = 1
	end If
	If FlashLevel9 < 0 Then
		Flasherflash9.TimerEnabled = False
		Flasherflash9.visible = 0
	End If
End Sub


Sub FlasherFlash10_Timer()
	dim flashx3, matdim
	If not Flasherflash10.TimerEnabled Then
		Flasherflash10.TimerEnabled = True
		Flasherflash10.visible = 1
		Flasherlit10.visible = 1
	End If
	flashx3 = FlashLevel10 * FlashLevel10 * FlashLevel10
	Flasherflash10.opacity = 1500 * flashx3
	Flasherlit10.BlendDisableLighting = 10 * flashx3
	Flasherbase10.BlendDisableLighting =  flashx3
	Flasherlight10.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel10)
	Flasherlit10.material = "domelit" & matdim
	FlashLevel10 = FlashLevel10 * 0.9 - 0.01
	If FlashLevel10 < 0.15 Then
		Flasherlit10.visible = 0
	Else
		Flasherlit10.visible = 1
	end If
	If FlashLevel10 < 0 Then
		Flasherflash10.TimerEnabled = False
		Flasherflash10.visible = 0
	End If
End Sub


Sub FlasherFlash11_Timer()
	dim flashx3, matdim
	If not Flasherflash11.TimerEnabled Then
		Flasherflash11.TimerEnabled = True
		Flasherflash11.visible = 1
		Flasherlit11.visible = 1
	End If
	flashx3 = FlashLevel11 * FlashLevel11 * FlashLevel11
	Flasherflash11.opacity = 1500 * flashx3
	Flasherlit11.BlendDisableLighting = 10 * flashx3
	Flasherbase11.BlendDisableLighting =  flashx3
	Flasherlight11.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel11)
	Flasherlit11.material = "domelit" & matdim
	FlashLevel11 = FlashLevel11 * 0.9 - 0.01
	If FlashLevel11 < 0.15 Then
		Flasherlit11.visible = 0
	Else
		Flasherlit11.visible = 1
	end If
	If FlashLevel11 < 0 Then
		Flasherflash11.TimerEnabled = False
		Flasherflash11.visible = 0
	End If
End Sub


Sub FlasherFlash12_Timer()
	dim flashx3, matdim
	If not Flasherflash12.TimerEnabled Then
		Flasherflash12.TimerEnabled = True
		Flasherflash12.visible = 1
		Flasherlit12.visible = 1
	End If
	flashx3 = FlashLevel12 * FlashLevel12 * FlashLevel12
	Flasherflash12.opacity = 8000 * flashx3
	Flasherlit12.BlendDisableLighting = 10 * flashx3
	Flasherbase12.BlendDisableLighting =  flashx3
	Flasherlight12.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel12)
	Flasherlit12.material = "domelit" & matdim
	FlashLevel12 = FlashLevel12 * 0.85 - 0.01
	If FlashLevel12 < 0.15 Then
		Flasherlit12.visible = 0
	Else
		Flasherlit12.visible = 1
	end If
	If FlashLevel12 < 0 Then
		Flasherflash12.TimerEnabled = False
		Flasherflash12.visible = 0
	End If
End Sub


Sub FlasherFlash13_Timer()
	dim flashx3, matdim
	If not Flasherflash13.TimerEnabled Then
		Flasherflash13.TimerEnabled = True
		Flasherflash13.visible = 1
		Flasherlit13.visible = 1
	End If
	flashx3 = FlashLevel13 * FlashLevel13 * FlashLevel13
	Flasherflash13.opacity = 8000 * flashx3
	Flasherlit13.BlendDisableLighting = 10 * flashx3
	Flasherbase13.BlendDisableLighting =  flashx3
	Flasherlight13.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel13)
	Flasherlit13.material = "domelit" & matdim
	FlashLevel13 = FlashLevel13 * 0.85 - 0.01
	If FlashLevel13 < 0.15 Then
		Flasherlit13.visible = 0
	Else
		Flasherlit13.visible = 1
	end If
	If FlashLevel13 < 0 Then
		Flasherflash13.TimerEnabled = False
		Flasherflash13.visible = 0
	End If
End Sub


Sub FlasherFlash14_Timer()
	dim flashx3, matdim
	If not Flasherflash14.TimerEnabled Then
		Flasherflash14.TimerEnabled = True
		Flasherflash14.visible = 1
		Flasherlit14.visible = 1
	End If
	flashx3 = FlashLevel14 * FlashLevel14 * FlashLevel14
	Flasherflash14.opacity = 8000 * flashx3
	Flasherlit14.BlendDisableLighting = 10 * flashx3
	Flasherbase14.BlendDisableLighting =  flashx3
	Flasherlight14.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel14)
	Flasherlit14.material = "domelit" & matdim
	FlashLevel14 = FlashLevel14 * 0.85 - 0.01
	If FlashLevel14 < 0.15 Then
		Flasherlit14.visible = 0
	Else
		Flasherlit14.visible = 1
	end If
	If FlashLevel14 < 0 Then
		Flasherflash14.TimerEnabled = False
		Flasherflash14.visible = 0
	End If
End Sub


Sub FlasherFlash15_Timer()
	dim flashx3, matdim
	If not Flasherflash15.TimerEnabled Then
		Flasherflash15.TimerEnabled = True
		Flasherflash15.visible = 1
		Flasherlit15.visible = 1
	End If
	flashx3 = FlashLevel15 * FlashLevel15 * FlashLevel15
	Flasherflash15.opacity = 8000 * flashx3
	Flasherlit15.BlendDisableLighting = 10 * flashx3
	Flasherbase15.BlendDisableLighting =  flashx3
	Flasherlight15.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel15)
	Flasherlit15.material = "domelit" & matdim
	FlashLevel15 = FlashLevel15 * 0.85 - 0.01
	If FlashLevel15 < 0.15 Then
		Flasherlit15.visible = 0
	Else
		Flasherlit15.visible = 1
	end If
	If FlashLevel15 < 0 Then
		Flasherflash15.TimerEnabled = False
		Flasherflash15.visible = 0
	End If
End Sub


Sub FlasherFlash16_Timer()
	dim flashx3, matdim
	If not Flasherflash16.TimerEnabled Then
		Flasherflash16.TimerEnabled = True
		Flasherflash16.visible = 1
		Flasherlit16.visible = 1
	End If
	flashx3 = FlashLevel16 * FlashLevel16 * FlashLevel16
	Flasherflash16.opacity = 8000 * flashx3
	Flasherlit16.BlendDisableLighting = 10 * flashx3
	Flasherbase16.BlendDisableLighting =  flashx3
	Flasherlight16.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel16)
	Flasherlit16.material = "domelit" & matdim
	FlashLevel16 = FlashLevel16 * 0.85 - 0.01
	If FlashLevel16 < 0.15 Then
		Flasherlit16.visible = 0
	Else
		Flasherlit16.visible = 1
	end If
	If FlashLevel16 < 0 Then
		Flasherflash16.TimerEnabled = False
		Flasherflash16.visible = 0
	End If
End Sub


Sub FlasherFlash17_Timer()
	dim flashx3, matdim
	If not Flasherflash17.TimerEnabled Then
		Flasherflash17.TimerEnabled = True
		Flasherflash17.visible = 1
		Flasherlit17.visible = 1
	End If
	flashx3 = FlashLevel17 * FlashLevel17 * FlashLevel17
	Flasherflash17.opacity = 8000 * flashx3
	Flasherlit17.BlendDisableLighting = 10 * flashx3
	Flasherbase17.BlendDisableLighting =  flashx3
	Flasherlight17.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel17)
	Flasherlit17.material = "domelit" & matdim
	FlashLevel17 = FlashLevel17 * 0.85 - 0.01
	If FlashLevel17 < 0.15 Then
		Flasherlit17.visible = 0
	Else
		Flasherlit17.visible = 1
	end If
	If FlashLevel17 < 0 Then
		Flasherflash17.TimerEnabled = False
		Flasherflash17.visible = 0
	End If
End Sub


Sub FlasherFlash18_Timer()
	dim flashx3, matdim
	If not Flasherflash18.TimerEnabled Then
		Flasherflash18.TimerEnabled = True
		Flasherflash18.visible = 1
		Flasherlit18.visible = 1
	End If
	flashx3 = FlashLevel18 * FlashLevel18 * FlashLevel18
	Flasherflash18.opacity = 8000 * flashx3
	Flasherlit18.BlendDisableLighting = 10 * flashx3
	Flasherbase18.BlendDisableLighting =  flashx3
	Flasherlight18.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel18)
	Flasherlit18.material = "domelit" & matdim
	FlashLevel18 = FlashLevel18 * 0.85 - 0.01
	If FlashLevel18 < 0.15 Then
		Flasherlit18.visible = 0
	Else
		Flasherlit18.visible = 1
	end If
	If FlashLevel18 < 0 Then
		Flasherflash18.TimerEnabled = False
		Flasherflash18.visible = 0
	End If
End Sub


Sub FlasherFlash19_Timer()
	dim flashx3, matdim
	If not Flasherflash19.TimerEnabled Then
		Flasherflash19.TimerEnabled = True
		Flasherflash19.visible = 1
		Flasherlit19.visible = 1
	End If
	flashx3 = FlashLevel19 * FlashLevel19 * FlashLevel19
	Flasherflash19.opacity = 8000 * flashx3
	Flasherlit19.BlendDisableLighting = 10 * flashx3
	Flasherbase19.BlendDisableLighting =  flashx3
	Flasherlight19.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel19)
	Flasherlit19.material = "domelit" & matdim
	FlashLevel19 = FlashLevel19 * 0.85 - 0.01
	If FlashLevel19 < 0.15 Then
		Flasherlit19.visible = 0
	Else
		Flasherlit19.visible = 1
	end If
	If FlashLevel19 < 0 Then
		Flasherflash19.TimerEnabled = False
		Flasherflash19.visible = 0
	End If
End Sub


Sub FlasherFlash20_Timer()
	dim flashx3, matdim
	If not Flasherflash20.TimerEnabled Then
		Flasherflash20.TimerEnabled = True
		Flasherflash20.visible = 1
		Flasherlit20.visible = 1
	End If
	flashx3 = FlashLevel20 * FlashLevel20 * FlashLevel20
	Flasherflash20.opacity = 1500 * flashx3
	Flasherlit20.BlendDisableLighting = 10 * flashx3
	Flasherbase20.BlendDisableLighting =  flashx3
	Flasherlight20.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel20)
	Flasherlit20.material = "domelit" & matdim
	FlashLevel20 = FlashLevel20 * 0.9 - 0.01
	If FlashLevel20 < 0.15 Then
		Flasherlit20.visible = 0
	Else
		Flasherlit20.visible = 1
	end If
	If FlashLevel20 < 0 Then
		Flasherflash20.TimerEnabled = False
		Flasherflash20.visible = 0
	End If
End Sub


Sub FlasherFlash21_Timer()
	dim flashx3, matdim
	If not Flasherflash21.TimerEnabled Then
		Flasherflash21.TimerEnabled = True
		Flasherflash21.visible = 1
		Flasherlit21.visible = 1
	End If
	flashx3 = FlashLevel21 * FlashLevel21 * FlashLevel21
	Flasherflash21.opacity = 1500 * flashx3
	Flasherlit21.BlendDisableLighting = 10 * flashx3
	Flasherbase21.BlendDisableLighting =  flashx3
	Flasherlight21.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel21)
	Flasherlit21.material = "domelit" & matdim
	FlashLevel21 = FlashLevel21 * 0.9 - 0.01
	If FlashLevel21 < 0.15 Then
		Flasherlit21.visible = 0
	Else
		Flasherlit21.visible = 1
	end If
	If FlashLevel21 < 0 Then
		Flasherflash21.TimerEnabled = False
		Flasherflash21.visible = 0
	End If
End Sub


Sub FlasherFlash22_Timer()
	dim flashx3, matdim
	If not Flasherflash22.TimerEnabled Then
		Flasherflash22.TimerEnabled = True
		Flasherflash22.visible = 1
		Flasherlit22.visible = 1
	End If
	flashx3 = FlashLevel22 * FlashLevel22 * FlashLevel22
	Flasherflash22.opacity = 8000 * flashx3
	Flasherlit22.BlendDisableLighting = 10 * flashx3
	Flasherbase22.BlendDisableLighting =  flashx3
	Flasherlight22.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel22)
	Flasherlit22.material = "domelit" & matdim
	FlashLevel22 = FlashLevel22 * 0.85 - 0.01
	If FlashLevel22 < 0.15 Then
		Flasherlit22.visible = 0
	Else
		Flasherlit22.visible = 1
	end If
	If FlashLevel22 < 0 Then
		Flasherflash22.TimerEnabled = False
		Flasherflash22.visible = 0
	End If
End Sub


Sub FlasherFlash23_Timer()
	dim flashx3, matdim
	If not Flasherflash23.TimerEnabled Then
		Flasherflash23.TimerEnabled = True
		Flasherflash23.visible = 1
		Flasherlit23.visible = 1
	End If
	flashx3 = FlashLevel23 * FlashLevel23 * FlashLevel23
	Flasherflash23.opacity = 1500 * flashx3
	Flasherlit23.BlendDisableLighting = 10 * flashx3
	Flasherbase23.BlendDisableLighting =  flashx3
	Flasherlight23.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel23)
	Flasherlit23.material = "domelit" & matdim
	FlashLevel23 = FlashLevel23 * 0.9 - 0.01
	If FlashLevel23 < 0.15 Then
		Flasherlit23.visible = 0
	Else
		Flasherlit23.visible = 1
	end If
	If FlashLevel23 < 0 Then
		Flasherflash23.TimerEnabled = False
		Flasherflash23.visible = 0
	End If
End Sub


Sub FlasherFlash24_Timer()
	dim flashx3, matdim
	If not Flasherflash24.TimerEnabled Then
		Flasherflash24.TimerEnabled = True
		Flasherflash24.visible = 1
		Flasherlit24.visible = 1
	End If
	flashx3 = FlashLevel24 * FlashLevel24 * FlashLevel24
	Flasherflash24.opacity = 8000 * flashx3
	Flasherlit24.BlendDisableLighting = 10 * flashx3
	Flasherbase24.BlendDisableLighting =  flashx3
	Flasherlight24.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel24)
	Flasherlit24.material = "domelit" & matdim
	FlashLevel24 = FlashLevel24 * 0.85 - 0.01
	If FlashLevel24 < 0.15 Then
		Flasherlit24.visible = 0
	Else
		Flasherlit24.visible = 1
	end If
	If FlashLevel24 < 0 Then
		Flasherflash24.TimerEnabled = False
		Flasherflash24.visible = 0
	End If
End Sub


Sub FlasherFlash25_Timer()
	dim flashx3, matdim
	If not Flasherflash25.TimerEnabled Then
		Flasherflash25.TimerEnabled = True
		Flasherflash25.visible = 1
		Flasherlit25.visible = 1
	End If
	flashx3 = FlashLevel25 * FlashLevel25 * FlashLevel25
	Flasherflash25.opacity = 1500 * flashx3
	Flasherlit25.BlendDisableLighting = 10 * flashx3
	Flasherbase25.BlendDisableLighting =  flashx3
	Flasherlight25.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel25)
	Flasherlit25.material = "domelit" & matdim
	FlashLevel25 = FlashLevel25 * 0.9 - 0.01
	If FlashLevel25 < 0.15 Then
		Flasherlit25.visible = 0
	Else
		Flasherlit25.visible = 1
	end If
	If FlashLevel25 < 0 Then
		Flasherflash25.TimerEnabled = False
		Flasherflash25.visible = 0
	End If
End Sub


Sub FlasherFlash26_Timer()
	dim flashx3, matdim
	If not Flasherflash26.TimerEnabled Then
		Flasherflash26.TimerEnabled = True
		Flasherflash26.visible = 1
		Flasherlit26.visible = 1
	End If
	flashx3 = FlashLevel26 * FlashLevel26 * FlashLevel26
	Flasherflash26.opacity = 8000 * flashx3
	Flasherlit26.BlendDisableLighting = 10 * flashx3
	Flasherbase26.BlendDisableLighting =  flashx3
	Flasherlight26.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel26)
	Flasherlit26.material = "domelit" & matdim
	FlashLevel26 = FlashLevel26 * 0.85 - 0.01
	If FlashLevel26 < 0.15 Then
		Flasherlit26.visible = 0
	Else
		Flasherlit26.visible = 1
	end If
	If FlashLevel26 < 0 Then
		Flasherflash26.TimerEnabled = False
		Flasherflash26.visible = 0
	End If
End Sub


Sub FlasherFlash27_Timer()
	dim flashx3, matdim
	If not Flasherflash27.TimerEnabled Then
		Flasherflash27.TimerEnabled = True
		Flasherflash27.visible = 1
		Flasherlit27.visible = 1
	End If
	flashx3 = FlashLevel27 * FlashLevel27 * FlashLevel27
	Flasherflash27.opacity = 1500 * flashx3
	Flasherlit27.BlendDisableLighting = 10 * flashx3
	Flasherbase27.BlendDisableLighting =  flashx3
	Flasherlight27.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel27)
	Flasherlit27.material = "domelit" & matdim
	FlashLevel27 = FlashLevel27 * 0.9 - 0.01
	If FlashLevel27 < 0.15 Then
		Flasherlit27.visible = 0
	Else
		Flasherlit27.visible = 1
	end If
	If FlashLevel27 < 0 Then
		Flasherflash27.TimerEnabled = False
		Flasherflash27.visible = 0
	End If
End Sub


Sub FlasherFlash28_Timer()
	dim flashx3, matdim
	If not Flasherflash28.TimerEnabled Then
		Flasherflash28.TimerEnabled = True
		Flasherflash28.visible = 0
		Flasherlit28.visible = 1
	End If
	flashx3 = FlashLevel28 * FlashLevel28 * FlashLevel28
	Flasherflash28.opacity = 8000 * flashx3
	Flasherlit28.BlendDisableLighting = 10 * flashx3
	Flasherbase28.BlendDisableLighting =  flashx3
	Flasherlight28.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel28)
	Flasherlit28.material = "domelit" & matdim
	FlashLevel28 = FlashLevel28 * 0.85 - 0.01
	If FlashLevel28 < 0.15 Then
		Flasherlit28.visible = 0
	Else
		Flasherlit28.visible = 1
	end If
	If FlashLevel28 < 0 Then
		Flasherflash28.TimerEnabled = False
		Flasherflash28.visible = 0
	End If
End Sub

Sub FlasherFlash29_Timer()
	dim flashx3, matdim
	If not Flasherflash29.TimerEnabled Then
		Flasherflash29.TimerEnabled = True
		Flasherflash29.visible = 0
		Flasherlit29.visible = 1
	End If
	flashx3 = FlashLevel29 * FlashLevel29 * FlashLevel29
	Flasherflash29.opacity = 8000 * flashx3
	Flasherlit29.BlendDisableLighting = 10 * flashx3
	Flasherbase29.BlendDisableLighting =  flashx3
	Flasherlight29.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel29)
	Flasherlit29.material = "domelit" & matdim
	FlashLevel29 = FlashLevel29 * 0.85 - 0.01
	If FlashLevel29 < 0.15 Then
		Flasherlit29.visible = 0
	Else
		Flasherlit29.visible = 1
	end If
	If FlashLevel29 < 0 Then
		Flasherflash29.TimerEnabled = False
		Flasherflash29.visible = 0
	End If
End Sub

Sub FlasherFlash30_Timer()
	dim flashx3, matdim
	If not Flasherflash30.TimerEnabled Then
		Flasherflash30.TimerEnabled = True
		Flasherflash30.visible = 0
		Flasherlit30.visible = 1
	End If
	flashx3 = FlashLevel30 * FlashLevel30 * FlashLevel30
	Flasherflash30.opacity = 8000 * flashx3
	Flasherlit30.BlendDisableLighting = 10 * flashx3
	Flasherbase30.BlendDisableLighting =  flashx3
	Flasherlight30.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel29)
	Flasherlit30.material = "domelit" & matdim
	FlashLevel30 = FlashLevel30 * 0.85 - 0.01
	If FlashLevel30 < 0.15 Then
		Flasherlit30.visible = 0
	Else
		Flasherlit30.visible = 1
	end If
	If FlashLevel30 < 0 Then
		Flasherflash30.TimerEnabled = False
		Flasherflash30.visible = 0
	End If
End Sub

'**********************************************************************************************************

Dim TotalBump, BumpNow, TotalBumpb, Bumpnowb

TotalBump = 8

Sub Change4Bump
	BumpNow = BumpNow + 1
	If BumpNow > 1 Then
		Primitive6.image = "bumper-cap-t2-"&BumpNow
		Primitive8.image = "bumper-cap-t2-"&BumpNow
		Primitive9.image = "bumper-cap-t2-"&BumpNow
		Primitive40.image = "bumper-cap-t2-"&BumpNow
	Else
		Primitive6.image = "bumper-cap-t2"
		Primitive8.image = "bumper-cap-t2"
		Primitive9.image = "bumper-cap-t2"
		Primitive40.image = "bumper-cap-t2"
	End If
	If BumpNow = TotalBump Then BumpNow = 0
End Sub

TotalBumpb = 4

Sub Change2Bump
	BumpNowb = BumpNowb + 1
	If BumpNowb > 1 Then
		Primitive10.image = "bumper-cap-t4-"&BumpNowb
		Primitive46.image = "bumper-cap-t4-"&BumpNowb
	Else
		Primitive10.image = "bumper-cap-t4"
		Primitive46.image = "bumper-cap-t4"
	End If
	If BumpNowb = TotalBumpb Then BumpNowb = 0
End Sub

'**********************************************************************************************************



'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1) = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(2) = "dtCDrop.SolDropUp"
SolCallback(3) = "bsSaucer.SolOut"
SolCallback(4) = "bsSaucer2.SolOut"
SolCallback(5) = "dtTDrop.SolDropUp"
SolCallback(7) = "bsTrough.SolOut"
SolCallback(19)= "vpmNudge.SolGameOn"

	SolCallback(sLRFlipper) = "SolRFlipper"
	SolCallback(sLLFlipper) = "SolLFlipper"

	Sub SolLFlipper(Enabled)
		 If Enabled Then
			 PlaySoundAtVol SoundFx("FlipperUp",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
		 Else
			 PlaySoundAtVol SoundFx("FlipperDown",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
		 End If
	  End Sub

	Sub SolRFlipper(Enabled)
		 If Enabled Then
			 PlaySoundAtVol SoundFx("FlipperUp",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
		 Else
			 PlaySoundAtVol SoundFx("FlipperDown",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
		 End If
	End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next
 '**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough,bsSaucer,bsSaucer2,dtCDrop,dtTDrop
Dim BIP
Dim AAA
Dim BBB
Dim CCC

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
         .SplashInfoLine = "Star Wars Empire Strikes Back (Hankin 1980) v2.2.1" & vbNewLine & "VPX Table By 32assassin" & vbNewLine & "Flasher Pic Sound Mod By Xenonph"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 1
        Controller.Games("empsback").Settings.Value("sound")=0
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
   		Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
		Controller.Run
	If Err Then MsgBox Err.Description
	On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1:PlayMusic"0ESB01.mp3":sw39.timerinterval=100000:sw39.timerenabled=1
	vpmNudge.TiltSwitch=2
	vpmNudge.Sensitivity=1
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,Bumper4,Bumper5,Bumper6, LeftSlingshot, RightSlingshot)

	Set bsTrough=New cvpmBallStack
	bsTrough.InitSw 0,40,0,0,0,0,0,0
	bsTrough.InitKick ballrelease,90,8
	bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("solenoid",DOFContactors)
	bsTrough.Balls=1

	Set bsSaucer=New cvpmBallStack
	bsSaucer.InitSaucer LeftHole,32,157,8
	bsSaucer.KickForceVar=2.5
	bsSaucer.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("solenoid",DOFContactors)

	Set bsSaucer2=New cvpmBallStack
	bsSaucer2.InitSaucer RightHole,30,180,6
	bsSaucer.KickForceVar=2.5
	bsSaucer2.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("solenoid",DOFContactors)

  	Set dtCDrop=New cvpmDropTarget
	dtCDrop.InitDrop Array(sw25,sw26,sw27),Array(25,26,27)
	dtCDrop.InitSnd SoundFX("DTDrop",DOFContactors), SoundFX("DTReset",DOFContactors)

  	Set dtTDrop=New cvpmDropTarget
	dtTDrop.InitDrop sw28,28
	dtTDrop.InitSnd SoundFX("DTDrop",DOFContactors), SoundFX("DTReset",DOFContactors)

End Sub
 '**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
 Sub Table1_KeyDown(ByVal keycode)
	If KeyCode=LeftMagnaSave Then Change2Bump
	If KeyCode=RightMagnaSave Then Change4Bump
    If keycode=rightflipperkey then controller.switch(1)=0
	If KeyCode= 49 Then Starpic=0:Starstop
	If KeyCode= 50 Then Starpic=1
	If vpmKeyDown(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull",plunger,1
    If keycode=AddCreditKey then vpmTimer.pulseSW (swCoin1)
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If keycode=rightflipperkey then controller.switch(1)=1
    If KeyCode = 6 and Starpic=0 Then
            FlashLevel1 = 1 : FlasherFlash1_Timer:FlashLevel4 = 1 : FlasherFlash4_Timer:FlashLevel20 = 1 : FlasherFlash20_Timer:FlashLevel25 = 1 : FlasherFlash25_Timer:FlashLevel19 = 1 : FlasherFlash19_Timer:FlashLevel26 = 1 : FlasherFlash26_Timer
            FlashLevel27 = 1 : FlasherFlash27_Timer
            sw17.timerinterval=200:sw17.timerenabled=1:sw11.timerenabled=0:Starstop
			Dim x
			x = INT(15 * RND(1) )
			Select Case x
			Case 0:PlaySound"coin":StopESBSounds:PlaySound("0ESBC01"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 1:PlaySound"coin":StopESBSounds:PlaySound("0ESBC02"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 2:PlaySound"coin":StopESBSounds:PlaySound("0ESBC03"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 3:PlaySound"coin":StopESBSounds:PlaySound("0ESBC04"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 4:PlaySound"coin":StopESBSounds:PlaySound("0ESBC05"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 5:PlaySound"coin":StopESBSounds:PlaySound("0ESBC06"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 6:PlaySound"coin":StopESBSounds:PlaySound("0ESBC07"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 7:PlaySound"coin":StopESBSounds:PlaySound("0ESBC08"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 8:PlaySound"coin":StopESBSounds:PlaySound("0ESBC09"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 9:PlaySound"coin":StopESBSounds:PlaySound("0ESB00zi"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 10:PlaySound"coin":StopESBSounds:PlaySound("0ESB00zj"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 11:PlaySound"coin":StopESBSounds:PlaySound("0ESBC10"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 12:PlaySound"coin":StopESBSounds:PlaySound("0ESBC11"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 13:PlaySound"coin":StopESBSounds:PlaySound("0ESBC12"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 14:PlaySound"coin":StopESBSounds:PlaySound("0ESBC13"):sw29.timerinterval=3000:sw29.timerenabled=1
			End Select
            End If

    If KeyCode = 6 and Starpic=1 Then
            FlashLevel1 = 1 : FlasherFlash1_Timer:FlashLevel4 = 1 : FlasherFlash4_Timer:FlashLevel20 = 1 : FlasherFlash20_Timer:FlashLevel25 = 1 : FlasherFlash25_Timer:FlashLevel19 = 1 : FlasherFlash19_Timer:FlashLevel26 = 1 : FlasherFlash26_Timer
            FlashLevel27 = 1 : FlasherFlash27_Timer
            sw17.timerinterval=200:sw17.timerenabled=1:sw11.timerenabled=0:Starstop
			Dim s
			s = INT(15 * RND(1) )
			Select Case s
			Case 0:PlaySound"coin":StopESBSounds:PlaySound("0ESBC01"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash38.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 1:PlaySound"coin":StopESBSounds:PlaySound("0ESBC02"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash41.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 2:PlaySound"coin":StopESBSounds:PlaySound("0ESBC03"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash34.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 3:PlaySound"coin":StopESBSounds:PlaySound("0ESBC04"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 4:PlaySound"coin":StopESBSounds:PlaySound("0ESBC05"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash28.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 5:PlaySound"coin":StopESBSounds:PlaySound("0ESBC06"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 6:PlaySound"coin":StopESBSounds:PlaySound("0ESBC07"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 7:PlaySound"coin":StopESBSounds:PlaySound("0ESBC08"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 8:PlaySound"coin":StopESBSounds:PlaySound("0ESBC09"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 9:PlaySound"coin":StopESBSounds:PlaySound("0ESB00zi"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash28.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 10:PlaySound"coin":StopESBSounds:PlaySound("0ESB00zj"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash33.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 11:PlaySound"coin":StopESBSounds:PlaySound("0ESBC10"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash38.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 12:PlaySound"coin":StopESBSounds:PlaySound("0ESBC11"):sw29.timerinterval=3000:sw29.timerenabled=1
			Case 13:PlaySound"coin":StopESBSounds:PlaySound("0ESBC12"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash35.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 14:PlaySound"coin":StopESBSounds:PlaySound("0ESBC13"):sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash30.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			End Select
            End If


	If KeyCode = 4 and Starpic=0 Then
            FlashLevel1 = 1 : FlasherFlash1_Timer:FlashLevel4 = 1 : FlasherFlash4_Timer:FlashLevel20 = 1 : FlasherFlash20_Timer:FlashLevel25 = 1 : FlasherFlash25_Timer:FlashLevel19 = 1 : FlasherFlash19_Timer:FlashLevel26 = 1 : FlasherFlash26_Timer
            FlashLevel27 = 1 : FlasherFlash27_Timer
            sw17.timerinterval=200:sw17.timerenabled=1:sw11.timerenabled=0:Starstop
			Dim y
			y = INT(15 * RND(1) )
			Select Case y
			Case 0:StopESBSounds:PlaySound"0ESBC01":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 1:StopESBSounds:PlaySound"0ESBC02":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 2:StopESBSounds:PlaySound"0ESBC03":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 3:StopESBSounds:PlaySound"0ESBC04":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 4:StopESBSounds:PlaySound"0ESBC05":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 5:StopESBSounds:PlaySound"0ESBC06":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 6:StopESBSounds:PlaySound"0ESBC07":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 7:StopESBSounds:PlaySound"0ESBC08":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 8:StopESBSounds:PlaySound"0ESBC09":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 9:StopESBSounds:PlaySound"0ESB00zi":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 10:StopESBSounds:PlaySound"0ESB00zj":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 11:StopESBSounds:PlaySound"0ESBC10":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 12:StopESBSounds:PlaySound"0ESBC11":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 13:StopESBSounds:PlaySound"0ESBC12":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 14:StopESBSounds:PlaySound"0ESBC13":sw29.timerinterval=3000:sw29.timerenabled=1
			End Select
			end if

	If KeyCode = 4 and Starpic=1 Then
            FlashLevel1 = 1 : FlasherFlash1_Timer:FlashLevel4 = 1 : FlasherFlash4_Timer:FlashLevel20 = 1 : FlasherFlash20_Timer:FlashLevel25 = 1 : FlasherFlash25_Timer:FlashLevel19 = 1 : FlasherFlash19_Timer:FlashLevel26 = 1 : FlasherFlash26_Timer
            FlashLevel27 = 1 : FlasherFlash27_Timer
            sw17.timerinterval=200:sw17.timerenabled=1:sw11.timerenabled=0:Starstop
			Dim r
			r = INT(15 * RND(1) )
			Select Case r
			Case 0:StopESBSounds:PlaySound"0ESBC01":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash38.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 1:StopESBSounds:PlaySound"0ESBC02":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash41.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 2:StopESBSounds:PlaySound"0ESBC03":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash34.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 3:StopESBSounds:PlaySound"0ESBC04":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 4:StopESBSounds:PlaySound"0ESBC05":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash28.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 5:StopESBSounds:PlaySound"0ESBC06":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 6:StopESBSounds:PlaySound"0ESBC07":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 7:StopESBSounds:PlaySound"0ESBC08":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 8:StopESBSounds:PlaySound"0ESBC09":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 9:StopESBSounds:PlaySound"0ESB00zi":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash28.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 10:StopESBSounds:PlaySound"0ESB00zj":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash33.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 11:StopESBSounds:PlaySound"0ESBC10":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash38.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 12:StopESBSounds:PlaySound"0ESBC11":sw29.timerinterval=3000:sw29.timerenabled=1
			Case 13:StopESBSounds:PlaySound"0ESBC12":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash35.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 14:StopESBSounds:PlaySound"0ESBC13":sw29.timerinterval=3000:sw29.timerenabled=1:Flasherflash30.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			End Select
			end if

	If KeyCode = 2 and Starpic=0 Then
		EndMusic:sw39.timerenabled=0:Drain.timerenabled=0:sw12.timerinterval=200:sw12.timerenabled=1:sw11.timerenabled=0:Starstop
        Dim v
        v = INT(38 * RND(1) )
        Select Case v
        Case 0:StopESBSounds:Playsound("0ESB00a"):Drain.timerinterval=6000:Drain.timerenabled=1
        Case 1:StopESBSounds:Playsound("0ESB00b"):Drain.timerinterval=3000:Drain.timerenabled=1
        Case 2:StopESBSounds:Playsound("0ESB00c"):Drain.timerinterval=3000:Drain.timerenabled=1
        Case 3:StopESBSounds:Playsound("0ESB00d"):Drain.timerinterval=2000:Drain.timerenabled=1
        Case 4:StopESBSounds:Playsound("0ESB00e"):Drain.timerinterval=2000:Drain.timerenabled=1
        Case 5:StopESBSounds:Playsound("0ESB00f"):Drain.timerinterval=4000:Drain.timerenabled=1
        Case 6:StopESBSounds:Playsound("0ESB00g"):Drain.timerinterval=4000:Drain.timerenabled=1
        Case 7:StopESBSounds:Playsound("0ESB00h"):Drain.timerinterval=5000:Drain.timerenabled=1
        Case 8:StopESBSounds:Playsound("0ESB00i"):Drain.timerinterval=4000:Drain.timerenabled=1
        Case 9:StopESBSounds:Playsound("0ESB00j"):Drain.timerinterval=5000:Drain.timerenabled=1
        Case 10:StopESBSounds:Playsound("0ESB00k"):Drain.timerinterval=9000:Drain.timerenabled=1
        Case 11:StopESBSounds:Playsound("0ESB00l"):Drain.timerinterval=3000:Drain.timerenabled=1
        Case 12:StopESBSounds:Playsound("0ESB00m"):Drain.timerinterval=4000:Drain.timerenabled=1
        Case 13:StopESBSounds:Playsound("0ESB00n"):Drain.timerinterval=4000:Drain.timerenabled=1
        Case 14:StopESBSounds:Playsound("0ESB00o"):Drain.timerinterval=5000:Drain.timerenabled=1
        Case 15:StopESBSounds:Playsound("0ESB00p"):Drain.timerinterval=4000:Drain.timerenabled=1
        Case 16:StopESBSounds:Playsound("0ESB00za"):Drain.timerinterval=6500:Drain.timerenabled=1
        Case 17:StopESBSounds:Playsound("0ESB00zb"):Drain.timerinterval=5000:Drain.timerenabled=1
        Case 18:StopESBSounds:Playsound("0ESB00zc"):Drain.timerinterval=4000:Drain.timerenabled=1
        Case 19:StopESBSounds:Playsound("0ESB00zd"):Drain.timerinterval=4000:Drain.timerenabled=1
        Case 20:StopESBSounds:Playsound("0ESB00ze"):Drain.timerinterval=6000:Drain.timerenabled=1
        Case 21:StopESBSounds:Playsound("0ESB00zk"):Drain.timerinterval=3000:Drain.timerenabled=1
        Case 22:StopESBSounds:Playsound("0ESB00zl"):Drain.timerinterval=6000:Drain.timerenabled=1
        Case 23:StopESBSounds:Playsound("0ESB00zm"):Drain.timerinterval=3000:Drain.timerenabled=1
        Case 24:StopESBSounds:Playsound("0ESB00zn"):Drain.timerinterval=4000:Drain.timerenabled=1
        Case 25:StopESBSounds:Playsound("0ESB00za"):Drain.timerinterval=7000:Drain.timerenabled=1
        Case 26:StopESBSounds:Playsound("0ESB00zp"):Drain.timerinterval=7000:Drain.timerenabled=1
        Case 27:StopESBSounds:Playsound("0ESB00zq"):Drain.timerinterval=4000:Drain.timerenabled=1
        Case 28:StopESBSounds:Playsound("0ESB00zr"):Drain.timerinterval=5000:Drain.timerenabled=1
        Case 29:StopESBSounds:Playsound("0ESB00zs"):Drain.timerinterval=6000:Drain.timerenabled=1
        Case 30:StopESBSounds:Playsound("0ESB00zt"):Drain.timerinterval=5000:Drain.timerenabled=1
        Case 31:StopESBSounds:Playsound("0ESB00zu"):Drain.timerinterval=3000:Drain.timerenabled=1
        Case 32:StopESBSounds:Playsound("0ESB00zv"):Drain.timerinterval=6000:Drain.timerenabled=1
        Case 33:StopESBSounds:Playsound("0ESB00zw"):Drain.timerinterval=3000:Drain.timerenabled=1
        Case 34:StopESBSounds:Playsound("0ESB00zx"):Drain.timerinterval=7000:Drain.timerenabled=1
        Case 35:StopESBSounds:Playsound("0ESB00zy"):Drain.timerinterval=7000:Drain.timerenabled=1
        Case 36:StopESBSounds:Playsound("0ESB00zz"):Drain.timerinterval=6000:Drain.timerenabled=1
        Case 37:StopESBSounds:Playsound("0ESB00zza"):Drain.timerinterval=7000:Drain.timerenabled=1
        End Select
	End If

	If KeyCode = 2 and Starpic=1 Then
		EndMusic:sw39.timerenabled=0:Drain.timerenabled=0:sw12.timerinterval=200:sw12.timerenabled=1:sw11.timerenabled=0:Starstop
        Dim q
        q = INT(38 * RND(1) )
        Select Case q
        Case 0:StopESBSounds:Playsound("0ESB00a"):Drain.timerinterval=6000:Drain.timerenabled=1:Flasherflash37.visible = 1
        Case 1:StopESBSounds:Playsound("0ESB00b"):Drain.timerinterval=3000:Drain.timerenabled=1:Flasherflash35.visible = 1
        Case 2:StopESBSounds:Playsound("0ESB00c"):Drain.timerinterval=3000:Drain.timerenabled=1:Flasherflash28.visible = 1:Flasherflash43.visible = 1
        Case 3:StopESBSounds:Playsound("0ESB00d"):Drain.timerinterval=2000:Drain.timerenabled=1:Flasherflash36.visible = 1
        Case 4:StopESBSounds:Playsound("0ESB00e"):Drain.timerinterval=2000:Drain.timerenabled=1:Flasherflash41.visible = 1
        Case 5:StopESBSounds:Playsound("0ESB00f"):Drain.timerinterval=4000:Drain.timerenabled=1:Flasherflash36.visible = 1
        Case 6:StopESBSounds:Playsound("0ESB00g"):Drain.timerinterval=4000:Drain.timerenabled=1:Flasherflash31.visible = 1
        Case 7:StopESBSounds:Playsound("0ESB00h"):Drain.timerinterval=5000:Drain.timerenabled=1:Flasherflash36.visible = 1:Flasherflash40.visible = 1
        Case 8:StopESBSounds:Playsound("0ESB00i"):Drain.timerinterval=4000:Drain.timerenabled=1:Flasherflash28.visible = 1
        Case 9:StopESBSounds:Playsound("0ESB00j"):Drain.timerinterval=5000:Drain.timerenabled=1:Flasherflash28.visible = 1
        Case 10:StopESBSounds:Playsound("0ESB00k"):Drain.timerinterval=9000:Drain.timerenabled=1:Flasherflash30.visible = 1
        Case 11:StopESBSounds:Playsound("0ESB00l"):Drain.timerinterval=3000:Drain.timerenabled=1:Flasherflash30.visible = 1
        Case 12:StopESBSounds:Playsound("0ESB00m"):Drain.timerinterval=4000:Drain.timerenabled=1:Flasherflash30.visible = 1
        Case 13:StopESBSounds:Playsound("0ESB00n"):Drain.timerinterval=4000:Drain.timerenabled=1:Flasherflash30.visible = 1
        Case 14:StopESBSounds:Playsound("0ESB00o"):Drain.timerinterval=5000:Drain.timerenabled=1:Flasherflash38.visible = 1
        Case 15:StopESBSounds:Playsound("0ESB00p"):Drain.timerinterval=4000:Drain.timerenabled=1:Flasherflash28.visible = 1
        Case 16:StopESBSounds:Playsound("0ESB00za"):Drain.timerinterval=6500:Drain.timerenabled=1:Flasherflash31.visible = 1
        Case 17:StopESBSounds:Playsound("0ESB00zb"):Drain.timerinterval=5000:Drain.timerenabled=1:Flasherflash30.visible = 1
        Case 18:StopESBSounds:Playsound("0ESB00zc"):Drain.timerinterval=4000:Drain.timerenabled=1:Flasherflash31.visible = 1
        Case 19:StopESBSounds:Playsound("0ESB00zd"):Drain.timerinterval=4000:Drain.timerenabled=1:Flasherflash28.visible = 1
        Case 20:StopESBSounds:Playsound("0ESB00ze"):Drain.timerinterval=6000:Drain.timerenabled=1:Flasherflash28.visible = 1
        Case 21:StopESBSounds:Playsound("0ESB00zk"):Drain.timerinterval=3000:Drain.timerenabled=1:Flasherflash41.visible = 1
        Case 22:StopESBSounds:Playsound("0ESB00zl"):Drain.timerinterval=6000:Drain.timerenabled=1:Flasherflash31.visible = 1:Flasherflash39.visible = 1
        Case 23:StopESBSounds:Playsound("0ESB00zm"):Drain.timerinterval=3000:Drain.timerenabled=1
        Case 24:StopESBSounds:Playsound("0ESB00zn"):Drain.timerinterval=4000:Drain.timerenabled=1:Flasherflash39.visible = 1:Flasherflash41.visible = 1
        Case 25:StopESBSounds:Playsound("0ESB00za"):Drain.timerinterval=7000:Drain.timerenabled=1:Flasherflash31.visible = 1
        Case 26:StopESBSounds:Playsound("0ESB00zp"):Drain.timerinterval=7000:Drain.timerenabled=1:Flasherflash28.visible = 1
        Case 27:StopESBSounds:Playsound("0ESB00zq"):Drain.timerinterval=4000:Drain.timerenabled=1:Flasherflash41.visible = 1
        Case 28:StopESBSounds:Playsound("0ESB00zr"):Drain.timerinterval=5000:Drain.timerenabled=1:Flasherflash29.visible = 1
        Case 29:StopESBSounds:Playsound("0ESB00zs"):Drain.timerinterval=6000:Drain.timerenabled=1:Flasherflash30.visible = 1:Flasherflash42.visible = 1
        Case 30:StopESBSounds:Playsound("0ESB00zt"):Drain.timerinterval=5000:Drain.timerenabled=1:Flasherflash30.visible = 1:Flasherflash42.visible = 1
        Case 31:StopESBSounds:Playsound("0ESB00zu"):Drain.timerinterval=3000:Drain.timerenabled=1:Flasherflash40.visible = 1
        Case 32:StopESBSounds:Playsound("0ESB00zv"):Drain.timerinterval=6000:Drain.timerenabled=1:Flasherflash37.visible = 1
        Case 33:StopESBSounds:Playsound("0ESB00zw"):Drain.timerinterval=3000:Drain.timerenabled=1:Flasherflash39.visible = 1:Flasherflash41.visible = 1
        Case 34:StopESBSounds:Playsound("0ESB00zx"):Drain.timerinterval=7000:Drain.timerenabled=1:Flasherflash28.visible = 1:Flasherflash42.visible = 1
        Case 35:StopESBSounds:Playsound("0ESB00zy"):Drain.timerinterval=7000:Drain.timerenabled=1:Flasherflash31.visible = 1:Flasherflash39.visible = 1
        Case 36:StopESBSounds:Playsound("0ESB00zz"):Drain.timerinterval=6000:Drain.timerenabled=1:Flasherflash29.visible = 1
        Case 37:StopESBSounds:Playsound("0ESB00zza"):Drain.timerinterval=7000:Drain.timerenabled=1:Flasherflash32.visible = 1
        End Select
	End If

	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Fire:playsoundAtVol"plunger",plunger,1:sw12.timerenabled=0:Starstop
	If KeyCode=PlungerKey and BIP=1 and AAA=0 Then StopSound"0ESB00":PlaySound"0ESB00":Drain.timerenabled=0:sw39.timerenabled=0
	If KeyCode=PlungerKey and BIP=1 and AAA=0 Then
              Dim z
              z = INT(12 * RND(1) )
              Select Case z
              Case 0:PlayMusic"0ESB02.mp3":sw14.timerinterval=181000:sw14.timerenabled=1
              Case 1:PlayMusic"0ESB03.mp3":sw14.timerinterval=181000:sw14.timerenabled=1
              Case 2:PlayMusic"0ESB04.mp3":sw14.timerinterval=181000:sw14.timerenabled=1
              Case 3:PlayMusic"0ESB05.mp3":sw14.timerinterval=181000:sw14.timerenabled=1
              Case 4:PlayMusic"0ESB06.mp3":sw14.timerinterval=181000:sw14.timerenabled=1
              Case 5:PlayMusic"0ESB07.mp3":sw14.timerinterval=181000:sw14.timerenabled=1
              Case 6:PlayMusic"0ESB08.mp3":sw14.timerinterval=144000:sw14.timerenabled=1
              Case 7:PlayMusic"0ESB09.mp3":sw14.timerinterval=123000:sw14.timerenabled=1
              Case 8:PlayMusic"0ESB14.mp3":sw14.timerinterval=168000:sw14.timerenabled=1
              Case 9:PlayMusic"0ESB15.mp3":sw14.timerinterval=91000:sw14.timerenabled=1
              Case 10:PlayMusic"0ESB16.mp3":sw14.timerinterval=106000:sw14.timerenabled=1
              Case 11:PlayMusic"0ESB17.mp3":sw14.timerinterval=58000:sw14.timerenabled=1
              End Select
              End If
	If KeyCode=PlungerKey Then
              Dim t
              t = INT(3 * RND(1) )
              Select Case t
              Case 0:FlashLevel5 = 1 : FlasherFlash5_Timer
              Case 1:FlashLevel23 = 1 : FlasherFlash23_Timer
              Case 2:FlashLevel24 = 1 : FlasherFlash24_Timer
              End Select
              End If
	If KeyCode=PlungerKey and BIP=1 and AAA=1 Then playsoundAtVol"plunger",plunger,1:BBB=0
    If KeyCode=PlungerKey and BIP=0 and AAA=0 Then
			  playsoundAtVol"plungerreleasefree",plunger,1
              End If
End Sub
'**********************************************************************************************************


 ' Switches & Targets
'**********************************************************************************************************

Sub Starstop()

   Flasherflash28.visible = 0
   Flasherflash29.visible = 0
   Flasherflash30.visible = 0
   Flasherflash31.visible = 0
   Flasherflash32.visible = 0
   Flasherflash33.visible = 0
   Flasherflash34.visible = 0
   Flasherflash35.visible = 0
   Flasherflash36.visible = 0
   Flasherflash37.visible = 0
   Flasherflash38.visible = 0
   Flasherflash39.visible = 0
   Flasherflash40.visible = 0
   Flasherflash41.visible = 0
   Flasherflash42.visible = 0
   Flasherflash43.visible = 0
   Flasherflash44.visible = 0

End Sub

Sub StopESBSounds()

   StopSound"0ESB00a"
   StopSound"0ESB00b"
   StopSound"0ESB00c"
   StopSound"0ESB00d"
   StopSound"0ESB00e"
   StopSound"0ESB00f"
   StopSound"0ESB00g"
   StopSound"0ESB00h"
   StopSound"0ESB00i"
   StopSound"0ESB00j"
   StopSound"0ESB00k"
   StopSound"0ESB00l"
   StopSound"0ESB00m"
   StopSound"0ESB00n"
   StopSound"0ESB00o"
   StopSound"0ESB00p"
   StopSound"0ESB00q"
   StopSound"0ESB00r"
   StopSound"0ESB00s"
   StopSound"0ESB00t"
   StopSound"0ESB00u"
   StopSound"0ESB00v"
   StopSound"0ESB00w"
   StopSound"0ESB00x"
   StopSound"0ESB00y"
   StopSound"0ESB00z"
   StopSound"0ESB00za"
   StopSound"0ESB00zb"
   StopSound"0ESB00zc"
   StopSound"0ESB00zd"
   StopSound"0ESB00ze"
   StopSound"0ESB00zf"
   StopSound"0ESB00zg"
   StopSound"0ESB00zh"
   StopSound"0ESB00zi"
   StopSound"0ESB00zj"
   StopSound"0ESB00zk"
   StopSound"0ESB00zl"
   StopSound"0ESB00zm"
   StopSound"0ESB00zn"
   StopSound"0ESB00zo"
   StopSound"0ESB00zp"
   StopSound"0ESB00zq"
   StopSound"0ESB00zr"
   StopSound"0ESB00zs"
   StopSound"0ESB00zt"
   StopSound"0ESB00zu"
   StopSound"0ESB00zv"
   StopSound"0ESB00zw"
   StopSound"0ESB00zx"
   StopSound"0ESB00zy"
   StopSound"0ESB00zz"
   StopSound"0ESB00zza"
   StopSound"0ESBC01"
   StopSound"0ESBC02"
   StopSound"0ESBC03"
   StopSound"0ESBC04"
   StopSound"0ESBC05"
   StopSound"0ESBC06"
   StopSound"0ESBC07"
   StopSound"0ESBC08"
   StopSound"0ESBC09"
   StopSound"0ESBC10"
   StopSound"0ESBC11"
   StopSound"0ESBC12"
   StopSound"0ESBC13"

End Sub


'Kickers
Sub RightHole_Hit:bsSaucer2.AddBall 0 : playsoundAtVol "popper_ball",RightHole,1:FlashLevel26 = 1 : FlasherFlash26_Timer:FlashLevel19 = 1 : FlasherFlash19_Timer:sw6.timerinterval=400:sw6.timerenabled=1
    If Starpic=0 Then
			Dim x
			x = INT(14 * RND(1) )
			Select Case x
			Case 0:PlaySound"0ESBC01"
			Case 1:PlaySound"0ESBC02"
			Case 2:PlaySound"0ESBC03"
			Case 3:PlaySound"0ESBC04"
			Case 4:PlaySound"0ESB00l"
			Case 5:PlaySound"0ESBC06"
			Case 6:PlaySound"0ESBC07"
			Case 7:PlaySound"0ESBC08"
			Case 8:PlaySound"0ESBC09"
			Case 9:PlaySound"0ESB00zf"
			Case 10:PlaySound"0ESB00zg"
			Case 11:PlaySound"0ESB00zh"
			Case 12:PlaySound"0ESB00zi"
			Case 13:PlaySound"0ESB00zj"
			End Select
     Else
			Dim o
			o = INT(14 * RND(1) )
			Select Case o
			Case 0:PlaySound"0ESBC01":Flasherflash38.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 1:PlaySound"0ESBC02":Flasherflash41.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 2:PlaySound"0ESBC03":Flasherflash34.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 3:PlaySound"0ESBC04"
			Case 4:PlaySound"0ESB00l":Flasherflash30.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 5:PlaySound"0ESBC06":Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 6:PlaySound"0ESBC07":Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 7:PlaySound"0ESBC08":Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 8:PlaySound"0ESBC09":Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 9:PlaySound"0ESB00zf":Flasherflash30.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 10:PlaySound"0ESB00zg":Flasherflash30.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 11:PlaySound"0ESB00zh":Flasherflash28.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1:Flasherflash43.visible = 1
			Case 12:PlaySound"0ESB00zi":Flasherflash28.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 13:PlaySound"0ESB00zj":Flasherflash33.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			End Select
            End If
End Sub

Sub RightHole_unHit:sw6.timerenabled=0:End Sub

Sub LeftHole_Hit:bsSaucer.AddBall 0 : playsoundAtVol "popper_ball",Lefthole,1:FlashLevel1 = 1 : FlasherFlash1_Timer:FlashLevel4 = 1 : FlasherFlash4_Timer:sw5.timerinterval=400:sw5.timerenabled=1
      If Starpic=0 Then
			Dim x
			x = INT(14 * RND(1) )
			Select Case x
			Case 0:PlaySound"0ESBC01"
			Case 1:PlaySound"0ESBC02"
			Case 2:PlaySound"0ESBC03"
			Case 3:PlaySound"0ESBC04"
			Case 4:PlaySound"0ESB00l"
			Case 5:PlaySound"0ESBC06"
			Case 6:PlaySound"0ESBC07"
			Case 7:PlaySound"0ESBC08"
			Case 8:PlaySound"0ESBC09"
			Case 9:PlaySound"0ESB00zf"
			Case 10:PlaySound"0ESB00zg"
			Case 11:PlaySound"0ESB00zh"
			Case 12:PlaySound"0ESB00zi"
			Case 13:PlaySound"0ESB00zj"
			End Select
      Else
			Dim n
			n = INT(14 * RND(1) )
			Select Case n
			Case 0:PlaySound"0ESBC01":Flasherflash38.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 1:PlaySound"0ESBC02":Flasherflash41.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 2:PlaySound"0ESBC03":Flasherflash34.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 3:PlaySound"0ESBC04"
			Case 4:PlaySound"0ESB00l":Flasherflash30.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 5:PlaySound"0ESBC06":Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 6:PlaySound"0ESBC07":Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 7:PlaySound"0ESBC08":Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 8:PlaySound"0ESBC09":Flasherflash39.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 9:PlaySound"0ESB00zf":Flasherflash30.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 10:PlaySound"0ESB00zg":Flasherflash30.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 11:PlaySound"0ESB00zh":Flasherflash28.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1:Flasherflash43.visible = 1
			Case 12:PlaySound"0ESB00zi":Flasherflash28.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			Case 13:PlaySound"0ESB00zj":Flasherflash33.visible = 1:sw11.timerinterval=2000:sw11.timerenabled=1
			End Select
            End If
End Sub

Sub LeftHole_unHit:sw5.timerenabled=0:sw6.timerenabled=0:End Sub



Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain",drain,1 :BBB=0:AAA=0:BIP=0:FlashLevel21 = 1 : FlasherFlash21_Timer:Flasherlight30.IntensityScale = 0:Flasherlight31.IntensityScale = 0:Flasherlight32.IntensityScale = 0:Flasherlight33.IntensityScale = 0:Flasherlight34.IntensityScale = 0:Flasherlight35.IntensityScale = 0:Flasherlight36.IntensityScale = 0
      If Starpic=0 Then
			Dim x
			x = INT(22 * RND(1) )
			Select Case x
			Case 0:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD01":Drain.timerinterval=2000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 1:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD02":Drain.timerinterval=7000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 2:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD03":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 3:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD04":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 4:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD05":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 5:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD06":Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 6:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD07":Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 7:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD08":Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 8:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD09":Drain.timerinterval=2000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 9:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD10":Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 10:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD11":Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 11:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD12":Drain.timerinterval=9000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 12:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD13":Drain.timerinterval=2000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 13:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD14":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 14:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD15":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 15:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD16":Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 16:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD17":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 17:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD18":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 18:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD19":Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 19:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD20":Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 20:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD21":Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			Case 21:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD22":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
			End Select
            sw4.timerenabled=0:sw14.timerenabled=0
      Else
			Dim m
			m = INT(22 * RND(1) )
			Select Case m
			Case 0:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD01":Drain.timerinterval=2000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash41.visible = 1:Flasherflash39.visible = 1
			Case 1:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD02":Drain.timerinterval=7000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash30.visible = 1:Flasherflash37.visible = 1
			Case 2:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD03":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash30.visible = 1
			Case 3:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD04":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash31.visible = 1
			Case 4:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD05":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash30.visible = 1
			Case 5:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD06":Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash42.visible = 1
			Case 6:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD07":Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash30.visible = 1
			Case 7:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD08":Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash35.visible = 1
			Case 8:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD09":Drain.timerinterval=2000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash35.visible = 1
			Case 9:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD10":Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash34.visible = 1
			Case 10:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD11":Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1
			Case 11:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD12":Drain.timerinterval=9000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1
			Case 12:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD13":Drain.timerinterval=2000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash31.visible = 1
			Case 13:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD14":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1
			Case 14:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD15":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1
			Case 15:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD16":Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash36.visible = 1
			Case 16:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD17":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1
			Case 17:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD18":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash44.visible = 1
			Case 18:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD19":Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash35.visible = 1
			Case 19:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD20":Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash31.visible = 1
			Case 20:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD21":Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash43.visible = 1
			Case 21:EndMusic:StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESBD22":Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1

			End Select
            sw4.timerenabled=0:sw14.timerenabled=0
            End If
End Sub

Sub drain_timer
    Dim x
    x = INT(3 * RND(1) )
    Select Case x
    Case 0:PlayMusic"0ESB10.mp3":sw39.timerinterval=236000:sw39.timerenabled=1:sw12.timerenabled=0:Starstop
    Case 1:PlayMusic"0ESB11.mp3":sw39.timerinterval=127000:sw39.timerenabled=1:sw12.timerenabled=0:Starstop
    Case 2:PlayMusic"0ESB12.mp3":sw39.timerinterval=130000:sw39.timerenabled=1:sw12.timerenabled=0:Starstop
    End Select
Drain.timerenabled=0
End sub

Sub BallRelease_unHit():PlaySound "0ESB00z":BIP=1:AAA=0
Flasherlight30.IntensityScale = 1:Flasherlight31.IntensityScale = 1:Flasherlight32.IntensityScale = 1:Flasherlight33.IntensityScale = 1:Flasherlight34.IntensityScale = 1:Flasherlight35.IntensityScale = 1:Flasherlight36.IntensityScale = 1
End Sub

'Lower Trigger wires
Sub sw4_Hit:Controller.Switch(4)=1 : playsoundAtVol"rollover" ,ActiveBall, 1:PlaySound"0ESBC05":FlashLevel10 = 1 : FlasherFlash10_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer: End Sub
Sub sw4_UnHit:Controller.Switch(4)=0:End Sub
Sub sw4_timer
    Dim x
    x = INT(12 * RND(1) )
    Select Case x
              Case 0:PlayMusic"0ESB02.mp3":sw14.timerinterval=181000:sw14.timerenabled=1
              Case 1:PlayMusic"0ESB03.mp3":sw14.timerinterval=181000:sw14.timerenabled=1
              Case 2:PlayMusic"0ESB04.mp3":sw14.timerinterval=181000:sw14.timerenabled=1
              Case 3:PlayMusic"0ESB05.mp3":sw14.timerinterval=181000:sw14.timerenabled=1
              Case 4:PlayMusic"0ESB06.mp3":sw14.timerinterval=181000:sw14.timerenabled=1
              Case 5:PlayMusic"0ESB07.mp3":sw14.timerinterval=181000:sw14.timerenabled=1
              Case 6:PlayMusic"0ESB08.mp3":sw14.timerinterval=144000:sw14.timerenabled=1
              Case 7:PlayMusic"0ESB09.mp3":sw14.timerinterval=123000:sw14.timerenabled=1
              Case 8:PlayMusic"0ESB14.mp3":sw14.timerinterval=168000:sw14.timerenabled=1
              Case 9:PlayMusic"0ESB15.mp3":sw14.timerinterval=91000:sw14.timerenabled=1
              Case 10:PlayMusic"0ESB16.mp3":sw14.timerinterval=106000:sw14.timerenabled=1
              Case 11:PlayMusic"0ESB17.mp3":sw14.timerinterval=58000:sw14.timerenabled=1
    End Select
sw4.timerenabled=0
End sub
Sub sw5_Hit:Controller.Switch(5)=1 : playsoundAtVol"rollover" ,ActiveBall, 1:PlaySound"0ESBC05":FlashLevel9 = 1 : FlasherFlash9_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer: End Sub
Sub sw5_UnHit:Controller.Switch(5)=0:End Sub
Sub sw5_timer
    FlashLevel1 = 1 : FlasherFlash1_Timer:FlashLevel4 = 1 : FlasherFlash4_Timer

End sub
Sub sw6_Hit:Controller.Switch(6)=1 : playsoundAtVol"rollover",ActiveBall, 1:FlashLevel10 = 1 : FlasherFlash10_Timer
    Dim x
    x = INT(2 * RND(1) )
    Select Case x
    Case 0:PlaySound"0ESB00q"
    Case 1:PlaySound"0ESBC10"
    End Select
End Sub
Sub sw6_unHit:Controller.Switch(6)=0:End Sub
Sub sw6_timer
    FlashLevel18 = 1 : FlasherFlash18_Timer:FlashLevel19 = 1 : FlasherFlash19_Timer

End sub
Sub sw12_Hit:Controller.Switch(12)=1 : playsoundAtVol"rollover" ,ActiveBall, 1:PlaySound"0ESB00r":FlashLevel10 = 1 : FlasherFlash10_Timer: End Sub
Sub sw12_unHit:Controller.Switch(12)=0:End Sub
Sub sw12_timer
    Dim x
    x = INT(4 * RND(1) )
    Select Case x
    Case 0:FlashLevel1 = 1 : FlasherFlash1_Timer:FlashLevel4 = 1 : FlasherFlash4_Timer
    Case 1:FlashLevel26 = 1 : FlasherFlash26_Timer:FlashLevel19 = 1 : FlasherFlash19_Timer
    Case 2:FlashLevel20 = 1 : FlasherFlash20_Timer:FlashLevel25 = 1 : FlasherFlash25_Timer
    Case 3:FlashLevel27 = 1 : FlasherFlash27_Timer
    End Select

End sub
Sub sw17_Hit:Controller.Switch(17)=1 : playsoundAtVol"rollover",ActiveBall, 1:FlashLevel9 = 1 : FlasherFlash9_Timer
    Dim x
    x = INT(2 * RND(1) )
    Select Case x
    Case 0:PlaySound"0ESB00q"
    Case 1:PlaySound"0ESBC10"
    End Select
End Sub
Sub sw17_unHit:Controller.Switch(17)=0:End Sub
Sub sw17_timer
    Dim x
    x = INT(6 * RND(1) )
    Select Case x
    Case 0:FlashLevel2 = 1 : FlasherFlash2_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
    Case 1:FlashLevel7 = 1 : FlasherFlash7_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer
    Case 2:FlashLevel11 = 1 : FlasherFlash11_Timer:FlashLevel25 = 1 : FlasherFlash25_Timer
    Case 3:FlashLevel9 = 1 : FlasherFlash9_Timer:FlashLevel10 = 1 : FlasherFlash10_Timer
    Case 4:FlashLevel21 = 1 : FlasherFlash21_Timer
    Case 5:FlashLevel27 = 1 : FlasherFlash27_Timer
    End Select

End sub
Sub sw29_Hit:Controller.Switch(29)=1 : playsoundAtVol"rollover" ,ActiveBall, 1:PlaySound"0ESB00r":FlashLevel9 = 1 : FlasherFlash9_Timer: End Sub
Sub sw29_unHit:Controller.Switch(29)=0:End Sub
Sub sw29_timer
    sw17.timerenabled=0

sw29.timerenabled=0
End sub
'Top wire triggers
Sub sw11_Hit:Controller.Switch(11)=1 : playsoundAtVol"rollover",ActiveBall, 1:FlashLevel6 = 1 : FlasherFlash6_Timer
    Dim x
    x = INT(2 * RND(1) )
    Select Case x
    Case 0:PlaySound"0ESB00q"
    Case 1:PlaySound"0ESBC10"
    End Select
End Sub
Sub sw11_unHit:Controller.Switch(11)=0:End Sub
Sub sw11_timer
    Starstop

sw11.timerenabled=0
End sub
Sub sw13_Hit:Controller.Switch(13)=1 : playsoundAtVol"rollover",ActiveBall, 1:FlashLevel6 = 1 : FlasherFlash6_Timer
    Dim x
    x = INT(2 * RND(1) )
    Select Case x
    Case 0:PlaySound"0ESB00q"
    Case 1:PlaySound"0ESBC10"
    End Select
End Sub
Sub sw13_unHit:Controller.Switch(13)=0:End Sub
Sub sw14_Hit:Controller.Switch(14)=1 : playsoundAtVol"rollover",ActiveBall, 1:FlashLevel2 = 1 : FlasherFlash2_Timer
    Dim x
    x = INT(2 * RND(1) )
    Select Case x
    Case 0:PlaySound"0ESB00q"
    Case 1:PlaySound"0ESBC10"
    End Select
End Sub
Sub sw14_unHit:Controller.Switch(14)=0:End Sub
Sub sw14_timer
    Dim x
    x = INT(12 * RND(1) )
    Select Case x
              Case 0:PlayMusic"0ESB02.mp3":sw4.timerinterval=181000:sw4.timerenabled=1
              Case 1:PlayMusic"0ESB03.mp3":sw4.timerinterval=181000:sw4.timerenabled=1
              Case 2:PlayMusic"0ESB04.mp3":sw4.timerinterval=181000:sw4.timerenabled=1
              Case 3:PlayMusic"0ESB05.mp3":sw4.timerinterval=181000:sw4.timerenabled=1
              Case 4:PlayMusic"0ESB06.mp3":sw4.timerinterval=181000:sw4.timerenabled=1
              Case 5:PlayMusic"0ESB07.mp3":sw4.timerinterval=181000:sw4.timerenabled=1
              Case 6:PlayMusic"0ESB08.mp3":sw4.timerinterval=144000:sw4.timerenabled=1
              Case 7:PlayMusic"0ESB09.mp3":sw4.timerinterval=123000:sw4.timerenabled=1
              Case 8:PlayMusic"0ESB14.mp3":sw4.timerinterval=168000:sw4.timerenabled=1
              Case 9:PlayMusic"0ESB15.mp3":sw4.timerinterval=91000:sw4.timerenabled=1
              Case 10:PlayMusic"0ESB16.mp3":sw4.timerinterval=106000:sw4.timerenabled=1
              Case 11:PlayMusic"0ESB17.mp3":sw4.timerinterval=58000:sw4.timerenabled=1
    End Select
sw14.timerenabled=0
End sub
Sub sw15_Hit:Controller.Switch(15)=1 : playsoundAtVol"rollover",ActiveBall, 1:FlashLevel2 = 1 : FlasherFlash2_Timer
    Dim x
    x = INT(2 * RND(1) )
    Select Case x
    Case 0:PlaySound"0ESB00q"
    Case 1:PlaySound"0ESBC10"
    End Select
End Sub
Sub sw15_unHit:Controller.Switch(15)=0:End Sub

'Plunger lane trigger wire
Sub sw39_Hit:Controller.Switch(39)=1 : playsoundAtVol"rollover",ActiveBall,1
If AAA=1 Then BIP=1 Else BIP=0
If BBB=1 and Starpic=0 Then PlaySound"0ESBD17"
If BBB=1 and Starpic=1 Then PlaySound"0ESBD17":Flasherflash28.visible = 1
End Sub
Sub sw39_unHit:Controller.Switch(39)=0:End Sub

Sub sw39_timer
   If Starpic=0 Then
    Dim x
    x = INT(38 * RND(1) )
    Select Case x
        Case 0:StopESBSounds:Playsound("0ESB00a"):Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 1:StopESBSounds:Playsound("0ESB00b"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 2:StopESBSounds:Playsound("0ESB00c"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 3:StopESBSounds:Playsound("0ESB00d"):Drain.timerinterval=2000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 4:StopESBSounds:Playsound("0ESB00e"):Drain.timerinterval=2000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 5:StopESBSounds:Playsound("0ESB00f"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 6:StopESBSounds:Playsound("0ESB00g"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 7:StopESBSounds:Playsound("0ESB00h"):Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 8:StopESBSounds:Playsound("0ESB00i"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 9:StopESBSounds:Playsound("0ESB00j"):Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 10:StopESBSounds:Playsound("0ESB00k"):Drain.timerinterval=9000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 11:StopESBSounds:Playsound("0ESB00l"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 12:StopESBSounds:Playsound("0ESB00m"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 13:StopESBSounds:Playsound("0ESB00n"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 14:StopESBSounds:Playsound("0ESB00o"):Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 15:StopESBSounds:Playsound("0ESB00p"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 16:StopESBSounds:Playsound("0ESB00za"):Drain.timerinterval=6500:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 17:StopESBSounds:Playsound("0ESB00zb"):Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 18:StopESBSounds:Playsound("0ESB00zc"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 19:StopESBSounds:Playsound("0ESB00zd"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 20:StopESBSounds:Playsound("0ESB00ze"):Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 21:StopESBSounds:Playsound("0ESB00zk"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 22:StopESBSounds:Playsound("0ESB00zl"):Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 23:StopESBSounds:Playsound("0ESB00zm"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 24:StopESBSounds:Playsound("0ESB00zn"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 25:StopESBSounds:Playsound("0ESB00za"):Drain.timerinterval=7000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 26:StopESBSounds:Playsound("0ESB00zp"):Drain.timerinterval=7000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 27:StopESBSounds:Playsound("0ESB00zq"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 28:StopESBSounds:Playsound("0ESB00zr"):Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 29:StopESBSounds:Playsound("0ESB00zs"):Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 30:StopESBSounds:Playsound("0ESB00zt"):Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 31:StopESBSounds:Playsound("0ESB00zu"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 32:StopESBSounds:Playsound("0ESB00zv"):Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 33:StopESBSounds:Playsound("0ESB00zw"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 34:StopESBSounds:Playsound("0ESB00zx"):Drain.timerinterval=7000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 35:StopESBSounds:Playsound("0ESB00zy"):Drain.timerinterval=7000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 36:StopESBSounds:Playsound("0ESB00zz"):Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 37:StopESBSounds:Playsound("0ESB00zza"):Drain.timerinterval=7000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        End Select
        sw39.timerenabled=0
   Else
        Dim l
        l = INT(38 * RND(1) )
        Select Case l
        Case 0:StopESBSounds:Playsound("0ESB00a"):Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash37.visible = 1
        Case 1:StopESBSounds:Playsound("0ESB00b"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash35.visible = 1
        Case 2:StopESBSounds:Playsound("0ESB00c"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1:Flasherflash43.visible = 1
        Case 3:StopESBSounds:Playsound("0ESB00d"):Drain.timerinterval=2000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash36.visible = 1
        Case 4:StopESBSounds:Playsound("0ESB00e"):Drain.timerinterval=2000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash41.visible = 1
        Case 5:StopESBSounds:Playsound("0ESB00f"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash36.visible = 1
        Case 6:StopESBSounds:Playsound("0ESB00g"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash31.visible = 1
        Case 7:StopESBSounds:Playsound("0ESB00h"):Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash36.visible = 1:Flasherflash40.visible = 1
        Case 8:StopESBSounds:Playsound("0ESB00i"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1
        Case 9:StopESBSounds:Playsound("0ESB00j"):Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1
        Case 10:StopESBSounds:Playsound("0ESB00k"):Drain.timerinterval=9000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash30.visible = 1
        Case 11:StopESBSounds:Playsound("0ESB00l"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash30.visible = 1
        Case 12:StopESBSounds:Playsound("0ESB00m"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash30.visible = 1
        Case 13:StopESBSounds:Playsound("0ESB00n"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash30.visible = 1
        Case 14:StopESBSounds:Playsound("0ESB00o"):Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash38.visible = 1
        Case 15:StopESBSounds:Playsound("0ESB00p"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1
        Case 16:StopESBSounds:Playsound("0ESB00za"):Drain.timerinterval=6500:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash31.visible = 1
        Case 17:StopESBSounds:Playsound("0ESB00zb"):Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash30.visible = 1
        Case 18:StopESBSounds:Playsound("0ESB00zc"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash31.visible = 1
        Case 19:StopESBSounds:Playsound("0ESB00zd"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1
        Case 20:StopESBSounds:Playsound("0ESB00ze"):Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1
        Case 21:StopESBSounds:Playsound("0ESB00zk"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash41.visible = 1
        Case 22:StopESBSounds:Playsound("0ESB00zl"):Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash31.visible = 1:Flasherflash39.visible =1
        Case 23:StopESBSounds:Playsound("0ESB00zm"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1
        Case 24:StopESBSounds:Playsound("0ESB00zn"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash39.visible = 1:Flasherflash41.visible = 1
        Case 25:StopESBSounds:Playsound("0ESB00za"):Drain.timerinterval=7000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash31.visible = 1
        Case 26:StopESBSounds:Playsound("0ESB00zp"):Drain.timerinterval=7000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1
        Case 27:StopESBSounds:Playsound("0ESB00zq"):Drain.timerinterval=4000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash41.visible = 1
        Case 28:StopESBSounds:Playsound("0ESB00zr"):Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash29.visible = 1
        Case 29:StopESBSounds:Playsound("0ESB00zs"):Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash30.visible = 1:Flasherflash42.visible = 1
        Case 30:StopESBSounds:Playsound("0ESB00zt"):Drain.timerinterval=5000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash30.visible = 1:Flasherflash42.visible = 1
        Case 31:StopESBSounds:Playsound("0ESB00zu"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash40.visible = 1
        Case 32:StopESBSounds:Playsound("0ESB00zv"):Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash37.visible = 1
        Case 33:StopESBSounds:Playsound("0ESB00zw"):Drain.timerinterval=3000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash39.visible = 1:Flasherflash41.visible = 1
        Case 34:StopESBSounds:Playsound("0ESB00zx"):Drain.timerinterval=7000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash28.visible = 1:Flasherflash42.visible = 1
        Case 35:StopESBSounds:Playsound("0ESB00zy"):Drain.timerinterval=7000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash31.visible = 1:Flasherflash39.visible = 1
        Case 36:StopESBSounds:Playsound("0ESB00zz"):Drain.timerinterval=6000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash29.visible = 1
        Case 37:StopESBSounds:Playsound("0ESB00zza"):Drain.timerinterval=7000:Drain.timerenabled=1:sw12.timerinterval=200:sw12.timerenabled=1:Flasherflash32.visible = 1
        End Select
        sw39.timerenabled=0
        End If
End sub

'Spinners
  Sub Spinner1_Spin():vpmTimer.pulsesw 7 : playsoundAtVol"fx_spinner",spinner1,VolSpin :FlashLevel25 = 1 : FlasherFlash25_Timer:FlashLevel20 = 1 : FlasherFlash20_Timer:FlashLevel2 = 1 : FlasherFlash2_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer: End Sub
  Sub Spinner2_Spin():vpmTimer.pulsesw 8 : playsoundAtVol"fx_spinner",spinner2,VolSpin :FlashLevel1 = 1 : FlasherFlash1_Timer:FlashLevel4 = 1 : FlasherFlash4_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer: End Sub

'Stand Up targets
Sub sw18_Hit:vpmTimer.PulseSw 18:FlashLevel22 = 1 : FlasherFlash22_Timer
    Dim x
    x = INT(3 * RND(1) )
    Select Case x
    Case 0:PlaySound"0ESB00s":StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESB00u"
    Case 1:PlaySound"0ESB00s":StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESB00w"
    Case 2:PlaySound"0ESB00s":StopSound"0ESB00u":StopSound"0ESB00w":StopSound"0ESB00x":PlaySound"0ESB00x"
    End Select
End Sub

Sub sw19_Hit:vpmTimer.PulseSw 19 :PlaySound"0ESB00s":FlashLevel8 = 1 : FlasherFlash8_Timer: End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22 :PlaySound"0ESB00s":FlashLevel8 = 1 : FlasherFlash8_Timer: End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23 :PlaySound"0ESB00s":FlashLevel8 = 1 : FlasherFlash8_Timer: End Sub

Sub sw20_Hit:vpmTimer.PulseSw 20 :PlaySound"0ESB00s":FlashLevel7 = 1 : FlasherFlash7_Timer: End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21 :PlaySound"0ESB00s":FlashLevel7 = 1 : FlasherFlash7_Timer: End Sub
Sub sw24_Hit:vpmTimer.PulseSw 24 :PlaySound"0ESB00s":FlashLevel7 = 1 : FlasherFlash7_Timer: End Sub


'Drop Targets
 Sub Sw25_Dropped:   dtCDrop.Hit 1 :PlaySound"0ESB00t":FlashLevel11 = 1 : FlasherFlash11_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer:  End Sub
 Sub Sw26_Dropped:   dtCDrop.Hit 2 :PlaySound"0ESB00t":FlashLevel11 = 1 : FlasherFlash11_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer:  End Sub
 Sub Sw27_Dropped:   dtCDrop.Hit 3 :PlaySound"0ESB00t":FlashLevel11 = 1 : FlasherFlash11_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer:  End Sub

 Sub sw28_Dropped:   dtTDrop.Hit 1 :PlaySound"0ESB00z":AAA=1:BBB=1:FlashLevel20 = 1 : FlasherFlash20_Timer:FlashLevel25 = 1 : FlasherFlash25_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer:  End Sub

'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw(31) : playsound SoundFX("fx_bumper1",DOFContactors):PlaySound"0ESB00y":FlashLevel17 = 1 : FlasherFlash17_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer: End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw(31) : playsound SoundFX("fx_bumper1",DOFContactors):PlaySound"0ESB00y":FlashLevel18 = 1 : FlasherFlash18_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer: End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw(33) : playsound SoundFX("fx_bumper1",DOFContactors):PlaySound"0ESB00v":FlashLevel3 = 1 : FlasherFlash3_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer: End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw(34) : playsound SoundFX("fx_bumper1",DOFContactors):PlaySound"0ESB00v":FlashLevel12 = 1 : FlasherFlash12_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer: End Sub
Sub Bumper5_Hit:vpmTimer.PulseSw(35) : playsound"metalhit2" :PlaySound"0ESB00v":FlashLevel13 = 1 : FlasherFlash13_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer: End Sub
Sub Bumper6_Hit:vpmTimer.PulseSw(38) : playsound"metalhit2" :PlaySound"0ESB00v":FlashLevel14 = 1 : FlasherFlash14_Timer:FlashLevel27 = 1 : FlasherFlash27_Timer: End Sub

'**********************************************************************************************************

'Map lights to an array
'**********************************************************************************************************
Set Lights(1)=Light1
Set Lights(2)=Light2
Set Lights(3)=Light3
Lights(4)=array(Light4,Light4a)
Set Lights(5)=Light5 'bumper3
Set Lights(6)=Light6
Set Lights(9)=Light9
Set Lights(11)=Light11
Set Lights(12)=Light12
Set Lights(14)=Light14
Set Lights(17)=Light17
Set Lights(18)=Light18
Set Lights(19)=Light19
Lights(20)=array(Light20,Light20a)
Set Lights(21)=Light21 'bumpe6
Set Lights(22)=Light22
Set Lights(23)=Light23
Set Lights(24)=Light24
Set Lights(25)=Light25
Set Lights(26)=Light26
Set Lights(28)=Light28
Set Lights(30)=Light30
Set Lights(33)=Light33
Set Lights(34)=Light34
Set Lights(35)=Light35
Set Lights(36)=Light36
Set Lights(37)=Light37
Set Lights(38)=Light38
Set Lights(39)=Light39
Set Lights(40)=Light40
Set Lights(41)=Light41
Set Lights(42)=Light42
Set Lights(43)=Light43
Set Lights(44)=Light44 'bumper4
Set Lights(46)=Light46
Set Lights(49)=Light49
Set Lights(50)=Light50
Set Lights(51)=Light51
Set Lights(52)=Light52
Set Lights(53)=Light53
Set Lights(54)=Light54
Set Lights(55)=Light55
Set Lights(56)=Light56
Set Lights(57)=Light57
Set Lights(58)=Light58
Set Lights(59)=Light59
Set Lights(60)=Light60 'bumper5
Set Lights(62)=Light62
Set Lights(63)=Light63

'Backglass
'Set Lights(13)=Light13
'Set Lights(29)=Light29
'Set Lights(45)=Light45
'Set Lights(61)=Light61

Dim Digits(32)
Digits(0)=Array(a00,a01,a02,a03,a04,a05,a06,n,a08)
Digits(1)=Array(a10,a11,a12,a13,a14,a15,a16,n,a18)
Digits(2)=Array(a20,a21,a22,a23,a24,a25,a26,n,a28)
Digits(3)=Array(a30,a31,a32,a33,a34,a35,a36,n,a38)
Digits(4)=Array(a40,a41,a42,a43,a44,a45,a46,n,a48)
Digits(5)=Array(a50,a51,a52,a53,a54,a55,a56,n,a58)
Digits(6)=Array(b00,b01,b02,b03,b04,b05,b06,n,b08)
Digits(7)=Array(b10,b11,b12,b13,b14,b15,b16,n,b18)
Digits(8)=Array(b20,b21,b22,b23,b24,b25,b26,n,b28)
Digits(9)=Array(b30,b31,b32,b33,b34,b35,b36,n,b38)
Digits(10)=Array(b40,b41,b42,b43,b44,b45,b46,n,b48)
Digits(11)=Array(b50,b51,b52,b53,b54,b55,b56,n,b58)
Digits(12)=Array(c00,c01,c02,c03,c04,c05,c06,n,c08)
Digits(13)=Array(c10,c11,c12,c13,c14,c15,c16,n,c18)
Digits(14)=Array(c20,c21,c22,c23,c24,c25,c26,n,c28)
Digits(15)=Array(c30,c31,c32,c33,c34,c35,c36,n,c38)
Digits(16)=Array(c40,c41,c42,c43,c44,c45,c46,n,c48)
Digits(17)=Array(c50,c51,c52,c53,c54,c55,c56,n,c58)
Digits(18)=Array(d00,d01,d02,d03,d04,d05,d06,n,d08)
Digits(19)=Array(d10,d11,d12,d13,d14,d15,d16,n,d18)
Digits(20)=Array(d20,d21,d22,d23,d24,d25,d26,n,d28)
Digits(21)=Array(d30,d31,d32,d33,d34,d35,d36,n,d38)
Digits(22)=Array(d40,d41,d42,d43,d44,d45,d46,n,d48)
Digits(23)=Array(d50,d51,d52,d53,d54,d55,d56,n,d58)
Digits(26)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(27)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)
Digits(24)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(25)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)



Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
		If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 32) then
				For Each obj In Digits(num)
					If chg And 1 Then obj.State = stat And 1
					chg = chg\2 : stat = stat\2
				Next
			else
				'if char(stat) > "" then msg(num) = char(stat)
			end if
		next
		end if
end if
End Sub


'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 37
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    PlaySound"0ESB00y":FlashLevel16 = 1 : FlasherFlash16_Timer
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 36
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors),sling2, 1
    PlaySound"0ESB00y":FlashLevel15 = 1 : FlasherFlash15_Timer
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub Table1_Exit():Controller.Stop:Controller.Games("empsback").Settings.Value("sound")=1:End Sub

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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Const tnob = 5 ' total number of balls
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
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

