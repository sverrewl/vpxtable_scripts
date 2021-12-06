' Scorpion Williams 1980
' Allknowing2012 - PF, plastics from internet
' IPDB for switches and VP9 table
' Script segments from various tables and authors
' 1.0b - hide the backglass lights when no int desktop mode
' 1.0c - thanks to Wolis who also reached out with some code/setting ideas (unreleased)
' 2.0 - big graphix updates by Thalamus and Hauntfreaks. Flashers and multiball bug fix
' 2.1 - added flipper and ball shadows from Hauntfreaks Cartoons, PMD Sounds
' 2.1a - remove lane switch pulse - also tweak one rubber so it doesnt alwayus exit to the right :-)

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
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


'**********************************************************
'********     OPTIONS   *******************************
'**********************************************************

Dim BallShadows: Ballshadows=1      '****************** set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1    '***********    set to 1 to turn on Flipper shadows

If Table.ShowDT = False then
  light64.visible=False
  light57.visible=False
  light58.visible=False
  light59.visible=False
  light60.visible=False
  light50.visible=False
  light51.visible=False
  light52.visible=False
  light53.visible=False
  SideWood.visible=False
  Topwood.visible=False
  Flasherpic1.visible=True
End If

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "scrpn_l1"

Dim xx
Dim dtl,dtC1,dtC2,dtR,bsSaucer,bsSaucer2,bsTrough
Dim BIP,aa,bb,cc,dd,ee,ff

'**********************************************
'Fluppers Flashers

Dim FlashLevel1, FlashLevel2, FlashLevel3, FlashLevel4, FlashLevel5, FlashLevel6, FlashLevel7, FlashLevel8, FlashLevel9, FlashLevel10
FlasherLight1.IntensityScale = 0
FlasherLight2.IntensityScale = 0
FlasherLight3.IntensityScale = 0
FlasherLight4.IntensityScale = 0
FlasherLight5.IntensityScale = 0
FlasherLight6.IntensityScale = 0
FlasherLight7.IntensityScale = 0
FlasherLight8.IntensityScale = 0
FlasherLight9.IntensityScale = 0
FlasherLight10.IntensityScale = 0

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

'*** white flasher ***
Sub FlasherFlash2_Timer()
  dim flashx3, matdim
  If not FlasherFlash2.TimerEnabled Then
    FlasherFlash2.TimerEnabled = True
    FlasherFlash2.visible = 1
    FlasherLit2.visible = 1
  End If
  flashx3 = FlashLevel2 * FlashLevel2 * FlashLevel2
  Flasherflash2.opacity = 1000 * flashx3
  FlasherLit2.BlendDisableLighting = 10 * flashx3
  Flasherbase2.BlendDisableLighting =  flashx3
  FlasherLight2.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel2)
  FlasherLit2.material = "domelit" & matdim
  FlashLevel2 = FlashLevel2 * 0.85 - 0.01
  If FlashLevel2 < 0.15 Then
    FlasherLit2.visible = 0
  Else
    FlasherLit2.visible = 1
  end If
  If FlashLevel2 < 0 Then
    FlasherFlash2.TimerEnabled = False
    FlasherFlash2.visible = 0
  End If
End Sub

'*** white flasher ***
Sub FlasherFlash3_Timer()
  dim flashx3, matdim
  If not FlasherFlash3.TimerEnabled Then
    FlasherFlash3.TimerEnabled = True
    FlasherFlash3.visible = 1
    FlasherLit3.visible = 1
  End If
  flashx3 = FlashLevel3 * FlashLevel3 * FlashLevel3
  Flasherflash3.opacity = 1000 * flashx3
  FlasherLit3.BlendDisableLighting = 10 * flashx3
  Flasherbase3.BlendDisableLighting =  flashx3
  FlasherLight3.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel3)
  FlasherLit3.material = "domelit" & matdim
  FlashLevel3 = FlashLevel3 * 0.85 - 0.01
  If FlashLevel3 < 0.15 Then
    FlasherLit3.visible = 0
  Else
    FlasherLit3.visible = 1
  end If
  If FlashLevel3 < 0 Then
    FlasherFlash3.TimerEnabled = False
    FlasherFlash3.visible = 0
  End If
End Sub

'*** white flasher ***
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

'*** white flasher ***
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


'*** white flasher ***
Sub FlasherFlash6_Timer()
  dim flashx3, matdim
  If not FlasherFlash6.TimerEnabled Then
    FlasherFlash6.TimerEnabled = True
    FlasherFlash6.visible = 1
    FlasherLit6.visible = 1
  End If
  flashx3 = FlashLevel6 * FlashLevel6 * FlashLevel6
  Flasherflash6.opacity = 1000 * flashx3
  FlasherLit6.BlendDisableLighting = 10 * flashx3
  Flasherbase6.BlendDisableLighting =  flashx3
  FlasherLight6.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel6)
  FlasherLit6.material = "domelit" & matdim
  FlashLevel6 = FlashLevel6 * 0.85 - 0.01
  If FlashLevel6 < 0.15 Then
    FlasherLit6.visible = 0
  Else
    FlasherLit6.visible = 1
  end If
  If FlashLevel6 < 0 Then
    FlasherFlash6.TimerEnabled = False
    FlasherFlash6.visible = 0
  End If
End Sub


'*** blue flasher ***
Sub FlasherFlash7_Timer()
  dim flashx3, matdim
  If not Flasherflash7.TimerEnabled Then
    Flasherflash7.TimerEnabled = True
    Flasherflash7.visible = 1
    Flasherlit7.visible = 1
  End If
  flashx3 = FlashLevel7 * FlashLevel7 * FlashLevel7
  Flasherflash7.opacity = 8000 * flashx3
  Flasherlit7.BlendDisableLighting = 10 * flashx3
  Flasherbase7.BlendDisableLighting =  flashx3
  Flasherlight7.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel7)
  Flasherlit7.material = "domelit" & matdim
  FlashLevel7 = FlashLevel7 * 0.85 - 0.01
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


'*** white flasher ***
Sub FlasherFlash8_Timer()
  dim flashx3, matdim
  If not FlasherFlash8.TimerEnabled Then
    FlasherFlash8.TimerEnabled = True
    FlasherFlash8.visible = 1
    FlasherLit8.visible = 1
  End If
  flashx3 = FlashLevel8 * FlashLevel8 * FlashLevel8
  Flasherflash8.opacity = 1000 * flashx3
  FlasherLit8.BlendDisableLighting = 10 * flashx3
  Flasherbase8.BlendDisableLighting =  flashx3
  FlasherLight8.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel8)
  FlasherLit8.material = "domelit" & matdim
  FlashLevel8 = FlashLevel8 * 0.85 - 0.01
  If FlashLevel8 < 0.15 Then
    FlasherLit8.visible = 0
  Else
    FlasherLit8.visible = 1
  end If
  If FlashLevel8 < 0 Then
    FlasherFlash8.TimerEnabled = False
    FlasherFlash8.visible = 0
  End If
End Sub


'*** Red flasher ***
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


'*** white flasher ***
Sub FlasherFlash10_Timer()
  dim flashx3, matdim
  If not FlasherFlash10.TimerEnabled Then
    FlasherFlash10.TimerEnabled = True
    FlasherFlash10.visible = 1
    FlasherLit10.visible = 1
  End If
  flashx3 = FlashLevel10 * FlashLevel10 * FlashLevel10
  Flasherflash10.opacity = 1000 * flashx3
  FlasherLit10.BlendDisableLighting = 10 * flashx3
  Flasherbase10.BlendDisableLighting =  flashx3
  FlasherLight10.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel10)
  FlasherLit10.material = "domelit" & matdim
  FlashLevel10 = FlashLevel10 * 0.85 - 0.01
  If FlashLevel10 < 0.15 Then
    FlasherLit10.visible = 0
  Else
    FlasherLit10.visible = 1
  end If
  If FlashLevel10 < 0 Then
    FlasherFlash10.TimerEnabled = False
    FlasherFlash10.visible = 0
  End If
End Sub


'**********************************************
'STAT's Modified Change Team script

Dim TotalBump, BumpNow, TotalBumpb, Bumpnowb, TotalBumpc, Bumpnowc, TotalBumpd, Bumpnowd


TotalBumpd = 2

Sub Change8Bump
  BumpNowd = BumpNowd + 1
  If BumpNowd > 1 Then
    WallApron.image = "scorpion-apron-"&BumpNowd
    WallApronLower.image = "scorpion-apron-"&BumpNowd
  Else
    WallApron.image = "scorpion-apron"
    WallApronLower.image = "scorpion-apron"
  End If
  If BumpNowd = TotalBumpd Then BumpNowd = 0
End Sub


TotalBumpc = 6

Sub Change6Bump
  BumpNowc = BumpNowc + 1
  If BumpNowc > 1 Then
    Target1.image = "DropTarget-scorpion-"&BumpNowc
    Target2.image = "DropTarget-scorpion-"&BumpNowc
    Target3.image = "DropTarget-scorpion-"&BumpNowc
    Target4.image = "DropTarget-scorpion-"&BumpNowc
    Target5.image = "DropTarget-scorpion-"&BumpNowc
    Target6.image = "DropTarget-scorpion-"&BumpNowc
    Target7.image = "DropTarget-scorpion-"&BumpNowc
    Target8.image = "DropTarget-scorpion-"&BumpNowc
    Target9.image = "DropTarget-scorpion-"&BumpNowc
    Target10.image = "DropTarget-scorpion-"&BumpNowc
    Target11.image = "DropTarget-scorpion-"&BumpNowc
  Else
    Target1.image = "DropTarget-scorpion"
    Target2.image = "DropTarget-scorpion"
    Target3.image = "DropTarget-scorpion"
    Target4.image = "DropTarget-scorpion"
    Target5.image = "DropTarget-scorpion"
    Target6.image = "DropTarget-scorpion"
    Target7.image = "DropTarget-scorpion"
    Target8.image = "DropTarget-scorpion"
    Target9.image = "DropTarget-scorpion"
    Target10.image = "DropTarget-scorpion"
    Target11.image = "DropTarget-scorpion"
  End If
  If BumpNowc = TotalBumpc Then BumpNowc = 0
End Sub


TotalBump = 10

Sub Change4Bump
  BumpNow = BumpNow + 1
  If BumpNow > 1 Then
    PL3.image = "scorpion-bumpercap-"&BumpNow
    PL2.image = "scorpion-bumpercap-"&BumpNow
    PL1.image = "scorpion-bumpercap-"&BumpNow
  Else
    PL3.image = "scorpion-bumpercap"
    PL2.image = "scorpion-bumpercap"
    PL1.image = "scorpion-bumpercap"
  End If
  If BumpNow = TotalBump Then BumpNow = 0
End Sub

TotalBumpb = 18

Sub Change2Bump
  BumpNowb = BumpNowb + 1
  If BumpNowb = 1 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 1:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 2 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 1:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 3 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 1:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 4 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 1:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 5 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 1
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 6 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 1:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 7 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 1:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 8 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 1:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 9 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 1:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 10 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 1:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 11 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 1
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 12 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 1:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 13 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 1:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 14 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 1:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 15 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 1:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
  If BumpNowb = 16 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 1:Flasherpic18.visible = 0
  End If
  If BumpNowb = 17 Then
    Flasherpic1.visible = 0:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 1
  End If

  If BumpNowb = TotalBumpb Then BumpNowb = 0
  If BumpNowb = 0 Then
    Flasherpic1.visible = 1:Flasherpic2.visible = 0:Flasherpic3.visible = 0:Flasherpic4.visible = 0:Flasherpic5.visible = 0:Flasherpic6.visible = 0
    Flasherpic7.visible = 0:Flasherpic8.visible = 0:Flasherpic9.visible = 0:Flasherpic10.visible = 0:Flasherpic11.visible = 0:Flasherpic12.visible = 0
    Flasherpic13.visible = 0:Flasherpic14.visible = 0:Flasherpic15.visible = 0:Flasherpic16.visible = 0:Flasherpic17.visible = 0:Flasherpic18.visible = 0
  End If
End Sub

'********************************************


' Thalamus, heavier ball please

Const BallMass = 1.5

LoadVPM "01560000", "S6.VBS", 3.36
Dim DesktopMode: DesktopMode = Table.ShowDT

Const UseSolenoids=2,UseLamps=1,UseSync=1,UseGI=0
Const SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="coin3"


Const sOutHole=1
Const sC1DTReset=2
Const sC2DTReset=3
Const sLDTReset=4
Const sRDTReset=5         ' solenoid for Top Right Targets
Const sBallRelease=6
Const sKicker1=7
Const sKicker3=8

Const sKnocker=14
Const sbuzzer=15
Const sCoinLockout=16
Const sEnable=23

SolCallback(sOutHole)="bsTrough.SolIn"
SolCallback(sKnocker)="vpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(sLDTReset)="SolRaiseDropL"  'dtL.SolDropUp"  ' Add a delay
SolCallback(sRDTReset)="SolRaiseDropR"  'dtR.SolDropUp"  ' Add a delay
SolCallback(sBallRelease)="bsTrough.SolOut"
SolCallback(sC1DTReset)="SolRaiseDropC1"  'Add a delay
SolCallback(sC2DTReset)="SolRaiseDropC2"  'Add a delay
SolCallback(sKicker1)="bsSaucer.SolOut"
SolCallback(sKicker3)="bsSaucer2.SolOut"

SolCallback(sLRFlipper)="SolRFlipper"
SolCallback(sLLFlipper)="SolLFlipper"

'Fantasy Bumper Lights
solCallback(17)= "vpmFlasher array(Flasher1,Flasher2),"
SolCallback(18)= "vpmFlasher array(Flasher3,Flasher4),"
SolCallback(19)= "vpmFlasher array(Flasher5,Flasher6),"

SolCallback(sEnable)="GameOn"

Sub GameOn(enabled)
  vpmNudge.SolGameOn(enabled)
  If Enabled Then
    GIOn:ee=1
  Else
    GIOff:ee=0:PlaySound"0BSGA19":Target4.timerinterval=34000:Target4.timerenabled=1:Target6.timerenabled=0
  End If
End Sub

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("fx_flipperupLeft",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToEnd:LeftFlipperUpper.RotateToEnd
         PlaySoundAtVol "fx_flipperupLeft",LeftFlipperUpper,VolFlip
     Else
         PlaySoundAtVol SoundFx("fx_flipperdown",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToStart::LeftFlipperUpper.RotateToStart
         PlaySoundAtVol "fx_flipperdown",LeftFlipperUpper,VolFlip
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("fx_flipperupRight",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToEnd:RightFlipperUpper.RotateToEnd
         PlaySoundAtVol "fx_flipperupRight",RightFlipperUpper,VolFlip
      '   vpmTimer.PulseSw(17)   ' Lane Switch  not needed here .. called via solenoid calls
     Else
         PlaySoundAtVol SoundFx("fx_flipperdown",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToStart:RightFlipperUpper.RotateToStart
         PlaySoundAtVol "fx_flipperdown",RightFlipperUpper,VolFlip
     End If
End Sub

Sub SolRaiseDropL(enabled)
   debug.print "SolRaiseDropL"
   If enabled Then
      TargetResetL.interval=500
      TargetResetL.enabled=True
   End if
End Sub

Sub TargetResetL_Timer()
  debug.print "Raise Left Targets"
  TargetResetL.enabled=False
  dtL.DropSol_On
End Sub

Sub SolRaiseDropR(enabled)
   debug.print "SolRaiseDropR"
   If enabled Then
      TargetResetR.interval=500
      TargetResetR.enabled=True
   End if
End Sub

Sub TargetResetR_Timer()
  debug.print "Raise Right Targets"
  TargetResetR.enabled=False
  dtR.DropSol_On
End Sub

Sub SolRaiseDropC1(enabled)
   If enabled Then
      TargetResetC1.interval=500
      TargetResetC1.enabled=True
   End if
End Sub

Sub TargetResetC1_Timer()
  TargetResetC1.enabled=False
  dtC1.DropSol_On
End Sub

Sub SolRaiseDropC2(enabled)
   If enabled Then
      TargetResetC2.interval=500
      TargetResetC2.enabled=True
   End if
End Sub

Sub TargetResetC2_Timer()
  TargetResetC2.enabled=False
  dtC2.DropSol_On
End Sub

Sub StopSounds()

   StopSound"0BSG01"
   StopSound"0BSG02"
   StopSound"0BSG03"
   StopSound"0BSG04"
   StopSound"0BSG05"
   StopSound"0BSG06"
   StopSound"0BSG07"
   StopSound"0BSG08"
   StopSound"0BSG09"
   StopSound"0BSG10"
   StopSound"0BSG11"
   StopSound"0BSG12"
   StopSound"0BSG13"
   StopSound"0BSG14"
   StopSound"0BSG15"
   StopSound"0BSG16"
   StopSound"0BSG17"
   StopSound"0BSG18"
   StopSound"0BSG19"
   StopSound"0BSG20"
   StopSound"0BSG21"
   StopSound"0BSG22"
   StopSound"0BSG23"
   StopSound"0BSG24"
   StopSound"0BSG25"
   StopSound"0BSG26"
   StopSound"0BSG27"
   StopSound"0BSG28"
   StopSound"0BSG29"
   StopSound"0BSG30"
   StopSound"0BSG31"
   StopSound"0BSG32"
   StopSound"0BSG33"
   StopSound"0BSG34"
   StopSound"0BSG35"
   StopSound"0BSG36"
   StopSound"0BSG37"
   StopSound"0BSG37a"
   StopSound"0BSG37b"
   StopSound"0BSG37c"
   StopSound"0BSG37d"
   StopSound"0BSG37e"
   StopSound"0BSG37f"
   StopSound"0BSG37g"
   StopSound"0BSG37h"
   StopSound"0BSG38"
   StopSound"0BSG39"
   StopSound"0BSG40"
   StopSound"0BSG41"
   StopSound"0BSG42"
   StopSound"0BSG43"
   StopSound"0BSG44"
   StopSound"0BSG45"
   StopSound"0BSG46"
   StopSound"0BSG47"
   StopSound"0BSG48"
   StopSound"0BSG49"
   StopSound"0BSG50"
   StopSound"0BSG51"
   StopSound"0BSG52"
   StopSound"0BSG53"
   StopSound"0BSG54"
   StopSound"0BSG55"
   StopSound"0BSG56"
   StopSound"0BSG57"
   StopSound"0BSG58"
   StopSound"0BSG59"
   StopSound"0BSG60"
   StopSound"0BSG61"
   StopSound"0BSG62"
   StopSound"0BSG63"
   StopSound"0BSG64"
   StopSound"0BSG65"
   StopSound"0BSG66"
   StopSound"0BSGA19"
   StopSound"0BGC01"
   StopSound"0BGC02"
   StopSound"0BGC03"
   StopSound"0BGC04"
   StopSound"0BGC05"
   StopSound"0BGC06"
   StopSound"0BGC07"
   StopSound"0BGC08"
   StopSound"0BGC09"
   StopSound"0BGC10"
   StopSound"0BGC11"
   StopSound"0BGC12"
   StopSound"0BGC13"
   StopSound"0BGC14"
   StopSound"0BGC15"
   StopSound"0BGC16"
   StopSound"0BGC17"
   StopSound"0BGC18"
   StopSound"0BGC19"
   StopSound"0BGC20"
   StopSound"0BGC21"
   StopSound"0BGC22"
   StopSound"0BGC23"
   StopSound"0BGC24"
   StopSound"0BGCC01"
   StopSound"0BGCC02"
   StopSound"0BGCC03"
   StopSound"0BGCC04"
   StopSound"0BGCC05"
   StopSound"0BGL01"
   StopSound"0BGL02"
   StopSound"0BGL03"
   StopSound"0BGL04"
   StopSound"0BGL05"
   StopSound"0BGL06"
   StopSound"0BGL07"
   StopSound"0BGL08"
   StopSound"0BGL09"
   StopSound"0BGL10"
   StopSound"0BGL12"
   StopSound"0BGL13"
End Sub

Sub StopMusic()

   StopSound"0BGC01"
   StopSound"0BGC02"
   StopSound"0BGC03"
   StopSound"0BGC04"
   StopSound"0BGC05"
   StopSound"0BGC06"
   StopSound"0BGC07"
   StopSound"0BGC08"
   StopSound"0BGC09"
   StopSound"0BGC10"
   StopSound"0BGC11"
   StopSound"0BGC12"
   StopSound"0BGC13"
   StopSound"0BGC14"
   StopSound"0BGC15"
   StopSound"0BGC16"
   StopSound"0BGC17"
   StopSound"0BGC18"
   StopSound"0BGC19"
   StopSound"0BGC20"
   StopSound"0BGC21"
   StopSound"0BGC22"
   StopSound"0BGC23"
   StopSound"0BGC24"
   StopSound"0BGCC01"
   StopSound"0BGCC02"
   StopSound"0BGCC03"
   StopSound"0BGCC04"
   StopSound"0BGCC05"

End Sub


Sub Table_Init
  vpmInit Me
  On Error Resume Next
  With Controller
    .GameName=cGameName
    .SplashInfoLine="Battlestar Galactica (Williams 1980) v1.02F" & vbNewLine & "VPX table by Allknowing2012" & vbNewLine & "Graphics and Sound Mod by Xenonph"
    .HandleKeyboard=0
    .ShowTitle=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .Hidden=1
    .Run
    'On Error Goto 0
  End With

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=True:PlayMusic"0BG04.mp3":Target2.timerinterval=77000:Target2.timerenabled=1

  vpmNudge.TiltSwitch=swTilt
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingShot,RightSlingShot)

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,9,10,0,0,0,0,0
    bsTrough.InitKick BallRelease,45,6
  bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors),SoundFX("drain",DOFContactors)
    bsTrough.Balls=2

  Set dtL=new cvpmDropTarget
  dtL.InitDrop Array(target1,target2,target3),Array(25,26,27)
  dtL.InitSnd SoundFX("target",DOFContactors),SoundFX("target",DOFContactors)

  Set dtR=new cvpmDropTarget
  dtR.InitDrop Array(target4,target5,target6),Array(28,29,30)
  dtR.InitSnd SoundFX("target",DOFContactors), SoundFX("target",DOFContactors)
  dtR.AllDownSw=31
    dtR.LinkedTo=dtL
    dtL.LinkedTo=dtR

  Set dtC1=new cvpmDropTarget
  dtC1.InitDrop Array(target7,target8),Array(33,34)
  dtC1.InitSnd SoundFX("target",DOFContactors), SoundFX("target",DOFContactors)

  if ballshadows=1 then
        BallShadowUpdate.enabled=1
      else
        BallShadowUpdate.enabled=0
    end if

    if flippershadows=1 then
        FlipperLSh.visible=1
        FlipperRSh.visible=1
       else
        FlipperLSh.visible=0
        FlipperRSh.visible=0
     end if

  Set dtC2=new cvpmDropTarget
  dtC2.InitDrop Array(target9,target10,target11),Array(35,36,37)
  dtC2.InitSnd SoundFX("target",DOFContactors), SoundFX("target",DOFContactors)
  dtC2.AllDownSw=38
    dtC2.LinkedTo=dtC1
    dtC1.LinkedTo=dtC2

  Set bsSaucer=New cvpmBallStack
  bsSaucer.InitSaucer Kicker1,11,165,4
  bsSaucer.InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)
  bsSaucer.KickForceVar=5

  Set bsSaucer2=New cvpmBallStack
  bsSaucer2.InitSaucer Kicker3,12,205,4
  bsSaucer2.InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)
  bsSaucer2.KickForceVar=8.5

    vpmMapLights InsertLights

    GIOff
End Sub

Sub table_Paused:Controller.Pause = 1:End Sub
Sub table_unPaused:Controller.Pause = 0:End Sub

Sub table_KeyDown(ByVal keycode)
  If KeyCode=LeftMagnaSave Then Change2Bump
  If KeyCode=RightMagnaSave Then Change4Bump
  If KeyCode=50 Then Change6Bump
  If KeyCode=49 Then Change8Bump
    If KeyDownHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Pullback:playsound "plungerpull"
End Sub

Sub table_KeyUp(ByVal keycode)

    If KeyCode = 6 Then
      Dim s:StopMusic:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=2000:Target10.timerenabled=1
      s = INT(24 * RND(1) )
      Select Case s
      Case 0:PlaySound"coin":PlaySound("0BGC01")
      Case 1:PlaySound"coin":PlaySound("0BGC02")
      Case 2:PlaySound"coin":PlaySound("0BGC03")
      Case 3:PlaySound"coin":PlaySound("0BGC04")
      Case 4:PlaySound"coin":PlaySound("0BGC05")
      Case 5:PlaySound"coin":PlaySound("0BGC06")
      Case 6:PlaySound"coin":PlaySound("0BGC07")
      Case 7:PlaySound"coin":PlaySound("0BGC08")
      Case 8:PlaySound"coin":PlaySound("0BGC09")
      Case 9:PlaySound"coin":PlaySound("0BGC10")
      Case 10:PlaySound"coin":PlaySound("0BGC11")
      Case 11:PlaySound"coin":PlaySound("0BGC12")
      Case 12:PlaySound"coin":PlaySound("0BGC13")
      Case 13:PlaySound"coin":PlaySound("0BGC14")
      Case 14:PlaySound"coin":PlaySound("0BGC15")
      Case 15:PlaySound"coin":PlaySound("0BGC16")
      Case 16:PlaySound"coin":PlaySound("0BGC17")
      Case 17:PlaySound"coin":PlaySound("0BGC18")
      Case 18:PlaySound"coin":PlaySound("0BGC19")
      Case 19:PlaySound"coin":PlaySound("0BGC20")
      Case 20:PlaySound"coin":PlaySound("0BGC21")
      Case 21:PlaySound"coin":PlaySound("0BGC22")
      Case 22:PlaySound"coin":PlaySound("0BGC23")
      Case 23:PlaySound"coin":PlaySound("0BGC24")
      End Select
            End If

    If KeyCode = 4 Then
      Dim y:StopMusic:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=2000:Target10.timerenabled=1
      y = INT(24 * RND(1) )
      Select Case y
      Case 0:PlaySound"coin":PlaySound("0BGC01")
      Case 1:PlaySound"coin":PlaySound("0BGC02")
      Case 2:PlaySound"coin":PlaySound("0BGC03")
      Case 3:PlaySound"coin":PlaySound("0BGC04")
      Case 4:PlaySound"coin":PlaySound("0BGC05")
      Case 5:PlaySound"coin":PlaySound("0BGC06")
      Case 6:PlaySound"coin":PlaySound("0BGC07")
      Case 7:PlaySound"coin":PlaySound("0BGC08")
      Case 8:PlaySound"coin":PlaySound("0BGC09")
      Case 9:PlaySound"coin":PlaySound("0BGC10")
      Case 10:PlaySound"coin":PlaySound("0BGC11")
      Case 11:PlaySound"coin":PlaySound("0BGC12")
      Case 12:PlaySound"coin":PlaySound("0BGC13")
      Case 13:PlaySound"coin":PlaySound("0BGC14")
      Case 14:PlaySound"coin":PlaySound("0BGC15")
      Case 15:PlaySound"coin":PlaySound("0BGC16")
      Case 16:PlaySound"coin":PlaySound("0BGC17")
      Case 17:PlaySound"coin":PlaySound("0BGC18")
      Case 18:PlaySound"coin":PlaySound("0BGC19")
      Case 19:PlaySound"coin":PlaySound("0BGC20")
      Case 20:PlaySound"coin":PlaySound("0BGC21")
      Case 21:PlaySound"coin":PlaySound("0BGC22")
      Case 22:PlaySound"coin":PlaySound("0BGC23")
      Case 23:PlaySound"coin":PlaySound("0BGC24")
      End Select
            End If

  If KeyCode = 2 Then
      Dim w:StopMusic:StopSound"0BSGA19":Target1.timerenabled=0:Target2.timerenabled=0:Target3.timerenabled=0:Target4.timerenabled=0:Target6.timerenabled=0
      w = INT(5 * RND(1) )
      Select Case w
      Case 0:PlaySound("0BGCC01")
      Case 1:PlaySound("0BGCC02")
      Case 2:PlaySound("0BGCC03")
      Case 3:PlaySound("0BGCC04")
      Case 4:PlaySound("0BGCC05")
      End Select
            End If
  If KeyUpHandler(keycode) Then Exit Sub
  If KeyCode=PlungerKey and BIP=1 Then
              Dim z:StopSounds:StopSound"0BGL00":StopSound"0BGL11":PlaySound"0BGL00":PlaySound"0BGL11":Drain.timerenabled=0:Kicker1.timerenabled=0:Kicker3.timerenabled=0:sw41.timerenabled=0:sw42.timerenabled=0
              z = INT(3 * RND(1) )
              Select Case z
              Case 0:PlayMusic"0BG01.mp3":Drain.timerinterval=57200:Drain.timerenabled=1
              Case 1:PlayMusic"0BG02.mp3":Drain.timerinterval=50000:Drain.timerenabled=1
              Case 2:PlayMusic"0BG03a.mp3":Drain.timerinterval=76300:Drain.timerenabled=1
              End Select
              End If
    If keycode = PlungerKey And BIP=1 Then Plunger.Fire:PlaySoundAtVol "Plunger",plunger,1
    If keycode = PlungerKey And BIP=0 Then Plunger.Fire:PlaySoundAtVol "plungerreleasefree",plunger,1
End Sub

Sub GIOn
  dim bulb
  for each bulb in GILights
  bulb.state = LightStateOn
  next
  Flasher7.visible=1
End Sub

Sub GIOff
  dim bulb
  for each bulb in GILights
  bulb.state = LightStateOff
  next
  Flasher7.visible=0
End Sub

'Set LampCallback=GetRef("UpdateMultipleLamps")

Sub ballreleased_hit:BIP=1:GIOn():End Sub

Sub ballreleased_UnHit():BIP=0:End Sub

Sub Gate3_Hit:Kicker1.timerenabled=0:Drain.timerenabled=0:ff=1:Target6.timerenabled=0
              Dim t:EndMusic:Target7.timerenabled=0:Target8.timerenabled=0
              t = INT(4 * RND(1) )
              Select Case t
              Case 0:PlaySound"0BGL01":Kicker3.timerinterval=27000:Kicker3.timerenabled=1
              Case 1:PlaySound"0BGL02":Kicker3.timerinterval=18000:Kicker3.timerenabled=1
              Case 2:PlaySound"0BGL05":Kicker3.timerinterval=8000:Kicker3.timerenabled=1
              Case 3:PlaySound"0BGL10":Kicker3.timerinterval=10000:Kicker3.timerenabled=1
              End Select

End Sub

Sub Gate2_Hit:FlashLevel6 = 1 : FlasherFlash6_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer:End Sub

Sub Gate1_Hit:FlashLevel6 = 1 : FlasherFlash6_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer:End Sub


'Sub Drain_Hit:bsTrough.AddBall Me:GIOff:End Sub  'switch9 Trough
Sub Drain_Hit:bsTrough.AddBall Me
              If dd=0 then EndMusic:Kicker1.timerenabled=0:Drain.timerenabled=0
              If dd=1 then dd=0
End Sub     'switch9 Thal, don't turn off GI it might be a multiball

Sub Drain_timer
    Dim x
    x = INT(3 * RND(1) )
    Select Case x
    Case 0:PlayMusic"0BG01.mp3":Kicker1.timerinterval=57200:Kicker1.timerenabled=1
    Case 1:PlayMusic"0BG02.mp3":Kicker1.timerinterval=50000:Kicker1.timerenabled=1
    Case 2:PlayMusic"0BG03a.mp3":Kicker1.timerinterval=76300:Kicker1.timerenabled=1
    End Select
Drain.timerenabled=0
End sub

Sub Kicker1_Hit:bsSaucer.AddBall 0
            aa=1
            If aa=1 and bb=0 then StopSounds:Playsound"0BGC05":FlashLevel7 = 1 : FlasherFlash7_Timer:Target7.timerinterval=300:Target7.timerenabled=1
            If aa=1 and bb=1 then cc=1:FlashLevel7 = 1 : FlasherFlash7_Timer:Target7.timerinterval=300:Target7.timerenabled=1:FlashLevel9 = 1 : FlasherFlash9_Timer:Target8.timerinterval=300:Target8.timerenabled=1
            If cc=1 then StopSounds:EndMusic:PlayMusic"0BG03.mp3":Playsound"0BGK02":aa=0:bb=0:cc=0:dd=1:Drain.timerenabled=0:Kicker1.timerenabled=0:Target6.timerinterval=22000:Target6.timerenabled=1

End Sub
Sub Kicker1_UnHit:aa=0:Target7.timerenabled=0:End Sub
Sub Kicker1_timer
    Dim x
    x = INT(3 * RND(1) )
    Select Case x
    Case 0:PlayMusic"0BG01.mp3":Drain.timerinterval=57200:Drain.timerenabled=1
    Case 1:PlayMusic"0BG02.mp3":Drain.timerinterval=50000:Drain.timerenabled=1
    Case 2:PlayMusic"0BG03a.mp3":Drain.timerinterval=76300:Drain.timerenabled=1
    End Select
Kicker1.timerenabled=0
End sub
Sub Kicker3_Hit:bsSaucer2.AddBall 0
            bb=1
            If bb=1 and aa=0 then StopSounds:Playsound"0BGC05":FlashLevel9 = 1 : FlasherFlash9_Timer:Target8.timerinterval=300:Target8.timerenabled=1
            If aa=1 and bb=1 then cc=1:FlashLevel9 = 1 : FlasherFlash9_Timer:Target8.timerinterval=300:Target8.timerenabled=1:FlashLevel7 = 1 : FlasherFlash7_Timer:Target7.timerinterval=300:Target7.timerenabled=1
            If cc=1 then StopSounds:EndMusic:PlayMusic"0BG03.mp3":Playsound"0BGK02":aa=0:bb=0:cc=0:dd=1:Drain.timerenabled=0:Kicker1.timerenabled=0:Target6.timerinterval=22000:Target6.timerenabled=1

End Sub     'switch12
Sub Kicker3_UnHit:bb=0:Target8.timerenabled=0:End Sub

Sub Kicker3_timer
    Dim x
    x = INT(7 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BGL03":sw41.timerinterval=8000:sw41.timerenabled=1
    Case 1:PlaySound"0BGL08":sw41.timerinterval=10000:sw41.timerenabled=1
    Case 2:PlaySound"0BGL06":sw41.timerinterval=10000:sw41.timerenabled=1
    Case 3:PlaySound"0BGL07":sw41.timerinterval=8000:sw41.timerenabled=1
    Case 4:PlaySound"0BGL09":sw41.timerinterval=7000:sw41.timerenabled=1
    Case 5:PlaySound"0BGL12":sw41.timerinterval=8000:sw41.timerenabled=1
    Case 6:PlaySound"0BGL13":sw41.timerinterval=9000:sw41.timerenabled=1
    End Select
Kicker3.timerenabled=0
End sub

Sub sw41_Hit:vpmTimer.PulseSw(41):PlaySound"0BGC05":End Sub     'switch41

Sub sw41_timer
    Dim x
    x = INT(2 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BGL04":sw42.timerinterval=8000:sw42.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=3000:Target10.timerenabled=1
    Case 1:PlaySound"0BGL04":sw42.timerinterval=8000:sw42.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=3000:Target10.timerenabled=1
    End Select
sw41.timerenabled=0
End sub

Sub sw42_Hit:vpmTimer.PulseSw(42):PlaySound"0BGC05":End Sub     'switch42

Sub sw42_timer
    Dim x
    x = INT(4 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BGL01":Kicker3.timerinterval=27000:Kicker3.timerenabled=1
    Case 1:PlaySound"0BGL02":Kicker3.timerinterval=18000:Kicker3.timerenabled=1
    Case 2:PlaySound"0BGL05":Kicker3.timerinterval=8000:Kicker3.timerenabled=1
    Case 3:PlaySound"0BGL10":Kicker3.timerinterval=10000:Kicker3.timerenabled=1
    End Select
sw42.timerenabled=0
End sub

Sub Spinner1_Spin:vpmTimer.PulseSw(32):PlaySoundAtVol "fx_spinner",Spinner1,VolSpin:FlashLevel6 = 1 : FlasherFlash6_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer:End Sub       'switch32 Left Spinner

Sub target1_Hit:dtL.Hit 1:FlashLevel6 = 1 : FlasherFlash6_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer:End Sub             'switch25

Sub target1_timer
    Dim x
    x = INT(5 * RND(1) )
    Select Case x
    Case 0:PlayMusic"0BG04.mp3":Target2.timerinterval=77000:Target2.timerenabled=1
    Case 1:PlayMusic"0BG05.mp3":Target2.timerinterval=79000:Target2.timerenabled=1
    Case 2:PlayMusic"0BG06.mp3":Target2.timerinterval=157000:Target2.timerenabled=1
    Case 3:PlayMusic"0BG07.mp3":Target2.timerinterval=109000:Target2.timerenabled=1
    Case 4:PlayMusic"0BGA14.mp3":Target2.timerinterval=50000:Target2.timerenabled=1
    End Select
target1.timerenabled=0
End sub

Sub target2_Hit:dtL.Hit 2:FlashLevel6 = 1 : FlasherFlash6_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer:End Sub             'switch26

Sub target2_timer
    Dim x
    x = INT(13 * RND(1) )
    Select Case x
    Case 0:PlayMusic"0BGA01.mp3":Target3.timerinterval=22000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=17000:Target10.timerenabled=1
    Case 1:PlayMusic"0BGA02.mp3":Target3.timerinterval=19000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=14000:Target10.timerenabled=1
    Case 2:PlayMusic"0BGA03.mp3":Target3.timerinterval=15000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=10000:Target10.timerenabled=1
    Case 3:PlayMusic"0BGA04.mp3":Target3.timerinterval=18000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=13000:Target10.timerenabled=1
    Case 4:PlayMusic"0BGA05.mp3":Target3.timerinterval=12000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=7000:Target10.timerenabled=1
    Case 5:PlayMusic"0BGA06.mp3":Target3.timerinterval=9000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=5000:Target10.timerenabled=1
    Case 6:PlayMusic"0BGA07.mp3":Target3.timerinterval=17000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=12000:Target10.timerenabled=1
    Case 7:PlayMusic"0BGA08.mp3":Target3.timerinterval=13000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=8000:Target10.timerenabled=1
    Case 8:PlayMusic"0BGA09.mp3":Target3.timerinterval=32000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=27000:Target10.timerenabled=1
    Case 9:PlayMusic"0BGA10.mp3":Target3.timerinterval=17000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=12000:Target10.timerenabled=1
    Case 10:PlayMusic"0BGA11.mp3":Target3.timerinterval=31000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=26000:Target10.timerenabled=1
    Case 11:PlayMusic"0BGA12.mp3":Target3.timerinterval=16000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=11000:Target10.timerenabled=1
    Case 12:PlayMusic"0BGA13.mp3":Target3.timerinterval=17000:Target3.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=12000:Target10.timerenabled=1
    End Select
target2.timerenabled=0
End sub

Sub target3_Hit:dtL.Hit 3:FlashLevel6 = 1 : FlasherFlash6_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer:End Sub             'switch27

Sub target3_timer
    Dim x
    x = INT(4 * RND(1) )
    Select Case x
    Case 0:PlayMusic"0BG08.mp3":Target4.timerinterval=34000:Target4.timerenabled=1
    Case 1:PlayMusic"0BG09.mp3":Target4.timerinterval=53000:Target4.timerenabled=1
    Case 2:PlayMusic"0BG10.mp3":Target4.timerinterval=49000:Target4.timerenabled=1
    Case 3:PlayMusic"0BG11.mp3":Target4.timerinterval=70000:Target4.timerenabled=1
    End Select
target3.timerenabled=0
End sub

Sub target4_Hit:dtR.Hit 1:FlashLevel6 = 1 : FlasherFlash6_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer:End Sub             'switch28

Sub target4_timer
    Dim x
    x = INT(13 * RND(1) )
    Select Case x
    Case 0:PlayMusic"0BGA01.mp3":Target1.timerinterval=22000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=17000:Target10.timerenabled=1
    Case 1:PlayMusic"0BGA02.mp3":Target1.timerinterval=19000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=14000:Target10.timerenabled=1
    Case 2:PlayMusic"0BGA03.mp3":Target1.timerinterval=15000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=10000:Target10.timerenabled=1
    Case 3:PlayMusic"0BGA04.mp3":Target1.timerinterval=18000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=13000:Target10.timerenabled=1
    Case 4:PlayMusic"0BGA05.mp3":Target1.timerinterval=12000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=7000:Target10.timerenabled=1
    Case 5:PlayMusic"0BGA06.mp3":Target1.timerinterval=9000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=5000:Target10.timerenabled=1
    Case 6:PlayMusic"0BGA07.mp3":Target1.timerinterval=17000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=12000:Target10.timerenabled=1
    Case 7:PlayMusic"0BGA08.mp3":Target1.timerinterval=13000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=8000:Target10.timerenabled=1
    Case 8:PlayMusic"0BGA09.mp3":Target1.timerinterval=32000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=27000:Target10.timerenabled=1
    Case 9:PlayMusic"0BGA10.mp3":Target1.timerinterval=17000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=12000:Target10.timerenabled=1
    Case 10:PlayMusic"0BGA11.mp3":Target1.timerinterval=31000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=26000:Target10.timerenabled=1
    Case 11:PlayMusic"0BGA12.mp3":Target1.timerinterval=16000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=11000:Target10.timerenabled=1
    Case 12:PlayMusic"0BGA13.mp3":Target1.timerinterval=17000:Target1.timerenabled=1:Target9.timerinterval=300:Target9.timerenabled=1:Target10.timerinterval=12000:Target10.timerenabled=1
    End Select
target4.timerenabled=0
End sub

Sub target5_Hit:dtR.Hit 2:FlashLevel6 = 1 : FlasherFlash6_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer:End Sub             'switch29

Sub target5_timer
      ff=1
target5.timerenabled=0
End sub

Sub target6_Hit:dtR.Hit 3:FlashLevel6 = 1 : FlasherFlash6_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer:End Sub             'switch30

Sub target6_timer
    Dim x
    x = INT(3 * RND(1) )
    Select Case x
    Case 0:PlayMusic"0BG03.mp3":Drain.timerinterval=72500:Drain.timerenabled=1
    Case 1:PlayMusic"0BG03.mp3":Drain.timerinterval=72500:Drain.timerenabled=1
    Case 2:PlayMusic"0BG03.mp3":Drain.timerinterval=72500:Drain.timerenabled=1
    End Select
target6.timerenabled=0
End sub

Sub target7_Hit:FlashLevel10 = 1 : FlasherFlash10_Timer:dtC1.Hit 1:End Sub              'switch33

Sub target7_timer
    FlashLevel7 = 1 : FlasherFlash7_Timer
End sub

Sub target8_Hit:FlashLevel10 = 1 : FlasherFlash10_Timer:dtC1.Hit 2:End Sub              'switch34

Sub target8_timer
    FlashLevel9 = 1 : FlasherFlash9_Timer
End sub


Sub target9_Hit:FlashLevel10 = 1 : FlasherFlash10_Timer:dtC2.Hit 1:End Sub              'switch35

Sub target9_timer
    Dim x
    x = INT(7 * RND(1) )
    Select Case x
    Case 0:FlashLevel6 = 1 : FlasherFlash6_Timer
    Case 1:FlashLevel7 = 1 : FlasherFlash7_Timer
    Case 2:FlashLevel8 = 1 : FlasherFlash8_Timer
    Case 3:FlashLevel9 = 1 : FlasherFlash9_Timer
    Case 4:FlashLevel6 = 1 : FlasherFlash6_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer
    Case 5:FlashLevel7 = 1 : FlasherFlash7_Timer:FlashLevel9 = 1 : FlasherFlash9_Timer
    Case 6:FlashLevel10 = 1 : FlasherFlash10_Timer
    End Select

End sub

Sub target10_Hit:FlashLevel10 = 1 : FlasherFlash10_Timer:dtC2.Hit 2:End Sub             'switch36

Sub target10_timer
    target9.timerenabled=0
target10.timerenabled=0
End sub

Sub target11_Hit:FlashLevel10 = 1 : FlasherFlash10_Timer:dtC2.Hit 3:End Sub             'switch37

'***********************************************
'***********************************************
          'Switches
'***********************************************
'***********************************************


Sub sw14_Hit:Controller.Switch(14)=1

    Dim x:StopSounds
    x = INT(11 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BSG35"
    Case 1:PlaySound"0BSG36"
    Case 2:PlaySound"0BSG37"
    Case 3:PlaySound"0BSG37a"
    Case 4:PlaySound"0BSG37b"
    Case 5:PlaySound"0BSG37c"
    Case 6:PlaySound"0BSG37d"
    Case 7:PlaySound"0BSG37e"
    Case 8:PlaySound"0BSG37f"
    Case 9:PlaySound"0BSG37g"
    Case 10:PlaySound"0BSG37h"
    End Select

End Sub
Sub sw14_unHit:Controller.Switch(14)=0:End Sub
Sub sw18_Hit:Controller.Switch(18)=1
    If ee=1 And ff=1 then
    Dim x:StopSounds:ff=0
    x = INT(43 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BSG18":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 1:PlaySound"0BSG19":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 2:PlaySound"0BSG20":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 3:PlaySound"0BSG21":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 4:PlaySound"0BSG22":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 5:PlaySound"0BSG23":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 6:PlaySound"0BSG24":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 7:PlaySound"0BSG25":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 8:PlaySound"0BSG26":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 9:PlaySound"0BSG27":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 10:PlaySound"0BSG28":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 11:PlaySound"0BSG29":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 12:PlaySound"0BSG30":Target5.timerinterval=2000:Target5.timerenabled=1
    Case 13:PlaySound"0BSG31":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 14:PlaySound"0BSG32":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 15:PlaySound"0BSG33":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 16:PlaySound"0BSG34":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 17:PlaySound"0BSG38":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 18:PlaySound"0BSG39":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 19:PlaySound"0BSG40":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 20:PlaySound"0BSG41":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 21:PlaySound"0BSG42":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 22:PlaySound"0BSG43":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 23:PlaySound"0BSG44":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 24:PlaySound"0BSG45":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 25:PlaySound"0BSG46":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 26:PlaySound"0BSG47":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 27:PlaySound"0BSG48":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 28:PlaySound"0BSG49":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 29:PlaySound"0BSG50":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 30:PlaySound"0BSG51":Target5.timerinterval=6500:Target5.timerenabled=1
    Case 31:PlaySound"0BSG52":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 32:PlaySound"0BSG53":Target5.timerinterval=7000:Target5.timerenabled=1
    Case 33:PlaySound"0BSG54":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 34:PlaySound"0BSG55":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 35:PlaySound"0BSG56":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 36:PlaySound"0BSG57":Target5.timerinterval=4200:Target5.timerenabled=1
    Case 37:PlaySound"0BSG58":Target5.timerinterval=4100:Target5.timerenabled=1
    Case 38:PlaySound"0BSG59":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 39:PlaySound"0BSG60":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 40:PlaySound"0BSG61":Target5.timerinterval=4700:Target5.timerenabled=1
    Case 41:PlaySound"0BSG62":Target5.timerinterval=3800:Target5.timerenabled=1
    Case 42:PlaySound"0BSG63":Target5.timerinterval=2500:Target5.timerenabled=1
    End Select
    End If

End Sub
Sub sw18_unHit:Controller.Switch(18)=0:End Sub
Sub sw19_Hit:Controller.Switch(19)=1
    If ee=1 And ff=1 then
    Dim x:StopSounds:ff=0
    x = INT(43 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BSG19":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 1:PlaySound"0BSG18":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 2:PlaySound"0BSG20":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 3:PlaySound"0BSG21":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 4:PlaySound"0BSG22":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 5:PlaySound"0BSG23":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 6:PlaySound"0BSG24":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 7:PlaySound"0BSG25":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 8:PlaySound"0BSG26":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 9:PlaySound"0BSG27":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 10:PlaySound"0BSG28":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 11:PlaySound"0BSG29":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 12:PlaySound"0BSG30":Target5.timerinterval=2000:Target5.timerenabled=1
    Case 13:PlaySound"0BSG31":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 14:PlaySound"0BSG32":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 15:PlaySound"0BSG33":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 16:PlaySound"0BSG34":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 17:PlaySound"0BSG38":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 18:PlaySound"0BSG39":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 19:PlaySound"0BSG40":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 20:PlaySound"0BSG41":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 21:PlaySound"0BSG42":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 22:PlaySound"0BSG43":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 23:PlaySound"0BSG44":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 24:PlaySound"0BSG45":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 25:PlaySound"0BSG46":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 26:PlaySound"0BSG47":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 27:PlaySound"0BSG48":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 28:PlaySound"0BSG49":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 29:PlaySound"0BSG50":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 30:PlaySound"0BSG51":Target5.timerinterval=6500:Target5.timerenabled=1
    Case 31:PlaySound"0BSG52":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 32:PlaySound"0BSG53":Target5.timerinterval=7000:Target5.timerenabled=1
    Case 33:PlaySound"0BSG54":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 34:PlaySound"0BSG56":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 35:PlaySound"0BSG55":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 36:PlaySound"0BSG57":Target5.timerinterval=4200:Target5.timerenabled=1
    Case 37:PlaySound"0BSG58":Target5.timerinterval=4100:Target5.timerenabled=1
    Case 38:PlaySound"0BSG59":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 39:PlaySound"0BSG60":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 40:PlaySound"0BSG61":Target5.timerinterval=4700:Target5.timerenabled=1
    Case 41:PlaySound"0BSG62":Target5.timerinterval=3800:Target5.timerenabled=1
    Case 42:PlaySound"0BSG63":Target5.timerinterval=2500:Target5.timerenabled=1
    End Select
    End If
End Sub
Sub sw19_unHit:Controller.Switch(19)=0:End Sub

Sub sw20_Hit:Controller.Switch(20)=1
    If ee=1 And ff=1 then
    Dim x:StopSounds:ff=0
    x = INT(43 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BSG20":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 1:PlaySound"0BSG19":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 2:PlaySound"0BSG18":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 3:PlaySound"0BSG21":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 4:PlaySound"0BSG22":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 5:PlaySound"0BSG23":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 6:PlaySound"0BSG24":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 7:PlaySound"0BSG25":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 8:PlaySound"0BSG26":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 9:PlaySound"0BSG27":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 10:PlaySound"0BSG28":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 11:PlaySound"0BSG29":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 12:PlaySound"0BSG30":Target5.timerinterval=2000:Target5.timerenabled=1
    Case 13:PlaySound"0BSG31":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 14:PlaySound"0BSG32":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 15:PlaySound"0BSG33":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 16:PlaySound"0BSG34":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 17:PlaySound"0BSG38":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 18:PlaySound"0BSG39":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 19:PlaySound"0BSG40":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 20:PlaySound"0BSG41":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 21:PlaySound"0BSG42":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 22:PlaySound"0BSG43":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 23:PlaySound"0BSG44":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 24:PlaySound"0BSG45":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 25:PlaySound"0BSG46":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 26:PlaySound"0BSG47":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 27:PlaySound"0BSG48":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 28:PlaySound"0BSG49":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 29:PlaySound"0BSG50":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 30:PlaySound"0BSG51":Target5.timerinterval=6500:Target5.timerenabled=1
    Case 31:PlaySound"0BSG52":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 32:PlaySound"0BSG53":Target5.timerinterval=7000:Target5.timerenabled=1
    Case 33:PlaySound"0BSG56":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 34:PlaySound"0BSG55":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 35:PlaySound"0BSG54":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 36:PlaySound"0BSG57":Target5.timerinterval=4200:Target5.timerenabled=1
    Case 37:PlaySound"0BSG58":Target5.timerinterval=4100:Target5.timerenabled=1
    Case 38:PlaySound"0BSG59":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 39:PlaySound"0BSG60":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 40:PlaySound"0BSG61":Target5.timerinterval=4700:Target5.timerenabled=1
    Case 41:PlaySound"0BSG62":Target5.timerinterval=3800:Target5.timerenabled=1
    Case 42:PlaySound"0BSG63":Target5.timerinterval=2500:Target5.timerenabled=1
    End Select
    End If
End Sub
Sub sw20_unHit:Controller.Switch(20)=0:End Sub
Sub sw21_Hit:Controller.Switch(21)=1
    If ee=1 And ff=1 then
    Dim x:StopSounds:ff=0
    x = INT(43 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BSG18":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 1:PlaySound"0BSG19":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 2:PlaySound"0BSG20":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 3:PlaySound"0BSG21":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 4:PlaySound"0BSG22":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 5:PlaySound"0BSG23":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 6:PlaySound"0BSG24":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 7:PlaySound"0BSG25":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 8:PlaySound"0BSG26":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 9:PlaySound"0BSG27":Target5.timerinterval=3000:Target5.timerenabled=1
    Case 10:PlaySound"0BSG28":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 11:PlaySound"0BSG29":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 12:PlaySound"0BSG30":Target5.timerinterval=2000:Target5.timerenabled=1
    Case 13:PlaySound"0BSG31":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 14:PlaySound"0BSG32":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 15:PlaySound"0BSG33":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 16:PlaySound"0BSG34":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 17:PlaySound"0BSG38":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 18:PlaySound"0BSG39":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 19:PlaySound"0BSG40":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 20:PlaySound"0BSG41":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 21:PlaySound"0BSG42":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 22:PlaySound"0BSG43":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 23:PlaySound"0BSG44":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 24:PlaySound"0BSG45":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 25:PlaySound"0BSG46":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 26:PlaySound"0BSG47":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 27:PlaySound"0BSG48":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 28:PlaySound"0BSG49":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 29:PlaySound"0BSG50":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 30:PlaySound"0BSG51":Target5.timerinterval=6500:Target5.timerenabled=1
    Case 31:PlaySound"0BSG52":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 32:PlaySound"0BSG53":Target5.timerinterval=7000:Target5.timerenabled=1
    Case 33:PlaySound"0BSG54":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 34:PlaySound"0BSG55":Target5.timerinterval=6000:Target5.timerenabled=1
    Case 35:PlaySound"0BSG56":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 36:PlaySound"0BSG57":Target5.timerinterval=4200:Target5.timerenabled=1
    Case 37:PlaySound"0BSG58":Target5.timerinterval=4100:Target5.timerenabled=1
    Case 38:PlaySound"0BSG59":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 39:PlaySound"0BSG60":Target5.timerinterval=5000:Target5.timerenabled=1
    Case 40:PlaySound"0BSG61":Target5.timerinterval=4700:Target5.timerenabled=1
    Case 41:PlaySound"0BSG62":Target5.timerinterval=3800:Target5.timerenabled=1
    Case 42:PlaySound"0BSG63":Target5.timerinterval=2500:Target5.timerenabled=1
    End Select
    End If
End Sub
Sub sw21_unHit:Controller.Switch(21)=0:End Sub
Sub sw15_Hit:Controller.Switch(15)=1

    Dim x:StopSounds
    x = INT(11 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BSG35"
    Case 1:PlaySound"0BSG36"
    Case 2:PlaySound"0BSG37"
    Case 3:PlaySound"0BSG37a"
    Case 4:PlaySound"0BSG37b"
    Case 5:PlaySound"0BSG37c"
    Case 6:PlaySound"0BSG37d"
    Case 7:PlaySound"0BSG37e"
    Case 8:PlaySound"0BSG37f"
    Case 9:PlaySound"0BSG37g"
    Case 10:PlaySound"0BSG37h"
    End Select

End Sub
Sub sw15_unHit:Controller.Switch(15)=0:End Sub

Sub sw46_Hit:Controller.Switch(46)=1
    If ff=1 then
    Dim x:StopSounds:ff=0
    x = INT(20 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BSG01":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 1:PlaySound"0BSG02":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 2:PlaySound"0BSG03":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 3:PlaySound"0BSG04":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 4:PlaySound"0BSG05":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 5:PlaySound"0BSG06":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 6:PlaySound"0BSG07":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 7:PlaySound"0BSG08":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 8:PlaySound"0BSG09":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 9:PlaySound"0BSG10":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 10:PlaySound"0BSG11":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 11:PlaySound"0BSG12":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 12:PlaySound"0BSG13":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 13:PlaySound"0BSG14":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 14:PlaySound"0BSG15":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 15:PlaySound"0BSG16":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 16:PlaySound"0BSG17":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 17:PlaySound"0BSG64":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 18:PlaySound"0BSG65":Target5.timerinterval=5400:Target5.timerenabled=1
    Case 19:PlaySound"0BSG66":Target5.timerinterval=2500:Target5.timerenabled=1
    End Select
    End If
End Sub
Sub sw46_unHit:Controller.Switch(46)=0:End Sub
Sub sw47_Hit:Controller.Switch(47)=1
    If ff=1 then
    Dim x:StopSounds:ff=0
    x = INT(20 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BSG02":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 1:PlaySound"0BSG01":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 2:PlaySound"0BSG03":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 3:PlaySound"0BSG04":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 4:PlaySound"0BSG05":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 5:PlaySound"0BSG06":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 6:PlaySound"0BSG07":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 7:PlaySound"0BSG08":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 8:PlaySound"0BSG09":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 9:PlaySound"0BSG10":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 10:PlaySound"0BSG11":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 11:PlaySound"0BSG12":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 12:PlaySound"0BSG13":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 13:PlaySound"0BSG14":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 14:PlaySound"0BSG15":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 15:PlaySound"0BSG16":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 16:PlaySound"0BSG17":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 17:PlaySound"0BSG64":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 18:PlaySound"0BSG65":Target5.timerinterval=5400:Target5.timerenabled=1
    Case 19:PlaySound"0BSG66":Target5.timerinterval=2500:Target5.timerenabled=1
    End Select
    End If
End Sub
Sub sw47_unHit:Controller.Switch(47)=0:End Sub
Sub sw48_Hit:Controller.Switch(48)=1
    If ff=1 then
    Dim x:StopSounds:ff=0
    x = INT(20 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BSG03":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 1:PlaySound"0BSG02":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 2:PlaySound"0BSG01":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 3:PlaySound"0BSG04":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 4:PlaySound"0BSG05":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 5:PlaySound"0BSG06":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 6:PlaySound"0BSG07":Target5.timerinterval=3500:Target5.timerenabled=1
    Case 7:PlaySound"0BSG08":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 8:PlaySound"0BSG09":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 9:PlaySound"0BSG10":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 10:PlaySound"0BSG11":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 11:PlaySound"0BSG12":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 12:PlaySound"0BSG13":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 13:PlaySound"0BSG14":Target5.timerinterval=4000:Target5.timerenabled=1
    Case 14:PlaySound"0BSG15":Target5.timerinterval=2500:Target5.timerenabled=1
    Case 15:PlaySound"0BSG16":Target5.timerinterval=4500:Target5.timerenabled=1
    Case 16:PlaySound"0BSG17":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 17:PlaySound"0BSG64":Target5.timerinterval=5500:Target5.timerenabled=1
    Case 18:PlaySound"0BSG65":Target5.timerinterval=5400:Target5.timerenabled=1
    Case 19:PlaySound"0BSG66":Target5.timerinterval=2500:Target5.timerenabled=1
    End Select
    End If
End Sub
Sub sw48_unHit:Controller.Switch(48)=0:End Sub

Sub sw49_Hit:Controller.Switch(49)=1:PlaySound"0BGC05":End Sub
Sub sw49_unHit:Controller.Switch(49)=0:End Sub

Sub sw43_Hit:vpmTimer.PulseSw(43):End Sub
Sub sw44_Hit:vpmTimer.PulseSw(44):End Sub
Sub sw45_Hit:vpmTimer.PulseSw(45):End Sub


Dim bump1,bump2,bump3

Sub Bumper1_Hit:vpmTimer.PulseSw 24:bump1 = 1:Me.TimerEnabled = 1

    Dim x:FlashLevel1 = 1 : FlasherFlash1_Timer
    x = INT(5 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BGB01"
    Case 1:PlaySound"0BGB02"
    Case 2:PlaySound"0BGB03"
    Case 3:PlaySound"0BGB04"
    Case 4:PlaySound"0BGB05"
    End Select

End Sub
Sub Bumper1_Timer()
  Select Case bump1
        Case 1:Ring1.Z = -30:bump1 = 2
        Case 2:Ring1.Z = -20:bump1 = 3
        Case 3:Ring1.Z = -10:bump1 = 4
        Case 4:Ring1.Z = 0:Me.TimerEnabled = 0
  End Select
End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 22:bump2 = 1:Me.TimerEnabled = 1

    Dim x:FlashLevel2 = 1 : FlasherFlash2_Timer
    x = INT(5 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BGB01"
    Case 1:PlaySound"0BGB02"
    Case 2:PlaySound"0BGB03"
    Case 3:PlaySound"0BGB04"
    Case 4:PlaySound"0BGB05"
    End Select

End Sub
Sub Bumper2_Timer()
  Select Case bump2
        Case 1:Ring2.Z = -30:bump2 = 2
        Case 2:Ring2.Z = -20:bump2 = 3
        Case 3:Ring2.Z = -10:bump2 = 4
        Case 4:Ring2.Z = 0:Me.TimerEnabled = 0
  End Select
End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 23:bump3 = 1:Me.TimerEnabled = 1

    Dim x:FlashLevel3 = 1 : FlasherFlash3_Timer
    x = INT(5 * RND(1) )
    Select Case x
    Case 0:PlaySound"0BGB01"
    Case 1:PlaySound"0BGB02"
    Case 2:PlaySound"0BGB03"
    Case 3:PlaySound"0BGB04"
    Case 4:PlaySound"0BGB05"
    End Select

End Sub
Sub Bumper3_Timer()
  Select Case bump3
        Case 1:Ring3.Z = -30:bump3 = 2
        Case 2:Ring3.Z = -20:bump3 = 3
        Case 3:Ring3.Z = -10:bump3 = 4
        Case 4:Ring3.Z = 0:Me.TimerEnabled = 0
  End Select
End Sub


Sub sw50_Hit :vpmTimer.PulseSw 50:End Sub
Sub sw51_Hit :vpmTimer.PulseSw 51:End Sub
Sub sw52_Hit :vpmTimer.PulseSw 52:End Sub
Sub sw53_Hit :vpmTimer.PulseSw 53:End Sub
Sub sw54_Hit :vpmTimer.PulseSw 54:End Sub
Sub sw55_Hit :vpmTimer.PulseSw 55:End Sub
Sub sw56_Hit :vpmTimer.PulseSw 56:End Sub
Sub sw57_Hit :vpmTimer.PulseSw 57:End Sub

 Dim dBall, dZpos

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(28)

' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)

' 2nd Player
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)

' 3rd Player
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)

' 4th Player
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)

' Credits
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)

' Balls
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 28) then
        For Each obj In Digits(num)
          If chg And 1 Then obj.State = stat And 1
          chg = chg\2 : stat = stat\2
        Next
      else

      end if
    next
    end if
  end if
End Sub

'*****************************************
'     BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

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
        If BOT(b).X < Table.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Table.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Table.Width/2))/17))' - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

Sub Pins_Hit (idx)
   debug.prnit "pin"
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound SoundFX("target",DOFContactors), 0, Vol(ActiveBall)*VolTarg, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub TargetBankWalls_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
    debug.print "metals_medium"
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
   debug.print "rubber"
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
    debug.print "Posts"
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub Rubbers_Hit(idx)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub Bumpers_Hit(idx)
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySound SoundFx("fx_bumper1",DOFContactors), 0, Vol(ActiveBall)*VolBump, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound SoundFx("fx_bumper2",DOFContactors), 0, Vol(ActiveBall)*VolBump, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound SoundFx("fx_bumper3",DOFContactors), 0, Vol(ActiveBall)*VolBump, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 4 : PlaySound SoundFx("fx_bumper4",DOFContactors), 0, Vol(ActiveBall)*VolBump, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table.width-1
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

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep


Sub LeftSlingShot_Slingshot
    PlaySoundAtVol Soundfx("left_slingshot",DOFContactors), sling2, 1
    FlashLevel4 = 1 : FlasherFlash4_Timer
    vpmTimer.PulseSw 39
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

Sub RightSlingShot_Slingshot
    PlaySoundAtVol Soundfx("right_slingshot",DOFContactors), sling1, 1
    FlashLevel5 = 1 : FlasherFlash5_Timer
    vpmTimer.PulseSw 40
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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / table.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / table.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table.width-1
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

Const tnob = 3 ' total number of balls
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

Sub GraphicsTimer_Timer()
  batleft.objrotz = LeftFlipper.CurrentAngle + 1
  batright.objrotz = RightFlipper.CurrentAngle - 1
    batleft1.objrotz = LeftFlipperUpper.CurrentAngle + 1
  batright1.objrotz = RightFlipperUpper.CurrentAngle - 1

  if FlipperShadows=1 then
    FlipperLSh.RotZ = batleft.objrotz
    FlipperLSh1.RotZ = batleft1.objrotz
    FlipperRSh.RotZ = batright.objrotz
    FlipperRSh1.RotZ = batright1.objrotz
  end if

End Sub
' Thalamus : Exit in a clean and proper way
Sub Table_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

