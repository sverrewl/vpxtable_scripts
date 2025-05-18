' Panthera - Gottlieb 1980
' IPD No. 1745 / June, 1980 / 4 Players
' VPX - version by JPSalas 2017, version 1.0.0

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "sys80.VBS", 3.37

'Variables
Dim bsTrough, dtL, dtR, dtT, bsLSaucer, x

Dim PeriodCorrect

'Time Period Option per BorgDogs request.
'******************************
   PeriodCorrect=0  'Change to 1 to remove voice sound effects and background music.
'******************************

Const cGameName = "panther7" '7 digits
'Const cGameName = "panthera"

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

Dim VarHidden
If Table1.ShowDT = True then
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
Else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next

    lrail.Visible = 0
    rrail.Visible = 0
    Flasherlight25.Visible = 0
    Flasherlight26.Visible = 0
    Flasherlight27.Visible = 0
    Flasherlight28.Visible = 0
    Flasherlight29.Visible = 0
    Flasherlight30.Visible = 0
    Flasherlight31.Visible = 0
End If

if B2SOn = true then VarHidden = 1

'*************
'   JP'S LUT
'*************

Dim LUTImage
Dim xa
Sub LoadLUT
    xa = LoadValue(cGameName, "LUTImage")
    If(xa <> "") Then LUTImage = xa Else LUTImage = 0
  UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT: LUTImage = (LUTImage +1 ) MOD 13: UpdateLUT: SaveLUT: End Sub

Sub UpdateLUT
Select Case LutImage
Case 0: Table1.ColorGradeImage = "LUT0"
Case 1: Table1.ColorGradeImage = "LUT1"
Case 2: Table1.ColorGradeImage = "LUT2"
Case 3: Table1.ColorGradeImage = "LUT3"
Case 4: Table1.ColorGradeImage = "LUT4"
Case 5: Table1.ColorGradeImage = "LUT5"
Case 6: Table1.ColorGradeImage = "LUT6"
Case 7: Table1.ColorGradeImage = "LUT7"
Case 8: Table1.ColorGradeImage = "LUT8"
Case 9: Table1.ColorGradeImage = "LUT9"
Case 10: Table1.ColorGradeImage = "LUT10"
Case 11: Table1.ColorGradeImage = "LUT11"
Case 12: Table1.ColorGradeImage = "LUT12"
End Select
End Sub

'**********************************************************************************************************
'Fluppers Flasher Script
Dim FlashLevel1, FlashLevel2, FlashLevel3, FlashLevel4, FlashLevel5, FlashLevel6, FlashLevel7, FlashLevel8, FlashLevel9, FlashLevel10
Dim FlashLevel11, FlashLevel12, FlashLevel13, FlashLevel14, FlashLevel24, FlashLevel25
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


'*** blue flasher ***
Sub FlasherFlash1_Timer()
  dim flashx3, matdim
  If not Flasherflash1.TimerEnabled Then
    Flasherflash1.TimerEnabled = True
    Flasherflash1.visible = 1
    Flasherlit1.visible = 1
  End If
  flashx3 = FlashLevel1 * FlashLevel1 * FlashLevel1
  Flasherflash1.opacity = 8000 * flashx3
  Flasherlit1.BlendDisableLighting = 10 * flashx3
  Flasherbase1.BlendDisableLighting =  flashx3
  Flasherlight1.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel1)
  Flasherlit1.material = "domelit" & matdim
  FlashLevel1 = FlashLevel1 * 0.85 - 0.01
  If FlashLevel1 < 0.15 Then
    Flasherlit1.visible = 0
  Else
    Flasherlit1.visible = 1
  end If
  If FlashLevel1 < 0 Then
    Flasherflash1.TimerEnabled = False
    Flasherflash1.visible = 0
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


'*** Red flasher ***
Sub FlasherFlash4_Timer()
  dim flashx3, matdim
  If not Flasherflash4.TimerEnabled Then
    Flasherflash4.TimerEnabled = True
    Flasherflash4.visible = 1
    Flasherlit4.visible = 1
  End If
  flashx3 = FlashLevel4 * FlashLevel4 * FlashLevel4
  Flasherflash4.opacity = 1500 * flashx3
  Flasherlit4.BlendDisableLighting = 10 * flashx3
  Flasherbase4.BlendDisableLighting =  flashx3
  Flasherlight4.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel4)
  Flasherlit4.material = "domelit" & matdim
  FlashLevel4 = FlashLevel4 * 0.9 - 0.01
  If FlashLevel4 < 0.15 Then
    Flasherlit4.visible = 0
  Else
    Flasherlit4.visible = 1
  end If
  If FlashLevel4 < 0 Then
    Flasherflash4.TimerEnabled = False
    Flasherflash4.visible = 0
  End If
End Sub


'*** blue flasher ***
Sub FlasherFlash5_Timer()
  dim flashx3, matdim
  If not Flasherflash5.TimerEnabled Then
    Flasherflash5.TimerEnabled = True
    Flasherflash5.visible = 1
    Flasherlit5.visible = 1
  End If
  flashx3 = FlashLevel5 * FlashLevel5 * FlashLevel5
  Flasherflash5.opacity = 8000 * flashx3
  Flasherlit5.BlendDisableLighting = 10 * flashx3
  Flasherbase5.BlendDisableLighting =  flashx3
  Flasherlight5.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel5)
  Flasherlit5.material = "domelit" & matdim
  FlashLevel5 = FlashLevel5 * 0.85 - 0.01
  If FlashLevel5 < 0.15 Then
    Flasherlit5.visible = 0
  Else
    Flasherlit5.visible = 1
  end If
  If FlashLevel5 < 0 Then
    Flasherflash5.TimerEnabled = False
    Flasherflash5.visible = 0
  End If
End Sub


'*** blue flasher ***
Sub FlasherFlash6_Timer()
  dim flashx3, matdim
  If not Flasherflash6.TimerEnabled Then
    Flasherflash6.TimerEnabled = True
    Flasherflash6.visible = 1
    Flasherlit6.visible = 1
  End If
  flashx3 = FlashLevel6 * FlashLevel6 * FlashLevel6
  Flasherflash6.opacity = 8000 * flashx3
  Flasherlit6.BlendDisableLighting = 10 * flashx3
  Flasherbase6.BlendDisableLighting =  flashx3
  Flasherlight6.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel6)
  Flasherlit6.material = "domelit" & matdim
  FlashLevel6 = FlashLevel6 * 0.85 - 0.01
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


'*** blue flasher ***
Sub FlasherFlash8_Timer()
  dim flashx3, matdim
  If not Flasherflash8.TimerEnabled Then
    Flasherflash8.TimerEnabled = True
    Flasherflash8.visible = 1
    Flasherlit8.visible = 1
  End If
  flashx3 = FlashLevel8 * FlashLevel8 * FlashLevel8
  Flasherflash8.opacity = 8000 * flashx3
  Flasherlit8.BlendDisableLighting = 10 * flashx3
  Flasherbase8.BlendDisableLighting =  flashx3
  Flasherlight8.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel8)
  Flasherlit8.material = "domelit" & matdim
  FlashLevel8 = FlashLevel8 * 0.85 - 0.01
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


'*** blue flasher ***
Sub FlasherFlash9_Timer()
  dim flashx3, matdim
  If not Flasherflash9.TimerEnabled Then
    Flasherflash9.TimerEnabled = True
    Flasherflash9.visible = 1
    Flasherlit9.visible = 1
  End If
  flashx3 = FlashLevel9 * FlashLevel9 * FlashLevel9
  Flasherflash9.opacity = 8000 * flashx3
  Flasherlit9.BlendDisableLighting = 10 * flashx3
  Flasherbase9.BlendDisableLighting =  flashx3
  Flasherlight9.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel9)
  Flasherlit9.material = "domelit" & matdim
  FlashLevel9 = FlashLevel9 * 0.85 - 0.01
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


'*** blue flasher ***
Sub FlasherFlash10_Timer()
  dim flashx3, matdim
  If not Flasherflash10.TimerEnabled Then
    Flasherflash10.TimerEnabled = True
    Flasherflash10.visible = 1
    Flasherlit10.visible = 1
  End If
  flashx3 = FlashLevel10 * FlashLevel10 * FlashLevel10
  Flasherflash10.opacity = 8000 * flashx3
  Flasherlit10.BlendDisableLighting = 10 * flashx3
  Flasherbase10.BlendDisableLighting =  flashx3
  Flasherlight10.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel10)
  Flasherlit10.material = "domelit" & matdim
  FlashLevel10 = FlashLevel10 * 0.85 - 0.01
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


'*** Red flasher ***
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


'*** Red flasher ***
Sub FlasherFlash12_Timer()
  dim flashx3, matdim
  If not Flasherflash12.TimerEnabled Then
    Flasherflash12.TimerEnabled = True
    Flasherflash12.visible = 1
    Flasherlit12.visible = 1
  End If
  flashx3 = FlashLevel12 * FlashLevel12 * FlashLevel12
  Flasherflash12.opacity = 1500 * flashx3
  Flasherlit12.BlendDisableLighting = 10 * flashx3
  Flasherbase12.BlendDisableLighting =  flashx3
  Flasherlight12.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel12)
  Flasherlit12.material = "domelit" & matdim
  FlashLevel12 = FlashLevel12 * 0.9 - 0.01
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


'*** Red flasher ***
Sub FlasherFlash13_Timer()
  dim flashx3, matdim
  If not Flasherflash13.TimerEnabled Then
    Flasherflash13.TimerEnabled = True
    Flasherflash13.visible = 1
    Flasherlit13.visible = 1
  End If
  flashx3 = FlashLevel13 * FlashLevel13 * FlashLevel13
  Flasherflash13.opacity = 1500 * flashx3
  Flasherlit13.BlendDisableLighting = 10 * flashx3
  Flasherbase13.BlendDisableLighting =  flashx3
  Flasherlight13.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel13)
  Flasherlit13.material = "domelit" & matdim
  FlashLevel13 = FlashLevel13 * 0.9 - 0.01
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


'*** Red flasher ***
Sub FlasherFlash14_Timer()
  dim flashx3, matdim
  If not Flasherflash14.TimerEnabled Then
    Flasherflash14.TimerEnabled = True
    Flasherflash14.visible = 1
    Flasherlit14.visible = 1
  End If
  flashx3 = FlashLevel14 * FlashLevel14 * FlashLevel14
  Flasherflash14.opacity = 1500 * flashx3
  Flasherlit14.BlendDisableLighting = 10 * flashx3
  Flasherbase14.BlendDisableLighting =  flashx3
  Flasherlight14.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel14)
  Flasherlit14.material = "domelit" & matdim
  FlashLevel14 = FlashLevel14 * 0.9 - 0.01
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


'*** Red flasher ***
Sub FlasherFlash24_Timer()
  dim flashx3, matdim
  If not Flasherflash24.TimerEnabled Then
    Flasherflash24.TimerEnabled = True
    Flasherflash24.visible = 1
    Flasherlit24.visible = 1
  End If
  flashx3 = FlashLevel24 * FlashLevel24 * FlashLevel24
  Flasherflash24.opacity = 1500 * flashx3
  Flasherlit24.BlendDisableLighting = 10 * flashx3
  Flasherbase24.BlendDisableLighting =  flashx3
  Flasherlight24.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel24)
  Flasherlit24.material = "domelit" & matdim
  FlashLevel24 = FlashLevel24 * 0.9 - 0.01
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


'*** Red flasher ***
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


'**********************************************************************************************************
'STATs pic change script (modified)

Dim TotalSpin, SpinNow

TotalSpin = 8

Sub ChangeSpin
  SpinNow = SpinNow + 1
  If SpinNow > 1 Then
    sw25.image = "spinner-"&SpinNow

  Else
    sw25.image = "spinner"

  End If
  If SpinNow = TotalSpin Then SpinNow = 0
End Sub

'**********************************************************************************************************




' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

'Table Init
Sub table1_Init
    vpmInit me:LoadLUT:NVOffset (13)
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "TRON Classic (Original 2022) v2.1" & vbNewLine & "VPX table by JPSalas + HF" & vbNewLine & "Graphics by HiRez00-Sounds Flashers by Xenonph"
        '.Games(cGameName).Settings.Value("rol")=0
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        .Games(cGameName).Settings.Value("dmd_red") = 0
        .Games(cGameName).Settings.Value("dmd_green") = 223
        .Games(cGameName).Settings.Value("dmd_blue") = 223
        .Games(cGameName).Settings.Value("sound")=0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        On Error Resume Next
        .SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - Then add the timer to renable all the solenoids after 2 seconds
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    'Nudging
    vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 4
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot, RightSlingShot, RightSlingShot2)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 67, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    'Left Saucer
'    Set bsLSaucer = New cvpmBallStack
'    With bsLSaucer
'        .InitSaucer sw35, 35, 60, 13
'        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
'        .KickAngleVar = 2
'    End With

    'Droptargets
    Set dtL = New cvpmDropTarget
    dtL.InitDrop Array(sw00, sw10, sw20, sw30), Array(00, 10, 20, 30)
    dtL.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    Set dtT = New cvpmDropTarget
    dtT.InitDrop Array(sw01, sw11, sw21, sw31), Array(01, 11, 21, 31)
    dtT.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    Set dtR = New cvpmDropTarget
    dtR.InitDrop Array(sw02, sw12, sw22, sw32), Array(02, 12, 22, 32)
    dtR.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1:If PeriodCorrect=0 then PlayMusic"TRON\0ta00.ogg":Drain.timerinterval=79000:Drain.timerenabled=1:Else Exit Sub
End Sub

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal KeyCode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySoundAtVol "fx_PlungerPull", Plunger, 1:Plunger.Pullback
  If KeyCode = 3 Then NextLUT
  If KeyCode = 50 Then ChangeSpin
  If KeyCode = RightMagnaSave and PeriodCorrect=0 Then PeriodCorrect=1:EndMusic:PlaySound"0tc06"
  If KeyCode = LeftMagnaSave and PeriodCorrect=1 Then PeriodCorrect=0:PlaySound"0tc02"
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If KeyCode = 6 and PeriodCorrect=0 Then
            FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=2000:sw13b.timerenabled=1
      Dim x
      x = INT(3 * RND(1) )
      Select Case x
      Case 0:StopSounds:PlaySound("0tc01")
      Case 1:StopSounds:PlaySound("0tc02")
      Case 2:StopSounds:PlaySound("0tc01")
      End Select
            End If

    If KeyCode = 6 and PeriodCorrect=1 Then
            FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=2000:sw13b.timerenabled=1
      Dim t
      t = INT(4 * RND(1) )
      Select Case t
      Case 0:StopSounds:PlaySound("0tc01")
      Case 1:StopSounds:PlaySound("0tc01")
      Case 2:StopSounds:PlaySound("0tc01")
      Case 3:StopSounds:PlaySound("0tc01")
      End Select
            End If

  If KeyCode = 4 and PeriodCorrect=0 Then
            FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=2000:sw13b.timerenabled=1
      Dim y
      y = INT(3 * RND(1) )
      Select Case y
      Case 0:StopSounds:PlaySound("0tc01")
      Case 1:StopSounds:PlaySound("0tc02")
      Case 2:StopSounds:PlaySound("0tc01")
      End Select
      end if

  If KeyCode = 4 and PeriodCorrect=1 Then
            FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=2000:sw13b.timerenabled=1
      Dim r
      r = INT(4 * RND(1) )
      Select Case r
      Case 0:StopSounds:PlaySound("0tc01")
      Case 1:StopSounds:PlaySound("0tc01")
      Case 2:StopSounds:PlaySound("0tc01")
      Case 3:StopSounds:PlaySound("0tc01")
      End Select
      end if


  If KeyCode = 2 and PeriodCorrect=0 Then
      Dim zz
      zz = INT(15 * RND(1) )
      Select Case zz
      Case 0:StopSounds:PlaySound("0ts01")
      Case 1:StopSounds:PlaySound("0ts02")
      Case 2:StopSounds:PlaySound("0ts03")
      Case 3:StopSounds:PlaySound("0ts04")
      Case 4:StopSounds:PlaySound("0ts05")
      Case 5:StopSounds:PlaySound("0ts06")
      Case 6:StopSounds:PlaySound("0ts07")
      Case 7:StopSounds:PlaySound("0ts08")
      Case 8:StopSounds:PlaySound("0ts09")
      Case 9:StopSounds:PlaySound("0ts10")
      Case 10:StopSounds:PlaySound("0ts11")
      Case 11:StopSounds:PlaySound("0ts12")
      Case 12:StopSounds:PlaySound("0ts13")
      Case 13:StopSounds:PlaySound("0ts14")
      Case 14:StopSounds:PlaySound("0ts15")
      End Select
    End If
  If KeyCode = 2 and PeriodCorrect=1 Then PlaySound"0tc06"



    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAtVol "fx_plunger", Plunger, 1:Plunger.Fire

End Sub


Sub StopSounds()

   StopSound"0tc01"
   StopSound"0tc02"
   StopSound"0td01"
   StopSound"0td02"
   StopSound"0td03"
   StopSound"0td04"
   StopSound"0td05"
   StopSound"0td06"
   StopSound"0td07"
   StopSound"0td08"
   StopSound"0td09"
   StopSound"0td10"
   StopSound"0td11"
   StopSound"0td12"
   StopSound"0td13"
   StopSound"0td14"
   StopSound"0td15"
   StopSound"0td16"
   StopSound"0td17"
   StopSound"0td18"
   StopSound"0td19"
   StopSound"0tpa01"
   StopSound"0tpa02"
   StopSound"0tpa03"
   StopSound"0tpa04"
   StopSound"0ts01"
   StopSound"0ts02"
   StopSound"0ts03"
   StopSound"0ts04"
   StopSound"0ts05"
   StopSound"0ts06"
   StopSound"0ts07"
   StopSound"0ts08"
   StopSound"0ts09"
   StopSound"0ts10"
   StopSound"0ts11"
   StopSound"0ts12"
   StopSound"0ts13"
   StopSound"0ts14"
   StopSound"0ts15"

End Sub


Dim BIP


' Slings
Dim LStep, RStep, RStep2

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), ActiveBall, 1:DOF 104, DOFPulse:FlashLevel9 = 1 : FlasherFlash9_Timer:FlashLevel25 = 1 : FlasherFlash25_Timer
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 34
    LeftSlingShot.TimerEnabled = 1
      Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"0tb03"
      Case 1:PlaySound"0tb04"
      Case 2:PlaySound"0tb05"
      End Select
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select

    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), ActiveBall, 1:DOF 103, DOFPulse:FlashLevel10 = 1 : FlasherFlash10_Timer:FlashLevel25 = 1 : FlasherFlash25_Timer
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 34
    RightSlingShot.TimerEnabled = 1
      Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"0tb03"
      Case 1:PlaySound"0tb04"
      Case 2:PlaySound"0tb05"
      End Select
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select

    RStep = RStep + 1
End Sub

Sub RightSlingShot2_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), ActiveBall, 1:DOF 105, DOFPulse:FlashLevel3 = 1 : FlasherFlash3_Timer:FlashLevel25 = 1 : FlasherFlash25_Timer
    RightSling8.Visible = 1
    Remk2.RotX = 26
    RStep2 = 0
    vpmTimer.PulseSw 34
    RightSlingShot2.TimerEnabled = 1
      Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"0tb03"
      Case 1:PlaySound"0tb04"
      Case 2:PlaySound"0tb05"
      End Select
End Sub

Sub RightSlingShot2_Timer
    Select Case RStep2
        Case 1:RightSLing8.Visible = 0:RightSLing7.Visible = 1:Remk2.RotX = 14
        Case 2:RightSLing7.Visible = 0:RightSLing6.Visible = 1:Remk2.RotX = 2
        Case 3:RightSLing6.Visible = 0:Remk2.RotX = -10:RightSlingShot2.TimerEnabled = 0
    End Select

    RStep2 = RStep2 + 1
End Sub

' Bumpers

Sub Bumper1_Hit:vpmTimer.PulseSw 24:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors), ActiveBall, 1:DOF 102, DOFPulse:FlashLevel7 = 1 : FlasherFlash7_Timer:FlashLevel25 = 1 : FlasherFlash25_Timer
      Dim w
      w = INT(2 * RND(1) )
      Select Case w
      Case 0:PlaySound"0tb01"
      Case 1:PlaySound"0tb02"
      End Select
End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 24:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors), ActiveBall, 1:DOF 101, DOFPulse:FlashLevel8 = 1 : FlasherFlash8_Timer:FlashLevel25 = 1 : FlasherFlash25_Timer
      Dim w
      w = INT(2 * RND(1) )
      Select Case w
      Case 0:PlaySound"0tb01"
      Case 1:PlaySound"0tb02"
      End Select
End Sub


'Plunger Switch

Sub Trigger1_Hit()
    If PeriodCorrect=0 Then
       EndMusic:StopSounds:Gate1.timerenabled=0:Gate2.timerenabled=0:sw03.timerenabled=0:sw13.timerenabled=0:Drain.timerenabled=0:sw33.timerenabled=0
       sw03b.timerenabled=0:sw13b.timerenabled=0:sw23.timerenabled=0:sw23.timerenabled=1
    Else
            Gate2.timerenabled=0
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tp00d"
      Case 1:PlaySound"0tp00e"
      Case 2:PlaySound"0tp00h"
      Case 3:PlaySound"0tp00i"
      End Select
     End If
End Sub


' Rollovers
Sub sw03_Hit:Controller.Switch(03) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:PlaySound"0tl01":End Sub
Sub sw03_UnHit:Controller.Switch(03) = 0:End Sub

Sub sw03_timer
      Dim z
      z = INT(19 * RND(1) )
      Select Case z
      Case 0:PlaySound"0td01":sw13.timerinterval=2000:sw13.timerenabled=1
      Case 1:PlaySound"0td02":sw13.timerinterval=2000:sw13.timerenabled=1
      Case 2:PlaySound"0td03":sw13.timerinterval=3000:sw13.timerenabled=1
      Case 3:PlaySound"0td04":sw13.timerinterval=3000:sw13.timerenabled=1
      Case 4:PlaySound"0td05":sw13.timerinterval=4000:sw13.timerenabled=1
      Case 5:PlaySound"0td06":sw13.timerinterval=3000:sw13.timerenabled=1
      Case 6:PlaySound"0td07":sw13.timerinterval=4000:sw13.timerenabled=1
      Case 7:PlaySound"0td08":sw13.timerinterval=2000:sw13.timerenabled=1
      Case 8:PlaySound"0td09":sw13.timerinterval=2000:sw13.timerenabled=1
      Case 9:PlaySound"0td10":sw13.timerinterval=4000:sw13.timerenabled=1
      Case 10:PlaySound"0td11":sw13.timerinterval=3000:sw13.timerenabled=1
      Case 11:PlaySound"0td12":sw13.timerinterval=2000:sw13.timerenabled=1
      Case 12:PlaySound"0td13":sw13.timerinterval=2000:sw13.timerenabled=1
      Case 13:PlaySound"0td14":sw13.timerinterval=2000:sw13.timerenabled=1
      Case 14:PlaySound"0td15":sw13.timerinterval=2000:sw13.timerenabled=1
      Case 15:PlaySound"0td16":sw13.timerinterval=4000:sw13.timerenabled=1
      Case 16:PlaySound"0td17":sw13.timerinterval=2000:sw13.timerenabled=1
      Case 17:PlaySound"0td18":sw13.timerinterval=2000:sw13.timerenabled=1
      Case 18:PlaySound"0td19":sw13.timerinterval=2000:sw13.timerenabled=1
      End Select
sw03.timerenabled=0
End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:PlaySound"0tl01":End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw13_timer
 If BIP=0 Then
 PlayMusic"TRON\0ta00.ogg":Drain.timerinterval=79000:Drain.timerenabled=1
 sw13.timerenabled=0
 End If
End Sub


Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:PlaySound"0tl01":End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw23_timer
      Dim z
      z = INT(9 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tp00a":PlayMusic"TRON\0tp02.ogg":sw33.timerinterval=192900:sw33.timerenabled=1
      Case 1:PlaySound"0tp00b":PlayMusic"TRON\0tp02.ogg":sw33.timerinterval=192900:sw33.timerenabled=1
      Case 2:PlaySound"0tp00c":PlayMusic"TRON\0tp02.ogg":sw33.timerinterval=192900:sw33.timerenabled=1
      Case 3:PlaySound"0tp00d":PlayMusic"TRON\0tp02.ogg":sw33.timerinterval=192900:sw33.timerenabled=1
      Case 4:PlaySound"0tp00e":PlayMusic"TRON\0tp02.ogg":sw33.timerinterval=192900:sw33.timerenabled=1
      Case 5:PlaySound"0tp00f":PlayMusic"TRON\0tp02.ogg":sw33.timerinterval=192900:sw33.timerenabled=1
      Case 6:PlaySound"0tp00g":PlayMusic"TRON\0tp02.ogg":sw33.timerinterval=192900:sw33.timerenabled=1
      Case 7:PlaySound"0tp00h":PlayMusic"TRON\0tp02.ogg":sw33.timerinterval=192900:sw33.timerenabled=1
      Case 8:PlaySound"0tp00i":PlayMusic"TRON\0tp02.ogg":sw33.timerinterval=192900:sw33.timerenabled=1
      End Select
sw23.timerenabled=0
End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:PlaySound"0tl01":End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw33_timer
     PlayMusic"TRON\0tp02.ogg":sw23.timerinterval=192900:sw23.timerenabled=1

sw33.timerenabled=0
End Sub

Sub sw05_Hit
        If PeriodCorrect=0 Then
            Controller.Switch(05) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
      Dim z
      z = INT(6 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tl01":PlaySound"0tfx01"
      Case 1:PlaySound"0tl03":PlaySound"0tfx02"
      Case 2:PlaySound"0tl03":PlaySound"0tfx01"
      Case 3:PlaySound"0tl01":PlaySound"0tfx02"
      Case 4:PlaySound"0tl01"
      Case 5:PlaySound"0tl03"
      End Select
        Else
            Controller.Switch(05) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
      Dim zt
      zt = INT(6 * RND(1) )
      Select Case zt
      Case 0:PlaySound"0tl01"
      Case 1:PlaySound"0tl03"
      Case 2:PlaySound"0tl03"
      Case 3:PlaySound"0tl01"
      Case 4:PlaySound"0tl01"
      Case 5:PlaySound"0tl03"
      End Select
        End If
End Sub

Sub sw05_UnHit:Controller.Switch(05) = 0:End Sub

Sub sw03b_Hit
        If PeriodCorrect=0 Then
            Controller.Switch(03) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:FlashLevel12 = 1 : FlasherFlash12_Timer:EndMusic:StopSounds
      Dim z
      z = INT(8 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tsd01":PlaySound"0tsd06"
      Case 1:PlaySound"0tsd02"
      Case 2:PlaySound"0tsd03":PlaySound"0tsd06"
      Case 3:PlaySound"0tsd04":PlaySound"0tsd06"
      Case 4:PlaySound"0tsd01":PlaySound"0tsd07"
      Case 5:PlaySound"0tsd06"
      Case 6:PlaySound"0tsd03":PlaySound"0tsd07"
      Case 7:PlaySound"0tsd04":PlaySound"0tsd07"
      End Select
        Else
            Controller.Switch(03) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:FlashLevel12 = 1 : FlasherFlash12_Timer:EndMusic:StopSounds
      Dim zc
      zc = INT(8 * RND(1) )
      Select Case zc
      Case 0:PlaySound"0tsd06"
      Case 1:PlaySound"0tsd06"
      Case 2:PlaySound"0tsd06"
      Case 3:PlaySound"0tsd06"
      Case 4:PlaySound"0tsd06"
      Case 5:PlaySound"0tsd06"
      Case 6:PlaySound"0tsd07"
      Case 7:PlaySound"0tsd07"
      End Select
        End If
End Sub

Sub sw03b_UnHit:Controller.Switch(03) = 0:End Sub

Sub sw03b_timer
    Dim x
    x = INT(11 * RND(1) )
    Select Case x
    Case 0:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
    Case 1:FlashLevel7 = 1 : FlasherFlash7_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer
    Case 2:FlashLevel4 = 1 : FlasherFlash4_Timer:FlashLevel11 = 1 : FlasherFlash11_Timer
    Case 3:FlashLevel2 = 1 : FlasherFlash2_Timer:FlashLevel24 = 1 : FlasherFlash24_Timer
    Case 4:FlashLevel12 = 1 : FlasherFlash12_Timer:FlashLevel13 = 1 : FlasherFlash13_Timer
    Case 5:FlashLevel14 = 1 : FlasherFlash14_Timer
    Case 6:FlashLevel9 = 1 : FlasherFlash9_Timer:FlashLevel10 = 1 : FlasherFlash10_Timer
    Case 7:FlashLevel7 = 1 : FlasherFlash7_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer
    Case 8:FlashLevel2 = 1 : FlasherFlash2_Timer:FlashLevel3 = 1 : FlasherFlash3_Timer
    Case 9:FlashLevel12 = 1 : FlasherFlash12_Timer:FlashLevel13 = 1 : FlasherFlash13_Timer:FlashLevel14 = 1 : FlasherFlash14_Timer:FlashLevel24 = 1 : FlasherFlash24_Timer:FlashLevel11 = 1 : FlasherFlash11_Timer:FlashLevel4 = 1 : FlasherFlash4_Timer
    Case 10:FlashLevel9 = 1 : FlasherFlash9_Timer:FlashLevel10 = 1 : FlasherFlash10_Timer:FlashLevel3 = 1 : FlasherFlash3_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer:FlashLevel7 = 1 : FlasherFlash7_Timer:FlashLevel8 = 1 : FlasherFlash8_Timer
    End Select

End sub

Sub sw13b_Hit:Controller.Switch(13) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:FlashLevel12 = 1 : FlasherFlash12_Timer
      Dim z
      z = INT(3 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tl01"
      Case 1:PlaySound"0tl01"
      Case 2:PlaySound"0tl03"
      End Select
End Sub
Sub sw13b_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw13b_timer
         sw03b.timerenabled=0
         sw13b.timerenabled=0
End sub

Sub sw23b_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:FlashLevel13 = 1 : FlasherFlash13_Timer
      Dim z
      z = INT(3 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tl01"
      Case 1:PlaySound"0tl01"
      Case 2:PlaySound"0tl03"
      End Select
End Sub
Sub sw23b_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw23b_timer
      Dim w
      w = INT(2 * RND(1) )
      Select Case w
      Case 0:PlaySound"0tbonus01"
      Case 1:PlaySound"0tbonus02"

      End Select
sw23b.timerenabled=0
End Sub

Sub sw33b_Hit
         If PeriodCorrect=0 Then
            Controller.Switch(33) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:FlashLevel13 = 1 : FlasherFlash13_Timer:EndMusic:StopSounds
      Dim z
      z = INT(8 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tsd01":PlaySound"0tsd06"
      Case 1:PlaySound"0tsd02"
      Case 2:PlaySound"0tsd03":PlaySound"0tsd06"
      Case 3:PlaySound"0tsd04":PlaySound"0tsd06"
      Case 4:PlaySound"0tsd01":PlaySound"0tsd07"
      Case 5:PlaySound"0tsd06"
      Case 6:PlaySound"0tsd03":PlaySound"0tsd07"
      Case 7:PlaySound"0tsd04":PlaySound"0tsd07"
      End Select
         Else
            Controller.Switch(33) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:FlashLevel13 = 1 : FlasherFlash13_Timer:EndMusic:StopSounds
      Dim zb
      zb = INT(8 * RND(1) )
      Select Case zb
      Case 0:PlaySound"0tsd06"
      Case 1:PlaySound"0tsd06"
      Case 2:PlaySound"0tsd06"
      Case 3:PlaySound"0tsd06"
      Case 4:PlaySound"0tsd06"
      Case 5:PlaySound"0tsd06"
      Case 6:PlaySound"0tsd07"
      Case 7:PlaySound"0tsd07"
      End Select
          End If
End Sub

Sub sw33b_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:DOF 106, DOFOn
      Dim z
      z = INT(3 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tl01"
      Case 1:PlaySound"0tl01"
      Case 2:PlaySound"0tl03"
      End Select
End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:DOF 106, DOFOff:End Sub

Sub sw15b_Hit
          If PeriodCorrect=0 Then
            Controller.Switch(15) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:DOF 107, DOFOn
      Dim z
      z = INT(7 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tfx02"
      Case 1:PlaySound"0tfx02"
      Case 2:PlaySound"0tfx02"
      Case 3:PlaySound"0tfx02"
      Case 4:PlaySound"0tfx02"
      Case 5:PlaySound"0tfx02"
      Case 6:PlaySound"0tfx02"
      End Select
          Else
            Controller.Switch(15) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:DOF 107, DOFOn
      Dim zc
      zc = INT(7 * RND(1) )
      Select Case zc
      Case 0:PlaySound"0tspin"
      Case 1:PlaySound"0tspin"
      Case 2:PlaySound"0tspin"
      Case 3:PlaySound"0tspin"
      Case 4:PlaySound"0tspin"
      Case 5:PlaySound"0tspin"
      Case 6:PlaySound"0tspin"
      End Select
          End If
End Sub

Sub sw15b_UnHit:Controller.Switch(15) = 0:DOF 107, DOFOff:End Sub

Sub sw25b_Hit:Controller.Switch(25) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:DOF 106, DOFOn
      Dim z
      z = INT(3 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tl01"
      Case 1:PlaySound"0tl01"
      Case 2:PlaySound"0tl03"
      End Select
End Sub
Sub sw25b_UnHit:Controller.Switch(25) = 0:DOF 106, DOFOff:End Sub

Sub sw14a_Hit:Controller.Switch(14) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:lsw14a.Duration 2, 1000, 0
      Dim z
      z = INT(6 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tr01"
      Case 1:PlaySound"0tr02"
      Case 2:PlaySound"0tr03"
      Case 3:PlaySound"0tr04"
      Case 4:PlaySound"0tr05"
      Case 5:PlaySound"0tr06"
      End Select
End Sub
Sub sw14a_UnHit:Controller.Switch(14) = 0:end sub

Sub sw14b_Hit:Controller.Switch(14) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:lsw14b.Duration 2, 1000, 0
      Dim z
      z = INT(8 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tr06"
      Case 1:PlaySound"0tr02"
      Case 2:PlaySound"0tr06"
      Case 3:PlaySound"0tr04"
      Case 4:PlaySound"0tr05"
      Case 5:PlaySound"0tr06"
      Case 6:PlaySound"0tr06"
      Case 7:PlaySound"0tr05"
      End Select
End Sub
Sub sw14b_UnHit:Controller.Switch(14) = 0:end sub

'Standup target

Sub sw04_Hit
         If PeriodCorrect=0 Then
            vpmTimer.PulseSw 4:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, 1:FlashLevel24 = 1 : FlasherFlash24_Timer
      Dim z
      z = INT(6 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tsd05"
      Case 1:PlaySound"0tsd05":PlaySound"0tfx01"
      Case 2:PlaySound"0tsd05":PlaySound"0tfx02"
      Case 3:PlaySound"0tsd05"
      Case 4:PlaySound"0tsd05"
      Case 5:PlaySound"0tsd05"
      End Select
          Else
            vpmTimer.PulseSw 4:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, 1:FlashLevel24 = 1 : FlasherFlash24_Timer
      Dim zg
      zg = INT(6 * RND(1) )
      Select Case zg
      Case 0:PlaySound"0tsd05"
      Case 1:PlaySound"0tsd05"
      Case 2:PlaySound"0tsd05"
      Case 3:PlaySound"0tsd05"
      Case 4:PlaySound"0tsd05"
      Case 5:PlaySound"0tsd05"
      End Select
          End If
End Sub

'rubbers
'Sub sw34a_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
'Sub sw34b_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
'Sub sw34c_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
'Sub sw34d_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
'Sub sw34e_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
'Sub sw34f_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
'Sub sw34g_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw34a_Hit():vpmTimer.PulseSw 34:End Sub
Sub sw34b_Hit():vpmTimer.PulseSw 34:End Sub

Sub sw34c_Hit()
  vpmTimer.PulseSw 34
  r34c.visible=0
  r34c1.visible=1
  me.uservalue=1
  Me.timerenabled=1
End Sub

sub sw34c_timer                 'default 50 timer
  select case me.uservalue
    Case 1: r34c1.visible=0: r34c.visible=1
    case 2: r34c.visible=0: r34c2.visible=1
    Case 3: r34c2.visible=0: r34c.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

Sub sw34d_Hit():vpmTimer.PulseSw 34:End Sub

Sub sw34e_Hit()
  vpmTimer.PulseSw 34
  r34e.visible=0
  r34e1.visible=1
  me.uservalue=1
  Me.timerenabled=1
End Sub

sub sw34e_timer                 'default 50 timer
  select case me.uservalue
    Case 1: r34e1.visible=0: r34e.visible=1
    case 2: r34e.visible=0: r34e2.visible=1
    Case 3: r34e2.visible=0: r34e.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

Sub sw34f_Hit():vpmTimer.PulseSw 34:End Sub
Sub sw34g_Hit()
  vpmTimer.PulseSw 34
  r34g.visible=0
  r34g1.visible=1
  me.uservalue=1
  Me.timerenabled=1
End Sub

sub sw34g_timer                 'default 50 timer
  select case me.uservalue
    Case 1: r34g1.visible=0: r34g.visible=1
    case 2: r34g.visible=0: r34g2.visible=1
    Case 3: r34g2.visible=0: r34g.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

'Spinner
Sub sw25_Spin():vpmTimer.PulseSw 25:PlaySoundAtVol "fx_spinner", sw25, 1:DOF 108, DOFPulse:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer:PlaySound"0tspin":End Sub

'Drop-Targets
Sub sw00_Hit():dtL.hit 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:FlashLevel2 = 1 : FlasherFlash2_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tc01"
      Case 1:PlaySound"0tc01"
      Case 2:PlaySound"0tc01"
      Case 3:PlaySound"0tc01"
      End Select
End Sub

Sub sw10_Hit():dtL.hit 2:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:FlashLevel2 = 1 : FlasherFlash2_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tc01"
      Case 1:PlaySound"0tc01"
      Case 2:PlaySound"0tc01"
      Case 3:PlaySound"0tc01"
      End Select
End Sub

Sub sw20_Hit():dtL.hit 3:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:FlashLevel2 = 1 : FlasherFlash2_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tc01"
      Case 1:PlaySound"0tc01"
      Case 2:PlaySound"0tc01"
      Case 3:PlaySound"0tc01"
      End Select
End Sub

Sub sw30_Hit():dtL.hit 4:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:FlashLevel2 = 1 : FlasherFlash2_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tc01"
      Case 1:PlaySound"0tc01"
      Case 2:PlaySound"0tc01"
      Case 3:PlaySound"0tc01"
      End Select
End Sub

Sub sw01_Hit():dtT.hit 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:FlashLevel4 = 1 : FlasherFlash4_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tc01"
      Case 1:PlaySound"0tc01"
      Case 2:PlaySound"0tc01"
      Case 3:PlaySound"0tc01"
      End Select
End Sub

Sub sw11_Hit():dtT.hit 2:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:FlashLevel4 = 1 : FlasherFlash4_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tc01"
      Case 1:PlaySound"0tc01"
      Case 2:PlaySound"0tc01"
      Case 3:PlaySound"0tc01"
      End Select
End Sub

Sub sw21_Hit():dtT.hit 3:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:FlashLevel4 = 1 : FlasherFlash4_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tc01"
      Case 1:PlaySound"0tc01"
      Case 2:PlaySound"0tc01"
      Case 3:PlaySound"0tc01"
      End Select
End Sub

Sub sw31_Hit():dtT.hit 4:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:FlashLevel4 = 1 : FlasherFlash4_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tc01"
      Case 1:PlaySound"0tc01"
      Case 2:PlaySound"0tc01"
      Case 3:PlaySound"0tc01"
      End Select
End Sub

Sub sw02_Hit():dtR.hit 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:FlashLevel11 = 1 : FlasherFlash11_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tc01"
      Case 1:PlaySound"0tc01"
      Case 2:PlaySound"0tc01"
      Case 3:PlaySound"0tc01"
      End Select
End Sub

Sub sw12_Hit():dtR.hit 2:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:FlashLevel11 = 1 : FlasherFlash11_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tc01"
      Case 1:PlaySound"0tc01"
      Case 2:PlaySound"0tc01"
      Case 3:PlaySound"0tc01"
      End Select
End Sub

Sub sw22_Hit():dtR.hit 3:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:FlashLevel11 = 1 : FlasherFlash11_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tc01"
      Case 1:PlaySound"0tc01"
      Case 2:PlaySound"0tc01"
      Case 3:PlaySound"0tc01"
      End Select
End Sub

Sub sw32_Hit():dtR.hit 4:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:FlashLevel11 = 1 : FlasherFlash11_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
      Dim z
      z = INT(4 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tc01"
      Case 1:PlaySound"0tc01"
      Case 2:PlaySound"0tc01"
      Case 3:PlaySound"0tc01"
      End Select
End Sub

Sub sw32_timer
      Dim w
      w = INT(2 * RND(1) )
      Select Case w
      Case 0:PlayMusic"TRON\0ta11.ogg":Drain.timerinterval=127000:Drain.timerenabled=1
      Case 1:PlayMusic"TRON\0ta12.ogg":Drain.timerinterval=220000:Drain.timerenabled=1

      End Select
sw32.timerenabled=0
End Sub


' Drain & Holes
Sub Drain_Hit
         If PeriodCorrect=0 Then
          bsTrough.AddBall Me:PlaySoundAtVol "fx_drain", Drain, 1:FlashLevel14 = 1 : FlasherFlash14_Timer:sw33.timerenabled=0:sw23.timerenabled=0
          Flasherlight15.IntensityScale = 0:Flasherlight16.IntensityScale = 0:Flasherlight17.IntensityScale = 0:Flasherlight18.IntensityScale = 0:Flasherlight1.IntensityScale = 0:Flasherlight19.IntensityScale = 0:Flasherlight20.IntensityScale = 0:Flasherlight21.IntensityScale = 0:Flasherlight22.IntensityScale = 0
          Flasherlight23.IntensityScale = 0:Flasherlight3.IntensityScale = 0
      Dim w
      w = INT(2 * RND(1) )
      Select Case w
      Case 0:EndMusic:PlaySound"0td00a":sw03.timerinterval=1500:sw03.timerenabled=1
      Case 1:EndMusic:PlaySound"0td00b":sw03.timerinterval=1500:sw03.timerenabled=1
      End Select
         Else
          bsTrough.AddBall Me:PlaySoundAtVol "fx_drain", Drain, 1:FlashLevel14 = 1 : FlasherFlash14_Timer:sw33.timerenabled=0:sw23.timerenabled=0
          Flasherlight15.IntensityScale = 0:Flasherlight16.IntensityScale = 0:Flasherlight17.IntensityScale = 0:Flasherlight18.IntensityScale = 0:Flasherlight1.IntensityScale = 0:Flasherlight19.IntensityScale = 0:Flasherlight20.IntensityScale = 0:Flasherlight21.IntensityScale = 0:Flasherlight22.IntensityScale = 0
          Flasherlight23.IntensityScale = 0:Flasherlight3.IntensityScale = 0
      Dim ll
      ll = INT(2 * RND(1) )
      Select Case ll
      Case 0:EndMusic:PlaySound"0tb02"
      Case 1:EndMusic:PlaySound"0tc06"
      End Select
          End If
End Sub

Sub drain_timer
      Dim z
      z = INT(8 * RND(1) )
      Select Case z
      Case 0:PlayMusic"TRON\0ta03.ogg":sw35.timerinterval=12000:sw35.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=12000:sw13b.timerenabled=1
      Case 1:PlayMusic"TRON\0ta04.ogg":sw35.timerinterval=33000:sw35.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=33000:sw13b.timerenabled=1
      Case 2:PlayMusic"TRON\0ta05.ogg":sw35.timerinterval=19000:sw35.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=19000:sw13b.timerenabled=1
      Case 3:PlayMusic"TRON\0ta06.ogg":sw35.timerinterval=62000:sw35.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=62000:sw13b.timerenabled=1
      Case 4:PlayMusic"TRON\0ta07.ogg":sw35.timerinterval=16000:sw35.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=16000:sw13b.timerenabled=1
      Case 5:PlayMusic"TRON\0ta08.ogg":sw35.timerinterval=22000:sw35.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=22000:sw13b.timerenabled=1
      Case 6:PlayMusic"TRON\0ta09.ogg":sw35.timerinterval=14000:sw35.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=14000:sw13b.timerenabled=1
      Case 7:PlayMusic"TRON\0ta10.ogg":sw35.timerinterval=21000:sw35.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=21000:sw13b.timerenabled=1

      End Select
Drain.timerenabled=0
End Sub

Sub sw35_Hit
         If PeriodCorrect=0 Then
          controller.switch(35) = 1:PlaySoundAtVol "fx_kicker_enter", sw35, 1:PlaySound"0tk01":sw03b.timerinterval=100:sw03b.timerenabled=1  ':bsLSaucer.AddBall 0
      Dim z
      z = INT(26 * RND(1) )
      Select Case z
      Case 0:PlaySound"0tv01":sw23b.timerinterval=2000:sw23b.timerenabled=1
      Case 1:PlaySound"0tv02":sw23b.timerinterval=3000:sw23b.timerenabled=1
      Case 2:PlaySound"0tv03":sw23b.timerinterval=3000:sw23b.timerenabled=1
      Case 3:PlaySound"0tv04":sw23b.timerinterval=4000:sw23b.timerenabled=1
      Case 4:PlaySound"0tv05":sw23b.timerinterval=1000:sw23b.timerenabled=1
      Case 5:PlaySound"0tv06":sw23b.timerinterval=4000:sw23b.timerenabled=1
      Case 6:PlaySound"0tv07":sw23b.timerinterval=2000:sw23b.timerenabled=1
      Case 7:PlaySound"0tv08":sw23b.timerinterval=3000:sw23b.timerenabled=1
      Case 8:PlaySound"0tv09":sw23b.timerinterval=4000:sw23b.timerenabled=1
      Case 9:PlaySound"0tv10":sw23b.timerinterval=2000:sw23b.timerenabled=1
      Case 10:PlaySound"0tv11":sw23b.timerinterval=3000:sw23b.timerenabled=1
      Case 11:PlaySound"0tv12":sw23b.timerinterval=3000:sw23b.timerenabled=1
      Case 12:PlaySound"0tv13":sw23b.timerinterval=3000:sw23b.timerenabled=1
      Case 13:PlaySound"0tv14":sw23b.timerinterval=3000:sw23b.timerenabled=1
      Case 14:PlaySound"0tv15":sw23b.timerinterval=3000:sw23b.timerenabled=1
      Case 15:PlaySound"0tv16":sw23b.timerinterval=4000:sw23b.timerenabled=1
      Case 16:PlaySound"0tv17":sw23b.timerinterval=4000:sw23b.timerenabled=1
      Case 17:PlaySound"0tv18":sw23b.timerinterval=4000:sw23b.timerenabled=1
      Case 18:PlaySound"0tv19":sw23b.timerinterval=4000:sw23b.timerenabled=1
      Case 19:PlaySound"0tv20":sw23b.timerinterval=2000:sw23b.timerenabled=1
      Case 20:PlaySound"0tv21":sw23b.timerinterval=1000:sw23b.timerenabled=1
      Case 21:PlaySound"0tv22":sw23b.timerinterval=1000:sw23b.timerenabled=1
      Case 22:PlaySound"0tv23":sw23b.timerinterval=2000:sw23b.timerenabled=1
      Case 23:PlaySound"0tv24":sw23b.timerinterval=2000:sw23b.timerenabled=1
      Case 24:PlaySound"0tv25":sw23b.timerinterval=1000:sw23b.timerenabled=1
      Case 25:PlaySound"0tv26":sw23b.timerinterval=2000:sw23b.timerenabled=1
      End Select
        Else
          controller.switch(35) = 1:PlaySoundAtVol "fx_kicker_enter", sw35, 1:PlaySound"0tk01":sw03b.timerinterval=100:sw03b.timerenabled=1:sw23b.timerinterval=500:sw23b.timerenabled=1  ':bsLSaucer.AddBall 0
        End If
End Sub

Sub sw35exit(enabled)
  if enabled then
  controller.switch(35) = 0
  sw35.kick 90, 16, 20
  PlaySoundAtVol SoundFX("fx_kicker",DOFContactors), Pkickarm, 1
  Pkickarm.rotz=18
  gi33.uservalue=1
  gi33.timerenabled=1
  end if
end Sub

sub gi33_timer
  Select Case me.uservalue
    Case 2:
    Pkickarm.rotz=0
    me.timerenabled=0
  End Select
  me.uservalue=me.uservalue+1
end Sub

Sub sw35_UnHit:PlaySound"0tk02":sw03b.timerenabled=0:sw23b.timerenabled=0:StopSound"0tbonus01":StopSound"0tbonus02":End Sub

Sub sw35_timer
      Dim w
      w = INT(2 * RND(1) )
      Select Case w
      Case 0:PlayMusic"TRON\0ta01.ogg":Gate3.timerinterval=101000:Gate3.timerenabled=1
      Case 1:PlayMusic"TRON\0ta02.ogg":Gate3.timerinterval=70000:Gate3.timerenabled=1

      End Select
sw35.timerenabled=0
End Sub


Sub Gate1_Hit()
BIP=0:FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
End Sub

Sub Gate1_timer:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=2000:sw13b.timerenabled=1
      Dim w
      w = INT(4 * RND(1) )
      Select Case w
      Case 0:PlaySound"0tpa01":PlayMusic"TRON\0tp01.ogg":Gate2.timerinterval=27000:Gate2.timerenabled=1
      Case 1:PlaySound"0tpa02":PlayMusic"TRON\0tp01.ogg":Gate2.timerinterval=27000:Gate2.timerenabled=1
      Case 2:PlaySound"0tpa03":PlayMusic"TRON\0tp01.ogg":Gate2.timerinterval=27000:Gate2.timerenabled=1
      Case 3:PlaySound"0tpa04":PlayMusic"TRON\0tp01.ogg":Gate2.timerinterval=27000:Gate2.timerenabled=1
      End Select
Gate1.timerenabled=0
End Sub

Sub Gate2_Hit()
     If PeriodCorrect=0 Then
      EndMusic:PlayMusic"TRON\0tp01.ogg":Gate1.timerinterval=27000:Gate1.timerenabled=1
      Flasherlight15.IntensityScale = 1:Flasherlight16.IntensityScale = 1:Flasherlight17.IntensityScale = 1:Flasherlight18.IntensityScale = 1:Flasherlight1.IntensityScale = 1:Flasherlight19.IntensityScale = 1:Flasherlight20.IntensityScale = 1:Flasherlight21.IntensityScale = 1:Flasherlight22.IntensityScale = 1
      Flasherlight23.IntensityScale = 1:BIP=1
      Drain.timerenabled=0:sw35.timerenabled=0:sw13.timerenabled=0:sw13b.timerenabled=0:sw03b.timerenabled=0:sw32.timerenabled=0:Gate3.timerenabled=0
     Else
      Flasherlight15.IntensityScale = 1:Flasherlight16.IntensityScale = 1:Flasherlight17.IntensityScale = 1:Flasherlight18.IntensityScale = 1:Flasherlight1.IntensityScale = 1:Flasherlight19.IntensityScale = 1:Flasherlight20.IntensityScale = 1:Flasherlight21.IntensityScale = 1:Flasherlight22.IntensityScale = 1
      Flasherlight23.IntensityScale = 1
      Drain.timerenabled=0:sw35.timerenabled=0:sw13.timerenabled=0:sw13b.timerenabled=0:sw03b.timerenabled=0:sw32.timerenabled=0:Gate3.timerenabled=0
      BIP=1
     End If
End Sub

Sub Gate2_timer:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=2000:sw13b.timerenabled=1
      Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"0tpa01":PlayMusic"TRON\0tp01.ogg":Gate1.timerinterval=27000:Gate1.timerenabled=1
      Case 1:PlaySound"0tpa02":PlayMusic"TRON\0tp01.ogg":Gate1.timerinterval=27000:Gate1.timerenabled=1
      Case 2:PlaySound"0tpa03":PlayMusic"TRON\0tp01.ogg":Gate1.timerinterval=27000:Gate1.timerenabled=1
      End Select
Gate2.timerenabled=0
End Sub

Sub Gate3_Hit()
      If PeriodCorrect=0 Then
           FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
     Dim w
     w = INT(7 * RND(1) )
     Select Case w
     Case 0:PlaySound"0tv18"
     Case 1:PlaySound"0tp00d"
     Case 2:PlaySound"0tp00e"
     Case 3:PlaySound"0tp00f"
     Case 4:PlaySound"0tp00g"
     Case 5:PlaySound"0tp00h"
     Case 6:PlaySound"0tp00i"
     End Select
      Else
           FlashLevel5 = 1 : FlasherFlash5_Timer:FlashLevel6 = 1 : FlasherFlash6_Timer
     Dim ww
     ww = INT(4 * RND(1) )
     Select Case ww
     Case 0:PlaySound"0tp00e"
     Case 1:PlaySound"0tp00h"
     Case 2:PlaySound"0tp00i"
     Case 3:PlaySound"0tt01"
     End Select
      End If
End Sub

Sub Gate3_timer
      Dim z
      z = INT(8 * RND(1) )
      Select Case z
      Case 0:PlayMusic"TRON\0ta03.ogg":sw32.timerinterval=12000:sw32.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=12000:sw13b.timerenabled=1
      Case 1:PlayMusic"TRON\0ta04.ogg":sw32.timerinterval=33000:sw32.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=33000:sw13b.timerenabled=1
      Case 2:PlayMusic"TRON\0ta05.ogg":sw32.timerinterval=19000:sw32.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=19000:sw13b.timerenabled=1
      Case 3:PlayMusic"TRON\0ta06.ogg":sw32.timerinterval=62000:sw32.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=62000:sw13b.timerenabled=1
      Case 4:PlayMusic"TRON\0ta07.ogg":sw32.timerinterval=16000:sw32.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=16000:sw13b.timerenabled=1
      Case 5:PlayMusic"TRON\0ta08.ogg":sw32.timerinterval=22000:sw32.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=22000:sw13b.timerenabled=1
      Case 6:PlayMusic"TRON\0ta09.ogg":sw32.timerinterval=14000:sw32.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=14000:sw13b.timerenabled=1
      Case 7:PlayMusic"TRON\0ta10.ogg":sw32.timerinterval=21000:sw32.timerenabled=1:sw03b.timerinterval=100:sw03b.timerenabled=1:sw13b.timerinterval=21000:sw13b.timerenabled=1

      End Select
Gate3.timerenabled=0
End Sub


'****Solenoids

SolCallback(5) = "dtT.soldropup"
SolCallback(2) = "dtL.soldropup"
SolCallback(1) = "dtR.soldropup"
SolCallback(6) = "sw35exit"   '"bsLSaucer.SolOut"
SolCallback(8) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(9) = "bsTrough.SolOut"

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers),  LeftFlipper, 1
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers),  LeftFlipper, 1
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers),  RightFlipper, 1
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers),  RightFlipper, 1
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtVol "fx_rubber_flipper", LeftFlipper, .2
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtVol "fx_rubber_flipper", RightFlipper, .2
End Sub

'*****************
'   Gi Effects
'*****************

Dim OldGiState
OldGiState = -1 'start witht he Gi off

Sub GiON
    For each x in aGiLights
        GiEffect
    Next
    Flasherlight26.IntensityScale = 1:Flasherlight27.IntensityScale = 1:Flasherlight28.IntensityScale = 1:Flasherlight29.IntensityScale = 1:Flasherlight30.IntensityScale = 1:Flasherlight31.IntensityScale = 1
End Sub

Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
    Flasherlight26.IntensityScale = 0:Flasherlight27.IntensityScale = 0:Flasherlight28.IntensityScale = 0:Flasherlight29.IntensityScale = 0:Flasherlight30.IntensityScale = 0:Flasherlight31.IntensityScale = 0
End Sub

Sub GiEffect
    For each x in aGiLights
        x.Duration 2, 1000, 1
    Next
End Sub

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then
            GiOff
        Else
            GiOn
        End If
    End If
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200), FlashRepeat(200)

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

    If VarHidden Then
        UpdateLeds
    End If
    UpdateLamps
    GIUpdate
    RollingUpdate
End Sub

Sub UpdateLamps()
    NFadeT 1, l1, "TILT"
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    'NFadeL 8, l8
    NFadeT 10, l10, "High Score to Date"
    NFadeT 11, l11, "Game Over"
    ' NFadeL 12, l12
    ' NFadeL 13, l13
    ' NFadeL 14, l14
    ' NFadeL 15, l15
    ' NFadeL 16, l16
    ' NFadeL 17, l17
    ' NFadeL 18, l18
    ' NFadeL 19, l19
    ' NFadeL 20, l20
    ' NFadeL 21, l21
    ' NFadeL 22, l22
    ' NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    NFadeL 27, l27
    NFadeL 28, l28
    NFadeLm 29, l29b
    NFadeL 29, l29
    NFadeLm 30, l30b
    NFadeL 30, l30
    NFadeLm 31, l31b
    NFadeL 31, l31
    NFadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeLm 44, l44d
    NFadeLm 44, l44b
    NFadeL 44, l44
    NFadeLm 45, l45d
    NFadeLm 45, l45b
    NFadeL 45, l45
    NFadeLm 46, l46d
    NFadeLm 46, l46b
    NFadeL 46, l46
    NFadeLm 47, l47d
    NFadeLm 47, l47b
    NFadeL 47, l47
    NFadeLm 48, l48b
    NFadeL 48, l48
    NFadeLm 49, l49b
    NFadeL 49, l49
    NFadeLm 50, l50b
    NFadeL 50, l50
    NFadeLm 51, l51b
    NFadeL 51, l51
End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
        FlashRepeat(x) = 20     ' how many times the flash repeats
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

'Lights, Ramps & Primitives used as 4 step fading lights
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
            If FlashLevel(nr) <FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change anything, it just follows the main flasher
    Select Case FadingLevel(nr)
        Case 4, 5
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub FlashBlink(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) <FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 0 AND FlashRepeat(nr) Then 'repeat the flash
                FlashRepeat(nr) = FlashRepeat(nr) -1
                If FlashRepeat(nr) Then FadingLevel(nr) = 5
            End If
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 1 AND FlashRepeat(nr) Then FadingLevel(nr) = 4
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

Sub NFadeTm(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'      Gottlieb led Patterns
'************************************

Dim Digits(32)
Dim Patterns(11) 'normal numbers
Dim Patterns2(11) 'numbers with a comma
Dim Patterns3(11) 'the credits and ball in play

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 768   '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 124   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 103  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 896  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 231 '9

Patterns3(0) = 0     'empty
Patterns3(1) = 63    '0
Patterns3(2) = 6     '1
Patterns3(3) = 91    '2
Patterns3(4) = 79    '3
Patterns3(5) = 102   '4
Patterns3(6) = 109   '5
Patterns3(7) = 124   '6
Patterns3(8) = 7     '7
Patterns3(9) = 127   '8
Patterns3(10) = 103  '9

'Assign 7-digit output to reels
Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5
Set Digits(6) = a6

Set Digits(7) = b0
Set Digits(8) = b1
Set Digits(9) = b2
Set Digits(10) = b3
Set Digits(11) = b4
Set Digits(12) = b5
Set Digits(13) = b6

Set Digits(14) = c0
Set Digits(15) = c1
Set Digits(16) = c2
Set Digits(17) = c3
Set Digits(18) = c4
Set Digits(19) = c5
Set Digits(20) = c6

Set Digits(21) = d0
Set Digits(22) = d1
Set Digits(23) = d2
Set Digits(24) = d3
Set Digits(25) = d4
Set Digits(26) = d5
Set Digits(27) = d6

Set Digits(28) = e0
Set Digits(29) = e1
Set Digits(30) = e2
Set Digits(31) = e3

Sub UpdateLeds
    On Error Resume Next
    dim oldstat
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
if oldstat <> STAT then
debug.print STAT
oldstat = stat
end if
                If(stat = Patterns(jj) ) OR (stat = Patterns2(jj) ) OR (stat = Patterns3(jj) ) then Digits(chgLED(ii, 0) ).SetValue jj
            Next
        Next
    End IF
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' Thalamus - replacing this

' ' *********************************************************************
' '                      Supporting Ball & Sound Functions
' ' *********************************************************************
'
' Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
'     Vol = Csng(BallVel(ball) ^2 / 2000)
' End Function
'
' Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
'     Dim tmp
'     tmp = ball.x * 2 / table1.width-1
'     If tmp> 0 Then
'         Pan = Csng(tmp ^10)
'     Else
'         Pan = Csng(-((- tmp) ^10) )
'     End If
' End Function
'
' Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
'     Pitch = BallVel(ball) * 20
' End Function
'
' Function BallVel(ball) 'Calculates the ball speed
'     BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
' End Function
'
' Function AudioFade(ball) 'only on VPX 10.4 and newer
'     Dim tmp
'     tmp = ball.y * 2 / Table1.height-1
'     If tmp> 0 Then
'         AudioFade = Csng(tmp ^10)
'     Else
'         AudioFade = Csng(-((- tmp) ^10) )
'     End If
' End Function
'
'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls
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
        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 15000 'increase the pitch on a ramp or elevated surface
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
End Sub

'*****************************************
'     FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  LFlip.RotY = LeftFlipper.CurrentAngle
  RFlip.RotY = RightFlipper.CurrentAngle
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
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

' Thalamus - replacing this
' '**********************
' ' Ball Collision Sound
' '**********************
'
' Sub OnBallBallCollision(ball1, ball2, velocity)
'     PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
' End Sub

'Les transistors inutilisés qui servent pour les lampes alimentent la memoire des DropTargets
Dim N1, O1, N2, O2, N3, O3, N4, O4, N5, O5, N6, O6, N7, O7, N8, O8, N9, O9, N10, O10, N11, O11, N12, O12
N1 = 0:O1 = 0:N2 = 0:O2 = 0:N3 = 0:O3 = 0:N4 = 0:O4 = 0:N5 = 0:O5 = 0:N6 = 0:O6 = 0:N7 = 0:O7 = 0:N8 = 0:O8 = 0:N9 = 0:O9 = 0:N10 = 0:O10 = 0:N11 = 0:O11 = 0:N12 = 0:O12 = 0

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
    '1ere bank
    N1 = Controller.Lamp(12)
    If N1 <> O1 Then
        If N1 Then sw00.isdropped = 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
        O1 = N1
    End If
    N2 = Controller.Lamp(13)
    If N2 <> O2 Then
        If N2 Then sw10.isdropped = 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
        O2 = N2
    End If
    N3 = Controller.Lamp(14)
    If N3 <> O3 Then
        If N3 Then sw20.isdropped = 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
        O3 = N3
    End If
    N4 = Controller.Lamp(15)
    If N4 <> O4 Then
        If N4 Then sw30.isdropped = 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
        O4 = N4
    End If

    '2eme bank
    N5 = Controller.Lamp(16)
    If N5 <> O5 Then
        If N5 Then sw01.isdropped = 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
        O5 = N5
    End If
    N6 = Controller.Lamp(17)
    If N6 <> O6 Then
        If N6 Then sw11.isdropped = 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
        O6 = N6
    End If
    N7 = Controller.Lamp(18)
    If N7 <> O7 Then
        If N7 Then sw21.isdropped = 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
        O7 = N7
    End If
    N8 = Controller.Lamp(19)
    If N8 <> O8 Then
        If N8 Then sw31.isdropped = 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
        O8 = N8
    End If

    '3eme bank
    N9 = Controller.Lamp(20)
    If N9 <> O9 Then
        If N9 Then sw02.isdropped = 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
        O9 = N9
    End If
    N10 = Controller.Lamp(21)
    If N10 <> O10 Then
        If N10 Then sw12.isdropped = 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
        O10 = N10
    End If
    N11 = Controller.Lamp(22)
    If N11 <> O11 Then
        If N11 Then sw22.isdropped = 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
        O11 = N11
    End If
    N12 = Controller.Lamp(23)
    If N12 <> O12 Then
        If N12 Then sw32.isdropped = 1:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
        O12 = N12
    End If
End Sub

'Gottlieb Panthera
'added by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Panthera - DIP switches"
        .AddFrame 2, 10, 190, "Maximum credits", 49152, Array("8 credits", 0, "10 credits", 32768, "15 credits", &H00004000, "25 credits", 49152)                                                                                  'dip 15&16
        .AddFrame 2, 86, 190, "Coin chute 1 and 2 control", &H00002000, Array("seperate", 0, "same", &H00002000)                                                                                                                   'dip 14
        .AddFrame 2, 132, 190, "Playfield special", &H00200000, Array("replay", 0, "extra ball", &H00200000)                                                                                                                       'dip 22
        .AddFrame 2, 178, 190, "Hole for special", &H80000000, Array("alternating", 0, "stays lit", &H80000000)                                                                                                                    'dip32
        .AddFrame 2, 224, 190, "High score to date awards", &H00C00000, Array("not displayed and no award", 0, "displayed and no award", &H00800000, "displayed and 2 replays", &H00400000, "displayed and 3 replays", &H00C00000) 'dip 23&24
        .AddChk 2, 300, 190, Array("Sound when scoring?", &H01000000)                                                                                                                                                              'dip 25
        .AddChk 2, 315, 190, Array("Replay button tune?", &H02000000)                                                                                                                                                              'dip 26
        .AddChk 2, 330, 190, Array("Coin switch tune?", &H04000000)                                                                                                                                                                'dip 27
        .AddChk 2, 345, 190, Array("Credits displayed?", &H08000000)                                                                                                                                                               'dip 28
        .AddChk 2, 360, 190, Array("Match feature", &H00020000)                                                                                                                                                                    'dip 18
        .AddChk 2, 375, 190, Array("Attract features", &H20000000)                                                                                                                                                                 'dip 30
        .AddChkExtra 2, 390, 190, Array("Background sound off", &H0100)                                                                                                                                                            'S-board dip 1
        .AddFrameExtra 205, 10, 190, "Attract tune", &H0200, Array("no attract tune", 0, "attract tune played every 6 minutes", &H0200)                                                                                            'S-board dip 2
        .AddFrame 205, 56, 190, "Balls per game", &H00010000, Array("5 balls", 0, "3 balls", &H00010000)                                                                                                                           'dip 17
        .AddFrame 205, 102, 190, "Replay limit", &H00040000, Array("no limit", 0, "one per ball", &H00040000)                                                                                                                      'dip 19
        .AddFrame 205, 148, 190, "Novelty", &H00080000, Array("normal game mode", 0, "50,000 points for special/extra ball", &H0080000)                                                                                            'dip 20
        .AddFrame 205, 194, 190, "Game mode", &H00100000, Array("replay", 0, "extra ball", &H00100000)                                                                                                                             'dip 21
        .AddFrame 205, 240, 190, "3rd coin chute credits control", &H00001000, Array("no effect", 0, "add 9", &H00001000)                                                                                                          'dip 13
        .AddFrame 205, 286, 190, "Tilt penalty", &H10000000, Array("game over", 0, "ball in play", &H10000000)                                                                                                                     'dip 29
        .AddFrame 205, 332, 190, "Extra ball target adjust", &H40000000, Array("alternating", 0, "stays lit", &H40000000)                                                                                                          'dip 31
        .AddLabel 50, 420, 300, 20, "After hitting OK, press F3 to reset game with new settings."
    End With

    Dim extra
    extra = Controller.Dip(4) + Controller.Dip(5) * 256
    extra = vpmDips.ViewDipsExtra(extra)
    Controller.Dip(4) = extra And 255
    Controller.Dip(5) = (extra And 65280) \ 256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")

Sub Table1_Exit():Controller.Games(cGameName).Settings.Value("sound")=1:Controller.Stop:End Sub


' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
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

' set position as table object and Vol + RndPitch manually

Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
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

Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' requires rampbump1 to 7 in Sound Manager

Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'*********************************************************************
'            Supporting Ball, Sound Functions and Math
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
' Thalamus, AudioFade - Patched
  If tmp > 0 Then
    AudioFade = Csng(tmp ^5) 'was 10
  Else
    AudioFade = Csng(-((- tmp) ^5) ) 'was 10
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
' Thalamus, AudioPan - Patched
  If tmp > 0 Then
    AudioPan = Csng(tmp ^5) 'was 10
  Else
    AudioPan = Csng(-((- tmp) ^5) ) 'was 10
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
' Thalamus, Pan - Patched
  If tmp > 0 Then
    Pan = Csng(tmp ^5) 'was 10
  Else
    Pan = Csng(-((- tmp) ^5) ) ' was 10
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function VolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  VolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
End Function

Function DVolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  DVolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
  debug.print DVolMulti
End Function

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
  BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
  BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
  VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
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

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

