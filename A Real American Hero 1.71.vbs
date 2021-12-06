Option Explicit
Randomize

' Thalamus 2018-11-01 : Improved directional sounds
' Took away old code for ballrolling and collition - use the newer OnBallBallCollision
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 200     ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolTarg   = 1    ' Targets volume.
Const VolFlip   = 1    ' Flipper volume.

Const cGameName = "SexyGirl"
Const cloneGameName = "playboyb"
Const UseSolenoids=1,UseLamps=True,UseGI=0,UseSyn=1,SSolenoidOn="SolOn",SSolenoidOff="Soloff",SFlipperOn="FlipperUpLeft",SFlipperOff="FlipperDown"
Const SCoin="coin3",cCredits="A Real American Hero - OPERATION P.I.N.B.A.L.L.(Original 2018) by Mickey-Lizard- MOD by Xenonph v1.71"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01560000","Bally.vbs", 3.36
Dim bsTrough, dtDrop, bsSaucer, balls, scount, brain, h, screens(190),plungerIM, xx
Dim Language
Dim NumOfBalls
Dim CDMD

'*****************************************************
'*                     Options                       *
'*****************************************************

Const GI_ON         = 1         ' 0 or 1 to disable or enable GI

'Instruction Cards
Language = 1  ' 0 = German   1 = English
NumOfBalls = 1   ' 0 = 3ball    1 = 5ball

'ColorDMD   Default is Green
CDMD = 0   '0=Green  1=Red  2=White  3=Blue  4=Reg Orange

'*****************************************************
'*                     Solonoids                     *
'*****************************************************
SolCallback(7) = "bsTrough.SolOut"
SolCallback(6)     = "vpmSolSound ""Knocker"","
'SolCallback(12)      = "vpmSolSound ""Sling"","
'SolCallback(14)      = "vpmSolSound ""Sling"","
'SolCallback(9)        = "vpmSolSound ""Bumper"","
'SolCallback(10)        = "vpmSolSound ""Bumper"","
'SolCallback(11)        = "vpmSolSound ""Bumper"","
SolCallback(8)      = "bsSaucer.SolOut"
SolCallback(13) = "SolTargetReset"
SolCallback(19)       = "solGI"
SolCallback(sLLFlipper) = "solLFlipper"
SolCallback(sLRFlipper) = "solRFlipper"

'*****************************************************
'*                    Table Init                     *
'*****************************************************
Sub Table1_Init
    With Controller
    .GameName = cloneGameName
    .SplashInfoLine = cCredits
    .HandleKeyboard = False
    .ShowTitle = False
    .ShowDMDOnly = True
      .ShowFrame = False
        If CDMD=0 Then Controller.Games("playboyb").Settings.Value("dmd_red") = 0:Controller.Games("playboyb").Settings.Value("dmd_green") = 255:Controller.Games("playboyb").Settings.Value("dmd_blue") = 0
        If CDMD=1 Then Controller.Games("playboyb").Settings.Value("dmd_red") = 255:Controller.Games("playboyb").Settings.Value("dmd_green") = 0:Controller.Games("playboyb").Settings.Value("dmd_blue") = 0
        If CDMD=2 Then Controller.Games("playboyb").Settings.Value("dmd_red") = 255:Controller.Games("playboyb").Settings.Value("dmd_green") = 255:Controller.Games("playboyb").Settings.Value("dmd_blue") = 255
        If CDMD=3 Then Controller.Games("playboyb").Settings.Value("dmd_red") = 0:Controller.Games("playboyb").Settings.Value("dmd_green") = 0:Controller.Games("playboyb").Settings.Value("dmd_blue") = 255
        If CDMD=4 Then Controller.Games("playboyb").Settings.Value("dmd_red") = 255:Controller.Games("playboyb").Settings.Value("dmd_green") = 69:Controller.Games("playboyb").Settings.Value("dmd_blue") = 0
        Controller.Games("playboyb").Settings.Value("sound")=0
    SetBallDip
    On Error Resume Next
      .Run
      '.Hidden=0
      If Err Then MsgBox Err.Description
    On Error Goto 0
  End With
'
' Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = True
          Dim x
          x = INT(3 * RND(1) )
          Select Case x
          Case 0:PlayMusic"0joe00.mp3":ITrigger.timerinterval=60500:ITrigger.timerenabled=1
          Case 1:PlayMusic"0joe21.mp3":ITrigger.timerinterval=54000:ITrigger.timerenabled=1
          Case 2:PlayMusic"0joe20.mp3":ITrigger.timerinterval=188000:ITrigger.timerenabled=1
          End Select

' Nudging
  vpmNudge.TiltSwitch = 7
  vpmNudge.Sensitivity = 1
  vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingShot)
  Set bsTrough = New cvpmBallStack                        ' Trough handler
  bsTrough.InitSw 0,8,0,0,0,0,0,0
  bsTrough.InitKick Kickout, 60, 4
  bsTrough.Balls = 1
  Set dtDrop = New cvpmDropTarget
  dtDrop.InitDrop Array(Target1,Target2,Target3,Target4,Target5),Array(1,2,3,4,5)
  'dtDrop.InitSnd "Drop","Reset"
  Set bsSaucer = New cvpmBallStack
  bsSaucer.InitSaucer GrottoKicker,32,0,20
  'bsSaucer.InitExitSnd "Saucer","Saucer"
  if Controller.Dip(0)+Controller.Dip(1)+Controller.Dip(2)+Controller.Dip(3)  < 1 then Controller.Pause=true : msgBox "Please select your game options!" & vbNewLine & "You will only have to do this once!" : vpmShowDips : Controller.Pause=false
CheckInstructionCards

Dim DesktopMode:DesktopMode = Table1.ShowDT

If DesktopMode = True Then
    For each xx in aSDO
      xx.visible = 0
    Next

    Light1Player.visible = 1
    Light2Player.visible = 1
    Light3Player.visible = 1
    Light4Player.visible = 1
    BallInPlayLight.visible = 1
    MatchLight.visible = 1
    GameOverLight.visible = 1
    Light57.visible = 1
        Light58.visible = 1
        Light59.visible = 1
    Ramp1730.visible = True
    Ramp1729.visible = True
    Ramp1731.visible = True
    Ramp1732.visible = True
    Ramp1.visible = True
        Flasherlight5.visible= True
        Controller.Hidden=0
End If
If DesktopMode = False Then
    For each xx in aSDO
      xx.visible = 0
    Next

    Light1Player.visible = 0
    Light2Player.visible = 0
    Light3Player.visible = 0
    Light4Player.visible = 0
    BallInPlayLight.visible = 0
    MatchLight.visible = 0
    GameOverLight.visible = 0
    Light57.visible = 0
        Light58.visible = 0
        Light59.visible = 0
    Ramp1730.visible = False
    Ramp1729.visible = False
    Ramp1731.visible = False
    Ramp1732.visible = False
    Ramp1.visible = False
        Flasherlight5.visible= False
        Controller.Hidden=1
End If

End Sub

Sub SetBallDip
  With Controller
    .Dip(0) = &HE2    '.Dip(0) = &H44
    .Dip(2) = &HD1    '.Dip(2) = &HDD
    .Dip(3) = &H01    '.Dip(3) = &H01

    if balls < 4 then   ' For 3 balls per game
      .Dip(1) = &H6A
    else        ' For 5 balls per game
      .Dip(1) = &HEA
    end if
  end with
End Sub

'*****************************************************
'*               Fluppers Flashers                   *
'*****************************************************

Dim FlashLevel1, FlashLevel2, FlashLevel3, FlashLevel4, FlashLevel5
FlasherLight1.IntensityScale = 0
Flasherlight2.IntensityScale = 0
Flasherlight3.IntensityScale = 0
Flasherlight4.IntensityScale = 0
Flasherlight5.IntensityScale = 0

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

'*** blue flasher ***
Sub FlasherFlash2_Timer()
  dim flashx3, matdim
  If not Flasherflash2.TimerEnabled Then
    Flasherflash2.TimerEnabled = True
    Flasherflash2.visible = 1
    Flasherlit2.visible = 1
  End If
  flashx3 = FlashLevel2 * FlashLevel2 * FlashLevel2
  Flasherflash2.opacity = 8000 * flashx3
  Flasherlit2.BlendDisableLighting = 10 * flashx3
  Flasherbase2.BlendDisableLighting =  flashx3
  Flasherlight2.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel2)
  Flasherlit2.material = "domelit" & matdim
  FlashLevel2 = FlashLevel2 * 0.85 - 0.01
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

'*** Red flasher ***
Sub FlasherFlash3_Timer()
  dim flashx3, matdim
  If not Flasherflash3.TimerEnabled Then
    Flasherflash3.TimerEnabled = True
    Flasherflash3.visible = 1
    Flasherlit3.visible = 1
  End If
  flashx3 = FlashLevel3 * FlashLevel3 * FlashLevel3
  Flasherflash3.opacity = 1500 * flashx3
  Flasherlit3.BlendDisableLighting = 10 * flashx3
  Flasherbase3.BlendDisableLighting =  flashx3
  Flasherlight3.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel3)
  Flasherlit3.material = "domelit" & matdim
  FlashLevel3 = FlashLevel3 * 0.9 - 0.01
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

'*** Red flasher ***
Sub FlasherFlash5_Timer()
  dim flashx3, matdim
  If not Flasherflash5.TimerEnabled Then
    Flasherflash5.TimerEnabled = True
    Flasherflash5.visible = 1
    Flasherlit5.visible = 1
  End If
  flashx3 = FlashLevel5 * FlashLevel5 * FlashLevel5
  Flasherflash5.opacity = 1500 * flashx3
  Flasherlit5.BlendDisableLighting = 10 * flashx3
  Flasherbase5.BlendDisableLighting =  flashx3
  Flasherlight5.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel5)
  Flasherlit5.material = "domelit" & matdim
  FlashLevel5 = FlashLevel5 * 0.9 - 0.01
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


'***********************************************

Dim AA
AA=0

Sub Flr_Timer
    Dim x
    x = INT(4 * RND(1) )
    Select Case x
    Case 0:FlashLevel1 = 1 : FlasherFlash1_Timer
    Case 1:FlashLevel3 = 1 : FlasherFlash3_Timer
    Case 2:FlashLevel2 = 1 : FlasherFlash2_Timer
    Case 3:FlashLevel4 = 1 : FlasherFlash4_Timer
    End Select
end sub

Sub Fll_Timer
    FlashLevel3 = 1 : FlasherFlash3_Timer
end sub

Sub Flu_Timer
    If AA=0 Then FlashLevel2 = 1 : FlasherFlash2_Timer:AA=1:Exit Sub
    If AA=1 Then FlashLevel1 = 1 : FlasherFlash1_Timer:AA=0
end sub

Sub Flashoff_Timer
    Fll.enabled=false
    Flu.enabled=false
    Flr.enabled=false
    Flashoff.enabled=false
end sub


Sub GiEffect(value) ' value is the duration of the blink
    Dim x
    For each x in aGi
        x.Duration 2, 500 * value, 1
    Next
End Sub


'*****************************************************
'*                   Drop Targets                    *
'*****************************************************
Sub Target1_Hit   : PlaySoundAtVol SoundFX("Drop",DOFContactors) , ActiveBall, VolTarg: Controller.Switch(1) = True :PlaySound"0joe00zd": End Sub
Sub Target2_Hit   : PlaySoundAtVol SoundFX("Drop",DOFContactors) , ActiveBall, VolTarg: Controller.Switch(2) = True :PlaySound"0joe00zd": End Sub
Sub Target3_Hit   : PlaySoundAtVol SoundFX("Drop",DOFContactors) , ActiveBall, VolTarg: Controller.Switch(3) = True :PlaySound"0joe00zd": End Sub
Sub Target4_Hit   : PlaySoundAtVol SoundFX("Drop",DOFContactors) , ActiveBall, VolTarg: Controller.Switch(4) = True :PlaySound"0joe00zd": End Sub
Sub Target5_Hit   : PlaySoundAtVol SoundFX("Drop",DOFContactors) , ActiveBall, VolTarg: Controller.Switch(5) = True :PlaySound"0joe00zd": End Sub
Sub SolTargetReset(enabled)
  if enabled then PlaySoundAtVol SoundFX("Reset",DOFContactors), Target3, VolTarg
  dtDrop.SolDropUp enabled
End Sub

'*****************************************************
'*                Keyboard Handlers                  *
'*****************************************************
Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("nudge", 0), 0, 1, 0, 0.25
  If KeyCode=PlungerKey Then Plunger.Pullback:PlaySoundAtVol SoundFX("PlungerPull",DOFContactors), Plunger, 1
  If vpmKeyDown(KeyCode) Then Exit Sub

End Sub
Sub Table1_KeyUp(ByVal keycode)
  If KeyCode = 6 Then
            PlaySoundAtVol"coin", Drain, 1
            Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=1500
      Dim x
      x = INT(10 * RND(1) )
      Select Case x
      Case 0:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00a"
      Case 1:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00b"
      Case 2:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00c"
      Case 3:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00d"
      Case 4:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00v"
      Case 5:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00q"
      Case 6:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00r"
      Case 7:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00y"
      Case 8:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00z"
      Case 9:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joekick09a"
      End Select
      end if
  If KeyCode = 4 Then
            PlaySoundAtVol"coin", Drain, 1
            Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=1500
      Dim y
      y = INT(10 * RND(1) )
      Select Case y
      Case 0:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00a"
      Case 1:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00b"
      Case 2:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00c"
      Case 3:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00d"
      Case 4:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00v"
      Case 5:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00q"
      Case 6:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00r"
      Case 7:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00y"
      Case 8:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joe00z"
      Case 9:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00p":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joe00v":StopSound"0joe00y":StopSound"0joe00z":StopSound"0joekick09a":PlaySound"0joekick09a"
      End Select
      end if
  If KeyCode = 2 Then
    EndMusic:StopSound"0joe00a":StopSound"0joe00b":StopSound"0joe00c":StopSound"0joe00d":StopSound"0joe00v":StopSound"0joe00q":StopSound"0joe00r":StopSound"0joekick04":StopSound"0joe00zb":StopSound"0joe00zc"
        Dim v
        v = INT(3 * RND(1) )
        Select Case v
        Case 0:Flr.enabled=false:Playsound("0joekick04")
        Case 1:Flr.enabled=false:Playsound("0joe00zb")
        Case 2:Flr.enabled=false:Playsound("0joe00zc")
        End Select
  End If
  If KeyCode=PlungerKey Then Plunger.Fire:PlaySoundAtVol SoundFX("Plunger",DOFContactors), ActiveBall, 1
    If KeyCode = LeftMagnaSave  And RightMagnaSave Then ScreenTimer.Enabled = True
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'Plunger Switch

Sub Trigger1_Hit():Drain.timerenabled=0:BallReleaseGate.timerenabled=0:GTrigger.timerenabled=0:RTrigger.timerenabled=0:ITrigger.timerenabled=0:GateR.timerenabled=0:StopSound"0joe00w":StopSound"0joe00p":PlaySound"0joe00l"
              Flu.enabled=true:Flu.interval=300
              Flr.enabled=False
              Dim z
              z = INT(4 * RND(1) )
              Select Case z
              Case 0:PlayMusic"0joe01ab.mp3":GateL.timerinterval=76050:GateL.timerenabled=1
              Case 1:PlayMusic"0joe02ab.mp3":GateL.timerinterval=57400:GateL.timerenabled=1
              Case 2:PlayMusic"0joe03ab.mp3":GateL.timerinterval=34360:GateL.timerenabled=1
              Case 3:PlayMusic"0joe04ab.mp3":GateL.timerinterval=61400:GateL.timerenabled=1

              End Select
End Sub


'*****************************************************
'*                     Flippers                      *
'*****************************************************

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("FlipperUp",DOFContactors),Flipper1,VolFlip:Flipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("FlipperDown",DOFContactors),Flipper1,VolFlip:Flipper1.RotateToStart
     End If
End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("FlipperUp",DOFContactors),Flipper2,VolFlip:Flipper2.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("FlipperDown",DOFContactors),Flipper2,VolFlip:Flipper2.RotateToStart
     End If
End Sub


Sub FlipperTimer_Timer()
   LFlip.RotY = Flipper1.CurrentAngle
   RFlip.RotY = Flipper2.CurrentAngle
End Sub

'*****************************************************
'*                    Sling Shots                    *
'*****************************************************

Dim RStep, Lstep

Sub RightSlingShot_Slingshot : vpmTimer.PulseSwitch 36, 0, 0
    PlaySoundAtVol SoundFX("SlingShot",DOFContactors), sling1, 1
    PlaySound"0joe00i"
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled=0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot : vpmTimer.PulseSwitch 37, 0, 0
    PlaySoundAtVol SoundFX("SlingShot",DOFContactors), sling2, 1
    PlaySound"0joe00i"
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled=0
    End Select
    LStep = LStep + 1
End Sub

'*****************************************************
'*                       Gates                       *
'*****************************************************

Sub GateTimer_Timer()
   GateRP.RotZ = ABS(GateR.currentangle)
   GateLP.RotZ = ABS(GateL.currentangle)
End Sub

 Sub GateR_Hit():PlaySoundAtVol "gate2", ActiveBall, 1:BallReleaseGate.timerenabled=0:GTrigger.timerenabled=0

 End Sub
 Sub GateR_timer
    Dim x
    x = INT(3 * RND(1) )
    Select Case x
    Case 0:PlayMusic"0joe00.mp3":ITrigger.timerinterval=60500:ITrigger.timerenabled=1
    Case 1:PlayMusic"0joe20.mp3":ITrigger.timerinterval=188000:ITrigger.timerenabled=1
    Case 2:PlayMusic"0joe21.mp3":ITrigger.timerinterval=54000:ITrigger.timerenabled=1
    End Select
    GateR.timerenabled=0
 End sub
 Sub GateL_Hit():PlaySoundAtVol "gate2", ActiveBall, 1

 End Sub
 Sub GateL_timer
      Dim x
      x = INT(4 * RND(1) )
      Select Case x
            Case 0:PlayMusic"0joe01ab.mp3":GrottoGate.timerinterval=76050:GrottoGate.timerenabled=1
            Case 1:PlayMusic"0joe02ab.mp3":GrottoGate.timerinterval=57400:GrottoGate.timerenabled=1
            Case 2:PlayMusic"0joe03ab.mp3":GrottoGate.timerinterval=34360:GrottoGate.timerenabled=1
            Case 3:PlayMusic"0joe04ab.mp3":GrottoGate.timerinterval=61400:GrottoGate.timerenabled=1

      End Select
 GateL.timerenabled=0
 End sub

 Sub GrottoGate_Hit():PlaySoundAtVol "gate2", ActiveBall, 1:End Sub
 Sub GrottoGate_timer
      Dim x
      x = INT(4 * RND(1) )
      Select Case x
            Case 0:PlayMusic"0joe01ab.mp3":GateL.timerinterval=76050:GateL.timerenabled=1
            Case 1:PlayMusic"0joe02ab.mp3":GateL.timerinterval=57400:GateL.timerenabled=1
            Case 2:PlayMusic"0joe03ab.mp3":GateL.timerinterval=34360:GateL.timerenabled=1
            Case 3:PlayMusic"0joe04ab.mp3":GateL.timerinterval=61400:GateL.timerenabled=1
      End Select
 GrottoGate.timerenabled=0
 End sub
 Sub BallReleaseGate_Hit():PlaySoundAtVol "gate2", ActiveBall, 1:GateR.timerenabled=0:EndMusic:PlaySound"0joe00h":PlaySound"0joe00p":GTrigger.timerinterval=4000:GTrigger.timerenabled=1:ITrigger.timerenabled=0:RTrigger.timerinterval=20000:RTrigger.timerenabled=1:GiEffect 3:End Sub

'*****************************************************
'*                  Switch Handling                  *
'*****************************************************
Sub GTrigger_Hit   : PlaySoundAtVol "Rollover" , ActiveBall, 1: Controller.Switch(21) = True :GateR.timerenabled=0
            Flu.enabled=false
      Dim x
      x = INT(5 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00e"
      Case 1:PlaySound"0joe00f"
      Case 2:PlaySound"0joe00h"
      Case 3:PlaySound"0joe00u"
      Case 4:PlaySound"0joe00m"
      End Select
End Sub

Sub GTrigger_Timer:PlaySound"0joe00p":End Sub

Sub GTrigger_UnHit : Controller.Switch(21) = False : End Sub
Sub ITrigger_Hit   : PlaySoundAtVol "Rollover" , ActiveBall, 1:Controller.Switch(20) = True :GateR.timerenabled=0
            Flu.enabled=false
      Dim x
      x = INT(5 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00e"
      Case 1:PlaySound"0joe00f"
      Case 2:PlaySound"0joe00h"
      Case 3:PlaySound"0joe00u"
      Case 4:PlaySound"0joe00m"
      End Select
End Sub

 Sub ITrigger_Timer
            Dim x
            x = INT(12 * RND(1) )
            Select Case x
            Case 0:PlayMusic"0joe22.mp3":GateR.timerinterval=38000:GateR.timerenabled=1:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=38000
            Case 1:PlayMusic"0joe23.mp3":GateR.timerinterval=26000:GateR.timerenabled=1:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=26000
            Case 2:PlayMusic"0joe24.mp3":GateR.timerinterval=51000:GateR.timerenabled=1:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=51000
            Case 3:PlayMusic"0joe25.mp3":GateR.timerinterval=46000:GateR.timerenabled=1:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=46000
            Case 4:PlayMusic"0joe26.mp3":GateR.timerinterval=43000:GateR.timerenabled=1:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=43000
            Case 5:PlayMusic"0joe27.mp3":GateR.timerinterval=54000:GateR.timerenabled=1:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=54000
            Case 6:PlayMusic"0joe28.mp3":GateR.timerinterval=36000:GateR.timerenabled=1:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=36000
            Case 7:PlayMusic"0joe29.mp3":GateR.timerinterval=23000:GateR.timerenabled=1:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=23000
            Case 8:PlayMusic"0joe30.mp3":GateR.timerinterval=31000:GateR.timerenabled=1:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=31000
            Case 9:PlayMusic"0joe31.mp3":GateR.timerinterval=31000:GateR.timerenabled=1:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=31000
            Case 10:PlayMusic"0joe32.mp3":GateR.timerinterval=31000:GateR.timerenabled=1:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=31000
            Case 11:PlayMusic"0joe33.mp3":GateR.timerinterval=31000:GateR.timerenabled=1:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=31000
            End Select
            ITrigger.timerenabled=0
 End Sub

Sub ITrigger_UnHit : Controller.Switch(20) = False : End Sub
Sub RTrigger_Hit   : PlaySoundAtVol "Rollover" , ActiveBall, 1:Controller.Switch(19) = True :GateR.timerenabled=0
            Flu.enabled=false
      Dim x
      x = INT(5 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00e"
      Case 1:PlaySound"0joe00f"
      Case 2:PlaySound"0joe00h"
      Case 3:PlaySound"0joe00u"
      Case 4:PlaySound"0joe00m"
      End Select
End Sub

 Sub RTrigger_Timer:Dim x:x = INT(2 * RND(1) ):Select Case x:Case 0:PlaySound"0joe00w":Case 1:PlaySound"0joe00x":End Select:Flr.enabled=true:Flr.interval=100:Flashoff.enabled=true:Flashoff.interval=3000: End Sub

Sub RTrigger_UnHit : Controller.Switch(19) = False : End Sub
Sub LTrigger_Hit   : PlaySoundAtVol "Rollover" , ActiveBall, 1:Controller.Switch(18) = True :GateR.timerenabled=0
            Flu.enabled=false
      Dim x
      x = INT(5 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00e"
      Case 1:PlaySound"0joe00f"
      Case 2:PlaySound"0joe00h"
      Case 3:PlaySound"0joe00u"
      Case 4:PlaySound"0joe00m"
      End Select
End Sub
Sub LTrigger_UnHit : Controller.Switch(18) = False : End Sub
Sub RightOutlane_Hit   :PlaySoundAtVol "Rollover" , ActiveBall, 1:Controller.Switch(22) = True :PlaySound"0joe00j": End Sub
Sub RightOutlane_UnHit : Controller.Switch(22) = False : End Sub
Sub LeftOutlane_Hit   : PlaySoundAtVol "Rollover" , ActiveBall, 1:Controller.Switch(23) = True :PlaySound"0joe00j": End Sub
Sub LeftOutlane_UnHit : Controller.Switch(23) = False : End Sub
Sub RightInlane_Hit   : PlaySoundAtVol "Rollover" , ActiveBall, 1:Controller.Switch(24) = True
      Dim x
      x = INT(5 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00e"
      Case 1:PlaySound"0joe00f"
      Case 2:PlaySound"0joe00h"
      Case 3:PlaySound"0joe00u"
      Case 4:PlaySound"0joe00m"
      End Select
End Sub
Sub RightInlane_UnHit : Controller.Switch(24) = False : End Sub
Sub StarTrigger_Hit   : Controller.Switch(30) = True
      Dim x
      x = INT(3 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00m"
      Case 1:PlaySound"0joe00n"
      Case 2:PlaySound"0joe00o"
      End Select
End Sub
Sub StarTrigger_UnHit : Controller.Switch(30) = False : End Sub
Sub FivekTrigger_Hit   : PlaySoundAtVol "Rollover" , ActiveBall, 1:Controller.Switch(31) = True
      Dim x
      x = INT(8 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00e"
      Case 1:PlaySound"0joe00d"
      Case 2:PlaySound"0joe00h"
      Case 3:PlaySound"0joe00o"
      Case 4:PlaySound"0joe00m"
      Case 5:PlaySound"0joe00t"
      Case 6:PlaySound"0joe00u"
      Case 7:PlaySound"0joe00za"
      End Select
End Sub
Sub FivekTrigger_UnHit : Controller.Switch(31) = False : End Sub
Sub GrottoTrigger_Hit   : PlaySoundAtVol "Rollover" , ActiveBall, 1:Controller.Switch(24) = True : End Sub
Sub GrottoTrigger_UnHit : Controller.Switch(24) = False : End Sub
' Sub Drain_Hit : ClearBallID : bsTrough.AddBall Me : PlaySound "Drain" :FlashLevel4 = 1 : FlasherFlash4_Timer
Sub Drain_Hit : bsTrough.AddBall Me : PlaySoundAtVol "Drain" , Drain, 1:FlashLevel4 = 1 : FlasherFlash4_Timer
      Dim x
      x = INT(22 * RND(1) )
      Select Case x
      Case 0:EndMusic:PlaySound"0joedrain01":Drain.timerinterval=6000:Drain.timerenabled=1
      Case 1:EndMusic:PlaySound"0joedrain02":Drain.timerinterval=3000:Drain.timerenabled=1
      Case 2:EndMusic:PlaySound"0joedrain03":Drain.timerinterval=3000:Drain.timerenabled=1
      Case 3:EndMusic:PlaySound"0joedrain04":Drain.timerinterval=5000:Drain.timerenabled=1
      Case 4:EndMusic:PlaySound"0joedrain05":Drain.timerinterval=4000:Drain.timerenabled=1
      Case 5:EndMusic:PlaySound"0joedrain06":Drain.timerinterval=6000:Drain.timerenabled=1
      Case 6:EndMusic:PlaySound"0joedrain07":Drain.timerinterval=5000:Drain.timerenabled=1
      Case 7:EndMusic:PlaySound"0joedrain08":Drain.timerinterval=6000:Drain.timerenabled=1
      Case 8:EndMusic:PlaySound"0joedrain09":Drain.timerinterval=4000:Drain.timerenabled=1
      Case 9:EndMusic:PlaySound"0joedrain10":Drain.timerinterval=7000:Drain.timerenabled=1
      Case 10:EndMusic:PlaySound"0joedrain11":Drain.timerinterval=6000:Drain.timerenabled=1
      Case 11:EndMusic:PlaySound"0joedrain12":Drain.timerinterval=6000:Drain.timerenabled=1
      Case 12:EndMusic:PlaySound"0joedrain13":Drain.timerinterval=6000:Drain.timerenabled=1
      Case 13:EndMusic:PlaySound"0joedrain14":Drain.timerinterval=5000:Drain.timerenabled=1
      Case 14:EndMusic:PlaySound"0joedrain15":Drain.timerinterval=6000:Drain.timerenabled=1
      Case 15:EndMusic:PlaySound"0joedrain16":Drain.timerinterval=4000:Drain.timerenabled=1
      Case 16:EndMusic:PlaySound"0joedrain17":Drain.timerinterval=6000:Drain.timerenabled=1
      Case 17:EndMusic:PlaySound"0joedrain18":Drain.timerinterval=5000:Drain.timerenabled=1
      Case 18:EndMusic:PlaySound"0joedrain19":Drain.timerinterval=4000:Drain.timerenabled=1
      Case 19:EndMusic:PlaySound"0joedrain20":Drain.timerinterval=6000:Drain.timerenabled=1
      Case 20:EndMusic:PlaySound"0joedrain21":Drain.timerinterval=5000:Drain.timerenabled=1
      Case 21:EndMusic:PlaySound"0joedrain22":Drain.timerinterval=9000:Drain.timerenabled=1
      End Select

BallReleaseGate.timerenabled=1

GrottoGate.timerenabled=0
GateL.timerenabled=0

End Sub

Sub drain_timer
    Dim x
    x = INT(2 * RND(1) )
    Select Case x
    Case 0:PlayMusic"0joe18.mp3":ITrigger.timerinterval=43000:ITrigger.timerenabled=1
    Case 1:PlayMusic"0joe19.mp3":ITrigger.timerinterval=42000:ITrigger.timerenabled=1
    End Select
Drain.timerenabled=0
End sub

Sub GrottoKicker_Hit : bsSaucer.AddBall 0
            GiEffect 8
            Fll.enabled=true:Fll.interval=400
      Dim x
      x = INT(42 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joekick01"
      Case 1:PlaySound"0joekick02"
      Case 2:PlaySound"0joekick03"
      Case 3:PlaySound"0joekick04"
      Case 4:PlaySound"0joekick05"
      Case 5:PlaySound"0joekick06"
      Case 6:PlaySound"0joekick07"
      Case 7:PlaySound"0joekick08"
      Case 8:PlaySound"0joekick09"
      Case 9:PlaySound"0joekick10"
      Case 10:PlaySound"0joekick11"
      Case 11:PlaySound"0joekick12"
      Case 12:PlaySound"0joekick13"
      Case 13:PlaySound"0joekick14"
      Case 14:PlaySound"0joekick15"
      Case 15:PlaySound"0joekick16"
      Case 16:PlaySound"0joekick17"
      Case 17:PlaySound"0joekick18"
      Case 18:PlaySound"0joekick19"
      Case 19:PlaySound"0joekick20"
      Case 20:PlaySound"0joekick21"
      Case 21:PlaySound"0joekick22"
      Case 22:PlaySound"0joekick23"
      Case 23:PlaySound"0joekick24"
      Case 24:PlaySound"0joekick25"
      Case 25:PlaySound"0joekick26"
      Case 26:PlaySound"0joekick27"
      Case 27:PlaySound"0joekick28"
      Case 28:PlaySound"0joekick29"
      Case 29:PlaySound"0joekick30"
      Case 30:PlaySound"0joekick31"
      Case 31:PlaySound"0joekick32"
      Case 32:PlaySound"0joekick33"
      Case 33:PlaySound"0joekick34"
      Case 34:PlaySound"0joekick35"
      Case 35:PlaySound"0joekick36"
      Case 36:PlaySound"0joekick37"
      Case 37:PlaySound"0joekick38"
      Case 38:PlaySound"0joekick39"
      Case 39:PlaySound"0joekick40"
      Case 40:PlaySound"0joekick41"
      Case 41:PlaySound"0joekick42"
      End Select
End Sub
Sub GrottoKicker_UnHit :Fll.enabled=False:Flu.enabled=true:Flu.interval=300:PlaySoundAtVol SoundFX("ballrelease",DOFContactors), ActiveBall, 1

End Sub


Sub KickOut_UnHit :PlaySoundAtVol SoundFX("ballrelease",DOFContactors), ActiveBall, 1:Drain.timerenabled=0: End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                      Switch Handling                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DropTargetSlingshot_Slingshot : vpmTimer.PulseSwitch 33, 0, 0 : End Sub
Sub Bumper2_Hit : PlaySoundAtVol SoundFX("Bumper",DOFContactors),Bumper2,VolBump:vpmTimer.PulseSwitch 38, 0, 0 :PlaySound"0joe00k":FlashLevel1 = 1 : FlasherFlash1_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer: End Sub
Sub Bumper1_Hit : PlaySoundAtVol SoundFX("Bumper",DOFContactors),Bumper1,VolBump:vpmTimer.PulseSwitch 39, 0, 0 :PlaySound"0joe00k":FlashLevel2 = 1 : FlasherFlash2_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer: End Sub
Sub Bumper3_Hit : PlaySoundAtVol SoundFX("Bumper",DOFContactors),Bumper3,VolBump:vpmTimer.PulseSwitch 40, 0, 0 :PlaySound"0joe00k":FlashLevel1 = 1 : FlasherFlash1_Timer:FlashLevel2 = 1 : FlasherFlash2_Timer:FlashLevel5 = 1 : FlasherFlash5_Timer: End Sub

'*****************************************************
'*                Model And S Targets                *
'*****************************************************
 Sub sw17_Hit:vpmTimer.PulseSw 17:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("Target",DOFContactors), ActiveBall, 1:
  If h = True Then
    ScreenTimer.Enabled = True:GiEffect 2
  End If
      Dim x
      x = INT(8 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00s"
      Case 1:PlaySound"0joe00a"
      Case 2:PlaySound"0joe00k"
      Case 3:PlaySound"0joe00b"
      Case 4:PlaySound"0joe00v"
      Case 5:PlaySound"0joe00c"
      Case 6:PlaySound"0joe00s"
      Case 7:PlaySound"0joe00y"
      End Select
FlashLevel3 = 1 : FlasherFlash3_Timer
End Sub
 Sub sw25_Hit:vpmTimer.PulseSw 25:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("Target",DOFContactors), ActiveBall, 1:
  If h = True Then
    ScreenTimer.Enabled = True:GiEffect 2
  End If
      Dim x
      x = INT(8 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00s"
      Case 1:PlaySound"0joe00a"
      Case 2:PlaySound"0joe00k"
      Case 3:PlaySound"0joe00b"
      Case 4:PlaySound"0joe00v"
      Case 5:PlaySound"0joe00c"
      Case 6:PlaySound"0joe00s"
      Case 7:PlaySound"0joe00y"
      End Select
FlashLevel3 = 1 : FlasherFlash3_Timer
End Sub
 Sub sw26_Hit:vpmTimer.PulseSw 26:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("Target",DOFContactors), ActiveBall, 1:
  If h = True Then
    ScreenTimer.Enabled = True:GiEffect 2
  End If
      Dim x
      x = INT(8 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00s"
      Case 1:PlaySound"0joe00a"
      Case 2:PlaySound"0joe00k"
      Case 3:PlaySound"0joe00b"
      Case 4:PlaySound"0joe00v"
      Case 5:PlaySound"0joe00c"
      Case 6:PlaySound"0joe00s"
      Case 7:PlaySound"0joe00y"
      End Select
FlashLevel3 = 1 : FlasherFlash3_Timer
End Sub
 Sub sw27_Hit:vpmTimer.PulseSw 27:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("Target",DOFContactors), ActiveBall, 1:
  If h = True Then
    ScreenTimer.Enabled = True:GiEffect 2
  End If
      Dim x
      x = INT(8 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00s"
      Case 1:PlaySound"0joe00a"
      Case 2:PlaySound"0joe00k"
      Case 3:PlaySound"0joe00b"
      Case 4:PlaySound"0joe00v"
      Case 5:PlaySound"0joe00c"
      Case 6:PlaySound"0joe00s"
      Case 7:PlaySound"0joe00y"
      End Select
FlashLevel3 = 1 : FlasherFlash3_Timer
End Sub
 Sub sw28_Hit:vpmTimer.PulseSw 28:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("Target",DOFContactors), ActiveBall, 1:
  If h = True Then
    ScreenTimer.Enabled = True:GiEffect 2
  End If
      Dim x
      x = INT(8 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00s"
      Case 1:PlaySound"0joe00a"
      Case 2:PlaySound"0joe00k"
      Case 3:PlaySound"0joe00b"
      Case 4:PlaySound"0joe00v"
      Case 5:PlaySound"0joe00c"
      Case 6:PlaySound"0joe00s"
      Case 7:PlaySound"0joe00y"
      End Select
FlashLevel3 = 1 : FlasherFlash3_Timer
End Sub
 Sub sw29_Hit:vpmTimer.PulseSw 29:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("Target",DOFContactors), ActiveBall, 1:
  If h = True Then
    ScreenTimer.Enabled = True:GiEffect 2
  End If
      Dim x
      x = INT(8 * RND(1) )
      Select Case x
      Case 0:PlaySound"0joe00s"
      Case 1:PlaySound"0joe00a"
      Case 2:PlaySound"0joe00k"
      Case 3:PlaySound"0joe00b"
      Case 4:PlaySound"0joe00v"
      Case 5:PlaySound"0joe00c"
      Case 6:PlaySound"0joe00s"
      Case 7:PlaySound"0joe00y"
      End Select
FlashLevel3 = 1 : FlasherFlash3_Timer
End Sub

'*****************************************************
'*               Projector Handling                  *
'*****************************************************
balls = 3
scount = 388    ' How many different screens?
brain = 1   ' Contains the present screen
h = False   ' Projector off


Sub ScreenTimer_Timer         'here are the changing of pics
' screens(brain - 1).IsDropped = True
    If brain = scount Then
      brain = 1
    End If
' screens(brain).IsDropped = False
  Projector.Image = "pic"&brain
  ScreenTimer.Enabled = False
  brain = brain + 1
End Sub

Sub ProjectorOff()            'cut-off projector
' screens(brain-1).IsDropped = True
' screens(0).IsDropped = False
  Projector.Image = "pic0"
  brain = 1             'reset to first pic at cut-off (like original???)
  h = False
End Sub

Sub ProjectorOn()           'switch projector on
  h = True
End Sub

'*****************************************************
'*                 Instruction Cards                 *
'*****************************************************
Sub CheckInstructionCards()
  If Language = 0 Then
    InstructionsL.image = "IGL"
    If NumOfBalls = 0 Then
      InstructionsR.image = "IGR3-ball"
    Else
      InstructionsR.image = "IGR5-ball"
    End If
  Else
    InstructionsL.image = "IEL"
    If NumOfBalls = 0 Then
      InstructionsR.image = "IER3-ball"
    Else
      InstructionsR.image = "IER5-ball"
    End If
  End If
End Sub



'*****************************************************
'*                      Lights                       *
'*****************************************************
Set Lights(15)  = Light1Player
Set Lights(31)  = Light2Player
Set Lights(47)  = Light3Player
Set Lights(63)  = Light4Player
Set Lights(13)  = BallInPlayLight
Set Lights(1)  = Light1k
Set Lights(17) = Light2k
Set Lights(33) = Light3k
Set Lights(49) = Light14k
Set Lights(2)  = Light15k
Set Lights(18) = Light16k
Set Lights(34) = Light17k
Set Lights(50) = Light18k
Set Lights(3)  = Light19k
Set Lights(19) = Light10k
Set Lights(35) = Light120k
Set Lights(55) = Light2X
Set Lights(39) = Light3X
Set Lights(23) = Light5X
Set Lights(4)  = GLight
Set Lights(20) = ILight
Set Lights(36) = RLight
Set Lights(52) = LLight
Set Lights(5)  = SLight
Set Lights(6)  = KickerGLight1
Set Lights(22) = KickerILight1
Set Lights(38) = KickerRLight1
Set Lights(54) = KickerLLight1
Set Lights(7)  = KickerSLight1
Set Lights(21) = Kicker25kLight
Set Lights(40) = JeanLight
Set Lights(8)  = JoyLight
Set Lights(24) = LinnLight
Set Lights(56) = SusyLight
Set Lights(9)  = EvaLight
Set Lights(37) = Arrow5kLight
Set Lights(57) = ArrowExtraLight
Set Lights(41) = ArrowSpecialLight
Set Lights(25) = TriggerLight
Set Lights(44) = RightOutlaneLight
Set Lights(60) = LeftOutlaneLight
Set Lights(53) = TargetSpecialLight
Set Lights(51) = GirlsSpecialLight
Set Lights(43) = ExtraballLight
Set Lights(59) = CreditLight
Set Lights(45) = GameOverLight
Set Lights(11) = ExtraballLight2
Set Lights(29) = HighscoreLight
Set Lights(27) = MatchLight
Set Lights(61) = TiltLight
Set LampCallBack = GetRef("UpdateMultipleLamps")

Public Sub UpdateMultipleLamps

If GameOverLight.State = 0 Then
   BallInPlayLight.State = 1
   Else
   BallInPlayLight.State = 0
End If
If  GameOverLight.State = 1 Then
  ProjectorOff()
End If
If  Light2X.State = 1 Then
  ProjectorOn()
'Else
  'ProjectorOff()                 'mode of operation "B"??
End If
If  Light3X.State = 1 Then
  ProjectorOn()
'Else
  'ProjectorOff()             'mode of operation "B"??
End If
If  Light5X.State = 1 Then
  ProjectorOn()
'Else
  'ProjectorOff()             'mode of operation "B"??
End If
End Sub

'*****************************************************
'*                    GI Lights                      *
'*****************************************************
Sub solGI (enabled)
  Dim xx
  debug.print enabled
  If GI_ON = 1 Then
    For each xx in aGI
      xx.state = enabled
    Next
  Else
    For each xx in aGI
      xx.state = 1
    Next
  End If
End Sub

'*****************************************************
'*                Dip Switch Settings                *
'*****************************************************

Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    'How to implent projector mode of operation ("A" or "B") to Dips ???
    .AddForm  370,500,"Sexy Girl - DIP switch settings"
    .AddFrame   2,  5, 115,"Balls per game",32768,Array("3 balls",0,"5 balls",32768)
    .AddFrame   2, 53, 115,"Credits Display",&H00080000,Array("On",&H00080000,"Off",0)
    .AddFrame   2,102, 115,"Match feature",&H00100000,Array("On",&H00100000,"Off",0)
    .AddFrame   2,150, 115,"Drop Target Special",&H00200000,Array("Lit until collected",&H200000,"Lit until ball lost",0)
    .AddFrame   2,198, 115,"5 Girls",&H00400000,Array("Held until made",&H400000,"Not held",0)
    .AddFrame   2,248, 115,"G && L Triggers",&H20000000,Array("Tied together",&H20000000,"Not tied",0)
    .AddFrame   2,298, 115,"I && R Triggers",&H10000000,Array("Tied together",&H10000000,"Not tied",0)
    .AddFrame   2,348, 115,"Outlanes",&H00800000,Array("Both lit",&H800000,"Alternating",0)
    .AddFrame   2,398, 115,"Sounds-Scoring", &H80, Array("Chime",0,"Noise",&H80)
    .AddFrame   2,448, 115,"Sounds-Coin (no credit)",&H80000000,Array("Chime",0,"Noise",&H80000000)
    .AddFrame 130,  5, 118,"High Score Award",&H00006000,Array("Replay",&H6000,"Extra Ball",&H4000,"No Award",0)
    .AddFrame 130, 75, 118,"Rollover Extra && Special",&H40000000,Array("Hold until made",&H40000000,"Not Held",0)
    .AddFrame 130,130, 118,"High game to date",&H00000060,Array("No award",0,"1 credit",&H20,"2 credits",&H40,"3 credits",&H60)
    .AddFrame 130,211, 118,"Max. credits",&H00070000,Array("5 credits",0,"10 credits",&H10000,"15 credits",&H20000,"20 credits",&H30000,"25 credits",&H40000,"30 credits",&H50000,"35 credits",&H60000,"40 credits",&H70000)
    .AddFrame 130,350, 118,"Coin Slot 2 (Key 3)",&H0F000000,Array("Same as Coin 1",&H0000000,"1/coin",&H1000000,"2/coin",&H2000000,_
                 "3/coin",&H3000000,"4/coin",&H4000000,"5/coin",&H5000000,"6/coin",&H6000000,"7/coin",&H7000000,_
                 "8/coin",&H8000000)
    .AddFrame 260,  5, 90,"Coin Slot 1 (Key 4)",&H000001F,Array("3/2 coins",&H00,"1 coin",&H02,"1/2 coins",&H03,"2/coin",&H04,"2/2 coins",&H05,_
                "3/coin",&H06,"3/2 coins",&H07,"4/coin",&H08,"4/2 coins",&H09,"5/coin",&H0A,"5/2 coins",&H0B,_
                "6/coin",&H0C,"6/2 coins",&H0D,"7/coin",&H0E,"7/2 coins",&H0F)
    .AddFrame 260,265, 90,"Coin Slot 3 (Key 5)",&H1F00,Array("3/2 coins",&H0000,"1 coin",&H0200,"1/2 coins",&H0300,"2/coin",&H0400,"2/2 coins",&H05000000,_
                "3/coin",&H0600,"3/2 coins",&H0700,"4/coin",&H0800,"4/2 coins",&H0900,"5/coin",&H0A00,"5/2 coins",&H0B00,_
                "6/coin",&H0C00,"6/2 coins",&H0D00,"7/coin",&H0E00,"7/2 coins",&H0F00)
    .ViewDips
  End With
End Sub
Set vpmShowDips = GetRef("editDips")

'*****************************************************
'*               DesktopMode Displays                *
'*****************************************************


Dim DisplayPatterns(11)
Dim DigStorage(32)
Dim LEDMod, ARMod

Sub CheckAR
  If Table1.BackdropImage = "backglass-4-3" Then
    LEDMod=22:ARMod=2
  Else
    LEDMod=0:ARMod=0
  End If
  For each obj in aSDO:obj.setvalue(LEDMod):Next
  For each obj in aSDO2:obj.setvalue(LEDMod):Next
  For each obj in aReels:obj.setvalue(ARMod):Next
End Sub

'Binary/Hex Pattern Recognition Array
DisplayPatterns(0) = 0    '0000000 Blank
DisplayPatterns(1) = 63   '0111111 zero
DisplayPatterns(2) = 6    '0000110 one
DisplayPatterns(3) = 91   '1011011 two
DisplayPatterns(4) = 79   '1001111 three
DisplayPatterns(5) = 102  '1100110 four
DisplayPatterns(6) = 109  '1101101 five
DisplayPatterns(7) = 125  '1111101 six
DisplayPatterns(8) = 7    '0000111 seven
DisplayPatterns(9) = 127  '1111111 eight
DisplayPatterns(10)= 111  '1101111 nine

Sub DisplayTimer_Timer ' 7-Digit output
  On Error Resume Next
  Dim ChgLED,ii,chg,stat,obj,TempCount,temptext,adj

  ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF) 'hex of binary (display 111111, or first 6 digits)
  If Not IsEmpty(ChgLED) Then
    For ii = 0 To UBound(ChgLED)
      chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      For TempCount = 0 to 10
        If stat = DisplayPatterns(TempCount) then
          aSDO(chgLED(ii, 0)).SetValue(TempCount+LEDMod)
          DigStorage(chgLED(ii, 0)) = TempCount
        End If
        If stat = (DisplayPatterns(TempCount) + 128) then
          aSDO(chgLED(ii, 0)).SetValue(TempCount+11+LEDMod)
          DigStorage(chgLED(ii, 0)) = TempCount
        End If
      Next
    Next
  End IF
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

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingSoundTimer_Timer()
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

Sub Table1_Exit()
      Controller.Games("playboyb").Settings.Value("sound")=1
      Controller.Games("playboyb").Settings.Value("dmd_red") = 255
      Controller.Games("playboyb").Settings.Value("dmd_green") = 69
      Controller.Games("playboyb").Settings.Value("dmd_blue") = 0
      Controller.Stop
End Sub

