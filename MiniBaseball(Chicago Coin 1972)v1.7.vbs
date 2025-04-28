' A crude Recreation of Chicago Coins Mini Baseball
' by TGX 2023.
' Version 1.7
'   - High Score Reset Function (f6 Key) fix
'   - Improvement to skill gage lamps (cosmetic to accomadate dual playfields)
'   - Graphic adjustments to overlay to support dual playfields
'   - Script changes to support mechanical plunger lever animation
'   - Object changes to support mechanical Plunger
'   - Improvements for sound when ball does not reach field of play
'   - Added ball return trigger (cosmetic emulation of original cabinet)
'   - Marquee adjustments
'   - Run Scored Ding Added per original
'   - More original score incrementing
' Version 1.6
'     - New option for cabinet art switch! Go for the fence with the full Lounge Lizard treatment!
'   - Fix for missing 'zero' in high score tape
'   - Fix for edge case in plunger sound
' Version 1,5
'   - Launcher active sound fix
'   - New cabinet lighting activates with Game Start/End
'   - Increased size of High Score tape
'   - New Skill Gage lighting
'   - Improved Playfield Art and cabinet by Lounge Lizard
'   - Revamped Out lamps based on new art with new illumination.
'   - 10 Run arrows now correctly cancelled if Out is hit.
'.
' Note: 'Gage' spelling is from the original game. 'Gage' is an alternative
' spelling of "Gauge', pronounced the same but 'Gage' is rarely used...
' I use it in the script below to adhere to the usage on the game playfield
' and to pass on a bit of trivia about the game..
' Thanks to JPSalas for inspiration and bearing with my inane coding questions.
' Some LUT's included credit goes to 3rdAxis, Fleep and Skitso
' High Score code by Darquayle
' Animated Shooter by STAT
' All new Reel-O-Matic
' Thanks to Kiwi for mechanical plunger script suggestions

'***Note: There is a bug in VPX regarding reels, which still exists as ofAugust 2023), which requires extra coding and graphic work to rotate the
'built in VPX reesls. The issue occurs when the table is rotated to 270 which is what most cabinet owners run.I have reported the flaw to Toxie in
'hopes that it will be fixed in a future VP release.
'At this time, I have implemented what I call the Reel-O-Matic which works around the problems with reel rotation. It simply uses a series
'of flashers which rotate properly with the table instead of the built in reel system which does not.This eliminates one half of the graphics
'and code required to support the errant VPX code.

'Another bug: line660 is completely ignored, no idea why.

' The position of the reel flashers may need adjustment for your specific case.....

'Initialize Variables
Dim GameStarted : GameStarted=False     'Game Started?
Dim Ready: Ready=0              'Trigger plunger sound
Dim GameInit : GameInit=False       'Used to Initialize Game
Dim TotalOuts : TotalOuts=0         'Track Number of Outs
Dim Accel : Accel=877            'Sets a speed for plunger action based on pull back
Dim SkillGageLampOff : SkillGageLampOff=0  'Skill Gage Lamp Control
Dim RunnerOnFirst : RunnerOnFirst=0     'Runner on First Status  1=True
Dim RunnerOnSecond : RunnerOnSecond=0   'Runner on Second Status 1=True
Dim RunnerOnThird : RunnerOnThird=0     'Runner on Third Status  1=True
Dim LtHomeRunLit : LtHomeRunLit=0     'Home Run Lane Light
Dim RtHomeRunLit : RtHomeRunLit=0       'Home Run Lane Light
Dim Itsahit : Itsahit=0           'Records Type of Hit 4=Home Run 3=Triple 2=Double 1=Single
Dim Runners : Runners=0           'Tracks number of runners on Base
Dim Sngle : Sngle=False           'Used to Trigger Base Runner Lamps
Dim Dbl : Dbl=False             'Used to Trigger Base Runner Lamps
Dim Trpl : Trpl=False           'Used to Trigger Base Runner Lamps
Dim HomeRun : HomeRun=False         'Used to Trigger Base Runner Lamps
Dim TheScore : TheScore=0         'Keep track of the score
Dim luts                  'Array of LUT settings
Dim lutpos                  'Index position of luts
Dim Power                 'Used to Calculate Ball Power
Dim bpower : bpower = 0
Dim Rscore : Rscore=0           'Used for Single Digit Reel Increment
Dim TensScore : TensScore = 0       'Used for Tens Reel Increment
Dim HndsScore : HndsScore = 0       'Used for Hundreds Reel Increment
Dim Bonus : Bonus =0            'Track 10 Run Bonus

'Options
Dim B2SEnable : B2SEnable = 0       'Set to 0 to disable B2S
Dim GameType: GameType=1          '1 for standard 100 point game 2 for 999 point game
Dim LL : LL=0               'Set to 1 for Lounge LIzard enhanced cabinet
Dim PlungerEnabled : PlungerEnabled=0   'Set to 1 to enable lever animation with mechanical plunger. 0 for keyboard

' darquayle OPTIONS:
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim HSArray
Dim HSScore100, HSScore10, HSScore1, HSScorex 'Define 3 different score values for each reel to use
dim hiscore
dim DTCount, hsa1, hsa2, hsa3
dim hisc
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")
Const hsFlashDelay = 4
Const pKey = 25 'Key to turn on/off PlungerKeyButton (P key is 25).  Useful for cabinets vs desktops when using real plunger vs. plunger button
Const ResetHSKey = 64 'Key to reset the high score values (F6 key is 64)
Const hisc_default=82 'Default High Score 82
Const HSA1_default=20 'Default Initials T
Const HSA2_default=7  'Default Initials G
Const HSA3_default=24 'Default Initials X
'/darquayle

'Lighting adjustments
luts = array("Fleep Natural Dark 1", "Fleep Natural Dark 2", "Fleep Warm Bright", "Fleep Warm Dark", "3rdaxis Referenced THX Standard", "Skitso Natural and Balanced", "Skitso Natural High Contrast")
'lutpos = 0           '  Set a default LUT? 0-10. This can be set and Magnasave disabled if desired.
Const EnableMagnasave = 1   '  Use magnasave button for LUT Adjustment?
SetLUT

'B2S
If B2SEnable = 1 Then
  Const cController = 1
  Dim Controller
  Set Controller = CreateObject("B2S.Server")
  Controller.B2SName = "Mini Baseball"
  Controller.Run
End If

'Set the starting table values
Sub Table1_Init
  If PlungerEnabled=0 Then
    ShooterAnimate.Enabled = False
  End If
  GageBarometer.CreateBall
  GageBarometer.Kick 180,10
  GageBarometer.Enabled=0
    BallStart.CreateBall
  GameOver.Visible=1
  GameOverLight.State=1
  If LL=1 Then
    Dim l: l=0
    For each l in Enh:l.Visible = 1: Next
    For each l in HSTape:l.Visible = 0: Next
  End If

  If LL=0 Then
    Dim i: i=0
    For each i in Enh:i.Visible = 0: Next
    lowercab.Visible=1
    lowercaboff.Visible=1
  End If
  loadhs    '*LOADS SAVED OPTIONS AND HIGH SCORE IF EXISTING
  UpdatePostIt  '*UPDATE HIGH SCORE STICKY
  Light022.State=0
  Light023.State=0
  Light024.State=0
  diverter.IsDropped = 1
End Sub

'Define Keys and Actions

Sub Table1_KeyDown(ByVal keycode)
    If Keycode=2 AND GameStarted = False And Not HSEnterMode=true Then
        NewGame
    End If
  'F6 to Reset High Score
  If keycode = ResetHSKey Then
    hisc = hisc_default
    TheScore = hisc_default
    HSA1 = HSA1_default
    HSA2 = HSA2_default
    HSA3 = HSA3_default
    savehs
    UpdatePostIt
  End If
  'End Reset
  If keycode = PlungerKey Then
    ShooterAnimate.Enabled = True
  End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
  Dim l : l=0
  If keycode = PlungerKey Then
  If PlungerEnabled =0 Then
    bpower=0
    ShooterAnimate.Enabled = False
    Plunger.Fire
    GageMeter.Fire
    Lever.Rotz = 0
  End If
    For each l in PFGageL:l.Visible = 0: Next
  End If

' ****** Watch for High Score Entry Mode ****
  If HSEnterMode Then HighScoreProcessKey(keycode) : end if

' ****** LUT Keydown ************************
  If keycode = RightMagnaSave then
    lutpos = lutpos + 1 : If lutpos > ubound(luts) Then lutpos = 0 : end if
        call SetLut
    playsound "Ding3"
    Debug.Print(luts(lutpos))
  End if

  If keycode = LeftMagnaSave then
    lutpos = lutpos - 1 : If lutpos < 0 Then lutpos = ubound(luts) : end if
        call SetLut
    playsound "Ding3"
    End If
End Sub

Sub ShooterAnimate_Timer()
  Lever.RotZ = Plunger.Position/25*90
    If PlungerEnabled = 0 Then
      bpower=bpower+1
      Lever.RotAndTra2=bpower
      Plunger.PullBack
      GageMeter.Pullback
        If bpower > 80 Then
          ShooterAnimate.Enabled = False
        End If
    End If
  Plunger.FireSpeed=Accel+Plunger.Position*Power
End Sub

Sub BaseRunner(Itsahit)
    Select Case Itsahit
      Case 1 ' ******************* Single ********************************
        Sngle=True
        If RunnerOnFirst=1 AND RunnerOnSecond=1 AND RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            TBTimer1.Enabled=1
            HBTimer1.Enabled=1
            RunnerOnFirst=1
            RunnerOnSecond=1
            RunnerOnThird=1
            Runners=3
            Exit Sub
        End If
        If RunnerOnFirst=1 AND RunnerOnSecond=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            TBTimer1.Enabled=1
            RunnerOnFirst=1
            RunnerOnSecond=1
            RunnerOnThird=1
            Runners=3
            Exit Sub
        End If
        If RunnerOnFirst=1 AND RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            HBTimer1.Enabled=1
            RunnerOnFirst=1
            RunnerOnSecond=1
            RunnerOnThird=0
            Runners=2
            Exit Sub
        End If
        If RunnerOnSecond=1 AND RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            TBTimer1.Enabled=1
            HBTimer1.Enabled=1
            RunnerOnFirst=1
            RunnerOnSecond=0
            RunnerOnThird=1
            Runners=2
            Exit Sub
        End If
        If RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            HBTimer1.Enabled=1
            RunnerOnFirst=1
            RunnerOnSecond=0
            RunnerOnThird=0
            Runners=1
            Exit Sub
        End If
        If RunnerOnSecond=1 Then
            FBTimer1.Enabled=1
            TBTimer1.Enabled=1
            RunnerOnFirst=1
            RunnerOnSecond=0
            RunnerOnThird=1
            Runners=2
            Exit Sub
        End If
        If RunnerOnFirst=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            RunnerOnFirst=1
            RunnerOnSecond=1
            Runners=2
            Exit Sub
        End If
        If Runners=0 Then
            FBTimer1.Enabled=1
            RunnerOnFirst=1
            RunnerOnSecond=0
            RunnerOnThird=0
            Runners=1
            Exit Sub
        End If
      Case 2 '********************* Double ******************************
        Dbl=True
        If RunnerOnFirst=1 AND RunnerOnSecond=1 AND RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            WaitTimerTB.Enabled=1
            HBTimer1.Enabled=1
            TBTimer1.Enabled=1
            WaitTimerHB.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=1
            RunnerOnThird=1
            Runners=2
            Exit Sub
        End If
        If RunnerOnFirst=1 and RunnerOnSecond=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            TBTimer1.Enabled=1
            WaitTimerHB.Enabled=1
            WaitTimerTB.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=1
            RunnerOnThird=1
            Runners=2
            Exit Sub
        End If
        If RunnerOnFirst=1 AND RunnerOnThird=1 Then
            HBTimer1.Enabled=1
            SBTimer1.Enabled=1
            FBTimer1.Enabled=1
            WaitTimerTB.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=1
            RunnerOnThird=1
            Runners=2
            Exit Sub
        End If
        If RunnerOnSecond=1 AND RunnerOnThird=1 Then
            HBTimer1.Enabled=1
            FBTimer1.Enabled=1
            WaitTimerTBQ.Enabled=1
            HBTimer1.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=1
            RunnerOnThird=0
            Runners=1
            Exit Sub
        End If
        If RunnerOnThird=1 Then
            FBTImer1.Enabled=1
            HBTimer1.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=1
            RunnerOnThird=0
            Runners=1
            Exit Sub
        End If
        If RunnerOnSecond=1 Then
            FBTImer1.Enabled=1
            TBTimer1.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=1
            RunnerOnThird=0
            Runners=1
            Exit Sub
        End If
        If RunnerOnFirst=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            WaitTimerTB.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=1
            RunnerOnThird=1
            Runners=2
            Exit Sub
        End If
        If Runners=0 Then
            FBTimer1.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=1
            RunnerOnThird=0
            Runners=1
            Exit Sub
        End If
      Case 3 '*************************** Triple *************************
        Trpl=True
        If RunnerOnFirst=1 AND RunnerOnSecond=1 AND RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            TBTimer1.Enabled=1
            HBTimer1.Enabled=1
            SBTimer1.Enabled=1
            TBTimer1.Enabled=1
            HBTimer1.Enabled=1
            WaitTimerHBL.Enabled=1
            WaitTimerTBRL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=1
            Runners=1
            Exit Sub
        End If
      If RunnerOnFirst=1 AND RunnerOnSecond=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            TBTimer1.Enabled=1
            SBTimer1.Enabled=1
            TBTimer1.Enabled=1
            HBTimer1.Enabled=1
            WaitTimerHBL.Enabled=1
            WaitTimerTBRL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=1
            Runners=1
            Exit Sub
        End If
      If RunnerOnFirst=1 AND RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            HBTimer1.Enabled=1
            SBTimer1.Enabled=1
            TBTimer1.Enabled=1
            HBTimer1.Enabled=1
            TBTimer1.Enabled=1
            WaitTimerHBL.Enabled=1
            WaitTimerTBRL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=1
            Runners=1
            Exit Sub
        End If
      If RunnerOnSecond=1 and RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            HBTimer1.Enabled=1
            TBTimer1.Enabled=1
            WaitTimerHBL.Enabled=1
            WaitTimerTBRL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=1
            Runners=1
            Exit Sub
        End If
        If RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            HBTimer1.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=1
            Runners=1
            Exit Sub
        End If
        If RunnerOnSecond=1 Then
            FBTimer1.Enabled=1
            TBTimer1.Enabled=1
            WaitTimerTBRL.Enabled=1
            WaitTimerHBL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=1
            Runners=1
            Exit Sub
        End If
        If RunnerOnFirst=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            WaitTimerTBL.Enabled=1
            WaitTimerHBL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=1
            Runners=1
            Exit Sub
        End If
        If Runners=0 Then
            FBTimer1.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=1
            Runners=1
            Exit Sub
        End If
      Case 4 '******************** HomeRun ********************************
        HomeRun=True
        If RunnerOnFirst=1 AND RunnerOnSecond=1 AND RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            TBTimer1.Enabled=1
            HBTimer1.Enabled=1
            WaitTimerHB.Enabled=1
            WaitTimerTB.Enabled=1
            WaitTimerSB.Enabled=1
            WaitTimerHBRM.Enabled=1
            WaitTimerTBRL.Enabled=1
            WaitTimerHBRL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=0
            Runners=0
            HomeRun=False
            Exit Sub
        End If
        If RunnerOnFirst=1 and RunnerOnSecond=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            TBTimer1.Enabled=1
            WaitTimerHB.Enabled=1
            WaitTimerTB.Enabled=1
            WaitTimerHBRM.Enabled=1
            WaitTimerTBRL.Enabled=1
            WaitTimerHBRL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=0
            Runners=0
            HomeRun=False
            Exit Sub
        End If
      If RunnerOnFirst=1 AND RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            HBTimer1.Enabled=1
            WaitTimerTB.Enabled=1
            WaitTimerHBRM.Enabled=1
            WaitTimerTBRL.Enabled=1
            WaitTimerHBRL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=0
            Runners=0
            HomeRun=False
            Exit Sub
        End If
      If RunnerOnSecond=1 and RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            TBTimer1.Enabled=1
            HBTimer1.Enabled=1
            WaitTimerTBRM.Enabled=1
            WaitTimerHBRM.Enabled=1
            WaitTimerTBRL.Enabled=1
            WaitTimerHBRL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=0
            Runners=0
            HomeRun=False
            Exit Sub
        End If
        If RunnerOnThird=1 Then
            FBTimer1.Enabled=1
            HBTimer1.Enabled=1
            WaitTimerTBRL.Enabled=1
            WaitTimerHBRL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=0
            Runners=0
            Exit Sub
        End If
        If RunnerOnSecond=1 Then
            FBTimer1.Enabled=1
            TBTimer1.Enabled=1
            WaitTimerSB.Enabled=1
            WaitTimerHBL.Enabled=1
            WaitTimerTBRL.Enabled=1
            WaitTimerHBRL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=0
            Runners=0
            HomeRun=False
            Exit Sub
        End If
        If RunnerOnFirst=1 Then
            FBTimer1.Enabled=1
            SBTimer1.Enabled=1
            WaitTimerTB.Enabled=1
            WaitTimerHBL.Enabled=1
            WaitTimerTBRL.Enabled=1
            WaitTimerHBRL.Enabled=1
            RunnerOnFirst=0
            RunnerOnSecond=0
            RunnerOnThird=0
            Runners=0
            HomeRun=False
            Exit Sub
        End If
        If Runners=0 Then
            FBTimer1.Enabled=1
            If GameInit=True Then
                HomeRun=False
            End If
        End If
    End Select
End Sub

'Switches

Sub Strength1_Hit
  If Ready=1 Then
    Flasher1LL.Visible=1
    Gage1.State=1
    Power=4
  End If
End Sub

Sub Strength1_UnHit
  Flasher1LL.Visible=0
  Gage1.State=0
End Sub

Sub Strength2_Hit
  If Ready=1 Then
    Flasher2LL.Visible=1
    Gage2.State=1
    Power=5
  End If
End Sub

Sub Strength2_UnHit
  Flasher2LL.Visible=0
  Gage2.State=0
End Sub

Sub Strength3_Hit
  If Ready=1 Then
    Flasher3LL.Visible=1
    Gage3.State=1
    Power=5
  End If
End Sub

Sub Strength3_UnHit
  Flasher3LL.Visible=0
  Gage3.State=0
End Sub

Sub Strength4_Hit
  If Ready=1 Then
    Flasher4LL.Visible=1
    Gage4.State=1
    Power=6
  End If
End Sub

Sub Strength4_UnHit
  Flasher4LL.Visible=0
  Gage4.State=0
End Sub

Sub Strength5_Hit
  If Ready=1 Then
    Flasher5LL.Visible=1
    Gage5.State=1
    Power=6
  End If
End Sub

Sub Strength5_UnHit
  Flasher5LL.Visible=0
  Gage5.State=0
End Sub

Sub Bullseye_Hit
  If Ready = 1 Then
    PlaySoundAtVol "shooter", ActiveBall, 1
    Plunger.Fire
  Else
    PlaySoundAtVol "launch", ActiveBall, 1
  End If
End Sub

Sub ReadyTrigger_Hit
  diverter.IsDropped = 1
End Sub

Sub ReadyTrigger_UnHit
  Ready = 1
  If GameInit=True Then
    WaitTimer.Enabled=1
    Exit Sub
  End If
End Sub

Sub NotReadyTrigger_Hit
    Ready=0
    diverter.IsDropped = 0
End Sub

Sub SingleLeft_Hit
  'Single
    BaseRunner(1)
  PlaySoundAtVol "crowdcheerweak", ActiveBall, 1
End Sub

Sub OutLeft_Hit
  'You are out!
  Outs
End Sub

Sub Triple_Hit
  'Triple
    BaseRunner(3)
    PlaySoundAtVol "crowdcheerstrong", ActiveBall, 1
End Sub

Sub HomeRunLeft_Hit
  'HomeRun
    If LtHomeRunLit=0 Then
      LeftHomeRunLit.Visible=1
      LH.State=1
      LtHomeRunLit=1
    End If
    If LtHomeRunLit + RtHomeRunLit = 2 Then
      TenRunsLit.Visible=1
      TH.State=1
    End If
    PlaySoundAtVol "crowdcheerstrong", ActiveBall, 1
    BaseRunner(4)
End Sub

Sub Center_Hit
  'Evaluate Base Runner Status
  If LtHomeRunLit + RtHomeRunLit = 2 Then
    LeftHomeRunLit.Visible=0
    LH.State=0
    RightHomeRunLit.Visible=0
    RH.State=0
    TenRunsLit.Visible=0
    TH.State=0
    LtHomeRunLit=0
    RtHomeRunLit=0
    PlaySoundAtVol "crowdcheerstrong", ActiveBall, 1
    Bonus=1
    BaseRunner(4)
   Else
    If TotalOuts > 0 Then TotalOuts=TotalOuts - 1 End If
      If TotalOuts=1 Then
        LitTwo.Visible=0
        Lit2L.State=0
        LitOne.Visible=1
        Lit1L.State=1
      Else
        LitOne.Visible=0
        Lit1L.State=0
      End If
    BaseRunner(2)
    PlaySoundAtVol "crowdcheerweak", ActiveBall, 1
   End If
End Sub

Sub HomeRunRight_Hit
  'HomeRun
    If RtHomeRunLit=0 Then
      RightHomeRunLit.Visible=1
      RH.State=1
      RtHomeRunLit=1
    End If
    If LtHomeRunLit + RtHomeRunLit = 2 Then
      TenRunsLit.Visible=1
      TH.State=1
    End If
    PlaySoundAtVol "crowdcheerstrong", ActiveBall, 1
    BaseRunner(4)
End Sub

Sub Double_Hit
  'Double
    BaseRunner(2)
    PlaySoundAtVol "crowdcheerweak", ActiveBall, 1
End Sub

Sub OutRight_Hit
  'You are out!
  Outs
End Sub

Sub SingleRight_Hit
  'Single
    BaseRunner(1)
    PlaySoundAtVol "crowdcheerweak", ActiveBall, 1
End Sub

'************ Misc Sounds ****************

Sub Post1_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
End Sub
Sub Post2_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
End Sub
Sub Post3_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
End Sub
Sub Post4_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
End Sub
Sub Post5_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
End Sub
Sub Post6_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
End Sub
Sub Post7_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
End Sub
Sub Post8_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
End Sub
Sub Post9_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
End Sub
Sub Post10_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
End Sub
Sub RightBarrier_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
End Sub
Sub LeftBarrier_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
End Sub
Sub DrainSwitch_Hit
  PlaySoundAtVol "gamestart", ActiveBall, 1
  If TotalOuts=3 Then
    BallRelease.IsDropped=0
  Else
    BallRelease.IsDropped=0
    BallReleaseTimer.Enabled=1
  End If
End Sub


'*************** Outs *************
Sub Outs
  TotalOuts=TotalOuts+1
  If TH.State=1 OR RH.State=1 OR LH.State=1 Then
    RightHomeRunLit.Visible=0
    RH.State=0
    LeftHomeRunLit.Visible=0
    LH.State=0
    TenRunsLit.Visible=0
    TH.State=0
    LtHomeRunLit=0
    RtHomeRunLit=0
  End If
  Select Case TotalOuts
    Case 1
      LitOne.Visible=1
      Lit1L.State=1
    Case 2
      LitTwo.Visible=1
      Lit2L.State=1
      LitOne.Visible=0
      Lit1L.State=0
    Case 3
      LitThree.Visible=1
      Lit3L.State=1
      LitTwo.Visible=0
      Lit2L.State=0
  End Select
  PlaySound "buzzer"
  If TotalOuts=3 Then EndGame
End Sub

'*************** New Game ******************

Sub NewGame
    PlaySound "gamestart"
    GameStarted=True
  GameInit=True
  GameOver.Visible=0
  GameOverLight.State=0
  Credit.Visible=1
'Reset Special Home Run Lamps
  LtHomeRunLit=0
  RtHomeRunLit=0
  LeftHomeRunLit.Visible=0
  LH.State=0
  RightHomeRunLit.Visible=0
  RH.State=0
  TenRunsLit.Visible=0
  TH.State=0
'Reset Base Runners
  BaseRunner(4)
'Reset Outs
  OutTimer.Enabled=1
    TotalOuts=0
'Reset Score Reel
 Rscore = 0
  TensScore = 0
  HndsScore = 0
  TheScore = 0
  AddReelSCore 0
'Drop Balls and Begin
 BallRelease.IsDropped=1
'Light Marquis
  Dim i: i=0
  For each i in Marquis:i.State = 1: Next
  If LL=0 Then
    lowercablit.Visible=1
  End If
'Light Reels
  Light022.State=1
  Light023.State=1
  Light024.State=1
 End Sub

'*************** End of Game *****************

Sub EndGame
'Note last playfield positions are left displayed
   BallRelease.IsDropped=0
     PlaySound "Ding3"
   GameStarted=False
   GameOver.Visible=1
   GameOverLight.State=1
   Credit.Visible=0
   If TheScore > hisc Then
    hisc = TheScore
    HighScoreEntryInit
   End If
'Turn Off Marquis
  Dim i: i=0
  For each i in Marquis:i.State = 0: Next
  If LL=0 Then
    lowercablit.Visible=0
  End If
    Light022.State=0
  Light023.State=0
  Light024.State=0
End Sub

'************** Advance Runner Timers **************

'Running to First Timers
Sub FBTimer1_Timer()
  FBTimer1.Enabled=False
  Batter.Visible=0
  HomePlate.State=0
  RunningtoFirstA.Visible=1
  FBTimer2.Enabled=1
End Sub
Sub FBTimer2_Timer
    FBTimer2.Enabled=False
  RunningtoFirstA.Visible=0
  FBTimer3.Enabled=1
End Sub
Sub FBTimer3_Timer
    FBTimer3.Enabled=False
  RunningtoFirstB.Visible=1
  FBTimer4.Enabled=1
End Sub
Sub FBTimer4_Timer
  FBTimer4.Enabled=False
  RunningtoFirstB.Visible=0
  FBTimer5.Enabled=1
End Sub
Sub FBTimer5_Timer
  FBTimer5.Enabled=False
  RunningtoFirstC.Visible=1
  FBTimer6.Enabled=1
End Sub
Sub FBTimer6_Timer
  FBTimer6.Enabled=False
  RunningtoFirstC.Visible=0
  If Sngle=True Then
      FirstBaseRunner.Visible=1
      Firstbase.State=1
      Sngle=False
      Exit Sub
  Else
      FirstBaseRunner.Visible=0
      Firstbase.State=0
      SBTimer1.Enabled=1
      Exit Sub
  End If
  If RunnerOnFirst=1 AND RunnerOnThird=1 Then
      FirstBaseRunner.Visible=0
      Firstbase.State=0
      Exit Sub
  End If
  If Dbl=True Then
      FirstBaseRunner.Visible=0
      Firstbase.State=0
      WaitTimerSBQ.Enabled=1
      Exit Sub
  End If
End Sub

'Running to Second Timers
Sub SBTimer1_Timer()
  SBTimer1.Enabled=False
  FirstBaseRunner.Visible=0
  Firstbase.State=0
  RunningtoSecondA.Visible=1
  SBTimer2.Enabled=1
End Sub
Sub SBTimer2_Timer
    SBTimer2.Enabled=False
  RunningtoSecondA.Visible=0
  SBTimer3.Enabled=1
End Sub
Sub SBTimer3_Timer
    SBTimer3.Enabled=False
  RunningtoSecondB.Visible=1
  SBTimer4.Enabled=1
End Sub
Sub SBTimer4_Timer
  SBTimer4.Enabled=False
  RunningtoSecondB.Visible=0
  SBTimer5.Enabled=1
End Sub
Sub SBTimer5_Timer
  SBTimer5.Enabled=False
  RunningtoSecondC.Visible=1
  SBTimer6.Enabled=1
End Sub
Sub SBTimer6_Timer
  SBTimer6.Enabled=False
  RunningtoSecondC.Visible=0
  If Trpl=True OR GameInit=True OR HomeRun=True Then
    SecondBaseRunner.Visible=0
    Secondbase.State=0
    TBTimer1.Enabled=1
    Exit Sub
  End If
  If Dbl=True Then
    SecondBaseRunner.Visible=1
    Secondbase.State=1
    Dbl=False
    Exit Sub
  End If
SecondBaseRunner.Visible=1
Secondbase.State=1
End Sub

'Running to Third Timers
Sub TBTimer1_Timer()
  TBTimer1.Enabled=False
  SecondBaseRunner.Visible=0
  Secondbase.State=0
  RunningtoThirdA.Visible=1
  TBTimer2.Enabled=1
End Sub
Sub TBTimer2_Timer
    TBTimer2.Enabled=False
  RunningtoThirdA.Visible=0
  TBTimer3.Enabled=1
End Sub
Sub TBTimer3_Timer
    TBTimer3.Enabled=False
  RunningtoThirdB.Visible=1
  TBTimer4.Enabled=1
End Sub
Sub TBTimer4_Timer
  TBTimer4.Enabled=False
  RunningtoThirdB.Visible=0
  TBTimer5.Enabled=1
End Sub
Sub TBTimer5_Timer
  TBTimer5.Enabled=False
  RunningtoThirdC.Visible=1
  TBTimer6.Enabled=1
End Sub
Sub TBTimer6_Timer
  TBTimer6.Enabled=False
  RunningtoThirdC.Visible=0
  If Sngle=True AND Runners=3 Then
    ThirdBaseRunner.Visible=1
    Thirdbase.State=1
    Exit Sub
  End If
  If Dbl=True Then
    If RunnerOnFirst=1 Then
      ThirdBaseRunner.Visible=1
      Thirdbase.State=1
      SecondBaseRunner.Visible=1
      Secondbase.State=1
      Dbl=False
      Exit Sub
    Else
      ThirdBaseRunner.Visible=0
      ThirdBase.State=0
      HBTimer1.Enabled=1
      Dbl=False
      Exit Sub
    End If
  End If
  If Trpl=True Then
    ThirdBaseRunner.Visible=1
    Thirdbase.State=1
    Trpl=False
    Exit Sub
  End If
  If GameInit=True OR Sngle=True OR HomeRun=True Then
    ThirdBaseRunner.Visible=0
    Thirdbase.State=0
    HBTimer1.Enabled=1
    Sngle=False
    Exit Sub
  End If
  ThirdBaseRunner.Visible=1
  Thirdbase.State=1
End Sub

'Running to Home Timers
Sub HBTimer1_Timer()
  HBTimer1.Enabled=False
  ThirdBaseRunner.Visible=0
  Thirdbase.State=0
  RunningtoHomeA.Visible=1
  HBTimer2.Enabled=1
End Sub
Sub HBTimer2_Timer
    HBTimer2.Enabled=False
  RunningtoHomeA.Visible=0
  HBTimer3.Enabled=1
End Sub
Sub HBTimer3_Timer
    HBTimer3.Enabled=False
  RunningtoHomeB.Visible=1
  HBTimer4.Enabled=1
End Sub
Sub HBTimer4_Timer
  HBTimer4.Enabled=False
  RunningtoHomeB.Visible=0
  HBTimer5.Enabled=1
End Sub
Sub HBTimer5_Timer
  HBTimer5.Enabled=False
  RunningtoHomeC.Visible=1
  HBTimer6.Enabled=1
End Sub
Sub HBTimer6_Timer
  HBTimer6.Enabled=False
  RunningtoHomeC.Visible=0
    If GameInit = False Then
      PlaySound "Ding3"
      AddReelScore 1
      If Bonus = 1 Then
        AddReelScore 9
        Bonus = 0
      End If
    End If
  GameInit=False
  HomeRun=False
End Sub

'************Out Lamp Timers*************
'Out Lamp Reset Timers
Sub OutTimer_Timer
  OutTimer.Enabled=False
  LitThree.Visible=1
  Lit3L.State=1
  OutTimer0.Enabled=True
End Sub
Sub OutTimer0_Timer
  OutTimer0.Enabled=False
  LitThree.Visible=0
  Lit3L.State=0
  OutTimer1.Enabled=True
End Sub
Sub OutTimer1_Timer
  OutTimer1.Enabled=False
  LitTwo.VIsible=1
  Lit2L.State=1
  OutTimer2.Enabled=True
End Sub
Sub OutTimer2_Timer
  OutTimer2.Enabled=False
  LitTwo.VIsible=0
  Lit2L.State=0
  OutTimer3.Enabled=True
End Sub
Sub OutTimer3_Timer
  OutTimer3.Enabled=False
  LitOne.Visible=1
  Lit1L.State=1
  OutTimer4.Enabled=True
End Sub
Sub OutTimer4_Timer
  OutTimer4.Enabled=False
  LitOne.Visible=0
  Lit1L.State=0
End Sub

'****************Runner Wait Timers****************

'Second Base
Sub WaitTimerSB_Timer
  WaitTimerSB.Enabled=False
  SBTimer1.Enabled=1
End Sub
Sub WaitTimerSBQ_Timer
  WaitTimerSBQ.Enabled=False
  SBTimer1.Enabled=1
End Sub
Sub WaitTimerSBL_Timer
  WaitTimerSBL.Enabled=False
  SecondBaseRunner.Visible=0
  SecondBase.State=0
  SBTimer1.Enabled=1
End Sub
Sub WaitTimerSBRL_Timer
  WaitTimerSBRL.Enabled=False
  SecondBaseRunner.Visible=0
  SecondBase.State=0
  SBTimer1.Enabled=1
End Sub

'Third Base
Sub WaitTimerTB_Timer
  WaitTimerTB.Enabled=False
  TBTimer1.Enabled=1
End Sub
Sub WaitTimerTBQ_Timer
  WaitTimerTBQ.Enabled=False
  TBTimer1.Enabled=1
End Sub
Sub WaitTimerTBL_Timer
  WaitTimerTBL.Enabled=False
  TBTimer1.Enabled=1
End Sub
Sub WaitTimerTBRM_Timer
  WaitTimerTBRM.Enabled=False
  ThirdBaseRunner.Visible=0
  ThirdBase.State=0
  TBTimer1.Enabled=1
End Sub
Sub WaitTimerTBRL_Timer
  WaitTimerTBRL.Enabled=False
  ThirdBaseRunner.Visible=0
  ThirdBase.State=0
  TBTimer1.Enabled=1
End Sub

'Home Base
Sub WaitTimerHB_Timer
  WaitTimerHB.Enabled=False
  HBTimer1.Enabled=1
End Sub

Sub WaitTimerHBQ_Timer
  WaitTimerHBQ.Enabled=False
  HBTimer1.Enabled=1
End Sub

Sub WaitTimerHBL_Timer
  WaitTimerHBL.Enabled=False
  ThirdBaseRunner.Visible=0
  ThirdBase.State=0
  HBTimer1.Enabled=1
End Sub

Sub WaitTimerHBRM_Timer
  WaitTimerHBRM.Enabled=False
  ThirdBaseRunner.Visible=0
  ThirdBase.State=0
  HBTimer1.Enabled=1
End Sub

Sub WaitTimerHBRL_Timer
  WaitTimerHBRL.Enabled=False
  ThirdBaseRunner.Visible=0
  ThirdBase.State=0
  HBTimer1.Enabled=1
End Sub

'Home Plate Batter
Sub WaitTimer_Timer
  WaitTimer.Enabled=False
      Batter.Visible=1
    HomePlate.State=1
End Sub

Sub WaitTimer2_Timer
  WaitTimer2.Enabled=False
      Batter.Visible=1
    HomePlate.State=1
End Sub

Sub BallReleaseTimer_Timer
  BallReleaseTimer.Enabled=False
    BallRelease.IsDropped=1
  WaitTimer2.Enabled=1
End Sub

'************* Set LUT ***************
Sub SetLUT
  table1.ColorGradeImage = luts(lutpos)
End Sub

'darquayle
Sub loadhs
    dim temp
  temp = LoadValue("MiniB", "hiscore")
    If (temp <> "") then hisc = CDbl(temp) else hisc = hisc_default
    temp = LoadValue("MiniB", "score1")
    If (temp <> "") then DTcount = CDbl(temp)
    temp = LoadValue("MiniB", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp) else HSA1 = HSA1_default
    temp = LoadValue("MiniB", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp) else HSA2 = HSA2_default
    temp = LoadValue("MiniB", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp) else HSA3 = HSA3_default
End Sub

'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD: And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
'TO SHOW THE SCORE ON POST-IT ADD LINE AT RELEVENT LOCATION THAT HAS:  UpdatePostIt
'TO INITIATE ADDING INITIALS ADD LINE AT RELEVENT LOCATION THAT HAS:  HighScoreEntryInit()
'ADD THE FOLLOWING LINES TO TABLE_INIT TO SETUP POSTIT
' if HSA1="" then HSA1=25
' if HSA2="" then HSA2=25
' if HSA3="" then HSA3=25
' UpdatePostIt
'ADD HSA1, HSA2 AND HSA3 TO SAVE AND LOAD VALUES FOR TABLE
'ADD A TIMER NAMED HighScoreFlashTimer WITH INTERVAL 100 TO TABLE
'SET HSSSCOREX BELOW TO WHATEVER VARIABLE YOU USE FOR HIGH SCORE.
'ADD OBJECTS TO PLAYFIELD (EASIEST TO JUST COPY FROM THIS TABLE)
'IMPORT POST-IT IMAGES


'Dim HSA1, HSA2, HSA3


' ***********************************************************
'  HiScore DISPLAY
' ***********************************************************

Sub UpdatePostIt
  dim tempscore
  HSScorex = hisc
  TempScore = HSScorex
  HSScore1 = 0
  HSScore10 = 0
  HSScore100 = 0
  If len(TempScore) > 0 Then
    HSScore1 = cint(right(Tempscore,1))
  End If
  If len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore10 = cint(right(Tempscore,1))
  End If
  If len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore100 = cint(right(Tempscore,1))
  End If

  PScore2.image = HSArray(HSScore100):If HSScorex<100 Then PScore2.image = HSArray(10)
  PScore2LL.image = HSArray(HSScore100):If HSScorex<100 Then PScore2LL.image = HSArray(10)
  PScore1.image = HSArray(HSScore10):If HSScorex<10 Then PScore1.image = HSArray(10)
  PScore1LL.image = HSArray(HSScore10):If HSScorex<10 Then PScore1LL.image = HSArray(10)
  PScore0.image = HSArray(HSScore1):If HSScorex<1 Then PScore0.image = HSArray(10)
  PScore0LL.image = HSArray(HSScore1):If HSScorex<1 Then PScore0LL.image = HSArray(10)
  HSName1.image = ImgFromCode(HSA1, 1)
  HSName1LL.image = ImgFromCode(HSA1, 1)
  HSName2.image = ImgFromCode(HSA2, 2)
  HSName2LL.image = ImgFromCode(HSA2, 2)
  HSName3.image = ImgFromCode(HSA3, 3)
  HSName3LL.image = ImgFromCode(HSA3, 3)
End Sub

Function ImgFromCode(code, digit)
  Dim Image
  If (HighScoreFlashTimer.Enabled = True and hsLetterFlash = 1 and digit = hsCurrentLetter) Then
    Image = "postitBL"
  ElseIf (code + ASC("A") - 1) >= ASC("A") and (code + ASC("A") - 1) <= ASC("Z") Then
    Image = "postit" & chr(code + ASC("A") - 1)
  ElseIf code = 27 Then
    Image = "PostitLT"
    ElseIf code = 0 Then
    image = "PostitSP"
    Else
      msgbox("Unknown display code: " & code)
  End If
  ImgFromCode = Image
End Function

Sub HighScoreEntryInit()
  HSA1=0:HSA2=0:HSA3=0
  HSEnterMode = True
  hsCurrentDigit = 0
  hsCurrentLetter = 1:HSA1=1
  HighScoreFlashTimer.Interval = 250
  HighScoreFlashTimer.Enabled = True
  hsLetterFlash = hsFlashDelay
End Sub

Sub HighScoreFlashTimer_Timer()
  hsLetterFlash = hsLetterFlash-1
  UpdatePostIt
  If hsLetterFlash=0 Then 'switch back
    hsLetterFlash = hsFlashDelay
  End If
End Sub

' *************************************************************
'  HiScore ENTER INITIALS
' *************************************************************

Sub HighScoreProcessKey(keycode)
    If keycode = LeftFlipperKey Then
    hsLetterFlash = hsFlashDelay
    Select Case hsCurrentLetter
      Case 1:
        HSA1=HSA1-1:If HSA1=-1 Then HSA1=26 'no backspace on 1st digit
        UpdatePostIt
      Case 2:
        HSA2=HSA2-1:If HSA2=-1 Then HSA2=27
        UpdatePostIt
      Case 3:
        HSA3=HSA3-1:If HSA3=-1 Then HSA3=27
        UpdatePostIt
     End Select
    End If

  If keycode = RightFlipperKey Then
    hsLetterFlash = hsFlashDelay
    Select Case hsCurrentLetter
      Case 1:
        HSA1=HSA1+1:If HSA1>26 Then HSA1=0
        UpdatePostIt
      Case 2:
        HSA2=HSA2+1:If HSA2>27 Then HSA2=0
        UpdatePostIt
      Case 3:
        HSA3=HSA3+1:If HSA3>27 Then HSA3=0
        UpdatePostIt
     End Select
  End If

    If keycode = StartGameKey Then
    Select Case hsCurrentLetter
      Case 1:
        hsCurrentLetter=2 'ok to advance
        HSA2=HSA1 'start at same alphabet spot
      Case 2:
        If HSA2=27 Then 'bksp
          HSA2=0
          hsCurrentLetter=1
        Else
          hsCurrentLetter=3 'enter it
          HSA3=HSA2 'start at same alphabet spot
        End If
      Case 3:
        If HSA3=27 Then 'bksp
          HSA3=0
          hsCurrentLetter=2
        Else
          savehs 'enter it
          HighScoreFlashTimer.Enabled = False
          HSEnterMode = False
        End If
    End Select
    UpdatePostIt
    End If
End Sub

Sub savehs
    savevalue "MiniB", "hiscore", hisc
    savevalue "MiniB", "score1", TheScore
  savevalue "MiniB", "hsa1", HSA1
  savevalue "MiniB", "hsa2", HSA2
  savevalue "MiniB", "hsa3", HSA3
End Sub

Sub HStimer_timer
  playsoundat SoundFXDOF("Ding3",140,DOFPulse,DOFKnocker), Plunger
  HStimer.uservalue=HStimer.uservalue+1
  If HStimer.uservalue=3 Then me.enabled=0
End Sub
'/darquayle

'*********** Reel-O-Matic **********
Sub AddReelScore(Points)
    Dim i : i=0
    Dim j : j=0
    Dim l : l=0
    TheScore=TheScore+Points
    'Playsound"cycle"
    If TheScore=100 and GameType=1 Then
      EndGame
    Else
      If TheScore=999 Then
        EndGame
      End If
    End If
    Rscore=RScore + Points
    If Rscore < 10 Then
      For each i in Ones:i.visible = False: Next
      Ones(Rscore).Visible = True
    End If
    If Rscore > 10 Then
      Rscore=Rscore-10
      For each i in Ones:i.visible = False: Next
      Ones(Rscore).Visible = True
      TensScore=TensScore + 1
    End If
    If Rscore=10 Then
      For each i in Ones:i.visible = False: Next
      Ones(Rscore-10).Visible = True
      For each i in Tens:i.visible = False: Next
      Tens(TensScore).Visible = True
      RScore=0
      TensScore=TensScore +1
    End If
    If TensScore < 10 Then
      For each j in Tens:j.visible = False: Next
      Tens(TensScore).Visible = True
    End If
    If TensScore > 10 Then
      TensScore=TensScore-10
      For each i in Tens:i.visible = False: Next
      Tens(TensScore).Visible = True
      HndsScore=HndsScore + 1
    End If
    If TensScore = 10 Then
      For each j in Tens:j.visible = False: Next
      For each l in Hnds:l.Visible = False: Next
      Hnds(HndsScore).Visible = True
      TensScore=0
      Tens(TensScore).Visible = True
      HndsScore=HndsScore +1
    End If
    If HndsScore < 10 Then
      For each l in Hnds:l.visible = False: Next
      Hnds(HndsScore).Visible = True
    End If
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


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
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
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

