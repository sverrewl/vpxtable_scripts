
' __      __.___.____    .____    .___   _____      _____    _________  ____ ________  ______   ________
'/  \    /  \   |    |   |    |   |   | /  _  \    /     \  /   _____/ /_   /   __   \/  __  \ /  _____/
'\   \/\/   /   |    |   |    |   |   |/  /_\  \  /  \ /  \ \_____  \   |   \____    />      </   __  \
' \        /|   |    |___|    |___|   /    |    \/    Y    \/        \  |   |  /    //   --   \  |__\  \
'  \__/\  / |___|_______ \_______ \___\____|__  /\____|__  /_______  /  |___| /____/ \______  /\_____  /
'       \/              \/       \/           \/         \/        \/                       \/       \/
'  __________________    _____    _______  ________    .____    ._____________  _____ __________________
' /  _____/\______   \  /  _  \   \      \ \______ \   |    |   |   \____    / /  _  \\______   \______ \
'/   \  ___ |       _/ /  /_\  \  /   |   \ |    |  \  |    |   |   | /     / /  /_\  \|       _/|    |  \
'\    \_\  \|    |   \/    |    \/    |    \|    `   \ |    |___|   |/     /_/    |    \    |   \|    `   \
' \______  /|____|_  /\____|__  /\____|__  /_______  / |_______ \___/_______ \____|__  /____|_  /_______  /
'        \/        \/         \/         \/        \/          \/           \/       \/       \/        \/
'

'                                                   Grand Lizard V1.8.5
'
'                                            A 3rdaxis & Slydog43 Collaboration

'
'-Special Thanks
'-Thanks to JP Salas for his VP9 version
'-Thanks to Rosve for his VP9 version
'-Thanks to G5K for the flippers
'-Thanks to VP Dev Team for making VPX so awesome!!!

Option Explicit
Randomize

On Error Resume Next
  ExecuteGlobal GetTextFile("controller.vbs")
  If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "S11.VBS", 3.26

Dim PRpull,PRdir,InPlunger,LizardFlash
 Const cGameName="grand_l4"
 'Const UseSolenoids=1
 Const UseSolenoids=2 'FastFlips
 Const UseLamps=1
 ' Standard sounds
 Const SCoin="CoinIn3"

Dim bsTrough, dt3B, dt4L, dt4R, mLeftMagnet, mRightMagnet, i
Dim myImage, myObject, myMaterial, myLightColor, myLightColorFull, myLightIntensity, myPrefs_DynamicBackground
Dim myFlipperLeftMatterial, myFlipperLeftRubberMaterial, myFlipperLeftUpperMatterial, myFlipperLeftUpperRubberMaterial
Dim myFlipperRightMatterial, myFlipperRightRubberMaterial, myFlipperRightUpperMatterial, myFlipperRghtUpperRubberMaterial
Dim myPrefs_BallRollingSound, myPrefs_PlayMusic, myPrefs_ShowJungleBlades, myPrefs_HideCenterPost, myPrefs_GIColor, myPrefs_RubberColor
Dim myPrefs_RubberSleevesColor, myPrefs_OutlaneDifficulty, myPrefs_PlayfieldGlass, myPrefs_Bounce,  myPlayingEndGameSnd
Dim myPrefs_Drain, myPrefs_Soundonoff, myPrefs_FlipperBatType, myPrefs_InstructionCardType, myPrefs_ScratchGlass
myPlayingEndGameSnd = False

'***********************************************************************************************************************************
' _________   ___  __   ____  ____  ___  ______________  _  ______                                                                 |
'/_  __/ _ | / _ )/ /  / __/ / __ \/ _ \/_  __/  _/ __ \/ |/ / __/                                                                 |
' / / / __ |/ _  / /__/ _/  / /_/ / ___/ / / _/ // /_/ /    /\ \                                                                   |
'/_/ /_/ |_/____/____/___/  \____/_/    /_/ /___/\____/_/|_/___/                                                                   |
'***********************************************************************************************************************************

'********* DIFFICULTY ********************************************
myPrefs_OutlaneDifficulty = 2         '0=Medium(Factory) 1=Easy 2=Hard
myPrefs_HideCenterPost = 0               '0=Factory 1=Hide Post
myPrefs_Bounce = 1                               '0=Off 1=On Allows ball to bounce into the air when it strikes a rubber post. (Experimental)
'********* TABLE MODS ********************************************
myPrefs_RubberColor = 0                     '0=White(Factory) 1=Random 2=Black 3=Green 4=Red 5=Blue
myPrefs_RubberSleevesColor = 0       '0=Yellow(Factory 1=Random 2=Black 3=Green 4=Red 5=Blue
myPrefs_GIColor = 0                             '0=Warm White(Factory) 1=Random 2=Cool White 3=Green 4=Blue 5=Purple 6=Red
myPrefs_ShowJungleBlades = 1           '0=Off(Factory) 1=On
myPrefs_InstructionCardType = 0     '0=Factory(Factory) 1=Modded Cards 2=THX Card  (See THX Setup for directions)
myPrefs_DynamicBackground = 1         '0=Off 1=On
myPrefs_PlayMusic = 1                         '0=Off(Default) 1=On Extra Sounds and music.
myPrefs_ScratchGlass = 1
myPrefs_PlayfieldGlass = 1 '(New) '0=No glass (flippers and other sounds will be much brither) 1=Glass (Sounds will be muffled)
                                                                 'Use the "G" key (34) on keyboard to toggle On/Off during game.
myPrefs_FlipperBatType = 3               'Flipper Bat Type 0=Factory
'*****************************************************************
'0= White Bat Red Rubber Upper & Lower(Factory)
'1= White Bat Black Rubber Upper & Lower
'2= White Bat Green Rubber Upper & Lower
'3= White Bat Red Rubber Upper & Blue Rubber Lower

'4= Yellow Bat Red Rubber Upper & Lower
'5= Yellow Bat Black Rubber Upper & Lower
'6= Yellow Bat Green Rubber Upper & Lower
'7= Yellow Bat Green Rubber Upper & Red Rubber Lower
'*******************************
'THX Setup. (Select Instruction Card Type 2)
'(Basic adjustment)
'Adjusting your tables "Evironment Emmision" under the Camera/Light Edit Mode (F6)
'is the easiest way to get you calibrated to the correct brightness for this table only.
'During adjustment, the THX Logo's shadow should just barely be visible on the right player card.
'(Advanced adjustment)
'The left card is for adjusting your TV's "Contrast"
'This provides a good contrast test pattern that you should
'be able to clearly see all eight rectangles but keep in mind,
'this will take some back and forth between brightness,
'contrast and backlight to get right and also
'maintain the right card being correctly adjusted as well. Enjoy;)
'**********************************************************************************************************************************


'        ______   _______  ___  _______  ____
'*****  / __/ /  /  _/ _ \/ _ \/ __/ _ \/ __/*****
'***** / _// /___/ // ___/ ___/ _// , _/\ \  ***** THANK YOU G5K
'*****/_/ /____/___/_/  /_/  /___/_/|_/___/  *****

  Sub Frametimer_Timer()
      BallShadowUpdate
      'LightCopy
      'RollingTimer
  End Sub

  Select Case myPrefs_FlipperBatType  'Upper Flipper Primitive Material
    Case 0
    batleft1.image = "_flipper_white_red"
    batright1.image = "_flipper_white_red"
    Case 1
          batleft1.image = "_flipper_white_black"
    batright1.image = "_flipper_white_black"
    Case 2
          batleft1.image = "_flipper_white_green"
    batright1.image = "_flipper_white_green"
    Case 3
    batleft1.image = "_flipper_white_red"
    batright1.image = "_flipper_white_red"
    Case 4
    batleft1.image = "_flipper_yellow_red"
    batright1.image = "_flipper_yellow_red"
    Case 5
                batleft1.image = "_flipper_yellow_black"
    batright1.image = "_flipper_yellow_black"
    Case 6
    batleft1.image = "_flipper_yellow_green"
    batright1.image = "_flipper_yellow_green"
     Case 7
    batleft1.image = "_flipper_yellow_green"
    batright1.image = "_flipper_yellow_green"
  End Select

  Select Case myPrefs_FlipperBatType 'Lower Flipper Primitive Material
    Case 0
    batleft.image = "_flipper_white_red"
    batright.image = "_flipper_white_red"
    Case 1
          batleft.image = "_flipper_white_black"
    batright.image = "_flipper_white_black"
    Case 2
          batleft.image = "_flipper_white_green"
    batright.image = "_flipper_white_green"
    Case 3
    batleft.image = "_flipper_white_blue"
    batright.image = "_flipper_white_blue"
    Case 4
    batleft.image = "_flipper_yellow_red"
    batright.image = "_flipper_yellow_red"
    Case 5
                batleft.image = "_flipper_yellow_black"
    batright.image = "_flipper_yellow_black"
    Case 6
    batleft.image = "_flipper_yellow_green"
    batright.image = "_flipper_yellow_green"
    Case 7
    batleft.image = "_flipper_yellow_red"
    batright.image = "_flipper_yellow_red"
  End Select


'********************************************************
  Sub myAdjustTableToPrefs

  If myPrefs_PlayMusic then
    PlaySound "GrandLizardIntro"
    Playsound "Lion-Roar"
   End If

  If myPrefs_ShowJungleBlades Then
    BladeR.Visible = True
    BladeL.Visible = True
  End If

  If myPrefs_HideCenterPost Then
    CenterPost.Visible = False
    CPostRubber.Visible = False
    CPostRubber.Collidable = False
  End If

  If ShowDT then
  SideRails.Visible = True
  LockdownBar. Visible = True
  LizardHeadFS.rotx = -93
  LizardHeadFS.transy = -10
  Else
  LizardHeadFS.rotx = -90
  LizardHeadFS.transy = 7.1
  End If

  Call mySetInstructionCards(myPrefs_InstructionCardType)

  Select Case myPrefs_OutlaneDifficulty
    Case 0 'Default
       PostOutlaneLeft2.Visible = True
       RubberOutlaneLeft2.Visible = True
       RubberOutlaneLeft2.Collidable = True
       PostOutlaneRight2.Visible = True
       RubberOutlaneRight2.Visible = True
       RubberOutlaneRight2.Collidable = True
       PostScrewLeft1.Visible = True
       PostScrewRight1.Visible = True
    Case 1 'Easy
       PostOutlaneLeft1.Visible = True
       RubberOutlaneLeft1.Visible = True
       RubberOutlaneLeft1.Collidable = True
       PostOutlaneRight1.Visible = True
       RubberOutlaneRight1.Visible = True
       RubberOutlaneRight1.Collidable = True
       PostScrewLeft.Visible = True
       PostScrewRight.Visible = True
    Case 2 'Hard
       PostOutlaneLeft3.Visible = True
       RubberOutlaneLeft3.Visible = True
       RubberOutlaneLeft3.Collidable = True
       PostOutlaneRight3.Visible = True
       RubberOutlaneRight3.Visible = True
       RubberOutlaneRight3.Collidable = True
  End Select

call myAdjustGIColor()

' ****** End GI Color
Call mySetRubberColor(True)
Call mySetRubberSleeveColor(True)
Call mySetPostColor(True)
End Sub

Sub mySetPostColor(myGIOn)
      myMaterial = "PostRed"
If NOT myGIOn then
  myMaterial = myMaterial & "D" 'for dark
End if

  For each myObject in posts
    myObject.Material = myMaterial
  Next
End Sub


Sub mySetRubberColor(myGIOn)

' ****** Set Rubber Color
  Select Case myPrefs_RubberColor
    Case 0 'White
      myMaterial = "Rubber White"
    Case 1 'Random
    Select Case Int(Rnd()*4) '(Rnd()*3)AXS
          Case 0 'White
            myMaterial = "Rubber White"
          Case 1 'Black
            myMaterial = "Rubber Black"
          Case 2 'Green
            myMaterial = "Rubber Green"
          Case 3 'Red
            myMaterial = "Rubber Red"
          Case 4 'Blue
            myMaterial = "Rubber Blue"
    End Select
    Case 2 'Black
      myMaterial = "Rubber Black"
    Case 3 'Green
      myMaterial = "Rubber Green"
    Case 4 'Red
      myMaterial = "Rubber Red"
    Case 5 'Blue
      myMaterial = "Rubber Blue"
  End Select

If NOT myGIOn then
  myMaterial = myMaterial & "D" 'for dark
End if

  For each myObject in Rubbers
    myObject.Material = myMaterial
  Next
  CPostRubber.Material = myMaterial
' ****** End Set Rubber Color
End Sub

  Sub mySetInstructionCards(myCardType)
    Select Case myCardType
      Case 0
        LCardFactory.visible = 1
        RCardFactory.visible = 1
      Case 1
        LCardMod.visible = 1
        RCardMod.visible = 1
      Case 2
        LCardTHX.visible = 1
        RCardTHX.visible = 1
      Case -1
        LCardFactory.visible = 0
        RCardFactory.visible = 0
        LCardMod.visible = 0
        RCardMod.visible = 0
        LCardTHX.visible = 0
        RCardTHX.visible = 0
      End Select
  End Sub

  Sub mySetRubberSleeveColor(myGIOn)
  ' ****** Set Rubber Sleeves Color ******

    Select Case myPrefs_RubberSleevesColor
    Case 0 'Yellow
      myMaterial = "Sleeve Yellow"
      myImage = "RubberSleeve Yellow"
    Case 1 'Random
       Select Case Int(Rnd()*4)
       Case 0 'Yellow
         myMaterial = "Sleeve Yellow"
         myImage = "RubberSleeve Yellow"
       Case 2 'Black
         myMaterial = "Sleeve Black"
         myImage = "RubberSleeve Black"
       Case 3 'Green
         myMaterial = "Sleeve Green"
         myImage = "RubberSleeve Green"
       Case 4 'Red
         myMaterial = "Sleeve Red"
         MyImage = "RubberSleeve Red"
       Case 5 'Blue
         myMaterial = "Sleeve Blue"
         MyImage = "RubberSleeve Blue"
       End Select
    Case 2 'Black
      myMaterial = "Sleeve Black"
      myImage = "RubberSleeve Black"
    Case 3 'Green
      myMaterial = "Sleeve Green"
      myImage = "RubberSleeve Green"
    Case 4 'Red
     myMaterial = "Sleeve Red"
      MyImage = "RubberSleeve Red"
    Case 5 'Blue
     myMaterial = "Sleeve Blue"
      MyImage = "RubberSleeve Blue"
    End Select

'******************************************

  If NOT myGIOn then
    myMaterial = myMaterial & "D" 'for dark
  End if

    For each myObject in RubberSleeves
      myObject.Material = myMaterial
      myObject.Image = myImage
   Next
  End Sub

  Sub myAdjustGIColor()
    Select Case myPrefs_GIColor
    Case 0 'Warm White
      myLightColor = RGB(225,190,160)
      myLightColorFull = RGB(225,190,160)
      myLightIntensity = 250
    Case 1 'Random
      Select Case Int(Rnd()*6)
        Case 0 'Warm White
          myLightColor = RGB(225,190,160)
          myLightColorFull = RGB(225,190,160)
          myLightIntensity = 250
        Case 1 'Cool White
          myLightColor = RGB(230,230,255)
          myLightColorFull = RGB(230,230,255)
          myLightIntensity = 250
        Case 2 'Green
          myLightColor = RGB(100,255,0)
          myLightColorFull = RGB(100,255,0)
          myLightIntensity = 150
        Case 3 'Blue
          myLightColor = RGB(25,25,255)
          myLightColorFull = RGB(25,25,255)
          myLightIntensity = 650
        Case 4 'Purple
          myLightColor = RGB(128,0,255)
          myLightColorFull = RGB(128,0,255)
          myLightIntensity = 250
        Case 5 'Red
          myLightColor = RGB(255,5,5)
          myLightColorFull = RGB(255,5,5)
          myLightIntensity = 100
      End Select
    Case 2 'Cool White
      myLightColor = RGB(230,230,255)
      myLightColorFull = RGB(230,230,255)
      myLightIntensity = 250
    Case 3 'Green
      myLightColor = RGB(100,255,0)
      myLightColorFull = RGB(100,255,0)
      myLightIntensity = 150
    Case 4 'Blue
      myLightColor = RGB(25,25,255)
      myLightColorFull = RGB(25,25,255)
      myLightIntensity = 650
    Case 5 'Purple
      myLightColor = RGB(128,0,255)
      myLightColorFull = RGB(128,0,255)
      myLightIntensity = 250
    Case 6 'Red
      myLightColor = RGB(255,5,5)
      myLightColorFull = RGB(255,5,5)
      myLightIntensity = 100
    End Select
    For each myObject in GI
    myObject.Color = myLightColor
    myObject.ColorFull = myLightColorFull
    myObject.Intensity = myLightIntensity
    Next
  End Sub

 '***********
 ' Table Init
 '***********

   Sub Table1_Init
    On Error Resume Next
    vpminit Me
    With Controller
       .GameName=cGameName
       If Err Then MsgBox "Can't start Game: " & cGameName & vbNewLine & Err.Description:Exit Sub
       .SplashInfoLine = "Grand Lizard" & vbNewLine & " BY 3rdAxis, Slydog43"
       .HandleKeyboard=0
       .ShowTitle=0
       .ShowDMDOnly=1
       .ShowFrame=0
       .HandleMechanics=0
      if ShowDT Then
        .Hidden=0
      Else
        .Hidden=1
      End If

       '.SetDisplayPosition 0, 0, GetPlayerHWnd
       On Error Resume Next
       .Run GetPlayerHWnd
       If Err Then MsgBox Err.Description
    End With
    On Error Goto 0

    Call myAdjustTableToPrefs()

    ' Trough handler
    Set bsTrough=New cvpmBallStack ' Trough handler
    bsTrough.InitSw 10, 12, 13, 14, 0, 0, 0, 0
    bsTrough.InitKick BallRelease, 80, 7
    if myPrefs_PlayfieldGlass = 1 Then
    bsTrough.InitExitSnd "Shooter_Lane", "Solenoid"
    end If
    if myPrefs_PlayfieldGlass = 0 Then
    bsTrough.InitExitSnd "Shooter_LaneNG", "SolenoidNG"
    end If
    bsTrough.Balls=3

    ' 3 Bank
    Set dt3B=New cvpmDropTarget
      dt3B.InitDrop Array(sw25, sw26, sw27), Array(25, 26, 27):DT3bGlass = 1

    ' 4 Bank Left
    Set dt4L=New cvpmDropTarget
      dt4L.InitDrop Array(sw20, sw21), Array(20, 21):DT4LGlass = 1

    ' 4 Bank Right
    Set dt4R=New cvpmDropTarget
    dt4R.InitDrop Array(sw22, sw23), Array(22, 23):DT4RGlass = 1

    ' Magnets
    Set mLeftMagnet=New cvpmMagnet
    With mLeftMagnet
       .InitMagnet LeftMagnet, 12
       .Solenoid=9
       .GrabCenter=0
       .CreateEvents "mLeftMagnet"
    End With

    Set mRightMagnet=New cvpmMagnet
    With mRightMagnet
       .InitMagnet RightMagnet, 12
       .Solenoid=10
       .GrabCenter=0
       .CreateEvents "mRightMagnet"
    End With

    ' Nudge
    vpmNudge.TiltSwitch=swTilt
    vpmNudge.Sensitivity=1
    'vpmNudge.TiltObj=Array(sw43, sw44)

  if myPrefs_ScratchGlass = 1 Then
    GlassImpurities.ImageA = "ScratchTest5"
  Else
    GlassImpurities.ImageA = "NewGlass"
  end if

If myPrefs_Bounce = 0 Then
     Wall6.Collidable = False
     Wall7.Collidable = False
     Wall8.Collidable = False
     Wall9.Collidable = False
Else
     Wall6.Collidable = True
     Wall7.Collidable = True
     Wall8.Collidable = True
     Wall9.Collidable = True
End If

  End Sub 'Table Init

   Sub Table1_Paused:Controller.Pause=1:End Sub
   Sub Table1_unPaused:Controller.Pause=0:End Sub
   Sub Table1_Exit():Controller.Stop:End Sub


'*************************************************************************************
'*****************               KEY EVENTS                ***************************
'*************************************************************************************
   Sub Table1_KeyDown(ByVal keycode)
'Msgbox keycode
    If Keycode = 5 then'AddCreditKey or keycode = 4 or keycode = 5 or keycode = 6 Then
        playsound"AXSCoinL"
    End if

    If Keycode = 6 then'AddCreditKey or keycode = 4 or keycode = 5 or keycode = 6 Then
        playsound"AXSCoinR"
    End if

    If keycode=RightFlipperKey Then Controller.Switch(48)=1
    If keycode=LeftFlipperKey Then Controller.Switch(47)=1
    If keycode= RightMagnaSave Then Controller.Switch(46)=1
    If keycode= LeftMagnaSave Then Controller.Switch(45)=1
    If keycode = LeftTiltKey Then LeftNudge 80, .8, 20:PlaySound "nudge_left"
    If keycode = RightTiltKey Then RightNudge 280, .8, 20:PlaySound "nudge_right"
    If keycode = CenterTiltKey Then CenterNudge 0, 1.2, 25:PlaySound "nudge_forward"

    If keycode = PlungerKey And myPrefs_PlayfieldGlass = 0 Then
      Plunger1.PullBack
      PlaySoundAtVol "Plunger_Pull_1", Plunger1, 1
    End If
    If keycode = PlungerKey And myPrefs_PlayfieldGlass = 1 Then
      Plunger1.PullBack
      PlaySoundAtVol "Plunger_Pull_1G", Plunger1, 1
    End If
    If keycode = 34 Then
      myPrefs_PlayfieldGlass = myPrefs_PlayfieldGlass + 1
      if myPrefs_PlayfieldGlass = 1 then:Playsound "GlassOn":end if
    End If
      If myPrefs_PlayfieldGlass > 1 Then
        Playsound "GlassOff"
        myPrefs_PlayfieldGlass = 0
      End If

'**********************************************************************

    If vpmKeyDown(keycode) Then Exit Sub
   End Sub

   Sub Table1_KeyUp(ByVal keycode)

    If keycode=RightFlipperKey Then Controller.Switch(48)=0
    If keycode=LeftFlipperKey Then Controller.Switch(47)=0
    If keycode=  RightMagnaSave Then Controller.Switch(46)=0
    If keycode= LeftMagnaSave Then Controller.Switch(45)=0

    If keycode = PlungerKey And myPrefs_PlayfieldGlass = 0 Then
      Plunger1.Fire
      If controller.switch(11) = True Then
        PlaySoundAtVol "Plunger_Release_Ball", Plunger1, 1
      Else
        PlaySoundAtVol "Plunger_Release_No_Ball", Plunger1, 1
      End If
    End If

    If keycode = PlungerKey And myPrefs_PlayfieldGlass = 1 Then
      Plunger1.Fire
      If controller.switch(11) = True Then
        PlaySoundAtVol "Plunger_Release_BallG", Plunger1, 1
      Else
        PlaySoundAtVol "Plunger_Release_No_BallG", Plunger1, 1
      End If
    End If

    If vpmKeyUp(keycode) Then Exit Sub

  End Sub

'===============================================


Function myBallCount
  Dim iCounter
  Dim myCount

  iCounter = 0
  For iCounter = 1 to 4
    myCount = RndNum(1,4)
  Next
    myBallCount = myCount
End Function
'************************************************************************************
'*****************                  NUDGE                 ***************************
'*****************     based on Noah's nudgetest table    ***************************
'************************************************************************************
 Dim LeftNudgeEffect, RightNudgeEffect, NudgeEffect

 Sub LeftNudge(angle, strength, delay)
     vpmNudge.DoNudge angle, (strength * (delay-LeftNudgeEffect) / delay) + RightNudgeEffect / delay
     LeftNudgeEffect = delay
     RightNudgeEffect = 0
     RightNudgeTimer.Enabled = 0
     LeftNudgeTimer.Interval = delay
     LeftNudgeTimer.Enabled = 1
 End Sub

 Sub RightNudge(angle, strength, delay)
     vpmNudge.DoNudge angle, (strength * (delay-RightNudgeEffect) / delay) + LeftNudgeEffect / delay
     RightNudgeEffect = delay
     LeftNudgeEffect = 0
     LeftNudgeTimer.Enabled = 0
     RightNudgeTimer.Interval = delay
     RightNudgeTimer.Enabled = 1
 End Sub

 Sub CenterNudge(angle, strength, delay)
     vpmNudge.DoNudge angle, strength * (delay-NudgeEffect) / delay
     NudgeEffect = delay
     NudgeTimer.Interval = delay
     NudgeTimer.Enabled = 1
 End Sub

 Sub LeftNudgeTimer_Timer()
     LeftNudgeEffect = LeftNudgeEffect-1
     If LeftNudgeEffect = 0 then LeftNudgeTimer.Enabled = 0
 End Sub

 Sub RightNudgeTimer_Timer()
     RightNudgeEffect = RightNudgeEffect-1
     If RightNudgeEffect = 0 then RightNudgeTimer.Enabled = 0
 End Sub

 Sub NudgeTimer_Timer()
     NudgeEffect = NudgeEffect-1
     If NudgeEffect = 0 then NudgeTimer.Enabled = 0
 End Sub

'**********Sling Shot Animations*********************************
' Rstep and Lstep  are the variables that increment the animation
'****************************************************************

Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 44
  If myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol SoundFX("right_slingshotNG",DOFContactors), ActiveBall, 1
  End If
  If myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), ActiveBall, 1
  End If
    RSling.Visible = 0
    RSling1.Visible = 1
    RStep = 0
    SlingArmR.TransX = -10
    SlingArmR.Transz = -24
    RightSlingShot.TimerInterval = 20 'Speed of Animation
    RightSlingShot.TimerEnabled = True
End Sub

Sub RightSlingShot_Timer
    RStep = RStep + 1
    Select Case RStep
        Case 3
          RSLing1.Visible = 0
          RSLing2.Visible = 1
        Case 4
          RSLing2.Visible = 0
          RSLing.Visible = 1
          SlingArmR.TransX = 0
          SlingArmR.Transz = 0
        RightSlingShot.TimerEnabled = False
   End Select
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 43
  If myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol SoundFX("left_slingshotNG",DOFContactors), ActiveBall, 1
  End If
  If myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), ActiveBall, 1
  End If
    LSling.Visible = 0
    LSling1.Visible = 1
    LStep = 0
    SlingArmL.TransX = 10
    SlingArmL.Transz = -24
    LeftSlingShot.TimerInterval = 20 'Speed of Animation
    LeftSlingShot.TimerEnabled = True
End Sub

Sub LeftSlingShot_Timer
    LStep = LStep + 1
    Select Case LStep
        Case 3
          LSLing1.Visible = 0
          LSLing2.Visible = 1
        Case 4
         LSLing2.Visible = 0
         LSLing.Visible = 1
         SlingArmL.TransX = 0
         SlingArmL.Transz = 0
       LeftSlingShot.TimerEnabled = False
    End Select
End Sub

 Sub SlingHopL_Hit
   If myPrefs_Bounce = 1 Then
     'Msgbox "left Hit"
     If Activeball.velX > 10 Then
     'Msgbox Activeball.velx
     SlingHopLTimer.Enabled = 1
     Activeball.velZ = 10
   End If
  End If
 End Sub

 Sub SlingHopLTimer_Timer
    If myPrefs_PlayfieldGlass = 0 Then
      PlaySoundAtVol "ball_bounce", SlingHopL, 1
    Else
      PlaySoundAtVol "ball_bounceG", SlingHopL, 1
    End If
     SlingHopLTimer.Enabled = 0
   End Sub

 Sub SlingHopR_Hit
   If myPrefs_Bounce = 1 Then
     'Msgbox "Right hit"
     If Activeball.velX > 10 Then
     'Msgbox Activeball.velx
     SlingHopRTimer.Enabled = 1
     Activeball.velZ = 10
   End If
  End If
 End Sub

 Sub SlingHopRTimer_Timer
    If myPrefs_PlayfieldGlass = 0 Then
      iwPlaySoundAtVollaysound "ball_bounce", SlingHopR, 1
    Else
      PlaySoundAtVol "ball_bounceG", SlingHopR, 1
    End If
     SlingHopRTimer.Enabled = 0
   End Sub

'**************************************
'********       TARGETS        ********
'**************************************

Sub TargetD_7M_Hit
  Call myTargetHit(TargetD_7, TargetButtonD_10, 38)
End Sub
Sub TargetR_7M_Hit
  Call myTargetHit(TargetR_7, TargetButtonR_10, 39)
End Sub
Sub TargetA_7M_Hit
  Call myTargetHit(TargetA_7, TargetButtonA_10, 40)
End Sub
Sub TargetZ_7M_Hit
  Call myTargetHit(TargetZ_7, TargetButtonZ_10, 37)
End Sub
Sub TargetI_7M_Hit
  Call myTargetHit(TargetI_7, TargetButtonI_10, 36)
End Sub
Sub TargetL_7M_Hit
  Call myTargetHit(TargetL_7, TargetButtonL_10, 35)
End Sub
Sub TargetE_7_Hit
  Call myTargetHit(TargetE_7, TargetButtonE_10, 41)
End Sub
Sub TargetM_7_Hit
  Call myTargetHit(TargetM_7, TargetButtonM_10, 24)
End Sub

Sub myTargetHit( myTarget1, myTarget2, mySwitchNum)
  vpmTimer.PulseSwitch mySwitchNum, 0, ""
  myTarget1.playanim 0,.6
  myTarget2.playanim 0,.6
  If myPrefs_PlayfieldGlass = 1 Then
      PlaySoundAtVol "target", myTarget1, 1
  End If
  If myPrefs_PlayfieldGlass = 0 Then
      PlaySoundAtVol "targetNG", myTarget1, 1
  End If
End Sub

'****************************************************************************
'*****************               FLIPPERS                 *******************
'****************************************************************************

 Sub LFlipper(Enabled)
   If Enabled Then
  If myPrefs_PlayfieldGlass = 1 Then
    Select Case Int(Rnd()*6)
      Case 0
        PlaySoundAtVol"Flipper_L01", FlipperLeft, 1
        PlaySoundAtVol"Flipper_L01", FlipperLeftUpper, 1
      Case 1
        PlaySoundAtVol"Flipper_L02", FlipperLeft, 1
        PlaySoundAtVol"Flipper_L02", FlipperLeftUpper, 1
      Case 2
        PlaySoundAtVol"Flipper_L07", FlipperLeft, 1
        PlaySoundAtVol"Flipper_L07", FlipperLeftUpper, 1
      Case 3
        PlaySoundAtVol"Flipper_L08", FlipperLeft, 1
        PlaySoundAtVol"Flipper_L08", FlipperLeftUpper, 1
      Case 4
        playsoundAtVol"Flipper_L09", FlipperLeft, 1
        playsoundAtVol"Flipper_L09", FlipperLeftUpper, 1
      Case 5
        PlaySoundAtVol"Flipper_L10", FlipperLeft, 1
        PlaySoundAtVol"Flipper_L10", FlipperLeftUpper, 1
    End Select
  End If
  If myPrefs_PlayfieldGlass = 0 Then
    Select Case Int(Rnd()*6)
      Case 0
        PlaySoundAtVol"Flipper_L01NG", FlipperLeft, 1
        PlaySoundAtVol"Flipper_L01NG", FlipperLeftUpper, 1
      Case 1
        PlaySoundAtVol"Flipper_L02NG", FlipperLeft, 1
        PlaySoundAtVol"Flipper_L02NG", FlipperLeftUpper, 1
      Case 2
        PlaySoundAtVol"Flipper_L07NG", FlipperLeft, 1
        PlaySoundAtVol"Flipper_L07NG", FlipperLeftUpper, 1
      Case 3
        PlaySoundAtVol"Flipper_L08NG", FlipperLeft, 1
        PlaySoundAtVol"Flipper_L08NG", FlipperLeftUpper, 1
      Case 4
        PlaySoundAtVol"Flipper_L09NG", FlipperLeft, 1
        PlaySoundAtVol"Flipper_L09NG", FlipperLeftUpper, 1
      Case 5
        PlaySoundAtVol"Flipper_L10NG", FlipperLeft, 1
        PlaySoundAtVol"Flipper_L10NG", FlipperLeftUpper, 1
    End Select
  End If
    FlipperLeft.RotateToEnd
    FlipperLeftUpper.RotateToEnd
                'myLightIntensity = 0 'SDS
                Call myChangeGIForFlip()
        'LightIntensityTimer.Enabled = 1
   Else
    PlaySoundAtVol "flipperdown", FlipperLeft, 1
    FlipperLeft.RotateToStart
    FlipperLeftUpper.RotateToStart
   End If
 End Sub

 Sub RFlipper(Enabled)
   If Enabled Then
  If myPrefs_PlayfieldGlass = 1 Then
    Select Case Int(Rnd()*6)
      Case 0
        PlaySoundAtVol"Flipper_R01", FlipperRight, 1
        PlaySoundAtVol"Flipper_R01", FlipperRightUpper, 1
      Case 1
        PlaySoundAtVol"Flipper_R02", FlipperRight, 1
        PlaySoundAtVol"Flipper_R02", FlipperRightUpper, 1
      Case 2
        PlaySoundAtVol"Flipper_R03", FlipperRight, 1
        PlaySoundAtVol"Flipper_R03", FlipperRightUpper, 1
      Case 3
        PlaySoundAtVol"Flipper_R04", FlipperRight, 1
        PlaySoundAtVol"Flipper_R04", FlipperRightUpper, 1
      Case 4
        PlaySoundAtVol"Flipper_R05", FlipperRight, 1
        PlaySoundAtVol"Flipper_R05", FlipperRightUpper, 1
      Case 5
        PlaySoundAtVol"Flipper_R06", FlipperRight, 1
        PlaySoundAtVol"Flipper_R06", FlipperRightUpper, 1
    End Select
  End If
  If myPrefs_PlayfieldGlass = 0 Then
    Select Case Int(Rnd()*6)
      Case 0
        PlaySoundAtVol"Flipper_R01NG", FlipperRight, 1
        PlaySoundAtVol"Flipper_R01NG", FlipperRightUpper, 1
      Case 1
        PlaySoundAtVol"Flipper_R02NG", FlipperRight, 1
        PlaySoundAtVol"Flipper_R02NG", FlipperRightUpper, 1
      Case 2
        PlaySoundAtVol"Flipper_R03NG", FlipperRight, 1
        PlaySoundAtVol"Flipper_R03NG", FlipperRightUpper, 1
      Case 3
        PlaySoundAtVol"Flipper_R04NG", FlipperRight, 1
        PlaySoundAtVol"Flipper_R04NG", FlipperRightUpper, 1
      Case 4
        PlaySoundAtVol"Flipper_R05NG", FlipperRight, 1
        PlaySoundAtVol"Flipper_R05NG", FlipperRightUpper, 1
      Case 5
        PlaySoundAtVol"Flipper_R06NG", FlipperRight, 1
        PlaySoundAtVol"Flipper_R06NG", FlipperRightUpper, 1
    End Select
  End If
    FlipperRight.RotateToEnd
    FlipperRightUpper.RotateToEnd
                'myLightIntensity = 0
                'LightIntensityTimer.Enabled = 1
                Call myChangeGIForFlip()
   Else
    PlaySoundAtVol "flipperdown", FlipperRight, 1
    FlipperRight.RotateToStart
    FlipperRightUpper.RotateToStart
   End If
 End Sub

Sub myChangeGIForFlip()
  For Each myObject in GI
  myObject.Intensity = 205
  GI_14.Intensity = 8
  GI_2.Intensity = 22
  LightIntensityTimer.Enabled = 1
    Next
End Sub

Sub LightIntensityTimer_Timer
  For Each myObject in GI
  myObject.Intensity = 250
  GI_14.Intensity = 10
  GI_2.Intensity = 30
  LightIntensityTimer.Enabled = 0
  Next
End Sub

 Sub FlipperLeft_Collide(parm)
  if myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol "rubber_flipper", ActiveBall, 1
  end if
  if myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol "rubber_flipperG", ActiveBall, 1
  end if
 End Sub

 Sub FlipperRight_Collide(parm)
  if myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol "rubber_flipper", ActiveBall, 1
  end if
  if myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol "rubber_flipperG", ActiveBall, 1
  end if
 End Sub

 Sub FlipperUpperLeft_Collide(parm)
  if myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol "rubber_flipper", ActiveBall, 1
  end if
  if myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol "rubber_flipperG", ActiveBall, 1
  end if
 End Sub

 Sub FlipperUpperRight_Collide(parm)
  if myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol "rubber_flipper", ActiveBall, 1
  end if
  if myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol "rubber_flipperG", ActiveBall, 1
  end if
 End Sub



'************************************************************************************
'*****************               SOLENOIDS                ***************************
'************************************************************************************

 SolCallback(sLRFlipper) = "RFlipper"
 SolCallback(sLLFlipper) = "LFlipper"
 SolCallback(1)="bsTrough.SolIn"
 SolCallBack(8)="bsTrough.SolOut"
 SolCallBack(2)="dt3B.SolDropUp"
 SolCallBack(3)="dt4L.SolDropUp"
 SolCallBack(4)="dt4R.SolDropUp"
 SolCallBack(5)="LockKick"
 SolCallBack(11)="PFGI" 'GI
 SolCallBack(13)= "SetFlasher66" 'Eyes
 SolCallBack(12)= "SetFlasher67" '3 Bank
 SolCallBack(6)= "SetFlasher68" 'Lock
 SolCallBack(7)= "SetFlasher69" 'Lmag
 SolCallBack(19)= "SetFlasher70" 'Rmag
 SolCallback(17)="vpmSolSound ""Slingshot"","
 SolCallback(18)="vpmSolSound ""Slingshot"","
 SolCallback(14)="vpmSolSoundknocker"
 SolCallback(15)="vpmSolSoundBell"

Sub sw20_Hit
  if myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol"droptarget", ActiveBall, 1
  end if
  if myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol"droptargetNG", ActiveBall, 1
  end if
End Sub

Sub sw21_Hit
  if myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol"droptarget", ActiveBall, 1
  end if
  if myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol"droptargetNG", ActiveBall, 1
  end if
End Sub

Sub sw22_Hit
  if myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol"droptarget", ActiveBall, 1
  end if
  if myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol"droptargetNG", ActiveBall, 1
  end if
End Sub

Sub sw23_Hit
  if myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol"droptarget", ActiveBall, 1
  end if
  if myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol"droptargetNG", ActiveBall, 1
  end if
End Sub

Sub sw25_Hit
  if myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol"droptarget", ActiveBall, 1
  end if
  if myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol"droptargetNG", ActiveBall, 1
  end if
End Sub

Sub sw26_Hit
  if myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol"droptarget", ActiveBall, 1
  end if
  if myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol"droptargetNG", ActiveBall, 1
  end if
End Sub

Sub sw27_Hit
  if myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol"droptarget", ActiveBall, 1
  end if
  if myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol"droptargetNG", ActiveBall, 1
  end if
End Sub

 Sub vpmSolSoundBell(enabled) 'Bell
    If enabled Then
    Playsound "Bell"
    Playsound "Bell"
    Playsound "Bell"
  End If
End Sub

 Sub vpmSolSoundknocker(enabled) 'Bell
    If enabled Then
    Playsound "knocker"
    Playsound "knocker"
    Playsound "knocker"
  End If
End Sub

'*********************************************
'**********       FLASHERS        ************
'*********************************************



 Sub SetFlasher66(enabled)'Eyes
  If Enabled Then
  LizardHeadFS.disablelighting=1
  LizardHeadFS.Image="GrandLizard_Flasher66"
  LizardHeadFS.Material= "LizardheadFlasher"
  Flasher1.visible = true
  Flasher66b.state = 1
  Flasher66c.visible = true
  Else
  LizardHeadFS.disablelighting=0
  LizardHeadFS.Image="GrandLizard_LPDiffuseMap"
  LizardHeadFS.Material= "Lizardhead"
  Flasher1.visible = false
  Flasher66b.state = 0
  Flasher66c.visible = False
  End If
End Sub

 Sub SetFlasher67(enabled)'3 Bank
  If Enabled Then
  LizardHeadFS.image="GrandLizard_Flasher67_69"
  LizardHeadFS.Material= "LizardheadFlasher"
  Flasher2.visible = true
  Flasher67.visible = true
  Flasher67b.state = true
  Else
  LizardHeadFS.image="GrandLizard_LPDiffuseMap"
  LizardHeadFS.Material= "Lizardhead"
  Flasher2.visible = False
  Flasher67.visible = false
  Flasher67b.state = False
  End If
End Sub

 Sub SetFlasher68(enabled)' Lock
  If Enabled Then
    Flasher3.visible = true
    Flasher68.visible = true
    Flasher68b.state = true
  Else
    Flasher3.visible = False
    Flasher68.visible = false
    Flasher68b.state = False
  End If
End Sub

 Sub SetFlasher69(enabled)'Lmag
  If Enabled Then
  LizardHeadFS.image="GrandLizard_Flasher67_69"
  LizardHeadFS.Material= "LizardheadFlasher"
  LizardHeadFS.Material= "Lizardhead"
  Flasher4.visible = True
  Flasher69.visible = true
  Flasher69b.state = true
  Else
  LizardHeadFS.image="GrandLizard_LPDiffuseMap"
    Flasher4.visible = False
    Flasher69.visible = false
    Flasher69b.state = False
  End If
End Sub

 Sub SetFlasher70(enabled)'Rmag
  If Enabled Then
  LizardHeadFS.image="GrandLizard_Flasher70"
    Flasher5.visible = True
    Flasher70.visible = true
    Flasher70b.state = true
  Else
  LizardHeadFS.image="GrandLizard_LPDiffuseMap"
    Flasher5.visible = False
    Flasher70.visible = false
    Flasher70b.state = False
  End If
End Sub

Sub LockKick(enabled)
    if enabled then
       sw34.Kick 100, 15
  If myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol "Kickout_0_Balls_Locked", sw34, 1
    End If
  If myPrefs_PlayfieldGlass = 0 Then
      PlaySoundAtVol "Kickout_0_Balls_LockedNG", sw34, 1
    End If
  If myBallCount = 2 Then
  If myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol "Kickout_1_Balls_Locked", sw34, 1
    End If
  If myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol "Kickout_1_Balls_LockedNG", sw34, 1
    End If
  End If
  If myBallCount = 3 Then
    If myPrefs_PlayfieldGlass = 1 Then
      PlaysoundAtVol "Kickout_2_Balls_Locked", sw34, 1
      End If
    If myPrefs_PlayfieldGlass = 0 Then
      PlaySoundAtVol "Kickout_2_Balls_LockedNG", sw34, 1
      End If
  End If
    end if
 End Sub

'*********************************************
'**********  GENERAL ILLUMINATION  ***********
'*********************************************

Sub PFGI(Enabled)
  If Enabled Then
      GI_13.State = 0
      GI_2.State = 0
      GI_14.State = 0
      'Light101.Intensity = .1
      Flasher6.visible = False
      BladeR.Material="BladesD"
      BladeL.Material="BladesD"
      batleft.Material="g_Plastic2"
      batright.Material="g_Plastic2"
      batleft1.Material="g_Plastic2"
      batright1.Material="g_Plastic2"
  if myPrefs_DynamicBackground = 0 then
      Ramp1.visible = False
      Ramp2.visible = False
                        SideRails.Material="Metal2"
                        LockdownBar.Material="Metal2"
    Else
                        SideRails.Material="Metal21"
                        LockdownBar.Material="Metal21"
                  Ramp1.visible = True
      Ramp2.visible = True
    End If
      CabinetWalls.Material="BladesD"
                        Gates.Material="Metal2"
      LockgateBracket.Material="Metal2"
      wireform.Material="Metal3"
                        MetalWork.Material="Metal3"
                        BounceBackBars.Material="Metal2"
                        PlasticNuts.Material = "PlasticNutsD"
                        PostOutlaneRight1.Material = "PostRedD"
                        PostOutlaneRight2.Material = "PostRedD"
                        PostOutlaneRight3.Material = "PostRedD"
                        PostOutlaneLeft1.Material = "PostRedD"
                        PostOutlaneLeft2.Material = "PostRedD"
                        PostOutlaneLeft3.Material = "PostRedD"
                        sw20.Material = "PlasticD"
      sw21.Material = "PlasticD"
      sw22.Material = "PlasticD"
      sw23.Material = "PlasticD"
      sw25.Material = "PlasticD"
      sw26.Material = "PlasticD"
      sw27.Material = "PlasticD"
                        RampBlades.Material= "RampBladesD"
                        PlasticSlingshots.Material= "PlasticTriggerD"
      CPlastic.Material= "ClearPlasticD"
      Plastic2.Material= "PlasticTriggerD"
      Plastic3.Material= "PlasticTriggerD"
      Plastic5.Material= "PlasticTriggerD"
      Plastic6.Material= "PlasticTriggerD"
      Plastic7.Material= "PlasticTriggerD"
      Plastic8.Material= "PlasticTriggerD"
      Plastic9.Material= "PlasticTriggerD"
      Plastic11.Material= "PlasticTriggerD"
      LizardHeadFS.Material= "LizardheadD"
      Ramp.Material="RightRampMAT2"
    dim xx
  For each xx in GI:xx.State = 0: Next
    if myPrefs_PlayfieldGlass = 0 Then
                 PlaySound "fx_relay"
    end if
    if myPrefs_PlayfieldGlass = 1 Then
                 PlaySound "fx_relayG"
    end if
                 Call mySetRubberColor(False)
           Call mySetRubberSleeveColor(False)
           Call mySetPostColor(False)
  Else
      CabinetWalls.Material="Blades"
      Light101.Intensity = 5
      GI_13.State = 1
      GI_2.State = 1
      GI_14.State = 1
      Flasher6.visible = True
      BladeR.Material="Blades"
      BladeL.Material="Blades"
      batleft.Material="g_Plastic"
      batright.Material="g_Plastic"
      batleft1.Material="g_Plastic"
      batright1.Material="g_Plastic"
  if myPrefs_DynamicBackground = 0 then
'                        SideRails.Material="Metal21"
'                        LockdownBar.Material="Metal21"
'                 Ramp1.visible = True
'     Ramp2.visible = True
    Else
      Ramp1.visible = False
      Ramp2.visible = False
                        SideRails.Material="Metal2"
                        LockdownBar.Material="Metal2"
    End If
      Gates.Material="Metal1"
      LockgateBracket.Material="Metal1"
      wireform.Material="Metal2"
      MetalWork.Material="Metal2"
                        BounceBackBars.Material="Metal1"
      PlasticNuts.Material = "PlasticNuts"
                        PostOutlaneRight1.Material = "PostRed"
                        PostOutlaneRight2.Material = "PostRed"
                        PostOutlaneRight3.Material = "PostRed"
                        PostOutlaneLeft1.Material = "PostRed"
                        PostOutlaneLeft2.Material = "PostRed"
                        PostOutlaneLeft3.Material = "PostRed"
                        sw20.Material = "Plastic"
      sw21.Material = "Plastic"
      sw22.Material = "Plastic"
      sw23.Material = "Plastic"
      sw25.Material = "Plastic"
      sw26.Material = "Plastic"
      sw27.Material = "Plastic"
                        RampBlades.Material= "RampBlades"
      CPlastic.Material = "ClearPlastic"
      PlasticSlingshots.Material= "PlasticTrigger"
      Plastic2.Material= "PlasticTrigger"
      Plastic3.Material= "PlasticTrigger"
      Plastic5.Material= "PlasticTrigger"
      Plastic6.Material= "PlasticTrigger"
      Plastic7.Material= "PlasticTrigger"
      Plastic8.Material= "PlasticTrigger"
      Plastic9.Material= "PlasticTrigger"
      Plastic11.Material= "PlasticTrigger"
      LizardHeadFS.Material="Lizardhead"
      Ramp.Material="RightRampMAT1"
    For each xx in GI:xx.State = 1: Next
    if myPrefs_PlayfieldGlass = 0 Then
                 PlaySound "fx_relay"
    end if
    if myPrefs_PlayfieldGlass = 1 Then
                 PlaySound "fx_relayG"
    end if
      Call mySetRubberColor(True)
            Call mySetRubberSleeveColor(True)
           Call mySetPostColor(True)
    End If
End Sub


'************************************************************************************
'*****************               SWITCHES                ****************************
'************************************************************************************

 Sub sw11_Hit():Controller.Switch(11)=1:End Sub
 Sub sw11_Unhit():Controller.Switch(11)=0:End Sub
 Sub sw15_Hit(): vpmTimer.PulseSwitch 15, 0, "": End Sub
 Sub sw16_Hit(): vpmTimer.PulseSwitch 16, 0, "":End Sub
 Sub sw17_Hit(): vpmTimer.PulseSwitch 17, 0, "":End Sub
 Sub sw18_Hit(): vpmTimer.PulseSwitch 18, 0, "":End Sub
 Sub sw19_Hit(): vpmTimer.PulseSwitch 19, 0, "":End Sub
 Sub sw20_Hit():dt4L.Hit 1:End Sub
 Sub sw21_Hit():dt4L.Hit 2:End Sub
 Sub sw22_Hit():dt4R.Hit 1:End Sub
 Sub sw23_Hit():dt4R.Hit 2:End Sub
 Sub sw25_Hit():dt3B.Hit 1:End Sub
 Sub sw26_Hit():dt3B.Hit 2:End Sub
 Sub sw27_Hit():dt3B.Hit 3:End Sub
 Sub sw28_Spin()
vpmTimer.PulseSwitch 28, 0, ""
If myPrefs_PlayfieldGlass = 0 Then
PlaySoundAtVol "spinner", sw28, 1
End If
If myPrefs_PlayfieldGlass = 1 Then
PlaySoundAtVol "spinnerG", sw28, 1
End If
End Sub
 Sub sw29_Hit(): vpmTimer.PulseSwitch 29, 0, "": End Sub
 Sub sw30_Hit(): vpmTimer.PulseSwitch 30, 0, "": End Sub
 Sub sw31_Hit():vpmTimer.PulseSwitch 31, 0, "":End Sub
 Sub sw32_Hit():Controller.Switch(32)=1:End Sub
 Sub sw32_Unhit():Controller.Switch(32)=0:End Sub
 Sub sw33_Hit():Controller.Switch(33)=1:End Sub
 Sub sw33_Unhit():Controller.Switch(33)=0:End Sub
 Sub sw34_Hit():Controller.Switch(34)=1:End Sub
 Sub sw34_Unhit():Controller.Switch(34)=0:End Sub
 Sub sw42_Hit(): vpmTimer.PulseSwitch 42, 0, "": End Sub
 Sub sw43_Slingshot():vpmTimer.PulseSwitch 43, 0, "":End Sub
 Sub sw44_Slingshot():vpmTimer.PulseSwitch 44, 0, "":End Sub
 Sub Drain_Hit()
'  ClearBallID
   bsTrough.AddBall Me
  If myPrefs_PlayfieldGlass = 0 Then
    Select Case Int(Rnd()*4)
      Case 0
        PlaySoundAtVol "Drain_1", Drain, 1
      Case 1
        PlaySoundAtVol "Drain_2", Drain, 1
      Case 2
        PlaySoundAtVol "Drain_3", Drain, 1
      Case 3
        PlaySoundAtVol "Drain_4", Drain, 1
    End Select
  End If
  If myPrefs_PlayfieldGlass = 1 Then
    Select Case Int(Rnd()*4)
      Case 0
        PlaySoundAtVol "Drain_1NG", Drain, 1
      Case 1
        PlaySoundAtVol "Drain_2NG", Drain, 1
      Case 2
        PlaySoundAtVol "Drain_3NG", Drain, 1
      Case 3
        PlaySoundAtVol "Drain_4NG", Drain, 1
    End Select
  End If
End Sub

 Sub StackingOpto_Hit()
  stopsound("JungleJam2") 'Ending Song
  stopsound("ConanEnding")
 myPlayingEndGameSnd = False
End Sub


'*************************************************************************************
'******************              SOUNDS                   ****************************
'*************************************************************************************

Dim RubberHitLoud, RubberHitSoft

Sub SoundTimer_Timer()

  If RubberHitLoud = 1 Then
    RubberHitLoud = 0
    If myPrefs_PlayfieldGlass = 0 Then
      Select Case Int(Rnd()*8)
        Case 0
          PlaySound "Rubber_1"
        Case 1
          PlaySound "Rubber_2"
        Case 2
          PlaySound "Rubber_3"
        Case 3
          PlaySound "Rubber_4"
        Case 4
          PlaySound "Rubber_5"
        Case 5
          PlaySound "Rubber_6"
        Case 6
          PlaySound "Rubber_7"
        Case 7
          PlaySound "Rubber_8"
      End Select
    end If
    If myPrefs_PlayfieldGlass = 1 Then
      Select Case Int(Rnd()*8)
        Case 0
          PlaySound "Rubber_1G"
        Case 1
          PlaySound "Rubber_2G"
        Case 2
          PlaySound "Rubber_3G"
        Case 3
          PlaySound "Rubber_4G"
        Case 4
          PlaySound "Rubber_5G"
        Case 5
          PlaySound "Rubber_6G"
        Case 6
          PlaySound "Rubber_7G"
        Case 7
          PlaySound "Rubber_8G"
      End Select
    end If
  end if
  If RubberHitSoft = 1 Then
    RubberHitSoft = 0
    If myPrefs_PlayfieldGlass = 0 Then
      Select Case Int(Rnd()*8)
        Case 0
          PlaySound "Rubber_1"
        Case 1
          PlaySound "Rubber_2"
        Case 2
          PlaySound "Rubber_3"
        Case 3
          PlaySound "Rubber_4"
        Case 4
          PlaySound "Rubber_5"
        Case 5
          PlaySound "Rubber_6"
        Case 6
          PlaySound "Rubber_7"
        Case 7
          PlaySound "Rubber_8"
      End Select
      end If
    If myPrefs_PlayfieldGlass = 1 Then
      Select Case Int(Rnd()*8)
        Case 0
          PlaySound "Rubber_1G"
        Case 1
          PlaySound "Rubber_2G"
        Case 2
          PlaySound "Rubber_3G"
        Case 3
          PlaySound "Rubber_4G"
        Case 4
          PlaySound "Rubber_5G"
        Case 5
          PlaySound "Rubber_6G"
        Case 6
          PlaySound "Rubber_7G"
        Case 7
          PlaySound "Rubber_8G"
      End Select
    end If
  end if
End Sub

 Sub RubberSleeves_Hit(index)
  If (Abs(int(Activeball.velx)) or Abs(int(Activeball.vely))) > 5 then
    RubberHitLoud = 1
  Else
    RubberHitSoft = 1
  End If
 End Sub

 Sub Rubbers_Hit(index)
  If (Abs(int(Activeball.velx)) or Abs(int(Activeball.vely))) > 5 then
    RubberHitLoud = 1
  Else
    RubberHitSoft = 1
  end if
 End Sub

 Sub M02_Hit
  If myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol "FX_metalhit2", ActiveBall, 1
  End If
  If myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol "FX_metalhit2NG", ActiveBall, 1
  End If
 End Sub

 Sub M03_Hit
  If myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol "FX_metalhit2", ActiveBall, 1
  End If
  If myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol "FX_metalhit2NG", ActiveBall, 1
  End If
 End Sub

 Sub M06_Hit
  If myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol "FX_metalhit2", ActiveBall, 1
  End If
  If myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol "FX_metalhit2NG", ActiveBall, 1
  End If
 End Sub

 Sub Wall860_Hit
  If myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol "FX_metalhit2", ActiveBall, 1
  End If
  If myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol "FX_metalhit2NG", ActiveBall, 1
  End If
 End Sub

 Sub Wall862_Hit
  If myPrefs_PlayfieldGlass = 1 Then
    PlaySoundAtVol "FX_metalhit2", ActiveBall, 1
  End If
  If myPrefs_PlayfieldGlass = 0 Then
    PlaySoundAtVol "FX_metalhit2NG", ActiveBall, 1
  End If
 End Sub


 Sub Trigger1_Hit
    Playsound "Lion-Roar"
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
				  ' PlaySound("fx_ballrolling" & b), -1, BallRollVol(BOT(b) )*.8, AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
				  ' PlaySound("fx_ballrolling" & b), -1, BallRollVol(BOT(b) )*.2, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*****************
'***** LAMPS *****
'*****************

 'Set Lights(1)  = Light1
  Set Lights(2)  = Light2
 'Set Lights(3)  = Light3
 'Set Lights(4)  = Light4
 'Set Lights(5)  = Light5
 'Set Lights(6)  = Light6
  Set Lights(7)  = Light7
  Set Lights(8)  = Light8
  Set Lights(9)  = Light9
  Set Lights(10) = Light10
  Set Lights(11) = Light11
  Set Lights(12) = Light12
  Set Lights(13) = Light13
  Set Lights(14) = Light14
  Set Lights(15) = Light15
  Set Lights(16) = Light16
  Set Lights(17) = Light17
  Set Lights(18) = Light18
  Set Lights(19) = Light19
  Set Lights(20) = Light20
  Set Lights(21) = Light21
  Set Lights(22) = Light22
  Set Lights(23) = Light23
  Set Lights(24) = Light24
  Set Lights(25) = Light25
  Set Lights(26) = Light26
  Set Lights(27) = Light27
  Set Lights(28) = Light28
  Set Lights(29) = Light29
  Set Lights(30) = Light30
  Set Lights(31) = Light31
  Set Lights(32) = Light32
  Set Lights(33) = Light33
  Set Lights(34) = Light34
  Set Lights(35) = Light35
  Set Lights(36) = Light36
  Set Lights(37) = Light37
  Set Lights(38) = Light38
  Set Lights(39) = Light39
  Set Lights(40) = Light40
  Set Lights(41) = Light41
  Set Lights(42) = Light42
  Set Lights(43) = Light43
  Set Lights(44) = Light44
  Set Lights(45) = Light45
  Set Lights(46) = Light46
  Set Lights(47) = Light47
  Set Lights(48) = Light48
  Set Lights(49) = Light49
  Set Lights(50) = Light50
 'Set Lights(51) = Light51
  Set Lights(52) = Light52
  Set Lights(53) = Light53
  Set Lights(54) = Light54
  Set Lights(55) = Light55
  Set Lights(56) = Light56
  Set Lights(57) = Light57
  Set Lights(58) = Light58
  Set Lights(59) = Light59
  Set Lights(60) = Light60
  Set Lights(61) = Light61
  Set Lights(62) = Light62
  Set Lights(63) = Light63
  Set Lights(64) = Light64

Sub LightCopy_TIMER

   Light7b.State = Light7.State
   Light8b.State = Light8.State
   Light9b.State = Light9.State
   Light10b.State = Light10.State
   Light11b.State = Light11.State
   Light12b.State = Light12.State
   Light13b.State = Light13.State
   Light14b.State = Light14.State
   Light15b.State = Light15.State
   Light16b.State = Light16.State
   Light17b.State = Light17.State
   Light18b.State = Light18.State
   Light19b.State = Light19.State
   Light20b.State = Light20.State
   Light21b.State = Light21.State
   Light22b.State = Light22.State
   Light23b.State = Light23.State
   Light24b.State = Light24.State
   Light25b.State = Light25.State
   Light26b.State = Light26.State
   Light27b.State = Light27.State
   Light28b.State = Light28.State
   Light29b.State = Light29.State
   Light30b.State = Light30.State
   Light31b.State = Light31.State
   Light32b.State = Light32.State
   Light33b.State = Light33.State
   Light34b.State = Light34.State
   Light35b.State = Light35.State
   Light36b.State = Light36.State
   Light37b.State = Light37.State
   Light38b.State = Light38.State
   Light39b.State = Light39.State
   Light40b.State = Light40.State
   Light41b.State = Light41.State
   Light42b.State = Light42.State
   Light43b.State = Light43.State
   Light44b.State = Light44.State
   Light45b.State = Light45.State
   Light46b.State = Light46.State
   Light47b.State = Light47.State
   Light48b.State = Light48.State
   Light49b.State = Light49.State
   Light50b.State = Light50.State
   Light52b.State = Light52.State
   Light53b.State = Light53.State
   Light54b.State = Light54.State
   Light55b.State = Light55.State
   Light56b.State = Light56.State
   Light57b.State = Light57.State
   Light58b.State = Light58.State
   Light59b.State = Light59.State
   Light60b.State = Light60.State
   Light61b.State = Light61.State
   Light62b.State = Light62.State
   Light63b.State = Light63.State
   Light64b.State = Light64.State

  If Light7.state=1 Then:Light7.image="pf_AXS5":Else:Light7.image="pf_AXS6":End If
  If Light8.state=1 Then:Light8.image="pf_AXS5":Else:Light8.image="pf_AXS6":End If
  If Light9.state=1 Then:Light9.image="pf_AXS5":Else:Light9.image="pf_AXS6":End If
  If Light9.state=1 Then:Light9.image="pf_AXS5":Else:Light9.image="pf_AXS6":End If

  If Light10.state=1 Then:Light10.image="pf_AXS5":Else:Light10.image="pf_AXS6":End If
  If Light11.state=1 Then:Light11.image="pf_AXS5":Else:Light11.image="pf_AXS6":End If
  If Light12.state=1 Then:Light12.image="pf_AXS5":Else:Light12.image="pf_AXS6":End If
  If Light13.state=1 Then:Light13.image="pf_AXS5":Else:Light13.image="pf_AXS6":End If
  If Light14.state=1 Then:Light14.image="pf_AXS5":Else:Light14.image="pf_AXS6":End If
  If Light15.state=1 Then:Light15.image="pf_AXS5":Else:Light15.image="pf_AXS6":End If
  If Light16.state=1 Then:Light16.image="pf_AXS5":Else:Light16.image="pf_AXS6":End If
  If Light17.state=1 Then:Light17.image="pf_AXS5":Else:Light17.image="pf_AXS6":End If
  If Light18.state=1 Then:Light18.image="pf_AXS5":Else:Light18.image="pf_AXS6":End If
  If Light19.state=1 Then:Light19.image="pf_AXS5":Else:Light19.image="pf_AXS6":End If
  If Light20.state=1 Then:Light20.image="pf_AXS5":Else:Light20.image="pf_AXS6":End If

  If Light21.state=1 Then:Light21.image="pf_AXS5":Else:Light21.image="pf_AXS6":End If
  If Light22.state=1 Then:Light22.image="pf_AXS5":Else:Light22.image="pf_AXS6":End If
  If Light23.state=1 Then:Light23.image="pf_AXS5":Else:Light23.image="pf_AXS6":End If
  If Light24.state=1 Then:Light24.image="pf_AXS5":Else:Light24.image="pf_AXS6":End If
  If Light25.state=1 Then:Light25.image="pf_AXS5":Else:Light25.image="pf_AXS6":End If
  If Light26.state=1 Then:Light26.image="pf_AXS5":Else:Light26.image="pf_AXS6":End If
  If Light27.state=1 Then:Light27.image="pf_AXS5":Else:Light27.image="pf_AXS6":End If
  If Light28.state=1 Then:Light28.image="pf_AXS5":Else:Light28.image="pf_AXS6":End If
  If Light28.state=1 Then:Light28.image="pf_AXS5":Else:Light28.image="pf_AXS6":End If
  If Light29.state=1 Then:Light29.image="pf_AXS5":Else:Light29.image="pf_AXS6":End If
  If Light30.state=1 Then:Light30.image="pf_AXS5":Else:Light30.image="pf_AXS6":End If

  If Light31.state=1 Then:Light31.image="pf_AXS5":Else:Light31.image="pf_AXS6":End If
  If Light32.state=1 Then:Light32.image="pf_AXS5":Else:Light32.image="pf_AXS6":End If
  If Light33.state=1 Then:Light33.image="pf_AXS5":Else:Light33.image="pf_AXS6":End If
  If Light34.state=1 Then:Light34.image="pf_AXS5":Else:Light34.image="pf_AXS6":End If
  If Light35.state=1 Then:Light35.image="pf_AXS5":Else:Light35.image="pf_AXS6":End If
  If Light36.state=1 Then:Light36.image="pf_AXS5":Else:Light36.image="pf_AXS6":End If
  If Light37.state=1 Then:Light37.image="pf_AXS5":Else:Light37.image="pf_AXS6":End If
  If Light38.state=1 Then:Light38.image="pf_AXS5":Else:Light38.image="pf_AXS6":End If
  If Light38.state=1 Then:Light38.image="pf_AXS5":Else:Light38.image="pf_AXS6":End If
  If Light39.state=1 Then:Light39.image="pf_AXS5":Else:Light39.image="pf_AXS6":End If
  If Light40.state=1 Then:Light40.image="pf_AXS5":Else:Light40.image="pf_AXS6":End If

  If Light41.state=1 Then:Light41.image="pf_AXS5":Else:Light41.image="pf_AXS6":End If
  If Light42.state=1 Then:Light42.image="pf_AXS5":Else:Light42.image="pf_AXS6":End If
  If Light43.state=1 Then:Light43.image="pf_AXS5":Else:Light43.image="pf_AXS6":End If
  If Light44.state=1 Then:Light44.image="pf_AXS5":Else:Light44.image="pf_AXS6":End If
  If Light45.state=1 Then:Light45.image="pf_AXS5":Else:Light45.image="pf_AXS6":End If
  If Light46.state=1 Then:Light46.image="pf_AXS5":Else:Light46.image="pf_AXS6":End If
  If Light47.state=1 Then:Light47.image="pf_AXS5":Else:Light47.image="pf_AXS6":End If
  If Light48.state=1 Then:Light48.image="pf_AXS5":Else:Light48.image="pf_AXS6":End If
  If Light48.state=1 Then:Light48.image="pf_AXS5":Else:Light48.image="pf_AXS6":End If
  If Light49.state=1 Then:Light49.image="pf_AXS5":Else:Light49.image="pf_AXS6":End If
  If Light50.state=1 Then:Light50.image="pf_AXS5":Else:Light50.image="pf_AXS6":End If

  'If Light51.state=1 Then:Light51.image="pf_AXS5":Else:Light51.image="pf_AXS6":End If
  If Light52.state=1 Then:Light52.image="pf_AXS5":Else:Light52.image="pf_AXS6":End If
  If Light53.state=1 Then:Light53.image="pf_AXS5":Else:Light53.image="pf_AXS6":End If
  If Light54.state=1 Then:Light54.image="pf_AXS5":Else:Light54.image="pf_AXS6":End If
  If Light55.state=1 Then:Light55.image="pf_AXS5":Else:Light55.image="pf_AXS6":End If
  If Light56.state=1 Then:Light56.image="pf_AXS5":Else:Light56.image="pf_AXS6":End If
  If Light57.state=1 Then:Light57.image="pf_AXS5":Else:Light57.image="pf_AXS6":End If
  If Light58.state=1 Then:Light58.image="pf_AXS5":Else:Light58.image="pf_AXS6":End If
  If Light58.state=1 Then:Light58.image="pf_AXS5":Else:Light58.image="pf_AXS6":End If
  If Light59.state=1 Then:Light59.image="pf_AXS5":Else:Light59.image="pf_AXS6":End If
  If Light60.state=1 Then:Light60.image="pf_AXS5":Else:Light60.image="pf_AXS6":End If

  If Light61.state=1 Then:Light61.image="pf_AXS5":Else:Light61.image="pf_AXS6":End If
  If Light62.state=1 Then:Light62.image="pf_AXS5":Else:Light62.image="pf_AXS6":End If
  If Light63.state=1 Then:Light63.image="pf_AXS5":Else:Light63.image="pf_AXS6":End If
  If Light64.state=1 Then:Light64.image="pf_AXS5":Else:Light64.image="pf_AXS6":End If

'***************************************************
'**********   MUSIC,SOUND AND MOTION    ************
'***************************************************

  If Light2.State = 1 then 'End Game
    If  NOT myPlayingEndGameSnd Then
      If myPrefs_PlayMusic then
        PlaySound "JungleJam2"
        myPlayingEndGameSnd = True
      End If
    End If
  End If

End Sub

'==================================================================

Set MotorCallback = GetRef("RealTimeUpdates")

  Sub RealTimeUpdates
    GateLinkage.objRotx  = -10 -  Gate6.Currentangle / -20
    GateLinkage.objRoty  = -1 -  Gate6.Currentangle / -20
    WireGateR.rotx= -90 - Gate6.Currentangle / 1
    WireGateL.rotx= -90 - Gate5.Currentangle / 1
    PTounge.rotx = -90 - Gate1.Currentangle / 4.5
    Spinner.rotx = -90 - sw28.Currentangle/1
    Spinnerarm.rotx = -90 - sw28.Currentangle/1
    Gateflap.objRotx  = -10 - Gate3.Currentangle / 1
  End Sub

Set MotorCallback = GetRef("RealTimeUpdates")

'===================================================================

  Sub Wall6_hit
    If Activeball.velY > 4 Then
      Activeball.velZ = 8
      BallBounceTimer.Enabled = 1
      RubberHitLoud = 1
    Else
      RubberHitSoft = 1
    End if
  End sub

  Sub Wall7_hit
    If Activeball.velY > 3 Then
      Activeball.velZ = 8
      BallBounceTimer.Enabled = 1
      RubberHitLoud = 1
    Else
      RubberHitSoft = 1
    End if
  End sub

  Sub Wall8_hit
    If Activeball.velY > 3 Then
      Activeball.velZ = 8
      BallBounceTimer.Enabled = 1
      RubberHitLoud = 1
    Else
      RubberHitSoft = 1
    End if
  End sub

  Sub Wall9_hit
    If Activeball.velY > 4 Then
      Activeball.velZ = 8
      BallBounceTimer.Enabled = 1
      RubberHitLoud = 1
    Else
      RubberHitSoft = 1
    End if
  End sub

  Sub BallBounceTimer_Timer
    If myPrefs_PlayfieldGlass = 0 Then
      Playsound "ball_bounce"
    Else
      Playsound "ball_bounceG"
    End If
    BallBounceTimer.Enabled = 0
  End Sub


'**********************************************************************
'************************* G5K Ball Shadows ***************************
'**********************************************************************

Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate()
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 5
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 5
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

Sub L20_Init()

End Sub

Sub L20_Timer()

End Sub

Dim DT3bGlass, DT4LGlass, DT4RGlass

  Sub  GraphicsTimer1_Timer() 'UpperFlipperPrimitive Motion
      batleft1.objrotz = FlipperLeftUpper.CurrentAngle - 8
      batright1.objrotz = FlipperRightUpper.CurrentAngle + 18
      batleft.objrotz = FlipperLeft.CurrentAngle + 1
      batright.objrotz = FlipperRight.CurrentAngle - 1
    If myPrefs_PlayfieldGlass = 1 Then
      GlassImpurities.Visible=1
    Else
      GlassImpurities.Visible=0
    End If

    if DT3bGlass = 1 Then
      if myPrefs_PlayfieldGlass = 1 Then
      dt3B.InitSnd "fx_Ldroptarget", "fx_Lresetdrop"
      else
      dt3B.InitSnd "fx_LdroptargetNG", "fx_LresetdropNG"
      end If
    end If

    if DT4LGlass = 1 Then
      if myPrefs_PlayfieldGlass = 1 Then
        dt4L.InitSnd "fx_droptarget", "fx_resetdrop"
      else
        dt4L.InitSnd "fx_droptargetNG", "fx_resetdropNG"
      end If
    end if

    if DT4RGlass = 1 Then
      if myPrefs_PlayfieldGlass = 1 Then
        dt4R.InitSnd "fx_droptarget", "fx_resetdrop"
      else
        dt4R.InitSnd "fx_droptargetNG", "fx_resetdropNG"
      end If
    end if

  End Sub
'********************************************************


' Thalamus, table used a very old routine to figure out if ball is colliding.
' This is not needed as vp now has a callback for this.
