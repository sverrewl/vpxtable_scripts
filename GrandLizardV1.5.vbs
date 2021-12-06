
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
'                                         Grand Lizard  V1.5
'
'                                 A 3rdaxis & Slydog43 Collaboration

'
'-Special Thanks
'-Thanks to JP Salas for his VP9 version
'-Thanks to Rosve for his VP9 version
'-Thanks to G5K for the flippers
'-Thanks to VP Dev Team for making VPX so awesome!!!
'
'///////////////////////////////////////////////
'//Layers                                     //
'//1=Primitive, Walls, Rubbers, Lamps, Timers //
'//2=Lamps B                                  //
'//3=GI                                       //
'//4=Flashers                                 //
'//5=Silkscreened Plastics                    //
'//6=Clear Plastics, Hal 9000                 //
'//7=Flashers Large                           //
'//8=Drop Targets                             //
'//9=Outlane Posts                            //
'//10=Outlane Posts                           //
'//11=Outlane Posts                           //
'///////////////////////////////////////////////


Option Explicit
Randomize

' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!
' I added JP ballrolling, std ssf routines and changed the placement of sounds

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
Const VolFlip   = .8   ' Flipper volume.


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
Dim myImage, myObject, myMaterial, myLightColor, myLightColorFull, myLightIntensity
Dim myFlipperLeftMatterial, myFlipperLeftRubberMaterial, myFlipperLeftUpperMatterial, myFlipperLeftUpperRubberMaterial
Dim myFlipperRightMatterial, myFlipperRightRubberMaterial, myFlipperRightUpperMatterial, myFlipperRghtUpperRubberMaterial
Dim myPrefs_BallRollingSound, myPrefs_PlayMusic, myPrefs_ShowJungleBlades, myPrefs_HideCenterPost, myPrefs_GIColor, myPrefs_RubberColor
Dim myPrefs_RubberSleevesColor, myPrefs_DynamicBackground
Dim myPrefs_Drain, myPrefs_SoundEffects, myPrefs_Disablecheating, myPrefs_AutoPlay, myPrefs_Soundonoff, myPrefs_FlipperBatType, myPrefs_InstructionCardType
Dim myPrefs_OutlaneDifficulty, myPlayingEndGameSnd , myPrefs_AutoPlayToggleKeyCode, myPrefs_AutoPlaySoundToggleKeyCode, myPrefs_AutoPlayCheatingToggleKeyCode
Dim myGamePlaying, AnimateCard1Counter

myGamePlaying = False
myPlayingEndGameSnd = False
myPrefs_Drain = 0
myPrefs_SoundEffects = 1

myPrefs_AutoPlayToggleKeyCode = 30                'a
myPrefs_AutoPlaySoundToggleKeyCode = 31       's
myPrefs_AutoPlayCheatingToggleKeyCode = 46    'c

' _________   ___  __   ____  ____  ___  ______________  _  ______
'/_  __/ _ | / _ )/ /  / __/ / __ \/ _ \/_  __/  _/ __ \/ |/ / __/
' / / / __ |/ _  / /__/ _/  / /_/ / ___/ / / _/ // /_/ /    /\ \ 
'/_/ /_/ |_/____/____/___/  \____/_/    /_/ /___/\____/_/|_/___/
'***********************************************************************************************************************************


myPrefs_RubberColor = 0                     '0=White(Default) 1=Random 2=Black 3=Green 4=Red 5=Blue
myPrefs_RubberSleevesColor = 0       '0=Black 1=Random 2=Yellow(Default) 3=Green 4=Red 5=Blue
myPrefs_GIColor = 0                             '0=Warm White(Default) 1=Random 2=Cool White 3=Green 4=Blue 5=Purple 6=Red
myPrefs_ShowJungleBlades = 0           '0=Off(Default) 1=On
myPrefs_InstructionCardType = 0     '0=Factory(Default) 1=Modded Cards 2=THX Card  (See THX Setup for directions)
myPrefs_OutlaneDifficulty = 0         '0=Factory(Default) 1=Easy 2=Hard
myPrefs_HideCenterPost = 0               '0=Factory 1=Hide Post
myPrefs_PlayMusic = 1                         '0=Off(Default) 1=On
myPrefs_BallRollingSound = 1           '0=Off(Default) 1=On
myPrefs_DynamicBackground = 1         '0=Off 1=On
myPrefs_Disablecheating = 0             '0=Off 1=On (If set to 0 and cheating is enabled during Autoplay scores will not be saved)
myPrefs_FlipperBatType = 0               'Flipper Bat Type (G5K)

'0= White Bat Red Rubber Upper & Lower(Default)
'1= White Bat Black Rubber Upper & Lower
'2= White Bat Green Rubber Upper & Lower
'3= White Bat Red Rubber Upper & Blue Rubber Lower

'4= Yellow Bat Red Rubber Upper & Lower
'5= Yellow Bat Black Rubber Upper & Lower
'6= Yellow Bat Green Rubber Upper & Lower
'7= Yellow Bat Green Rubber Upper & Red Rubber Lower

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
'                         )  (   (                )
'   (            *   ) ( /(  )\ ))\ )   (      ( /(
'   )\       ( ` )  /( )\())(()/(()/(   )\     )\())
'((((_)(     )\ ( )(_)|(_)\  /(_))(_)|(((_)(  ((_)\ 
' )\ _ )\ _ ((_|_(_())  ((_)(_))(_))  )\ _ )\__ ((_)
' (_)_\(_) | | |_   _| / _ \| _ \ |   (_)_\(_) \ / /
'  / _ \ | |_| | | |  | (_) |  _/ |__  / _ \  \ V /
' /_/ \_\ \___/  |_|   \___/|_| |____|/_/ \_\  |_|

'Instructions:
'* Select the number of players desired then use the "A" key to engage and toggle AutoPlay on and off.
'* Use the "S" key to toggle on and off Autoplay's sounds and special effects during Autoplay.
'* Use the "C" key to toggle on and off cheating during Autoplay (Game will play indefinitely no scoring)
'* Status Lights:(Just Above Apron To The Left) Yellow=(Autoplay On) Blue=(Sound On) Green=(Cheat On)

'<<<<<***************  WARNING!!!ONCE "CHEAT" IS ENABLED SCORES WILL NOT BE SAVED! WARNING!!!********************>>>>>

'Also, apron credit light will flash and plunger will shoot when AutoPlay is activated.
'A new game or aditional players added will deactivate AutoPlay.

'***********************************************************************************************************************************

'        ______   _______  ___  _______  ____
'*****  / __/ /  /  _/ _ \/ _ \/ __/ _ \/ __/*****
'***** / _// /___/ // ___/ ___/ _// , _/\ \  ***** THANK YOU G5K
'*****/_/ /____/___/_/  /_/  /___/_/|_/___/  *****

   Sub  GraphicsTimer1_Timer() 'UpperFlipperPrimitive Motion
    batleft1.objrotz = FlipperLeftUpper.CurrentAngle - 8
    batright1.objrotz = FlipperRightUpper.CurrentAngle + 18
   End Sub

   Sub GraphicsTimer_Timer() 'Lower Flipper Primitive Motion
    batleft.objrotz = FlipperLeft.CurrentAngle + 1
    batright.objrotz = FlipperRight.CurrentAngle - 1
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

Sub myAdjustTableToPrefs
  If myPrefs_PlayMusic then
    PlaySound "GrandLizardIntro"
    Playsound "Lion-Roar"
   End If

  If myPrefs_ShowJungleBlades Then
    BladeR.Visible = True
    BladeL.Visible = True
  End If

  If ShowDT then
    SideRails.Visible = True
    LizardHeadDT.Visible = True
    LizardHeadFS.Visible = False
    LockdownBar. Visible = True
  End If

  If myPrefs_HideCenterPost Then
    Primitive43.Visible = False
    CPostRubber.Visible = False
    CPostRubber.Collidable = False
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
' ****** Set Rubber Sleeves Color

  Select Case myPrefs_RubberSleevesColor
    Case 0 'Black
      myMaterial = "Sleeve Black"
      myImage = "RubberSleeve Black"
    Case 1 'Random
       Select Case Int(Rnd()*4)
         Case 0 'Black
           myMaterial = "Sleeve Black"
           myImage = "RubberSleeve Black"
         Case 2 'Yellow
           myMaterial = "Sleeve Yellow"
           myImage = "RubberSleeve Yellow"
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
    Case 2 'Yellow
      myMaterial = "Sleeve Yellow"
      myImage = "RubberSleeve Yellow"
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

If NOT myGIOn then
  myMaterial = myMaterial & "D" 'for dark
End if

  For each myObject in RubberSleeves
      myObject.Material = myMaterial
      myObject.Image = myImage
 Next
' ****** End Set Rubber Sleeves Color
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
    bsTrough.InitExitSnd "SSNewBall", "Solenoid"
    bsTrough.Balls=3

    ' 3 Bank
    Set dt3B=New cvpmDropTarget
    dt3B.InitDrop Array(sw25, sw26, sw27), Array(25, 26, 27)
    dt3B.InitSnd "Droptarget", "resetdrop"

    ' 4 Bank Left
    Set dt4L=New cvpmDropTarget
    dt4L.InitDrop Array(sw20, sw21), Array(20, 21)
    dt4L.InitSnd "Droptarget", "resetdrop"

    ' 4 Bank Right
    Set dt4R=New cvpmDropTarget
    dt4R.InitDrop Array(sw22, sw23), Array(22, 23)
    dt4R.InitSnd "Droptarget", "resetdrop"

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

  If NOT Me.ShowDT  Then
    TextBoxWarning.Visible = False
    End If



 End Sub 'Table Init

 Sub Table1_Paused:Controller.Pause=1:End Sub
 Sub Table1_unPaused:Controller.Pause=0:End Sub
 Sub Table1_Exit():Controller.Stop:End Sub


'*************************************************************************************
'*****************               KEY EVENTS                ***************************
'*************************************************************************************
 Sub Table1_KeyDown(ByVal keycode)
    If keycode=RightFlipperKey Then Controller.Switch(48)=1
    If keycode=LeftFlipperKey Then Controller.Switch(47)=1
  If keycode= RightMagnaSave Then Controller.Switch(46)=1
    If keycode= LeftMagnaSave Then Controller.Switch(45)=1
    If keycode = LeftTiltKey Then LeftNudge 80, .8, 20:PlaySound "nudge_left"
    If keycode = RightTiltKey Then RightNudge 280, .8, 20:PlaySound "nudge_right"
    If keycode = CenterTiltKey Then CenterNudge 0, 1.2, 25:PlaySound "nudge_forward"
  If keycode = PlungerKey Then
     Plunger1.PullBack
         PlaySoundAtVol "fx_plungerpull", Plunger1, 1
  End If
  If keycode = StartGameKey Then
    Call myStartGame
  End If

  'sds
    ' If keycode = myPrefs_AutoPlayToggleKeyCode and myBallCount > 0 then
    If keycode = myPrefs_AutoPlayToggleKeyCode then
    myPrefs_AutoPlay = myPrefs_AutoPlay + 1
    If myPrefs_AutoPlay > 1 Then
      myPrefs_AutoPlay = 0
    End If
    Call myAutoPlay(myPrefs_AutoPlay)
  End If

  If keycode = myPrefs_AutoPlaySoundToggleKeyCode then
    call myAutoPlaySoundToggle
  End If

   If keycode = myPrefs_AutoPlayCheatingToggleKeyCode then
    call myAutoPlayCheatToggle
  End If


    If vpmKeyDown(keycode) Then Exit Sub
 End Sub

 Sub Table1_KeyUp(ByVal keycode)
  If keycode=RightFlipperKey Then Controller.Switch(48)=0
  If keycode=LeftFlipperKey Then Controller.Switch(47)=0
  If keycode=  RightMagnaSave Then Controller.Switch(46)=0
  If keycode= LeftMagnaSave Then Controller.Switch(45)=0
  If keycode = PlungerKey Then
    Plunger1.Fire
        PlaySoundAtVol "fx_plunger", Plunger1, 1
  End If

  If keycode=keyInsertCoin1 or keycode=keyInsertCoin2 or keycode=keyInsertCoin3 or keycode=keyInsertCoin4 Then
    TextBoxWarning.Visible = False
  End If

  If vpmKeyUp(keycode) Then Exit Sub
 End Sub


'===============================================
Sub myStartGame
    GameOverTimer.Enabled = 0
    Gameover6.Collidable = 0
    StopSound ""
    StopSound ""
    CreditLight.State = 0
    CreditLight1.State = 0
    Hal9000Eye.Visible = 0
    Hal9000Ring.Visible = 0
    HalLight2.State = 0
    TWall.enabled = 0
    TWall1.enabled = 0
    TriggerLF.enabled=0
    TriggerRF.enabled=0
    TriggerULF.enabled=0
    TriggerURF.enabled=0
    TriggerLMS.enabled=0
    TriggerRMS.enabled=0
    TriggerAutoFlipPlunger.enabled = 0
    stopsound "AutoplayTrack1"
    stopsound "AutoplayTrack2"
         stopsound "AutoplayTrack3"
    AutoplayTrack1Timer.enabled=0
    AutoplayTrack2Timer.enabled=0
         AutoplayTrack3Timer.enabled=0
  myGamePlaying = True
End Sub

Sub myAutoPlaySoundToggle
    If Hal9000Eye.Visible = True then
      If myPrefs_SoundEffects = 1 Then
        myPrefs_SoundEffects = 0
        Hal9000Eye.Image = "__BlueEye"
        Playsound "button-click"
        Playsound "spin-down"
        Stopsound "boot1"
        stopsound "AutoplayTrack1"
        stopsound "AutoplayTrack2"
                stopsound "AutoplayTrack3"
        AutoplayTrack1Timer.enabled=0
        AutoplayTrack2Timer.enabled=0
                                AutoplayTrack3Timer.enabled=0
                                liB.State=0
       Else
        myPrefs_SoundEffects = 1
        Hal9000Eye.Image = "__RedEye"
        If myPrefs_Drain = 1 Then
            Hal9000Eye.Image = "__GreenEye"
        End If
        Playsound "button-click"
        Playsound "boot1"
        Stopsound "spin-down"
                                liB.State=1
        Select Case Int(Rnd()*3)
          Case 0
            playsound "AutoplayTrack1"
                                                AutoplayTrack1Timer.enabled=1
          Case 1
            playsound "AutoplayTrack2"
                                                AutoplayTrack2Timer.enabled=1
                                         Case 2
            playsound "AutoplayTrack3"
                                                AutoplayTrack3Timer.enabled=1
        End Select
      End If
  End If
End Sub

Sub myAutoPlayCheatToggle
  If Hal9000Eye.Visible = True then
    If myPrefs_Drain = 1 Then
      myPrefs_Drain = 0
      Wall11.Collidable = 0
      Wall12.Collidable = 0
      Wall13.Collidable = 0
                        BarrierLeft.Collidable = 0
                        BarrierRight.Collidable = 0
      Playsound "button-click"
      Stopsound "cheater"
      StrobeLight2.State=0
                        liG.State=0
       If myPrefs_SoundEffects = 1 Then
        Hal9000Eye.image = "__RedEye"
      End If
       If myPrefs_SoundEffects = 0 Then
        Hal9000Eye.image = "__BlueEye"
       End If
     Else
      myPrefs_Drain = 1
      Wall11.Collidable = 1
      Wall12.Collidable = 1
      Wall13.Collidable = 1
                        BarrierLeft.Collidable = 1
                        BarrierRight.Collidable = 1
      If myPrefs_Disablecheating = 0 Then
         Playsound "button-click"
         Playsound "cheater"
         Playsound "cheater"
         StrobeLight2.State=2
         Hal9000Eye.image = "__GreenEye"
         Drain.enabled=False
         TWall.enabled=False
         Gameover6.Collidable = 1
                                 liG.State=1
      End If
    End If
       End If
End Sub

'Sub StatusLightsVisible
'           If Hal9000Ring.Visible=0 Then
'                liY.Visible=0
'                liB.Visible=0
'                liG.Visible=0
'       If Hal9000Ring.Visible=1 Then
'                liY.Visible=1
'                liB.Visible=1
'                liG.Visible=1
'        End If
'     End If
'End Sub

' Function myBallCount
'   Dim iCounter
'   Dim myCount
'
'   iCounter = 0
'   For iCounter = 1 to 4
'     myCount = myCount + ballStatus(iCounter)
'   Next
'     myBallCount = myCount
' End Function
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
    'PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
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
     'PlaySound SoundFX("left_slingshot",DOFContactors), 0, 1, 0.05, 0.05
  vpmTimer.PulseSw 43
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
     'Msgbox "left Hit"
     If Activeball.velX > 10 Then
     'Msgbox Activeball.velx
     SlingHopLTimer.Enabled = 1
     Activeball.velZ = 10
   End If
 End Sub

 Sub SlingHopLTimer_Timer
     Playsound "ball_bounce"
     SlingHopLTimer.Enabled = 0
   End Sub

 Sub SlingHopR_Hit
     'Msgbox "Right hit"
     If Activeball.velX > 10 Then
     'Msgbox Activeball.velx
     SlingHopRTimer.Enabled = 1
     Activeball.velZ = 10
   End If
 End Sub

 Sub SlingHopRTimer_Timer
     Playsound "ball_bounce"
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
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSwitch mySwitchNum, 0, ""
    myTarget1.playanim 0,.25
    myTarget2.playanim 0,.25
End Sub

'****************************************************************************
'*****************               FLIPPERS                 *******************
'****************************************************************************

 Sub LFlipper(Enabled)
   If Enabled Then
    PlaySoundAtVol "flipperup", FlipperLeftUpper, VolFlip
    PlaySoundAtVol "flipperup", FlipperLeft, VolFlip
    FlipperLeft.RotateToEnd
    FlipperLeftUpper.RotateToEnd
                'myLightIntensity = 0 'SDS
                Call myChangeGIForFlip()
        'LightIntensityTimer.Enabled = 1
   Else
    PlaySoundAtVol "flipperdown", FlipperLeftUpper, VolFlip
    PlaySoundAtVol "flipperdown", FlipperLeft, VolFlip
    FlipperLeft.RotateToStart
    FlipperLeftUpper.RotateToStart
   End If
 End Sub

 Sub RFlipper(Enabled)
   If Enabled Then
    PlaySoundAtVol "flipperup", FlipperRightUpper, VolFlip
    PlaySoundAtVol "flipperup", FlipperRight, VolFlip
    FlipperRight.RotateToEnd
    FlipperRightUpper.RotateToEnd
                'myLightIntensity = 0
                'LightIntensityTimer.Enabled = 1
                Call myChangeGIForFlip()
   Else
    PlaySoundAtVol "flipperdown", FlipperRightUpper, VolFlip
    PlaySoundAtVol "flipperdown", FlipperRight, VolFlip
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
     PlaySoundAtBallVol "rubber_flipper", 1
 End Sub

 Sub FlipperRight_Collide(parm)
     PlaySoundAtBallVol "rubber_flipper", 1
 End Sub

 Sub FlipperUpperLeft_Collide(parm)
     PlaySoundAtBallVol "rubber_flipper", 1
 End Sub

 Sub FlipperUpperRight_Collide(parm)
     PlaySoundAtBallVol "rubber_flipper", 1
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
    LizardHeadDT.image="lizheadL2AXS"
    LizardHeadFS.image="lizheadL2AXS"
    Flasher1.visible = true
    Flasher66b.state = true
Select Case Int(Rnd()*3)
 Case 0
       Flasher66c.visible = true
 Case 1
       Flasher66c1.visible = true
 Case 2
       Flasher66c2.visible = true
' Case 3
'       Flasher66c3.visible = true
End Select
  Else
    LizardHeadDT.image="lizheadAXS"
    LizardHeadFS.image="lizheadAXS"
    Flasher1.visible = false
    Flasher66b.state = False
    Flasher66c.visible = False
    Flasher66c1.visible = False
    Flasher66c2.visible = False
    Flasher66c3.visible = False
  End If
End Sub

 Sub SetFlasher67(enabled)'3 Bank
  If Enabled Then
    LizardHeadDT.image="lizhead67AXS"
    LizardHeadFS.image="lizhead67AXS"
    Flasher2.visible = true
    Flasher67.visible = true
    Flasher67b.state = true
  Else
    LizardHeadDT.image="lizheadAXS"
    LizardHeadFS.image="lizheadAXS"
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
    Flasher4.visible = True
    Flasher69.visible = true
    Flasher69b.state = true
  Else
    Flasher4.visible = False
    Flasher69.visible = false
    Flasher69b.state = False
  End If
End Sub

 Sub SetFlasher70(enabled)'Rmag
  If Enabled Then
    Flasher5.visible = True
    Flasher70.visible = true
    Flasher70b.state = true
  Else
    Flasher5.visible = False
    Flasher70.visible = false
    Flasher70b.state = False
  End If
End Sub

Sub LockKick(enabled)
    if enabled then
       sw34.Kick 100, 15
       PlaySoundAtVol "fpopper", sw34, VolKick
    end if
 End Sub

'*********************************************
'**********  GENERAL ILLUMINATION  ***********
'*********************************************

Sub PFGI(Enabled)
  If Enabled Then
            TextBoxWarning.Visible = False
      CabinetWalls.image = "LargeWood2"
      GI_13.State = 0
      GI_2.State = 0
      GI_14.State = 0
      Light101.Intensity = 1
      Flasher6.visible = False
      LizardHeadDT.image="lizheadDAXS"
      LizardHeadFS.image="lizheadDAXS"
      BladeR.image="1373633971-sota-forestD"
      BladeL.image="1373633971-sota-forestD"
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
                        Gates.Material="Metal2"
                        MetalWork.Material="Metal2"
                        BounceBackBars.Material="Metal2"
                        PlasticNuts.Material = "PlasticNutsD"
                        PostOutlaneRight1.Material = "PostRedD"
                        PostOutlaneRight2.Material = "PostRedD"
                        PostOutlaneRight3.Material = "PostRedD"
                        PostOutlaneLeft1.Material = "PostRedD"
                        PostOutlaneLeft2.Material = "PostRedD"
                        PostOutlaneLeft3.Material = "PostRedD"
                        ApronLight1.Intensity = 0
                        ApronLight2.Intensity = 0
                        sw20.Material = "PlasticD"
      sw21.Material = "PlasticD"
      sw22.Material = "PlasticD"
      sw23.Material = "PlasticD"
      sw25.Material = "PlasticD"
      sw26.Material = "PlasticD"
      sw27.Material = "PlasticD"
                        RampBlades.Material= "RampBladesD"
                        PlasticSlingshots.Material= "PlasticTriggerD"
      Plastic2.Material= "PlasticTriggerD"
      Plastic3.Material= "PlasticTriggerD"
      Plastic5.Material= "PlasticTriggerD"
      Plastic6.Material= "PlasticTriggerD"
      Plastic7.Material= "PlasticTriggerD"
      Plastic8.Material= "PlasticTriggerD"
      Plastic9.Material= "PlasticTriggerD"
                        'iBall.Image= "_SuperBallD"
                       'wireform.Material="Metal21" (Comment out for "Static" status)
    dim xx
  For each xx in GI:xx.State = 0: Next
                 PlaySound "fx_relay"
                 Call mySetRubberColor(False)
           Call mySetRubberSleeveColor(False)
           Call mySetPostColor(False)
  Else
      CabinetWalls.image = "LargeWood"
      Light101.Intensity = 6
      GI_13.State = 1
      GI_2.State = 1
      GI_14.State = 1
      Flasher6.visible = True
      LizardHeadDT.image="lizheadAXS"
      LizardHeadFS.image="lizheadAXS"
      BladeR.image="1373633971-sota-forest"
      BladeL.image="1373633971-sota-forest"
      batleft.Material="g_Plastic"
      batright.Material="g_Plastic"
      batleft1.Material="g_Plastic"
      batright1.Material="g_Plastic"
      Ramp1.visible = False
      Ramp2.visible = False
      SideRails.Material="Metal2"
      LockdownBar.Material="Metal2"
      Gates.Material="Metal1"
      MetalWork.Material="Metal1"
                        BounceBackBars.Material="Metal1"
      PlasticNuts.Material = "PlasticNuts"
                        PostOutlaneRight1.Material = "PostRed"
                        PostOutlaneRight2.Material = "PostRed"
                        PostOutlaneRight3.Material = "PostRed"
                        PostOutlaneLeft1.Material = "PostRed"
                        PostOutlaneLeft2.Material = "PostRed"
                        PostOutlaneLeft3.Material = "PostRed"
                        ApronLight1.Intensity = 2
                        ApronLight2.Intensity = 2
                        sw20.Material = "Plastic"
      sw21.Material = "Plastic"
      sw22.Material = "Plastic"
      sw23.Material = "Plastic"
      sw25.Material = "Plastic"
      sw26.Material = "Plastic"
      sw27.Material = "Plastic"
                        RampBlades.Material= "RampBlades"
      PlasticSlingshots.Material= "PlasticTrigger"
      Plastic2.Material= "PlasticTrigger"
      Plastic3.Material= "PlasticTrigger"
      Plastic5.Material= "PlasticTrigger"
      Plastic6.Material= "PlasticTrigger"
      Plastic7.Material= "PlasticTrigger"
      Plastic8.Material= "PlasticTrigger"
      Plastic9.Material= "PlasticTrigger"
                        'iBall.Image= "_SuperBall"
      'wireform.Material="Metal2" (Comment out for "Static" status)
    For each xx in GI:xx.State = 1: Next
      PlaySound "fx_relay"
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
 Sub sw28_Spin():vpmTimer.PulseSwitch 28, 0, "":PlaySoundAtVol "spinner", sw28, VolSpin:End Sub
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
 ' Sub Drain_Hit():Playsound "drain":ClearBallID:bsTrough.AddBall Me::End Sub
 Sub Drain_Hit():PlaysoundAtVol "drain", Drain, 1:bsTrough.AddBall Me::End Sub
 Sub StackingOpto_Hit()
  stopsound("JungleJam2") 'Ending Song
  stopsound("ConanEnding")
  myPlayingEndGameSnd = False
  ' NewBallid
End Sub


'*************************************************************************************
'******************              SOUNDS                   ****************************
'*************************************************************************************

 Sub RubberSleeves_Hit(index)
    If (Abs(int(Activeball.velx)) or Abs(int(Activeball.vely))) > 5 then
       PlaySoundAtVol "rubberS", ActiveBall, 1
    Else
       PlaySoundAtVol "rubberQ", ActiveBall, 1
    End If
 End Sub

 Sub Rubbers_Hit(index)
    If (Abs(int(Activeball.velx)) or Abs(int(Activeball.vely))) > 5 then
       PlaySoundAtVol "rubberS", ActiveBall, 1
    Else
       PlaySoundAtVol "rubberQ", ActiveBall, 1
    End If
 End Sub

 Sub M03_Hit
     PlaySoundAtVol "FX_metalhit2", ActiveBall, 1
 End Sub

 Sub M02_Hit
     PlaySoundAtVol "FX_metalhit2", ActiveBall, 1
 End Sub

 Sub Trigger1_Hit
    Playsound "Lion-Roar"
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


' fix
' Dim iCounter
' For iCounter = 7 to 64
'   Eval("Light" & iCounter & "b.State") = Eval("Light" & iCounter & ".State")
' Next
'***************************************************
'**********   MUSIC,SOUND AND MOTION    ************
'***************************************************

    If Light2.State = 1 then 'End Game
    If  NOT myPlayingEndGameSnd Then
      If myPrefs_PlayMusic then
        If myPrefs_AutoPlay Then
          PlaySound "ConanEnding"
          stopsound "JungleJam2"
        Else
          stopsound "AutoplayTrack1"
          stopsound "AutoplayTrack2"
                                        stopsound "AutoplayTrack3"
          PlaySound "JungleJam2"
          Playsound "GrandLizardIntro2"
          Playsound "MonsterGrowl"
          AutoplayTrack1Timer.enabled=0
          AutoplayTrack2Timer.enabled=0
                                        AutoplayTrack3Timer.enabled=0
        End IF
        myPlayingEndGameSnd = True
      End If
    End If
    If  myGamePlaying then 'Turn Off Autoplay
      myPrefs_AutoPlay = 0
      Call myAutoPlay(myPrefs_AutoPlay)
      myGamePlaying = False
    End If
  End If

End Sub

'==================================================================

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    WireGateR.rotx= -90 - Gate6.Currentangle / 1
    WireGateL.rotx= -90 - Gate5.Currentangle / 1
    PTounge.rotx = -90 - Gate1.Currentangle / 4.5
    Spinner.rotx = -90 - sw28.Currentangle/1
    Spinnerarm.rotx = -90 - sw28.Currentangle/1
    Gateflap.objRotx  = -10 - Gate3.Currentangle / 1
   End Sub

Set MotorCallback = GetRef("RealTimeUpdates")

'===================================================================

Sub Wall6_hit:PlaysoundAtVol "rubberS", ActiveBall, 1
    'Msgbox Activeball.velY
    If Activeball.velY > 4 Then
    Activeball.velZ = 12
    Wall6Timer.Enabled = 1
   End if
     End sub

Sub Wall6Timer_Timer
    Playsound "ball_bounce"
    Wall6Timer.Enabled = 0
   End Sub

Sub Wall7_hit:PlaysoundAtVol "rubberS", ActiveBall, 1
    If Activeball.velY > 3 Then
    Activeball.velZ = 12
    Wall7Timer.Enabled = 1
   End if
     End sub

Sub Wall7Timer_Timer
    Playsound "ball_bounce"
    Wall7Timer.Enabled = 0
   End Sub

Sub Wall8_hit:PlaysoundAtVol "rubberS", ActiveBall, 1
    If Activeball.velY > 3 Then
    Activeball.velZ = 12
    Wall8Timer.Enabled = 1
   End if
     End sub

Sub Wall8Timer_Timer
    Playsound "ball_bounce"
    Wall8Timer.Enabled = 0
   End Sub

Sub Wall9_hit:PlaysoundAtVol "rubberS", ActiveBall, 1
    If Activeball.velY > 4 Then
    Activeball.velZ = 12
    Wall9Timer.Enabled = 1
   End if
     End sub

Sub Wall9Timer_Timer
    Playsound "ball_bounce"
    Wall9Timer.Enabled = 0
   End Sub

'===================================================================

''*************************************
''*** Accompanying Sounds & Scripts *** AXS
''*************************************
'
Sub   AutoplayTrack1Timer_Timer
         Playsound "AutoplayTrack2"
         Stopsound "AutoplayTrack1"
         AutoplayTrack1Timer.enabled=0
         AutoplayTrack2Timer.enabled=1
End Sub

Sub   AutoplayTrack2Timer_Timer
         Playsound "AutoplayTrack3"
         Stopsound "AutoplayTrack2"
         AutoplayTrack2Timer.enabled=0
         AutoplayTrack3Timer.enabled=1
End Sub

Sub   AutoplayTrack3Timer_Timer
         Playsound "AutoplayTrack1"
         Stopsound "AutoplayTrack3"
         AutoplayTrack3Timer.enabled=0
         AutoplayTrack1Timer.enabled=1
End Sub

Sub TWall_hit
  If myPrefs_SoundEffects Then
     Select Case Int(Rnd()*5)
    Case 0
      Call myAutoPlayEndBallSound("WizardScream1")
    Case 1
      Call myAutoPlayEndBallSound("Lizard-scream1")
    Case 2
      Call myAutoPlayEndBallSound("Glass Breakin")
    Case 3
      Call myAutoPlayEndBallSound("Fatality")
    Case 4
      Call myAutoPlayEndBallSound("WizardScream3")
    End Select
  End If
End Sub

Sub myAutoPlayEndBallSound(mySoundFile)
  Playsound mySoundFile
    Playsound mySoundFile
    Playsound mySoundFile
    CreditLight1.State = 2
    CreditLight.State = 2
End Sub

 Sub TWall1_hit
    If myPrefs_SoundEffects Then
     Select Case Int(Rnd()*6)
       Case 0
      Call myAutoPlaySound("Fight")
       Case 1
      Call myAutoPlaySound("EnoughTalk")
       Case 3
      Call myAutoPlaySound("Revenge")
       Case 4
      Call myAutoPlaySound("Notyet")
       Case 5
      Playsound ""
      End Select
    End If
 End Sub

Sub myAutoPlaySound(mySoundFile)
  Playsound mySoundFile
    Playsound mySoundFile
    Playsound mySoundFile
End Sub
'====================================================================================================================
Sub myAutoPlay(myOnOff)
    debug.print "OnOff=" & myOnOff
    Select Case myOnOff
      Case 0 'Turn Off AutoPlay
          Call mySetInstructionCards(myPrefs_InstructionCardType)
          RCardFilm1.visible = False
          RCardFilm2.visible = False
          TimerAnimateCard1.Enabled = False
          Select Case Int(Rnd()*3)
             Case 0
                playsound ""
                playsound ""
              Case 1
                playsound ""
              Case 2
                playsound ""
                playsound ""
          End Select
          playsound "button-click"
          stopsound "AutoplayTrack1"
                stopsound "AutoplayTrack2"
          stopsound "AutoplayTrack3"
          Wall11.Collidable = 0
          Wall12.Collidable = 0
          Wall13.Collidable = 0
                                        BarrierLeft.Collidable = 0
                                        BarrierRight.Collidable = 0
          StrobeLight2.State=0
          CreditLight.State = 0
          CreditLight1.State = 0
          HalLight2.State = 0
          Hal9000Eye.Visible = 0
          Hal9000Ring.Visible = 0
          TWall.enabled = 0
          TWall1.enabled = 0
          TriggerLF.enabled=0
          TriggerRF.enabled=0
          TriggerULF.enabled=0
          TriggerURF.enabled=0
                    TriggerLMS.enabled=0
                    TriggerRMS.enabled=0
          TriggerAutoFlipPlunger.enabled = 0
                    AutoplayTrack1Timer.enabled=0
                    AutoplayTrack2Timer.enabled=0
                    AutoplayTrack3Timer.enabled=0
                    liY.State=0
                    liB.State=0
                    liG.State=0
      Case 1 'Turn On AutoPlay
        RCardFilm1.visible = True
        RCardFilm2.visible = True
        Call mySetInstructionCards(-1)

        TimerAnimateCard1.Enabled = True
        AnimateCard1Counter = 0

        Select Case Int(Rnd()*3)
          Case 0
            playsound "AutoplayTrack1"
                        AutoplayTrack1Timer.enabled=1
          Case 1
            playsound "AutoplayTrack2"
                        AutoplayTrack2Timer.enabled=1
          Case 2
            playsound "AutoplayTrack3"
                        AutoplayTrack3Timer.enabled=1
        End Select
        playsound "0CB00k"
        playsound ""
        playsound "fx_plungerpull"
        stopsound ""
        stopsound "button-click"
        stopsound ""
        stopsound ""
        CreditLight.State = 2
        CreditLight1.State = 2
        Hal9000Eye.Visible = 1
        Hal9000Ring.Visible = 1
        HalLight2.State = 2
        TWall.enabled = 1
        TWall1.enabled = 1
        TriggerLF.enabled=1
        TriggerRF.enabled=1
        TriggerULF.enabled=1
        TriggerURF.enabled=1
                TriggerLMS.enabled=1
                TriggerRMS.enabled=1
        TriggerAutoFlipPlunger.enabled = 1
        StrobeLight.State=1
        StrobeLight1.State=1
        StrobeLightTimer.Enabled=1
        Plunger1.PullBack
        TimerAutoFlipPlunger.Enabled = True
        liY.State=1
        liB.State=1
    End Select
End Sub


 Sub TriggerAutoFlipPlunger_Hit()
    CreditLight1.State = 2
    CreditLight.State = 0
  'msgbox("Great")
  Plunger1.Pullback()
    PlaysoundAtVol "fx_plungerpull", Plunger1, 1
  TimerAutoFlipPlunger.Enabled = True
 End Sub
Sub TimerAutoFlipPlunger_Timer()
  Plunger1.Fire()
    PlaysoundAtVol "fx_plunger", Plunger1, 1
  TimerAutoFlipPlunger.Enabled = False
  TimerBallStuck.Enabled = True
 End Sub

Sub TimerAnimateCard1_Timer()
  RCardFilm1.Image = "IC 00" & Eval(Int(Rnd()*9)+1)
    RCardFilm2.Image = "IC 00" & Eval(Int(Rnd()*9)+1)
End Sub

''******************** Flippers ************************

 Sub TriggerLF_Hit
    If myPrefs_SoundEffects Then
    Call myHitSoundFlipper2
    Select Case Int(Rnd()*3)
       Case 0
'        StrobeLight2.State=1
'        StrobeLightTimer.Enabled=1
       Case 1
        StrobeLight3.State=1
                    myLightColor = RGB(100,255,0)
            myLightColorFull = RGB(100,255,0)
              myLightIntensity = 50
        StrobeLightTimer.Enabled=1
       Case 2
        StrobeLight4.State=1
                          myLightColor = RGB(255,5,5)
            myLightColorFull = RGB(255,5,5)
              myLightIntensity = 40
        StrobeLightTimer.Enabled=1

      End Select
     End If
        'Playsound "flipperup"
    FlipperLeft.RotateToEnd
        'TriggerLF.enabled=0
        TimerLF.Enabled=1
 End Sub

 Sub TimerLF_Timer
        'PlaySound "flipperdown"
    FlipperLeft.RotateToStart
        'TriggerLF.enabled=1
        TimerLF.Enabled=0
 End Sub

 Sub TriggerRF_Hit
   If myPrefs_SoundEffects Then
  Call myHitSoundFlipper1
    Select Case Int(Rnd()*3)
       Case 0
'        StrobeLight2.State=1
'        StrobeLightTimer.Enabled=1
       Case 1
        StrobeLight3.State=1
                            myLightColor = RGB(100,255,0)
            myLightColorFull = RGB(100,255,0)
              myLightIntensity = 50
        StrobeLightTimer.Enabled=1
       Case 2
        StrobeLight4.State=1
                                  myLightColor = RGB(255,5,5)
            myLightColorFull = RGB(255,5,5)
              myLightIntensity = 40
        StrobeLightTimer.Enabled=1
      End Select
     End If
        'Playsound "flipperup"
    FlipperRight.RotateToEnd
        'TriggerRF.enabled=0
        TimerRF.Enabled=1
 End Sub

 Sub TimerRF_Timer
        'PlaySound "flipperdown"
    FlipperRight.RotateToStart
        'TriggerRF.enabled=1
        TimerRF.enabled=0
 End Sub



 Sub Gameover6_Hit      'Should also toggle Autoplay "off"  FIX
      Playsound "WizardScream3"
      vpmTimer.PulseSwitch 7, 0, "":
      resettimer.Enabled=1
      Gameover6.Collidable = 0
      Stopsound "AutoplayTrack1"
      Stopsound "AutoplayTrack2"
      Stopsound "AutoplayTrack3"
      liY.State=0
      liB.State=0
      liG.State=0
  myPrefs_AutoPlay = 0
  Call myAutoPlay(myPrefs_AutoPlay)

  End Sub

  Sub resettimer_timer
       Drain.enabled=True
       Wall11.Collidable = 0
       Wall12.Collidable = 0
       Wall13.Collidable = 0
       BarrierLeft.Collidable = 0
       BarrierRight.Collidable = 0
       StrobeLight2.State=0
       Hal9000Eye.image = "__RedEye"
       resettimer.enabled=0
   End Sub

   Sub StrobeLightTimer_Timer
        StrobeLight.State=0
        StrobeLight1.State=0
        StrobeLight3.State=0
        StrobeLight4.State=0
        StrobeLightTimer.Enabled=0
    End Sub

'************************** Upper Flippers ******************************

 Sub TriggerULF_Hit
    If myPrefs_SoundEffects Then
  Call myHitSoundFlipper1
    Select Case Int(Rnd()*3)
       Case 0
'        StrobeLight2.State=1
'        StrobeLightTimer.Enabled=1
       Case 1
        StrobeLight3.State=1
                          myLightColor = RGB(100,255,0)
            myLightColorFull = RGB(100,255,0)
              myLightIntensity = 50
        StrobeLightTimer.Enabled=1
       Case 2
        StrobeLight4.State=1
                                  myLightColor = RGB(255,5,5)
            myLightColorFull = RGB(255,5,5)
              myLightIntensity = 40
        StrobeLightTimer.Enabled=1
      End Select
     End If
        FlipperLeftUpper.RotateToEnd
        TimerULF.Enabled=1
 End Sub

 Sub TimerULF_Timer
  FlipperLeftUpper.RotateToStart
        TimerULF.enabled=0
 End Sub

 Sub TriggerURF_Hit'
   If myPrefs_SoundEffects Then
    Call myHitSoundFlipper1  ' added ???
      Select Case Int(Rnd()*3)
       Case 0
  '        StrobeLight2.State=1
  '        StrobeLightTimer.Enabled=1
       Case 1
      StrobeLight3.State=1
                myLightColor = RGB(100,255,0)
          myLightColorFull = RGB(100,255,0)
          myLightIntensity = 50
      StrobeLightTimer.Enabled=1
       Case 2
      StrobeLight4.State=1
                    myLightColor = RGB(255,5,5)
          myLightColorFull = RGB(255,5,5)
          myLightIntensity = 40
      StrobeLightTimer.Enabled=1
      End Select
     End If
  FlipperRightUpper.RotateToEnd
     TimerURF.Enabled=1
 End Sub

 Sub TimerURF_Timer
  FlipperRightUpper.RotateToStart
        TimerURF.enabled=0
 End Sub

Sub myHitSoundFlipper1
  Select Case Int(Rnd()*13)
       Case 0
        Playsound "sword1"
       Case 1
        PlaySound "sword2"
       Case 2
        Playsound "sword3"
       Case 3
        playsound "sword4"
       Case 4
        Playsound "sword5"
       Case 5
        Playsound "sword6"
       Case 6
        Playsound "sword7"
       Case 7
        Playsound "sword8"
       Case 8
        PlaySound "sword9"
       Case 9
        Playsound "sword10"
       Case 10
        Playsound "Lizard-scream1"
       Case 11
        Playsound "Lizard-scream2"
       Case 12
        Playsound "WizardScream2"
     End Select
End Sub

Sub myHitSoundFlipper2
     Select Case Int(Rnd()*10)
       Case 0
        Playsound "energy-short-sword-1"
       Case 1
        PlaySound "energy-short-sword-2"
       Case 2
        Playsound "energy-short-sword-3"
       Case 3
        Playsound "energy-short-sword-4"
       Case 4
        Playsound "sword1"
       Case 5
        Playsound "sword2"
       Case 6
        Playsound "sword3"
       Case 7
        Playsound "sword4"
       Case 8
        PlaySound "sword5"
       Case 9
        Playsound "sword6"
      End Select
End Sub

'************* Auto Magna Save *************

Sub TriggerLMS_Hit
    If myPrefs_AutoPlay then
     vpmTimer.PulseSwitch 45, 0, "":
   End If
End Sub

Sub TriggerRMS_Hit
    If myPrefs_AutoPlay then
     vpmTimer.PulseSwitch 46, 0, "":
   End If
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

