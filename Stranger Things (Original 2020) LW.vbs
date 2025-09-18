 '   -::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::-
 '    `-/+++//:-.---------.--------.`    ----.  `-----. .--``-::::---  .--------`.-sssssssssss-.
 '   .oss+-.-:oo/o:/sss:/o.osso-:osso   `ssss/   +ssss+`.s::oss:--:+o` .sss+--:+/ +ssss/:+ssss/
 '   +sss-    `--` -sss` .-+ss+ `+ss/   /o+sss`  ++:oss+.s/sss-     -` `sss/../ : /sss+  `osss/
 '   +sssso/:.`    -sss`   +sso/oso.   .s-.sss/  ++ .+ssss/sso   `:///-.oss+:/o`  /sss+ `/sso:`
 '   `+sssssssso:` -sss`   +ss+:sss+` `ooo+ssss. ++  `:sss/sso    -sso``oss:  .` `/ssso+ssso.
 '     `-:+osssss/ -sss`   +ss/ -sss+ :s` `:sss+ ++    .os:sss:   -sso``oss:    .:+sss+-ossso.
 '   --     `/ssso :sss.   +ss+  /sss/s/    +sss-++     -s:.+ss/-.+sso`.sss+..-/o`+sss/ .ossss.
 '   /o-    `/sss-.-:::-` .:::-. -:::::-`  `-:::---.    -:-` `.-::::::.-::::::::- +sss/  .ossso`
 '   +sso+++oso/.   ````````.`                                ``..```      ```   `ossso.  /ssss+`
 '   ````....`      oo/ssso/o/:ooo/``/ooo:`:ooo:`/oooo+`.++.:oss+//os-  `:oo++oo-....... `.......
 '   :////////////-.-  sss: `:.sss:``:sss. -sss. -s+sss/`++/sso.    -: `oss/  `::://////////////:
 '                     sss:   .sss+//+sss. -sss. -o.-ossoo++ss:   .----.osss+/-.`
 '                     sss:   .sss-  -sss. -sss. -o` `+sss++ss-   `oss: .+ossssso.
 '                     sss:   .sss-  -sss. -sss. -o`   -os++ss+    oss:-` ``-/sss+
 '                     sss/   .sss-  -sss. -sss. :o`    .s+`+ss/```oss::+-` `-sss-
 '                   `-+++/.  -+++:``:+++-`:+++:`/+-    .++.`./////+++/-+++++++:.
 '*************************************************************************************************                                                                        ````
 ' Original Pinball Table Created by ScottyWic
 ' Dedicated to my wife, without her consistent disapproval, there's no way I would have finished.
 ' Thanks to Netflix and all who worked on making this unbelievable show.
 ' I hope you enjoy, and please kill the demogorgon won't you?
 ' All Art and resources have been sourced online and we hold no rights to any media used.

' Thanks for those who helped with the Stranger Edition:

'  - My Wife for her amazing voice over work (yes that is her pitch shifted up to sound like a teen)
'  on the game and her patience with me while i was working on this table for far too many hours.
'  - And also for giving us a beautiful baby boy =).
'  - Nailbuster for the amazing pinup player and his endless patience working with me to add features
'  and additions to make all this come together.
'  - The VPX team, without this amazing platform none of this would be possible.
'  - TerryRed for all his work on the DOF and MXleds, the complexity of this work and his dedication is  inspiring.
'  - TomTower for his guidance and the amazing Demogorgon animation - seriously amazing.
'  - Ninuzzu for his help on the Demogorgon extension rod and al blinking wall lights scripting.
'  - DJRobx for the amazing SSF work, making the clunks and thuds sound unbelievably realistic.
'  - JPsalas for his guidance getting me started coding pinball tables he is the godfather.
'  - Peter, Jens and Clint for their endless bug checking and help dialing this game in.end
'  - Flupper for his tutorials on clear plastic ramps and his awesome dome scripts which i used heavily.
'  lauminet for the awesome demogorgon model
'  - Jrsprad & Stat for their help on version 1 of the table.

 '*************************************************************************************************


Option Explicit
Randomize


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  Player Options
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X

  Dim osbactive:osbactive = 0 'set to 0 for off, 1 for only player 1 to be sent, 2 for all scores to be sent.
  Dim osbid:osbid ="" ' your orbital scoreboard login name
  Dim osbkey:osbkey="enter your key" ' your orbital scoreboard api key
  Dim osbdefinit:osbdefinit = "" ' your default initials to use
  Rockmusic = 1 'change this to 1 to switch out the background music with 80s rock/pop songs from the show
  soundtrackvol = 80 'Set the background audio volume to whatever you'd like out of 100
  videovol = 100 'set the volme you'd like for the videos
  calloutvol = 80 ' set this to whatever you're like your callouts to be
  calloutlowermusicvol = 1 'set to 1 if you want music volume lowered during audio callouts
  turnoffrules = 0 ' change to 1 to take off the backglass helper rules text during a game
  turnonultradmd = 1 ' change to 1 to turn on ultradmd, 2 to turn on ultradmd expiramental (only works if pc region set to USA)
  ballrolleron = 0 ' set to 0 to turn off the ball roller if you use the "c" key in your cabinet
  toppervideo = 0 'set to 1 to turn on the topper. This is only for use who have a dedicated Topper screen.



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  FRAMEWORK VARIABLES
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'
  'Constructions
  Const BallSize = 50   'Ball size must be 50
    Const BallMass = 1    'Ball mass must be 1
  Const cGameName = "STLE"
  Const TableName = "STLE"
  Const myVersion = "1.47"
  Const MaxPlayers = 4
  Const BallSaverTime = 15
  Const MaxMultiplier = 10
  Const MaxMultiballs = 5
  Const bpgcurrent = 3

  ' Define Global Variables
  Dim calloutlowermusicvol
  Dim soundtrackvol
  Dim toppervideo
  Dim videovol
  Dim calloutvol
  Dim rockmusic
  Dim ballrolleron
  Dim turnonultradmd
  Dim turnoffrules
  Dim PlayersPlayingGame
  Dim CurrentPlayer
  Dim Credits
  Dim BonusPoints(4)
  Dim BonusHeldPoints(4)
  Dim BonusMultiplier(4)
  Dim bBonusHeld
  Dim BallsRemaining(4)
  Dim ExtraBallsAwards(4)
  Dim Score(4)
  Dim HighScore(4)
  Dim HighScoreName(4)
  Dim WaffleScore(4)
  Dim WaffleScoreName(4)
  Dim Jackpot
  Dim SuperJackpot
  Dim Tilt
  Dim TiltSensitivity
  Dim Tilted
  Dim TotalGamesPlayed
  Dim mBalls2Eject
  Dim SkillshotValue(4)
  Dim bAutoPlunger
  Dim bInstantInfo
  Dim bromconfig
  Dim bAttractMode

  Const typefont = "ITC Avant Garde Gothic LT Bold"
  Const numberfont = "ITC Avant Garde Gothic LT Bold"

  ' Define Game Control Variables
  Dim LastSwitchHit
  Dim BallsOnPlayfield

  Dim BallsInHole

  ' Define Game Flags
  Dim bFreePlay
  Dim bGameInPlay
  Dim bOnTheFirstBall
  Dim bBallInPlungerLane
  Dim bBallSaverActive
  Dim bBallSaverReady
  Dim bMultiBallMode
  Dim bMusicOn
  Dim bSkillshotReady
  Dim bExtraBallWonThisBall
  Dim bJustStarted

  Dim plungerIM

Dim BallHandlingQueue : Set BallHandlingQueue = New vpwQueueManager
Dim xmasQueue : Set xmasQueue = New vpwQueueManager
Dim GeneralPupQueue: Set GeneralPupQueue = New vpwQueueManager

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  Orbital Scoreboard Code
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X

  '****************************
  ' POST SCORES
  '****************************
  Dim osbtemp:osbtemp = osbdefinit
  Dim osbtempscore:osbtempscore = 0
  Sub SubmitOSBScore
    On Error Resume Next
    if osbactive = 1 or osbactive = 2 Then
    Dim objXmlHttpMain, Url, strJSONToSend

    Url = "https://hook.integromat.com/82bu988v9grj31vxjklh2e4s6h97rnu0"

    strJSONToSend = "{""auth"":""" & osbkey &""",""player id"": """ & osbid & """,""player initials"": """ & osbtemp &""",""score"": " & CStr(osbtempscore) & ",""table"":"""& TableName & """,""version"":""" & myVersion & """}"

    Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP")
    objXmlHttpMain.open "PUT",Url, False
    objXmlHttpMain.setRequestHeader "Content-Type", "application/json"
    objXmlHttpMain.setRequestHeader "application", "application/json"

    objXmlHttpMain.send strJSONToSend
    end if
  End Sub



  '****************************
  ' GET SCORES
  '****************************
  dim worldscores

  Sub GetScores()
    if osbkey="" then exit sub
    On Error Resume Next
    Dim objXmlHttpMain, Url, strJSONToSend

    Url = "https://hook.integromat.com/kj765ojs42ac3w4915elqj5b870jrm5c"

    strJSONToSend = "{""auth"":"""& osbkey &""", ""table"":""STLE""}"

    Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP")
    objXmlHttpMain.open "PUT",Url, False
    objXmlHttpMain.setRequestHeader "Content-Type", "application/json"
    objXmlHttpMain.setRequestHeader "application", "application/json"

    objXmlHttpMain.send strJSONToSend

    worldscores = objXmlHttpMain.responseText
    vpmtimer.addtimer 3000, "showsuccess '"
    debug.print "got the scores"
    'debug.print worldscores
    splitscores
  End Sub

  Dim scorevar(22)
  Dim dailyvar(22)
  Dim weeklyvar(22)
  Dim alltimevar(42)
  sub emptyscores
    dim i
    For i = 0 to 42
      alltimevar(i) = "0"
    Next
    For i = 0 to 22
      weeklyvar(i) = "0"
      dailyvar(i) = "0"
    Next
  End Sub
  emptyscores


  Sub splitscores
    On Error Resume Next
    dim a,scoreset,subset,subit,myNum,daily,weekly,alltime,x
    a = Split(worldscores,": {")
    subset = Split(a(1),"[")

'   debug.print subset(1)
'   debug.print subset(2)
'   debug.print subset(3)
' daily scores
    myNum = 0
    daily = Split(subset(1),": ")
    for each x in daily
      myNum = MyNum + 1
      x = Replace(x, vbCr, "")
      x = Replace(x, vbLf, "")
      x = Replace(x, ",", "")
      x = Replace(x, """", "")
      x = Replace(x, "{", "")
      x = Replace(x, "}", "")
      x = Replace(x, "score", "")
      x = Replace(x, "initials", "")
      x = Replace(x, "weekly", "")
      x = Replace(x, "]", "")
      x = Replace(x, "alltime", "")
      dailyvar(MyNum) = x
      if dailyvar(MyNum) = "" Then
        if MyNum = 2 or 4 or 6 or 8 or 10 or 12 or 14 or 16 or 18 or 20 Then
          dailyvar(MyNum) = "OBS"
        Else
          dailyvar(MyNum) = 0
        end if
      end if
      debug.print "dailyvar(" &MyNum & ")=" & x
    Next

' weekly scores
    myNum = 0
    weekly = Split(subset(2),": ")
    for each x in weekly
      myNum = MyNum + 1
      x = Replace(x, vbCr, "")
      x = Replace(x, vbLf, "")
      x = Replace(x, ",", "")
      x = Replace(x, """", "")
      x = Replace(x, "{", "")
      x = Replace(x, "}", "")
      x = Replace(x, "score", "")
      x = Replace(x, "initials", "")
      x = Replace(x, "weekly", "")
      x = Replace(x, "]", "")
      x = Replace(x, "alltime", "")
      weeklyvar(MyNum) = x
      if weeklyvar(MyNum) = "" Then
        if MyNum = 2 or 4 or 6 or 8 or 10 or 12 or 14 or 16 or 18 or 20 Then
          weeklyvar(MyNum) = "OBS"
        Else
          weeklyvar(MyNum) = 0
        end if
      end if
      debug.print "weeklyvar(" &MyNum & ")=" & x
    Next

' alltime scores
    myNum = 0
    alltime = Split(subset(3),": ")
    for each x in alltime
      myNum = MyNum + 1
      x = Replace(x, vbCr, "")
      x = Replace(x, vbLf, "")
      x = Replace(x, ",", "")
      x = Replace(x, """", "")
      x = Replace(x, "{", "")
      x = Replace(x, "}", "")
      x = Replace(x, "score", "")
      x = Replace(x, "initials", "")
      x = Replace(x, "weekly", "")
      x = Replace(x, "]", "")
      x = Replace(x, "alltime", "")
      alltimevar(MyNum) = x
      if alltimevar(MyNum) = "" Then
        if MyNum = 2 or 4 or 6 or 8 or 10 or 12 or 14 or 16 or 18 or 20 or 22 or 24 or 26 or 28 or 30 or 32 or 34 or 36 or 38 or 40 Then
          alltimevar(MyNum) = "OBS"
        Else
          alltimevar(MyNum) = "0"
        end if
      end if
      debug.print "alltimevar(" &MyNum & ")=" & x
    Next

  end Sub


  sub showsuccess
    pNote "Scoreboard","Updated"
    'pupDMDDisplay "-","Scoreboard^Updated",dmdnote,3,0,10
  end sub

  GetScores


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   TABLE INITS & MATHS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X


  Sub Table1_Init()
    SetLocale(1033)
    Spot1.opacity = 0
    resetbackglass
    StartXMAS
    LoadEM
    DMD_Init
    Dim i
    help.opacity = 0
    Randomize
    Const IMPowerSetting = 45 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
      .InitImpulseP swplunger, IMPowerSetting, IMTime
      .Random 1.5
      .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_solenoid", DOFContactors)
      .CreateEvents "plungerIM"
    End With


   Set MagnetR = New cvpmMagnet
    With MagnetR
      .InitMagnet Magnet11, 60
      .GrabCenter = True
      .MagnetOn = 0
      .CreateEvents "MagnetR"
    End With


   Set MagnetU = New cvpmMagnet
    With MagnetU
      .InitMagnet MagnetEscape, 70
      .GrabCenter = True
      .MagnetOn = 0
      .CreateEvents "MagnetU"
    End With

   Set MagnetW = New cvpmMagnet
    With MagnetW
      .InitMagnet Magnetwall, 60
      .GrabCenter = True
      .MagnetOn = 0
      .CreateEvents "MagnetW"
    End With

   Set spinner = New cvpmTurntable
    With spinner
      .InitTurntable TurnTable, 25
      .SpinDown = 10
      .CreateEvents "spinner"
    End With

    Loadhs
    bAttractMode = False
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    bGameInPlay = False
    bAutoPlunger = False
    bMusicOn = True
    BallsOnPlayfield = 0
    BallsInHole = 0
    LastSwitchHit = ""
    Tilt = 0
    Saves = 0
    Drains = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJustStarted = True
    bInstantInfo = False
    bromconfig = False
    'bumps(CurrentPlayer) = 0
    GiOff
    PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":4, ""FS"":1 }"

    StartAttractMode


    ' Remove the cabinet rails if in FS mode
  End Sub
    FlasherEL.opacity = 0
    Flasher1.opacity = 100

  LoadCoreFiles
  Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    'ExecuteGlobal GetTextFile("controller.vbs")
    'If Err Then MsgBox "Can't open controller.vbs"
    On Error Goto 0
  End Sub


Dim SidewallChoice: SidewallChoice = 0
Dim RailChoice: RailChoice = 0
Dim GuitarChoice: GuitarChoice = True
'//////////////F12 Menu//////////////
' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
  If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

    RailChoice = Table1.Option("Desktop Options", 0, 2, 1, 0, 0, Array("OffCab", "Desktop1", "Desktop2"))
  SetRails RailChoice

  SidewallChoice = Table1.Option("Cabinet Options", 0, 2, 1, 0, 0, Array("Desktop", "PinCab1", "PinCab2"))
  SetSidewall SidewallChoice
        If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If

    GuitarChoice = Table1.Option("Guitar Visible", 0, 1, 1, 1, 0, Array("False", "True"))
  SetGuitar GuitarChoice
    End Sub

Sub SetRails(Opt)
  Select Case Opt
    Case 0:
      Ramp41.Visible = 0
      Ramp42.Visible = 0
      RightCab.visible= 0
            LeftCab.visible= 0
    Case 1:
      Ramp41.Visible = 1
      Ramp42.Visible = 1
      RightCab.visible= 1
            LeftCab.visible= 1
            RightCab.image = "sidecabR"
            LeftCab.image = "sidecabL"
        Case 2:
      Ramp41.Visible = 1
      Ramp42.Visible = 1
      RightCab.visible= 1
            LeftCab.visible= 1
            RightCab.image = "sidecabR2"
            LeftCab.image = "sidecabL2"
  End Select
End Sub


Sub SetSidewall(Opt)
  Select Case Opt
    Case 0:
      Wall001.sidevisible = 0
      Wall002.sidevisible = 0
      Wall013.sidevisible = 0
      Wall014.sidevisible = 0
    Case 1:
      Wall001.sidevisible = 1
      Wall002.sidevisible = 1
            Wall013.sidevisible = 0
      Wall014.sidevisible = 0
      Wall001.Image = "blader"
      Wall002.Image = "bladel"
    Case 2:
            Wall001.sidevisible = 0
      Wall002.sidevisible = 0
      Wall013.sidevisible = 1
      Wall014.sidevisible = 1
      Wall013.Image = "blader_2"
      Wall014.Image = "bladel_2"
  End Select
End Sub

Sub SetGuitar(Opt)
  Select Case Opt
    Case 0:
      Guitar.Visible = 0

    Case 1:
      Guitar.Visible = 1
  End Select
End Sub
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  Controller VBS stuff, but with b2s not started
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'***Controller.vbs version 1.2***'
'
'by arngrim
'
'This script was written to have a generic way to define a controller, no matter if the table is EM or SS based.
'It will also try to load the B2S.Server and if it is not present (or forced off),
'just the standard VPinMAME.Controller is loaded for SS generation games, or no controller for EM ones.
'
'At the first launch of a table using Controller.vbs, it will write into the registry with these default values
'ForceDisableB2S = 0
'DOFContactors   = 2
'DOFKnocker      = 2
'DOFChimes       = 2
'DOFBell         = 2
'DOFGear         = 2
'DOFShaker       = 2
'DOFFlippers     = 2
'DOFTargets      = 2
'DOFDropTargets  = 2
'
'Note that the value can be 0,1 or 2 (0 enables only digital sound, 1 only DOF and 2 both)
'
'If B2S.Server is setup but one doesn't want to use it, one should change the registry entry for ForceDisableB2S to 1
'
'
'Table script usage:
'
'This needs to be added on top of the script on both SS and EM tables:
'
'  On Error Resume Next
'  ExecuteGlobal GetTextFile("Controller.vbs")
'  If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the same folder as this table."
'  On Error Goto 0
'
'In addition the name of the rom (or the fake rom name for EM tables) is needed, because we need it for B2S (and loading VPM):
'
'  cGameName = "rom_name"
'
'For SS tables, the traditional LoadVPM method must be -removed- from the script
'as it is fully integrated into this script (leave the actual call in the script, of course),
'so search for something like this in the table script and -comment out or delete-:
'
'  Sub LoadVPM(VPMver, VBSfile, VBSver)
'    On Error Resume Next
'    If ScriptEngineMajorVersion <5 Then MsgBox "VB Script Engine 5.0 or higher required"
'    ExecuteGlobal GetTextFile(VBSfile)
'    If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
'
'    Set Controller = CreateObject("B2S.Server")
'    'Set Controller = CreateObject("VPinMAME.Controller")
'
'    If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
'    If VPMver> "" Then If Controller.Version <VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
'    If VPinMAMEDriverVer <VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
'    On Error Goto 0
'  End Sub
'
'For SS tables with bad/outdated support by B2S Server (unsupported solenoids, lamps) one can call:
'
'  LoadVPMALT
'
'For EM tables, in the table_init, call:
'
'  LoadEM
'
'For PROC tables, in the table_init, call:
'
'  LoadPROC
'
'Finally, all calls to the B2S.Server Controller properties must be surrounded by a B2SOn check, so for example:
'
'  If B2SOn Then Controller.B2SSetGameOver 1
'
'Or "If ... End If" for multiple script lines that feature the B2S.Server Controller properties, for example:
'
'  If B2SOn Then
'    Controller.B2SSetTilt 0
'    Controller.B2SSetCredits Credits
'    Controller.B2SSetGameOver 1
'  End If
'
'That's all :)
'
'
'Optionally, if one wants to add the automatic ability to mute sounds and switch to DOF calls instead
'(based on the toy configuration that is set at the first run of a table), one can use three variants:
'
'For SS tables:
'
'  PlaySound SoundFX("sound", DOF_toy_category)
'
'If the specific DOF_toy_category (knocker, chimes, etc) is set to 1 in the Controller.txt,
'it will not play the sound but play "" instead.
'
'For EM tables, usually DOF calls are scripted and directly linked with a sound, so SoundFX and DOF can be combined to one method:
'
'  PlaySound SoundFXDOF("sound", DOFevent, State, DOF_toy_category)
'
'If the specific DOF_toy_category (knocker, chimes, etc) is set to 1 in the Controller.txt,
'it will not play the sound but just trigger the DOF call instead.
'
'For pure DOF calls without any sound (lights for example), the DOF method can be used:
'
'  DOF(DOFevent, State)
'

Const directory = "HKEY_CURRENT_USER\SOFTWARE\Visual Pinball\Controller\"
Dim objShell
Dim PopupMessage
Dim B2SController
Dim Controller
Const DOFContactors = 1
Const DOFKnocker = 2
Const DOFChimes = 3
Const DOFBell = 4
Const DOFGear = 5
Const DOFShaker = 6
Const DOFFlippers = 7
Const DOFTargets = 8
Const DOFDropTargets = 9
Const DOFOff = 0
Const DOFOn = 1
Const DOFPulse = 2

Dim DOFeffects(9)
Dim B2SOn
Dim B2SOnALT

Sub LoadEM
  LoadController("EM")
End Sub

Sub LoadPROC(VPMver, VBSfile, VBSver)
  LoadVBSFiles VPMver, VBSfile, VBSver
  LoadController("PROC")
End Sub

Sub LoadVPM(VPMver, VBSfile, VBSver)
  LoadVBSFiles VPMver, VBSfile, VBSver
  LoadController("VPM")
End Sub

'This is used for tables that need 2 controllers to be launched, one for VPM and the second one for B2S.Server
'Because B2S.Server can't handle certain solenoid or lamps, we use this workaround to communicate to B2S.Server and DOF
'By scripting the effects using DOFAlT and SoundFXDOFALT and B2SController
Sub LoadVPMALT(VPMver, VBSfile, VBSver)
  LoadVBSFiles VPMver, VBSfile, VBSver
  LoadController("VPMALT")
End Sub

Sub LoadVBSFiles(VPMver, VBSfile, VBSver)
  On Error Resume Next
  If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
  ExecuteGlobal GetTextFile(VBSfile)
  If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
  InitializeOptions
End Sub

Sub LoadVPinMAME
  Set Controller = CreateObject("VPinMAME.Controller")
  If Err Then MsgBox "Can't load VPinMAME." & vbNewLine & Err.Description
  If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
  If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
  On Error Goto 0
End Sub

'Try to load b2s.server and if not possible, load VPinMAME.Controller instead.
'The user can put a value of 1 for ForceDisableB2S, which will force to load VPinMAME or no controller for EM tables.
'Also defines the array of toy categories that will either play the sound or trigger the DOF effect.
Sub LoadController(TableType)
  Dim FileObj
  Dim DOFConfig
  Dim TextStr2
  Dim tempC
  Dim count
  Dim ISDOF
  Dim Answer

  B2SOn = False
  B2SOnALT = False
  tempC = 0
  on error resume next
  Set objShell = CreateObject("WScript.Shell")
  objShell.RegRead(directory & "ForceDisableB2S")
  If Err.number <> 0 Then
    PopupMessage = "This latest version of Controller.vbs stores its settings in the registry. To adjust the values, you must use VP 10.2 (or newer) and setup your configuration in the DOF section of the -Keys, Nudge and DOF- dialog of Visual Pinball."
    objShell.RegWrite directory & "ForceDisableB2S",0, "REG_DWORD"
    objShell.RegWrite directory & "DOFContactors",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFKnocker",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFChimes",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFBell",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFGear",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFShaker",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFFlippers",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFTargets",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFDropTargets",2, "REG_DWORD"
    MsgBox PopupMessage
  End If
  tempC = objShell.RegRead(directory & "ForceDisableB2S")
  DOFeffects(1)=objShell.RegRead(directory & "DOFContactors")
  DOFeffects(2)=objShell.RegRead(directory & "DOFKnocker")
  DOFeffects(3)=objShell.RegRead(directory & "DOFChimes")
  DOFeffects(4)=objShell.RegRead(directory & "DOFBell")
  DOFeffects(5)=objShell.RegRead(directory & "DOFGear")
  DOFeffects(6)=objShell.RegRead(directory & "DOFShaker")
  DOFeffects(7)=objShell.RegRead(directory & "DOFFlippers")
  DOFeffects(8)=objShell.RegRead(directory & "DOFTargets")
  DOFeffects(9)=objShell.RegRead(directory & "DOFDropTargets")
  Set objShell = nothing

  If TableType = "PROC" or TableType = "VPMALT" Then
    If TableType = "PROC" Then
      Set Controller = CreateObject("VPROC.Controller")
      If Err Then MsgBox "Can't load PROC"
    Else
      LoadVPinMAME
    End If
    If tempC = 0 Then
      On Error Resume Next
      If Controller is Nothing Then
        Err.Clear
      Else
        Set B2SController = CreateObject("B2S.Server")
        If B2SController is Nothing Then
          Err.Clear
        Else
          B2SController.B2SName = B2ScGameName
          B2SController.Run()
          On Error Goto 0
          B2SOn = True
          B2SOnALT = True
        End If
      End If
    End If
  Else
    If tempC = 0 Then
      On Error Resume Next
      Set Controller = CreateObject("B2S.Server")
      If Controller is Nothing Then
        Err.Clear
        If TableType = "VPM" Then
          LoadVPinMAME
        End If
      Else
        Controller.B2SName = cGameName
        If TableType = "EM" Then
          Controller.Run()
        End If
        On Error Goto 0
        B2SOn = True
      End If
    Else
      If TableType = "VPM" Then
        LoadVPinMAME
      End If
    End If
    Set DOFConfig=Nothing
    Set FileObj=Nothing
  End If
End sub

'Additional DOF sound vs toy/effect helpers:

'Mostly used for SS tables, returns the sound to be played or no sound,
'depending on the toy category that is set to play the sound or not.
'The trigger of the DOF Effect is set at the DOF method level
'because for SS tables we usually don't need to script the DOF calls.
'Just map the Solenoid, Switches and Lamps in the ini file directly.
Function SoundFX (Sound, Effect)
  If ((Effect = 0 And B2SOn) Or DOFeffects(Effect)=1) Then
    SoundFX = ""
  Else
    SoundFX = Sound
  End If
End Function

'Mostly used for EM tables, because in EM there is most often a direct link
'between a sound and DOF Trigger, DOFevent is the ID reference of the DOF Call
'that is used in the DOF ini file and State defines if it pulses, goes on or off.
'Example based on the constants that must be present in the table script:
'SoundFXDOF("flipperup",101,DOFOn,contactors)
Function SoundFXDOF (Sound, DOFevent, State, Effect)
  If DOFeffects(Effect)=1 Then
    SoundFXDOF = ""
    DOF DOFevent, State
  ElseIf DOFeffects(Effect)=2 Then
    SoundFXDOF = Sound
    DOF DOFevent, State
  Else
    SoundFXDOF = Sound
  End If
End Function

'Method used to communicate to B2SController instead of the usual Controller
Function SoundFXDOFALT (Sound, DOFevent, State, Effect)
  If DOFeffects(Effect)=1 Then
    SoundFXDOFALT = ""
    DOFALT DOFevent, State
  ElseIf DOFeffects(Effect)=2 Then
    SoundFXDOFALT = Sound
    DOFALT DOFevent, State
  Else
    SoundFXDOFALT = Sound
  End If
End Function

'Pure method that makes it easier to call just a DOF Event.
'Example DOF 123, DOFOn
'Where 123 refers to E123 in a line in the DOF ini.
Sub DOF(DOFevent, State)
  If B2SOn Then
    If State = 2 Then
      Controller.B2SSetData DOFevent, 1:Controller.B2SSetData DOFevent, 0
    Else
      Controller.B2SSetData DOFevent, State
    End If
  End If
End Sub

'If PROC or B2SController is used, we need to pass B2S events to the B2SController instead of the usual Controller
'Use this method to pass information to DOF instead of the Sub DOF
Sub DOFALT(DOFevent, State)
  If B2SOnALT Then
    If State = 2 Then
      B2SController.B2SSetData DOFevent, 1:B2SController.B2SSetData DOFevent, 0
    Else
      B2SController.B2SSetData DOFevent, State
    End If
  End If
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   KEYS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X


  Sub Table1_KeyDown(ByVal Keycode)


  If ballrolleron = 1 then
    if keycode = 46 then ' C Key
       If contball = 1 Then
          contball = 0
       Else
          contball = 1
       End If
    End If
  End If
  if keycode = 48 then 'B Key
     If bcboost = 1 Then
        bcboost = bcboostmulti
     Else
        bcboost = 1
     End If
  End If
  if keycode = 203 then bcleft = 1 ' Left Arrow
  if keycode = 200 then bcup = 1 ' Up Arrow
  if keycode = 208 then bcdown = 1 ' Down Arrow
  if keycode = 205 then bcright = 1 ' Right Arrow


    If keycode = PlungerKey Then
      PlaySoundAt "fx_plungerpull", Plunger
      Plunger.Pullback
    End If

    If hsbModeActive = True Then
      EnterHighScoreKey(keycode)
    elseif bGameInPlay Then

      If inhighscore = False Then
      If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25:CheckTilt
      If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25:CheckTilt
      If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound SoundFX("fx_nudge",0), 0, 1, 1, 0.25:CheckTilt
      End If
      If NOT Tilted Then
      If keycode = LeftFlipperKey Then SolLFlipper 1:SolULFlipper 1:ldown = 1:checkdown
      If keycode = RightFlipperKey Then SolRFlipper 1:SolURFlipper 1:rdown = 1:checkdown

      If keycode = StartGameKey Then
        If((PlayersPlayingGame <MaxPlayers) AND(bOnTheFirstBall = True) ) Then
            PlayersPlayingGame = PlayersPlayingGame + 1
            PuPlayer.playlistplayex pCallouts,"audiocallouts","player1.wav",80,1
    chilloutthemusic
            If PlayersPlayingGame = 2 Then
              PuPlayer.LabelSet pBackglass,"Play2","PLAYER 2",1,"{'mt':2,'color':16777215, 'size': 1.5, 'xpos': 93.3, 'xalign': 0}"
              pUpdateScores
            PuPlayer.playlistplayex pCallouts,"audiocallouts","player2.wav",80,1
    chilloutthemusic
            End If
            If PlayersPlayingGame = 3 Then
              PuPlayer.LabelSet pBackglass,"Play3","PLAYER 3",1,"{'mt':2,'color':16777215, 'size': 1.5, 'xpos': 93.3, 'xalign': 0}"
              pUpdateScores
            PuPlayer.playlistplayex pCallouts,"audiocallouts","player3.wav",80,1
    chilloutthemusic
            End If
            If PlayersPlayingGame = 4 Then
              PuPlayer.LabelSet pBackglass,"Play4","PLAYER 4",1,"{'mt':2,'color':16777215, 'size': 1.5, 'xpos': 93.3, 'xalign': 0}"
              pUpdateScores
            PuPlayer.playlistplayex pCallouts,"audiocallouts","player4.wav",80,1
    chilloutthemusic
            End If
            TotalGamesPlayed = TotalGamesPlayed + 1
            savegp
            'DMDFlush
            'DMD "black.png", " ", PlayersPlayingGame & " PLAYERS",  500
            PlaySound "so_fanfare1"
        End If
      End If
      End If
      Else
      If NOT Tilted Then
   ' If (GameInPlay)
          'If keycode = RightFlipperKey Then 'DMDFlush
        If keycode = LeftFlipperKey Then SolLFlipper 0:SolULFlipper 0:helptime.enabled = true:DMDintroloop:introtime = 0
        If keycode = RightFlipperKey Then SolRFlipper 0:SolURFlipper 0:helptime.enabled = true:DMDintroloop:introtime = 0
          If keycode = StartGameKey Then
              If(BallsOnPlayfield = 0) Then
                ResetForNewGame()
              End If
          End If
      End If
      End If ' If (GameInPlay)

  End Sub

  sub fireitbitch
Plunger.Fire
      bAutoPlunger = False
      Plunger.AutoPlunger = false
  end sub

  Sub Table1_KeyUp(ByVal keycode)


  if keycode = 203 then bcleft = 0 ' Left Arrow
  if keycode = 200 then bcup = 0 ' Up Arrow
  if keycode = 208 then bcdown = 0 ' Down Arrow
  if keycode = 205 then bcright = 0 ' Right Arrow

    If keycode = PlungerKey Then
      PlaySoundAt "fx_plunger", Plunger
      Plunger.Fire
    End If

    ' Table specific

    If bGameInPLay and hsbModeActive <> True Then


      If keycode = LeftFlipperKey Then
        SolLFlipper 0
        SolULFlipper 0
        ldown = 0
        'InstantInfoTimer.Enabled = False
        'If bInstantInfo Then
          'DMDScoreNow
          'bInstantInfo = False
        'End If
      End If
      If keycode = RightFlipperKey Then
        SolRFlipper 0
        SolURFlipper 0
        rdown = 0
        'InstantInfoTimer.Enabled = False
        'If bInstantInfo Then
          'DMDScoreNow
          'bInstantInfo = False
        'End If
      End If
    Else
      If keycode = LeftFlipperKey Then SolLFlipper 0:SolULFlipper 0:helptime.enabled = false
      If keycode = RightFlipperKey Then SolRFlipper 0:SolURFlipper 0:helptime.enabled = false
    End If

  End Sub


  '*************
  ' Pause Table
  '*************

  Sub table1_Paused
  End Sub

  Sub table1_unPaused
  End Sub

  Sub table1_Exit
    Savehs
  End Sub



  '********************
  '     Flippers
  '********************


  Sub SolLFlipper(Enabled)
    If finalflips = False Then
    If lowerflippersoff = True Then
    If Enabled Then
            FlipperActivate LeftFlipper, LFPress
            LF.Fire
      PlaySoundAt SoundFXDOF("Flipper_L01", 101, DOFOn, DOFFlippers), lane3
      If bSkillshotReady = False Then
        RotateLaneLightsLeft
      End If
    Else
            FlipperDeActivate LeftFlipper, LFPress
      PlaySoundAt SoundFXDOF("Flipper_LD", 101, DOFOff, DOFFlippers), lane3
      LeftFlipper.RotateToStart
    End If
    End If
    End If
  End Sub

  Dim lowerflippersoff

  Sub SolULFlipper(Enabled)
    If finalflips = False Then
    If lowerflippersoff = False Then
    If Enabled Then
         FlipperActivate Flipper2, LFPress1
      PlaySoundAt SoundFXDOF("Flipper_R01", 101, DOFOn, DOFFlippers), Flipper2
      Flipper2.RotateToEnd
    Else
            FlipperDeActivate Flipper2, LFPress1
      PlaySoundAt SoundFXDOF("Flipper_RD", 101, DOFOff, DOFFlippers), Flipper2
      Flipper2.RotateToStart
    End If
    End If
    End If
  End Sub

  Sub SolRFlipper(Enabled)
    If finalflips = False Then
    If lowerflippersoff = True Then
    If Enabled Then
            FlipperActivate RightFlipper, RFPress
            RF.Fire
      PlaySoundAt SoundFXDOF("Flipper_R01", 102, DOFOn, DOFFlippers), lane5
      If bSkillshotReady = False Then
        RotateLaneLightsRight
      End If
    Else
            FlipperDeActivate RightFlipper, RFPress
      PlaySoundAt SoundFXDOF("Flipper_RD", 102, DOFOff, DOFFlippers), lane5
      RightFlipper.RotateToStart
    End If
    End If
    End If
  End Sub

  Sub SolURFlipper(Enabled)
    If finalflips = False Then
    If lowerflippersoff = False Then
    If Enabled Then
            FlipperActivate Flipper2, RFPress1
      PlaySoundAt SoundFXDOF("Flipper_L01", 102, DOFOn, DOFFlippers), Flipper1
      Flipper1.RotateToEnd
    Else
            FlipperDeActivate Flipper2, RFPress1
      PlaySoundAt SoundFXDOF("Flipper_LD", 102, DOFOff, DOFFlippers), Flipper1
      Flipper1.RotateToStart
    End If
    End If
    End If
  End Sub

  ' flippers hit Sound

  Sub LeftFlipper_Collide(parm)
        CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
      LF.ReProcessBalls ActiveBall
    PlaySoundAtVol "fx_rubber_flipper", ActiveBall, parm
  End Sub

  Sub RightFlipper_Collide(parm)
        CheckLiveCatch Activeball, RightFlipper, RFCount, parm
      RF.ReProcessBalls ActiveBall
    PlaySoundAtVol "fx_rubber_flipper", ActiveBall, parm
  End Sub

  Sub Flipper2_Collide(parm)
    PlaySoundAtVol "fx_rubber_flipper", ActiveBall, parm
  End Sub

  Sub Flipper1_Collide(parm)
    PlaySoundAtVol "fx_rubber_flipper", ActiveBall, parm
  End Sub

  Sub RotateLaneLightsLeft
    Dim TempState
    'flipper lanes
    TempState = ll1.State
    ll1.State = ll2.State
    ll2.State = ll3.State
    ll3.State = ll4.State
    ll4.State = ll5.State
    ll5.State = TempState
    ll8.state = ll1.state
    ll7.state = ll2.state
    ll6.state = ll3.state
    ll9.state = ll4.state
    ll10.state = ll5.state
  End Sub

  Sub RotateLaneLightsRight
    Dim TempState
    'flipperlanes
    TempState = ll5.State
    ll5.State = ll4.State
    ll4.State = ll3.State
    ll3.State = ll2.State
    ll2.State = ll1.State
    ll1.State = TempState
    ll8.state = ll1.state
    ll7.state = ll2.state
    ll6.state = ll3.state
    ll9.state = ll4.state
    ll10.state = ll5.state
  End Sub

  Dim helpavail
  helpavail = 0
  helptime.enabled = false
  Dim helpon
  helpon = false

  Sub helptime_timer
    helpavail = helpavail + 1
    Select Case helpavail
    Case 2
      If helpon = false Then
        help.opacity = 100
        helpon = true
        helptime.enabled = false
        helpavail = 0
      Else
        help.opacity = 0
        helpon = false
        helptime.enabled = false
        helpavail = 0
      End If
    End Select
  End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   TILT
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X


  'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

  Sub CheckTilt                                    'Called when table is nudged
    Tilt = Tilt + TiltSensitivity                'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt> TiltSensitivity) AND(Tilt <15) Then 'show a warning
      pNote "CAREFUL!","MOUTHBREATHER"
        PlaySound "buzz"
      PuPlayer.playlistplayex pBackglass,"videotilt","",100,1
      DOF 131, DOFPulse
      DOF 311, DOFPulse  'DOF MX - Tilt Warning
    End if
    If Tilt> 15 Then 'If more that 15 then TILT the table
      Tilted = True
      pNote "TILT","MOUTHBREATHER"
        PlaySound "powerdownn"
      PuPlayer.playlistplayex pBackglass,"videotilt","",100,4
      DOF 310, DOFPulse   'DOF MX - TILT
      DisableTable True
      tilttableclear.enabled = true
      TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
    End If
  End Sub

  Dim tilttime:tilttime = 0

  sub tilttableclear_timer
    tilttime = tilttime + 1
    Select Case tilttime
      Case 10
        tableclearing
    End Select
  End Sub

  Sub tableclearing
    dropwallskip = 1
    barricadetime.Enabled = True
    diverter.RotZ = -33
    diverteropen.collidable = True
    diverter.collidable = False
    exitav
    exitcastle
    vpmtimer.addtimer 2000, "posttiltreset'"
    BallLockRunExit
    BallLockBarbExit
    wallrelease
    MagnetR.MagnetON = False
    KickerEL.Kick 90, 4
    BallRelease.Kick 90, 4
    BallRelease.Kick 90, 4
  End Sub

  Sub posttiltreset
    CloseGates
    barricadetime.Enabled = True
    tilttime = 0
  End Sub

  Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt> 0 Then
      Tilt = Tilt - 0.1
    Else
      TiltDecreaseTimer.Enabled = False
    End If
  End Sub

  Sub DisableTable(Enabled)
    If Enabled Then
      'turn off GI and turn off all the lights
      GiOff
      LightSeqTilt.Play SeqAllOff
      'Disable slings, bumpers etc
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      'Bumper1.Force = 0

      LeftSlingshot.Disabled = 1
      RightSlingshot.Disabled = 1
    PuPlayer.playresume 4
    PuPlayer.playlistplayex pAudio,"audiomodes","clear.mp3",100,1
    PuPlayer.playlistplayex pMusic,"audioclear","clear.mp3",100, 1

    Else
      'turn back on GI and the lights
      GiOn
      LightSeqTilt.StopPlay
      'Bumper1.Force = 6
      LeftSlingshot.Disabled = 0
      RightSlingshot.Disabled = 0
      'clean up the buffer display
      'DMDFlush
    End If
  End Sub

  Sub TiltRecoveryTimer_Timer()
    ' if all the balls have been drained then..
    If(BallsOnPlayfield = 0) Then
      ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
      EndOfBall()
      TiltRecoveryTimer.Enabled = False
    End If
  ' else retry (checks again in another second or so)
  End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   BALL SHADOW
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X


  Dim BallShadow
  BallShadow = Array (BallShadow1)

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
        BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
      Else
        BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
      End If
      ballShadow(b).Y = BOT(b).Y + 20
      If BOT(b).Z > 20 Then
        BallShadow(b).visible = 1
      Else
        BallShadow(b).visible = 0
      End If
    Next
  End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   SOUND FUNCTIONS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X

  '********************
  ' Music as wav sounds
  '********************

  Dim Song
  Song = ""

  Sub PlaySong(name)
    If bMusicOn Then
      If Song <> name Then
        StopSound Song
        Song = name
        If Song = "m_end" Then
          PlaySound Song, 0, 0.1  'this last number is the volume, from 0 to 1
        Else
          PlaySound Song, -1, 0.1 'this last number is the volume, from 0 to 1
        End If
      End If
    End If
  End Sub

  Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 200)
  End Function

  Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp> 0 Then
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

  '********************
  ' SSF supporting functions
  '********************

  function AudioFade(ball)
    Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
      AudioFade = Csng(tmp ^10)
    Else
      AudioFade = Csng(-((- tmp) ^10) )
    End If
  End Function

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
    PlaySound sound, 1, Vol, Pan(tableobj), 0, 0, 0, 1, AudioFade(tableobj)
  End Sub

  Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
    PlaySound sound, -1, Vol, Pan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
  End Sub

  'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3
  Sub PlaySoundAtBallVol(sound, VolMult)
    PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
  End Sub

  'Set position as bumperX and Vol manually.
  Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
    PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
  End Sub

  Dim NextOrbitHit:NextOrbitHit = 0
  Sub PlasticRampBumps_Hit(idx)
    if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
      RandomBump 10, -20000
      ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
      ' Lowering these numbers allow more closely-spaced clunks.
      NextOrbitHit = Timer + .1 + (Rnd * .2)
    end if
  End Sub

  Sub PlasticBumps_Hit(idx)
    PlaySoundAtBall "fx_plastichit"
  End Sub

  Sub MetalWallBumps_Hit(idx)
    'debug.print "MetalWall"
         PlaySoundAtBall "fx_Metalhit"
    if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
      RandomBump 3, 10000 'Increased pitch to simulate metal wall
      ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
      ' Lowering these numbers allow more closely-spaced clunks.
      NextOrbitHit = Timer + .8 + (Rnd * .8)
    end if
  End Sub

  Sub WireRampBumps_Hit(idx)
    if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
      dim BumpSnd:BumpSnd= "wirerampbump" & CStr(Int(Rnd*5)+1)
      PlaySound BumpSnd, 0, Vol(ActiveBall) * .5, Pan(ActiveBall), 0, 30000, 0, 1, AudioFade(ActiveBall)
      NextOrbitHit = Timer + .2 + (Rnd * .2)
    end if
  End Sub

  ' Requires rampbump1 to 7 in Sound Manager
  Sub RandomBump(voladj, freq)
    dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
  End Sub



  ' Stop Bump Sounds
  Sub BumpSTOPMetal ()
  dim i:for i=1 to 7:StopSound "RampBump" & i:next
  NextOrbitHit = Timer + 1
  End Sub

  Sub BumpSTOPWire ()
  dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
  NextOrbitHit = Timer + 1
  End Sub


  '->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->
  '-> Real Time updates using the GameTimer
  '->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->

' Thalamus - sub used twice - not a big issue though, both are the same.
' Sub GameTimer_Timer
'   RollingUpdate
' End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 20 ' total number of balls
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
    If UBound(BOT) = -1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) * 100
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ) * .5, Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
End Sub

  '->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->
  '->  Sound FX Groupings
  '->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->
  ' Make sure to make collections for each of these to fire the sound fxs

  Sub OnBallBallCollision(ball1, ball2, velocity)
        FlipperCradleCollision ball1, ball2, velocity
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 500, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  End Sub

  Sub Pins_Hit (idx)
    PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  End Sub

  Sub Targets_Hit (idx)
    PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  End Sub

  Sub Gates_Hit (idx)
    PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Sub


  Sub Rubbers_Hit(idx)
    debug.print "Rubbers"
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
      PlaySound "fx_rubber2", 0, Vol(ActiveBall)*3, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    else
      RandomSoundRubber()
    End If
  End Sub

  Sub Posts_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then
      RandomSoundHole
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
      RandomSoundRubber()
    End If
  End Sub

  '->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->
  '-> Sound Randomizers
  '->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->
  Sub RandomSoundFlipper()
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySound "fx_flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
      Case 2 : PlaySound "fx_flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
      Case 3 : PlaySound "fx_flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
  End Sub

  Sub RandomSoundRubber()
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySound "fx_rubber_hit_1", 0, Vol(ActiveBall)*3, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
      Case 2 : PlaySound "fx_rubber_hit_2", 0, Vol(ActiveBall)*3, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
      Case 3 : PlaySound "fx_rubber_hit_3", 0, Vol(ActiveBall)*3, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
  End Sub

  Sub RandomSoundHole()
    Select Case RndNum(1,2)
      Case 1 : PlaySoundAt "fx_kicker_enter", ActiveBall
      Case 2 : PlaySoundAt "fx_kicker_catch", ActiveBall
    End Select
  End Sub


  '

  '******************************
  ' Diverse Collection Hit Sounds
  '******************************


  ' Random quotes from the game

  Sub PlayQuote_timer() 'one quote each 2 minutes
    Dim Quote
    Quote = "di_quote" & INT(RND * 56) + 1
    PlaySound Quote
  End Sub

  ' Ramp Soundss

  Sub rrend_Hit(): StopSound "fx_metalrolling": BumpSTOPWire(): vpmTimer.AddTimer 100, "PlaySoundAt ""fx_balldrop"",rrend'":End Sub
  Sub lrend_Hit(): StopSound "fx_metalrolling": BumpSTOPWire(): vpmTimer.AddTimer 100, "PlaySoundAt ""fx_balldrop"",lrend'":End Sub


  '->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->
  '->  Ramp Sounds, Use as needed
  '->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->->
  'shooter ramp
  Sub ShooterStart_Hit():StopSound "fx_launchball":If ActiveBall.VelY < 0 Then PlaySoundAt "fx_launchball", ShooterStart:End If:End Sub 'ball is going up
  Sub ShooterEnd_Hit:If ActiveBall.Z > 50  Then Me.TimerInterval=100:Me.TimerEnabled=1:End If:End Sub           'ball is flying
  Sub ShooterEnd_Timer(): Me.TimerEnabled=0 : PlaySoundAt "fx_balldrop", ShooterEnd : End Sub
  'center ramp
  Sub CREnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAt "fx_lrenter", CREnter:End If
    spinsign
  End Sub     'ball is going up
  Sub CREnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_metalrolling":StopSound "fx_lrenter":End If:End Sub   'ball is going down
  Sub CREnter1_Hit():If ActiveBall.VelY < 0 Then PlaySoundAt "fx_lrenter", CRenter1:End If:End Sub      'ball is going up
  Sub CREnter1_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_metalrolling":StopSound "fx_lrenter":End If:End Sub    'ball is going down
  Sub CREnter2_Hit():If ActiveBall.VelY < 0 Then PlaySoundAt "fx_lrenter", CRenter2:End If:End Sub      'ball is going up
  Sub CREnter2_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_metalrolling":StopSound "fx_lrenter":End If:End Sub    'ball is going down





'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   START GAME, END GANE
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X

'
  Sub ResetForNewGame()
    Dim i


    bGameInPLay = True

    'resets the score display, and turn off attract mode
    StopAttractMode

    clearhslabels
    clearosblabels

    GiOn

    TotalGamesPlayed = TotalGamesPlayed + 1
    savegp
    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
      Score(i) = 0
      BonusPoints(i) = 0
      BonusHeldPoints(i) = 0
      BonusMultiplier(i) = 1
      BallsRemaining(i) = 3
      ExtraBallsAwards(i) = 0
    Next
    'For x = 0 to UBOUND(aLights)
    ' lamps(CurrentPlayer,x) = x.State
    'Next
    ' initialise any other flags
    Tilt = 0

    ' initialise Game variables
    Game_Init()

    ' you may wish to start some music, play a sound, do whatever at this point

    ' set up the start delay to handle any Start of Game Attract Sequence
    vpmtimer.addtimer 1500, "FirstBall '"



  End Sub


  Sub EndOfGame()
    PuPlayer.playlistplayex pMusic,"audioclear","clear.mp3",100, 1
    'pNote "GAME OVER","PLAY AGAIN"
    PuPlayer.playlistplayex pBackglass,"videogameover","",100,1
    StartAttractMode
    'debug.print "End Of Game"
    introposition = 0
    bGameInPLay = False
    ' just ended your game then play the end of game tune
    If NOT bJustStarted Then
      PlaySong "m_end"
    End If
    bJustStarted = False
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0
    SolULFlipper 0
    SolURFlipper 0
    BallsInLock(CurrentPlayer) = 0
    BallsInRunLock(CurrentPlayer) = 0
    beaconoff
    Raisewall



    ' terminate all Mode - eject locked balls
    ' most of the Mode/timers terminate at the end of the ball
    'PlayQuote.Enabled = 0
    ' show game over on the 'DMD
    'DMD "black.png", "Game Over", "",  2000
    Dim i
    If Score(1) Then
      'DMD "black.png", "PLAYER 1", Score(1), 3000
      ' Submit Player 1 Score to Orbital Scoreboard
    End If
    If Score(2) Then
      'DMD "black.png", "PLAYER 2", Score(2), 3000
    End If
    If Score(3) Then
      'DMD "black.png", "PLAYER 3", Score(3), 3000
    End If
    If Score(4) Then
      'DMD "black.png", "PLAYER 4", Score(4), 3000
    End If

    ' set any lights for the attract mode
    GiOff

    dim cn, rs

    i = 0



  ' you may wish to light any Game Over Light you may have
  End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   ULTRADMD SCRIPTING
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X

  Dim UltraDMD

  ' DMD using UltraDMD calls

  Sub DMD(background, toptext, bottomtext, duration)
    If turnonultradmd = 0 then exit sub
    UltraDMD.DisplayScene00 background, toptext, 15, bottomtext, 15, 14, duration, 14
    UltraDMDTimer.Enabled = 1 'to show the score after the animation/message
  End Sub

  Sub DMDScore
    If turnonultradmd = 0 then exit sub
    UltraDMD.SetScoreboardBackgroundImage "scoreboard-background.jpg", 15, 7

    If turnonultradmd = 1 Then
    UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Player " & CurrentPlayer, "Ball " & Balls
    End If

    If turnonultradmd = 2 Then
    UltraDMD.DisplayScoreboard00 PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Player " & CurrentPlayer, "Ball " & Balls
    End If
  End Sub

  Sub DMDScoreNow
    If turnonultradmd = 0 then exit sub
    DMDFlush
    DMDScore
  End Sub

  Sub DMDFLush
    If turnonultradmd = 0 then exit sub
    UltraDMDTimer.Enabled = 0
    UltraDMD.CancelRendering
  End Sub

  Sub DMDScrollCredits(background, text, duration)
    If turnonultradmd = 0 then exit sub
    UltraDMD.ScrollingCredits background, text, 15, 14, duration, 14
  End Sub

  Sub DMDId(id, background, toptext, bottomtext, duration)
    If turnonultradmd = 0 then exit sub
    UltraDMD.DisplayScene00ExwithID id, False, background, toptext, 15, 0, bottomtext, 15, 0, 14, duration, 14
  End Sub

  Sub DMDMod(id, toptext, bottomtext, duration)
    If turnonultradmd = 0 then exit sub
    UltraDMD.ModifyScene00Ex id, toptext, bottomtext, duration
  End Sub

  Sub UltraDMDTimer_Timer 'used for the attrackmode and the instant info.
    If turnonultradmd = 0 then exit sub
    If bInstantInfo Then
      InstantInfo
    ElseIf bAttractMode Then
    ElseIf NOT UltraDMD.IsRendering Then
      DMDScoreNow
    ElseIf bromconfig Then
      romconfig
    End If
  End Sub

  Sub DMD_Init
    If turnonultradmd = 0 then exit sub
    Set UltraDMD = CreateObject("UltraDMD.DMDObject")
    If UltraDMD is Nothing Then
      MsgBox "No UltraDMD found.  This table will NOT run without it."
      Exit Sub
    End If

    UltraDMD.Init
    If turnonultradmd = 0 then exit sub
    If Not UltraDMD.GetMajorVersion = 1 Then
      MsgBox "Incompatible Version of UltraDMD found."
      Exit Sub
    End If

    If UltraDMD.GetMinorVersion <1 Then
    If turnonultradmd = 0 then exit sub
      MsgBox "Incompatible Version of UltraDMD found. Please update to version 1.1 or newer."
      Exit Sub
    End If

    Dim fso:Set fso = CreateObject("Scripting.FileSystemObject")
    Dim curDir:curDir = fso.GetAbsolutePathName(".")

    Dim DirName
    DirName = curDir& "\" &TableName& ".UltraDMD"

    If Not fso.FolderExists(DirName) Then _
        Msgbox "UltraDMD userfiles directory '" & DirName & "' does not exist." & CHR(13) & "No graphic images will be displayed on the DMD"
    UltraDMD.SetProjectFolder DirName

    ' wait for the animation to end
    While UltraDMD.IsRendering = True
    WEnd

  End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   PUPDMD - WIP
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X

'**************************
'   PinUp Player Config
'   Change HasPuP = True if using PinUp Player Videos
'**************************

  Dim HasPup:HasPuP = True

  Dim PuPlayer

      Const pTopper=0
      Const pDMD=1
      Const pBackglass=2
      Const pPlayfield=3
      Const pMusic=4
      Const pAudio=7
      Const pCallouts=8

  if HasPuP Then
  on error resume next
  Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")
  on error goto 0
  if not IsObject(PuPlayer) then HasPuP = False
  end If

  if HasPuP Then

  PuPlayer.Init pBackglass,"STLE"
  PuPlayer.Init pMusic,"STLE"
  PuPlayer.Init pAudio,"STLE"
  PuPlayer.Init pCallouts,"STLE"
  If toppervideo = 1 Then
  PuPlayer.Init pTopper,"STLE"
  End If

  PuPlayer.SetScreenex pBackglass,0,0,0,0,0       'Set PuPlayer DMD TO Always ON    <screen number> , xpos, ypos, width, height, POPUP
  PuPlayer.SetScreenex pAudio,0,0,0,0,2
  PuPlayer.hide pAudio
  PuPlayer.SetScreenex pMusic,0,0,0,0,2
  PuPlayer.hide pMusic
  PuPlayer.SetScreenex pCallouts,0,0,0,0,2
  PuPlayer.hide pCallouts

  Sub chilloutthemusic
    If calloutlowermusicvol = 1 Then
      PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":11, ""VL"":40 }"
      PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":40 }"
      PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 7, ""FN"":11, ""VL"":40 }"
      PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 8, ""FN"":11, ""VL"":"&(calloutvol)&" }"
      vpmtimer.addtimer 2200, "turnitbackup'"
    End If
  End Sub

  Sub turnitbackup
    If calloutlowermusicvol = 1 Then
      PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":11, ""VL"":"&(videovol)&" }"
      PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":"&(soundtrackvol)&" }"
      PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 7, ""FN"":11, ""VL"":"&(soundtrackvol)&" }"
    End If
  End Sub


  PuPlayer.playlistadd pMusic,"audioattract", 1 , 0
  PuPlayer.playlistadd pMusic,"audiobg", 1 , 0
  PuPlayer.playlistadd pMusic,"audioclear", 1 , 0
  PuPlayer.playlistadd pMusic,"audiobgrock", 1 , 0
  PuPlayer.playlistadd pAudio,"audioevents", 1 , 0
  PuPlayer.playlistadd pAudio,"audiomodes", 1 , 0
  PuPlayer.playlistadd pAudio,"audiomultiballs", 1 , 0
  PuPlayer.playlistadd pCallouts,"audiocallouts", 1 , 0
  PuPlayer.playlistadd pCallouts,"audiojackpot", 1 , 0
  PuPlayer.playlistadd pCallouts,"audiosuperjackpot", 1 , 0
  PuPlayer.playlistadd pCallouts,"audiocheers", 1 , 0
  PuPlayer.playlistadd pBackglass,"bgs", 1 , 0
  PuPlayer.playlistadd pBackglass,"scene", 1 , 0
  PuPlayer.playlistadd pBackglass,"videoattract", 1 , 0
  PuPlayer.playlistadd pBackglass,"videobarb", 1 , 0
  PuPlayer.playlistadd pBackglass,"videobarblock", 1 , 0
  PuPlayer.playlistadd pBackglass,"videobarblit", 1 , 0
  PuPlayer.playlistadd pBackglass,"videoavopen", 1 , 0
  PuPlayer.playlistadd pBackglass,"videobadmenlit", 1 , 0
  PuPlayer.playlistadd pBackglass,"videobadmenlock", 1 , 0
  PuPlayer.playlistadd pBackglass,"videobadmenmb", 1 , 0
  PuPlayer.playlistadd pBackglass,"videoballsaved", 1 , 0
  PuPlayer.playlistadd pBackglass,"videobyerscastle", 1 , 0
  PuPlayer.playlistadd pBackglass,"videocompass", 1 , 0
  PuPlayer.playlistadd pBackglass,"videodrain", 1 , 0
  PuPlayer.playlistadd pBackglass,"videoeggos", 1 , 0
  PuPlayer.playlistadd pBackglass,"videoextraball", 1 , 0
  PuPlayer.playlistadd pBackglass,"videogameover", 1 , 0
  PuPlayer.playlistadd pBackglass,"videohighscore", 1 , 0
  PuPlayer.playlistadd pBackglass,"videohighscorescreen", 1 , 0
  PuPlayer.playlistadd pBackglass,"videokickback", 1 , 0
  PuPlayer.playlistadd pBackglass,"videolevitate", 1 , 0
  PuPlayer.playlistadd pBackglass,"videomystery", 1 , 0
  PuPlayer.playlistadd pBackglass,"videoparty", 1 , 0
  PuPlayer.playlistadd pBackglass,"videoplayers", 1 , 0
  PuPlayer.playlistadd pBackglass,"videopool", 1 , 0
  PuPlayer.playlistadd pBackglass,"videoquotes", 1 , 0
  PuPlayer.playlistadd pBackglass,"videoradio", 1 , 0
  PuPlayer.playlistadd pBackglass,"videoskillshot", 1 , 0
  PuPlayer.playlistadd pBackglass,"videotilt", 1 , 0
  PuPlayer.playlistadd pBackglass,"videowill", 1 , 0
  PuPlayer.playlistadd pBackglass,"videowizard", 1 , 0
  If toppervideo = 1 Then
  PuPlayer.playlistadd pTopper,"topper", 1 , 0
  End If

      'pNote "A.V CLUB","NOW OPEN"
      'PuPlayer.playlistplayex pBackglass,"videoavopen","",100,1

  'Set Background video on DMD
    PuPlayer.playlistplayex pBackglass,"scene","base.mov",0,1  'should be an attract background (no text is displayed)
    PuPlayer.SetBackground pBackglass,1

  End if

  If toppervideo = 1 Then
    PuPlayer.playlistplayex pTopper,"topper","topper.mp4",0,1  'should be an attract background (no text is displayed)
    PuPlayer.SetBackground pTopper,1
  End If


  'Init TextOverlay on PUP Screen

  PuPlayer.LabelInit pBackglass

  'Setup Pages.  Note if you use fonts they must be in FONTS folder of the pupVideos\tablename\FONTS
  ' syntax - PuPlayer.LabelNew <screen# or pDMD>,<Labelname>,<fontName>,<size%>,<colour>,<rotation>,<xAlign>,<yAlign>,<xpos>,<ypos>,<PageNum>,<visible>

  'Page 1 (default score display)
  PuPlayer.LabelNew pBackglass,"Play1","AvantGarde-Book",       2,16777215  ,0,2,1,98,67,1,0
  PuPlayer.LabelNew pBackglass,"Play1score","AvantGarde LT Medium", 3,16777215  ,0,2,1,98,70,1,0
  PuPlayer.LabelNew pBackglass,"Play2","AvantGarde-Book",       2,16777215  ,0,2,1,98,73,1,0
  PuPlayer.LabelNew pBackglass,"Play2score","AvantGarde LT Medium", 3,16777215  ,0,2,1,98,76,1,0
  PuPlayer.LabelNew pBackglass,"Play3","AvantGarde-Book",       2,16777215  ,0,2,1,98,79,1,0
  PuPlayer.LabelNew pBackglass,"Play3score","AvantGarde LT Medium", 3,16777215  ,0,2,1,98,82,1,0
  PuPlayer.LabelNew pBackglass,"Play4","AvantGarde-Book",       2,16777215  ,0,2,1,98,85,1,0
  PuPlayer.LabelNew pBackglass,"Play4score","AvantGarde LT Medium", 3,16777215  ,0,2,1,98,88,1,0
  PuPlayer.LabelNew pBackglass,"Ball","AvantGarde-Book",        2,16777215  ,0,2,1,98,63,1,1
  PuPlayer.LabelNew pBackglass,"hstitle","AvantGarde-Book",     1,16777215  ,0,2,1,98,92,1,1
  PuPlayer.LabelNew pBackglass,"hs","AvantGarde-Book",        2,16777215  ,0,2,1,98,94,1,1
  PuPlayer.LabelNew pBackglass,"gptitle","AvantGarde-Book",     1,16777215  ,0,2,1,98,96,1,1
  PuPlayer.LabelNew pBackglass,"gp","AvantGarde-Book",        2,16777215  ,0,2,1,98,98,1,1
  PuPlayer.LabelNew pBackglass,"Willh","AvantGarde LT Medium",    2,16777215  ,0,1,1,8,70,1,1
  PuPlayer.LabelNew pBackglass,"Willj","AvantGarde LT Medium",    2,16777215  ,0,2,1,13,70,1,1
  PuPlayer.LabelNew pBackglass,"badh","AvantGarde LT Medium",     2,16777215  ,0,2,1,8,76,1,1
  PuPlayer.LabelNew pBackglass,"badj","AvantGarde LT Medium",     2,16777215  ,0,2,1,13,76,1,1
  PuPlayer.LabelNew pBackglass,"barbh","AvantGarde LT Medium",    2,16777215  ,0,2,1,8,80,1,1
  PuPlayer.LabelNew pBackglass,"barbj","AvantGarde LT Medium",    2,16777215  ,0,2,1,13,80,1,1
  PuPlayer.LabelNew pBackglass,"partyh","AvantGarde LT Medium",   2,16777215  ,0,2,1,8,85,1,1
  PuPlayer.LabelNew pBackglass,"partyj","AvantGarde LT Medium",   2,16777215  ,0,2,1,13,85,1,1
  PuPlayer.LabelNew pBackglass,"mav","AvantGarde LT Medium",      2,16777215  ,0,2,1,8,85,1,1
  PuPlayer.LabelNew pBackglass,"mwaf","AvantGarde LT Medium",     2,16777215  ,0,1,1,13,85,1,1
  PuPlayer.LabelNew pBackglass,"notetitle","AvantGarde LT Medium",  4,16777215  ,0,1,1,50,87,1,1
  PuPlayer.LabelNew pBackglass,"notecopy","AvantGarde-Book",      2,16777215  ,0,1,1,50,92,1,1
  PuPlayer.LabelNew pBackglass,"titlebg","Fundamental 3D  Brigade", 9,0  ,0,1,1,50,50,1,1
  PuPlayer.LabelNew pBackglass,"title","Fundamental  Brigade",    9,16777215  ,0,1,1,50,50,1,1
  PuPlayer.LabelNew pBackglass,"titlebg2","Fundamental 3D  Brigade",  6,0  ,0,1,1,50,50,1,1
  PuPlayer.LabelNew pBackglass,"title2","Fundamental  Brigade",   6,16777215  ,0,1,1,50,50,1,1
  PuPlayer.LabelNew pBackglass,"modetitle","AvantGarde-Book",     2,16777215  ,0,1,1,80,74,1,1
  PuPlayer.LabelNew pBackglass,"modetimer","AvantGarde LT Medium",  6,16777215  ,0,1,1,80,78,1,1
  PuPlayer.LabelNew pBackglass,"high1name","AvantGarde LT Medium",  5,16777215  ,0,1,1,22,30,1,1
  PuPlayer.LabelNew pBackglass,"high1score","AvantGarde LT Medium", 5,16777215  ,0,1,1,36,30,1,1
  PuPlayer.LabelNew pBackglass,"high2name","AvantGarde LT Medium",  5,16777215  ,0,1,1,22,38,1,1
  PuPlayer.LabelNew pBackglass,"high2score","AvantGarde LT Medium", 5,16777215  ,0,1,1,36,38,1,1
  PuPlayer.LabelNew pBackglass,"high3name","AvantGarde LT Medium",  5,16777215  ,0,1,1,22,46,1,1
  PuPlayer.LabelNew pBackglass,"high3score","AvantGarde LT Medium", 5,16777215  ,0,1,1,36,46,1,1
  PuPlayer.LabelNew pBackglass,"high4name","AvantGarde LT Medium",  5,16777215  ,0,1,1,22,54,1,1
  PuPlayer.LabelNew pBackglass,"high4score","AvantGarde LT Medium", 5,16777215  ,0,1,1,36,54,1,1
  PuPlayer.LabelNew pBackglass,"waf1name","AvantGarde LT Medium",   5,16777215  ,0,1,1,56,30,1,1
  PuPlayer.LabelNew pBackglass,"waf1score","AvantGarde LT Medium",  5,16777215  ,0,1,1,63,30,1,1
  PuPlayer.LabelNew pBackglass,"waf2name","AvantGarde LT Medium",   5,16777215  ,0,1,1,56,38,1,1
  PuPlayer.LabelNew pBackglass,"waf2score","AvantGarde LT Medium",  5,16777215  ,0,1,1,63,38,1,1
  PuPlayer.LabelNew pBackglass,"waf3name","AvantGarde LT Medium",   5,16777215  ,0,1,1,56,46,1,1
  PuPlayer.LabelNew pBackglass,"waf3score","AvantGarde LT Medium",  5,16777215  ,0,1,1,63,46,1,1
  PuPlayer.LabelNew pBackglass,"waf4name","AvantGarde LT Medium",   5,16777215  ,0,1,1,56,54,1,1
  PuPlayer.LabelNew pBackglass,"waf4score","AvantGarde LT Medium",  5,16777215  ,0,1,1,63,54,1,1
  PuPlayer.LabelNew pBackglass,"HighScore","AvantGarde LT Medium",  6,16777215  ,0,0,1,20,30,1,1
  PuPlayer.LabelNew pBackglass,"HighScoreL1","AvantGarde LT Medium",  8,16777215  ,0,0,1,20,40,1,1
  PuPlayer.LabelNew pBackglass,"HighScoreL2","AvantGarde LT Medium",  8,16777215  ,0,0,1,24,40,1,1
  PuPlayer.LabelNew pBackglass,"HighScoreL3","AvantGarde LT Medium",  8,16777215  ,0,0,1,28,40,1,1
  PuPlayer.LabelNew pBackglass,"HighScoreL4","AvantGarde LT Medium",  4,16777215  ,0,0,1,20,50,1,1

'obslabels
  'day
  PuPlayer.LabelNew pBackglass,"dh1n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,30,1,1
  PuPlayer.LabelNew pBackglass,"dh1s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,30,1,1
  PuPlayer.LabelNew pBackglass,"dh2n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,34,1,1
  PuPlayer.LabelNew pBackglass,"dh2s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,34,1,1
  PuPlayer.LabelNew pBackglass,"dh3n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,38,1,1
  PuPlayer.LabelNew pBackglass,"dh3s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,38,1,1
  PuPlayer.LabelNew pBackglass,"dh4n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,42,1,1
  PuPlayer.LabelNew pBackglass,"dh4s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,42,1,1
  PuPlayer.LabelNew pBackglass,"dh5n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,46,1,1
  PuPlayer.LabelNew pBackglass,"dh5s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,46,1,1
  PuPlayer.LabelNew pBackglass,"dh6n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,50,1,1
  PuPlayer.LabelNew pBackglass,"dh6s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,50,1,1
  PuPlayer.LabelNew pBackglass,"dh7n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,54,1,1
  PuPlayer.LabelNew pBackglass,"dh7s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,54,1,1
  PuPlayer.LabelNew pBackglass,"dh8n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,58,1,1
  PuPlayer.LabelNew pBackglass,"dh8s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,58,1,1
  PuPlayer.LabelNew pBackglass,"dh9n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,62,1,1
  PuPlayer.LabelNew pBackglass,"dh9s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,62,1,1
  PuPlayer.LabelNew pBackglass,"dh10n","AvantGarde LT Medium",      5,16777215  ,0,1,1,21,66,1,1
  PuPlayer.LabelNew pBackglass,"dh10s","AvantGarde LT Medium",      5,16777215  ,0,1,1,32,66,1,1
  'week
  PuPlayer.LabelNew pBackglass,"wh1n","AvantGarde LT Medium",       5,16777215  ,0,1,1,54,30,1,1
  PuPlayer.LabelNew pBackglass,"wh1s","AvantGarde LT Medium",       5,16777215  ,0,1,1,65,30,1,1
  PuPlayer.LabelNew pBackglass,"wh2n","AvantGarde LT Medium",       5,16777215  ,0,1,1,54,34,1,1
  PuPlayer.LabelNew pBackglass,"wh2s","AvantGarde LT Medium",       5,16777215  ,0,1,1,65,34,1,1
  PuPlayer.LabelNew pBackglass,"wh3n","AvantGarde LT Medium",       5,16777215  ,0,1,1,54,38,1,1
  PuPlayer.LabelNew pBackglass,"wh3s","AvantGarde LT Medium",       5,16777215  ,0,1,1,65,38,1,1
  PuPlayer.LabelNew pBackglass,"wh4n","AvantGarde LT Medium",       5,16777215  ,0,1,1,54,42,1,1
  PuPlayer.LabelNew pBackglass,"wh4s","AvantGarde LT Medium",       5,16777215  ,0,1,1,65,42,1,1
  PuPlayer.LabelNew pBackglass,"wh5n","AvantGarde LT Medium",       5,16777215  ,0,1,1,54,46,1,1
  PuPlayer.LabelNew pBackglass,"wh5s","AvantGarde LT Medium",       5,16777215  ,0,1,1,65,46,1,1
  PuPlayer.LabelNew pBackglass,"wh6n","AvantGarde LT Medium",       5,16777215  ,0,1,1,54,50,1,1
  PuPlayer.LabelNew pBackglass,"wh6s","AvantGarde LT Medium",       5,16777215  ,0,1,1,65,50,1,1
  PuPlayer.LabelNew pBackglass,"wh7n","AvantGarde LT Medium",       5,16777215  ,0,1,1,54,54,1,1
  PuPlayer.LabelNew pBackglass,"wh7s","AvantGarde LT Medium",       5,16777215  ,0,1,1,65,54,1,1
  PuPlayer.LabelNew pBackglass,"wh8n","AvantGarde LT Medium",       5,16777215  ,0,1,1,54,58,1,1
  PuPlayer.LabelNew pBackglass,"wh8s","AvantGarde LT Medium",       5,16777215  ,0,1,1,65,58,1,1
  PuPlayer.LabelNew pBackglass,"wh9n","AvantGarde LT Medium",       5,16777215  ,0,1,1,54,62,1,1
  PuPlayer.LabelNew pBackglass,"wh9s","AvantGarde LT Medium",       5,16777215  ,0,1,1,65,62,1,1
  PuPlayer.LabelNew pBackglass,"wh10n","AvantGarde LT Medium",      5,16777215  ,0,1,1,54,66,1,1
  PuPlayer.LabelNew pBackglass,"wh10s","AvantGarde LT Medium",      5,16777215  ,0,1,1,65,66,1,1
  ' all-time
  PuPlayer.LabelNew pBackglass,"ah1n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,30,1,1
  PuPlayer.LabelNew pBackglass,"ah1s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,30,1,1
  PuPlayer.LabelNew pBackglass,"ah2n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,34,1,1
  PuPlayer.LabelNew pBackglass,"ah2s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,34,1,1
  PuPlayer.LabelNew pBackglass,"ah3n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,38,1,1
  PuPlayer.LabelNew pBackglass,"ah3s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,38,1,1
  PuPlayer.LabelNew pBackglass,"ah4n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,42,1,1
  PuPlayer.LabelNew pBackglass,"ah4s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,42,1,1
  PuPlayer.LabelNew pBackglass,"ah5n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,46,1,1
  PuPlayer.LabelNew pBackglass,"ah5s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,46,1,1
  PuPlayer.LabelNew pBackglass,"ah6n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,50,1,1
  PuPlayer.LabelNew pBackglass,"ah6s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,50,1,1
  PuPlayer.LabelNew pBackglass,"ah7n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,54,1,1
  PuPlayer.LabelNew pBackglass,"ah7s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,54,1,1
  PuPlayer.LabelNew pBackglass,"ah8n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,58,1,1
  PuPlayer.LabelNew pBackglass,"ah8s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,58,1,1
  PuPlayer.LabelNew pBackglass,"ah9n","AvantGarde LT Medium",       5,16777215  ,0,1,1,21,62,1,1
  PuPlayer.LabelNew pBackglass,"ah9s","AvantGarde LT Medium",       5,16777215  ,0,1,1,32,62,1,1
  PuPlayer.LabelNew pBackglass,"ah10n","AvantGarde LT Medium",      5,16777215  ,0,1,1,21,66,1,1
  PuPlayer.LabelNew pBackglass,"ah10s","AvantGarde LT Medium",      5,16777215  ,0,1,1,32,66,1,1
  PuPlayer.LabelNew pBackglass,"ah11n","AvantGarde LT Medium",      5,16777215  ,0,1,1,54,30,1,1
  PuPlayer.LabelNew pBackglass,"ah11s","AvantGarde LT Medium",      5,16777215  ,0,1,1,65,30,1,1
  PuPlayer.LabelNew pBackglass,"ah12n","AvantGarde LT Medium",      5,16777215  ,0,1,1,54,34,1,1
  PuPlayer.LabelNew pBackglass,"ah12s","AvantGarde LT Medium",      5,16777215  ,0,1,1,65,34,1,1
  PuPlayer.LabelNew pBackglass,"ah13n","AvantGarde LT Medium",      5,16777215  ,0,1,1,54,38,1,1
  PuPlayer.LabelNew pBackglass,"ah13s","AvantGarde LT Medium",      5,16777215  ,0,1,1,65,38,1,1
  PuPlayer.LabelNew pBackglass,"ah14n","AvantGarde LT Medium",      5,16777215  ,0,1,1,54,42,1,1
  PuPlayer.LabelNew pBackglass,"ah14s","AvantGarde LT Medium",      5,16777215  ,0,1,1,65,42,1,1
  PuPlayer.LabelNew pBackglass,"ah15n","AvantGarde LT Medium",      5,16777215  ,0,1,1,54,46,1,1
  PuPlayer.LabelNew pBackglass,"ah15s","AvantGarde LT Medium",      5,16777215  ,0,1,1,65,46,1,1
  PuPlayer.LabelNew pBackglass,"ah16n","AvantGarde LT Medium",      5,16777215  ,0,1,1,54,50,1,1
  PuPlayer.LabelNew pBackglass,"ah16s","AvantGarde LT Medium",      5,16777215  ,0,1,1,65,50,1,1
  PuPlayer.LabelNew pBackglass,"ah17n","AvantGarde LT Medium",      5,16777215  ,0,1,1,54,54,1,1
  PuPlayer.LabelNew pBackglass,"ah17s","AvantGarde LT Medium",      5,16777215  ,0,1,1,65,54,1,1
  PuPlayer.LabelNew pBackglass,"ah18n","AvantGarde LT Medium",      5,16777215  ,0,1,1,54,58,1,1
  PuPlayer.LabelNew pBackglass,"ah18s","AvantGarde LT Medium",      5,16777215  ,0,1,1,65,58,1,1
  PuPlayer.LabelNew pBackglass,"ah19n","AvantGarde LT Medium",      5,16777215  ,0,1,1,54,62,1,1
  PuPlayer.LabelNew pBackglass,"ah19s","AvantGarde LT Medium",      5,16777215  ,0,1,1,65,62,1,1
  PuPlayer.LabelNew pBackglass,"ah20n","AvantGarde LT Medium",      5,16777215  ,0,1,1,54,66,1,1
  PuPlayer.LabelNew pBackglass,"ah20s","AvantGarde LT Medium",      5,16777215  ,0,1,1,65,66,1,1

  Sub ruleshelperon
    rulestime.enabled = 1
  End Sub

  Sub ruleshelperoff
    rulestime.enabled = 0
  End Sub

  Dim rulesposition
  rulesposition = 0

  Sub rulestime_timer
    If turnoffrules = 1 then exit sub end if
    rulesposition = rulesposition + 1
    Select Case rulesposition
    Case 1
      PuPlayer.LabelSet pBackglass,"notetitle","Find Barb",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","Hit the B A R B lane targets to light Barb Lock.",1,""
    Case 8
      PuPlayer.LabelSet pBackglass,"notetitle","Save Will",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","Take down the keep out barricade, then bash the targets 15 times \r enter the upside and escape with Will",1,""
    Case 15
      PuPlayer.LabelSet pBackglass,"notetitle","Escape the Bad Men",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","Hit the Bad Men targets on the left and lock balls to escape the Bad Men",1,""
    Case 24
      PuPlayer.LabelSet pBackglass,"notetitle","Gather the Adventuring Party",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","Every shot on the table corresponds to a party member in the middle \r Collect them all and hit castle byers for multiball",1,""
    Case 32
      PuPlayer.LabelSet pBackglass,"notetitle","Visit the A.V. Club",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","The AV Club Is where you'll find fun modes Shoot the ramps to power up the battery \r Enter the A.V. Club to get a mode started- Finish 2 to light extra ball",1,""
    Case 40
      PuPlayer.LabelSet pBackglass,"notetitle","Collect Eggos for Mystery Roll",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","The mystery is super valuable- Collect waffles to light a mystery roll in castle byers",1,""
    Case 48
      PuPlayer.LabelSet pBackglass,"notetitle","Need Another Kickback?",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","If your kickback is empty then fill up your lane light to reactivate it!",1,""
    Case 56
      rulesposition = 0
  End Select
  End Sub


  'Page 2 (default Text Splash 1 Big Line)
  PuPlayer.LabelNew pBackglass,"Splash"  ,"avantgarde",40,77749231,0,1,1,0,0,2,0

  'Page 3 (default Text Splash 2 Lines)
  PuPlayer.LabelNew pBackglass,"Splash2a","avantgarde",40,77749231,0,1,1,0,25,3,0
  PuPlayer.LabelNew pBackglass,"Splash2b","avantgarde",40,77749231,0,1,1,0,75,3,0

  Sub resetbackglass
  Loadhs
  PuPlayer.LabelShowPage pBackglass,1,0,""
  PuPlayer.playlistplayex pBackglass,"scene","base.mov",0,1
  PuPlayer.SetBackground pBackglass,1
  PuPlayer.LabelSet pBackglass,"hstitle","HIGH SCORE",1,""
  PuPlayer.LabelSet pBackglass,"hs","" & FormatNumber(HighScore(0),0),1,""
  PuPlayer.LabelSet pBackglass,"gptitle","GAMES PLAYED",1,""
  PuPlayer.LabelSet pBackglass,"gp","" & FormatNumber(TotalGamesPlayed,0),1,""
  PuPlayer.LabelSet pBackglass,"Willh","0",1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 8.1, 'xalign': 1}"
  PuPlayer.LabelSet pBackglass,"Willj","0",1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0}"
  PuPlayer.LabelSet pBackglass,"badh","0",1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
  PuPlayer.LabelSet pBackglass,"badj","0",1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
  PuPlayer.LabelSet pBackglass,"barbh","0",1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
  PuPlayer.LabelSet pBackglass,"barbj","0",1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
  PuPlayer.LabelSet pBackglass,"partyh","0",1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
  PuPlayer.LabelSet pBackglass,"partyj","0",1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
  PuPlayer.LabelSet pBackglass,"mav","0",1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 94.6, 'yalign': 0}"
  PuPlayer.LabelSet pBackglass,"mwaf","000",1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 13.3, 'xalign': 1, 'ypos': 94.6, 'yalign': 0}"
  PuPlayer.LabelSet pBackglass,"Ball","PRESS START TO PLAY",1,""
  End Sub

  Dim titlepos
  titlepos = 0
  titletimer.enabled = 0
  dim title
  title = ""
  dim subtitle
  subtitle = ""

  Sub titletimer_timer
    titlepos = titlepos + 1
    Select Case titlepos
      Case 1
        PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':2565927, 'size': 0, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 0, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':2565927, 'size': 0, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 0, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
      Case 2
        PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':5066061, 'size': 0.6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 0.6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':5066061, 'size': 0.4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 0.4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
      Case 3
        PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':7960953, 'size': 1.2, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 1.2, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':7960953, 'size': 0.8, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 0.8, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
      Case 4
        PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':9671571, 'size': 1.8, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 1.8, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':9671571, 'size': 1.2, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 1.2, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
      Case 5
        PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':11842740, 'size': 2.4, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 2.4, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':11842740, 'size': 1.6, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 1.6, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
      Case 6
        PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':13224393, 'size': 3, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 3, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':13224393, 'size': 2, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 2, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
      Case 7
        PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':14671839, 'size': 3.6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 3.6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':14671839, 'size': 2.4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 2.4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
      Case 8
        PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':15790320, 'size': 4.2, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 4.2, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':15790320, 'size': 2.8, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 2.8, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
      Case 9
        PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':16316664, 'size': 4.8, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 4.8, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':16316664, 'size': 3.2, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 3.2, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
      Case 10
        PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':16777215, 'size': 5.4, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 5.4, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':16777215, 'size': 3.6, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 3.6, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
      Case 11
        PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':16777215, 'size': 6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
      Case 150
        PuPlayer.LabelSet pBackglass,"title","",1,"{'mt':2,'color':16777215, 'size': 6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg","",1,"{'mt':2,'color':0, 'size': 6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"title2","",1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        PuPlayer.LabelSet pBackglass,"titlebg2","",1,"{'mt':2,'color':0, 'size': 4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
        titlepos = 0
        titletimer.enabled = 0
    End Select
  End Sub


  Sub pNote(msgText,msg2text)
    title = msgText
    subtitle = msg2text
    If titlepos = 0 Then
      titletimer.enabled = 1
    Else
      titlepos = 0
      titletimer.enabled = 1
      PuPlayer.LabelSet pBackglass,"title","",1,"{'mt':2,'color':16777215, 'size': 6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
      PuPlayer.LabelSet pBackglass,"titlebg","",1,"{'mt':2,'color':0, 'size': 6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
      PuPlayer.LabelSet pBackglass,"title2","",1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
      PuPlayer.LabelSet pBackglass,"titlebg2","",1,"{'mt':2,'color':0, 'size': 4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
    End If
  End Sub




  Sub currentplayerbackglass
    PuPlayer.LabelSet pBackglass,"Willh","" & udhits(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 8.1, 'xalign': 1}"
    PuPlayer.LabelSet pBackglass,"Willj","" & WillHits(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0}"
    PuPlayer.LabelSet pBackglass,"badh","" & BallsInRunLock(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
    PuPlayer.LabelSet pBackglass,"badj","" & RunHits(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
    PuPlayer.LabelSet pBackglass,"barbh","" & BallsInLock(CurrentPlayer) ,1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
    PuPlayer.LabelSet pBackglass,"barbj","" & barbjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
    PuPlayer.LabelSet pBackglass,"partyh","" & partycollected(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
    PuPlayer.LabelSet pBackglass,"partyj","" & partyjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
    PuPlayer.LabelSet pBackglass,"mav","" & avsdone(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 94.6, 'yalign': 0}"
    PuPlayer.LabelSet pBackglass,"mwaf", bumps(CurrentPlayer),1,""
    PuPlayer.LabelSet pBackglass,"modetitle","Timer",1,"{'mt':2,'color':16777215, 'size': 0, 'xpos': 80.7, 'xalign': 1, 'ypos': 72.6, 'yalign': 0}"
    PuPlayer.LabelSet pBackglass,"modetimer","",1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 80.7, 'xalign': 1, 'ypos': 75, 'yalign': 0}"
  End Sub


  Sub pUpdateScores
    If CurrentPlayer = 1 Then
      PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(CurrentPlayer),0),1,"{'mt':2,'color':16777215, 'size': 2.4 }"
      PuPlayer.LabelSet pBackglass,"Play1","PLAYER 1",1,"{'mt':2,'color':16777215 }"
      'make other scores red (inactive)
      If PlayersPlayingGame = 2 Then
        PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(2),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play2","PLAYER 2",1,"{'mt':2,'color':3875550}"
      End If
      If PlayersPlayingGame = 3 Then
        PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(2),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play2","PLAYER 2",1,"{'mt':2,'color':3875550}"
        PuPlayer.LabelSet pBackglass,"Play3score","" & FormatNumber(Score(3),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play3","PLAYER 3",1,"{'mt':2,'color':3875550}"
      End If
      If PlayersPlayingGame = 4 Then
        PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(2),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play2","PLAYER 2",1,"{'mt':2,'color':3875550}"
        PuPlayer.LabelSet pBackglass,"Play3score","" & FormatNumber(Score(3),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play3","PLAYER 3",1,"{'mt':2,'color':3875550}"
        PuPlayer.LabelSet pBackglass,"Play4score","" & FormatNumber(Score(4),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play4","PLAYER 4",1,"{'mt':2,'color':3875550}"
      End If
    End If
    If CurrentPlayer = 2 Then
      PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(CurrentPlayer),0),1,"{'mt':2,'color':16777215, 'size': 2.4 }"
      PuPlayer.LabelSet pBackglass,"Play2","PLAYER 2",1,"{'mt':2,'color':16777215 }"
      If PlayersPlayingGame = 2 Then
        PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(1),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play1","PLAYER 1",1,"{'mt':2,'color':3875550}"
      End If
      If PlayersPlayingGame = 3 Then
        PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(1),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play1","PLAYER 1",1,"{'mt':2,'color':3875550}"
        PuPlayer.LabelSet pBackglass,"Play3score","" & FormatNumber(Score(3),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play3","PLAYER 3",1,"{'mt':2,'color':3875550}"
      End If
      If PlayersPlayingGame = 4 Then
        PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(1),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play1","PLAYER 1",1,"{'mt':2,'color':3875550}"
        PuPlayer.LabelSet pBackglass,"Play3score","" & FormatNumber(Score(3),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play3","PLAYER 3",1,"{'mt':2,'color':3875550}"
        PuPlayer.LabelSet pBackglass,"Play4score","" & FormatNumber(Score(4),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play4","PLAYER 4",1,"{'mt':2,'color':3875550}"
      End If
    End If
    If CurrentPlayer = 3 Then
      PuPlayer.LabelSet pBackglass,"Play3score","" & FormatNumber(Score(CurrentPlayer),0),1,"{'mt':2,'color':16777215, 'size': 2.4 }"
      PuPlayer.LabelSet pBackglass,"Play3","PLAYER 3",1,"{'mt':2,'color':16777215 }"
      If PlayersPlayingGame = 3 Then
        PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(1),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play1","PLAYER 1",1,"{'mt':2,'color':3875550}"
        PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(2),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play2","PLAYER 2",1,"{'mt':2,'color':3875550}"
      End If
      If PlayersPlayingGame = 4 Then
        PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(1),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play1","PLAYER 1",1,"{'mt':2,'color':3875550}"
        PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(2),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play2","PLAYER 2",1,"{'mt':2,'color':3875550}"
        PuPlayer.LabelSet pBackglass,"Play4score","" & FormatNumber(Score(4),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play4","PLAYER 4",1,"{'mt':2,'color':3875550}"
      End If
    End If
    If CurrentPlayer = 4 Then
      PuPlayer.LabelSet pBackglass,"Play4score","" & FormatNumber(Score(CurrentPlayer),0),1,"{'mt':2,'color':16777215, 'size': 2.4 }"
      PuPlayer.LabelSet pBackglass,"Play4","PLAYER 4",1,"{'mt':2,'color':16777215 }"
      If PlayersPlayingGame = 4 Then
        PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(1),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play1","PLAYER 1",1,"{'mt':2,'color':3875550}"
        PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(2),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play2","PLAYER 2",1,"{'mt':2,'color':3875550}"
        PuPlayer.LabelSet pBackglass,"Play3score","" & FormatNumber(Score(3),0),1,"{'mt':2,'color':3875550, 'size': 2.4 }"
        PuPlayer.LabelSet pBackglass,"Play3","PLAYER 3",1,"{'mt':2,'color':3875550}"
      End If
    End If
  PuPlayer.LabelSet pBackglass,"Ball","Ball "  &  bpgcurrent - BallsRemaining(CurrentPlayer) + 1 & "/" & bpgcurrent,1,""
  end Sub


  Sub pDMDLabelHide(labName)
  PuPlayer.LabelSet pBackglass,labName,"",0,""
  end sub


  Sub pDMDSplashBig(msgText, msg2text, timeSec)
  PuPlayer.LabelShowPage pBackglass,2,timeSec,""
  PuPlayer.LabelSet pBackglass,"Splash",msgText,0,"{'mt':1,'at':1,'fq':250,'len':" & (timeSec*1000) & ",'fc':" & mColor & "}"
  end sub


  Sub pDMDScrollBig(msgText,timeSec,mColor)
  PuPlayer.LabelShowPage pBackglass,2,timeSec,""
  PuPlayer.LabelSet pBackglass,"Splash",msgText,0,"{'mt':1,'at':2,'xps':1,'xpe':-1,'len':" & (timeSec*1000) & ",'mlen':" & (timeSec*1000) & ",'tt':0,'fc':" & mColor & "}"
  end sub

  Sub pDMDScrollBigV(msgText,timeSec,mColor)
  PuPlayer.LabelShowPage pBackglass,2,timeSec,""
  PuPlayer.LabelSet pBackglass,"Splash",msgText,0,"{'mt':1,'at':2,'yps':1,'ype':-1,'len':" & (timeSec*1000) & ",'mlen':" & (timeSec*1000) & ",'tt':0,'fc':" & mColor & "}"
  end sub

  Sub pDMDSplashLines(msgText,msgText2,timeSec,mColor)
  PuPlayer.LabelShowPage pBackglass,3,timeSec,""
  PuPlayer.LabelSet pBackglass,"Splash2a",msgText,0,"{'mt':1,'at':1,'fq':250,'len':" & (timeSec*1000) & ",'fc':" & mColor & "}"
  PuPlayer.LabelSet pBackglass,"Splash2b",msgText2,0,"{'mt':1,'at':1,'fq':250,'len':" & (timeSec*1000) & ",'fc':" & mColor & "}"
  end Sub


  Sub pDMDSplashScore(msgText,timeSec,mColor)
  PuPlayer.LabelSet pBackglass,"MsgScore",msgText,0,"{'mt':1,'at':1,'fq':250,'len':"& (timeSec*1000) &",'fc':" & mColor & "}"
  end Sub

  Sub pDMDSplashScoreScroll(msgText,timeSec,mColor)
  PuPlayer.LabelSet pBackglass,"MsgScore",msgText,0,"{'mt':1,'at':2,'xps':1,'xpe':-1,'len':"& (timeSec*1000) &", 'mlen':"& (timeSec*1000) &",'tt':0, 'fc':" & mColor & "}"
  end Sub





'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  High Scores
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X

  ' load em up


  Dim hschecker:hschecker = 0

  Sub Loadhs
    Dim x
    x = LoadValue(TableName, "HighScore1")
    If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 1300000 End If

    x = LoadValue(TableName, "HighScore1Name")
    If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "SBW" End If

    x = LoadValue(TableName, "HighScore2")
    If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 1200000 End If

    x = LoadValue(TableName, "HighScore2Name")
    If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "JPS" End If

    x = LoadValue(TableName, "HighScore3")
    If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 1100000 End If

    x = LoadValue(TableName, "HighScore3Name")
    If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "011" End If

    x = LoadValue(TableName, "HighScore4")
    If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 1000000 End If

    x = LoadValue(TableName, "HighScore4Name")
    If(x <> "") then HighScoreName(3) = x Else HighScoreName(3) = "011" End If

    x = LoadValue(TableName, "WaffleScore1")
    If(x <> "") Then WaffleScore(0) = CDbl(x) Else WaffleScore(0) = 100 End If

    x = LoadValue(TableName, "WaffleScore1Name")
    If(x <> "") Then WaffleScoreName(0) = x Else WaffleScoreName(0) = "SBW" End If

    x = LoadValue(TableName, "WaffleScore2")
    If(x <> "") then WaffleScore(1) = CDbl(x) Else WaffleScore(1) = 99 End If

    x = LoadValue(TableName, "WaffleScore2Name")
    If(x <> "") then WaffleScoreName(1) = x Else WaffleScoreName(1) = "011" End If

    x = LoadValue(TableName, "WaffleScore3")
    If(x <> "") then WaffleScore(2) = CDbl(x) Else WaffleScore(2) = 98 End If

    x = LoadValue(TableName, "WaffleScore3Name")
    If(x <> "") then WaffleScoreName(2) = x Else WaffleScoreName(2) = "DSN" End If

    x = LoadValue(TableName, "WaffleScore4")
    If(x <> "") then WaffleScore(3) = CDbl(x) Else WaffleScore(3) = 97 End If

    x = LoadValue(TableName, "WaffleScore4Name")
    If(x <> "") then WaffleScoreName(3) = x Else WaffleScoreName(3) = "MIK" End If

    x = LoadValue(TableName, "Credits")
    If(x <> "") then Credits = CInt(x) Else Credits = 0 End If

    x = LoadValue(TableName, "TotalGamesPlayed")
    If(x <> "") then TotalGamesPlayed = CInt(x) Else TotalGamesPlayed = 0 End If

    If hschecker = 0 Then
    checkorder
    checkwaforder
    End If
  End Sub

  Dim hs3,hs2,hs1,hs0,hsn3,hsn2,hsn1,hsn0
  Dim wf3,wf2,wf1,wf0,wfn3,wfn2,wfn1,wfn0

  Sub checkorder
    hschecker = 1
    hs3 = HighScore(3)
    hs2 = HighScore(2)
    hs1 = HighScore(1)
    hs0 = HighScore(0)
    hsn3 = HighScoreName(3)
    hsn2 = HighScoreName(2)
    hsn1 = HighScoreName(1)
    hsn0 = HighScoreName(0)
    If hs3 > hs0 Then
      HighScore(0) = hs3
      HighScoreName(0) = hsn3
      HighScore(1) = hs0
      HighScoreName(1) = hsn0
      HighScore(2) = hs1
      HighScoreName(2) = hsn1
      HighScore(3) = hs2
      HighScoreName(3) = hsn2

    ElseIf hs3 > hs1 Then
      HighScore(0) = hs0
      HighScoreName(0) = hsn0
      HighScore(1) = hs3
      HighScoreName(1) = hsn3
      HighScore(2) = hs1
      HighScoreName(2) = hsn1
      HighScore(3) = hs2
      HighScoreName(3) = hsn2
    ElseIf hs3 > hs2 Then
      HighScore(0) = hs0
      HighScoreName(0) = hsn0
      HighScore(1) = hs1
      HighScoreName(1) = hsn1
      HighScore(2) = hs3
      HighScoreName(2) = hsn3
      HighScore(3) = hs2
      HighScoreName(3) = hsn2
    ElseIf hs3 < hs2 Then
      HighScore(0) = hs0
      HighScoreName(0) = hsn0
      HighScore(1) = hs1
      HighScoreName(1) = hsn1
      HighScore(2) = hs2
      HighScoreName(2) = hsn2
      HighScore(3) = hs3
      HighScoreName(3) = hsn3
    End If

    savehs
  End Sub


  sub checkwaforder
    wf3 = WaffleScore(3)
    wf2 = waffleScore(2)
    wf1 = WaffleScore(1)
    wf0 = waffleScore(0)
    wfn3 = waffleScoreName(3)
    wfn2 = waffleScoreName(2)
    wfn1 = waffleScoreName(1)
    wfn0 = waffleScoreName(0)
    If wf3 > wf0 Then
      WaffleScore(0) = wf3
      waffleScoreName(0) = wfn3
      waffleScore(1) = wf0
      waffleScoreName(1) = wfn0
      waffleScore(2) = wf1
      waffleScoreName(2) = wfn1
      waffleScore(3) = wf2
      waffleScoreName(3) = wfn2
    ElseIf wf3 > wf1 Then
      WaffleScore(0) = wf0
      WaffleScoreName(0) = wfn0
      WaffleScore(1) = wf3
      WaffleScoreName(1) = wfn3
      WaffleScore(2) = wf1
      WaffleScoreName(2) = wfn1
      WaffleScore(3) = wf2
      WaffleScoreName(3) = wfn2
    ElseIf wf3 > wf2 Then
      WaffleScore(0) = wf0
      WaffleScoreName(0) = wfn0
      WaffleScore(1) = wf1
      WaffleScoreName(1) = wfn1
      WaffleScore(2) = wf3
      WaffleScoreName(2) = wfn3
      WaffleScore(3) = wf2
      WaffleScoreName(3) = wfn2
    ElseIf wf3 < wf2 Then
      WaffleScore(0) = wf0
      WaffleScoreName(0) = wfn0
      WaffleScore(1) = wf1
      WaffleScoreName(1) = wfn1
      WaffleScore(2) = wf2
      WaffleScoreName(2) = wfn2
      WaffleScore(3) = wf3
      WaffleScoreName(3) = wfn3
    End If
    Savewafs
  End Sub


  Sub Savehs
    SaveValue TableName, "HighScore1", HighScore(0)
    SaveValue TableName, "HighScore1Name", HighScoreName(0)
    SaveValue TableName, "HighScore2", HighScore(1)
    SaveValue TableName, "HighScore2Name", HighScoreName(1)
    SaveValue TableName, "HighScore3", HighScore(2)
    SaveValue TableName, "HighScore3Name", HighScoreName(2)
    SaveValue TableName, "HighScore4", HighScore(3)
    SaveValue TableName, "HighScore4Name", HighScoreName(3)
  End Sub


  Sub Savewafs
    SaveValue TableName, "WaffleScore1", WaffleScore(0)
    SaveValue TableName, "WaffleScore1Name", WaffleScoreName(0)
    SaveValue TableName, "WaffleScore2", WaffleScore(1)
    SaveValue TableName, "WaffleScore2Name", WaffleScoreName(1)
    SaveValue TableName, "WaffleScore3", WaffleScore(2)
    SaveValue TableName, "WaffleScore3Name", WaffleScoreName(2)
    SaveValue TableName, "WaffleScore4", WaffleScore(3)
    SaveValue TableName, "WaffleScore4Name", WaffleScoreName(3)
  End Sub



  Sub Savegp
    SaveValue TableName, "TotalGamesPlayed", TotalGamesPlayed
    vpmtimer.addtimer 1000, "Loadhs'"
  End Sub


  ' Initials

  Dim hsbModeActive:hsbModeActive = False
  Dim hsEnteredName
  Dim hsEnteredDigits(3)
  Dim hsCurrentDigit
  Dim hsValidLetters
  Dim hsCurrentLetter
  Dim hsLetterFlash
  Dim wafEnteredName
  Dim wafEnteredDigits(3)
  Dim wafCurrentDigit
  Dim wafValidLetters
  Dim wafCurrentLetter
  Dim wafLetterFlash

  ' Check the scores to see if you got one

  Sub CheckHighscore()
    Dim tmp
    tmp = Score(CurrentPlayer)
    osbtempscore = Score(CurrentPlayer)
    If tmp > HighScore(3) Then
      AwardSpecial
      vpmtimer.addtimer 2000, "PlaySound ""vo_contratulationsgreatscore"" '"
      HighScore(3) = tmp
      'enter player's name
      HighScoreEntryInit()
      DOF 403, DOFPulse   'DOF MX - Hi Score
    Else
      osbtemp = osbdefinit
      if osbkey="" Then
      Else
      SubmitOSBScore
      end If
      EndOfBallComplete()
      'checkwaffles()
    End If
  End Sub





  Sub HighScoreEntryInit()
    hsbModeActive = True
    inhighscore = True
    PlaySound "vo_enteryourinitials"

    hsEnteredDigits(1) = "A"
    hsEnteredDigits(2) = " "
    hsEnteredDigits(3) = " "

    hsCurrentDigit = 1

    pNote "YOU GOT","A HIGH SCORE!"
    PuPlayer.playlistplayex pCallouts,"audiocallouts","yougotahighscore.wav",100,1
    chilloutthemusic
    PuPlayer.playlistplayex pBackglass,"videohighscore","hs1.mov",100,7
    PuPlayer.SetLoop 2,1
    Playsound "bellhs"
    HighScoreDisplayName()
    HighScorelabels
  End Sub

  ' flipper moving around the letters

  Sub EnterHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
      Playsound "fx_Previous"
      If wafflemode = false Then
        If hsletter = 0 Then
          hsletter = 27
        Else
          hsLetter = hsLetter - 1
        End If
        HighScoreDisplayName()
      Else
        If wafletter = 0 Then
          wafletter = 27
        Else
          wafLetter = wafLetter - 1
        End If
        WaffleScoreDisplayName()
      End If
    End If

    If keycode = RightFlipperKey Then
      Playsound "fx_Next"
      If wafflemode = false Then
        If hsletter = 27 Then
          hsletter = 0
        Else
          hsLetter = hsLetter + 1
        End If
        HighScoreDisplayName()
      Else
        If wafletter = 27 Then
          wafletter = 0
        Else
          wafLetter = wafLetter + 1
        End If
        WaffleScoreDisplayName()
      End If
    End If

    If keycode = StartGameKey or keycode = PlungerKey Then
      PlaySound "success"
      If wafflemode = false Then
        If hsCurrentDigit = 3 Then
          If hsletter = 0 Then
            hsCurrentDigit = hsCurrentDigit -1
          Else
            assignletter
            vpmtimer.addtimer 700, "HighScoreCommitName()'"
            PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,7
          End If
        End If
      Else
        If wafCurrentDigit = 3 Then
          If wafletter = 0 Then
            wafCurrentDigit = wafCurrentDigit -1
          Else
            wafassignletter
            vpmtimer.addtimer 700, "WaffleScoreCommitName()'"
            PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,7
            wafflemode = False
          End If
        End If
      End If


      If wafflemode = false Then
        If hsCurrentDigit < 3 Then
          If hsletter = 0 Then
            If hsCurrentDigit = 1 Then
            Else
              hsCurrentDigit = hsCurrentDigit -1
            End If
          Else
            assignletter
            hsCurrentDigit = hsCurrentDigit + 1
            HighScoreDisplayName()

          End If
        End If
      Else
        If wafCurrentDigit < 3 Then
          If wafletter = 0 Then
            If wafCurrentDigit = 1 Then
            Else
              wafCurrentDigit = wafCurrentDigit -1
            End If
          Else
            wafassignletter
            wafCurrentDigit = wafCurrentDigit + 1
            WaffleScoreDisplayName()

          End If
        End If
      End If
    End if
  End Sub

  Dim hsletter
  hsletter = 1

  dim hsdigit:hsdigit = 1

  Sub assignletter
    if hscurrentdigit = 1 Then
      hsdigit = 1
    End If
    if hscurrentdigit = 2 Then
      hsdigit = 2
    End If
    if hscurrentdigit = 3 Then
      hsdigit = 3
    End If
    If hsletter = 1 Then
      hsEnteredDigits(hsdigit) = "A"
    End If
    If hsletter = 2 Then
      hsEnteredDigits(hsdigit) = "B"
    End If
    If hsletter = 3 Then
      hsEnteredDigits(hsdigit) = "C"
    End If
    If hsletter = 4 Then
      hsEnteredDigits(hsdigit) = "D"
    End If
    If hsletter = 5 Then
      hsEnteredDigits(hsdigit) = "E"
    End If
    If hsletter = 6 Then
      hsEnteredDigits(hsdigit) = "F"
    End If
    If hsletter = 7 Then
      hsEnteredDigits(hsdigit) = "G"
    End If
    If hsletter = 8 Then
      hsEnteredDigits(hsdigit) = "H"
    End If
    If hsletter = 9 Then
      hsEnteredDigits(hsdigit) = "I"
    End If
    If hsletter = 10 Then
      hsEnteredDigits(hsdigit) = "J"
    End If
    If hsletter = 11 Then
      hsEnteredDigits(hsdigit) = "K"
    End If
    If hsletter = 12 Then
      hsEnteredDigits(hsdigit) = "L"
    End If
    If hsletter = 13 Then
      hsEnteredDigits(hsdigit) = "M"
    End If
    If hsletter = 14 Then
      hsEnteredDigits(hsdigit) = "N"
    End If
    If hsletter = 15 Then
      hsEnteredDigits(hsdigit) = "O"
    End If
    If hsletter = 16 Then
      hsEnteredDigits(hsdigit) = "P"
    End If
    If hsletter = 17 Then
      hsEnteredDigits(hsdigit) = "Q"
    End If
    If hsletter = 18 Then
      hsEnteredDigits(hsdigit) = "R"
    End If
    If hsletter = 19 Then
      hsEnteredDigits(hsdigit) = "S"
    End If
    If hsletter = 20 Then
      hsEnteredDigits(hsdigit) = "T"
    End If
    If hsletter = 21 Then
      hsEnteredDigits(hsdigit) = "U"
    End If
    If hsletter = 22 Then
      hsEnteredDigits(hsdigit) = "V"
    End If
    If hsletter = 23 Then
      hsEnteredDigits(hsdigit) = "W"
    End If
    If hsletter = 24 Then
      hsEnteredDigits(hsdigit) = "X"
    End If
    If hsletter = 25 Then
      hsEnteredDigits(hsdigit) = "Y"
    End If
    If hsletter = 26 Then
      hsEnteredDigits(hsdigit) = "Z"
    End If
    If hsletter = 27 Then
      hsEnteredDigits(hsdigit) = " "
    End If

  End Sub

  Sub HighScorelabels
    PuPlayer.LabelSet pBackglass,"HighScore","YOU GOT A\rHIGH SCORE!",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL1","A",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL2"," ",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL3"," ",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL4",Score(CurrentPlayer),1,""
    hsletter = 1
  End Sub

  Sub HighScoreDisplayName()

    Select case hsLetter
    Case 0
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","<",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","<",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","<",1,""
    Case 1
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","A",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","A",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","A",1,""
    Case 2
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","B",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","B",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","B",1,""
    Case 3
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","C",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","C",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","C",1,""
    Case 4
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","D",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","D",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","D",1,""
    Case 5
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","E",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","E",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","E",1,""
    Case 6
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","F",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","F",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","F",1,""
    Case 7
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","G",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","G",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","G",1,""
    Case 8
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","H",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","H",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","H",1,""
    Case 9
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","I",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","I",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","I",1,""
    Case 10
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","J",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","J",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","J",1,""
    Case 11
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","K",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","K",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","K",1,""
    Case 12
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","L",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","L",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","L",1,""
    Case 13
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","M",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","M",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","M",1,""
    Case 14
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","N",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","N",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","N",1,""
    Case 15
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","O",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","O",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","O",1,""
    Case 16
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","P",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","P",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","P",1,""
    Case 17
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","Q",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","Q",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","Q",1,""
    Case 18
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","R",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","R",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","R",1,""
    Case 19
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","S",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","S",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","S",1,""
    Case 20
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","T",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","T",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","T",1,""
    Case 21
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","U",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","U",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","U",1,""
    Case 22
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","V",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","V",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","V",1,""
    Case 23
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","W",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","W",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","W",1,""
    Case 24
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","X",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","X",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","X",1,""
    Case 25
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","Y",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","Y",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","Y",1,""
    Case 26
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","Z",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","Z",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","Z",1,""
    Case 27
      if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1"," ",1,""
      if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2"," ",1,""
      if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3"," ",1,""
    End Select
  End Sub

  ' post the high score letters



  Sub HighScoreCommitName()
    PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,7
    hsEnteredName = hsEnteredDigits(1) & hsEnteredDigits(2) & hsEnteredDigits(3)
    HighScoreName(3) = hsEnteredName
    checkorder
    osbtemp = hsEnteredName
    if osbkey="" Then
    Else
    SubmitOSBScore
    end If
    inhighscore = False
    checkwaffles()
  End Sub



  Sub CheckWaffles()
    Dim tmp2
    tmp2 = bumps(CurrentPlayer)

    If tmp2 > WaffleScore(3) Then
      wafflemode = true
      vpmtimer.addtimer 2000, "PlaySound ""vo_contratulationsgreatscore"" '"
      WaffleScore(3) = tmp2
      'enter player's name
      WaffleScoreEntryInit()
      DOF 403, DOFPulse   'DOF MX - Hi Score
    Else
      EndOfBallComplete()
      hsbModeActive = False
    End If
  End Sub

  Dim wafflemode:wafflemode = False

  Sub WaffleScoreEntryInit()
    inhighscore = True
    hsbModeActive = True

    PlaySound "vo_enteryourinitials"

    wafEnteredDigits(1) = "A"
    wafEnteredDigits(2) = " "
    wafEnteredDigits(3) = " "

    wafCurrentDigit = 1

    pNote "YOU ARE AN","EGGO CHAMPION"
    PuPlayer.playlistplayex pCallouts,"audiocallouts","youreaneggochampion.wav",100,1
    chilloutthemusic
    PuPlayer.playlistplayex pBackglass,"videohighscore","hs1.mov",100,7
    PuPlayer.SetLoop 2,1
    Playsound "bellhs"

    WaffleScoreDisplayName()
    WaffleScorelabels
  End Sub



  Dim wafletter
  wafletter = 1

  dim wafdigit:wafdigit = 1

  Sub wafassignletter
    if wafcurrentdigit = 1 Then
      wafdigit = 1
    End If
    if wafcurrentdigit = 2 Then
      wafdigit = 2
    End If
    if wafcurrentdigit = 3 Then
      wafdigit = 3
    End If
    If wafletter = 1 Then
      wafEnteredDigits(wafdigit) = "A"
    End If
    If wafletter = 2 Then
      wafEnteredDigits(wafdigit) = "B"
    End If
    If wafletter = 3 Then
      wafEnteredDigits(wafdigit) = "C"
    End If
    If wafletter = 4 Then
      wafEnteredDigits(wafdigit) = "D"
    End If
    If wafletter = 5 Then
      wafEnteredDigits(wafdigit) = "E"
    End If
    If wafletter = 6 Then
      wafEnteredDigits(wafdigit) = "F"
    End If
    If wafletter = 7 Then
      wafEnteredDigits(wafdigit) = "G"
    End If
    If wafletter = 8 Then
      wafEnteredDigits(wafdigit) = "H"
    End If
    If wafletter = 9 Then
      wafEnteredDigits(wafdigit) = "I"
    End If
    If wafletter = 10 Then
      wafEnteredDigits(wafdigit) = "J"
    End If
    If wafletter = 11 Then
      wafEnteredDigits(wafdigit) = "K"
    End If
    If wafletter = 12 Then
      wafEnteredDigits(wafdigit) = "L"
    End If
    If wafletter = 13 Then
      wafEnteredDigits(wafdigit) = "M"
    End If
    If wafletter = 14 Then
      wafEnteredDigits(wafdigit) = "N"
    End If
    If wafletter = 15 Then
      wafEnteredDigits(wafdigit) = "O"
    End If
    If wafletter = 16 Then
      wafEnteredDigits(wafdigit) = "P"
    End If
    If wafletter = 17 Then
      wafEnteredDigits(wafdigit) = "Q"
    End If
    If wafletter = 18 Then
      wafEnteredDigits(wafdigit) = "R"
    End If
    If wafletter = 19 Then
      wafEnteredDigits(wafdigit) = "S"
    End If
    If wafletter = 20 Then
      wafEnteredDigits(wafdigit) = "T"
    End If
    If wafletter = 21 Then
      wafEnteredDigits(wafdigit) = "U"
    End If
    If wafletter = 22 Then
      wafEnteredDigits(wafdigit) = "V"
    End If
    If wafletter = 23 Then
      wafEnteredDigits(wafdigit) = "W"
    End If
    If wafletter = 24 Then
      wafEnteredDigits(wafdigit) = "X"
    End If
    If wafletter = 25 Then
      wafEnteredDigits(wafdigit) = "Y"
    End If
    If wafletter = 26 Then
      wafEnteredDigits(wafdigit) = "Z"
    End If
    If wafletter = 27 Then
      wafEnteredDigits(wafdigit) = " "
    End If

  End Sub

  Sub WaffleScorelabels
    PuPlayer.LabelSet pBackglass,"HighScore","YOU ARE AN\rEGGO CHAMPION",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL1","A",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL2"," ",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL3"," ",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL4",bumps(CurrentPlayer) & "\rEGGOS COLLECTED",1,""
    wafletter = 1
  End Sub

  Sub WaffleScoreDisplayName()

    Select case wafLetter
    Case 0
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","<",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","<",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","<",1,""
    Case 1
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","A",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","A",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","A",1,""
    Case 2
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","B",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","B",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","B",1,""
    Case 3
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","C",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","C",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","C",1,""
    Case 4
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","D",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","D",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","D",1,""
    Case 5
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","E",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","E",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","E",1,""
    Case 6
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","F",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","F",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","F",1,""
    Case 7
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","G",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","G",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","G",1,""
    Case 8
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","H",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","H",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","H",1,""
    Case 9
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","I",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","I",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","I",1,""
    Case 10
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","J",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","J",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","J",1,""
    Case 11
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","K",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","K",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","K",1,""
    Case 12
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","L",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","L",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","L",1,""
    Case 13
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","M",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","M",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","M",1,""
    Case 14
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","N",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","N",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","N",1,""
    Case 15
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","O",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","O",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","O",1,""
    Case 16
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","P",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","P",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","P",1,""
    Case 17
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","Q",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","Q",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","Q",1,""
    Case 18
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","R",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","R",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","R",1,""
    Case 19
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","S",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","S",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","S",1,""
    Case 20
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","T",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","T",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","T",1,""
    Case 21
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","U",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","U",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","U",1,""
    Case 22
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","V",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","V",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","V",1,""
    Case 23
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","W",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","W",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","W",1,""
    Case 24
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","X",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","X",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","X",1,""
    Case 25
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","Y",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","Y",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","Y",1,""
    Case 26
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","Z",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","Z",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","Z",1,""
    Case 27
      if(wafCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1"," ",1,""
      if(wafCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2"," ",1,""
      if(wafCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3"," ",1,""
    End Select
  End Sub


  Sub WaffleScoreCommitName()
    inhighscore = False
    PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,7
    hsbModeActive = False
    wafEnteredName = wafEnteredDigits(1) & wafEnteredDigits(2) & wafEnteredDigits(3)
    WaffleScoreName(3) = wafEnteredName
    checkwaforder
    EndOfBallComplete()
    PuPlayer.LabelSet pBackglass,"HighScore","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL1","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL2"," ",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL3"," ",1,""
  End Sub




'****************************************
' Real Time updates using the GameTimer
'****************************************
'used for all the real time updates

Sub GameTimer_Timer
    RollingUpdate
' add any other real time update subs, like gates or diverters
End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  ATTRACT MODE
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'


  Dim introposition
  introposition = 0

  Sub DMDintroloop
    PuPlayer.LabelSet pBackglass,"modetitle","",1,"{'mt':2,'color':16777215, 'size': 0, 'xpos': 80.7, 'xalign': 1, 'ypos': 72.6, 'yalign': 0}"
    introtime = 0
    introposition = introposition + 1
    Select Case introposition
    Case 1
    PuPlayer.LabelSet pBackglass,"HighScore","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL1","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL2","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL3","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL4","",1,""
    PuPlayer.LabelSet pBackglass,"modetimer","",1,""
      PuPlayer.playlistplayex pMusic,"audioclear","clear.mp3",100, 1
      PuPlayer.playlistplayex pBackglass,"videoattract","intro-smaller.mov",100,1
      PuPlayer.LabelSet pBackglass,"notetitle","Need Rules Help?",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","Hold both flippers for 2 seconds while in attract mode for rules overlay \r Hold them again (or start a game) to remove it",1,""
      PuPlayer.LabelSet pBackglass,"high1name","",1,""
      PuPlayer.LabelSet pBackglass,"high1score","",1,""
      PuPlayer.LabelSet pBackglass,"high2name","",1,""
      PuPlayer.LabelSet pBackglass,"high2score","",1,""
      PuPlayer.LabelSet pBackglass,"high3name","",1,""
      PuPlayer.LabelSet pBackglass,"high3score","",1,""
      PuPlayer.LabelSet pBackglass,"high4name","",1,""
      PuPlayer.LabelSet pBackglass,"high4score","",1,""
      PuPlayer.LabelSet pBackglass,"waf1name","",1,""
      PuPlayer.LabelSet pBackglass,"waf1score","",1,""
      PuPlayer.LabelSet pBackglass,"waf2name","",1,""
      PuPlayer.LabelSet pBackglass,"waf2score","",1,""
      PuPlayer.LabelSet pBackglass,"waf3name","",1,""
      PuPlayer.LabelSet pBackglass,"waf3score","",1,""
      PuPlayer.LabelSet pBackglass,"waf4name","",1,""
      PuPlayer.LabelSet pBackglass,"waf4score","",1,""
    Case 2
    loadhs
      PuPlayer.playlistplayex pMusic,"audiobgrock","",soundtrackvol,1
    PuPlayer.LabelSet pBackglass,"HighScore","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL1","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL2","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL3","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL4","",1,""
    PuPlayer.LabelSet pBackglass,"modetimer","",1,""
      PuPlayer.LabelSet pBackglass,"notetitle","Got A killer good High Score?",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","Post it to #stlevpx on instagram to brag",1,""
      PuPlayer.playlistplayex pBackglass,"videoattract","hs3-sm.mov",100,1
      PuPlayer.LabelSet pBackglass,"high1name",HighScoreName(0),1,""
      PuPlayer.LabelSet pBackglass,"high1score",FormatNumber(HighScore(0),0),1,""
      PuPlayer.LabelSet pBackglass,"high2name",HighScoreName(1),1,""
      PuPlayer.LabelSet pBackglass,"high2score",FormatNumber(HighScore(1),0),1,""
      PuPlayer.LabelSet pBackglass,"high3name",HighScoreName(2),1,""
      PuPlayer.LabelSet pBackglass,"high3score",FormatNumber(HighScore(2),0),1,""
      PuPlayer.LabelSet pBackglass,"high4name",HighScoreName(3),1,""
      PuPlayer.LabelSet pBackglass,"high4score",FormatNumber(HighScore(3),0),1,""
      PuPlayer.LabelSet pBackglass,"waf1name",WaffleScoreName(0),1,""
      PuPlayer.LabelSet pBackglass,"waf1score",FormatNumber(WaffleScore(0),0),1,""
      PuPlayer.LabelSet pBackglass,"waf2name",WaffleScoreName(1),1,""
      PuPlayer.LabelSet pBackglass,"waf2score",FormatNumber(WaffleScore(1),0),1,""
      PuPlayer.LabelSet pBackglass,"waf3name",WaffleScoreName(2),1,""
      PuPlayer.LabelSet pBackglass,"waf3score",FormatNumber(WaffleScore(2),0),1,""
      PuPlayer.LabelSet pBackglass,"waf4name",WaffleScoreName(3),1,""
      PuPlayer.LabelSet pBackglass,"waf4score",FormatNumber(WaffleScore(3),0),1,""

    Case 3
      PuPlayer.playlistplayex pBackglass,"videoattract","thanks-sm.mov",100,1
      PuPlayer.LabelSet pBackglass,"notetitle","Did You Know?",1,""
    PuPlayer.LabelSet pBackglass,"HighScore","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL1","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL2","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL3","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL4","",1,""
    PuPlayer.LabelSet pBackglass,"modetimer","",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","You can change the music to 80s jams in the script? Check it out",1,""
      PuPlayer.LabelSet pBackglass,"high1name","",1,""
      PuPlayer.LabelSet pBackglass,"high1score","",1,""
      PuPlayer.LabelSet pBackglass,"high2name","",1,""
      PuPlayer.LabelSet pBackglass,"high2score","",1,""
      PuPlayer.LabelSet pBackglass,"high3name","",1,""
      PuPlayer.LabelSet pBackglass,"high3score","",1,""
      PuPlayer.LabelSet pBackglass,"high4name","",1,""
      PuPlayer.LabelSet pBackglass,"high4score","",1,""
      PuPlayer.LabelSet pBackglass,"waf1name","",1,""
      PuPlayer.LabelSet pBackglass,"waf1score","",1,""
      PuPlayer.LabelSet pBackglass,"waf2name","",1,""
      PuPlayer.LabelSet pBackglass,"waf2score","",1,""
      PuPlayer.LabelSet pBackglass,"waf3name","",1,""
      PuPlayer.LabelSet pBackglass,"waf3score","",1,""
      PuPlayer.LabelSet pBackglass,"waf4name","",1,""
      PuPlayer.LabelSet pBackglass,"waf4score","",1,""
    Case 4
      if osbkey="" then
        clearhslabels
        clearosblabels
        DMDintroloop
      Else
      clearhslabels
      clearosblabels
      PuPlayer.playlistplayex pBackglass,"videoattract","osb-d10-w10.mp4",100,1
      PuPlayer.LabelSet pBackglass,"notetitle","Here's the Daily and Weekly OSB Scores",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","Keep playing and make your way to the top!",1,""
      PuPlayer.LabelSet pBackglass,"dh1n","1. " & dailyvar(2),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh1s",FormatNumber(dailyvar(3),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh2n","2. " & dailyvar(4),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh2s",FormatNumber(dailyvar(5),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh3n","3. " & dailyvar(6),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh3s",FormatNumber(dailyvar(7),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh4n","4. " & dailyvar(8),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh4s",FormatNumber(dailyvar(9),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh5n","5. " & dailyvar(10),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh5s",FormatNumber(dailyvar(11),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh6n","6. " & dailyvar(12),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh6s",FormatNumber(dailyvar(13),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh7n","7. " & dailyvar(14),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh7s",FormatNumber(dailyvar(15),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh8n","8. " & dailyvar(16),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh8s",FormatNumber(dailyvar(17),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh9n","9. " & dailyvar(18),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh9s",FormatNumber(dailyvar(19),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh10n","10. " & dailyvar(20),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"dh10s",FormatNumber(dailyvar(21),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh1n","1. " & weeklyvar(2),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh1s",FormatNumber(weeklyvar(3),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh2n","2. " & weeklyvar(4),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh2s",FormatNumber(weeklyvar(5),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh3n","3. " & weeklyvar(6),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh3s",FormatNumber(weeklyvar(7),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh4n","4. " & weeklyvar(8),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh4s",FormatNumber(weeklyvar(9),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh5n","5. " & weeklyvar(10),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh5s",FormatNumber(weeklyvar(11),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh6n","6. " & weeklyvar(12),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh6s",FormatNumber(weeklyvar(13),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh7n","7. " & weeklyvar(14),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh7s",FormatNumber(weeklyvar(15),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh8n","8. " & weeklyvar(16),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh8s",FormatNumber(weeklyvar(17),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh9n","9. " & weeklyvar(18),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh9s",FormatNumber(weeklyvar(19),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh10n","10. " & weeklyvar(20),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"wh10s",FormatNumber(weeklyvar(21),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      end If
    Case 5
      if osbkey="" then
        clearhslabels
        clearosblabels
        DMDintroloop
      Else
      clearhslabels
      clearosblabels
      PuPlayer.playlistplayex pBackglass,"videoattract","osb-at20.mp4",100,1
      PuPlayer.LabelSet pBackglass,"notetitle","All Time Top 20!!!",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","Keep playing and earn you spot in history",1,""
      PuPlayer.LabelSet pBackglass,"ah1n","1. " & alltimevar(2),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah1s",FormatNumber(alltimevar(3),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah2n","2. " & alltimevar(4),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah2s",FormatNumber(alltimevar(5),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah3n","3. " & alltimevar(6),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah3s",FormatNumber(alltimevar(7),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah4n","4. " & alltimevar(8),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah4s",FormatNumber(alltimevar(9),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah5n","5. " & alltimevar(10),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah5s",FormatNumber(alltimevar(11),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah6n","6. " & alltimevar(12),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah6s",FormatNumber(alltimevar(13),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah7n","7. " & alltimevar(14),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah7s",FormatNumber(alltimevar(15),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah8n","8. " & alltimevar(16),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah8s",FormatNumber(alltimevar(17),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah9n","9. " & alltimevar(18),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah9s",FormatNumber(alltimevar(19),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah10n","10. " & alltimevar(20),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah10s",FormatNumber(alltimevar(21),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah11n","11. " & alltimevar(22),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah11s",FormatNumber(alltimevar(23),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah12n","12. " & alltimevar(24),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah12s",FormatNumber(alltimevar(25),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah13n","13. " & alltimevar(26),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah13s",FormatNumber(alltimevar(27),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah14n","14. " & alltimevar(28),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah14s",FormatNumber(alltimevar(29),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah15n","15. " & alltimevar(30),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah15s",FormatNumber(alltimevar(31),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah16n","16. " & alltimevar(32),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah16s",FormatNumber(alltimevar(33),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah17n","17. " & alltimevar(34),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah17s",FormatNumber(alltimevar(35),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah18n","18. " & alltimevar(36),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah18s",FormatNumber(alltimevar(37),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah19n","19. " & alltimevar(38),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah19s",FormatNumber(alltimevar(39),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah20n","20. " & alltimevar(40),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      PuPlayer.LabelSet pBackglass,"ah20s",FormatNumber(alltimevar(41),0),1,"{'mt':2, 'size': 2.5, 'xalign': 0, 'yalign': 1}"
      end If
    Case 6
      clearhslabels
      clearosblabels
      PuPlayer.LabelSet pBackglass,"notetitle","Want to code your own original table?",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","Check Scottywics YouTube Channel for helpful guides",1,""
      PuPlayer.playlistplayex pBackglass,"videoattract","logos-sm.mov",100,1

    Case 7
      PuPlayer.playlistplayex pBackglass,"videoattract","nohatesm.mov",100,1
      PuPlayer.LabelSet pBackglass,"notetitle","Hope You Enjoy The Game!",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","Table Update By LoadedWeapon",1,""
      introposition = 0
  End Select
  End Sub

  sub clearosblabels
    PuPlayer.LabelSet pBackglass,"dh1n","",1,""
    PuPlayer.LabelSet pBackglass,"dh1s","",1,""
    PuPlayer.LabelSet pBackglass,"dh2n","",1,""
    PuPlayer.LabelSet pBackglass,"dh2s","",1,""
    PuPlayer.LabelSet pBackglass,"dh3n","",1,""
    PuPlayer.LabelSet pBackglass,"dh3s","",1,""
    PuPlayer.LabelSet pBackglass,"dh4n","",1,""
    PuPlayer.LabelSet pBackglass,"dh4s","",1,""
    PuPlayer.LabelSet pBackglass,"dh5n","",1,""
    PuPlayer.LabelSet pBackglass,"dh5s","",1,""
    PuPlayer.LabelSet pBackglass,"dh6n","",1,""
    PuPlayer.LabelSet pBackglass,"dh6s","",1,""
    PuPlayer.LabelSet pBackglass,"dh7n","",1,""
    PuPlayer.LabelSet pBackglass,"dh7s","",1,""
    PuPlayer.LabelSet pBackglass,"dh8n","",1,""
    PuPlayer.LabelSet pBackglass,"dh8s","",1,""
    PuPlayer.LabelSet pBackglass,"dh9n","",1,""
    PuPlayer.LabelSet pBackglass,"dh9s","",1,""
    PuPlayer.LabelSet pBackglass,"dh10n","",1,""
    PuPlayer.LabelSet pBackglass,"dh10s","",1,""
    PuPlayer.LabelSet pBackglass,"wh1n","",1,""
    PuPlayer.LabelSet pBackglass,"wh1s","",1,""
    PuPlayer.LabelSet pBackglass,"wh2n","",1,""
    PuPlayer.LabelSet pBackglass,"wh2s","",1,""
    PuPlayer.LabelSet pBackglass,"wh3n","",1,""
    PuPlayer.LabelSet pBackglass,"wh3s","",1,""
    PuPlayer.LabelSet pBackglass,"wh4n","",1,""
    PuPlayer.LabelSet pBackglass,"wh4s","",1,""
    PuPlayer.LabelSet pBackglass,"wh5n","",1,""
    PuPlayer.LabelSet pBackglass,"wh5s","",1,""
    PuPlayer.LabelSet pBackglass,"wh6n","",1,""
    PuPlayer.LabelSet pBackglass,"wh6s","",1,""
    PuPlayer.LabelSet pBackglass,"wh7n","",1,""
    PuPlayer.LabelSet pBackglass,"wh7s","",1,""
    PuPlayer.LabelSet pBackglass,"wh8n","",1,""
    PuPlayer.LabelSet pBackglass,"wh8s","",1,""
    PuPlayer.LabelSet pBackglass,"wh9n","",1,""
    PuPlayer.LabelSet pBackglass,"wh9s","",1,""
    PuPlayer.LabelSet pBackglass,"wh10n","",1,""
    PuPlayer.LabelSet pBackglass,"wh10s","",1,""
    PuPlayer.LabelSet pBackglass,"ah1n","",1,""
    PuPlayer.LabelSet pBackglass,"ah1s","",1,""
    PuPlayer.LabelSet pBackglass,"ah2n","",1,""
    PuPlayer.LabelSet pBackglass,"ah2s","",1,""
    PuPlayer.LabelSet pBackglass,"ah3n","",1,""
    PuPlayer.LabelSet pBackglass,"ah3s","",1,""
    PuPlayer.LabelSet pBackglass,"ah4n","",1,""
    PuPlayer.LabelSet pBackglass,"ah4s","",1,""
    PuPlayer.LabelSet pBackglass,"ah5n","",1,""
    PuPlayer.LabelSet pBackglass,"ah5s","",1,""
    PuPlayer.LabelSet pBackglass,"ah6n","",1,""
    PuPlayer.LabelSet pBackglass,"ah6s","",1,""
    PuPlayer.LabelSet pBackglass,"ah7n","",1,""
    PuPlayer.LabelSet pBackglass,"ah7s","",1,""
    PuPlayer.LabelSet pBackglass,"ah8n","",1,""
    PuPlayer.LabelSet pBackglass,"ah8s","",1,""
    PuPlayer.LabelSet pBackglass,"ah9n","",1,""
    PuPlayer.LabelSet pBackglass,"ah9s","",1,""
    PuPlayer.LabelSet pBackglass,"ah10n","",1,""
    PuPlayer.LabelSet pBackglass,"ah10s","",1,""
    PuPlayer.LabelSet pBackglass,"ah11n","",1,""
    PuPlayer.LabelSet pBackglass,"ah11s","",1,""
    PuPlayer.LabelSet pBackglass,"ah12n","",1,""
    PuPlayer.LabelSet pBackglass,"ah12s","",1,""
    PuPlayer.LabelSet pBackglass,"ah13n","",1,""
    PuPlayer.LabelSet pBackglass,"ah13s","",1,""
    PuPlayer.LabelSet pBackglass,"ah14n","",1,""
    PuPlayer.LabelSet pBackglass,"ah14s","",1,""
    PuPlayer.LabelSet pBackglass,"ah15n","",1,""
    PuPlayer.LabelSet pBackglass,"ah15s","",1,""
    PuPlayer.LabelSet pBackglass,"ah16n","",1,""
    PuPlayer.LabelSet pBackglass,"ah16s","",1,""
    PuPlayer.LabelSet pBackglass,"ah17n","",1,""
    PuPlayer.LabelSet pBackglass,"ah17s","",1,""
    PuPlayer.LabelSet pBackglass,"ah18n","",1,""
    PuPlayer.LabelSet pBackglass,"ah18s","",1,""
    PuPlayer.LabelSet pBackglass,"ah19n","",1,""
    PuPlayer.LabelSet pBackglass,"ah19s","",1,""
    PuPlayer.LabelSet pBackglass,"ah20n","",1,""
    PuPlayer.LabelSet pBackglass,"ah20s","",1,""
  end Sub


  sub clearhslabels
    PuPlayer.LabelSet pBackglass,"HighScore","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL1","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL2","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL3","",1,""
    PuPlayer.LabelSet pBackglass,"HighScore","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL1","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL2","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL3","",1,""
    PuPlayer.LabelSet pBackglass,"high1name","",1,""
    PuPlayer.LabelSet pBackglass,"high1score","",1,""
    PuPlayer.LabelSet pBackglass,"high2name","",1,""
    PuPlayer.LabelSet pBackglass,"high2score","",1,""
    PuPlayer.LabelSet pBackglass,"high3name","",1,""
    PuPlayer.LabelSet pBackglass,"high3score","",1,""
    PuPlayer.LabelSet pBackglass,"high4name","",1,""
    PuPlayer.LabelSet pBackglass,"high4score","",1,""
  end Sub

  Dim introtime
  introtime = 0

  Sub intromover_timer
    introtime = introtime + 1
    If introposition = 1 Then
      If introtime = 47 Then
        DMDintroloop
      End If
    End If
    If introposition = 2 Then
      If introtime = 13 Then
        DMDintroloop
      End If
    End If
    If introposition = 3 Then
      If introtime = 13 Then
        DMDintroloop
      End If
    End If
    If introposition = 4 Then
      If introtime = 9 Then
        DMDintroloop
      End If
    End If
    If introposition = 5 Then
      If introtime = 10 Then
        DMDintroloop
      End If
    End If
    If introposition = 6 Then
      If introtime = 10 Then
        DMDintroloop
      End If
    End If
    If introposition = 7 Then
      If introtime = 10 Then
        DMDintroloop
      End If
    End If
    If introposition = 0 Then
      If introtime = 9 Then
        introposition = 0
        DMDintroloop
      End If
    End If
  End Sub


  Sub StartAttractMode()
    DOF 323, DOFOn   'DOF MX - Attract Mode ON
    bAttractMode = True
    UltraDMDTimer.Enabled = 1
    StartLightSeq
    'ShowTableInfo
    DMDintroloop
    StartRainbow aLights
    DMDattract.Enabled = 1
    intromover.enabled = true
    ruleshelperoff
  End Sub

  Sub StopAttractMode()
    DOF 323, DOFOff   'DOF MX - Attract Mode Off
    bAttractMode = False
    DMDScoreNow
    LightSeqAttract.StopPlay
    LightSeqFlasher.StopPlay
    StopRainbow
    ResetAllLightsColor
    DMDattract.Enabled = 0
    intromover.enabled = false

  'StopSong
  End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  CINEMATIC SKIPPING
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X


  Dim ldown:ldown = 0
  Dim rdown:rdown = 0

  Sub checkdown
    If ldown + rdown = 2 Then
      skipscene
    End If
  End Sub

  Dim dropwallskip:dropwallskip = 0
  Dim levitateskip:levitateskip = 0
  Dim radioskip:radioskip = 0
  Dim compassskip:compassskip = 0
  Dim poolskip:poolskip = 0
  Dim badmenskip:badmenskip = 0
  Dim partyskip:partyskip = 0
  Dim willskip:willskip = 0
  Dim barbskip:barbskip = 0
  Dim extraballskip:extraballskip = 0

  Sub skipscene
    If dropwallskip = 1 Then
      Dropwall
      PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,1
      dropwallskip = 0
    End If

    If levitateskip = 1 Then
      levitatestart
    End If

    If radioskip = 1 Then
      radiostart
    End If

    If poolskip = 1 Then
      poolstart
    End If

    If compassskip = 1 Then
      compassstart
    End If

    If badmenskip = 1 Then
      StartRun
    End If

    If partyskip = 1 Then
      startparty
    End If

    If willskip = 1 Then
      EnterUpsideDown
    End If

    If barbskip = 1 Then
      StartBarb
    End If

    If extraballskip = 1 Then
      exitav
    End If

  End Sub






'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  LIGHTING / RAINBOW LIGHTS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'


  '********************************************************************************************
  ' Only for VPX 10.2 and higher.
  ' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
  ' When TotalPeriod done, light or flasher will be set to FinalState value where
  ' Final State values are:   0=Off, 1=On, 2=Return to previous State
  '********************************************************************************************

  Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first version

    If TypeName(MyLight) = "Light" Then

      If FinalState = 2 Then
        FinalState = MyLight.State 'Keep the current light state
      End If
      MyLight.BlinkInterval = BlinkPeriod
      MyLight.Duration 2, TotalPeriod, FinalState
    ElseIf TypeName(MyLight) = "Flasher" Then

      Dim steps

      ' Store all blink information
      steps = Int(TotalPeriod / BlinkPeriod + .5) 'Number of ON/OFF steps to perform
      If FinalState = 2 Then                      'Keep the current flasher state
        FinalState = ABS(MyLight.Visible)
      End If
      MyLight.UserValue = steps * 10 + FinalState 'Store # of blinks, and final state

      ' Start blink timer and create timer subroutine
      MyLight.TimerInterval = BlinkPeriod
      MyLight.TimerEnabled = 0
      MyLight.TimerEnabled = 1
      ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
    End If
  End Sub

  '******************************************
  ' Change light color - simulate color leds
  ' changes the light color and state
  ' 10 colors: red, orange, amber, yellow...
  '******************************************
  ' in this table this colors are use to keep track of the progress during the acts and battles

  'colors
  Dim red, orange, amber, yellow, darkgreen, green, blue, darkblue, purple, white, base

  red = 10
  orange = 9
  amber = 8
  yellow = 7
  darkgreen = 6
  green = 5
  blue = 4
  darkblue = 3
  purple = 2
  white = 1
  base = 11

  Sub SetLightColor(n, col, stat)
    Select Case col
      Case red
        n.color = RGB(18, 0, 0)
        n.colorfull = RGB(255, 0, 0)
      Case orange
        n.color = RGB(18, 3, 0)
        n.colorfull = RGB(255, 64, 0)
      Case amber
        n.color = RGB(193, 49, 0)
        n.colorfull = RGB(255, 153, 0)
      Case yellow
        n.color = RGB(18, 18, 0)
        n.colorfull = RGB(255, 255, 0)
      Case darkgreen
        n.color = RGB(0, 8, 0)
        n.colorfull = RGB(0, 64, 0)
      Case green
        n.color = RGB(0, 18, 0)
        n.colorfull = RGB(0, 255, 0)
      Case blue
        n.color = RGB(0, 18, 18)
        n.colorfull = RGB(0, 255, 255)
      Case darkblue
        n.color = RGB(0, 8, 8)
        n.colorfull = RGB(0, 64, 64)
      Case purple
        n.color = RGB(128, 0, 128)
        n.colorfull = RGB(255, 0, 255)
      Case white
        n.color = RGB(255, 252, 224)
        n.colorfull = RGB(193, 91, 0)
      Case white
        n.color = RGB(255, 252, 224)
        n.colorfull = RGB(193, 91, 0)
      Case base
        n.color = RGB(255, 197, 143)
        n.colorfull = RGB(255, 255, 236)
    End Select
    If stat <> -1 Then
      n.State = 0
      n.State = stat
    End If
  End Sub

  Sub ResetAllLightsColor ' Called at a new game
    'shoot again
    SetLightColor LightShootAgain, red, -1
    SetLightColor LightShootAgain1, red, -1
    ' lanes
    SetLightColor ll1, blue, -1
    SetLightColor ll8, blue, -1
    SetLightColor ll2, blue, -1
    SetLightColor ll7, blue, -1
    SetLightColor ll3, blue, -1
    SetLightColor ll6, blue, -1
    SetLightColor ll4, blue, -1
    SetLightColor ll9, blue, -1
    SetLightColor ll5, blue, -1
    SetLightColor ll10, blue, -1
    ' Run Lights
    SetLightColor lr1, darkblue, -1
    SetLightColor lr8, darkblue, -1
    SetLightColor lr2, darkblue, -1
    SetLightColor lr9, darkblue, -1
    SetLightColor lr3, darkblue, -1
    SetLightColor lr12, darkblue, -1
    SetLightColor lr4, darkblue, -1
    SetLightColor lr11, darkblue, -1
    SetLightColor lr5, darkblue, -1
    SetLightColor lr10, darkblue, -1
    SetLightColor lr6, darkblue, -1
    SetLightColor lr7, darkblue, -1
    ' Mode Lights
    SetLightColor lm1, yellow, -1
    SetLightColor lm9, yellow, -1
    SetLightColor lm2, blue, -1
    SetLightColor lm8, blue, -1
    SetLightColor lm3, darkgreen, -1
    SetLightColor lm7, darkgreen, -1
    SetLightColor lm4, red, -1
    SetLightColor lm6, red, -1
    ' escape lights row 1
    SetLightColor le2, blue, -1
    SetLightColor le16, blue, -1
    SetLightColor le4, blue, -1
    SetLightColor le19, blue, -1
    SetLightColor le6, blue, -1
    SetLightColor le21, blue, -1
    SetLightColor le11, blue, -1
    SetLightColor le25, blue, -1
    SetLightColor le13, blue, -1
    SetLightColor le28, blue, -1
    SetLightColor le15, blue, -1
    SetLightColor le29, blue, -1
    ' escape lights row 2
    SetLightColor le1, purple, -1
    SetLightColor le17, purple, -1
    SetLightColor le3, purple, -1
    SetLightColor le18, purple, -1
    SetLightColor le5, purple, -1
    SetLightColor le20, purple, -1
    SetLightColor le10, purple, -1
    SetLightColor le26, purple, -1
    SetLightColor le12, purple, -1
    SetLightColor le27, purple, -1
    SetLightColor le14, purple, -1
    SetLightColor le30, purple, -1
    ' escape Go Lights
    SetLightColor le7, yellow, -1
    SetLightColor le23, yellow, -1
    SetLightColor le8, orange, -1
    SetLightColor le22, orange, -1
    SetLightColor le9, red, -1
    SetLightColor le24, red, -1
    ' barb Lights
    SetLightColor lb1, yellow, -1
    SetLightColor lb6, yellow, -1
    SetLightColor lb2, yellow, -1
    SetLightColor lb5, yellow, -1
    SetLightColor lb3, yellow, -1
    SetLightColor lb8, yellow, -1
    SetLightColor lb4, yellow, -1
    SetLightColor lb7, yellow, -1
    SetLightColor leftrampred, yellow, -1
    SetLightColor leftrampred2, yellow, -1
    SetLightColor leftrampred1, yellow, -1
    SetLightColor leftrampred3, yellow, -1
    SetLightColor rightrampred, yellow, -1
    SetLightColor rightrampred1, yellow, -1
    ' lock Lights
    SetLightColor llo2, darkblue, -1
    SetLightColor llo8, darkblue, -1
    'Extra Ball
    SetLightColor llo5, orange, -1
    SetLightColor llo10, orange, -1
    SetLightColor lro1, yellow, -1
    SetLightColor lro3, yellow, -1
    ' Orbit & Ramp Lights
    SetLightColor llo1, red, -1
    SetLightColor llo9, red, -1
    SetLightColor llr1, red, -1
    SetLightColor llr2, red, -1
    SetLightColor llo3, red, -1
    SetLightColor llo7, red, -1
    SetLightColor lc1, red, -1
    SetLightColor lc4, red, -1
    SetLightColor lro2, red, -1
    SetLightColor lro6, red, -1
    'SetLightColor lrr1, red, -1
    SetLightColor lro4, red, -1
    SetLightColor lro5, red, -1
    SetLightColor wallll, red, -1
    SetLightColor wallll1, red, -1
    SetLightColor wallml, red, -1
    SetLightColor wallml1, red, -1
    SetLightColor wallrl, red, -1
    SetLightColor wallrl1, red, -1
    ' lightning bolts
    SetLightColor lbolt1, white, -1
    SetLightColor lbolt7, white, -1
    SetLightColor lbolt2, white, -1
    SetLightColor lbolt8, white, -1
    SetLightColor lbolt3, white, -1
    SetLightColor lbolt10, white, -1
    SetLightColor lbolt4, white, -1
    SetLightColor lbolt9, white, -1
    SetLightColor lbolt5, white, -1
    SetLightColor lbolt11, white, -1
    SetLightColor lbolt6, white, -1
    SetLightColor lbolt12, white, -1
    SetLightColor avclubready, white, -1
    SetLightColor avclubready1, white, -1
    ' battery
    SetLightColor batterygreen, green, -1
    SetLightColor batterygreen1, green, -1
    SetLightColor batteryorange, orange, -1
    SetLightColor batteryorange1, orange, -1
    SetLightColor batteryyellow, yellow, -1
    SetLightColor batteryyellow1, yellow, -1
    SetLightColor batteryred, red, -1
    SetLightColor batteryred1, red, -1
    ' modes
    SetLightColor levitate, purple, -1
    SetLightColor levitate1, purple, -1
    SetLightColor compass, purple, -1
    SetLightColor compass1, purple, -1
    SetLightColor pool, purple, -1
    SetLightColor pool1, purple, -1
    SetLightColor radio, purple, -1
    SetLightColor radio1, purple, -1
    ' extra ball
    SetLightColor extraball, green, -1
    SetLightColor extraball1, green, -1
    ' mode Lights
    SetLightColor llo4, purple, -1
    SetLightColor llo6, purple, -1
    SetLightColor lc2, purple, -1
    SetLightColor lc3, purple, -1
    SetLightColor llo5, purple, -1
    SetLightColor llo10, purple, -1
    ' characters
    SetLightColor lucas, white, -1
    SetLightColor jonathan, white, -1
    SetLightColor nancy, white, -1
    SetLightColor joyce, white, -1
    SetLightColor will, white, -1
    SetLightColor elevencrew, white, -1
    SetLightColor mikey, white, -1
    SetLightColor hopper, white, -1
    SetLightColor dustin, white, -1
    SetLightColor partylock, orange, -1
    SetLightColor partylock1, orange, -1

    ' upsidedown Lights
    SetLightColor upsidedownarrow, darkgreen, -1
    SetLightColor upsidedowncircle, darkgreen, -1
    ' mystery Light
    SetLightColor mysterylight, orange, -1
    SetLightColor mysterylight1, orange, -1
    ' kickback
    SetLightColor kickbacklight, green, -1
    SetLightColor kickbacklight1, green, -1
    SetLightColor lm5, orange, -1
    SetLightColor lm10, orange, -1



  End Sub

  Sub UpdateBonusColors
  End Sub

  '*************************
  ' Rainbow Changing Lights
  '*************************

  Dim RGBStep, RGBFactor, rRed, rGreen, rBlue, RainbowLights

  Sub StartRainbow(n)
    set RainbowLights = n
    RGBStep = 0
    RGBFactor = 5
    rRed = 255
    rGreen = 0
    rBlue = 0
    RainbowTimer.Enabled = 1
  End Sub

  Sub StopRainbow()
    Dim obj
    RainbowTimer.Enabled = 0
    RainbowTimer.Enabled = 0
      For each obj in RainbowLights
        SetLightColor obj, "white", 0
      Next
  End Sub

  Sub RainbowTimer_Timer 'rainbow led light color changing
    Dim obj
    Select Case RGBStep
      Case 0 'Green
        rGreen = rGreen + RGBFactor
        If rGreen > 255 then
          rGreen = 255
          RGBStep = 1
        End If
      Case 1 'Red
        rRed = rRed - RGBFactor
        If rRed < 0 then
          rRed = 0
          RGBStep = 2
        End If
      Case 2 'Blue
        rBlue = rBlue + RGBFactor
        If rBlue > 255 then
          rBlue = 255
          RGBStep = 3
        End If
      Case 3 'Green
        rGreen = rGreen - RGBFactor
        If rGreen < 0 then
          rGreen = 0
          RGBStep = 4
        End If
      Case 4 'Red
        rRed = rRed + RGBFactor
        If rRed > 255 then
          rRed = 255
          RGBStep = 5
        End If
      Case 5 'Blue
        rBlue = rBlue - RGBFactor
        If rBlue < 0 then
          rBlue = 0
          RGBStep = 0
        End If
    End Select
      For each obj in RainbowLights
        obj.color = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
        obj.colorfull = RGB(rRed, rGreen, rBlue)
      Next
  End Sub


  Sub StartLightSeq()

    On Error Resume Next
    'lights sequences
'   LightSeqaxmas.UpdateInterval = 150
'   LightSeqaxmas.Play SeqRandom, 10, , 50000
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
    LightSeqAttract.UpdateInterval = 25
    LightSeqAttract.Play SeqBlinking, , 5, 150
    LightSeqAttract.Play SeqRandom, 40, , 4000
    LightSeqAttract.Play SeqAllOff
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqCircleOutOn, 15, 3
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqLeftOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 40, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 40, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqRightOn, 30, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqLeftOn, 30, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqCircleOutOn, 15, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqStripe1VertOn, 50, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe1VertOn, 50, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe2VertOn, 50, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe1VertOn, 25, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe2VertOn, 25, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
  End Sub

  Sub LightSeqAttract_PlayDone()
    StartLightSeq()
  End Sub

  Sub LightSeqTilt_PlayDone()
    LightSeqTilt.Play SeqAllOff
  End Sub

  Sub LightSeqSkillshot_PlayDone()
    LightSeqSkillshot.Play SeqAllOff
  End Sub



  '**********************
  '     GI effects
  ' independent routine
  ' it turns on the gi
  ' when there is a ball
  ' in play
  '**********************

  Dim OldGiState
  OldGiState = -1   'start witht the Gi off

  Sub ChangeGi(col) 'changes the gi color
    Dim bulb
    For each bulb in aGILights
      SetLightColor bulb, col, -1
    Next
  End Sub

  Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
      OldGiState = Ubound(tmp)
      If UBound(tmp) = 3 Then 'we have 4 captive balls on the table (-1 means no balls, 0 is the first ball, 1 is the second..)
        'GiOff               ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
      Else
        'Gion
      End If
    End If
  End Sub

  Sub GiOn
    DOF 126, DOFOn  'RGB Undercab - Default Colour
    Dim bulb
    For each bulb in aGiLights
      SetLightColor bulb, base, -1
      bulb.State = 1
    Next
    flasher20.opacity = 100
    flasher21.opacity = 2000
    spot2.opacity = 1000
    spot3.opacity = 400
    spot4.opacity = 1000
    spot5.opacity = 400
    backspot.opacity = 2000
    reflefton.opacity = 100
    refleftoff.opacity = 0
    refrightooff.opacity = 0
    refrightoon.opacity = 100
  End Sub

  Sub GiRed
    DOF 500, DOFOn  'RGB Undercab - Red
    Dim bulb
    For each bulb in aGiLights
      SetLightColor bulb, red, -1
      bulb.State = 1
    Next
    flasherred.opacity = 100
    Flasherred2.opacity = 2000
    Spotred1.opacity = 400
    Spotred2.opacity = 1000
    spotred3.opacity = 400
    spotred4.opacity = 1000
    backspot.opacity = 2000
    refleftred.opacity = 100
    refleftoff.opacity = 0
    refrightooff.opacity = 0
    refrightred.opacity = 100
  End Sub

  Sub GiYellow
    DOF 501, DOFOn  'RGB Undercab - Yellow
    Dim bulb
    For each bulb in aGiLights
      SetLightColor bulb, yellow, -1
      bulb.State = 1
    Next
    Flasheryellow.opacity = 100
    Flasheryellow2.opacity = 2000
    Spotyellow1.opacity = 400
    Spotyellow2.opacity = 1000
    spotyellow3.opacity = 400
    spotyellow4.opacity = 1000
    backspot.opacity = 2000
    refleftyellow.opacity = 100
    refleftoff.opacity = 0
    refrightooff.opacity = 0
    refrightyellow.opacity = 100
  End Sub

  Sub GiPurple
    DOF 502, DOFOn  'RGB Undercab - Purple
    Dim bulb
    For each bulb in aGiLights
      SetLightColor bulb, purple, -1
      bulb.State = 1
    Next
    Flasherpurple.opacity = 2000
    Flasherpurple2.opacity = 60
    Spotpurple1.opacity = 400
    Spotpurple2.opacity = 1000
    spotpurple3.opacity = 400
    spotpurple4.opacity = 1000
    backspot.opacity = 2000
    refleftpurple.opacity = 100
    refleftoff.opacity = 0
    refrightooff.opacity = 0
    refrightpurple.opacity = 100
  End Sub

  Sub GiOrange
    DOF 503, DOFOn  'RGB Undercab - Orange
    Dim bulb
    For each bulb in aGiLights
      SetLightColor bulb, orange, -1
      bulb.State = 1
      bulb.State = 1
    Next
    Flasherorange.opacity = 100
    Flasherorange2.opacity = 2000
    Spotorange1.opacity = 400
    Spotorange2.opacity = 1000
    spotorange3.opacity = 400
    spotorange4.opacity = 1000
    backspot.opacity = 2000
    refleftorange.opacity = 100
    refleftoff.opacity = 0
    refrightooff.opacity = 0
    refrightorange.opacity = 100
  End Sub

  Sub GiGreen
    DOF 504, DOFOn  'RGB Undercab - Green
    Dim bulb
    For each bulb in aGiLights
      SetLightColor bulb, green, -1
      bulb.State = 1
    Next
    Flashergreen.opacity = 100
    Flashergreen2.opacity = 2000
    Spotgreen1.opacity = 400
    Spotgreen2.opacity = 1000
    spotgreen3.opacity = 400
    spotgreen4.opacity = 1000
    backspot.opacity = 2000
    refleftgreen.opacity = 100
    refleftoff.opacity = 0
    refrightooff.opacity = 0
    refrightgreen.opacity = 100
  End Sub

  Sub GiBlue
    DOF 505, DOFOn  'RGB Undercab - Cyan
    Dim bulb
    For each bulb in aGiLights
      SetLightColor bulb, blue, -1
      bulb.State = 1
    Next
    Flasherblue.opacity = 100
    Flasherblue2.opacity = 2000
    Spotblue1.opacity = 400
    Spotblue2.opacity = 1000
    spotblue3.opacity = 400
    spotblue4.opacity = 1000
    backspot.opacity = 2000
    refleftblue.opacity = 100
    refleftoff.opacity = 0
    refrightooff.opacity = 0
    refrightblue.opacity = 100
  End Sub

  Sub GiLowerOn
    Flasher1.opacity = 0
    Dim bulb
    For each bulb in lowergi
      bulb.State = 1
    Next

  End Sub

  Sub GiLowerOff
    Flasher1.opacity = 100
    Dim bulb
    For each bulb in lowergi
      bulb.State = 0
    Next

  End Sub

  Sub GiOff
    DOF 126, DOFOff  'RGB Undercab Off
    DOF 500, DOFOff  'RGB Undercab Off
    DOF 501, DOFOff  'RGB Undercab Off
    DOF 502, DOFOff  'RGB Undercab Off
    DOF 503, DOFOff  'RGB Undercab Off
    DOF 504, DOFOff  'RGB Undercab Off
    DOF 505, DOFOff  'RGB Undercab Off
    Dim bulb
    For each bulb in aGiLights
      bulb.State = 0
    Next
    flasher20.opacity = 0
    flasher21.opacity = 0
    flasherred.opacity = 0
    Flasherred2.opacity = 0
    Flasheryellow.opacity = 0
    Flasheryellow2.opacity = 0
    Flasherpurple.opacity = 0
    Flasherpurple2.opacity = 0
    Flasherorange.opacity = 0
    Flasherorange2.opacity = 0
    Flashergreen.opacity = 0
    Flashergreen2.opacity = 0
    Flasherblue.opacity = 0
    Flasherblue2.opacity = 0
    spot2.opacity = 0
    spot3.opacity = 0
    spot4.opacity = 0
    spot5.opacity = 0
    Spotpurple1.opacity = 0
    Spotpurple2.opacity = 0
    spotpurple3.opacity = 0
    spotpurple4.opacity = 0
    Spotred1.opacity = 0
    Spotred2.opacity = 0
    spotred3.opacity = 0
    spotred4.opacity = 0
    Spotorange1.opacity = 0
    Spotorange2.opacity = 0
    spotorange3.opacity = 0
    spotorange4.opacity = 0
    Spotyellow1.opacity = 0
    Spotyellow2.opacity = 0
    spotyellow3.opacity = 0
    spotyellow4.opacity = 0
    Spotgreen1.opacity = 0
    Spotgreen2.opacity = 0
    spotgreen3.opacity = 0
    spotgreen4.opacity = 0
    Spotblue1.opacity = 0
    Spotblue2.opacity = 0
    spotblue3.opacity = 0
    spotblue4.opacity = 0
    backspot.opacity = 0
    reflefton.opacity = 0
    refleftoff.opacity = 100
    refleftgreen.opacity = 0
    refleftred.opacity = 0
    refleftblue.opacity = 0
    refleftyellow.opacity = 0
    refleftorange.opacity = 0
    refleftpurple.opacity = 0
    refrightooff.opacity = 100
    refrightoon.opacity = 0
    refrightgreen.opacity = 0
    refrightred.opacity = 0
    refrightblue.opacity = 0
    refrightyellow.opacity = 0
    refrightorange.opacity = 0
    refrightpurple.opacity = 0
  End Sub


  'xmaslights
Redim fspeed(axmas.Count),intensity(axmas.Count),fadeDir(axmas.Count)

'**************** this is the random fading routine for only one object, then below you have the function for every object inside the collection, using ExecuteGlobal
'Sub myFlasher_timer:_
'If Intensity(0) <=0 Then fspeed(0) = RndNum (1,10) * 20: fadeDir(0)=1:End If:_
'If Intensity(0) >= 3000 Then fadeDir(0)=-1 :End If:_
'Intensity(0) = Intensity(0) + fSpeed(0) * fadeDir(0):_
'myFlasher.opacity = intensity(0):_
'End Sub
'************************************************************************************************************************************


'**************** Starts the fading
Sub StartXMAS
exit Sub

Dim i:For i = 0 to axmas.Count-1
axmas(i).timerInterval=-1
axmas(i).timerenabled=1
axmas(i).opacity= 0
ExecuteGlobal ("Sub " & axmas(i).Name & "_timer:" & _
"If Intensity("& i &") <=0 Then fspeed("& i &") = RndNum (1,10) * 2: fadeDir("& i &")=1:End If:" & _
"If Intensity("& i &") >= 3000 Then fadeDir("& i &")=-1 :End If:" & _
"Intensity("& i &") = Intensity("& i &") + fSpeed("& i &") * fadeDir("& i &"):" & _
"axmas("& i &").opacity = intensity("& i &"):" & _
"End Sub")
Next
End Sub


'**************** Stops the fading and sets all lights On
Sub StopXMAS
Dim i:For i = 0 to axmas.Count-1
axmas(i).timerenabled=0
axmas(i).opacity= 0
Next
End Sub


  ' GI & light sequence effects


  Sub GiEffect(n)
    Select Case n
      Case 0 'all off
        LightSeqGi.Play SeqAlloff
      Case 1 'all blink
        LightSeqGi.UpdateInterval = 4
        LightSeqGi.Play SeqBlinking, , 5, 100
      Case 2 'random
        LightSeqGi.UpdateInterval = 10
        LightSeqGi.Play SeqRandom, 5, , 1000
      Case 3 'upon
        LightSeqGi.UpdateInterval = 4
        LightSeqGi.Play SeqUpOn, 5, 1
      Case 4 ' left-right-left
        LightSeqGi.UpdateInterval = 5
        LightSeqGi.Play SeqLeftOn, 10, 1
        LightSeqGi.UpdateInterval = 5
        LightSeqGi.Play SeqRightOn, 10, 1
    End Select
  End Sub

  Sub LightEffect(n)
    Select Case n
      Case 0 ' all off
        LightSeqInserts.Play SeqAlloff
      Case 1 'all blink
        LightSeqInserts.UpdateInterval = 4
        LightSeqInserts.Play SeqBlinking, , 5, 100
      Case 2 'random
        LightSeqInserts.UpdateInterval = 10
        LightSeqInserts.Play SeqRandom, 5, , 1000
      Case 3 'upon
        LightSeqInserts.UpdateInterval = 4
        LightSeqInserts.Play SeqUpOn, 10, 1
      Case 4 ' left-right-left
        LightSeqInserts.UpdateInterval = 5
        LightSeqInserts.Play SeqLeftOn, 10, 1
        LightSeqInserts.UpdateInterval = 5
        LightSeqInserts.Play SeqRightOn, 10, 1
      Case 5 'random
        LightSeqbumper.UpdateInterval = 4
        LightSeqbumper.Play SeqBlinking, , 5, 10
      Case 6 'random
        LightSeqRSling.UpdateInterval = 4
        LightSeqRSling.Play SeqBlinking, , 5, 6
      Case 7 'random
        LightSeqLSling.UpdateInterval = 4
        LightSeqLSling.Play SeqBlinking, , 5, 6
      Case 8 'random
        LightSeqBack.UpdateInterval = 4
        LightSeqBack.Play SeqBlinking, , 5, 6
      Case 9 'random
        LightSeqTruck.UpdateInterval = 4
        LightSeqTruck.Play SeqBlinking, , 5, 10
      Case 10 'random
        LightSeqbarb.UpdateInterval = 4
        'LightSeqbarb.Play SeqBlinking, , 5, 10
      Case 11 'random
        LightSeqdice.UpdateInterval = 4
        LightSeqdice.Play SeqBlinking, , 5, 10
      Case 12 'random
        LightSeqlr.UpdateInterval = 4
        LightSeqlr.Play SeqBlinking, , 5, 10
      Case 13 'random
        LightSeqrr.UpdateInterval = 4
        LightSeqrr.Play SeqBlinking, , 5, 10
    End Select
  End Sub

  ' Flasher Effects using lights

  Dim FEStep, FEffect
  FEStep = 0
  FEffect = 0

  Sub FlashEffect(n)
    Select case n
      Case 0 ' all off
        LightSeqFlasher.Play SeqAlloff
      Case 1 'all blink
        LightSeqFlasher.UpdateInterval = 4
        LightSeqFlasher.Play SeqBlinking, , 5, 100
      Case 2 'random
        LightSeqFlasher.UpdateInterval = 10
        LightSeqFlasher.Play SeqRandom, 5, , 1000
      Case 3 'upon
        LightSeqFlasher.UpdateInterval = 4
        LightSeqFlasher.Play SeqUpOn, 10, 1
      Case 4 ' left-right-left
        LightSeqFlasher.UpdateInterval = 5
        LightSeqFlasher.Play SeqLeftOn, 10, 1
        LightSeqFlasher.UpdateInterval = 5
        LightSeqFlasher.Play SeqRightOn, 10, 1
      Case 5 ' top flashers blink fast
    End Select
  End Sub


  '****************************
  ' Flashers - Thanks Flupper
  '****************************

  Dim FlashLevel1, FlashLevel2, FlashLevel3, FlashLevel4, FlashLevel5, FlashLevel6
  Flasherlight2.IntensityScale = 0
  Flasherlight3.IntensityScale = 0
  Flasherlight4.IntensityScale = 0
  Flasherlight5.IntensityScale = 0
  Flasherlight6.IntensityScale = 0
  Flasherlight7.IntensityScale = 0

  '*** top center Red flasher ***
  Sub Flasherflash2_Timer()
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

  '*** top left Red flasher ***
  Sub Flasherflash1_Timer()
    dim flashx3, matdim
    If not Flasherflash1.TimerEnabled Then
      Flasherflash1.TimerEnabled = True
      Flasherflash1.visible = 1
      Flasherlit1.visible = 1
    End If
    flashx3 = FlashLevel1 * FlashLevel1 * FlashLevel1
    Flasherflash1.opacity = 1500 * flashx3
    Flasherlit1.BlendDisableLighting = 10 * flashx3
    Flasherbase1.BlendDisableLighting =  flashx3
    Flasherlight3.IntensityScale = flashx3
    matdim = Round(10 * FlashLevel1)
    Flasherlit1.material = "domelit" & matdim
    FlashLevel1 = FlashLevel1 * 0.9 - 0.01
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


  '*** top right Red flasher ***
  Sub Flasherflash4_Timer()
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
    Flasherlight5.IntensityScale = flashx3
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


  '*** center Blue flasher ***
  Sub Flasherflash6_Timer()
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
    Flasherlight7.IntensityScale = flashx3
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


  '*** right red flasher ***
  Sub Flasherflash3_Timer()
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
    Flasherlight4.IntensityScale = flashx3
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


  '*** left white flasher ***
  Sub Flasherflash5_Timer()
    dim flashx3, matdim
    If not Flasherflash5.TimerEnabled Then
      Flasherflash5.TimerEnabled = True
      Flasherflash5.visible = 1
      Flasherlit5.visible = 1
    End If
    flashx3 = FlashLevel5 * FlashLevel5 * FlashLevel5
    Flasherflash5.opacity = 500 * flashx3
    Flasherlit5.BlendDisableLighting = 10 * flashx3
    Flasherbase5.BlendDisableLighting =  flashx3
    Flasherlight6.IntensityScale = flashx3
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

  Dim flashseq
  flashseq = 0

  Sub flashflash_Timer()
    flashseq = flashseq + 1
    Select Case flashseq
      Case 1
        FlashLevel5 = 1 : Flasherflash5_Timer
      Case 2
        FlashLevel1 = 1 : Flasherflash1_Timer
      Case 3
        FlashLevel2 = 1 : Flasherflash2_Timer
      Case 4
        FlashLevel4 = 1 : Flasherflash4_Timer
      Case 5
        FlashLevel6 = 1 : Flasherflash6_Timer
      Case 6
        FlashLevel3 = 1 : Flasherflash3_Timer
      Case 7
        FlashLevel5 = 1 : Flasherflash5_Timer
      Case 8
        FlashLevel1 = 1 : Flasherflash1_Timer
      Case 9
        FlashLevel2 = 1 : Flasherflash2_Timer
      Case 10
        FlashLevel4 = 1 : Flasherflash4_Timer
      Case 11
        FlashLevel6 = 1 : Flasherflash6_Timer
      Case 12
        FlashLevel3 = 1 : Flasherflash3_Timer
      Case 13
        FlashLevel5 = 1 : Flasherflash5_Timer
      Case 14
        FlashLevel1 = 1 : Flasherflash1_Timer
      Case 15
        FlashLevel2 = 1 : Flasherflash2_Timer
      Case 16
        FlashLevel4 = 1 : Flasherflash4_Timer
      Case 17
        FlashLevel6 = 1 : Flasherflash6_Timer
      Case 18
        FlashLevel3 = 1 : Flasherflash3_Timer
      Case 19
        FlashLevel5 = 1 : Flasherflash5_Timer
      Case 20
        FlashLevel1 = 1 : Flasherflash1_Timer
      Case 21
        FlashLevel2 = 1 : Flasherflash2_Timer
      Case 22
        FlashLevel4 = 1 : Flasherflash4_Timer
      Case 23
        FlashLevel6 = 1 : Flasherflash6_Timer
      Case 24
        FlashLevel3 = 1 : Flasherflash3_Timer
      Case 25
        FlashLevel5 = 1 : Flasherflash5_Timer
      Case 26
        FlashLevel1 = 1 : Flasherflash1_Timer
      Case 27
        FlashLevel2 = 1 : Flasherflash2_Timer
      Case 28
        FlashLevel4 = 1 : Flasherflash4_Timer
      Case 29
        FlashLevel6 = 1 : Flasherflash6_Timer
      Case 30
        FlashLevel3 = 1 : Flasherflash3_Timer
        flashseq = 0
        flashflash.Enabled = False
    End Select
  End Sub

  Sub Flashxmas(n)
    Select case n
      Case 0 ' all off
        LightSeqaxmas.Play SeqAlloff
      Case 1 'all blink
        LightSeqaxmas.UpdateInterval = 4
        LightSeqaxmas.Play SeqBlinking, , 5, 100
      Case 2 'random
        LightSeqaxmas.UpdateInterval = 10
        LightSeqaxmas.Play SeqRandom, 5, , 1000
    End Select
  End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  DOF MX LEDS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'

  '****** DOF MX led extra Triggers and Commands (TerryRed) *******'

  'This section is only used for DOF MX addressable leds. Any commands
  'throughout the rest of the script labelled  'DOF MX'  is also only used
  'for addressable leds.


  Sub DOF_MX_OrbitOuterLeft_Hit()  'DOF MX - OrbitOuterLeft_Hit
    If barbMultiball = True Then
      DOF 455, DOFPulse  'DOF MX - Yellow
    End If
    If runMultiball = True Then
      DOF 460, DOFPulse  'DOF MX - Cyan
    End If
    If willMultiball = True Then
      DOF 465, DOFPulse  'DOF MX - Green
    End If
    If demoMultiball = True Then
      DOF 470, DOFPulse  'DOF MX - Red
    End If
    If PartyMultiball = True Then
      DOF 475, DOFPulse  'DOF MX - orange
    End If
    If (levitateactive = 1) or (radioactive = 1) or (poolactive = 1) or (compassactive = 1) Then
      DOF 480, DOFPulse  'DOF MX - Purple
    End If
    If (levitateactive = 0) and (radioactive = 0) and (poolactive = 0) and (compassactive = 0) and (partymultiball = false) and (barbMultiball = False) and (runMultiball = False) and (willMultiball = False) and (demoMultiball = False) Then
      DOF 450, DOFPulse  'DOF MX - Default
    End If
  End Sub

  Sub DOF_MX_OrbitLeft_Hit()        'DOF MX - OrbitLeft_Hit
    If barbMultiball = True Then
      DOF 456, DOFPulse   'DOF MX - Yellow
    End If
    If runMultiball = True Then
      DOF 461, DOFPulse   'DOF MX - Cyan
    End If
    If willMultiball = True Then
      DOF 466, DOFPulse   'DOF MX - Green
    End If
    If demoMultiball = True Then
      DOF 471, DOFPulse   'DOF MX - Red
    End If
    If PartyMultiball = True Then
      DOF 476, DOFPulse  'DOF MX - orange
    End If
    If (levitateactive = 1) or (radioactive = 1) or (poolactive = 1) or (compassactive = 1) Then
      DOF 481, DOFPulse  'DOF MX - Purple
    End If
    If (levitateactive = 0) and (radioactive = 0) and (poolactive = 0) and (compassactive = 0) and (partymultiball = false) and (barbMultiball = False) and (runMultiball = False) and (willMultiball = False) and (demoMultiball = False) Then
      DOF 451, DOFPulse   'DOF MX - Default
    End If
  End Sub

  Sub DOF_MX_OrbitCenter_Hit()       'DOF MX - OrbitCenter_Hit
    If barbMultiball = True Then
      DOF 457, DOFPulse   'DOF MX - Yellow
    End If
    If runMultiball = True Then
      DOF 462, DOFPulse   'DOF MX - Cyan
    End If
    If willMultiball = True Then
      DOF 467, DOFPulse   'DOF MX - Green
    End If
    If demoMultiball = True Then
      DOF 472, DOFPulse   'DOF MX - Red
    End If
    If PartyMultiball = True Then
      DOF 477, DOFPulse  'DOF MX - orange
    End If
    If (levitateactive = 1) or (radioactive = 1) or (poolactive = 1) or (compassactive = 1) Then
      DOF 482, DOFPulse  'DOF MX - Purple
    End If
    If (levitateactive = 0) and (radioactive = 0) and (poolactive = 0) and (compassactive = 0) and (partymultiball = false) and (barbMultiball = False) and (runMultiball = False) and (willMultiball = False) and (demoMultiball = False) Then
      DOF 452, DOFPulse   'DOF MX - Default
    End If
  End Sub

  Sub DOF_MX_OrbitRight_Hit()        'DOF MX - OrbitRight_Hit
    If barbMultiball = True Then
      DOF 458, DOFPulse   'DOF MX - Yellow
    End If
    If runMultiball = True Then
      DOF 463, DOFPulse   'DOF MX - Cyan
    End If
    If willMultiball = True Then
      DOF 468, DOFPulse   'DOF MX - Green
    End If
    If demoMultiball = True Then
      DOF 473, DOFPulse   'DOF MX - Red
    End If
    If PartyMultiball = True Then
      DOF 478, DOFPulse  'DOF MX - orange
    End If
    If (levitateactive = 1) or (radioactive = 1) or (poolactive = 1) or (compassactive = 1) Then
      DOF 483, DOFPulse  'DOF MX - Purple
    End If
    If (levitateactive = 0) and (radioactive = 0) and (poolactive = 0) and (compassactive = 0) and (partymultiball = false) and (barbMultiball = False) and (runMultiball = False) and (willMultiball = False) and (demoMultiball = False) Then
      DOF 453, DOFPulse   'DOF MX - Default
    End If
  End Sub

  Sub DOF_MX_OrbitOuterRight_Hit()   'DOF MX - OrbitOuterRight_Hit
    If barbMultiball = True Then
      DOF 459, DOFPulse   'DOF MX - Yellow
    End If
    If runMultiball = True Then
      DOF 464, DOFPulse   'DOF MX - Cyan
    End If
    If willMultiball = True Then
      DOF 469, DOFPulse   'DOF MX - Green
    End If
    If demoMultiball = True Then
      DOF 474, DOFPulse 'DOF MX - Red
    End If
    If PartyMultiball = True Then
      DOF 479, DOFPulse  'DOF MX - orange
    End If
    If (levitateactive = 1) or (radioactive = 1) or (poolactive = 1) or (compassactive = 1) Then
      DOF 484, DOFPulse  'DOF MX - Purple
    End If
    If (levitateactive = 0) and (radioactive = 0) and (poolactive = 0) and (compassactive = 0) and (partymultiball = false) and (barbMultiball = False) and (runMultiball = False) and (willMultiball = False) and (demoMultiball = False) Then
      DOF 454, DOFPulse   'DOF MX - Default
    End If
  End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  TOY SHAKES
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'

  '*************  TOYS SHAKE ON NUDGE  *************
  Dim mMagnet, cBall, RatioMultiplier

  Sub WobbleMagnet_Init
     Set mMagnet = new cvpmMagnet
     With mMagnet
      .InitMagnet WobbleMagnet, 1.5
      .Size = 100
      .CreateEvents mMagnet
      .MagnetOn = True
     End With
    Set cBall = ckicker.createball:ckicker.Kick 0,0:mMagnet.addball cball
   End Sub

  Sub ToysUpdate_timer
    If NOT Toysbouncing.Enabled Then
      RatioMultiplier = (thedemogorgonmodel.TransY)/370*RndNum(5,10)*0.1    'multiply it with a random number between 0.5 and 1
      thedemogorgonmodel.rotx=(ckicker.y - cball.y)*RatioMultiplier
      thedemogorgonmodel.roty=(cball.x - ckicker.x)*RatioMultiplier*0.5
    End If
  End Sub

  '*************  TOYS SHAKE ON HIT  *************

  'On Hit add this line and the toy will shake:
  'Toysbouncing.Enabled = 1:brake=0:perc=3:sbou=(Rndnum(-10,10))/10

  Dim bou,brake,perc,sbou
  Sub Toysbouncing_timer
    bou=bou+0.3:brake=brake+0.02
    If (perc-(brake*(perc/6)))<0 Then Me.Enabled = 0 :bou=0 :brake=0 :perc=0
    thedemogorgonmodel.rotx=sin(bou)*(perc-(brake*(perc/6)))
    thedemogorgonmodel.roty=cos((bou)*sbou)*(perc-(brake*(perc/6)))
  End Sub


  'Couldn't find an hit event for the 2 Demogorgon targets so I'll add it here:




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  MANUAL BALLCONTROL
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
  '

  Sub StartControl_Hit()
     Set ControlBall = ActiveBall
     contballinplay = true
  End Sub

  Sub StopControl_Hit()
     contballinplay = false
  End Sub

  Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
  Dim bcvel, bcyveloffset, bcboostmulti

  bcboost = 1 'Do Not Change - default setting
  bcvel = 4 'Controls the speed of the ball movement
  bcyveloffset = -0.01 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
  bcboostmulti = 3 'Boost multiplier to ball veloctiy (toggled with the B key)

  Sub BallControl_Timer()
     If Contball and ContBallInPlay then
        If bcright = 1 Then
           ControlBall.velx = bcvel*bcboost
        ElseIf bcleft = 1 Then
           ControlBall.velx = - bcvel*bcboost
        Else
           ControlBall.velx=0
        End If

       If bcup = 1 Then
          ControlBall.vely = -bcvel*bcboost
       ElseIf bcdown = 1 Then
          ControlBall.vely = bcvel*bcboost
       Else
          ControlBall.vely= bcyveloffset
       End If
     End If
  End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   SCORING FUNCTIONS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X

  Sub AddScore(points)
    If(Tilted = False) Then
      ' add the points to the current players score variable
      Score(CurrentPlayer) = Score(CurrentPlayer) + points
      DMDScore
      pUpdateScores
    End if

  ' you may wish to check to see if the player has gotten a replay
  End Sub

  ' Add bonus to the bonuspoints AND update the score board

  Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
      'DMD "black.png", "EXTRA", "BALL",  2000
      LightShootAgain.State = 1
      LightShootAgain1.State = 1
      flashflash.Enabled = True
      LightSeqFlasher.UpdateInterval = 150
      LightSeqFlasher.Play SeqRandom, 10, , 10000
      ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
      bExtraBallWonThisBall = True
      DOF 402, DOFPulse   'DOF MX - Extra Ball
      extraballready(CurrentPlayer) = 0
      If bMultiBallMode = false Then
      PuPlayer.playlistplayex pBackglass,"videoextraball","extraballsm.mov",100,1
      PuPlayer.playlistplayex pCallouts,"audiocallouts","extraball.wav",100,1
      chilloutthemusic
      extraball.state = 0
      extraball1.state = 0
      End If
    Else
    AddScore 4000000
    END If
  End Sub

  Sub AwardSpecial()
    Credits = Credits + 1
    PlaySound SoundFXDOF("knocker",136,DOFPulse,DOFKnocker)
    DOF 115, DOFPulse
    GiEffect 1
    LightEffect 1
    DOF 400, DOFPulse   'DOF MX - Special
  End Sub


  Sub AwardSkillshot()
    Dim i
    ResetSkillShotTimer_Timer
    run1on(CurrentPlayer) = 1
    run2on(CurrentPlayer) = 1
    run3on(CurrentPlayer) = 1
    lr1.State = 1
    lr8.State = 1
    lr2.State = 1
    lr9.State = 1
    lr3.State = 1
    lr12.State = 1
    AddScore SkillshotValue(CurrentPLayer)

    'do some light show
    CheckRunTargets
    GiEffect 1
    LightEffect 2
    LightEffect 9
    DOF 401, DOFPulse   'DOF MX - Skillshot
    PuPlayer.playlistplayex pCallouts,"audiocallouts","skillshot.wav",100,1
    chilloutthemusic
    PuPlayer.playlistplayex pBackglass,"videoskillshot","",100,3
    pNote "SKILLSHOT",SkillShotValue(CurrentPLayer)
    SkillShotValue(CurrentPLayer) = SkillShotValue(CurrentPLayer) + 1000000
  End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   BALL FUNCTIONS & DRAINS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X


  Function Balls
    Dim tmp
    tmp = bpgcurrent - BallsRemaining(CurrentPlayer) + 1
    If tmp> bpgcurrent Then
      Balls = bpgcurrent
    Else
      Balls = tmp
    End If
  End Function


  Sub FirstBall
    ResetForNewPlayerBall()
    CreateNewBall()
  End Sub


  Sub ResetForNewPlayerBall()
    PuPlayer.playresume 4
    If PlayersPlayingGame > 1 Then
      If CurrentPlayer = 1 Then
        PuPlayer.playlistplayex pBackglass,"videoplayers","player1sm1.mov",40,1
            PuPlayer.playlistplayex pCallouts,"audiocallouts","player1.wav",70,1
    chilloutthemusic
      Elseif currentplayer = 2 Then
        PuPlayer.playlistplayex pBackglass,"videoplayers","player2sm.mov",40,1
            PuPlayer.playlistplayex pCallouts,"audiocallouts","player2.wav",70,1
    chilloutthemusic
      Elseif currentplayer = 3 Then
        PuPlayer.playlistplayex pBackglass,"videoplayers","player3sm.mov",40,1
            PuPlayer.playlistplayex pCallouts,"audiocallouts","player3.wav",70,1
    chilloutthemusic
      Elseif currentplayer = 4 Then
        PuPlayer.playlistplayex pBackglass,"videoplayers","player4sm.mov",40,1
            PuPlayer.playlistplayex pCallouts,"audiocallouts","player4.wav",70,1
    chilloutthemusic
      End If
    Else
      pNote "BALL " & Balls,"LAUNCH BALL"
    PlaySound "flute"
    End If
    AddScore 0
    BonusPoints(CurrentPlayer) = 0
    bBonusHeld = False
    bExtraBallWonThisBall = False
    ResetNewBallLights()
    ResetNewBallVariables
    bBallSaverReady = True
    bSkillShotReady = True
    PuPlayer.LabelSet pBackglass,"HighScore","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL1","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL2","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL3","",1,""
    PuPlayer.LabelSet pBackglass,"HighScoreL4","",1,""
    PuPlayer.LabelSet pBackglass,"modetimer","",1,""
  End Sub


  Sub CreateNewBall()
    BallRelease.CreateSizedball BallSize / 2
    BallsOnPlayfield = BallsOnPlayfield + 1
    PlaySoundAt SoundFXDOF("fx_Ballrel", 114, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 90, 4
    If BallsOnPlayfield > 1 Then
      bMultiBallMode = True
      bAutoPlunger = True
    End If

    If barbMultiball = True Then
      'PlaySong "m_barb"  'this last number is the volume, from 0 to 1
    End If

    If runMultiball = True Then
      'PlaySong "m_run"  'this last number is the volume, from 0 to 1
    End If

    If willMultiball = True Then
      'PlaySong "m_multiball"  'this last number is the volume, from 0 to 1
    End If

    If demoMultiball = True Then
      'PlaySong "m_multiball"  'this last number is the volume, from 0 to 1
    End If
  End Sub


  Sub AddMultiball(nballs)
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
  End Sub

  Sub CreateMultiballTimer_Timer()
    If bBallInPlungerLane Then
      Exit Sub
    Else
      If BallsOnPlayfield <MaxMultiballs Then
        CreateNewBall()
        mBalls2Eject = mBalls2Eject -1
        If mBalls2Eject = 0 Then
          Me.Enabled = False
        End If
      Else
        mBalls2Eject = 0
        Me.Enabled = False
      End If
    End If
  End Sub

  Sub EndOfBall()
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1


    bMultiBallMode = False
    bOnTheFirstBall = False
    If NOT Tilted Then
      vpmtimer.addtimer 500, "EndOfBall2 '"
    Else
      vpmtimer.addtimer 500, "EndOfBall2 '"
    End If
  End Sub


  Sub EndOfBall2()
    Tilted = False
    Tilt = 0
    DisableTable False
    tilttableclear.enabled = False
    tilttime = 0
    If(ExtraBallsAwards(CurrentPlayer) <> 0) Then
      ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1
      If(ExtraBallsAwards(CurrentPlayer) = 0) Then
        LightShootAgain.State = 0
        LightShootAgain1.State = 0
      End If
      LightSeqFlasher.UpdateInterval = 150
      LightSeqFlasher.Play SeqRandom, 10, , 2000

      CreateNewBall()
      ResetForNewPlayerBall
      PuPlayer.playlistplayex pBackglass,"videoextraball","isthatsupposedtoimpress.mov",100,2
      pNote "SHOOT AGAIN","SAME PLAYER"
        PuPlayer.playlistplayex pCallouts,"audiocallouts","shootagain.wav",100,1
    chilloutthemusic
    Else
      BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1
      If(BallsRemaining(CurrentPlayer) <= 0) Then
        CheckHighScore()
      Else
        EndOfBallComplete()
      End If
    End If
  End Sub

  Sub EndOfBallComplete()
    ResetNewBallVariables
    Dim NextPlayer
    If(PlayersPlayingGame> 1) Then
      NextPlayer = CurrentPlayer + 1
      If(NextPlayer> PlayersPlayingGame) Then
        NextPlayer = 1
      End If
    Else
      NextPlayer = CurrentPlayer
    End If
    If((BallsRemaining(CurrentPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then
      EndOfGame()
    Else
      CurrentPlayer = NextPlayer
      AddScore 0
      ResetForNewPlayerBall()
      CreateNewBall()
      If PlayersPlayingGame> 1 Then
        PlaySound "vo_player" &CurrentPlayer
      End If
    End If
  End Sub

  Sub Balldrained
    DOF 312, DOFPulse  'DOF MX - Drained
    PuPlayer.playlistplayex pMusic,"audioclear","clear.mp3",100, 1
    PuPlayer.playlistplayex pBackglass,"videodrain","",100,1
  End Sub

  Sub Drain_Hit()
    Drain.DestroyBall
    BallsOnPlayfield = BallsOnPlayfield - 1
    PlaySoundAt "fx_drain", Drain
    If Tilted Then
      StopEndOfBallMode
    End If
    If(bGameInPLay = True) AND(Tilted = False) Then
      If(bBallSaverActive = True) Then
        AddMultiball 1
        bAutoPlunger = True
        If bMultiBallMode = False Then
          Ballsaved
          PuPlayer.playlistplayex pCallouts,"audiocallouts","ballsaved.wav",100,1
    chilloutthemusic
          'pNote "BALL SAVED",""
          PuPlayer.playlistplayex pBackglass,"videoballsaved","",100,1
        End If
      Else
        If(BallsOnPlayfield = 1) Then
          If(bMultiBallMode = True) then
            CheckKIDSLane
            If barbMultiball = True Then
              EndBarb
            End If
            If RunMultiball = True Then
              EndRun
            End If
            If PartyMultiball = True Then
              endparty
            End If
            If WillMultiball = True Then
              EndWill
            End If
            If DemoMultiball = True Then
              EndDemo
            End If
            bMultiBallMode = False
            ChangeGi "white"
          End If
          bMultiBallMode = False
          ChangeGi "white"
          CheckKIDSLane
      End If
        If(BallsOnPlayfield = 0) Then
          bMultiBallMode = False
          PlaySong "m_wait"
          ChangeGi "white"
    If poolactive = 1 Then
      poolend
    End If
    If compassactive = 1 Then
      compassend
    End If
    If radioactive = 1 Then
      radioend
    End If
    If levitateactive = 1 Then
      levitateend
    End If
    If barbMultiball = True Then
      EndBarb
    End If
    If RunMultiball = True Then
      EndRun
    End If
    If PartyMultiball = True Then
      endparty
    End If
    If WillMultiball = True Then
      EndWill
    End If
    If DemoMultiball = True Then
      EndDemo
    End If
          vpmtimer.addtimer 1000, "Balldrained '"
          vpmtimer.addtimer 4000, "EndOfBall '"
          StopEndOfBallMode
        End If
      End If
    End If
  End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  BALL SAVE & LAUNCH
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'
  Sub Ballsaved
    Saves = Saves + 1
    Select Case Saves
      Case 1 'DMD "ballsave1-2.wmv", "", "", 3000
      Case 2 'DMD "ballsave2-6.wmv", "", "", 7000:Saves = 0
    End Select
  End Sub

  Sub ballsavestarttrigger_hit
    FlashLevel3 = 1 : Flasherflash3_Timer
    DOF 360, DOFPulse  'DOF MX - Right Red Flasher
    ' if there is a need for a ball saver, then start off a timer       ' if there is a need for a ball saver, then start off a timer
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True) AND(15 <> 0) And(bBallSaverActive = False) Then
      EnableBallSaver 15
    End If
  End Sub

  Sub swPlungerRest_Hit()
    PlaySoundAt "fx_sensor", ActiveBall
    bBallInPlungerLane = True
    If bAutoPlunger Then
      PlungerIM.Strength = 45
      'PlungerIM.AutoFire
      PlungerIM.Strength = Plunger.MechStrength
      Plunger.AutoPlunger = true
      Plunger.Pullback
      Plunger.Fire
      DOF 114, DOFPulse
      DOF 115, DOFPulse
      DOF 318, DOFPulse  'DOF MX - Ball Launched (AutoPlunger)
      bAutoPlunger = False
      Plunger.AutoPlunger = false
    End If
    DOF 141, DOFOn  'Launch Ball Button Flashing - ON
    DOF 317, DOFOn 'DOF MX - Ball Ready to Launch - ON
    If bSkillShotReady Then
      swPlungerRest.TimerEnabled = 1
      UpdateSkillshot()
    End If
    LastSwitchHit = "swPlungerRest"
  End Sub

  sub pleasework

  End Sub


  Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
    swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
    If bSkillShotReady Then
      ResetSkillShotTimer.Enabled = 1
    End If
    DOF 317, DOFOff   'DOF MX - Ball is ready to Launch - OFF
    DOF 141, DOFOff   'Launch Ball Button Flashing - OFF
  End Sub


  Sub swPlungerRest_Timer
    'DMD "start-5.wmv", "", "",  6000
    swPlungerRest.TimerEnabled = 0
  End Sub

  Sub EnableBallSaver(seconds)
    bBallSaverActive = True
    bBallSaverReady = False
    BallSaverTimerExpired.Interval = 1000 * seconds
    BallSaverTimerExpired.Enabled = True
    BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
    BallSaverSpeedUpTimer.Enabled = True
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    LightShootAgain.BlinkInterval = 160
    LightShootAgain.State = 2
    LightShootAgain1.BlinkInterval = 160
    LightShootAgain1.State = 2
  End Sub

  ' The ball saver timer has expired.  Turn it off AND reset the game flag
  '
  Sub BallSaverTimerExpired_Timer()
    'debug.print "Ballsaver ended"
    BallSaverTimerExpired.Enabled = False
    ' clear the flag
    Dim waittime
    waittime = 4000
    vpmtimer.addtimer waittime, "ballsavegrace'"
    ' if you have a ball saver light then turn it off at this point
    If bExtraBallWonThisBall = True Then
      LightShootAgain.State = 1
      LightShootAgain1.State = 1
    Else
      LightShootAgain.State = 0
      LightShootAgain1.State = 0
    End If
  End Sub

  Sub ballsavegrace
    bBallSaverActive = False
  End Sub

  Sub BallSaverSpeedUpTimer_Timer()
    'debug.print "Ballsaver Speed Up Light"
    BallSaverSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    LightShootAgain.BlinkInterval = 80
    LightShootAgain.State = 2
    LightShootAgain1.BlinkInterval = 80
    LightShootAgain1.State = 2
  End Sub






'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  TABLE VARIABLES
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'

  Dim MagnetR
  Dim MagnetU
  Dim MagnetW
  Dim spinner

  ' Game Variables
  Dim LaneBonus
  Dim TargetBonus
  Dim RampBonus
  Dim OrbitBonus
  Dim spinvalue
  Dim barbMultiball
  Dim barbHits(10)
  Dim LookForBarb
  Dim RunMultiball
  Dim RunAway
  Dim RunHits(9)
  Dim BallsInRunLock(4)
  Dim LRHits(10)
  Dim RRHits(7)
  Dim WillMultiball
  Dim WillSuper
  Dim WillSuperReady
  Dim DemoMultiball
  Dim DemoHits
  Dim MonsterFinalBlow
  Dim BarbJackpots
  Dim finalflips
  Dim changetrack
  Dim Saves
  Dim Drains
  Dim energy
  Dim partymultiball
  Dim DGs
  Dim inmode

  ' player specific variables
  Dim barb1on(4)
  Dim barb2on(4)
  Dim barb3on(4)
  Dim barb4on(4)
  Dim BSuper(4)
  Dim barbcompleted(4)
  Dim barbjacks(4)
  Dim bumps(4)
  Dim mysteryready(4)
  Dim run1lock(4)
  Dim run1on(4)
  Dim run2on(4)
  Dim run3on(4)
  Dim run4on(4)
  Dim run5on(4)
  Dim run6on(4)
  Dim levitatedone(4)
  Dim pooldone(4)
  Dim compassdone(4)
  Dim radiodone(4)
  Dim avclubreadyon(4)
  Dim avsdone(4)
  Dim lucason(4)
  Dim jonathanon(4)
  Dim nancyon(4)
  Dim joyceon(4)
  Dim willon(4)
  Dim elevencrewon(4)
  Dim mikeyon(4)
  Dim hopperon(4)
  Dim dustinon(4)
  Dim partycollected(4)
  Dim BallsInLock(4)
  Dim batteryon(4)
  Dim partyjacks(4)
  Dim partydone(4)
  Dim udhits(4)
  Dim udtargetdown(4)
  Dim upsidedownlights(4)
  Dim barricadedown(4)
  Dim willcompleted(4)
  Dim mysteryaward1(4)
  Dim mysteryaward2(4)
  Dim mysteryaward3(4)
  Dim mysteryaward4(4)
  Dim mysteryaward5(4)
  Dim mysteryaward6(4)
  Dim currentmystery(4)
  Dim le1on(4)
  Dim le2on(4)
  Dim le3on(4)
  Dim le4on(4)
  Dim le5on(4)
  Dim le6on(4)
  Dim le7on(4)
  Dim le8on(4)
  Dim le9on(4)
  Dim le10on(4)
  Dim le11on(4)
  Dim le12on(4)
  Dim le13on(4)
  Dim le14on(4)
  Dim le15on(4)
  Dim walllon(4)
  Dim wallmon(4)
  Dim wallron(4)
  Dim WillHits(4)
  Dim extraballready(4)
  Dim modeebawarded(4)
  Dim batt1on(4)
  Dim batt2on(4)
  Dim batt3on(4)
  Dim batt4on(4)
  Dim kickbackon(4)
  Dim partyready(4)
  Dim udfirst(4)
  Dim badmensuper(4)
  Dim escapegateopen(4)
  Dim demoisdead(4)
  Dim inhighscore


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  GAME STARTING & RESETS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'

  Sub Game_Init() 'called at the start of a new game
    Dim i
    ' set base current player values
    For i = 0 to 4
      udfirst(i) = 0
      BallsInRunLock(i) = 0
      barb1on(i) = 0
      barb2on(i) = 0
      barb3on(i) = 0
      barb4on(i) = 0
      BSuper(i) = 0
      bumps(i) = 0
      run1lock(i) = 0
      run1on(i) = 0
      run2on(i) = 0
      run3on(i) = 0
      run4on(i) = 0
      run5on(i) = 0
      run6on(i) = 0
      demoisdead(i) = 0
      levitatedone(i) = 0
      radiodone(i) = 0
      pooldone(i) = 0
      compassdone(i) = 0
      avclubreadyon(i) = 0
      avsdone(i) = 0
      barbcompleted(i) = 0
      barbjacks(i) = 0
      lucason(i) = 0
      jonathanon(i) = 0
      nancyon(i) = 0
      joyceon(i) = 0
      willon(i) = 0
      elevencrewon(i) = 0
      mikeyon(i) = 0
      hopperon(i) = 0
      dustinon(i) = 0
      partycollected(i) = 0
      BallsInLock(i) = 0
      batteryon(i) = 0
      partyjacks(i) = 0
      partydone(i) = 0
      udhits(i) = 0
      udtargetdown(i) = 0
      upsidedownlights(i) = 0
      barricadedown(i) = 0
      willcompleted(i) = 0
      mysteryaward1(i) = 0
      mysteryaward2(i) = 0
      mysteryaward3(i) = 0
      mysteryaward4(i) = 0
      mysteryaward5(i) = 0
      mysteryaward6(i) = 0
      mysteryready(i) = 0
      currentmystery(i) = 0
      le1on(i) = 0
      le2on(i) = 0
      le3on(i) = 0
      le4on(i) = 0
      le5on(i) = 0
      le6on(i) = 0
      le7on(i) = 0
      le8on(i) = 0
      le9on(i) = 0
      le10on(i) = 0
      le11on(i) = 0
      le12on(i) = 0
      le13on(i) = 0
      le14on(i) = 0
      le15on(i) = 0
      walllon(i) = 0
      wallmon(i) = 0
      wallron(i) = 0
      willhits(i) = 0
      extraballready(i) = 0
      modeebawarded(i) = 0
      batt1on(i) = 1
      batt2on(i) = 1
      batt3on(i) = 0
      batt4on(i) = 0
      kickbackon(i) = 1
      partyready(i) = 0
      badmensuper(i) = 0
      escapegateopen(i) = 0
    Next
      inhighscore = False
    spinner.MotorOn = False
    partymultiball = False
    lrflashtime.Enabled = False
    bExtraBallWonThisBall = False
    'Init Variables
    For i = 0 to 4
      SkillshotValue(i) = 1000000 ' increases by 1000000 each time it is collected
    Next
      PuPlayer.LabelSet pBackglass,"high1name","",1,""
      PuPlayer.LabelSet pBackglass,"high1score","",1,""
      PuPlayer.LabelSet pBackglass,"high2name","",1,""
      PuPlayer.LabelSet pBackglass,"high2score","",1,""
      PuPlayer.LabelSet pBackglass,"high3name","",1,""
      PuPlayer.LabelSet pBackglass,"high3score","",1,""
      PuPlayer.LabelSet pBackglass,"high4name","",1,""
      PuPlayer.LabelSet pBackglass,"high4score","",1,""
      PuPlayer.LabelSet pBackglass,"waf1name","",1,""
      PuPlayer.LabelSet pBackglass,"waf1score","",1,""
      PuPlayer.LabelSet pBackglass,"waf2name","",1,""
      PuPlayer.LabelSet pBackglass,"waf2score","",1,""
      PuPlayer.LabelSet pBackglass,"waf3name","",1,""
      PuPlayer.LabelSet pBackglass,"waf3score","",1,""
      PuPlayer.LabelSet pBackglass,"waf4name","",1,""
      PuPlayer.LabelSet pBackglass,"waf4score","",1,""
    PuPlayer.LabelShowPage pBackglass,1,0,""
    pUpdateScores
    PuPlayer.playlistplayex pBackglass,"scene","base.mov",0,1
    PuPlayer.SetBackground pBackglass,1
    PuPlayer.LabelSet pBackglass,"Play1","PLAYER 1",1,"{'mt':2,'color':16777215, 'size': 1.5, 'xpos': 93.3, 'xalign': 0}"
      PuPlayer.LabelSet pBackglass,"notetitle","",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","",1,""
    ruleshelperon
    MagnetW.MagnetON = False
    beacontime.Enabled = False
    lowerflippersoff = True
    CloseGates
    beaconoff
    Raisewall
    walllt.Collidable = True
    wallmt.Collidable = True
    wallrt.Collidable = True
    help.opacity = 0
    barbMultiball = false
  For i = 0 to 10
    barbHits(i) = 0
  Next
    LookForBarb = False
    ' Run Multiball resets
    RunMultiball = False
    RunAway = False
  For i = 0 to 9
    RunHits(i) = 0
  Next
    finalflips = False
    ' Will Resets
    WillMultiball = False
    WillSuperReady = False
    'Monster Resets
    DemoMultiball = False
    DemoHits = 0
    MonsterFinalBlow = False
    BarbJackpots = False


  For i = 0 to 100
    orbithit(i) = 0
  Next

      resetbackglass
  End Sub

  Sub StopEndOfBallMode()              'this sub is called after the last ball is drained
    ResetSkillShotTimer_Timer
  End Sub

  Sub ResetNewBallVariables()          'reset variables for a new ball or player
    If rockmusic = 1 Then
    PuPlayer.playlistplayex pMusic,"audiobgrock","",soundtrackvol,1
    PuPlayer.SetLoop 4,1
    Else
    PuPlayer.playlistplayex pMusic,"audiobg","",soundtrackvol,1
    PuPlayer.SetLoop 4,1
    End if
  End Sub

  Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
      a.State = 0
    Next
    For each a in aLightsGlow
      a.State = 0
    Next
  End Sub

  Sub ResetNewBallLights() 'turn on or off the needed lights before a new ball is released
    TurnOffPlayfieldLights()
    MagnetU.MagnetON = False
    SetLightColor pool, purple, -1
    SetLightColor pool1, purple, -1
    SetLightColor compass, purple, -1
    SetLightColor compass1, purple, -1
    SetLightColor radio, purple, -1
    SetLightColor radio1, purple, -1
    SetLightColor levitate, purple, -1
    SetLightColor levitate1, purple, -1
    CloseGates
    gate8.open = false
    lbolt1.state = 2
    lbolt7.state = 2
    lbolt2.state = 2
    lbolt8.state = 2
    lbolt3.state = 2
    lbolt10.state = 2
    lbolt4.state = 2
    lbolt9.state = 2
    lbolt5.state = 2
    lbolt11.state = 2
    lbolt6.state = 2
    lbolt12.state = 2
    batteryred.state = 0
    batteryred1.state = 0
    batteryorange.state = 0
    batteryorange1.state = 0
    batteryyellow.state = 0
    batteryyellow1.state = 0
    batterygreen.state = 0
    batterygreen1.state = 0
    avclubready.state = 0
    avclubready1.state = 0
    energy = 0
    kbon.opacity = 0
    kboff.opacity = 0
    If escapegateopen(currentplayer) = 1 Then
      gate8.open = True
    End If
    If barb1on(CurrentPlayer) = 1 Then
      lb1.State = 1
      lb6.State = 1
    End If
    If barb2on(CurrentPlayer) = 1 Then
      lb2.State = 1
      lb5.State = 1
    End If
    If demoisdead(CurrentPlayer) = 1 Then
    lm4.State = 1
    lm6.State = 1
    End If
    If barb3on(CurrentPlayer) = 1 Then
      lb3.State = 1
      lb8.State = 1
    End If
    If barb4on(CurrentPlayer) = 1 Then
      lb4.State = 1
      lb7.State = 1
    End If
    If mysteryready(CurrentPlayer) = 1 Then
      mysterylight.state = 2
      mysterylight1.state = 2
    End If
    If run1on(CurrentPlayer) = 1 Then
      lr1.state = 1
      lr8.state = 1
    End If
    If run2on(CurrentPlayer) = 1 Then
      lr2.state = 1
      lr9.state = 1
    End If
    If run3on(CurrentPlayer) = 1 Then
      lr3.state = 1
      lr12.state = 1
    End If
    If run4on(CurrentPlayer) = 1 Then
      lr4.state = 1
      lr11.state = 1
    End If
    If run5on(CurrentPlayer) = 1 Then
      lr5.state = 1
      lr10.state = 1
    End If
    If run6on(CurrentPlayer) = 1 Then
      lr6.state = 1
      lr7.state = 1
    End If
    If levitatedone(CurrentPlayer) = 1 Then
      SetLightColor levitate, amber, -1
      levitate.state = 1
      SetLightColor levitate1, amber, -1
      levitate1.state = 1
    End If
    If radiodone(CurrentPlayer) = 1 Then
      SetLightColor radio, amber, -1
      radio.state = 1
      SetLightColor radio1, amber, -1
      radio1.state = 1
    End If
    If compassdone(CurrentPlayer) = 1 Then
      SetLightColor compass, amber, -1
      compass.state = 1
      SetLightColor compass1, amber, -1
      compass1.state = 1
    End If
    If pooldone(CurrentPlayer) = 1 Then
      SetLightColor pool, amber, -1
      pool.state = 1
      SetLightColor pool1, amber, -1
      pool1.state = 1
    End If
    If avclubreadyon(CurrentPlayer) = 1 Then
      avclubready.state = 2
      avclubready1.state = 2
    End If
    If barbcompleted(CurrentPlayer) = 1 Then
      lm1.state = 1
      lm9.state = 1
    End If
    If badmensuper(CurrentPlayer) = 1 Then
      lm2.state = 1
      lm8.state = 1
    End If
    If lucason(CurrentPlayer) = 1 Then
      lucas.state = 1
    End If
    If jonathanon(CurrentPlayer) = 1 Then
      jonathan.state = 1
    End If
    If nancyon(CurrentPlayer) = 1 Then
      nancy.state = 1
    End If
    If joyceon(CurrentPlayer) = 1 Then
      joyce.state = 1
    End If
    If willon(CurrentPlayer) = 1 Then
      will.state = 1
    End If
    If elevencrewon(CurrentPlayer) = 1 Then
      elevencrew.state = 1
    End If
    If mikeyon(CurrentPlayer) = 1 Then
      mikey.state = 1
    End If
    If hopperon(CurrentPlayer) = 1 Then
      hopper.state = 1
    End If
    If dustinon(CurrentPlayer) = 1 Then
      dustin.state = 1
    End If
    If partydone(CurrentPlayer) = 1 Then
      lm5.state = 1
      lm10.state = 1
    End If
    If upsidedownlights(CurrentPlayer) = 1 Then
      upsidedownarrow.state = 2
      upsidedowncircle.state = 2
    End If
    If willcompleted(CurrentPlayer) = 1 Then
      lm3.State = 1
      lm7.State = 1
    End If
    If le1on(CurrentPlayer) = 1 Then
      le1.State = 1
      le17.State = 1
    End If
    If le2on(CurrentPlayer) = 1 Then
      le2.State = 1
      le16.State = 1
    End If
    If le3on(CurrentPlayer) = 1 Then
      le3.State = 1
      le18.State = 1
    End If
    If le4on(CurrentPlayer) = 1 Then
      le4.State = 1
      le19.State = 1
    End If
    If le5on(CurrentPlayer) = 1 Then
      le5.State = 1
      le20.State = 1
    End If
    If le6on(CurrentPlayer) = 1 Then
      le6.State = 1
      le21.State = 1
    End If
    If le7on(CurrentPlayer) = 1 Then
      le7.State = 2
      le23.State = 2
    End If
    If le8on(CurrentPlayer) = 1 Then
      le8.State = 2
      le22.State = 2
    End If
    If le9on(CurrentPlayer) = 1 Then
      le9.State = 2
      le24.State = 2
    End If
    If le10on(CurrentPlayer) = 1 Then
      le10.State = 1
      le26.State = 1
    End If
    If le11on(CurrentPlayer) = 1 Then
      le11.State = 1
      le25.State = 1
    End If
    If le12on(CurrentPlayer) = 1 Then
      le12.State = 1
      le27.State = 1
    End If
    If le13on(CurrentPlayer) = 1 Then
      le13.State = 1
      le28.State = 1
    End If
    If le14on(CurrentPlayer) = 1 Then
      le14.State = 1
      le30.State = 1
    End If
    If le15on(CurrentPlayer) = 1 Then
      le15.State = 1
      le29.State = 1
    End If
    If walllon(CurrentPlayer) = 1 Then
      wallll.state = 1
      wallll1.state = 1
    End If
    If wallmon(CurrentPlayer) = 1 Then
      wallml.state = 1
      wallml1.state = 1
    End If
    If wallron(CurrentPlayer) = 1 Then
      wallrl.state = 1
      wallrl1.state = 1
    End If
    If extraballready(CurrentPlayer) = 1 Then
      extraball.state = 2
      extraball1.state = 2
    End If
    If batt1on(CurrentPlayer) = 1 Then
      batteryred.state = 1
      batteryred1.state = 1
      energy = 1
    End If
    If batt2on(CurrentPlayer) = 1 Then
      batteryorange.state = 1
      batteryorange1.state = 1
      energy = 2
    End If
    If batt3on(CurrentPlayer) = 1 Then
      batteryyellow.state = 1
      batteryyellow1.state = 1
      energy = 3
    End If
    If batteryon(CurrentPlayer) = 1 Then
      batteryorange.state = 2
      batteryorange1.state = 2
      batterygreen.state = 2
      batterygreen1.state = 2
      batteryred.state = 2
      batteryred1.state = 2
      batteryyellow.state = 2
      batteryyellow1.state = 2
      energy = 4
    End If
    If PlayersPlayingGame > 1 Then
      resetawall
    End If

    If kickbackon(CurrentPlayer) = 1 Then
      kickbacklight.state = 1
      kickbacklight1.state = 1
      kickbackgate.open = True
      kbon.opacity = 1000
      kboff.opacity = 0
    End If
    If partyready(CurrentPlayer) = 1 Then
      partylock.state = 2
      partylock1.state = 2
    End if

    movemodes
    currentplayerbackglass
    CheckBARBTargets
    CheckRunTargets
    CheckESCAPETargets
    udtarget.TransZ = -22
    udtargetdown(CurrentPlayer) = 0
    udtarget.Collidable = True
    udtargettime.Enabled = False
  End Sub

  Sub resetawall
      If barricadedown(CurrentPlayer) = 0 Then
        If wallll.State + wallml.State + wallrl.State = 3 Then
          Dropwall
        Else
          Raisewall
        End If
      Else
        If barricadedowncurrently = False Then
          Dropwall
        End If
      End If
  End Sub

  Sub startamultiball
    Dim waittime
    waittime = 1000
    vpmtimer.addtimer waittime, "closeupshop'"
  End Sub

  Sub closeupshop
    barbgate.open = False
    BallLockEscape.enabled = False
    diverter.RotZ = 0
    diverteropen.collidable = False
    diverter.collidable = True
    boltsoff
  End Sub

  Sub endamultiball
    CheckBARBTargets
    checkhits
    CheckKIDSLane
    checkparty
    CheckRunTargets
    If llo2.State = 2 Then
      diverter.RotZ = -33
      diverteropen.collidable = True
      diverter.collidable = False
    End If
    boltson
  End Sub


  Sub UpdateSkillShot() 'Updates the skillshot light
    LightSeqSkillshot.Play SeqAllOff
    lr1.State = 2
    lr8.State = 2
    lr6.state = 2
    lr7.state = 2
  End Sub

  Sub SkillshotOff_Hit 'trigger to stop the skillshot due to a weak plunger shot
    If bSkillShotReady Then
      ResetSkillShotTimer_Timer
    End If
  End Sub

  Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
    ResetSkillShotTimer.Enabled = 0
    bSkillShotReady = False
    LightSeqSkillshot.StopPlay
    If lr1.State = 2 Then lr1.State = 0
    If lr6.State = 2 Then lr6.State = 0
    lr7.state = 0
    lr8.state = 0

    If run1on(CurrentPlayer) = 1 Then
      lr1.state = 1
      lr8.state = 1
    End If
    If run6on(CurrentPlayer) = 1 Then
      lr6.state = 1
      lr7.state = 1
    End If
  End Sub

  Sub spinsign
    spinthesign.enabled = 1
  End Sub

  Dim spinspot
  spinspot = 0

  Sub spinthesign_timer
    sign.RotY = sign.RotY + 1
    spinspot = spinspot + 1
    Select Case spinspot
      case 500
      playsoundat "fx_sign", sign
      spinthesign.enabled = 0
      spinspot = 0
    End Select
  End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  SECONDARY HIT EVENTS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'

  Dim LStep, RStep

  Sub LeftSlingShot_Slingshot
        LS.VelocityCorrect(Activeball)
    If Tilted Then Exit Sub
    movemodes
    LightEffect 7
    PlaySoundAt SoundFXDOF("Sling_L1", 103, DOFPulse, DOFContactors), lane3
    PlaySound "whomp"
    DOF 104, DOFPulse 'DOF OuterLeft Flasher Red
    DOF 300, DOFPulse   'DOF MX - Left Slingshot
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 110
    ' add some effect to the table?
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot"
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
        RS.VelocityCorrect(Activeball)
    If Tilted Then Exit Sub
    LightEffect 6
    movemodes
    PlaySoundAt SoundFXDOF("Sling_R1", 105, DOFPulse, DOFContactors), lane5
    PlaySound "whomp"
    DOF 106, DOFPulse 'DOF OuterRight Flasher Red
    DOF 301, DOFPulse   'DOF MX - Right Slingshot
    RightSling4.Visible = 1
    Lemk1.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 110
    ' add some effect to the table?
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot"
  End Sub

  Sub RightSlingShot_Timer
    Select Case RStep
      Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Lemk1.RotX = 14
      Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Lemk1.RotX = 2
      Case 3:RightSLing2.Visible = 0:Lemk1.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select

    RStep = RStep + 1
  End Sub

  '*********
  ' Gates
  '*********
  ' lets get these gates all closed up.

    Gate8.Open = False


  Sub CloseGates
    diverter.RotZ = 0
    diverteropen.collidable = False
    diverter.collidable = True
    barbgate.Open = False

  End Sub


  '***************************
  'block the upside down during multiballs
  '***************************

  Sub udblock_hit
  'Gate7.Open = True
  Gate9.Open = True
  If barbMultiball = True Then
  Gate9.Open = False
  End If
  If RunMultiball = True Then
  Gate9.Open = False
  End If
  End Sub

  '***********
  ' Spinner
  '***********
  ' var spinvalue
  ' spinner increase its value by 1000 each time is hit
  ' max value is 10.000

  Sub Spinner1_Spin
    DOF 131, DOFPulse
    DOF 350, DOFPulse   'DOF MX - Left Spinner
    PlaySoundAt "fx_spinner", Spinner1
    If Not Tilted Then
      ' any light effect?
      ' any ''''DMD display?
      AddScore spinvalue
    End If
  End Sub

  Sub SpinCounterLO_Hit
    If Not Tilted Then
      If spinvalue <10000 Then
        spinvalue = spinvalue + 1000
      End If
    End If
  End Sub

  Sub Spinner2_Spin
    DOF 130, DOFPulse
    DOF 351, DOFPulse   'DOF MX - Right Spinner
    PlaySoundAt "fx_spinner", Spinner2
    If Not Tilted Then
      ' any light effect?
      ' any ''''DMD display?
      AddScore spinvalue
    End If
  End Sub

  Sub SpinCounterRO_Hit
    If Not Tilted Then
      If spinvalue <10000 Then
        spinvalue = spinvalue + 1000
      End If
    End If
  End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  MAIN SHOTS - PRIMARY HIT EVENTS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'

  Dim quotenum:quotenum = 0
  Sub randomquote
    quotenum = RndNum(1,5)
    debug.print "Q:" & quotenum
    Select Case quotenum
      Case 1
        PuPlayer.playlistplayex pBackglass,"videoquotes","",100,1
      Case 2
        RandomLightQuote
    End Select
  End Sub

  '****************
  '     ORBITS
  '****************

  Sub leftorbitdone_Hit
    Addscore 7250
    FlashLevel2 = 1 : Flasherflash2_Timer
    OrbitAward
    If bMultiBallMode = False Then
      PlaySound "ro2"
      randomquote
    End If

    If partymultiball = False Then
      If jonathan.state = 0 Then
        jonathan.state = 1
        jonathanon(CurrentPlayer) = 1
        partycollected(CurrentPlayer) = partycollected(CurrentPlayer) + 1
        PuPlayer.LabelSet pBackglass,"partyh","" & partycollected(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
        checkparty
      End If
    End If
    If jonathan.state = 2 Then
      jonathan.state = 0
      awardparty
    End If

    If Tilted Then Exit Sub

    If barbMultiball = True Then
      If llo1.state = 2 Then
        If BSuper(CurrentPlayer) = 1 Then
          barbfound
        Else
          AwardBarb
        End If
        If BSuper(CurrentPlayer) = 0 Then
        Else
        llo1.state = 0
        llo9.state = 0
        End If
      End If
    End If
  End Sub'

  Sub leftorbitdone1_hit
    FlashLevel1 = 1 : Flasherflash1_Timer
  End Sub


  Sub rightorbitdone_Hit
    Addscore 7250
    PlaySound "steve"
    FlashLevel4 = 1 : Flasherflash4_Timer
    If Tilted Then Exit Sub

    If partymultiball = False Then
      If dustin.state = 0 Then
        dustin.state = 1
        dustinon(CurrentPlayer) = 1
        partycollected(CurrentPlayer) = partycollected(CurrentPlayer) + 1
        PuPlayer.LabelSet pBackglass,"partyh","" & partycollected(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
        checkparty
      End If
    End If
    If dustin.state = 2 Then
      dustin.state = 0
      awardparty
    End If

    If bMultiBallMode = False Then
      randomquote
    End If

    If barbMultiball = True Then
      If lro4.state = 2 Then
        If BSuper(CurrentPlayer) = 5 Then
          barbfound
        Else
          AwardBarb
        End If
        If BSuper(CurrentPlayer) = 0 Then
        Else
        lro4.state = 0
        lro5.state = 0
        End If
      End If
    End If
  End Sub


  dim orbithit(100)
  Sub OrbitAward
    orbithit(CurrentPlayer) = orbithit(CurrentPlayer) + 1
    Select Case orbithit(CurrentPlayer)
        Case 1
           DMD "black.png", "30 More Shots", "For Super Orbits",  1000
        Case 31
           DMD "black.png", "Super Orbits", "1000000 Award",  1000
          AddScore 1000000
        Case 34
           DMD "black.png", "30 More Shots", "For Double Super Orbits",  1000
        Case 64
           DMD "black.png", "Double Super", "Orbits 2000000 Award",  1000
          AddScore 2000000
        Case 68
           DMD "black.png", "30 More Shots", "For Ultra Orbits",  1000
        Case 98
           DMD "black.png", "Ultra Orbits", "4000000 Award",  1000
          AddScore 4000000
          Orbithit(CurrentPlayer) = 0
    End Select
  END Sub

  '****************
  '     RAMPS
  '****************

  Sub leftrampdone_Hit
    Addscore 13333
    If WillMultiball = True Then
      If llo3.state = 2 Then
        AwardWill
        DOF 380, DOFPulse  ' DOF MX - Left Ramp Green
      End If
    End If

    DOF 142, DOFPulse
    LightEffect 12
    FlashLevel1 = 1 : Flasherflash1_Timer
    PlaySoundAt "fx_metalrolling", ActiveBall

    If Tilted Then Exit Sub

    If RunMultiball = True Then
    AwardRun
    DOF 381, DOFPulse  ' DOF MX - Left Ramp Cyan
    End If
    If partymultiball = False Then
      If nancy.state = 0 Then
        nancy.state = 1
        nancyon(CurrentPlayer) = 1
        partycollected(CurrentPlayer) = partycollected(CurrentPlayer) + 1
        PuPlayer.LabelSet pBackglass,"partyh","" & partycollected(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
        checkparty
      End If
    End If
    If nancy.state = 2 Then
      nancy.state = 0
      awardparty
      DOF 382, DOFPulse  ' DOF MX - Left Ramp Orange
    End If

    If barbMultiball = True Then
      If leftrampred.state = 2 Then
        If BSuper(CurrentPlayer) = 2 Then
          barbfound
        Else
          AwardBarb
          DOF 383, DOFPulse  ' DOF MX - Left Ramp Yellow
        End If
        If BSuper(CurrentPlayer) = 0 Then
        Else
        leftrampred.state = 0
        leftrampred2.state = 0
        End If
      End If
    End If

    'modes
    If llo4.state = 2 Then
      DOF 384, DOFPulse  ' DOF MX - Left Ramp Purple
      If levitateactive = 1 Then
        levitateaward
      End If
      If radioactive = 1 Then
        radioaward
      End If
      If poolactive = 1 Then
        poolaward
      End If
      If compassactive = 1 Then
        llo4.state = 0
        llo6.state = 0
        compasscheck
        compassaward
      End If
    End If


    If bMultiBallMode = False Then
      DOF 385, DOFPulse  ' DOF MX - Left Ramp Red
      randomquote
      If lbolt1.state = 1 Then
        lbolt2.state = 1
        lbolt8.state = 1
        addenergy
      End If
      If lbolt1.state = 2 Then
        lbolt1.state = 1
        lbolt7.state = 1
        addenergy
      End If
    End If
  End Sub


  Sub centerrampdone_Hit
    Addscore 16666
    If WillMultiball = True Then
      If lc1.state = 2 Then
        AwardWill
        DOF 380, DOFPulse  ' DOF MX - Center Ramp Green
      End If
    End If
    DOF 142, DOFPulse
    LightEffect 12
    FlashLevel2 = 1 : Flasherflash2_Timer
    PlaySoundAt "fx_metalrolling", ActiveBall

    If Tilted Then Exit Sub

    If RunMultiball = True Then
      If RunAway = True Then
        RUNSuper
        DOF 381, DOFPulse  ' DOF MX - Center Ramp Cyan
      End If
    End If
    If partymultiball = False Then
      If joyce.state = 0 Then
        joyce.state = 1
        joyceon(CurrentPlayer) = 1
        partycollected(CurrentPlayer) = partycollected(CurrentPlayer) + 1
        PuPlayer.LabelSet pBackglass,"partyh","" & partycollected(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
        checkparty
      End If
    End If
    If joyce.state = 2 Then
      joyce.state = 0
      awardparty
      DOF 382, DOFPulse  ' DOF MX - Center Ramp Orange
    End If

    If barbMultiball = True Then
      If leftrampred1.state = 2 Then
        If BSuper(CurrentPlayer) = 3 Then
          barbfound
        Else
          AwardBarb
          DOF 383, DOFPulse  ' DOF MX - Center Ramp Yellow
        End If
        If BSuper(CurrentPlayer) = 0 Then
        Else
        leftrampred1.state = 0
        leftrampred3.state = 0
        End If
      End If
    End If

    If bMultiBallMode = False Then
      DOF 385, DOFPulse  ' DOF MX - Center Ramp Red
      randomquote
      If lbolt3.state = 1 Then
        lbolt4.state = 1
        lbolt9.state = 1
        addenergy
      End If
      If lbolt3.state = 2 Then
        lbolt3.state = 1
        lbolt10.state = 1
        addenergy
      End If
    End If

    If lc2.state = 2 Then
      DOF 384, DOFPulse  ' DOF MX - Center Ramp Purple
      If levitateactive = 1 Then
        levitateaward
      End If
      If radioactive = 1 Then
        radioaward
      End If
      If poolactive = 1 Then
        poolaward
      End If
      If compassactive = 1 Then
        lc2.state = 0
        lc3.state = 0
        compasscheck
        compassaward
      End If
    End If

  End Sub


  Sub rightrampdone_Hit
    Addscore 13333
    If WillMultiball = True Then
      If llr1.state = 2 Then
        DOF 386, DOFPulse  ' DOF MX - Right Ramp Green
        AwardWill
      End If
    End If
    DOF 143, DOFPulse
    LightEffect 13
    FlashLevel4 = 1 : Flasherflash4_Timer
    FlashLevel6 = 1 : Flasherflash6_Timer
    PlaySoundAt "fx_metalrolling", ActiveBall
    PlaySound "portalopen"

    If bMultiBallMode = False Then
      DOF 391, DOFPulse  ' DOF MX - Right Ramp Red
      randomquote
      If lbolt5.state = 1 Then
        lbolt6.state = 1
        lbolt12.state = 1
        addenergy
      End If
      If lbolt5.state = 2 Then
        lbolt5.state = 1
        lbolt11.state = 1
        addenergy
      End If
    End If

    If Tilted Then Exit Sub

    If RunMultiball = True Then
    DOF 387, DOFPulse  ' DOF MX - Right Ramp Cyan
    AwardRun
    End If
    If partymultiball = False Then
      If hopper.state = 0 Then
        hopper.state = 1
        hopperon(CurrentPlayer) = 1
        partycollected(CurrentPlayer) = partycollected(CurrentPlayer) + 1
        PuPlayer.LabelSet pBackglass,"partyh","" & partycollected(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
        checkparty
      End If
    End If
    If hopper.state = 2 Then
      hopper.state = 0
      awardparty
      DOF 388, DOFPulse  ' DOF MX - Right Ramp Orange
    End If

    If barbMultiball = True Then
      If rightrampred.state = 2 Then
        If BSuper(CurrentPlayer) = 4 Then
          barbfound
        Else
          AwardBarb
          DOF 389, DOFPulse  ' DOF MX - Right Ramp Yellow
        End If
        If BSuper(CurrentPlayer) = 0 Then
        Else
        rightrampred.state = 0
        rightrampred1.state = 0
        End If
      End If
    End If

    ' modes
    If llo5.state = 2 Then
      DOF 390, DOFPulse  ' DOF MX - Right Ramp Purple
      If levitateactive = 1 Then
        levitateaward
      End If
      If radioactive = 1 Then
        radioaward
      End If
      If poolactive = 1 Then
        poolaward
      End If
      If compassactive = 1 Then
        llo5.state = 0
        llo10.state = 0
        compasscheck
        compassaward
      End If
    End If
  End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  MODES
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'

  Sub boltsoff
    lbolt1.state = 0
    lbolt7.state = 0
    lbolt2.state = 0
    lbolt8.state = 0
    lbolt3.state = 0
    lbolt10.state = 0
    lbolt4.state = 0
    lbolt9.state = 0
    lbolt5.state = 0
    lbolt11.state = 0
    lbolt6.state = 0
    lbolt12.state = 0
  End Sub

  Sub boltson
    If avclubready.state = 2 Then Exit Sub
    lbolt1.state = 2
    lbolt7.state = 2
    lbolt2.state = 2
    lbolt8.state = 2
    lbolt3.state = 2
    lbolt10.state = 2
    lbolt4.state = 2
    lbolt9.state = 2
    lbolt5.state = 2
    lbolt11.state = 2
    lbolt6.state = 2
    lbolt12.state = 2
  End Sub

  Sub addenergy
    energy = energy +1
    Select Case energy
      Case 1
      batteryred.state = 1
      batteryred1.state = 1
      batt1on(CurrentPlayer) = 1
      Case 2
      batteryorange.state = 1
      batteryorange1.state = 1
      batt2on(CurrentPlayer) = 1
      Case 3
      batteryyellow.state = 1
      batteryyellow1.state = 1
      batt3on(CurrentPlayer) = 1
      Case 4
      batterygreen.state = 2
      batterygreen1.state = 2
      batteryorange.state = 2
      batteryorange1.state = 2
      batteryyellow.state = 2
      batteryyellow1.state = 2
      batteryred.state = 2
      batteryred1.state = 2
      avclubready.state = 2
      avclubready1.state = 2
      avclubreadyon(CurrentPlayer) = 1
      batteryon(CurrentPlayer) = 1
      batt1on(CurrentPlayer) = 0
      batt1on(CurrentPlayer) = 0
      batt1on(CurrentPlayer) = 0
      lbolt1.state = 0
      lbolt7.state = 0
      lbolt2.state = 0
      lbolt8.state = 0
      lbolt3.state = 0
      lbolt10.state = 0
      lbolt4.state = 0
      lbolt9.state = 0
      lbolt5.state = 0
      lbolt11.state = 0
      lbolt6.state = 0
      lbolt12.state = 0
      PuPlayer.playlistplayex pCallouts,"audiocallouts","theavclubisnowopen.wav",100,1
    chilloutthemusic
      pNote "A.V CLUB","NOW OPEN"
      PuPlayer.playlistplayex pBackglass,"videoavopen","",100,1
    End Select

  End Sub

  Sub extraballmode
    extraball.state = 2
    extraball1.state = 2
    extraballready(CurrentPlayer) = 1
    PuPlayer.playlistplayex pCallouts,"audiocallouts","extraballislit.wav",100,1
    chilloutthemusic
    PuPlayer.playlistplayex pBackglass,"videoextraball","extraballlit.mov",100,1
  End Sub


  Sub Kicker1_hit

    waittime = 1000
    RandomSoundHole
    If partymultiball = False Then
      If elevencrew.state = 0 Then
        elevencrew.state = 1
        elevencrewon(CurrentPlayer) = 1
        partycollected(CurrentPlayer) = partycollected(CurrentPlayer) + 1
        PuPlayer.LabelSet pBackglass,"partyh","" & partycollected(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
        checkparty
      End If
    End If

    If extraball.state = 2 and avclubready.state = 0 Then
      extraballskip = 1
      If bMultiBallMode = False Then
        waittime = 12000
      Else
        waittime = 1000
      End If
      AwardExtraBall
      extraball.state = 0
      vpmtimer.addtimer waittime, "exitav'"
    End if

    If elevencrew.state = 2 Then
      elevencrew.state = 0
      awardparty
    End If


    If bMultiBallMode = False Then
      Dim waittime
      If avclubready.state = 2 Then
        If activemode = 1 Then
          If extraball.state = 2 Then
            AwardExtraBall
            extraball.state = 0
            waittime = 10000
            vpmtimer.addtimer waittime, "levitateprestart'"
          Else
            levitateprestart
          End If
        End If
        If activemode = 2 Then
          If extraball.state = 2 Then
            AwardExtraBall
            extraball.state = 0
            waittime = 10000
            vpmtimer.addtimer waittime, "poolprestart'"
          Else
            poolprestart
          End If
        End If
        If activemode = 3 Then
          If extraball.state = 2 Then
            AwardExtraBall
            extraball.state = 0
            waittime = 10000
            vpmtimer.addtimer waittime, "radioprestart'"
          Else
            radioprestart
          End If
        End If
        If activemode = 4 Then
          If extraball.state = 2 Then
            AwardExtraBall
            extraball.state = 0
            waittime = 10000
            vpmtimer.addtimer waittime, "compassprestart'"
          Else
            compassprestart
          End If
        End If
      Else
        vpmtimer.addtimer waittime, "exitav'"
      End If
    Else
      vpmtimer.addtimer waittime, "exitav'"
    End If
  End Sub

  Sub levitateprestart
    PuPlayer.playpause 4
    Dim waittime
    avclubready.state = 0
    avclubready1.state = 0
    avclubreadyon(CurrentPlayer) = 0
    ruleshelperoff
    GiOff
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Levitation Mode",1,""
    PuPlayer.LabelSet pBackglass,"notecopy","Quickshot score builder- Scores build every shot you make \r Only 20 seconds to get the shot, hurry keep it levitating",1,""
    waittime = 15000
    PuPlayer.playlistplayex pBackglass,"videolevitate","levitatestart.mov",100,1
    vpmtimer.addtimer waittime, "levitatestart'"
    levitateskip = 1
  End Sub

  Sub poolprestart
    PuPlayer.playpause 4
    Dim waittime
    waittime = 8000
    avclubready.state = 0
    avclubready1.state = 0
    avclubreadyon(CurrentPlayer) = 0
    ruleshelperoff
    GiOff
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Sensory Deprevation Pool",1,""
    PuPlayer.LabelSet pBackglass,"notecopy","Help 11 fill the pool and locate Will \r You have 90 seconds to hit 10 shots and locate Will",1,""
    PuPlayer.playlistplayex pBackglass,"videopool","poolstart.mov",100,1
    vpmtimer.addtimer waittime, "poolstart'"
    poolskip = 1
  End Sub

  Sub radioprestart
    PuPlayer.playpause 4
    Dim waittime
    waittime = 18000
    avclubready.state = 0
    avclubready1.state = 0
    avclubreadyon(CurrentPlayer) = 0
    ruleshelperoff
    GiOff
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Radio Mode",1,""
    PuPlayer.LabelSet pBackglass,"notecopy","Catch the roving dots to find the signal to Will \r You have 60 seconds to hit 5 shots and locate Will",1,""
    PuPlayer.playlistplayex pBackglass,"videoradio","radiostart.mov",100,1
    vpmtimer.addtimer waittime, "radiostart'"
    radioskip = 1
  End Sub

  Sub compassprestart
    PuPlayer.playpause 4
    Dim waittime
    avclubready.state = 0
    avclubready1.state = 0
    avclubreadyon(CurrentPlayer) = 0
    ruleshelperoff
    GiOff
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Compass Mode",1,""
    PuPlayer.LabelSet pBackglass,"notecopy","Clear the purple dots 2x within 60 seconds to find the portal \r Each shot scores 1,000,000 points",1,""
    waittime = 22000
    PuPlayer.playlistplayex pBackglass,"videocompass","compassstart.mov",100,1
    vpmtimer.addtimer waittime, "compassstart'"
    compassskip = 1
  End Sub

  Sub exitav
    FlashLevel6 = 1 : Flasherflash6_Timer
    Dim waittime
    waittime = 400
    vpmtimer.addtimer waittime, "exitav2'"
    DOF 361, DOFPulse  'DOF MX - Center Blue Flasher
    PlaySound "zing"
    extraballskip = 0
  End Sub

  Sub exitav2
    FlashLevel6 = 1 : Flasherflash6_Timer
    Dim waittime
    waittime = 400
    vpmtimer.addtimer waittime, "exitav3'"
    DOF 361, DOFPulse  'DOF MX - Center Blue Flasher
    PlaySound "zing"
  End Sub

  Sub exitav3
    FlashLevel6 = 1 : Flasherflash6_Timer
    kicker1.Kick 200, 20
    PlaySoundAt SoundFXDOF("fx_kicker", 200, DOFPulse, DOFContactors), kicker1
    DOF 361, DOFPulse  'DOF MX - Center Blue Flasher
    DOF 115, DOFPulse
    PlaySound "zing"
  End Sub

  Dim modenum
  Modenum = 1
  Dim activemode

  Sub movemodes
    If DemoMultiball = True Then Exit Sub
    If levitate.state + pool.state + radio.state + compass.state = 4 Then
      Exit Sub
    End If
    modenum = modenum + 1
    Select Case modenum
      Case 1
        If levitate.state = 0 Then
          levitate.state = 1
          levitate1.state = 1
          activemode = 1
          If compassdone(CurrentPlayer) = 0 Then
            compass.state = 0
            compass1.state = 0
          End If
          If radiodone(CurrentPlayer) = 0 Then
            radio.state = 0
            radio1.state = 0
          End If
          If pooldone(CurrentPlayer) = 0 Then
            pool.state = 0
            pool1.state = 0
          End If
        Else
          movemodes
        End If
      Case 2
        If pool.state = 0 Then
          pool.state = 1
          pool1.state = 1
          activemode = 2
          If compassdone(CurrentPlayer) = 0 Then
            compass.state = 0
            compass1.state = 0
          End If
          If radiodone(CurrentPlayer) = 0 Then
            radio.state = 0
            radio1.state = 0
          End If
          If levitatedone(CurrentPlayer) = 0 Then
            levitate.state = 0
            levitate1.state = 0
          End If
        Else
          movemodes
        End If
      Case 3
        If radio.state = 0 Then
          radio.state = 1
          radio1.state = 1
          activemode = 3
          If compassdone(CurrentPlayer) = 0 Then
            compass.state = 0
            compass1.state = 0
          End If
          If levitatedone(CurrentPlayer) = 0 Then
            levitate.state = 0
            levitate1.state = 0
          End If
          If pooldone(CurrentPlayer) = 0 Then
            pool.state = 0
            pool1.state = 0
          End If
        Else
          movemodes
        End If
      Case 4
        If compass.state = 0 Then
          compass.state = 1
          compass1.state = 1
          activemode = 4
          If levitatedone(CurrentPlayer) = 0 Then
            levitate.state = 0
            levitate1.state = 0
          End If
          If radiodone(CurrentPlayer) = 0 Then
            radio.state = 0
            radio1.state = 0
          End If
          If pooldone(CurrentPlayer) = 0 Then
            pool.state = 0
            pool1.state = 0
          End If
        Else
          movemodes
        End If
      Case 5
        modenum = 0
        movemodes
    End Select
  End Sub

  Sub resetmodecheck
    If compassdone(CurrentPlayer) + radiodone(CurrentPlayer) + pooldone(CurrentPlayer) + levitatedone(CurrentPlayer) = 4 Then
      compassdone(CurrentPlayer) = 0
      compass.state = 0
      compass1.state = 0
      radiodone(CurrentPlayer) = 0
      radio.state = 0
      radio1.state = 0
      pooldone(CurrentPlayer) = 0
      pool.state = 0
      pool1.state = 0
      levitatedone(CurrentPlayer) = 0
      levitate.state = 0
      levitate1.state = 0
      movemodes
    End If

    lbolt1.state = 2
    lbolt7.state = 2
    lbolt2.state = 2
    lbolt8.state = 2
    lbolt3.state = 2
    lbolt10.state = 2
    lbolt4.state = 2
    lbolt9.state = 2
    lbolt5.state = 2
    lbolt11.state = 2
    lbolt6.state = 2
    lbolt12.state = 2
    energy = 0
    movemodes
    batterygreen.state = 0
    batterygreen1.state = 0
    batteryorange.state = 0
    batteryorange1.state = 0
    batteryred.state = 0
    batteryred1.state = 0
    batteryyellow.state = 0
    batteryyellow1.state = 0
    llo4.state = 0
    llo6.state = 0
    lc2.state = 0
    lc3.state = 0
    llo5.state = 0
    llo10.state = 0
  End Sub




  '****************
  '     LEVITATE
  '****************
  ' timed quick shot score builder
  ' one shot at a time - go as long as you can.20 sec intervals

  Dim levitateactive
    levitateactive = 0

  Sub levitatestart
    If levitateskip = 0 then exit Sub
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Levitation Mode | 0/10",1,""
    avsdone(CurrentPlayer) = avsdone(CurrentPlayer) + 1
    exitav
    BallLockEscape.enabled = False
    inmode = 1
    pNote "A.V CLUB","LEVITATE MODE"
    PuPlayer.playlistplayex pBackglass,"videolevitate","levitate.mov",100,2
    PuPlayer.SetLoop 2,1
    GiOff
    GiPurple
    PuPlayer.playpause 4
    PuPlayer.playlistplayex pAudio,"audiomodes","levitate.mp3",100,1
    PuPlayer.SetLoop 7,1
    levitatetime.enabled = True
    batterygreen.state = 0
    batterygreen1.state = 0
    batteryorange.state = 0
    batteryorange1.state = 0
    batteryyellow.state = 0
    batteryyellow1.state = 0
    batteryred.state = 0
    batteryred1.state = 0
    batteryon(CurrentPlayer) = 0
    batt1on(CurrentPlayer) = 0
    batt2on(CurrentPlayer) = 0
    batt3on(CurrentPlayer) = 0
    levitateactive = 1
    levitatemovelight
    levitateskip = 0
  End Sub

  Dim levitatepos
  levitatepos = 0

  Sub levitatetime_Timer()
    levitatepos = levitatepos + 1
    PuPlayer.LabelSet pBackglass,"modetimer",20 - levitatepos,1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 80.7, 'xalign': 1, 'ypos': 75, 'yalign': 0}"
    Select Case levitatepos
      Case 1

      Case 20
        levitateend
        pNote "LEVITATE MODE","OVER"
        PuPlayer.playlistplayex pBackglass,"videolevitate","levitateend.mov",100,1
        Dim waittime
        waittime = 2000
        vpmtimer.addtimer waittime, "levitateend'"
    End Select
  End Sub

  Dim levitatelight
  levitatelight = 0

  Sub levitatemovelight
    levitatelight = levitatelight + 1
    Select Case levitatelight
      Case 1
        llo4.state = 2
        llo6.state = 2
        lc2.state = 0
        lc3.state = 0
        llo5.state = 0
        llo10.state = 0
      Case 2
        llo4.state = 0
        llo6.state = 0
        lc2.state = 2
        lc3.state = 2
        llo5.state = 0
        llo10.state = 0
      Case 3
        llo4.state = 0
        llo6.state = 0
        lc2.state = 0
        lc3.state = 0
        llo5.state = 2
        llo10.state = 2
        levitatelight = 0
    End Select
  End Sub

  Dim levitateawarded
  levitateawarded = 0

  Sub levitateaward
    levitateawarded = levitateawarded + 1
    levitatepos = -1
    Select Case levitateawarded
      Case 1
        AddScore 200000
        levitatemovelight
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "LEVITATION EXTENDED","200,000"
        PuPlayer.playlistplayex pCallouts,"audiocallouts","timeadded.wav",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Levitation Mode | 1/10",1,""
      Case 2
        AddScore 400000
        levitatemovelight
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "LEVITATION EXTENDED","400,000"
        PuPlayer.playlistplayex pCallouts,"audiocallouts","timeadded.wav",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Levitation Mode | 2/10",1,""
      Case 3
        AddScore 600000
        levitatemovelight
        pNote "LEVITATION EXTENDED","600,000"
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        PuPlayer.playlistplayex pCallouts,"audiocallouts","timeadded.wav",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Levitation Mode | 3/10",1,""
      Case 4
        AddScore 800000
        levitatemovelight
        pNote "LEVITATION EXTENDED","800,000"
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        PuPlayer.playlistplayex pCallouts,"audiocallouts","timeadded.wav",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Levitation Mode | 4/10",1,""
      Case 5
        AddScore 1000000
        levitatemovelight
        pNote "LEVITATION EXTENDED","1,000,000"
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        PuPlayer.playlistplayex pCallouts,"audiocallouts","timeadded.wav",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Levitation Mode | 5/10",1,""
      Case 6
        AddScore 1200000
        levitatemovelight
        pNote "LEVITATION EXTENDED","1,200,000"
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        PuPlayer.playlistplayex pCallouts,"audiocallouts","timeadded.wav",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Levitation Mode | 6/10",1,""
      Case 7
        AddScore 2000000
        levitatemovelight
        pNote "LEVITATION EXTENDED","2,000,000"
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        PuPlayer.playlistplayex pCallouts,"audiocallouts","timeadded.wav",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Levitation Mode | 7/10",1,""
      Case 8
        AddScore 3000000
        levitatemovelight
        pNote "LEVITATION EXTENDED","3,000,000"
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        PuPlayer.playlistplayex pCallouts,"audiocallouts","timeadded.wav",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Levitation Mode | 8/10",1,""
      Case 9
        AddScore 4000000
        levitatemovelight
        pNote "LEVITATION EXTENDED","4,000,000"
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        PuPlayer.playlistplayex pCallouts,"audiocallouts","timeadded.wav",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Levitation Mode | 9/10",1,""
      Case 10
        AddScore 8000000
        levitatemovelight
        pNote "LEVITATION LEARNED","8,000,000"
        DOF 422, DOFPulse  ' DOF MX - AV Club Mode Complete
        PuPlayer.playlistplayex pBackglass,"videolevitate","levitateend.mov",100,1
    chilloutthemusic
        PuPlayer.playlistplayex pCallouts,"audiocallouts","levitationcomplete.wav",100,1
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Levitation Mode | 10/10",1,""
        levitateend
    End Select
  End Sub

  Sub levitateend

    inmode = 0
    PuPlayer.LabelSet pBackglass,"modetimer"," ",1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 80.7, 'xalign': 1, 'ypos': 75, 'yalign': 0}"
    checkhits
    ruleshelperon
    PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,2
    PuPlayer.SetLoop 2,0
    PuPlayer.LabelSet pBackglass,"modetimer","",1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 80.7, 'xalign': 1, 'ypos': 75, 'yalign': 0}"
    If bMultiBallMode = False Then
    GiOff
    GiOn
    End If
    PuPlayer.playresume 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1
    PuPlayer.SetLoop 7,0
    levitatetime.enabled = False
    levitatemove.enabled = False
    levitatepos = 0
    levitatelight = 0
    levitateawarded = 0
    levitateactive = 0
    levitatedone(CurrentPlayer) = 1
    SetLightColor levitate, amber, -1
    levitate.state = 1
    SetLightColor levitate1, amber, -1
    levitate1.state = 1
    resetmodecheck

    If avsdone(CurrentPlayer) = 2 and extraball.state = 0 Then
      extraballmode
    End If
    PuPlayer.LabelSet pBackglass,"mav","" & avsdone(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 94.6, 'yalign': 0}"
  End Sub



  '****************
  '     POOL
  '****************
  ' all dots lit, small points - 10 shots fill the pool
  ' 45 seconds to complete
  Dim poolactive
  poolactive = 0

  Sub poolstart
    If poolskip = 0 then exit Sub
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Sensory Deprevation Pool | 0/10",1,""
    exitav
    BallLockEscape.enabled = False
    inmode = 1
    pNote "A.V CLUB","SENSORY DEPRIVATION POOL"
    PuPlayer.playlistplayex pBackglass,"videopool","pool-all.mov",100,2
    PuPlayer.SetLoop 2,1
    GiOff
    GiPurple
    PuPlayer.playpause 4
    PuPlayer.playlistplayex pAudio,"audiomodes","pool.mp3",100,1
    PuPlayer.SetLoop 7,1
    pooltime.enabled = True
    llo4.state = 2
    llo6.state = 2
    lc2.state = 2
    lc3.state = 2
    llo5.state = 2
    llo10.state = 2
    poolactive = 1
    batterygreen.state = 0
    batterygreen1.state = 0
    batteryorange.state = 0
    batteryorange1.state = 0
    batteryyellow.state = 0
    batteryyellow1.state = 0
    batteryred.state = 0
    batteryred1.state = 0
    batteryon(CurrentPlayer) = 0
    batt1on(CurrentPlayer) = 0
    batt2on(CurrentPlayer) = 0
    batt3on(CurrentPlayer) = 0
    avsdone(CurrentPlayer) = avsdone(CurrentPlayer) + 1
    poolskip = 0
  End Sub

  Dim poolpos
  poolpos = 0

  Sub pooltime_Timer()
    PuPlayer.LabelSet pBackglass,"modetimer",90 - poolpos,1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 80.7, 'xalign': 1, 'ypos': 75, 'yalign': 0}"
    poolpos = poolpos + 1
    Select Case poolpos
      Case 1

      Case 90
        pNote "POOL MODE","OVER"
        poolend
    End Select
  End Sub

  Dim poolawarded
  poolawarded = 0

  Sub poolaward
    poolawarded = poolawarded + 1
    Select Case poolawarded
      Case 1
        AddScore 400000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "POOL FILLED","400,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Sensory Deprevation Pool | 1/10",1,""
      Case 2
        AddScore 400000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "POOL FILLED","400,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Sensory Deprevation Pool | 2/10",1,""
      Case 3
        AddScore 400000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "POOL FILLED","400,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Sensory Deprevation Pool | 3/10",1,""
      Case 4
        AddScore 400000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "POOL FILLED","400,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Sensory Deprevation Pool | 4/10",1,""
      Case 5
        AddScore 400000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "POOL FILLED","400,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Sensory Deprevation Pool | 5/10",1,""
      Case 6
        AddScore 400000
        pNote "POOL FILLED","400,000"
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Sensory Deprevation Pool | 6/10",1,""
      Case 7
        AddScore 400000
        pNote "POOL FILLED","400,000"
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Sensory Deprevation Pool | 7/10",1,""
      Case 8
        AddScore 400000
        pNote "POOL FILLED","400,000"
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Sensory Deprevation Pool | 8/10",1,""
      Case 9
        AddScore 400000
        pNote "POOL FILLED","400,000"
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Sensory Deprevation Pool | 9/10",1,""
      Case 10
        AddScore 1500000
        pNote "WILL LOCATED","1,500,000"
        PuPlayer.playlistplayex pCallouts,"audiocallouts","poolcomplete.wav",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Sensory Deprevation Pool | 10/10",1,""
        poolend
        DOF 422, DOFPulse  ' DOF MX - AV Club Mode Complete
    End Select
  End Sub

  Sub poolend

    inmode = 0
    PuPlayer.LabelSet pBackglass,"modetimer"," ",1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 80.7, 'xalign': 1, 'ypos': 75, 'yalign': 0}"
    checkhits
    ruleshelperon
    PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,2
    PuPlayer.SetLoop 2,0
    If bMultiBallMode = False Then
    GiOff
    GiOn
    End If
    PuPlayer.playresume 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1
    PuPlayer.SetLoop 7,0
    pooltime.enabled = False
    SetLightColor pool, amber, -1
    SetLightColor pool1, amber, -1
    pooldone(CurrentPlayer) = 1
    pool.state = 1
    pool1.state = 1
    poolpos = 0
    poolawarded = 0
    poolactive = 0
    resetmodecheck
    If avsdone(CurrentPlayer) = 2 and extraball.state = 0 Then
      extraballmode
    End If
    PuPlayer.LabelSet pBackglass,"mav","" & avsdone(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 94.6, 'yalign': 0}"
  End Sub

  '****************
  '     RADIO
  '****************
  ' roving dot - 5 shots to complete
  ' 60 seconds to complete

  Dim radioactive
  radioactive = 0

  Sub radiostart
    If radioskip = 0 then exit Sub
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Radio Mode | 0/5",1,""
    exitav
    BallLockEscape.enabled = False
    inmode = 1
    pNote "A.V CLUB","RADIO MODE"
    PuPlayer.playlistplayex pBackglass,"videoradio","radiomid.mov",100,2
    PuPlayer.SetLoop 2,1
    GiOff
    GiPurple
    PuPlayer.playpause 4
    PuPlayer.playlistplayex pAudio,"audiomodes","radio.mp3",100,1
    PuPlayer.SetLoop 7,1
    radiomove.enabled = True
    radiotime.enabled = True
    radioactive = 1
    batterygreen.state = 0
    batterygreen1.state = 0
    batteryorange.state = 0
    batteryorange1.state = 0
    batteryyellow.state = 0
    batteryyellow1.state = 0
    batteryred.state = 0
    batteryred1.state = 0
    batteryon(CurrentPlayer) = 0
    batt1on(CurrentPlayer) = 0
    batt2on(CurrentPlayer) = 0
    batt3on(CurrentPlayer) = 0
    avsdone(CurrentPlayer) = avsdone(CurrentPlayer) + 1
    radioskip = 0
  End Sub

  Dim radiolight
  radiolight = 0

  Sub radiomove_Timer()
    radiolight = radiolight + 1
    Select Case radiolight
      Case 1
        llo4.state = 2
        llo6.state = 2
        lc2.state = 0
        lc3.state = 0
        llo5.state = 0
        llo10.state = 0
      Case 6
        llo4.state = 0
        llo6.state = 0
        lc2.state = 2
        lc3.state = 2
        llo5.state = 0
        llo10.state = 0
      Case 12
        llo4.state = 0
        llo6.state = 0
        lc2.state = 0
        lc3.state = 0
        llo5.state = 2
        llo10.state = 2
      Case 17
        radiolight = 0
    End Select
  End Sub

  Dim radiopos
  radiopos = 0

  Sub radiotime_Timer()
    PuPlayer.LabelSet pBackglass,"modetimer",60 - radiopos,1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 80.7, 'xalign': 1, 'ypos': 75, 'yalign': 0}"
    radiopos = radiopos + 1
    Select Case radiopos
      Case 1

      Case 60
        radioend
    End Select
  End Sub

  Dim radioawarded
  radioawarded = 0

  Sub radioaward
    radioawarded = radioawarded + 1
    Select Case radioawarded
      Case 1
        AddScore 1000000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "SIGNAL FOUND","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Radio Mode | 1/5",1,""
      Case 2
        AddScore 1000000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "SIGNAL FOUND","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Radio Mode | 2/5",1,""
      Case 3
        AddScore 1000000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "SIGNAL FOUND","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Radio Mode | 3/5",1,""
      Case 4
        AddScore 1000000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "SIGNAL FOUND","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Radio Mode | 4/5",1,""
      Case 5
        AddScore 2000000
        DOF 422, DOFPulse  ' DOF MX - AV Club Mode Complete
        pNote "WILL CONTACTED","2,000,000"
        PuPlayer.playlistplayex pCallouts,"audiocallouts","radiocomplete.wav",100,1
    chilloutthemusic
        PuPlayer.playlistplayex pBackglass,"videoradio","radioend.mov",100,2
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Radio Mode | 5/5",1,""
        radioend
      Case 6
        radioend
    End Select
  End Sub

  Sub radioend

    inmode = 0
    PuPlayer.LabelSet pBackglass,"modetimer"," ",1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 80.7, 'xalign': 1, 'ypos': 75, 'yalign': 0}"
    checkhits
    ruleshelperon
    PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,2
    PuPlayer.SetLoop 2,0
    If bMultiBallMode = False Then
    GiOff
    GiOn
    End If
    PuPlayer.playresume 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1
    PuPlayer.SetLoop 7,0
    radiodone(CurrentPlayer) = 1
    SetLightColor radio, amber, -1
    radio.state = 1
    SetLightColor radio1, amber, -1
    radio1.state = 1
    resetmodecheck
    radiolight = 0
    radiopos = 0
    radioawarded = 0
    radioactive = 0
    radiomove.enabled = False
    radiotime.enabled = False
    If avsdone(CurrentPlayer) = 2 and extraball.state = 0 Then
      extraballmode
    End If
    PuPlayer.LabelSet pBackglass,"mav","" & avsdone(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 94.6, 'yalign': 0}"
  End Sub

  '****************
  '     COMPASS
  '****************
  ' dot clear 2x clear the dots, then do it again. 6 total shots to complete.
  ' 60 seconds to complete

  Dim compassactive
  compassactive = 0

  Sub compassstart
    If compassskip = 0 then exit Sub
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Compass Mode | 0/6",1,""
    PuPlayer.LabelSet pBackglass,"notecopy","Clear the purple dots 2x within 60 seconds to find the portal \r Each shot scores 1,000,000 points",1,""

    exitav
    BallLockEscape.enabled = False
    inmode = 1
    pNote "A.V CLUB","COMPASS MODE"
    PuPlayer.playlistplayex pBackglass,"videocompass","compassmid.mov",100,2
    GiOff
    GiPurple
    PuPlayer.playpause 4
    PuPlayer.playlistplayex pAudio,"audiomodes","compass.mp3",100,1
    PuPlayer.SetLoop 7,1
    compassactive = 1
    llo4.state = 2
    llo6.state = 2
    lc2.state = 2
    lc3.state = 2
    llo5.state = 2
    llo10.state = 2
    compasstime.enabled = True
    batterygreen.state = 0
    batterygreen1.state = 0
    batteryorange.state = 0
    batteryorange1.state = 0
    batteryyellow.state = 0
    batteryyellow1.state = 0
    batteryred.state = 0
    batteryred1.state = 0
    batteryon(CurrentPlayer) = 0
    batt1on(CurrentPlayer) = 0
    batt2on(CurrentPlayer) = 0
    batt3on(CurrentPlayer) = 0
    avsdone(CurrentPlayer) = avsdone(CurrentPlayer) + 1
    compassskip = 0
  End Sub

  Dim compasspos
  compasspos = 0

  Sub compasstime_Timer()
    compasspos = compasspos + 1
    PuPlayer.LabelSet pBackglass,"modetimer",60 - compasspos,1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 80.7, 'xalign': 1, 'ypos': 75, 'yalign': 0}"
    Select Case compasspos
      Case 1

      Case 60
        compassend
        pNote "COMPASS MODE","OVER"
    End Select
  End Sub

  Dim compassawarded
  compassawarded = 0

  Sub compassaward
    compassawarded = compassawarded + 1
    Select Case compassawarded
      Case 1
        AddScore 1000000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "COMPASS FOLLOWED","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Compass Mode | 1/6",1,""
      Case 2
        AddScore 1000000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "COMPASS FOLLOWED","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Compass Mode | 2/6",1,""
      Case 3
        AddScore 1000000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "COMPASS FOLLOWED","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Compass Mode | 3/6",1,""
      Case 4
        AddScore 1000000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "COMPASS FOLLOWED","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Compass Mode | 4/6",1,""
      Case 5
        AddScore 1000000
        DOF 421, DOFPulse  ' DOF MX - AV Club Award
        pNote "COMPASS FOLLOWED","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiocheers","",100,1
    chilloutthemusic
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Compass Mode | 5/6",1,""
      Case 6
        AddScore 3000000
        DOF 422, DOFPulse  ' DOF MX - AV Club Mode Complete
        pNote "PORTAL FOUND","3,000,000"
        PuPlayer.playlistplayex pBackglass,"videocompass","compassend.mov",100,1
    chilloutthemusic
        PuPlayer.playlistplayex pCallouts,"audiocallouts","compasscomplete.wav",100,1
    PuPlayer.LabelSet pBackglass,"notetitle","A.V. CLUB - Compass Mode | 6/6",1,""
        compassend
        compassactive = 0
    End Select
  End Sub

  Sub compasscheck
      If llo4.state + lc2.state + llo5.state = 0 Then
        llo4.state = 2
        llo6.state = 2
        lc2.state = 2
        lc3.state = 2
        llo5.state = 2
        llo10.state = 2
      End If
  End Sub

  Sub compassend
    inmode = 0
    checkhits
    ruleshelperon
    PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,2
    PuPlayer.SetLoop 2,0
    PuPlayer.LabelSet pBackglass,"modetimer"," ",1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 80.7, 'xalign': 1, 'ypos': 75, 'yalign': 0}"
    If bMultiBallMode = False Then
    GiOff
    GiOn
    End If
    PuPlayer.playresume 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1
    PuPlayer.SetLoop 7,0
    compasstime.enabled = False
    compassdone(CurrentPlayer) = 1
    SetLightColor compass, amber, -1
    compass.state = 1
    SetLightColor compass1, amber, -1
    compass1.state = 1
    compassactive = 0
    compassawarded = 0
    compasspos = 0
    llo4.state = 0
    llo6.state = 0
    lc2.state = 0
    lc3.state = 0
    llo5.state = 0
    llo10.state = 0
    resetmodecheck
    If avsdone(CurrentPlayer) = 2 and extraball.state = 0 Then
      extraballmode
    End If
    PuPlayer.LabelSet pBackglass,"mav","" & avsdone(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 94.6, 'yalign': 0}"
  End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  WAFFLES & MYSTERY
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'


  Sub Bumper1_Hit
    WaffleShake()
    LightEffect 5
    If NOT Tilted Then
      BumperRewards
      PlaySoundAt SoundFXDOF("fx_bumper", 107, DOFPulse, DOFContactors), ActiveBall
      DOF 110, DOFPulse   'Left Flasher Gold
      DOF 302, DOFPulse   'DOF MX - Bumper 1
      PlaySound "wave"
      AddScore 1000
    End If
  End Sub

  Sub Bumper2_Hit
    WaffleShake2()
    LightEffect 5
    If NOT Tilted Then
      BumperRewards
      PlaySoundAt SoundFXDOF("fx_bumper", 109, DOFPulse, DOFContactors), ActiveBall
      DOF 111, DOFPulse   'Center Flasher Gold
      DOF 303, DOFPulse   'DOF MX - Bumper 2
      PlaySound "wave"
      AddScore 1000
    End If
  End Sub

  Sub Bumper3_Hit
    WaffleShake3()
    LightEffect 5
    If NOT Tilted Then
      BumperRewards
      PlaySoundAt SoundFXDOF("fx_bumper", 108, DOFPulse, DOFContactors), ActiveBall
      DOF 112, DOFPulse   'Right Flasher Gold
      DOF 304, DOFPulse   'DOF MX - Bumper 3
      PlaySound "wave"
      AddScore 1000
    End If
  End Sub

  Sub BumperRewards
      PuPlayer.LabelSet pBackglass,"mwaf", bumps(CurrentPlayer),1,""
      bumps(CurrentPlayer) = bumps(CurrentPlayer) + 1
      Select Case bumps(CurrentPlayer)
        Case 10
          pNote "MMMM EGGOS", bumps(CurrentPlayer) & " EGGOS COLLECTED"
          PuPlayer.playlistplayex pBackglass,"videoeggos","",100,1
        Case 25
          mysteryready(CurrentPlayer) = 1
          PuPlayer.playlistplayex pCallouts,"audiocallouts","getthemystery.wav",100,1
    chilloutthemusic
          mysterylight.state = 2
          mysterylight1.state = 2
          currentmystery(CurrentPlayer) = 1
        Case 40
          pNote "MMMM EGGOS", bumps(CurrentPlayer) & " EGGOS COLLECTED"
          PuPlayer.playlistplayex pBackglass,"videoeggos","",100,1
        Case 60
          pNote "MMMM EGGOS", bumps(CurrentPlayer) & " EGGOS COLLECTED"
          PuPlayer.playlistplayex pBackglass,"videoeggos","",100,1
        Case 70
          mysteryready(CurrentPlayer) = 1
          mysterylight.state = 2
          mysterylight1.state = 2
          PuPlayer.playlistplayex pCallouts,"audiocallouts","getthemystery.wav",100,1
    chilloutthemusic
          currentmystery(CurrentPlayer) = 2
        Case 90
          pNote "MMMM EGGOS", bumps(CurrentPlayer) & " EGGOS COLLECTED"
          PuPlayer.playlistplayex pBackglass,"videoeggos","",100,1
        Case 120
          mysteryready(CurrentPlayer) = 1
          mysterylight.state = 2
          mysterylight1.state = 2
          PuPlayer.playlistplayex pCallouts,"audiocallouts","getthemystery.wav",100,1
    chilloutthemusic
          currentmystery(CurrentPlayer) = 3
        Case 130
          pNote "MMMM EGGOS", bumps(CurrentPlayer) & " EGGOS COLLECTED"
          PuPlayer.playlistplayex pBackglass,"videoeggos","",100,1
        Case 180
          pNote "MMMM EGGOS", bumps(CurrentPlayer) & " EGGOS COLLECTED"
          PuPlayer.playlistplayex pBackglass,"videoeggos","",100,1
        Case 200
          mysteryready(CurrentPlayer) = 1
          mysterylight.state = 2
          mysterylight1.state = 2
          PuPlayer.playlistplayex pCallouts,"audiocallouts","getthemystery.wav",100,1
    chilloutthemusic
          currentmystery(CurrentPlayer) = 4
        Case 230
          pNote "MMMM EGGOS", bumps(CurrentPlayer) & " EGGOS COLLECTED"
          PuPlayer.playlistplayex pBackglass,"videoeggos","",100,1
        Case 280
          pNote "MMMM EGGOS", bumps(CurrentPlayer) & " EGGOS COLLECTED"
          PuPlayer.playlistplayex pBackglass,"videoeggos","",100,1
        Case 300
          mysteryready(CurrentPlayer) = 1
          mysterylight.state = 2
          mysterylight1.state = 2
          PuPlayer.playlistplayex pCallouts,"audiocallouts","getthemystery.wav",100,1
    chilloutthemusic
          currentmystery(CurrentPlayer) = 5
        Case 330
          pNote "MMMM EGGOS", bumps(CurrentPlayer) & " EGGOS COLLECTED"
          PuPlayer.playlistplayex pBackglass,"videoeggos","",100,1
        Case 400
          pNote "MMMM EGGOS", bumps(CurrentPlayer) & " EGGOS COLLECTED"
          PuPlayer.playlistplayex pBackglass,"videoeggos","",100,1
        Case 450
          pNote "MMMM EGGOS", bumps(CurrentPlayer) & " EGGOS COLLECTED"
          PuPlayer.playlistplayex pBackglass,"videoeggos","",100,1
        Case 500
          mysteryready(CurrentPlayer) = 1
          mysterylight.state = 2
          mysterylight1.state = 2
          PuPlayer.playlistplayex pCallouts,"audiocallouts","getthemystery.wav",100,1
    chilloutthemusic
          currentmystery(CurrentPlayer) = 6
      End Select
  End Sub

  Dim mysteryreward

  Sub mysteryvalue
    mysteryreward = RndNum(1,20)
    mysteryawarded
  End Sub


  Sub Mysteryawarded
    mysteryready(CurrentPlayer) = 0
    mysterylight.state = 0
    mysterylight1.state = 0
    dim waittime
    waittime = 6000
    PuPlayer.playlistplayex pBackglass,"videomystery","mystery.mov",100,1
    Select Case mysteryreward
      Case 1
        pNote "EXTRA BALL","IS LIT"
        vpmtimer.addtimer waittime, "extraballmode'"
      Case 2 'Big Points
        AddScore 1000000
        pNote "BIG POINTS","1,000,000"
      Case 3 'little points
        AddScore 200000
        pNote "LITTLE POINTS","200,000"
      Case 4 'barb lock lit or if lit then lock given
        If lro1.state = 0 Then
          lb1.State = 1
          lb6.State = 1
          barb1on(CurrentPlayer) = 1
          barb2on(CurrentPlayer) = 1
          barb3on(CurrentPlayer) = 1
          barb4on(CurrentPlayer) = 1
          lb2.State = 1
          lb5.State = 1
          lb3.State = 1
          lb8.State = 1
          lb4.state = 1
          lb7.state = 1
          pNote "WHERE'S BARB","LOCK LIT AWARDED"
          vpmtimer.addtimer waittime, "CheckBARBTargets'"

        Else
    BallsInLock(CurrentPlayer) = BallsInLock(CurrentPlayer) + 1
      Select Case BallsInLock(CurrentPlayer)
        Case 1
          ResetBARBLights
          DOF 405, DOFPulse   'DOF MX - Ball Locked
          waittime = 1000
          PuPlayer.playlistplayex pCallouts,"audiocallouts","ball1locked.wav",100,1
    chilloutthemusic
          PuPlayer.playlistplayex pBackglass,"videobarblock","",100,1
          pNote "WHERE'S BARB","BALL 1 LOCKED"
          vpmtimer.addtimer waittime, "BallLockBarbExit '"
          PuPlayer.LabelSet pBackglass,"barbh","" & BallsInLock(CurrentPlayer) ,1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
        Case 2
          barbskip = 1
          ResetBARBLights
          ruleshelperoff
          GiOff
          PuPlayer.LabelSet pBackglass,"notetitle","Where's Barb Multiball",1,""
          PuPlayer.LabelSet pBackglass,"notecopy","Hit the Yellow Shots - 1 Randomly will light Super Jackpot \r Hit Castle Byers when Super Jackpot List to Find Barb",1,""
          If inmode = 1 Then
          waittime = 1000
          Else
          waittime = 16000
          End If
          PuPlayer.playpause 4
          PuPlayer.playlistplayex pBackglass,"videobarb","barbmdstart.mov",100,1
          vpmtimer.addtimer waittime, "StartBarb'"
          PuPlayer.LabelSet pBackglass,"barbh","" & BallsInLock(CurrentPlayer) ,1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
      End Select
        End If
      Case 5 'badmen lock
        BallsInRunLock(CurrentPlayer) = BallsInRunLock(CurrentPlayer) + 1
      Select Case BallsInRunLock(CurrentPlayer)
        Case 1
          run1lock(CurrentPlayer) = 1
          PuPlayer.playlistplayex pCallouts,"audiocallouts","ball1locked.wav",100,1
    chilloutthemusic
          pNote "BAD MEN","BALL 1 LOCKED"
          PuPlayer.playlistplayex pBackglass,"videobadmenlock","",100,1
          DOF 405, DOFPulse   'DOF MX - Ball Locked
          PuPlayer.LabelSet pBackglass,"badh","" & BallsInRunLock(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
          waittime = 1000
          vpmtimer.addtimer waittime, "BallLockRunExit'"
        Case 3
          ResetRunLights
          run1lock(CurrentPlayer) = 2
          PuPlayer.playlistplayex pCallouts,"audiocallouts","ball2locked.wav",100,1
    chilloutthemusic
          pNote "BAD MEN","BALL 2 LOCKED"
          PuPlayer.playlistplayex pBackglass,"videobadmenlock","",100,1
          waittime = 1000
          DOF 405, DOFPulse   'DOF MX - Ball Locked
          PuPlayer.LabelSet pBackglass,"badh","" & BallsInRunLock(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
          vpmtimer.addtimer waittime, "BallLockRunExit '"
        Case 2
          badmenskip = 1
          PuPlayer.playpause 4
          ResetRunLights
          run1lock(CurrentPlayer) = 3
          ruleshelperoff
          GiOff
          PuPlayer.LabelSet pBackglass,"notetitle","Bad Men Multiball",1,""
          PuPlayer.LabelSet pBackglass,"notecopy","Hit the outside ramps 5 times to light Super Jackpot \r Hit the inside ramp to get the Super Jackpot and escape!",1,""
          PuPlayer.playlistplayex pBackglass,"videobadmenmb","badmenstart.mov",100,1
          If inmode = 1 Then
            waittime = 1000
          Else
            waittime = 11000
          End If
          PuPlayer.LabelSet pBackglass,"badh","" & BallsInRunLock(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
          vpmtimer.addtimer waittime, "StartRun'"
      End Select

      Case 6 'Mode ready NEED TO SET RULES
        batterygreen.state = 2
        batterygreen1.state = 2
        batteryorange.state = 2
        batteryorange1.state = 2
        batteryyellow.state = 2
        batteryyellow1.state = 2
        batteryred.state = 2
        batteryred1.state = 2
        avclubready.state = 2
        avclubready1.state = 2
        avclubreadyon(CurrentPlayer) = 1
        batteryon(CurrentPlayer) = 1
        lbolt1.state = 0
        lbolt7.state = 0
        lbolt2.state = 0
        lbolt8.state = 0
        lbolt3.state = 0
        lbolt10.state = 0
        lbolt4.state = 0
        lbolt9.state = 0
        lbolt5.state = 0
        lbolt11.state = 0
        lbolt6.state = 0
        lbolt12.state = 0
        pNote "MODE","IS LIT"
      Case 7 'super ball saver
        EnableBallSaver 30
        pNote "SUPER BALL SAVER","30 SECONDS"
      Case 8  'zilch
        Addscore 0
        pNote "ZILCH","NO AWARD"
      Case 9 'award party member NEED TO SET RULES
        lucas.state = 1
        jonathan.state = 1
        nancy.state = 1
        joyce.state = 1
        will.state = 1
        elevencrew.state = 1
        mikey.state = 1
        hopper.state = 1
        dustin.state = 1
        pNote "AWARD PARTY MEMBERS","ALL MEMBERS FILLED"
        vpmtimer.addtimer waittime, "checkparty'"

      Case 10 'badmen lock
        BallsInRunLock(CurrentPlayer) = BallsInRunLock(CurrentPlayer) + 1
      Select Case BallsInRunLock(CurrentPlayer)
        Case 1
          run1lock(CurrentPlayer) = 1
          PuPlayer.playlistplayex pCallouts,"audiocallouts","ball1locked.wav",100,1
    chilloutthemusic
          pNote "BAD MEN","BALL 1 LOCKED"
          PuPlayer.playlistplayex pBackglass,"videobadmenlock","",100,1
          DOF 405, DOFPulse   'DOF MX - Ball Locked
          PuPlayer.LabelSet pBackglass,"badh","" & BallsInRunLock(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
          waittime = 1000
          vpmtimer.addtimer waittime, "BallLockRunExit'"
        Case 3
          ResetRunLights
          run1lock(CurrentPlayer) = 2
          PuPlayer.playlistplayex pCallouts,"audiocallouts","ball2locked.wav",100,1
    chilloutthemusic
          pNote "BAD MEN","BALL 2 LOCKED"
          PuPlayer.playlistplayex pBackglass,"videobadmenlock","",100,1
          waittime = 1000
          DOF 405, DOFPulse   'DOF MX - Ball Locked
          PuPlayer.LabelSet pBackglass,"badh","" & BallsInRunLock(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
          vpmtimer.addtimer waittime, "BallLockRunExit '"
        Case 2
          badmenskip = 1
          PuPlayer.playpause 4
          ResetRunLights
          run1lock(CurrentPlayer) = 3
          ruleshelperoff
          GiOff
          PuPlayer.LabelSet pBackglass,"notetitle","Bad Men Multiball",1,""
          PuPlayer.LabelSet pBackglass,"notecopy","Hit the outside ramps 5 times to light Super Jackpot \r Hit the inside ramp to get the Super Jackpot and escape!",1,""
          PuPlayer.playlistplayex pBackglass,"videobadmenmb","badmenstart.mov",100,1
          If inmode = 1 Then
            waittime = 1000
          Else
            waittime = 11000
          End If
          PuPlayer.LabelSet pBackglass,"badh","" & BallsInRunLock(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
          vpmtimer.addtimer waittime, "StartRun'"
      End Select
      Case 11 ' drop wall
        If barricadedowncurrently = False Then
          wallll.State = 1
          wallll1.State = 1
          wallml.State = 1
          wallml1.State = 1
          wallrl.State = 1
          wallrl1.State = 1
          walllon(CurrentPlayer) = 1
          wallmon(CurrentPlayer) = 1
          wallron(CurrentPlayer) = 1
          pNote "DROP BARRICADE","TIME TO BASH DEMOGORGON"
          vpmtimer.addtimer waittime, "checkbank'"
          checkbank
        Else
          udhits(CurrentPlayer) = udhits(CurrentPlayer) + 8
          pNote "AWARD DEMOGORGON HITS","8 HITS ADDED"
          vpmtimer.addtimer waittime, "checkhits'"

        End If
      Case 12
        If barricadedowncurrently = False Then
          wallll.State = 1
          wallll1.State = 1
          wallml.State = 1
          wallml1.State = 1
          wallrl.State = 1
          wallrl1.State = 1
          walllon(CurrentPlayer) = 1
          wallmon(CurrentPlayer) = 1
          wallron(CurrentPlayer) = 1
          pNote "DROP BARRICADE","TIME TO BASH DEMOGORGON"
          vpmtimer.addtimer waittime, "checkbank'"
          checkbank
        Else
          udhits(CurrentPlayer) = udhits(CurrentPlayer) + 8
          pNote "AWARD DEMOGORGON HITS","8 HITS ADDED"
          vpmtimer.addtimer waittime, "checkhits'"

        End If
      Case 13 ' tiny points
        AddScore 20000
        pNote "TINY POINTS","20,000"
      Case 14 'Huge Points
        AddScore 5000000
        pNote "HUGE POINTS","5,000,000"
      Case 15 'Ultra ball saver
        EnableBallSaver 60
        pNote "ULTRA BALL SAVER","60 SECONDS"
      Case 16
        If lro1.state = 0 Then
          lb1.State = 1
          lb6.State = 1
          barb1on(CurrentPlayer) = 1
          barb2on(CurrentPlayer) = 1
          barb3on(CurrentPlayer) = 1
          barb4on(CurrentPlayer) = 1
          lb2.State = 1
          lb5.State = 1
          lb3.State = 1
          lb8.State = 1
          lb4.state = 1
          lb7.state = 1
          pNote "WHERE'S BARB","LOCK LIT AWARDED"
          vpmtimer.addtimer waittime, "CheckBARBTargets'"
        Else
    BallsInLock(CurrentPlayer) = BallsInLock(CurrentPlayer) + 1
      Select Case BallsInLock(CurrentPlayer)
        Case 1
          ResetBARBLights
          DOF 405, DOFPulse   'DOF MX - Ball Locked
          waittime = 1000
          PuPlayer.playlistplayex pCallouts,"audiocallouts","ball1locked.wav",100,1
    chilloutthemusic
          PuPlayer.playlistplayex pBackglass,"videobarblock","",100,1
          pNote "WHERE'S BARB","BALL 1 LOCKED"
          vpmtimer.addtimer waittime, "BallLockBarbExit '"
          PuPlayer.LabelSet pBackglass,"barbh","" & BallsInLock(CurrentPlayer) ,1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
        Case 2
          barbskip = 1
          ResetBARBLights
          ruleshelperoff
          GiOff
          PuPlayer.LabelSet pBackglass,"notetitle","Where's Barb Multiball",1,""
          PuPlayer.LabelSet pBackglass,"notecopy","Hit the Yellow Shots - 1 Randomly will light Super Jackpot \r Hit Castle Byers when Super Jackpot List to Find Barb",1,""
          If inmode = 1 Then
          waittime = 1000
          Else
          waittime = 16000
          End If
          PuPlayer.playpause 4
          PuPlayer.playlistplayex pBackglass,"videobarb","barbmdstart.mov",100,1
          vpmtimer.addtimer waittime, "StartBarb'"
          PuPlayer.LabelSet pBackglass,"barbh","" & BallsInLock(CurrentPlayer) ,1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
      End Select
        End If
      Case 17
        pNote "EXTRA BALL","IS LIT"
        vpmtimer.addtimer waittime, "extraballmode'"
      Case 18  'zilch
        Addscore 0
        pNote "ZILCH","NO AWARD"
      Case 19
        lucas.state = 1
        jonathan.state = 1
        nancy.state = 1
        joyce.state = 1
        will.state = 1
        elevencrew.state = 1
        mikey.state = 1
        hopper.state = 1
        dustin.state = 1
        pNote "AWARD PARTY MEMBERS","ALL MEMBERS FILLED"
        vpmtimer.addtimer waittime, "checkparty'"

      Case 20 'Huge Points
        AddScore 5000000
        pNote "HUGE POINTS","5,000,000"

    End Select
  End Sub

  Sub castlekicker_hit
    Dim waittime
    waittime = 1000
    RandomSoundHole

    If partymultiball = False Then
      If will.state = 0 Then
        will.state = 1
        willon(CurrentPlayer) = 1
        partycollected(CurrentPlayer) = partycollected(CurrentPlayer) + 1
        PuPlayer.LabelSet pBackglass,"partyh","" & partycollected(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
        checkparty
      End If
    End If

    If will.state = 2 Then
      will.state = 0
      awardparty
    End If



    If bMultiBallMode = False Then
      If mysterylight.state = 0 and partylock.state = 0 Then
        waittime = 1000
        PuPlayer.playlistplayex pBackglass,"videobyerscastle","",100,1
        vpmtimer.addtimer waittime, "exitcastle'"
      End If

      If partylock.state = 2 and mysterylight.state = 0 Then
        partyprestart
      End If

      If mysterylight.state = 2 Then'
        If partylock.state = 2 Then
          waittime = 3000
          mysteryvalue
          vpmtimer.addtimer waittime, "partyprestart'"
        Else
          waittime = 3000
          mysteryvalue
          vpmtimer.addtimer waittime, "exitcastle'"
        End If
      End If

    Else
      If barbMultiball = True Then
        If lro2.state = 2 Then
          awardbarbsuper
          lro2.state = 0
          lro6.state = 0
        End If
      End If
      vpmtimer.addtimer waittime, "exitcastle'"
    End If
  End Sub

  Sub partyprestart
    partyskip = 1
    PuPlayer.playpause 4
    gioff
    Dim waittime
    If inmode = 1 Then
    waittime = 1000
    Else
    waittime = 15000
    End If
    PuPlayer.playlistplayex pBackglass,"videoparty","party.mov",100,3
    partyready(CurrentPlayer) = 0
    vpmtimer.addtimer waittime, "startparty'"
    ruleshelperoff
    PuPlayer.LabelSet pBackglass,"notetitle","Adventuring Party Multiball",1,""
    PuPlayer.LabelSet pBackglass,"notecopy","Hit all the shots on the table to collect party members \r Balls added every 3 members collected Collect All To Complete",1,""

  End sub

  Sub exitcastle
    DOF 360, DOFPulse  'DOF MX - Right Red Flasher
    FlashLevel3 = 1 : Flasherflash3_Timer
    Dim waittime
    waittime = 400
    vpmtimer.addtimer waittime, "exitcastle1'"
    PlaySound "flute"
  End Sub

  Sub exitcastle1
    DOF 360, DOFPulse  'DOF MX - Right Red Flasher
    FlashLevel3 = 1 : Flasherflash3_Timer
    Dim waittime
    waittime = 400
    vpmtimer.addtimer waittime, "exitcastle2'"
    PlaySound "flute"
  End Sub

  Sub exitcastle2
    DOF 115, DOFPulse  'DOF - Strobe
    DOF 360, DOFPulse  'DOF MX - Right Red Flasher
    FlashLevel3 = 1 : Flasherflash3_Timer
    castlekicker.Kick 230, 22
    PlaySoundAt SoundFXDOF("fx_kicker", 137, DOFPulse, DOFContactors), castlekicker
    PlaySound "flute"
  End Sub


  '***************************
  'Waffle Animation / shake
  '***************************

  Dim WafflePos

  Sub WaffleShake()
    WafflePos = 3
    waffletime.Enabled = True
  End Sub

  Sub waffletime_Timer()
    waffle1.RotY = WafflePos
    If WafflePos <= 0.1 AND WafflePos >= -0.1 Then Me.Enabled = False:Exit Sub
    If WafflePos < 0 Then
      WafflePos = ABS(WafflePos)- 0.1
    Else
      WafflePos = - WafflePos + 0.1
    End If
  End Sub

  Dim WafflePos2

  Sub WaffleShake2()
    WafflePos2 = 3
    waffletime2.Enabled = True
  End Sub

  Sub waffletime2_Timer()
    waffle2.RotY = WafflePos2
    If WafflePos2 <= 0.1 AND WafflePos2 >= -0.1 Then Me.Enabled = False:Exit Sub
    If WafflePos2 < 0 Then
      WafflePos2 = ABS(WafflePos2)- 0.1
    Else
      WafflePos2 = - WafflePos2 + 0.1
    End If
  End Sub

  Dim WafflePos3

  Sub WaffleShake3()
    WafflePos3 = 3
    waffletime3.Enabled = True
  End Sub

  Sub waffletime3_Timer()
    waffle3.RotY = WafflePos3
    If WafflePos3 <= 0.1 AND WafflePos3 >= -0.1 Then Me.Enabled = False:Exit Sub
    If WafflePos3 < 0 Then
      WafflePos3 = ABS(WafflePos3)- 0.1
    Else
      WafflePos3 = - WafflePos3 + 0.1
    End If
  End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  ADVENTURING PARTY MULTIBALL
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'

  Sub checkparty
    If lucas.state + jonathan.state + nancy.state + joyce.state + will.state + elevencrew.state + mikey.state + hopper.state + dustin.state = 9 Then
      If bMultiBallMode = True Then
      Else
        If partylock.state = 0 Then
          PuPlayer.playlistplayex pCallouts,"audiocallouts","partymultiballlit.wav",100,1
    chilloutthemusic
        End If
        partylock.state = 2
        partylock1.state = 2
        partyready(CurrentPlayer) = 1
      End If
    End If
  End Sub

  Sub startparty
    If partyskip = 0 then exit Sub
    exitcastle
    DOF 420, DOFPulse  'DOF MX - PARTY FLASH
    pNote "ADVENTURING PARTY","MULTIBALL"
    PuPlayer.playlistplayex pBackglass,"videoparty","partymb.mov",100,3
    PuPlayer.SetLoop 2,1
    GiOff
    GiOrange
    PuPlayer.playpause 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","party.mp3",100,1
    PuPlayer.SetLoop 7,1
    partylock.state = 0
    partylock1.state = 0
    partyready(CurrentPlayer) = 0
    PuPlayer.LabelSet pBackglass,"barbh","0",1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
    bMultiBallMode = True
    partyMultiball = True
    AddMultiball 1
    EnableBallSaver 15
    flashflash.Enabled = True
'   StartDemoLightSeq
'   GeneralPupQueue.Add "StopDemoLightSeq","StopDemoLightSeq",80,100000,0,0,0,False

  dim k, newLoopDelay

    PrepSpellWord 40
    SetLoopDelay "adventuring party"

    SpellWord "adventuring party"

    for k = 1 to 7

      NewLoopDelay = LoopDelay*k
    debug.print "loop delay :" &k &":" &newloopdelay

      xmasQueue.Add "SpellWord-"&k,"SpellWord ""adventuring party"" ",80,NewLoopDelay,0,0,0,False
    Next

    xmasQueue.Add "SpellEnd","giOn:bInLightQuote = False",75,LoopDelay*8+(delayInc*myFSpeed*2),0,0,0,False

'   LightSeqaxmas.UpdateInterval = 150
'   LightSeqaxmas.Play SeqRandom, 10, , 100000
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 100000
    boltsoff
    lucas.state = 2
    jonathan.state = 2
    nancy.state = 2
    joyce.state = 2
    will.state = 2
    elevencrew.state = 2
    mikey.state = 2
    hopper.state = 2
    dustin.state = 2
    startamultiball
    partyskip = 0
  End Sub


  Sub awardparty
    If DemoMultiball = True then exit sub
    partyjacks(CurrentPlayer) = partyjacks(CurrentPlayer) + 1
    Select Case partyjacks(CurrentPlayer)
      Case 1
        AddScore 1000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
        PuPlayer.LabelSet pBackglass,"partyj","" & partyjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
      Case 2
        AddScore 1000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
        PuPlayer.LabelSet pBackglass,"partyj","" & partyjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
      Case 3
        AddScore 1000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","1,000,000"
        AddMultiball 1
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
        PuPlayer.LabelSet pBackglass,"partyj","" & partyjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"

      Case 4
        AddScore 1000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
        PuPlayer.LabelSet pBackglass,"partyj","" & partyjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"

      Case 5
        AddScore 1000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
        PuPlayer.LabelSet pBackglass,"partyj","" & partyjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"

      Case 6
        AddScore 1000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
        PuPlayer.LabelSet pBackglass,"partyj","" & partyjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"

        AddMultiball 1
      Case 7
        AddScore 1000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","1,000,000"
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
        PuPlayer.LabelSet pBackglass,"partyj","" & partyjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"

      Case 8
        AddScore 1000000
        DOF 429, DOFPulse  ' DOF MX - Party Get SJ
        pNote "JACKPOT","1,000,000"
        PuPlayer.LabelSet pBackglass,"partyj","" & partyjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
        PuPlayer.playlistplayex pCallouts,"audiocallouts","getthesuperjackpot.wav",100,1
    chilloutthemusic
      Case 9
        partysuper
        DOF 423, DOFPulse  ' DOF MX - Party Mode Complete
        PuPlayer.LabelSet pBackglass,"partyj","" & partyjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"

    End Select
  End Sub

  Sub partysuper
    AddScore 10000000
    pNote "ADVENTURING PARTY COLLECTED","10,000,000"
    lm5.state = 1
    lm10.state = 1
    partyjacks(CurrentPlayer) = 0
    lucas.state = 2
    jonathan.state = 2
    nancy.state = 2
    joyce.state = 2
    will.state = 2
    elevencrew.state = 2
    mikey.state = 2
    hopper.state = 2
    dustin.state = 2
    partydone(CurrentPlayer) = 1
    PuPlayer.playlistplayex pCallouts,"audiocallouts","adventurepartyassembled.wav",100,1
    chilloutthemusic
  End Sub

  Sub endparty
    ruleshelperon
    PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,3
    PuPlayer.SetLoop 2,0
    GiOff
    GiOn
    PuPlayer.playresume 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1
    PuPlayer.SetLoop 7,0
    ResetPartyLights
    bMultiBallMode = False
    PartyMultiball = False
    CheckMONSTER
    GiOn
    FlashEffect 0
    Flashxmas 0
    LightSeqFlasher.StopPlay

    xmasqueue.RemoveAll(True)
'   LightSeqaxmas.StopPlay
    endamultiball

  End Sub

  Sub ResetPartyLights
    lucas.state = 0
    jonathan.state = 0
    nancy.state = 0
    joyce.state = 0
    will.state = 0
    elevencrew.state = 0
    mikey.state = 0
    hopper.state = 0
    dustin.state = 0
    lucason(CurrentPlayer) = 0
    jonathanon(CurrentPlayer) = 0
    nancyon(CurrentPlayer) = 0
    joyceon(CurrentPlayer) = 0
    willon(CurrentPlayer) = 0
    elevencrewon(CurrentPlayer) = 0
    mikeyon(CurrentPlayer) = 0
    hopperon(CurrentPlayer) = 0
    dustinon(CurrentPlayer) = 0
  End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  FLIPPER LANES
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'


  Sub lane1_Hit
    PlaySoundAt "fx_sensor", ActiveBall
    If bMultiBallMode = False Then
      PlaySound "lane"
      DOF 144, DOFPulse
      DOF 313, DOFPulse   'DOF MX - Left Outer Lane
    End If
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    ll1.State = 1
    ll8.State = 1
    AddScore 50050
    ' Do some sound or light effect

    LastSwitchHit = "lane1"
    ' do some check
    CheckKIDSLane
  End Sub

  Sub lane2_Hit
    PlaySoundAt "fx_sensor", ActiveBall
    If bMultiBallMode = False Then
      PlaySound "lane"
      DOF 145, DOFPulse
      DOF 315, DOFPulse   'DOF MX - Left Inner Lane
    End If
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    ll2.State = 1
    ll7.State = 1
    AddScore 10010
    ' Do some sound or light effect

    LastSwitchHit = "lane2"
    ' do some check
    CheckKIDSLane
  End Sub

  Sub lane3_Hit
    PlaySoundAt "fx_sensor", ActiveBall
    If bMultiBallMode = False Then
      PlaySound "lane"
      DOF 146, DOFPulse
      DOF 315, DOFPulse   'DOF MX - Left Inner Lane
    End If
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    ll3.State = 1
    ll6.State = 1
    AddScore 10010
    ' Do some sound or light effect

    LastSwitchHit = "lane3"
    ' do some check
    CheckKIDSLane
  End Sub

  Sub lane4_Hit
    PlaySoundAt "fx_sensor", ActiveBall
    If bMultiBallMode = False Then
      PlaySound "lane"
      DOF 147, DOFPulse
      DOF 316, DOFPulse   'DOF MX - Right Inner Lane
    End If
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    ll4.State = 1
    ll9.State = 1
    AddScore 10010
    ' Do some sound or light effect

    LastSwitchHit = "lane4"
    ' do some check
    CheckKIDSLane
  End Sub

  Sub lane5_Hit
    PlaySoundAt "fx_sensor", ActiveBall
    If bMultiBallMode = False Then
      PlaySound "lane"
      DOF 148, DOFPulse
      DOF 314, DOFPulse   'DOF MX - Right Outer Lane
    End If
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    ll5.State = 1
    ll10.State = 1
    AddScore 50050
    ' Do some sound or light effect

    LastSwitchHit = "lane5"
    ' do some check
    CheckKIDSLane
  End Sub

  Sub CheckKIDSLane() 'use the lane lights
    If ll1.State + ll2.State + ll3.State + ll4.State + ll5.State = 5 Then
      'Activate Ball Saver
      If kickbacklight.state = 0 Then
        PuPlayer.playlistplayex pCallouts,"audiocallouts","kickback.wav",100,1
    chilloutthemusic
        pNote "KICKBACK","ACTIVATED"
      ResetKidsLights
      End If
      kickinback

      ''''DMD "black.png", "Ball Save","Activated",  500
    End If
  End Sub


  Sub kickinback
    kickbackon(CurrentPlayer) = 1
    kickbackgate.open = True
    kickbacklight.state = 1
    kickbacklight1.state = 1
      kbon.opacity = 1000
      kboff.opacity = 0
  End Sub

  Sub Kickback_hit
    if bBallSaverActive = false Then
      EnableBallSaver 5
    end if
    kickback.Kick 0, 60
    PlaySoundAt SoundFXDOF("fx_kicker", 103, DOFPulse, DOFContactors), kickback
    DOF 309, DOFPulse   'DOF MX - Kickback
    PuPlayer.playlistplayex pBackglass,"videokickback","",100,1
    kickbackon(CurrentPlayer) = 0
    kickbackgate.open = False
    kickbacklight.state = 0
    kickbacklight1.state = 0
      kbon.opacity = 0
      kboff.opacity = 1000
  End Sub

  Sub ResetKidsLights
    ll1.State = 0
    ll8.State = 0
    ll2.State = 0
    ll7.State = 0
    ll3.State = 0
    ll6.State = 0
    ll4.State = 0
    ll9.State = 0
    ll5.State = 0
    ll10.State = 0
  End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  FIND WILL
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'

  Sub walllt_hit
    walllon(CurrentPlayer) = 1
    wallll.State = 1
    wallll1.State = 1
    checkbank
    PlaySoundAt SoundFXDOF("fx_target", 201, DOFPulse, DOFTargets), ActiveBall
    DOF 340, DOFPulse  'DOF MX - Barricade Targets
    If WillMultiball = True Then
    Else
      beaconon
      Dim waittime
      waittime = 700
      vpmtimer.addtimer waittime, "beaconoff'"
    End If
  End Sub

  Sub wallmt_hit
    wallmon(CurrentPlayer) = 1
    wallml.State = 1
    wallml1.State = 1
    checkbank
    PlaySoundAt SoundFXDOF("fx_target", 201, DOFPulse, DOFTargets), ActiveBall
    DOF 340, DOFPulse  'DOF MX - Barricade Targets
    If WillMultiball = True Then
    Else
      beaconon
      Dim waittime
      waittime = 700
      vpmtimer.addtimer waittime, "beaconoff'"
    End If
  End Sub

  Sub wallrt_hit
    wallron(CurrentPlayer) = 1
    wallrl.State = 1
    wallrl1.State = 1
    checkbank
    PlaySoundAt SoundFXDOF("fx_target", 201, DOFPulse, DOFTargets), ActiveBall
    DOF 340, DOFPulse  'DOF MX - Barricade Targets
    If WillMultiball = True Then
    Else
      beaconon
      Dim waittime
      waittime = 700
      vpmtimer.addtimer waittime, "beaconoff'"
    End If
  End Sub

  Sub checkbank
    If WillMultiball = True Then
      If wallml.state = 2 or wallll.state = 2 or wallrl.state = 2 Then
        AwardWill
      End If
    End If

    If partymultiball = False Then
      If mikey.state = 0 Then
        mikey.state = 1
        mikeyon(CurrentPlayer) = 1
        partycollected(CurrentPlayer) = partycollected(CurrentPlayer) + 1
        PuPlayer.LabelSet pBackglass,"partyh","" & partycollected(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
        checkparty
      End If
    End If

    If mikey.state = 2 Then
      mikey.state = 0
      awardparty
    End If

    If WillMultiball = False Then
        If wallll.State + wallml.State + wallrl.State = 3 Then
          'PuPlayer.playlistplayex pCallouts,"audiocallouts","bashthedemogorgon.wav",100,1
          If bMultiBallMode = false and inmode = 0 Then
          dropwallskip = 1
          Dim waittime
          waittime = 16000
          gioff
          'vpmtimer.addtimer waittime, "Dropwall'"
          BallHandlingQueue.Add "Dropwall","Dropwall",95,waittime,0,0,0,True
          MagnetW.MagnetON = True
          spinner.MotorOn = True
          PuPlayer.playpause 4

          myFSpeed = 80

          PuPlayer.playlistplayex pBackglass,"videowill","barricade down.mov",100,1
          GeneralPupQueue.Add "StopXmas","StopXmas",95,2000,0,0,0,True
          GeneralPupQueue.Add "Spell-R","Spell ""r"" ",95,3200,0,0,0,True
          GeneralPupQueue.Add "Spell-U","Spell ""u"" ",95,5500,0,0,0,True
          GeneralPupQueue.Add "Spell-N","Spell ""n"" ",95,8700,0,0,0,True
          'GeneralPupQueue.Add "StartDemoLightSeq","StartDemoLightSeq",95,12600,0,0,0,True


          DOF_Shaker_R.enabled = true  'DOF - Shaker pulse for "R"
          DOF_Shaker_U.enabled = true  'DOF - Shaker pulse for "U"
          DOF_Shaker_N.enabled = true  'DOF - Shaker pulse for "N"
          DOF_Shaker_DemoWallBurst.enabled = true  'DOF - Shaker pulse for "DemoWallBurst"
          Else
          Dropwall
          End If
        End If
    End If

  End Sub

  '***************************
  'Beacon spinner
  '***************************


  Sub beacontime_Timer()
    beaconflash.RotZ = beaconflash.RotZ + 2
    beaconflash1.RotZ = beaconflash1.RotZ + 2
    reflector2.RotZ = reflector2.RotZ + 3
    reflector1.RotZ = reflector1.RotZ + 3
  End Sub

  Sub beaconon
    PlaySound "klaxon3"
    beacontime.Enabled = True
    beaconlight1.State = 1
    beaconlight.State = 1
    beaconflash.opacity = 30
    beaconflash1.opacity = 30
    DOF 319, DOFON ' DOF MX - Beacon ON
    DOF 127, DOFOn   'DOF - Beacon - ON
  End Sub

  Sub beaconoff
    beacontime.Enabled = False
    beaconlight1.State = 0
    beaconlight.State = 0
    beaconflash.opacity = 0
    beaconflash1.opacity = 0
    DOF 319, DOFOFF ' DOF MX - Beacon OFF
    DOF 127, DOFOff   'DOF - Beacon - OFF
  End Sub

  Dim openwall
  openwall = false
  Dim openwall2
  openwall2 = false
  Dim demoout
  demoout = False
  Dim barricadedowncurrently
  barricadedowncurrently = False



  '*********************
  ' Demogorgon walk
  '**********************

  '*****************
  ' Maths
  '*****************
  'Const Pi = 3.1415927
' Function dSin(degrees)
'   dsin = sin(degrees * Pi/180)
' End Function
'
' Function dCos(degrees)
'   dcos = cos(degrees * Pi/180)
' End Function

  Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
  End Function


  dim twalk,rndD,Ddir:ddir=1
  Sub demWalk_Timer

    twalk=twalk+0.8*Ddir

    if dgs= 1 then
      rndD=RndNum(1,3):if rndD=1 then Ddir=ddir*-1
    end if

    Dbody.rotz=dsin(twalk)*20
    Dleft.rotz=dsin(twalk)*20
    Dright.rotz=dsin(twalk)*20
    Dhead.rotz=dsin(twalk)*-10+(rndD*3)
        Dhead001.rotz=dsin(twalk)*-10+(rndD*3)
        Dhead002.rotz=dsin(twalk)*-10+(rndD*3)
    Dleft.rotx=dsin(twalk)*20
    Dright.rotx=dsin(twalk+180)*20

  End Sub


  Sub demogorgontime_Timer()
    If demoout = False Then
      dbody.y = dbody.y + 1
      dleft.y = dbody.y
      Dright.y = dbody.y
      Dhead.y = dbody.y
      Dlegs.y = dbody.y
      thedemogorgonmodel1.y = dbody.y
      If  dbody.y = 280 Then
        demogorgontime.Enabled = False
        demoout = True
      End If
    End If
    If demoout = True Then
      dbody.y = dbody.y - 1
      dleft.y = dbody.y
      Dright.y = dbody.y
      Dhead.y = dbody.y
      Dlegs.y = dbody.y
      thedemogorgonmodel1.y = dbody.y
      If  dbody.y = -85 Then
        demogorgontime.Enabled = False
        closewall
        demoout = False
      End If
    End If
  End Sub

' Sub demogorgontime_Timer()
'   If demoout = False Then
'     thedemogorgonmodel.TransY = thedemogorgonmodel.TransY + 1
'     thedemogorgonmodel1.TransY= thedemogorgonmodel.TransY
'     If  thedemogorgonmodel.TransY = 370 Then
'       demogorgontime.Enabled = False
'       demoout = True
'     End If
'   End If
'   If demoout = True Then
'     thedemogorgonmodel.TransY = thedemogorgonmodel.TransY - 1
'     thedemogorgonmodel1.TransY= thedemogorgonmodel.TransY
'     If  thedemogorgonmodel.TransY = 0 Then
'       demogorgontime.Enabled = False
'       demoout = False
'     End If
'   End If
' End Sub

  Sub barricadetime_Timer()
    If barricadedowncurrently = False Then
      spikebat.TransZ = spikebat.TransZ - 1
      backbank.TransZ = backbank.TransZ - 1
      walllt.TransZ = walllt.TransZ - 1
      wallmt.TransZ = wallmt.TransZ - 1
      wallrt.TransZ = wallrt.TransZ - 1
      If spikebat.TransZ = -180 Then
        barricadetime.Enabled = False
        barricadedowncurrently = True
        walllt.Collidable = False
        wallmt.Collidable = False
        wallrt.Collidable = False
      End If
    End If
    If barricadedowncurrently = True Then
      spikebat.TransZ = spikebat.TransZ + 1
      backbank.TransZ = backbank.TransZ + 1
      walllt.TransZ = walllt.TransZ + 1
      wallmt.TransZ = wallmt.TransZ + 1
      wallrt.TransZ = wallrt.TransZ + 1
      If spikebat.TransZ = 0 Then
        barricadetime.Enabled = False
        barricadedowncurrently = False
        walllt.Collidable = True
        wallmt.Collidable = True
        wallrt.Collidable = True
      End If
    End If
  End Sub

  Sub udtargettime_Timer
    If udtargetdown(CurrentPlayer) = 0 Then
      udtarget.TransZ = udtarget.TransZ - 1
      If udtarget.TransZ = -82 Then
        udtargetdown(CurrentPlayer) = 1
        udtarget.Collidable = False
        udtargettime.Enabled = False
      End If
    End If
    If udtargetdown(CurrentPlayer) = 1 Then
      udtarget.TransZ = udtarget.TransZ + 1
      If udtarget.TransZ = -22 Then
        udtargetdown(CurrentPlayer) = 0
        udtarget.Collidable = True
        udtargettime.Enabled = False
      End If
    End If
  End Sub

Sub StartDemoLightSeq

Debug.print "Start DEMO Light"
  dim i
  for i = 0 to axmas.count-1
    axmas(i).opacity = 2000
  Next

  LightSeqaxmas.UpdateInterval = 150
  LightSeqaxmas.Play SeqRandom, 10, , 6000
  GeneralPupQueue.Add "ResetDemoLightSeq","ResetDemoLightSeq",95,6050,0,0,0,True
End Sub

Sub ResetDemoLightSeq
dim i
  for i = 0 to axmas.count-1
    axmas(i).opacity = 0
  Next

  LightSeqaxmas.StopPlay
  GeneralPupQueue.Add "StartDemoLightSeq","StartDemoLightSeq",95,100,0,0,0,True

End Sub

Sub StopDemoLightSeq
  dim i
  for i = 0 to axmas.count-1
    axmas(i).opacity = 0
  Next

  LightSeqaxmas.StopPlay

End Sub

  Sub Dropwall
    if dropwallskip = 0 then exit Sub
    GeneralPupQueue.Add "StopXMAS","StopXMAS",80,100,0,0,0,False
    Spot1.opacity = 1000
    DOF 116, DOFPulse  'DOF - Gear Motor
    PlaySound "Bridge_Move"
'   LightSeqaxmas.UpdateInterval = 150
'   LightSeqaxmas.Play SeqRandom, 10, , 3000
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 3000
    MagnetW.MagnetON = True
    spinner.MotorOn = True
    beaconon
    backwallleft.Enabled = True
    backwallright.Enabled = True
    demogorgontime.Enabled = True
    barricadetime.Enabled = True
    Dim waittime
    waittime = 3000
    barricadedown(CurrentPlayer) = 1
    vpmtimer.addtimer waittime, "wallrelease'"
  End Sub

  Sub wallrelease
    dropwallskip = 0
    If bMultiBallMode = False and inmode = 0 then
    PuPlayer.playresume 4
    gion
    End If
    PlaySound "Bridge_Stop"
    Dim waittime
    waittime = 1000
    vpmtimer.addtimer waittime, "spinoff'"
    MagnetW.MagnetON = False
    If WillMultiball = true Then
    Else
      beaconoff
    End If
  End Sub

  Sub spinoff
    StopDemoLightSeq
    LightSeqFlasher.StopPlay
    LightSeqaxmas.StopPlay
    spinner.MotorOn = False
  End Sub

  Sub Raisewall
    StopDemoLightSeq
    'StartXMAS
    Spot1.opacity = 0
    DOF 116, DOFPulse  'DOF - Gear Motor
    PlaySound "Bridge_Move"
    If bMultiBallMode = false Then
    GiOn
    End If
    MagnetW.MagnetON = False
    If udtargetdown(CurrentPlayer) = 1 Then
      udtargettime.Enabled = True
    End If
    If openwall = True Then
      demogorgontime.Enabled = True
    End If
  End Sub

Sub closewall
  backwallleft.Enabled = True
  backwallright.Enabled = True
  barricadetime.Enabled = True
End Sub


  Sub backwallleft_Timer()
    If openwall = False Then
      wallleft.TransX = wallleft.TransX - 1
      If wallleft.TransX = -120 Then
      backwallleft.Enabled = False
      openwall = True
      End If
    End If
    If openwall = True Then
      wallleft.TransX = wallleft.TransX + 1
      If wallleft.TransX = 0 Then
      backwallleft.Enabled = False
      openwall = False
      End If
    End If
  End Sub

  Sub backwallright_Timer()
    If openwall2 = False Then

      wallright.TransX = wallright.TransX + 1
      If wallright.TransX = 120 Then
      backwallright.Enabled = False
      openwall2 = True
      End If
    End If
    If openwall2 = True Then

      wallright.TransX = wallright.TransX - 1
      If wallright.TransX = 0 Then
      backwallright.Enabled = False
      openwall2 = False
      End If
    End If
  End Sub


  Sub udtarget_hit
    beaconon
    DGs=1: playsound "dgshoutsoft"
    DOF 203, DOFPulse  'DOF - Fan
    DOF 341, DOFPulse  'DOF MX - Upside Down Targets
    PlaySoundAt SoundFXDOF("fx_target", 202, DOFPulse, DOFTargets), ActiveBall
    Dim waittime
    waittime = 700
    vpmtimer.addtimer waittime, "beaconoff'"
    If NOT Toysbouncing.Enabled Then
      Toysbouncing.Enabled = 1:brake=0:perc=3:sbou=(Rndnum(-10,10))/10
    End If
    checkhits
  End Sub


  Sub udtarget1_hit
    beaconon
    DOF 341, DOFPulse  'DOF MX - Upside Down Targets
    PlaySoundAt SoundFXDOF("fx_target", 202, DOFPulse, DOFTargets), ActiveBall
    Dim waittime
    waittime = 700
    vpmtimer.addtimer waittime, "beaconoff'"
    If NOT Toysbouncing.Enabled Then
      Toysbouncing.Enabled = 1:brake=0:perc=3:sbou=(Rndnum(-10,10))/10
    End If
    checkhits
  End Sub


  Sub udtarget2_hit
    beaconon
    DOF 341, DOFPulse  'DOF MX - Upside Down Targets
    PlaySoundAt SoundFXDOF("fx_target", 202, DOFPulse, DOFTargets), ActiveBall
    Dim waittime
    waittime = 700
    vpmtimer.addtimer waittime, "beaconoff'"
    If NOT Toysbouncing.Enabled Then
      Toysbouncing.Enabled = 1:brake=0:perc=3:sbou=(Rndnum(-10,10))/10
    End If
    checkhits
  End Sub


  Sub checkhits
    If mikey.state = 2 Then
      mikey.state = 0
      awardparty
    End If

    If DemoMultiball = True Then
      KillMonster
    End If

    udhits(CurrentPlayer) = udhits(CurrentPlayer) + 1
    PuPlayer.LabelSet pBackglass,"Willh","" & udhits(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 8.1, 'xalign': 1}"
    If udfirst(CurrentPlayer) = 0 Then
      If udhits(CurrentPlayer) > 15 Then
        openupsidedown
      End If
    Else
      If udhits(CurrentPlayer) > 30 Then
        openupsidedown
      End If
    End If
    Dim waittime
    waittime = 700
    vpmtimer.addtimer waittime, "stopscreaming'"
  End Sub


  Sub stopscreaming
    DGs=0
  End Sub

  Sub openupsidedown
    If udtargetdown(CurrentPlayer) = 0 Then
      udtargettime.Enabled = True
    End If
    If bMultiBallMode = False and inmode = 0 Then
      If upsidedownarrow.state = 0 Then
        PuPlayer.playlistplayex pCallouts,"audiocallouts","hittheupsidedown.wav",100,1
    chilloutthemusic
      End If
      upsidedownarrow.state = 2
      upsidedowncircle.state = 2
      BallLockEscape.enabled = True
      upsidedownlights(CurrentPlayer) = 1
    End If
  End Sub



  'upside down


  '*****************
  '  ESCAPE Targets
  '*****************

  Sub escape1_Hit
    DOF 352, DOFPulse   'DOF MX - Escape E
    PlaySoundAt SoundFXDOF("fx_target", 135, DOFPulse, DOFTargets), ActiveBall
    If Tilted Then Exit Sub
    If le2.State = 0 Then
    le2.State = 1
    le16.State = 1
    le1.State = 0
    le17.State = 0
    le2on(CurrentPlayer) = 1
    Exit Sub
    End If
    If le2.State = 1 Then
    le2.State = 1
    le16.State = 1
    le1.State = 1
    le17.State = 1
    le1on(CurrentPlayer) = 1
    End If
    AddScore 25010
    LastSwitchHit = "escape1"
    ' do some check
    CheckESCAPETargets
  End Sub

  Sub escape2_Hit
    DOF 353, DOFPulse   'DOF MX - Escape S
    PlaySoundAt SoundFXDOF("fx_target", 135, DOFPulse, DOFTargets), ActiveBall
    If Tilted Then Exit Sub
    If le4.State = 0 Then
    le4.State = 1
    le19.State = 1
    le3.State = 0
    le18.State = 0
    le4on(CurrentPlayer) = 1
    Exit Sub
    End If
    If le4.State = 1 Then
    le4.State = 1
    le19.State = 1
    le3.State = 1
    le18.State = 1
    le3on(CurrentPlayer) = 1
    End If
    AddScore 25010

    LastSwitchHit = "escape2"
    ' do some check
    CheckESCAPETargets
  End Sub

  Sub escape3_Hit
    DOF 354, DOFPulse   'DOF MX - Escape C
    PlaySoundAt SoundFXDOF("fx_target", 135, DOFPulse, DOFTargets), ActiveBall
    If Tilted Then Exit Sub
    If le6.State = 0 Then
    le6.State = 1
    le21.State = 1
    le5.State = 0
    le20.State = 0
    le6on(CurrentPlayer) = 1
    Exit Sub
    End If
    If le6.State = 1 Then
    le6.State = 1
    le21.State = 1
    le5.State = 1
    le20.State = 1
    le5on(CurrentPlayer) = 1
    End If
    AddScore 25010
    LastSwitchHit = "escape3"
    ' do some check
    CheckESCAPETargets
  End Sub

  Sub escape4_Hit
    DOF 355, DOFPulse   'DOF MX - Escape A
    PlaySoundAt SoundFXDOF("fx_target", 135, DOFPulse, DOFTargets), ActiveBall
    If Tilted Then Exit Sub
    If le11.State = 0 Then
    le11.State = 1
    le25.State = 1
    le10.State = 0
    le26.State = 0
    le11on(CurrentPlayer) = 1
    Exit Sub
    End If
    If le11.State = 1 Then
    le11.State = 1
    le25.State = 1
    le10.State = 1
    le26.State = 1
    le10on(CurrentPlayer) = 1
    End If
    AddScore 25010
    LastSwitchHit = "escape4"
    ' do some check
    CheckESCAPETargets
  End Sub

  Sub escape5_Hit
    DOF 356, DOFPulse   'DOF MX - Escape P
    PlaySoundAt SoundFXDOF("fx_target", 135, DOFPulse, DOFTargets), ActiveBall
    If Tilted Then Exit Sub
    If le13.State = 0 Then
    le13.State = 1
    le28.State = 1
    le12.State = 0
    le27.State = 0
    le13on(CurrentPlayer) = 1
    Exit Sub
    End If
    If le13.State = 1 Then
    le13.State = 1
    le28.State = 1
    le12.State = 1
    le27.State = 1
    le12on(CurrentPlayer) = 1
    End If
    AddScore 25010
    LastSwitchHit = "escape5"
    ' do some check
    CheckESCAPETargets
  End Sub

  Sub escape6_Hit
    DOF 357, DOFPulse   'DOF MX - Escape E
    PlaySoundAt SoundFXDOF("fx_target", 135, DOFPulse, DOFTargets), ActiveBall
    If Tilted Then Exit Sub
    If le15.State = 0 Then
    le15.State = 1
    le29.State = 1
    le15on(CurrentPlayer) = 1
    le14.State = 0
    le30.State = 0
    Exit Sub
    End If
    If le15.State = 1 Then
    le15.State = 1
    le29.State = 1
    le14.State = 1
    le30.State = 1
    le14on(CurrentPlayer) = 1
    End If
    AddScore 25010
    LastSwitchHit = "escape6"
    ' do some check
    CheckESCAPETargets
  End Sub

  Sub CheckESCAPETargets
  If le1.State + le2.State + le3.State + le4.State + le5.State + le6.State + le10.State + le11.State + le12.State + le13.State + le14.State + le15.State = 12 Then
    Gate8.Open = True
    escapegateopen(currentplayer) = 1
    If le9.state = 0 Then
      PuPlayer.playlistplayex pCallouts,"audiocallouts","hurryescapetheupsidedown.wav",100,1
    chilloutthemusic
      PlaySound "bell"
    End If
    MagnetU.MagnetON = True ' Magnet On
    le7.State = 2
    le23.State = 2
    le7on(CurrentPlayer) = 1
    le8.State = 2
    le22.State = 2
    le8on(CurrentPlayer) = 1
    le9.State = 2
    le24.State = 2
    le9on(CurrentPlayer) = 1
  End If
  End Sub


  Sub ResetESCAPELights
    le1on(CurrentPlayer) = 0
    le2on(CurrentPlayer) = 0
    le3on(CurrentPlayer) = 0
    le4on(CurrentPlayer) = 0
    le5on(CurrentPlayer) = 0
    le6on(CurrentPlayer) = 0
    le7on(CurrentPlayer) = 0
    le8on(CurrentPlayer) = 0
    le9on(CurrentPlayer) = 0
    le10on(CurrentPlayer) = 0
    le11on(CurrentPlayer) = 0
    le12on(CurrentPlayer) = 0
    le13on(CurrentPlayer) = 0
    le14on(CurrentPlayer) = 0
    le15on(CurrentPlayer) = 0
    le1.State = 0
    le17.State = 0
    le2.State = 0
    le16.State = 0
    le3.State = 0
    le18.State = 0
    le4.State = 0
    le19.State = 0
    le5.State = 0
    le20.State = 0
    le6.State = 0
    le21.State = 0
    le7.State = 0
    le23.State = 0
    le8.State = 0
    le22.State = 0
    le9.State = 0
    le24.State = 0
    le10.State = 0
    le26.State = 0
    le11.State = 0
    le25.State = 0
    le12.State = 0
    le27.State = 0
    le13.State = 0
    le28.State = 0
    le14.State = 0
    le30.State = 0
    le15.State = 0
    le29.State = 0
  End Sub



  Sub BallLockEscape_Hit

    ruleshelperoff
    GiOff
    PuPlayer.playpause 4
    PuPlayer.LabelSet pBackglass,"notetitle","The Upside Down",1,""
    PuPlayer.LabelSet pBackglass,"notecopy","Collect all Escape target to open the exit \r Escape to start Will Multiball",1,""
    If udfirst(CurrentPlayer) = 0 Then
      willskip = 1
      PuPlayer.playlistplayex pBackglass,"videowill","enterupsidedown.mov",100,1
      DOF_Shaker_WillHiding.enabled = true  'DOF - Shaker pulse for "WillHiding"
      udfirst(CurrentPlayer) = 1
      dim waittime
      waittime = 26000
      vpmtimer.addtimer 22000, "udwarning'"
      vpmtimer.addtimer waittime, "EnterUpsideDown'"
    Else
      willskip = 1
      EnterUpsideDown

    End If
  End Sub


  Sub udwarning
    pNote "BEWARE UPSIDE DOWN KICKOUT","FLIPPERS REVERSED!"
  End Sub


  Sub EnterUpsideDown
    If willskip = 0 then exit Sub
    PuPlayer.playlistplayex pBackglass,"videowill","upsidedown1.mov",100,3
    PuPlayer.SetLoop 2,1
    PuPlayer.playpause 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","upsidedown.mp3",100,1
    PuPlayer.SetLoop 7,1
    SolULFlipper 1
    SolULFlipper 0
    SolULFlipper 1
    SolULFlipper 0
    SolURFlipper 1
    SolURFlipper 0
    SolURFlipper 1
    SolURFlipper 0
    SolLFlipper 0
    SolRFlipper 0
    BallLockEscape.DestroyBall
    lowerflippersoff = False
    If le1.State + le2.State + le3.State + le4.State + le5.State + le6.State + le10.State + le11.State + le12.State + le13.State + le14.State + le15.State = 0 Then
    Dim waittime
    waittime = 1000
    vpmtimer.addtimer waittime, "UpsideDown'"
    PlaySound "monster2"
    ''''DMD "will_start-14.wmv", "", "",  14000
    DOF 415, DOFPulse  'DOF MX - Enter Portal
    DOF 203, DOFPulse  'DOF - Fan
    GiOff
    GiLowerOn
    Else
    waittime = 1000
    vpmtimer.addtimer waittime, "UpsideDown'"
    PlaySound "bell"
    DOF 410, DOFOn   'DOF MX - Upside Down - ON
    ''''DMD "black.png", "Entering", "Upside Down",  1000
    GiOff
    GiLowerOn
    End If
    willskip = 0
  End Sub

  Sub UpsideDown
    BallEscapeRelease.CreateBall
    BallEscapeRelease.Kick 50, 7
    PlaySoundAt "fx_kicker", BallEscapeRelease
    DOF 410, DOFOn   'DOF MX - Upside Down - ON
  End Sub

  Sub BallEscapeDrain_Hit
        MagnetU.MagnetON = False ' Magnet On
    Flipper2.RotateToStart
    Flipper1.RotateToStart
    PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,3
    PuPlayer.SetLoop 2,0
    PuPlayer.playresume 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1
    PuPlayer.SetLoop 7,0
    BallEscapeDrain.DestroyBall
    BallLockEscape.CreateBall
    BallLockEscape.Kick 190, 17
    PlaySoundAt "fx_kicker", BallEscapeDrain
    PlaySoundAt SoundFXDOF("fx_Ballrel", 138, DOFPulse, DOFContactors), BallLockEscape
    DOF 115, DOFPulse
    GiOn
    GiLowerOff
    ruleshelperon
    DOF 410, DOFOff   'DOF MX - Upside Down - OFF
    lowerflippersoff = True
  End Sub



  Sub BallEscape_Hit
    Flipper2.RotateToStart
    Flipper1.RotateToStart
    PuPlayer.playresume 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1
    PuPlayer.SetLoop 7,0
    RandomSoundHole
    ResetESCAPELights
    PlaySound "demogorgon"
    BallEscape.DestroyBall
    BallLockEscape.CreateBall
    BallLockEscape.Kick 190, 17
    PlaySoundAt "fx_kicker",BallLockEscape
    PlaySoundAt SoundFXDOF("fx_Ballrel", 138, DOFPulse, DOFContactors), BallLockEscape
    DOF 115, DOFPulse
    DOF 410, DOFOff   'DOF MX - Upside Down - OFF
    Gate8.Open = False
    escapegateopen(currentplayer) = 0
    MagnetU.MagnetON = False
    le7.State = 0
    le23.State = 0
    le8.State = 0
    le22.State = 0
    le9.State = 0
    le24.State = 0
    SavedWill
    GiLowerOff
    lowerflippersoff = True
  End Sub

  Sub SavedWill() 'Multiball
    pNote "SAVE WILL","MULTIBALL"
    PuPlayer.playlistplayex pBackglass,"videowill","willmbmid.mov",100,3
    PuPlayer.SetLoop 2,1
    ruleshelperoff
    PuPlayer.LabelSet pBackglass,"notetitle","Save Will Multiball",1,""
    PuPlayer.LabelSet pBackglass,"notecopy","Follow the sequence of shots between ramps and barricade \r Complete all 12 Shots to save Will!",1,""
    PuPlayer.playpause 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","will.mp3",100,1
    PuPlayer.SetLoop 7,1
    Raisewall
    barricadedown(CurrentPlayer) = 0
    udhits(CurrentPlayer) = 0
    upsidedownarrow.state = 0
    upsidedowncircle.state = 0
    upsidedownlights(CurrentPlayer) = 0
    SetLightColor llo3, darkgreen, -1
    SetLightColor llo7, darkgreen, -1
    SetLightColor leftrampred, darkgreen, -1
    SetLightColor leftrampred2, darkgreen, -1
    SetLightColor lc1, darkgreen, -1
    SetLightColor lc4, darkgreen, -1
    SetLightColor leftrampred1, darkgreen, -1
    SetLightColor leftrampred3, darkgreen, -1
    SetLightColor rightrampred, darkgreen, -1
    SetLightColor rightrampred1, darkgreen, -1
    SetLightColor llr1, darkgreen, -1
    SetLightColor llr2, darkgreen, -1
    SetLightColor wallll, darkgreen, -1
    SetLightColor wallll1, darkgreen, -1
    SetLightColor wallml, darkgreen, -1
    SetLightColor wallml1, darkgreen, -1
    SetLightColor wallrl, darkgreen, -1
    SetLightColor wallrl1, darkgreen, -1
    wallll.state = 2
    wallll1.state = 2
    wallml.state = 2
    wallml1.state = 2
    wallrl.state = 2
    wallrl1.state = 2
    DOF 418, DOFPulse    'DOF MX - WILL Flash
    bMultiBallMode = True
    WillMultiball = True
    AddMultiball 2
    EnableBallSaver 15
    GiOff
    GiGreen
    flashflash.Enabled = True


    StartDemoLightSeq
    GeneralPupQueue.Add "StopDemoLightSeq","StopDemoLightSeq",80,50000,0,0,0,False

'   LightSeqaxmas.UpdateInterval = 150
'   LightSeqaxmas.Play SeqRandom, 10, , 50000
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
    startamultiball
    beaconon
      Dim waittime
      waittime = 1500
      vpmtimer.addtimer waittime, "beaconon'"
  End Sub

  Sub AwardWill
    WillHits(CurrentPlayer) = WillHits(CurrentPlayer) + 1
    PuPlayer.LabelSet pBackglass,"Willj","" & WillHits(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0}"
    Select Case WillHits(CurrentPlayer)
      Case 1
        AddScore 1000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","1,000,000"
        PlaySound "nancy"
        llo3.state = 2
        llo7.state = 2
        leftrampred.state = 2
        leftrampred2.state = 2
        wallll.state = 0
        wallll1.state = 0
        wallml.state = 0
        wallml1.state = 0
        wallrl.state = 0
        wallrl1.state = 0
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 2
        AddScore 4000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","4,000,000"
        PlaySound "nancy"
        wallll.state = 2
        wallll1.state = 2
        wallml.state = 2
        wallml1.state = 2
        wallrl.state = 2
        wallrl1.state = 2
        llo3.state = 0
        llo7.state = 0
        leftrampred.state = 0
        leftrampred2.state = 0
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 3
        AddScore 1000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","1,000,000"
        PlaySound "nancy"
        lc1.state = 2
        lc4.state = 2
        leftrampred1.state = 2
        leftrampred3.state = 2
        wallll.state = 0
        wallll1.state = 0
        wallml.state = 0
        wallml1.state = 0
        wallrl.state = 0
        wallrl1.state = 0
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 4
        AddScore 4000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","4,000,000"
        PlaySound "nancy"
        wallll.state = 2
        wallll1.state = 2
        wallml.state = 2
        wallml1.state = 2
        wallrl.state = 2
        wallrl1.state = 2
        lc1.state = 0
        lc4.state = 0
        leftrampred1.state = 0
        leftrampred3.state = 0
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 5
        AddScore 1000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","1,000,000"
        PlaySound "nancy"
        llr1.state = 2
        llr2.state = 2
        rightrampred.state = 2
        rightrampred1.state = 2
        wallll.state = 0
        wallll1.state = 0
        wallml.state = 0
        wallml1.state = 0
        wallrl.state = 0
        wallrl1.state = 0
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 6
        AddScore 4000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","4,000,000"
        PlaySound "nancy"
        wallll.state = 2
        wallll1.state = 2
        wallml.state = 2
        wallml1.state = 2
        wallrl.state = 2
        wallrl1.state = 2
        llr1.state = 0
        llr2.state = 0
        rightrampred.state = 0
        rightrampred1.state = 0
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 7
        AddScore 1000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","1,000,000"
        PlaySound "nancy"
        llo3.state = 2
        llo7.state = 2
        leftrampred.state = 2
        leftrampred2.state = 2
        wallll.state = 0
        wallll1.state = 0
        wallml.state = 0
        wallml1.state = 0
        wallrl.state = 0
        wallrl1.state = 0
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 8
        AddScore 4000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","4,000,000"
        PlaySound "nancy"
        wallll.state = 2
        wallll1.state = 2
        wallml.state = 2
        wallml1.state = 2
        wallrl.state = 2
        wallrl1.state = 2
        llo3.state = 0
        llo7.state = 0
        leftrampred.state = 0
        leftrampred2.state = 0
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 9
        AddScore 1000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","1,000,000"
        PlaySound "nancy"
        lc1.state = 2
        lc4.state = 2
        leftrampred1.state = 2
        leftrampred3.state = 2
        wallll.state = 0
        wallll1.state = 0
        wallml.state = 0
        wallml1.state = 0
        wallrl.state = 0
        wallrl1.state = 0
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 10
        AddScore 4000000
        DOF 404, DOFPulse   'DOF MX - Jackpot
        pNote "JACKPOT","4,000,000"
        PlaySound "nancy"
        wallll.state = 2
        wallll1.state = 2
        wallml.state = 2
        wallml1.state = 2
        wallrl.state = 2
        wallrl1.state = 2
        lc1.state = 0
        lc4.state = 0
        leftrampred1.state = 0
        leftrampred3.state = 0
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 11
        AddScore 1000000
        DOF 432, DOFPulse  ' DOF MX - Will Get SJ
        pNote "JACKPOT","1,000,000"
        PlaySound "nancy"
        llr1.state = 2
        llr2.state = 2
        rightrampred.state = 2
        rightrampred1.state = 2
        wallll.state = 0
        wallll1.state = 0
        wallml.state = 0
        wallml1.state = 0
        wallrl.state = 0
        wallrl1.state = 0
        PuPlayer.playlistplayex pCallouts,"audiocallouts","getthesuperjackpot.wav",100,1
    chilloutthemusic
      Case 12
        willsj
        wallll.state = 2
        wallll1.state = 2
        wallml.state = 2
        wallml1.state = 2
        wallrl.state = 2
        wallrl1.state = 2
        llr1.state = 0
        llr2.state = 0
        rightrampred.state = 0
        rightrampred1.state = 0
    End Select
  End Sub


  Sub WillSJ
    PuPlayer.playlistplayex pBackglass,"videowill","willend.mov",100,3
    Dim waittime
    waittime = 34000
    vpmtimer.addtimer waittime, "willrestart'"
    PlaySound "nancy"
    DOF 138, DOFPulse 'DOF Solenoid - Rear Center
    DOF 115, DOFPulse
    DOF 426, DOFPulse  ' DOF MX - Will Mode Complete
    AddScore 10000000
    pNote "WILL SAVED!","10,000,000"
    WillHits(CurrentPlayer) = 0
    willcompleted(CurrentPlayer) = 1
    SetLightColor lm3, darkgreen, -1
    SetLightColor lm7, darkgreen, -1
    lm3.State = 1
    lm7.State = 1
    WillSuperReady = False
    PuPlayer.playlistplayex pCallouts,"audiocallouts","yousavedwill.wav",100,1
    chilloutthemusic
  End Sub

  Sub willrestart
    If willmultiball = true Then
    PuPlayer.playlistplayex pBackglass,"videowill","willmbmid.mov",100,3
    PuPlayer.SetLoop 2,1
    End If
  End Sub

  Sub EndWill()
    ruleshelperon
    PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,3
    PuPlayer.SetLoop 2,0
    beaconoff
    PuPlayer.playresume 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1
    PuPlayer.SetLoop 7,0
    WillMultiball = False
    WillSuperReady = False
    bMultiBallMode = False
    lc1.State = 0
    lc4.State = 0
    lc2.State = 0
    lc3.State = 0
    llr1.state = 0
    llr2.state = 0
    rightrampred.state = 0
    rightrampred1.state = 0
    SetLightColor wallll, red, -1
    SetLightColor wallll1, red, -1
    SetLightColor wallml, red, -1
    SetLightColor wallml1, red, -1
    SetLightColor wallrl, red, -1
    SetLightColor wallrl1, red, -1
    wallll.state = 0
    wallll1.state = 0
    wallml.state = 0
    wallml1.state = 0
    wallrl.state = 0
    wallrl1.state = 0
    wallmon(CurrentPlayer) = 0
    walllon(CurrentPlayer) = 0
    wallron(CurrentPlayer) = 0
    lc1.state = 0
    lc4.state = 0
    leftrampred1.state = 0
    leftrampred3.state = 0
    llo3.state = 0
    llo7.state = 0
    leftrampred.state = 0
    leftrampred2.state = 0
    CheckMONSTER
    GiOff
    GiOn
    FlashEffect 0
    Flashxmas 0
    LightSeqFlasher.StopPlay
    LightSeqaxmas.StopPlay
    endamultiball
    barricadedown(CurrentPlayer) = 0
    Raisewall
  End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  BARB
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'
' TARGETS
'****************
  Sub barb1_Hit
    LightEffect 10
    DOF 334, DOFPulse   'DOF MX - Barb B
    PlaySoundAt SoundFXDOF("fx_target", 133, DOFPulse, DOFTargets), ActiveBall
    If Tilted Then Exit Sub
    lb1.State = 1
    lb6.State = 1
    AddScore 25010
    CheckBARBTargets
    barb1on(CurrentPlayer) = 1
  End Sub

  Sub barb2_Hit
    LightEffect 10
    DOF 335, DOFPulse   'DOF MX - Barb A
    PlaySoundAt SoundFXDOF("fx_target", 133, DOFPulse, DOFTargets), ActiveBall
    If Tilted Then Exit Sub
    lb2.State = 1
    lb5.State = 1
    AddScore 25010
    CheckBARBTargets
    barb2on(CurrentPlayer) = 1
  End Sub

  Sub barb3_Hit
    LightEffect 10
    DOF 336, DOFPulse   'DOF MX - Barb R
    PlaySoundAt SoundFXDOF("fx_target", 133, DOFPulse, DOFTargets), ActiveBall
    If Tilted Then Exit Sub
    lb3.State = 1
    lb8.State = 1
    AddScore 25010
    LastSwitchHit = "barb3"
    CheckBARBTargets
    barb3on(CurrentPlayer) = 1
  End Sub

  Sub barb4_Hit
    LightEffect 10
    DOF 337, DOFPulse   'DOF MX - Barb B
    PlaySoundAt SoundFXDOF("fx_target", 133, DOFPulse, DOFTargets), ActiveBall
    If Tilted Then Exit Sub
    lb4.State = 1
    lb7.State = 1
    AddScore 25010
    LastSwitchHit = "barb4"
    CheckBARBTargets
    barb4on(CurrentPlayer) = 1
  End Sub

  Sub CheckBARBTargets
    If bMultiBallMode = True Then Exit Sub
    If lb1.State + lb2.State + lb3.State + lb4.State = 4 Then
      barbgate.Open = True
        If lro1.state = 0 then
          PuPlayer.playlistplayex pCallouts,"audiocallouts","barblockislit.wav",100,1
    chilloutthemusic
          PuPlayer.playlistplayex pBackglass,"videobarblit","",100,1
          pNote "WHERE'S BARB","LOCK IS LIT"
        End If
      lro1.State = 2
      lro3.State = 2
      lro4.State = 2
      lro5.State = 2
    End If
  End Sub

  Sub ResetBARBLights
    lb1.State = 0
    lb6.State = 0
    lb2.State = 0
    lb5.State = 0
    lb3.State = 0
    lb8.State = 0
    lb4.State = 0
    lb7.State = 0
    lro1.State = 0
    lro3.State = 0
    lro4.State = 0
    lro5.State = 0
    barb1on(CurrentPlayer) = 0
    barb2on(CurrentPlayer) = 0
    barb3on(CurrentPlayer) = 0
    barb4on(CurrentPlayer) = 0
  End Sub

'
' LOCKS
'****************

  Sub BallLockBarb_Hit
    RandomSoundHole
    Dim waittime
    waittime = 100
    If LookForBarb = True Then
    'BarbSuper
    Else
    BallsInLock(CurrentPlayer) = BallsInLock(CurrentPlayer) + 1
      Select Case BallsInLock(CurrentPlayer)
        Case 1
          ResetBARBLights
          DOF 405, DOFPulse   'DOF MX - Ball Locked
          waittime = 1000
          PuPlayer.playlistplayex pCallouts,"audiocallouts","ball1locked.wav",100,1
    chilloutthemusic
          PuPlayer.playlistplayex pBackglass,"videobarblock","",100,1
          pNote "WHERE'S BARB","BALL 1 LOCKED"
          vpmtimer.addtimer waittime, "BallLockBarbExit '"
          PuPlayer.LabelSet pBackglass,"barbh","" & BallsInLock(CurrentPlayer) ,1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
        Case 2
          barbskip = 1
          ResetBARBLights
          ruleshelperoff
          GiOff
          PuPlayer.LabelSet pBackglass,"notetitle","Where's Barb Multiball",1,""
          PuPlayer.LabelSet pBackglass,"notecopy","Hit the Yellow Shots - 1 Randomly will light Super Jackpot \r Hit Castle Byers when Super Jackpot List to Find Barb",1,""
          If inmode = 1 Then
          waittime = 1000
          Else
          waittime = 16000
          End If
          PuPlayer.playpause 4
          PuPlayer.playlistplayex pBackglass,"videobarb","barbmdstart.mov",100,1
          vpmtimer.addtimer waittime, "StartBarb'"
          PuPlayer.LabelSet pBackglass,"barbh","" & BallsInLock(CurrentPlayer) ,1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
      End Select
    End If
  End Sub

  Sub BallLockBarbExit()
    BallLockBarb.Kick 80, 9
      PlaySoundAt SoundFXDOF("fx_kicker", 119, DOFPulse, DOFContactors), BallLockBarb
      DOF 115, DOFPulse
    barbgate.Open = False
  End Sub

'
' MULTIBALL
'****************

  Sub StartBarb() 'Multiball
  Dim i
    If barbskip = 0 then exit Sub
    BallLockBarbExit
    DOF 338, DOFPulse   'DOF MX - BARB Flash
    pNote "WHERE'S BARB","MULTIBALL"
    PuPlayer.playlistplayex pBackglass,"videobarb","barbmbmid.mov",100,3
    PuPlayer.SetLoop 2,1
    PuPlayer.playpause 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","barb.mp3",100,1
    PuPlayer.SetLoop 7,1
    flashflash.Enabled = True
    PuPlayer.LabelSet pBackglass,"barbh","0",1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
    BallsInLock(CurrentPlayer) = 0
    bMultiBallMode = True
    barbMultiball = True
    AddMultiball 2
    EnableBallSaver 15
    GiOff
    GiYellow
    AssignBarbSuper


  dim k, newLoopDelay

    PrepSpellWord 40
    SetLoopDelay "find barb"

    SpellWord "find barb"

    for k = 1 to 7

      NewLoopDelay = LoopDelay*k
    debug.print "loop delay :" &k &":" &newloopdelay

      xmasQueue.Add "SpellWord-"&k,"SpellWord ""find barb"" ",80,NewLoopDelay,0,0,0,False
    Next

    xmasQueue.Add "SpellEnd","giOn:bInLightQuote = False",75,LoopDelay*8+(delayInc*myFSpeed*2),0,0,0,False


'   StartDemoLightSeq
'   GeneralPupQueue.Add "StopDemoLightSeq","StopDemoLightSeq",80,50000,0,0,0,False
'   LightSeqaxmas.UpdateInterval = 150
'   LightSeqaxmas.Play SeqRandom, 10, , 50000
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
    SetLightColor llo1, yellow, -1
    SetLightColor llo9, yellow, -1
    SetLightColor lro4, yellow, -1
    SetLightColor lro5, yellow, -1
    SetLightColor leftrampred, yellow, -1
    SetLightColor leftrampred2, yellow, -1
    SetLightColor leftrampred1, yellow, -1
    SetLightColor leftrampred3, yellow, -1
    SetLightColor rightrampred, yellow, -1
    SetLightColor rightrampred1, yellow, -1
    llo1.State = 2
    llo9.State = 2
    lro4.State = 2
    lro5.State = 2
    leftrampred.state = 2
    leftrampred2.state = 2
    leftrampred1.state = 2
    leftrampred3.state = 2
    rightrampred.state = 2
    rightrampred1.state = 2
    startamultiball
    barbskip = 0
  End Sub

  Sub AssignBarbSuper
    BSuper(CurrentPlayer) = RndNum(1,5)
    '1 - leftorbitdone 2 - leftrampdone 3 - centerrampdone 4 - rightrampdone 5 - rightorbitdone
  End Sub

  Sub AwardBarb
    barbjacks(CurrentPlayer) = barbjacks(currentplayer) + 1
    PuPlayer.LabelSet pBackglass,"barbj","" & barbjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"
    AddScore 1000000
    pNote "JACKPOT","1,000,000"
    PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    DOF 404, DOFPulse   'DOF MX - Jackpot
    chilloutthemusic
  End Sub

  Sub barbfound
    DOF 431, DOFPulse  ' DOF MX - Barb Get SJ
    PuPlayer.playlistplayex pCallouts,"audiocallouts","getthesuperjackpot.wav",100,1
    chilloutthemusic
    AddScore 2000000
    pNote "SUPER JACKPOT IS LIT","2,000,000"
    llo1.State = 0
    llo9.State = 0
    rightrampred.State = 0
    rightrampred1.State = 0
    leftrampred.State = 0
    leftrampred2.State = 0
    leftrampred1.State = 0
    leftrampred3.State = 0
    lro4.State = 0
    lro5.State = 0
    SetLightColor lro2, yellow, -1
    SetLightColor lro6, yellow, -1
    lro2.state = 2
    lro6.state = 2
    barbjacks(CurrentPlayer) = barbjacks(currentplayer) + 1
    PuPlayer.LabelSet pBackglass,"barbj","" & barbjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"

  End Sub

  Sub barbrestart
    If barbMultiball = True Then
    PuPlayer.playlistplayex pBackglass,"videobarb","barbmbmid.mov",100,1
    PuPlayer.SetLoop 2,1
    End If
  End Sub

  Sub awardbarbsuper
    PuPlayer.playlistplayex pCallouts,"audiocallouts","youfoundbarb.wav",100,1
    chilloutthemusic
    PuPlayer.playlistplayex pBackglass,"videobarb","barbendshort.mov",100,1
    AddScore 8000000
    pNote "BARB FOUND","8,000,000"
    DOF 425, DOFPulse  ' DOF MX - BARB Mode Complete
    Dim waittime
    waittime = 7500
    vpmtimer.addtimer waittime, "barbrestart'"
    lm1.state = 1
    lm9.state = 1
    barbcompleted(CurrentPlayer) = 1
    BSuper(CurrentPlayer) = 0
    llo1.State = 2
    llo9.State = 2
    lro4.State = 2
    lro5.State = 2
    leftrampred.state = 2
    leftrampred2.state = 2
    leftrampred1.state = 2
    leftrampred3.state = 2
    rightrampred.state = 2
    rightrampred1.state = 2
    barbjacks(CurrentPlayer) = barbjacks(currentplayer) + 1
    PuPlayer.LabelSet pBackglass,"barbj","" & barbjacks(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 81.5, 'yalign': 0}"

  End Sub

  Sub EndBarb()
    ruleshelperon
    PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,3
    PuPlayer.SetLoop 2,0
    PuPlayer.playresume 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1
    PuPlayer.SetLoop 7,0

    xmasqueue.RemoveAll(True)

    LookForBarb = False
    barbMultiball = False
    BarbJackpots = False
    bMultiBallMode = True
    llo1.State = 0
    llo9.State = 0
    llo3.State = 0
    llo7.State = 0
    lro2.State = 0
    lro6.State = 0
    lro4.State = 0
    lro5.State = 0
    CheckMONSTER
    GiOff
    GiOn
    LightSeqFlasher.StopPlay
    LightSeqaxmas.StopPlay
    llo1.State = 0
    llo9.State = 0
    lro4.State = 0
    lro5.State = 0
    leftrampred.state = 0
    leftrampred2.state = 0
    leftrampred1.state = 0
    leftrampred3.state = 0
    rightrampred.state = 0
    rightrampred1.state = 0
    lro2.state = 0
    lro6.state = 0
    SetLightColor llo1, red, -1
    SetLightColor llo9, red, -1
    SetLightColor lro4, red, -1
    SetLightColor lro5, red, -1
    SetLightColor lro2, red, -1
    SetLightColor lro6, red, -1
    endamultiball
  End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  BAD MEN
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'
' Light all BAD MEN lights to lock Balls
' Lock 3 balls to light start multiball
' During multiball hit ramps to escape - 6 shots
'
' TARGETS
'****************

  Sub run1_Hit
    FlashLevel5 = 1 : Flasherflash5_Timer
    LightEffect 9
    TruckShake()
    PlaySoundAt SoundFXDOF("fx_target", 134, DOFPulse, DOFTargets), ActiveBall
    PlaySound "runtarget"
    If Tilted Then Exit Sub
    If lr3.state = 0 Then
      lr3.State = 1
      lr12.State = 1
      run3on(CurrentPlayer) = 1
    Else
      lr4.state = 1
      lr11.state = 1
      run4on(CurrentPlayer) = 1
    End If
    AddScore 25010
    ' Do some sound or light effect
    DOF 330, DOFPulse   'DOF MX - Bad Men Target Left Splash
    ' do some check
    CheckRunTargets
  End Sub

  Sub run2_Hit
    FlashLevel5 = 1 : Flasherflash5_Timer
    LightEffect 9
    TruckShake()
    PlaySoundAt SoundFXDOF("fx_target", 134, DOFPulse, DOFTargets), ActiveBall
    PlaySound "runtarget"
    If Tilted Then Exit Sub

    If lr2.state = 0 Then
      lr2.State = 1
      lr9.State = 1
      run2on(CurrentPlayer) = 1
    Else
      lr5.state = 1
      lr10.state = 1
      run5on(CurrentPlayer) = 1
    End If
    AddScore 25010
    ' Do some sound or light effect
    DOF 330, DOFPulse   'DOF MX - Bad Men Target Left Splash
    ' do some check
    CheckRunTargets
  End Sub

  Sub run3_Hit
    FlashLevel5 = 1 : Flasherflash5_Timer
    If bSkillShotReady Then
      AwardSkillshot
      Else
      DOF 330, DOFPulse   'DOF MX - Bad Men Target Left Splash
    End If
    LightEffect 9
    TruckShake()
    PlaySoundAt SoundFXDOF("fx_target", 134, DOFPulse, DOFTargets), ActiveBall
    PlaySound "runtarget"
    If Tilted Then Exit Sub

    If lr1.state = 0 then
      lr1.State = 1
      lr8.State = 1
      run1on(CurrentPlayer) = 1
    End If

    If lr1.state = 1 Then
      lr6.state = 1
      lr7.state = 1
      run6on(CurrentPlayer) = 1
    End If

    AddScore 25010
    ' Do some sound or light effect

    LastSwitchHit = "run3"
    CheckRunTargets
  End Sub


  Sub CheckRunTargets
    If partymultiball = False Then
      If lucas.state = 0 Then
        lucas.state = 1
        lucason(CurrentPlayer) = 1
        partycollected(CurrentPlayer) = partycollected(CurrentPlayer) + 1
        PuPlayer.LabelSet pBackglass,"partyh","" & partycollected(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 88, 'yalign': 0}"
        checkparty
      End If
    End If
    If lucas.state = 2 Then
      lucas.state = 0
      awardparty
    End If

  If bMultiBallMode = True Then Exit Sub

  If run1lock(CurrentPlayer) = 0 or run1lock(CurrentPlayer) = 2 Then
    If lr1.state = 1 Then
      If lr1.State + lr2.State + lr3.State = 3 Then
        diverter.RotZ = -33
        diverteropen.collidable = True
        diverter.collidable = False
        If llo1.state = 0 then
          PuPlayer.playlistplayex pCallouts,"audiocallouts","badmenlockislit.wav",100,1
    chilloutthemusic
          pNote "BAD MEN","LOCK IS LIT"
          PuPlayer.playlistplayex pBackglass,"videobadmenlit","",100,1
        End If
        PlaySoundAt "fx_diverter", Diverter
        ''''DMD "lock-3.wmv", "", "", 4000
        ''''DMD "black.png", "Run Lock","is Lit",  500
        llo1.State = 2
        llo9.State = 2
        llo2.State = 2
        llo8.State = 2
      End If
    End If
  Else
    If lr4.State + lr5.State + lr6.State = 3 Then
      diverter.RotZ = -33
      diverteropen.collidable = True
      diverter.collidable = False
        If llo1.state = 0 then
          PuPlayer.playlistplayex pCallouts,"audiocallouts","badmenlockislit.wav",100,1
    chilloutthemusic
          pNote "BAD MEN","LOCK IS LIT"
          PuPlayer.playlistplayex pBackglass,"videobadmenlit","",100,1
        End If
        PlaySoundAt "fx_diverter", diverter
      llo1.State = 2
      llo9.State = 2
      llo2.State = 2
      llo8.State = 2
    End If
  End If
  End Sub

'
'Hawkins Electric Truck Animation / shake
'***************************

  Dim TruckPos

  Sub TruckShake()
    TruckPos = 3
    DOF 128, DOFPulse  'Shaker Motor
    TruckShakeTimer.Enabled = True
  End Sub

  Sub TruckShakeTimer_Timer()
    Hawkins.RotY = TruckPos
    If TruckPos <= 0.1 AND TruckPos >= -0.1 Then Me.Enabled = False:Exit Sub
    If TruckPos < 0 Then
      TruckPos = ABS(TruckPos)- 0.1
    Else
      TruckPos = - TruckPos + 0.1
    End If
  End Sub


'
' LOCKS
'****************

  Sub BallLockRun_Hit
    RandomSoundHole
    Dim waittime
    waittime = 100
    If RunAway = True Then
    RunSuper
    Else
    BallsInRunLock(CurrentPlayer) = BallsInRunLock(CurrentPlayer) + 1
      Select Case BallsInRunLock(CurrentPlayer)
        Case 1
          run1lock(CurrentPlayer) = 1
          PuPlayer.playlistplayex pCallouts,"audiocallouts","ball1locked.wav",100,1
    chilloutthemusic
          pNote "BAD MEN","BALL 1 LOCKED"
          PuPlayer.playlistplayex pBackglass,"videobadmenlock","",100,1
          DOF 405, DOFPulse   'DOF MX - Ball Locked
          PuPlayer.LabelSet pBackglass,"badh","" & BallsInRunLock(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
          waittime = 1000
          vpmtimer.addtimer waittime, "BallLockRunExit'"
        Case 3
          ResetRunLights
          run1lock(CurrentPlayer) = 2
          PuPlayer.playlistplayex pCallouts,"audiocallouts","ball2locked.wav",100,1
    chilloutthemusic
          pNote "BAD MEN","BALL 2 LOCKED"
          PuPlayer.playlistplayex pBackglass,"videobadmenlock","",100,1
          waittime = 1000
          DOF 405, DOFPulse   'DOF MX - Ball Locked
          PuPlayer.LabelSet pBackglass,"badh","" & BallsInRunLock(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
          vpmtimer.addtimer waittime, "BallLockRunExit '"
        Case 2
          badmenskip = 1
          PuPlayer.playpause 4
          ResetRunLights
          run1lock(CurrentPlayer) = 3
          ruleshelperoff
          GiOff
          PuPlayer.LabelSet pBackglass,"notetitle","Bad Men Multiball",1,""
          PuPlayer.LabelSet pBackglass,"notecopy","Hit the outside ramps 5 times to light Super Jackpot \r Hit the inside ramp to get the Super Jackpot and escape!",1,""
          PuPlayer.playlistplayex pBackglass,"videobadmenmb","badmenstart.mov",100,1
          If inmode = 1 Then
            waittime = 1000
          Else
            waittime = 11000
          End If
          PuPlayer.LabelSet pBackglass,"badh","" & BallsInRunLock(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 7.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
          vpmtimer.addtimer waittime, "StartRun'"
      End Select
    End If
    'vpmtimer.addtimer waittime, "BallLockRunExit '"
  End Sub

  Sub ResetRunLights
  lr1.State = 0
  lr8.State = 0
  lr2.State = 0
  lr9.State = 0
  lr3.State = 0
  lr12.State = 0
  lr4.State = 0
  lr11.State = 0
  lr5.State = 0
  lr10.State = 0
  lr6.State = 0
  lr7.State = 0
  run1on(CurrentPlayer) = 0
  run2on(CurrentPlayer) = 0
  run3on(CurrentPlayer) = 0
  run4on(CurrentPlayer) = 0
  run5on(CurrentPlayer) = 0
  run6on(CurrentPlayer) = 0
  End Sub


  Sub BallLockRunExit
    FlashLevel1 = 1 : Flasherflash1_Timer
    BallLockRun.Kick 90, 7
    PlaySoundAt SoundFXDOF("fx_kicker", 117, DOFPulse, DOFContactors), BallLockRun
    DOF 115, DOFPulse
    llo1.State = 0
    llo9.State = 0
    llo2.State = 0
    llo8.State = 0
    CheckRunTargets
    Dim waittime
    waittime = 1000
    vpmtimer.addtimer waittime, "delayerdiverterclose'"
  End Sub

  Sub delayerdiverterclose
    diverter.RotZ = 0
    diverteropen.collidable = False
    diverter.collidable = True
    PlaySoundAt "fx_diverter", diverter
    CheckRunTargets
  End Sub

'
' MULTIBALL
'****************

  Sub StartRun() 'Multiball
    If badmenskip = 0 then exit Sub
    BallLockRunExit
    DOF 333, DOFPulse   'DOF MX - Flash Bad Men
    pNote "BAD MEN","MULTIBALL"
    PuPlayer.playlistplayex pBackglass,"videobadmenmb","badmenmid.mov",100,3
    PuPlayer.SetLoop 2,1
    PuPlayer.playpause 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","badmen.mp3",100,1
    PuPlayer.SetLoop 7,1
    flashflash.Enabled = True
    BallsInRunLock(CurrentPlayer) = 0
    bMultiBallMode = True
    RunMultiball = True
    AddMultiball 2
    EnableBallSaver 10
    GiOff
    GiBlue
    run1lock(CurrentPlayer) = 0
'   LightSeqaxmas.UpdateInterval = 150
'   LightSeqaxmas.Play SeqRandom, 10, , 50000

  dim k, newLoopDelay

    PrepSpellWord 40
    SetLoopDelay "bad men"

    SpellWord "bad men"

    for k = 1 to 7

      NewLoopDelay = LoopDelay*k
    debug.print "loop delay :" &k &":" &newloopdelay

      xmasQueue.Add "SpellWord-"&k,"SpellWord ""bad men"" ",80,NewLoopDelay,0,0,0,False
    Next

    xmasQueue.Add "SpellEnd","giOn:bInLightQuote = False",75,LoopDelay*8+(delayInc*myFSpeed*2),0,0,0,False


'   StartDemoLightSeq
'   GeneralPupQueue.Add "StopDemoLightSeq","StopDemoLightSeq",80,50000,0,0,0,False
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
    ' turn up the lights and yell baby
    SetLightColor llo3, darkblue, -1
    SetLightColor llo7, darkblue, -1
    SetLightColor llr1, darkblue, -1
    SetLightColor llr2, darkblue, -1
    SetLightColor lc1, darkblue, -1
    SetLightColor lc4, darkblue, -1
    llo3.State = 2
    llo7.State = 2
    llr1.State = 2
    llr2.State = 2
    startamultiball
    If RunHits(CurrentPlayer) = 5 Then
      TrytoRun
    End If
    badmenskip = 0
  End Sub

  Sub AwardRun
    If RunHits(CurrentPlayer) = 5 Then
      TrytoRun
    End If
    If RunAway = False Then
    RunHits(CurrentPlayer) = RunHits(CurrentPlayer) + 1
    End If
    Select Case RunHits(CurrentPlayer)
      Case 1
        AddScore 2000000
        pNote "JACKPOT","2,000,000"
        PuPlayer.LabelSet pBackglass,"badj","" & RunHits(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
        PlaySound "steve"
        DOF 404, DOFPulse   'DOF MX - Jackpot
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 2
        AddScore 2000000
        pNote "JACKPOT","2,000,000"
        PuPlayer.LabelSet pBackglass,"badj","" & RunHits(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
        PlaySound "steve"
        DOF 404, DOFPulse   'DOF MX - Jackpot
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 3
        AddScore 2000000
        pNote "JACKPOT","2,000,000"
        PuPlayer.LabelSet pBackglass,"badj","" & RunHits(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
        PlaySound "steve"
        DOF 404, DOFPulse   'DOF MX - Jackpot
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 4
        AddScore 2000000
        pNote "JACKPOT","2,000,000"
        PuPlayer.LabelSet pBackglass,"badj","" & RunHits(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
        PlaySound "steve"
        DOF 404, DOFPulse   'DOF MX - Jackpot
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
      Case 5
        AddScore 2000000
        pNote "JACKPOT","2,000,000"
        PuPlayer.LabelSet pBackglass,"badj","" & RunHits(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
        PlaySound "steve"
        DOF 430, DOFPulse  ' DOF MX - Bad Men Get SJ
        PuPlayer.playlistplayex pCallouts,"audiojackpot","",100,1
    chilloutthemusic
        TrytoRun
    End Select
  End Sub

  Sub TrytoRun
    RunAway = True
    lc1.State = 2
    lc4.State = 2
    llo3.State = 0
    llo7.State = 0
    llr1.State = 0
    llr2.State = 0
    PuPlayer.playlistplayex pCallouts,"audiocallouts","getthesuperjackpot.wav",100,1
    chilloutthemusic
  End Sub

  Sub runrestart
    If RunMultiball = True Then
    PuPlayer.playlistplayex pBackglass,"videobadmenmb","badmenmid.mov",100,1
    PuPlayer.SetLoop 2,1
    End If
  End Sub

  Sub RUNSuper
    RunAway = False
    AddScore 6000000
    pNote "BAD MEN ESCAPED","6,000,000"
    PuPlayer.playlistplayex pBackglass,"videobadmenmb","badmenend.mov",100,1
    Dim waittime
    waittime = 24000
    vpmtimer.addtimer waittime, "runrestart'"
    PuPlayer.playlistplayex pCallouts,"audiocallouts","badmenescaped.wav",100,1
    chilloutthemusic
    PlaySound "steve"
    DOF 424, DOFPulse  ' DOF MX - Bad Men Mode Complete
    RunHits(CurrentPlayer) = 0
    PuPlayer.LabelSet pBackglass,"badj","" & RunHits(CurrentPlayer),1,"{'mt':2,'color':16777215, 'size': 2, 'xpos': 12.7, 'xalign': 0, 'ypos': 74.6, 'yalign': 0}"
    lm2.State = 1
    lm8.State = 1
    lc1.State = 0
    lc4.State = 0
    llo3.State = 2
    llo7.State = 2
    llr1.State = 2
    llr2.State = 2
    badmensuper(currentplayer) = 1
  End Sub

  Sub EndRun()
    ruleshelperon
    PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,3
    PuPlayer.SetLoop 2,0
    PuPlayer.playresume 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1
    PuPlayer.SetLoop 7,0
    ResetRunLights
    RunAway = False
    RunMultiball = False
    bMultiBallMode = True
    diverter.RotZ = 0
    diverteropen.collidable = False
    diverter.collidable = True
    PlaySoundAt "fx_diverter", diverter
    lc1.State = 0
    lc4.State = 0
    llo3.State = 0
    llo7.State = 0
    llr1.State = 0
    llr2.State = 0
    CheckMONSTER
    GiOff
    GiOn
    FlashEffect 0
    Flashxmas 0
    SetLightColor llo3, red, -1
    SetLightColor llo7, red, -1
    SetLightColor llr1, red, -1
    SetLightColor llr2, red, -1
    SetLightColor lc1, red, -1
    SetLightColor lc4, red, -1
    LightSeqFlasher.StopPlay
    LightSeqaxmas.StopPlay
    bMultiBallMode = False
    endamultiball
  End Sub





'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  KILL THE DEMOGORGON
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'

  Sub CheckMONSTER
    If lm1.State + lm2.State + lm3.State + lm5.State = 4 Then
      Hit11
    End If
  End Sub

  Sub Hit11
    FlasherEL.opacity = 500
    MagnetR.MagnetON = True ' Magnet On
    KickerEL.enabled = True
    GiOff
  End Sub

  Dim demodefeated

  Sub KickerEL_hit
    finalflips = True
    ruleshelperoff
    Dim waittime
    If MonsterFinalBlow = True Then
      finalscene
      RightFlipper.RotateToStart
      LeftFlipper.RotateToStart
    Else
      dropwallskip = 1
      RightFlipper.RotateToStart
      LeftFlipper.RotateToStart
      Dropwall
      waittime = 9000
      pNote "KILL THE DEMOGORGON","HIT THE DEMOGORGON 20 TIMES IN 40 SECS"
      PuPlayer.playlistplayex pBackglass,"videowizard","wizardmode.mov",100,3
      vpmtimer.addtimer waittime, "kickit '"
      vpmtimer.addtimer waittime, "StartMonster '"
      MagnetR.MagnetON = False ' Magnet Off
      PuPlayer.LabelSet pBackglass,"notetitle","Kill the Demogorgon",1,""
      PuPlayer.LabelSet pBackglass,"notecopy","It comes down to this Hit the demogorgon 20 times in 40 seconds \r Become our champion and get 100 million, or die defeated",1,""
    End If
    GiOff
  End Sub


  Sub kickit
    finalflips = False
    KickerEL.Kick -1, 50
    DOF 115, DOFPulse
    KickerEL.enabled = False
    spinner.MotorOn = True
    PlaySoundAt SoundFXDOF("fx_kicker", 118, DOFPulse, DOFContactors), KickerEL
  End Sub


  Sub bell
    PlaySound "bell"
  End Sub

  Sub finalscene
    GiOff
    bBallSaverActive = False
    ruleshelperoff
    PuPlayer.playlistplayex pBackglass,"videowizard","wizardfinal.mov",100,3
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1
    PuPlayer.playlistplayex pMusic,"audioclear","clear.mp3",100, 1
'   LightSeqaxmas.UpdateInterval = 150
'   LightSeqaxmas.Play SeqRandom, 10, , 116000
    StartDemoLightSeq
    GeneralPupQueue.Add "StopDemoLightSeq","StopDemoLightSeq",80,116000,0,0,0,False
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 116000
    lm4.State = 1
    lm6.State = 1
    vpmtimer.addtimer 108000, "bell '"
    vpmtimer.addtimer 108000, "final1'"
    vpmtimer.addtimer 111000, "bell '"
    vpmtimer.addtimer 111000, "final2'"
    vpmtimer.addtimer 114000, "bell '"
    vpmtimer.addtimer 114000, "final3'"
    vpmtimer.addtimer 116000, "bell '"

    vpmtimer.addtimer 108000, "DemoSuper '"
    vpmtimer.addtimer 120000, "kickit2 '"
    vpmtimer.addtimer 119000, "final4'"
    vpmtimer.addtimer 117000, "EndDemo '"
    demodefeated = True
    MagnetR.MagnetON = False ' Magnet Off
  End Sub

  Sub final1
    DOF 427, DOFPulse  ' DOF MX - Monster Mode Complete
    pNote "DEMOGORGON DEFEATED","100,000,000"
  End Sub

  Sub final2
    DOF 427, DOFPulse  ' DOF MX - Monster Mode Complete
    pNote "WE WILL ALL","BE SAFE"
  End Sub

  Sub final3
    DOF 427, DOFPulse  ' DOF MX - Monster Mode Complete
    pNote "OR WILL WE","..."
  End Sub

  Sub final4
    pNote "BALL RELEASING","GET READY"
  End Sub


  Sub kickit2
    finalflips = False
    KickerEL.Kick -40, 50
    DOF 115, DOFPulse
    KickerEL.enabled = False
    PlaySoundAt SoundFXDOF("fx_kicker", 118, DOFPulse, DOFContactors), KickerEL
  End Sub

  Dim wizardpos
  wizardpos = 0
    wizardtimer.enabled = false
  Sub wizardtimer_Timer()
    wizardpos = wizardpos + 1
    PuPlayer.LabelSet pBackglass,"modetimer",40 - wizardpos,1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 80.7, 'xalign': 1, 'ypos': 75, 'yalign': 0}"
    Select Case wizardpos
      Case 1

      Case 40
        If MonsterFinalBlow = False Then
        dim waittime
        waittime = 6000
        vpmtimer.addtimer waittime, "EndDemo '"
        finalflips = True
        pNote "YOU FAILED","YOU HAVE DIED"
        End If
        PuPlayer.LabelSet pBackglass,"modetimer","0",1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 80.7, 'xalign': 1, 'ypos': 75, 'yalign': 0}"
        wizardtimer.enabled = false
    End Select
  End Sub


  Sub StartMonster() 'Multiball
    If openwall = False Then
      dropwallskip = 1
      Dropwall
    End If
    wizardtimer.enabled = true
    DemoMultiball = True
    PuPlayer.playpause 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","wizard.mp3",100,1
    PuPlayer.SetLoop 7,1
    bMultiBallMode = True
    DOF 419, DOFPulse    'DOF MX - Monster Flash
    DemoMultiball = True
    AddMultiball 2
    EnableBallSaver 27
    FlasherEL.opacity = 0
    DemoHits = 0
    FlashEffect 2
'   StartDemoLightSeq
'   GeneralPupQueue.Add "StopDemoLightSeq","StopDemoLightSeq",80,65000,0,0,0,False

  dim k, newLoopDelay

    PrepSpellWord 40
    SetLoopDelay "demogorgon"

    SpellWord "demogorgon"

    for k = 1 to 7

      NewLoopDelay = LoopDelay*k
    debug.print "loop delay :" &k &":" &newloopdelay

      xmasQueue.Add "SpellWord-"&k,"SpellWord ""demogorgon"" ",80,NewLoopDelay,0,0,0,False
    Next

    xmasQueue.Add "SpellEnd","giOn:bInLightQuote = False",75,LoopDelay*8+(delayInc*myFSpeed*2),0,0,0,False

'   LightSeqaxmas.UpdateInterval = 150
'   LightSeqaxmas.Play SeqRandom, 10, , 65000
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 65000
    ' turn up the lights and yell baby
      Dim a
    For each a in aLights
      a.State = 2
    Next
    StartRainbow aLights
    startamultiball
    GiOff
    GiRed
    udtarget.Collidable = True
  End Sub


  Sub KillMonster
    udtarget.Collidable = True
    If MonsterFinalBlow = False Then
    DemoHits = DemoHits + 1
    End If
    Select Case DemoHits
      Case 1
        AddScore 4000000
        PlaySound "demohit"
        DOF 221, DOFPulse  ' DOF - Shaker Increases Intensity level with each hit
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 1","4,000,000"
      Case 2
        AddScore 4000000
        PlaySound "demohit"
        DOF 222, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 2","4,000,000"
      Case 3
        AddScore 4000000
        PlaySound "demohit"
        DOF 223, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 3","4,000,000"
      Case 4
        AddScore 4000000
        PlaySound "demohit"
        DOF 224, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 4","4,000,000"
      Case 5
        AddScore 4000000
        PlaySound "demohit"
        DOF 225, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 5","4,000,000"
      Case 6
        AddScore 4000000
        PlaySound "demohit"
        DOF 226, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 6","4,000,000"
      Case 7
        AddScore 4000000
        PlaySound "demohit"
        DOF 227, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 7","4,000,000"
      Case 8
        AddScore 4000000
        PlaySound "demohit"
        DOF 228, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 8","4,000,000"
      Case 9
        AddScore 4000000
        PlaySound "demohit"
        DOF 229, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 9","4,000,000"
      Case 10
        AddScore 4000000
        PlaySound "demohit"
        DOF 230, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 10","4,000,000"
      Case 11
        AddScore 4000000
        PlaySound "demohit"
        DOF 231, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 11","4,000,000"
      Case 12
        AddScore 4000000
        PlaySound "demohit"
        DOF 232, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 12","4,000,000"
      Case 13
        AddScore 4000000
        PlaySound "demohit"
        DOF 233, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 13","4,000,000"
      Case 14
        AddScore 4000000
        PlaySound "demohit"
        DOF 234, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 14","4,000,000"
      Case 15
        AddScore 4000000
        PlaySound "demohit"
        DOF 235, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 15","4,000,000"
      Case 16
        AddScore 4000000
        PlaySound "demohit"
        DOF 236, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 16","4,000,000"
      Case 17
        AddScore 4000000
        PlaySound "demohit"
        DOF 237, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 17","4,000,000"
      Case 18
        AddScore 4000000
        PlaySound "demohit"
        DOF 238, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 18","4,000,000"
      Case 19
        AddScore 4000000
        PlaySound "demohit"
        DOF 239, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        pNote "DEMOGORGON HIT 19","4,000,000"
      Case 20
        AddScore 4000000
        PlaySound "demohit"
        DOF 203, DOFPulse  ' DOF - Fan
        DOF 240, DOFPulse  ' DOF - Shaker
        DOF 428, DOFPulse  ' DOF MX - Monster Hit
        TryForFinalBlow
        pNote "DEMOGORGON HIT 20","4,000,000"
    End Select
  End Sub

  Sub TryForFinalBlow
    MonsterFinalBlow = True
    TurnOffPlayfieldLights()
    GiRed
    FlasherEL.opacity = 500
    MagnetR.MagnetON = True ' Magnet On
    KickerEL.enabled = True
  End Sub

  Sub DemoSuper
    DOF 427, DOFPulse  ' DOF MX - Monster Mode Complete
    AddScore 100000000
    lm4.State = 1
    lm6.State = 1
    DemoHits = 0
    MonsterFinalBlow = False
  End Sub


  Sub EndDemo()
    If MonsterFinalBlow = True Then Exit Sub
    wizardtimer.enabled = false
    ruleshelperon
    bBallSaverActive = False
    PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,3
    PuPlayer.SetLoop 2,0
    PuPlayer.playresume 4
    PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1
    PuPlayer.SetLoop 7,0
    spinner.MotorOn = False
    Dim lamp
    MonsterFinalBlow = False
    DemoMultiball = False
    lm1.State = 0
    lm9.State = 0
    lm2.State = 0
    lm8.State = 0
    lm3.State = 0
    lm7.State = 0
    barbcompleted(CurrentPlayer) = 0
    badmensuper(CurrentPlayer) = 0
    partydone(CurrentPlayer) = 0
    willcompleted(CurrentPlayer) = 0
    demoisdead(CurrentPlayer) = 1
    WillHits(CurrentPlayer) = 0
    StopRainbow
    ResetAllLightsColor
    FlasherEL.opacity = 0
    GiOff
    GiOn
    wizardpos = 0
    FlashEffect 0

    xmasqueue.RemoveAll(True)

'   LightSeqFlasher.StopPlay
    LightSeqaxmas.StopPlay
    If demodefeated = True Then
    lm4.State = 1
    lm6.State = 1
    End If
    finalflips = false
    endamultiball
    Raisewall
    If rockmusic = 1 Then
    PuPlayer.playlistplayex pMusic,"audiobgrock","",soundtrackvol,1
    PuPlayer.SetLoop 4,1
    Else
    PuPlayer.playlistplayex pMusic,"audiobg","",soundtrackvol,1
    PuPlayer.SetLoop 4,1
    End if
    bMultiBallMode = False
    MagnetR.MagnetON = false ' Magnet On
    KickerEL.enabled = false
  End Sub


  '**********************************************
  'DOF Shaker for Videos  'TerryRed

  'These subs and Timers are used to trigger
  'DOF Shaker motor pulses to match video events.
  '**********************************************


  Sub DOF_Shaker_R_Timer()  'Shaker Motor Burst when "R" appears in video
    if dropwallskip = 0 then DOF_Shaker_R.enabled = false:exit sub
      DOF 250, DOFPulse   'DOF Shaker   "R"
      DOF_Shaker_R.enabled = false
  End Sub

  Sub DOF_Shaker_U_Timer()  'Shaker Motor Burst when "U" appears in video
    if dropwallskip = 0 then DOF_Shaker_U.enabled = false:exit sub
      DOF 251, DOFPulse   'DOF Shaker   "U"
      DOF_Shaker_U.enabled = false
  End Sub

  Sub DOF_Shaker_N_Timer()  'Shaker Motor Burst when "N" appears in video
    if dropwallskip = 0 then DOF_Shaker_N.enabled = false:exit sub
      DOF 252, DOFPulse   'DOF Shaker   "N"
      DOF_Shaker_N.enabled = false
  End Sub

  Sub DOF_Shaker_DemoWallBurst_Timer()  'Shaker Motor Burst when "Demogorgon Wall Burst" appears in video
    if dropwallskip = 0 then DOF_Shaker_DemoWallBurst.enabled = false:exit sub
      DOF 253, DOFPulse   'DOF Shaker   "Demogorgon Wall Burst"
      DOF_Shaker_DemoWallBurst.enabled = false
  End Sub

  Sub DOF_Shaker_WillHiding_Timer()  'Shaker Motor Burst when "Will is Hiding" appears in video
    if willskip = 0 then DOF_Shaker_WillHiding.enabled = false:exit sub
      DOF 254, DOFPulse   'DOF Shaker   "Will is Hiding"
      DOF_Shaker_WillHiding.enabled = false
  End Sub



'**********************************
'   ZMAT: General Math Functions
'**********************************
' These get used throughout the script.

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
  dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
  dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    End If
  ElseIf dx = 0 Then
    If dy = 0 Then
      Atn2 = 0
    Else
      Atn2 = Sgn(dy) * pi / 2
    End If
  End If
End Function

Function ArcCos(x)
  If x = 1 Then
    ArcCos = 0/180*PI
  ElseIf x = -1 Then
    ArcCos = 180/180*PI
  Else
    ArcCos = Atn(-x/Sqr(-x * x + 1)) + 2 * Atn(1)
  End If
End Function

Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

Function min(a,b)
  If a > b Then
    min = b
  Else
    min = a
  End If
End Function

' Used for drop targets
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy) 'Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
  Dim AB, BC, CD, DA
  AB = (bx * py) - (by * px) - (ax * py) + (ay * px) + (ax * by) - (ay * bx)
  BC = (cx * py) - (cy * px) - (bx * py) + (by * px) + (bx * cy) - (by * cx)
  CD = (dx * py) - (dy * px) - (cx * py) + (cy * px) + (cx * dy) - (cy * dx)
  DA = (ax * py) - (ay * px) - (dx * py) + (dy * px) + (dx * ay) - (dy * ax)

  If (AB <= 0 And BC <= 0 And CD <= 0 And DA <= 0) Or (AB >= 0 And BC >= 0 And CD >= 0 And DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
  Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
  Dim rotxy
  rotxy = RotPoint(ax,ay,angle)
  rax = rotxy(0) + px
  ray = rotxy(1) + py
  rotxy = RotPoint(bx,by,angle)
  rbx = rotxy(0) + px
  rby = rotxy(1) + py
  rotxy = RotPoint(cx,cy,angle)
  rcx = rotxy(0) + px
  rcy = rotxy(1) + py
  rotxy = RotPoint(dx,dy,angle)
  rdx = rotxy(0) + px
  rdy = rotxy(1) + py

  InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
  Dim rx, ry
  rx = x * dCos(angle) - y * dSin(angle)
  ry = x * dSin(angle) + y * dCos(angle)
  RotPoint = Array(rx,ry)
End Function

'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. and, special scripting
'
' TriggerLF and RF should now be 27 vp units from the flippers. In addition, 3 degrees should be added to the end angle
' when creating these triggers.
'
' RF.ReProcessBalls Activeball and LF.ReProcessBalls Activeball must be added the flipper_collide subs.
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.4            | 0.4                   | 0.375                  | 0.375              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity


'*******************************************
' Early 90's and after

Sub InitPolarity()
  Dim x, a
  a = Array(LF, RF)
  For Each x In a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 5.5
    x.AddPt "Polarity", 2, 0.16, - 5.5
    x.AddPt "Polarity", 3, 0.20, - 0.75
    x.AddPt "Polarity", 4, 0.25, - 1.25
    x.AddPt "Polarity", 5, 0.3, - 1.75
    x.AddPt "Polarity", 6, 0.4, - 3.5
    x.AddPt "Polarity", 7, 0.5, - 5.25
    x.AddPt "Polarity", 8, 0.7, - 4.0
    x.AddPt "Polarity", 9, 0.75, - 3.5
    x.AddPt "Polarity", 10, 0.8, - 3.0
    x.AddPt "Polarity", 11, 0.85, - 2.5
    x.AddPt "Polarity", 12, 0.9, - 2.0
    x.AddPt "Polarity", 13, 0.95, - 1.5
    x.AddPt "Polarity", 14, 1, - 1.0
    x.AddPt "Polarity", 15, 1.05, -0.5
    x.AddPt "Polarity", 16, 1.1, 0
    x.AddPt "Polarity", 17, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.23, 0.85
    x.AddPt "Velocity", 2, 0.27, 1
    x.AddPt "Velocity", 3, 0.3, 1
    x.AddPt "Velocity", 4, 0.35, 1
    x.AddPt "Velocity", 5, 0.6, 1 '0.982
    x.AddPt "Velocity", 6, 0.62, 1.0
    x.AddPt "Velocity", 7, 0.702, 0.968
    x.AddPt "Velocity", 8, 0.95,  0.968
    x.AddPt "Velocity", 9, 1.03,  0.945
    x.AddPt "Velocity", 10, 1.5,  0.945

  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
  LF.SetObjects "LF", LeftFlipper, TriggerLF
  RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub

'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before\
' modified 2024 by rothbauerw
' Added Reprocessballs for flipper collisions (LF.Reprocessballs Activeball and RF.Reprocessballs Activeball must be added to the flipper collide subs
' Improved handling to remove correction for backhand shots when the flipper is raised

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef, FlipStartAngle
  Private Balls(20), balldata(20)
  Private Name

  Dim PolarityIn, PolarityOut
  Dim VelocityIn, VelocityOut
  Dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    ReDim PolarityIn(0)
    ReDim PolarityOut(0)
    ReDim VelocityIn(0)
    ReDim VelocityOut(0)
    ReDim YcoefIn(0)
    ReDim YcoefOut(0)
    Enabled = True
    TimeDelay = 50
    LR = 1
    Dim x
    For x = 0 To UBound(balls)
      balls(x) = Empty
      Set Balldata(x) = new SpoofBall
    Next
  End Sub

  Public Sub SetObjects(aName, aFlipper, aTrigger)

    If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: .SetObjects error: first argument must be a String (And name of Object). Found:" & TypeName(aName) End If
    If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: .SetObjects error: Second argument must be a flipper. Found:" & TypeName(aFlipper) End If
    If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & TypeName(aTrigger) End If
    If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
    Name = aName
    Set Flipper = aFlipper
    FlipperStart = aFlipper.x
    FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
    FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y

    Dim str
    str = "Sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
    ExecuteGlobal(str)
    str = "Sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
    ExecuteGlobal(str)

  End Sub

  ' Legacy: just no op
  Public Property Let EndPoint(aInput)

  End Property

  Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      Case "Polarity"
        ShuffleArrays PolarityIn, PolarityOut, 1
        PolarityIn(aIDX) = aX
        PolarityOut(aIDX) = aY
        ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity"
        ShuffleArrays VelocityIn, VelocityOut, 1
        VelocityIn(aIDX) = aX
        VelocityOut(aIDX) = aY
        ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef"
        ShuffleArrays YcoefIn, YcoefOut, 1
        YcoefIn(aIDX) = aX
        YcoefOut(aIDX) = aY
        ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
  End Sub

  Public Sub AddBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If IsEmpty(balls(x)) Then
        Set balls(x) = aBall
        Exit Sub
      End If
    Next
  End Sub

  Private Sub RemoveBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If TypeName(balls(x) ) = "IBall" Then
        If aBall.ID = Balls(x).ID Then
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
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x)) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x)) Then
        balldata(x).Data = balls(x)
      End If
    Next
    FlipStartAngle = Flipper.currentangle
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub

  Public Sub ReProcessBalls(aBall) 'save data of balls in flipper range
    If FlipperOn() Then
      Dim x
      For x = 0 To UBound(balls)
        If Not IsEmpty(balls(x)) Then
          if balls(x).ID = aBall.ID Then
            If isempty(balldata(x).ID) Then
              balldata(x).Data = balls(x)
            End If
          End If
        End If
      Next
    End If
  End Sub

  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef, BalltoFlip, BalltoBase, NoCorrection, checkHit
      Ycoef = 1

      'y safety Exit
      If aBall.VelY > -8 Then 'ball going down
        RemoveBall aBall
        Exit Sub
      End If

      'Find balldata. BallPos = % on Flipper
      For x = 0 To UBound(Balls)
        If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          BalltoFlip = DistanceFromFlipperAngle(BallData(x).x, BallData(x).y, Flipper, FlipStartAngle)
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
        NoCorrection = 1
      Else
        checkHit = 50 + (20 * BallPos)

        If BalltoFlip > checkHit or (PartialFlipCoef < 0.5 and BallPos > 0.22) Then
          NoCorrection = 1
        Else
          NoCorrection = 0
        End If
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        'If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled and NoCorrection = 0 Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef*VelCoef)
      End If
      If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  Dim x, aCount
  aCount = 0
  ReDim a(UBound(aArray) )
  For x = 0 To UBound(aArray)   'Shuffle objects in a temp array
    If Not IsEmpty(aArray(x) ) Then
      If IsObject(aArray(x)) Then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  If offset < 0 Then offset = 0
  ReDim aArray(aCount-1+offset)   'Resize original array
  For x = 0 To aCount-1       'set objects back into original array
    If IsObject(a(x)) Then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
  BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)    'Set up line via two points, no clamping. Input X, output Y
  Dim x, y, b, m
  x = input
  m = (Y2 - Y1) / (X2 - X1)
  b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x
      y = .y
      z = .z
      velx = .velx
      vely = .vely
      velz = .velz
      id = .ID
      mass = .mass
      radius = .radius
    End With
  End Property
  Public Sub Reset()
    x = Empty
    y = Empty
    z = Empty
    velx = Empty
    vely = Empty
    velz = Empty
    id = Empty
    mass = Empty
    radius = Empty
  End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  Dim y 'Y output
  Dim L 'Line
  'find active line
  Dim ii
  For ii = 1 To UBound(xKeyFrame)
    If xInput <= xKeyFrame(ii) Then
      L = ii
      Exit For
    End If
  Next
  If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)    'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )     'Clamp lower
  If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )    'Clamp upper

  LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************
' To add the flipper tricks you must
'  - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'  - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines
'  - Include FlipperActivate and FlipperDeactivate in the Flipper solenoid subs

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
    FlipperTricks Flipper2, LFPress1, LFCount1, LFEndAngle1, LFState1
  FlipperTricks Flipper1, RFPress1, RFCount1, RFEndAngle1, RFState1
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
     Dim gBOT
     gBOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
          gBOT(b).velx = gBOT(b).velx / 1.3
          gBOT(b).vely = gBOT(b).vely - 0.5
        End If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
  End If
End Sub


Dim FCCDamping: FCCDamping = 0.4

Sub FlipperCradleCollision(ball1, ball2, velocity)
  if velocity < 0.7 then exit sub   'filter out gentle collisions
    Dim DoDamping, coef
    DoDamping = false
    'Check left flipper
    If LeftFlipper.currentangle = LFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, LeftFlipper) OR FlipperTrigger(ball2.x, ball2.y, LeftFlipper) Then DoDamping = true
    End If
    'Check right flipper
    If RightFlipper.currentangle = RFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, RightFlipper) OR FlipperTrigger(ball2.x, ball2.y, RightFlipper) Then DoDamping = true
    End If
    If DoDamping Then
    coef = FCCDamping
        ball1.velx = ball1.velx * coef: ball1.vely = ball1.vely * coef: ball1.velz = ball1.velz * coef
        ball2.velx = ball2.velx * coef: ball2.vely = ball2.vely * coef: ball2.velz = ball2.velz * coef
    End If
End Sub





'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
  DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function DistanceFromFlipperAngle(ballx, bally, Flipper, Angle)
  DistanceFromFlipperAngle = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Angle + 90)) + Flipper.x, Sin(Radians(angle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function

'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle
Dim LFPress1, LFCount1, LFEndAngle1, LFState1
Dim RFPress1, RFCount1, RFEndAngle1, RFState1

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)
LFState = 1
LFState1 = 1
RFState = 1
RFState1 = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0
    SOSRampup = 2.5
  Case 1
    SOSRampup = 6
  Case 2
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle
LFEndAngle = Leftflipper.endangle
LFEndAngle1 = Flipper2.endangle
RFEndAngle = RightFlipper.endangle
RFEndAngle1 = Flipper1.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
  FlipperPress = 0
  Flipper.eostorqueangle = EOSA
  Flipper.eostorque = EOST * EOSReturn / FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b, gBOT
        gBOT = GetBalls

    For b = 0 To UBound(gBOT)
      If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3 * Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
    If FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If
  End If
End Sub

Const LiveDistanceMin = 5  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)
Const BaseDampen = 0.55

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
    Dim Dir, LiveDist
    Dir = Flipper.startangle / Abs(Flipper.startangle)    '-1 for Right Flipper
    Dim LiveCatchBounce   'If live catch is not perfect, it won't freeze ball totally
    Dim CatchTime
    CatchTime = GameTime - FCount
    LiveDist = Abs(Flipper.x - ball.x)

    If CatchTime <= LiveCatch And parm > 3 And LiveDist > LiveDistanceMin And LiveDist < LiveDistanceMax Then
        If CatchTime <= LiveCatch * 0.5 Then   'Perfect catch only when catch time happens in the beginning of the window
            LiveCatchBounce = 0
        Else
            LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)  'Partial catch when catch happens a bit late
        End If

        If LiveCatchBounce = 0 And ball.velx * Dir > 0 And LiveDist > 30 Then ball.velx = 0

        If ball.velx * Dir > 0 And LiveDist < 30 Then
            ball.velx = BaseDampen * ball.velx
            ball.vely = BaseDampen * ball.vely
            ball.angmomx = BaseDampen * ball.angmomx
            ball.angmomy = BaseDampen * ball.angmomy
            ball.angmomz = BaseDampen * ball.angmomz
        Elseif LiveDist > 30 Then
            ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
            ball.angmomx = 0
            ball.angmomy = 0
            ball.angmomz = 0
        End If
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
    End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************





'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen ActiveBall
  TargetBouncer ActiveBall, 0.7
End Sub

Dim RubbersD        'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False    'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1    'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64    'there's clamping so interpolate up to 56 at least

Dim SleevesD  'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False    'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn   'tbpOut.text
  Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
  End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
  End Sub

  Public Sub Dampen(aBall)
    If threshold Then
      If BallSpeed(aBall) < threshold Then Exit Sub
    End If
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
    "actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
    If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)

    aBall.velx = aBall.velx * coef
    aBall.vely = aBall.vely * coef
    aBall.velz = aBall.velz * coef
    If debugOn Then TBPout.text = str
  End Sub

  Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
      aBall.velx = aBall.velx * coef
      aBall.vely = aBall.vely * coef
      aBall.velz = aBall.velz * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    Dim x
    For x = 0 To UBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
    Next
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub
End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
  Public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize
    ReDim ballvel(0)
    ReDim ballvelx(0)
    ReDim ballvely(0)
  End Sub

  Public Sub Update() 'tracks in-ball-velocity
    Dim str, b, AllBalls, highestID
    allBalls = GetBalls

    For Each b In allballs
      If b.id >= HighestID Then highestID = b.id
    Next

    If UBound(ballvel) < highestID Then ReDim ballvel(highestID)  'set bounds
    If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)  'set bounds
    If UBound(ballvely) < highestID Then ReDim ballvely(highestID)  'set bounds

    For Each b In allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.

Sub RDampen_Timer
  Cor.Update
End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9  'Level of bounces. Recommmended value of 0.7-1

Sub TargetBouncer(aBall,defvalue)
  Dim zMultiplier, vel, vratio
  If TargetBouncerEnabled = 1 And aball.z < 30 Then
    '   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    vel = BallSpeed(aBall)
    If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
    Select Case Int(Rnd * 6) + 1
      Case 1
        zMultiplier = 0.2 * defvalue
      Case 2
        zMultiplier = 0.25 * defvalue
      Case 3
        zMultiplier = 0.3 * defvalue
      Case 4
        zMultiplier = 0.4 * defvalue
      Case 5
        zMultiplier = 0.45 * defvalue
      Case 6
        zMultiplier = 0.5 * defvalue
    End Select
    aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
    aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
    aBall.vely = aBall.velx * vratio
    '   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    '   debug.print "conservation check: " & BallSpeed(aBall)/vel
  End If
End Sub

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer ActiveBall, 1
End Sub



'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS
Set LS = New SlingshotCorrection
Dim RS
Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00, - 4
  AddSlingsPt 1, 0.45, - 7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4
End Sub

Sub AddSlingsPt(idx, aX, aY)    'debugger wrapper for adjusting flipper script In-game
  Dim a
  a = Array(LS, RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function
'
'Function RotPoint(x,y,angle)
' dim rx, ry
' rx = x*dCos(angle) - y*dSin(angle)
' ry = x*dSin(angle) + y*dCos(angle)
' RotPoint = Array(rx,ry)
'End Function

Class SlingshotCorrection
  Public DebugOn, Enabled
  Private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut

  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
    Enabled = True
  End Sub

  Public Property Let Object(aInput)
    Set Slingshot = aInput
  End Property

  Public Property Let EndPoint1(aInput)
    SlingX1 = aInput.x
    SlingY1 = aInput.y
  End Property

  Public Property Let EndPoint2(aInput)
    SlingX2 = aInput.x
    SlingY2 = aInput.y
  End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    Dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1
      YL = SlingY1
      XR = SlingX2
      YR = SlingY2
    Else
      XL = SlingX2
      YL = SlingY2
      XR = SlingX1
      YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If Abs(XR - XL) > Abs(YR - YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If Not IsEmpty(ModIn(0) ) Then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      '   debug.print " BallPos=" & BallPos &" Angle=" & Angle
      '   debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled Then aBall.Velx = RotVxVy(0)
      If Enabled Then aBall.Vely = RotVxVy(1)
      '   debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      '   debug.print " "
    End If
  End Sub
End Class

Sub Spell(tmpLetter)
Dim i
  Select Case tmpLetter
    Case "a"
      i = 0
    Case "b"
      i = 1
    Case "c"
      i = 2
    Case "d"
      i = 3
    Case "e"
      i = 4
    Case "f"
      i = 5
    Case "g"
      i = 6
    Case "h"
      i = 7
    Case "i"
      i = 8
    Case "j"
      i = 9
    Case "k"
      i = 10
    Case "l"
      i = 11
    Case "m"
      i = 12
    Case "n"
      i = 13
    Case "o"
      i = 14
    Case "p"
      i = 15
    Case "q"
      i = 16
    Case "r"
      i = 17
    Case "s"
      i = 18
    Case "t"
      i = 19
    Case "u"
      i = 20
    Case "v"
      i = 21
    Case "w"
      i = 22
    Case "x"
      i = 23
    Case "y"
      i = 24
    Case "z"
      i = 25
    Case Else
      Exit Sub
  End Select



axmas(i).timerInterval= 10
axmas(i).opacity = 100
fadedir(i) = 1
fspeed(i) = myFSpeed
Intensity(i) = 1
axmas(i).timerenabled=1
'vpmtimer.addtimer 2000 , "axmas(i).timerenabled=0:axmas(i).opacity = 0 '"

ExecuteGlobal ("Sub " & axmas(i).Name & "_timer:" & _
"If Intensity("& i &") <=0 Then me.timerenabled=0:End If:" & _
"If Intensity("& i &") >= 3000 Then fadeDir("& i &")=-1 :End If:" & _
"Intensity("& i &") = Intensity("& i &") + fSpeed("& i &") * fadeDir("& i &"):" & _
"axmas("& i &").opacity = intensity("& i &"):" & _
"End Sub")

End Sub

Dim bInLightQuote : bInLightQuote = False
Sub RandomLightQuote

  if bInLightQuote then Exit Sub

  PrepSpellWord 60

  bInLightQuote = True

  Select Case (Int(Rnd*5)+1)
    Case 1
      Spellword "friends dont lie"
    Case 2
      Spellword "a neverending story"
    Case 3
      Spellword "mornings are for coffee and contemplation"
    Case 4
      Spellword "i dump your ass"
    Case 5
      Spellword "she will not be able to resist these pearls"
  End Select

End Sub


Sub SetLoopDelay(xString)
  LoopDelay =  (Len(xString)*delayInc*myFSpeed) + (delayInc*myFSpeed*2)
  debug.print "loop delay " &loopdelay
End Sub

Sub TestMerlin
  debug.print "in test"


  dim k, newLoopDelay

    PrepSpellWord 40
    SetLoopDelay "adventuring party"

    SpellWord "adventuring party"

    for k = 1 to 7

      NewLoopDelay = LoopDelay*k
    debug.print "loop delay :" &k &":" &newloopdelay

      xmasQueue.Add "SpellWord-"&k,"SpellWord ""adventuring party"" ",80,NewLoopDelay,0,0,0,False
    Next

    xmasQueue.Add "SpellEnd","giOn:bInLightQuote = False",75,LoopDelay*8+(delayInc*myFSpeed*2),0,0,0,False


Exit Sub
  dim k2, newLoopDelay2

    PrepSpellWord 40
    SetLoopDelay "find barb"

    SpellWord "find barb"

    for k = 1 to 7

      NewLoopDelay = LoopDelay*k
    debug.print "loop delay :" &k &":" &newloopdelay

      xmasQueue.Add "SpellWord-"&k,"SpellWord ""find barb"" ",80,NewLoopDelay,0,0,0,False
    Next

    xmasQueue.Add "SpellEnd","giOn:bInLightQuote = False",75,LoopDelay*8+(delayInc*myFSpeed*2),0,0,0,False
End Sub


Dim LoopDelay
dim delay,myFSpeed
delay = 0
LoopDelay = 0

Sub PrepSpellWord (tmpSpeed)
  myFSpeed = tmpSpeed
  gioff
  'StopXMAS
  GeneralPupQueue.Add "StopXMAS","StopXMAS",80,10,0,0,0,False
End Sub

Dim DelayInc : DelayInc = 25

Sub SpellWord(xString)
dim i,tmpX, delayInc

debug.print "WORD:" & xString

  delay = 0
  delayInc = 25
  For i = 1 To Len(xString)
    tmpX = chr(34) &Mid(xString,i,1) & chr(34)
    delay = delay + delayInc*myFSpeed
    GeneralPupQueue.Add "Spell-"&i,"Spell " &tmpX,95,delay,0,0,0,True
  Next
' delay = delay + delayInc*myFSpeed
  'GeneralPupQueue.Add "Spell-"&i+1,"giOn:bInLightQuote = False",95,delay,0,0,0,True
End Sub


'  "If Intensity("& i &") <=0 Then axmas(i).timerenabled=0:axmas(i).opacity = 0:End If:" & _


'***************************************************************
' ZQUE: VPIN WORKSHOP ADVANCED QUEUING SYSTEM - 1.2.0
'***************************************************************
' WHAT IS IT?
' The VPin Workshop Advanced Queuing System allows table authors
' to put sub routine calls in a queue without creating a bunch
' of timers. There are many use cases for this: queuing sequences
' for light shows and DMD scenes, delaying solenoids until the
' DMD is finished playing all its sequences (such as holding a
' ball in a scoop), managing what actions take priority over
' others (e.g. an extra ball sequence is probably more important
' than a small jackpot), and many more.
'
' This system uses Scripting.Dictionary, a single timer, and the
' GameTime global to keep track of everything in the queue.
' This allows for better stability and a virtually unlimited
' number of items in the queue. It also allows for greater
' versatility, like pre-delays, queue delays, priorities, and
' even modifying items in the queue.
'
' The VPin Workshop Queuing System can replace vpmTimer as a
' proper queue system (each item depends on the previous)
' whereas vpmTimer is a collection of virtual timers that run
' in parallel. It also adds on other advanced functionality.
' However, this queue system does not have ROM support out of
' the box like vpmTimer does.
'
' I recommend reading all the comments before you implement the
' queuing system into your table.
'
' WHAT YOU NEED to use the queuing system:
' 1) Put this VBS file in your scripts folder, or copy / paste
'    the code into your table script (and skip step 2).
' 2) Include this file via Scripting.FileSystemObject, and
'    ExecuteGlobal it.
' 3) Make one or more queues by constructing the vpwQueueManager:
'    Dim queue : Set queue = New vpwQueueManager
' 4) Create (or use) a timer that is always enabled and
'    preferably has an interval of 1 millisecond. Use a
'    higher number for less time precision but less resource
'    use. You only need one timer even if you
'    have multiple queues.
' 5) For each queue you created, call its Tick routine in
'    the timer's *_timer() routine:
'    queue.Tick
' 6) You're done! Refer to the routines in vpwQueueManager to
'    learn how to use the queuing system.
'
' TUTORIAL: https://youtu.be/kpPYgOiUlxQ
'***************************************************************

'===========================================
' vpwQueueManager
' This class manages a queue of
' vpwQueueItems and executes them.
'===========================================
Class vpwQueueManager
  Public qItems ' A dictionary of vpwQueueItems in the queue (do NOT use native Scripting.Dictionary.Add/Remove; use the vpwQueueManager's Add/Remove methods instead!)
  Public preQItems ' A dictionary of vpwQueueItems pending to be added to qItems
  Public debugOn 'Null = no debug. String = activate debug by using this unique label for the queue. REQUIRES baldgeek's error logs.

  '----------------------------------------------------------
  ' vpwQueueManager.qCurrentItem
  ' This contains a string of the key currently active / at
  ' the top of the queue. An empty string means no items are
  ' active right now.
  ' This is an important property; it should be monitored
  ' in another timer or routine whenever you Add a queue item
  ' with a -1 (indefinite) preDelay or postDelay. Then, for
  ' preDelay, ExecuteCurrentItem should be called to run the
  ' queue item. And for postDelay, DoNextItem should be
  ' called to move to the next item in the queue.
  '
  ' For example, let's say you add a queue item with the
  ' key "kickTheBall" and an indefinite preDelay. You want
  ' to wait until another timer fires before this queue item
  ' executes and kicks the ball out of a scoop. In the other
  ' timer, you will monitor qCurrentItem. Once it equals
  ' "kickTheBall", call ExecuteCurrentItem, which will run
  ' the queue item and presumably kick out the ball.
  '
  ' WARNING!: If you do not properly execute one of these
  ' callback routines on an indefinite delayed item, then
  ' the queue will effectively freeze / stop until you do.
  '---------------------------------------------------------
  Public qCurrentItem

  Public preDelayTime ' The GameTime the preDelay for the qCurrentItem was started
  Public postDelayTime ' The GameTime the postDelay for the qCurrentItem was started

  Private onQueueEmpty ' A string or object to be called every time the queue empties (use the QueueEmpty property to get/set this)
  Private queueWasEmpty ' Boolean to determine if the queue was already empty when firing DoNextItem
  Private preDelayTransfer ' Number of milliseconds of preDelay to transfer over to the next queue item when doNextItem is called

  Private Sub Class_Initialize
    Set qItems = CreateObject("Scripting.Dictionary")
    Set preQItems = CreateObject("Scripting.Dictionary")
    qCurrentItem = ""
    onQueueEmpty = ""
    queueWasEmpty = True
    debugOn = Null
    preDelayTransfer = 0
  End Sub

  '----------------------------------------------------------
  ' vpwQueueManager.Tick
  ' This is where all the magic happens! Call this method in
  ' your timer's _timer routine to check the queue and
  ' execute the necessary methods. We do not iterate over
  ' every item in the queue here, which allows for superior
  ' performance even if you have hundreds of items in the
  ' queue.
  '----------------------------------------------------------
  Public Sub Tick()
    Dim item
    If qItems.Count > 0 Then ' Don't waste precious resources if we have nothing in the queue

      ' If no items are active, or the currently active item no longer exists, move to the next item in the queue.
      ' (This is also a failsafe to ensure the queue continues to work even if an item gets manually deleted from the dictionary).
      If qCurrentItem = "" Or Not qItems.Exists(qCurrentItem) Then
        DoNextItem
      Else ' We are good; do stuff as normal
        Set item = qItems.item(qCurrentItem)

        If item.Executed Then
          ' If the current item was executed and the post delay passed, go to the next item in the queue
          If item.postDelay >= 0 And GameTime >= (postDelayTime + item.postDelay) Then
            DebugLog qCurrentItem & " - postDelay of " & item.postDelay & " passed."
            DoNextItem
          End If
        Else
          ' If the current item expires before it can be executed, go to the next item in the queue
          If item.timeToLive > 0 And GameTime >= (item.queuedOn + item.timeToLive) Then
            DebugLog qCurrentItem & " - expired (Time To live). Moving To the Next queue item."
            DoNextItem
          End If

          ' If the current item was not executed yet and the pre delay passed, then execute it
          If item.preDelay >= 0 And GameTime >= (preDelayTime + item.preDelay) Then
            DebugLog qCurrentItem & " - preDelay of " & item.preDelay & " passed. Executing callback."
            item.Execute
            preDelayTime = 0
            postDelayTime = GameTime
          End If
        End If
      End If
    End If

    ' Loop through each item in the pre-queue to find any that is ready to be added
    If preQItems.Count > 0 Then
      Dim k, key
      k = preQItems.Keys
      For Each key In k
        Set item = preQItems.Item(key)

        ' If a queue item was pre-queued and is ready to be considered as actually in the queue, add it
        If GameTime >= (item.queuedOn + item.preQueueDelay) Then
          DebugLog key & " (preQueue) - preQueueDelay of " & item.preQueueDelay & " passed. Item added To the main queue."
          preQItems.Remove key
          Me.Add key, item.Callback, item.priority, 0, item.preDelay, item.postDelay, item.timeToLive, item.executeNow
        End If
      Next
    End If
  End Sub

  '----------------------------------------------------------
  ' vpwQueueManager.DoNextItem
  ' Goes to the next item in the queue and deletes the
  ' currently active one.
  '----------------------------------------------------------
  Public Sub DoNextItem()
    If Not qCurrentItem = "" Then
      If qItems.Exists(qCurrentItem) Then qItems.Remove qCurrentItem ' Remove the current item from the queue if it still exists
      qCurrentItem = ""
    End If

    If qItems.Count > 0 Then
      Dim k, key
      Dim nextItem
      Dim nextItemPriority
      Dim item
      nextItemPriority = 0
      nextItem = ""

      ' Find which item needs to run next based on priority first, queue order second (ignore items with an active preQueueDelay)
      k = qItems.Keys
      For Each key In k
        Set item = qItems.Item(key)

        If item.preQueueDelay <= 0 And item.priority > nextItemPriority Then
          nextItem = key
          nextItemPriority = item.priority
        End If
      Next

      If qItems.Exists(nextItem) Then
        Set item = qItems.Item(nextItem)
        DebugLog "DoNextItem - checking " & nextItem & " (priority " & item.priority & ")"

        ' Make sure the item is not expired and not already executed. If it is, remove it and re-call doNextItem
        If (item.timeToLive > 0 And GameTime >= (item.queuedOn + item.timeToLive + preDelayTransfer)) Or item.executed = True Then
          DebugLog "DoNextItem - " & nextItem & " expired (Time To live) Or already executed. Removing And going To the Next item."
          qItems.Remove nextItem
          DoNextItem
          Exit Sub
        End If

        'Transfer preDelay time when applicable
        If preDelayTransfer > 0 And item.preDelay > -1 Then
          DebugLog "DoNextItem " & nextItem & " - Transferred remaining postDelay of " & preDelayTransfer & " milliseconds from previously overridden queue item To its preDelay And timeToLive"
          qItems.Item(nextItem).preDelay = item.preDelay + preDelayTransfer
          If item.timeToLive > 0 Then qItems.Item(nextItem).timeToLive = item.timeToLive + preDelayTransfer
          preDelayTransfer = 0
        End If

        ' Set item as current / active, and execute if it has no pre-delay (otherwise Tick will take care of pre-delay)
        qCurrentItem = nextItem
        If item.preDelay = 0 Then
          DebugLog "DoNextItem - " & nextItem & " Now active. It has no preDelay, so executing callback immediately."
          item.Execute
          preDelayTime = 0
          postDelayTime = GameTime
        Else
          DebugLog "DoNextItem - " & nextItem & " Now active. Waiting For a preDelay of " & item.preDelay & " before executing."
          preDelayTime = GameTime
          postDelayTime = 0
        End If
      End If
    ElseIf queueWasEmpty = False Then
      DebugLog "DoNextItem - Queue Is Now Empty; executing queueEmpty callback."
      CallQueueEmpty() ' Call QueueEmpty if this was the last item in the queue
    End If
  End Sub

  '----------------------------------------------------------
  ' vpwQueueManager.ExecuteCurrentItem
  ' Helper routine that can be used when the current item is
  ' on an indefinite preDelay. Call this when you are ready
  ' for that item to execute.
  '----------------------------------------------------------
  Public Sub ExecuteCurrentItem()
    If Not qCurrentItem = "" And qItems.Exists(qCurrentItem) Then
      DebugLog "ExecuteCurrentItem - Executing the callback For " & qCurrentItem & "."
      Dim item
      Set item = qItems.Item(qCurrentItem)
      item.Execute
      preDelayTime = 0
      postDelayTime = GameTime
    End If
  End Sub

  '----------------------------------------------------------
  ' vpwQueueManager.Add
  ' REQUIRES Class vpwQueueItem
  '
  ' Add an item to the queue.
  '
  ' PARAMETERS:
  '
  ' key (string) - Unique name for this queue item
  ' WARNING: Specifying a key that already exists will
  ' overwrite the item in the queue. This is by design. Also
  ' note the following behaviors:
  ' * Tickers / clocks for tracking delay times will NOT be
  ' restarted for this item (but the total duration will be
  ' updated. For example, if the old preDelay was 3 seconds
  ' and 2 seconds elapsed, but Add was called to update
  ' preDelay to 5 seconds, then the queue item will now
  ' execute in 3 more seconds (new preDelay - time elapsed)).
  ' However, timeToLive WILL be restarted.
  ' * Items will maintain their same place in the queue.
  ' * If key = qCurrentItem (overwriting the currently active
  ' item in the queue) and qCurrentItem already executed
  ' the callback (but is waiting for a postDelay), then the
  ' current queue item's remaining postDelay will be added to
  ' the preDelay of the next item, and this item will be
  ' added to the bottom of the queue for re-execution.
  ' If you do not want it to re-execute, then add an If
  ' guard on your call to the Add method checking
  ' "If Not vpwQueueManager.qCurrentItem = key".
  '
  ' qCallback (object|string) - An object to be called,
  ' or string to be executed globally, when this queue item
  ' runs. I highly recommend making sub routines for groups
  ' of things that should be executed by the queue so that
  ' your qCallback string does not get long, and you can
  ' easily organize your callbacks. Also, use double
  ' double-quotes when the call itself has quotes in it
  ' (VBScript escaping).
  ' Example: "playsound ""Plunger"""
  '
  ' priority (number) - Items in the queue will be executed
  ' in order from highest priority to lowest. Items with the
  ' same priority will be executed in order according to
  ' when they were added to the queue. Use any number
  ' greater than 0. My recommendation is to make a plan for
  ' your table on how you will prioritize various types of
  ' queue items and what priority number each type should
  ' have. Also, you should reserve priority 1 (lowest) to
  ' items which should wait until everything else in the
  ' queue is done (such as ejecting a ball from a scoop).
  '
  ' preQueueDelay (number) - The number of
  ' milliseconds before the queue actually considers this
  ' item as "in the queue" (pretend you started a timer to
  ' add this item into the queue after this delay; this
  ' logically works in a similar way; the only difference is
  ' timeToLive is still considered even when an item is
  ' pre-queued.) Set to 0 to add to the queue immediately.
  ' NOTE: this should be less than timeToLive.
  '
  ' preDelay (number) - The number of milliseconds before
  ' the qCallback executes once this item is active (top)
  ' in the queue. Set this to 0 to immediately execute the
  ' qCallback when this item becomes active.
  ' Set this to -1 to have an indefinite delay until
  ' vpwQueueManager.ExecuteCurrentItem is called (see the
  ' comment for qCurrentItem for more information).
  ' NOTE: this should be less than timeToLive. And, if
  ' timeToLive runs out before preDelay runs out, the item
  ' will be removed and will not execute.
  '
  ' postDelay (number) - After the qCallback executes, the
  ' number of milliseconds before moving on to the next item
  ' in the queue. Set this to -1 to have an indefinite delay
  ' until vpwQueueManager.DoNextItem is called (see the
  ' comment for qCurrentItem for more information).
  '
  ' timeToLive (number) - After this item is added to the
  ' queue, the number of milliseconds before this queue item
  ' expires / is removed if the qCallback is not executed by
  ' then. Set to 0 to never expire. NOTE: If not 0, this
  ' should be greater than preDelay + preQueueDelay or the
  ' item will expire before the qCallback is executed.
  ' Example use case: Maybe a player scored a jackpot, but
  ' it would be awkward / irrelevant to play that jackpot
  ' sequence if it hasn't played after a few seconds (e.g.
  ' other items in the queue took priority).
  '
  ' executeNow (boolean) - Specify true if this item
  ' should interrupt the queue and run immediately. This
  ' will only happen, however, if the currently active item
  ' has a priority less than or equal to the item you are
  ' adding. Note this does not bypass preQueueDelay nor
  ' preDelay if set.
  ' Example: If a player scores an extra ball, you might
  ' want that to interrupt everything else going on as it
  ' is an important milestone.
  '----------------------------------------------------------
  Public Sub Add(key, qCallback, priority, preQueueDelay, preDelay, postDelay, timeToLive, executeNow)
    DebugLog "Adding queue item " & key

    'Construct the item class
    Dim newClass
    Set newClass = New vpwQueueItem
    With newClass
      .Callback = qCallback
      .priority = priority
      .preQueueDelay = preQueueDelay
      .preDelay = preDelay
      .postDelay = postDelay
      .timeToLive = timeToLive
      .executeNow = executeNow
    End With

    'If we are attempting to overwrite the current queue item which already executed, take the remaining postDelay and add it to the preDelay of the next item. And set us up to immediately go to the next item while re-adding this item to the queue.
    If preQueueDelay <= 0 And qItems.Exists(key) And qCurrentItem = key Then
      If qItems.Item(key).executed = True Then
        DebugLog key & " (Add) - Attempting To overwrite the current queue item which already executed. Immediately re-queuing this item To the bottom of the queue, transferring the remaining postDelay To the Next item, And going To the Next item."
        If qItems.Item(key).postDelay >= 0 Then
          preDelayTransfer = ((postDelayTime + qItems.Item(key).postDelay) - GameTime)
        End If

        'Remove current queue item so we can go to the next item, this can be re-queued to the bottom, and the remaining postDelay transferred to the preDelay of the next item
        qItems.Remove qCurrentItem
        qCurrentItem = ""
      End If
    End If

    ' Determine execution stuff if this item does not have a pre-queue delay
    If preQueueDelay <= 0 Then
      If executeNow = True Then
        ' Make sure this item does not immediately execute if the current item has a higher priority
        If Not qCurrentItem = "" And qItems.Exists(qCurrentItem) Then
          Dim item
          Set item = qItems.Item(qCurrentItem)
          If item.priority <= priority Then
            DebugLog key & " (Add) - Execute Now was Set To True And this item's priority (" & priority & ") Is >= the active item's priority (" & item.priority & " from " & qCurrentItem & "). Making it the current active queue item."
            qCurrentItem = key
            If preDelay = 0 And preDelayTransfer = 0 Then
              DebugLog key & " (Add) - No pre-delay. Executing the callback immediately."
              newClass.Execute
              preDelayTime = 0
              postDelayTime = GameTime
            Else
              DebugLog key & " (Add) - Waiting For a pre-delay of " & (preDelay + preDelayTransfer) & " before executing the callback."
              preDelayTime = GameTime
              postDelayTime = 0
            End If
          Else
            DebugLog key & " (Add) - Execute Now was Set To True, but this item's priority (" & priority & ") Is Not >= the active item's priority (" & item.priority & " from " & qCurrentItem & "). This item will Not be executed Now And will be added To the queue normally."
          End If
        Else
          DebugLog key & " (Add) - Execute Now was Set To True And no item was active In the queue. Making it the current active queue item."
          qCurrentItem = key
          If preDelay = 0 Then
            DebugLog key & " (Add) - No pre-delay. Executing the callback immediately."
            preDelayTransfer = 0 'No preDelay transfer if we are immediately re-executing the same queue item
            newClass.Execute
            preDelayTime = 0
            postDelayTime = GameTime
          Else
            DebugLog key & " (Add) - Waiting For a pre-delay of " & preDelay & " before executing the callback."
            preDelayTime = GameTime
            postDelayTime = 0
          End If
        End If
      End If
      If qItems.Exists(key) Then 'Overwrite existing item in the queue if it exists
        DebugLog key & " (Add) - Already exists In the queue. Updating the item With the new parameters passed In Add."
        Set qItems.Item(key) = newClass
      Else
        DebugLog key & " (Add) - Added To the queue."
        qItems.Add key, newClass
      End If
      queueWasEmpty = False
    Else
      If preQItems.Exists(key) Then 'Overwrite existing item in the preQueue if it exists
        DebugLog key & " (Add) - Already exists In the preQueue. Updating the item With the new parameters passed In Add."
        Set preQItems.Item(key) = newClass
      Else
        DebugLog key & " (Add) - Added To the preQueue."
        preQItems.Add key, newClass
      End If
    End If
  End Sub

  '----------------------------------------------------------
  ' vpwQueueManager.Remove
  '
  ' Removes an item from the queue. It is better to use this
  ' than to remove the item from qItems directly as this sub
  ' will also call DoNextItem to advance the queue if
  ' the item removed was the active item.
  ' NOTE: This only removes items from qItems; to remove
  ' an item from preQItems, use the standard
  ' Scripting.Dictionary Remove method.
  '
  ' PARAMETERS:
  '
  ' key (string) - Unique name of the queue item to remove.
  '----------------------------------------------------------
  Public Sub Remove(key)
    If qItems.Exists(key) Then
      DebugLog key & " (Remove)"
      qItems.Remove key
      If qCurrentItem = key Or qCurrentItem = "" Then DoNextItem ' Ensure the queue does not get stuck
    End If
  End Sub

  '----------------------------------------------------------
  ' vpwQueueManager.RemoveAll
  '
  ' Removes all items from the queue / clears the queue.
  ' It is better to call this sub than to remove all items
  ' from qItems directly because this sub cleans up the queue
  ' to ensure it continues to work properly.
  '
  ' PARAMETERS:
  '
  ' preQueue (boolean) - Also clear the pre-queue.
  '----------------------------------------------------------
  Public Sub RemoveAll(preQueue)
    DebugLog "Queue was emptied via RemoveAll."

    ' Loop through each item in the queue and remove it
    Dim k, key
    k = qItems.Keys
    For Each key In k
      qItems.Remove key
    Next
    qCurrentItem = ""

    If queueWasEmpty = False Then CallQueueEmpty() ' Queue is now empty, so call our callback if applicable

    If preQueue Then
      k = preQItems.Keys
      For Each key In k
        preQItems.Remove key
      Next
    End If
  End Sub

  '----------------------------------------------------------
  ' Get vpwQueueManager.QueueEmpty
  ' Get the current callback for when the queue is empty.
  '----------------------------------------------------------
  Public Property Get QueueEmpty()
    If IsObject(onQueueEmpty) Then
      Set QueueEmpty = onQueueEmpty
    Else
      QueueEmpty = onQueueEmpty
    End If
  End Property

  '----------------------------------------------------------
  ' Let vpwQueueManager.QueueEmpty
  ' Set the callback to call every time the queue empties.
  ' This could be useful for setting a sub routine to be
  ' called each time the queue empties for doing things such
  ' as ejecting balls from scoops. Unlike using the Add
  ' method, this callback is immune from getting removed by
  ' higher priority items in the queue and will be called
  ' every time the queue is emptied, not just once.
  '
  ' PARAMETERS:
  '
  ' callback (object|string) - The callback to call every
  ' time the queue empties.
  '----------------------------------------------------------
  Public Property Let QueueEmpty(callback)
    If IsObject(callback) Then
      Set onQueueEmpty = callback
    ElseIf VarType(callback) = vbString Then
      onQueueEmpty = callback
    End If
  End Property

  '----------------------------------------------------------
  ' Get vpwQueueManager.CallQueueEmpty
  ' Private method that actually calls the QueueEmpty
  ' callback.
  '----------------------------------------------------------
  Private Sub CallQueueEmpty()
    If queueWasEmpty = True Then Exit Sub
    queueWasEmpty = True

    If IsObject(onQueueEmpty) Then
      Call onQueueEmpty(0)
    ElseIf VarType(onQueueEmpty) = vbString Then
      If onQueueEmpty > "" Then ExecuteGlobal onQueueEmpty
    End If
  End Sub

  '----------------------------------------------------------
  ' DebugLog
  ' Log something if debugOn is not null.
  ' REQUIRES / uses the WriteToLog sub from Baldgeek's
  ' error log library.
  '----------------------------------------------------------
  Private Sub DebugLog(message)
    If Not IsNull(debugOn) Then
      WriteToLog "VPW Queue " & debugOn, message
    End If
  End Sub
End Class

'===========================================
' vpwQueueItem
' Represents a single item for the queue
' system. Do NOT use this class directly.
' Instead, use the vpwQueueManager.Add
' routine.

' You can, however, access an individual
' item in the queue via
' vpwQueueManager.qItems and then modify
' its properties while it is still in the
' queue.
'===========================================
Class vpwQueueItem  ' Do not construct this class directly; use vpwQueueManager.Add instead, and vpwQueueManager.qItems.Item(key) to modify an item's properties.
  Public priority ' The item's set priority
  Public timeToLive ' The item's set timeToLive milliseconds requested
  Public preQueueDelay ' The item's pre-queue milliseconds requested
  Public preDelay ' The item's pre delay milliseconds requested
  Public postDelay ' The item's post delay milliseconds requested
  Public executeNow ' Whether the item was set to Execute immediately
  Private qCallback ' The item's callback object or string (use the Callback property on the class to get/set it)

  Public executed ' Whether or not this item's qCallback was executed yet
  Public queuedOn ' The game time this item was added to the queue
  Public executedOn ' The game time this item was executed

  Private Sub Class_Initialize
    ' Defaults
    priority = 0
    timeToLive = 0
    preQueueDelay = 0
    preDelay = 0
    postDelay = 0
    qCallback = ""
    executeNow = False

    queuedOn = GameTime
    executedOn = 0
  End Sub

  '----------------------------------------------------------
  ' vpwQueueItem.Execute
  ' Executes the qCallback on this item if it was not yet
  ' already executed.
  Public Sub Execute()
    If executed Then Exit Sub ' Do not allow an item's qCallback to ever Execute more than one time

    'Mark as execute before actually executing callback; that way, if callback recursively adds the item back into the queue, then we can properly handle it.
    executed = True
    executedOn = GameTime

    ' Execute qCallback
    If IsObject(qCallback) Then
      Call qCallback(0)
    ElseIf VarType(qCallback) = vbString Then
      If qCallback > "" Then ExecuteGlobal qCallback
    End If
  End Sub

  Public Property Get Callback()
    If IsObject(qCallback) Then
      Set Callback = qCallback
    Else
      Callback = qCallback
    End If
  End Property

  Public Property Let Callback(cb)
    If IsObject(cb) Then
      Set qCallback = cb
    ElseIf VarType(cb) = vbString Then
      qCallback = cb
    End If
  End Property
End Class



Sub WipeAllQueues
  GeneralPupQueue.RemoveAll(True)
  xmasQueue.RemoveAll(True)
  BallHandlingQueue.RemoveAll(True)
End Sub

Sub QueueTimer_Timer()
  BallHandlingQueue.Tick
  GeneralPupQueue.Tick
  xmasQueue.Tick
End Sub
'***************************************************************
' END VPIN WORKSHOP ADVANCED QUEUING SYSTEM
'***************************************************************
