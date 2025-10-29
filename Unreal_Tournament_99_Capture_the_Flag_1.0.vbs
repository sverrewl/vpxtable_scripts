'*
'* Unreal Tournament 99 Capture The Flag Pinball  2021
'*
'* OK, OK, OK! take two !
'* Sometimes Leo Getz more than he asked for !
'*
'* I GOT THE FLAG PINBALL !
'*
'* Based on Lethal Weapon 3 vpw modv1.02
'* no ROM is needed nor used for this table
'* db2s included
'*
'* FlexDMD is needed so go get:
'*  https://github.com/freezy/dmd-extensions/releases
'*    Download the installer and run it.
'*
'*  https://github.com/vbousquet/flexdmd/releases
'*  copy files to vpinmame dir and run the flexDMDUI to register dlls
'*
'* Oqqsan         First King Of Morpheus ! Full Script, new ramps, inserts, a finger on everything
'* Major Frenchy      Playfield flippers and laneguides gfx
'* SiXtoe         VR-room
'* Tomate, iaakki     Nfozzy physics
'* Apophis          DOF  +  FleepPackage
'*
'* Testing: Hard to manage without !
'* Apophis
'* fluffhead35
'* RIK
'* VPW team
'*
'*
'* Features:
'* 1-4 Players
'* 900 Different Sound Files
'* FlexDMD
'* FlupperDomes 2
'* Nfozzy phsics
'* FleepSounds
'* OrbitalScoreBoard
'* press "R" Will magnify Instructions
'* left magnasave  - select player TEAM color
'* right magnesave - select player main taunt voice
'*
'*
'* Playd alot on LW3 vpw mod and saw the potential, this table is very good !
'* If you havent still playd that one go get it, you will not be disapointed !
'*
'*
'*
'* OK, OK, OK!
'* Lethal Weapon 3 table Tune-Up by VPin Workshop.
'* It all started by again by just thinking about adding nFozzys physics and some missing textures in the original table,
'* then we started replacing some primitives and reworking the existing ones, then the apron, cabinet, metal ramp, bats and a few other things were replaced.
'* Totally new 3d inserts. Then different lamps were added to have a better lighting of all the new baked textures and finally adding a built in VR room...
'*
'* - New/reworked primitives: tomate
'* - Baked ON/OFF textures: tomate
'* - Inserts: Sheltemke, oqqsan
'* - Scripting: oqqsan
'* - nFozzys physics:
'* - Fleep sounds: apophis
'* - Flashes & Lighting Overhaul: tomate, Sixtoe
'* - VR Room and fixes: Sixtoe, 360 Room by pattyg234.
'* - Miscellaneous tweaks: Sixtoe, tomate, oqqsan
'* - Testing: Bord, Rik, oqqsan, VPW team
'*
'* This release would not have been possible without the legacy of those who came before us, including most notably, Javier for the original vpx table, EBisLit for the new playfield image and 32Assassin for the rework. Thank you.
'* Thanks to Flupper for his beautiful domes and Bumpers caps, Rothbauerw for nFozzy physics and Fleep for sounds.
'*
'*
'* Enjoy! x2


Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


'*********** General Sound Options *********************
'*   Recommended values should be no greater than 1   *
Const VolumeDial =      0.9 ' global volume multiplier for the mechanical sounds.
Const BG_Volume =     1.0 ' backglass volume
Const Music_Volume =    0.7 ' music volume
'***********        Options        *********************
Const OSBactive =     0   ' 1=on 0=off    .. orbital scoreboard : NOt on by default !! if you have this set it to 1
Dim DB2S_on : DB2S_on = 0   ' 1=on 0=auto  : dB2S
RainFlasher.visible =     1   ' 0=off 1=on for raining mod
Const Gunshake =      1   ' 0=off 1=on for shaking PF when gun fire , probaly best to disable this if you go VRROOOM!
Const VRRoom =        0   ' 0=off 1=on 2=Ultra Minimal



Const DOFdebug =      0         ' set to 1 to turn on debug on dof events
Const StaticScoreboard =  0         ' if the vidoes on scoreboard are bad for you set this to 1
Table1.BallPlayfieldReflectionScale=  150   ' any value, bigger number = Ball will get brighter : default 150

'DOF id
'E101 0/1 LeftFlipper
'E102 0/1 RightFlipper
'E103 2 Leftslingshot
'E104 2 Rightslingshot
'E105 2 top bumper left
'E106 2 top bumper center
'E107 2 top bumper right
'E108 2 front blue Flasher  8
'E109 2 middle blu Flasher 3
'E110 2 back blu Flasher   9
'E111 2 front red Flasher   1
'E112 2 middle red Flasher  2
'E113 2 back red Flasher    7
'E114 0/1 Lauchbutton ready  ( guncontrol )
'E115 0/1 Startgamebutton ready
'E116 2 startgame
'E117 2 autolaunch/lauchbutton Fire
'E118 2 kickback fire
'E119 2 drain Hit
'E120 2 ballsaved
'E121 2 Left Spinner
'E122 2 Right spinner
'E123 2 skillshot
'E124 2 skillshot perfect
'E125 2 skillshot 2
'E126 2 award extraball
'E127 2 add credits
'E128 2 middle DT
'E129 2 right DT
'E130 2 Pwnage
'E131 2 lite Kickback
'E133 2 invisibility awarded
'E134 2 left saucer
'E135 2 left orbit
'E136 2 TopVUK
'E137 2 right saucer
'E138 2 right orbit
'E139 2 mainramp
'E140 2 redeemer
'E141 2 Jackpot
'E142 2 HeadShot
'E143 2 ball locked
'E144 0/1 Multiball
'E145 2 Godlike
'E146 2 Monsterkill
'E147 0/1 redundercab
'E148 0/1 blueundercab
'E149 0/1 goldundercabdo
'E149 0/1 goldundercab
'E150 0/1 greenundercab
'E151 2 Ballrelease ( simulated )
'E152 2 4 skillshot lights hit
'E153 2 Capture
'E154 2 ENEMY Capture
'E155 2 Capture Denied



' enemy flag ... auto cap after 50 seconds
' lvl 3 will have automatic opponent cap at around 75 seconds. lvl 4 will have another at 180sec
'
'
' Select Player team With left magna save
Dim Team(4)
Dim Taunt(4)
Dim highscorename(3)
Dim hScore(3)
Dim LastScore
Dim GamesPlayd
Const cGameName="ut99ctf"
Const OSDversion = "1.00"
gatesprims.blenddisablelighting=2
cab.blenddisablelighting=0.05
Dim DesktopMode:DesktopMode = Table1.ShowDT
const RubberizerEnabled = 1


If DB2S_on=1 or not Desktopmode Then startcontroller
If OSBactive = 1 Then
  On Error Resume Next
  ExecuteGlobal GetTextFile("osb.vbs")
End If
Sub Table1_init
  If NOT Desktopmode Then Flasher006.visible=0 : Flasher007.visible=0
  FlexDMDInit
  Loadhighscore
  UmainDMD.DisplayScene00ExWithId "attract1",true, "introvideo.wmv", "  ", 15, 5, "  ", 15, 0, 2, 85000, 14
  Attract1.enabled=1
  PlaySound "AAOpeningVO", 1, 0.6 * BG_Volume, 0, 0,0,0, 0, 0
  Playmusic "Unrealtournament/UT-Title.mp3"
  musicvolume = Music_Volume
  SetTeamColors
  InitLightsXY
End Sub

'**********************************************************************************************************
'*
'**********************************************************************************************************

Dim SLS(200,14)' 1-lIGHTS  FLAG 0-1-2 , 2FADELEVEL ,3BLINKING PAUSE , 4BLINKING COUNTDOWN, 5SPECIALSTATE
            ' 6SP BLINK PAUSE, 7SP COUNTDOWN, 8SP nrofblinks(0=noend),9Normal nrofblinks : 10=SP off afterxx frames  11=DELAYSTART !  12=turn off special after done ( 12 = always so mby takeitaway )
      '13 and 14 coodinates

Dim SC(4,50) ' player flags
'   SC(PN,x)=1
'    0=bounsscore for after
'  1  Score
'  2R Bonus1  3=Bonus2  4=Bonus3  11=MULTIPLYERLEVEL (2,3,4,6,8)
'  5R  DT1 Sw25 6=DT2 Sw26  7=DT3 Sw27  8=DT4 Sw33  9=DT5 Sw34  10=DT6 Sw35  ... all DT/collected armor is lost each ball
' 12  Rightorbit
' 13  Leftorbit nr(1-5=nr of lights lit) + more states ?
' 14  kicker1
' 15  kicker2
' 16  kicker3
' 17  Lockedballs
' 18  Needed to start Locking for multiball ( 5 lights )
' 19  Killingspree
'   20 HUMAN MAIN TAUNTS    <<< no reset
'   21 Skilshotlevel
'
'
' 22= 65 and 66   blinks ' for reentry after lost ball so this can stay
' 23= 42 capture  blinks
' 24= 43 upgrade  blinks
' 25= 50 enemyF blinks
' 26= 68 teamplay blinks
'
'   27=Map Playing
' 28= map score left Counter
' 29= map score Right Counter
'   30= enemyflagcounter for next lit EF
' 31=GameTimerCheck
'   32= GFBlevel  get flag back
' 33= save last countdown for reset next ball





Dim Controller
Sub startcontroller
  Set Controller = CreateObject("B2S.Server")
  Controller.B2SName = "ut99ctf"
  Controller.Run
  B2SOn = True
  DB2S_on=1
End Sub

Dim B21
Sub B2STimer1_Timer
  If B2STimer1.enabled=0 Then
    B2STimer1.enabled=1
    B21=6
  Else
    B21=B21-1
    Select Case B21
      case 0,2,4 : Controller.B2SSetData 10,0
      case 1,3,5 : Controller.B2SSetData 10,1
    End Select
    If B21<0 Then B2STimer1.enabled=0
  End If
End Sub
Dim B22
Sub B2STimer2_Timer
  If B2STimer2.enabled=0 Then
    B2STimer2.enabled=1
    B22=6
  Else
    B22=B22-1
    Select Case B22
      case 0,2,4 : Controller.B2SSetData 11,0
      case 1,3,5 : Controller.B2SSetData 11,1
    End Select
    If B22<0 Then B2STimer2.enabled=0
  End If
End Sub


Sub SavePlayerLights
  If countdown30.enabled=1 Then
    countdown30.enabled=0
    SC(PN,33)=lastcd+1000
  End If
  If SLS(65,1)>0 Then SLS(65,1)=2
  SC(PN,22)=SLS(65,1)

  If SLS(42,1)>0 Then SLS(42,1)=2
  SC(PN,23)=SLS(42,1)

  If SLS(43,1)>0 Then SLS(43,1)=2
  SC(PN,24)=SLS(43,1)

  If SLS(41,1)>0 Then SLS(41,1)=2
  SC(PN,25)=SLS(41,1)

  If SLS(68,1)>0 Then SLS(68,1)=2 '
  SC(PN,26)=SLS(68,1)
End Sub



Dim PN , MaxPlayers ' current player,maxplayer=total entered

Dim tempstr, tempstr2, tempnr, tempnr2, startgame         'misc



Dim MBActive  , Locked1 , Locked2
Dim Firstblood , FastFrags
Dim AutoKick
Dim Tilt, Tilted,TiltedBalls
Const TiltSens = 50
Dim Ownage, BallInPlay, Credits
Dim superbumpers
Dim Invisiblilty(5)



'init ' must reset all
FFlevel=0
KSlevel=0
LostBalls=0
MBActive=0
AutoKick=0
Ownage=0

superbumpers=0



'**********************************************************************************************************
'* CAPTURE DISPLAYS *
'**********************************************************************************************************

Sub BlueDisplay(nr)
  Select Case nr
    Case 0 : B1.state=1 : B2.state=1 : B3.state=1 : B4.state=1 : B5.state=1 : B6.state=1 : B7.state=0
    Case 1 : B1.state=0 : B2.state=1 : B3.state=1 : B4.state=0 : B5.state=0 : B6.state=0 : B7.state=0
    Case 2 : B1.state=1 : B2.state=1 : B3.state=0 : B4.state=1 : B5.state=1 : B6.state=0 : B7.state=1
    Case 3 : B1.state=1 : B2.state=1 : B3.state=1 : B4.state=1 : B5.state=0 : B6.state=0 : B7.state=1
    Case 4 : B1.state=0 : B2.state=1 : B3.state=1 : B4.state=0 : B5.state=0 : B6.state=1 : B7.state=1
    Case 5 : B1.state=1 : B2.state=0 : B3.state=1 : B4.state=1 : B5.state=0 : B6.state=1 : B7.state=1
    Case 6 : B1.state=0 : B2.state=0 : B3.state=1 : B4.state=1 : B5.state=1 : B6.state=1 : B7.state=1
    Case 7 : B1.state=1 : B2.state=1 : B3.state=1 : B4.state=0 : B5.state=0 : B6.state=0 : B7.state=0
    Case 8 : B1.state=1 : B2.state=1 : B3.state=1 : B4.state=1 : B5.state=1 : B6.state=1 : B7.state=1
    Case 9 : B1.state=1 : B2.state=1 : B3.state=1 : B4.state=0 : B5.state=0 : B6.state=1 : B7.state=1
    Case 10 :B1.state=0 : B2.state=0 : B3.state=0 : B4.state=0 : B5.state=0 : B6.state=0 : B7.state=0
    Case 11: B1.state=int(rnd(1)*2)
        B2.state=int(rnd(1)*2)
        B3.state=int(rnd(1)*2)
        B4.state=int(rnd(1)*2)
        B5.state=int(rnd(1)*2)
        B6.state=int(rnd(1)*2)
        B7.state=int(rnd(1)*2)
  End Select
End Sub
Sub RedDisplay(nr)
  Select Case nr
    Case 0 : R1.state=1 : R2.state=1 : R3.state=1 : R4.state=1 : R5.state=1 : R6.state=1 : R7.state=0
    Case 1 : R1.state=0 : R2.state=1 : R3.state=1 : R4.state=0 : R5.state=0 : R6.state=0 : R7.state=0
    Case 2 : R1.state=1 : R2.state=1 : R3.state=0 : R4.state=1 : R5.state=1 : R6.state=0 : R7.state=1
    Case 3 : R1.state=1 : R2.state=1 : R3.state=1 : R4.state=1 : R5.state=0 : R6.state=0 : R7.state=1
    Case 4 : R1.state=0 : R2.state=1 : R3.state=1 : R4.state=0 : R5.state=0 : R6.state=1 : R7.state=1
    Case 5 : R1.state=1 : R2.state=0 : R3.state=1 : R4.state=1 : R5.state=0 : R6.state=1 : R7.state=1
    Case 6 : R1.state=0 : R2.state=0 : R3.state=1 : R4.state=1 : R5.state=1 : R6.state=1 : R7.state=1
    Case 7 : R1.state=1 : R2.state=1 : R3.state=1 : R4.state=0 : R5.state=0 : R6.state=0 : R7.state=0
    Case 8 : R1.state=1 : R2.state=1 : R3.state=1 : R4.state=1 : R5.state=1 : R6.state=1 : R7.state=1
    Case 9 : R1.state=1 : R2.state=1 : R3.state=1 : R4.state=0 : R5.state=0 : R6.state=1 : R7.state=1
    Case 10: R1.state=0 : R2.state=0 : R3.state=0 : R4.state=0 : R5.state=0 : R6.state=0 : R7.state=0
    Case 11: R1.state=int(rnd(1)*2)
        R2.state=int(rnd(1)*2)
        R3.state=int(rnd(1)*2)
        R4.state=int(rnd(1)*2)
        R5.state=int(rnd(1)*2)
        R6.state=int(rnd(1)*2)
        R7.state=int(rnd(1)*2)
  End Select
End Sub




'**********************************************************************************************************
'* attrackt repeater *
'**********************************************************************************************************

'   If FlexDMD then UMainDMD.ModifyScene00 "attract1", "TIMED EVENT", "10 RAMPS 99S"
Dim Attrackt1X
Sub Attract1_Timer   '200ms
  If Attract1.enabled=0 Then
    Attract1.enabled=1
  Else
    If startgame=1 Then Attract1.enabled=0
    If not UMainDMD.isrendering Then UmainDMD.DisplayScene00ExWithId "attract1",true, "introvideo.wmv", "  ", 15, 5, "  ", 15, 0, 2, 85000, 14

    Attrackt1X=Attrackt1X+1

    Select Case Attrackt1X

      Case  20  : UMainDMD.ModifyScene00 "attract1", "Unreal", "Tournament"
      Case  40  : UMainDMD.ModifyScene00 "attract1", " ", " "

      Case  45  : UMainDMD.ModifyScene00 "attract1", "Last Score", Formatscore(Lastscore)
      Case  60  : UMainDMD.ModifyScene00 "attract1", " ", " "

      Case  65  : UMainDMD.ModifyScene00 "attract1", "Credits:" , Credits   '  to start here Set Attrackt1X=case-1
      Case  80  : UMainDMD.ModifyScene00 "attract1", " ", " "

      Case  85  : If Credits<1 Then
              UMainDMD.ModifyScene00 "attract1", "Insert", "Coins"
            Else
              UMainDMD.ModifyScene00 "attract1", "Press Start", "To Play !"
            End If
      Case 100  : UMainDMD.ModifyScene00 "attract1", " ", " "


      Case 105  : UMainDMD.ModifyScene00 "attract1", "HISCORE 1 by "& highscorename(1), FormatScore(hScore(1))
      Case 120  : UMainDMD.ModifyScene00 "attract1", " ", " "

      Case 125  : UMainDMD.ModifyScene00 "attract1", "HISCORE 2 by "& highscorename(2), FormatScore(hScore(2))
      Case 140  : UMainDMD.ModifyScene00 "attract1", " ", " "

      Case 145  : UMainDMD.ModifyScene00 "attract1", "HISCORE 3 by "& highscorename(3), FormatScore(hScore(3))
      Case 160  : UMainDMD.ModifyScene00 "attract1", " ", " "

      Case 165  : If TodaysBest=0 Then Exit Sub
            UMainDMD.ModifyScene00 "attract1", "Todays Best ",FormatScore(TodaysBest)
        Case 180  : UMainDMD.ModifyScene00 "attract1", " ", " "

      Case 185  : UMainDMD.ModifyScene00 "attract1", "Games Playd ",GamesPlayd
      Case 200  : UMainDMD.ModifyScene00 "attract1", " ", " "
            Attrackt1X = 0
    End Select
  End If
End Sub




'**********************************************************************************************************
'* Start Next Player or Next Ball *
'**********************************************************************************************************

Dim playchamp
Sub GameOVER
  countdown30.enabled=0
  Finfo=0 : RFinfo=0
  playchamp=0
  If SC(1,27)>2 Then playchamp=1
  If SC(2,27)>2 Then playchamp=1
  If SC(3,27)>2 Then playchamp=1
  If SC(4,27)>2 Then playchamp=1
  If playchamp=1 Then
    endmusic
    Playmusic "Unrealtournament/UT-Champions.mp3"
    MusicVolume=Music_Volume
  Else
    EndMusic
    Playmusic "Unrealtournament/UT-Menu.mp3"
    MusicVolume=Music_Volume

  End If
  for i = 0 to 200
    SLS(i,1)=0
  Next
  SLS(100,1)=1 ' plungerRed always on
  for i = 179 to 190

    SLS(i,1)=1    ' TURNgi ON
    SLS(i,2)=333  ' blink GI alittle
  Next

  startgame=2
  leftflipper.rotatetostart
  rightflipper.rotatetostart

  FlexFlashText2scl "GAME", "OVER" , 70
  LockBolt1.Z=90
  LockBolt1.collidable=False
  debug.Print "Boltmoved DOWN"
  GameOverTimer_Timer
  If SC(1,1)>0 And osbactive = 1 Then SubmitOSBScore SC(1,1)
  If SC(2,1)>0 And osbactive = 1 Then SubmitOSBScore SC(2,1)
  If SC(3,1)>0 And osbactive = 1 Then SubmitOSBScore SC(3,1)
  If SC(4,1)>0 And osbactive = 1 Then SubmitOSBScore SC(4,1)
End Sub

'*orbital scoreboard
'****************************
' POST SCORES
'****************************



' Dim osbtempscore:osbtempscore = 0
  Sub SubmitOSBScore(osbtempscore)
    If osbkey="" Then Exit Sub
'   pupDMDDisplay "-","Uploading Your^Score to OSB",dmdnote,3,0,10
    On Error Resume Next
    if osbactive = 1 Then
      Dim objXmlHttpMain, Url, strJSONToSend

      Url = "https://hook.integromat.com/82bu988v9grj31vxjklh2e4s6h97rnu0"

      strJSONToSend = "{""auth"":""" & osbkey &""",""player id"": """ & osbid & """,""player initials"": """ & osbdefinit &""",""score"": " & CStr(osbtempscore) & ",""table"":"""& cGameName & """,""version"":""" & OSDversion & """}"

      Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP")
      objXmlHttpMain.open "PUT",Url, False
      objXmlHttpMain.setRequestHeader "Content-Type", "application/json"
      objXmlHttpMain.setRequestHeader "application", "application/json"

      objXmlHttpMain.send strJSONToSend
    end if
  End Sub



'****************************
'*GAMEOVER TIMER*
'****************************

Dim checkHS : checkHS=9  ' must be 9 to alow startbutton to work for first game
Dim HSname, TodaysBest

Sub GameOverTimer_Timer   ' 500ms

  If GameOverTimer.enabled=0 Then
    GameOverTimer.enabled=1
    checkHS=0
    Lastscore=SC(MaxPlayers,1)
    If SC(1,1)>TodaysBest Then TodaysBest=SC(1,1)
    If SC(2,1)>TodaysBest Then TodaysBest=SC(2,1)
    If SC(3,1)>TodaysBest Then TodaysBest=SC(3,1)
    If SC(4,1)>TodaysBest Then TodaysBest=SC(4,1)

  Else

    If checkHS=0 Then  ' player 1
      LockBolt1.Z=90
      LockBolt1.collidable=False
      Locked1=0
      Locked2=0
      debug.Print "Boltmoved DOWN"
      If SC(1,1)> HScore(1) Then
        HSNAME=1
        hScore(3)=hScore(2) : highscorename(3)=highscorename(2)
        hScore(2)=hScore(1) : highscorename(2)=highscorename(1)
        hScore(1)=SC(1,1)
        lastscore=SC(1,1)
        InitialPlayer=1
        FlexInitInit
        checkHS=1 '  stop everything until entered name is done
        Exit Sub
      Elseif SC(1,1)>hScore(2) Then
        HSNAME=2
        hScore(3)=hScore(2) : highscorename(3)=highscorename(2)
        hScore(2)=SC(1,1)
        lastscore=SC(1,1)
        InitialPlayer=1
        FlexInitInit
        checkHS=1 '  stop everything until entered name is done
        Exit Sub

      Elseif SC(1,1)>hScore(3) Then
        HSNAME=3
        hScore(3)=SC(1,1)
        lastscore=SC(1,1)
        InitialPlayer=1
        FlexInitInit
        checkHS=1 '  stop everything until entered name is done
        Exit Sub
      End If
      checkHS=2 ' next player
    End If

    If checkHS=2 Then  ' player 2
      If SC(2,1)> HScore(1) Then
        HSNAME=1
        hScore(3)=hScore(2) : highscorename(3)=highscorename(2)
        hScore(2)=hScore(1) : highscorename(2)=highscorename(1)
        hScore(1)=SC(2,1)
        lastscore=SC(2,1)
        InitialPlayer=2
        FlexInitInit
        checkHS=3 '  stop everything until entered name is done
        Exit Sub
      Elseif SC(2,1)>hScore(2) Then
        HSNAME=2
        hScore(3)=hScore(2) : highscorename(3)=highscorename(2)
        hScore(2)=SC(2,1)
        lastscore=SC(2,1)
        InitialPlayer=2
        FlexInitInit
        checkHS=3 '  stop everything until entered name is done
        Exit Sub
      Elseif SC(2,1)>hScore(3) Then
        HSNAME=3
        hScore(3)=SC(2,1)
        lastscore=SC(2,1)
        InitialPlayer=2
        FlexInitInit
        checkHS=3 '  stop everything until entered name is done
        Exit Sub
      End If
      checkHS=4 ' next player
    End If

    If checkHS=4 Then  ' player 2
      If SC(3,1)> HScore(1) Then
        HSNAME=1
        hScore(3)=hScore(2) : highscorename(3)=highscorename(2)
        hScore(2)=hScore(1) : highscorename(2)=highscorename(1)
        hScore(1)=SC(3,1)
        lastscore=SC(3,1)
        InitialPlayer=3
        FlexInitInit
        checkHS=5 '  stop everything until entered name is done
        Exit Sub
      Elseif SC(3,1)>hScore(2) Then
        HSNAME=2
        hScore(3)=hScore(2) : highscorename(3)=highscorename(2)
        hScore(2)=SC(3,1)
        InitialPlayer=3
        lastscore=SC(3,1)
        FlexInitInit
        checkHS=5 '  stop everything until entered name is done
        Exit Sub
      Elseif SC(3,1)>hScore(3) Then
        HSNAME=3
        hScore(3)=SC(3,1)
        lastscore=SC(3,1)
        InitialPlayer=3
        FlexInitInit
        checkHS=5 '  stop everything until entered name is done
        Exit Sub
      End If
      checkHS=6 ' next player
    End If

    If checkHS=6 Then  ' player 2
      If SC(4,1)> HScore(1) Then
        HSNAME=1
        hScore(3)=hScore(2) : highscorename(3)=highscorename(2)
        hScore(2)=hScore(1) : highscorename(2)=highscorename(1)
        hScore(1)=SC(4,1)
        lastscore=SC(4,1)
        InitialPlayer=4
        FlexInitInit
        checkHS=7 '  stop everything until entered name is done
        Exit Sub
      Elseif SC(4,1)>hScore(2) Then
        HSNAME=2
        hScore(3)=hScore(2) : highscorename(3)=highscorename(2)
        hScore(2)=SC(4,1)
        lastscore=SC(4,1)
        InitialPlayer=4
        FlexInitInit
        checkHS=7 '  stop everything until entered name is done
        Exit Sub
      Elseif SC(4,1)>hScore(3) Then
        HSNAME=3
        hScore(3)=SC(4,1)
        lastscore=SC(4,1)
        InitialPlayer=4
        FlexInitInit
        checkHS=7 '  stop everything until entered name is done
        Exit Sub
      End If
      checkHS=8 ' next player
      exit Sub
    End If

    If checkhS=8 Then
      for i = 0 to 50
        SC(1,i)=0
        SC(2,i)=0
        SC(3,i)=0
        SC(4,i)=0
      Next' reset all players
      checkhS=9
      If UMainDMD.isrendering Then UMainDMD.cancelrendering
      UmainDMD.DisplayScene00ExWithId "attract1",true, "introvideo.wmv", "  ", 15, 5, "  ", 15, 0, 2, 85000, 14
      Attract1_Timer
      Exit Sub

    End If

    If checkHS=9 Then checkHS=10 : Exit Sub
    If checkHS=10 Then checkHS=11 :  Exit Sub
    If checkHS=11 Then checkHS=12 : Startgame=3 : Exit Sub



    If checkhS=12 Then
      Locked1=0 : Locked2=0
      If startgame=1 Then
        GameOverTimer.enabled=0
        If UMainDMD.isrendering Then UMainDMD.cancelrendering
      End If
    End If

  End If
End Sub

Sub StartNewPlayer
  TrippleHS=0
  DoubleHS=0

  Lastcap=0
  Lastcap2=0
  lastcap3=0


  LFinfo=0 : RFinfo=0
  Debug.Print "StartNewPlayer"
  pwnage=0
  Invisiblilty(1)=0
  Invisiblilty(2)=0
  Invisiblilty(3)=0
  Invisiblilty(4)=0
  Invisiblilty(5)=0


  If Extraball=2 Then
    ExtraBall=3 ' prevent getting another on same BIP
    FlexFlashText3 "EXTRA", "LIFE" , 70
    SLS(85,1)=0
    SPB1 85,85,4,1,1,1
    RESETFORNEWBALL
    SetTeamColors
    SLS(73,1)=1
    SLS(74,1)=1
    SPB1 71,74,10,2,0,1
    RedDisplay SC(PN,28)
    BlueDisplay SC(PN,29)
    If L063.timerenabled=1  Then
      L063.timerenabled=0
      FlagOn10=0
      L063.timerenabled=1
    End If
    Exit Sub
  End If



  Extraball=0
  PN=PN+1

  IF PN>MAXplayers Then
    PN=1
    Ballinplay=Ballinplay+1
    If Ballinplay=4 then GAMEOVER : exit Sub
  End If

  If Ballinplay=1 Then
    Firstblood=0 ' ONCE FOR ALL PLAYERS
  Else
    Firstblood=1
  End If


  RESETFORNEWBALL
  SetTeamColors
  SLS(73,1)=1
  SLS(74,1)=1
  SPB1 71,74,10,2,0,1
  RedDisplay SC(PN,28)
  BlueDisplay SC(PN,29)

  If SC(PN,27)=0 Then ' 0 = level1
    SLS( 7,1)=1  ' handicap on level 1, make sure those are on
    SLS(15,1)=1
    SLS(20,1)=1
    SLS(25,1)=1
    SLS(30,1)=1
    IF SC(PN,12)=0 Then SC(PN,12)=1
    IF SC(PN,13)=0 Then SC(PN,13)=1
    IF SC(PN,14)=0 Then SC(PN,14)=1
    IF SC(PN,15)=0 Then SC(PN,15)=1
    IF SC(PN,16)=0 Then SC(PN,16)=1

  End If
  If  SLS(41,1)>0 Then
    L063.timerenabled=0
    FlagOn10=0
    L063.timerenabled=1
  End If

End Sub

Sub RESETFORNEWBALL
  lostballs=0
  Lightsout
  SPB1 1,178,3,7,0,1


  L035.timerenabled=0
  L036.timerenabled=0
  L037.timerenabled=0
  L038.timerenabled=0
  L039.timerenabled=0
  L040.timerenabled=0
  Sw25.isdropped=False
  Sw26.isdropped=False
  Sw27.isdropped=False
  Sw33.isdropped=False
  Sw34.isdropped=False
  Sw35.isdropped=False

  L049.timerenabled=0
  KBStatus=1
  debug.print "kickback ON"
  SLS(49,1)=1 ' KB ON
  LiteKBstatus=0
  SLS(47,1)=0
'L047.timerenabled=0
    '  restore plyers Lights

    SC(PN,2)=0 ' bonus Reset
    SC(PN,3)=0
    SC(PN,4)=0
    SC(PN,11)=0
    SC(PN,4)=0  'DT Reset
    SC(PN,5)=0
    SC(PN,6)=0
    SC(PN,7)=0
    SC(PN,8)=0
    SC(PN,9)=0
    SC(PN,10)=0

    If SC(PN,13)>0 Then SLS( 7,1)=1'  12  Rightorbit
    If SC(PN,13)>1 Then SLS( 8,1)=1
    If SC(PN,13)>2 Then SLS( 9,1)=1
    If SC(PN,13)>3 Then SLS(10,1)=1
    If SC(PN,13)>4 Then SLS(11,1)=1

    If SC(PN,12)>0 Then SLS(15,1)=1'  13  Leftorbit nr(1-5=nr of lights lit) + more states ?
    If SC(PN,12)>1 Then SLS(16,1)=1
    If SC(PN,12)>2 Then SLS(17,1)=1
    If SC(PN,12)>3 Then SLS(18,1)=1
    If SC(PN,12)>4 Then SLS(19,1)=1

    If SC(PN,14)>0 Then SLS(20,1)=1'  14  kicker1
    If SC(PN,14)>1 Then SLS(21,1)=1
    If SC(PN,14)>2 Then SLS(22,1)=1
    If SC(PN,14)>3 Then SLS(23,1)=1
    If SC(PN,14)>4 Then SLS(24,1)=1

    If SC(PN,15)>0 Then SLS(25,1)=1'  15  kicker2
    If SC(PN,15)>1 Then SLS(26,1)=1
    If SC(PN,15)>2 Then SLS(27,1)=1
    If SC(PN,15)>3 Then SLS(28,1)=1
    If SC(PN,15)>4 Then SLS(29,1)=1

    If SC(PN,16)>0 Then SLS(30,1)=1'  16  kicker3
    If SC(PN,16)>1 Then SLS(31,1)=1
    If SC(PN,16)>2 Then SLS(32,1)=1
    If SC(PN,16)>3 Then SLS(33,1)=1
    If SC(PN,16)>4 Then SLS(34,1)=1

'   SC(PN,18)=0
    SC(PN,19)=0

    SLS(65,1)=SC(PN,22)
    SLS(66,1)=SC(PN,22) ' CP

    SLS(43,1)=SC(PN,24) ' EB
    SLS(41,1)=SC(PN,25) ' EF
    SLS(70,1)=SC(PN,25)
    SLS(48,1)=SC(PN,25)
    SLS(63,1)=SC(PN,25)



    SLS(42,1)=SC(PN,23) ' CAP
    If SLS(41,1)>1 and SC(PN,23)>1 Then SLS(42,1)=1
    SLS(68,1)=SC(PN,26) ' TP



    ' LockBolt  up/down and divider on if light 65 is flashing
    If locked1=1 or Locked2=1 Then
      LockBolt1.Z=140
      LockBolt1.collidable=True
debug.Print "Boltmoved UP"
    Else
      LockBolt1.Z=90
      LockBolt1.collidable=False
debug.Print "Boltmoved DOWN"
    End If
    If SLS(65,1)=2 Then
      LockBolt1.Z=140
      LockBolt1.collidable=True
debug.Print "Boltmoved UP"

      DividerWall1_Timer ' open if light is on
    Else
      DividerWall0_Timer ' must close this
    End If


    MusicVolume=Music_Volume*0.7
    Superbumpers=0
    SC(PN,11)=0
    FFlevel=0
    KSlevel=0
    LostBalls=0
    MBActive=0
    Ownage=0

    PlaySound "Prepare", 1, 0.3 * BG_Volume, 0, 0,0,0, 0, 0 '  .... red leader and red balls that sor of thing, according to operation manual ?
    If UMainDMD.isrendering Then UMainDMD.cancelrendering
    FlexFlashText3 "PLAYER", PN , 140

    AutoKick=0  ' wait for plunger key
    RespawnWaiting=RespawnWaiting+1
      DOF 151,2
      If DOFdebug=1 Then debug.print "DOF 151,2 ballrelease"
    L062.timerenabled=0
    L062_Timer
    L063_Timer
End Sub

'**********************************************************************************************************
'* Lightsout *
'**********************************************************************************************************

Sub Lightsout

  for i = 0 to 200
      SLS(i,1)=0
      SLS(i,2)=0
      SLS(i,3)=0
      SLS(i,4)=0
      SLS(i,5)=0
      SLS(i,6)=0
      SLS(i,7)=0
      SLS(i,8)=0
      SLS(i,9)=0
      SLS(i,10)=0
      SLS(i,11)=0
      SLS(i,12)=0
  Next ' TURNoff everything
  SLS(100,1)=1 ' plungerRed always on
  for i = 179 to 190

    SLS(i,1)=1    ' TURNgi ON
    SLS(i,2)=123  ' blink for 123 updates = 1 solid blink ``?
  Next
  PlaySound "fx_relay", 1,0.3*VolumeDial,0,0,0,0,0,0

  kickback.PULLBACK
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, kickback
  L049.timerenabled=0
  KBStatus=1
debug.print "kickback ON"

  SLS(49,1)=1 ' KB ON
  LiteKBstatus=0
  SLS(47,1)=0
'L047.timerenabled=0
  If StaticScoreboard=1 Then
    if sc(PN,27)=0 then UMainDMD.SetScoreboardBackgroundImage "sb0.png", 15, 15
    if sc(PN,27)=1 then UMainDMD.SetScoreboardBackgroundImage "sb1.png", 15, 15
    if sc(PN,27)=2 then UMainDMD.SetScoreboardBackgroundImage "sb2.png", 15, 15
    if sc(PN,27)=3 then UMainDMD.SetScoreboardBackgroundImage "sb3.png", 15, 15
    SBtimer=0
  Else

    If SC(PN,27)=0 Then UMainDMD.SetScoreboardBackgroundImage "Coret.wmv", 15, 15 : SBtimer=196
    If SC(PN,27)=1 Then UMainDMD.SetScoreboardBackgroundImage "November.wmv", 15, 15 : SBtimer=220
    If SC(PN,27)=2 Then UMainDMD.SetScoreboardBackgroundImage "Dreary.wmv", 15, 15 : SBtimer=228
    If SC(PN,27)=3 Then UMainDMD.SetScoreboardBackgroundImage "Face.wmv", 15, 15 : SBtimer=234

  End If

End Sub



Dim SBtimer
Sub SBrestartVid_timer
  If StaticScoreboard=0 Then

    If SBtimer>1 Then SBtimer=SBtimer-1
    If SBtimer=1 Then

      If SC(PN,27)=0 Then UMainDMD.SetScoreboardBackgroundImage "Coret.wmv", 15, 15 : SBtimer=196
      If SC(PN,27)=1 Then UMainDMD.SetScoreboardBackgroundImage "November.wmv", 15, 15 : SBtimer=220
      If SC(PN,27)=2 Then UMainDMD.SetScoreboardBackgroundImage "Dreary.wmv", 15, 15 : SBtimer=228
      If SC(PN,27)=3 Then UMainDMD.SetScoreboardBackgroundImage "Face.wmv", 15, 15 : SBtimer=234
    End If
  End If

End Sub





'**********************************************************************************************************
'* GAMESTART *
'**********************************************************************************************************

SUB GAMESTARTER
  TrippleHS=0
  DoubleHS=0
  Lastcap=0
  Lastcap2=0
  lastcap3=0

  StopAddingPlayers=0

  locked1=0
  locked2=0
  pwnage=0
  Invisiblilty(1)=0
  Invisiblilty(2)=0
  Invisiblilty(3)=0
  Invisiblilty(4)=0
  Invisiblilty(5)=0
  If DB2S_on Then B2STimer2_Timer  ' or b2stimer2 for big one
  DividerWall0_Timer ' be sure to turn it to normal up there
  LockBolt1.Z=90
  LockBolt1.collidable=False
  debug.Print "Boltmoved down"

  initialmusic=1
  Startgame=1
  BallInPlay=1
  LightsOut
  lostballs=0
  RespawnWaiting=RespawnWaiting+1
  DOF 151,2
  If DOFdebug=1 Then debug.print "DOF 151,2 ballrelease"
  If UMainDMD.isrendering Then UMainDMD.cancelrendering

  PlaySound "Prepare", 1, 0.3 * BG_Volume, 0, 0,0,0, 0, 0 '  .... red leader and red balls that sor of thing, according to operation manual ?
  FlexFlashText2 "PREPARE", "FOR BATTLE" , 110
  MusicOn  ' change music


  AutoKick=0  ' wait for plunger key
  L062.timerenabled=0
  L062_Timer
  If UMainDMD.isrendering Then UMainDMD.cancelrendering

  'Player1
  PN=1
    MaxPLayers=1

  SetTeamColors
  SLS(73,1)=1
  SLS(74,1)=1
  SPB1 71,74,5,2,0,1

  RedDisplay 0
  BlueDisplay 0

  SC(1,32)=5 ' GETFLAGBACK LIMIT SET ON ALL PLAYERS
  SC(2,32)=5
  SC(3,32)=5
  SC(4,32)=5

  SLS( 7,1)=1  ' handicap on level 1
  SLS(15,1)=1
  SLS(20,1)=1
  SLS(25,1)=1
  SLS(30,1)=1

  SC(PN,12)=1
  SC(PN,13)=1
  SC(PN,14)=1
  SC(PN,15)=1
  SC(PN,16)=1


  SC(1,20) = Taunt(1)
  SC(2,20) = Taunt(2)
  SC(3,20) = Taunt(3)
  SC(4,20) = Taunt(4)


  If StaticScoreboard=1 Then
    UMainDMD.SetScoreboardBackgroundImage "sb0.png", 15, 15
    SBtimer=0
  Else
    UMainDMD.SetScoreboardBackgroundImage "coret.wmv", 15, 15 : SBtimer=196
  End If

End Sub



'**********************************************************************************************************
'* BONUS TIME AT BALL LOST *
'**********************************************************************************************************

Dim BonusCounter,BonusX
Sub BonusTime_Timer ' 60ms
  If BonusTime.enabled=0 Then
    If  SLS(41,1)>0 Then
      L063.timerenabled=0
      FlagOn10=0
      L063.timerenabled=1 ' keeping this alive but resetting it
    End If

    BonusTime.Enabled=1
    PlaySound "UT_TitleB", 1, 0.7 * Music_Volume, 0, 0,0,0, 0, 0
    MusicVolume=0.05
    If SC(PN,11)=1 Then  SC(PN,0)=SC(PN,0)*2
    If SC(PN,11)=2 Then  SC(PN,0)=SC(PN,0)*3
    If SC(PN,11)=3 Then  SC(PN,0)=SC(PN,0)*4
    If SC(PN,11)=4 Then  SC(PN,0)=SC(PN,0)*6
    If SC(PN,11)>4 Then  SC(PN,0)=SC(PN,0)*8
    debug.print "BONUS" & SC(PN,0)

    FlexFlashText3 "BONUS", (" " & FormatScore(SC(PN,0)) & " ") , 68
    BonusCounter=0
    LFinfo=0 : RFinfo=0
  Else
    If countdown30.enabled=1 Then
      SC(PN,33)=lastcd+1000 ' save countdown timer and stopping It
      countdown30.enabled=0
    End If
    BonusCounter=BonusCounter+1
    If BonusCounter=60 Then PlaySound "hum4", 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0
    If BonusCounter<60 Then exit Sub
    If SC(PN,0)>1000000 Then SC(PN,0)=SC(PN,0)-1000000 : SC(PN,1)=SC(PN,1)+1000000 : exit Sub
    If SC(PN,0)>100000 Then SC(PN,0)=SC(PN,0)-100000 : SC(PN,1)=SC(PN,1)+100000 : exit Sub
    If SC(PN,0)>20000  Then SC(PN,0)=SC(PN,0)-10000  : SC(PN,1)=SC(PN,1)+10000  : exit Sub
    If SC(PN,0)>1000   Then SC(PN,0)=SC(PN,0)-1000   : SC(PN,1)=SC(PN,1)+1000   : exit Sub
    If SC(PN,0)>200    Then SC(PN,0)=SC(PN,0)-100    : SC(PN,1)=SC(PN,1)+100    : exit Sub
    If SC(PN,0)>10     Then SC(PN,0)=SC(PN,0)-10     : SC(PN,1)=SC(PN,1)+10     : exit Sub
    If SC(PN,0)>0      Then SC(PN,0)=SC(PN,0)-1      : SC(PN,1)=SC(PN,1)+1      : exit Sub
    SC(PN,0)=0
    BonusTime.enabled=0
    BonusEndTimer.enabled=1
  End If
End Sub

Sub BonusEndTimer_timer   '1.2 sec
  BonusEndTimer.enabled=0
  SavePlayerLights
  StartNewPlayer
End Sub


'**********************************************************************************************************
'* SKILLSHOT *
'**********************************************************************************************************

Dim Skill1 , Skill2, SSLevel
Sub L062_Timer   ' 80ms

  If L062.timerenabled=0 Then ' skillshot level 1
    L062.timerenabled=1
    Skill1=0
    Skill2=0
    SLS(59,5)=2 ' off=2 on SP
    SLS(60,5)=2
    SLS(61,5)=2
    SLS(62,5)=2
    SLS(56,1)=2
    Debug.print"SSON"
    If SC(PN,27)=0 Then L062.TimerInterval=83
    If SC(PN,27)=1 Then L062.TimerInterval=77 ' up speed
    If SC(PN,27)=2 Then L062.TimerInterval=71 ' even faster on last maps
    If SC(PN,27)=3 Then L062.TimerInterval=65

  Else
'   If SC(PN,27)=0 Then  ' map playing
'     Skill1=Skill1+1
'
'     Select Case Skill1
'       Case  1 : SLS(59,5)=1 : SLS(62,5)=2 :
'       Case  7 : SLS(60,5)=1 : SLS(59,5)=2 :
'       Case 13 : SLS(61,5)=1 : SLS(60,5)=2 :
'       Case 19 : SLS(62,5)=1 : SLS(61,5)=2 :
'       Case 24 : Skill1=0
'     End Select
'   End If
'
'   If SC(PN,27)>0 Then ' atm for map2+3+4   goes faster
      Skill1=skill1+1
      Select Case Skill1
        Case  1 : SLS(59,5)=1 : SLS(62,5)=2 :
        Case  2 : Skill1=3
        Case  7 : SLS(60,5)=1 : SLS(59,5)=2
        Case  8 : Skill1=9
        Case 13 : SLS(61,5)=1 : SLS(60,5)=2
        Case 14 : skill1=15
        Case 19 : SLS(62,5)=1 : SLS(61,5)=2
        Case 23 : Skill1=0
      End Select

'   End If
  End If

End Sub


Sub L061_Timer
  SkillshotOFF
End Sub


Sub SkillshotOFF
  SLS(56,1)=0
  Skill1=0
  L062.timerenabled=0
  L061.timerenabled=0
  SLS(59,5)=0
  SLS(60,5)=0
  SLS(61,5)=0
  SLS(62,5)=0
Debug.print"SSOFF"
End Sub


Sub Skillshot2

  Scoring 50000,5000
  FFT=FFTscore
  PlaySound "AANali" & Int(rnd(1)*6)+1 , 1, 1.2*BG_Volume, 0, 0,0,0, 0, 0
  PlaySound "mach15", 1, 0.3 * BG_Volume, 0, 0,0,0, 0, 0
  FlexFlashText1 "PERFECT2",88
End Sub


Sub SkillShotPerfect
  DOF 124,2
  If DOFdebug=1 Then debug.print "DOF 124,2 skilshot perfect"
  SkillshotOFF
  Scoring 20000,2000
  FFT=FFTscore
  PlaySound "AANali" & Int(rnd(1)*6)+1 , 1, 1.2*BG_Volume, 0, 0,0,0, 0, 0
  PlaySound "mach15", 1, 0.3 * BG_Volume, 0, 0,0,0, 0, 0
  FlexFlashText1 "PERFECT",88
Debug.print"PerfectHit"
  SPB1 44,45,35,2,9,1
  SPB1 66,67,35,2,9,1
  L059.timerenabled=1
  SS2=1
End Sub


DIM SS2
Sub L059_Timer '  2nd skillshot timer
  SS2=0
  L059.timerenabled=0
End Sub


'123 skillshot
'124 skillshot perfect
'125 skillshot 2
Sub SkillShotGood
  DOF 123,2
  If DOFdebug=1 Then debug.print "DOF 123,2 skilshotgood"
  SkillshotOFF
  Scoring 10000,1000
  FFT=FFTscore
  PlaySound "mach15", 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0
  FlexFlashText1 "SKILLSHOT",88
  Debug.print"DirectHit"
  SPB1 44,45,35,2,5,1
  SPB1 66,67,35,2,5,1
  L059.timerenabled=1
  SS2=1

End Sub



Sub SkillshotClose
  SkillshotOFF
  Scoring 3000,300
  FlexFlashText1 "Almost",88
Debug.print"AlmostSS"

End Sub



Sub SkillshotBad
  SkillshotOFF
  Scoring 500,50
  FlexFlashText1 "Missed",88
Debug.print"MissedSS"

End Sub




'**********************************************************************************************************
'* FLEX INITIALS *
'**********************************************************************************************************


Dim EnterCurrent ' the letter flippers changes
Dim EnteredName ' what has been entered
Dim InitialPlayer '  nr of player entering name
Dim EnterBlink

Sub FlexInitInit
  EnteredName=""
  EnterCurrent=1
  FlexEnterName=1
End Sub


Sub FlexInitials
  tempstr2="ABCDEFGHIJKLMNOPQRSTUVWXYZ-0123456789"
  tempstr="   PLAYER " & InitialPlayer & " NAME = " & EnteredName    ' & EnterCurrent
  EnterBlink=EnterBlink+1
  If EnterBlink=4 Then EnterBlink=0

  If enterblink <2 Then
    UMainDMD.DisplayScoreboard 4, InitialPlayer , SC(1,1) , SC(2,1) , SC(3,1) , SC(4,1) , tempstr, " "
  Else
    UMainDMD.DisplayScoreboard 4, InitialPlayer , SC(1,1) , SC(2,1) , SC(3,1) , SC(4,1) , tempstr & Mid(tempstr2 ,EnterCurrent,1), " "
  End If

  If len (EnteredName)=3 Then
    FlexEnterName=0
    highscorename(HSname) = EnteredName
    entercurrent=0
    enteredname=""
    checkHS=checkHS+1 ' NEXT
  End If

End Sub




'**********************************************************************************************************
'* FLEX SCORING *
'**********************************************************************************************************

Dim FlexEnterName
Sub FlexTimer_Timer

  If FlexEnterName=1 Then FlexInitials : Exit Sub
  If Startgame>0 And NOT UMainDMD.IsRendering() Then FlexScoring
End Sub


dim FlexCounter,FlexPause,FlexMsg,FlexMsg2
FlexCounter=56  '' set it to off straightaway now
dim ScoreX
Sub FlexScoring

  If FlexCounter>26 And FlexMsg2<>"" Then
    FlexMsg = FlexMsg2
    FlexMsg2=""
    FlexCounter=0
    FlexPause=0
  End If

  If FlexCounter<56 Then
    FlexCounter=FlexCounter+1
'd ebug.print"counter="& FlexCounter & " cp=" & FlexPause
    If FlexCounter=25 and FlexPause=0 Then FlexPause=20
    If FlexPause>1 then FlexPause=FlexPause-1 : FlexCounter=25
    If FlexPause=1 Then FlexCounter = 26 : FlexPause=0
'   UMainDMD.SetScoreboardBackgroundImage "sb" & PN-1 & ".png", 15, 15
    UMainDMD.DisplayScoreboard MaxPlayers, PN , SC(1,1) , SC(2,1) , SC(3,1) , SC(4,1) ,  mid ( FlexMsg , FlexCounter , 26), " "

  Else
    ScoreX=ScoreX+1
    If ScoreX>48 then ScoreX=0
    tempnr2=1
    If SLS(35,1) > 0 Then tempnr2=tempnr2+0.20
    If SLS(36,1) > 0 Then tempnr2=tempnr2+0.70
    If SLS(37,1) > 0 Then tempnr2=tempnr2+0.20
    If SLS(38,1) > 0 Then tempnr2=tempnr2+0.40
    If SLS(39,1) > 0 Then tempnr2=tempnr2+0.40
    If SLS(40,1) > 0 Then tempnr2=tempnr2+0.40

    If tempnr2>1 Then
      If ScoreX<24 Then
        UMainDMD.DisplayScoreboard MaxPlayers, PN , SC(1,1) , SC(2,1) , SC(3,1) , SC(4,1) ,  "Ball " & BallInPlay & "   Score" , " Credits " & Credits
      Else
        UMainDMD.DisplayScoreboard MaxPlayers, PN , SC(1,1) , SC(2,1) , SC(3,1) , SC(4,1) ,  "Ball " & BallInPlay & "   X "& tempnr2 , " Credits " & Credits
      End If
    Else
      UMainDMD.DisplayScoreboard MaxPlayers, PN , SC(1,1) , SC(2,1) , SC(3,1) , SC(4,1) ,  "Ball " & BallInPlay , " Credits " & Credits
    End If
  End If

End Sub


Sub FlexFlashText3( newstr, newstr2 ,Pause )   ' 2 liners
  If DB2S_on Then B2STimer1_Timer  ' or b2stimer2 for big one

  Dim FlexPic
  FlexPic = "sb" & Int(RND(1)*34) & ".png"

  If FFT>0 Then
    UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2,  FormatScore(FFT) , 4, 15, 14, Pause*6, 14
  End If

  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2,  newstr2, 4, 15, 14, Pause*4, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2,    " ", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2, newstr2, 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2,    " ", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2, newstr2, 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5,    " ", 4, 5, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5,  newstr2, 4, 5, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5,    " ", 4, 5, 14, Pause, 14

  If FFT=0 Then
    UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5, newstr2, 4, 15, 14, Pause, 14
    UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5,    " ", 4, 5, 14, Pause, 14
    UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5, newstr2, 4, 15, 14, Pause, 14
  End If
  FFT=0
End Sub



Sub FlexFlashText2( newstr, newstr2 ,Pause )   ' 2 liners
  If DB2S_on Then B2STimer1_Timer  ' or b2stimer2 for big one

  Dim FlexPic
  FlexPic = "sb" & Int(RND(1)*34) & ".png"

  If FFT>0 Then
    UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5,  FormatScore(FFT) , 4, 15, 14, Pause*6, 14
  End If

  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2,  newstr2, 4, 15, 14, Pause*4, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2,    " ", 0, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 2, newstr2, 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2,    " ", 0, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 2, newstr2, 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5,    " ", 0, 5, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 5, newstr2, 4, 5, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5,    " ", 0, 5, 14, Pause, 14

  If FFT=0 Then
    UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 5, newstr2, 4, 15, 14, Pause, 14
    UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5,    " ", 4, 5, 14, Pause, 14
    UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 5, newstr2, 4, 15, 14, Pause, 14
  End If
  FFT=0
End Sub

Sub FlexFlashText2cap( newstr, score ,Pause )   ' 2 liners
  If DB2S_on Then B2STimer1_Timer  ' or b2stimer2 for big one

  Dim FlexPic
  FlexPic = "sb" & Int(RND(1)*34) & ".png"

  tmp1=FormatScore(score)

  For i = 1 to len(newstr)
    UmainDMD.DisplayScene00Ex FlexPic , left(newstr,i)+ space(len(newstr)-i), 15, 2, tmp1, 4, 15, 14, Pause/3, 14
    UmainDMD.DisplayScene00Ex FlexPic , " ", 15, 0, tmp1 , 4, 15, 14, Pause/3, 14
  Next

  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2,    " ", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 2,    tmp1, 2, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 0,    " ", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 0,    tmp1, 4, 5,  14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 0,     tmp1, 4, 15, 14, Pause*2, 14
  UmainDMD.DisplayScene00Ex FlexPic , " ", 15, 5,         " ", 2, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 0,     tmp1, 4, 15, 14, Pause*2, 14
  FFT=0
End Sub



Sub FlexFlashText2scl( newstr, newstr2 ,Pause )   ' 2 liners
  If DB2S_on Then B2STimer1_Timer  ' or b2stimer2 for big one

  Dim FlexPic
  FlexPic = "sb" & Int(RND(1)*34) & ".png"


  For i = 1 to len(newstr)
    UmainDMD.DisplayScene00Ex FlexPic , left(newstr,i) + space(len(newstr)-i), 15, 2,                   " "    , 4, 15, 14, Pause, 14
    UmainDMD.DisplayScene00Ex FlexPic , " ",                   15, 2, space(len(newstr2)-i) + right(newstr2,i) , 4, 15, 14, Pause, 14
  Next

  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2,    " ", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ",15, 2,newstr2, 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 5, 15,    " ", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ",15, 2,newstr2, 4, 5, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2,    " ", 4, 15,14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,     " ",15, 2,newstr2, 4, 15,14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2,    " ", 4, 15,14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr,15, 0,newstr2, 0, 15,14, Pause*2, 14
  UmainDMD.DisplayScene00Ex FlexPic , " ",15, 0," ", 0, 15,14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr,15, 0,newstr2, 0, 15,14, Pause*2, 14
  FFT=0
End Sub


Function FormatScore(ByVal Num)
    dim NumString
    NumString = CStr(abs(Num) )
    if len(NumString)>9 then NumString = left(NumString, Len(NumString)-9) & "," & right(NumString,9)
    if len(NumString)>6 then NumString = left(NumString, Len(NumString)-6) & "," & right(NumString,6)
    if len(NumString)>3 then NumString = left(NumString, Len(NumString)-3) & "," & right(NumString,3)
  FormatScore = NumString
End function


Sub FlexFlashText1( newstr ,Pause )   ' 1 liners
  If DB2S_on Then B2STimer1_Timer  ' or b2stimer2 for big one

  Dim FlexPic
  FlexPic = "sb" & Int(RND(1)*34) & ".png"

  If FFT>0 Then
    UmainDMD.DisplayScene00Ex FlexPic , FormatScore(FFT) , 15, 5, "" , 4, 15, 14, Pause*6, 14
  End If

  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2, "", 4, 15, 14, Pause*4, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 2, "", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2, "", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 2, "", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2, "", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 2, "", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5, "", 4, 5, 14, Pause, 14

  If FFT=0 Then
    UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 5, "", 4, 5, 14, Pause, 14
    UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5, "", 4, 5, 14, Pause, 14
    UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5, "", 4, 5, 14, Pause, 14
    UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 5, "", 4, 15, 14, Pause, 14
  End If
  FFT=0
End Sub

Sub FlexFlashText1tw( newstr ,Pause ) ' typewriter
  If DB2S_on Then B2STimer1_Timer  ' or b2stimer2 for big one

  Dim FlexPic
  FlexPic = "sb" & Int(RND(1)*34) & ".png"

For i = 1 to len(newstr)
  UmainDMD.DisplayScene00Ex FlexPic , left(newstr,i), 15, 2, "", 4, 15, 14, Pause/2, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 2, "", 4, 15, 14, Pause/2, 14
Next
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 2, "", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 2, "", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 2, "", 4, 15, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5, "", 4, 5, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 5, "", 4, 5, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic , newstr, 15, 5, "", 4, 5, 14, Pause, 14
  UmainDMD.DisplayScene00Ex FlexPic ,   " ", 15, 5, "", 4, 15, 14, Pause, 14
End Sub


Sub FlexNewText( newstr )  ' will keep 2nd msg ( 3rd will overwrite 2nd but most likely be fine !
  FlexMsg2=space(37-len(newstr)/2) & newstr & space(52)
End Sub




'**********************************************************************************************************
'* dtbg animate *
'**********************************************************************************************************
Dim desktopbg
Sub dtbg_timer
  If dtbg.enabled=0 Then
    dtbg.enabled=1
    desktopbg=0

  Else
    desktopbg=desktopbg+1
    Select case desktopbg

      case 1 : Flasher006.imageA="UTLeftBG2" : Flasher007.imageA="UTRightBG2"
      case 2 : Flasher006.imageA="UTLeftBG3" : Flasher007.imageA="UTRightBG3"
      case 7 : Flasher006.imageA="UTLeftBG2" : Flasher007.imageA="UTRightBG2"
      case 8 : Flasher006.imageA="UTLeftBG1" : Flasher007.imageA="UTRightBG1"
           dtbg.enabled=0
    End Select
  End If

End Sub

Dim dtbg2(5)
Sub dtbg001_Timer
  dtbg2(1)=dtbg2(1)+4
  If dtbg2(1)>360 Then dtbg2(1)=dtbg2(1)-360

  dtbg2(2)=dtbg2(2) + sin((dtbg2(1)-180)/360)
  flasher006.height=100 + dtbg2(2)
  flasher007.height=100 - dtbg2(2)
  flasher006.roty=  dtbg2(2)/16
  flasher007.roty=  dtbg2(2)/16

End Sub





'**********************************************************************************************************
'* Skarj animate *
'**********************************************************************************************************

Dim Skarj, skarj2
Sub scarjitimer_Timer
  If scarjitimer.enabled=0 Then
    scarjitimer.Enabled=1
    skarj=0
    skarj2=0

    PlaySound "roam1WL" , 1, BG_Volume, 0, 0,0,0, 0, 0
    PlaySound "mcreak6" , 1, 0.33 * BG_Volume, 0, 0,0,0, 0, 0
  Else
    skarj = skarj + 10
    skarj2 = skarj2 - sin((Skarj-180)/360)
    Primitive012.objrotz=skarj2 * 20
    If skarj=180 Then PlaySound "mcreak7" , 1, 0.33 * BG_Volume, 0, 0,0,0, 0, 0
    If skarj>359 Then scarjitimer.enabled=0
  End If

End Sub


Sub sw49_hit
  scarjitimer_Timer
End Sub




'**********************************************************************************************************
'* Frame Timer *
'**********************************************************************************************************

Sub Frametimer_Timer()  '15ms

  If LSAllUP=1 Then
    LShead=LShead+LSaddH
    If LShead<-300 Then
      LSAllUP=0
    End If
  End If

  If LSAllDOWN=1 Then
    LShead=LShead+LSaddH
    If LShead>2456 Then
      LSAllUP=0
    End If
  End If

  UpdateLights
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
    PrimGate.RotX = Sw49.CurrentAngle + 90
  gate1p.RotX =(gate1.currentangle)

  BallShadowUpdate
  RollingTimer
  RDampen

End Sub

dim testingvar
Sub lighttesting
    testingvar=1
FOR I = 1 TO 166 : SLS(I,1)=1 : NEXT ' all inserts on
  Team(0)=1 : SetTeamColors
End Sub


Dim TESTALLLIGHT
'**********************************************************************************************************
'* WORLD TIMER *
'**********************************************************************************************************
dim i
Sub World_Timer ' 40ms
'SLS(86,1)=2
'SLS(87,1)=2
'SLS(88,1)=2
  If Tilt > 0 Then
    Tilt = Tilt - 1
    If Tilt < 0 Then Tilt = 0
    Debug.print "Tiltrecovery ="&Tilt
  End If

' If testingvar=0 Then lighttesting
  If startgame=0 Then lightshow
  RainingMOD

  i=LiftTop.rotY+2
  If i > 360 then i = i -360
  Lifttop.rotY=i
  Primitive033.roty=i

  If startgame=0 And int(rnd(1)*80)=23 Then dtbg_timer

End Sub




'**********************************************************************************************************
'* WAITING FOR PLUNGER *
'**********************************************************************************************************




Dim longrange
Sub plungertoolong_Timer ' 1250
  If startgame=0 Then plungertoolong_Timer=0
  SPB1 100,100,7,4,0,1
  SPB1 75,77,2,4,10,1
  longrange=longrange+1
  If longrange>6 Then longrange=0
  If longrange=5 Then PlaySound "draw", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
  If longrange=2 Then PlaySound "disabled", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
End Sub


'**********************************************************************************************************
'* Fade all Lights *
'**********************************************************************************************************


Sub InitLightsXY ' lightname + SLS(nr,?)
  SetSLSXY L002,  2
  SetSLSXY L003,  3
  SetSLSXY L004,  4
  SetSLSXY L005,  5
  SetSLSXY L006,  6
  SetSLSXY L007,  7
  SetSLSXY L008,  8
  SetSLSXY L009,  9
  SetSLSXY L010, 10
  SetSLSXY L011, 11
  SetSLSXY L012, 12
  SetSLSXY L013, 13
  SetSLSXY L014, 14
  SetSLSXY L015, 15
  SetSLSXY L016, 16
  SetSLSXY L017, 17
  SetSLSXY L018, 18
  SetSLSXY L019, 19
  SetSLSXY L020, 20
  SetSLSXY L021, 21
  SetSLSXY L022, 22
  SetSLSXY L023, 23
  SetSLSXY L024, 24
  SetSLSXY L025, 25
  SetSLSXY L026, 26
  SetSLSXY L027, 27
  SetSLSXY L028, 28
  SetSLSXY L029, 29
  SetSLSXY L030, 30
  SetSLSXY L031, 31
  SetSLSXY L032, 32
  SetSLSXY L033, 33
  SetSLSXY L034, 34
  SetSLSXY L035, 35
  SetSLSXY L036, 36
  SetSLSXY L037, 37
  SetSLSXY L038, 38
  SetSLSXY L039, 39
  SetSLSXY L040, 40
  SetSLSXY L041, 41
  SetSLSXY L042, 42
  SetSLSXY L043, 43
  SetSLSXY L044, 44
  SetSLSXY L045, 45
  SetSLSXY L046, 46
  SetSLSXY L047, 47
  SetSLSXY L048, 48
  SetSLSXY L049, 49
  SetSLSXY L050, 50
  SetSLSXY L051, 51
  SetSLSXY L052, 52
  SetSLSXY L053, 53
' SetSLSXY L054, 54
' SetSLSXY L055, 55
  SetSLSXY L056, 56
  SetSLSXY L057, 57
  SetSLSXY L058, 58
  SetSLSXY L059, 59
  SetSLSXY L060, 60
  SetSLSXY L061, 61
  SetSLSXY L062, 62
  SetSLSXY L063, 63
  SetSLSXY L064, 64
  SetSLSXY L065, 65
  SetSLSXY L066, 66
  SetSLSXY L067, 67
  SetSLSXY L068, 68
  SetSLSXY L34C, 69
  SetSLSXY L070, 70
  SetSLSXY L071, 71
  SetSLSXY L072, 72
  SetSLSXY L073, 73
  SetSLSXY L074, 74
' SetSLSXY L075, 75
' SetSLSXY L076, 76
' SetSLSXY L077, 77
  SetSLSXY L078, 78
  SetSLSXY L079, 79
  SetSLSXY L080, 80
  SetSLSXY L081, 81
  SetSLSXY L082, 82
  SetSLSXY L083, 83
  SetSLSXY L085, 85
' SetSLSXY LiftTopRedLight, 99
' SetSLSXY PlungerRed,100
  SetSLSXY LO001,101
  SetSLSXY LO002,102
  SetSLSXY LO003,103
  SetSLSXY LO004,104
  SetSLSXY LO005,105
  SetSLSXY LO006,106
  SetSLSXY LO007,107
  SetSLSXY LO008,108
  SetSLSXY LO009,109
  SetSLSXY LO010,110
  SetSLSXY LO011,111
' SetSLSXY Rspwn001,120
' SetSLSXY Rspwn002,121
' SetSLSXY Rspwn003,122
' SetSLSXY Rspwn004,123
' SetSLSXY Rspwn005,124
' SetSLSXY Rspwn006,125
' SetSLSXY Rspwn007,126
' SetSLSXY Rspwn008,127
' SetSLSXY Rspwn009,128
' SetSLSXY Rspwn010,129
' SetSLSXY Rspwn011,130
' SetSLSXY Rspwn012,131
' SetSLSXY Rspwn013,132
' SetSLSXY Rspwn014,133
' SetSLSXY Rspwn015,134
' SetSLSXY Rspwn016,135
' SetSLSXY Rspwn017,136
' SetSLSXY Rspwn018,137
' SetSLSXY Rspwn019,138
' SetSLSXY Rspwn020,139
End Sub

Sub SetSLSXY(object,nr)
  SLS(nr,13)=object.x
  SLS(nr,14)=object.y
End Sub

Sub UpdateLights

  FadeL 2  : SetL L002,5 : SetL LT002,3 : SetP p002,2 : Setp2 p002off,1 ' GREEN bonuslightsbottom
  FadeL 3  : SetL L003,5 : SetL LT003,3 : SetP p003,2 : Setp2 p003off,1
  FadeL 4  : SetL L004,5 : SetL LT004,3 : SetP p004,2 : Setp2 p004off,1
  FadeL 5  : SetL L005,5 : SetL LT005,3 : SetP p005,2 : Setp2 p005off,1
  FadeL 6  : SetL L006,5 : SetL LT006,3 : SetP p006,2 : Setp2 p006off,1

  FadeL 7  : SetL L007,3 : SetL LT007,2 : SetP p007,2 : Setp2 p007off,1 ' BLUE outlane x5
  FadeL 8  : SetL L008,3 : SetL LT008,2 : SetP p008,2 : Setp2 p008off,1
  FadeL 9  : SetL L009,3 : SetL LT009,2 : SetP p009,2 : Setp2 p009off,1
  FadeL 10 : SetL L010,3 : SetL LT010,2 : SetP p010,2 : Setp2 p010off,1
  FadeL 11 : SetL L011,3 : SetL LT011,2 : SetP p011,2 : Setp2 p011off,1

  FadeL 12 : SetL L012,2 : SetL LT012,4 : SetP p012,2 : Setp2 p012off,1 ' Bonus X3 top
  FadeL 13 : SetL L013,2 : SetL LT013,4 : SetP p013,2 : Setp2 p013off,1
  FadeL 14 : SetL L014,2 : SetL LT014,4 : SetP p014,2 : Setp2 p014off,1

  FadeL 15 : SetL L015,5 : SetL LT015,2 : SetP p015,1 : Setp2 p015off,1 ' RED outlane x5    was /3 middle and *4 first
  FadeL 16 : SetL L016,4 : SetL LT016,1 : SetP p016,1 : Setp2 p016off,1
  FadeL 17 : SetL L017,4 : SetL LT017,1 : SetP p017,1 : Setp2 p017off,1
  FadeL 18 : SetL L018,3 : SetL LT018,1 : SetP p018,1 : Setp2 p018off,1
  FadeL 19 : SetL L019,3 : SetL LT019,1 : SetP p019,1 : Setp2 p019off,1

  FadeL 20 : SetL L020,2 : SetL LT020,1 : SetP p020,2 : Setp2 p020off,1 '  kicker 01
  FadeL 21 : SetL L021,2 : SetL LT021,1 : SetP p021,2 : Setp2 p021off,1
  FadeL 22 : SetL L022,2 : SetL LT022,1 : SetP p022,2 : Setp2 p022off,1
  FadeL 23 : SetL L023,2 : SetL LT023,1 : SetP p023,2 : Setp2 p023off,1
  FadeL 24 : SetL L024,2 : SetL LT024,1 : SetP p024,2 : Setp2 p024off,1

  FadeL 25 : SetL L025,2 : SetL LT025,1 : SetP p025,2 : Setp2 p025off,1   '  kicker 02
  FadeL 26 : SetL L026,2 : SetL LT026,1 : SetP p026,2 : Setp2 p026off,1
  FadeL 27 : SetL L027,2 : SetL LT027,1 : SetP p027,2 : Setp2 p027off,1
  FadeL 28 : SetL L028,2 : SetL LT028,1 : SetP p028,2 : Setp2 p028off,1
  FadeL 29 : SetL L029,2 : SetL LT029,1 : SetP p029,2 : Setp2 p029off,1

  FadeL 30 : SetL L030,2 : SetL LT030,1 : SetP p030,2 : Setp2 p030off,1 '  kicker 03
  FadeL 31 : SetL L031,2 : SetL LT031,1 : SetP p031,2 : Setp2 p031off,1
  FadeL 32 : SetL L032,2 : SetL LT032,1 : SetP p032,2 : Setp2 p032off,1
  FadeL 33 : SetL L033,2 : SetL LT033,1 : SetP p033,2 : Setp2 p033off,1
  FadeL 34 : SetL L034,2 : SetL LT034,1 : SetP p034,2 : Setp2 p034off,1

  FadeL 35 : SetL L035,1 : SetL LT035,1 : SetP p035,2 : Setp2 p035off,1   ' DT middle
  FadeL 36 : SetL L036,1 : SetL LT036,1 : SetP p036,2 : Setp2 p036off,1
  FadeL 37 : SetL L037,2 : SetL LT037,1 : SetP p037,2 : Setp2 p037off,1

  FadeL 38 : SetL L038,1 : SetL LT038,1 : SetP p038,1.5 : Setp2 p038off,1   '  DT RIGHT
  FadeL 39 : SetL L039,1 : SetL LT039,1 : SetP p039,1.5 : Setp2 p039off,1
  FadeL 40 : SetL L040,1 : SetL LT040,1 : SetP p040,1.5 : Setp2 p040off,1

  FadeL 41 : SetL L041,2 : SetL LT041,2 : SetP p041,6  : Setp2 p041off,0.2 '  ENEMY FLAG1
  FadeL 42 : SetL L042,1.5:SetL LT042,1.5 : SetP p042,6 : Setp2 p042off,0.3 '  CAPTUREkicker2 main
  FadeL 43 : SetL L043,2 : SetL LT043,3 : SetP p043,2 : Setp2 p043off,2'  EXTRABALLkicker3 main
  FadeL 44 : SetL L044,2 : SetL LT044,3 : SetP p044,3 : Setp2 p044off,2 '  high ground
  FadeL 45 : SetL L045,2 : SetL LT045,3 : SetP p045,3 : Setp2 p045off,2'  jackpot
  FadeL 46 : SetL L046,1.5:SetL LT046,2 : SetP p046,2 : Setp2 p046off,2'  million1
  FadeL 47 : SetL L047,1 : SetL LT047,2 : SetP p047,4 : Setp2 p047off,2'  lite kickback
  FadeL 48 : SetL L048,1 : SetL LT048,1 : SetP p048,6  : Setp2 p048off,0.2 '  million2
  FadeL 49 : SetL L049,1 : SetL LT049,2 : SetP p049,4 : Setp2 p049off,2'  kickback
  FadeL 63 : SetL L063,1 : SetL LT063,1 : SetP p063,6   : Setp2 p063off,0.2 '  million3
  FadeL 64 : SetL L064,3 : SetL LT064,2 : SetP p064,2 : Setp2 p064off,3'  respawn

  FadeL 50 : SetL L050,1 : SetL LT050,2 : SetP p050,3   ' coret       4x CTF maps
  FadeL 51 : SetL L051,1 : SetL LT051,2 : SetP p051,3   ' face
  FadeL 52 : SetL L052,1 : SetL LT052,2 : SetP p052,3   ' november
  FadeL 53 : SetL L053,1 : SetL LT053,2 : SetP p053,3   ' dreary

' FadeL 54 : SetL L054,1 : SetL LT054,1 : SetL LT054b,1
' FadeL 55 : SetL L055,1 : SetL LT055,1 : SetL LT055b,1

  FadeL 56 : SetL L056,4 : Flasher056.opacity=i*2     ' Spot MidL
  FadeL 57 : SetL L057,4 : Flasher057.opacity=i*2     ' Spot Left
  FadeL 58 : SetL L058,4 : Flasher058.opacity=i*2   ' Spot Right

  FadeL 59 : SetL L059,0.6 : SetL LT059,1.5  : SetP p059,1 : Setp2 p059off,1' WHITE skillshots x4
  FadeL 60 : SetL L060,0.6 : SetL LT060,1.5  : SetP p060,1 : Setp2 p060off,1
  FadeL 61 : SetL L061,0.4 : SetL LT061,1.5  : SetP p061,1 : Setp2 p061off,1
  FadeL 62 : SetL L062,0.2 : SetL LT062,1.5  : SetP p062,1 : Setp2 p062off,1

  FadeL 65 : SetL L065,2 : SetL LT065,1  : SetP p065,2 : Setp2 p065off,0.5  ' cntrP1
  FadeL 66 : SetL L066,2 : SetL LT066,1  : SetP p066,2 : Setp2 p066off,0.5  ' cntrP2
  FadeL 67 : SetL L067,1 : SetL LT067,1  : SetP p067,6   : Setp2 p067off,0.2        '  Redeemer
  FadeL 68 : SetL L068,2 : SetL LT068,1  : SetP p068,2 : Setp2 p068off,0.5  ' Teamplay

  FadeL 69 : SetL L34C,4 :  SetL L60C,4 :  SetL L61C,4  '  3 extra bumper Lights , ONLY BLINKING SO ARE STRONG ?

  FadeL 70 : SetL L070,1 : SetL LT070,1  : SetP p070,3   : Setp2 p070off,0.2   'EnemyFlag2

  FadeL 71 : SetL L071,1 : SetL LT071,2  : SetP p071,2   'RedScore
  FadeL 73 : SetL L073,12
  FadeL 72 : SetL L072,1 : SetL LT072,2  : SetP p072,2   'BlueScore
  FadeL 74 : SetL L074,12

  FadeL 75 : SetL L075,1 : F075.opacity=i*27 : p075.blenddisablelighting=i*2   'GunLights Bottom
  FadeL 76 : SetL L076,1 : F076.opacity=i*27 : p076.blenddisablelighting=i*2   'GunLights middle
  FadeL 77 : SetL L077,1 : F077.opacity=i*27 : p077.blenddisablelighting=i*2   'GunLights top


  FadeL 78 :  SetL L078,1 : SetL LT078,0.4  : SetP p078,1    'INVISIBILITY
  FadeL 79 :  SetL L079,1 : SetL LT079,0.4  : SetP p079,1
  FadeL 80 :  SetL L080,1 : SetL LT080,0.4  : SetP p080,1
  FadeL 81 :  SetL L081,1 : SetL LT081,0.4  : SetP p081,1
  FadeL 82 :  SetL L082,1 : SetL LT082,0.4  : SetP p082,1
  FadeL 83 :  SetL L083,2 : SetL LT083,0.8  : SetP p083,2

  FadeL 85 :  SetL L085,0.5 : SetL LT085,0.8  : SetP p085,0.8  : Setp2 p085off,0.5 ' EXTRA LIFE
  FadeL 86 :   SetP2 RedBaseLight001,30 :  If i<1 Then primitive011.material="metal2" else primitive011.material="metal0.8"  ' SetP2 Primitive011,0.05
  FadeL 87 :   SetP2 RedBaseLight002,10 : SetP2 RedBaseLight004,100
  FadeL 88 :   SetP2 RedBaseLight003,10

  FadeL 99 : SetL LiftTopRedLight,10    ' : SetP2 Primitive011,0.05


  FadeL 100 : SetL PlungerRed,2


  FadeL 101 : SetL LO001,4
  FadeL 102 : SetL LO002,4
  FadeL 103 : SetL LO003,4
  FadeL 104 : SetL LO004,4
  FadeL 105 : SetL LO005,4
  FadeL 106 : SetL LO006,4
  FadeL 107 : SetL LO007,4
  FadeL 108 : SetL LO008,4
  FadeL 109 : SetL LO009,4
  FadeL 110 : SetL LO010,4
  FadeL 111 : SetL LO011,4

  FadeL 120 : SetL Rspwn001,2
  FadeL 121 : SetL Rspwn002,2
  FadeL 122 : SetL Rspwn003,2
  FadeL 123 : SetL Rspwn004,2
  FadeL 124 : SetL Rspwn005,2
  FadeL 125 : SetL Rspwn006,2
  FadeL 126 : SetL Rspwn007,2
  FadeL 127 : SetL Rspwn008,2
  FadeL 128 : SetL Rspwn009,2
  FadeL 129 : SetL Rspwn010,2
  FadeL 130 : SetL Rspwn011,2
  FadeL 131 : SetL Rspwn012,2
  FadeL 132 : SetL Rspwn013,2
  FadeL 133 : SetL Rspwn014,2
  FadeL 134 : SetL Rspwn015,2
  FadeL 135 : SetL Rspwn016,2
  FadeL 136 : SetL Rspwn017,2
  FadeL 137 : SetL Rspwn018,2
  FadeL 138 : SetL Rspwn019,2
  FadeL 139 : SetL Rspwn020,2




'GI
  SetGI 190,Gi001 : Gi002.state=i : Gi003.state=i : Gi004.state=i ' laneguide R
  SetGI 189,Gi009 : Gi010.state=i : Gi011.state=i : Gi012.state=i ' laneguide L
  SetGI 188,Gi005 : Gi006.state=i : Gi007.state=i : Gi008.state=i ' Sling R
  SetGI 187,Gi013 : Gi014.state=i : Gi015.state=i : Gi016.state=i ' Sling L

  SetGI 186,Gi021 : Gi022.state=i : Gi023.state=i : Gi024.state=i ' lonlytargetleft
  SetGI 185,Gi025 : Gi026.state=i                 ' leftTargets 's
  SetGI 184,Gi027 : Gi028.state=i : Gi029.state=i : Gi030.state=i :  Gi031.state=i : ' DT right
                                    '  032 extra "flash" when DT down
  SetGI 183,Gi035 : Gi036.state=i : Gi037.state=i : Gi038.state=i ' DT Mid
  SetGI 182,Gi035 : Gi036.state=i : Gi037.state=i : Gi038.state=i ' top left
  SetGI 181,GI039 : GI040.state=i : GI041.state=i : GI042.state=i
    GI043.state=i : GI044.state=i : GI045.state=i : GI046.state=i ' TOP LEFT
  SetGI 180,GI047 : GI048.state=i : GI049.state=i : GI050.state=i :   GI055.state=i
    GI051.state=i : GI052.state=i :  GI053.state=i : GI054.state=i ' top Right
  SetGI 179,GI056 : GI057.state=i : GI058.state=i : GI059.state=i : GI060.state=i : GI061.state=i
    GI062.state=i : GI063.state=i :  GI064.state=i : GI065.state=i : GI066.state=i : GI067.state=i
    GI068.state=i : GI069.state=i
': GI070.state=i : GI071.state=i   '  TOP middle

End Sub



'**********************************************************************************************************
'* Special blinkers *
'**********************************************************************************************************

Sub SPB1( tmp1,tmp5,tmp2,tmp6,tmp3,tmp4)

  tempnr=0 : tempnr2=1
  If tmp1>tmp5 then tempnr2=-1
  for i=tmp1 to tmp5 Step tempnr2
    SLS(i, 5) = 3
    SLS(i, 6) = tmp6
    SLS(i, 8) = tmp2
    SLS(i,11) = tempnr
    SLS(i,12) = tmp4
    tempnr = tempnr + tmp3
  Next

End Sub



Dim LShead, LStail, LSAllUP, LSaddH , LSblinks
Sub SPBallup(tmp2,tmp1,tmp3) '  tail in updates
  If LSALLDOWN=0 Then  ' only if the other is not running

    LSALLUP = 1
    LShead  = 2150
    LSaddH  =-tmp2
    LStail  = tmp1 ' how long light should be on before going out
    LSblinks= tmp3

    for i = 1 to 170
      If sls(i,14)>0 Then SLS(i,5)=2 ' off all that has stored XY
    Next
  End If
End Sub

Dim LSALLDOWN
Sub SPBalldown(tmp2,tmp1,tmp3) '  tail in updates
  If LSALLUP=0 Then  ' only if the other is not running
    LSALLDOWN = 1
    LShead  = 0
    LSaddH  = tmp2
    LStail  = tmp1 ' how long light should be on before going out
    LSblinks= tmp3
    for i = 1 to 170
      If sls(i,14)>0 Then SLS(i,5)=2 ' off all that has stored XY
    Next
  End If
End Sub

Sub SPBredside
  for i = 1 to 130
    If SLS(i,13)>475 Then
      SLS(i,5)=2 'off
    Elseif SLS(i,13)>0 Then
      SLS(i,5)=1 'on
    End If
  Next
  SPBred.enabled=1
End Sub


Sub SPBred_Timer ' 88ms
  SPBred.enabled=0
  for i = 1 to 139
    If SLS(i,13)>0 Then SLS(i,5)=0
  Next
End Sub

Sub SPBblueside
  for i = 1 to 130
    If SLS(i,13)>475 Then
      SLS(i,5)=1 'On
    Elseif SLS(i,13)>0 Then
      SLS(i,5)=0 'off
    End If
  Next
  SPBblue.enabled=1
End Sub


Sub SPBblue_Timer ' 88ms
  SPBblue.enabled=0
  for i = 1 to 139
    If SLS(i,13)>0 Then SLS(i,5)=0
  Next
End Sub

Dim LShalfandhalf
Sub SPBhalfandhalf(tmp1) ' nr of blinks
  LShalfandhalf = tmp1
  LShalfcounter=0
  SPBleftright.enabled=1
End Sub

Dim LShalfcounter
Sub SPBleftright_Timer ' 30ms
  LShalfcounter=LShalfcounter+1

  Select Case  LShalfcounter
    Case 1 :
      for i = 1 to 130
        If SLS(i,13)>475 Then
          SLS(i,5)=1 'On
        Elseif SLS(i,13)>0 Then
          SLS(i,5)=0 'off
        End If
      Next

    Case 4 :
      for i = 1 to 130
        If SLS(i,13)>475 Then
          SLS(i,5)=2 'off
        Elseif SLS(i,13)>0 Then
          SLS(i,5)=1 'on
        End If
      Next

    Case 6 :
      LShalfcounter=0
      LShalfandhalf=LShalfandhalf-1
      If LShalfandhalf<1 Then
        SPBleftright.enabled=0
        for i = 1 to 170
          If sls(i,14)>0 Then SLS(i,5)=0 ' off Spesial state all that has stored XY
        Next
      End If
  End Select

End Sub




'**********************************************************************************************************
'* Fade level *
'**********************************************************************************************************
Sub FadeL ( nr )
  If LSALLUP=1 And SLS(nr,14)<LShead And SLS(nr,14)>LShead+LSaddH And SLS(nr,5)=2 Then
    SPB1 nr,nr,2, LStail ,0, 1
  Elseif LSALLDOWN=1 And SLS(nr,14)>LShead And SLS(nr,14)<LShead+LSaddH And SLS(nr,5)=2 Then
    SPB1 nr,nr,2, LStail ,0, 1
  End If
  i = SLS(nr,2)

  If SLS(nr,5) < 1 Then ' no special light state
    Select Case SLS(nr,1)
      Case 1 : i=i+1
          If i = 1 then i = 2
          If i > 6 Then i = 6
      Case 0 : i=i-1 : If i < 0 Then i = 0
      Case 2 : i=i-1 : If i < 0 Then i = 0 :          SLS(nr,1)=3             ' 2+ = blinking
      Case 3 : i=i+1 : If i > 6 Then i = 6          : SLS(nr,1)=4 : SLS(nr,4)=SLS(nr,3)
      Case 4 : SLS(nr,4) = SLS(nr,4) -1 : If SLS(nr,4)<0 Then SLS(nr,1)=5
      Case 5 : i=i-1 : If i < 0 Then i = 1 :          SLS(nr,1)=6 : SLS(nr,4)=SLS(nr,3)
      Case 6 : SLS(nr,4) = SLS(nr,4) -1
            If SLS(nr,4)<0 Then
              If SLS(nr,9)>0 Then
                SLS(nr,9)=SLS(nr,9)-1
                If SLS(nr,9)=0 Then
                  SLS(nr,1)=1 'turn on after blinking is done
                Else
                  SLS(nr,1)=3  ' next blink
                End If
              Else
                SLS(nr,1)=3  ' state 2 : no number of blinks set = always blink
              End If
            End If
    End select
'SLS=1 LFS=2 SLSP=3 SLSQ=4 LSP=5 LSPP=6 LSPQ=7 8=sp nrblinks   9= normal numberof blinks ( if 8 or 9 = 0 ) then repeat blinking forever 10=special state ends after xxx frames ( a blink is 12 frames )

  Else
    If SLS(NR,11)>0 Then SLS(nr,11)=SLS(nr,11)-1 : If SLS(nr,11)>0 Then Exit Sub
    If SLS(nr,10)>0 Then SLS(nr,10)=SLS(nr,10)-1 : If SLS(nr,10)=0 Then SLS(nr,5)=0
    Select Case SLS(nr,5)
      ' special takes over  2= off 3=blink
      Case 1 : i=i+1
          If i = 1 then i = 2
          If i > 6 Then i = 6
      Case 2 : i=i-1 : If i < 0 Then i = 0
      Case 3 : i=i-1 : If i < 0 Then i = 0 :          SLS(nr,5)=4             ' 2+ = blinking
      Case 4 : i=i+1 :
           If i=1 then i=2
           If i > 6 Then i = 6                : SLS(nr,5)=5 : SLS(nr,7)=SLS(nr,6)
      Case 5 : SLS(nr,7) = SLS(nr,7) -1 : If SLS(nr,7)<0 Then SLS(nr,5)=6
      Case 6 : i=i-1 : If i < 0 Then i = 1 :          SLS(nr,5)=7 : SLS(nr,7)=SLS(nr,6)
      Case 7 : SLS(nr,7) = SLS(nr,7) -1
            If SLS(nr,7)<0 Then
              If SLS(nr,8)>0 Then
                SLS(nr,8)=SLS(nr,8)-1
                If SLS(nr,8)=0 Then
                  If SLS(nr,12)=0 Then
                    SLS(nr,5)=1 'turn on after blinking is done . special state still on
                  Else
                    SLS(nr,12)=0 'reset flag 12
                    SLS(nr,5)=0 'turn off special if flag 12 is set
                  End If
                Else
                  SLS(nr,5)=4  ' next blink
                End If
              Else
                SLS(nr,5)=4  ' state 2 : no number of blinks set = always blink
              End If
            End If
    End select

  End If

  SLS(nr,2) = i

End Sub

Sub SetL( object,Multi)
  object.intensity = i * Multi * 0.75
End Sub

Sub SetP2 ( object,Multi)
  object.BlendDisableLighting = i*multi
End Sub

Sub SetP ( object,Multi)
  object.BlendDisableLighting = i*Multi
  object.visible=i
End Sub

Dim GIsound
Sub SetGI(nr,object)
  i= SLS(nr,1)
  If SLS(nr,2) > 0 Then ' delayed turnoff
    GIsound=GIsound+1
    If GIsound=260 Then GIsound=0 : PlaySound "fx_relay", 1,0.15*VolumeDial,0,0,0,0,0,0
    i=2
    SLS(nr,2) = SLS(nr,2) - 1
  End If
  object.state = i
End Sub




'**********************************************************************************************************
'* Lightshow at Gameover *
'**********************************************************************************************************

DIM LS1,stopgi,count01,count02,GIswitch,FlashCount
Sub lightshow
'IF count01=10 Then  ' ONLY TO TURN ODFF


  count01=count01+1
  If count01=22 And Int(Rnd(1)*3)=1 Then count01=0
  If count01=43 And Int(Rnd(1)*2)=1 Then count01=22
  If count01>44 Then
    count01=0
    If Giswitch=0 Then
      GIswitch=1
    Else
      GIswitch=0
  End if
    For i = 179 To 190
      SLS(i,1)=GiSwitch
    Next
    PlaySound "fx_relay", 1,0.3*VolumeDial,0,0,0,0,0,0
  End if

  select case count01
    case 1 ,37 : SLS(2,1)=1 : SLS(7,1) =1 : SLS(15,1) =1  : SLS(50,1) =1 : SLS(24,1) =1 : SLS(29,1) =1 : SLS(34,1) =1 : SLS(35,1) =1 : SLS(41,1) =1 : SLS(46,1) =1 : SLS(47,1) =1
    case 4 ,34 : SLS(3,1)=1 : SLS(8,1) =1 : SLS(16,1) =1  : SLS(51,1) =1 : SLS(23,1) =1 : SLS(28,1) =1 : SLS(33,1) =1 : SLS(36,1) =1 : SLS(42,1) =1 : SLS(48,1) =1 : SLS(71,1) =1
    case 7 ,31 : SLS(4,1)=1 : SLS(9,1) =1 : SLS(17,1) =1  : SLS(52,1) =1 : SLS(22,1) =1 : SLS(27,1) =1 : SLS(32,1) =1 : SLS(37,1) =1 : SLS(43,1) =1 : SLS(49,1) =1 : SLS(85,1) =1
    case 10,28 : SLS(5,1)=1 : SLS(10,1)=1 : SLS(18,1) =1  : SLS(53,1) =1 : SLS(21,1) =1 : SLS(26,1) =1 : SLS(31,1) =1 : SLS(38,1) =1 : SLS(44,1) =1 : SLS(63,1) =1
    case 13,25 : SLS(6,1)=1 : SLS(11,1)=1 : SLS(19,1) =1  : SLS(54,1) =1 : SLS(20,1) =1 : SLS(25,1) =1 : SLS(30,1) =1 : SLS(39,1) =1 : SLS(45,1) =1 : SLS(64,1) =1 : SLS(72,1) =1
    case 16,22 :                    SLS(55,1) =1 :                          SLS(40,1) =1:        SLS(70,1) =1

    case 6 ,42 : SLS(2,1)=0 : SLS(7,1) =0 : SLS(15,1) =0  : SLS(50,1) =0 : SLS(24,1) =0 : SLS(29,1) =0 : SLS(34,1) =0 : SLS(35,1) =0 : SLS(41,1) =0 : SLS(46,1) =0 : SLS(47,1) =0
    case 9 ,39 : SLS(3,1)=0 : SLS(8,1) =0 : SLS(16,1) =0  : SLS(51,1) =0 : SLS(23,1) =0 : SLS(28,1) =0 : SLS(33,1) =0 : SLS(36,1) =0 : SLS(42,1) =0 : SLS(48,1) =0 : SLS(71,1) =0
    case 12,36 : SLS(4,1)=0 : SLS(9,1) =0 : SLS(17,1) =0  : SLS(52,1) =0 : SLS(22,1) =0 : SLS(27,1) =0 : SLS(32,1) =0 : SLS(37,1) =0 : SLS(43,1) =0 : SLS(49,1) =0 : SLS(85,1) =0
    case 15,33 : SLS(5,1)=0 : SLS(10,1)=0 : SLS(18,1) =0  : SLS(53,1) =0 : SLS(21,1) =0 : SLS(26,1) =0 : SLS(31,1) =0 : SLS(38,1) =0 : SLS(44,1) =0 : SLS(63,1) =0
    case 18,30 : SLS(6,1)=0 : SLS(11,1)=0 : SLS(19,1) =0  : SLS(54,1) =0 : SLS(20,1) =0 : SLS(25,1) =0 : SLS(30,1) =0 : SLS(39,1) =0 : SLS(45,1) =0 : SLS(64,1) =0 : SLS(72,1) =0
    case 21,27 :                    SLS(55,1) =0 :                      SLS(40,1) =0 :        SLS(70,1) =0

  End Select

  count02=count02+1
    If count02>14 Then
      count02=0
      FlashCount=FlashCount+1
      Select Case FlashCount
        Case 1 : Objlevel(1) = 1 : FlasherFlash1_Timer : Objlevel(8) = 1 : FlasherFlash8_Timer :  playSound "fx_relay", 1,0.3*VolumeDial,0,0,0,0,0,0
            DOF 108,2
            If DOFdebug=1 Then debug.print "DOF 108,2 flasherB1"
            DOF 111,2
            If DOFdebug=1 Then debug.print "DOF 111,2 flasherR1"

      for i = 78 to 82 : SLS(i,1)=1 : Next

        Case 2 : Objlevel(2) = 1 : FlasherFlash2_Timer : Objlevel(3) = 1 : FlasherFlash3_Timer :  playSound "fx_relay", 1,0.3*VolumeDial,0,0,0,0,0,0

            DOF 112,2
            If DOFdebug=1 Then debug.print "DOF 112,2 flasherb2"
            DOF 109,2
            If DOFdebug=1 Then debug.print "DOF 109,2 flasher1"

      for i = 78 to 82 : SLS(i,1)=0 : Next
      SLS(83,1)=1

        Case 3 : Objlevel(7) = 1 : FlasherFlash7_Timer : Objlevel(9) = 1 : FlasherFlash9_Timer :  playSound "fx_relay", 1,0.3*VolumeDial,0,0,0,0,0,0

            DOF 113,2
            If DOFdebug=1 Then debug.print "DOF 113,2 flasher7"
            DOF 110,2
            If DOFdebug=1 Then debug.print "DOF 110,2 flasher1"
      SLS(83,1)=0

        Case 4 : FlashCount=0
      End Select
    End If
  Select Case count02
    Case 1  :  SLS(59,1)=1 : SLS(14,1)=1 : SLS(65,1)=1 : BlueDisplay(8)
    Case 3  :  SLS(60,1)=1 : SLS(13,1)=1 : SLS(66,1)=1 : SLS(73,1)=1
    Case 5  :  SLS(61,1)=1 : SLS(12,1)=1 : SLS(67,1)=1 : REDDisplay(8)
    Case 7  :  SLS(62,1)=1 :         SLS(68,1)=1 : SLS(74,1)=1

    Case 6  :  SLS(59,1)=0 : SLS(14,1)=0 : SLS(65,1)=0 : BlueDisplay(10)
    Case 8  :  SLS(60,1)=0 : SLS(13,1)=0 : SLS(66,1)=0 : SLS(73,1)=0
    Case 10 :  SLS(61,1)=0 : SLS(12,1)=0 : SLS(67,1)=0 : REDDisplay(10)
    Case 12 :  SLS(62,1)=0 :         SLS(68,1)=0 : SLS(74,1)=0
  End Select

  If Int(Rnd(1)*17)=2 Then

    SLS( 56 + LS1 ,1 ) = 0

    LS1 = LS1 + 1
    If LS1 > 2 Then LS1 = 0

      SLS( 56 + LS1 ,1 ) = 1
  End If


  If Int(Rnd(1)*10)=3 Then
    PlaySound "fx_relay", 1,0.1*VolumeDial,0,0,0,0,0,0
    i= Int(Rnd*12)
    If SLS(179+i,1) = 1 Then
      SLS(179+i,1) = 0
    Else
      SLS(179+i,1) = 1
    End If
  End If

'   If Int(Rnd(1)*40)=3 Then SPB1 1,200,1,0,0,1 did not fit there

End Sub


'**********************************************************************************************************
'* Taunting * And * Randomsounds * Random background noise *
'**********************************************************************************************************

Sub Taunting_Timer  ' male 1,2   female 1,2

  If tilted=0 Then
    If Taunting.enabled=0 Then
      Taunting.enabled=1 '  2sec delay
    Else

      If SC(PN,20)=0 Then playsound "a" & int(rnd(1)*47)+1 , 1, 0.9*BG_Volume, 0, 0,0,0, 0, 0
      If SC(PN,20)=1 Then playsound "m" & int(rnd(1)*39)+1 , 1, 0.9*BG_Volume, 0, 0,0,0, 0, 0

      If SC(PN,20)=2 Then playsound "b" & int(rnd(1)*46)+1 , 1, 0.9*BG_Volume, 0, 0,0,0, 0, 0
      If SC(PN,20)=3 Then playsound "c" & int(rnd(1)*50)+1 , 1, 0.9*BG_Volume, 0, 0,0,0, 0, 0
      Taunting.enabled=0
    End If
  End If

End Sub



Sub Taunting2_Timer
  Dim weaponsounds
  weaponsounds=Array("aarocket_exp","aarocketload","aarocket_fire","aasniperfire1","aasniperfire2","aareload","aashock_fire","aashockalt_fire","aamini_fire","aaflakload","aaflak_fire","aadrawsniper","biofire","bioimpact","controlsound","minialtfire","pulsebolt","pulsefire","ripperfire")
'19
  If Startgame=1 And Tilted=0 And plungertoolong.enabled=0 Then

    i=int(rnd(1)*35)
    If i=SC(PN,20) Then i=i+1
    tempnr=int(Rnd(1)*3)-1
    tempnr2=rnd(1) * 0.3 * BG_Volume
    Select Case i
      Case 0 : playsound "a" & int(rnd(1)*47)+1 , 1, tempnr2, 0, 0,0,0, 0, 0
      Case 1 : playsound "m" & int(rnd(1)*39)+1 , 1, tempnr2, 0, 0,0,0, 0, 0
      Case 2 : playsound "b" & int(rnd(1)*46)+1 , 1, tempnr2, 0, 0,0,0, 0, 0
      Case 3 : playsound "c" & int(rnd(1)*50)+1 , 1, tempnr2, 0, 0,0,0, 0, 0
      Case 4,5,6,7,8,9,10,11,12,13    : playsound Weaponsounds(int(rnd(1)*19)),1,  tempnr2, Rnd(1)*tempnr, 0,0,0, 0, 0
      Case 14,15,24,25,32,33,34 :      playsound "f" & int(rnd(1)*36)+1,1,  tempnr2, Rnd(1)*tempnr, 0,0,0, 0, 0

      Case 16,17,18,19,20,21,22,23    : playsound "d" & int(rnd(1)*96)+1 , 1, tempnr2, Rnd(1)*tempnr, 0,0,0, 0, 0
      Case 26,27,28,29,30,31        : playsound Weaponsounds(int(rnd(1)*19)),1,  tempnr2/2.5 , Rnd(1)*tempnr, 0,0,0, 0, 0
    End Select

  End If

End Sub

Sub perfection_Timer
  playsound "perfection", 1, 0.7*BG_Volume, 0, 0,0,0, 0, 0
  perfection.enabled=0
End Sub

'**********************************************************************************************************
'* KICKERS *
'**********************************************************************************************************
Dim TrippleHS
Dim DoubleHS
Sub HeadShot

  playsound "AAheadshot", 1, 0.7*BG_Volume, 0, 0,0,0, 0, 0

  If TrippleHS>0 then
    TrippleHS=TrippleHS+1
  Else
    TrippleHS=1
  End If

  If TrippleHS=3 And doubleHS=0 Then
    doubleHS=2
    ' fix lights anim ation ?
    '  text ACTIVE 2x , HEADSHOT fast
    SPBallup 35,6,3  ' speed up, lightstayontime, blinks
    If UMainDMD.isrendering Then UMainDMD.cancelrendering
    FlexFlashText2 "SUPER","HEADSHOT",60
    perfection.enabled=1

    Scoring 50000,6000
    FFT=FFTscore
    FlexFlashText2cap "HeadShotX2" , FFT , 60
    If DB2S_on Then B2STimer1_Timer  ' or b2stimer2 for big one
    DOF 142,2
    If DOFdebug=1 Then debug.print "DOF 142,2 Headshot"
  End If

  If doubleHS=0 Then
    Scoring 25000,3000
    FFT=FFTscore
    FlexFlashText2cap "HeadShot" , FFT , 75
    If DB2S_on Then B2STimer1_Timer  ' or b2stimer2 for big one
    DOF 142,2
    If DOFdebug=1 Then debug.print "DOF 142,2 Headshot"
  End If
  If doubleHS=1 Then
    Scoring 50000,6000
    FFT=FFTscore
    FlexFlashText2cap "HeadShot2X" , FFT , 70
    If DB2S_on Then B2STimer1_Timer  ' or b2stimer2 for big one
    DOF 142,2
    If DOFdebug=1 Then debug.print "DOF 142,2 Headshot"
  End If
  If DoubleHS=2 Then DoubleHS=1


End Sub

Sub Frag
  Scoring 1000,200
End Sub



Sub Jackpot

  Dim JP1
  JP1=Array("perfection","stepaside","wantsome","inferior","myhouse","superior","obsolete","diehuman","imonfire","hadtohurt","fearme","bowdown")
  If UMainDMD.isrendering Then UMainDMD.cancelrendering
  Scoring 150000,15000
  FFT=FFTscore
  FlexFlashText2 "Camper","Bonus" , 95
  SPB1 1,139,2,1,0,1 ' megablinks
  PlaySound JP1(int(Rnd(1)*12)) , 1, 0.9 * BG_Volume, 0, 0,0,0, 0, 0

End Sub


Sub FirstbloodX
  PlaySound "AAFirstblood" , 1, BG_Volume, 0, 0,0,0, 0, 0 : Firstblood=1
  FlexFlashText1 "FirstBlood" , 80
End Sub

Sub SetlightsRedeemer

    If SC(PN,13)>0 Then SLS( 7,1)=1 : SPB1 7,7,3,2,1,1 '  12  Rightorbit
    If SC(PN,13)>1 Then SLS( 8,1)=1 : SPB1 8,8,3,2,1,1
    If SC(PN,13)>2 Then SLS( 9,1)=1 : SPB1 9,9,3,2,1,1
    If SC(PN,13)>3 Then SLS(10,1)=1 : SPB1 10,10,2,3,1,1
    If SC(PN,13)>4 Then SLS(11,1)=1 : SPB1 11,11,2,3,1,1

    If SC(PN,12)>0 Then SLS(15,1)=1 : SPB1 15,15,2,3,1,1' 13  Leftorbit nr(1-5=nr of lights lit) + more states ?
    If SC(PN,12)>1 Then SLS(16,1)=1 : SPB1 16,16,2,3,1,1
    If SC(PN,12)>2 Then SLS(17,1)=1 : SPB1 17,17,2,3,1,1
    If SC(PN,12)>3 Then SLS(18,1)=1 : SPB1 18,18,2,3,1,1
    If SC(PN,12)>4 Then SLS(19,1)=1 : SPB1 19,19,2,3,1,1

    If SC(PN,14)>0 Then SLS(20,1)=1 : SPB1 20,20,2,3,1,1' 14  kicker1
    If SC(PN,14)>1 Then SLS(21,1)=1 : SPB1 21,21,2,3,1,1
    If SC(PN,14)>2 Then SLS(22,1)=1 : SPB1 22,22,2,3,1,1
    If SC(PN,14)>3 Then SLS(23,1)=1 : SPB1 23,23,2,3,1,1
    If SC(PN,14)>4 Then SLS(24,1)=1 : SPB1 24,24,2,3,1,1

    If SC(PN,15)>0 Then SLS(25,1)=1 : SPB1 25,25,2,3,1,1' 15  kicker2
    If SC(PN,15)>1 Then SLS(26,1)=1 : SPB1 26,26,2,3,1,1
    If SC(PN,15)>2 Then SLS(27,1)=1 : SPB1 27,27,2,3,1,1
    If SC(PN,15)>3 Then SLS(28,1)=1 : SPB1 28,28,2,3,1,1
    If SC(PN,15)>4 Then SLS(29,1)=1 : SPB1 29,29,2,3,1,1

    If SC(PN,16)>0 Then SLS(30,1)=1 : SPB1 30,30,2,3,1,1' 16  kicker3
    If SC(PN,16)>1 Then SLS(31,1)=1 : SPB1 31,31,2,3,1,1
    If SC(PN,16)>2 Then SLS(32,1)=1 : SPB1 32,32,2,3,1,1
    If SC(PN,16)>3 Then SLS(33,1)=1 : SPB1 33,33,2,3,1,1
    If SC(PN,16)>4 Then SLS(34,1)=1 : SPB1 34,34,2,3,1,1

    CaptureCheck

End Sub


Sub Kicker001_Hit   ' JUST A TEMP ONE...
  SPB1 99,99,4,10,0,1
  SPB1 86,88,5,5,0,1

  If KBstatus=0 Then LiteKBstatus=1 : SLS(47,1)=2



'E139 mainramp done
'E140 redeemer
'E141 Jackpot
'E142 HeadShot  done
  If  SS2=1 Then
    SS2=0
    Debug.print "SS2 awarded"
    L059.timerenabled=0
    Skillshot2
  End If

  If Lockingthisball=1 Then
    kicker001.kick 90,48
  SPBblueside

    SoundSaucerKick 1, kicker001
    exit sub
  End If

  If SLS(44,1)>1 Then
    HeadShot
    i = int(rnd(1)*3)
    If SC(PN,12+i)<5 Then
      SC(PN,12+i)=SC(PN,12+i)+1
    elseIf SC(PN,13+i)<5 then
      SC(PN,13+i)=SC(PN,13+i)+1
    elseIf SC(PN,14+i)<5 then
      SC(PN,14+i)=SC(PN,14+i)+1
    End If
    If SLS(67,1)=0 Then EnemyFlag

    SetlightsRedeemer
  End If
  TrippleHS=0


  If SLS(45,1)>1 Then
    Jackpot
    DOF 141,2
    If DOFdebug=1 Then debug.print "DOF 141,2 jackpot"

  End If

  If Invisiblilty(1)=1 And Invisiblilty(2)=1 And Invisiblilty(3)=1 And Invisiblilty(4)=1 And Invisiblilty(5)=1 Then

    SPB1 78,83,10,2,5,1
    Respawn=Respawn+60
    SPB1 120,139,5,1,3,1
    For i = 120 to 139
      SLS(i,1)=1
    Next
    SLS(64,1)=2
    SLS(64,3)=6
    Scoring 200000,30000
    FFT=FFTscore
    FlexFlashText1 "INVISIBILITY",66

    DOF 133,2
    If DOFdebug=1 Then debug.print "DOF 133,2 invisibility awarded"
'133 invisibility awarded
    Invisiblilty(1)=0
    Invisiblilty(2)=0
    Invisiblilty(3)=0
    Invisiblilty(4)=0
    Invisiblilty(5)=0
    SLS (78,1)=0
    SLS (79,1)=0
    SLS (80,1)=0
    SLS (81,1)=0
    SLS (82,1)=0
    SPB1 79,82,3,2,0,1
  End If




  HeadShotON
  FastFrags=FastFrags+1
  Frag
  SC(PN,19)=SC(PN,19)+1
  CheckFastFrags

  SoundSaucerLock()


  If SLS(65,1)>1 Then
    DOF 143,2
    If DOFdebug=1 Then debug.print "DOF 143,2 ball locked"
    'E143 ball locked 1
    SPB1 86,88,10,1,0,1
    If locked1=0 or locked2=0 Then
    ' lock this ball and shoot out a new one
    DividerWall0_Timer
    LockThisBall1.enabled=1  ' 1sec then kick to lock Position
    ' turn on lock1 or 2
    SPBalldown 40,15,4
    Else
    SPBalldown 40,15,4
      Kicker001.destroyball
      RandomSoundDrain(Kicker001)

      SC(PN,17)=SC(PN,17)+1
      If SC(PN,17)=2 Then
        SLS(65,1)=0
        SLS(66,1)=0
        SPB1 65,66,3,2,1,1  ' blinks,timer,special off after done
        SLS(68,1)=2 '  last for multiball set ! kicker 3?
      End If
      PlaySound "searchdestroy" & int(rnd(1)*2)+1 , 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0

      AutoKick=0
      RespawnWaiting=RespawnWaiting+1
      DOF 151,2
      If DOFdebug=1 Then debug.print "DOF 151,2 ballrelease"
      'E151 2 Ballrelease ( simulated )
      DOF 151,2
      If DOFdebug=1 Then debug.print "DOF 151,2 ballrelease"
      L062_Timer
      'RandomSoundBallRelease(BallRelease)
      PlaySound "AARespawn", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
      FlexFlashText3 "READY","SET GO",70
    End If

  elseIf SLS(67,1)>0 Then ' redeemer is on !
    DOF 140,2
    If DOFdebug=1 Then debug.print "DOF 140,2 Redeemer"
    SPB1 67,67,7,2,1,1 ' 3 blinks rEDEEMER
    SLS(67,1)=0 ' turn off redeemer light
    Scoring 20000,4000
    For i = 179 to 190
      SLS(i,2)=300  ' blink for 300 updates
    Next
    ' redeemer award 2 "random" Lights
    i = int(rnd(1)*3)
    If SC(PN,12+i)<5 Then
      SC(PN,12+i)=SC(PN,12+i)+1
    elseIf SC(PN,13+i)<5 Then
       SC(PN,13+i)=SC(PN,13+i)+1
    elseIf SC(PN,14+i)<5 Then
       SC(PN,14+i)=SC(PN,14+i)+1
    End If
    i = int(rnd(1)*3)
    If SC(PN,12+i)<5 Then
      SC(PN,12+i)=SC(PN,12+i)+1
    elseIf SC(PN,13+i)<5 Then
       SC(PN,13+i)=SC(PN,13+i)+1
    elseIf SC(PN,14+i)<5 Then
       SC(PN,14+i)=SC(PN,14+i)+1
    End If
    EnemyFlag

    SetlightsRedeemer


    If SLS(48,1)>1 Then   ' return enemy flag
    ' turn off enemyflag  l046 1500ms
      L041.timerenabled=1
      SLS(41,1)=0
        If SLS(42,1)=1 Then SLS(42,1)=2
      SLS(48,1)=0
      SLS(63,1)=0
      SLS(70,1)=0
      FlexNewText "Flag returned!"
      Scoring 5000,1000


    End If

    Gate6.timerenabled=1 ' delayed sound 2sec

    kicker001.destroyball
    kicker003.createball

    PlaySound "AARedeemer", 1, 0.7 * BG_Volume, 0, 0,0,0, 0, 0
    FlexFlashText1 "INCOMMING" , 80
    DividerWall0_Timer '  open for normal biz
  Else ' normal operation
    kicker001.destroyball
    kicker002.createball
    kicker002.kick 260,11
    SPB1 86,88,1,2,0,1
    SPBredside ' nr of blinks
  End If

End Sub


Sub HeadShotON
  MapBlinker_Timer
  SLS(44,1)=2 ' blinking
  L044.timerenabled=0
  L044.timerenabled=1 ' headshot timer 10 sec
End Sub


Sub Kicker003_hit
  kicker003.kick 40,40
  SPB1 86,88,1,2,0,1
  SoundSaucerKick 1, kicker003
End Sub


Sub Gate6_Timer 'redeemer

    FastFrags=FastFrags+2
  Frag
  Frag
    HeadShotON
    SC(PN,19)=SC(PN,19)+2
    CheckFastFrags

    kicker003.kick 30,30+int(rnd(1)*10)
    SPB1 86,88,1,2,0,1
    SoundSaucerKick 1, kicker003
    PlaySound "AAlift1act", 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0
    Gate6.timerenabled=0
End Sub


'invisibilty (1) = sw40
'invisibilty (2) = 21
'invisibilty (3) = topvuk
'invisibilty (4) = 32
'invisibilty (5) = 22

Sub Sw40_hit  ' kicker 1
  TrippleHS=0
  Objlevel(2) = 1 : FlasherFlash2_Timer : playSound "fx_relay", 1,0.1*VolumeDial,0,0,0,0,0,0
  DOF 112,2
  If DOFdebug=1 Then debug.print "DOF 112,2 flasherb2"

  'flak
  DOF 134,2
  If DOFdebug=1 Then debug.print "DOF 134,2 left saucer"
  '133 invisibility awarded
  '134 left saucer
  '135 left orbit
  '136 TopVUK
  '137 right saucer
  '138 right orbit

  If Invisiblilty(1)=1 Then
    Invisiblilty(1)=1
    Invisiblilty(2)=0
    Invisiblilty(3)=0
    Invisiblilty(4)=0
    Invisiblilty(5)=0
    SLS (78,1)=1
    SLS (79,1)=0
    SLS (80,1)=0
    SLS (81,1)=0
    SLS (82,1)=0
    SLS (83,1)=0
    SPB1 78,82,2,2,0,1
  Else
    Invisiblilty(1)=1
    SLS (78,1)=1
    SPB1 78,78,5,2,0,1
  End If


  If Firstblood=0 Then FirstBloodX
  HeadShotON
  FastFrags=FastFrags+1
  Frag
  SC(PN,19)=SC(PN,19)+1
  CheckFastFrags

  If SLS(43,1)>0 Then
    SLS(43,1)=0
    Extraball = 2 ' 2 = gotten
    SLS(85,1)=1
    SPB1 85,85,8,8,1,1
    SPB1 43,43,3,2,1,1

    PlaySound "AANali" & Int(rnd(1)*6)+1 , 1, BG_Volume, 0, 0,0,0, 0, 0
    FlexFlashText3 "EXTRA","LIFE" , 80
    DOF 126,2
    If DOFdebug=1 Then debug.print "DOF 126,2 award extraball"
  End If


  If SLS(48,1)>1 Then
    ' turn off enemyflag  l041 1500ms
    L041.timerenabled=1
    SLS(41,1)=0
    If SLS(42,1)=1 Then SLS(42,1)=2
    SLS(48,1)=0
    SLS(63,1)=0
    FlexNewText "Flag returned!"
    SLS(70,1)=0
    Scoring 5000,1000

  End If

  EnemyFlag
  GameTimerCheck

    ' check to see if lock lights should be turned on
  If MBactive=0 Then
    SC(PN,18) = SC(PN,18) + 1   ' 18 is lights done
    If SC(PN,18)=7 Then
      DividerWall1_Timer ' divider=on ramp goes to tower
      PlaySound "AApointsecure" & Int(rnd(1)*4)+1 , 1, BG_Volume, 0, 0,0,0, 0, 0
      FlexFlashText2 "Point 1","Secure" , 80
      SLS(65,1)=2
      SLS(66,1)=2
      LockBolt1.Z=140
      LockBolt1.collidable=True
      debug.Print "Boltmoved UP"
    End If
  End If



  SC(PN,14) = SC(PN,14) + 1
  If SC(PN,14) = 5 Then
    '!Turn on capture!
    SLS(42,1)=2 ' blinking
    If SLS(41,1)>1 Then SLS(42,1)=1
    SPB1 42,42,7,2,2,1
    PlaySound "wdend04", 1, BG_Volume, 0, 0,0,0, 0, 0
    If UMainDMD.isrendering Then UMainDMD.cancelrendering
    FlexFlashText3 "Capture","Lit",80
    PlaySound "AAassist", 1, BG_Volume, 0, 0,0,0, 0, 0
  Elseif SC(PN,14) > 4 Then
    SC(PN,14) = 5
  End If

  Select Case SC(PN,14)
    Case 1: SLS(20,1)=1
    Case 2: SLS(21,1)=1
    Case 3: SLS(22,1)=1
    Case 4: SLS(23,1)=1
        PlaySound "AASpreeSound", 1, BG_Volume, 0, 0,0,0, 0, 0
    Case 5: SLS(24,1)=1
        Scoring 1000,100
  End Select

  SoundSaucerLock()
  Sw40.timerenabled=1
  SPB1 20,24,3,2, 7,1  ' light 7,specialstate, 3 blinks,timer,special off after done
  SLS(182,2)=33 ' blink middletopgi  1 solid blink ``?

End Sub

Sub Sw40_Timer
  Sw40.Kick 135,3
  SoundSaucerKick 1, Sw40
  PlaySound "AAFlak_fire", 1, 0.7*BG_Volume, 0, 0,0,0, 0, 0
  Sw40.timerenabled=0
  Objlevel(1) = 1 : FlasherFlash1_Timer : playSound "fx_relay", 1,0.3*VolumeDial, AudioPan(Flasherbase1),0,0,0,0, AudioPan(Flasherbase1)
            DOF 111,2
            If DOFdebug=1 Then debug.print "DOF 111,2 flasherR1"

End Sub





Sub TopVUK_Hit ' kicker2
    TrippleHS=0

    SPB1 99,99,4,10,0,1
    SPB1 86,88,1,2,0,1

    DOF 136,2
    If DOFdebug=1 Then debug.print "DOF 136,2 Top VUK"
'133 invisibility awarded
'134 left saucer
'135 left orbit
'136 TopVUK
'137 right saucer
'138 right orbit
  If Invisiblilty(3)=1 Then
    Invisiblilty(1)=0
    Invisiblilty(2)=0
    Invisiblilty(3)=1
    Invisiblilty(4)=0
    Invisiblilty(5)=0
    SLS (78,1)=0
    SLS (79,1)=0
    SLS (80,1)=1
    SLS (81,1)=0
    SLS (82,1)=0
    SLS (83,1)=0
    SPB1 78,82,2,2,0,1
  Else
    Invisiblilty(3)=1
    SLS (80,1)=1
    SPB1 80,80,5,2,0,1
  End If

  If Firstblood=0 Then FirstBloodX
  HeadShotON
  FastFrags=FastFrags+1
  Frag
  SC(PN,19)=SC(PN,19)+1
  CheckFastFrags
  SoundSaucerLock()

  If SLS(65,1)>1 Then
    SPB1 86,88,10,1,0,1
    DOF 143,2
    If DOFdebug=1 Then debug.print "DOF 143,2 ball locked"
    Scoring 2000,1000
    If locked1=0 or locked2=0 Then
      ' lock this ball and shoot out a new one
      DividerWall0_Timer
      LockThisBall2.enabled=1  ' 1sec then kick to lock Position
      ' turn on lock1 or 2
      SPBalldown 40,15,4
    Else
      SPBalldown 40,15,4
      TopVUK.destroyball
      RandomSoundDrain(TopVUK)

      SC(PN,17)=SC(PN,17)+1
      If SC(PN,17)=2 Then
        SLS(65,1)=0
        SLS(66,1)=0
        SPB1 65,66,3,2,1,1  ' blinks,timer,special off after done
        SLS(68,1)=2 '  last for multiball set ! kicker 3?
      End If
      PlaySound "searchdestroy" & int(rnd(1)*2)+1 , 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0

      AutoKick=0
      RespawnWaiting=RespawnWaiting+1
      DOF 151,2
      If DOFdebug=1 Then debug.print "DOF 151,2 ballrelease"
      L062.timerenabled=0
      L062_Timer

      PlaySound "AARespawn", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
      FlexFlashText3 "READY","SET GO",70

    End If
  Else
    TopVUK.destroyball
    Kicker002.createball

    TopVUK.timerenabled=1
  End If


    ' check to see if lock lights should be turned on
    If MBactive=0 Then
      SC(PN,18) = SC(PN,18) + 1   ' 18 is lights done
      If SC(PN,18)=7 Then
        DividerWall1_Timer ' divider=on ramp goes to tower
        PlaySound "AApointsecure" & Int(rnd(1)*4)+1 , 1, BG_Volume, 0, 0,0,0, 0, 0
        FlexFlashText2 "Point 1","Secure" , 80
        SLS(65,1)=2
        SLS(66,1)=2
      LockBolt1.Z=140
      LockBolt1.collidable=True
debug.Print "Boltmoved UP"
      End If
    End If


  If SLS(42,1)>1 Then  CaptureTheFlag 'if capture blinks dont reward lighg


  EnemyFlag
  GameTimerCheck

  SC(PN,15) = SC(PN,15) + 1
  If SC(PN,15) =5 Then
    '!Turn on capture!
    SLS(42,1)=2 ' blinking
    If SLS(41,1)>1 Then SLS(42,1)=1
    SPB1 42,42,7,2,2,1
    PlaySound "wdend04", 1, BG_Volume, 0, 0,0,0, 0, 0
    If UMainDMD.isrendering Then UMainDMD.cancelrendering
    FlexFlashText3 "Capture","Lit",80
    PlaySound "AAassist", 1, BG_Volume, 0, 0,0,0, 0, 0
  elseif SC(PN,15) > 4 Then
    SC(PN,15) = 5
  End If


  Select Case SC(PN,15)
    Case 1: SLS(25,1)=1
    Case 2: SLS(26,1)=1
    Case 3: SLS(27,1)=1
    Case 4: SLS(28,1)=1
    PlaySound "AASpreeSound", 1, BG_Volume, 0, 0,0,0, 0, 0
    Case 5: SLS(29,1)=1
    Scoring 1000,100
  End Select

  SPB1 25,29,3,2, 7,1  ' light 7,specialstate, 3 blinks,timer,special off after done
  SLS(181,2)=33 ' blink middletopgi  1 solid blink ``?

End Sub

Sub GameTimerCheck
  SC(PN,31)=SC(PN,31)+1 ' each player keeps this going
  debug.print "progress" & SC(PN,31) & " if3=5min"

' if map playing = 1 '' so this is level1 coretff
  If SC(PN,27)=0 Then
    If SC(PN,31)=  3 Then PlaySound "cd5min", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
    If SC(PN,31)=  9 Then PlaySound "cd3min", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
    If SC(PN,31)= 19 Then PlaySound "cd1min", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
    If SC(PN,31)= 23 or SC(PN,31)=24 Then
      SC(PN,31)=24
      PlaySound "cd30sec", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
      'start 1sec timer with last countdown' restart at 32 if new ball !!!!
      Lastcd=0 : countdown30.enabled=1
    End If
  End If

  If SC(PN,27)=1 Then
    If SC(PN,31)=  3 Then PlaySound "cd5min", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
    If SC(PN,31)= 12 Then PlaySound "cd3min", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
    If SC(PN,31)= 25 Then PlaySound "cd1min", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
    If SC(PN,31)= 30  or SC(PN,31)=31 Then
      SC(PN,31)=31

      PlaySound "cd30sec", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
      'start 1sec timer with last countdown' restart at 32 if new ball !!!!
      Lastcd=0 : countdown30.enabled=1
    End If
  End If

  If SC(PN,27)=2 Then ' dreary
    If SC(PN,31)=  3 Then PlaySound "cd5min", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0

    If SC(PN,31)= 6 Then
      PlaySound "AACaptureSound" & int(rnd(1)*3)+1 , 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
      FlexNewText "! ENEMY CAPTURE !"
      dtbg_timer
      If SC(PN,28)=SC(PN,29)+1 Then LOSTLEAD.enabled=1

      SC(PN,29)=SC(PN,29)+1
      'E154 2 ENEMY Capture
      DOF 154,2
      If DOFdebug=1 Then debug.print "DOF 154,2 enemy capture"

      BlueDisplay SC(PN,29)
      SPB1 72,72,7,2,1,1
      SPB1 73,74,5,2,1,1
    End If

    If SC(PN,31)= 14 Then PlaySound "cd3min", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
    If SC(PN,31)= 28 Then PlaySound "cd1min", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
    If SC(PN,31)= 33 or SC(PN,31)=34 Then
      SC(PN,31)=34

      PlaySound "cd30sec", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
      'start 1sec timer with last countdown' restart at 32 if new ball !!!!
      Lastcd=0 : countdown30.enabled=1
    End If
  End If

  If SC(PN,27)=3 Then ' FACE
    If SC(PN,31)=  3 Then PlaySound "cd5min", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0

    If SC(PN,31)= 6 Then
      PlaySound "AACaptureSound" & int(rnd(1)*3)+1 , 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
      FlexNewText "! ENEMY CAPTURE !"
      dtbg_timer
      If SC(PN,28)=SC(PN,29)+1 Then LOSTLEAD.enabled=1

      SC(PN,29)=SC(PN,29)+1
      'E154 2 ENEMY Capture
      DOF 154,2
      If DOFdebug=1 Then debug.print "DOF 154,2 enemy capture"

      BlueDisplay SC(PN,29)
      SPB1 72,72,7,2,1,1
      SPB1 73,74,5,2,1,1
    End If

    If SC(PN,31)= 22 Then
      PlaySound "AACaptureSound" & int(rnd(1)*3)+1 , 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
      FlexNewText "! ENEMY CAPTURE !"
      dtbg_timer
      If SC(PN,28)=SC(PN,29)+1 Then LOSTLEAD.enabled=1

      SC(PN,29)=SC(PN,29)+1
      'E154 2 ENEMY Capture
      DOF 154,2
      If DOFdebug=1 Then debug.print "DOF 154,2 enemy capture"

      BlueDisplay SC(PN,29)
      SPB1 72,72,7,2,1,1
      SPB1 73,74,5,2,1,1
    End If


    If SC(PN,31)= 19 Then PlaySound "cd3min", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
    If SC(PN,31)= 33 Then PlaySound "cd1min", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
    If SC(PN,31)= 39 or SC(PN,31)=40 Then
      SC(PN,31)=40

      PlaySound "cd30sec", 1, 1.1*BG_Volume, 0, 0,0,0, 0, 0
      'start 1sec timer with last countdown' restart at 32 if new ball !!!!
      Lastcd=0 : countdown30.enabled=1
    End If
  End If
End Sub

Dim lastcd
Sub countdown30_Timer
  lastcd=lastcd+1
  If lastcd=20 Then PlaySound "cd10", 1, BG_Volume, 0, 0,0,0, 0, 0
  If lastcd=21 Then PlaySound  "cd9", 1, BG_Volume, 0, 0,0,0, 0, 0
  If lastcd=22 Then PlaySound  "cd8", 1, BG_Volume, 0, 0,0,0, 0, 0
  If lastcd=23 Then PlaySound  "cd7", 1, BG_Volume, 0, 0,0,0, 0, 0
  If lastcd=24 Then PlaySound  "cd6", 1, BG_Volume, 0, 0,0,0, 0, 0
  If lastcd=25 Then PlaySound  "cd5", 1, BG_Volume, 0, 0,0,0, 0, 0
  If lastcd=26 Then PlaySound  "cd4", 1, BG_Volume, 0, 0,0,0, 0, 0
  If lastcd=27 Then PlaySound  "cd3", 1, BG_Volume, 0, 0,0,0, 0, 0
  If lastcd=28 Then PlaySound  "cd2", 1, BG_Volume, 0, 0,0,0, 0, 0
  If lastcd=29 Then PlaySound  "cd1", 1, BG_Volume, 0, 0,0,0, 0, 0
  If lastcd=30 Then
    ENDMAP
    countdown30.enabled=0
  End If
End Sub


Sub ENDMAP
  TrippleHS=0
  DoubleHS=0
  Lastcap=0
  Lastcap2=0
  lastcap3=0

  LFinfo=0 : RFinfo=0
  If DB2S_on Then B2STimer2_Timer  ' or b2stimer2 for big one
  pwnage=0
  Invisiblilty(1)=0
  Invisiblilty(2)=0
  Invisiblilty(3)=0
  Invisiblilty(4)=0
  Invisiblilty(5)=0

  L049.timerenabled=0 ' reenable kickback both places
  KBstatus=1
  SLS(49,1)=1
  LiteKBstatus=0
  SLS(47,1)=0
' L047.timerenabled=0
  If SC(PN,28)>SC(PN,29) Then'check score  28 vs 29
    ' you win
    PlaySound  "winner", 1, 0.8*BG_Volume, 0, 0,0,0, 0, 0
    '12-16
    Respawn=Respawn+25
      SPB1 120,139,5,1,3,1
      For i = 120 to 139
        SLS(i,1)=1
      Next
    SLS(64,1)=2 ' respawn blinks
    SLS(64,3)=6
'   L064.timerenabled=1

    SLS(42,1)=0 ' Cap off
    SLS(41,1)=0 ' EF
    SLS(48,1)=0 ' EF
    SLS(70,1)=0 ' EF
    SLS(63,1)=0 ' EF


    SC(PN,12)=0
    SC(PN,13)=0
    SC(PN,14)=0
    SC(PN,15)=0
    SC(PN,16)=0
    For i = 7 to 11
      SLS(i,1)=0
    Next
    For i = 15 to 34
      SLS(i,1)=0
    Next
    SPB1 1,77,7,2,0,1 ' blink all 7 times

    For i = 179 to 190
    SLS(i,2)=800  ' blink for 230 updates
    Next
    Scoring 20000 * SC(PN,28)+50000,0
    FFT=FFTscore
    If UMainDMD.isrendering Then UMainDMD.cancelrendering
    FlexFlashText2cap "CAP BONUS",FFT,100
'   If SC(PN,27)=1 then FlexFlashText3 "BONUS NOVMBR",FFT,100
'   If SC(PN,27)=2 then FlexFlashText3 "BONUS DREARY,FFT,100
'   If SC(PN,27)=3 then FlexFlashText3 "BONUS FACE",FFT,100

    'reset PF Score
    SC(PN,28)=0
    SC(PN,29)=0
    BlueDisplay 0
    RedDisplay 0

    SC(PN,27) = SC(PN,27) +1 ' next map
    If SC(PN,27) >3 Then SC(PN,27)=3
    If SC(pn,27)=1 Then SC(PN,32)=4 ' november ONLY NEED TO SET GFB HERE
    If SC(pn,27)=2 Then SC(PN,32)=3 ' dreary
    If SC(pn,27)=3 Then SC(PN,32)=3 ' face

    SC(PN,31)=0 'reset gametimer

  ' turn off lights too
  ' make everything blink ?=
  ' CP can stay
  ' teamplay can stay
  ' redeemer can stay


  Else
    ' you lost .. even on draw and need to replay level
    PlaySound  "lostmatch", 1, 0.8*BG_Volume, 0, 0,0,0, 0, 0
    Respawn=Respawn+20
      SPB1 120,139,5,1,3,1
      For i = 120 to 139
        SLS(i,1)=1
      Next
    SLS(64,1)=2 ' respawn blinks
    SLS(64,3)=6

'   reset lights and start next level

    SLS(42,1)=0 ' Cap off
    SLS(41,1)=0 ' EF
    SLS(48,1)=0 ' EF
    SLS(70,1)=0 ' EF
    SLS(63,1)=0 ' EF

    SC(PN,12)=0
    SC(PN,13)=0
    SC(PN,14)=0
    SC(PN,15)=0
    SC(PN,16)=0
    For i = 7 to 11
      SLS(i,1)=0
    Next
    For i = 15 to 34
      SLS(i,1)=0
    Next
    SPB1 1,77,7,2,0,1 ' blink all 7 times

    For i = 179 to 190
    SLS(i,2)=800  ' blink for 230 updates
    Next

    If SC(PN,27)=3 then
      Scoring 20000 * SC(PN,28),0
      FFT=FFTscore
      If UMainDMD.isrendering Then UMainDMD.cancelrendering
      FlexFlashText3 "MAP BONUS !","Face ",111 ' this wont come
    End If
    'reset PF Score
    SC(PN,28)=0
    SC(PN,29)=0
    BlueDisplay 0
    RedDisplay 0

    SC(PN,31)=0 'reset gametimer
  End If


End Sub

Sub TopVUK_Timer
  SPB1 86,88,1,2,1,1
  Kicker002.kick 260,11
  SPBredside ' nr of blinks
  SoundSaucerKick 1, TopVUK
  PlaySound "AAsniperfire" & int(rnd(1)*2)+1 , 1, BG_Volume, 0, 0,0,0, 0, 0
  TopVUK.timerenabled=0
  Objlevel(7) = 1 : FlasherFlash7_Timer : playSound "fx_relay", 1,0.3*VolumeDial, AudioPan(Flasherbase7),0,0,0,0, AudioPan(Flasherbase7)
            DOF 113,2
            If DOFdebug=1 Then debug.print "DOF 113,2 flasher7"
End Sub
Sub TakenLead_Timer
  TAKENLEAD.enabled=0
  PlaySound "takenlead", 1, 0.9 * BG_Volume, 0, 0,0,0, 0, 0
End Sub
Sub LostLead_Timer
  LostLEAD.enabled=0
  PlaySound "lostlead", 1, 0.9 * BG_Volume, 0, 0,0,0, 0, 0
End Sub

Dim Lastcap,Lastcap2,lastcap3
Dim FFT 'used in flexflashtext2 atlest for now
Sub CaptureTheFlag

  SPB1  7,11,4,1,1,1 ' all blink alittle and 10 for the lane that capture
  SPB1 15,34,4,1,1,1

  SLS(42,1)=0 ' turn off capture Flag ( will get checked on capturecheck and reset if needed

  If SC(PN,14)>4 Then
    Lastcap=14
    SC(PN,14)=0
    CaptureCheck
    If SC(PN,27)=0 Then SLS(20,1)=1 : SC(PN,14)=1 : Else SLS(20,1)=0
    SLS(21,1)=0
    SLS(22,1)=0
    SLS(23,1)=0
    SLS(24,1)=0
    SPB1 20,24,10,1,1,1
    If lastcap2 = 14 or lastcap3=14 Then CapisDenied : exit Sub
    CapisGood
    Exit Sub
  End If

  If SC(PN,15)>4 Then
    Lastcap=15
    SC(PN,15)=0
    CaptureCheck
    If SC(PN,27)=0 Then SLS(25,1)=1 : SC(PN,15)=1 : Else SLS(25,1)=0
    SLS(26,1)=0
    SLS(27,1)=0
    SLS(28,1)=0
    SLS(29,1)=0
    SPB1 25,29,10,1,1,1
    If lastcap2 = 15 or lastcap3=15 Then CapisDenied : exit Sub
    CapisGood
    Exit Sub
  End If


  If SC(PN,16)>4 Then
    Lastcap=16
    SC(PN,16)=0
    CaptureCheck
    If SC(PN,27)=0 Then SLS(30,1)=1 : SC(PN,16)=1 : Else SLS(30,1)=0
    SLS(31,1)=0
    SLS(32,1)=0
    SLS(33,1)=0
    SLS(34,1)=0
    SPB1 30,34,10,1,1,1
    If lastcap2 = 16 or lastcap3=16 Then CapisDenied : exit Sub
    CapisGood
    Exit Sub
  End If

  If SC(PN,13)>4 Then
    Lastcap=13
    SC(PN,13)=0
    CaptureCheck
    If SC(PN,27)=0 Then SLS( 7,1)=1 : SC(PN,13)=1 : Else SLS( 7,1)=0
    SLS( 8,1)=0
    SLS( 9,1)=0
    SLS(10,1)=0
    SLS(11,1)=0
    SPB1  7,11,10,1,1,1
    If lastcap2 = 13 or lastcap3=13 Then CapisDenied : exit Sub
    CapisGood
    Exit Sub
  End If

  If SC(PN,12)>4 Then
    Lastcap=12
    SC(PN,12)=0
    CaptureCheck
    If SC(PN,27)=0 Then SLS(15,1)=1 : SC(PN,12)=1 : Else SLS(15,1)=0
    SLS(16,1)=0
    SLS(17,1)=0
    SLS(18,1)=0
    SLS(19,1)=0
    SPB1 15,19,10,1,1,1
    If lastcap2 = 12 or lastcap3=12 Then CapisDenied : exit Sub
    CapisGood
    Exit Sub
  End If

End Sub


Sub CapisDenied

  DOF 155,2
  If DOFdebug=1 Then debug.print "DOF 155,2 capture DENIED"
  SPB1 86,88,3,12,0,1 ' towerlights x3 long blinks
  dtbg_timer

  SPB1 73,74,5,2,1,1 ' 2 smal flags

  If UMainDMD.isrendering Then UMainDMD.cancelrendering

  Scoring 10000,1000
  FlexFlashText2scl "CAPTURE" , " DENIED" , 100
  LiftTopRedLight.timerenabled=1 ' 1sec delay denied
End Sub

Sub LiftTopRedLight_Timer
  LiftTopRedLight.timerenabled=0
  PlaySound "Denied", 1, 1.5*BG_Volume, 0, 0,0,0, 0, 0
End Sub

Sub CapisGood
  lastcap3=lastcap2
  lastcap2=lastcap
  'E153 2 Capture
  'E154 2 ENEMY Capture
  DOF 153,2
  If DOFdebug=1 Then debug.print "DOF 153,2 capture"

  SPB1 86,88,10,1,0,1 ' towerlights x10 blinks

  dtbg_timer
  ' 28= map score left Counter
  ' 29= map score Right Counter
  If SC(PN,28)=SC(PN,29) Then TAKENLEAD.enabled=1
  SC(PN,28)=SC(PN,28)+1
  RedDisplay SC(PN,28)
  SPB1 71,71,7,2,1,1
  SPB1 73,74,5,2,1,1 ' 2 smal flags

  If UMainDMD.isrendering Then UMainDMD.cancelrendering

  Scoring 50000,5000
  FFT=FFTscore
  FlexFlashText2cap "CAPTURED" , FFT , 80

  PlaySound "Capture", 1, BG_Volume, 0, 0,0,0, 0, 0
  PlaySound "AACaptureSound" & int(rnd(1)*3)+1 , 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0


End Sub



Sub CaptureCheck
  If SC(PN,12)>4 Then SPB1 42,42,5,2,2,1 : SLS(42,1)=2 : If SLS(41,1)>1 Then SLS(42,1)=1
  If SC(PN,13)>4 Then SPB1 42,42,5,2,2,1 : SLS(42,1)=2 : If SLS(41,1)>1 Then SLS(42,1)=1
  If SC(PN,14)>4 Then SPB1 42,42,5,2,2,1 : SLS(42,1)=2 : If SLS(41,1)>1 Then SLS(42,1)=1
  If SC(PN,15)>4 Then SPB1 42,42,5,2,2,1 : SLS(42,1)=2 : If SLS(41,1)>1 Then SLS(42,1)=1
  If SC(PN,16)>4 Then SPB1 42,42,5,2,2,1 : SLS(42,1)=2 : If SLS(41,1)>1 Then SLS(42,1)=1
End Sub



'**********************  RIGHT KICKER ****************************



Sub Sw32_hit  ' kicker 3
  TrippleHS=0
    DOF 137,2
    If DOFdebug=1 Then debug.print "DOF 137,2 right saucer"
'133 invisibility awarded
'134 left saucer
'135 left orbit
'136 TopVUK
'137 right saucer
'138 right orbit
  If Invisiblilty(4)=1 Then
    Invisiblilty(1)=0
    Invisiblilty(2)=0
    Invisiblilty(3)=0
    Invisiblilty(4)=1
    Invisiblilty(5)=0
    SLS (78,1)=0
    SLS (79,1)=0
    SLS (80,1)=0
    SLS (81,1)=1
    SLS (82,1)=0
    SLS (83,1)=0
    SPB1 78,82,2,2,0,1
  Else
    Invisiblilty(4)=1
    SLS (81,1)=1
    SPB1 81,81,5,2,0,1
  End If

  If Firstblood=0 Then FirstBloodX
  HeadShotON
  FastFrags=FastFrags+1
  Frag

  SC(PN,19)=SC(PN,19)+1


  CheckFastFrags

  If SLS(48,1)>1 Then
    ' turn off enemyflag  l046 1500ms
    L041.timerenabled=1
    SLS(41,1)=0
    If SLS(42,1)=1 Then SLS(42,1)=2
    SLS(48,1)=0
    SLS(63,1)=0
    SLS(70,1)=0
    FlexNewText "Flag returned!"
    Scoring 5000,1000
  End If

  EnemyFlag
  GameTimerCheck


    ' check to see if lock lights should be turned on
    If MBactive=0 Then
      SC(PN,18) = SC(PN,18) + 1   ' 18 is lights done
      If SC(PN,18)=7 Then
        DividerWall1_Timer ' divider=on ramp goes to tower
        PlaySound "AApointsecure" & Int(rnd(1)*4)+1 , 1, BG_Volume, 0, 0,0,0, 0, 0
        FlexFlashText2 "Point 1","Secure" , 80
        SLS(65,1)=2
        SLS(66,1)=2
      LockBolt1.Z=140
      LockBolt1.collidable=True
debug.Print "Boltmoved UP"
      End If
    End If



  SC(PN,16) = SC(PN,16) + 1
  If SC(PN,16) =5 Then
    '!Turn on capture!
    SLS(42,1)=2 ' blinking
    If SLS(41,1)>1 Then SLS(42,1)=1
    SPB1 42,42,7,2,2,1
    PlaySound "wdend04" , 1, BG_Volume, 0, 0,0,0, 0, 0
    If UMainDMD.isrendering Then UMainDMD.cancelrendering
    FlexFlashText3 "Capture","Lit",80
    PlaySound "AAassist", 1, BG_Volume, 0, 0,0,0, 0, 0
  elseif SC(PN,16) > 4 Then
    SC(PN,16) = 5

  End If


  Select Case SC(PN,16)
    Case 1: SLS(30,1)=1
    Case 2: SLS(31,1)=1
    Case 3: SLS(32,1)=1
    Case 4: SLS(33,1)=1
        PlaySound "AASpreeSound", 1, BG_Volume, 0, 0,0,0, 0, 0
    Case 5: SLS(34,1)=1
    Scoring 1000,100
  End Select

  SoundSaucerLock()
  SPB1 30,34,3,2, 7,1  ' light 7,specialstate, 3 blinks,timer,special off after done
  SLS(180,2)=33 ' blink middletopgi  1 solid blink ``?
  Objlevel(3) = 1 : FlasherFlash3_Timer : playSound "fx_relay", 1,0.3*VolumeDial, AudioPan(Sling1),0,0,0,0, AudioPan(Sling1)
            DOF 112,2
            If DOFdebug=1 Then debug.print "DOF 112,2 flasherb2"
  If SLS(68,1)>0 Then
  DOF 144,1
  If DOFdebug=1 Then debug.print "DOF 144,1 multiball on"
'E144 multiball
'E145 godlike
'E146 monsterkill
    PlaySound "objectivedest" & int(rnd(1)*5)+1 , 1, 1.5* BG_Volume, 0, 0,0,0, 0, 0
    If UMainDMD.isrendering Then UMainDMD.cancelrendering
    FlexFlashText3 "OBJECTIVE","DESTROYD",100
    Scoring 10000,2000
    L068.timerenabled=1
    ' start multiball
    SPB1 65,66,6,2,0,1
    SPB1 68,68,6,2,0,1
    SLS(68,1)=0
    For i = 179 to 190
    SLS(i,2)=700  ' blink for 230 updates
    Next

    Respawn=Respawn+25 ' 20 sec respawn timer
      SPB1 120,139,5,1,3,1
      For i = 120 to 139
        SLS(i,1)=1
      Next
    SLS(64,1)=2 ' respawn blinks
    SLS(64,3)=6
'   L064.timerenabled=1

  Else
    Sw32.timerenabled=1
  End If
End Sub

Sub Sw32_Timer
  Sw32.Kick 270,10
  SoundSaucerKick 1, Sw32
  PlaySound "AAMini_Fire", 1,1.5*BG_Volume, 0, 0,0,0, 0, 0
  sw32.timerenabled=0
End Sub

Dim MBalls,LockingBalls
Sub L068_timer
  MBalls=Mballs+1
  MBActive=1

  Select Case MBalls
   Case 1 :   SPB1 1,68,1,2,5,1
   Case 2 :   Sw32.Kick 270,10
        SoundSaucerKick 1, Sw32
        PlaySound "AAMini_Fire", 1, 0.6 * BG_Volume, 0, 0,0,0, 0, 0

   Case 3 : SLS(45,1)=2
        SPBhalfandhalf 6
        SPB1 86,88,20,0,0,1


        If Team(PN)=0 Then
          PlaySound "RedLeader" & int(rnd(1)*5)+1 ,1 , 1.5*BG_Volume, 0, 0,0,0, 0, 0
          FlexFlashText1 "RED LEADER",100
        End If

        If Team(PN)=1 Then
          PlaySound "BlueLeader" & int(rnd(1)*4)+1 ,1 , 1.5*BG_Volume, 0, 0,0,0, 0, 0
          FlexFlashText1 "BLUE LEADER",100
          End If

        If Team(PN)=2 Then
          PlaySound "GoldLeader" & int(rnd(1)*5)+1 ,1 , 1.5*BG_Volume, 0, 0,0,0, 0, 0
          FlexFlashText1 "GOLD LEADER",100
        End If

        If Team(PN)=3 Then
          PlaySound "GreenLeader" & int(rnd(1)*5)+1 ,1 , 1.5*BG_Volume, 0, 0,0,0, 0, 0
          FlexFlashText1 "GREEN LEADER",100
        End If

   Case 4 :     SPBhalfandhalf 6
          If Locked1=1 Then
          Locked1=0
          LockBolt1.Z=90
          LockBolt1.collidable=False
          debug.Print "Boltmoved Down"
        Else
          RespawnWaiting=RespawnWaiting+1
          DOF 151,2
          If DOFdebug=1 Then debug.print "DOF 151,2 ballrelease"
          PlaySound "AARespawn", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
'         RandomSoundBallRelease(BallRelease)
        End If
  '     SPB1 1,68,3,2,1,1


   Case 6 :   If Locked2=1 Then
          Locked2=0
          LockBolt1.Z=90
          LockBolt1.collidable=False
          debug.Print "Boltmoved Down"

        Else
          RespawnWaiting=RespawnWaiting+1
      DOF 151,2
      If DOFdebug=1 Then debug.print "DOF 151,2 ballrelease"
          PlaySound "AARespawn", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
'         RandomSoundBallRelease(BallRelease)
        End If
        SPB1 1,68,3,2,1,1

   Case 5 :   SPBhalfandhalf 6
        PlaySound "Assault" & int(rnd(1)*4)+1,1 , BG_Volume*1.5, 0, 0,0,0, 0, 0
        FlexFlashText2 "ASSAULT","THE BASE",100
        SPB1 86,88,10,0,0,1

   Case 7 : SPBhalfandhalf 6
        MBalls=0
        SPB1 45,45,5,2,1,1
        L068.timerenabled=0
        PlaySound "engage" & int(rnd(1)*5)+1,1 , BG_Volume*1.5, 0, 0,0,0, 0, 0
        FlexFlashText2 "ENGAGE","ENEMY",100
        SPB1 86,88,10,0,0,1
  End Select
End Sub


'E139 mainramp done
'E140 redeemer
'E141 Jackpot
'E142 HeadShot  done
Sub RampR_fx2_unhit
    SPBblueside
    SPB1 99,99,1,10,0,1
    SPB1 86,88,1,1,0,1

  If KBstatus=0 Then LiteKBstatus=1 : SLS(47,1)=2

  DOF 139,2
  If DOFdebug=1 Then debug.print "DOF 139,2 mainramp"
  If Invisiblilty(1)=1 And Invisiblilty(2)=1 And Invisiblilty(3)=1 And Invisiblilty(4)=1 And Invisiblilty(5)=1 Then
    SPB1 78,83,10,2,5,1
    Respawn=Respawn+60
      SPB1 120,139,5,1,3,1
      For i = 120 to 139
        SLS(i,1)=1
      Next
    SLS(64,1)=2
    SLS(64,3)=6
    Scoring 200000,30000
    FFT=FFTscore
    FlexFlashText1 "INVISIBILITY",66
    PlaySound "CloakOn", BG_Volume*1.5, 0, 0,0,0, 0, 0

    DOF 133,2
    If DOFdebug=1 Then debug.print "DOF 133,2 invisibility awarded"

'133 invisibility awarded
    SPB1 79,82,3,2,0,1
    Invisiblilty(1)=0
    Invisiblilty(2)=0
    Invisiblilty(3)=0
    Invisiblilty(4)=0
    Invisiblilty(5)=0
    SLS (78,1)=0
    SLS (79,1)=0
    SLS (80,1)=0
    SLS (81,1)=0
    SLS (82,1)=0
  End If






  If  SS2=1 Then
    SS2=0
    Debug.print "SS2 awarded"
    L059.timerenabled=0
    Skillshot2
  End If


  If Firstblood=0 Then FirstBloodX

  SC(PN,19)=SC(PN,19)+1
  CheckFastFrags

  If SLS(44,1)>1 Then
    HeadShot
    i = int(rnd(1)*3)
    If SC(PN,12+i)<5 Then
      SC(PN,12+i)=SC(PN,12+i)+1
    elseIf SC(PN,13+i)<5 then
      SC(PN,13+i)=SC(PN,13+i)+1
    elseIf SC(PN,14+i)<5 then
      SC(PN,14+i)=SC(PN,14+i)+1
    End If
    If SLS(67,1)=0 Then EnemyFlag

    SetlightsRedeemer
  End If
  TrippleHS=0

  HeadShotON
  FastFrags=FastFrags+1
  Frag
  If SLS(45,1)>1 Then
    Jackpot
  DOF 141,2
  If DOFdebug=1 Then debug.print "DOF 141,2 jackpot"
  End If

' If SLS(65,1)=0 Then
  If locked1=1 or locked2=1 Then
      LockBolt1.Z=90
      LockBolt1.collidable=False
    debug.print "BOLT DOWN NORMALRAMP"
  End If
' End If  ' this lets out 1 ball from locking station, if another is on the way from normal ramp


  Objlevel(9) = 1 : FlasherFlash9_Timer : playSound "fx_relay", 1,0.3*VolumeDial, AudioPan(Flasherbase9),0,0,0,0, AudioPan(Flasherbase9)
  DOF 110,2
  If DOFdebug=1 Then debug.print "DOF 110,2 flasher1"
End Sub

'**********************************************************************************************************
'* DROPTARGETS *
'**********************************************************************************************************
'  5=DT1 Sw25
'  6=DT2 Sw26
'  7=DT3 Sw27
'  8=DT4 Sw33
'  9=DT5 Sw34
' 10=DT6 Sw35
Sub Sw25_hit 'DT1
  'E128 middle DT
  'E129 right DT
  DOF 128,2
  If DOFdebug=1 Then debug.print "DOF 128,2 1/3 middle DT"
  MapBlinker_Timer
  If SC(PN,5)=0 Then Scoring 1000,100
  SC(PN,5)=1 ' score (x,5) = dt nr 1
  PlaySoundAtLevelStatic SoundFX("DTDrop",DOFKnocker), 2, Sw25
  PlaySound "AAArmorUT", 1, 0.35 * BG_Volume, 0, 0,0,0, 0, 0
  sw25.timerenabled=1
  SPB1 35,37,3, 2,7,1  ' light 7,specialstate, 3 blinks,timer,special off after done
  SLS (35,1)=1
  L035.timerenabled=0
  L035.timerenabled=1
  SPB1 35,37,3,2, 7,1  ' light 7,specialstate, 3 blinks,timer,special off after done
  SLS (35,1)=1
  SLS(182,2)=33 '  blink middleDT GI
End Sub

Sub L035_Timer  ' 15 sec an armor gone
  SLS(35,1)=0
  SPB1 35,35,1,2,7,1
  L035.timerenabled=0
End Sub

Sub sw25_Timer ' 10sec
  sw25.isdropped=false
  PlaySoundAtLevelStatic SoundFX("DTReset",DOFKnocker), 2, Sw25
  sw25.timerenabled=0
  SC(PN,5)=0
End Sub




Sub Sw26_hit 'DT2
  DOF 128,2
  If DOFdebug=1 Then debug.print "DOF 128,2 2/3 middle DT"
  MapBlinker_Timer
  If SC(PN,6)=0 Then Scoring 2000,400
  SC(PN,6)=1
  PlaySoundAtLevelStatic SoundFX("DTDrop",DOFKnocker), 2, Sw26
  PlaySound "AAAmpPickup", 1, 0.35 * BG_Volume, 0, 0,0,0, 0, 0
  Sw26.timerenabled=1
  SPB1 35,37,3, 2,7,1  ' light 7,specialstate, 3 blinks,timer,special off after done
  SLS (36,1)=1
  L036.timerenabled=0
  L036.timerenabled=1
  SLS(182,2)=33 ' blink middleDT GI
End Sub

Sub L036_Timer  ' 15 sec an amp
  SLS(36,1)=0
  SPB1 36,36,1,2,7,1
  L036.timerenabled=0
  PlaySound "AmpOut", 1, 0.7 * BG_Volume, 0, 0,0,0, 0, 0
End Sub

Sub sw26_Timer ' 15sec
  Sw26.isdropped=false
  PlaySoundAtLevelStatic SoundFX("DTReset",DOFKnocker), 2, Sw26
  Sw26.timerenabled=0
  SC(PN,6)=0
End Sub




Sub Sw27_hit 'DT3
  DOF 128,2
  If DOFdebug=1 Then debug.print "DOF 128,2 3/3 middle DT"
  MapBlinker_Timer
  If SC(PN,7)=0 Then Scoring 1000,100
  SC(PN,7)=1
  PlaySoundAtLevelStatic SoundFX("DTDrop",DOFKnocker), 2, Sw27
  PlaySound "AABootSnd", 1, 0.45 * BG_Volume, 0, 0,0,0, 0, 0
  Sw27.timerenabled=1
  SPB1 35,37,3, 2,7,1  ' light 7,specialstate, 3 blinks,timer,special off after done
  SLS (37,1)=1
  L037.timerenabled=0
  L037.timerenabled=1
  SLS(182,2)=33 ' blink middleDT GI
End Sub

Sub L037_Timer  ' 15 sec shield
  SLS(37,1)=0
  SPB1 37,37,1,2,7,1
  L037.timerenabled=0
End Sub

Sub sw27_Timer ' 7sec
  Sw27.isdropped=false
  PlaySoundAtLevelStatic SoundFX("DTReset",DOFKnocker), 2, Sw27
  Sw27.timerenabled=0
  SC(PN,7)=0
End Sub




Sub Sw33_hit 'DT4
  DOF 129,2
  If DOFdebug=1 Then debug.print "DOF 129,2 1/3 right DT"
  MapBlinker_Timer
  If SC(PN,8)=0 Then Scoring 1000,200
  SC(PN,8)=1
  PlaySoundAtLevelStatic SoundFX("DTDrop",DOFKnocker), 2, Sw33
  PlaySound "AArocketload", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
  Sw33.timerenabled=1
  SPB1 38,40,3,2, 7,1  ' light 7,specialstate, 3 blinks,timer,special off after done
  SLS (38,1)=1
  L038.timerenabled=0
  L038.timerenabled=1
  SLS(184,2)=33 ' blink leftDT  GI
End Sub

Sub L038_Timer  ' 15 sec shield
  SLS(38,1)=0
  SPB1 38,38,1,2,7,1
  L038.timerenabled=0
End Sub

Sub sw33_Timer
  Sw33.isdropped=false
  PlaySoundAtLevelStatic SoundFX("DTReset",DOFKnocker), 2, Sw33
  Sw33.timerenabled=0
  SC(PN,8)=0
End Sub




Sub Sw34_hit 'DT5
  DOF 129,2
  If DOFdebug=1 Then debug.print "DOF 129,2 2/3 right DT"
  MapBlinker_Timer
  If SC(PN,9)=0 Then Scoring 1000,200

  SC(PN,9)=1
  PlaySoundAtLevelStatic SoundFX("DTDrop",DOFKnocker), 2, Sw34
  PlaySound "AAFlakload", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
  Sw34.timerenabled=1
  SPB1 38,40,3,2, 7,1  ' light 7,specialstate, 3 blinks,timer,special off after done
  SLS (39,1)=1
  L039.timerenabled=0
  L039.timerenabled=1
  SLS(184,2)=33 ' blink leftDT  GI
End Sub

Sub L039_Timer  ' 15 sec shield
  SLS(39,1)=0
  SPB1 39,39,1,2,7,1
  L039.timerenabled=0
End Sub

Sub Sw34_Timer
  Sw34.isdropped=false
  PlaySoundAtLevelStatic SoundFX("DTReset",DOFKnocker), 2, Sw34
  Sw34.timerenabled=0
  SC(PN,9)=0
End Sub




Sub Sw35_hit 'DT6
  DOF 129,2
  If DOFdebug=1 Then debug.print "DOF 129,2 3/3 right DT"
  MapBlinker_Timer
  If SC(PN,10)=0 Then Scoring 1000,200

  SC(PN,10)=1
  PlaySoundAtLevelStatic SoundFX("DTDrop",DOFKnocker), 2, Sw35
  PlaySound "AAReload", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
  sw35.timerenabled=1
  SPB1 38,40,3, 2,7,1  ' light 7,specialstate, 3 blinks,timer,special off after done
  SLS (40,1)=1
  L040.timerenabled=0
  L040.timerenabled=1
  SLS(184,2)=33 ' blink leftDT  GI
End Sub

Sub L040_Timer  ' 15 sec shield
  SLS(40,1)=0
  SPB1 40,40,1,2,7,1
  L040.timerenabled=0
End Sub

Sub sw35_Timer
  sw35.isdropped=false
  PlaySoundAtLevelStatic SoundFX("DTReset",DOFKnocker), 2, Sw35
  sw35.timerenabled=0
  SC(PN,10)=0
End Sub


'**********************************************************************************************************
'* SWITCHES *
'**********************************************************************************************************

'skillshot lights
Sub sw20_Hit 'bottom one
  MapBlinker_Timer
  PlaySound "mclang2", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
  SLS(59,1)=1
  SPB1 59,59,3,2,8,1
  CheckRedeemer
  Scoring 500,0
  DOF 152,2
  If DOFdebug=1 Then debug.print "DOF 152,2 4skillshot lights hit"

  If skill1>0 Then
    Select Case Skill1
      Case 1,2,10,11 : SkillshotClose : Exit Sub
      Case 3,4,8,9 : Skillshotgood : Exit Sub
      Case 5,6,7 : SkillShotPerfect : Exit Sub
    End Select
    SkillshotBAD
  End If

End Sub

Sub sw19_Hit
MapBlinker_Timer
  PlaySound "mclang2", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
  SLS(60,1)=1
  SPB1 60,60,3,2,8,1
  CheckRedeemer
  Scoring 500,0
  DOF 152,2
  If DOFdebug=1 Then debug.print "DOF 152,2 4skillshot lights hit"

  If skill1>0 Then
    Select Case Skill1
      Case 1,2,10,11 : SkillshotClose : Exit Sub
      Case 3,4,8,9 : Skillshotgood : Exit Sub
      Case 5,6,7 : SkillShotPerfect : Exit Sub
    End Select
    SkillshotBAD
  End If


End Sub

Sub sw18_Hit
  MapBlinker_Timer
  PlaySound "mclang2", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
  SLS(61,1)=1
  SPB1 61,61,3,2,8,1
  CheckRedeemer
  Scoring 500,0
  DOF 152,2
  If DOFdebug=1 Then debug.print "DOF 152,2 4skillshot lights hit"

  If skill1>0 Then
    Select Case Skill1
      Case 1,2,10,11 : SkillshotClose : Exit Sub
      Case 3,4,8,9 : Skillshotgood : Exit Sub
      Case 5,6,7 : SkillShotPerfect : Exit Sub
    End Select
    SkillshotBAD
  End If


End Sub

Sub sw17_Hit 'top one
  MapBlinker_Timer
  PlaySound "mclang2", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
  SLS(62,1)=1
  SPB1 62,62,3,2,8,1
  CheckRedeemer
  Scoring 500,0
  DOF 152,2
  If DOFdebug=1 Then debug.print "DOF 152,2 4skillshot lights hit"

End Sub

Sub CheckRedeemer
  If SLS(59,1)=1 And SLS(60,1)=1 And SLS(61,1)=1 And SLS(62,1)=1 Then
    SPB1 59,62,5,2,1,1 ' 5 blinks
    SLS(59,1)=0 'turn off whities
    SLS(60,1)=0
    SLS(61,1)=0
    SLS(62,1)=0

    SLS(67,1)=2 ' TURN ON REDEEMER
    FlexFlashText3 "REDEEMER","LiT",100

    PlaySound "AAlift3act", 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0
    Dividerwall1_Timer ' activate divider
    ' WHEN ENTER KICK OUT behind tower
  End If

End Sub

Dim OrbDir
Sub Sw54_hit 'roundabout bothways
  If OrbDir=4 Then
    PlaySound "AARocket_Exp", 1, 0.2 * BG_Volume, 0, 0,0,0, 0, 0
  Else
    PlaySound "AAShockalt_fire", 1, 0.2 * BG_Volume, 0, 0,0,0, 0, 0
  End If
  OrbDir=1
End Sub



'invisibilty (1) = sw40
'invisibilty (2) = 22
'invisibilty (3) = topvuk
'invisibilty (4) = 32
'invisibilty (5) = 21

Sub Sw21_Hit
  TrippleHS=0 '  turn off headshot triple
  If OrbDir=1 Then
    DOF 135,2
    If DOFdebug=1 Then debug.print "DOF 135,2 left orbit"
'133 invisibility awarded
'134 left saucer
'135 left orbit
'136 TopVUK
'137 right saucer
'138 right orbit

  If Invisiblilty(2)=1 Then
    Invisiblilty(1)=0
    Invisiblilty(2)=1
    Invisiblilty(3)=0
    Invisiblilty(4)=0
    Invisiblilty(5)=0
    SLS (78,1)=0
    SLS (79,1)=1 ' award first one and just delete all others
    SLS (80,1)=0
    SLS (81,1)=0
    SLS (82,1)=0
    SLS (83,1)=0
    SPB1 78,82,2,2,0,1
  Else
    Invisiblilty(2)=1
    SLS (79,1)=1
    SPB1 79,79,5,2,0,1
  End If


    OrbDir=3
    SPB1 111,101,1,2,3,1
    If Firstblood=0 Then FirstBloodX
    HeadShotON
    FastFrags=FastFrags+1
    Frag

    SC(PN,19)=SC(PN,19)+1
    CheckFastFrags


    ' check to see if lock lights should be turned on
    If MBactive=0 Then
      SC(PN,18) = SC(PN,18) + 1   ' 18 is lights done
      If SC(PN,18)=7 Then
        DividerWall1_Timer ' divider=on ramp goes to tower
        PlaySound "AApointsecure" & Int(rnd(1)*4)+1 , 1, BG_Volume, 0, 0,0,0, 0, 0
        FlexFlashText2 "Point 1","Secure" , 80
        SLS(65,1)=2
        SLS(66,1)=2
      LockBolt1.Z=140
      LockBolt1.collidable=True
debug.Print "Boltmoved UP"
      End If
    End If

    EnemyFlag
    GameTimerCheck







    SC(PN,12) = SC(PN,12) + 1
    If SC(PN,12) = 5 Then
      '!Turn on capture!
      SLS(42,1)=2 ' blinking
      If SLS(41,1)>1 Then SLS(42,1)=1
      SPB1 42,42,7,2,2,1
      PlaySound "wdend04" , 1, BG_Volume, 0, 0,0,0, 0, 0
      If UMainDMD.isrendering Then UMainDMD.cancelrendering
      FlexFlashText3 "Capture","Lit",80
      PlaySound "AAassist", 1, BG_Volume, 0, 0,0,0, 0, 0
    elseif SC(PN,12) > 4 Then
      SC(PN,12) = 5
    End If

    Select Case SC(PN,12)
      Case 1: SLS(15,1)=1
      Case 2: SLS(16,1)=1
      Case 3: SLS(17,1)=1
      Case 4: SLS(18,1)=1
      PlaySound "AASpreeSound", 1, BG_Volume, 0, 0,0,0, 0, 0
      Case 5: SLS(19,1)=1
      Scoring 1000,100
    End Select

    SPB1 15,19,3,2, 7,1  ' light 7,specialstate, 3 blinks,timer,special off after done
  End If

  Objlevel(2) = 1 : FlasherFlash2_Timer : playSound "fx_relay", 1,0.3*VolumeDial, AudioPan(Sling1),0,0,0,0, AudioPan(Sling1)
  DOF 112,2
  If DOFdebug=1 Then debug.print "DOF 112,2 flasherb2"
End Sub

Sub Sw55_hit 'roundabout both ways
  If OrbDir=3 Then
    PlaySound "AAShock_fire", 1, 0.2 * BG_Volume, 0, 0,0,0, 0, 0
  Else
    PlaySound "AARocket_fire", 1, 0.2 * BG_Volume, 0, 0,0,0, 0, 0
  End If
  OrbDir=2
End Sub

Sub Sw22_Hit
  TrippleHS=0

  If OrbDir=2 Then
    DOF 138,2
    If DOFdebug=1 Then debug.print "DOF 138,2 right saucer"
'133 invisibility awarded
'134 left saucer
'135 left orbit
'136 TopVUK
'137 right saucer
'138 right orbit

'invisibilty (1) = sw40
'invisibilty (2) = 22
'invisibilty (3) = topvuk
'invisibilty (4) = 32
'invisibilty (5) = 21
  If Invisiblilty(5)=1 Then
    Invisiblilty(1)=0
    Invisiblilty(2)=0
    Invisiblilty(3)=0
    Invisiblilty(4)=0
    Invisiblilty(5)=1
    SLS (78,1)=0
    SLS (79,1)=0
    SLS (80,1)=0
    SLS (81,1)=0
    SLS (82,1)=1
    SLS (83,1)=0
    SPB1 78,82,2,2,0,1
  Else
    Invisiblilty(5)=1
    SLS (82,1)=1
    SPB1 82,82,5,2,0,1
  End If

    OrbDir=4
    SPB1 101,111,1,2,3,1



    If Firstblood=0 Then FirstBloodX
    HeadShotON
    FastFrags=FastFrags+1
    Frag

    SC(PN,19)=SC(PN,19)+1
    CheckFastFrags


    ' check to see if lock lights should be turned on
    If MBactive=0 Then
      SC(PN,18) = SC(PN,18) + 1   ' 18 is lights done
      If SC(PN,18)=7 Then
        DividerWall1_Timer ' divider=on ramp goes to tower
        PlaySound "AApointsecure" & Int(rnd(1)*4)+1 , 1, BG_Volume, 0, 0,0,0, 0, 0
        FlexFlashText2 "Point 1","Secure" , 80
        SLS(65,1)=2
        SLS(66,1)=2
      LockBolt1.Z=140
      LockBolt1.collidable=True
debug.Print "Boltmoved UP"

      End If
    End If



    EnemyFlag
    GameTimerCheck
    SC(PN,13) = SC(PN,13) + 1
    If SC(PN,13) = 5 Then
      '!Turn on capture!
      SLS(42,1)=2 ' blinking
      If SLS(41,1)>1 Then SLS(42,1)=1
      SPB1 42,42,7,2,2,1
      PlaySound "wdend04", 1, BG_Volume, 0, 0,0,0, 0, 0
      If UMainDMD.isrendering Then UMainDMD.cancelrendering
      FlexFlashText3 "Capture","Lit",80
      PlaySound "AAassist", 1, BG_Volume, 0, 0,0,0, 0, 0
    ElseIf SC(PN,13) > 4 Then
      SC(PN,13) = 5
    End If


    Select Case SC(PN,13)
      Case 1: SLS( 7,1)=1
      Case 2: SLS( 8,1)=1
      Case 3: SLS( 9,1)=1
      Case 4: SLS(10,1)=1
      PlaySound "AASpreeSound", 1, BG_Volume, 0, 0,0,0, 0, 0
      Case 5: SLS(11,1)=1
      Scoring 1000,100

    End Select

    SPB1  7,11,3, 2,7,1 ' light 7,specialstate, 3 blinks,timer,special off after done
  End If

  Objlevel(3) = 1 : FlasherFlash3_Timer : playSound "fx_relay", 1,0.3*VolumeDial, AudioPan(Sling1),0,0,0,0, AudioPan(Sling1)
  DOF 109,2
  If DOFdebug=1 Then debug.print "DOF 109,2 flasher1"

End Sub




Dim Leftdrop
Sub Trigger001_Hit : Leftdrop=1 : End sub
Dim Rightdrop
Sub RampR_fx3_Hit : Rightdrop=1 : End Sub

Sub Sw29_hit
  If Leftdrop=1 Then Leftdrop=0 : RandomSoundDelayedBallDropOnPlayfield sw29

  Scoring 1000,0
  PlaySound "AAAmmoPick", 1, 0.4 * BG_Volume, 0, 0,0,0, 0, 0
End Sub


Sub Sw37_hit
  If Rightdrop=1 Then Rightdrop=0 : RandomSoundDelayedBallDropOnPlayfield sw37

  Scoring 1000,0
  PlaySound "AAAmmoPick", 1, 0.4 * BG_Volume, 0, 0,0,0, 0, 0
End Sub


Dim KBstatus
Sub Sw28_hit
  If SLS(48,1)>1 Then
  ' turn off enemyflag  l046 1500ms
    L041.timerenabled=1
    SLS(41,1)=0
    If SLS(42,1)=1 Then SLS(42,1)=2
    SLS(48,1)=0
    SLS(63,1)=0
    SLS(70,1)=0
    FlexNewText "Flag returned!"
    Scoring 5000,1000

  End If


  If KBStatus=1 Then
    If tilted=0 Then
      If L049.timerenabled=0 Then
        L049.timerenabled=1
        SPB1 49,49,20,2,8,1 ' KB 12 blinks
      End If
      kickback.Fire
      Scoring 1000,100
      DOF 118,2
      If DOFdebug=1 Then debug.print "DOF 118,2 Kickback Fire"
      kickback.pullback
      PlaySound "AAUTsuperHeal", 1, 0.3 * BG_Volume, 0, 0,0,0, 0, 0
      PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, kickback
    Else
      PlaySound "AAeliminated", 1, 0.4 * BG_Volume, 0, 0,0,0, 0, 0
      Scoring 50000,0
    End If
  End If
End Sub


Sub L041_Timer  ' return flag sound

  L041.timerenabled=0
  PlaySound "AAReturnSound", 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0

End Sub


Sub L049_Timer    '2sec then turn off KB
  L049.timerenabled=0
' L047.timerenabled=1
  KBStatus=0
    debug.print "kickback OFF"
  SLS(49,1)=0 ' KB Off
  SLS(49,5)=0 ' off with special too
End Sub


Dim LiteKBstatus
'Sub L047_Timer
' L047.timerenabled=0
' SLS(47,1)=2 ' litekickback On
' LiteKBstatus=1
'End Sub

Dim Superspi
Sub Superspinner_Timer
  Superspi=superspi+1
  l27a.uservalue=0
  L27a_timer
  PlaySound "fx_spinner"
  l27b.uservalue=0
  L27b_Timer
  If superspi=121 Then  superspinner.enabled=0 : superspi=0
End Sub


Dim BumpFlash
Sub BumperFlasher_Timer
  If BumperFlasher.enabled=0 Then
    BumperFlasher.enabled=1
    BumpFlash=0
    L34a_Timer
    L60a_Timer
    L61a_Timer
  Else
    BumpFlash=BumpFlash+1
    L34a_Timer
    L60a_Timer
    L61a_Timer
    If BumpFlash>20 Then BumperFlasher.enabled=0
  End If
End Sub


Dim pwnage
Sub sw39_hit
    PlaySound "ReturnTarget", 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0

'PWNAGE LIGHT   !!!
  If SLS(46,1)>0 Then
    DOF 130,2
    If DOFdebug=1 Then debug.print "DOF 130,2 Pwnage"
    'E130 pwnage
    'E131 lite Kickback
    SLS(46,1)=0
    SPB1 46,46,4,2,1,1
    PlaySound "ownage", 1, 0.45 * BG_Volume, 0, 0,0,0, 0, 0
    FlexFlashText1 "PWNAGE",70
    Scoring 9000+pwnage*1000,1800+pwnage*200
'   If superbumpers=16 Then FlexNewText "SUPER BUMPERS ACTIVE"
    pwnage=pwnage+1
    Select case pwnage

      Case 2: Superbumpers=16
          FlexNewText "SUPER BUMPERS ACTIVATED"
          BumperFlasher_Timer

      Case 3 : SC(PN,11) = SC(PN,11)+1
          If SC(PN,11)>5 Then
            SC(PN,11)=5  '  reward for extra bonus !!!!
            scoring 100000,1000
            FFT=FFTscore
            FlexFlashText2 "MULTI","BONUS",70
          Else
            If SC(PN,11) = 1 Then FlexFlashText3 "BONUS","X 2",100
            If SC(PN,11) = 2 Then FlexFlashText3 "BONUS","X 3",100
            If SC(PN,11) = 3 Then FlexFlashText3 "BONUS","X 4",100
            If SC(PN,11) = 4 Then FlexFlashText3 "BONUS","X 6",100
            If SC(PN,11) = 5 Then FlexFlashText3 "BONUS","X 8",100
            scoring 5000,500
          End If

          SLS(2,1)=2 : SLS(2,9)=7
          If SC(PN,11)>1 Then SLS(3,1)=2 : SLS(3,9)=7
          If SC(PN,11)>2 Then SLS(4,1)=2 : SLS(4,9)=7
          If SC(PN,11)>3 Then SLS(5,1)=2 : SLS(5,9)=7
          If SC(PN,11)>4 Then SLS(6,1)=2 : SLS(6,9)=7

      Case 5,7,9,11,13,15,17,19,21,23,25,27,29,31 :
          FlexNewText "SUPER SPINNERS 60Sec"
          superspinner.enabled=1


      Case 6 : SC(PN,11) = SC(PN,11)+1
          If SC(PN,11)>5 Then
            SC(PN,11)=5  '  reward for extra bonus !!!!
            scoring 100000,1000
            FFT=FFTscore
            FlexFlashText2 "MULTI","BONUS",70

          Else
            If SC(PN,11) = 1 Then FlexFlashText3 "BONUS","X 2",100
            If SC(PN,11) = 2 Then FlexFlashText3 "BONUS","X 3",100
            If SC(PN,11) = 3 Then FlexFlashText3 "BONUS","X 4",100
            If SC(PN,11) = 4 Then FlexFlashText3 "BONUS","X 6",100
            If SC(PN,11) = 5 Then FlexFlashText3 "BONUS","X 8",100
            scoring 5000,500
          End If

          SLS(2,1)=2 : SLS(2,9)=7
          If SC(PN,11)>1 Then SLS(3,1)=2 : SLS(3,9)=7
          If SC(PN,11)>2 Then SLS(4,1)=2 : SLS(4,9)=7
          If SC(PN,11)>3 Then SLS(5,1)=2 : SLS(5,9)=7
          If SC(PN,11)>4 Then SLS(6,1)=2 : SLS(6,9)=7



      Case 1,4,8,10,12,14,16,18,20,22,24,26,28,30,32 :
          FlexFlashText3 "ADVANCE","LIGHTS",60
          i = int(rnd(1)*3)
          If SC(PN,12+i)<5 Then
            SC(PN,12+i)=SC(PN,12+i)+1
          elseIf SC(PN,13+i)<5 then
            SC(PN,13+i)=SC(PN,13+i)+1
          elseIf SC(PN,14+i)<5 then
            SC(PN,14+i)=SC(PN,14+i)+1
          End If
          SetlightsRedeemer
    End Select
  End If


  SLS(186,2)=33 :  ' blink target GI 1 solid blink
  If LiteKBstatus=1 Then
    Objlevel(1) = 1 : FlasherFlash1_Timer : playSound "fx_relay", 1,0.3*VolumeDial, AudioPan(psw39),0,0,0,0, AudioPan(psw39)
            DOF 111,2
            If DOFdebug=1 Then debug.print "DOF 111,2 flasherR1"
    SLS(47,1)=0 ' litekickback off
    LiteKBstatus=0
    SPB1 47,47,3,2,8,1 ' lithe 3 blinks
    SLS(49,1)=1 ' KB on
    SLS(49,5)=0 'SPstateoff too
    DOF 131,2
    If DOFdebug=1 Then debug.print "DOF 131,2 Lite kickback"
    SPB1 49,49,4,2,8,1 ' KB 3 blinks
    L049.timerenabled=0
    KBstatus=1
debug.print "kickback ON"

    Scoring 1000,0
  End If

End Sub

Sub Sw36_hit
  If SLS(48,1)>1 Then
  ' turn off enemyflag  l046 1500ms
    L041.timerenabled=1 ' delaydsound only here
    SLS(41,1)=0
    If SLS(42,1)=1 Then SLS(42,1)=2
    SLS(48,1)=0
    SLS(63,1)=0
    SLS(70,1)=0
    FlexNewText "Flag returned!"

    Scoring 2500,1000


  End If

  PlaySound "AAeliminated", 1, 0.4 * BG_Volume, 0, 0,0,0, 0, 0
    Scoring 50000,0
End Sub


'**********************************************************************************************************
'*  DRAIN  *
'**********************************************************************************************************

Dim lostballs  ' for multiball to end need to loose 3


Dim RespawnWaiting,RespawnT1
Sub Respawntimer_Timer ' 600 mSwcopy


  if RespawnWaiting>0 then  debug.print "R_wait=" & RespawnWaiting

  If RespawnWaiting>0 Then
    If autokick=1 Then

      If RespawnT1=0 Then ' always wait 1200ms max
        RespawnT1=1
      ElseIf RespawnT1=1 Then
        RespawnT1=2
      Else
        RespawnT1=0
'       NExtBallGO=1
        RespawnWaiting=RespawnWaiting-1
        Ballrelease.createball
        BallRelease.kick 0,33
        SPBallup 35,12,2  ' speed up, lightstayontime, blinks

        SPB1 100,100,5,2,0,1
'E117 2 autolaunch/lauchbutton Fire
        DOF 117,2
        If DOFdebug=1 Then debug.print "DOF 117,2 EnforcerFire"
        PlaySound "Aenforcerfire", 0, 0.8 * BG_Volume, 0, 0.25
        SPB1 75,77,5,2,0,1
        L2plunger.timerenabled=1
      End If
    Else
      plungertoolong.enabled=1
      DOF 114,1
      If DOFdebug=1 Then debug.print "DOF 114,1 LaunchReady"


    End If
  End If
End Sub


dim StopAddingPlayers
Sub Drain_hit
    'E119 2 drain Hit
    'E120 2 ballsaved
    DOF 119,2
    If DOFdebug=1 Then debug.print "DOF 119,2 Drain_Hit"

    RandomSoundDrain(drain)
    Drain.destroyball
    If startgame>1 Then exit Sub
    For i = 179 to 190
      SLS(i,2)=33 ' blink GI
    Next
    If Tilted>0 Then L063.timerenabled=0 : SLS(63,1)=0 : Respawn=0

    If Respawn>0 Then
      RespawnWaiting=RespawnWaiting+1
      DOF 151,2
      If DOFdebug=1 Then debug.print "DOF 151,2 ballrelease"
      PlaySound "AARespawn", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
      FlexFlashText1tw "RESPAWN",100
      DOF 120,2
      If DOFdebug=1 Then debug.print "DOF 120,2 BALLSAVED"
    ElseIf MBActive=1 Then

      If tilted=0 Then PlaySound "AARespawn", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
      lostballs=lostballs+1

      If lostballs=2 Then
        'If SLS(65,1)=1 Then SLS(65,1)=2 : SLS(66,1)=2
        lostballs=0 ' must reset
        MBActive=0  ' turn off multiball
        DOF 144,0
        If DOFdebug=1 Then debug.print "DOF 144,0 multiball off"
        'E144 multiball

        SLS(45,1)=0 ' sniperspot offset
        If Tilted=0 Then
          SPB1 45,45,3,2,1,1 ' blinkit 3 times
          For i = 179 to 190
            SLS(i,2)=200 ' extra blinking GI
          Next
        End If
        SC(PN,18)=0 ' reset needed lights for starting lock again
        SC(PN,17)=0 ' reset lockedballs
      End If

    Else
      AutoKick=0
      If tilted=0 Then PlaySound "AAfailed", 1, 0.27 * BG_Volume, 0, 0,0,0, 0, 0

      If countdown30.enabled=1 Then
        SC(PN,33)=lastcd+1000 ' save countdown timer and stopping It
        countdown30.enabled=0
      End If
      StopAddingPlayers=1
      If Tilted=0 Then BonusTime_Timer
      SPB1 1,139,3,1,0,1
    End If

End Sub



'**********************************************************************************************************
'* BUMPERS *
'**********************************************************************************************************

Sub L60a_Timer ' BUMPER1
  If L60a.timerenabled=0 Then
    L60a.timerenabled=1
    L60a.uservalue=0

  Else
    i = L60a.uservalue
    Select Case i
      Case 0 : L60a.intensity=7  : L60b.intensity=11 : Bumper1cap.blenddisablelighting=1   : If superbumpers>15 Then  L60c.intensity=25
      Case 1 : L60a.intensity=14 : L60b.intensity=22 : Bumper1cap.blenddisablelighting=2   : If superbumpers>15 Then  L60c.intensity=55
      Case 2 : L60a.intensity=12 : L60b.intensity=18 : Bumper1cap.blenddisablelighting=1.5 : If superbumpers>15 Then  L60c.intensity=40
      Case 3 : L60a.intensity=8  : L60b.intensity=11 : Bumper1cap.blenddisablelighting=1   : If superbumpers>15 Then  L60c.intensity=25
      Case 4 : L60a.intensity=4  : L60b.intensity=6  : Bumper1cap.blenddisablelighting=0.7 : If superbumpers>15 Then  L60c.intensity=12
      Case 5 : L60a.intensity=2  : L60b.intensity=3  : Bumper1cap.blenddisablelighting=0.3 : If superbumpers>15 Then  L60c.intensity=6
      Case 6 : L60a.intensity=0  : L60b.intensity=0  : Bumper1cap.blenddisablelighting=0   : L60c.intensity=0
           L60a.timerenabled=0
    End Select
    L60a.uservalue=i+1
  End If
End Sub
Sub L34a_Timer ' BUMPER2
  If L34a.timerenabled=0 Then
    L34a.timerenabled=1
    L34a.uservalue=0
  Else
    i = L34a.uservalue
    Select Case i
      Case 0 : L34a.intensity=7  : L34b.intensity=11 : Bumper2cap.blenddisablelighting=1: If superbumpers>15 Then  L34c.intensity=25
      Case 1 : L34a.intensity=14 : L34b.intensity=22 : Bumper2cap.blenddisablelighting=2: If superbumpers>15 Then  L34c.intensity=50
      Case 2 : L34a.intensity=12 : L34b.intensity=18 : Bumper2cap.blenddisablelighting=1.5: If superbumpers>15 Then  L34c.intensity=40
      Case 3 : L34a.intensity=8  : L34b.intensity=11 : Bumper2cap.blenddisablelighting=1: If superbumpers>15 Then  L34c.intensity=25
      Case 4 : L34a.intensity=4  : L34b.intensity=6  : Bumper2cap.blenddisablelighting=0.7: If superbumpers>15 Then  L34c.intensity=12
      Case 5 : L34a.intensity=2  : L34b.intensity=3  : Bumper2cap.blenddisablelighting=0.3: If superbumpers>15 Then  L34c.intensity=6
      Case 6 : L34a.intensity=0  : L34b.intensity=0  : Bumper2cap.blenddisablelighting=0:  L34C.intensity=0
           L34a.timerenabled=0
    End Select
    L34a.uservalue=i+1
  End If
End Sub
Sub L61a_Timer ' BUMPER1
  If L61a.timerenabled=0 Then
    L61a.timerenabled=1
    L61a.uservalue=0
  Else
    i = L61a.uservalue
    Select Case i
      Case 0 : L61a.intensity=7  : L61b.intensity=11 : Bumper3cap.blenddisablelighting=1: If superbumpers>15 Then  L61c.intensity=25
      Case 1 : L61a.intensity=14 : L61b.intensity=22 : Bumper3cap.blenddisablelighting=2: If superbumpers>15 Then  L61c.intensity=50
      Case 2 : L61a.intensity=12 : L61b.intensity=18 : Bumper3cap.blenddisablelighting=1.5: If superbumpers>15 Then  L61c.intensity=40
      Case 3 : L61a.intensity=8  : L61b.intensity=11 : Bumper3cap.blenddisablelighting=1: If superbumpers>15 Then  L61c.intensity=25
      Case 4 : L61a.intensity=4  : L61b.intensity=6  : Bumper3cap.blenddisablelighting=0.7: If superbumpers>15 Then  L61c.intensity=12
      Case 5 : L61a.intensity=2  : L61b.intensity=3  : Bumper3cap.blenddisablelighting=0.3: If superbumpers>15 Then  L61c.intensity=6
      Case 6 : L61a.intensity=0  : L61b.intensity=0  : Bumper3cap.blenddisablelighting=0: L61c.intensity=0
           L61a.timerenabled=0
    End Select
    L61a.uservalue=i+1
  End If
End Sub



'**********************************************************************************************************
'* LOCK BALLS *
'**********************************************************************************************************

Sub LockThisBall1_Timer
' SPB1 1,139,3,12,0,1 ' megablinks


  LockThisBall1.enabled=0
  SC(PN,17)=SC(PN,17)+1
  If locked1=0 Then
    locked1=1
  elseIf Locked2=0 Then
    Locked2=1
  End If
  If SC(PN,17)=2 Then
    Locked2=1
    'turn off 2 locklights and start last one in kiicker 3
    SLS(65,1)=0
    SLS(66,1)=0
    SPB1 65,66,3,2,1,1  ' blinks,timer,special off after done
    SLS(68,1)=2 '  last for multiball set ! kicker 3?
'   SPB1 68,68,7,2,1,1  ' 7blinks
  Else
    FlexFlashText2 "Point 2","Secure" , 80
  End If
  kicker001.kick 90,47
  SPBblueside

  SoundSaucerKick 1, kicker001
  Lockingthisball=1
  PlaySound "searchdestroy" & int(rnd(1)*2)+1 , 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0
End Sub

Sub LockThisBalL2_Timer
' SPB1 1,139,3,12,0,1 ' megablinks


  LockThisBall2.enabled=0
  SC(PN,17)=SC(PN,17)+1
  If locked1=0 Then
    locked1=1
  elseIf Locked2=0 Then
    Locked2=1
  End If
  If SC(PN,17)=2 Then
    Locked2=1
    FlexFlashText3 "TeamPlay","LiT" , 100  ' only bottom blink

    'turn off 2 locklights and start last one in kiicker 3
    SLS(65,1)=0
    SLS(66,1)=0
    SPB1 65,66,3,2,1,1  ' blinks,timer,special off after done
    SLS(68,1)=2 '  last for multiball set ! kicker 3?
'   SPB1 68,68,7,2,1,1  ' 7blinks
  Else
    FlexFlashText2 "Point 2","Secure" , 80

  End If

  TOPVUK.destroyball
  kicker001.createball
  kicker001.kick 90,47
  SPBblueside

  SoundSaucerKick 1, kicker001

  Lockingthisball=1
  PlaySound "searchdestroy" & int(rnd(1)*2)+1 , 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0
End Sub

' wait for this before knocking out from ballrelease
Dim Lockingthisball
Sub sw003_hit
  If Lockingthisball=1 Then
    Lockingthisball=0

    If SLS(67,1)>0 Or SLS(65,1)>0 Then DividerWall1_Timer
    AutoKick=0
    RespawnWaiting=RespawnWaiting+1
      DOF 151,2
      If DOFdebug=1 Then debug.print "DOF 151,2 ballrelease"
    L062.timerenabled=0
    L062_Timer
'   RandomSoundBallRelease(BallRelease)
    PlaySound "AARespawn", 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
    FlexFlashText3 "READY","SET GO",70

  End If
End Sub



Sub L044_Timer
  L044.timerenabled=0
  SLS(44,1)=0
  FastFrags=0 '  turn off mmmonsterkill
  FFlevel=0
End Sub


Sub sw001_Hit
  If locked1>0 or Locked2>0 Then
      LockBolt1.Z=140
      LockBolt1.collidable=True
    debug.print "BOLT UP ball goes over bolt"

  End If
End Sub

'Dim Bolt1Dir
'Sub Bolt1down_Timer'
' If Bolt1Down.enabled=0 Then
'   bolt1down.enabled=1
'     LockBolt1.Z=115
'debug.Print "Boltmiddle"
' Else
'   If Bolt1Dir=0 Then
'     LockBolt1.Z=90
'     LockBolt1.collidable=False
'debug.Print "Boltmoved DOWN"
'debug.print "LOCKED1="&locked1&"LOCKED2="&locked2
'   End If
'   If Bolt1Dir=1 Then
'     LockBolt1.Z=140
'     LockBolt1.collidable=True
'debug.Print "Boltmoved UP"
'debug.print "LOCKED1="&locked1&"LOCKED2="&locked2
'
'
'   End If
'   Bolt1Down.enabled=0
' End If
'End Sub



'**********************************************************************************************************
'* BUMPERS *
'**********************************************************************************************************
'E105 2 top bumper left ( blue )
'E106 2 top bumper center ( red )
'E107 2 top bumper right ( yellow )
Sub Bumper1_Hit
  DOF 105,2
  If DOFdebug=1 Then debug.print "DOF 105,2 bumper1left"

  If Tilted=0 Then

    If superbumpers>15 Then Scoring 1000,0 : Else : Scoring 100,0

    MapBlinker_Timer
    RandomSoundBumperTop Bumper1
    L60a_Timer
    SLS(179,2)=33 ' blink middletopgi  1 solid blink ``?

  End If
End Sub

Sub Bumper2_Hit
  DOF 107,2
  If DOFdebug=1 Then debug.print "DOF 107,2 bumper3right"
  If Tilted=0 Then

    If superbumpers>15 Then Scoring 1000,0 : else : Scoring 100,0

    MapBlinker_Timer
    RandomSoundBumperMiddle Bumper2
    L34a_Timer
    SLS(179,2)=33 ' blink middletopgi  1 solid blink ``?

  End If
End Sub

Sub Bumper3_Hit
  DOF 106,2
  If DOFdebug=1 Then debug.print "DOF 106,2 bumper1bottom"
  If Tilted=0 Then
    If superbumpers>15 Then Scoring 1000,0 : Else : Scoring 100,0

    MapBlinker_Timer
    RandomSoundBumperBottom Bumper3
    L61a_Timer
    SLS(179,2)=33 ' blink middletopgi  1 solid blink ``?

  End If
End Sub



'**********************************************************************************************************
'* SPINNERS *
'**********************************************************************************************************


Sub Sw47_Spin
'E121 2 Left Spinner
'E122 2 Right spinner
  DOF 121,2
  If DOFdebug=1 Then debug.print "DOF 121,2 left spinner"
  PlaySound "fx_spinner"
  l27a.uservalue=0
  L27a_timer
  If Superspinner.enabled=1 Then scoring 3000,0 : Else :  Scoring 200,0
End sub

Sub Sw48_Spin
  DOF 122,2
  If DOFdebug=1 Then debug.print "DOF 122,2 Right spinner"
  PlaySound "fx_spinner"
  l27b.uservalue=0
  L27b_Timer
  If Superspinner.enabled=1 Then scoring  3000,0 : Else :   Scoring 200,0
 End sub

Sub L27a_timer
  If l27a.timerenabled=0 Then l27a.timerenabled=1
  i=l27a.uservalue
  Select Case i
    Case 0 : l27a.intensity=12 : flasher27a.opacity=600  : Redbasesign.blenddisablelighting=0.5 : RedBaseLight.blenddisablelighting=4
    Case 1 : l27a.intensity=20 : flasher27a.opacity=1000 : Redbasesign.blenddisablelighting=1   : RedBaseLight.blenddisablelighting=8
    Case 2 : l27a.intensity=17 : flasher27a.opacity=800  : Redbasesign.blenddisablelighting=0.8 : RedBaseLight.blenddisablelighting=7
    Case 3 : l27a.intensity=14 : flasher27a.opacity=600  : Redbasesign.blenddisablelighting=0.6 : RedBaseLight.blenddisablelighting=5
    Case 4 : l27a.intensity=11 : flasher27a.opacity=400  : Redbasesign.blenddisablelighting=0.4 : RedBaseLight.blenddisablelighting=3
    Case 5 : l27a.intensity=8  : flasher27a.opacity=300  : Redbasesign.blenddisablelighting=0.2 : RedBaseLight.blenddisablelighting=2
    Case 6 : l27a.intensity=4  : flasher27a.opacity=200  : Redbasesign.blenddisablelighting=0.1 : RedBaseLight.blenddisablelighting=1
    Case 7 : l27a.intensity=0  : flasher27a.opacity=0    : Redbasesign.blenddisablelighting=0   : RedBaseLight.blenddisablelighting=0
        l27a.timerenabled=0
  End Select
  l27a.uservalue=i+1
End Sub

Sub L27b_timer
  If l27b.timerenabled=0 Then l27b.timerenabled=1
    i=l27b.uservalue
    Select Case i
      Case 0 : l27b.intensity=17 : flasher27b.opacity=800  : Bluebasesign.blenddisablelighting=1   : BlueBaseLight.blenddisablelighting=5
      Case 1 : l27b.intensity=30 : flasher27b.opacity=1600 : Bluebasesign.blenddisablelighting=2   : BlueBaseLight.blenddisablelighting=10
      Case 2 : l27b.intensity=25 : flasher27b.opacity=1300 : Bluebasesign.blenddisablelighting=1.6 : BlueBaseLight.blenddisablelighting=7
      Case 3 : l27b.intensity=20 : flasher27b.opacity=1000 : Bluebasesign.blenddisablelighting=1.2 : BlueBaseLight.blenddisablelighting=5
      Case 4 : l27b.intensity=15 : flasher27b.opacity=700  : Bluebasesign.blenddisablelighting=0.8 : BlueBaseLight.blenddisablelighting=3
      Case 5 : l27b.intensity=10 : flasher27b.opacity=400  : Bluebasesign.blenddisablelighting=0.4 : BlueBaseLight.blenddisablelighting=2
      Case 6 : l27b.intensity=5  : flasher27b.opacity=200  : Bluebasesign.blenddisablelighting=0.2 : BlueBaseLight.blenddisablelighting=1
      Case 7 : l27b.intensity=0  : flasher27b.opacity=0    : Bluebasesign.blenddisablelighting=0   : BlueBaseLight.blenddisablelighting=0
          l27b.timerenabled=0
    End Select
    l27b.uservalue=i+1
End Sub

'**********************************************************************************************************
'* FLIPPER INFO *
'**********************************************************************************************************

dim Flippinfo
Sub FlipperInfo
  Flippinfo=Flippinfo+1

  If tilted>0 Then LFinfo=0 : RFinfo=0 : Exit Sub

  Select case Flippinfo

    Case  1 : FlexFlashText3 "BONUS",SC(PN,0),105
          PlaySound "AAUTHealth", 1, 0.2 * BG_Volume, 0, 0,0,0, 0, 0
    Case  3 : If SC(PN,18) < 7 Then
          FlexFlashText3 "M/B START IN",7 - SC(PN,18),105
        Else
          FlexFlashText3 "MULTIBALL","ACTIVE",105
        End If
    Case 5 : IF SC(PN,19) < 10 Then
          FlexFlashText3 "E/LIFE LIT IN",10 - SC(PN,19),105
        Else
          FlexFlashText3 "E/LIFE","GOTTEN",105
        End If
    Case 7: FlexFlashText3 "PWNAGE",Pwnage,85
    Case 9: FlexFlashText3 "PWN 3+6=","BONUSX UP",85


    Case 11 : FlexFlashText3 "GET BACK","IN THERE",95
    Case 13 : Flippinfo=0
  End select

End Sub
'     If SC(PN,27)=0 then FlexFlashText3 "DO NOT IGNORE","ENEMY FLAGS",70
'     If SC(PN,27)=1 then FlexFlashText3 "READY","SET GO",70
'     If SC(PN,27)=2 then FlexFlashText3 "READY","SET GO",70
'     If SC(PN,27)=3 then FlexFlashText3 "READY","SET GO",70
' 0=coret   1=november   2=dreary    3=face





'**********************************************************************************************************
'* RESPAWN TIMER 1000ms *
'**********************************************************************************************************


Dim Respawn
Sub L064_Timer

  If LFinfo>0 Then LFinfo=LFinfo+1
  If RFinfo>0 Then RFinfo=RFinfo+1
  If LFinfo=0 And RFinfo=0 Then Flippinfo=0
  If LFinfo>5 Or RFinfo>5 Then FLipperInfo

  Respawn=Respawn-1
  Select case Respawn
    case  1 : SLS(120,1)=0 :
    case  2 : SLS(121,1)=0 : SPB1 120,120,3,0,0,1
    case  3 : SLS(122,1)=0 : SPB1 121,121,3,0,0,1
    case  4 : SLS(123,1)=0 : SPB1 122,122,3,0,0,1
    case  5 : SLS(124,1)=0 : SPB1 123,123,3,0,0,1
    case  6 : SLS(125,1)=0 : SPB1 124,124,3,0,0,1
    case  7 : SLS(126,1)=0 : SPB1 125,125,3,0,0,1
    case  8 : SLS(127,1)=0 : SPB1 126,126,3,0,0,1
    case  9 : SLS(128,1)=0 : SPB1 127,127,3,0,0,1
    case 10 : SLS(129,1)=0 : SPB1 128,128,3,0,0,1
    case 11 : SLS(130,1)=0 : SPB1 129,129,3,0,0,1
    case 12 : SLS(131,1)=0 : SPB1 130,130,3,0,0,1
    case 13 : SLS(132,1)=0 : SPB1 131,131,3,0,0,1
    case 14 : SLS(133,1)=0 : SPB1 132,132,3,0,0,1
    case 15 : SLS(134,1)=0 : SPB1 133,133,3,0,0,1
    case 16 : SLS(135,1)=0 : SPB1 134,134,3,0,0,1
    case 17 : SLS(136,1)=0 : SPB1 135,135,3,0,0,1
    case 18 : SLS(137,1)=0 : SPB1 136,136,3,0,0,1
    case 19 : SLS(138,1)=0 : SPB1 136,136,3,0,0,1 : SLS(139,1)=0
    case 20 : SLS(139,1)=0 : SPB1 136,136,3,0,0,1
  end Select


  If respawn<=0 Then
    respawn=0
'   L064.timerenabled=0
    SLS(64,1)=0 ' respawn off
'   If extraball=2 Then SLS(64,1)=1
  End If

End Sub




'**********************************************************************************************************
'*  KEYS / Flippers *
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)

  If keycode = 19 then ScoreCard=1 : CardTimer.enabled=1

  If keycode = PlungerKey or keycode = LockBarKey Then
    If AutoKick=0 and RespawnWaiting>0  Then
      AutoKick=1
      L061.timerenabled=0
      L061.timerenabled=1 ' activate Skillshotturnoff 2500ms
      Gunflash_Timer
      RespawnWaiting=RespawnWaiting-1
      BallRelease.createball
      BallRelease.kick 0,33
      SPBallup 40,10,1  ' speed up, lightstayontime, blinks
      plungertoolong.enabled=0
      DOF 117,2
      If DOFdebug=1 Then debug.print "DOF 117,2 EnforcerFire"
      DOF 114,0
      If DOFdebug=1 Then debug.print "DOF 114,0 AutoLaunchOn"
      PlaySound "Aenforcerfire", 0, 0.8 * BG_Volume, 0, 0.25
      SPB1 75,77,5,2,0,1
      L2plunger.timerenabled=1
      Respawn=20 ' L064   ' respawn timer on each plunger that is not autokicker
      SPB1 120,139,5,1,3,1
      For i = 120 to 139
        SLS(i,1)=1
      Next
      If UMainDMD.isrendering Then UMainDMD.cancelrendering
      SLS(64,1)=2 ' respawn blinks
      SLS(64,3)=6
      L064.timerenabled=1
      'startlastcountdown if it was turned off somewhere
      If SC(pn,33)>999 Then
        lastcd=SC(PN,33)-1000
        debug.print "lastcountdownreset " & lastcd & "sec left"
        countdown30.enabled=1
      End If
    Else
      PlaySound "AEmptyGun", 0, 0.8 * BG_Volume, 0, 0.25
    End If
  End If


  If keycode = StartGameKey And checkHS>8 Then
    soundStartButton()
    If credits>0 Then
      If startgame=0 Then
        startgame=1
        GameOverTimer.enabled=0
        Attract1.enabled=0
          SPB1 1,139,4,3,0,1 ' megablinks

        DOF 115,0
        If DOFdebug=1 Then debug.print "DOF 115,0 StartbuttonOff"
        If UMainDMD.isrendering Then UMainDMD.cancelrendering
        Credits=Credits-1
        GAMESTARTER
        GamesPlayd=GamesPlayd+1
        DOF 116,2
        If DOFdebug=1 Then debug.print "DOF 116,2 Gamestarted"

      Else

        If ballinplay=1 And StopAddingPlayers=0 Then  ' more players
          Maxplayers=MaxPlayers+1
          GamesPlayd=GamesPlayd+1
          If MaxPlayers>4 Then
            MaxPlayers=4
          Else
            If UMainDMD.isrendering Then UMainDMD.cancelrendering
            Flexflashtext1 "Players = " & MaxPlayers , 70
            PlaySound "Ric" & Int(rnd(1)*3)+1, 0, 0.9 * BG_Volume, 0, 0.25
            Credits=Credits-1
          End If

        End If

      End If
    Else
      PlaySound "call4", 0, 0.8 * BG_Volume, 0, 0.25
    End If
  End If

  If startgame=1 And Tilted=0 and BonusTime.enabled=0 And GunFL=0 Then
    If keycode = LeftTiltKey   Then Nudge  90, 5 : SoundNudgeLeft()   : Tilting
    If keycode = RightTiltKey  Then Nudge 270, 5 : SoundNudgeRight()  : Tilting
    If keycode = CenterTiltKey Then Nudge   0, 3 : SoundNudgeCenter() : Tilting
  End If

    If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    DOF 127,2
    If DOFdebug=1 Then debug.print "DOF 127,2 add credits"
    Credits=Credits+1
    Attrackt1X=64 ' start attract at credits ( only works if its active )
    If credits>30 Then Credits=30
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

If Tilted=0 And Startgame=1 Then
  'nFozzy begin
  If keycode = LeftFlipperKey Then
    LFinfo=1
    DOF 101,1
    If DOFdebug=1 Then debug.print "DOF 101,1 leftflipperdown"
    'E101 0/1 LeftFlipper
    'E102 0/1 RightFlipper
    FlipperActivate LeftFlipper, LFPress
      LF.fire 'LeftFlipper.RotateToEnd
        If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  End If


  If keycode = RightFlipperKey Then
    RFinfo=1
    DOF 102,1
    If DOFdebug=1 Then debug.print "DOF 102,1 rightflipperdown"
    FlipperActivate RightFlipper, RFPress
      RF.fire 'RightFlipper.RotateToEnd
        If RightFlipper.currentangle > RightFlipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  End If
End If
  'nFozzy finish

End Sub

Dim LFinfo,RFinfo


Sub Table1_KeyUp(ByVal KeyCode)
  if keycode = 19 then ScoreCard=0

  If Tilted=0 Then

    If keycode= leftmagnasave Then

      SPB1 86,88,2,5,0,1
      MusicON  ' just random music
      Team(PN)=Team(PN)+1
      If Team(PN)>3 Then Team(PN)=0
      SetTeamColors

    End If


    If keycode= rightmagnasave Then


      SC(PN,20)=SC(PN,20)+1 ' tauting ingame
      If SC(PN,20)=4 Then SC(PN,20)=0
      Taunt(PN)=SC(PN,20)

      Taunting.enabled=1 ' on first, cause we dont want delay here
      Taunting_Timer
    End If


    'nFozzy begin
  If Startgame=1 Then
    If keycode = LeftFlipperKey Then
    LFinfo=0

      DOF 101,0
      If DOFdebug=1 Then debug.print "DOF 101,0 leftflipperup"

      TempNr=SC(PN,2)' swapping bonuslights at top when releasing flippers
      SC(PN,2)= SC(PN,3)
      SC(PN,3)= SC(PN,4)
      SC(PN,4)=tempnr
      MultiplyerUpdate


        FlipperDeActivate LeftFlipper, LFPress
        LeftFlipper.RotateToStart
        If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
          RandomSoundFlipperDownLeft LeftFlipper
        End If
        FlipperLeftHitParm = FlipperUpSoundLevel
      End If


      If keycode = RightFlipperKey Then
    RFinfo=0
        DOF 102,0
        If DOFdebug=1 Then debug.print "DOF 102,0 rightflipperup"

        TempNr=SC(PN,4)' swapping bonuslights at top when releasing flippers
        SC(PN,4)=SC(PN,3)
        SC(PN,3)=SC(PN,2)
        SC(PN,2)=tempnr
        MultiplyerUpdate

        FlipperDeActivate RightFlipper, RFPress
        RightFlipper.RotateToStart
        If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
          RandomSoundFlipperDownRight RightFlipper
        End If
        FlipperRightHitParm = FlipperUpSoundLevel
      End If
        'nFozzy finish
    End If
  End If

  If checkHS>0 and checkHS<9 Then   'fixign  enterinitial status
    If keycode = LeftFlipperKey Then
      EnterCurrent=EnterCurrent-1
      If EnterCurrent<1 Then EnterCurrent=37
    End If
    If keycode = RightFlipperKey Then
      EnterCurrent=EnterCurrent+1
      If EnterCurrent>37 Then EnterCurrent=1
    End If
    If keycode = StartGameKey and EnterCurrent>0 Then
      If len(EnteredName)<3 Then
        EnteredName=EnteredName+ mid ( "ABCDEFGHIJKLMNOPQRSTUVWXYZ-0123456789",EnterCurrent,1)
      End If
    End If
  End If
End Sub


Dim plungershake
Sub L2plunger_Timer
  If Gunshake=1 Then
    plungershake=plungershake+1

    select case plungershake
      case 1 : Nudge 0,4
      case 2 : Nudge 0,2
      case 3 : Nudge 0,1 : L2plunger.timerenabled=0 : plungershake=0
    End Select
  Else
    L2plunger.timerenabled=0
  End If
End Sub





'Bonus lights at Top
Sub sw41_unhit
  MapBlinker_Timer
  Scoring 200,0


  If SC(PN,2)=0 Then SC(PN,2)=1 : PlaySound "AANewBeep", 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0
  ' special blink state on these
  MultiplyerUpdate
  SLS(12,5)=3 : SLS(12,10)=45   ' 2 ON NORMAL 3 ON SPECIAL OVERRIDE

End Sub
Sub sw42_unhit
  MapBlinker_Timer
  Scoring 200,0

  If SC(PN,3)=0 Then SC(PN,3)=1 : PlaySound "AANewBeep", 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0
  MultiplyerUpdate
    SLS(13,5)=3 : SLS(13,10)=45

End Sub
Sub sw43_unhit
  MapBlinker_Timer
  Scoring 200,0

  If SC(PN,4)=0 Then SC(PN,4)=1 : PlaySound "AANewBeep", 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0
  MultiplyerUpdate
    SLS(14,5)=3 : SLS(14,10)=45
End Sub

Sub MultiplyerUpdate
  SLS(12,1)=SC(PN,2)
  SLS(13,1)=SC(PN,3)
  SLS(14,1)=SC(PN,4)
  ' check for all 3 and add multiplyer
  If SC(PN,2)=1 And SC(PN,3)=1 And SC(PN,4)=1 Then
    SC(PN,2) = 0
    SC(PN,3) = 0
    SC(PN,4) = 0
    SLS(12,1)=0
    SLS(13,1)=0
    SLS(14,1)=0
    SLS(12,5)=3 : SLS(12,10)=45
    SLS(13,5)=3 : SLS(13,10)=45
    SLS(14,5)=3 : SLS(14,10)=45

    SC(PN,11) = SC(PN,11)+1
    If SC(PN,11)>5 Then
      SC(PN,11)=5  '  reward for extra bonus !!!!
      PlaySound "proceed", 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0
      scoring 50000,1000
      FFT=FFTscore
      FlexFlashText2 "MULTI","BONUS",100

    Else
      PlaySound "nicecatch", 1, 0.8 * BG_Volume, 0, 0,0,0, 0, 0
      If SC(PN,11) = 1 Then FlexFlashText3 "BONUS","X 2",100
      If SC(PN,11) = 2 Then FlexFlashText3 "BONUS","X 3",100
      If SC(PN,11) = 3 Then FlexFlashText3 "BONUS","X 4",100
      If SC(PN,11) = 4 Then FlexFlashText3 "BONUS","X 6",100
      If SC(PN,11) = 5 Then FlexFlashText3 "BONUS","X 8",100
      scoring 5000,500
    End If

'   If SC(PN,11)>0 Then SLS(2,1)=2 : SLS(2,9)=7  ' blink 7 times activated upto the last one reached
                      SLS(2,1)=2 : SLS(2,9)=7
    If SC(PN,11)>1 Then SLS(3,1)=2 : SLS(3,9)=7
    If SC(PN,11)>2 Then SLS(4,1)=2 : SLS(4,9)=7
    If SC(PN,11)>3 Then SLS(5,1)=2 : SLS(5,9)=7
    If SC(PN,11)>4 Then SLS(6,1)=2 : SLS(6,9)=7


  End If
End Sub





' #####################################
' ###### Fluppers Domes V2.2 #####
' #####################################



Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.25  ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.9   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherBloomIntensity = 0.9
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20)
Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "red" :
InitFlasher 2, "red" :
InitFlasher 3, "blue"
'InitFlasher 4, "blue" :
'InitFlasher 5, "red" :
'InitFlasher 6, "blue"
InitFlasher 7, "red" :
InitFlasher 8, "blue"
InitFlasher 9, "blue" :
'InitFlasher 10, "red"
'InitFlasher 11, "white"

Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)
  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ =  atn( (tablewidth/2 - objbase(nr).x) / (objbase(nr).y - tableheight*1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ : objflasher(nr).height = objbase(nr).z + 60
  End If
  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0 : objlit(nr).visible = 0 : objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX : objlit(nr).RotY = objbase(nr).RotY : objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX : objlit(nr).ObjRotY = objbase(nr).ObjRotY : objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x : objlit(nr).y = objbase(nr).y : objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness
  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :  objflasher(nr).imageA = "domeblueflash" :  objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4)
    Case "red" : objflasher(nr).imageA = "domeredflash" :  objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4): objlight(nr).intensity = 3900
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)

  If not objflasher(nr).TimerEnabled Then objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1 : End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 50 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 50 * ObjLevel(nr)^2

'*** Ramp flash with color flashers   blue=prim034   red=prim062
  If nr=3 or nr=8 or nr=9 Then primitive034.BlendDisableLighting=10 * ObjLevel(nr)^2
  If nr=1 or nr=2 or nr=7 Then primitive062.BlendDisableLighting=10 * ObjLevel(nr)^2

  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then
    objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0
    primitive034.BlendDisableLighting=0
    primitive062.BlendDisableLighting=0
  End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub
Sub FlasherFlash5_Timer() : FlashFlasher(5) : End Sub
Sub FlasherFlash6_Timer() : FlashFlasher(6) : End Sub
Sub FlasherFlash7_Timer() : FlashFlasher(7) : End Sub
Sub FlasherFlash8_Timer() : FlashFlasher(8) : End Sub
Sub FlasherFlash9_Timer() : FlashFlasher(9) : End Sub
Sub FlasherFlash10_Timer() : FlashFlasher(10) : End Sub
Sub FlasherFlash11_Timer() : FlashFlasher(11) : End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
'E103 2 Leftslingshot
'E104 2 Rightslingshot
  DOF 104,2
  If DOFdebug=1 Then debug.print "DOF 104,2 rightslingshot"
  Scoring 50,0

  RSling.Visible = 0
  RSling1.Visible = 1
  sling1.rotx = 20
  RStep = 0
  RightSlingShot.TimerEnabled = 1
  InstructionCard001.transY=10

  If Tilted=0 Then

    SLS(58,1)=0
    SLS(57,1)=1
    SPB1 57,57,2,2,1,1

    ownage=ownage+1
    If ownage>3 Then
      ownage=0
      SLS(46,1)=1
      SPB1 46,46,5,2,1,1
    End If

    MapBlinker_Timer

    Objlevel(8) = 1 : FlasherFlash8_Timer : playSound "fx_relay", 1,0.3*VolumeDial, AudioPan(Sling1),0,0,0,0, AudioPan(Sling1)
            DOF 108,2
            If DOFdebug=1 Then debug.print "DOF 108,2 flasher1"
    RandomSoundSlingshotRight Sling1

    SLS(190,2)=33 : SLS(188,2)=33  ' blink bottom R GI for 33 updates = 1 solid blink ``?

    PlaySound "fx_relay", 1,0.1*VolumeDial,0,0,0,0,0,0
  End If
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 3 : InstructionCard001.transY=5
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0: InstructionCard001.transY=0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  DOF 103,2
  If DOFdebug=1 Then debug.print "DOF 103,2 leftslingshot"
  Scoring 50,0

  LSling.Visible = 0
  LSling1.Visible = 1
  sling2.rotx = 20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
  InstructionCard002.transY=10
  LeftSlingShot.TimerEnabled = 1

  If Tilted=0 Then

    SLS(58,1)=1
    SLS(57,1)=0
    SPB1 58,58,2,2,1,1

    ownage=ownage+1
    If ownage>3 Then
      ownage=0
      SLS(46,1)=1
      SPB1 46,46,5,2,1,1
    End If

    MapBlinker_Timer

    Objlevel(1) = 1 : FlasherFlash1_Timer : playSound "fx_relay", 1,0.3*VolumeDial, AudioPan(Sling2),0,0,0,0, AudioPan(Sling2)
            DOF 111,2
            If DOFdebug=1 Then debug.print "DOF 111,2 flasherR1"
    RandomSoundSlingshotLeft Sling2

    SLS(189,2)=33 : SLS(187,2)=33  ' blink bottom L GI for 33 updates = 1 solid blink ``?

    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    InstructionCard002.transY=10

  End If
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10 : InstructionCard002.transY=5
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0: InstructionCard002.transY=0
    End Select
    LStep = LStep + 1
End Sub




'******************************************************
'   BALL ROLLING AND DROP SOUNDS
'******************************************************
Const tnob = 5
ReDim rolling(tnob)
InitRolling
Dim isBallOnWireRamp : isBallOnWireRamp = False

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 to (tnob)
    rolling(i) = False
  Next
End Sub

Sub RollingTimer()

  Dim BOT, b
  BOT = GetBalls
  TiltedBalls = Ubound(BOT)
  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    StopSound "BallRoll_" & b
        StopSound "fx_plasticrolling" & b
    StopSound "fx_metalrolling" & b
  Next
  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then
    If startgame=3 Then RespawnWaiting=0 : Startgame=0
    If credits>0 Then
        DOF 115,1
        If DOFdebug=1 Then debug.print "DOF 115,1 Startbuttonready"
    End If
    Exit Sub
  End If
  ' play the rolling sound for each ball
  For b = 0 to UBound(BOT)
      if BOT(b).id < 10000 then debug.print "Ballid-" & BOT(b).id
      If BOT(b).ID<10000 Then
        BOT(b).ID = BOT(b).ID+10000*(Team(PN)+1)
        debug.print "Ballid" & BOT(b).ID
        If Team(PN)=0 then BOT(b).color = RGB (255,48,32) ' Red
        If Team(PN)=1 then BOT(b).color = RGB (32,48,255) ' blue
        If Team(PN)=2 then BOT(b).color = RGB (255,215,10)  ' gold
        If Team(PN)=3 then BOT(b).color = RGB (32,255,32) ' green
      End If

    If BallVel(BOT(b)) > 1.5 Then
      rolling(b) = True
'     if BOT(B).z <> 25 Then debug.print BOT(b).Z
      If BOT(b).Z > 28 Then
        ' ball on wire ramp
        StopSound "BallRoll_" & b
        StopSound "fx_plasticrolling" & b
        PlaySound "fx_metalrolling" & b, -1, VolPlayfieldRoll(BOT(b)) * 3 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
        PlaySound "fx_plasticrolling" & b, -1, VolPlayfieldRoll(BOT(b)) * 3 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
'     ElseIf BOT(b).Z >80 Then
'       ' ball on plastic ramp
'       StopSound "BallRoll_" & b
'       StopSound "fx_metalrolling" & b
'       PlaySound "fx_plasticrolling" & b, -1, VolPlayfieldRoll(BOT(b)) * 2 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
      Else
        ' ball on playfield
        StopSound "fx_plasticrolling" & b
        StopSound "fx_metalrolling" & b
        PlaySound "BallRoll_" & b, -1, VolPlayfieldRoll(BOT(b)) * 1.5 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
      End If
    Else
      If rolling(b) = True Then
        StopSound "BallRoll_" & b
        StopSound "fx_plasticrolling" & b
        StopSound "fx_metalrolling" & b
        rolling(b) = False
      End If
    End If

    '***Ball Drop Sounds***
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz > -7 Then
          RandomSoundBallBouncePlayfieldSoft BOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard BOT(b)
        End If
      End If
    End If
    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If
  Next
End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow001,BallShadow002,BallShadow003,BallShadow004,BallShadow005)

Sub BallShadowUpdate
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
        If BOT(b).X < 476 Then
            BallShadow(b).X = BOT(b).X + 2
        Else
            BallShadow(b).X = BOT(b).X + 2
        End If
        ballShadow(b).Y = BOT(b).Y + 8
    If BOT(b).z <40 Then
      BallShadow(b).visible = 0
    Else
      BallShadow(b).visible = 0
    End If
    Next
End Sub



'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
  Next

  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.05, -5.5
  AddPt "Polarity", 2, 0.4, -5.5
  AddPt "Polarity", 3, 0.6, -5.0
  AddPt "Polarity", 4, 0.65, -4.5
  AddPt "Polarity", 5, 0.7, -4.0
  AddPt "Polarity", 6, 0.75, -3.5
  AddPt "Polarity", 7, 0.8, -3.0
  AddPt "Polarity", 8, 0.85, -2.5
  AddPt "Polarity", 9, 0.9,-2.0
  AddPt "Polarity", 10, 0.95, -1.5
  AddPt "Polarity", 11, 1, -1.0
  AddPt "Polarity", 12, 1.05, -0.5
  AddPt "Polarity", 13, 1.1, 0
  AddPt "Polarity", 14, 1.3, 0

  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,  1.05
  addpt "Velocity", 3, 0.53,  1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,  0.945

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub



'******************************************************
'     FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
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
  Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property
  Public Property Get EndPointY: EndPointY = FlipperEndY : End Property

  Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
    if gametime > 100 then Report aChooseArray
  End Sub

  Public Sub Report(aChooseArray)   'debug, reports all coords in tbPL.text
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
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        RemoveBall aBall
        exit Sub
      end if

      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)            'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
   :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
        'playsound "knocker"
      End If
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'   FLIPPER POLARITY AND RUBBER DAMPENER
'     SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
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
  redim aArray(aCount-1+offset) 'Resize original array
  for x = 0 to aCount-1   'set objects back into original array
    if IsObject(a(x)) then
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
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
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

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function

' Used for drop targets and flipper tricks
Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function


'******************************************************
'     FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim BOT, b

  If abs(Flipper1.currentangle) < abs(Endangle1) + 3 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    If Flipper2.currentangle = EndAngle2 Then
      BOT = GetBalls
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          exit Sub
        end If
      Next
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          'debug.print "flippernudge!!"
          BOT(b).velx = BOT(b).velx /1.5
          BOT(b).vely = BOT(b).vely - 1
        end If
      Next
    End If
  Else
    If abs(Flipper1.currentangle) > abs(Endangle1) + 30 then EOSNudge1 = 0
  End If
End Sub

'*************************************************
' Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function

' Used for drop targets and stand up targets
Function Atn2(dy, dx)
  dim pi
  pi = 4*Atn(1)

  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    end if
  ElseIf dx = 0 Then
    if dy = 0 Then
      Atn2 = 0
    else
      Atn2 = Sgn(dy) * pi / 2
    end if
  End If
End Function

'*****************
' Maths
'*****************
Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

'*************************************************
' End - Check ball distance from Flipper for Rem
'*************************************************


dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 0.8
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveCatch = 24
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.033

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
  FlipperPress = 0
  Flipper.eostorqueangle = EOSA
  Flipper.eostorque = EOST*EOSReturn/FReturn


  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim BOT, b
    BOT = GetBalls

    For b = 0 to UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)  '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3*Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
    if FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If

  End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
  Dim LiveCatchBounce                             'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime : CatchTime = GameTime - FCount

  if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
    if CatchTime <= LiveCatch*0.8 Then            'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    else
      LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)  'Partial catch when catch happens a bit late
    end If
    debug.print "Live catch! Bounce: " & LiveCatchBounce

    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (16 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
  End If
End Sub

'*****************************************************************************************************
'*******************************************************************************************************
'END nFOZZY FLIPPERS'


Sub LeftFlipper_Collide(parm)
  LeftFlipperCollide parm
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  if RubberizerEnabled <> 0 then Rubberizer(parm)
End Sub

Sub RightFlipper_Collide(parm)
  RightFlipperCollide parm
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  If RubberizerEnabled <> 0 then Rubberizer(parm)
End Sub



sub Rubberizer(parm)
  if parm < 10 And parm > 2 And Abs(activeball.angmomz) < 10 then
    'debug.print "parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = activeball.angmomz * 2
    activeball.vely = activeball.vely * 1.2
    'debug.print ">> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  Elseif parm <= 2 and parm > 0.2 Then
    'debug.print "* parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = activeball.angmomz * 0.5
    activeball.vely = activeball.vely * (1.2 + rnd(1)/3 )
    'debug.print "**** >> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  end if
end sub


'****************************************************************************
'PHYSICS DAMPENERS
'****************************************************************************

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub


dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False  'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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


  Public Sub Report()   'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


End Class


'******************************************************
'                TRACK ALL BALL VELOCITIES
'                 FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
        public ballvel, ballvelx, ballvely

        Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

        Public Sub Update()        'tracks in-ball-velocity
                dim str, b, AllBalls, highestID : allBalls = getballs

                for each b in allballs
                        if b.id >= HighestID then highestID = b.id
                Next

                if uBound(ballvel) < highestID then redim ballvel(highestID)        'set bounds
                if uBound(ballvelx) < highestID then redim ballvelx(highestID)        'set bounds
                if uBound(ballvely) < highestID then redim ballvely(highestID)        'set bounds

                for each b in allballs
                        ballvel(b.id) = BallSpeed(b)
                        ballvelx(b.id) = b.velx
                        ballvely(b.id) = b.vely
                Next
        End Sub
End Class


Sub RDampen()
Cor.Update
End Sub








'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                                                                                                                'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                                                                                                        'volume level; range [0, 1]
NudgeRightSoundLevel = 1                                                                                                'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                                                                                                'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                                                                                                'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                                                                                        'volume level; range [0, 1]
PlungerPullSoundLevel = 1                                                                                                'volume level; range [0, 1]
RollingSoundFactor = 1.1/5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                                                           'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                                                                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                                                        'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                                                      'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
SlingshotSoundLevel = 0.95                                                                                                'volume level; range [0, 1]
BumperSoundFactor = 4.25                                                                                                'volume multiplier; must not be zero
KnockerSoundLevel = 1                                                                                                         'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                                                                        'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                                                                                        'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                                                                                        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                                                                                'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                                                                        'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                                                                        'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                                                                        'volume level; range [0, 1]
WallImpactSoundFactor = 0.055                                                                                        'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.055/3
SaucerLockSoundLevel = 4
SaucerKickSoundLevel = 4

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                                                                                                        'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                                                                                        'volume multiplier; must not be zero
DTSoundLevel = 1                                                                                                                'volume multiplier; must not be zero
RolloverSoundLevel = 0.5                                                                      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                                                                                                                'volume level; range [0, 1]
BallReleaseSoundLevel = 1                                                                                                'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                                                                        'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                                                                                'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                                                                                                        'volume multiplier; must not be zero


'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
    PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
        Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
        Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
        PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


' *********************************************************************
'                     Fleep  Supporting Ball & Sound Functions
' *********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
        Dim tmp
    tmp = tableobj.y * 2 / tableheight-1
    If tmp > 0 Then
                AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
  tmp = tableobj.x * 2 / tablewidth-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
        Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
        Volz = Csng((ball.velz) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
        VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
    PitchPlayfieldRoll = BallVel(ball) ^2 * 15
End Function

Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
    RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
        PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub


Sub SoundPlungerPull()
        PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, kickback
End Sub'

Sub SoundPlungerReleaseBall()
        PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, kickback
End Sub

Sub SoundPlungerReleaseNoBall()
        PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, kickback
End Sub


'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////
Sub KnockerSolenoid()
        PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////
Sub RandomSoundDrain(drainswitch)
        PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
        PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*7)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
        PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, Sling
    If tilted=0 Then        PlaySoundAtLevelStatic "AAimpactdraw", 0.06, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
        PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
    If tilted=0 Then PlaySoundAtLevelStatic "AAimpactdraw", 0.06, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
        PlaySoundAtLevelStatic SoundFX("AAimpactfire",DOFContactors), 0.06 , Bump


End Sub

Sub RandomSoundBumperMiddle(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
        PlaySoundAtLevelStatic SoundFX("AAimpactfire",DOFContactors), 0.06 , Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
        PlaySoundAtLevelStatic SoundFX("AAimpactfire",DOFContactors), 0.06 , Bump
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
        FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
        PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
        FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
                PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
        FlipperLeftHitParm = parm/10
        If FlipperLeftHitParm > 1 Then
                FlipperLeftHitParm = 1
        End If
        FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
        RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
        FlipperRightHitParm = parm/10
        If FlipperRightHitParm > 1 Then
                FlipperRightHitParm = 1
        End If
        FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
         RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
        PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
        PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
        RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 5 then
                 RandomSoundRubberStrong 1
        End if
        If finalspeed <= 5 then
                 RandomSoundRubberWeak()
         End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
        Select Case Int(Rnd*10)+1
                Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
        End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
'        PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
        PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * RubberWeakSoundFactor


End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 5 then
                 RandomSoundRubberStrong 1
        End if
        If finalspeed <= 5 then
                 RandomSoundRubberWeak()
         End If
End Sub

Sub RandomSoundWall()
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 16 then
                Select Case Int(Rnd*5)+1
                        Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
                End Select
        End if
        If finalspeed >= 6 AND finalspeed <= 16 then
                Select Case Int(Rnd*4)+1
                        Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
                End Select
         End If
        If finalspeed < 6 Then
                Select Case Int(Rnd*3)+1
                        Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
                End Select
        End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
        PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall)/2 * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
        RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
        RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////
Sub RandomSoundBottomArchBallGuide()
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 16 then
                PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
        End if
        If finalspeed >= 6 AND finalspeed <= 16 then
                 Select Case Int(Rnd*2)+1
                        Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
                        Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
                End Select
         End If
        If finalspeed < 6 Then
                 Select Case Int(Rnd*2)+1
                        Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
                        Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
                End Select
        End if
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
        PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
        If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
                RandomSoundBottomArchBallGuideHardHit()
        Else
                RandomSoundBottomArchBallGuide
        End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 16 then
                 Select Case Int(Rnd*2)+1
                        Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
                        Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
                End Select
        End if
        If finalspeed >= 6 AND finalspeed <= 16 then
                PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
         End If
        If finalspeed < 6 Then
                PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
        End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
        PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
        PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 10 then
                 RandomSoundTargetHitStrong()
                RandomSoundBallBouncePlayfieldSoft Activeball
        Else
                 RandomSoundTargetHitWeak()
         End If
End Sub

Sub Targets_Hit (idx)
        PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
        Select Case Int(Rnd*9)+1
                Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
                Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
                Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
                Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
                Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
                Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
                Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
                Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
                Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
        End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
        PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
        Select Case Int(Rnd*5)+1
                Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
                Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
                Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
                Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
                Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
        End Select
End Sub



'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
        PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate()
        PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
        SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
        SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
        PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
        PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
        If Activeball.velx > 1 Then SoundPlayfieldGate
        StopSound "Arch_L1"
        StopSound "Arch_L2"
        StopSound "Arch_L3"
        StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
        If activeball.velx < -8 Then
                RandomSoundRightArch
        End If
End Sub

Sub Arch2_hit()
        If Activeball.velx < 1 Then SoundPlayfieldGate
        StopSound "Arch_R1"
        StopSound "Arch_R2"
        StopSound "Arch_R3"
        StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
        If activeball.velx > 10 Then
                RandomSoundLeftArch
        End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
        PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), 1, Activeball
End Sub

Sub SoundSaucerKick(scenario, saucer)
        Select Case scenario
                Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), 1, saucer
                Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), 2, saucer
            PlaySoundAtLevelStatic SoundFX("fx_kicker", DOFContactors), 2, saucer
        End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
'Debug.print velocity

    Dim snd
  If velocity>1 Then
        Select Case Int(Rnd*7)+1
                Case 1 : snd = "Ball_Collide_1"
                Case 2 : snd = "Ball_Collide_2"
                Case 3 : snd = "Ball_Collide_3"
                Case 4 : snd = "Ball_Collide_4"
                Case 5 : snd = "Ball_Collide_5"
                Case 6 : snd = "Ball_Collide_6"
                Case 7 : snd = "Ball_Collide_7"
        End Select
        PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  End If
End Sub


'/////////////////////////////////////////////////////////////////
'                                        End Mechanical Sounds
'/////////////////////////////////////////////////////////////////




'**********************************************************************************************************
'* MUSIC *
'**********************************************************************************************************

Dim initialmusic
Sub Table1_musicdone
  If initialmusic=0 or playchamp=1 Then
    playchamp=2
    initialmusic=1
    Playmusic "Unrealtournament/UT-Menu.mp3"
    MusicVolume=Music_Volume
  Else
    MusicOn
  End If
End Sub

Sub MusicON
  EndMusic
  StopSound "AAOpeningVO"
  PlayMusic "Unrealtournament/UT" & int(rnd(1)*13)+1 & ".mp3"
  MusicVolume=Music_Volume*0.6
End Sub


'**********************************************************************************************************
'* InstructionCard *
'**********************************************************************************************************

Dim CardCounter, ScoreCard
Sub CardTimer_Timer
        If scorecard=1 Then
                CardCounter=CardCounter+2
                If CardCounter>50 Then CardCounter=50
        Else
                CardCounter=CardCounter-4
                If CardCounter<0 Then CardCounter=0
        End If
        InstructionCard.transX = CardCounter*6
        InstructionCard.transY = CardCounter*6
        InstructionCard.transZ = -cardcounter*2
'        InstructionCard.objRotX = -cardcounter/2
        InstructionCard.size_x = 1+CardCounter/25
        InstructionCard.size_y = 1+CardCounter/25
        If CardCounter=0 Then
                CardTimer.Enabled=0
                InstructionCard.visible=0
        Else
                InstructionCard.visible=1
        End If
End Sub



'**********************************************************************************************************
'* Raining Effect *
'**********************************************************************************************************

Dim Raining
Sub RainingMod
  Raining=Raining+1
  if Raining>8 then Raining=0
  RainFlasher.imageA="rain" & Raining
End Sub



'**********************************************************************************************************
'* Gun Flasher *
'**********************************************************************************************************

Dim GunFL
Sub GunFlash_timer
  If GunFlash.timerenabled=0 Then
    GunFlash.timerenabled=1
    GunFL=100
  Else
    GunFL=GunFL-10
    GunFlash.opacity=GunFL
    If GunFL =0 Then GunFlash.timerenabled=0
  End If
End Sub




'**********************************************************************************************************
'* Lane Divider *
'**********************************************************************************************************

Sub DividerWall1_Timer ' divider=on ramp goes to tower

  If DividerWall1.timerenabled=0 Then
    DividerWall1.timerenabled=1
    DividerWall0.timerenabled=0
    DividerWall1.collidable=1
    DividerWall0.collidable=0
    DividerWall3.collidable=0
    If Primitive100.rotY<180 Then playSound "AAldoorsopen", 1, BG_Volume*0.4, 0,0,0,0,0,0
  Else
    DividerWall1.collidable=1
    DividerWall0.collidable=0
    DividerWall3.collidable=0
    i = Primitive100.rotY
    i = i + 10
    If i > 180 Then i = 180
    Primitive100.rotY = i
    If i = 180 Then
      DividerWall1.timerenabled=0
    End If
  End If
End Sub


Sub DividerWall0_Timer ' divider=off ramp goes normal

  If DividerWall0.timerenabled=0 Then
    DividerWall0.timerenabled=1
    DividerWall1.timerenabled=0
    DividerWall1.collidable=0
    DividerWall0.collidable=1
    DividerWall3.collidable=1
    If Primitive100.rotY>115 Then playSound "AAldoorsopen", 1, BG_Volume*0.4, 0,0,0,0,0,0
  Else
    DividerWall1.collidable=0
    DividerWall0.collidable=1
    DividerWall3.collidable=1
    i = Primitive100.rotY
    i = i - 10
    If i < 115 Then i = 115
    Primitive100.rotY = i
    If i = 115 Then
      DividerWall0.timerenabled=0
    End If
  End If
End Sub




'**********************************************************************************************************
'* Enemy Flag *
'**********************************************************************************************************

Sub EnemyFlag

  SC(PN,30)=SC(PN,30)+1 ' counter for get flag back
  debug.print "enemyflag"&SC(pn,30)
  If SC(PN,30)>=SC(PN,32) Then  ' get flag back level
    If SLS(41,1)>1 Then
      PlaySound "AACaptureSound" & int(rnd(1)*3)+1 , 1, 0.5 * BG_Volume, 0, 0,0,0, 0, 0
      FlexNewText "! ENEMY CAPTURE !"
      'E154 2 ENEMY Capture
      DOF 154,2
      If DOFdebug=1 Then debug.print "DOF 154,2 enemy capture"
      If SC(PN,28)=SC(PN,29)+1 Then LOSTLEAD.enabled=1

      SC(PN,29)=SC(PN,29)+1

      BlueDisplay SC(PN,29)
      SPB1 72,72,7,2,1,1
      SPB1 73,74,5,2,1,1
      PlaySound "getflag" & int(rnd(1)*5)+1, 1, 0.7 * BG_Volume, 0, 0,0,0, 0, 0
      SC(PN,30)=0
      FlagOn10=0 ' RESET AUTOCAP
      Exit Sub
    End If
' 28= map score left Counter
' 29= map score Right Counter
    PlaySound "getflag" & int(rnd(1)*5)+1, 1, 0.33 * BG_Volume, 0, 0,0,0, 0, 0
    SC(PN,30)=0
  debug.print "enemyflag"&SC(pn,30)&"reset"
    'light enemy flags
    SLS(41,1)=2
    If SLS(42,1)>1 Then SLS(42,1)=1
    SLS(48,1)=2
    SLS(63,1)=2
    SLS(70,1)=2
    L063.timerenabled=1
  End If

End Sub



Sub MapBlinker_Timer
  If startgame>0 And Tilted=0 Then
    if SC(PN,27)=0 Then SPB1 50,50,1,2,1,1
    if SC(PN,27)=1 Then SPB1 53,53,1,2,1,1
    if SC(PN,27)=2 Then SPB1 52,52,1,2,1,1
    if SC(PN,27)=3 Then SPB1 51,51,1,2,1,1
  End If
End Sub

Sub Monsterkill_Timer
  monsterkill.enabled=0
End Sub

Dim FFlevel, KSlevel
Sub CheckFastFrags

  Taunting_Timer

  Select Case KSlevel
    case 0 :
      If SC(PN,19)>4 Then
        delayspree.enabled=1
      End If
    case 1 :
      If SC(PN,19)>9 Then
        delayspree.enabled=1
      End If
    case 2 :
      If SC(PN,19)>14 Then
        delayspree.enabled=1
      End If
    case 3 :
      If SC(PN,19)>19 Then
        delayspree.enabled=1
      End If
    case 4 :
      If SC(PN,19)>24 Then
        delayspree.enabled=1
      End If
  End Select


  Select Case FFlevel

    case 3 :
      If FastFrags > 4.99 Then
        PlaySound "monsterkill", 1,  BG_Volume, 0, 0,0,0, 0, 0
  DOF 146,2
  If DOFdebug=1 Then debug.print "DOF 146,2 mmonsterkill"
'E145 godlike
'E146 monsterkill
        If monsterkill.enabled=0 Then
          Scoring 5000,1000
'         FFT=5000
          FlexFlashText3 "MonsterKill","",75
          monsterkill.enabled=1
        End If
        FastFrags=FastFrags-1
        Scoring 3000,600
        Exit Sub
      End If

    case 2 :
      If FastFrags > 3.99 Then
        PlaySound "ultrakill", 1, 0.9 * BG_Volume, 0, 0,0,0, 0, 0
        FlexFlashText3 "Ultra Kill","",75
        FFlevel=3
        Scoring 2500,500
        Exit Sub
      End If

    case 1 :
      If FastFrags > 2.99 Then
        PlaySound "AAmultikill", 1, 0.9 * BG_Volume, 0, 0,0,0, 0, 0
        FlexFlashText3 "Multi Kill","",75

        FFlevel=2
        Scoring 1500,300
        Exit Sub
      End If

    case 0 :
      If FastFrags > 1.99 Then
        PlaySound "doublekill", 1, 0.6 * BG_Volume, 0, 0,0,0, 0, 0
        FlexFlashText3 "Double Kill","",75
        FFlevel=1
        Scoring 1000,200
        Exit Sub
      End If
  End Select

End Sub
Dim Extraball

Sub Delayspree_Timer
  delayspree.enabled=0

  Select Case KSlevel
    case 0 : PlaySound "killingspree", 1, 0.6 * BG_Volume, 0, 0,0,0, 0, 0
      FlexFlashText1  "Spree !",88    ' 1 line  , pause = 100 for oki blink, lower for faster


    case 1 : PlaySound "rampage", 1, 0.7 * BG_Volume, 0, 0,0,0, 0, 0
        If Extraball=0 Then ExtraBall=1 : SLS(43,1)=2 : L043.timerenabled=1
      FlexFlashText1 "RAMPAGE",88

    case 2 : PlaySound "AAdominating", 1, 0.75 * BG_Volume, 0, 0,0,0, 0, 0
      FlexFlashText1 "Dominating",88

    case 3 : PlaySound "unstoppable", 1, 0.88 * BG_Volume, 0, 0,0,0, 0, 0
      FlexFlashText1 "Unstoppable",88

    case 4 : PlaySound "godlike", 1, BG_Volume, 0, 0,0,0, 0, 0
      FlexFlashText1 "GODLIKE",88
  DOF 145,2
  If DOFdebug=1 Then debug.print "DOF 145,2 GODLIKE"
'E145 godlike

  End Select
  KSlevel=KSlevel+1
End Sub

Dim FlagOn4
Dim FlagOn10 'autocap no new flag taken

Sub L063_Timer   ' enemyflag alarm
  If SLS(63,1)=0 Then
    L063.timerenabled=0
    FlagOn10=0
  Else
    PlaySound "Aflagtaken", 1, 0.6 * BG_Volume, 0, 0,0,0, 0, 0

    FlagOn10=FlagOn10+1
    If FlagOn10>9 Then  ' AUTOCAP TIMER AROUND 32 SEC
      PlaySound "AACaptureSound" & int(rnd(1)*3)+1 , 1, 0.45 * BG_Volume, 0, 0,0,0, 0, 0
      FlexNewText "! ENEMY CAPTURE !"
      SC(PN,29)=SC(PN,29)+1
      'E154 2 ENEMY Capture
      DOF 154,2
      If DOFdebug=1 Then debug.print "DOF 154,2 enemy capture"

      BlueDisplay SC(PN,29)
      SPB1 72,72,7,2,1,1
      SPB1 73,74,5,2,1,1
      SC(PN,30)=0
      FlagOn10=0 ' RESET AUTOCAP
        ' WILL CAP AGAIN IN 32 SEC IF NOT TURNED OFF
      Exit Sub
    End If

    FlagOn4=FlagOn4+1
    If FlagOn4= 4 Then
      FlagOn4=0
      PlaySound "getflag" & int(rnd(1)*5)+1, 1, 0.55 * BG_Volume, 0, 0,0,0, 0, 0
      FlexNewText "Get Our FLAG Back!"
    End If

  End If
End Sub



Sub L043_Timer  ' start extraballlight 1300ms nali at rampage
  L043.timerenabled=0
  PlaySound "AANali" & Int(rnd(1)*6)+1 , 1, BG_Volume, 0, 0,0,0, 0, 0
End Sub





'**********************************************************************************************************
'* TilT *
'**********************************************************************************************************

Sub Tilting    'Nudge comes here

  If Tilted=0 Then

    If Tilt > 115 Then
      TiltGame
      Exit Sub
    End If

    Tilt = Tilt + TiltSens
    If Tilt > TiltSens Then
    PlaySound "Danger", 1, 0.7 * BG_Volume, 0, 0,0,0, 0, 0
    FlexFlashText1 "Danger!",100
    End if

  End If

End Sub


Sub TiltGame

  SavePlayerLights

  StopSound "Danger"
  PlaySound "Tilt"

  debug.print"Tilted=1"
  SC(PN,0)=0 ' no bouns for you
  Tilted=1 ' set it to 2 so everything is off, control putting it on at end of ball stuff
  TiltedWait_timer ' .. wait for no balls in play ( excluded the ones in lockup )

  Tilt=200
  If UMainDMD.isrendering Then UMainDMD.cancelrendering

  FlexFlashText1 "-TILT-",250
  FlexFlashText1 "-TILT-",250


  Leftflipper.RotateTOStart
  Rightflipper.RotateTOStart
  For i = 1 to 200
    SLS(i,5)=0
    SLS(i,1)=0
  Next

  EndMusic

End Sub

Sub TiltedWait_Timer

  If tiltedwait.enabled=0 Then
    tiltedwait.enabled=1
  Else
    If tilt=0 Then
      If TiltedBalls - Locked1 - Locked2 = -1 Then
        Tilted=0
        debug.print"Tilted=0"
        tiltedwait.enabled=0
        musicon  ' need to restart this
        If UMainDMD.isrendering Then UMainDMD.cancelrendering
        StartNewPlayer
      End If
    End If
  End If

End Sub


'**********************************************************************************************************
'* ADD PLAYER SCORE AND BONUS * WITH BONUS FROM DROPTARGETS *
'**********************************************************************************************************
DIM FFTscore
Sub Scoring ( addtoscore , addtobonus )

  tempnr=2
  If SC(PN,27)=0 Then tempnr=2  ' coret
  If SC(PN,27)=1 Then tempnr=5  'november
  If SC(PN,27)=2 Then tempnr=10 'dreary
  If SC(PN,27)=3 Then tempnr=20 'face

  tempnr2=1
  If SLS(35,1) > 0 Then tempnr2=tempnr2+0.20
  If SLS(36,1) > 0 Then tempnr2=tempnr2+0.70
  If SLS(37,1) > 0 Then tempnr2=tempnr2+0.20
  If SLS(38,1) > 0 Then tempnr2=tempnr2+0.40
  If SLS(39,1) > 0 Then tempnr2=tempnr2+0.40
  If SLS(40,1) > 0 Then tempnr2=tempnr2+0.40
  FFTscore=int(addtoscore * tempnr * tempnr2)
  SC(PN,1)=SC(PN,1) + FFTscore
  SC(PN,0)=SC(PN,0) + int(addtobonus * tempnr * tempnr2)
  if addtobonus>0 Then debug.print "add"&addtobonus & ">" & tempnr*addtobonus &" BONUS"& SC(PN,0)
End Sub



'**********************************************************************************************************
'* FLEX DMD*
'**********************************************************************************************************



Dim MainDMD
Dim UMainDMD
Dim fso
dim curdir

Sub FlexDMDinit
  dim tmp
  Set MainDMD = CreateObject("FlexDMD.FlexDMD")
    If MainDMD is Nothing Then
        MsgBox "FlexDMD was not found.  This table will NOT run without it."
        Exit Sub
    End If

  With MainDMD
    .RenderMode = 2
    .Width = 128
    .Height = 32
    .GameName = "UT99_main"
    .Run = True
    .Color = RGB(255,150,88)
  End With


  Set UMainDMD = MainDMD.NewUltraDMD()
  UMainDMD.Init

  Set fso = CreateObject("Scripting.FileSystemObject")
    curDir = fso.GetAbsolutePathName(".")
    UMainDMD.SetProjectFolder curDir & "\UT99DMD"

End Sub



'**********************************************************************************************************
'* Set Team colors on miniscoreboard bottom pf *
'**********************************************************************************************************

'E147 0/1 redundercab
'E148 0/1 blueundercab
'E149 0/1 goldundercab
'E150 0/1 greenundercab


Sub SetTeamColors


  If Team(PN)=0 Then
    DOF 147,1 : If DOFdebug=1 Then debug.print "DOF 147,1 RED UNDERCAB ON"
    DOF 148,0
    DOF 149,0
    DOF 150,0
    LT071.ColorFull= RGB(255,0,0)
    L073.ColorFull= RGB(255,0,0)
    p071.material="InsertRedOnTridark"
'   p071off.material="InsertRedOFFTridark"
    R1.color=RGB(255,0,0) : R1.colorfull=RGB(255,0,0)
    R2.color=RGB(255,0,0) : R2.colorfull=RGB(255,0,0)
    R3.color=RGB(255,0,0) : R3.colorfull=RGB(255,0,0)
    R4.color=RGB(255,0,0) : R4.colorfull=RGB(255,0,0)
    R5.color=RGB(255,0,0) : R5.colorfull=RGB(255,0,0)
    R6.color=RGB(255,0,0) : R6.colorfull=RGB(255,0,0)
    R7.color=RGB(255,0,0) : R7.colorfull=RGB(255,0,0)
    Redbaselight001.image="Led_Red"
    Redbaselight002.image="Led_Red"
    Redbaselight003.image="Led_Red"
    Redbaselight001.material="Plastic laneguides Red"
    Redbaselight002.material="Plastic laneguides Red"
    Redbaselight003.material="Plastic laneguides Red"
    RedBaseLight004.material="Plastic laneguides Red"
    p042.material="insertRedonTri"
    LT042.colorfull=RGB(200,0,0)
    L042.colorfull=RGB(200,0,0)


    LT072.ColorFull= RGB(0,0,255) 'Opponent Side
    L074.ColorFull= RGB(0,0,255)
    p072.material="insertBlueonTri"
'   p072off.material="InsertDarkBlueOffTri"
    B1.color=RGB(0,0,255) : B1.colorfull=RGB(0,0,255)
    B2.color=RGB(0,0,255) : B2.colorfull=RGB(0,0,255)
    B3.color=RGB(0,0,255) : B3.colorfull=RGB(0,0,255)
    B4.color=RGB(0,0,255) : B4.colorfull=RGB(0,0,255)
    B5.color=RGB(0,0,255) : B5.colorfull=RGB(0,0,255)
    B6.color=RGB(0,0,255) : B6.colorfull=RGB(0,0,255)
    B7.color=RGB(0,0,255) : B7.colorfull=RGB(0,0,255)

    p067.material="insertBlueonTri"
    p070.material="insertBlueonTri"
    p048.material="insertBlueonTri"
    p063.material="insertBlueonTri"
    p041.material="insertBlueonTri"

    LT067.colorfull=RGB(0,0,255)
    LT070.colorfull=RGB(0,0,255)
    LT048.colorfull=RGB(0,0,255)
    LT063.colorfull=RGB(0,0,255)
    LT041.colorfull=RGB(0,0,255)
    L067.colorfull=RGB(0,0,255)
    L070.colorfull=RGB(0,0,255)
    L048.colorfull=RGB(0,0,255)
    L063.colorfull=RGB(0,0,255)
    L041.colorfull=RGB(0,0,255)


  End If

  If Team(PN)=1 Then
    DOF 147,0
    DOF 148,1 : If DOFdebug=1 Then debug.print "DOF 148,1 BLUE UNDERCAB ON"
    DOF 149,0
    DOF 150,0
    LT071.ColorFull= RGB(0,0,255)
    L073.ColorFull= RGB(0,0,255)
    p071.material="insertBlueonTri"
'   p071off.material="InsertDarkBlueOffTri"
    R1.color=RGB(0,0,255) : R1.colorfull=RGB(0,0,255)
    R2.color=RGB(0,0,255) : R2.colorfull=RGB(0,0,255)
    R3.color=RGB(0,0,255) : R3.colorfull=RGB(0,0,255)
    R4.color=RGB(0,0,255) : R4.colorfull=RGB(0,0,255)
    R5.color=RGB(0,0,255) : R5.colorfull=RGB(0,0,255)
    R6.color=RGB(0,0,255) : R6.colorfull=RGB(0,0,255)
    R7.color=RGB(0,0,255) : R7.colorfull=RGB(0,0,255)
    Redbaselight001.image="Led_Blue"
    Redbaselight002.image="Led_Blue"
    Redbaselight003.image="Led_Blue"
    Redbaselight001.material="Plastic laneguides Blue"
    Redbaselight002.material="Plastic laneguides Blue"
    Redbaselight003.material="Plastic laneguides Blue"
    RedBaseLight004.material="Plastic laneguides Blue"
    p042.material="insertBlueonTri"
    LT042.colorfull=RGB(0,0,255)
    L042.colorfull=RGB(0,0,255)

    LT072.ColorFull= RGB(255,0,0) 'Opponent Side
    L074.ColorFull= RGB(255,0,0)
    p072.material="InsertRedOnTri"
'   p072off.material="InsertRedOFFTridark"
    B1.color=RGB(255,0,0) : B1.colorfull=RGB(255,0,0)
    B2.color=RGB(255,0,0) : B2.colorfull=RGB(255,0,0)
    B3.color=RGB(255,0,0) : B3.colorfull=RGB(255,0,0)
    B4.color=RGB(255,0,0) : B4.colorfull=RGB(255,0,0)
    B5.color=RGB(255,0,0) : B5.colorfull=RGB(255,0,0)
    B6.color=RGB(255,0,0) : B6.colorfull=RGB(255,0,0)
    B7.color=RGB(255,0,0) : B7.colorfull=RGB(255,0,0)
    p067.material="insertRedonTri"
    p070.material="insertRedonTri"
    p048.material="insertRedonTri"
    p063.material="insertRedonTri"
    p041.material="insertRedonTri"

    LT067.colorfull=RGB(200,0,0)
    LT070.colorfull=RGB(200,0,0)
    LT048.colorfull=RGB(200,0,0)
    LT063.colorfull=RGB(200,0,0)
    LT041.colorfull=RGB(200,0,0)
    L067.colorfull=RGB(200,0,0)
    L070.colorfull=RGB(200,0,0)
    L048.colorfull=RGB(200,0,0)
    L063.colorfull=RGB(200,0,0)
    L041.colorfull=RGB(200,0,0)
  End If

  If Team(PN)=2 Then ' GOLD VS Red
    DOF 147,0
    DOF 148,0
    DOF 149,1 : If DOFdebug=1 Then debug.print "DOF 149,1 gold UNDERCAB ON"
    DOF 150,0
    LT071.ColorFull= RGB(255,215,0)
    L073.ColorFull= RGB(255,215,0)
    p071.material="InsertDarkYellowOnTri"
'   p071off.material="InsertDarkYellowOffTri"
    R1.color=RGB(255,215,0) : R1.colorfull=RGB(255,215,0)
    R2.color=RGB(255,215,0) : R2.colorfull=RGB(255,215,0)
    R3.color=RGB(255,215,0) : R3.colorfull=RGB(255,215,0)
    R4.color=RGB(255,215,0) : R4.colorfull=RGB(255,215,0)
    R5.color=RGB(255,215,0) : R5.colorfull=RGB(255,215,0)
    R6.color=RGB(255,215,0) : R6.colorfull=RGB(255,215,0)
    R7.color=RGB(255,215,0) : R7.colorfull=RGB(255,215,0)
    Redbaselight001.image="Led_Yellow"
    Redbaselight002.image="Led_Yellow"
    Redbaselight003.image="Led_Yellow"
    Redbaselight001.material="Plastic laneguides Yellow"
    Redbaselight002.material="Plastic laneguides Yellow"
    Redbaselight003.material="Plastic laneguides Yellow"
    RedBaseLight004.material="Plastic laneguides Yellow"
    p042.material="insertYellowonTri"
    LT042.colorfull=RGB(200,215,0)
    L042.colorfull=RGB(200,215,0)

    LT072.ColorFull= RGB(255,0,0) 'Opponent Side
    L074.ColorFull= RGB(255,0,0)
    p072.material="InsertRedOnTri"
'   p072off.material="InsertRedOFFTridark"
    B1.color=RGB(255,0,0) : B1.colorfull=RGB(255,0,0)
    B2.color=RGB(255,0,0) : B2.colorfull=RGB(255,0,0)
    B3.color=RGB(255,0,0) : B3.colorfull=RGB(255,0,0)
    B4.color=RGB(255,0,0) : B4.colorfull=RGB(255,0,0)
    B5.color=RGB(255,0,0) : B5.colorfull=RGB(255,0,0)
    B6.color=RGB(255,0,0) : B6.colorfull=RGB(255,0,0)
    B7.color=RGB(255,0,0) : B7.colorfull=RGB(255,0,0)
    p067.material="insertRedonTri"
    p070.material="insertRedonTri"
    p048.material="insertRedonTri"
    p063.material="insertRedonTri"
    p041.material="insertRedonTri"

    LT067.colorfull=RGB(200,0,0)
    LT070.colorfull=RGB(200,0,0)
    LT048.colorfull=RGB(200,0,0)
    LT063.colorfull=RGB(200,0,0)
    LT041.colorfull=RGB(200,0,0)
    L067.colorfull=RGB(200,0,0)
    L070.colorfull=RGB(200,0,0)
    L048.colorfull=RGB(200,0,0)
    L063.colorfull=RGB(200,0,0)
    L041.colorfull=RGB(200,0,0)
  End If

  If Team(PN)=3 Then ' gREEN VS Red
    DOF 147,0
    DOF 148,0
    DOF 149,0
    DOF 150,1 : If DOFdebug=1 Then debug.print "DOF 150,1 GREEN UNDERCAB ON"
    LT071.ColorFull= RGB(0,255,0)
    L073.ColorFull= RGB(0,255,0)
    p071.material="InsertDarkGreenOnTri"
'   p071off.material="InsertDarkGreenOffTri"
    R1.color=RGB(0,255,0) : R1.colorfull=RGB(0,255,0)
    R2.color=RGB(0,255,0) : R2.colorfull=RGB(0,255,0)
    R3.color=RGB(0,255,0) : R3.colorfull=RGB(0,255,0)
    R4.color=RGB(0,255,0) : R4.colorfull=RGB(0,255,0)
    R5.color=RGB(0,255,0) : R5.colorfull=RGB(0,255,0)
    R6.color=RGB(0,255,0) : R6.colorfull=RGB(0,255,0)
    R7.color=RGB(0,255,0) : R7.colorfull=RGB(0,255,0)
    Redbaselight001.image="Led_Green"
    Redbaselight002.image="Led_Green"
    Redbaselight003.image="Led_Green"
    Redbaselight001.material="Plastic Greenish"
    Redbaselight002.material="Plastic Greenish"
    Redbaselight003.material="Plastic Greenish"
    RedBaseLight004.material="Plastic Greenish"
    p042.material="insertGreenonTri"
    LT042.colorfull=RGB(0,200,0)
    L042.colorfull=RGB(0,200,0)

    LT072.ColorFull= RGB(255,0,0) 'Opponent Side
    L074.ColorFull= RGB(255,0,0)
    p072.material="InsertRedOnTri"
'   p072off.material="InsertRedOFFTridark"
    B1.color=RGB(255,0,0) : B1.colorfull=RGB(255,0,0)
    B2.color=RGB(255,0,0) : B2.colorfull=RGB(255,0,0)
    B3.color=RGB(255,0,0) : B3.colorfull=RGB(255,0,0)
    B4.color=RGB(255,0,0) : B4.colorfull=RGB(255,0,0)
    B5.color=RGB(255,0,0) : B5.colorfull=RGB(255,0,0)
    B6.color=RGB(255,0,0) : B6.colorfull=RGB(255,0,0)
    B7.color=RGB(255,0,0) : B7.colorfull=RGB(255,0,0)
    p067.material="insertRedonTri"
    p070.material="insertRedonTri"
    p048.material="insertRedonTri"
    p063.material="insertRedonTri"
    p041.material="insertRedonTri"

    LT067.colorfull=RGB(200,0,0)
    LT070.colorfull=RGB(200,0,0)
    LT048.colorfull=RGB(200,0,0)
    LT063.colorfull=RGB(200,0,0)
    LT041.colorfull=RGB(200,0,0)
    L067.colorfull=RGB(200,0,0)
    L070.colorfull=RGB(200,0,0)
    L048.colorfull=RGB(200,0,0)
    L063.colorfull=RGB(200,0,0)
    L041.colorfull=RGB(200,0,0)
  End If

End Sub



'**********************************************************************************************************
'* LOAD/SAVE HIGHSCORES *
'**********************************************************************************************************

Sub LoadHighScore
  Dim FileObj, ScoreFile, TextStr
  Dim i2, i3
  Dim SavedDataTemp(18)

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "UT99.txt") then
    DefaultHighScores
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "UT99.txt")
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      DefaultHighScores
      Exit Sub
    End if
    For i = 1 to 17
      If (TextStr.AtEndOfStream=True) then
        TextStr.Close
        DefaultHighScores
        Set ScoreFile = Nothing
        Set FileObj = Nothing
        Exit Sub
      End if
      SavedDataTemp(i)=TextStr.ReadLine
      If SavedDataTemp(i)="" Then
        TextStr.Close
        DefaultHighScores
        Set ScoreFile = Nothing
        Set FileObj = Nothing
        Exit Sub
      End If
    Next

    TextStr.Close
    Set ScoreFile = Nothing
      Set FileObj = Nothing

    hScore(1) =int ( SavedDataTemp(1) )
    highscorename(1)=SavedDataTemp(2)
    hScore(2) =int ( SavedDataTemp(3) )
    highscorename(2)=SavedDataTemp(4)
    hScore(3) =int ( SavedDataTemp(5) )
    highscorename(3)=SavedDataTemp(6)

    GamesPlayd=int ( SavedDataTemp(7) )
      LastScore =int ( SavedDataTemp(8) )
    Credits  = int ( SavedDataTemp(9) )

    Team(1) = int ( SavedDataTemp(10) )
    Team(2) = int ( SavedDataTemp(11) )
    Team(3) = int ( SavedDataTemp(12) )
    Team(4) = int ( SavedDataTemp(13) )

    Taunt(1)= int ( SavedDataTemp(14) )
    Taunt(2)= int ( SavedDataTemp(15) )
    Taunt(3)= int ( SavedDataTemp(16) )
    Taunt(4)= int ( SavedDataTemp(17) )

End Sub


Sub DefaultHighScores
  'highscore reset  . errors unavoidable
  FlexFlashText3 "Highscore","reset",100
  highscorename(1)="VPX"
  highscorename(2)="VPX"
    highscorename(3)="VPX"
  hscore(1)=500000
  hscore(2)=300000
  hscore(3)=100000

  Credits=0
  GamesPlayd=0
  LastScore=0

  Team(1)=0  ' 0-3 for each player ( 0=red 1blue 2gold 3green)
  Team(2)=1
  Team(3)=2
  Team(4)=3

  Taunt(1)=0
  Taunt(2)=1
  Taunt(3)=2
  Taunt(4)=3

End Sub

Dim ReplayGoals,ReplayGoalDown
Sub SaveHighScore

    Dim FileObj
    Dim ScoreFile
    Set FileObj=CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) then
      Exit Sub
    End if
    Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "UT99.txt",True)

      ScoreFile.WriteLine hscore(1)
      ScoreFile.WriteLine highscorename(1)
      ScoreFile.WriteLine hscore(2)
      ScoreFile.WriteLine highscorename(2)
      ScoreFile.WriteLine hscore(3)
      ScoreFile.WriteLine highscorename(3)
      ScoreFile.WriteLine GamesPlayd
      ScoreFile.WriteLine LastScore
      ScoreFile.WriteLine Credits
      ScoreFile.WriteLine team(1)
      ScoreFile.WriteLine team(2)
      ScoreFile.WriteLine team(3)
      ScoreFile.WriteLine team(4)
      ScoreFile.WriteLine Taunt(1)
      ScoreFile.WriteLine Taunt(2)
      ScoreFile.WriteLine Taunt(3)
      ScoreFile.WriteLine Taunt(4)


      ScoreFile.Close
    Set ScoreFile=Nothing
    Set FileObj=Nothing
End Sub


Sub table1_Exit
  SaveHighScore

  If Not UMainDMD is Nothing Then
    If UMainDMD.IsRendering Then
      UMainDMD.CancelRendering
    End If
    UMainDMD.Uninit
    UMainDMD = NULL
  End If
End Sub




'*Scoring
'
'CamperBonus 150000,15000
'headshot 25000,4000
'Frag 1000,200
'Redeemer 20000,4000
'return flag 5000,1000
'assist 1000,100
'lockball 2000,1000
'startMB 10000,2000
'endmap   20000 X caps,0
'winner  50000,0
'capture 50000,5000
'DT LightTimers Mid=30,30,30 Right=60,60,60sec
'DT 1000,100
'Amp 2000,400
'skillshotlights 500,0
'inlanes 1000,0
'outlanedrain 50000,0
'kickback fire 1000,100
'pwnage 10000,2000   +1000,200 for each time during same ball
'enable KB 1000,0
'superbumpers 1000,0 else 100,0p
'superspinnner = 3000,0 else 200,0
'ownage 2 = superbumpers until ball lost
'ownage 3,6 = award bonusmultiplyer
'ownage 5,7,9,11,13,15,17,19 = superspinners 60 seconds
'ownage 4,8,10,12,14,16,18,20 = extra light
'bonuslight 200 rollover
'3x bonuslight = 5000,500 + bonusmultiplyer
'more than 8x = 50000,1000
'slingshot 50
'double 1000,200
'multi  1500,300
'ultra  2500,500
'monstr 5000,1000
'skillshot
'bad     500,50
'almost  3000,300
'good    10000,1000
'perfect 20000,2000
'2nd skillshot mainramp 50000,5000 If good+ ss, possible to do more than once in 9,5 seconds
'invisibility minigame  200000,30000 + 30 second respawn
'max one each 5 lanes then main ramp ( 2 on any lane = restart, no restart on main ramp anymore )
'
' Coret  2x all scores
' november 5x
' dreary 10x
' face 20x
'
'
'
'
'*********************
'VR Mode
'*********************
DIM VRThings
If VRRoom > 0 Then
  scoretext.visible = 0
  DMD.visible = 1
  PinCab_Backglass.blenddisablelighting = 3
  PinCab_Rails.visible = 1
  Flasher006.visible = 0
  Flasher007.visible = 0
  If VRRoom = 1 Then
    for each VRThings in VR_Cab:VRThings.visible = 1:Next
    for each VRThings in VR_Min:VRThings.visible = 1:Next
  End If
  If VRRoom = 2 Then
    for each VRThings in VR_Cab:VRThings.visible = 0:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
  PinCab_Backbox.visible=1
  PinCab_Backglass.visible=1
  End If
Else
    for each VRThings in VR_Cab:VRThings.visible = 0:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
  if DesktopMode then
    scoretext.visible = 1
    PinCab_Rails.visible = 1
  else
    scoretext.visible = 0
    PinCab_Rails.visible = 0
  End if
End If
