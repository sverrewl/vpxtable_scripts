'Monster Bash - IPDB No. 4441
' © Williams 1998
'VPX recreation by tom tower & ninuzzu
'thanks to unclewilly and randr for previous versions
'thanks to flupper and zany for the domes and bumpers
'thanks to VPDev Team for the freaking amazing VPX

Option Explicit
Randomize

' Thalamus 2018-07-24
' Added "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Table doesn't use JP ball rolling standars
' Thalamus 2018-12-17 : Added FFv2
' Thalamus 2018-09-09 : Improved directional sounds

' !! NOTE : Table not verified yet !!

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolGates  = 1    ' Gates volume.
Const VolTarg   = 1    ' Targets volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 0.5  ' Flipper volume.


Const BallSize = 51
Const BallMass = 1.2

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode
Const UseVPMModSol = 1

' Rom Name
Const cGameName = "mb_106b"

LoadVPM "02000000", "WPC.VBS", 3.50

' Thalamus - for Fast Flip v2
' NoUpperRightFlipper
NoUpperLeftFlipper

' Standard Options
Const UseSolenoids = 2
' Thal : Added because of useSolenoid=2
Const cSingleLFlip = 0
Const cSingleRFlip = 0
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

'************************************************************************
'            INIT TABLE
'************************************************************************

Dim bsTrough, bsLKick, xx

Sub Table1_Init
  vpmInit Me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Monster Bash Williams 1998"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 4
    .Hidden = DesktopMode
    .dip(0)=&h00  'Set DIP to USA (this fixes the Phantom Flip lights)
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
    .Switch(22) = 1 'close coin door
    .Switch(24) = 0 'always closed
  End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Nudging
    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmTrough
    With bsTrough
    .size = 4
    .initSwitches Array(32, 33, 34, 35)
    .Initexit BallRelease, 90, 6
    .InitEntrySounds "fx_drain", "", ""
    .InitExitSounds SoundFX(SSolenoidOn,DOFContactors), SoundFX("BallRelease",DOFContactors)
    .Balls = 4
    .CreateEvents "bsTrough", Drain
    End With

  ' Creature Popper
  Set bsLKick = New cvpmSaucer
  With bsLKick
    .InitKicker sw28, 28, 162, 15, 1
        .InitExitVariance 1, 1
    .InitSounds "fx_saucerHit", SoundFX(SSolenoidOn,DOFContactors), SoundFX("LeftEject",DOFContactors)
    .CreateEvents "bsLKick", sw28
  End With

  FrankInit:TargetsInit:DracInit:CreatureInit
  UpdateGI 0,0:UpdateGI 1,0:UpdateGI 2,0
  lock.IsDropped = 1              'ramp lock post
End Sub

'************************************************************************
'             KEYS
'************************************************************************

Sub Table1_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(11) = 1
    If keycode = LockBarKey then controller.switch(11) = 1
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge",0)
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge",0)
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge",0)
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(11) = 0
    If keycode = LockBarKey then controller.switch(11) = 0
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Table1_Paused:Controller.Pause = True:End Sub
Sub Table1_unPaused:Controller.Pause = False:End Sub
Sub Table1_exit():Controller.Pause = False:Controller.Stop:End Sub

'************************************************************************
'            SOLENOIDS
'************************************************************************

SolCallback(1) = "AutoPlunger"                          'AutoPlunger
SolCallback(2) = "SolBride"                           'Bride Post
SolCallback(3) = "SolMummy"                           'Mummy Coffin
SolCallback(5) = "vpmSolGate LGate,false,"                    'Left Gate
SolCallback(6) = "vpmSolGate Rgate,false,"                    'Right Gate
SolCallback(7) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"        'Knocker
SolCallback(8) = "SolLockPost"                          'Ramp Lock Post
SolCallback(9) = "SolBallRelease"                       'Trough Eject
SolCallback(10) = "vpmSolSound SoundFX(""LeftSlingshot"",DOFContactors),"   'Left Sling
SolCallback(11) = "vpmSolSound SoundFX(""RightSlingshot"",DOFContactors),"    'Right Sling
SolCallback(12) = "vpmSolSound SoundFX(""LeftJet"",DOFContactors),"       'Left Bumper
SolCallback(13) = "vpmSolSound SoundFX(""RightJet"",DOFContactors),"      'Right Bumper
SolCallback(14) = "vpmSolSound SoundFX(""BottomJet"",DOFContactors),"     'Bottom Bumper
SolCallback(15) = "bsLKick.SolOut"                        'Left Eject
SolCallback(16) = "SolRightScoop"                       'Right Popper
SolModCallback(17) = "SetModLamp 117,"                      'WolfMan Flashers
SolModCallback(18) = "SetModLamp 118,"                      'Bride Flashers
SolModCallback(19) = "SetModLamp 119,"                      'Frank Flashers
SolModCallback(20) = "SetModLamp 120,"                      'Dracula Flashers
SolModCallback(21) = "SolCreature"                        'Creature Flashers
SolModCallback(22) = "SetModLamp 122,"                      'Mummy Flashers
SolModCallback(23) = "SetModLamp 123,"                      'Right Popper Flashers
SolModCallback(24) = "SetModLamp 124,"                      'Frank Arrow Flashers
SolModCallback(25) = "SetModLamp 125,"                      'Monsters of rock Flashers
SolCallback(26) = "SetLamp 126,"                        'WolfMan Loop Flashers
SolCallback(27) = "SolFrank"                          'Frank Motor
SolCallback(28) = "SolBank"                           'Bank Motor

SolCallback(41) = "SolDrac"                           'Dracula Motor

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'******************************************
'       FLIPPERS
'******************************************

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("fx_FlipperUp",DOFFlippers), LeftFlipper, VolFlip
    LeftFlipper.RotateToEnd
  Else
        PlaySoundAtVol SoundFX("fx_FlipperDown",DOFFlippers), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_FlipperUp",DOFFlippers), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_FlipperDown",DOFFlippers),RightFlipper, VolFlip
        RightFlipper.RotateToStart
     End If
 End Sub

'************************************************************************
'            AUTOPLUNGER
'************************************************************************

Sub AutoPlunger(Enabled)
    If Enabled Then
       Plunger.Fire
       PlaySoundAtVol SoundFX("AutoPlunger",DOFContactors), Plunger, 1
  End If
End Sub

'************************************************************************
'            DRAIN AND RELEASE
'************************************************************************

Sub SolBallRelease(Enabled)
    If Enabled Then
        If bsTrough.Balls Then vpmTimer.PulseSw 31
        bsTrough.ExitSol_On
    End If
End Sub

'************************************************************************
'            RIGHT SCOOP
'************************************************************************
Dim sbrake,sshake,BIK,Scoopforce,ballspeed
Sub SolRightScoop(enabled)
  RightHole.Enabled=0
  PlaysoundAtVol SoundFx("fx_mine_popper",DOFContactors), sw36, VolKick
  If BIK>1 Then Scoopforce=90:Else Scoopforce=120:End If  'change scoop force depending on how much balls inside the scoop.
  sw36.kick 0, Scoopforce, 1.56
  sw36.TimerInterval=500:sw36.TimerEnabled=1
End Sub

Sub sw36_hit:Controller.Switch(36)=1:PlaysoundAtVol "fx_kicker_catch", sw36, VolKick:End Sub
Sub sw36_unhit:Controller.Switch(36)=0:End Sub

Sub sw36_timer()
  Me.TimerEnabled=0:RightHole.Enabled=1     'after 0.5 seconds re-enable the "RightHole" kicker
  If BIK>0 Then
    BIK=BIK-1
    If sw36.BallCntOver=0 Then BIK=0
  End If
End Sub

Sub RightScoop_hit:PlaysoundAtVol "fx_kicker_catch", sw36,VolKick:End Sub   'sound when hitting the wall
Sub CheckBallSpeed_hit():ballspeed=ActiveBall.VelY:End Sub
Sub RightHole_hit:BIK=BIK+1:PlaysoundAtVol "fx_mine_enter",RightHole,VolKick:If ballspeed < -45 Then sbrake=5:ScoopShake.Enabled=1:End If:End Sub

Sub ScoopShake_timer
  if sbrake>0 then
    sshake=sshake+30
    sbrake=sbrake-0.1
    ScoopP.transy =((dsin(sshake)))*sbrake
  else
    me.enabled=0:ScoopP.transy=0:sbrake=0:sshake=0
  End if
End Sub

'************************************************************************
'           SLINGSHOTS
'************************************************************************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 51
  LSling.Visible = 0
  LSling1.Visible = 1
  sling1.TransZ = -20
  LStep = 0
  Me.TimerInterval = 20
  LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 52
  RSling.Visible = 0
  RSling1.Visible = 1
  sling2.TransZ = -20
  RStep = 0
  RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'************************************************************************
'             BUMPERS
'************************************************************************

Sub Bumper1_Hit:vpmTimer.PulseSw 53::End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 54:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 55:End Sub

'************************************************************************
'           BRIDE ANIMATION
'************************************************************************

Sub SolBride(Enabled)
  If enabled then
    PlaySoundAtVol SoundFX("IdolReleaseOn",DOFContactors), BrideH, 1
    BrideH.Z = BrideH.Z + 20
  Else
    PlaySoundAtVol SoundFX("IdolReleaseOff",DOFContactors), BrideH, 1
    BrideH.Z = BrideH.Z - 20
  End if
End Sub

'************************************************************************
'             MUMMY ANIMATION
'************************************************************************
Dim cofdir, cbrake, cshake

Sub SolMummy(Enabled)
  if enabled then
    PlaySoundAtVol SoundFX("IdolReleaseOn",DOFContactors),mumcoffin,1
    cofdir=1:mummycoffin.enabled=1
    CoffinShake.Enabled=0:Mumcoffin.transy=0:Mumcoffin3.transy=0:cbrake=0:cshake=0
  else
    PlaySoundAtVol SoundFX("IdolReleaseOff",DOFContactors),mumcoffin,1
    cofdir=-1:mummycoffin.enabled=1
    CoffinShake.Enabled=0:Mumcoffin.transy=0:Mumcoffin3.transy=0:cbrake=0:cshake=0
  End If
End Sub

  Sub MummyCoffin_Timer()
  Mumcoffin.RotY = Mumcoffin.RotY + cofdir*10
  mumcoffin2.roty=Mumcoffin.RotY:mumcoffin3.roty=Mumcoffin.RotY
  If Mumcoffin.RotY >75 and cofdir=1 then Me.Enabled=0:CoffinShake.Enabled=1:cbrake=5
  If Mumcoffin.RotY <0 and cofdir=-1 then Me.Enabled=0
End Sub

sub CoffinShake_timer()
  if cbrake>0 then
    cshake=cshake+30
    cbrake=cbrake-0.1
    Mumcoffin.transy =((dsin(cshake)))*cbrake
    mumcoffin3.transy=Mumcoffin.transy
  else
    me.enabled=0:Mumcoffin.transy=0:Mumcoffin3.transy=0:cbrake=0:cshake=0
  End if
End Sub


'************************************************************************
'         RAMP LOCK POST
'************************************************************************

Sub SolLockPost(Enabled)
  if enabled then
    lock.isdropped=0:PlaySoundAtVol SoundFX("TopPostUp",DOFContactors),LockP,1:LockP.Z = 50 : RRPost.Enabled = 1
  else
    lock.isdropped=1:PlaySoundAtVol SoundFX("TopPostDown",DOFContactors),LockP,1:LockP.Z = -50
  End if
End Sub

'************************************************************************
'         CREATURE ANIMATION
'************************************************************************
Dim crot,cBall,ShakeEnabled:ShakeEnabled = 0:crot=0

Sub CreatureInit
  Set cBall = ckicker.createball
  ckicker.Kick 0, 0
End Sub

Sub SolCreature(value)
  SetModLamp 121,value
  If value>0 then
    ShakeEnabled = 1
    CreatureCover.visible = 0
    Creatureshake.enabled = 1
  Else
    CreatureCover.visible = 1
    CreatureWait.Interval=2000  'wait 2 seconds then stop the creature shake
    CreatureWait.enabled = 1
  End if
End Sub

Sub Creatureshake_timer
     crot=crot+0.35
     creature.rotz=creature.rotz-(Sin(crot))*0.2
end Sub

Sub CreatureWait_timer  ' stop the creature shake
  Me.Enabled=0: Creatureshake.Enabled=0:crot=0:creature.RotZ=0:ShakeEnabled = 0
End Sub

'************************************************************************
'           FRANK TARGETS BANK
'************************************************************************

Sub sw85_Hit:vpmTimer.PulseSw 85:bankshake.enabled=1:tBRAKE=5:PlaySoundAtVol SoundFX("fx_target",DOFTargets),franky,VolTarg:End Sub
Sub sw86_Hit:vpmTimer.PulseSw 86:bankshake.enabled=1:tBRAKE=5:PlaySoundAtVol SoundFX("fx_target",DOFTargets),franky,VolTarg:End Sub

Sub TargetsInit
  'starts the bank in down position
    frankytargets.z=-11
  sw85.Isdropped = 0
  sw86.Isdropped = 0
  Controller.Switch(81) = 1
  Controller.Switch(82) = 0
End Sub

Sub SolBank(enabled)
  If enabled then
    PlaySound SoundFX("fx_mine_motor", DOFGear), -1, 0.5, -0.1, 0, 0, 1, 0 ' TODO
    BankMove.enabled=1
  Else
    StopSound "fx_mine_motor"
    BankMove.enabled=0
  End If
End Sub

dim tdir,ftargvel
ftargvel=0.2:tdir=-1

Sub BankMove_timer()
    if tdir=-1 and frankytargets.z>-75.5 then
          frankytargets.z=frankytargets.z-ftargvel
          if frankytargets.z<-75 then Controller.Switch(82) = 1:Controller.Switch(81) = 0:sw85.isdropped=1:sw86.isdropped=1:tdir=1:Me.enabled=0
    End if
    if tdir=1 and frankytargets.z<-10 then
          frankytargets.z=frankytargets.z+ftargvel
          if frankytargets.z>-11 then Controller.Switch(81) = 1:Controller.Switch(82) = 0:sw85.isdropped=0:sw86.isdropped=0:tdir=-1:Me.enabled=0
    End if
End Sub

dim tbrake,tshake

sub bankshake_timer()
  if tbrake>0 then
    tshake=tshake+30
    tbrake=tbrake-0.08
    frankytargets.transy =((dsin(tshake)))*tbrake
  else
    tbrake=0:tshake=0:me.enabled=0
  End if
End Sub

'************************************************************************
'       FRANKY ANIMATION
'************************************************************************
'83 table down
'84 table up
'87 frank hit

Dim FrankDir,velfranky,compyfrank,compzfrank,fmove,canplay:canplay=0:velfranky=0.3

Sub frankinit
  franktable.rotx=0
  Controller.Switch(84) = 1
  Controller.Switch(87) = 1
  fhitwall.isdropped=1
End Sub

Sub SolFrank(enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("Bridge_Move", DOFGear), franky, 1
    if franky.rotx<-40 then FrankDir=-1:FrankMove.Enabled = 1:canplay=1
    if franky.rotx>-40 then FrankDir=1:FrankMove.Enabled = 1:canplay=1
  Else
    FrankMove.Enabled = 0::canplay=0
    StopSound "Bridge_Move":PlaysoundAtVol SoundFX("Bridge_Stop", 0), franky, 1
  End If
End Sub

'*******Franky body/arms/bracket/table rotation****************************************

Sub FrankMove_timer
    fmove=1
  Select Case FrankDir
    Case 1    'move frank down
    '---------------------------------------------------------------------------------------
      if franky.rotx<-85 then Controller.Switch(83) = 1:Controller.Switch(84) = 0:fhitwall.isdropped=0:fhitwallhelp.isdropped=0 else fhitwall.isdropped=1:fhitwallhelp.isdropped=1
      if franky.rotx>-80 then Controller.Switch(87) = 1 else Controller.Switch(87) = 0
    '---------------------------------------------------------------------------------------
      if franktable.rotx>-75 then
         franktable.rotx=franktable.rotx-velfranky
      If franksh.size_Y> 0.3 then franksh.size_y = franksh.size_y - velfranky/120:franksh.transY= franksh.transY + velfranky/120
    '------------------------------------------------------
         franky.rotx=franky.rotx-velfranky
         franky.z=114*(dcos(franktable.rotx+270))+138
         franky.y=114*(dsin(franktable.rotx-270))+622
         franky.transx=(dsin(franky.rotx))*5
    '------------------------------------------------------
         frankybraket.x=351 + (frankybraket.y/12)
         frankybraket.rotx=franktable.rotx/5
         frankybraket.z=122*(dcos(franktable.rotx+90))+140
         frankybraket.y=122*(dsin(franktable.rotx+270))+623

    '------------------------------------------------------
        compyfrank=((-156)*(dSIN(franky.rotx+90))+FRANKY.Y)
        arms.y=(-39*(dSin(franky.rotx))+compyfrank)
        compzfrank =((-156)*(dsin(franky.rotx+0))+FRANKY.z)
        arms.z=(39*(dcos(franky.rotx))+ compzfrank)
        arms.x=349+(dsin(franky.rotx+90))*-12
    '------------------------------------------------------
      End if
      if franktable.rotx<-70 then
        if canplay=1 then PlaySoundAtVol SoundFX("frankhit",DOFShaker), franky,1:canplay=0
        if franky.rotx>-90 then
          franky.rotx=franky.rotx-2
          franky.transx=(dsin(franky.rotx))*5
          compyfrank=((-156)*(dSIN(franky.rotx+90))+FRANKY.Y)
          arms.y=(-39*(dSin(franky.rotx))+compyfrank)
          compzfrank =((-156)*(dsin(franky.rotx+0))+FRANKY.z)
          arms.z=(39*(dcos(franky.rotx))+ compzfrank)
          arms.x=349+(dsin(franky.rotx+90))*-12
        else
          fshake=0:fbrake=1 : mainshake.enabled=1:fmove=0:Me.enabled=0
         End if
      End if

    Case -1   'move frank up
    '---------------------------------------------------------------------------------------
      if franky.rotx>-80 then Controller.Switch(87) = 1 else Controller.Switch(87) = 0
      if franky.rotx>-5 then Controller.Switch(83) = 0:Controller.Switch(84) = 1:fhitwall.isdropped=1:fhitwallhelp.isdropped=1
    '---------------------------------------------------------------------------------------
      if franktable.rotx<0 then
        franktable.rotx=franktable.rotx+velfranky
        If franksh.size_Y< 1 then franksh.size_y = franksh.size_y + velfranky/120:franksh.transY= franksh.transY - velfranky/120
    '------------------------------------------------------
        franky.rotx=franky.rotx+velfranky
        franky.z=114*(dcos(franktable.rotx+270))+138
        franky.y=114*(dsin(franktable.rotx-270))+622
        franky.transx=(dsin(franky.rotx))*5
    '------------------------------------------------------
        frankybraket.x=351 + (frankybraket.y/12)
        frankybraket.rotx=franktable.rotx/5
        frankybraket.z=122*(dcos(franktable.rotx+90))+140
        frankybraket.y=122*(dsin(franktable.rotx+270))+623
    '------------------------------------------------------
        compyfrank=((-156)*(dSIN(franky.rotx+90))+FRANKY.Y)
        arms.y=(-39*(dSin(franky.rotx))+compyfrank)
        compzfrank =((-156)*(dsin(franky.rotx+0))+FRANKY.z)
        arms.z=(39*(dcos(franky.rotx))+ compzfrank)
        arms.x=349+(dsin(franky.rotx+90))*-12
    '------------------------------------------------------
        if franktable.rotx>-55 then
          if franky.rotx< franktable.rotx then
            franky.rotx=franky.rotx+1
            compyfrank=((-156)*(dSIN(franky.rotx+90))+FRANKY.Y)
            arms.y=(-39*(dSin(franky.rotx))+compyfrank)
            compzfrank =((-156)*(dsin(franky.rotx+0))+FRANKY.z)
            arms.z=(39*(dcos(franky.rotx))+ compzfrank)
            arms.x=349+(dsin(franky.rotx+90))*-12
          End if
        End if
      else
        fmove=0:Me.enabled=0
      End if
  End Select
End Sub

'*******Franky Shake****************************************
dim fhit, fhitspeed

Sub fhitwall_Hit
  fhitspeed=Int(SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely))
  frankyshake.enabled=1:fhit=1
  PlaySoundAtVol SoundFX("frankhit",DOFShaker), franky, 1
End Sub

Sub frankyshake_timer
  if fmove=0 then
    if franky.rotx>-76 then Controller.Switch(87) = 1 else Controller.Switch(87) = 0
    if franky.rotx<-75 and fhit=1 then franky.rotx=franky.rotx+(fhitspeed/5) else fhit=0
    if franky.rotx>(-90+(fhitspeed*3))then fhit=0
    if fhit=0 and franky.rotx>-90  then franky.rotx=franky.rotx-0.5
    if franky.rotx<-89 then fshake=0 : fbrake=1 : mainshake.enabled=1:frankyshake.enabled=0
    arms.rotx=(franky.rotx+90)*2
    compyfrank=((-156)*(dSIN(franky.rotx+90))+FRANKY.Y)
    arms.y=(-39*(dSin(franky.rotx))+compyfrank)
    compzfrank =((-156)*(dsin(franky.rotx+0))+FRANKY.z)
    arms.z=(39*(dcos(franky.rotx))+ compzfrank)
    arms.x=349+(dsin(franky.rotx+90))*-12
  End if
End Sub

dim fbrake,fshake

Sub mainshake_timer()
  if fbrake>0 then
    fshake=fshake+30
    fbrake=fbrake-0.04
    franky.rotx =franky.rotx + ((dsin(fshake)))*fbrake
  else
    fbrake=0:fshake=0:mainshake.enabled=0
  End if

  arms.rotx=(franky.rotx+90)*4
  compyfrank=((-156)*(dSIN(franky.rotx+90))+FRANKY.Y)
  arms.y=(-39*(dSin(franky.rotx))+compyfrank)
  compzfrank =((-156)*(dsin(franky.rotx+0))+FRANKY.z)
  arms.z=(39*(dcos(franky.rotx))+ compzfrank)
  arms.x=349+(dsin(franky.rotx+90))*-12
End Sub

'************************************************************************
'         DRAC ANIMATION
'************************************************************************

Dim MechPos, dOldPos, dNewPos, dDir, DracStart

Sub DracInit
  Controller.Switch(25) = 0 : DracStart=0
  if dDir= 1 then Drac.rotz = 72 - controller.getmech(2)
  if dDir=-1 then Drac.rotz = 77 - controller.getmech(2)
  DracSupport.RotZ=Drac.RotZ:DracSh.RotZ=Drac.RotZ
  For Each xx in DracTargets:xx.IsDropped = 1:Next
  DracTargets(Int((12 + Drac.RotZ)/2+1.25)).IsDropped = 0
  For Each xx in BackSide:xx.IsDropped = 1:Next
  BackSide(Int((12 + Drac.RotZ)/2+1.25)).IsDropped = 0
End Sub

'*******Drac Hit****************************************
Sub DracTargets_Hit(idx)
    vpmTimer.PulseSw 25:DBRAKE=3:DracShake.enabled=1:PlaySoundAtVol SoundFX("drachit",DOFShaker), Drac, 1
End Sub

'*******Drac Move****************************************
Sub SolDrac(enabled)
  If Enabled Then
    DracStart=1:PlaySoundAtVol SoundFX("IdolMotor",DOFGear), Drac, 1
  Else
    StopSound "IdolMotor"
  End If
End Sub

Sub DracUpdate
  dOldPos=MechPos
  MechPos=controller.getmech(2)
  dNewPos=MechPos
    if dOldPos>dNewPos then dDir=1:DOF 102, DOFPulse
    if dOldPos<dNewPos then dDir=-1:DOF 102, DOFPulse
  If DracStart=0 and (MechPos = 5  or MechPos = 85) then exit sub   'this is used on init so drac starts always at 72 degrees or -13
'   'constant movement
'   if dDir= 1 and Drac.RotZ < (77 - MechPos) then Drac.RotZ=Drac.RotZ+0.13
'   if dDir=-1 and Drac.RotZ > (72 - MechPos) then Drac.RotZ=Drac.RotZ-0.13
    'chasing the vpm mech (it accelerates/decelerates depending on how far the drac toy is from the vpm mech)
    if dDir= 1 and Drac.RotZ < (77 - MechPos) then Drac.rotz = Drac.rotz - (Drac.rotz - (77- MechPos))/20
    if dDir=-1 and Drac.RotZ > (72 - MechPos) then Drac.rotz = Drac.rotz - (Drac.rotz - (72- MechPos))/20
  DracSupport.RotZ=Drac.RotZ:DracSh.RotZ=Drac.RotZ
  For Each xx in DracTargets:xx.IsDropped = 1:Next
  DracTargets(Int((12 + Drac.RotZ)/2+1.25)).IsDropped = 0
  For Each xx in BackSide:xx.IsDropped = 1:Next
  BackSide(Int((12 + Drac.RotZ)/2+1.25)).IsDropped = 0
End Sub

'*******Drac Shake****************************************
dim dbrake,dshake

Sub DracShake_timer
  if dbrake>0 then
    dshake=dshake+30
    dbrake=dbrake-0.04
    Drac.transy =((dsin(dshake)))*dbrake
    DracSh.transY=drac.TransY
  else
    dbrake=0:dshake=0:Me.enabled=0
  End if
End Sub

'************************************************************************
'         Switches
'************************************************************************

'Rollovers
Sub sw16_Hit:Controller.Switch(16) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw56_Hit:Controller.Switch(56) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw56_UnHit:Controller.Switch(56) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw64_UnHit:Controller.Switch(64) = 0:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw65_UnHit:Controller.Switch(65) = 0:End Sub

Sub sw66_Hit:Controller.Switch(66) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolTarg:End Sub
Sub sw66_UnHit:Controller.Switch(66) = 0:End Sub

' Ramps
Sub sw67_Hit:Controller.Switch(67) = 1:PlaySoundAtVol "fx_gate",sw67p,VolGates:sw67p.rotZ=-25:Me.TimerEnabled=1:End Sub
Sub sw67_timer:Me.TimerEnabled=0:sw67p.rotZ=0:End Sub
Sub sw67_UnHit:Controller.Switch(67) = 0:End Sub

Sub sw68_Hit:Controller.Switch(68) = 1:PlaySoundAtVol "fx_gate4", sw68, VolGates:End Sub
Sub sw68_UnHit:Controller.Switch(68) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:PlaySoundAtVol "fx_gate4", sw71, VolGates:End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

Sub sw72_Hit:Controller.Switch(72) = 1:PlaySoundAtVol "fx_gate", sw72p, VolGates: sw72p.rotZ=25:Me.TimerEnabled=1:End Sub
Sub sw72_timer:Me.TimerEnabled=0:sw72p.rotZ=0:End Sub
Sub sw72_UnHit:Controller.Switch(72) = 0:End Sub

Sub sw73_Hit:Controller.Switch(73) = 1:PlaySoundAtVol "fx_gate",sw73p, VolGates:sw73p.rotZ=25:Me.TimerEnabled=1:End Sub
Sub sw73_timer:Me.TimerEnabled=0:sw73p.rotZ=0:End Sub
Sub sw73_UnHit:Controller.Switch(73) = 0:End Sub

' Gates
Sub sw63_Hit:Controller.Switch(63) = 1:PlaySoundAtVol "fx_gate4",sw63, VolGates:End Sub
Sub sw63_UnHit:Controller.Switch(63) = 0:End Sub

' Opto & proximity switches
Sub sw42_Hit:Controller.Switch(42) = 1:End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

' Spinner
Sub Spinner1_spin():vpmTimer.PulseSw 117:PlaySoundAtVol "fx_spinner",spinner1, VolSpin:End Sub

' Drac Targets
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw12p, 1:sw12p.TransY=-5:Me.TimerEnabled=1:End Sub
Sub sw12_timer:Me.TimerEnabled=0:sw12p.TransY=0:End Sub
Sub sw15_Hit:vpmTimer.PulseSw 15:PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw15p, 1:sw15p.TransY=-5:Me.TimerEnabled=1:End Sub
Sub sw15_timer:Me.TimerEnabled=0:sw15p.TransY=0:End Sub

Sub sw23_Hit:vpmTimer.PulseSw 23:PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw23p, 1:sw23p.TransY=-5:Me.TimerEnabled=1:End Sub
Sub sw23_timer:Me.TimerEnabled=0:sw23p.TransY=0:End Sub
' Blue Targets
Sub sw44_Hit:vpmTimer.PulseSw 44:PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw44,VolTarg:End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw45,VolTarg:End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46:PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw46,VolTarg:End Sub


' *********************************************************************
'     GENERAL ILLUMINATION
' *********************************************************************

Set GiCallBack2 = GetRef("UpdateGi")

Sub UpdateGi(nr,step)
  Dim ii
  Select Case nr
  Case 0    'Bottom Playfield
    If step=0 Then
      For each ii in GIBot:ii.state=0:Next
    Else
      For each ii in GIBot:ii.state=1:Next
    End If
    Reflect1.Amount = 100 * (1 + step/2)
    Reflect2.IntensityScale = 0.125 * step
    For each ii in GIBot:ii.IntensityScale = 0.125 * step:Next
    If Step>=7 Then Table1.ColorGradeImage = "ColorGrade_8":Else Table1.ColorGradeImage = "ColorGrade_" & (step+1):End If
  Case 1    'Top Right Playfield
    If step=0 Then
      For each ii in GITopRight:ii.state=0:Next
      For each ii in GIBumpers:ii.state=0:Next
    Else
      For each ii in GITopRight:ii.state=1:Next
      For each ii in GIBumpers:ii.state=1:Next
    End If
    BWR.IntensityScale = 0.125 * step
    Reflect5.Amount = 100 * (1 + step/2)
    Reflect6.Amount = 100 * (1 + step/2)
    For each ii in GITopRight:ii.IntensityScale = 0.125 * step:Next
    For each ii in GIBumpers:ii.IntensityScale = 0.125 * step:Next
    If step>4 Then DOF 103, DOFOn : Else DOF 103, DOFOff:End If
  Case 2    'Top Left Playfield
    If step=0 Then
      For each ii in GITopLeft:ii.state=0:Next
    Else
      For each ii in GITopLeft:ii.state=1:Next
    End If
    BWL.IntensityScale = 0.125 * step
    Reflect3.Amount = 100 * (1 + step/2)
    Reflect4.Amount = 100 * (1 + step/2)
    For each ii in GITopLeft:ii.IntensityScale = 0.125 * step:Next
  Case 3    'Top Insert Panel

  Case 4    'Bottom Insert Panel
  End Select
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
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
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, indepEndent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps
  NFadeL 11, l11
  NFadeLm 12, l12c
  NFadeL 12, l12
  NFadeL 13, l13
  NFadeL 14, l14
  NFadeL 15, l15
  NFadeL 16, l16
  NFadeL 17, l17
  NFadeL 18, l18

  NFadeL 21, l21
  NFadeL 22, l22
  NFadeL 23, l23
  NFadeLm 24, l24c
  NFadeL 24, l24
  NFadeL 25, l25
  NFadeL 26, l26
  NFadeL 27, l27
  NFadeL 28, l28

  NFadeLm 31, l31c
  NFadeL 31, l31
  NFadeLm 32, l32c
  NFadeL 32, l32
  NFadeLm 33, l33c
  NFadeL 33, l33
  NFadeL 34, l34
  NFadeL 35, l35
  NFadeL 36, l36
  NFadeL 37, l37
  NFadeL 38, l38

  NFadeL 41, l41
  NFadeL 42, l42
  NFadeLm 43, l43c
  NFadeL 43, l43
  NFadeLm 44, l44c
  NFadeL 44, l44
  NFadeL 45, l45
  NFadeL 46, l46
  NFadeL 47, l47
  NFadeL 48, l48

  NFadeL 51, l51
  NFadeL 52, l52
  NFadeL 53, l53
  NFadeL 54, l54
  NFadeL 55, l55
  NFadeL 56, l56
  NFadeL 57, l57
  NFadeLm 58, l58c
  NFadeL 58, l58

  NFadeL 61, l61
  NFadeL 62, l62
  NFadeL 63, l63
  NFadeL 64, l64
  NFadeL 65, l65
  NFadeL 66, l66
  NFadeL 67, l67
  NFadeL 68, l68

  NFadeLm 71, l71
  NFadeL 71, l71b

  NFadeLm 72, l72
  NFadeL 72, l72b
  NFadeL 73, l73
  NFadeL 74, l74
  NFadeLm 75, l75
  NFadeL 75, l75b
  NFadeLm 76, l76b
  NFadeL 76, l76
  NFadeL 77, l77

  NFadeLm 81, l81pp
  Flash 81, l81f
  NFadeLm 82, l82pp
  Flash 82, l82f
  NFadeLm 83, l83pp
  Flash 83, l83f
  NFadeLm 84, l84pp
  Flash 84, l84f
  NFadeL 85, l85
  NFadeL 86, l86

  '*****Flashers*****
  FadeModLamp 117, f17, 1     'WolfMan Flashers
  FadeModLamp 117, f17a, 1
  FadeModLamp 117, f17b, 1
  FadeModLamp 118, f18, 2     'Bride Flasher
  FadeModLamp 118, f18a, 2
  FadeModLamp 118, f18b, 2
  FadeModLamp 118, f18c, 2
  FadeModLamp 118, f18r, 2
  FadeModLamp 119, f19, 2     'Frank Flashers
  FadeModLamp 119, f19a, 2
  FadeModLamp 119, f19b, 2
  FadeModLamp 119, f19c, 2
  FadeModLamp 119, f19d, 2
  FadeModLamp 119, f19e, 2
  FadeModLamp 119, f19f, 2
  FadeModLamp 120, f20, 2     'Dracula Flasher
  FadeModLamp 120, f20a, 2
  FadeModLamp 120, f20b, 2
  FadeModLamp 121, f21, 2     'Creature Flashers
  FadeModLamp 121, f21a, 2
  FadeModLamp 121, f21b, 2
  FadeModLamp 122, f22, 2     'Mummy Flashers
  FadeModLamp 122, f22a, 2
  FadeModLamp 122, f22b, 2
  FadeModLamp 122, f22c, 2
  FadeModLamp 122, f22d, 2
  FadeModLamp 123, f23, 2     'Right Popper Flasher
  FadeModLamp 123, f23a, 2
  FadeModLamp 123, f23b, 2
  FadeModLamp 123, f23c, 2
  FadeModLamp 123, f23d, 2
  FadeModLamp 123, f23r, 2
  FadeModLamp 124, f24, 2     'Frank Arrow Flasher
  FadeModLamp 125, f25, 2     'Monsters of rock Flasher
  NFadeLm 126, f26        'WolfMan Loop Flashers
  NFadeL 126, f26a
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
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

' Modulated lights and flashers

Sub SetModLamp(nr, value)
    If value > 0 Then
    LampState(nr) = 1
  Else
    LampState(nr) = 0
  End If
  FadingLevel(nr) = value
End Sub

Sub FadeModLamp(nr, object, factor)
  Object.IntensityScale = FadingLevel(nr) * factor/255
  If TypeName(object) = "Light" Then
    Object.State = LampState(nr)
  End If
  If TypeName(object) = "Flasher" Then
    Object.visible = LampState(nr)
  End If
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / VolDiv)
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

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
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

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

' *********************************************************************
'             Other Sound FX
' *********************************************************************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

Sub LeftFlipper_Collide(parm)
    RandomSoundFlipper
End Sub

Sub Rightflipper_Collide(parm)
    RandomSoundFlipper
End Sub

Sub LGate_Hit: PlaySoundAtVol "fx_gate4", LGate, VolGates : End Sub
Sub RGate_Hit: PlaySoundAtVol "fx_gate4", RGate, VolGates : End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR((activeball.velx ^2) + (activeball.vely ^2))
  If finalspeed > 20 then
    PlaySound "fx_rubber", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case RndNum(1,3)
    Case 1 : PlaySound "fx_rubber_hit_1", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "fx_rubber_hit_2", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "fx_rubber_hit_3", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandomSoundFlipper()
  Select Case RndNum(1,3)
    Case 1 : PlaySound "fx_flip_hit_1", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "fx_flip_hit_2", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "fx_flip_hit_3", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandomSoundHole()
  Select Case RndNum(1,3)
    Case 1 : PlaySound "fx_hole1"
    Case 2 : PlaySound "fx_hole2"
    Case 3 : PlaySound "fx_hole3"
  End Select
End Sub

' *********************************************************************
'           Ball Drop & Ramp Sounds
' *********************************************************************
'shooter ramp
Sub ShooterStart_Hit():StopSound "fx_launchball":If ActiveBall.VelY < 0 Then PlaySoundAtVol "fx_launchball", ShooterStart,1:End If:End Sub  'ball is going up
Sub ShooterEnd_Hit:If ActiveBall.Z > 50  Then Me.TimerInterval=100:Me.TimerEnabled=1:End If:End Sub           'ball is flying
Sub ShooterEnd_Timer(): Me.TimerEnabled=0 : PlaySoundAtVol "fx_balldrop", ShooterEnd,1 : End Sub
'center ramp
Sub CREnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAtVol "fx_lrenter", CREnter,1:End If:End Sub     'ball is going up
Sub CREnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_metalrolling":StopSound "fx_lrenter":End If:End Sub   'ball is going down
Sub CRExit_Hit:StopSound "fx_lrenter" : Me.TimerInterval=100:Me.TimerEnabled=1:End Sub
Sub CRExit_Timer(): Me.TimerEnabled=0 : PlaySoundAtVol "fx_metalrolling", Primitive36,1 : End Sub
'left ramp
Sub LREnter_hit():PlaySoundAtVol "fx_metalrolling", ActiveBall,1:If Activeball.VelY<6 Then ActiveBall.VelY=12:End If:End Sub
Sub LRExit_Hit:StopSound "fx_metalrolling":Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub LRExit_timer:Me.TimerEnabled=0:PlaysoundAtVol "fx_wireramp_exit", sw26,1::End Sub
'right ramp
Sub RREnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAtVol "fx_lrenter",RREnter,1:End If:End Sub      'ball is going up
Sub RREnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_lrenter":End If:End Sub   'ball is going down
Sub RRail_Hit::If ActiveBall.VelY < 0 Then StopSound "fx_lrenter" : Me.TimerInterval=100:Me.TimerEnabled=1:End If:End Sub   'ball is going up
Sub RRail_UnHit:If ActiveBall.VelY > 0 Then StopSound "fx_metalrolling":End If:End Sub    'ball is going down
Sub RRail_Timer(): Me.TimerEnabled=0 : PlaySoundAtVol "fx_metalrolling",Primitive89,1 : End Sub
Sub RRPost_Hit():StopSound "fx_metalrolling":End Sub
Sub RRPost_UnHit():me.Enabled = 0:PlaySoundAtVol "fx_wireramp_exit", ActiveBall,1:End Sub
Sub RRExit_Hit:StopSound "fx_metalrolling":Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub RRExit_timer:Me.TimerEnabled=0:PlaysoundAtVol "fx_wireramp_exit",sw17,1:End Sub
' *********************************************************************
'           Left Ramp Hack
' *********************************************************************
'prevents the ball from stucking at ramp enter at low speeds
Sub LRHelp1_Hit:If ActiveBall.VelX<2 then ActiveBall.VelX=-5:End If:End Sub
Sub LRHelp2_Hit:If ActiveBall.VelY<2 then ActiveBall.VelX=-5:ActiveBall.VelY=-5:End If:End Sub
Sub LRHelp3_Hit:If ActiveBall.VelY<2 then ActiveBall.VelY=-5:End If:End Sub
Sub LRHelp4_Hit:If ActiveBall.VelY<1 then ActiveBall.VelX=4:ActiveBall.VelY=-4:End If:End Sub

' *********************************************************************
'       Left and Right Orbits Hack
' *********************************************************************
'slow down the ball when shooting at loops (orbis)
Sub LoopHelpL_Unhit():If ActiveBall.VelY > 20 Then ActiveBall.VelY = RndNum(18,20):End If:End Sub
Sub LoopHelpR_Unhit():If ActiveBall.VelY > 20 Then ActiveBall.VelY = RndNum(18,20):End If:End Sub

'******************************************************
'         RealTime Updates
'******************************************************
Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates()
  RollingSoundUpdate
  BallShadowUpdate
  DracUpdate
  If ShakeEnabled = 0 then creature.rotz = (ckicker.y - cball.y) / 8
  sw68p.RotX = GateLR.currentangle
  sw71p.RotX = GateRR.currentangle
  sw63p.RotX = GateR.currentangle
  'LeftFlipperSh.RotZ = LeftFlipper.currentangle
  'RightFlipperSh.RotZ = RightFlipper.currentangle
  LeftGateP.RotY = -LGate.currentangle*2/3
  RightGateP.RotY = -RGate.currentangle*2/3
  If gateLR.currentangle>75 Then sw68p1.RotX=-10 Else sw68p1.RotX=0
  If gateRR.currentangle>75 Then sw71p1.RotX=-10 Else sw71p1.RotX=0
  If gateR.currentangle>75 Then sw63p1.RotX=-10 Else sw63p1.RotX=0
End Sub

'*********** ROLLING SOUND *********************************
Const tnob = 5            ' total number of balls : 4 (trough) + 1 (Fake Ball)
Const fakeballs = 1         ' number of balls created on table start (rolling sound and ballshadow will be skipped)
ReDim rolling(tnob-fakeballs)

Sub InitRolling:Dim i:For i=0 to (tnob-fakeballs-1):rolling(i) = False:Next:End Sub

Sub RollingSoundUpdate()
    Dim BOT, b
    BOT = GetBalls
  ' stop the sound of deleted balls
  If UBound(BOT)<(tnob - 1) Then
    For b = (UBound(BOT)-fakeballs + 2) to (tnob-fakeballs)
      rolling(b-1) = False
      StopSound("fx_ballrolling" & b)
    Next
  End If
  ' exit the Sub if no balls on the table
    If UBound(BOT) = fakeballs-1 Then Exit Sub
  ' play the rolling sound for each ball
    For b = fakeballs to UBound(BOT)
        If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
            rolling(b-fakeballs) = True
            PlaySoundAtVol("fx_ballrolling" & (b-fakeballs+1)), -1, Vol(BOT(b) )/4, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(ActiveBall)
        Else
            If rolling(b-fakeballs) = True Then
                StopSound("fx_ballrolling" & (b-fakeballs+1))
                rolling(b-fakeballs) = False
            End If
        End If
    Next
End Sub

'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3,Ballshadow4)

Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls
  ' hide shadow of deleted balls
  If UBound(BOT)<(tnob-1) Then
    For b = (UBound(BOT)-fakeballs + 2) to (tnob-fakeballs)
      BallShadow(b-1).visible = 0
    Next
  End If
  ' exit the Sub if no balls on the table
    If UBound(BOT) = fakeballs-1 Then Exit Sub
  ' render the shadow for each ball
    For b = fakeballs to UBound(BOT)
    If BOT(b).X < Table1.Width/2 Then
      BallShadow(b-fakeballs).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) + 10
    Else
      BallShadow(b-fakeballs).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) - 10
    End If
      ballShadow(b-fakeballs).Y = BOT(b).Y + 15
    If BOT(b).Z > 20 Then
      BallShadow(b-fakeballs).visible = 1
    Else
      BallShadow(b-fakeballs).visible = 0
    End If
  Next
End Sub

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

Sub PlaySoundAtVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

