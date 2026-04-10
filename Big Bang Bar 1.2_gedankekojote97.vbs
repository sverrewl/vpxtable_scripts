''************ Big Bamg Bar
''************ Capcom, 1996
''************ vp9 version by unclewilly and jimmyfingers, artwork by grizz
''************ original 3dmodels by rom extracted from his Future Pinball release
''************ vpx conversion by ninuzzu, scanned textures, objects corrections and adjustments by clarkkent

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


'//////////////////////////////////////////////////////////////////////
'// OPTIONS
'//////////////////////////////////////////////////////////////////////

dim blacklight : blacklight = False 'False - Blacklight Off, True - Blacklight On

Const Siderails =  1               '0 - No Siderails, 1 - Siderails
Const BallBright = 1        '0 - Normal, 1 - Bright
Const cController = 3           '0=Use value defined in cController.txt, 1=VPinMAME, 2=UVP server, 3=B2S server, 4=B2S with DOF (disable VP mech sounds)
Const VolumeDial = 0.8
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      'Level of ramp rolling volume. Value between 0 and 1

Const BallSize = 50
Const BallMass = 1
Const AmbientBallShadowOn = 1
Const DynamicBallShadowsOn = 1
Const RubberizerEnabled = 1
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7
Const tnob = 10
Const lob = 0
Dim tablewidth: tablewidth = Table.width
Dim tableheight: tableheight = Table.height


Dim DesktopMode:DesktopMode = Table.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

   LoadVPM "01560000", "Capcom.VBS", 3.10
'
'Dim cNewController
'Sub LoadVPM(VPMver, VBSfile, VBSver)
' Dim FileObj, ControllerFile, TextStr
'
' On Error Resume Next
' If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
' ExecuteGlobal GetTextFile(VBSfile)
' If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
'
' cNewController = 1
' If cController = 0 then
'   Set FileObj=CreateObject("Scripting.FileSystemObject")
'   If Not FileObj.FolderExists(UserDirectory) then
'     Msgbox "Visual Pinball\User directory does not exist. Defaulting to vPinMame"
'   ElseIf Not FileObj.FileExists(UserDirectory & "cController.txt") then
'     Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
'     ControllerFile.WriteLine 1: ControllerFile.Close
'   Else
'     Set ControllerFile=FileObj.GetFile(UserDirectory & "cController.txt")
'     Set TextStr=ControllerFile.OpenAsTextStream(1,0)
'     If (TextStr.AtEndOfStream=True) then
'       Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
'       ControllerFile.WriteLine 1: ControllerFile.Close
'     Else
'       cNewController=Textstr.ReadLine: TextStr.Close
'     End If
'   End If
' Else
'   cNewController = cController
' End If
'
' Select Case cNewController
'   Case 1
'     Set Controller = CreateObject("VPinMAME.Controller")
'     If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
'     If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
'     If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
'   Case 2
'     Set Controller = CreateObject("UltraVP.BackglassServ")
'   Case 3,4
'     Set Controller = CreateObject("B2S.Server")
' End Select
' On Error Goto 0
'End Sub


  Const cGameName = "bbb109"
  Const UseSolenoids = 2
  Const UseLamps = 0
  Const UseSync = 1
  Const HandleMech = 0

  'Standard Sounds
  Const SSolenoidOn = "Solenoid"
  Const SSolenoidOff = ""
  Const SCoin = ""


'******Variables***************
  Dim xx
  Dim bsTrough, DTBank4, DTBank3, DTBank1, bsRHole
  Dim captb1, captb2, captb3, captb4
  Dim bsTHelp, bsSw46, bsSw47, bsSw48
  Dim bsALL, bsALR

'''*************************************************************
'''Toggle DOF sounds on/off based on cController value
'''*************************************************************
'Dim ToggleMechSounds
'Function SoundFX (sound)
'    If cNewController= 4 and ToggleMechSounds = 0 Then
'        SoundFX = ""
'    Else
'        SoundFX = sound
'    End If
'End Function
'
'Sub DOF(dofevent, dofstate)
' If cNewController>2 Then
'   If dofstate = 2 Then
'     Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
'   Else
'     Controller.B2SSetData dofevent, dofstate
'   End If
' End If
'End Sub

'********Table Init************
Sub Table_Init
    LoadLUT
    vpmInit Me
  With Controller
    .GameName = cGameName
    .SplashInfoLine = "Big Bang Bar, Capcom 1996"
    .HandleMechanics = 0
    .HandleKeyboard = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .ShowTitle = 0
    .Hidden = DesktopMode
  End With
    On Error Resume Next
    Controller.Run
    If Err Then MsgBox Err.Description
      On Error Goto 0

  'Nudging
      vpmNudge.TiltSwitch=10
      vpmNudge.Sensitivity=5
      vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

'********Trough***************
  Set bsTrough = New cvpmBallStack
  With bsTrough
    .InitSw 35,36,37,38,39,0,0,0
    .InitKick BallRelease, 90, 8
    .Balls = 4
  End With

'********Right Hole***********
  Set bsRHole = New cvpmBallStack
  With bsRHole
    .InitSaucer sw67, 67, 222, 18
    .KickZ = 0.4
    .InitExitSnd SoundFX("VUKOut",DOFContactors), SoundFX("Solenoid",DOFContactors)
    .KickForceVar = 2
  End With

     'Ball lock stacks
     Set bsSw46 = New cvpmBallStack
     With bsSw46
         .InitSw 0, 0, 0, 0, 0, 0, 0, 0
     End With

     Set bsSw47 = New cvpmBallStack
     With bsSw47
         .InitSw 0, 0, 0, 0, 0, 0, 0, 0
     End With

     Set bsSw48 = New cvpmBallStack
     With bsSw48
         .InitSw 0, 0, 0, 0, 0, 0, 0, 0
     End With

     Set bsTHelp = New cvpmBallStack  'To help if all 4 balls get caught in lower lock
     With bsTHelp
         .InitSw 0, 0, 0, 0, 0, 0, 0, 0
     End With

'********Alien Locks ***********
     Set bsALL = New cvpmBallStack
     With bsALL
         .InitSw 0, 0, 0, 0, 0, 0, 0, 0
         .InitKick sw61Out, 160, 1
     End With

     Set bsALR = New cvpmBallStack
     With bsALR
         .InitSw 0, 0, 0, 0, 0, 0, 0, 0
         .InitKick sw62Out, 190, 1
     End With

'********DropTargets
'     Set DTBank4 = New cvpmDropTarget
'       With DTBank4
'       .InitDrop Array(sw17,sw18,sw19,sw20), Array(17,18,19,20)
'        .Initsnd "", SoundFX("Drop_Target_Reset_1")
'       End With
'
'     Set DTBank3 = New cvpmDropTarget
'       With DTBank3
'       .InitDrop Array(sw49,sw50,sw51), Array(49,50,51)
'        .Initsnd "", SoundFX("Drop_Target_Reset_2")
'       End With
'
'     Set DTBank1 = New cvpmDropTarget
'       With DTBank1
'       .InitDrop sw58, 58
'        .Initsnd "", SoundFX("Drop_Target_Reset_3")
'       End With

'*******Captive balls
  ' Initialize Captive Ball System
  Controller.Switch(77) = True : Controller.Switch(78) = False : Controller.Switch(79) = True : Controller.Switch(80) = False
  BP1 = True : BP2 = True : BP3 = True : BP4 = True : BP1R = False : BP2R = False : BP3L = False : BP4L = False
  capmove1.enabled = true : capmove3.enabled = true : capmove3L.enabled = false: capmove1R.enabled = false
  capmove2.enabled = true : capmove4.enabled = true : capmove4L.enabled = true : capmove2R.enabled = true
  CreateBalls

'********Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

'********KickBack Init
  kickback.PullBack

'********Diverters Init
  DivLR.IsDropped=1
  DivTube1.isDropped=1

  End Sub

   Sub Table_Paused:Controller.Pause = 1:End Sub
   Sub Table_unPaused:Controller.Pause = 0:End Sub

'//////////////////////////////////////////////////////////////////////
'// Ball
'//////////////////////////////////////////////////////////////////////

If BallBright Then
  table.BallImage = "ball_HDR_brighter"
Else
  table.BallImage = "MRBallDark2b"
End If


'//////////////////////////////////////////////////////////////////////
'// Dynamic Ball Shadows
'//////////////////////////////////////////////////////////////////////

Const fovY          = 0   'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const DynamicBSFactor     = 0.95  '0 to 1, higher is darker
Const AmbientBSFactor     = 0.7 '0 to 1, higher is darker
Const AmbientMovement   = 2   '1 to 4, higher means more movement as the ball moves left and right
Const Wideness        = 20  'Sets how wide the dynamic ball shadows can get (20 +5 thinness should be most realistic for a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source


' *** This segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
' ' stop the sound of deleted balls
' For b = UBound(BOT) + 1 to tnob
'   If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
'   ...rolling(b) = False
'   ...StopSound("BallRoll_" & b)
' Next
'
'...rolling and drop sounds...

'   If DropCount(b) < 5 Then
'     DropCount(b) = DropCount(b) + 1
'   End If
'
'   ' "Static" Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     If BOT(b).Z > 30 Then
'       BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
'     Else
'       BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
'     End If
'     BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
'     BallShadowA(b).X = BOT(b).X
'     BallShadowA(b).visible = 1
'   End If

' *** Required Functions, enable these if they are not already present elswhere in your table
Function DistanceFast(x, y)
  dim ratio, ax, ay
  ax = abs(x)         'Get absolute value of each vector
  ay = abs(y)
  ratio = 1 / max(ax, ay)   'Create a ratio
  ratio = ratio * (1.29289 - (ax + ay) * ratio * 0.29289)
  if ratio > 0 then     'Quickly determine if it's worth using
    DistanceFast = 1/ratio
  Else
    DistanceFast = 0
  End if
end Function

Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

'Dim PI: PI = 4*Atn(1)

'Function Atn2(dy, dx)
' If dx > 0 Then
'   Atn2 = Atn(dy / dx)
' ElseIf dx < 0 Then
'   If dy = 0 Then
'     Atn2 = pi
'   Else
'     Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
'   end if
' ElseIf dx = 0 Then
'   if dy = 0 Then
'     Atn2 = 0
'   else
'     Atn2 = Sgn(dy) * pi / 2
'   end if
' End If
'End Function
'
'Function AnglePP(ax,ay,bx,by)
' AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******
Dim sourcenames, currentShadowCount, DSSources(30), numberofsources, numberofsources_hold
sourcenames = Array ("","","","","","","","","","","","")
currentShadowCount = Array (0,0,0,0,0,0,0,0,0,0,0,0)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objrtx1(12), objrtx2(12)
dim objBallShadow(12)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5,BallShadowA6,BallShadowA7,BallShadowA8,BallShadowA9,BallShadowA10,BallShadowA11)

DynamicBSInit

sub DynamicBSInit()
  Dim iii, source

  for iii = 0 to tnob                 'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = iii/1000 + 0.01
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = (iii)/1000 + 0.02
    objrtx2(iii).visible = 0

    currentShadowCount(iii) = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = iii/1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100*AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source in DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
    iii = iii + 1
  Next
  numberofsources = iii
  numberofsources_hold = iii
end sub


Sub DynamicBSUpdate
  Dim falloff:  falloff = 150     'Max distance to light sources, can be changed if you have a reason
  Dim ShadowOpacity, ShadowOpacity2
  Dim s, Source, LSd, currentMat, AnotherSource, BOT, iii
  BOT = GetBalls

  'Hide shadow of deleted balls
  For s = UBound(BOT) + 1 to tnob
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(BOT) < lob Then Exit Sub    'No balls in play, exit

'The Magic happens now
  For s = lob to UBound(BOT)

' *** Normal "ambient light" ball shadow
  'Layered from top to bottom. If you had an upper pf at for example 80 and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

    If AmbientBallShadowOn = 1 Then     'Primitive shadow on playfield, flasher shadow in ramps
      If BOT(s).Z > 30 Then             'The flasher follows the ball up ramps while the primitive is on the pf
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + BallSize/10 + fovY
        objBallShadow(s).visible = 1

        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf, primitive only
        objBallShadow(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + fovY
        BallShadowA(s).visible = 0
      Else                      'Under pf, no shadows
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 0
      end if

    Elseif AmbientBallShadowOn = 2 Then   'Flasher shadow everywhere
      If BOT(s).Z > 30 Then             'In a ramp
        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf
        BallShadowA(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          BallShadowA(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          BallShadowA(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        BallShadowA(s).Y = BOT(s).Y + Ballsize/10 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/2 + 5
      Else                      'Under pf
        BallShadowA(s).visible = 0
      End If
    End If

' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If BOT(s).Z < 30 Then 'And BOT(s).Y < (TableHeight - 200) Then 'Or BOT(s).Z > 105 Then    'Defining when and where (on the table) you can have dynamic shadows
        For iii = 0 to numberofsources - 1
          LSd=DistanceFast((BOT(s).x-DSSources(iii)(0)),(BOT(s).y-DSSources(iii)(1))) 'Calculating the Linear distance to the Source
          If LSd < falloff And gilvl > 0 Then               'If the ball is within the falloff range of a light and light is on (we will set numberofsources to 0 when GI is off)
            currentShadowCount(s) = currentShadowCount(s) + 1   'Within range of 1 or 2
            if currentShadowCount(s) = 1 Then           '1 dynamic shadow source
              sourcenames(s) = iii
              currentMat = objrtx1(s).material
              objrtx2(s).visible = 0 : objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01            'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-LSd)/falloff                 'Sets opacity/darkness of shadow by distance to light
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness           'Scales shape of shadow with distance/opacity
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^2,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-ShadowOpacity),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-ShadowOpacity)
              End If

            Elseif currentShadowCount(s) = 2 Then
                                  'Same logic as 1 shadow, but twice
              currentMat = objrtx1(s).material
              AnotherSource = sourcenames(s)
              objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(AnotherSource)(0),DSSources(AnotherSource)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-DistanceFast((BOT(s).x-DSSources(AnotherSource)(0)),(BOT(s).y-DSSources(AnotherSource)(1))))/falloff
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0

              currentMat = objrtx2(s).material
              objrtx2(s).visible = 1 : objrtx2(s).X = BOT(s).X : objrtx2(s).Y = BOT(s).Y + fovY
  '           objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx2(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity2 = (falloff-LSd)/falloff
              objrtx2(s).size_y = Wideness*ShadowOpacity2+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity2*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2))
              End If
            end if
          Else
            currentShadowCount(s) = 0
            BallShadowA(s).Opacity = 100*AmbientBSFactor
          End If
        Next
      Else                  'Hide dynamic shadows everywhere else
        objrtx2(s).visible = 0 : objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub

'//////////////////////////////////////////////////////////////////////
'// LUT
'//////////////////////////////////////////////////////////////////////

'//////////////---- LUT (Colour Look Up Table) ----//////////////
'0 = Fleep Natural Dark 1
'1 = Fleep Natural Dark 2
'2 = Fleep Warm Dark
'3 = Fleep Warm Bright
'4 = Fleep Warm Vivid Soft
'5 = Fleep Warm Vivid Hard
'6 = Skitso Natural and Balanced
'7 = Skitso Natural High Contrast
'8 = 3rdaxis Referenced THX Standard
'9 = CalleV Punchy Brightness and Contrast
'10 = HauntFreaks Desaturated
'11 = Tomate Washed Out
'12 = VPW Original 1 to 1
'13 = Bassgeige
'14 = Blacklight
'15 = B&W Comic Book

Dim LUTset, DisableLUTSelector, LutToggleSound, bLutActive
LutToggleSound = True
LoadLUT
'LUTset = 0     ' Override saved LUT for debug
SetLUT
DisableLUTSelector = 0  ' Disables the ability to change LUT option with magna saves in game when set to 1

Sub Table_exit:SaveLUT:Controller.Stop:End Sub

Sub SetLUT  'AXS
  Table.ColorGradeImage = "LUT" & LUTset
end sub

Sub LUTBox_Timer
  LUTBox.TimerEnabled = 0
  LUTBox.Visible = 0
End Sub

Sub ShowLUT
  LUTBox.visible = 1
  Select Case LUTSet
    Case 0: LUTBox.text = "Fleep Natural Dark 1"
    Case 1: LUTBox.text = "Fleep Natural Dark 2"
    Case 2: LUTBox.text = "Fleep Warm Dark"
    Case 3: LUTBox.text = "Fleep Warm Bright"
    Case 4: LUTBox.text = "Fleep Warm Vivid Soft"
    Case 5: LUTBox.text = "Fleep Warm Vivid Hard"
    Case 6: LUTBox.text = "Skitso Natural and Balanced"
    Case 7: LUTBox.text = "Skitso Natural High Contrast"
    Case 8: LUTBox.text = "3rdaxis Referenced THX Standard"
    Case 9: LUTBox.text = "CalleV Punchy Brightness and Contrast"
    Case 10: LUTBox.text = "HauntFreaks Desaturated"
      Case 11: LUTBox.text = "Tomate washed out"
        Case 12: LUTBox.text = "VPW original 1on1"
        Case 13: LUTBox.text = "bassgeige"
        Case 14: LUTBox.text = "blacklight"
        Case 15: LUTBox.text = "B&W Comic Book"
    Case 16: LUTBox.text = "Skitso New ColorLut"
  End Select
  LUTBox.TimerEnabled = 1
End Sub

Sub SaveLUT
  Dim FileObj
  Dim ScoreFile

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if

  if LUTset = "" then LUTset = 0 'failsafe

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "BigBangBarLUT.txt",True)
  ScoreFile.WriteLine LUTset
  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub
Sub LoadLUT
    bLutActive = False
  Dim FileObj, ScoreFile, TextStr
  dim rLine

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    LUTset=0
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "BigBangBarLUT.txt") then
    LUTset=0
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "BigBangBarLUT.txt")
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    rLine = TextStr.ReadLine
    If rLine = "" then
      LUTset=0
      Exit Sub
    End if
    LUTset = int (rLine)
    Set ScoreFile = Nothing
      Set FileObj = Nothing
End Sub

'***************************************
'        KEYS
'***************************************

 Sub Table_KeyDown(ByVal keycode)
'LUT controls
If keycode = LeftMagnaSave Then bLutActive = True
    If keycode = RightMagnaSave Then
            If bLutActive Then
                    if DisableLUTSelector = 0 then
                LUTSet = LUTSet  - 1
                if LutSet < 0 then LUTSet = 16
                SetLUT
                ShowLUT
            End If
        End If
        End If


If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
                Select Case Int(rnd*3)
                        Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
                        Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
                        Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
End Select
        End If
If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress End If

If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress End If
If keycode = PlungerKey Then Plunger.PullBack:SoundPlungerPull()
If keycode = LeftTiltKey Then Nudge 90, 5:SoundNudgeLeft()
If keycode = RightTiltKey Then Nudge 270, 5:SoundNudgeRight()
If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter()
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table_KeyUp(ByVal keycode)
'LUT controls
If keycode = LeftMagnaSave Then bLutActive = False
  If vpmKeyUp(keycode) Then Exit Sub

If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress End If

If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress End If
    If keycode = PlungerKey Then Plunger.Fire : SoundPlungerReleaseBall()
End Sub

'******************************************
'Solenoids
'******************************************

  SolCallback(1) = "bsTrough.SolIn"               'OutHole
  SolCallback(2) = "bsTrough.SolOut"                'Ball Release
  SolCallback(3) = "vpmSolSound SoundFX(""knocker"",DOFContactors),"      'Knocker
  SolCallback(4) = "vpmSolSound SoundFX(""""),"   'LeftSling
  SolCallback(5) = "vpmSolSound SoundFX(""""),"   'RightSling
  SolCallback(6) = "SolKickBack"                  'KickBack
  SolCallback(7) = "sol4Bank"                   '4-Bank drop target
  SolCallback(8) = "SolLowerLockPin"                'Lower Lock Post
' SolCallback(9) = "SolLFlipper"                  'Left Flipper
' SolCallback(10) = "SolRFlipper"                 'Right Flipper
' SolCallback(11) = "SolURFlipper"                'Upper Flipper
  SolCallback(12) = "bsRHole.SolOut"                'Eject Hole
  SolCallback(13) = "SolLRDIvert"                 'Island Diverter
  SolCallback(14) = "SolRDivert1"                 'Ramp Diverter 1
  SolCallback(15) = "SolRDivert2"                 'Ramp Diverter 2
  SolCallback(16) = "SolRDivert3"                 'Alien Lock Post
  SolCallback(17) = "sol3Bank"                  '3-Bank drop target
  SolCallback(18) = "vpmSolSound SoundFX("""",DOFContactors),"    'Left Bumper
  SolCallback(19) = "vpmSolSound SoundFX("""",DOFContactors),"    'Middle Bumper
  SolCallback(20) = "vpmSolSound SoundFX("""",DOFContactors),"    'Right Bumper
  SolCallback(21)  ="Flash1"                'BackBox Left Flasher
  SolCallback(22)  ="Flash4"                'Tube Dancer Flasher
  SolCallback(23)  ="Flash5"                'Dance Floor Flasher
  SolCallback(24)  ="Flash2"                'Eject Hole Flasher
  SolCallback(25)  ="setlamp 165,"                'Alien Lock Flasher
  SolCallback(26)  ="Flash3"                'Lower Lock Flasher
  SolCallback(27) = "GateLeft"                  'Loop Left Gate
  SolCallback(28) = "GateRight"                 'Loop Right Gate
  SolCallback(29) = "sol1Bank"                  '1-Bank drop target
  SolCallback(30) = "solDancer"                 'Tube Dancer Motor
  SolCallback(31) = "SolAlienForward"               'Alien Motor Forward
  SolCallback(32)  ="SolAlienReverse"               'ALien Motor Reverse

SolCallback(51)  ="SolGameOn" 'Game on/off control

'******************************************
'        Flippers
'******************************************

'******Flipper fix stuff till emulation is fixed

dim GameOnFF:GameOnFF = 0
dim BIPFF:BIPFF = 0

 Sub FFTimer_Timer()
  If bsTrough.Balls = 4 or RightSlingshot.SlingshotStrength = 0 or BIPFF = 0 Then
    GameOnFF = 0
  else
    GameOnFF = 1
  end if
  End Sub

dim GameStateOn
Sub SolGameOn(Enabled)
If Enabled Then
GameStateOn = True
Else
GameStateOn = False
End If
End Sub

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'//////////////////////////////////////////////////////////////////////
'// FLIPPERS
'//////////////////////////////////////////////////////////////////////

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
  If Enabled and GameStateOn = True Then
    LF.Fire  'leftflipper.rotatetoend
Controller.Switch(33) = 1
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
Controller.Switch(33) = 0
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled and GameStateOn = True Then
    RF.Fire : RightFlipper1.RotateToEnd'rightflipper.rotatetoend
Controller.Switch(34) = 1
Controller.Switch(68) = 1
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
Controller.Switch(34) = 0
Controller.Switch(68) = 0
        RightFlipper1.RotateToStart
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

' This subroutine updates the flipper shadows and visual primitives
Sub FlipperVisualUpdate
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
  FlipperRSh1.RotZ = RightFlipper1.CurrentAngle
End Sub

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -2.7
'        AddPt "Polarity", 2, 0.33, -2.7
'        AddPt "Polarity", 3, 0.37, -2.7
'        AddPt "Polarity", 4, 0.41, -2.7
'        AddPt "Polarity", 5, 0.45, -2.7
'        AddPt "Polarity", 6, 0.576,-2.7
'        AddPt "Polarity", 7, 0.66, -1.8
'        AddPt "Polarity", 8, 0.743, -0.5
'        AddPt "Polarity", 9, 0.81, -0.5
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -3.7
'        AddPt "Polarity", 2, 0.33, -3.7
'        AddPt "Polarity", 3, 0.37, -3.7
'        AddPt "Polarity", 4, 0.41, -3.7
'        AddPt "Polarity", 5, 0.45, -3.7
'        AddPt "Polarity", 6, 0.576,-3.7
'        AddPt "Polarity", 7, 0.66, -2.3
'        AddPt "Polarity", 8, 0.743, -1.5
'        AddPt "Polarity", 9, 0.81, -1
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub
'
'


'*******************************************
'  Late 80's early 90's

'Sub InitPolarity()
' dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'   x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
' Next
'
' AddPt "Polarity", 0, 0, 0
' AddPt "Polarity", 1, 0.05, -5
' AddPt "Polarity", 2, 0.4, -5
' AddPt "Polarity", 3, 0.6, -4.5
' AddPt "Polarity", 4, 0.65, -4.0
' AddPt "Polarity", 5, 0.7, -3.5
' AddPt "Polarity", 6, 0.75, -3.0
' AddPt "Polarity", 7, 0.8, -2.5
' AddPt "Polarity", 8, 0.85, -2.0
' AddPt "Polarity", 9, 0.9,-1.5
' AddPt "Polarity", 10, 0.95, -1.0
' AddPt "Polarity", 11, 1, -0.5
' AddPt "Polarity", 12, 1.1, 0
' AddPt "Polarity", 13, 1.3, 0
'
' addpt "Velocity", 0, 0,         1
' addpt "Velocity", 1, 0.16, 1.06
' addpt "Velocity", 2, 0.41,         1.05
' addpt "Velocity", 3, 0.53,         1'0.982
' addpt "Velocity", 4, 0.702, 0.968
' addpt "Velocity", 5, 0.95,  0.968
' addpt "Velocity", 6, 1.03,         0.945
'
' LF.Object = LeftFlipper
' LF.EndPoint = EndPointLp
' RF.Object = RightFlipper
' RF.EndPoint = EndPointRp
'End Sub



'
''*******************************************
'' Early 90's and after
'
Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
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

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        LF.Object = LeftFlipper
        LF.EndPoint = EndPointLp
        RF.Object = RightFlipper
        RF.EndPoint = EndPointRp
End Sub


' Flipper trigger hit subs
Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub




'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
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

  Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
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
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

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
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
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
  redim aArray(aCount-1+offset)        'Resize original array
  for x = 0 to aCount-1                'set objects back into original array
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
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
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
  dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

  LinearEnvelope = Y
End Function


'******************************************************
'  FLIPPER TRICKS
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
  Dim b, BOT
  BOT = GetBalls

  If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          exit Sub
        end If
      Next
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
        end If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 then EOSNudge1 = 0
  End If
End Sub

'*****************
' Maths
'*****************
Dim PI: PI = 4*Atn(1)

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

Function Atn2(dy, dx)
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

'*************************************************
'  Check ball distance from Flipper for Rem
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


'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0:
    SOSRampup = 2.5
  Case 1:
    SOSRampup = 6
  Case 2:
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

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
    Dim b, BOT
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
  Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

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
  Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime : CatchTime = GameTime - FCount

  if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
    if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    else
      LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
    end If

    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
  End If
End Sub


'******************************************
'Drop Targets
'******************************************

Dim tdir


   Sub sol4Bank(Enabled)
    If Enabled Then
      tdir=1
    debug.print "SolDTDropUp"
    DTRaise 20
    DTRaise 19
    DTRaise 18
        DTRaise 17
    RandomSoundDropTargetReset sw19p
  End If
   End Sub

   Sub sol3Bank(Enabled)
    If Enabled Then
      tdir=1
    debug.print "SolDTDropUp"
    DTRaise 51
    DTRaise 50
    DTRaise 49
    RandomSoundDropTargetReset sw50p
  End If
   End Sub

   Sub sol1Bank(Enabled)
    If Enabled Then
      tDir = 1
    debug.print "SolDTDropUp"
    DTRaise 58
    RandomSoundDropTargetReset sw58p
  End If
   End Sub

Sub sw20_hit
  DTHit 20
End Sub

Sub sw19_hit
  DTHit 19
End Sub

Sub sw18_hit
  DTHit 18
End Sub

Sub sw17_hit
  DTHit 17
End Sub

Sub sw51_hit
  DTHit 51
End Sub

Sub sw50_hit
  DTHit 50
End Sub

Sub sw49_hit
  DTHit 49
End Sub

Sub sw58_hit
  DTHit 58
End Sub

'******************************************
'        DANCER ANIMATION
'******************************************

dim dance

sub solDancer(enabled)
if enabled then
dancerT.enabled=1
else
dancerT.enabled=0
end if
end sub

sub dancerT_timer()

dancer.ShowFrame(dance)
'dance = dance + (sin(dance*0.2854)+1)/2 old script, less exaggerated
dance = dance + (sin(dance*0.2854)*1.3+0.7)/2
If dance > 11 Then dance = 0 End If

end sub

'******************************************
'         Gates
'******************************************

 dim GateStateL, GateStateR : GateStateL = 0 : GateStateR = 0
 '
 Sub GateLeft(enabled)  : If enabled then : GateL.Open = true : GateStateL = 1 : vpmTimer.AddTimer 1000, "GateCloseL" : End If : End Sub
 Sub GateCloseL(aSw)    : GateL.Open = false  : GateStateL = 0 : End Sub

 Sub GateRight(enabled) : If enabled then : GateR.Open = true : GateStateR = 1 : vpmTimer.AddTimer 1000, "GateCloseR" : End If : End Sub
 Sub GateCloseR(aSw)    : GateR.Open = false : GateStateR = 0 : End Sub


'******************************************
'     Captive ball script
'   based on pacdudes script
'******************************************

  dim mBallDir, mBallCos, mBallSin
  dim mTrig, mWall, mKickers, mVelX, mVelY, mVelX2, mVelY2
  dim ForceTransL, ForceTransR, MinForce
  dim BP1,BP2,BP3,BP4,BP1R,BP2R,BP3L,BP4L

'   Default Settings
  ForceTransL = 1 : ForceTransR = 1 : MinForce = 12.5 : mBallDir = 10
  mBallCos = Cos(mBallDir * 3.1415927/180) : mBallSin = Sin(mBallDir * 3.1415927/180)

' Trigger Leads (Left/Right Side)
  Sub CapLLead_Hit() : mVelX  = activeBall.VelX : mVelY  = activeBall.VelY : End Sub
  Sub CapRLead_Hit() : mVelX2 = activeBall.VelX : mVelY2 = activeBall.VelY : End Sub

'   Left Side Captive Ball Hits (Up Motion)
  Sub CapLHit_Hit()
    dim force, dx, dy
    dX = ActiveBall.X - Captive1.X : dY = ActiveBall.Y - Captive1.Y
    force = -ForceTransL * (dY * mVelY + dX * mVelX) * (dY * mBallCos + dX * mBallSin) / (dX*dX + dY*dY)
    If force < 1 Then Exit Sub
    If force < MinForce Then force = MinForce
    If BP3L Then : BP3L = False : ForceTransL = 0.2 :PlaySound "fx_collide": Captive3L.DestroyBall : CapMove3L.CreateBall : CapMove3L.Kick mBallDir, force : Else :_
    If BP4L Then : BP4L = False : ForceTransL = 0.3 : Controller.Switch(78) = False :PlaySound "fx_collide": Captive4L.DestroyBall : CapMove4L.CreateBall : CapMove3L.enabled = False : CapMove4L.Kick mBallDir, force : Else :_
    If BP2  Then : BP2  = False : ForceTransL = 0.4:PlaySound "fx_collide": Captive2.DestroyBall : CapMove2.CreateBall : CapMove4L.enabled = False : CapMove2.Kick mBallDir, force : Else :_
    If BP1  Then : BP1  = False : Controller.Switch(77) = False :PlaySound "fx_collide": Captive1.DestroyBall : CapMove1.CreateBall : CapMove2.enabled = False : CapMove1.Kick mBallDir, force :_
    End If : End If : End If : End If
  End Sub

'   Right Side Captive Ball Hits (Up Motion)
  Sub CapRHit_Hit()
    dim force2, dx2, dy2
    dX2 = ActiveBall.X - Captive3.X : dY2 = ActiveBall.Y - Captive3.Y
    force2 = -ForceTransR * (dY2 * mVelY2 + dX2 * mVelX2) * (dY2 * mBallCos + dX2 * mBallSin) / (dX2*dX2 + dY2*dY2)
    If force2 < 1 Then Exit Sub
    If force2 < MinForce Then force2 = MinForce
    If BP1R Then : BP1R = False : ForceTransR = 0.2 :PlaySound "fx_collide": Captive1R.DestroyBall : CapMove1R.CreateBall : CapMove1R.Kick mBallDir, force2 : Else :_
    If BP2R Then : BP2R = False : ForceTransR = 0.3 : Controller.Switch(80) = False :PlaySound "fx_collide": Captive2R.DestroyBall : CapMove2R.CreateBall : CapMove1R.enabled = False : CapMove2R.Kick mBallDir, force2 : Else :_
    If BP4  Then : BP4  = False : ForceTransR = 0.4 :PlaySound "fx_collide": Captive4.DestroyBall : CapMove4.CreateBall : CapMove2R.enabled = False : CapMove4.Kick mBallDir, force2 : Else :_
    If BP3  Then : BP3  = False : Controller.Switch(79) = False :PlaySound "fx_collide": Captive3.DestroyBall : CapMove3.CreateBall : CapMove4.enabled = False : CapMove3.Kick mBallDir, force2 :_
    End If : End If : End If : End If
  End Sub

'   Left Side Return Hits (Down Motion)
  Sub CapMove1_Hit()
  If BP1 = True Then
   CapMove1.DestroyBall : CapMove2_Hit
  Else
    BP1  = True : Controller.Switch(77) = True : CapMove1.DestroyBall : Captive1.CreateBall : CapMove2.enabled = True
    End If
  End Sub
  Sub CapMove2_Hit()
  If BP2 = True Then
   CapMove2.DestroyBall : CapMove4L_Hit
  Else
    BP2  = True : ForceTransL = 0.4 : CapMove2.DestroyBall : Captive2.CreateBall : CapMove4L.enabled = True
    end if
  End Sub
  Sub CapMove4L_Hit()
  If BP4L = True Then
   CapMove4L.DestroyBall : CapMove3L_Hit
  Else
    BP4L = True : ForceTransL = 0.3 : Controller.Switch(78) = True : CapMove4L.DestroyBall : Captive4L.CreateBall : CapMove3L.enabled = True
    End If
  End Sub
  Sub CapMove3L_Hit()
    BP3L = True : ForceTransL = 0.2 : CapMove3L.DestroyBall : Captive3L.CreateBall
  End Sub

' Right Side Return Hits (Down Motion)
  Sub CapMove3_Hit()
  If BP3 = True Then
   CapMove3.DestroyBall : CapMove4_Hit
  Else
    BP3  = True : Controller.Switch(79) = True : CapMove3.DestroyBall : Captive3.CreateBall : CapMove4.enabled = True
    End If
  End Sub
  Sub CapMove4_Hit()
  If BP4 = True Then
   CapMove4.DestroyBall : CapMove2R_Hit
  Else
    BP4  = True : ForceTransR = 0.4 : CapMove4.DestroyBall : Captive4.CreateBall : CapMove2R.enabled = True
    End If
  End Sub
  Sub CapMove2R_Hit()
  If BP2R = True Then
   CapMove2R.DestroyBall : CapMove1R_Hit
  Else
    BP2R = True : ForceTransR = 0.3 : Controller.Switch(80) = True : CapMove2R.DestroyBall : Captive2R.CreateBall : CapMove1R.enabled = True
    End If
  End Sub
  Sub CapMove1R_Hit()
    BP1R = True : ForceTransR = 0.2 : CapMove1R.DestroyBall : Captive1R.CreateBall
  End Sub

'   Destroy any Existing Balls and Create the 4 Captive Balls (either real or reel-based)
  Sub CreateBalls()
  Captive1.DestroyBall  : Captive2.DestroyBall  : Captive3.DestroyBall  : Captive4.DestroyBall
  Captive4L.DestroyBall : Captive3L.DestroyBall : Captive1R.DestroyBall : Captive2R.DestroyBall
    If BP1  Then : Captive1.CreateBall  : End If : If BP2  Then : Captive2.CreateBall  : End If : If BP4L Then : Captive4L.CreateBall : End If
    If BP3L Then : Captive3L.CreateBall : End If : If BP3  Then : Captive3.CreateBall  : End If : If BP4  Then : Captive4.CreateBall  : End If
    If BP2R Then : Captive2R.CreateBall : End If : If BP1R Then : Captive1R.CreateBall : End If
End Sub


'******************************************
'      LowerLock And Post Handling
'******************************************

dim postdwn : postdwn = 0
  Sub SolLowerLockPin(Enabled)
  if enabled then
    If MissionLockPin.IsDropped=0 Then:Playsound SoundFX("diverter",DOFContactors):End If
    If postdwn = 0 Then vpmTimer.AddTimer 500, "PostUp"
    MissionLockPin.IsDropped=1:postdwn = 1
    If bsSw48.Balls>0 Then
      sw48H.DestroyBall:sw48.CreateSizedballwithMass BSize,BMass:bsSw48.SolOut 1:sw48.Kick 210, 1
      If bsSw48.Balls = 0 Then
        controller.switch(48) = 0
      End If
    End If
    If bsSw47.Balls>0 Then
      sw47H.DestroyBall:sw47.CreateSizedballwithMass BSize,BMass:bsSw47.SolOut 1:sw47.Kick 180, 1
      If bsSw47.Balls = 0 Then
        controller.switch(47) = 0
      End If
    End If
    If bsSw46.Balls>0 Then
      sw46H.DestroyBall:sw46.CreateSizedballwithMass BSize,BMass:bsSw46.SolOut 1:sw46.Kick 180, 1
      If bsSw46.Balls = 0 Then
        controller.switch(46) = 0
      End If
    End If
    If bsTHelp.Balls>0 Then
      swTHelpH.DestroyBall:swTHelp.CreateSizedballwithMass BSize,BMass:bsTHelp.SolOut 1:swTHelp.Kick 180, 1
    End If
  End If
  End Sub

  Sub PostUp(aSw)
  postdwn = 0:MissionLockPin.IsDropped=0
  End Sub

  Sub swTHelp_Hit()
  Dim BSpeed
' BIPFF = BIPFF - 1
  If bsTHelp.Balls>0 Then
    swTHelp.DestroyBall:sw46_Hit
  Else
    BSpeed = ActiveBall.VelY
    If bsSw46.Balls = 0 then
      swTHelp.Kick 180, BSpeed
    Else
      swTHelp.DestroyBall:swTHelpH.CreateSizedballwithMass BSize,BMass:bsTHelp.AddBall 1
    End If
  End If
  End Sub

  Sub sw46_Hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
  Dim BSpeed
  If bsSw46.Balls>0 Then
    sw46.DestroyBall:sw48_Hit
  Else
    BSpeed = ActiveBall.VelY
    If bsSw47.Balls = 0 then
      vpmTimer.PulseSw 46
      sw46.Kick 180, BSpeed
    Else
      controller.switch(46) = 1
      sw46.DestroyBall:sw46H.CreateSizedballwithMass BSize,BMass:bsSw46.AddBall 1
    End If
  End If
  End Sub

  Sub sw47_Hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
  Dim BSpeed
  If bsSw47.Balls>0 Then
    sw47.DestroyBall:sw46_Hit
  Else
    BSpeed = ActiveBall.VelY
    If bsSw48.Balls = 0 then
      vpmTimer.PulseSw 47
      sw47.Kick 180, BSpeed
    Else
      controller.switch(47) = 1
      sw47.DestroyBall:sw47H.CreateSizedballwithMass BSize,BMass:bsSw47.AddBall 1
    End If
  End If
  End Sub

  Sub sw48_Hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
  Dim BSpeed
  BSpeed = ActiveBall.VelY
  If bsSw48.Balls>0 and bsSw47.Balls>0 and bsSw46.Balls>0 Then
    sw48.Kick 210, BSpeed
  Else
    If bsSw48.Balls > 0 then
      sw48.DestroyBall:sw47_Hit
    Else
      controller.switch(48) = 1
      sw48.DestroyBall:sw48H.CreateSizedballwithMass BSize,BMass:bsSw48.AddBall 1
    End If
  End If
  End Sub

'  Sub sw48_UnHit():BIPFF = BIPFF + 1:End Sub

'******************************************
'      Left Kickback
'******************************************

  Sub SolKickBack(enabled)
  If enabled then
    PlaySoundAt "Popper", kickback
    kickback.Fire
  else
    kickback.PullBack
  End If
  End Sub

'******************************************
'      Ramp diverters
'******************************************

  Sub SolLRDIvert(Enabled)
  If Enabled then
    DivLR.IsDropped=0
  Else
    DivLR.IsDropped=1
  End If
  End Sub

  Sub SolRDivert1(Enabled)
  If Enabled then
    DivTubef.RotateToEnd
    DivTube.isDropped=1
    DivTube1.isDropped=0
  PlaySoundAt "scoop_exit" , Diverter2Sound
    vpmTimer.AddTimer 3000, "UnRDivert"
  End If
  End Sub
  Sub UnRDivert(aSw)
    DivTubef.RotateToStart
    DivTube.isDropped=0
    DivTube1.isDropped=1
  End Sub

Sub SolRDivert2(Enabled)
  If Enabled then
    DivTube2f.RotateToEnd
    DivTube2.isDropped=1
  PlaySoundAt "scoop_exit" , Diverter1Sound
  Else
    DivTube2f.RotateToStart
    DivTube2.isDropped=0
  End If
  End Sub

  Sub SolRDivert3(Enabled)
  If Enabled then
    AliensLockPin.IsDropped=1
  PlaySoundAt "scoop_exit" , Diverter3Sound
  Else
    AliensLockPin.IsDropped=0
  End If
  End Sub

'******************************************
' Alien Lock Mech Handler
'******************************************

'Note : The Startup timer is a workaround for motor overshoot caused by no power control option in VPM for the solenoids.
' - It sends an extra quarter cycle of switch changes with no animation before animation engages at actual overshoot position)
' - This prevents the alien mech from looping 2-3 times every time it starts due to position loss caused by overshoot in the
' - forward direction.  The alien mech has also been modified to correct a mis-position problem that happens sometimes after
' - Looped In Space mode ends.  The correction ensures the alien mech always faces home when it stops.  Since during the game
' - it only ever stops at the home position, this makes it fully functional, although the operator test mode doesn't function
' - as one would expect it to (making full loops instead of small movements), but that has no bearing on normal gameplay.
dim startstatus, checkstart : startstatus = 0 : checkstart = 0

  Sub SolAlienForward(Enabled)
  If enabled then
    forward = 1 : ALockTimer.enabled = 1
  Else
    forward = 0
  End If
  End Sub

  Sub SolAlienReverse(Enabled)
  If enabled then
    reverse = 1 : ALockTimer.enabled = 1
  Else
    reverse = 0
  End If
  End Sub

Sub ALockStartup_Timer() : If direction = 0 then : If checkstart = 0 and startstatus = 1 then : vpmTimer.AddTimer 1000, "CheckStatus" : End If : checkstart = 1 : Else : checkstart = 0 : End If : End Sub
Sub CheckStatus(aSw) : if checkstart = 1 Then : startstatus = 0 : OldPos = 31 : NewPos = 0 : End If : End Sub
dim NewPos, forward, reverse, OldPos : OldPos = 31 : NewPos = 0 : forward = 0 : reverse = 0
'dim RAlien:RAlien = Array(RM0,RM1,RM2,RM3,RM4,RM5,RM6,RM7,RM8,RM9)
'dim LAlien:LAlien = Array(RM10,RM19,RM18,RM17,RM16,RM15,RM14,RM13,RM12,RM11)

dim direction
Sub ALockTimer_timer()
  if reverse = 1 then : direction = -1 : end if : If forward = 1 then direction = 1 : end if
  NewPos = OldPos + direction
  If (NewPos > OldPos) and NewPos >= 32 Then NewPos = 0
  If (NewPos < OldPos) and NewPos  <  0 Then NewPos = 31
If startstatus = 0 and forward = 1 then
  if NewPos >=  7   then startstatus = 1
  Select Case NewPos
    Case 0 : controller.switch(57) = 0 ' Home Notch 1 (Reset Trick)
    Case 1 : controller.switch(57) = 1
    Case 2 : controller.switch(57) = 0 ' Home Notch 2 (Rest Trick)
    Case 3 : controller.switch(57) = 1
    Case 7 : controller.switch(57) = 1
  End Select
Else
' If OldPos <> NewPos Then PlaySound Soundfx ("Motor")
  ' Handle Optos
  If forward = 1 or reverse = 1 Then
  Select Case NewPos
    Case 0 : controller.switch(57) = 0 ' Home Notch 1
    Case 1 : controller.switch(57) = 1
    Case 2 : controller.switch(57) = 0 ' Home Notch 2
    Case 3 : controller.switch(57) = 1
    Case 7 : controller.switch(57) = 1
    Case 8 : controller.switch(57) = 0 ' Quarter Turn
    Case 9 : controller.switch(57) = 1
    Case 15: controller.switch(57) = 1
    Case 16: controller.switch(57) = 0 ' Half Turn
    Case 17: controller.switch(57) = 1
    Case 23: controller.switch(57) = 1
    Case 24: controller.switch(57) = 0 ' Three Quarters Turn
    Case 25: controller.switch(57) = 1
    Case 31: controller.switch(57) = 1
  End Select
  End If
  ' Animate Aliens
  Select Case NewPos
    Case  2 : Alien1.rotZ=-255:Alien2.rotZ=345:AW0018624.image = "AW00186-24p27"    'AlienReel.setvalue 8
    Case  3 : Alien1.rotZ=-270:Alien2.rotZ=0:AW0018624.image = "AW00186-24p28"    'AlienReel.setvalue 9
    Case  5 : Alien1.rotZ=-290:Alien2.rotZ=20:AW0018624.image = "AW00186-24p30"     'AlienReel.setvalue 9
    Case  6 : Alien1.rotZ=-315:Alien2.rotZ=45:AW0018624.image = "AW00186-24p31"     'AlienReel.setvalue 0
    Case  9 : Alien1.rotZ=-330:Alien2.rotZ=60:AW0018624.image = "AW00186-24p2"    'AlienReel.setvalue 0
    Case 10 : Alien1.rotZ=0:Alien2.rotZ=90:AW0018624.image = "AW00186-24p3"       'AlienReel.setvalue 1
    Case 11 : Alien1.rotZ=-15:Alien2.rotZ=105:AW0018624.image = "AW00186-24p5"      'AlienReel.setvalue 1
    Case 12 : Alien1.rotZ=-30:Alien2.rotZ=120:AW0018624.image = "AW00186-24p6"    'AlienReel.setvalue 2
    Case 15 : Alien1.rotZ=-45:Alien2.rotZ=135:AW0018624.image = "AW00186-24p9"    'AlienReel.setvalue 2
    Case 16 : Alien1.rotZ=-60:Alien2.rotZ=150:AW0018624.image = "AW00186-24p10"     'AlienReel.setvalue 3
    Case 18 : Alien1.rotZ=-75:Alien2.rotZ=165:AW0018624.image = "AW00186-24p11"     'AlienReel.setvalue 3
    Case 19 : Alien1.rotZ=-90:Alien2.rotZ=180:AW0018624.image = "AW00186-24p12"     'AlienReel.setvalue 4
    Case 21 : Alien1.rotZ=-120:Alien2.rotZ=200:AW0018624.image = "AW00186-24p15"    'AlienReel.setvalue 4
    Case 22 : Alien1.rotZ=-135:Alien2.rotZ=225:AW0018624.image = "AW00186-24p16"    'AlienReel.setvalue 5
    Case 24 : Alien1.rotZ=-160:Alien2.rotZ=250:AW0018624.image = "AW00186-24p18"    'AlienReel.setvalue 5
    Case 25 : Alien1.rotZ=-180:Alien2.rotZ=270:AW0018624.image = "AW00186-24p19"    'AlienReel.setvalue 6
    Case 27 : Alien1.rotZ=-195:Alien2.rotZ=285:AW0018624.image = "AW00186-24p21"    'AlienReel.setvalue 6
    Case 28 : Alien1.rotZ=-210:Alien2.rotZ=300:AW0018624.image = "AW00186-24p22"    'AlienReel.setvalue 7
    Case 30 : Alien1.rotZ=-225:Alien2.rotZ=315:AW0018624.image = "AW00186-24p24"    'AlienReel.setvalue 7
    Case 31 : Alien1.rotZ=-240:Alien2.rotZ=330:AW0018624.image = "AW00186-24p25"    'AlienReel.setvalue 8
  End Select

  'Eject Lock
  If (NewPos >= 12 and NewPos <= 15) and (bsALL.balls > 0 or bsALR.balls > 0) Then
    AlienEjectR : vpmTimer.Addtimer 300, "AlienEjectL"
  End If
End If
  OldPos = NewPos
  If direction = 1 Then
    If forward = 0 and reverse = 0 and NewPos >= 6 and NewPos <= 9 Then : Direction = 0 : ALockTimer.enabled = False : StopSound "Motor": End If
  Else
    If forward = 0 and reverse = 0 Then : Direction = 0 : ALockTimer.enabled = False : End If
  End If
End Sub

  ' Eject Ball Subs
Sub AlienEjectL(aSw) : bsALL.SolOut 1 : End Sub
Sub AlienEjectR()    : bsALR.SolOut 1 : End Sub

 Sub sw61Out_UnHit():BIPFF = BIPFF + 1:PlaySoundAt "KickandWire",sw61Out :End Sub
 Sub sw62Out_UnHit():BIPFF = BIPFF + 1:PlaySoundAt "KickandWire",sw62Out:End Sub

'******************************************
'     Drains and Kickers
'******************************************

  Sub Drain_Hit():RandomSoundDrain Drain :BIPFF = BIPFF - 1:bsTrough.AddBall Me:End Sub
  Sub BallRelease_UnHit():BIPFF = BIPFF + 1:RandomSoundBallRelease ballrelease :End Sub
  Sub sw61_Hit(): WireRampOff : BIPFF = BIPFF - 1:bsALL.AddBall Me:PlaySoundAt "VUKEnter", sw61:End Sub
  Sub sw62_Hit(): WireRampOff : BIPFF = BIPFF - 1:bsALR.AddBall Me:PlaySoundAt "VUKEnter", sw62:End Sub
  Sub sw67_Hit:Playsound "VUKEnter":bsRHole.AddBall Me:End Sub


'************
' SlingShots
'************


Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotRight sling1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    vpmTimer.PulseSw 52
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotLeft sling2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    vpmTimer.PulseSw 51
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'************************************************
'     Bumpers Animation
'************************************************

Dim dirRing1:dirRing1 = -1
Dim dirRing2:dirRing2 = -1
Dim dirRing3:dirRing3 = -1

Sub Bumper1_Hit:vpmTimer.PulseSw 56:me.TimerEnabled = 1: RandomSoundBumperTop Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 54:me.TimerEnabled = 1:RandomSoundBumperMiddle Bumper2 :End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 55:me.TimerEnabled = 1:RandomSoundBumperBottom Bumper3 :End Sub

Sub Bumper1_timer()
  Bumper1Ring.Z = Bumper1Ring.Z + (5 * dirRing1)
  If Bumper1Ring.Z <= -40 Then dirRing1 = 1
  If Bumper1Ring.Z >= 0 Then
    dirRing1 = -1
    Bumper1Ring.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper2_timer()
  Bumper2Ring.Z = Bumper2Ring.Z + (5 * dirRing2)
  If Bumper2Ring.Z <= -40 Then dirRing2 = 1
  If Bumper2Ring.Z >= 0 Then
    dirRing2 = -1
    Bumper2Ring.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper3_timer()
  Bumper3Ring.Z = Bumper3Ring.Z + (5 * dirRing3)
  If Bumper3Ring.Z <= -40 Then dirRing3 = 1
  If Bumper3Ring.Z >= 0 Then
    dirRing3 = -1
    Bumper3Ring.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

'************************************************
'     Switches
'************************************************

'*******Rollover switches
 Sub sw26_Hit:Controller.Switch(26) = 1:End Sub
 Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

 Sub sw27_Hit:Controller.Switch(27) = 1:End Sub
 Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

 Sub sw28_Hit:Controller.Switch(28) = 1:End Sub
 Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

 Sub sw29_Hit:Controller.Switch(29) = 1:End Sub
 Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

 Sub sw30_Hit:Controller.Switch(30) = 1:End Sub
 Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub

 Sub sw43_Hit:Controller.Switch(43) = 1:plungerbulb.state=2:End Sub
 Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

 Sub sw44_Hit:Controller.Switch(44) = 1:End Sub
 Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

 Sub sw45_Hit:Controller.Switch(45) = 1:End Sub
 Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

 Sub sw59_Hit:Controller.Switch(59) = 1:End Sub
 Sub sw59_UnHit:Controller.Switch(59) = 0:End Sub

 Sub sw60_Hit:Controller.Switch(60) = 1:End Sub
 Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

 Sub sw65_Hit:Controller.Switch(65) = 1:End Sub
 Sub sw65_UnHit:Controller.Switch(65) = 0:End Sub

 Sub sw66_Hit:Controller.Switch(66) = 1:End Sub
 Sub sw66_UnHit:Controller.Switch(66) = 0:End Sub

'*******Ramp Switches
 Sub sw24_Hit():Controller.Switch(24) = 1:End Sub
 Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

 Sub sw32_Hit():Controller.Switch(32) = 1:End Sub
 Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

 Sub sw69_Hit():Controller.Switch(69) = 1:End Sub
 Sub sw69_UnHit:Controller.Switch(69) = 0:End Sub

 Sub sw70_Hit():Controller.Switch(70) = 1:End Sub
 Sub sw70_UnHit:Controller.Switch(70) = 0:End Sub

 Sub sw71_Hit():Controller.Switch(71) = 1:End Sub
 Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

 Sub sw31_Hit():Controller.Switch(31) = 1:End Sub
 Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

 Sub sw61a_Hit():Controller.Switch(61) = 1:End Sub
 Sub sw61a_UnHit:Controller.Switch(61) = 0:End Sub

 Sub sw62a_Hit():Controller.Switch(62) = 1:End Sub
 Sub sw62a_UnHit:Controller.Switch(62) = 0:End Sub

'*******Spinner
  Sub sw25_Spin():vpmTimer.PulseSw 25:SoundSpinner sw25:End Sub
Sub Spinner2_Spin : SoundSpinner Spinner2 : End Sub
Sub Spinner1_Spin : SoundSpinner Spinner1 : End Sub

'*******StandUp Targets
  Sub sw21_Hit():vpmTimer.PulseSw 21:sw21p.transY=-5:me.TimerEnabled=1:End Sub
  Sub sw21_Timer():sw21p.transY=0:ME.TimerEnabled=0:End Sub

  Sub sw22_Hit():vpmTimer.PulseSw 22:sw22p.transY=-5:me.TimerEnabled=1:End Sub
  Sub sw22_Timer():sw22p.transY=0:ME.TimerEnabled=0:End Sub

  Sub sw23_Hit():vpmTimer.PulseSw 23:sw23p.transX=-5:me.TimerEnabled=1:End Sub
  Sub sw23_Timer():sw23p.transX=0:ME.TimerEnabled=0:End Sub

  Sub sw52_Hit():vpmTimer.PulseSw 52:sw52p.transX=-5:me.TimerEnabled=1:End Sub
  Sub sw52_Timer():sw52p.transX=0:ME.TimerEnabled=0:End Sub

  Sub sw53_Hit():vpmTimer.PulseSw 53:sw53p.transX=-5:me.TimerEnabled=1:End Sub
  Sub sw53_Timer():sw53p.transX=0:ME.TimerEnabled=0:End Sub

  Sub BallCatcherUnderground_Hit:If ActiveBall.VelX >= 16 then ActiveBall.VelX = 10:End if:End Sub
  Sub BallCatcherLeftOrbit_Hit:If ActiveBall.VelY <= -16 then ActiveBall.VelY = -16:End if:End Sub
  Sub BallCatcherRightOrbit_Hit:If ActiveBall.VelY <= -16 then ActiveBall.VelY = -16:End if:End Sub

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

dim gilvl:gilvl = 1


 Sub UpdateLamps
'GI
If LampState(9) = 0 and LampState(10) = 0 and lampstate(11)= 0 and lampstate (25)=0 then GILeft.state=0 : gilvl = 0 : GIShadows.visible=0
If LampState(9) = 1 and LampState(10) = 1 and lampstate(11)= 1 and lampstate (25)=1 then GILeft.state=1 : gilvl = 1 : GIShadows.visible=1
If LampState(12) = 0 and LampState(14) = 0  then GIBLeft.state=0 : gilvl = 0 : GIShadows.visible=0
If LampState(12) = 1 and LampState(14) = 1  then GIBLeft.state=1 : gilvl = 1 : GIShadows.visible=1
If LampState(17) = 0 and LampState(18) = 0 then GIRight.state=0 : gilvl = 0 : GIShadows.visible=0
If LampState(17) = 1 and LampState(18) = 1 then GIRight.state=1 : gilvl = 1 : GIShadows.visible=1
If LampState(21) = 0 and LampState(22) = 0 and lampstate(23)= 0 and lampstate(24)= 0 then GIBRight.state=0 : gilvl = 0 : GIShadows.visible=0
If LampState(21) = 1 and LampState(22) = 1 and lampstate(23)= 1 and lampstate(24)= 1 then GIBRight.state=1 : gilvl = 1 : GIShadows.visible=1
If LampState(33) = 0 and LampState(34) = 0 and lampstate(35)= 0 and lampstate (36)=0 then GITop.state=0 : gilvl = 0 : GIShadows.visible=0
If LampState(33) = 1 and LampState(34) = 1 and lampstate(35)= 1 and lampstate (36)=1 then GITop.state=1 : gilvl = 1 : GIShadows.visible=1

If blacklight Then
  If LampState(62) = 0 Then
    table.ColorGradeImage = "LUT" & LUTset
  Else
    table.ColorGradeImage = "LUT14"
  End If
End If

  NFadeL  9, l9     ' 4-Bank GI 1
  NFadeL 10, l10      ' 4-Bank GI 2
  NFadeL 11, l11      ' 4-Bank GI 3
  NFadeL 12, l12      ' Left Sling GI 1
  NFadeL 14, l14      ' Left Flipper GI 1
  NFadeL 17, l17      ' Upper Right Flipper GI 1
  NFadeL 18, l18      ' Eject Hole GI 1
  NFadeL 21, l21      ' Right Sling GI 1
  NFadeL 22, l22      ' Right Sling GI 2
  NFadeL 23, l23      ' Right Flipper GI 1
  NFadeL 24, l24      ' Right Flipper GI 2
  NFadeL 25, l25      ' Tube GI 1
  NFadeL 26, l26      ' Tube GI 2
  NFadeL 27, l27      ' Tube GI 3
  NFadeL 28, l28      ' Tube GI 4
  NFadeL 29, l29      ' Tube GI 5
  NFadeLm 33, l33     ' Hoot GI 1
  NFadeLm 33, l33a
  NFadeLm 34, l34     ' Hoot GI 2
  NFadeLm 33, l33b
  NFadeLm 35, l35     ' Hoot GI 3
  NFadeLm 33, l33c
  NFadeLm 36, l36     ' Hoot GI 4
  NFadeLm 33, l33d
  NFadeL 37, l37      ' Alien GI 1
  NFadeL 38, l38      ' Alien GI 2
  NFadeL 39, l39      ' Alien GI 3
  NFadeL 40, l40      ' Captive GI 1
  NFadeL 128, l128    ' Upper Right Flipper GI 2
  'NFadeL 24, SHADOWS_gi
'Inserts
  NFadeL 19, l19      ' Spaceship 1
  NFadeL 20, l20      ' Spaceship 2
  NFadeL 30, l30      ' Left Orbit Chase 1
  NFadeL 31, l31      ' Left Orbit Chase 2
  NFadeL 32, l32      ' Left Orbit Chase 3
  NFadeL 41, l41      ' Right Orbit Chase 1
  NFadeL 42, l42      ' Right Orbit Chase 2
  NFadeL 43, l43      ' Right Orbit Chase 3
  NFadeLm 49, l49     ' Rollover B
  Flash 49, l49r
  NFadeLm 50, l50     ' Rollover A
  Flash 50, l50r
  NFadeLm 51, l51     ' Rollover R
  Flash 51, l51r
  NFadeL 65, l65      ' Bonus 2X
  NFadeL 66, l66      ' Bonus 3X
  NFadeL 67, l67      ' Mode: Underground
  NFadeL 68, l68      ' Mode: Big Bang
  NFadeL 69, l69      ' Mode: Bar Brawl
  NFadeL 70, l70      ' Mode: Ball Busters
  NFadeL 71, l71      ' Mode: Looped
  NFadeL 72, l72      ' Shoot Again
  NFadeL 73, l73      ' Mode: Babescanner
  NFadeL 74, l74      ' Mode: Waitress
  NFadeL 75, l75      ' Shoot: Cosmic Dartz
  NFadeLm 76, l76     ' Special (Outlane R)
  Flash 76, l76r
  NFadeL 77, l77      ' Inlane Right
  NFadeL 78, l78      ' Bonus 5X
  NFadeL 79, l79      ' Bonus 4X
  NFadeL 80, l80      ' Mode: Tube Dancer
  NFadeL 81, l81      ' Shoot: Left Orbit
  NFadeL 82, l82      ' Shoot: Babescanner
  NFadeL 83, l83      ' 4-Bank Mars
  NFadeL 84, l84      ' 4-Bank Pythos
  NFadeL 85, l85      ' 4-Bank Venus
  NFadeL 86, l86      ' 4-Bank Mercury
  NFadeLm 87, l87     ' Free Shot (Outlane L)
  Flash 87, l87r
  NFadeL 88, l88      ' Inlane Left
  NFadeL 89, l89      ' Mode: Cosmic Dartz
  NFadeL 90, l90      ' Mode: Tour de Bar
  NFadeL 91, l91      ' Mode: Mosh
  NFadeL 92, l92      ' Mode: Happy Hour
  NFadeL 93, l93      ' Mode: Extra Ball
  NFadeL 94, l94      ' Mode: Get Lucky
  NFadeL 95, l95      ' Mode: Loonapalooza
  NFadeL 97, l97      ' Ramp Jackpot
  NFadeL 98, l98      ' Ramp Standup Left
  NFadeL 99, l99      ' Ramp Standup Right
  NFadeL 100, l100    ' Ramp Standup Side
  NFadeL 101, l101    ' Double Jackpot
  NFadeL 102, l102    ' Shoot: Tour de Bar
  NFadeL 103, l103    ' Shoot: Underground 1
  NFadeL 105, l105    ' Captive Left 4
  NFadeL 106, l106    ' Captive Left 3
  NFadeL 107, l107    ' Captive Left 2
  NFadeL 108, l108    ' Captive Left 1
  NFadeL 109, l109    ' Captive Right 4
  NFadeL 110, l110    ' Captive Right 3
  NFadeL 111, l111    ' Captive Right 2
  NFadeL 112, l112    ' Captive Right 1
  NFadeL 113, l113    ' 3-Bank Uranus
  NFadeL 114, l114    ' 3-Bank Neptune
  NFadeL 115, l115    ' 3-Bank Pluto
  NFadeL 116, l116    ' Shoot: Right Orbit
  NFadeL 117, l117    ' DJ Eyes
  NFadeL 118, l118    ' Shoot: Lunapalooza
  NFadeL 121, l121    ' Shoot: Underground 2
  NFadeLm 122, l122   ' Star Bumper Left
  NFadeL 122, l122a
  NFadeLm 123, l123   ' Star Bumper Medium
  NFadeL 123, l123a
  NFadeLm 124, l124   ' Star Bumper Right
  NFadeL 124, l124a
  NFadeL 125, l125    ' Dance Floor
  NFadeL 126, l126    ' Shoot: Extra Ball
  NFadeL 127, l127    ' Shoot: Big Bang

' Ramp Arrows
  NFadeLm 57, L57     ' Low
  FadeObj 57, l57p, "l_on", "l_b", "l_a", "l_off"
  NFadeLm 58, L58     ' Mid
  FadeObj 58, l58p, "l_on", "l_b", "l_a", "l_off"
  NFadeLm 59, L59     ' Top
  FadeObj 59, l59p, "l_on", "l_b", "l_a", "l_off"

' Neon Light
  FadeObjm 62, NeonLamp, "BlackLOn", "BlackLb", "BlackLa", "BlackL"
  Flashm 62, l62
  Flashm 62, l62a
  Flash 62, l62b

'Bulbs
  Flash 44, l44     ' Alien Lock Left Bulb
  Flash 45, l45     ' Alien Lock Right Bulb
  Flash 52, l52     ' Tube Sign X-Ball Bulb
  Flash 53, l53     ' Tube Sign 10 Million Bulb
  Flash 54, l54     ' Tube Sign Jackpot Bulb
  NFadeL 104, l104    ' Qualify Mode Bulb
  NFadeL 119, l119    ' Island: Lock Ready Bulb
  NFadeL 120, l120    ' Island: Mode Ready Bulb

'Flashers
  NFadeLm 165, f25      'Alien Lock Flasher
  Flash 165, f25a

 End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
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

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub


'//////////////////////////////////////////////////////////////////////
'// TIMERS
'//////////////////////////////////////////////////////////////////////
Set MotorCallback = GetRef("GameTimer")

' The game timer interval is 10 ms
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking
  RollingUpdate         'update rolling sounds
  DoDTAnim
End Sub
Sub GameTimer
    UpdateMechs
End Sub

Sub UpdateMechs
GateSWsx.RotZ= -GateL.currentangle
GateSWdx.RotZ= -GateR.currentangle
RampEntryGate.rotx=-Spinner1.currentangle + 90
RampExitGate.rotz=-Spinner2.currentangle
DiverterLeft.rotz=DivTubef.currentangle
DiverterRight.rotz=DivTube2f.currentangle
LeftFlipperP.roty=LeftFlipper.currentangle
RightFlipperP.roty=RightFlipper.currentangle
RightUFlipperP.roty=RightFlipper1.currentangle
BallShadowUpdate
End Sub


' The frame timer interval is -1, so executes at the display frame rate
Sub FrameTimer_Timer()
FlipperVisualUpdate
If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub

' *********************************************************************
'                      Extra Sounds
' *********************************************************************

Sub ShooterEnd_Hit():plungerbulb.state=0:PlaysoundAt "Ball_Bounce_Playfield_Soft_1", ShooterEnd : End Sub

'//////////////////////////////////////////////////////////////////////
'// TargetBounce
'//////////////////////////////////////////////////////////////////////

sub TargetBouncer(aBall,defvalue)
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled = 1 and aball.z < 30 then
        'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        vel = BallSpeed(aBall)
        if aBall.velx = 0 then vratio = 1 else vratio = aBall.vely/aBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1: zMultiplier = 0.2*defvalue
      Case 2: zMultiplier = 0.25*defvalue
            Case 3: zMultiplier = 0.3*defvalue
      Case 4: zMultiplier = 0.4*defvalue
            Case 5: zMultiplier = 0.45*defvalue
            Case 6: zMultiplier = 0.5*defvalue
        End Select
        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
        aBall.vely = aBall.velx * vratio
        'debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        'debug.print "conservation check: " & BallSpeed(aBall)/vel
  end if
end sub

' Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit
  TargetBouncer activeball, 1
End Sub

'//////////////////////////////////////////////////////////////////////
'// PHYSICS DAMPENERS
'//////////////////////////////////////////////////////////////////////

' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR



Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
  TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
  TargetBouncer Activeball, 0.7
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'#########################    Adjust these values to increase or lessen the elasticity

dim FlippersD : Set FlippersD = new Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold         'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  public sub Dampenf(aBall, parm) 'Rubberizer is handle here
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
' Thalamus - patched :       aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()         'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub

End Class

'//////////////////////////////////////////////////////////////////////
'// TRACK ALL BALL VELOCITIES FOR RUBBER DAMPENER AND DROP TARGETS
'//////////////////////////////////////////////////////////////////////

dim cor : set cor = New CoRTracker

Class CoRTracker
  public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs

    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds
    if uBound(ballvelx) < highestID then redim ballvelx(highestID)  'set bounds
    if uBound(ballvely) < highestID then redim ballvely(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

'//////////////////////////////////////////////////////////////////////
'// RAMP ROLLING SFX
'//////////////////////////////////////////////////////////////////////

'Ball tracking ramp SFX 1.0
'   Reqirements:
'          * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'          * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'          * Create a Timer called RampRoll, that is enabled, with a interval of 100
'          * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'          * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'          * To stop tracking ball
'                 * call WireRampOff
'                 * Otherwise, the ball will auto remove if it's below 30 vp units
'

dim RampMinLoops : RampMinLoops = 4

' RampBalls
'      Setup:        Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RammBalls(6,2)
'      Description:
dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(6)

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID : End Sub


' WaddBall (Active Ball, Boolean)
'     Description: This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
Sub Waddball(input, RampInput)  'Add ball
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  dim x : for x = 1 to uBound(RampBalls)  'Check, don't add balls twice
    if RampBalls(x, 1) = input.id then
      if Not IsEmpty(RampBalls(x,1) ) then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 to uBound(RampBalls)
    if IsEmpty(RampBalls(x, 1)) then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      exit Sub
    End If
    if x = uBound(RampBalls) then   'debug
      Debug.print "WireRampOn error, ball queue is full: " & vbnewline & _
      RampBalls(0, 0) & vbnewline & _
      Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbnewline & _
      Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbnewline & _
      Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbnewline & _
      Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbnewline & _
      Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbnewline & _
      " "
    End If
  next
End Sub

' WRemoveBall (BallId)
'    Description: This subroutine is called from the RampRollUpdate subroutine
'                 and is used to remove and stop the ball rolling sounds
Sub WRemoveBall(ID)   'Remove ball
  'Debug.Print "In WRemoveBall() + Remove ball from loop array"
  dim ballcount : ballcount = 0
  dim x : for x = 1 to Ubound(RampBalls)
    if ID = RampBalls(x, 1) then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    if not IsEmpty(Rampballs(x,1)) then ballcount = ballcount + 1
  next
  if BallCount = 0 then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer():RampRollUpdate:End Sub

Sub RampRollUpdate()    'Timer update
  dim x : for x = 1 to uBound(RampBalls)
    if Not IsEmpty(RampBalls(x,1) ) then
      if BallVel(RampBalls(x,0) ) > 1 then ' if ball is moving, play rolling sound
        If RampType(x) then
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end if
  next
  if not RampBalls(0,0) then RampRoll.enabled = 0

End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbnewline & _
  "1 " & Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbnewline & _
  "2 " & Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbnewline & _
  "3 " & Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbnewline & _
  "4 " & Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbnewline & _
  "5 " & Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbnewline & _
  "6 " & Typename(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbnewline & _
  " "
End Sub


Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(ball), 1, -1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function


'//////////////////////////////////////////////////////////////////////
'// RAMP TRIGGERS
'//////////////////////////////////////////////////////////////////////

Sub ramptrigger01_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger001_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger02_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger02_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger0002_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger0002_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger00002_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger03_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger03_unhit()
  PlaySoundAt "WireRamp_Stop", ramptrigger03
End Sub

Sub ramptrigger003_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger003_unhit()
  PlaySoundAt "WireRamp_Stop", ramptrigger003
End Sub

Sub ramptrigger002_unhit()
  PlaySoundAt "WireRamp_Stop", ramptrigger002
End Sub

'Fixes

Sub swthelph_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub sw64h_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub sw47h_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub sw48h_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub sw61Out_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub sw62Out_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

'//////////////////////////////////////////////////////////////////////
'// Ball Rolling
'//////////////////////////////////////////////////////////////////////

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 to tnob
    rolling(i) = False
  Next
End Sub

Sub RollingUpdate()
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
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

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If BOT(b).Z > 30 Then
        BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
      Else
        BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
      End If
      BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
      BallShadowA(b).X = BOT(b).X
      BallShadowA(b).visible = 1
    End If
  Next
End Sub

'//////////////////////////////////////////////////////////////////////
'// Mechanic Sounds
'//////////////////////////////////////////////////////////////////////

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'   Metals (all metal objects, metal walls, metal posts, metal wire guides)
'   Apron (the apron walls and plunger wall)
'   Walls (all wood or plastic walls)
'   Rollovers (wire rollover triggers, star triggers, or button triggers)
'   Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'   Gates (plate gates)
'   GatesWire (wire gates)
'   Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
' - Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
' - Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
' - Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
' - Ball rolling sounds sub
'
' Tutorial vides by Apophis
' Part 1:   https://youtu.be/PbE2kNiam3g
' Part 2:   https://youtu.be/B5cm1Y8wQsk
' Part 3:   https://youtu.be/eLhWyuYOyGg


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                         'volume level; range [0, 1]
NudgeRightSoundLevel = 1                        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                       'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                       'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                     'volume level; range [0, 1]
PlungerPullSoundLevel = 1                       'volume level; range [0, 1]
RollingSoundFactor = 1.1/5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                      'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                  'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel               'sound helper; not configurable
SlingshotSoundLevel = 0.95                        'volume level; range [0, 1]
BumperSoundFactor = 4.25                        'volume multiplier; must not be zero
KnockerSoundLevel = 1                           'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                  'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                     'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                     'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                    'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                 'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                 'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                     'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                          'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                     'volume multiplier; must not be zero
DTSoundLevel = 0.25                           'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                       'volume level; range [0, 1]
SpinnerSoundLevel = 0.5                                       'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                           'volume level; range [0, 1]
BallReleaseSoundLevel = 1                       'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                   'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                         'volume multiplier; must not be zero


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


'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / tableheight-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

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
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
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
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////
Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
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
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
  RandomSoundWall()
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
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
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
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Activeball
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  Dim snd
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
End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd*6)+1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315                  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05                  'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025*RelayGISoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025*RelayGISoundLevel, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025*RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025*RelayFlashSoundLevel, obj
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

' *********************************************************************
'                      Ramp Helpers
' *********************************************************************

sub TubeHelper_hit(): If ActiveBall.VelY < -1 Then activeball.vely= -1 : end if:end sub

Sub LRHelpers_Hit(idx)
  If ActiveBall.VelY < 0 AND ActiveBall.VelY  > -18 Then
    ActiveBall.VelY=ActiveBall.VelY/1.08
    ActiveBall.VelX=ActiveBall.VelX/1.08
  End If
  If ActiveBall.VelY > 0 Then
    ActiveBall.VelY=ActiveBall.VelY/1.1
    ActiveBall.VelX=ActiveBall.VelX/1.1
  End If
End Sub

'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************
' To add these slingshot corrections:
'   - On the table, add the endpoint primitives that define the two ends of the Slingshot
' - Initialize the SlingshotCorrection objects in InitSlingCorrection
'   - Call the .VelocityCorrect methods from the respective _Slingshot event sub


dim LS : Set LS = New SlingshotCorrection
dim RS : Set RS = New SlingshotCorrection

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
  AddSlingsPt 0, 0.00,  -4
  AddSlingsPt 1, 0.45,  -7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4

End Sub


Sub AddSlingsPt(idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LS, RS)
  dim x : for each x in a
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
Function RotPoint(x,y,angle)
    dim rx, ry
    rx = x*dCos(angle) - y*dSin(angle)
    ry = x*dSin(angle) + y*dCos(angle)
    RotPoint = Array(rx,ry)
End Function

Class SlingshotCorrection
  Public DebugOn, Enabled
  private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): Enabled = True : End Sub

  Public Property let Object(aInput) : Set Slingshot = aInput : End Property
  Public Property Let EndPoint1(aInput) : SlingX1 = aInput.x: SlingY1 = aInput.y: End Property
  Public Property Let EndPoint2(aInput) : SlingX2 = aInput.x: SlingY2 = aInput.y: End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    If gametime > 100 then Report
  End Sub

  Public Sub Report()         'debug, reports all coords in tbPL.text
    If not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1 : YL = SlingY1 : XR = SlingX2 : YR = SlingY2
    Else
      XL = SlingX2 : YL = SlingY2 : XR = SlingX1 : YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If ABS(XR-XL) > ABS(YR-YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If not IsEmpty(ModIn(0) ) then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      'debug.print " BallPos=" & BallPos &" Angle=" & Angle
      'debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled then aBall.Velx = RotVxVy(0)
      If Enabled then aBall.Vely = RotVxVy(1)
      'debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      'debug.print " "
    End If
  End Sub

End Class

'******************************************************
'****  DROP TARGETS by Rothbauerw
'******************************************************
' This solution improves the physics for drop targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full drop target animation, including deflection on hit and a slight lift when the drop
' targets raise, switch handling, bricking, and popping the ball up if it's over the drop target when it raises.
'
'Add a Timer named DTAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
'DTAnim.interval = 10
'DTAnim.enabled = True

'Sub DTAnim_Timer
' DoDTAnim
' DoSTAnim
'End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.


'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

'Define a variable for each drop target
Dim DT58,DT51,DT50,DT49,DT20,DT19,DT18,DT17

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:      primary target wall to determine drop
' secondary:      wall used to simulate the ball striking a bent or offset target after the initial Hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             rotz must be used for orientation
'             rotx to bend the target back
'             transz to move it up and down
'             the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
' switch:       ROM switch number
' animate:      Arrary slot for handling the animation instrucitons, set to 0
'
' Values for annimate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target


DT58 = Array(sw58, sw58a, sw58p, 58, 0)
DT51 = Array(sw51, sw51a, sw51p, 51, 0)
DT50 = Array(sw50, sw50a, sw50p, 50, 0)
DT49 = Array(sw49, sw49a, sw49p, 49, 0)
DT20 = Array(sw20, sw20a, sw20p, 20, 0)
DT19 = Array(sw19, sw19a, sw19p, 19, 0)
DT18 = Array(sw18, sw18a, sw18p, 18, 0)
DT17 = Array(sw17, sw17a, sw17p, 17, 0)

Dim DTArray
DTArray = Array(DT58,DT51,DT50,DT49,DT20,DT19,DT18,DT17)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 50 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
  Dim i
  i = DTArrayID(switch)

  PlayTargetSound
  DTArray(i)(4) =  DTCheckBrick(Activeball,DTArray(i)(2))
  If DTArray(i)(4) = 1 or DTArray(i)(4) = 3 or DTArray(i)(4) = 4 Then
    DTBallPhysics Activeball, DTArray(i)(2).rotz, DTMass
  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i)(4) = -1
  DoDTAnim
End Sub

Sub DTDrop(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i)(4) = 1
  DoDTAnim
End Sub

Function DTArrayID(switch)
  Dim i
  For i = 0 to uBound(DTArray)
    If DTArray(i)(3) = switch Then DTArrayID = i:Exit Function
  Next
End Function


sub DTBallPhysics(aBall, angle, mass)
  dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
  calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

  aBall.velx = calc1 * cos(rangle) + calc2
  aBall.vely = calc1 * sin(rangle) + calc3
End Sub


'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - tan(bangle) * aball.x + tan(rangle2) * dtprim.x) / (tan(rangle2) - tan(bangle))
  Yintersect = tan(rangle2) * Xintersect + (dtprim.y - tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

  If perpvel > 0 and  perpvelafter <= 0 Then
    If DTEnableBrick = 1 and  perpvel > DTBrickVel and DTBrickVel <> 0 and cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
    DTCheckBrick = 4
  Else
    DTCheckBrick = 0
  End If
End Function


Sub DoDTAnim()
  Dim i
  For i=0 to Ubound(DTArray)
    DTArray(i)(4) = DTAnimate(DTArray(i)(0),DTArray(i)(1),DTArray(i)(2),DTArray(i)(3),DTArray(i)(4))
  Next
End Sub

Function DTAnimate(primary, secondary, prim, switch,  animate)
  dim transz, switchid
  Dim animtime, rangle

  switchid = switch

  rangle = prim.rotz * PI / 180

  DTAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If (animate = 1 or animate = 4) and animtime < DTDropDelay Then
    primary.collidable = 0
  If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    DTAnimate = animate
    Exit Function
    elseif (animate = 1 or animate = 4) and animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
  End If

  if animate = 2 Then
    transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
    if prim.transz > -DTDropUnits  Then
      prim.transz = transz
    end if

    prim.rotx = DTMaxBend * cos(rangle)/2
    prim.roty = DTMaxBend * sin(rangle)/2

    if prim.transz <= -DTDropUnits Then
      prim.transz = -DTDropUnits
      secondary.collidable = 0
      controller.Switch(Switchid) = 1
      primary.uservalue = 0
      DTAnimate = 0
      Exit Function
    Else
      DTAnimate = 2
      Exit Function
    end If
  End If

  If animate = 3 and animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
  elseif animate = 3 and animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  if animate = -1 Then
    transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1

    If prim.transz = -DTDropUnits Then
      Dim BOT, b
      BOT = GetBalls

      For b = 0 to UBound(BOT)
        If InRotRect(BOT(b).x,BOT(b).y,prim.x, prim.y, prim.rotz, -25,-10,25,-10,25,25,-25,25) and BOT(b).z < prim.z+DTDropUnits+25 Then
          BOT(b).velz = 20
        End If
      Next
    End If

    if prim.transz < 0 Then
      prim.transz = transz
    elseif transz > 0 then
      prim.transz = transz
    end if

    if prim.transz > DTDropUpUnits then
      DTAnimate = -2
      prim.transz = DTDropUpUnits
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = gametime
    end if
    primary.collidable = 0
    secondary.collidable = 1
    controller.Switch(Switchid) = 0

  End If

  if animate = -2 and animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits
    if prim.transz < 0 then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    end If
  End If
End Function


'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS
'******************************************************


' Used for drop targets
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

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
    Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
    Dim rotxy
    rotxy = RotPoint(ax,ay,angle)
    rax = rotxy(0)+px : ray = rotxy(1)+py
    rotxy = RotPoint(bx,by,angle)
    rbx = rotxy(0)+px : rby = rotxy(1)+py
    rotxy = RotPoint(cx,cy,angle)
    rcx = rotxy(0)+px : rcy = rotxy(1)+py
    rotxy = RotPoint(dx,dy,angle)
    rdx = rotxy(0)+px : rdy = rotxy(1)+py

    InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

'Function RotPoint(x,y,angle)
'    dim rx, ry
'    rx = x*dCos(angle) - y*dSin(angle)
'    ry = x*dSin(angle) + y*dCos(angle)
'    RotPoint = Array(rx,ry)
'End Function


'******************************************************
'****  END DROP TARGETS
'******************************************************

'******************************************************
'*****   FLUPPER DOMES
'******************************************************
' Based on FlupperDoms2.2

' What you need in your table to use these flashers:
' Open this table and your table both in VPX
' Export all the materials domebasemat, Flashermaterial0 - 20 and import them in your table
' Export all textures (images) starting with the name "dome" and "ronddome" and import them into your table with the same names
' Export all textures (images) starting with the name "flasherbloom" and import them into your table with the same names
' Copy a set of 4 objects flasherbase, flasherlit, flasherlight and flasherflash from layer 7 to your table
' If you duplicate the four objects for a new flasher dome, be sure that they all end with the same number (in the 0-20 range)
' Copy the flasherbloom flashers from layer 10 to your table. you will need to make one per flasher dome that you plan to make
' Select the correct flasherbloom texture for each flasherbloom flasher, per flasher dome
' Copy the script below

' Place your flasher base primitive where you want the flasher located on your Table
' Then run InitFlasher in the script with the number of your flasher objects and the color of the flasher.  This will align the flasher object, light object, and
' flasher lit primitive.  It will also assign the appropriate flasher bloom images to the flasher bloom object.
'
' Example: InitFlasher 1, "green"
'
' Color Options: "blue", "green", "red", "purple", "yellow", "white", and "orange"

' You can use the RotateFlasher call to align the Rotz/ObjRotz of the flasher primitives with "handles".  Don't set those values in the editor,
' call the RotateFlasher sub instead (this call will likely crash VP if it's call for the flasher primitives without "handles")
'
' Example: RotateFlasher 1, 180     'where 1 is the flasher number and 180 is the angle of Z rotation

' For flashing the flasher use in the script: "ObjLevel(1) = 1 : FlasherFlash1_Timer"
' This should also work for flashers with variable flash levels from the rom, just use ObjLevel(1) = xx from the rom (in the range 0-1)
'
' Notes (please read!!):
' - Setting TestFlashers = 1 (below in the ScriptsDirectory) will allow you to see how the flasher objects are aligned (need the targetflasher image imported to your table)
' - The rotation of the primitives with "handles" is done with a script command, not on the primitive itself (see RotateFlasher below)
' - Color of the objects are set in the script, not on the primitive itself
' - Screws are optional to copy and position manually
' - If your table is not named "Table1" then change the name below in the script
' - Every flasher uses its own material (Flashermaterialxx), do not use it for anything else
' - Lighting > Bloom Strength affects how the flashers look, do not set it too high
' - Change RotY and RotX of flasherbase only when having a flasher something other then parallel to the playfield
' - Leave RotX of the flasherflash object to -45; this makes sure that the flash effect is visible in FS and DT
' - If you want to resize a flasher, be sure to resize flasherbase, flasherlit and flasherflash with the same percentage
' - If you think that the flasher effects are too bright, change flasherlightintensity and/or flasherflareintensity below

' Some more notes for users of the v1 flashers and/or JP's fading lights routines:
' - Delete all textures/primitives/script/materials in your table from the v1 flashers and scripts before you start; they don't mix well with v2
' - Remove flupperflash(m) routines if you have them; they do not work with this new script
' - Do not try to mix this v2 script with the JP fading light routine (that is making it too complicated), just use the example script below

' example script for rom based tables (non modulated):

' SolCallback(25)="FlashRed"
'
' Sub FlashRed(flstate)
' If Flstate Then
'   Objlevel(1) = 1 : FlasherFlash1_Timer
' End If
' End Sub

' example script for rom based tables (modulated):

' SolModCallback(25)="FlashRed"
'
' Sub FlashRed(level)
' Objlevel(1) = level/255 : FlasherFlash1_Timer
' End Sub


 Sub Flash1(Enabled)
  If Enabled Then
    Objlevel(1) = 1 : FlasherFlash1_Timer
  End If
  Sound_Flash_Relay enabled, Flasherbase1
 End Sub

 Sub Flash2(Enabled)
  If Enabled Then
    Objlevel(2) = 1 : FlasherFlash2_Timer
  End If
  Sound_Flash_Relay enabled, Flasherbase2
 End Sub

 Sub Flash3(Enabled)
  If Enabled Then
    Objlevel(3) = 1 : FlasherFlash3_Timer
  End If
  Sound_Flash_Relay enabled, Flasherbase3
 End Sub

 Sub Flash4(Enabled)
  If Enabled Then
    Objlevel(4) = 1 : FlasherFlash4_Timer
  End If
  Sound_Flash_Relay enabled, Flasherbase4
 End Sub

 Sub Flash5(Enabled)
  If Enabled Then
    Objlevel(5) = 1 : FlasherFlash5_Timer
  End If
  Sound_Flash_Relay enabled, Flasherbase5
 End Sub


Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table      ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.1   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.3   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherBloomIntensity = 0.2   ' *** lower this, if the blooms are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"

InitFlasher 1, "red"
InitFlasher 2, "red"
InitFlasher 3, "yellow"
InitFlasher 4, "yellow"
InitFlasher 5, "yellow"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 1,17 : RotateFlasher 2,0 : RotateFlasher 3,90 : RotateFlasher 4,90


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

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objbase(nr).roty > 135 then
    objflasher(nr).y = objbase(nr).y + 50
    objflasher(nr).height = objbase(nr).z + 20
  Else
    objflasher(nr).y = objbase(nr).y + 20
    objflasher(nr).height = objbase(nr).z + 50
  End If
  objflasher(nr).x = objbase(nr).x

  'rothbauerw
  'Adjust the position of the light object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  objlight(nr).x = objbase(nr).x
  objlight(nr).y = objbase(nr).y
  objlight(nr).bulbhaloheight = objbase(nr).z -10

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  dim xthird, ythird
  xthird = tablewidth/3
  ythird = tableheight/3

  If objbase(nr).x >= xthird and objbase(nr).x <= xthird*2 then
    objbloom(nr).imageA = "flasherbloomCenter"
    objbloom(nr).imageB = "flasherbloomCenter"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird then
    objbloom(nr).imageA = "flasherbloomUpperLeft"
    objbloom(nr).imageB = "flasherbloomUpperLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird then
    objbloom(nr).imageA = "flasherbloomUpperRight"
    objbloom(nr).imageB = "flasherbloomUpperRight"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird*2 then
    objbloom(nr).imageA = "flasherbloomCenterLeft"
    objbloom(nr).imageB = "flasherbloomCenterLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird*2 then
    objbloom(nr).imageA = "flasherbloomCenterRight"
    objbloom(nr).imageB = "flasherbloomCenterRight"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird*3 then
    objbloom(nr).imageA = "flasherbloomLowerLeft"
    objbloom(nr).imageB = "flasherbloomLowerLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird*3 then
    objbloom(nr).imageA = "flasherbloomLowerRight"
    objbloom(nr).imageB = "flasherbloomLowerRight"
  end if

  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objbloom(nr).color = RGB(4,120,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4) : objbloom(nr).color = RGB(12,255,4)
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4) : objbloom(nr).color = RGB(255,32,4)
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255) : objbloom(nr).color = RGB(230,49,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50) : objbloom(nr).color = RGB(200,173,25)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59) : objbloom(nr).color = RGB(255,240,150)
    Case "orange" :  objlight(nr).color = RGB(255,70,0) : objflasher(nr).color = RGB(255,70,0) : objbloom(nr).color = RGB(255,70,0)
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
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0 : End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub
Sub FlasherFlash5_Timer() : FlashFlasher(5) : End Sub

'******************************************************
'******  END FLUPPER DOMES
'******************************************************

'/////////////////////////////////////////////////////////////////
'         Siderails
'/////////////////////////////////////////////////////////////////

If Siderails = 1 Then
sidewall2.visible = 1
sidewall1.visible = 1
Else
sidewall2.visible = 0
sidewall1.visible = 0
End if
