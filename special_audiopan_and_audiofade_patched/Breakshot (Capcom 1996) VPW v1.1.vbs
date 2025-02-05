'Breakshot (Capcom 1996)
'https://www.ipdb.org/machine.cgi?id=3784
' ______   _______  _______  _______  _       _______           _______ _________
'(  ___ \ (  ____ )(  ____ \(  ___  )| )   /\(  ____ \|\     /|(  ___  )\__   __/
'| (   ) )| (    )|| (    \/| (   ) || |  / /| (    \/| )   ( || (   ) |   ) (
'| (__/ / | (____)|| (__    | (___) || (_/ / | (_____ | (___) || |   | |   | |
'|  __ (  |     __)|  __)   |  ___  ||  _ (  (_____  )|  ___  || |   | |   | |
'| (  \ \ | (\ (   | (      | (   ) || ( \ \       ) || (   ) || |   | |   | |
'| )___) )| ) \ \__| (____/\| )   ( || |  \ \/\____) || )   ( || (___) |   | |
'|______/ |/   \__/(_______/|/     \||_)   \/\_______)|/     \|(_______)   )_(
'
'VPW Hustlers
'============
'Niwak - Graphics & Script Work
'Sixtoe - Table & Script Work
'Apophis - Script Work & Tweaks
'Tastywasps - VR Object Animations
'VPW Team - Testing

Option Explicit
Randomize
SetLocale 1033

'*******************************************
' Base setup & core script files
'*******************************************

Const BallSize = 50, BallMass = 1, tnob = 5, lob = 2
Const cGameName="bsv103", UseSolenoids=2, UseVPMModSol=2, UseLamps=1, UseGI=0, NTweens=23, SSolenoidOn="", SSolenoidOff="", SCoin=""
Dim DesktopMode: DesktopMode = Table1.ShowDT
Dim UseVPMDMD: If RenderingMode = 2 Then UseVPMDMD = True Else UseVPMDMD = DesktopMode

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You are missing the Script folder that comes along VPX distribution"
On Error Goto 0
LoadVPM "03060000", "CAPCOM.VBS", 3.61

'-------------
'cTween
'-------------

Dim Tween()
If IsEmpty(Eval("NTweens")) = False Then
  If NTweens > 0 Then
    Redim Tween(NTweens - 1)
    Dim x : For x = 0 to UBound(Tween) : Set Tween(x) = new cTween : Tween(x).Id = x : Next
  End If
End If

Public Sub UpdateAllTweens(elapsed) : Dim x : For x = 0 to UBound(Tween) : Tween(x).OnFrameTimer elapsed : Next : End Sub

Class cTween
  Public Id
  Private CallBackRef, CallBackParam1, LiveValue, ActionPos, NextPos, TweenSpeed, TweenLength
  Private ActionType(30), ActionLength(30), ActionParam1(30), ActionParam2(30), ActionParam3(30)

  Private Sub Class_Initialize()
    Set CallBackRef = Nothing
    LiveValue = 0.0
    Clear
  End Sub

  Public Property Get Value() : Value = LiveValue : End Property
  Public Property Let Callback(cbRef) : Set CallBackRef = cbRef : End Property
  Public Property Let CallbackParam(v) : If IsObject(v) Then : Set CallBackParam1 = v : Else : CallBackParam1 = v : End If : End Property

  Public Function Clear()
    Set Clear = Me
    NextAction
    ActionPos = 0
    ActionType(ActionPos) = 0
    NextPos = 0
  End Function

  Public Function TargetLength(v, length) ' v is target value, length is tween duration in ms
    Set TargetLength = Me
    If length <= 0 Then SetValue(v) : Exit Function
    ActionType(NextPos) = 1
    ActionParam1(NextPos) = length
    ActionParam2(NextPos) = v
    Set ActionParam3(NextPos) = Nothing
    NextPos = (NextPos + 1) Mod 30
    ActionType(NextPos) = 0
  End Function

  Public Function TargetSpeed(v, speed) ' v is target value, speed is VP unit per ms
    Set TargetSpeed = Me
    ActionType(NextPos) = 2
    ActionParam1(NextPos) = speed
    ActionParam2(NextPos) = v
    Set ActionParam3(NextPos) = Nothing
    NextPos = (NextPos + 1) Mod 30
    ActionType(NextPos) = 0
  End Function

  Public Function Delay(length)
    Set Delay = Me
    If length <= 0 Then Exit Function
    ActionType(NextPos) = 3
    ActionParam1(NextPos) = length
    NextPos = (NextPos + 1) Mod 30
    ActionType(NextPos) = 0
  End Function

  Public Function SetValue(v)
    Set SetValue = Me
    ActionType(NextPos) = 4
    ActionParam1(NextPos) = v
    NextPos = (NextPos + 1) Mod 30
    ActionType(NextPos) = 0
  End Function

  Public Function Run(subRef)
    Set Run = Me
    If subRef is Nothing Then Exit Function
    ActionType(NextPos) = 5
    Set ActionParam1(NextPos) = subRef
    NextPos = (NextPos + 1) Mod 30
    ActionType(NextPos) = 0
  End Function

  Public Function Run1(subRef, x)
    Set Run1 = Me
    If subRef is Nothing Then Exit Function
    ActionType(NextPos) = 6
    Set ActionParam1(NextPos) = subRef
    ActionParam2(NextPos) = x
    NextPos = (NextPos + 1) Mod 30
    ActionType(NextPos) = 0
  End Function

  Public Function Run2(subRef, x, y)
    Set Run2 = Me
    If subRef is Nothing Then Exit Function
    ActionType(NextPos) = 7
    Set ActionParam1(NextPos) = subRef
    ActionParam2(NextPos) = x
    ActionParam3(NextPos) = y
    NextPos = (NextPos + 1) Mod 30
    ActionType(NextPos) = 0
  End Function

  Public Function Done() ' No Op Sub to account for VBS not wanting parenthesis at end of call
    Set Done = Me
  End Function

  Public Sub OnFrameTimer(elapsed)
    While elapsed > 0
      Select Case ActionType(ActionPos)
      Case 0
        Exit Sub
      Case 1
        If ActionParam1(ActionPos) > elapsed Then
          If TweenSpeed > 1e10 Then TweenSpeed = (ActionParam2(ActionPos) - Value) / ActionParam1(ActionPos)
          ActionParam1(ActionPos) = ActionParam1(ActionPos) - elapsed
          LiveValue = LiveValue + elapsed * TweenSpeed
          OnValueChanged
          Exit Sub
        End If
        elapsed = elapsed - ActionParam1(ActionPos)
        LiveValue = ActionParam2(ActionPos)
        NextAction
        OnValueChanged
      Case 2
        If TweenLength > 1e10 Then TweenLength = (ActionParam2(ActionPos) - Value) / ActionParam1(ActionPos)
        If TweenLength > elapsed Then
          TweenLength = TweenLength - elapsed
          LiveValue = LiveValue + elapsed * ActionParam1(ActionPos)
          OnValueChanged
          Exit Sub
        End If
        elapsed = elapsed - TweenLength
        LiveValue = ActionParam2(ActionPos)
        NextAction
        OnValueChanged
      Case 3
        ActionParam1(ActionPos) = ActionParam1(ActionPos) - elapsed
        If ActionParam1(ActionPos) > 0 Then Exit Sub
        elapsed = -ActionParam1(ActionPos)
        NextAction
      Case 4
        LiveValue = ActionParam1(ActionPos)
        NextAction
        OnValueChanged
      Case 5
        Dim cb : Set cb = ActionParam1(ActionPos)
        NextAction
        cb Id, LiveValue
      Case 6
        Dim cb2, x : Set cb2 = ActionParam1(ActionPos) : x = ActionParam2(ActionPos)
        NextAction
        cb2 Id, LiveValue, x
      Case 7
        Dim cb3, x2, y2 : Set cb3 = ActionParam1(ActionPos) : x2 = ActionParam2(ActionPos) : y2 = ActionParam3(ActionPos)
        NextAction
        cb3 Id, LiveValue, x2, y2
      End Select
    Wend
  End Sub

  Private Sub NextAction
    TweenSpeed = 1e11
    TweenLength = 1e11
    ActionPos = (ActionPos + 1) Mod 30
  End Sub

  Private Sub OnValueChanged
    Dim cb : Set cb = CallBackRef
    If Not cb is Nothing Then
      If IsEmpty(CallBackParam1) Then
        cb Id, LiveValue
      Else
        cb Id, CallBackParam1, LiveValue
      End If
    End If
  End Sub
End Class

' Common Tween updaters used by callbacks
Sub PrimArrayTransYUpdate(id, prims, v) : Dim x: For Each x in prims: x.TransY = v: Next: End Sub
Sub PrimArrayTransZUpdate(id, prims, v) : Dim x: For Each x in prims: x.TransZ = v: Next: End Sub

'***********
' Solenoids
'***********
SolCallback(1)      = "SolTrough"     'Drain
SolCallback(2)      = "SolRelease"      'Ball Release
SolCallback(3)      = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
'4 Left Slingshot
'5 Right Slingshot
SolCallback(6)      = "Kickback"      'Left Outlane Kicker
SolCallback(7)      = "ResetDropsRight"   'Drop Targets Right
SolCallback(8)      = "RightSaucer"     'Right Saucer
SolCallback(sLRFlipper) = "SolLRFlipper"        '9 Left Flipper
SolCallback(sURFlipper) = "SolURFlipper"        '10 Right Flipper
SolCallback(sLLFlipper) = "SolLLFlipper"        '11 Upper Right Flippper
SolCallback(12)     = "PostUp"        'Center Post Up
SolCallback(13)     = "ResetDropsLeft"    'Drop Targets Left
SolCallback(14)     = "PostRelease"     'Center Post Release (also triggered while firing up)
SolCallback(15)     = "SolRightGate"    'Right Top Gate
SolCallback(16)     = "SolLeftGate"     'Left Top Gate
'17-24 Not Used
SolCallback(25)     = "CLeftPocket"     'Center Pocket Left
SolCallback(26)     = "CenterPocket"    'Center Pocket Center
'27 Not Used
SolModCallback(28)      = "FlashLight"      'Flasher Center Pocket
'29 Right Bumper
'30 Center Bumper
'31 Left Bumper
SolCallback(32)     = "CRightPocket"    'Center Pocket Right

Sub AdjustBulbTint(light, lightmaps)
  Dim p : p = light.GetInPlayIntensity / (light.Intensity * light.IntensityScale)
  Dim r : r = 1.0
  Dim g : g = p^0.15
  Dim b : b = (p - 0.04) / 0.96 : If b < 0 Then b = 0
  Dim i : i = .2126 * r^2.2 + .7152 * g^2.2 + .0722 * b^2.2
  Dim v : v = RGB(Int(255 * r), Int(255 * g), Int(255 * b))
  dim x : For Each x in lightmaps : x.color = v : x.opacity = 100.0 / i : Next
End Sub

Sub FlashLight(p) : s128a.State = p : s128b.State = p : AdjustBulbTint s128a, BL_Flashers : End Sub
Sub l25_Animate : AdjustBulbTint l25, BL_Inserts_l25 : End Sub
Sub l26_Animate : AdjustBulbTint l26, BL_Inserts_l26 : End Sub

'****************
' Lightmaps
'****************
' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BR1: BP_BR1=Array(BM_BR1, LM_GI_gi_BR1, LM_GI_gi10_BR1, LM_GI_gi11_BR1, LM_GI_gi12_BR1, LM_Inserts_l25_BR1, LM_Flashers_BR1)
Dim BP_BR2: BP_BR2=Array(BM_BR2, LM_GI_gi_BR2, LM_GI_gi09_BR2, LM_GI_gi10_BR2, LM_GI_gi13_BR2, LM_Inserts_l26_BR2, LM_Flashers_BR2)
Dim BP_BR3: BP_BR3=Array(BM_BR3, LM_GI_gi_BR3, LM_Inserts_l25_BR3, LM_Inserts_l26_BR3, LM_Flashers_BR3)
Dim BP_BS1: BP_BS1=Array(BM_BS1, LM_GI_gi_BS1, LM_GI_gi09_BS1, LM_GI_gi10_BS1, LM_GI_gi11_BS1, LM_GI_gi12_BS1, LM_Inserts_l25_BS1, LM_Inserts_l26_BS1, LM_Flashers_BS1)
Dim BP_BS2: BP_BS2=Array(BM_BS2, LM_GI_gi_BS2, LM_GI_gi09_BS2, LM_GI_gi10_BS2, LM_GI_gi11_BS2, LM_GI_gi12_BS2, LM_GI_gi13_BS2, LM_Inserts_l25_BS2, LM_Inserts_l26_BS2, LM_Flashers_BS2)
Dim BP_BS3: BP_BS3=Array(BM_BS3, LM_GI_gi_BS3, LM_GI_gi10_BS3, LM_GI_gi11_BS3, LM_GI_gi12_BS3, LM_Inserts_l25_BS3, LM_Inserts_l26_BS3, LM_Flashers_BS3)
Dim BP_Cap_Bot: BP_Cap_Bot=Array(BM_Cap_Bot, LM_GI_gi_Cap_Bot, LM_GI_gi09_Cap_Bot, LM_GI_gi11_Cap_Bot, LM_GI_gi12_Cap_Bot, LM_Inserts_l25_Cap_Bot, LM_Inserts_l26_Cap_Bot, LM_Flashers_Cap_Bot)
Dim BP_Cap_Top: BP_Cap_Top=Array(BM_Cap_Top, LM_GI_gi_Cap_Top, LM_GI_gi09_Cap_Top, LM_GI_gi10_Cap_Top, LM_GI_gi11_Cap_Top, LM_GI_gi12_Cap_Top, LM_Inserts_l26_Cap_Top, LM_Flashers_Cap_Top, LM_Inserts_l10_Cap_Top, LM_Inserts_l5_Cap_Top)
Dim BP_Diverter: BP_Diverter=Array(BM_Diverter, LM_GI_gi_Diverter, LM_Inserts_l25_Diverter, LM_Inserts_l26_Diverter)
Dim BP_Flap_L: BP_Flap_L=Array(BM_Flap_L, LM_GI_gi_Flap_L, LM_GI_gi11_Flap_L, LM_GI_gi12_Flap_L, LM_Inserts_l45_Flap_L)
Dim BP_Flap_R: BP_Flap_R=Array(BM_Flap_R, LM_GI_gi_Flap_R, LM_GI_gi09_Flap_R, LM_GI_gi10_Flap_R)
Dim BP_FlipperL: BP_FlipperL=Array(BM_FlipperL, LM_GI_gi01_FlipperL, LM_GI_gi02_FlipperL, LM_GI_gi04_FlipperL, LM_GI_gi05_FlipperL, LM_GI_gi06_FlipperL, LM_GI_gi07_FlipperL, LM_GI_gi08_FlipperL, LM_GI_gi_FlipperL)
Dim BP_FlipperLU: BP_FlipperLU=Array(BM_FlipperLU, LM_GI_gi01_FlipperLU, LM_GI_gi02_FlipperLU, LM_GI_gi05_FlipperLU, LM_GI_gi06_FlipperLU, LM_GI_gi07_FlipperLU, LM_GI_gi08_FlipperLU, LM_GI_gi_FlipperLU)
Dim BP_FlipperR: BP_FlipperR=Array(BM_FlipperR, LM_GI_gi01_FlipperR, LM_GI_gi02_FlipperR, LM_GI_gi03_FlipperR, LM_GI_gi04_FlipperR, LM_GI_gi05_FlipperR, LM_GI_gi06_FlipperR, LM_GI_gi08_FlipperR, LM_GI_gi_FlipperR)
Dim BP_FlipperR1: BP_FlipperR1=Array(BM_FlipperR1, LM_GI_gi_FlipperR1, LM_Inserts_l26_FlipperR1)
Dim BP_FlipperR1U: BP_FlipperR1U=Array(BM_FlipperR1U, LM_GI_gi_FlipperR1U)
Dim BP_FlipperRU: BP_FlipperRU=Array(BM_FlipperRU, LM_GI_gi01_FlipperRU, LM_GI_gi02_FlipperRU, LM_GI_gi03_FlipperRU, LM_GI_gi04_FlipperRU, LM_GI_gi05_FlipperRU, LM_GI_gi06_FlipperRU, LM_GI_gi_FlipperRU)
Dim BP_Gate_L_W: BP_Gate_L_W=Array(BM_Gate_L_W, LM_GI_gi_Gate_L_W)
Dim BP_Gate_R_W: BP_Gate_R_W=Array(BM_Gate_R_W)
Dim BP_LDT1: BP_LDT1=Array(BM_LDT1, LM_GI_gi08_LDT1, LM_GI_gi_LDT1)
Dim BP_LDT2: BP_LDT2=Array(BM_LDT2, LM_GI_gi_LDT2)
Dim BP_LDT3: BP_LDT3=Array(BM_LDT3, LM_GI_gi08_LDT3, LM_GI_gi_LDT3)
Dim BP_LEMK: BP_LEMK=Array(BM_LEMK, LM_GI_gi05_LEMK, LM_GI_gi06_LEMK, LM_GI_gi07_LEMK, LM_GI_gi08_LEMK)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GI_gi07_LSling1, LM_GI_gi08_LSling1, LM_GI_gi_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GI_gi07_LSling2, LM_GI_gi08_LSling2, LM_GI_gi_LSling2)
Dim BP_NewtBall: BP_NewtBall=Array(BM_NewtBall, LM_GI_gi_NewtBall, LM_GI_gi09_NewtBall, LM_GI_gi10_NewtBall, LM_GI_gi11_NewtBall, LM_GI_gi12_NewtBall, LM_Inserts_l25_NewtBall, LM_Inserts_l26_NewtBall, LM_Flashers_NewtBall, LM_Inserts_l12_NewtBall, LM_Inserts_l3_NewtBall, LM_Inserts_l4_NewtBall, LM_Inserts_l8_NewtBall)
Dim BP_NewtBall_1: BP_NewtBall_1=Array(BM_NewtBall_1, LM_GI_gi_NewtBall_1, LM_GI_gi09_NewtBall_1, LM_GI_gi10_NewtBall_1, LM_GI_gi11_NewtBall_1, LM_GI_gi12_NewtBall_1, LM_Inserts_l25_NewtBall_1, LM_Inserts_l26_NewtBall_1, LM_Flashers_NewtBall_1, LM_Inserts_l3_NewtBall_1, LM_Inserts_l4_NewtBall_1)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_GI_gi01_Parts, LM_GI_gi02_Parts, LM_GI_gi03_Parts, LM_GI_gi04_Parts, LM_GI_gi05_Parts, LM_GI_gi06_Parts, LM_GI_gi07_Parts, LM_GI_gi08_Parts, LM_GI_gi_Parts, LM_GI_gi09_Parts, LM_GI_gi10_Parts, LM_GI_gi11_Parts, LM_GI_gi12_Parts, LM_GI_gi13_Parts, LM_Inserts_l25_Parts, LM_Inserts_l26_Parts, LM_Flashers_Parts, LM_Inserts_l10_Parts, LM_Inserts_l11_Parts, LM_Inserts_l12_Parts, LM_Inserts_l13_Parts, LM_Inserts_l14_Parts, LM_Inserts_l15_Parts, LM_Inserts_l16_Parts, LM_Inserts_l17_Parts, LM_Inserts_l18_Parts, LM_Inserts_l19_Parts, LM_Inserts_l20_Parts, LM_Inserts_l21_Parts, LM_Inserts_l22_Parts, LM_Inserts_l23_Parts, LM_Inserts_l24_Parts, LM_Inserts_l27_Parts, LM_Inserts_l28_Parts, LM_Inserts_l29_Parts, LM_Inserts_l30_Parts, LM_Inserts_l31_Parts, LM_Inserts_l32_Parts, LM_Inserts_l33_Parts, LM_Inserts_l34_Parts, LM_Inserts_l35_Parts, LM_Inserts_l36_Parts, LM_Inserts_l37_Parts, LM_Inserts_l38_Parts, LM_Inserts_l39_Parts, LM_Inserts_l3_Parts, LM_Inserts_l40_Parts, _
  LM_Inserts_l41_Parts, LM_Inserts_l42_Parts, LM_Inserts_l43_Parts, LM_Inserts_l44_Parts, LM_Inserts_l45_Parts, LM_Inserts_l46_Parts, LM_Inserts_l47_Parts, LM_Inserts_l48_Parts, LM_Inserts_l49_Parts, LM_Inserts_l4_Parts, LM_Inserts_l50_Parts, LM_Inserts_l51_Parts, LM_Inserts_l52_Parts, LM_Inserts_l53_Parts, LM_Inserts_l54_Parts, LM_Inserts_l55_Parts, LM_Inserts_l56_Parts, LM_Inserts_l57_Parts, LM_Inserts_l58_Parts, LM_Inserts_l59_Parts, LM_Inserts_l5_Parts, LM_Inserts_l60_Parts, LM_Inserts_l61_Parts, LM_Inserts_l62_Parts, LM_Inserts_l63_Parts, LM_Inserts_l64_Parts, LM_Inserts_l6_Parts, LM_Inserts_l7_Parts, LM_Inserts_l8_Parts, LM_Inserts_l9_Parts)
Dim BP_Plastics: BP_Plastics=Array(BM_Plastics, LM_GI_gi01_Plastics, LM_GI_gi02_Plastics, LM_GI_gi03_Plastics, LM_GI_gi04_Plastics, LM_GI_gi05_Plastics, LM_GI_gi06_Plastics, LM_GI_gi07_Plastics, LM_GI_gi08_Plastics, LM_GI_gi_Plastics, LM_GI_gi09_Plastics, LM_GI_gi10_Plastics, LM_GI_gi11_Plastics, LM_GI_gi12_Plastics, LM_GI_gi13_Plastics, LM_Inserts_l25_Plastics, LM_Inserts_l26_Plastics, LM_Flashers_Plastics, LM_Inserts_l28_Plastics, LM_Inserts_l8_Plastics)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GI_gi01_Playfield, LM_GI_gi02_Playfield, LM_GI_gi03_Playfield, LM_GI_gi04_Playfield, LM_GI_gi05_Playfield, LM_GI_gi06_Playfield, LM_GI_gi07_Playfield, LM_GI_gi08_Playfield, LM_GI_gi_Playfield, LM_GI_gi09_Playfield, LM_GI_gi10_Playfield, LM_GI_gi11_Playfield, LM_GI_gi12_Playfield, LM_GI_gi13_Playfield, LM_Inserts_l25_Playfield, LM_Inserts_l26_Playfield, LM_Flashers_Playfield, LM_Inserts_l7_Playfield)
Dim BP_Post_R: BP_Post_R=Array(BM_Post_R, LM_GI_gi03_Post_R, LM_GI_gi04_Post_R, LM_GI_gi_Post_R)
Dim BP_RDT1: BP_RDT1=Array(BM_RDT1, LM_GI_gi_RDT1, LM_GI_gi13_RDT1, LM_Inserts_l26_RDT1, LM_Flashers_RDT1, LM_Inserts_l48_RDT1, LM_Inserts_l7_RDT1)
Dim BP_RDT2: BP_RDT2=Array(BM_RDT2, LM_GI_gi_RDT2, LM_GI_gi13_RDT2, LM_Inserts_l25_RDT2, LM_Inserts_l26_RDT2, LM_Flashers_RDT2, LM_Inserts_l47_RDT2, LM_Inserts_l7_RDT2)
Dim BP_RDT3: BP_RDT3=Array(BM_RDT3, LM_GI_gi_RDT3, LM_GI_gi13_RDT3, LM_Inserts_l25_RDT3, LM_Flashers_RDT3, LM_Inserts_l46_RDT3, LM_Inserts_l7_RDT3)
Dim BP_REMK: BP_REMK=Array(BM_REMK, LM_GI_gi01_REMK, LM_GI_gi02_REMK, LM_GI_gi03_REMK, LM_GI_gi04_REMK)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GI_gi02_RSling1, LM_GI_gi03_RSling1, LM_GI_gi04_RSling1, LM_GI_gi_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GI_gi03_RSling2, LM_GI_gi04_RSling2, LM_GI_gi_RSling2)
Dim BP_Rails: BP_Rails=Array(BM_Rails, LM_GI_gi02_Rails, LM_GI_gi03_Rails, LM_GI_gi04_Rails, LM_GI_gi07_Rails, LM_GI_gi08_Rails, LM_GI_gi_Rails, LM_GI_gi09_Rails, LM_Inserts_l25_Rails, LM_Inserts_l26_Rails, LM_Flashers_Rails)
Dim BP_Rails_DT: BP_Rails_DT=Array(BM_Rails_DT)
Dim BP_sw43: BP_sw43=Array(BM_sw43)
Dim BP_sw44: BP_sw44=Array(BM_sw44, LM_GI_gi08_sw44, LM_GI_gi_sw44)
Dim BP_sw45: BP_sw45=Array(BM_sw45, LM_GI_gi02_sw45, LM_GI_gi04_sw45, LM_GI_gi_sw45)
Dim BP_sw46: BP_sw46=Array(BM_sw46, LM_GI_gi06_sw46, LM_GI_gi07_sw46, LM_GI_gi08_sw46, LM_GI_gi_sw46)
Dim BP_sw47: BP_sw47=Array(BM_sw47, LM_GI_gi02_sw47, LM_GI_gi04_sw47, LM_GI_gi_sw47)
Dim BP_sw51: BP_sw51=Array(BM_sw51, LM_GI_gi13_sw51, LM_Inserts_l26_sw51, LM_Flashers_sw51)
Dim BP_sw55: BP_sw55=Array(BM_sw55, LM_GI_gi_sw55, LM_GI_gi12_sw55, LM_Flashers_sw55)
Dim BP_sw60: BP_sw60=Array(BM_sw60, LM_GI_gi_sw60, LM_GI_gi11_sw60, LM_GI_gi12_sw60)
Dim BP_sw61: BP_sw61=Array(BM_sw61, LM_GI_gi_sw61, LM_GI_gi10_sw61, LM_GI_gi11_sw61, LM_Flashers_sw61)
Dim BP_sw62: BP_sw62=Array(BM_sw62, LM_GI_gi_sw62, LM_GI_gi09_sw62, LM_GI_gi10_sw62, LM_Flashers_sw62)
Dim BP_sw63: BP_sw63=Array(BM_sw63, LM_GI_gi13_sw63)
Dim BP_sw76: BP_sw76=Array(BM_sw76, LM_GI_gi_sw76)
Dim BP_sw77: BP_sw77=Array(BM_sw77, LM_GI_gi_sw77)
Dim BP_sw78: BP_sw78=Array(BM_sw78, LM_GI_gi_sw78)
Dim BP_sw79: BP_sw79=Array(BM_sw79, LM_GI_gi_sw79, LM_Flashers_sw79)
' Arrays per lighting scenario
Dim BL_Flashers: BL_Flashers=Array(LM_Flashers_BR1, LM_Flashers_BR2, LM_Flashers_BR3, LM_Flashers_BS1, LM_Flashers_BS2, LM_Flashers_BS3, LM_Flashers_Cap_Bot, LM_Flashers_Cap_Top, LM_Flashers_NewtBall, LM_Flashers_NewtBall_1, LM_Flashers_Parts, LM_Flashers_Plastics, LM_Flashers_Playfield, LM_Flashers_RDT1, LM_Flashers_RDT2, LM_Flashers_RDT3, LM_Flashers_Rails, LM_Flashers_sw51, LM_Flashers_sw55, LM_Flashers_sw61, LM_Flashers_sw62, LM_Flashers_sw79)
Dim BL_GI_gi: BL_GI_gi=Array(LM_GI_gi_BR1, LM_GI_gi_BR2, LM_GI_gi_BR3, LM_GI_gi_BS1, LM_GI_gi_BS2, LM_GI_gi_BS3, LM_GI_gi_Cap_Bot, LM_GI_gi_Cap_Top, LM_GI_gi_Diverter, LM_GI_gi_Flap_L, LM_GI_gi_Flap_R, LM_GI_gi_FlipperL, LM_GI_gi_FlipperLU, LM_GI_gi_FlipperR, LM_GI_gi_FlipperR1, LM_GI_gi_FlipperR1U, LM_GI_gi_FlipperRU, LM_GI_gi_Gate_L_W, LM_GI_gi_LDT1, LM_GI_gi_LDT2, LM_GI_gi_LDT3, LM_GI_gi_LSling1, LM_GI_gi_LSling2, LM_GI_gi_NewtBall, LM_GI_gi_NewtBall_1, LM_GI_gi_Parts, LM_GI_gi_Plastics, LM_GI_gi_Playfield, LM_GI_gi_Post_R, LM_GI_gi_RDT1, LM_GI_gi_RDT2, LM_GI_gi_RDT3, LM_GI_gi_RSling1, LM_GI_gi_RSling2, LM_GI_gi_Rails, LM_GI_gi_sw44, LM_GI_gi_sw45, LM_GI_gi_sw46, LM_GI_gi_sw47, LM_GI_gi_sw55, LM_GI_gi_sw60, LM_GI_gi_sw61, LM_GI_gi_sw62, LM_GI_gi_sw76, LM_GI_gi_sw77, LM_GI_gi_sw78, LM_GI_gi_sw79)
Dim BL_GI_gi01: BL_GI_gi01=Array(LM_GI_gi01_FlipperL, LM_GI_gi01_FlipperLU, LM_GI_gi01_FlipperR, LM_GI_gi01_FlipperRU, LM_GI_gi01_Parts, LM_GI_gi01_Plastics, LM_GI_gi01_Playfield, LM_GI_gi01_REMK)
Dim BL_GI_gi02: BL_GI_gi02=Array(LM_GI_gi02_FlipperL, LM_GI_gi02_FlipperLU, LM_GI_gi02_FlipperR, LM_GI_gi02_FlipperRU, LM_GI_gi02_Parts, LM_GI_gi02_Plastics, LM_GI_gi02_Playfield, LM_GI_gi02_REMK, LM_GI_gi02_RSling1, LM_GI_gi02_Rails, LM_GI_gi02_sw45, LM_GI_gi02_sw47)
Dim BL_GI_gi03: BL_GI_gi03=Array(LM_GI_gi03_FlipperR, LM_GI_gi03_FlipperRU, LM_GI_gi03_Parts, LM_GI_gi03_Plastics, LM_GI_gi03_Playfield, LM_GI_gi03_Post_R, LM_GI_gi03_REMK, LM_GI_gi03_RSling1, LM_GI_gi03_RSling2, LM_GI_gi03_Rails)
Dim BL_GI_gi04: BL_GI_gi04=Array(LM_GI_gi04_FlipperL, LM_GI_gi04_FlipperR, LM_GI_gi04_FlipperRU, LM_GI_gi04_Parts, LM_GI_gi04_Plastics, LM_GI_gi04_Playfield, LM_GI_gi04_Post_R, LM_GI_gi04_REMK, LM_GI_gi04_RSling1, LM_GI_gi04_RSling2, LM_GI_gi04_Rails, LM_GI_gi04_sw45, LM_GI_gi04_sw47)
Dim BL_GI_gi05: BL_GI_gi05=Array(LM_GI_gi05_FlipperL, LM_GI_gi05_FlipperLU, LM_GI_gi05_FlipperR, LM_GI_gi05_FlipperRU, LM_GI_gi05_LEMK, LM_GI_gi05_Parts, LM_GI_gi05_Plastics, LM_GI_gi05_Playfield)
Dim BL_GI_gi06: BL_GI_gi06=Array(LM_GI_gi06_FlipperL, LM_GI_gi06_FlipperLU, LM_GI_gi06_FlipperR, LM_GI_gi06_FlipperRU, LM_GI_gi06_LEMK, LM_GI_gi06_Parts, LM_GI_gi06_Plastics, LM_GI_gi06_Playfield, LM_GI_gi06_sw46)
Dim BL_GI_gi07: BL_GI_gi07=Array(LM_GI_gi07_FlipperL, LM_GI_gi07_FlipperLU, LM_GI_gi07_LEMK, LM_GI_gi07_LSling1, LM_GI_gi07_LSling2, LM_GI_gi07_Parts, LM_GI_gi07_Plastics, LM_GI_gi07_Playfield, LM_GI_gi07_Rails, LM_GI_gi07_sw46)
Dim BL_GI_gi08: BL_GI_gi08=Array(LM_GI_gi08_FlipperL, LM_GI_gi08_FlipperLU, LM_GI_gi08_FlipperR, LM_GI_gi08_LDT1, LM_GI_gi08_LDT3, LM_GI_gi08_LEMK, LM_GI_gi08_LSling1, LM_GI_gi08_LSling2, LM_GI_gi08_Parts, LM_GI_gi08_Plastics, LM_GI_gi08_Playfield, LM_GI_gi08_Rails, LM_GI_gi08_sw44, LM_GI_gi08_sw46)
Dim BL_GI_gi09: BL_GI_gi09=Array(LM_GI_gi09_BR2, LM_GI_gi09_BS1, LM_GI_gi09_BS2, LM_GI_gi09_Cap_Bot, LM_GI_gi09_Cap_Top, LM_GI_gi09_Flap_R, LM_GI_gi09_NewtBall, LM_GI_gi09_NewtBall_1, LM_GI_gi09_Parts, LM_GI_gi09_Plastics, LM_GI_gi09_Playfield, LM_GI_gi09_Rails, LM_GI_gi09_sw62)
Dim BL_GI_gi10: BL_GI_gi10=Array(LM_GI_gi10_BR1, LM_GI_gi10_BR2, LM_GI_gi10_BS1, LM_GI_gi10_BS2, LM_GI_gi10_BS3, LM_GI_gi10_Cap_Top, LM_GI_gi10_Flap_R, LM_GI_gi10_NewtBall, LM_GI_gi10_NewtBall_1, LM_GI_gi10_Parts, LM_GI_gi10_Plastics, LM_GI_gi10_Playfield, LM_GI_gi10_sw61, LM_GI_gi10_sw62)
Dim BL_GI_gi11: BL_GI_gi11=Array(LM_GI_gi11_BR1, LM_GI_gi11_BS1, LM_GI_gi11_BS2, LM_GI_gi11_BS3, LM_GI_gi11_Cap_Bot, LM_GI_gi11_Cap_Top, LM_GI_gi11_Flap_L, LM_GI_gi11_NewtBall, LM_GI_gi11_NewtBall_1, LM_GI_gi11_Parts, LM_GI_gi11_Plastics, LM_GI_gi11_Playfield, LM_GI_gi11_sw60, LM_GI_gi11_sw61)
Dim BL_GI_gi12: BL_GI_gi12=Array(LM_GI_gi12_BR1, LM_GI_gi12_BS1, LM_GI_gi12_BS2, LM_GI_gi12_BS3, LM_GI_gi12_Cap_Bot, LM_GI_gi12_Cap_Top, LM_GI_gi12_Flap_L, LM_GI_gi12_NewtBall, LM_GI_gi12_NewtBall_1, LM_GI_gi12_Parts, LM_GI_gi12_Plastics, LM_GI_gi12_Playfield, LM_GI_gi12_sw55, LM_GI_gi12_sw60)
Dim BL_GI_gi13: BL_GI_gi13=Array(LM_GI_gi13_BR2, LM_GI_gi13_BS2, LM_GI_gi13_Parts, LM_GI_gi13_Plastics, LM_GI_gi13_Playfield, LM_GI_gi13_RDT1, LM_GI_gi13_RDT2, LM_GI_gi13_RDT3, LM_GI_gi13_sw51, LM_GI_gi13_sw63)
Dim BL_Inserts_l10: BL_Inserts_l10=Array(LM_Inserts_l10_Cap_Top, LM_Inserts_l10_Parts)
Dim BL_Inserts_l11: BL_Inserts_l11=Array(LM_Inserts_l11_Parts)
Dim BL_Inserts_l12: BL_Inserts_l12=Array(LM_Inserts_l12_NewtBall, LM_Inserts_l12_Parts)
Dim BL_Inserts_l13: BL_Inserts_l13=Array(LM_Inserts_l13_Parts)
Dim BL_Inserts_l14: BL_Inserts_l14=Array(LM_Inserts_l14_Parts)
Dim BL_Inserts_l15: BL_Inserts_l15=Array(LM_Inserts_l15_Parts)
Dim BL_Inserts_l16: BL_Inserts_l16=Array(LM_Inserts_l16_Parts)
Dim BL_Inserts_l17: BL_Inserts_l17=Array(LM_Inserts_l17_Parts)
Dim BL_Inserts_l18: BL_Inserts_l18=Array(LM_Inserts_l18_Parts)
Dim BL_Inserts_l19: BL_Inserts_l19=Array(LM_Inserts_l19_Parts)
Dim BL_Inserts_l20: BL_Inserts_l20=Array(LM_Inserts_l20_Parts)
Dim BL_Inserts_l21: BL_Inserts_l21=Array(LM_Inserts_l21_Parts)
Dim BL_Inserts_l22: BL_Inserts_l22=Array(LM_Inserts_l22_Parts)
Dim BL_Inserts_l23: BL_Inserts_l23=Array(LM_Inserts_l23_Parts)
Dim BL_Inserts_l24: BL_Inserts_l24=Array(LM_Inserts_l24_Parts)
Dim BL_Inserts_l25: BL_Inserts_l25=Array(LM_Inserts_l25_BR1, LM_Inserts_l25_BR3, LM_Inserts_l25_BS1, LM_Inserts_l25_BS2, LM_Inserts_l25_BS3, LM_Inserts_l25_Cap_Bot, LM_Inserts_l25_Diverter, LM_Inserts_l25_NewtBall, LM_Inserts_l25_NewtBall_1, LM_Inserts_l25_Parts, LM_Inserts_l25_Plastics, LM_Inserts_l25_Playfield, LM_Inserts_l25_RDT2, LM_Inserts_l25_RDT3, LM_Inserts_l25_Rails)
Dim BL_Inserts_l26: BL_Inserts_l26=Array(LM_Inserts_l26_BR2, LM_Inserts_l26_BR3, LM_Inserts_l26_BS1, LM_Inserts_l26_BS2, LM_Inserts_l26_BS3, LM_Inserts_l26_Cap_Bot, LM_Inserts_l26_Cap_Top, LM_Inserts_l26_Diverter, LM_Inserts_l26_FlipperR1, LM_Inserts_l26_NewtBall, LM_Inserts_l26_NewtBall_1, LM_Inserts_l26_Parts, LM_Inserts_l26_Plastics, LM_Inserts_l26_Playfield, LM_Inserts_l26_RDT1, LM_Inserts_l26_RDT2, LM_Inserts_l26_Rails, LM_Inserts_l26_sw51)
Dim BL_Inserts_l27: BL_Inserts_l27=Array(LM_Inserts_l27_Parts)
Dim BL_Inserts_l28: BL_Inserts_l28=Array(LM_Inserts_l28_Parts, LM_Inserts_l28_Plastics)
Dim BL_Inserts_l29: BL_Inserts_l29=Array(LM_Inserts_l29_Parts)
Dim BL_Inserts_l3: BL_Inserts_l3=Array(LM_Inserts_l3_NewtBall, LM_Inserts_l3_NewtBall_1, LM_Inserts_l3_Parts)
Dim BL_Inserts_l30: BL_Inserts_l30=Array(LM_Inserts_l30_Parts)
Dim BL_Inserts_l31: BL_Inserts_l31=Array(LM_Inserts_l31_Parts)
Dim BL_Inserts_l32: BL_Inserts_l32=Array(LM_Inserts_l32_Parts)
Dim BL_Inserts_l33: BL_Inserts_l33=Array(LM_Inserts_l33_Parts)
Dim BL_Inserts_l34: BL_Inserts_l34=Array(LM_Inserts_l34_Parts)
Dim BL_Inserts_l35: BL_Inserts_l35=Array(LM_Inserts_l35_Parts)
Dim BL_Inserts_l36: BL_Inserts_l36=Array(LM_Inserts_l36_Parts)
Dim BL_Inserts_l37: BL_Inserts_l37=Array(LM_Inserts_l37_Parts)
Dim BL_Inserts_l38: BL_Inserts_l38=Array(LM_Inserts_l38_Parts)
Dim BL_Inserts_l39: BL_Inserts_l39=Array(LM_Inserts_l39_Parts)
Dim BL_Inserts_l4: BL_Inserts_l4=Array(LM_Inserts_l4_NewtBall, LM_Inserts_l4_NewtBall_1, LM_Inserts_l4_Parts)
Dim BL_Inserts_l40: BL_Inserts_l40=Array(LM_Inserts_l40_Parts)
Dim BL_Inserts_l41: BL_Inserts_l41=Array(LM_Inserts_l41_Parts)
Dim BL_Inserts_l42: BL_Inserts_l42=Array(LM_Inserts_l42_Parts)
Dim BL_Inserts_l43: BL_Inserts_l43=Array(LM_Inserts_l43_Parts)
Dim BL_Inserts_l44: BL_Inserts_l44=Array(LM_Inserts_l44_Parts)
Dim BL_Inserts_l45: BL_Inserts_l45=Array(LM_Inserts_l45_Flap_L, LM_Inserts_l45_Parts)
Dim BL_Inserts_l46: BL_Inserts_l46=Array(LM_Inserts_l46_Parts, LM_Inserts_l46_RDT3)
Dim BL_Inserts_l47: BL_Inserts_l47=Array(LM_Inserts_l47_Parts, LM_Inserts_l47_RDT2)
Dim BL_Inserts_l48: BL_Inserts_l48=Array(LM_Inserts_l48_Parts, LM_Inserts_l48_RDT1)
Dim BL_Inserts_l49: BL_Inserts_l49=Array(LM_Inserts_l49_Parts)
Dim BL_Inserts_l5: BL_Inserts_l5=Array(LM_Inserts_l5_Cap_Top, LM_Inserts_l5_Parts)
Dim BL_Inserts_l50: BL_Inserts_l50=Array(LM_Inserts_l50_Parts)
Dim BL_Inserts_l51: BL_Inserts_l51=Array(LM_Inserts_l51_Parts)
Dim BL_Inserts_l52: BL_Inserts_l52=Array(LM_Inserts_l52_Parts)
Dim BL_Inserts_l53: BL_Inserts_l53=Array(LM_Inserts_l53_Parts)
Dim BL_Inserts_l54: BL_Inserts_l54=Array(LM_Inserts_l54_Parts)
Dim BL_Inserts_l55: BL_Inserts_l55=Array(LM_Inserts_l55_Parts)
Dim BL_Inserts_l56: BL_Inserts_l56=Array(LM_Inserts_l56_Parts)
Dim BL_Inserts_l57: BL_Inserts_l57=Array(LM_Inserts_l57_Parts)
Dim BL_Inserts_l58: BL_Inserts_l58=Array(LM_Inserts_l58_Parts)
Dim BL_Inserts_l59: BL_Inserts_l59=Array(LM_Inserts_l59_Parts)
Dim BL_Inserts_l6: BL_Inserts_l6=Array(LM_Inserts_l6_Parts)
Dim BL_Inserts_l60: BL_Inserts_l60=Array(LM_Inserts_l60_Parts)
Dim BL_Inserts_l61: BL_Inserts_l61=Array(LM_Inserts_l61_Parts)
Dim BL_Inserts_l62: BL_Inserts_l62=Array(LM_Inserts_l62_Parts)
Dim BL_Inserts_l63: BL_Inserts_l63=Array(LM_Inserts_l63_Parts)
Dim BL_Inserts_l64: BL_Inserts_l64=Array(LM_Inserts_l64_Parts)
Dim BL_Inserts_l7: BL_Inserts_l7=Array(LM_Inserts_l7_Parts, LM_Inserts_l7_Playfield, LM_Inserts_l7_RDT1, LM_Inserts_l7_RDT2, LM_Inserts_l7_RDT3)
Dim BL_Inserts_l8: BL_Inserts_l8=Array(LM_Inserts_l8_NewtBall, LM_Inserts_l8_Parts, LM_Inserts_l8_Plastics)
Dim BL_Inserts_l9: BL_Inserts_l9=Array(LM_Inserts_l9_Parts)
Dim BL_World: BL_World=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Cap_Bot, BM_Cap_Top, BM_Diverter, BM_Flap_L, BM_Flap_R, BM_FlipperL, BM_FlipperLU, BM_FlipperR, BM_FlipperR1, BM_FlipperR1U, BM_FlipperRU, BM_Gate_L_W, BM_Gate_R_W, BM_LDT1, BM_LDT2, BM_LDT3, BM_LEMK, BM_LSling1, BM_LSling2, BM_NewtBall, BM_NewtBall_1, BM_Parts, BM_Plastics, BM_Playfield, BM_Post_R, BM_RDT1, BM_RDT2, BM_RDT3, BM_REMK, BM_RSling1, BM_RSling2, BM_Rails, BM_Rails_DT, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw47, BM_sw51, BM_sw55, BM_sw60, BM_sw61, BM_sw62, BM_sw63, BM_sw76, BM_sw77, BM_sw78, BM_sw79)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Cap_Bot, BM_Cap_Top, BM_Diverter, BM_Flap_L, BM_Flap_R, BM_FlipperL, BM_FlipperLU, BM_FlipperR, BM_FlipperR1, BM_FlipperR1U, BM_FlipperRU, BM_Gate_L_W, BM_Gate_R_W, BM_LDT1, BM_LDT2, BM_LDT3, BM_LEMK, BM_LSling1, BM_LSling2, BM_NewtBall, BM_NewtBall_1, BM_Parts, BM_Plastics, BM_Playfield, BM_Post_R, BM_RDT1, BM_RDT2, BM_RDT3, BM_REMK, BM_RSling1, BM_RSling2, BM_Rails, BM_Rails_DT, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw47, BM_sw51, BM_sw55, BM_sw60, BM_sw61, BM_sw62, BM_sw63, BM_sw76, BM_sw77, BM_sw78, BM_sw79)
Dim BG_Lightmap: BG_Lightmap=Array(LM_Flashers_BR1, LM_Flashers_BR2, LM_Flashers_BR3, LM_Flashers_BS1, LM_Flashers_BS2, LM_Flashers_BS3, LM_Flashers_Cap_Bot, LM_Flashers_Cap_Top, LM_Flashers_NewtBall, LM_Flashers_NewtBall_1, LM_Flashers_Parts, LM_Flashers_Plastics, LM_Flashers_Playfield, LM_Flashers_RDT1, LM_Flashers_RDT2, LM_Flashers_RDT3, LM_Flashers_Rails, LM_Flashers_sw51, LM_Flashers_sw55, LM_Flashers_sw61, LM_Flashers_sw62, LM_Flashers_sw79, LM_GI_gi_BR1, LM_GI_gi_BR2, LM_GI_gi_BR3, LM_GI_gi_BS1, LM_GI_gi_BS2, LM_GI_gi_BS3, LM_GI_gi_Cap_Bot, LM_GI_gi_Cap_Top, LM_GI_gi_Diverter, LM_GI_gi_Flap_L, LM_GI_gi_Flap_R, LM_GI_gi_FlipperL, LM_GI_gi_FlipperLU, LM_GI_gi_FlipperR, LM_GI_gi_FlipperR1, LM_GI_gi_FlipperR1U, LM_GI_gi_FlipperRU, LM_GI_gi_Gate_L_W, LM_GI_gi_LDT1, LM_GI_gi_LDT2, LM_GI_gi_LDT3, LM_GI_gi_LSling1, LM_GI_gi_LSling2, LM_GI_gi_NewtBall, LM_GI_gi_NewtBall_1, LM_GI_gi_Parts, LM_GI_gi_Plastics, LM_GI_gi_Playfield, LM_GI_gi_Post_R, LM_GI_gi_RDT1, LM_GI_gi_RDT2, LM_GI_gi_RDT3, LM_GI_gi_RSling1, _
  LM_GI_gi_RSling2, LM_GI_gi_Rails, LM_GI_gi_sw44, LM_GI_gi_sw45, LM_GI_gi_sw46, LM_GI_gi_sw47, LM_GI_gi_sw55, LM_GI_gi_sw60, LM_GI_gi_sw61, LM_GI_gi_sw62, LM_GI_gi_sw76, LM_GI_gi_sw77, LM_GI_gi_sw78, LM_GI_gi_sw79, LM_GI_gi01_FlipperL, LM_GI_gi01_FlipperLU, LM_GI_gi01_FlipperR, LM_GI_gi01_FlipperRU, LM_GI_gi01_Parts, LM_GI_gi01_Plastics, LM_GI_gi01_Playfield, LM_GI_gi01_REMK, LM_GI_gi02_FlipperL, LM_GI_gi02_FlipperLU, LM_GI_gi02_FlipperR, LM_GI_gi02_FlipperRU, LM_GI_gi02_Parts, LM_GI_gi02_Plastics, LM_GI_gi02_Playfield, LM_GI_gi02_REMK, LM_GI_gi02_RSling1, LM_GI_gi02_Rails, LM_GI_gi02_sw45, LM_GI_gi02_sw47, LM_GI_gi03_FlipperR, LM_GI_gi03_FlipperRU, LM_GI_gi03_Parts, LM_GI_gi03_Plastics, LM_GI_gi03_Playfield, LM_GI_gi03_Post_R, LM_GI_gi03_REMK, LM_GI_gi03_RSling1, LM_GI_gi03_RSling2, LM_GI_gi03_Rails, LM_GI_gi04_FlipperL, LM_GI_gi04_FlipperR, LM_GI_gi04_FlipperRU, LM_GI_gi04_Parts, LM_GI_gi04_Plastics, LM_GI_gi04_Playfield, LM_GI_gi04_Post_R, LM_GI_gi04_REMK, LM_GI_gi04_RSling1, LM_GI_gi04_RSling2, _
  LM_GI_gi04_Rails, LM_GI_gi04_sw45, LM_GI_gi04_sw47, LM_GI_gi05_FlipperL, LM_GI_gi05_FlipperLU, LM_GI_gi05_FlipperR, LM_GI_gi05_FlipperRU, LM_GI_gi05_LEMK, LM_GI_gi05_Parts, LM_GI_gi05_Plastics, LM_GI_gi05_Playfield, LM_GI_gi06_FlipperL, LM_GI_gi06_FlipperLU, LM_GI_gi06_FlipperR, LM_GI_gi06_FlipperRU, LM_GI_gi06_LEMK, LM_GI_gi06_Parts, LM_GI_gi06_Plastics, LM_GI_gi06_Playfield, LM_GI_gi06_sw46, LM_GI_gi07_FlipperL, LM_GI_gi07_FlipperLU, LM_GI_gi07_LEMK, LM_GI_gi07_LSling1, LM_GI_gi07_LSling2, LM_GI_gi07_Parts, LM_GI_gi07_Plastics, LM_GI_gi07_Playfield, LM_GI_gi07_Rails, LM_GI_gi07_sw46, LM_GI_gi08_FlipperL, LM_GI_gi08_FlipperLU, LM_GI_gi08_FlipperR, LM_GI_gi08_LDT1, LM_GI_gi08_LDT3, LM_GI_gi08_LEMK, LM_GI_gi08_LSling1, LM_GI_gi08_LSling2, LM_GI_gi08_Parts, LM_GI_gi08_Plastics, LM_GI_gi08_Playfield, LM_GI_gi08_Rails, LM_GI_gi08_sw44, LM_GI_gi08_sw46, LM_GI_gi09_BR2, LM_GI_gi09_BS1, LM_GI_gi09_BS2, LM_GI_gi09_Cap_Bot, LM_GI_gi09_Cap_Top, LM_GI_gi09_Flap_R, LM_GI_gi09_NewtBall, LM_GI_gi09_NewtBall_1, _
  LM_GI_gi09_Parts, LM_GI_gi09_Plastics, LM_GI_gi09_Playfield, LM_GI_gi09_Rails, LM_GI_gi09_sw62, LM_GI_gi10_BR1, LM_GI_gi10_BR2, LM_GI_gi10_BS1, LM_GI_gi10_BS2, LM_GI_gi10_BS3, LM_GI_gi10_Cap_Top, LM_GI_gi10_Flap_R, LM_GI_gi10_NewtBall, LM_GI_gi10_NewtBall_1, LM_GI_gi10_Parts, LM_GI_gi10_Plastics, LM_GI_gi10_Playfield, LM_GI_gi10_sw61, LM_GI_gi10_sw62, LM_GI_gi11_BR1, LM_GI_gi11_BS1, LM_GI_gi11_BS2, LM_GI_gi11_BS3, LM_GI_gi11_Cap_Bot, LM_GI_gi11_Cap_Top, LM_GI_gi11_Flap_L, LM_GI_gi11_NewtBall, LM_GI_gi11_NewtBall_1, LM_GI_gi11_Parts, LM_GI_gi11_Plastics, LM_GI_gi11_Playfield, LM_GI_gi11_sw60, LM_GI_gi11_sw61, LM_GI_gi12_BR1, LM_GI_gi12_BS1, LM_GI_gi12_BS2, LM_GI_gi12_BS3, LM_GI_gi12_Cap_Bot, LM_GI_gi12_Cap_Top, LM_GI_gi12_Flap_L, LM_GI_gi12_NewtBall, LM_GI_gi12_NewtBall_1, LM_GI_gi12_Parts, LM_GI_gi12_Plastics, LM_GI_gi12_Playfield, LM_GI_gi12_sw55, LM_GI_gi12_sw60, LM_GI_gi13_BR2, LM_GI_gi13_BS2, LM_GI_gi13_Parts, LM_GI_gi13_Plastics, LM_GI_gi13_Playfield, LM_GI_gi13_RDT1, LM_GI_gi13_RDT2, LM_GI_gi13_RDT3, _
  LM_GI_gi13_sw51, LM_GI_gi13_sw63, LM_Inserts_l10_Cap_Top, LM_Inserts_l10_Parts, LM_Inserts_l11_Parts, LM_Inserts_l12_NewtBall, LM_Inserts_l12_Parts, LM_Inserts_l13_Parts, LM_Inserts_l14_Parts, LM_Inserts_l15_Parts, LM_Inserts_l16_Parts, LM_Inserts_l17_Parts, LM_Inserts_l18_Parts, LM_Inserts_l19_Parts, LM_Inserts_l20_Parts, LM_Inserts_l21_Parts, LM_Inserts_l22_Parts, LM_Inserts_l23_Parts, LM_Inserts_l24_Parts, LM_Inserts_l25_BR1, LM_Inserts_l25_BR3, LM_Inserts_l25_BS1, LM_Inserts_l25_BS2, LM_Inserts_l25_BS3, LM_Inserts_l25_Cap_Bot, LM_Inserts_l25_Diverter, LM_Inserts_l25_NewtBall, LM_Inserts_l25_NewtBall_1, LM_Inserts_l25_Parts, LM_Inserts_l25_Plastics, LM_Inserts_l25_Playfield, LM_Inserts_l25_RDT2, LM_Inserts_l25_RDT3, LM_Inserts_l25_Rails, LM_Inserts_l26_BR2, LM_Inserts_l26_BR3, LM_Inserts_l26_BS1, LM_Inserts_l26_BS2, LM_Inserts_l26_BS3, LM_Inserts_l26_Cap_Bot, LM_Inserts_l26_Cap_Top, LM_Inserts_l26_Diverter, LM_Inserts_l26_FlipperR1, LM_Inserts_l26_NewtBall, LM_Inserts_l26_NewtBall_1, LM_Inserts_l26_Parts, _
  LM_Inserts_l26_Plastics, LM_Inserts_l26_Playfield, LM_Inserts_l26_RDT1, LM_Inserts_l26_RDT2, LM_Inserts_l26_Rails, LM_Inserts_l26_sw51, LM_Inserts_l27_Parts, LM_Inserts_l28_Parts, LM_Inserts_l28_Plastics, LM_Inserts_l29_Parts, LM_Inserts_l3_NewtBall, LM_Inserts_l3_NewtBall_1, LM_Inserts_l3_Parts, LM_Inserts_l30_Parts, LM_Inserts_l31_Parts, LM_Inserts_l32_Parts, LM_Inserts_l33_Parts, LM_Inserts_l34_Parts, LM_Inserts_l35_Parts, LM_Inserts_l36_Parts, LM_Inserts_l37_Parts, LM_Inserts_l38_Parts, LM_Inserts_l39_Parts, LM_Inserts_l4_NewtBall, LM_Inserts_l4_NewtBall_1, LM_Inserts_l4_Parts, LM_Inserts_l40_Parts, LM_Inserts_l41_Parts, LM_Inserts_l42_Parts, LM_Inserts_l43_Parts, LM_Inserts_l44_Parts, LM_Inserts_l45_Flap_L, LM_Inserts_l45_Parts, LM_Inserts_l46_Parts, LM_Inserts_l46_RDT3, LM_Inserts_l47_Parts, LM_Inserts_l47_RDT2, LM_Inserts_l48_Parts, LM_Inserts_l48_RDT1, LM_Inserts_l49_Parts, LM_Inserts_l5_Cap_Top, LM_Inserts_l5_Parts, LM_Inserts_l50_Parts, LM_Inserts_l51_Parts, LM_Inserts_l52_Parts, _
  LM_Inserts_l53_Parts, LM_Inserts_l54_Parts, LM_Inserts_l55_Parts, LM_Inserts_l56_Parts, LM_Inserts_l57_Parts, LM_Inserts_l58_Parts, LM_Inserts_l59_Parts, LM_Inserts_l6_Parts, LM_Inserts_l60_Parts, LM_Inserts_l61_Parts, LM_Inserts_l62_Parts, LM_Inserts_l63_Parts, LM_Inserts_l64_Parts, LM_Inserts_l7_Parts, LM_Inserts_l7_Playfield, LM_Inserts_l7_RDT1, LM_Inserts_l7_RDT2, LM_Inserts_l7_RDT3, LM_Inserts_l8_NewtBall, LM_Inserts_l8_Parts, LM_Inserts_l8_Plastics, LM_Inserts_l9_Parts)
Dim BG_All: BG_All=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Cap_Bot, BM_Cap_Top, BM_Diverter, BM_Flap_L, BM_Flap_R, BM_FlipperL, BM_FlipperLU, BM_FlipperR, BM_FlipperR1, BM_FlipperR1U, BM_FlipperRU, BM_Gate_L_W, BM_Gate_R_W, BM_LDT1, BM_LDT2, BM_LDT3, BM_LEMK, BM_LSling1, BM_LSling2, BM_NewtBall, BM_NewtBall_1, BM_Parts, BM_Plastics, BM_Playfield, BM_Post_R, BM_RDT1, BM_RDT2, BM_RDT3, BM_REMK, BM_RSling1, BM_RSling2, BM_Rails, BM_Rails_DT, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw47, BM_sw51, BM_sw55, BM_sw60, BM_sw61, BM_sw62, BM_sw63, BM_sw76, BM_sw77, BM_sw78, BM_sw79, LM_Flashers_BR1, LM_Flashers_BR2, LM_Flashers_BR3, LM_Flashers_BS1, LM_Flashers_BS2, LM_Flashers_BS3, LM_Flashers_Cap_Bot, LM_Flashers_Cap_Top, LM_Flashers_NewtBall, LM_Flashers_NewtBall_1, LM_Flashers_Parts, LM_Flashers_Plastics, LM_Flashers_Playfield, LM_Flashers_RDT1, LM_Flashers_RDT2, LM_Flashers_RDT3, LM_Flashers_Rails, LM_Flashers_sw51, LM_Flashers_sw55, LM_Flashers_sw61, LM_Flashers_sw62, LM_Flashers_sw79, _
  LM_GI_gi_BR1, LM_GI_gi_BR2, LM_GI_gi_BR3, LM_GI_gi_BS1, LM_GI_gi_BS2, LM_GI_gi_BS3, LM_GI_gi_Cap_Bot, LM_GI_gi_Cap_Top, LM_GI_gi_Diverter, LM_GI_gi_Flap_L, LM_GI_gi_Flap_R, LM_GI_gi_FlipperL, LM_GI_gi_FlipperLU, LM_GI_gi_FlipperR, LM_GI_gi_FlipperR1, LM_GI_gi_FlipperR1U, LM_GI_gi_FlipperRU, LM_GI_gi_Gate_L_W, LM_GI_gi_LDT1, LM_GI_gi_LDT2, LM_GI_gi_LDT3, LM_GI_gi_LSling1, LM_GI_gi_LSling2, LM_GI_gi_NewtBall, LM_GI_gi_NewtBall_1, LM_GI_gi_Parts, LM_GI_gi_Plastics, LM_GI_gi_Playfield, LM_GI_gi_Post_R, LM_GI_gi_RDT1, LM_GI_gi_RDT2, LM_GI_gi_RDT3, LM_GI_gi_RSling1, LM_GI_gi_RSling2, LM_GI_gi_Rails, LM_GI_gi_sw44, LM_GI_gi_sw45, LM_GI_gi_sw46, LM_GI_gi_sw47, LM_GI_gi_sw55, LM_GI_gi_sw60, LM_GI_gi_sw61, LM_GI_gi_sw62, LM_GI_gi_sw76, LM_GI_gi_sw77, LM_GI_gi_sw78, LM_GI_gi_sw79, LM_GI_gi01_FlipperL, LM_GI_gi01_FlipperLU, LM_GI_gi01_FlipperR, LM_GI_gi01_FlipperRU, LM_GI_gi01_Parts, LM_GI_gi01_Plastics, LM_GI_gi01_Playfield, LM_GI_gi01_REMK, LM_GI_gi02_FlipperL, LM_GI_gi02_FlipperLU, LM_GI_gi02_FlipperR, _
  LM_GI_gi02_FlipperRU, LM_GI_gi02_Parts, LM_GI_gi02_Plastics, LM_GI_gi02_Playfield, LM_GI_gi02_REMK, LM_GI_gi02_RSling1, LM_GI_gi02_Rails, LM_GI_gi02_sw45, LM_GI_gi02_sw47, LM_GI_gi03_FlipperR, LM_GI_gi03_FlipperRU, LM_GI_gi03_Parts, LM_GI_gi03_Plastics, LM_GI_gi03_Playfield, LM_GI_gi03_Post_R, LM_GI_gi03_REMK, LM_GI_gi03_RSling1, LM_GI_gi03_RSling2, LM_GI_gi03_Rails, LM_GI_gi04_FlipperL, LM_GI_gi04_FlipperR, LM_GI_gi04_FlipperRU, LM_GI_gi04_Parts, LM_GI_gi04_Plastics, LM_GI_gi04_Playfield, LM_GI_gi04_Post_R, LM_GI_gi04_REMK, LM_GI_gi04_RSling1, LM_GI_gi04_RSling2, LM_GI_gi04_Rails, LM_GI_gi04_sw45, LM_GI_gi04_sw47, LM_GI_gi05_FlipperL, LM_GI_gi05_FlipperLU, LM_GI_gi05_FlipperR, LM_GI_gi05_FlipperRU, LM_GI_gi05_LEMK, LM_GI_gi05_Parts, LM_GI_gi05_Plastics, LM_GI_gi05_Playfield, LM_GI_gi06_FlipperL, LM_GI_gi06_FlipperLU, LM_GI_gi06_FlipperR, LM_GI_gi06_FlipperRU, LM_GI_gi06_LEMK, LM_GI_gi06_Parts, LM_GI_gi06_Plastics, LM_GI_gi06_Playfield, LM_GI_gi06_sw46, LM_GI_gi07_FlipperL, LM_GI_gi07_FlipperLU, _
  LM_GI_gi07_LEMK, LM_GI_gi07_LSling1, LM_GI_gi07_LSling2, LM_GI_gi07_Parts, LM_GI_gi07_Plastics, LM_GI_gi07_Playfield, LM_GI_gi07_Rails, LM_GI_gi07_sw46, LM_GI_gi08_FlipperL, LM_GI_gi08_FlipperLU, LM_GI_gi08_FlipperR, LM_GI_gi08_LDT1, LM_GI_gi08_LDT3, LM_GI_gi08_LEMK, LM_GI_gi08_LSling1, LM_GI_gi08_LSling2, LM_GI_gi08_Parts, LM_GI_gi08_Plastics, LM_GI_gi08_Playfield, LM_GI_gi08_Rails, LM_GI_gi08_sw44, LM_GI_gi08_sw46, LM_GI_gi09_BR2, LM_GI_gi09_BS1, LM_GI_gi09_BS2, LM_GI_gi09_Cap_Bot, LM_GI_gi09_Cap_Top, LM_GI_gi09_Flap_R, LM_GI_gi09_NewtBall, LM_GI_gi09_NewtBall_1, LM_GI_gi09_Parts, LM_GI_gi09_Plastics, LM_GI_gi09_Playfield, LM_GI_gi09_Rails, LM_GI_gi09_sw62, LM_GI_gi10_BR1, LM_GI_gi10_BR2, LM_GI_gi10_BS1, LM_GI_gi10_BS2, LM_GI_gi10_BS3, LM_GI_gi10_Cap_Top, LM_GI_gi10_Flap_R, LM_GI_gi10_NewtBall, LM_GI_gi10_NewtBall_1, LM_GI_gi10_Parts, LM_GI_gi10_Plastics, LM_GI_gi10_Playfield, LM_GI_gi10_sw61, LM_GI_gi10_sw62, LM_GI_gi11_BR1, LM_GI_gi11_BS1, LM_GI_gi11_BS2, LM_GI_gi11_BS3, LM_GI_gi11_Cap_Bot, _
  LM_GI_gi11_Cap_Top, LM_GI_gi11_Flap_L, LM_GI_gi11_NewtBall, LM_GI_gi11_NewtBall_1, LM_GI_gi11_Parts, LM_GI_gi11_Plastics, LM_GI_gi11_Playfield, LM_GI_gi11_sw60, LM_GI_gi11_sw61, LM_GI_gi12_BR1, LM_GI_gi12_BS1, LM_GI_gi12_BS2, LM_GI_gi12_BS3, LM_GI_gi12_Cap_Bot, LM_GI_gi12_Cap_Top, LM_GI_gi12_Flap_L, LM_GI_gi12_NewtBall, LM_GI_gi12_NewtBall_1, LM_GI_gi12_Parts, LM_GI_gi12_Plastics, LM_GI_gi12_Playfield, LM_GI_gi12_sw55, LM_GI_gi12_sw60, LM_GI_gi13_BR2, LM_GI_gi13_BS2, LM_GI_gi13_Parts, LM_GI_gi13_Plastics, LM_GI_gi13_Playfield, LM_GI_gi13_RDT1, LM_GI_gi13_RDT2, LM_GI_gi13_RDT3, LM_GI_gi13_sw51, LM_GI_gi13_sw63, LM_Inserts_l10_Cap_Top, LM_Inserts_l10_Parts, LM_Inserts_l11_Parts, LM_Inserts_l12_NewtBall, LM_Inserts_l12_Parts, LM_Inserts_l13_Parts, LM_Inserts_l14_Parts, LM_Inserts_l15_Parts, LM_Inserts_l16_Parts, LM_Inserts_l17_Parts, LM_Inserts_l18_Parts, LM_Inserts_l19_Parts, LM_Inserts_l20_Parts, LM_Inserts_l21_Parts, LM_Inserts_l22_Parts, LM_Inserts_l23_Parts, LM_Inserts_l24_Parts, LM_Inserts_l25_BR1, _
  LM_Inserts_l25_BR3, LM_Inserts_l25_BS1, LM_Inserts_l25_BS2, LM_Inserts_l25_BS3, LM_Inserts_l25_Cap_Bot, LM_Inserts_l25_Diverter, LM_Inserts_l25_NewtBall, LM_Inserts_l25_NewtBall_1, LM_Inserts_l25_Parts, LM_Inserts_l25_Plastics, LM_Inserts_l25_Playfield, LM_Inserts_l25_RDT2, LM_Inserts_l25_RDT3, LM_Inserts_l25_Rails, LM_Inserts_l26_BR2, LM_Inserts_l26_BR3, LM_Inserts_l26_BS1, LM_Inserts_l26_BS2, LM_Inserts_l26_BS3, LM_Inserts_l26_Cap_Bot, LM_Inserts_l26_Cap_Top, LM_Inserts_l26_Diverter, LM_Inserts_l26_FlipperR1, LM_Inserts_l26_NewtBall, LM_Inserts_l26_NewtBall_1, LM_Inserts_l26_Parts, LM_Inserts_l26_Plastics, LM_Inserts_l26_Playfield, LM_Inserts_l26_RDT1, LM_Inserts_l26_RDT2, LM_Inserts_l26_Rails, LM_Inserts_l26_sw51, LM_Inserts_l27_Parts, LM_Inserts_l28_Parts, LM_Inserts_l28_Plastics, LM_Inserts_l29_Parts, LM_Inserts_l3_NewtBall, LM_Inserts_l3_NewtBall_1, LM_Inserts_l3_Parts, LM_Inserts_l30_Parts, LM_Inserts_l31_Parts, LM_Inserts_l32_Parts, LM_Inserts_l33_Parts, LM_Inserts_l34_Parts, LM_Inserts_l35_Parts, _
  LM_Inserts_l36_Parts, LM_Inserts_l37_Parts, LM_Inserts_l38_Parts, LM_Inserts_l39_Parts, LM_Inserts_l4_NewtBall, LM_Inserts_l4_NewtBall_1, LM_Inserts_l4_Parts, LM_Inserts_l40_Parts, LM_Inserts_l41_Parts, LM_Inserts_l42_Parts, LM_Inserts_l43_Parts, LM_Inserts_l44_Parts, LM_Inserts_l45_Flap_L, LM_Inserts_l45_Parts, LM_Inserts_l46_Parts, LM_Inserts_l46_RDT3, LM_Inserts_l47_Parts, LM_Inserts_l47_RDT2, LM_Inserts_l48_Parts, LM_Inserts_l48_RDT1, LM_Inserts_l49_Parts, LM_Inserts_l5_Cap_Top, LM_Inserts_l5_Parts, LM_Inserts_l50_Parts, LM_Inserts_l51_Parts, LM_Inserts_l52_Parts, LM_Inserts_l53_Parts, LM_Inserts_l54_Parts, LM_Inserts_l55_Parts, LM_Inserts_l56_Parts, LM_Inserts_l57_Parts, LM_Inserts_l58_Parts, LM_Inserts_l59_Parts, LM_Inserts_l6_Parts, LM_Inserts_l60_Parts, LM_Inserts_l61_Parts, LM_Inserts_l62_Parts, LM_Inserts_l63_Parts, LM_Inserts_l64_Parts, LM_Inserts_l7_Parts, LM_Inserts_l7_Playfield, LM_Inserts_l7_RDT1, LM_Inserts_l7_RDT2, LM_Inserts_l7_RDT3, LM_Inserts_l8_NewtBall, LM_Inserts_l8_Parts, _
  LM_Inserts_l8_Plastics, LM_Inserts_l9_Parts)
' VLM  Arrays - End


'*******************************************
'  Options & Mods
'*******************************************

Dim LightLevel : LightLevel = -1
Dim PlungerVis : PlungerVis = 1       'Plunger position visualization (0-off, 1-On)
Dim VolumeDial : VolumeDial = 0.8           'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5   'Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5   'Level of ramp rolling volume. Value between 0 and 1

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

    Dim x, v


    ' Room Brightness
  v = Int(5 + 100 * (NightDay^2.2)/(100.0^2.2))
  LightLevel = v
  UpdateMaterial "VLM.Bake.Active", 0, 0, 0, 0, 0, 0, 1, RGB(v, v, v), RGB(0, 0, 0), RGB(0, 0, 0), False, True, 0, 0, 0, 0
  UpdateMaterial "VLM.Bake.Solid", 0, 0, 0, 0, 0, 0, 1, RGB(v, v, v), RGB(0, 0, 0), RGB(0, 0, 0), False, False, 0, 0, 0, 0
  UpdateMaterial "VLM.BumperCaps", 0, 0, 0, 0, 0, 0, 1, RGB(v, v, v), RGB(0, 0, 0), RGB(0, 0, 0), False, True, 0, 0, 0, 0
  UpdateMaterial "Playfield", 0, 0, 0, 0, 0, 0, 1, RGB(v, v, v), RGB(0, 0, 0), RGB(0, 0, 0), False, True, 0, 0, 0, 0

  ' Outlane  difficulty
  v = Table1.Option("Out Post Difficulty", 0, 2, 1, 1, 0, Array("Easy", "Medium", "Hard"))
  DifPost1.Collidable = 0: DifPost2.Collidable = 0: DifPost3.Collidable = 0
  Select Case v
    Case 2
      DifPost1.Collidable = 1
      For Each x in BP_Post_R: x.y = 1390: Next
    Case 1
      DifPost2.Collidable = 1
      For Each x in BP_Post_R: x.y = 1403: Next
    Case 0
      DifPost3.Collidable = 1
      For Each x in BP_Post_R: x.y = 1416: Next
  End Select

    ' Plunger Position Visualization
    PlungerVis = Table1.Option("Plunger Position Visualization", 0, 1, 1, 1, 0, Array("Disabled", "Enabled"))
  PlungerLine.visible = PlungerVis

  'Desktop DMD
  v = Table1.Option("Desktop DMD", 0, 1, 1, 1, 0, Array("Not Visible", "Visible"))
  If v = 1 Then
    Scoretext.Visible = 1
  Else
    Scoretext.Visible = 0
  End If

  'Ball Image
  v = Table1.Option("Ball Image", 0, 2, 1, 0, 0, Array("Default", "Mod1", "Mod2"))
  If v = 0 then
    For each x in gBOT: x.Image = "BallEnv-Light100-GIOff-Clamped": Next
  ElseIf v = 1 then
    For each x in gBOT: x.Image = "ball_Test8BWarm": Next
  Else
    For each x in gBOT: x.Image = "ball_Test8B": Next
  End If


  ' Desaturation
    v = Table1.Option("Desaturation", 0, 1, 0.1, 0, 1)
  if v < 0.1 Then
    Table1.ColorGradeImage = ""
  ElseIf v < 0.2 Then
    Table1.ColorGradeImage = "colorgradelut256x16-10"
  ElseIf v < 0.3 Then
    Table1.ColorGradeImage = "colorgradelut256x16-20"
  ElseIf v < 0.4 Then
    Table1.ColorGradeImage = "colorgradelut256x16-30"
  ElseIf v < 0.5 Then
    Table1.ColorGradeImage = "colorgradelut256x16-40"
  ElseIf v < 0.6 Then
    Table1.ColorGradeImage = "colorgradelut256x16-50"
  ElseIf v < 0.7 Then
    Table1.ColorGradeImage = "colorgradelut256x16-60"
  ElseIf v < 0.8 Then
    Table1.ColorGradeImage = "colorgradelut256x16-70"
  ElseIf v < 0.9 Then
    Table1.ColorGradeImage = "colorgradelut256x16-80"
  ElseIf v < 0.10 Then
    Table1.ColorGradeImage = "colorgradelut256x16-90"
  Else
    Table1.ColorGradeImage = "colorgradelut256x16-100"
  End If

  ' Side rails (always for VR, and if not in cabinet mode)
  If RenderingMode <> 2 Then
    v = Table1.Option("Side Rails", 0, 1, 1, 1, 0, Array("Off", "On"))
    For Each x in BP_Rails: x.Visible = v: Next
    v = v And Table1.ShowDT
    For Each x in BP_Rails_DT: x.Visible = v: Next
  End If

  If RenderingMode = 2 Then
    v = Table1.Option("Show Clock", 0, 1, 1, 1, 0, Array("Off", "On"))
    Dim VR_Obj : For Each VR_Obj in VRClock : VR_Obj.Visible = v : Next
    ClockTimer.Enabled = v
  End If

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    RampRollVolume = Table1.Option("Ramp Volume", 0, 1, 0.01, 0.5, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub


'****************
' Flippers
'****************
Const ReflipAngle = 20

Sub SolLLFlipper(Enabled)
  If Enabled Then
    PinCab_LeftFlipperButton.X = 4.448818 + 8
    LF.Fire: LFPress = 1
        FlipperActivate LeftFlipper, LFPress
    If LeftFlipper.currentangle < LeftFlipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    PinCab_LeftFlipperButton.X = 4.448818 - 8
        FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolLRFlipper(Enabled)
  If Enabled Then
    PinCab_RightFlipperButton.X = 10.3443 - 8
    RF.Fire: RFPress = 1
        FlipperActivate RightFlipper, RFPress
    If RightFlipper.currentangle > RightFlipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    PinCab_RightFlipperButton.X = 10.3443 + 8
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart: RFPress = 0
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolURFlipper(Enabled)
  If Enabled Then
    RightUpFlipper.RotateToEnd
    If RightUpFlipper.currentangle > RightUpFlipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightUpFlipper
    Else
      SoundFlipperUpAttackRight RightUpFlipper
      RandomSoundFlipperUpRight RightUpFlipper
    End If
  Else
    RightUpFlipper.RotateToStart
    If RightUpFlipper.currentangle > RightUpFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightUpFlipper
    End If
    'FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub LeftFlipper_Collide(parm)
    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
    LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
    CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
    RightFlipperCollide parm
End Sub

Sub LeftFlipper_Animate
  Dim lfa : lfa = LeftFlipper.CurrentAngle - 90
  FlipperLSh.RotZ = lfa + 90
  Dim v, LM ' Darken light from lane bulbs when bats are up
  v = CInt(100.0 * (LeftFlipper.StartAngle - LeftFlipper.CurrentAngle) / (LeftFlipper.StartAngle - LeftFlipper.EndAngle))
  For each LM in BP_FlipperL : LM.Rotz = lfa : LM.opacity = 100 - v : Next
  For each LM in BP_FlipperLU : LM.Rotz = lfa : LM.opacity = v : Next
  BM_FlipperL.visible  = v < 0.5
  BM_FlipperLU.visible = v > 0.5
End Sub

Sub RightFlipper_Animate
  Dim rfa : rfa = RightFlipper.CurrentAngle - 90
  FlipperRSh.RotZ = rfa + 90
  Dim v, LM ' Darken light from lane bulbs when bats are up
  v = CInt(100.0 * (RightFlipper.StartAngle - RightFlipper.CurrentAngle) / (RightFlipper.StartAngle - RightFlipper.EndAngle))
  For each LM in BP_FlipperR : LM.Rotz = rfa : LM.Opacity = 100 - v : Next
  For each LM in BP_FlipperRU : LM.Rotz = rfa : LM.Opacity = v : Next
  BM_FlipperR.visible  = v < 0.5
  BM_FlipperRU.visible = v > 0.5
End Sub

Sub RightUpFlipper_Animate
  Dim rfa : rfa = RightUpFlipper.CurrentAngle - 90
  FlipperRUSh.RotZ = rfa + 90
  Dim v, LM ' Darken light from lane bulbs when bats are up
  v = CInt(100.0 * (RightUpFlipper.StartAngle - RightUpFlipper.CurrentAngle) / (RightUpFlipper.StartAngle - RightUpFlipper.EndAngle))
  For each LM in BP_FlipperR1 : LM.Rotz = rfa : LM.Opacity = 100 - v : Next
  For each LM in BP_FlipperR1U : LM.Rotz = rfa : LM.Opacity = v : Next
  BM_FlipperR1.visible  = v < 0.5
  BM_FlipperR1U.visible = v > 0.5
End Sub

'**************
' Timers
'**************
Dim VRPlungerystart: VRPlungerystart = PinCab_Shooter.Y
Dim Plungerystart : Plungerystart= Plunger.Y

Dim Bumpers : Bumpers = Array(Bumper1, Bumper2, Bumper3)
Dim FrameTime, LastFrameTime: LastFrameTime = 0
Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  FrameTime = GameTime - LastFrameTime
  LastFrameTime = GameTime

  ' Core tween updates
  UpdateAllTweens FrameTime

  ' Update rolling sounds
  RollingUpdate

  ' Animate VR plunger
  If RenderingMode = 2 Then
    UpdateBeerBubbles FrameTime
    UpdateLavaLamp FrameTime
    PinCab_Shooter.Y =  (VRPlungerystart + Plungerystart - Plunger.y) + (5 * Plunger.Position)
  End If

  ' Plunger visualization line
  If PlungerVis = 1 then PlungerLine.Y = (1980 + Plungerystart - Plunger.y) + (5 * Plunger.Position)

  ' Ball shadow and brightness
  DynamicBSUpdate
  UpdateBallBrightness

  ' Newton baked balls position from hidden physics balls
  Dim a, x, y
  x = NewtonBall.x: y = NewtonBall.y: For Each a in BP_NewtBall: a.x = x: a.y = y: Next
  x = CaptiveBall.x: y = CaptiveBall.y
  For Each a in BP_NewtBall_1: a.x = x: a.y = y: Next

  ' Center post animation (slight moves when in up position and hit)
  If PostIsDropped = False Then PostBallUpdate

  ' Animate Bumper switch
  ' 83 = Bumper skirt radius + 25 => start pressing
  ' 63 = Bumper radius + 25 => fully pressed
  Dim i, nearest, b, z, s
  For i = 0 To 2
    nearest = 10000.
    For s = 0 to UBound(gBOT)
      If gBOT(s).z < 30 Then ' Ball on playfield
        x = Bumpers(i).x - gBOT(s).x
        y = Bumpers(i).y - gBOT(s).y
        b = x * x + y * y
        If b < nearest Then nearest = b
      End If
    Next
    z = 0
    If nearest < 83 * 83 Then
      z = (83.0 - sqr(nearest)) / (83.0 - 63.0)
      z = - z * z * 5.0
      if (z < -5) Then z = -5
    End If
    If i = 0 And BM_BS3.z <> z Then For Each x in BP_BS3: x.Z = z: Next
    If i = 1 And BM_BS1.z <> z Then For Each x in BP_BS1: x.Z = z: Next
    If i = 2 And BM_BS2.z <> z Then For Each x in BP_BS2: x.Z = z: Next
  Next
End Sub

'************
' Table Init
'************
Dim BSBall1, BSBall2, BSBall3, NewtonBall, CaptiveBall, CenterPostBall, gBOT

Sub Table1_Init
  vpmInit Me

  On Error Resume Next
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    If .Version < 03060000 Then MsgBox "This table needs PinMame version 3.6+" : Exit Sub
    .SplashInfoLine = "BreakShot (Capcom 1996)"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .Hidden = 0
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
    On Error Goto 0
  End With
  On Error Goto 0

  vpmNudge.TiltSwitch   = 30
  vpmNudge.Sensitivity  = 5
  vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  vpmMapLights AllLamps

  'Trough
  Set BSBall1 = sw38.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set BSBall2 = sw37.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set BSBall3 = sw36.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Controller.Switch(38) = 1
  Controller.Switch(37) = 1
  Controller.Switch(36) = 1

  'Captive Ball
  Set CaptiveBall = CBKicker.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set NewtonBall = NBKicker.CreateSizedballWithMass(Ballsize/2,Ballmass)
  CaptiveBall.visible = 0
  NewtonBall.visible = 0

  gBOT = Array(NewtonBall,CaptiveBall,BSBall1,BSBall2,BSBall3)

  vpmTimer.AddTimer 300, "CBKicker.kick 180,0 '"
  vpmTimer.AddTimer 310, "CBKicker.enabled= 0 '"

  'Center Post Ball
  Set CenterPostBall = CenterPostKicker.CreateSizedballWithMass(28,Ballmass*2)
  CenterPostBall.visible = 0
  vpmTimer.AddTimer 300, "CenterPostKicker.kick 180,0 '"
  vpmTimer.AddTimer 310, "CenterPostKicker.enabled= 0 '"

  PlungerKB.PullBack

  'Initialize center post at down position
  CenterPostIsDropped True
  Tween(0).SetValue(0)

  'Setup targets
  InitDT
  InitST

  '----- VR Room Auto-Detect -----
  Dim VR_Obj
  If RenderingMode = 2 Then
    For Each VR_Obj in VR_Cab : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in VR_Room : VR_Obj.Visible = 1 : Next
    BM_Playfield.ReflectionProbe = "Playfield Reflections VR" 'Roughness set to 0
  Else
    For Each VR_Obj in VR_Cab : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VR_Room : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRClock : VR_Obj.Visible = 0 : Next
    BM_Playfield.ReflectionProbe = "Playfield Reflections"
  End If

  PlungerLine.blenddisablelighting = 3
End Sub

'**********************************************************************************************************
' Keys
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If keycode = PlungerKey Then Plunger.Pullback:SoundPlungerPull
  If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 2:SoundNudgeCenter()
  If keycode = StartGameKey Then soundStartButton()
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If keycode = PlungerKey Then Plunger.Fire:SoundPlungerReleaseBall
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'**********************************************************************************************************

'Bumpers
Const BumperDrop = 30, BumperDropLength = 30.0, BumperRaiseLength = 100.0
Sub Bumper1_Hit   : vpmTimer.PulseSw(59) : RandomSoundBumperTop Bumper1 : Tween(3).Clear().TargetSpeed(-BumperDrop, -BumperDrop / BumperDropLength).TargetSpeed(0, BumperDrop / BumperRaiseLength).Done : End Sub
Tween(3).Callback = GetRef("PrimArrayTransZUpdate") : Tween(3).CallbackParam = BP_BR3
Sub Bumper2_Hit   : vpmTimer.PulseSw(58) : RandomSoundBumperTop Bumper2 : Tween(4).Clear().TargetSpeed(-BumperDrop, -BumperDrop / BumperDropLength).TargetSpeed(0, BumperDrop / BumperRaiseLength).Done : End Sub
Tween(4).Callback = GetRef("PrimArrayTransZUpdate") : Tween(4).CallbackParam = BP_BR1
Sub Bumper3_Hit   : vpmTimer.PulseSw(57) : RandomSoundBumperTop Bumper3 : Tween(5).Clear().TargetSpeed(-BumperDrop, -BumperDrop / BumperDropLength).TargetSpeed(0, BumperDrop / BumperRaiseLength).Done : End Sub
Tween(5).Callback = GetRef("PrimArrayTransZUpdate") : Tween(5).CallbackParam = BP_BR2

'Rollovers
Sub SW43_Hit:Controller.Switch(43)=1 : End Sub
Sub SW43_unHit:Controller.Switch(43)=0:End Sub
' plunger lane is not baked Sub SW43_Animate: Dim a, x: a = sw43.CurrentAnimOffset: For Each x in BP_sw43: x.transz = a: Next: End Sub

Sub SW44_Hit:Controller.Switch(44)=1 : End Sub
Sub SW44_unHit:Controller.Switch(44)=0:End Sub
Sub SW44_Animate: Dim a, x: a = sw44.CurrentAnimOffset: For Each x in BP_sw44: x.transz = a: Next: End Sub

Sub SW45_Hit:Controller.Switch(45)=1 : End Sub
Sub SW45_unHit:Controller.Switch(45)=0:End Sub
Sub SW45_Animate: Dim a, x: a = sw45.CurrentAnimOffset: For Each x in BP_sw45: x.transz = a: Next: End Sub

Sub SW46_Hit:Controller.Switch(46)=1 : End Sub
Sub SW46_unHit:Controller.Switch(46)=0:End Sub
Sub SW46_Animate: Dim a, x: a = sw46.CurrentAnimOffset: For Each x in BP_sw46: x.transz = a: Next: End Sub

Sub SW47_Hit:Controller.Switch(47)=1 : End Sub
Sub SW47_unHit:Controller.Switch(47)=0:End Sub
Sub SW47_Animate: Dim a, x: a = sw47.CurrentAnimOffset: For Each x in BP_sw47: x.transz = a: Next: End Sub

Sub SW60_Hit:Controller.Switch(60)=1 : End Sub
Sub SW60_unHit:Controller.Switch(60)=0:End Sub
Sub SW60_Animate: Dim a, x: a = sw60.CurrentAnimOffset: For Each x in BP_sw60: x.transz = a: Next: End Sub

Sub SW61_Hit:Controller.Switch(61)=1 : End Sub
Sub SW61_unHit:Controller.Switch(61)=0:End Sub
Sub SW61_Animate: Dim a, x: a = sw61.CurrentAnimOffset: For Each x in BP_sw61: x.transz = a: Next: End Sub

Sub SW62_Hit:Controller.Switch(62)=1 : End Sub
Sub SW62_unHit:Controller.Switch(62)=0:End Sub
Sub SW62_Animate: Dim a, x: a = sw62.CurrentAnimOffset: For Each x in BP_sw62: x.transz = a: Next: End Sub

Sub SW63_Hit:Controller.Switch(63)=1 : End Sub
Sub SW63_unHit:Controller.Switch(63)=0:End Sub
Sub SW63_Animate: Dim a, x: a = sw63.CurrentAnimOffset: For Each x in BP_sw63: x.transz = a: Next: End Sub

Sub SW79_Hit:Controller.Switch(79)=1 : End Sub
Sub SW79_unHit:Controller.Switch(79)=0:End Sub
Sub SW79_Animate: Dim a, x: a = sw79.CurrentAnimOffset: For Each x in BP_sw79: x.transz = a: Next: End Sub

'Opto's
Sub SW53_Hit:Controller.Switch(53)=1 : End Sub
Sub SW53_unHit:Controller.Switch(53)=0:End Sub
Sub SW54_Hit:Controller.Switch(54)=1 : End Sub
Sub SW54_unHit:Controller.Switch(54)=0:End Sub
Sub SW65_Hit:Controller.Switch(65)=1 : End Sub
Sub SW65_unHit:Controller.Switch(65)=0:End Sub


'********************************************
' ZTAR: Targets
'********************************************

Sub sw51_Hit : STHit 51 : End Sub
Sub sw51o_Hit : TargetBouncer ActiveBall, 1 : End Sub
Sub sw55_Hit : STHit 55 : End Sub
Sub sw55o_Hit : TargetBouncer ActiveBall, 1 : End Sub
Sub sw76_Hit : STHit 76 : End Sub
Sub sw76o_Hit : TargetBouncer ActiveBall, 1 : End Sub
Sub sw77_Hit : STHit 77 : End Sub
Sub sw77o_Hit : TargetBouncer ActiveBall, 1 : End Sub
Sub sw78_Hit : STHit 78 : End Sub
Sub sw78o_Hit : TargetBouncer ActiveBall, 1 : End Sub

'********************************************
'  Drop Target Controls
'********************************************

Sub sw48_Hit : DTHit 48 : End Sub
Sub sw49_Hit : DTHit 49 : End Sub
Sub sw50_Hit : DTHit 50 : End Sub
Sub sw73_Hit : DTHit 73 : End Sub
Sub sw74_Hit : DTHit 74 : End Sub
Sub sw75_Hit : DTHit 75 : End Sub

Sub ResetDropsRight(enabled)
  If enabled Then
    RandomSoundDropTargetReset BM_RDT2
    DTRaise 48
    DTRaise 49
    DTRaise 50
  End If
End Sub

Sub ResetDropsLeft(enabled)
  If enabled Then
    RandomSoundDropTargetReset BM_LDT2
    DTRaise 73
    DTRaise 74
    DTRaise 75
  End If
End Sub

'******************************************************
' Trough
'******************************************************

Sub sw38_Hit():Controller.Switch(38) = 1:UpdateTrough:End Sub
Sub sw38_UnHit():Controller.Switch(38) = 0:UpdateTrough:End Sub
Sub sw37_Hit():Controller.Switch(37) = 1:UpdateTrough:End Sub
Sub sw37_UnHit():Controller.Switch(37) = 0:UpdateTrough:End Sub
Sub sw36_Hit():Controller.Switch(36) = 1:UpdateTrough:End Sub
Sub sw36_UnHit():Controller.Switch(36) = 0:UpdateTrough:End Sub

Sub sw35_Hit():Controller.Switch(35) = 1:UpdateTrough:RandomSoundDrain sw35:End Sub
Sub sw35_UnHit():Controller.Switch(35) = 0:UpdateTrough:End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If sw36.BallCntOver = 0 Then sw37.kick 60, 9
  If sw37.BallCntOver = 0 Then sw38.kick 60, 9
  Me.Enabled = 0
End Sub

'******************************************************
' Drain & Release
'******************************************************

Sub SolTrough(enabled)
  If enabled Then
    sw35.kick 57, 20
  End If
End Sub

Sub SolRelease(enabled)
  If enabled Then
    sw36.kick 60, 9
    RandomSoundBallRelease sw36
  End If
End Sub

'**********************************************************************************************************
' Solenoid Subs
'**********************************************************************************************************

Sub Kickback(Enabled)
  If Enabled Then
    SoundKicker
    PlungerKB.Fire
  Else
    PlungerKB.PullBack
  End If
End Sub

'**********************************************************************************************************
' Center Post
'**********************************************************************************************************

Dim PostIsDropped : PostIsDropped = False

Tween(0).Callback = GetRef("CenterPostUpdate")
Sub PostUp(Enabled) : If Enabled Then CenterPostIsDropped False : Tween(0).Clear().TargetSpeed(30, 30.0/60.0).Run(GetRef("OnCenterPostUp")) : End If : End Sub
Sub PostRelease(Enabled) : If Enabled Then CenterPostIsDropped True : Tween(0).Clear().TargetSpeed(0, 37.0/180.0).Run(GetRef("OnCenterPostDown")) : End If : End Sub
Sub OnCenterPostUp(id, z) : Controller.Switch(66) = 1 : End Sub
Sub OnCenterPostDown(id, z) : Controller.Switch(66) = 0 : End Sub
Sub CenterPostUpdate(id, z) : Dim a : For Each a in BP_Diverter: a.z = z: Next : End Sub

Sub PostBallUpdate
  Dim x : x = CenterPostBall.x
  Dim y : y = CenterPostBall.y
  Dim a : For Each a in BP_Diverter: a.x = x: a.y = y:Next
End Sub

Sub CenterPostIsDropped(dropped)
  If dropped <> PostIsDropped Then
    PostIsDropped = dropped
    CenterPostWall001.IsDropped = dropped
    CenterPostWall002.IsDropped = dropped
    CenterPostWall003.IsDropped = dropped
    If dropped Then
      CenterPostBall.x = 142
      CenterPostBall.y = 2040
    Else
      PostTimer_Timer
    End If
  End If
End Sub

dim PostCount: PostCount = 0
Sub PostTimer_Timer
  If PostTimer.Enabled=False Then
    PostCount = 0
    PostTimer.Enabled = True
  End If
  CenterPostBall.x = 431
  CenterPostBall.y = 894
  PostCount = PostCount + 1
  If PostCount >= 10 Then PostTimer.Enabled = False
End Sub


'******************************************
'       Solenoid Controlled Gates
'******************************************

Dim RightGateSolState: RightGateSolState = False

Sub RightGate_Animate
  Dim a, r
  r = -RightGate.CurrentAngle / 1.5
  For Each a in BP_Flap_R: a.RotY = r: Next
End Sub

Sub RightGate_Timer
  RightGate.TimerEnabled = 0
  RightGate.Open = False
  Dim a: For Each a in BP_Gate_R_W: a.z = 0: Next
End Sub

Sub SolRightGate(Enabled)
  'Debug.print "Sol Right Gate " & Enabled
  Dim a
  RightGate.TimerEnabled = 0
  If Enabled Then
    RightGate.Open = True
    For Each a in BP_Gate_R_W: a.z = -10: Next
  Else ' Filters out PWM not handled by PinMame
    RightGate.TimerEnabled = 1
  End If
  RightGate_Animate
End Sub

Sub LeftGate_Animate
  Dim a, r
  r = -LeftGate.CurrentAngle / 1.5
  For Each a in BP_Flap_L: a.RotY = r: Next
End Sub

Sub LeftGate_Timer
  LeftGate.TimerEnabled = 0
  LeftGate.Open = False
  Dim a: For Each a in BP_Gate_L_W: a.z = 0: Next
End Sub

Sub SolLeftGate(Enabled)
  'Debug.print "Sol Left Gate " & Enabled
  Dim a
  LeftGate.TimerEnabled = 0
  If Enabled Then
    LeftGate.Open = True
    For Each a in BP_Gate_L_W: a.z = -10: Next
  Else ' Filters out PWM not handled by PinMame
    LeftGate.TimerEnabled = 1
  End If
  LeftGate_Animate
End Sub

'****************************************
'   VUKs
'****************************************
Dim KickerBall67, KickerBall68, KickerBall69, KickerBall70

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle : rangle = PI * (kangle - 90) / 180
  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

'Right VUK
Sub sw67_Hit : set KickerBall67 = activeball : Controller.Switch(67) = 1 : SoundSaucerLock : End Sub
Sub sw67_UnHit : Controller.Switch(67) = 0 : End Sub
Sub RightSaucer(Enable)
    If Enable then
    If Controller.Switch(67) <> 0 Then
      KickBall KickerBall67, 0, 0, 50, 0
      SoundSaucerKick 1, sw67
    End If
  End If
End Sub

'Centre Left Kicker
Sub sw68_Hit : set KickerBall68 = activeball : Controller.Switch(68) = 1 : SoundSaucerLock : End Sub
Sub sw68_UnHit : Controller.Switch(68) = 0 : End Sub
Sub CLeftPocket(Enable)
    If Enable then
    If Controller.Switch(68) <> 0 Then
      KickBall KickerBall68, 0, 0, 15, 0
      SoundSaucerKick 1, sw68
    End If
  End If
End Sub

'Centre Centre Kicker
Sub sw69_Hit : set KickerBall69 = activeball : Controller.Switch(69) = 1 : SoundSaucerLock : End Sub
Sub sw69_UnHit : Controller.Switch(69) = 0 : End Sub
Sub CenterPocket(Enable)
    If Enable then
    If Controller.Switch(69) <> 0 Then
      KickBall KickerBall69, 0, 0, 15, 0
      SoundSaucerKick 1, sw69
    End If
  End If
End Sub

'Centre Right Kicker
Sub sw70_Hit : set KickerBall70 = activeball : Controller.Switch(70) = 1 : SoundSaucerLock : End Sub
Sub sw70_UnHit : Controller.Switch(70) = 0 : End Sub
Sub CRightPocket(Enable)
    If Enable then
    If Controller.Switch(70) <> 0 Then
      KickBall KickerBall70, 0, 0, 15, 0
      SoundSaucerKick 1, sw70
    End If
  End If
End Sub

'****************************************************************
' ZSLG: Slingshots
'****************************************************************

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  RandomSoundSlingshotRight BM_REMK
  Tween(1).Clear().SetValue(2).Delay(30).SetValue(1).Delay(10).SetValue(0)
End Sub

Tween(1).Callback = GetRef("RightSlingShotUpdate")
Tween(1).SetValue(0)
Sub RightSlingShotUpdate(id, step)
  Debug.Print "RightSlingShotUpdate " & id & " / " & step
  Dim x, x1, x2, y : Select Case step
  Case 2: x1 = True  : x2 = False : y = 20 : Controller.Switch(41) = 1 ' Fully out
  Case 1: x1 = False : x2 = True  : y = 10 : Controller.Switch(41) = 1
  Case 0: x1 = False : x2 = False : y =  0 : Controller.Switch(41) = 0 ' Rest pose
  End Select
  For Each x in BP_RSling1: x.Visible = x1: Next
  For Each x in BP_RSling2: x.Visible = x2: Next
  For Each x in BP_REMK: x.transx = y: Next
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(Activeball)
  RandomSoundSlingshotLeft BM_LEMK
  Tween(2).Clear().SetValue(2).Delay(30).SetValue(1).Delay(10).SetValue(0)
End Sub

Tween(2).Callback = GetRef("LeftSlingShotUpdate")
Tween(2).SetValue(0)
Sub LeftSlingShotUpdate(id, step)
  Dim x, x1, x2, y : Select Case step
  Case 2: x1 = True  : x2 = False : y = 20 : Controller.Switch(42) = 1 ' Fully out
  Case 1: x1 = False : x2 = True  : y = 10 : Controller.Switch(42) = 1
  Case 0: x1 = False : x2 = False : y =  0 : Controller.Switch(42) = 0' Rest pose
  End Select
  For Each x in BP_LSling1: x.Visible = x1: Next
  For Each x in BP_LSling2: x.Visible = x2: Next
  For Each x in BP_LEMK: x.transx = y: Next
End Sub


'***************************************************************
' ZSHA: VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

Const AmbientBallShadowOn = 1   '0 = Static shadow under ball ("flasher" image, like JP's)
'                 '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'                 '2 = flasher image shadow, but it moves like ninuzzu's

' *** Required Functions, enable these if they are not already present elswhere in your table
Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

Function min(a,b)
  If a < b Then
    min = a
  Else
    min = b
  End If
End Function

' *** These define the appearance of shadows in your table  ***

'Ambient (Room light source)
Const AmbientBSFactor = 0.9  '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 5        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(5)
Dim OnPF(5)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4)

' *** The Shadow Dictionary
Dim bsDict
Set bsDict = New cvpmDictionary
Const bsNone = "None"
Const bsWire = "Wire"
Const bsRamp = "Ramp"
Const bsRampClear = "Clear"

'Initialization
DynamicBSInit

Sub DynamicBSInit()
  Dim iii, source
  For iii = 0 To tnob - 1
    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii / 1000 + 0.04
    objBallShadow(iii).visible = 0
    BallShadowA(iii).Opacity = 100 * AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next
End Sub

Sub BallOnPlayfieldNow(onPlayfield, ballNum)  'Whether a ball is currently on the playfield. Only update certain things once, save some cycles
  If onPlayfield Then
    OnPF(ballNum) = True
    bsRampOff gBOT(ballNum).ID
    '   debug.print "Back on PF"
    UpdateMaterial objBallShadow(ballNum).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(ballNum).size_x = 5
    objBallShadow(ballNum).size_y = 4.5
    objBallShadow(ballNum).visible = 1
    BallShadowA(ballNum).visible = 0
    BallShadowA(ballNum).Opacity = 100 * AmbientBSFactor
  Else
    OnPF(ballNum) = False
    '   debug.print "Leaving PF"
  End If
End Sub

Sub DynamicBSUpdate
  Dim falloff 'Max distance to light sources, can be changed dynamically if you have a reason
  falloff = 150
  Dim ShadowOpacity1, ShadowOpacity2
  Dim s, LSd, iii
  Dim dist1, dist2, src1, src2
  Dim bsRampType
  '   Dim gBOT: gBOT=getballs 'Uncomment if you're destroying balls - Not recommended! #SaveTheBalls

  'Hide shadow of deleted balls
  For s = UBound(gBOT) + 1 To tnob - 1
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(gBOT) < lob Then Exit Sub 'No balls in play, exit

  'The Magic happens now
  For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow
    'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your Elseif segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else (under 20)

    'Primitive shadow on playfield, flasher shadow in ramps
    If AmbientBallShadowOn = 1 Then
      '** Above the playfield
      If gBOT(s).Z > 30 Then
        If OnPF(s) Then BallOnPlayfieldNow False, s   'One-time update
        bsRampType = getBsRampType(gBOT(s).id)
        '   debug.print bsRampType

        If Not bsRampType = bsRamp Then 'Primitive visible on PF
          objBallShadow(s).visible = 1
          objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (Table1.Width / 2)) / (Ballsize / AmbientMovement) + offsetX
          objBallShadow(s).Y = gBOT(s).Y + offsetY
          objBallShadow(s).size_x = 5 * ((gBOT(s).Z + BallSize) / 80) 'Shadow gets larger and more diffuse as it moves up
          objBallShadow(s).size_y = 4.5 * ((gBOT(s).Z + BallSize) / 80)
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (30 / (gBOT(s).Z)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else 'Opaque, no primitive below
          objBallShadow(s).visible = 0
        End If

        If bsRampType = bsRampClear Or bsRampType = bsRamp Then 'Flasher visible on opaque ramp
          BallShadowA(s).visible = 1
          BallShadowA(s).X = gBOT(s).X + offsetX
          BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
          BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
          If bsRampType = bsRampClear Then BallShadowA(s).Opacity = 50 * AmbientBSFactor
        ElseIf bsRampType = bsWire Or bsRampType = bsNone Then 'Turn it off on wires or falling out of a ramp
          BallShadowA(s).visible = 0
        End If

        '** On pf, primitive only
      ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then
        If Not OnPF(s) Then BallOnPlayfieldNow True, s
        objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (Table1.Width / 2)) / (Ballsize / AmbientMovement) + offsetX
        objBallShadow(s).Y = gBOT(s).Y + offsetY
        '   objBallShadow(s).Z = gBOT(s).Z + s/1000 + 0.04    'Uncomment (and adjust If/Elseif height logic) if you want the primitive shadow on an upper/split pf

        '** Under pf, flasher shadow only
      Else
        If OnPF(s) Then BallOnPlayfieldNow False, s
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 1
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000
      End If

      'Flasher shadow everywhere
    ElseIf AmbientBallShadowOn = 2 Then
      If gBOT(s).Z > 30 Then 'In a ramp
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then 'On pf
        BallShadowA(s).visible = 1
        BallShadowA(s).X = gBOT(s).X + (gBOT(s).X - (Table1.Width / 2)) / (Ballsize / AmbientMovement) + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = 1.04 + s / 1000
      Else 'Under pf
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000
      End If
    End If
  Next
End Sub

' *** Ramp type definitions

Sub bsRampOnWire()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsWire
  Else
    bsDict.Add ActiveBall.ID, bsWire
  End If
End Sub

Sub bsRampOn()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsRamp
  Else
    bsDict.Add ActiveBall.ID, bsRamp
  End If
End Sub

Sub bsRampOnClear()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsRampClear
  Else
    bsDict.Add ActiveBall.ID, bsRampClear
  End If
End Sub

Sub bsRampOff(idx)
  If bsDict.Exists(idx) Then
    bsDict.Item(idx) = bsNone
  End If
End Sub

Function getBsRampType(id)
  Dim retValue
  If bsDict.Exists(id) Then
    retValue = bsDict.Item(id)
  Else
    retValue = bsNone
  End If
  getBsRampType = retValue
End Function

'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

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
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
  '   Dim BOT
  '   BOT = GetBalls

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


'*****************
' Maths
'*****************

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

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
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
Const EOSReturn = 0.035  'mid 80's to early 90's
'   Const EOSReturn = 0.025  'mid 90's and later

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
  Flipper.eostorque = EOST * EOSReturn / FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b', BOT
    '   BOT = GetBalls

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

' This collection contains the bottom sling posts. They are not in the dPosts collection so that the TargetBouncer is not applied to them, but they should still have dampening applied
' If you experience airballs with posts or targets, consider adding them to this collection
Sub NoTargetBouncer_Hit
    RubbersD.dampen ActiveBall
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

' FIXME for the time being, the cor timer interval must be 10 ms (so below 60FPS framerate)
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

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

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7  'Level of bounces. Recommmended value of 0.7

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






'******************************************************
'   ZRDT:  DROP TARGETS by Rothbauerw
'******************************************************
' This solution improves the physics for drop targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full drop target animation, including deflection on hit and a slight lift when the drop
' targets raise, switch handling, bricking, and popping the ball up if it's over the drop target when it raises.
'
' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.
'
' WARNING this is a heavily experimental modded version to try leveraging the tween class

'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

Class DropTarget
  Private m_primary, m_secondary, m_prims, m_sw, m_isDropped

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Secondary(): Set Secondary = m_secondary: End Property
  Public Property Let Secondary(input): Set m_secondary = input: End Property

  Public Property Get Prims(): Prims = m_prims: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
  Public Property Let IsDropped(input): m_isDropped = input: End Property

  Public default Function init(primary, secondary, prims, sw, isDropped)
    Set m_primary = primary
    Set m_secondary = secondary
    m_prims = prims
    m_sw = sw
    m_isDropped = isDropped
    Set Init = Me
  End Function
End Class

'Define a variable for each drop target
Dim DT48, DT49, DT50, DT73, DT74, DT75

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:  primary target wall to determine drop
'   secondary:  wall used to simulate the ball striking a bent or offset target after the initial Hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          rotz must be used for orientation
'          rotx to bend the target back
'          transz to move it up and down
'          the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
'   switch:  ROM switch number
'   isDropped:  Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.
'         Use the function DTDropped(switchid) to check a target's drop status.

Set DT48 = (new DropTarget)(sw48, sw48a, BP_RDT1, 48, False)
Set DT49 = (new DropTarget)(sw49, sw49a, BP_RDT2, 49, False)
Set DT50 = (new DropTarget)(sw50, sw50a, BP_RDT3, 50, False)
Set DT73 = (new DropTarget)(sw73, sw73a, BP_LDT1, 73, False)
Set DT74 = (new DropTarget)(sw74, sw74a, BP_LDT2, 74, False)
Set DT75 = (new DropTarget)(sw75, sw75a, BP_LDT3, 75, False)

Dim DTArray: DTArray = Array(DT48, DT49, DT50, DT73, DT74, DT75)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 54 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance
Const DTTween = 6 ' Index of first tween to use for drop targets

'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub InitDT
  Dim i: For i = 0 To UBound(DTArray)
    Tween(DTTween + i * 2 + 0).Callback = GetRef("DTBendCallback")
    Tween(DTTween + i * 2 + 0).CallbackParam = DTArray(i).Prims
    Tween(DTTween + i * 2 + 1).Callback = GetRef("PrimArrayTransZUpdate")
    Tween(DTTween + i * 2 + 1).CallbackParam = DTArray(i).Prims
  Next
End Sub

Sub DTBendCallback(id, prims, v)
  Dim rangle, c, s: rangle = prims(0).rotz * PI / 180 : c = v * Cos(rangle) : s = v * Sin(rangle)
  Dim a: For Each a in prims: a.rotx = c: a.roty = s: Next
End Sub

Function DTArrayID(switch)
  Dim i : For i = 0 To UBound(DTArray)
    If DTArray(i).sw = switch Then DTArrayID = i : Exit Function
  Next
End Function

Function DTDropped(switchid)
  Dim ind : ind = DTArrayID(switchid)
  DTDropped = DTArray(ind).isDropped
End Function

Sub DTBallPhysics(aBall, angle, mass)
  Dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * Cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Cos(rangle + 4 * Atn(1) / 2)
  calc3 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Sin(rangle + 4 * Atn(1) / 2)

  aBall.velx = calc1 * Cos(rangle) + calc2
  aBall.vely = calc1 * Sin(rangle) + calc3
End Sub

Sub DTHit(switch)
  Dim i: i = DTArrayID(switch)
  Dim prims: prims = DTArray(i).Prims
  Dim dtprim: Set dtprim = prims(0)
  Dim aBall : Set aBall = ActiveBall

  PlayTargetSound

  ' Check if target is hit on it's face or sides and whether a 'brick' occurred (high velocity resulting in the ball bouncing without the target dropping)
  Dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aBall.id),cor.ballvelx(aBall.id))
  bangleafter = Atn2(aBall.vely,aBall.velx)

  Xintersect = (aBall.y - dtprim.y - Tan(bangle) * aBall.x + Tan(rangle2) * dtprim.x) / (Tan(rangle2) - Tan(bangle))
  Yintersect = Tan(rangle2) * Xintersect + (dtprim.y - Tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aBall.id) * Cos(bangle - rangle)
  paravel = cor.BallVel(aBall.id) * Sin(bangle - rangle)

  perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

  If perpvel > 0 And  perpvelafter <= 0 Then
    If DTEnableBrick = 1 And  perpvel > DTBrickVel And DTBrickVel <> 0 And cdist < 8 Then
      ' Brick: Bend back then forth, toggling forth and back from primary to secondary collidables
      Tween(DTTween + i * 2 + 0).Clear().Run2(GetRef("SetDTCollidable"), i, 0).TargetLength(DTMaxBend, DTDropDelay / 2).TargetLength(0, DTDropDelay / 2).Run2(GetRef("SetDTCollidable"), i, 1).Done
    Else
      DTDropSwipe switch, False ' Normal Hit
    End If
    DTBallPhysics aBall, dtprim.rotz, DTMass
  ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
    DTDropSwipe switch, True ' Swipe Hit
    DTBallPhysics aBall, dtprim.rotz, DTMass
  End If
End Sub

Sub DTRaise(switch)
  Dim i, prims : i = DTArrayID(switch) : prims = DTArray(i).Prims
  DTArray(i).isDropped = False
  Controller.Switch(switch) = 0

  ' If raising from dropped position, push ball above target up
  If prims(0).transz = -DTDropUnits Then
    Dim b, gBOT : gBOT = GetBalls
    For b = 0 To UBound(gBOT)
      If InRotRect(gBOT(b).x,gBOT(b).y,Prims(0).x, Prims(0).y, Prims(0).rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < Prims(0).z + DTDropUnits + 25 Then
        gBOT(b).velz = 20
      End If
    Next
  End If

  ' Raise up slightly higher than raise position, wait a bit then drop back to rest position
  Tween(DTTween + i * 2 + 0).Clear().TargetLength(0, DTDropUpSpeed).Done
  Tween(DTTween + i * 2 + 1).Clear().Run2(GetRef("SetDTCollidable"), i, 0).TargetLength(DTDropUpUnits, DTDropUpSpeed).Delay(DTRaiseDelay - DTDropUpSpeed).Run2(GetRef("SetDTCollidable"), i, 1).TargetLength(0, DTDropSpeed).Done
End Sub

Sub DTDrop(switch) : DTDropSwipe switch, False : End Sub
Sub DTDropSwipe(switch, isSwipe)
  Dim i : i = DTArrayID(switch)
  DTArray(i).primary.collidable = 0
  If isSwipe Then DTArray(i).secondary.collidable = 0 Else DTArray(i).secondary.collidable = 1

  ' Bend back then drop adjusting bend while dropping
  Tween(DTTween + i * 2 + 0).Clear().TargetLength(DTMaxBend, DTDropDelay).TargetLength( DTMaxBend/2, DTDropSpeed).Done
  Tween(DTTween + i * 2 + 1).Clear().Delay(DTDropDelay).Run1(GetRef("OnDTDrop"), i).TargetLength(-DTDropUnits, DTDropSpeed).Run1(GetRef("OnDTDropped"), i).Done
End Sub

Sub SetDTCollidable(id, v, index, state) : DTArray(index).Primary.Collidable = state : DTArray(index).Secondary.Collidable = 1-state :  End Sub
Sub OnDTDrop(id, v, index) : Dim prims : prims = DTArray(index).Prims : SoundDropTargetDrop prims(0) :  End Sub
Sub OnDTDropped(id, v, index) : DTArray(index).Secondary.collidable = 0 : DTArray(index).isDropped = True : Controller.Switch(DTArray(index).Sw) = 1 : End Sub


'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS
'******************************************************

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
'****  END DROP TARGETS
'******************************************************




'******************************************************
' ZRST: STAND-UP TARGETS by Rothbauerw
'******************************************************

Class StandupTarget
  Private m_primary, m_prims, m_sw

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Prims(): Prims = m_prims: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public default Function init(primary, prims, sw)
    Set m_primary = primary
    m_prims = prims
    m_sw = sw
    Set Init = Me
  End Function
End Class

'Define a variable for each stand-up target
Dim ST51, ST55, ST76, ST77, ST78

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, switch)
'   primary:   vp target to determine target hit
'   prim:    primitive target used for visuals and animation. IMPORTANT!!! transy must be used to offset the target animation
'   switch:  ROM switch number
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST51 = (new StandupTarget)(sw51, BP_sw51, 51)
Set ST55 = (new StandupTarget)(sw55, BP_sw55, 55)
Set ST76 = (new StandupTarget)(sw76, BP_sw76, 76)
Set ST77 = (new StandupTarget)(sw77, BP_sw77, 77)
Set ST78 = (new StandupTarget)(sw78, BP_sw78, 78)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray : STArray = Array(ST51, ST55, ST76, ST77, ST78)

'Configure the behavior of Stand-up Targets
Const STAnimLength = 80  'length in ms of animation (time to get back)
Const STMaxOffset = 9   'max vp units target moves when hit
Const STMass = 0.2    'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance
Const STTween = 18 ' Index of first tween to use for stand up targets

'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub InitST
  Dim i: For i = 0 To UBound(STArray)
    Tween(STTween + i).Callback = GetRef("PrimArrayTransYUpdate")
    Tween(STTween + i).CallbackParam = STArray(i).Prims
  Next
End Sub

Function STArrayID(switch)
  Dim i : For i = 0 To UBound(STArray)
    If STArray(i).sw = switch Then STArrayID = i : Exit Function
  Next
End Function

Sub STHit(switch)
  Dim i : i = STArrayID(switch)
  Dim aBall : Set aBall = ActiveBall
  Dim target : Set target = STArray(i).Primary

  PlayTargetSound

  Dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
  rangle = (target.orientation - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
  paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

  perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

  If (perpvel > 0 And  perpvelafter <= 0) _
  Or (perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)))Then
    STArray(i).Primary.Collidable = 0
    Controller.Switch(switch) = 1
    Tween(STTween + i).Clear().SetValue(-STMaxOffset).TargetLength(0, STAnimLength).Run1(GetRef("OnSTBounced"), i).Done
    DTBallPhysics aBall, STArray(i).primary.orientation, STMass
  End If
End Sub

Sub OnSTBounced(id, v, index) : STArray(index).Primary.Collidable = 1 : Controller.Switch(STArray(index).Sw) = 0 : End Sub

'******************************************************
'****   END STAND-UP TARGETS
'******************************************************


'******************************************************
' ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i : For i = 0 To tnob
    rolling(i) = False
  Next
End Sub

Sub RollingUpdate()
  Dim b
  '   Dim BOT
  '   BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 To tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(gBOT)
    If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ <  - 1 And gBOT(b).z < 55 And gBOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz >  - 7 Then
          RandomSoundBallBouncePlayfieldSoft gBOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard gBOT(b)
        End If
      End If
    End If

    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If gBOT(b).Z > 30 Then
        BallShadowA(b).height = gBOT(b).z - BallSize / 4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      Else
        BallShadowA(b).height = 0.1
      End If
      BallShadowA(b).Y = gBOT(b).Y + offsetY
      BallShadowA(b).X = gBOT(b).X + offsetX
      BallShadowA(b).visible = 1
    End If
  Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************



'******************************************************
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'  Metals (all metal objects, metal walls, metal posts, metal wire guides)
'  Apron (the apron walls and plunger wall)
'  Walls (all wood or plastic walls)
'  Rollovers (wire rollover triggers, star triggers, or button triggers)
'  Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'  Gates (plate gates)
'  GatesWire (wire gates)
'  Rubbers (all rubbers including posts, sleeves, pegs, and bands)
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
' Audio : Adding Fleep Part 1       https://youtu.be/rG35JVHxtx4
' Audio : Adding Fleep Part 2       https://youtu.be/dk110pWMxGo
' Audio : Adding Fleep Part 3       https://youtu.be/ESXWGJZY_EI


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1        'volume level; range [0, 1]
NudgeRightSoundLevel = 1        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1        'volume level; range [0, 1]
StartButtonSoundLevel = 0.1      'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1        'volume level; range [0, 1]
RollingSoundFactor = 1.1 / 5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010    'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635    'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0            'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45          'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel    'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel    'sound helper; not configurable
SlingshotSoundLevel = 0.95            'volume level; range [0, 1]
BumperSoundFactor = 4.25            'volume multiplier; must not be zero
KnockerSoundLevel = 1              'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2      'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5      'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.375 / 5      'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025      'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025      'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8    'volume level; range [0, 1]
WallImpactSoundFactor = 0.075          'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075 / 3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5      'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10  'volume multiplier; must not be zero
DTSoundLevel = 0.25        'volume multiplier; must not be zero
RolloverSoundLevel = 0.25      'volume level; range [0, 1]
SpinnerSoundLevel = 0.5      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8          'volume level; range [0, 1]
BallReleaseSoundLevel = 1        'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015  'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025 / 5      'volume multiplier; must not be zero

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
  PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
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
  PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
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
  PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Table1.Height - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

' Thalamus, AudioFade patched
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 5) 'was 10
	Else
		AudioFade = CSng( - (( - tmp) ^ 5) ) 'was 10
	End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Table1.Width - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

' Thalamus, AudioPan patched
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 5) 'was 10
	Else
		AudioPan = CSng( - (( - tmp) ^ 5) ) 'was 10
	End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = CSng(BallVel(ball) ^ 2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = CSng((ball.velz) ^ 2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
End Function

Function RndInt(min, max) ' Sets a random number integer between min and max
  RndInt = Int(Rnd() * (max - min + 1) + min)
End Function

Function RndNum(min, max) ' Sets a random number between min and max
  RndNum = Rnd() * (max - min) + min
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////

Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * VolumeDial, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
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

Sub SoundKicker()
  PlaySoundAtLevelStatic ("Kickback_1"), PlungerPullSoundLevel, PlungerKB
End Sub

'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////

Sub KnockerSolenoid()
  PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 8) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////

Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////

Sub RandomSoundFlipperUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm / 10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm / 10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd * 4) + 1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////

Sub Rubbers_Hit(idx)
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 5 Then
    RandomSoundRubberStrong 1
  End If
  If finalspeed <= 5 Then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////

Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd * 10) + 1
    Case 1
      PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 2
      PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 3
      PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 4
      PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 5
      PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 6
      PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 7
      PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 8
      PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 9
      PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 10
      PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd * 9) + 1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
  RandomSoundWall()
End Sub

Sub RandomSoundWall()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 5) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 4) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 3) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

Sub RandomSoundMetal()
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ActiveBall) * MetalImpactSoundFactor
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
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Bounce_" & Int(Rnd * 2) + 1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd * 3) + 1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(ActiveBall.id) < 4) And cor.ballvely(ActiveBall.id) > 7 Then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd * 3) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd * 7) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

Sub RandomSoundTargetHitStrong()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 10 Then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft ActiveBall
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////

Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd * 9) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd * 7) + 1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////

Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd * 5) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, ActiveBall
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub Arch1_hit()
  If ActiveBall.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If ActiveBall.velx <  - 8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If ActiveBall.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If ActiveBall.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0
      PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1
      PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05    'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025 * RelayGISoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025 * RelayGISoundLevel, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025 * RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025 * RelayFlashSoundLevel, obj
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************


'********************************************
' Ball Collision, spinner collision and Sound
'********************************************

Sub OnBallBallCollision(ball1, ball2, velocity)

  FlipperCradleCollision ball1, ball2, velocity

  dim collAngle,bvelx,bvely,hitball, whichBall
  If ball1.radius < 23 or ball2.radius < 23 then

    If ball1.radius < 23 Then
      collAngle = GetCollisionAngle(ball1.x,ball1.y,ball2.x,ball2.y)
      set hitball = ball2
      If ball1.x = SpinnerBall1.x and ball1.y = SpinnerBall1.y Then
        whichball = 1
      Else
        whichball = 2
      End If
    else
      collAngle = GetCollisionAngle(ball2.x,ball2.y,ball1.x,ball1.y)
      set hitball = ball1
      If ball2.x = SpinnerBall1.x and ball2.y = SpinnerBall1.y Then
        whichball = 1
      Else
        whichball = 2
      End If
    End If

    dim discAngle

    If whichBall = 1 Then
      discAngle = NormAngle(spinAngle)
    Else
      discAngle = NormAngle(spinAngle2)
    End If

'   discSpinSpeed = discspinspeed + ecvel(0,1.5,sin(collAngle - discAngle)*velocity,BallMass * ABS(sin(collAngle - discAngle))) * cDiscSpeedMult

'   PlaySound "fx_lamphit", 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 1, 0, AudioFade(ball1)

    dim sineOfAngle, sineOfAngleSqr
    sineOfAngle = sin(collAngle - discAngle)

    discSpinSpeed = discspinspeed + ecvel(0,1.5,sineOfAngle*velocity,BallMass) * cDiscSpeedMult

    PlaySound "fx_lamphit", 0, Csng(velocity) ^2 / 2000 / 3, AudioPan(ball1), 0, Pitch(ball1), 1, 0, AudioFade(ball1)


  Else
    If ball1.z > 10 and ball2.z > 10 Then
'     PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
      '--- From Fleep code
      Dim snd : snd = "Ball_Collide_" & (Int(Rnd * 7) + 1)
      PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
      '--- End Fleep code
    End If

    'Newton ball code
    If (ball1.id = NewtonBall.id) and (ball2.id <> CaptiveBall.id) and (ball2.vely < 0) then KickCapBall ball2,velocity
    If (ball2.id = NewtonBall.id) and (ball1.id <> CaptiveBall.id) and (ball1.vely < 0) then KickCapBall ball1,velocity

  End If
End Sub

Const CBAngle = 1.5   '1.5 radians is the direction the captive ball can move
Sub KickCapBall(capball,velocity)
  dim angle, vel
  angle = CBAngle - GetCollisionAngle(NewtonBall.x, NewtonBall.y, capball.x, capball.x)
  vel = 0.8*velocity*cos(angle)
  CaptiveBall.velx = vel*cos(CBAngle)
  CaptiveBall.vely = -vel*sin(CBAngle)
End Sub

Function GetCollisionAngle(ax, ay, bx, by)
  Dim ang
  Dim collisionV:Set collisionV = new jVector
  collisionV.SetXY ax - bx, ay - by
  GetCollisionAngle = collisionV.ang
End Function

Function NormAngle(angle)
  NormAngle = angle
''  Dim pi:pi = 3.14159265358979323846
  Do While NormAngle>2 * pi
    NormAngle = NormAngle - 2 * pi
  Loop
  Do While NormAngle <0
    NormAngle = NormAngle + 2 * pi
  Loop
End Function

Class jVector
     Private m_mag, m_ang, pi

     Sub Class_Initialize
         m_mag = CDbl(0)
         m_ang = CDbl(0)
         pi = CDbl(3.14159265358979323846)
     End Sub

     Public Function add(anothervector)
         Dim tx, ty, theta
         If TypeName(anothervector) = "jVector" then
             Set add = new jVector
             add.SetXY x + anothervector.x, y + anothervector.y
         End If
     End Function

     Public Function multiply(scalar)
         Set multiply = new jVector
         multiply.SetXY x * scalar, y * scalar
     End Function

     Sub ShiftAxes(theta)
         ang = ang - theta
     end Sub

     Sub SetXY(tx, ty)

         if tx = 0 And ty = 0 Then
             ang = 0
          elseif tx = 0 And ty <0 then
             ang = - pi / 180 ' -90 degrees
          elseif tx = 0 And ty>0 then
             ang = pi / 180   ' 90 degrees
         else
             ang = atn(ty / tx)
             if tx <0 then ang = ang + pi ' Add 180 deg if in quadrant 2 or 3
         End if

         mag = sqr(tx ^2 + ty ^2)
     End Sub

     Property Let mag(nmag)
         m_mag = nmag
     End Property

     Property Get mag
         mag = m_mag
     End Property

     Property Let ang(nang)
         m_ang = nang
         Do While m_ang>2 * pi
             m_ang = m_ang - 2 * pi
         Loop
         Do While m_ang <0
             m_ang = m_ang + 2 * pi
         Loop
     End Property

     Property Get ang
         Do While m_ang>2 * pi
             m_ang = m_ang - 2 * pi
         Loop
         Do While m_ang <0
             m_ang = m_ang + 2 * pi
         Loop
         ang = m_ang
     End Property

     Property Get x
         x = m_mag * cos(ang)
     End Property

     Property Get y
         y = m_mag * sin(ang)
     End Property

     Property Get dump
         dump = "vector "
         Select Case CInt(ang + pi / 8)
             case 0, 8:dump = dump & "->"
             case 1:dump = dump & "/'"
             case 2:dump = dump & "/\"
             case 3:dump = dump & "'\"
             case 4:dump = dump & "<-"
             case 5:dump = dump & ":/"
             case 6:dump = dump & "\/"
             case 7:dump = dump & "\:"
         End Select

         dump = dump & " mag:" & CLng(mag * 10) / 10 & ", ang:" & CLng(ang * 180 / pi) & ", x:" & CLng(x * 10) / 10 & ", y:" & CLng(y * 10) / 10
     End Property
End Class

Function ECVel(Velocity1, Mass1, Velocity2, Mass2)
  ECVel = (Mass1 - Mass2)/(Mass1 + Mass2) * Velocity1  + 2 * Mass2/(Mass1 + Mass2)*Velocity2
End Function


'******************************************************
' BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  0.8       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)
Const PLOffset = 0.2
Dim PLGain: PLGain = (1-PLOffset)/(1260-2000)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 90 * BallBrightness + 65*LightLevel/100 + 100

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    d_w = b_base*(1 - (gBOT(s).z-25)/500)
    If d_w < 30 Then d_w = 30
    ' Handle plunger lane
    If InRect(gBOT(s).x,gBOT(s).y,870,2000,870,1260,930,1260,930,2000) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-2000))
    End If
    ' Assign color
    b_r = Int(d_w)
    b_g = Int(d_w)
    b_b = Int(d_w)
    If b_r > 255 Then b_r = 255
    If b_g > 255 Then b_g = 255
    If b_b > 255 Then b_b = 255
    gBOT(s).color = b_r + (b_g * 256) + (b_b * 256 * 256)
    'debug.print "--- ball.color level="&b_r
  Next
End Sub

' Make bubbles random speed
Sub UpdateBeerBubbles(elapsed)
  Const BubbleSpeed = 11.0
  VRBeerBubble1.z = VRBeerBubble1.z + Rnd(1)*0.5/BubbleSpeed
  if VRBeerBubble1.z > -771 then VRBeerBubble1.z = -955
  VRBeerBubble2.z = VRBeerBubble2.z + Rnd(1)*1/BubbleSpeed
  if VRBeerBubble2.z > -768 then VRBeerBubble2.z = -955
  VRBeerBubble3.z = VRBeerBubble3.z + Rnd(1)*1/BubbleSpeed
  if VRBeerBubble3.z > -768 then VRBeerBubble3.z = -955
  VRBeerBubble4.z = VRBeerBubble4.z + Rnd(1)*0.75/BubbleSpeed
  if VRBeerBubble4.z > -774 then VRBeerBubble4.z = -955
  VRBeerBubble5.z = VRBeerBubble5.z + Rnd(1)*1/BubbleSpeed
  if VRBeerBubble5.z > -771 then VRBeerBubble5.z = -955
  VRBeerBubble6.z = VRBeerBubble6.z + Rnd(1)*1/BubbleSpeed
  if VRBeerBubble6.z > -774 then VRBeerBubble6.z = -955
  VRBeerBubble7.z = VRBeerBubble7.z + Rnd(1)*0.8/BubbleSpeed
  if VRBeerBubble7.z > -768 then VRBeerBubble7.z = -955
  VRBeerBubble8.z = VRBeerBubble8.z + Rnd(1)*1/BubbleSpeed
  if VRBeerBubble8.z > -771 then VRBeerBubble8.z = -955
End Sub

'*****************************************************************************************************
' VR Lava Lamp
'*****************************************************************************************************
Dim Lspeed(7), Lbob(7), blobAng(7), blobRad(7), blobSiz(7), Blob, Bcnt   ' for VR Lava Lamp

' Lava Lamp code below.  Thank you STEELY!
Bcnt = 0

For Each Blob in Lspeed
  Lbob(Bcnt) = .2
  Lspeed(Bcnt) = Int((5 * Rnd) + 2) * .05
  Bcnt = Bcnt + 1
Next

Sub UpdateLavaLamp(elapsed)
  Bcnt = 0
  For Each Blob in VRLavaLamp
    If Blob.TransZ <= VRLavaLampBase.Size_Z * 1.5 Then  'Change blob direction to up
      Lspeed(Bcnt) = Int((5 * Rnd) + 2) * .05   'travel speed
      blobAng(Bcnt) = Int((359 * Rnd) + 1)      'blob location/angle from center
      blobRad(Bcnt) = Int((40 * Rnd) + 10)      'blob distance from center, radius
      blobSiz(Bcnt) = Int((150 * Rnd) + 100)    'blob size
      Blob.Size_x = blobSiz(Bcnt):Blob.Size_y = blobSiz(Bcnt):Blob.Size_z = blobSiz(Bcnt)
      Blob.X = Round(Cos(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaLampBase.X  'place blob
      Blob.Y = Round(Sin(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaLampBase.Y
    End If
    If Blob.TransZ => VRLavaLampBase.Size_Z*5 Then    'Change blob direction to down
      blobAng(Bcnt) = Int((359 * Rnd) + 1)      'blob location/angle from center
      blobRad(Bcnt) = Int((40 * Rnd) + 10)      'blob distance from center,radius
      Blob.X = Round(Cos(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaLampBase.X  'place blob
      Blob.Y = Round(Sin(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaLampBase.Y
      Lspeed(Bcnt) = Int((8 * Rnd) + 5) * .05:Lspeed(Bcnt) = Lspeed(Bcnt) * -1        'travel speed
    End If

    ' Make blob wobble
    If Blob.Size_x > blobSiz(Bcnt) + blobSiz(Bcnt)*.15 or Blob.Size_x < blobSiz(Bcnt) - blobSiz(Bcnt)*.15  Then
      Lbob(Bcnt) = Lbob(Bcnt) * -1
    End If

    Blob.Size_x = Blob.Size_x + Lbob(Bcnt) * elapsed / 10.0
    Blob.Size_y = Blob.Size_y + Lbob(Bcnt) * elapsed / 10.0
    Blob.Size_z = Blob.Size_Z - Lbob(Bcnt) * .66 * elapsed / 10.0
    Blob.TransZ = Blob.TransZ + Lspeed(Bcnt) * elapsed / 10.0 'Move blob
    Bcnt = Bcnt + 1
  Next
End Sub

Sub ClockTimer_Timer()
  Dim t : t = Now()
  Pminutes.RotAndTra2 = (Minute(t)+(Second(t)/100))*6
  Phours.RotAndTra2 = Hour(t)*30+(Minute(t)/2)
  Pseconds.RotAndTra2 = (Second(t))*6
End Sub

'0.001 - Sixtoe - Built Table
'0.002 - Niwak - Add first bake, wire up rollover and flipper bats, add option menu
'0.003 - Sixtoe - Added playfield mesh, added physical kickers, newton ball code done for captive ball, tweaked table layout, added some missing sounds, changed flipper angles
'0.004 - Sixtoe - Fixed kickers (it was bugging me)
'0.005 - Niwak - New batch bake, fix options menu, implement solenoid gates (flap and wire), fix flasher, better bumper skirt anim, drop/stand target
'0.006 - Niwak - New batch bake, add sling animation, fix drop target axis, fix plastic lighting, remove unused second newton ball bake, fix center pole up/down and controlled gates, add outpost difficulty, initial buggy simplified physics for the center post
'0.007 - Sixtoe - Fixed pop bumper ball traps, filled in some potential balltrap holes,
'0.008 - apophis - Reworked center post mech. Updated kicker code. Adjusted flipper triggers. Adjusted upper right flipper physics params. Added phys mtl to inlane guides.
'0.009 - Sixtoe - Tweaked some walls, tweaked centre ball mass, tweaked kickback plunger, added VR cabinet and room.
'0.010 - Sixtoe - Something?
'0.011 - Niwak - Added the GI shadows (and checked the spherical map of the ball for the time being
'0.012 - Sixtoe - Even More stuff?
'0.013 - Sixtoe - Tuned centre colllision post, tuned flippers, adjusted kicker, adjusted a few physics objects, added apophis flipper disable code.
'0.014 - Sixtoe - Something else.
'0.015 - ClarkKent - Adjusted Flippers
'0.016 - Apophis - Fixed flipper bake/lightmap positions. Reverted to v13 UR flipper size. Adjusted some flipper physics params be more consistent with nFozzy standards. Animated center post hits. Added FlipperCradleCollision. Reduced nudge strength. Fixed sling arm animations. Added UpdateBallBrightness. Fixed post passes.
'0.017 - Niwak - Use latest PWM outputs from PinMame 3.6, removed EnableFlips
'0.018 - Niwak - Adjust for latest PinMame, change ball, adjust GI reflection, move options to TweakUI, adjusted reflections
'0.019 - Sixtoe - Added and hooked up shooter rod to VR room, fixed and repositioned pincab DMD, setlocale
'0.020 - TastyWasps - Animations for VR toys and flipper buttons.
'0.021 - Sixtoe - Fixed VR cabinet, adjusted flipper positions and flipper primitive positions.
'0.022 - Niwak - Fix ball reflections, fix invalid timers, add roughness to reflections
'0.023 - Sixtoe - Removed roughness for playfield reflections in VR, added new default LUT from HF, added image to plunger
'0.024 - Niwak - Add rule card to tweak menu
'0.025 - Niwak - Remove timers below framerate except for cor tracker (to be done later)
'0.026 - apophis - Same as 025 but file not corrupt (Added CorTimer and copied script from 025 into o24 and resaved)
'0.027 - Niwak - Only enabled VR timer when VR is used
'0.028 - Niwak - Clean ups to account for latest VPX changes
'0.029 - Niwak - New Blender 4 bake: better lights, some bulb with filaments near the center post, tinting for these bulbs, still needs some light balancing (white inserts too strong ?)
'0.030 - apophis - Set bloom strength to 0.05. Set "Hide parts behind" for BM_Playfield. Fixed flipper triggers. Updated flipper corrections and tricks.
'0.031 - Niwak - New Blender 4 bake with better insert balance
'0.032 - Niwak - Use tweening instead of timers, adjust center post height, more script cleanups
'0.033 - Niwak - Fix flipper shadow, experiment with using tweens for Rothbauer targets and bumpers
'0.034 - Niwak - Simplify script again
'0.035 - Niwak - New bake: adjust plain inserts again, adjust bulb height, keep more bulbs for VR, do not remove backface for VR
'0.036 - apophis - Adjusted center post up position. Lowered drop targets down position. Added correction to aBall.velz in dampener code. Updated gate physics params, flipper bumper and sling strengths, PF fric. Reduced ball reflection of PF. Updated DisableStaticPreRendering functionality to be consistent with VPX 10.8.1 API.
'0.037 - apophis - Script fix for standalone compatibility.
'RC1 - apophis - Added plunger line visualization option and Desktop DMD visibility option.
'RC2 - Sixtoe - Script tidied up
'Release 1.0
'1.0.1 - apophis - Attempt to prevent center post from dislodging and rolling down table. Bumper strength set to 11.
'1.0.2 - apophis - Better aligned flipper bakemaps and lightmaps to phyiscal flippers. Added ball image option.
'Release 1.1

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub
