'  +              `s
' yNhssyyyydddhhyhhNo
'+/`      VPW      `+:                    ....`                  ```  ::  .//+/:
'         hMm        : `:/++.    ```.`    :mMd`    :ddo-    `/yyooooyhmm   +MM`  -sNNh/ `          `
'         hMN `/odmh:`   :NMy   .Nm:..     hMd`     yM/   `oNd.       -s+  /MM`    NM: `NhhyyhhssshN.
'         dMd    oMm`     hMM-  .N:`:hmh`  hMh      sM+  `dMm.          .  /MM`    NM. ss`  oMh    :+
'         NMd     dMh`   -VPWo `d+   mMs   hMs      sMo  yMM-           ```:MM+//++MM. `    yMy
'         mMd     `mMh   ds+MN`sy    mMy   hMs      oMo `NMm      -.:yNMN:`/MM/-.`.MM.      yMy
'         NMm      `hMy oh  NMdd     hMo   dMo      yM: `MMd         `NMd  /MM.   `MM`      hMs
'         MMd       `hMdN.  +Mm`     dMs   dMo     .ss+` dMN.         NMd  -Mm`   .MM`      mMo
'        `MMo        `dM:    d-      mMs   NMs   `.:/    .dMN/      `+Ns.  :MM`   -MM.      NM+
'        .MMy         `:            :MMs  :yyssyyhdm.      /ymmsooooo-    -dNd:` .+oo+/    `MM+`.
'        oMMy                       ..``                                                  ./+/:--
'      -+oo+/:-     ``
'                  -o.                  -
'                 /Nmdddddysysoooo+yVPWy.   `-+s+:-  .-//+/     -+sss:           +`
'                +s/:-.`         :dMMs.  `+ho:-:/yMNo`  /MMy      mM.   +hdysossyms
'               `              /dMm+`   oMy`      .dMm. `VPWm`    oN    .NM:      .:
'                           `oNMd/     yMh         `MMy `N+hMN:   +N    .MM-       `
'                         `oNMh:      `NM:          mMy  N+ oMMo  :M    `NMdhhhhdmmy
'                       .smMh-        -MM/         `NN.  m+  :NMh`-M`   `NM+``   `:
'                     :hMMy.           yMd-       `hy.   N+   .hMm:M.   `NM.
'                   :dMMs.          .+/ +dMdo:-:/oo-    `M/     oMMM-   `NM:       -
'                 :hMMd///++osyhhdddMy    .:++/-`       oNy+:.   -mM:   `MM- .-/+ym.
'                /o+/::--.`````     o                  -/-.`       s/  -ohysoo+/:/-
'
' Twilight Zone VPW Edition
' Bally/Midway 1993
' https://www.ipdb.org/machine.cgi?id=2684
'
' VPW Anomalies
' =============
' Project Leads - Niwak & Sixtoe
' Niwak - Blender Toolkit, Scripting, Inventing.
' Sixtoe - VPX Rebuild, Script Stuff, VR Stuff, General Dogsbody.
' rothbauer - Physics and Guidance.
' apophis - Script Updates & Various Tweaks.
' leojreimrc - FSS Backglass.
' Darth Vito - 360 room.
' mcarter78 - 360 room 3D text, script tweaks.
' Flupper - Graphics Advice.
' nfozzy - physics advice and scripting.
' Testing - AstroNasty, DGrimmReaper, Cliffy, Primtime5k, Studlygoorite, Lumi, somatik, BountyBob, RetroRitchie, DarthVito, Dr.Nobody, Rigo, Pinballfan2018, DaRdog, Manu, Flux, geradg
'
' VPW Edition Based on;
' ---------------------
' Remaster by Skitso, rothbauerw and Bord.
' Which was based on;
' VPX recreation by Ninuzzu & Tom Tower
' Which was based on;
' FP conversion by Coindropper
'
' Previous Mods Credits/Thanks
' ----------------------------
' Tom Tower for the mini clock, pyramid, piano, camera, gumballmachine models.
' nFozzy for scripting complicated things.
' Fleep for the audio system.
' Hauntfreaks for the Mystic Seer Toy.
' rom for "Robbie the Robot" model.
' Flupper for retexturing and remeshing the plastic ramp.
' JPSalas for the help with the inserts and the script.
' Clark Kent for early graphics and physics work.
'
' Required Software (64-bit versions)
' -----------------------------------
' VPX 10.8.0 RC4 or later
' VPinMame 3.6.0-998 or later
'
' Changelog at end of script.
'
Option Explicit
Randomize
SetLocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim Romset, PowerballStart

'******************************************************************************************
'* TABLE OPTIONS **************************************************************************
'* Except options shown below, all other options are accessed by pressing F12 in game.    *
'******************************************************************************************
Romset = 1              ' Choose the ROM: 0 = home rom with credits (tz_94ch),  1 = home rom (free play) (tz_94h), 2 = last arcade rom version (tz_92)
PowerBallStart = 7      ' Powerball Starting Location: 1-3 = gumball machine, 4-6 = trough, 0 = random, 7 = random gumball

'******************************************************************************************
'* END OF TABLE OPTIONS *******************************************************************
'******************************************************************************************

'******************************************************
' Standard definitions
'******************************************************

Const tnob = 5                      'Total number of balls
Const lob = 0                       'Locked balls

Dim BIPL : BIPL = False             'Ball in plunger lane

Dim cGameName
Select Case Romset
    Case 0: cGameName = "tz_94ch" 'Coin Based Home Version (Advanced Features)
    Case 1: cGameName = "tz_94h"  'Free Play Home Version (Advance Features)
    Case 2: cGameName = "tz_92"   'Last Arcade Romset (Basic Features)
End Select
Const UseSolenoids = 2
Const UseVPMModSol = 2
Const UseLamps = 1
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = ""

Const VRDesktopSim = False

Dim DesktopMode: DesktopMode = Table1.ShowDT
Dim UseVPMDMD
If Desktopmode or RenderingMode = 2 Then UseVPMDMD = 1 Else UseVPMDMD = 0

Const BallSize = 50                 'Ball size must be 50
Const BallMass = 1                  'Ball mass must be 1

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

LoadVPM "03060000", "WPC.VBS", 3.49

Set GiCallback2 = GetRef("UpdateGI")
Sub UpdateGI(no, v) ' GI Strings
    Select Case no
        '#0 Left
        Case 0 : l100a.State = v : l100b.State = v : l100c.State = v : l100d.State = v : l100e.State = v : l100.State = v : l100_1.State = v : l100_2.State = v : l100_3.State = v : l100_4.State = v : l100_5.State = v
        '#1 Mini PF
        Case 1 : l101.State = v
        '#2 Clock
        Case 2 : l102.State = v
        '#3 Inserts modulation ?
        Case 3 :
        '#4 Right
        Case 4 : l104a.State = v : l104b.State = v : l104c.State = v : l104d.State = v : l104.State = v : l104_1.State = v : l104_2.State = v : l104_3.State = v : l104_4.State = v : l104_5.State = v : l104_6.State = v : l104_7.State = v : l104_8.State = v
    End Select
End Sub

Const TW_SHOOTER_DIV = 0
Const TW_GUMBALL_DIV = 1
Const TW_LEFT_RAMP_DIV = 2
Const TW_RIGHT_RAMP_DIV = 3
Const NTweens = 4

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
Sub PrimArrayRotZUpdate(id, prims, v) : Dim x: For Each x in prims: x.RotZ = v: Next: End Sub

'*******************************************
'  Blender Bakes
'*******************************************

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BR1: BP_BR1=Array(BM_BR1, LM_Mod_l111_BR1, LM_Mod_l109_BR1, LM_Flashers_f17_BR1, LM_GI_Left_l100_BR1, LM_Bumpers_l61_BR1, LM_Bumpers_l62_BR1, LM_Bumpers_l63_BR1)
Dim BP_BR2: BP_BR2=Array(BM_BR2, LM_Mod_l110_BR2, LM_Mod_l109_BR2, LM_Flashers_f17_BR2, LM_Flashers_f19_BR2, LM_Flashers_f28_BR2, LM_GI_Left_l100_BR2, LM_GI_Left_l100b_BR2, LM_GI_Right_l104_BR2, LM_GI_Right_l104c_BR2, LM_Inserts_l17_BR2, LM_Bumpers_l61_BR2, LM_Bumpers_l62_BR2, LM_Bumpers_l63_BR2, LM_Inserts_l64_BR2, LM_Inserts_l65_BR2)
Dim BP_BR3: BP_BR3=Array(BM_BR3, LM_Mod_l109_BR3, LM_Flashers_f17_BR3, LM_Flashers_f19_BR3, LM_GI_Left_l100c_BR3, LM_GI_Left_l100e_BR3, LM_GI_Right_l104a_BR3, LM_GI_Right_l104c_BR3, LM_Inserts_l37_BR3, LM_Inserts_l38_BR3, LM_Bumpers_l61_BR3, LM_Bumpers_l62_BR3, LM_Bumpers_l63_BR3)
Dim BP_BS1: BP_BS1=Array(BM_BS1, LM_Flashers_f17_BS1, LM_Flashers_f38_BS1, LM_GI_Left_l100_BS1, LM_GI_Right_l104_BS1, LM_Bumpers_l61_BS1, LM_Bumpers_l62_BS1, LM_Bumpers_l63_BS1)
Dim BP_BS2: BP_BS2=Array(BM_BS2, LM_Mod_l110_BS2, LM_Mod_l109_BS2, LM_Flashers_f17_BS2, LM_Flashers_f20_BS2, LM_GI_Left_l100_BS2, LM_Inserts_l17_BS2, LM_Bumpers_l61_BS2, LM_Bumpers_l62_BS2, LM_Bumpers_l63_BS2, LM_Inserts_l64_BS2, LM_Inserts_l65_BS2)
Dim BP_BS3: BP_BS3=Array(BM_BS3, LM_Flashers_f17_BS3, LM_Flashers_f19_BS3, LM_GI_Left_l100_BS3, LM_GI_Left_l100b_BS3, LM_Inserts_l37_BS3, LM_Inserts_l38_BS3, LM_Bumpers_l61_BS3, LM_Bumpers_l62_BS3, LM_Bumpers_l63_BS3)
Dim BP_B_Caps_Bot_Amber: BP_B_Caps_Bot_Amber=Array(BM_B_Caps_Bot_Amber, LM_Mod_l110_B_Caps_Bot_Amber, LM_Mod_l109_B_Caps_Bot_Amber, LM_Flashers_f17_B_Caps_Bot_Ambe, LM_Flashers_f18_B_Caps_Bot_Ambe, LM_Flashers_f20_B_Caps_Bot_Ambe, LM_Flashers_f28_B_Caps_Bot_Ambe, LM_Flashers_f37_B_Caps_Bot_Ambe, LM_Inserts_l17_B_Caps_Bot_Amber, LM_Inserts_l52_B_Caps_Bot_Amber, LM_Inserts_l53_B_Caps_Bot_Amber, LM_Inserts_l56_B_Caps_Bot_Amber, LM_Bumpers_l61_B_Caps_Bot_Amber, LM_Bumpers_l63_B_Caps_Bot_Amber, LM_Inserts_l73_B_Caps_Bot_Amber, LM_Inserts_l75_B_Caps_Bot_Amber, LM_Inserts_l81_B_Caps_Bot_Amber)
Dim BP_B_Caps_Bot_Red: BP_B_Caps_Bot_Red=Array(BM_B_Caps_Bot_Red, LM_GI_Clock_l102_B_Caps_Bot_Red, LM_Mod_l109_B_Caps_Bot_Red, LM_Flashers_f17_B_Caps_Bot_Red, LM_GI_Left_l100_B_Caps_Bot_Red, LM_Bumpers_l61_B_Caps_Bot_Red, LM_Bumpers_l63_B_Caps_Bot_Red, LM_Inserts_l64_B_Caps_Bot_Red)
Dim BP_B_Caps_Bot_Yellow: BP_B_Caps_Bot_Yellow=Array(BM_B_Caps_Bot_Yellow, LM_Mod_l109_B_Caps_Bot_Yellow, LM_Flashers_f17_B_Caps_Bot_Yell, LM_Flashers_f18_B_Caps_Bot_Yell, LM_Bumpers_l61_B_Caps_Bot_Yello, LM_Bumpers_l62_B_Caps_Bot_Yello, LM_Bumpers_l63_B_Caps_Bot_Yello)
Dim BP_B_Caps_Top_Amber: BP_B_Caps_Top_Amber=Array(BM_B_Caps_Top_Amber, LM_GI_Clock_l102_B_Caps_Top_Amb, LM_Mod_l109_B_Caps_Top_Amber, LM_Flashers_f17_B_Caps_Top_Ambe, LM_Flashers_f28_B_Caps_Top_Ambe, LM_Flashers_f37_B_Caps_Top_Ambe, LM_GI_Left_l100_B_Caps_Top_Ambe, LM_GI_Left_l100b_B_Caps_Top_Amb, LM_GI_Left_l100c_B_Caps_Top_Amb, LM_GI_Left_l100d_B_Caps_Top_Amb, LM_GI_Right_l104_B_Caps_Top_Amb, LM_GI_Right_l104c_B_Caps_Top_Am, LM_Inserts_l13_B_Caps_Top_Amber, LM_Inserts_l16_B_Caps_Top_Amber, LM_Inserts_l17_B_Caps_Top_Amber, LM_Inserts_l18_B_Caps_Top_Amber, LM_Inserts_l24_B_Caps_Top_Amber, LM_Inserts_l25_B_Caps_Top_Amber, LM_Inserts_l26_B_Caps_Top_Amber, LM_Inserts_l27_B_Caps_Top_Amber, LM_Inserts_l28_B_Caps_Top_Amber, LM_Inserts_l32_B_Caps_Top_Amber, LM_Inserts_l38_B_Caps_Top_Amber, LM_Inserts_l51_B_Caps_Top_Amber, LM_Inserts_l52_B_Caps_Top_Amber, LM_Inserts_l53_B_Caps_Top_Amber, LM_Inserts_l55_B_Caps_Top_Amber, LM_Bumpers_l61_B_Caps_Top_Amber, LM_Bumpers_l62_B_Caps_Top_Amber, _
  LM_Bumpers_l63_B_Caps_Top_Amber, LM_Inserts_l64_B_Caps_Top_Amber, LM_Inserts_l65_B_Caps_Top_Amber, LM_Inserts_l71_B_Caps_Top_Amber, LM_Inserts_l72_B_Caps_Top_Amber, LM_Inserts_l73_B_Caps_Top_Amber, LM_Inserts_l75_B_Caps_Top_Amber, LM_Inserts_l82_B_Caps_Top_Amber)
Dim BP_B_Caps_Top_Red: BP_B_Caps_Top_Red=Array(BM_B_Caps_Top_Red, LM_Flashers_f17_B_Caps_Top_Red, LM_Flashers_f18_B_Caps_Top_Red, LM_GI_Left_l100_B_Caps_Top_Red, LM_GI_Left_l100c_B_Caps_Top_Red, LM_GI_Right_l104a_B_Caps_Top_Re, LM_Inserts_l14_B_Caps_Top_Red, LM_Inserts_l16_B_Caps_Top_Red, LM_Inserts_l17_B_Caps_Top_Red, LM_Inserts_l26_B_Caps_Top_Red, LM_Inserts_l31_B_Caps_Top_Red, LM_Inserts_l34_B_Caps_Top_Red, LM_Inserts_l36_B_Caps_Top_Red, LM_Bumpers_l61_B_Caps_Top_Red, LM_Bumpers_l62_B_Caps_Top_Red, LM_Bumpers_l63_B_Caps_Top_Red, LM_Inserts_l65_B_Caps_Top_Red)
Dim BP_B_Caps_Top_Yellow: BP_B_Caps_Top_Yellow=Array(BM_B_Caps_Top_Yellow, LM_GI_Clock_l102_B_Caps_Top_Yel, LM_Mod_l109_B_Caps_Top_Yellow, LM_Flashers_f17_B_Caps_Top_Yell, LM_Flashers_f18_B_Caps_Top_Yell, LM_Flashers_f28_B_Caps_Top_Yell, LM_Flashers_f37_B_Caps_Top_Yell, LM_GI_Left_l100a_B_Caps_Top_Yel, LM_GI_Left_l100c_B_Caps_Top_Yel, LM_GI_Left_l100d_B_Caps_Top_Yel, LM_GI_Left_l100e_B_Caps_Top_Yel, LM_GI_Right_l104a_B_Caps_Top_Ye, LM_GI_Right_l104c_B_Caps_Top_Ye, LM_Inserts_l11_B_Caps_Top_Yello, LM_Inserts_l12_B_Caps_Top_Yello, LM_Inserts_l13_B_Caps_Top_Yello, LM_Inserts_l14_B_Caps_Top_Yello, LM_Inserts_l15_B_Caps_Top_Yello, LM_Inserts_l16_B_Caps_Top_Yello, LM_Inserts_l17_B_Caps_Top_Yello, LM_Inserts_l21_B_Caps_Top_Yello, LM_Inserts_l22_B_Caps_Top_Yello, LM_Inserts_l23_B_Caps_Top_Yello, LM_Inserts_l24_B_Caps_Top_Yello, LM_Inserts_l25_B_Caps_Top_Yello, LM_Inserts_l28_B_Caps_Top_Yello, LM_Inserts_l31_B_Caps_Top_Yello, LM_Inserts_l32_B_Caps_Top_Yello, LM_Inserts_l33_B_Caps_Top_Yello, _
  LM_Inserts_l34_B_Caps_Top_Yello, LM_Inserts_l35_B_Caps_Top_Yello, LM_Inserts_l36_B_Caps_Top_Yello, LM_Inserts_l37_B_Caps_Top_Yello, LM_Inserts_l42_B_Caps_Top_Yello, LM_Inserts_l43_B_Caps_Top_Yello, LM_Inserts_l44_B_Caps_Top_Yello, LM_Inserts_l45_B_Caps_Top_Yello, LM_Bumpers_l61_B_Caps_Top_Yello, LM_Bumpers_l62_B_Caps_Top_Yello, LM_Bumpers_l63_B_Caps_Top_Yello, LM_Inserts_l64_B_Caps_Top_Yello, LM_Inserts_l65_B_Caps_Top_Yello, LM_Inserts_l71_B_Caps_Top_Yello)
Dim BP_BumperPegs: BP_BumperPegs=Array(BM_BumperPegs, LM_Mod_l109_BumperPegs, LM_Flashers_f17_BumperPegs, LM_GI_Left_l100_BumperPegs, LM_Inserts_l38_BumperPegs, LM_Bumpers_l61_BumperPegs, LM_Bumpers_l62_BumperPegs, LM_Bumpers_l63_BumperPegs, LM_Inserts_l64_BumperPegs)
Dim BP_Camera: BP_Camera=Array(BM_Camera, LM_Mod_l110_Camera, LM_Flashers_f17_Camera, LM_Flashers_f19_Camera, LM_Flashers_f20_Camera, LM_Flashers_f28_Camera, LM_Bumpers_l63_Camera)
Dim BP_ClockLarge: BP_ClockLarge=Array(BM_ClockLarge, LM_GI_Clock_l102_ClockLarge, LM_Flashers_f19_ClockLarge, LM_Flashers_f28_ClockLarge, LM_Inserts_l84_ClockLarge)
Dim BP_ClockShort: BP_ClockShort=Array(BM_ClockShort, LM_GI_Clock_l102_ClockShort, LM_Flashers_l74_ClockShort, LM_Inserts_l84_ClockShort)
Dim BP_ClockToy: BP_ClockToy=Array(BM_ClockToy, LM_GI_Clock_l102_ClockToy, LM_Flashers_f17_ClockToy, LM_Flashers_f18_ClockToy, LM_Flashers_l74_ClockToy, LM_Flashers_f19_ClockToy, LM_Flashers_f20_ClockToy, LM_Flashers_f28_ClockToy, LM_Flashers_f37_ClockToy, LM_Flashers_f38_ClockToy, LM_Flashers_f39_ClockToy, LM_Flashers_f40_ClockToy, LM_Flashers_f41_ClockToy, LM_GI_Left_l100_ClockToy, LM_GI_Left_l100a_ClockToy, LM_GI_Left_l100b_ClockToy, LM_GI_Right_l104_ClockToy, LM_GI_Right_l104a_ClockToy, LM_GI_Right_l104b_ClockToy, LM_Inserts_l51_ClockToy, LM_Inserts_l52_ClockToy, LM_Inserts_l53_ClockToy, LM_Inserts_l54_ClockToy, LM_Inserts_l55_ClockToy, LM_Inserts_l56_ClockToy, LM_Bumpers_l61_ClockToy, LM_Inserts_l72_ClockToy, LM_Inserts_l73_ClockToy, LM_Inserts_l81_ClockToy, LM_Inserts_l82_ClockToy, LM_Inserts_l83_ClockToy, LM_Inserts_l84_ClockToy)
Dim BP_Clock_Color: BP_Clock_Color=Array(BM_Clock_Color, LM_GI_Clock_l102_Clock_Color, LM_Flashers_f17_Clock_Color, LM_Flashers_f18_Clock_Color, LM_Flashers_l74_Clock_Color, LM_Flashers_f19_Clock_Color, LM_Flashers_f20_Clock_Color, LM_Flashers_f28_Clock_Color, LM_Flashers_f37_Clock_Color, LM_Flashers_f38_Clock_Color, LM_Flashers_f41_Clock_Color, LM_GI_Right_l104_Clock_Color, LM_Inserts_l51_Clock_Color, LM_Inserts_l52_Clock_Color, LM_Inserts_l53_Clock_Color, LM_Inserts_l55_Clock_Color, LM_Inserts_l56_Clock_Color, LM_Inserts_l58_Clock_Color, LM_Bumpers_l63_Clock_Color, LM_Inserts_l67_Clock_Color, LM_Inserts_l68_Clock_Color, LM_Inserts_l72_Clock_Color, LM_Inserts_l73_Clock_Color, LM_Inserts_l75_Clock_Color, LM_Inserts_l82_Clock_Color, LM_Inserts_l83_Clock_Color, LM_Inserts_l84_Clock_Color, LM_Inserts_l85_Clock_Color)
Dim BP_Clock_White: BP_Clock_White=Array(BM_Clock_White, LM_GI_Clock_l102_Clock_White, LM_Flashers_f17_Clock_White, LM_Flashers_f18_Clock_White, LM_Flashers_l74_Clock_White, LM_Flashers_f19_Clock_White, LM_Flashers_f20_Clock_White, LM_Flashers_f28_Clock_White, LM_Flashers_f37_Clock_White, LM_Flashers_f38_Clock_White, LM_Flashers_f39_Clock_White, LM_Flashers_f41_Clock_White, LM_GI_Right_l104_Clock_White, LM_Inserts_l51_Clock_White, LM_Inserts_l52_Clock_White, LM_Inserts_l53_Clock_White, LM_Inserts_l54_Clock_White, LM_Inserts_l55_Clock_White, LM_Inserts_l56_Clock_White, LM_Inserts_l58_Clock_White, LM_Bumpers_l63_Clock_White, LM_Inserts_l67_Clock_White, LM_Inserts_l68_Clock_White, LM_Inserts_l72_Clock_White, LM_Inserts_l73_Clock_White, LM_Inserts_l75_Clock_White, LM_Inserts_l82_Clock_White, LM_Inserts_l83_Clock_White, LM_Inserts_l84_Clock_White, LM_Inserts_l85_Clock_White)
Dim BP_DiverterP: BP_DiverterP=Array(BM_DiverterP, LM_GI_Clock_l102_DiverterP, LM_Flashers_f18_DiverterP, LM_Flashers_f28_DiverterP, LM_Flashers_f41_DiverterP, LM_GI_Right_l104_DiverterP)
Dim BP_DiverterP1: BP_DiverterP1=Array(BM_DiverterP1, LM_Flashers_f20_DiverterP1, LM_Flashers_f28_DiverterP1, LM_Flashers_f41_DiverterP1)
Dim BP_FlipperL: BP_FlipperL=Array(BM_FlipperL, LM_Flashers_f18_FlipperL, LM_Flashers_f20_FlipperL, LM_Flashers_f28_FlipperL, LM_GI_Left_l100a_FlipperL, LM_GI_Left_l100b_FlipperL, LM_GI_Left_l100e_FlipperL, LM_GI_Right_l104a_FlipperL, LM_GI_Right_l104c_FlipperL, LM_GI_Right_l104d_FlipperL, LM_Inserts_l41_FlipperL, LM_Inserts_l42_FlipperL, LM_Inserts_l43_FlipperL, LM_Inserts_l44_FlipperL, LM_Inserts_l45_FlipperL, LM_Inserts_l47_FlipperL)
Dim BP_FlipperL1: BP_FlipperL1=Array(BM_FlipperL1, LM_GI_Clock_l102_FlipperL1, LM_Flashers_f18_FlipperL1, LM_Flashers_f19_FlipperL1, LM_Flashers_f20_FlipperL1, LM_Flashers_f28_FlipperL1, LM_Flashers_f38_FlipperL1, LM_Flashers_f39_FlipperL1, LM_Flashers_f40_FlipperL1, LM_Flashers_f41_FlipperL1, LM_GI_Right_l104_FlipperL1, LM_Inserts_l52_FlipperL1, LM_Inserts_l53_FlipperL1, LM_Inserts_l55_FlipperL1, LM_Bumpers_l63_FlipperL1, LM_Inserts_l81_FlipperL1)
Dim BP_FlipperLU: BP_FlipperLU=Array(BM_FlipperLU, LM_Flashers_f17_FlipperLU, LM_Flashers_f18_FlipperLU, LM_Flashers_f20_FlipperLU, LM_Flashers_f28_FlipperLU, LM_GI_Left_l100a_FlipperLU, LM_GI_Left_l100b_FlipperLU, LM_GI_Left_l100c_FlipperLU, LM_GI_Left_l100d_FlipperLU, LM_GI_Left_l100e_FlipperLU, LM_GI_Right_l104a_FlipperLU, LM_GI_Right_l104c_FlipperLU, LM_GI_Right_l104d_FlipperLU, LM_Inserts_l11_FlipperLU, LM_Inserts_l12_FlipperLU, LM_Inserts_l41_FlipperLU, LM_Inserts_l42_FlipperLU, LM_Inserts_l43_FlipperLU, LM_Inserts_l44_FlipperLU, LM_Inserts_l45_FlipperLU, LM_Inserts_l47_FlipperLU)
Dim BP_FlipperR: BP_FlipperR=Array(BM_FlipperR, LM_Flashers_f20_FlipperR, LM_Flashers_f28_FlipperR, LM_Flashers_f39_FlipperR, LM_GI_Left_l100a_FlipperR, LM_GI_Left_l100b_FlipperR, LM_GI_Left_l100c_FlipperR, LM_GI_Left_l100d_FlipperR, LM_GI_Left_l100e_FlipperR, LM_GI_Right_l104a_FlipperR, LM_GI_Right_l104c_FlipperR, LM_GI_Right_l104d_FlipperR, LM_Inserts_l42_FlipperR, LM_Inserts_l43_FlipperR, LM_Inserts_l44_FlipperR, LM_Inserts_l45_FlipperR, LM_Inserts_l46_FlipperR, LM_Inserts_l47_FlipperR)
Dim BP_FlipperR1: BP_FlipperR1=Array(BM_FlipperR1, LM_Mod_l106_FlipperR1, LM_Mod_l107_FlipperR1, LM_Mod_l111_FlipperR1, LM_Flashers_f17_FlipperR1, LM_Flashers_f18_FlipperR1, LM_Flashers_f19_FlipperR1, LM_Flashers_f20_FlipperR1, LM_Flashers_f28_FlipperR1, LM_Flashers_f37_FlipperR1, LM_Flashers_f38_FlipperR1, LM_Flashers_f39_FlipperR1, LM_Flashers_f41_FlipperR1, LM_GI_Left_l100a_FlipperR1, LM_GI_Right_l104_FlipperR1, LM_Inserts_l25_FlipperR1, LM_Inserts_l86_FlipperR1)
Dim BP_FlipperR1U: BP_FlipperR1U=Array(BM_FlipperR1U, LM_GI_Clock_l102_FlipperR1U, LM_Mod_l106_FlipperR1U, LM_Mod_l107_FlipperR1U, LM_Mod_l111_FlipperR1U, LM_Flashers_f17_FlipperR1U, LM_Flashers_f18_FlipperR1U, LM_Flashers_f19_FlipperR1U, LM_Flashers_f20_FlipperR1U, LM_Flashers_f28_FlipperR1U, LM_Flashers_f37_FlipperR1U, LM_Flashers_f38_FlipperR1U, LM_Flashers_f39_FlipperR1U, LM_Flashers_f41_FlipperR1U, LM_GI_Left_l100_FlipperR1U, LM_GI_Left_l100b_FlipperR1U, LM_GI_Right_l104_FlipperR1U, LM_GI_Right_l104a_FlipperR1U, LM_GI_Right_l104d_FlipperR1U, LM_Inserts_l25_FlipperR1U, LM_Inserts_l26_FlipperR1U, LM_Inserts_l27_FlipperR1U, LM_Inserts_l48_FlipperR1U, LM_Inserts_l71_FlipperR1U, LM_Inserts_l85_FlipperR1U, LM_Inserts_l86_FlipperR1U)
Dim BP_FlipperRU: BP_FlipperRU=Array(BM_FlipperRU, LM_Flashers_f17_FlipperRU, LM_Flashers_f18_FlipperRU, LM_Flashers_f20_FlipperRU, LM_Flashers_f28_FlipperRU, LM_Flashers_f39_FlipperRU, LM_Flashers_f41_FlipperRU, LM_GI_Left_l100a_FlipperRU, LM_GI_Left_l100d_FlipperRU, LM_GI_Left_l100e_FlipperRU, LM_GI_Right_l104a_FlipperRU, LM_GI_Right_l104c_FlipperRU, LM_GI_Right_l104d_FlipperRU, LM_Inserts_l11_FlipperRU, LM_Inserts_l22_FlipperRU, LM_Inserts_l42_FlipperRU, LM_Inserts_l43_FlipperRU, LM_Inserts_l44_FlipperRU, LM_Inserts_l45_FlipperRU, LM_Inserts_l46_FlipperRU, LM_Inserts_l47_FlipperRU)
Dim BP_FlipperSpL: BP_FlipperSpL=Array(BM_FlipperSpL, LM_Flashers_f17_FlipperSpL, LM_Flashers_f18_FlipperSpL, LM_Flashers_f28_FlipperSpL, LM_GI_Left_l100a_FlipperSpL, LM_GI_Left_l100b_FlipperSpL, LM_GI_Left_l100e_FlipperSpL, LM_GI_Right_l104a_FlipperSpL, LM_GI_Right_l104c_FlipperSpL, LM_GI_Right_l104d_FlipperSpL, LM_Inserts_l41_FlipperSpL, LM_Inserts_l42_FlipperSpL, LM_Inserts_l43_FlipperSpL, LM_Inserts_l47_FlipperSpL)
Dim BP_FlipperSpL1: BP_FlipperSpL1=Array(BM_FlipperSpL1, LM_Flashers_f18_FlipperSpL1, LM_Flashers_f19_FlipperSpL1, LM_Flashers_f20_FlipperSpL1, LM_Flashers_f28_FlipperSpL1, LM_Flashers_f38_FlipperSpL1, LM_Flashers_f39_FlipperSpL1, LM_Flashers_f41_FlipperSpL1, LM_GI_Right_l104_FlipperSpL1, LM_Inserts_l55_FlipperSpL1, LM_Bumpers_l63_FlipperSpL1, LM_Inserts_l81_FlipperSpL1)
Dim BP_FlipperSpLU: BP_FlipperSpLU=Array(BM_FlipperSpLU, LM_Flashers_f17_FlipperSpLU, LM_Flashers_f18_FlipperSpLU, LM_Flashers_f28_FlipperSpLU, LM_GI_Left_l100a_FlipperSpLU, LM_GI_Left_l100b_FlipperSpLU, LM_GI_Left_l100c_FlipperSpLU, LM_GI_Left_l100d_FlipperSpLU, LM_GI_Left_l100e_FlipperSpLU, LM_GI_Right_l104a_FlipperSpLU, LM_GI_Right_l104c_FlipperSpLU, LM_GI_Right_l104d_FlipperSpLU, LM_Inserts_l41_FlipperSpLU, LM_Inserts_l42_FlipperSpLU, LM_Inserts_l43_FlipperSpLU, LM_Inserts_l44_FlipperSpLU, LM_Inserts_l45_FlipperSpLU, LM_Inserts_l47_FlipperSpLU)
Dim BP_FlipperSpR: BP_FlipperSpR=Array(BM_FlipperSpR, LM_Flashers_f28_FlipperSpR, LM_Flashers_f39_FlipperSpR, LM_GI_Left_l100a_FlipperSpR, LM_GI_Left_l100e_FlipperSpR, LM_GI_Right_l104_FlipperSpR, LM_GI_Right_l104a_FlipperSpR, LM_GI_Right_l104c_FlipperSpR, LM_GI_Right_l104d_FlipperSpR, LM_Inserts_l44_FlipperSpR, LM_Inserts_l45_FlipperSpR, LM_Inserts_l46_FlipperSpR, LM_Inserts_l47_FlipperSpR)
Dim BP_FlipperSpR1: BP_FlipperSpR1=Array(BM_FlipperSpR1, LM_Mod_l106_FlipperSpR1, LM_Mod_l107_FlipperSpR1, LM_Mod_l111_FlipperSpR1, LM_Flashers_f18_FlipperSpR1, LM_Flashers_f19_FlipperSpR1, LM_Flashers_f28_FlipperSpR1, LM_Flashers_f37_FlipperSpR1, LM_Flashers_f38_FlipperSpR1, LM_Flashers_f39_FlipperSpR1, LM_Flashers_f41_FlipperSpR1, LM_GI_Right_l104_FlipperSpR1, LM_Inserts_l86_FlipperSpR1)
Dim BP_FlipperSpR1U: BP_FlipperSpR1U=Array(BM_FlipperSpR1U, LM_Mod_l106_FlipperSpR1U, LM_Mod_l107_FlipperSpR1U, LM_Mod_l111_FlipperSpR1U, LM_Flashers_f18_FlipperSpR1U, LM_Flashers_f19_FlipperSpR1U, LM_Flashers_f20_FlipperSpR1U, LM_Flashers_f28_FlipperSpR1U, LM_Flashers_f37_FlipperSpR1U, LM_Flashers_f38_FlipperSpR1U, LM_Flashers_f39_FlipperSpR1U, LM_Flashers_f41_FlipperSpR1U, LM_GI_Right_l104_FlipperSpR1U, LM_GI_Right_l104a_FlipperSpR1U, LM_GI_Right_l104d_FlipperSpR1U, LM_Inserts_l26_FlipperSpR1U, LM_Inserts_l27_FlipperSpR1U, LM_Inserts_l71_FlipperSpR1U, LM_Inserts_l86_FlipperSpR1U)
Dim BP_FlipperSpRU: BP_FlipperSpRU=Array(BM_FlipperSpRU, LM_Flashers_f18_FlipperSpRU, LM_Flashers_f20_FlipperSpRU, LM_Flashers_f28_FlipperSpRU, LM_Flashers_f39_FlipperSpRU, LM_GI_Left_l100a_FlipperSpRU, LM_GI_Left_l100e_FlipperSpRU, LM_GI_Right_l104_FlipperSpRU, LM_GI_Right_l104a_FlipperSpRU, LM_GI_Right_l104c_FlipperSpRU, LM_GI_Right_l104d_FlipperSpRU, LM_Inserts_l42_FlipperSpRU, LM_Inserts_l43_FlipperSpRU, LM_Inserts_l44_FlipperSpRU, LM_Inserts_l45_FlipperSpRU, LM_Inserts_l46_FlipperSpRU, LM_Inserts_l47_FlipperSpRU)
Dim BP_GMKnob: BP_GMKnob=Array(BM_GMKnob, LM_Flashers_f18_GMKnob, LM_Flashers_f19_GMKnob, LM_Flashers_f20_GMKnob, LM_Flashers_f28_GMKnob, LM_Flashers_f38_GMKnob, LM_Flashers_f39_GMKnob, LM_Flashers_f40_GMKnob, LM_Flashers_f41_GMKnob, LM_GI_Left_l100_GMKnob, LM_GI_MiniPF_l101_GMKnob)
Dim BP_Gate1: BP_Gate1=Array(BM_Gate1, LM_Flashers_f18_Gate1, LM_GI_Right_l104_Gate1)
Dim BP_Gate2: BP_Gate2=Array(BM_Gate2)
Dim BP_Gumballs: BP_Gumballs=Array(BM_Gumballs, LM_Flashers_f18_Gumballs, LM_Flashers_f19_Gumballs, LM_Flashers_f20_Gumballs, LM_Flashers_f28_Gumballs, LM_Flashers_f38_Gumballs, LM_Flashers_f39_Gumballs, LM_Flashers_f40_Gumballs, LM_GI_Left_l100_Gumballs, LM_GI_MiniPF_l101_Gumballs)
Dim BP_InvaderToy: BP_InvaderToy=Array(BM_InvaderToy, LM_Mod_l105_InvaderToy, LM_Mod_l106_InvaderToy, LM_Mod_l107_InvaderToy, LM_Flashers_f18_InvaderToy, LM_Flashers_f20_InvaderToy, LM_Flashers_f37_InvaderToy, LM_GI_Right_l104_InvaderToy, LM_GI_Right_l104a_InvaderToy, LM_GI_Right_l104d_InvaderToy, LM_Inserts_l86_InvaderToy)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_Flashers_f18_LSling1, LM_Flashers_f28_LSling1, LM_GI_Left_l100a_LSling1, LM_GI_Left_l100b_LSling1, LM_GI_Right_l104a_LSling1, LM_Inserts_l11_LSling1, LM_Inserts_l12_LSling1, LM_Inserts_l13_LSling1, LM_Inserts_l41_LSling1, LM_Inserts_l42_LSling1, LM_Inserts_l43_LSling1, LM_Inserts_l44_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_Flashers_f18_LSling2, LM_GI_Left_l100a_LSling2, LM_GI_Left_l100b_LSling2, LM_GI_Right_l104a_LSling2, LM_Inserts_l11_LSling2, LM_Inserts_l12_LSling2, LM_Inserts_l13_LSling2, LM_Inserts_l14_LSling2, LM_Inserts_l41_LSling2, LM_Inserts_l42_LSling2, LM_Inserts_l43_LSling2)
Dim BP_MysticSeerToy: BP_MysticSeerToy=Array(BM_MysticSeerToy, LM_Flashers_f17_MysticSeerToy, LM_GI_Left_l100a_MysticSeerToy, LM_GI_Left_l100b_MysticSeerToy, LM_Inserts_l33_MysticSeerToy, LM_Bumpers_l61_MysticSeerToy, LM_Bumpers_l62_MysticSeerToy)
Dim BP_Outlanes: BP_Outlanes=Array(BM_Outlanes, LM_Mod_l111_Outlanes, LM_Flashers_f17_Outlanes, LM_GI_Left_l100a_Outlanes, LM_GI_Left_l100b_Outlanes, LM_GI_Left_l100e_Outlanes, LM_GI_Right_l104a_Outlanes, LM_GI_Right_l104c_Outlanes, LM_GI_Right_l104d_Outlanes, LM_Inserts_l31_Outlanes, LM_Inserts_l33_Outlanes, LM_Inserts_l35_Outlanes, LM_Inserts_l37_Outlanes, LM_Inserts_l48_Outlanes, LM_Bumpers_l62_Outlanes, LM_Inserts_l66_Outlanes)
Dim BP_Over1: BP_Over1=Array(BM_Over1, LM_GI_Clock_l102_Over1, LM_Mod_l105_Over1, LM_Mod_l106_Over1, LM_Flashers_f17_Over1, LM_Flashers_f18_Over1, LM_Flashers_l74_Over1, LM_Flashers_f19_Over1, LM_Flashers_f20_Over1, LM_Flashers_f28_Over1, LM_Flashers_f37_Over1, LM_Flashers_f38_Over1, LM_Flashers_f39_Over1, LM_Flashers_f41_Over1, LM_GI_Left_l100_Over1, LM_GI_Right_l104_Over1, LM_Inserts_l67_Over1, LM_Inserts_l68_Over1, LM_Inserts_l83_Over1, LM_Inserts_l84_Over1, LM_Inserts_l86_Over1)
Dim BP_Over2: BP_Over2=Array(BM_Over2, LM_GI_Clock_l102_Over2, LM_Flashers_f18_Over2, LM_Flashers_l74_Over2, LM_Flashers_f19_Over2, LM_Flashers_f20_Over2, LM_Flashers_f28_Over2, LM_Flashers_f38_Over2, LM_Flashers_f39_Over2, LM_Flashers_f40_Over2, LM_Flashers_f41_Over2, LM_GI_Left_l100_Over2, LM_GI_Right_l104_Over2, LM_Inserts_l57_Over2, LM_Inserts_l58_Over2, LM_Inserts_l81_Over2, LM_Inserts_l83_Over2, LM_Inserts_l84_Over2, LM_Inserts_l85_Over2)
Dim BP_Over3: BP_Over3=Array(BM_Over3, LM_GI_Clock_l102_Over3, LM_Flashers_f17_Over3, LM_Flashers_f18_Over3, LM_Flashers_l74_Over3, LM_Flashers_f19_Over3, LM_Flashers_f20_Over3, LM_Flashers_f28_Over3, LM_Flashers_f38_Over3, LM_Flashers_f39_Over3, LM_Flashers_f40_Over3, LM_Flashers_f41_Over3, LM_GI_Left_l100_Over3, LM_GI_Right_l104_Over3, LM_GI_Right_l104a_Over3, LM_Inserts_l53_Over3, LM_Inserts_l56_Over3, LM_Inserts_l57_Over3, LM_Bumpers_l61_Over3, LM_Bumpers_l62_Over3, LM_Bumpers_l63_Over3, LM_Inserts_l68_Over3, LM_Inserts_l75_Over3, LM_Inserts_l82_Over3, LM_Inserts_l83_Over3, LM_Inserts_l84_Over3, LM_Inserts_l85_Over3, LM_Inserts_l86_Over3)
Dim BP_Over4: BP_Over4=Array(BM_Over4, LM_Flashers_f20_Over4, LM_Flashers_f28_Over4, LM_Flashers_f39_Over4, LM_Flashers_f41_Over4, LM_GI_Left_l100_Over4)
Dim BP_Over5: BP_Over5=Array(BM_Over5, LM_Flashers_f20_Over5, LM_Flashers_f28_Over5, LM_GI_Left_l100_Over5)
Dim BP_PF: BP_PF=Array(BM_PF, LM_Mod_l110_PF, LM_GI_Clock_l102_PF, LM_Mod_l106_PF, LM_Mod_l107_PF, LM_Mod_l111_PF, LM_Mod_l108_PF, LM_Mod_l109_PF, LM_Flashers_f17_PF, LM_Flashers_f18_PF, LM_Flashers_l74_PF, LM_Flashers_f19_PF, LM_Flashers_f20_PF, LM_Flashers_f28_PF, LM_Flashers_f37_PF, LM_Flashers_f38_PF, LM_Flashers_f39_PF, LM_Flashers_f40_PF, LM_Flashers_f41_PF, LM_GI_Left_l100_PF, LM_GI_Left_l100a_PF, LM_GI_Left_l100b_PF, LM_GI_Left_l100c_PF, LM_GI_Left_l100d_PF, LM_GI_Left_l100e_PF, LM_GI_MiniPF_l101_PF, LM_GI_Right_l104_PF, LM_GI_Right_l104a_PF, LM_GI_Right_l104b_PF, LM_GI_Right_l104c_PF, LM_GI_Right_l104d_PF, LM_Inserts_l15_PF, LM_Inserts_l16_PF, LM_Inserts_l25_PF, LM_Inserts_l27_PF, LM_Inserts_l33_PF, LM_Inserts_l34_PF, LM_Inserts_l35_PF, LM_Inserts_l37_PF, LM_Inserts_l41_PF, LM_Inserts_l42_PF, LM_Inserts_l43_PF, LM_Inserts_l44_PF, LM_Inserts_l45_PF, LM_Inserts_l46_PF, LM_Inserts_l47_PF, LM_Inserts_l48_PF, LM_Inserts_l54_PF, LM_Inserts_l56_PF, LM_Inserts_l57_PF, LM_Bumpers_l61_PF, LM_Bumpers_l62_PF, _
  LM_Bumpers_l63_PF, LM_Inserts_l64_PF, LM_Inserts_l65_PF, LM_Inserts_l71_PF, LM_Inserts_l72_PF, LM_Inserts_l73_PF, LM_Inserts_l75_PF, LM_Inserts_l81_PF, LM_Inserts_l82_PF, LM_Inserts_l83_PF, LM_Inserts_l84_PF, LM_Inserts_l85_PF, LM_Inserts_l86_PF)
Dim BP_PF_Upper: BP_PF_Upper=Array(BM_PF_Upper, LM_Mod_l110_PF_Upper, LM_Mod_l109_PF_Upper, LM_Flashers_f17_PF_Upper, LM_Flashers_f18_PF_Upper, LM_Flashers_f19_PF_Upper, LM_Flashers_f20_PF_Upper, LM_Flashers_f28_PF_Upper, LM_Flashers_f38_PF_Upper, LM_Flashers_f39_PF_Upper, LM_Flashers_f40_PF_Upper, LM_Flashers_f41_PF_Upper, LM_GI_Left_l100_PF_Upper, LM_GI_Left_l100b_PF_Upper, LM_GI_MiniPF_l101_PF_Upper, LM_GI_Right_l104_PF_Upper, LM_Inserts_l15_PF_Upper, LM_Inserts_l16_PF_Upper, LM_Inserts_l17_PF_Upper, LM_Inserts_l18_PF_Upper, LM_Inserts_l38_PF_Upper, LM_Inserts_l52_PF_Upper, LM_Inserts_l53_PF_Upper, LM_Inserts_l54_PF_Upper, LM_Inserts_l55_PF_Upper, LM_Bumpers_l61_PF_Upper, LM_Bumpers_l62_PF_Upper, LM_Bumpers_l63_PF_Upper, LM_Inserts_l64_PF_Upper, LM_Inserts_l65_PF_Upper, LM_Inserts_l76_PF_Upper, LM_Inserts_l77_PF_Upper, LM_Inserts_l78_PF_Upper, LM_Inserts_l81_PF_Upper)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_Mod_l110_Parts, LM_GI_Clock_l102_Parts, LM_Mod_l105_Parts, LM_Mod_l106_Parts, LM_Mod_l107_Parts, LM_Mod_l111_Parts, LM_Mod_l108_Parts, LM_Mod_l109_Parts, LM_Flashers_f17_Parts, LM_Flashers_f18_Parts, LM_Flashers_l74_Parts, LM_Flashers_f19_Parts, LM_Flashers_f20_Parts, LM_Flashers_f28_Parts, LM_Flashers_f37_Parts, LM_Flashers_f38_Parts, LM_Flashers_f39_Parts, LM_Flashers_f40_Parts, LM_Flashers_f41_Parts, LM_GI_Left_l100_Parts, LM_GI_Left_l100a_Parts, LM_GI_Left_l100b_Parts, LM_GI_Left_l100c_Parts, LM_GI_Left_l100d_Parts, LM_GI_Left_l100e_Parts, LM_GI_MiniPF_l101_Parts, LM_GI_Right_l104_Parts, LM_GI_Right_l104a_Parts, LM_GI_Right_l104b_Parts, LM_GI_Right_l104c_Parts, LM_GI_Right_l104d_Parts, LM_Inserts_l11_Parts, LM_Inserts_l12_Parts, LM_Inserts_l13_Parts, LM_Inserts_l14_Parts, LM_Inserts_l15_Parts, LM_Inserts_l16_Parts, LM_Inserts_l17_Parts, LM_Inserts_l18_Parts, LM_Inserts_l21_Parts, LM_Inserts_l22_Parts, LM_Inserts_l23_Parts, LM_Inserts_l24_Parts, _
  LM_Inserts_l25_Parts, LM_Inserts_l26_Parts, LM_Inserts_l27_Parts, LM_Inserts_l28_Parts, LM_Inserts_l31_Parts, LM_Inserts_l32_Parts, LM_Inserts_l33_Parts, LM_Inserts_l34_Parts, LM_Inserts_l35_Parts, LM_Inserts_l36_Parts, LM_Inserts_l37_Parts, LM_Inserts_l38_Parts, LM_Inserts_l41_Parts, LM_Inserts_l42_Parts, LM_Inserts_l43_Parts, LM_Inserts_l44_Parts, LM_Inserts_l45_Parts, LM_Inserts_l46_Parts, LM_Inserts_l47_Parts, LM_Inserts_l48_Parts, LM_Inserts_l51_Parts, LM_Inserts_l52_Parts, LM_Inserts_l53_Parts, LM_Inserts_l54_Parts, LM_Inserts_l55_Parts, LM_Inserts_l56_Parts, LM_Inserts_l57_Parts, LM_Inserts_l58_Parts, LM_Bumpers_l61_Parts, LM_Bumpers_l62_Parts, LM_Bumpers_l63_Parts, LM_Inserts_l64_Parts, LM_Inserts_l65_Parts, LM_Inserts_l66_Parts, LM_Inserts_l67_Parts, LM_Inserts_l68_Parts, LM_Inserts_l71_Parts, LM_Inserts_l72_Parts, LM_Inserts_l73_Parts, LM_Inserts_l75_Parts, LM_Inserts_l76_Parts, LM_Inserts_l77_Parts, LM_Inserts_l78_Parts, LM_Inserts_l81_Parts, LM_Inserts_l82_Parts, LM_Inserts_l83_Parts, _
  LM_Inserts_l84_Parts, LM_Inserts_l85_Parts, LM_Inserts_l86_Parts)
Dim BP_Piano: BP_Piano=Array(BM_Piano, LM_GI_Clock_l102_Piano, LM_Flashers_f18_Piano, LM_Flashers_l74_Piano, LM_Flashers_f19_Piano, LM_Flashers_f20_Piano, LM_Flashers_f28_Piano, LM_Flashers_f38_Piano, LM_Flashers_f39_Piano, LM_Flashers_f41_Piano, LM_GI_Right_l104_Piano, LM_Inserts_l56_Piano, LM_Inserts_l75_Piano, LM_Inserts_l82_Piano, LM_Inserts_l83_Piano, LM_Inserts_l84_Piano)
Dim BP_Pyramid: BP_Pyramid=Array(BM_Pyramid, LM_Flashers_f18_Pyramid, LM_Flashers_f19_Pyramid, LM_Flashers_f20_Pyramid, LM_Flashers_f28_Pyramid, LM_Flashers_f38_Pyramid, LM_Flashers_f39_Pyramid, LM_Flashers_f41_Pyramid, LM_GI_MiniPF_l101_Pyramid, LM_Inserts_l76_Pyramid, LM_Inserts_l78_Pyramid)
Dim BP_RDiv: BP_RDiv=Array(BM_RDiv, LM_Flashers_f19_RDiv, LM_Flashers_f20_RDiv, LM_Flashers_f28_RDiv, LM_Flashers_f38_RDiv, LM_Flashers_f39_RDiv, LM_Flashers_f40_RDiv, LM_Flashers_f41_RDiv, LM_GI_Left_l100_RDiv)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_Mod_l111_RSling1, LM_Flashers_f20_RSling1, LM_Flashers_f28_RSling1, LM_Flashers_f38_RSling1, LM_Flashers_f39_RSling1, LM_GI_Left_l100a_RSling1, LM_GI_Right_l104a_RSling1, LM_Inserts_l11_RSling1, LM_Inserts_l22_RSling1, LM_Inserts_l23_RSling1, LM_Inserts_l43_RSling1, LM_Inserts_l44_RSling1, LM_Inserts_l45_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_Mod_l111_RSling2, LM_Flashers_f28_RSling2, LM_Flashers_f38_RSling2, LM_Flashers_f39_RSling2, LM_GI_Left_l100a_RSling2, LM_GI_Right_l104a_RSling2, LM_Inserts_l22_RSling2, LM_Inserts_l23_RSling2, LM_Inserts_l24_RSling2, LM_Inserts_l44_RSling2, LM_Inserts_l45_RSling2, LM_Inserts_l48_RSling2)
Dim BP_Rails: BP_Rails=Array(BM_Rails, LM_GI_Clock_l102_Rails, LM_Flashers_f17_Rails, LM_Flashers_f18_Rails, LM_Flashers_f19_Rails, LM_Flashers_f20_Rails, LM_Flashers_f28_Rails, LM_Flashers_f37_Rails, LM_Flashers_f38_Rails, LM_Flashers_f39_Rails, LM_Flashers_f40_Rails, LM_Flashers_f41_Rails, LM_GI_Left_l100_Rails, LM_GI_Left_l100a_Rails, LM_GI_Right_l104_Rails, LM_GI_Right_l104a_Rails, LM_GI_Right_l104c_Rails, LM_GI_Right_l104d_Rails, LM_Bumpers_l61_Rails, LM_Bumpers_l63_Rails)
Dim BP_Rails_DT: BP_Rails_DT=Array(BM_Rails_DT, LM_Flashers_f20_Rails_DT, LM_Flashers_f38_Rails_DT, LM_Flashers_f39_Rails_DT, LM_Flashers_f41_Rails_DT)
Dim BP_Robot: BP_Robot=Array(BM_Robot, LM_Flashers_f18_Robot, LM_Flashers_l74_Robot, LM_Flashers_f19_Robot, LM_Flashers_f20_Robot, LM_Flashers_f28_Robot, LM_Flashers_f38_Robot, LM_Flashers_f39_Robot, LM_Flashers_f40_Robot, LM_Flashers_f41_Robot, LM_GI_Left_l100_Robot, LM_GI_Right_l104_Robot, LM_GI_Right_l104b_Robot, LM_Inserts_l51_Robot, LM_Inserts_l82_Robot)
Dim BP_RocketToy: BP_RocketToy=Array(BM_RocketToy, LM_Mod_l111_RocketToy, LM_Flashers_f17_RocketToy, LM_Flashers_f18_RocketToy, LM_Flashers_f19_RocketToy, LM_Flashers_f28_RocketToy, LM_Flashers_f37_RocketToy, LM_Flashers_f41_RocketToy, LM_GI_Left_l100a_RocketToy, LM_GI_Left_l100e_RocketToy, LM_GI_Right_l104_RocketToy, LM_GI_Right_l104a_RocketToy, LM_GI_Right_l104d_RocketToy, LM_Inserts_l48_RocketToy, LM_Bumpers_l62_RocketToy, LM_Inserts_l66_RocketToy, LM_Inserts_l85_RocketToy)
Dim BP_SLING1: BP_SLING1=Array(BM_SLING1, LM_Mod_l111_SLING1, LM_Flashers_f39_SLING1, LM_GI_Right_l104a_SLING1, LM_GI_Right_l104c_SLING1, LM_GI_Right_l104d_SLING1, LM_Inserts_l23_SLING1, LM_Inserts_l45_SLING1)
Dim BP_SLING2: BP_SLING2=Array(BM_SLING2, LM_Flashers_f19_SLING2, LM_GI_Left_l100a_SLING2, LM_GI_Left_l100b_SLING2, LM_GI_Left_l100c_SLING2, LM_GI_Left_l100d_SLING2, LM_GI_Left_l100e_SLING2, LM_Inserts_l12_SLING2, LM_Inserts_l41_SLING2)
Dim BP_ShooterDiv: BP_ShooterDiv=Array(BM_ShooterDiv)
Dim BP_SideMod: BP_SideMod=Array(BM_SideMod, LM_GI_Clock_l102_SideMod, LM_Flashers_f17_SideMod, LM_Flashers_f18_SideMod, LM_Flashers_f19_SideMod, LM_Flashers_f20_SideMod, LM_Flashers_f28_SideMod, LM_Flashers_f38_SideMod, LM_Flashers_f39_SideMod, LM_Flashers_f40_SideMod, LM_Flashers_f41_SideMod, LM_GI_Left_l100_SideMod, LM_GI_Left_l100b_SideMod, LM_GI_Left_l100c_SideMod, LM_GI_MiniPF_l101_SideMod, LM_GI_Right_l104_SideMod, LM_Inserts_l31_SideMod, LM_Bumpers_l61_SideMod, LM_Bumpers_l62_SideMod)
Dim BP_Sign_Spiral: BP_Sign_Spiral=Array(BM_Sign_Spiral, LM_GI_Clock_l102_Sign_Spiral, LM_Flashers_f18_Sign_Spiral, LM_Flashers_f19_Sign_Spiral, LM_Flashers_f20_Sign_Spiral, LM_Flashers_f28_Sign_Spiral, LM_Flashers_f37_Sign_Spiral, LM_Flashers_f38_Sign_Spiral, LM_Flashers_f39_Sign_Spiral, LM_Flashers_f40_Sign_Spiral, LM_GI_Left_l100_Sign_Spiral, LM_GI_MiniPF_l101_Sign_Spiral, LM_GI_Right_l104_Sign_Spiral, LM_GI_Right_l104b_Sign_Spiral, LM_Inserts_l51_Sign_Spiral, LM_Inserts_l52_Sign_Spiral, LM_Inserts_l53_Sign_Spiral, LM_Inserts_l54_Sign_Spiral, LM_Inserts_l55_Sign_Spiral, LM_Inserts_l56_Sign_Spiral, LM_Bumpers_l63_Sign_Spiral, LM_Inserts_l72_Sign_Spiral, LM_Inserts_l73_Sign_Spiral, LM_Inserts_l81_Sign_Spiral, LM_Inserts_l82_Sign_Spiral)
Dim BP_SlotMachineToy: BP_SlotMachineToy=Array(BM_SlotMachineToy, LM_GI_Clock_l102_SlotMachineToy, LM_Mod_l106_SlotMachineToy, LM_Flashers_f17_SlotMachineToy, LM_Flashers_f18_SlotMachineToy, LM_Flashers_f19_SlotMachineToy, LM_Flashers_f20_SlotMachineToy, LM_Flashers_f28_SlotMachineToy, LM_Flashers_f37_SlotMachineToy, LM_Flashers_f38_SlotMachineToy, LM_Flashers_f39_SlotMachineToy, LM_Flashers_f41_SlotMachineToy, LM_GI_Left_l100_SlotMachineToy, LM_GI_Left_l100b_SlotMachineToy, LM_GI_Right_l104_SlotMachineToy, LM_GI_Right_l104b_SlotMachineTo, LM_Inserts_l28_SlotMachineToy, LM_Bumpers_l61_SlotMachineToy, LM_Bumpers_l62_SlotMachineToy, LM_Bumpers_l63_SlotMachineToy, LM_Inserts_l67_SlotMachineToy, LM_Inserts_l71_SlotMachineToy, LM_Inserts_l85_SlotMachineToy, LM_Inserts_l86_SlotMachineToy)
Dim BP_SpiralToy: BP_SpiralToy=Array(BM_SpiralToy, LM_Flashers_f19_SpiralToy, LM_Flashers_f20_SpiralToy, LM_GI_MiniPF_l101_SpiralToy)
Dim BP_TVtoy: BP_TVtoy=Array(BM_TVtoy, LM_Flashers_f18_TVtoy, LM_Flashers_f20_TVtoy, LM_Flashers_f28_TVtoy, LM_Flashers_f41_TVtoy, LM_GI_Right_l104_TVtoy)
Dim BP_TownSquarePost: BP_TownSquarePost=Array(BM_TownSquarePost, LM_Mod_l109_TownSquarePost, LM_Flashers_f17_TownSquarePost, LM_Flashers_f18_TownSquarePost, LM_GI_Left_l100_TownSquarePost, LM_GI_Left_l100b_TownSquarePost, LM_Inserts_l38_TownSquarePost, LM_Bumpers_l63_TownSquarePost, LM_Inserts_l64_TownSquarePost)
Dim BP_URMagnet: BP_URMagnet=Array(BM_URMagnet, LM_Flashers_f41_URMagnet, LM_GI_Right_l104_URMagnet)
Dim BP_sw11: BP_sw11=Array(BM_sw11, LM_Flashers_f28_sw11, LM_GI_Right_l104a_sw11, LM_GI_Right_l104c_sw11, LM_GI_Right_l104d_sw11, LM_Inserts_l66_sw11)
Dim BP_sw12: BP_sw12=Array(BM_sw12, LM_Flashers_f39_sw12, LM_GI_Right_l104a_sw12, LM_GI_Right_l104d_sw12)
Dim BP_sw27: BP_sw27=Array(BM_sw27)
Dim BP_sw36: BP_sw36=Array(BM_sw36, LM_Flashers_f17_sw36, LM_GI_Left_l100_sw36, LM_GI_Left_l100a_sw36, LM_GI_Left_l100b_sw36, LM_GI_Left_l100c_sw36, LM_GI_Left_l100d_sw36, LM_Inserts_l31_sw36, LM_Bumpers_l61_sw36)
Dim BP_sw37: BP_sw37=Array(BM_sw37, LM_Flashers_f17_sw37, LM_GI_Left_l100_sw37, LM_GI_Left_l100b_sw37, LM_GI_Left_l100c_sw37, LM_GI_Left_l100d_sw37, LM_Bumpers_l61_sw37)
Dim BP_sw38: BP_sw38=Array(BM_sw38, LM_Flashers_f17_sw38, LM_Flashers_f19_sw38, LM_Flashers_f28_sw38, LM_GI_Left_l100_sw38, LM_GI_Left_l100a_sw38, LM_GI_Left_l100b_sw38, LM_GI_Left_l100c_sw38, LM_GI_Left_l100d_sw38, LM_GI_Left_l100e_sw38)
Dim BP_sw47: BP_sw47=Array(BM_sw47, LM_Flashers_f17_sw47, LM_Flashers_f18_sw47, LM_Flashers_f19_sw47, LM_GI_Right_l104_sw47, LM_Inserts_l71_sw47)
Dim BP_sw47m: BP_sw47m=Array(BM_sw47m, LM_Flashers_f28_sw47m, LM_Inserts_l51_sw47m, LM_Inserts_l82_sw47m)
Dim BP_sw48: BP_sw48=Array(BM_sw48, LM_Flashers_f17_sw48, LM_GI_Left_l100a_sw48, LM_GI_Left_l100b_sw48, LM_GI_Left_l100c_sw48, LM_GI_Left_l100d_sw48, LM_GI_Left_l100e_sw48, LM_GI_Right_l104a_sw48, LM_Inserts_l13_sw48, LM_Inserts_l14_sw48, LM_Inserts_l15_sw48, LM_Inserts_l33_sw48, LM_Inserts_l34_sw48, LM_Inserts_l35_sw48, LM_Inserts_l37_sw48, LM_Bumpers_l62_sw48)
Dim BP_sw48m: BP_sw48m=Array(BM_sw48m, LM_Mod_l109_sw48m, LM_Flashers_f17_sw48m, LM_GI_Left_l100a_sw48m, LM_GI_Left_l100b_sw48m, LM_GI_Left_l100d_sw48m, LM_Inserts_l14_sw48m, LM_Inserts_l15_sw48m, LM_Inserts_l35_sw48m, LM_Inserts_l37_sw48m, LM_Inserts_l38_sw48m, LM_Bumpers_l62_sw48m)
Dim BP_sw53p: BP_sw53p=Array(BM_sw53p, LM_Flashers_f20_sw53p, LM_Flashers_f28_sw53p, LM_GI_Left_l100_sw53p)
Dim BP_sw54p: BP_sw54p=Array(BM_sw54p, LM_Flashers_f41_sw54p)
Dim BP_sw56: BP_sw56=Array(BM_sw56, LM_Flashers_f19_sw56, LM_Flashers_f20_sw56, LM_Flashers_f38_sw56, LM_Flashers_f39_sw56, LM_GI_Left_l100_sw56)
Dim BP_sw61: BP_sw61=Array(BM_sw61, LM_GI_Right_l104_sw61)
Dim BP_sw62: BP_sw62=Array(BM_sw62, LM_GI_Right_l104_sw62)
Dim BP_sw63: BP_sw63=Array(BM_sw63, LM_GI_Clock_l102_sw63, LM_Flashers_f18_sw63, LM_Flashers_f28_sw63, LM_Flashers_f37_sw63, LM_GI_Right_l104_sw63)
Dim BP_sw64: BP_sw64=Array(BM_sw64, LM_GI_Clock_l102_sw64, LM_Flashers_f18_sw64, LM_Flashers_l74_sw64, LM_Flashers_f19_sw64, LM_Flashers_f28_sw64, LM_Inserts_l56_sw64, LM_Inserts_l57_sw64, LM_Inserts_l75_sw64, LM_Inserts_l82_sw64, LM_Inserts_l83_sw64, LM_Inserts_l84_sw64)
Dim BP_sw64m: BP_sw64m=Array(BM_sw64m, LM_GI_Clock_l102_sw64m, LM_Flashers_f19_sw64m, LM_Flashers_f28_sw64m, LM_GI_Right_l104_sw64m, LM_Inserts_l75_sw64m, LM_Inserts_l83_sw64m, LM_Inserts_l84_sw64m)
Dim BP_sw65: BP_sw65=Array(BM_sw65, LM_GI_Clock_l102_sw65, LM_Flashers_f19_sw65, LM_Flashers_f28_sw65, LM_GI_Right_l104_sw65)
Dim BP_sw65a: BP_sw65a=Array(BM_sw65a, LM_GI_Clock_l102_sw65a, LM_Flashers_f19_sw65a, LM_Flashers_f28_sw65a, LM_GI_Right_l104_sw65a)
Dim BP_sw65am: BP_sw65am=Array(BM_sw65am, LM_GI_Clock_l102_sw65am, LM_Flashers_f18_sw65am, LM_Flashers_f19_sw65am, LM_Flashers_f20_sw65am, LM_Flashers_f28_sw65am, LM_GI_Right_l104_sw65am)
Dim BP_sw65m: BP_sw65m=Array(BM_sw65m, LM_GI_Clock_l102_sw65m, LM_Flashers_f18_sw65m, LM_Flashers_f19_sw65m, LM_Flashers_f28_sw65m, LM_GI_Right_l104_sw65m, LM_GI_Right_l104b_sw65m, LM_Inserts_l85_sw65m)
Dim BP_sw66: BP_sw66=Array(BM_sw66, LM_Flashers_l74_sw66, LM_Flashers_f19_sw66, LM_Flashers_f28_sw66, LM_GI_Right_l104_sw66, LM_GI_Right_l104b_sw66, LM_Inserts_l71_sw66, LM_Inserts_l72_sw66, LM_Inserts_l73_sw66)
Dim BP_sw66m: BP_sw66m=Array(BM_sw66m, LM_Flashers_f19_sw66m, LM_GI_Right_l104_sw66m, LM_GI_Right_l104b_sw66m, LM_Inserts_l73_sw66m)
Dim BP_sw67: BP_sw67=Array(BM_sw67, LM_Flashers_f19_sw67, LM_Flashers_f28_sw67, LM_GI_Right_l104_sw67, LM_GI_Right_l104b_sw67, LM_Inserts_l71_sw67, LM_Inserts_l72_sw67, LM_Inserts_l73_sw67)
Dim BP_sw67m: BP_sw67m=Array(BM_sw67m, LM_Flashers_f18_sw67m, LM_Flashers_f19_sw67m, LM_GI_Right_l104_sw67m, LM_GI_Right_l104b_sw67m, LM_Inserts_l71_sw67m, LM_Inserts_l73_sw67m, LM_Inserts_l85_sw67m)
Dim BP_sw68: BP_sw68=Array(BM_sw68, LM_Flashers_f18_sw68, LM_Flashers_f28_sw68, LM_GI_Left_l100_sw68, LM_Inserts_l51_sw68, LM_Inserts_l52_sw68, LM_Inserts_l56_sw68, LM_Inserts_l82_sw68)
Dim BP_sw68m: BP_sw68m=Array(BM_sw68m, LM_Mod_l108_sw68m, LM_GI_Right_l104_sw68m, LM_Inserts_l71_sw68m, LM_Inserts_l72_sw68m)
Dim BP_sw77: BP_sw77=Array(BM_sw77, LM_Mod_l109_sw77, LM_Flashers_f17_sw77, LM_GI_Left_l100_sw77, LM_GI_Left_l100a_sw77, LM_GI_Left_l100b_sw77, LM_Inserts_l14_sw77, LM_Inserts_l15_sw77, LM_Inserts_l16_sw77, LM_Inserts_l17_sw77, LM_Inserts_l18_sw77, LM_Inserts_l34_sw77, LM_Inserts_l36_sw77, LM_Inserts_l37_sw77, LM_Inserts_l38_sw77, LM_Bumpers_l62_sw77, LM_Bumpers_l63_sw77, LM_Inserts_l64_sw77, LM_Inserts_l65_sw77)
Dim BP_sw77m: BP_sw77m=Array(BM_sw77m, LM_Mod_l109_sw77m, LM_Flashers_f17_sw77m, LM_Inserts_l16_sw77m, LM_Inserts_l17_sw77m, LM_Inserts_l38_sw77m, LM_Bumpers_l62_sw77m, LM_Bumpers_l63_sw77m, LM_Inserts_l64_sw77m, LM_Inserts_l65_sw77m)
Dim BP_sw78: BP_sw78=Array(BM_sw78, LM_Mod_l109_sw78, LM_Flashers_f17_sw78, LM_Flashers_f20_sw78, LM_Flashers_f37_sw78, LM_GI_Right_l104_sw78, LM_Inserts_l16_sw78, LM_Inserts_l17_sw78, LM_Inserts_l18_sw78, LM_Inserts_l36_sw78, LM_Bumpers_l63_sw78, LM_Inserts_l65_sw78)
Dim BP_sw78m: BP_sw78m=Array(BM_sw78m, LM_Flashers_f17_sw78m, LM_Inserts_l17_sw78m, LM_Bumpers_l63_sw78m, LM_Inserts_l64_sw78m, LM_Inserts_l65_sw78m)
' Arrays per lighting scenario
Dim BL_Bumpers_l61: BL_Bumpers_l61=Array(LM_Bumpers_l61_B_Caps_Bot_Amber, LM_Bumpers_l61_B_Caps_Bot_Red, LM_Bumpers_l61_B_Caps_Bot_Yello, LM_Bumpers_l61_B_Caps_Top_Amber, LM_Bumpers_l61_B_Caps_Top_Red, LM_Bumpers_l61_B_Caps_Top_Yello, LM_Bumpers_l61_BR1, LM_Bumpers_l61_BR2, LM_Bumpers_l61_BR3, LM_Bumpers_l61_BS1, LM_Bumpers_l61_BS2, LM_Bumpers_l61_BS3, LM_Bumpers_l61_BumperPegs, LM_Bumpers_l61_ClockToy, LM_Bumpers_l61_MysticSeerToy, LM_Bumpers_l61_Over3, LM_Bumpers_l61_PF, LM_Bumpers_l61_PF_Upper, LM_Bumpers_l61_Parts, LM_Bumpers_l61_Rails, LM_Bumpers_l61_SideMod, LM_Bumpers_l61_SlotMachineToy, LM_Bumpers_l61_sw36, LM_Bumpers_l61_sw37)
Dim BL_Bumpers_l62: BL_Bumpers_l62=Array(LM_Bumpers_l62_B_Caps_Bot_Yello, LM_Bumpers_l62_B_Caps_Top_Amber, LM_Bumpers_l62_B_Caps_Top_Red, LM_Bumpers_l62_B_Caps_Top_Yello, LM_Bumpers_l62_BR1, LM_Bumpers_l62_BR2, LM_Bumpers_l62_BR3, LM_Bumpers_l62_BS1, LM_Bumpers_l62_BS2, LM_Bumpers_l62_BS3, LM_Bumpers_l62_BumperPegs, LM_Bumpers_l62_MysticSeerToy, LM_Bumpers_l62_Outlanes, LM_Bumpers_l62_Over3, LM_Bumpers_l62_PF, LM_Bumpers_l62_PF_Upper, LM_Bumpers_l62_Parts, LM_Bumpers_l62_RocketToy, LM_Bumpers_l62_SideMod, LM_Bumpers_l62_SlotMachineToy, LM_Bumpers_l62_sw48, LM_Bumpers_l62_sw48m, LM_Bumpers_l62_sw77, LM_Bumpers_l62_sw77m)
Dim BL_Bumpers_l63: BL_Bumpers_l63=Array(LM_Bumpers_l63_B_Caps_Bot_Amber, LM_Bumpers_l63_B_Caps_Bot_Red, LM_Bumpers_l63_B_Caps_Bot_Yello, LM_Bumpers_l63_B_Caps_Top_Amber, LM_Bumpers_l63_B_Caps_Top_Red, LM_Bumpers_l63_B_Caps_Top_Yello, LM_Bumpers_l63_BR1, LM_Bumpers_l63_BR2, LM_Bumpers_l63_BR3, LM_Bumpers_l63_BS1, LM_Bumpers_l63_BS2, LM_Bumpers_l63_BS3, LM_Bumpers_l63_BumperPegs, LM_Bumpers_l63_Camera, LM_Bumpers_l63_Clock_Color, LM_Bumpers_l63_Clock_White, LM_Bumpers_l63_FlipperL1, LM_Bumpers_l63_FlipperSpL1, LM_Bumpers_l63_Over3, LM_Bumpers_l63_PF, LM_Bumpers_l63_PF_Upper, LM_Bumpers_l63_Parts, LM_Bumpers_l63_Rails, LM_Bumpers_l63_Sign_Spiral, LM_Bumpers_l63_SlotMachineToy, LM_Bumpers_l63_TownSquarePost, LM_Bumpers_l63_sw77, LM_Bumpers_l63_sw77m, LM_Bumpers_l63_sw78, LM_Bumpers_l63_sw78m)
Dim BL_Flashers_f17: BL_Flashers_f17=Array(LM_Flashers_f17_B_Caps_Bot_Ambe, LM_Flashers_f17_B_Caps_Bot_Red, LM_Flashers_f17_B_Caps_Bot_Yell, LM_Flashers_f17_B_Caps_Top_Ambe, LM_Flashers_f17_B_Caps_Top_Red, LM_Flashers_f17_B_Caps_Top_Yell, LM_Flashers_f17_BR1, LM_Flashers_f17_BR2, LM_Flashers_f17_BR3, LM_Flashers_f17_BS1, LM_Flashers_f17_BS2, LM_Flashers_f17_BS3, LM_Flashers_f17_BumperPegs, LM_Flashers_f17_Camera, LM_Flashers_f17_Clock_Color, LM_Flashers_f17_Clock_White, LM_Flashers_f17_ClockToy, LM_Flashers_f17_FlipperLU, LM_Flashers_f17_FlipperR1, LM_Flashers_f17_FlipperR1U, LM_Flashers_f17_FlipperRU, LM_Flashers_f17_FlipperSpL, LM_Flashers_f17_FlipperSpLU, LM_Flashers_f17_MysticSeerToy, LM_Flashers_f17_Outlanes, LM_Flashers_f17_Over1, LM_Flashers_f17_Over3, LM_Flashers_f17_PF, LM_Flashers_f17_PF_Upper, LM_Flashers_f17_Parts, LM_Flashers_f17_Rails, LM_Flashers_f17_RocketToy, LM_Flashers_f17_SideMod, LM_Flashers_f17_SlotMachineToy, LM_Flashers_f17_TownSquarePost, LM_Flashers_f17_sw36, LM_Flashers_f17_sw37, _
  LM_Flashers_f17_sw38, LM_Flashers_f17_sw47, LM_Flashers_f17_sw48, LM_Flashers_f17_sw48m, LM_Flashers_f17_sw77, LM_Flashers_f17_sw77m, LM_Flashers_f17_sw78, LM_Flashers_f17_sw78m)
Dim BL_Flashers_f18: BL_Flashers_f18=Array(LM_Flashers_f18_B_Caps_Bot_Ambe, LM_Flashers_f18_B_Caps_Bot_Yell, LM_Flashers_f18_B_Caps_Top_Red, LM_Flashers_f18_B_Caps_Top_Yell, LM_Flashers_f18_Clock_Color, LM_Flashers_f18_Clock_White, LM_Flashers_f18_ClockToy, LM_Flashers_f18_DiverterP, LM_Flashers_f18_FlipperL, LM_Flashers_f18_FlipperL1, LM_Flashers_f18_FlipperLU, LM_Flashers_f18_FlipperR1, LM_Flashers_f18_FlipperR1U, LM_Flashers_f18_FlipperRU, LM_Flashers_f18_FlipperSpL, LM_Flashers_f18_FlipperSpL1, LM_Flashers_f18_FlipperSpLU, LM_Flashers_f18_FlipperSpR1, LM_Flashers_f18_FlipperSpR1U, LM_Flashers_f18_FlipperSpRU, LM_Flashers_f18_GMKnob, LM_Flashers_f18_Gate1, LM_Flashers_f18_Gumballs, LM_Flashers_f18_InvaderToy, LM_Flashers_f18_LSling1, LM_Flashers_f18_LSling2, LM_Flashers_f18_Over1, LM_Flashers_f18_Over2, LM_Flashers_f18_Over3, LM_Flashers_f18_PF, LM_Flashers_f18_PF_Upper, LM_Flashers_f18_Parts, LM_Flashers_f18_Piano, LM_Flashers_f18_Pyramid, LM_Flashers_f18_Rails, LM_Flashers_f18_Robot, _
  LM_Flashers_f18_RocketToy, LM_Flashers_f18_SideMod, LM_Flashers_f18_Sign_Spiral, LM_Flashers_f18_SlotMachineToy, LM_Flashers_f18_TVtoy, LM_Flashers_f18_TownSquarePost, LM_Flashers_f18_sw47, LM_Flashers_f18_sw63, LM_Flashers_f18_sw64, LM_Flashers_f18_sw65am, LM_Flashers_f18_sw65m, LM_Flashers_f18_sw67m, LM_Flashers_f18_sw68)
Dim BL_Flashers_f19: BL_Flashers_f19=Array(LM_Flashers_f19_BR2, LM_Flashers_f19_BR3, LM_Flashers_f19_BS3, LM_Flashers_f19_Camera, LM_Flashers_f19_Clock_Color, LM_Flashers_f19_Clock_White, LM_Flashers_f19_ClockLarge, LM_Flashers_f19_ClockToy, LM_Flashers_f19_FlipperL1, LM_Flashers_f19_FlipperR1, LM_Flashers_f19_FlipperR1U, LM_Flashers_f19_FlipperSpL1, LM_Flashers_f19_FlipperSpR1, LM_Flashers_f19_FlipperSpR1U, LM_Flashers_f19_GMKnob, LM_Flashers_f19_Gumballs, LM_Flashers_f19_Over1, LM_Flashers_f19_Over2, LM_Flashers_f19_Over3, LM_Flashers_f19_PF, LM_Flashers_f19_PF_Upper, LM_Flashers_f19_Parts, LM_Flashers_f19_Piano, LM_Flashers_f19_Pyramid, LM_Flashers_f19_RDiv, LM_Flashers_f19_Rails, LM_Flashers_f19_Robot, LM_Flashers_f19_RocketToy, LM_Flashers_f19_SLING2, LM_Flashers_f19_SideMod, LM_Flashers_f19_Sign_Spiral, LM_Flashers_f19_SlotMachineToy, LM_Flashers_f19_SpiralToy, LM_Flashers_f19_sw38, LM_Flashers_f19_sw47, LM_Flashers_f19_sw56, LM_Flashers_f19_sw64, LM_Flashers_f19_sw64m, LM_Flashers_f19_sw65, _
  LM_Flashers_f19_sw65a, LM_Flashers_f19_sw65am, LM_Flashers_f19_sw65m, LM_Flashers_f19_sw66, LM_Flashers_f19_sw66m, LM_Flashers_f19_sw67, LM_Flashers_f19_sw67m)
Dim BL_Flashers_f20: BL_Flashers_f20=Array(LM_Flashers_f20_B_Caps_Bot_Ambe, LM_Flashers_f20_BS2, LM_Flashers_f20_Camera, LM_Flashers_f20_Clock_Color, LM_Flashers_f20_Clock_White, LM_Flashers_f20_ClockToy, LM_Flashers_f20_DiverterP1, LM_Flashers_f20_FlipperL, LM_Flashers_f20_FlipperL1, LM_Flashers_f20_FlipperLU, LM_Flashers_f20_FlipperR, LM_Flashers_f20_FlipperR1, LM_Flashers_f20_FlipperR1U, LM_Flashers_f20_FlipperRU, LM_Flashers_f20_FlipperSpL1, LM_Flashers_f20_FlipperSpR1U, LM_Flashers_f20_FlipperSpRU, LM_Flashers_f20_GMKnob, LM_Flashers_f20_Gumballs, LM_Flashers_f20_InvaderToy, LM_Flashers_f20_Over1, LM_Flashers_f20_Over2, LM_Flashers_f20_Over3, LM_Flashers_f20_Over4, LM_Flashers_f20_Over5, LM_Flashers_f20_PF, LM_Flashers_f20_PF_Upper, LM_Flashers_f20_Parts, LM_Flashers_f20_Piano, LM_Flashers_f20_Pyramid, LM_Flashers_f20_RDiv, LM_Flashers_f20_RSling1, LM_Flashers_f20_Rails, LM_Flashers_f20_Rails_DT, LM_Flashers_f20_Robot, LM_Flashers_f20_SideMod, LM_Flashers_f20_Sign_Spiral, LM_Flashers_f20_SlotMachineToy, _
  LM_Flashers_f20_SpiralToy, LM_Flashers_f20_TVtoy, LM_Flashers_f20_sw53p, LM_Flashers_f20_sw56, LM_Flashers_f20_sw65am, LM_Flashers_f20_sw78)
Dim BL_Flashers_f28: BL_Flashers_f28=Array(LM_Flashers_f28_B_Caps_Bot_Ambe, LM_Flashers_f28_B_Caps_Top_Ambe, LM_Flashers_f28_B_Caps_Top_Yell, LM_Flashers_f28_BR2, LM_Flashers_f28_Camera, LM_Flashers_f28_Clock_Color, LM_Flashers_f28_Clock_White, LM_Flashers_f28_ClockLarge, LM_Flashers_f28_ClockToy, LM_Flashers_f28_DiverterP, LM_Flashers_f28_DiverterP1, LM_Flashers_f28_FlipperL, LM_Flashers_f28_FlipperL1, LM_Flashers_f28_FlipperLU, LM_Flashers_f28_FlipperR, LM_Flashers_f28_FlipperR1, LM_Flashers_f28_FlipperR1U, LM_Flashers_f28_FlipperRU, LM_Flashers_f28_FlipperSpL, LM_Flashers_f28_FlipperSpL1, LM_Flashers_f28_FlipperSpLU, LM_Flashers_f28_FlipperSpR, LM_Flashers_f28_FlipperSpR1, LM_Flashers_f28_FlipperSpR1U, LM_Flashers_f28_FlipperSpRU, LM_Flashers_f28_GMKnob, LM_Flashers_f28_Gumballs, LM_Flashers_f28_LSling1, LM_Flashers_f28_Over1, LM_Flashers_f28_Over2, LM_Flashers_f28_Over3, LM_Flashers_f28_Over4, LM_Flashers_f28_Over5, LM_Flashers_f28_PF, LM_Flashers_f28_PF_Upper, LM_Flashers_f28_Parts, _
  LM_Flashers_f28_Piano, LM_Flashers_f28_Pyramid, LM_Flashers_f28_RDiv, LM_Flashers_f28_RSling1, LM_Flashers_f28_RSling2, LM_Flashers_f28_Rails, LM_Flashers_f28_Robot, LM_Flashers_f28_RocketToy, LM_Flashers_f28_SideMod, LM_Flashers_f28_Sign_Spiral, LM_Flashers_f28_SlotMachineToy, LM_Flashers_f28_TVtoy, LM_Flashers_f28_sw11, LM_Flashers_f28_sw38, LM_Flashers_f28_sw47m, LM_Flashers_f28_sw53p, LM_Flashers_f28_sw63, LM_Flashers_f28_sw64, LM_Flashers_f28_sw64m, LM_Flashers_f28_sw65, LM_Flashers_f28_sw65a, LM_Flashers_f28_sw65am, LM_Flashers_f28_sw65m, LM_Flashers_f28_sw66, LM_Flashers_f28_sw67, LM_Flashers_f28_sw68)
Dim BL_Flashers_f37: BL_Flashers_f37=Array(LM_Flashers_f37_B_Caps_Bot_Ambe, LM_Flashers_f37_B_Caps_Top_Ambe, LM_Flashers_f37_B_Caps_Top_Yell, LM_Flashers_f37_Clock_Color, LM_Flashers_f37_Clock_White, LM_Flashers_f37_ClockToy, LM_Flashers_f37_FlipperR1, LM_Flashers_f37_FlipperR1U, LM_Flashers_f37_FlipperSpR1, LM_Flashers_f37_FlipperSpR1U, LM_Flashers_f37_InvaderToy, LM_Flashers_f37_Over1, LM_Flashers_f37_PF, LM_Flashers_f37_Parts, LM_Flashers_f37_Rails, LM_Flashers_f37_RocketToy, LM_Flashers_f37_Sign_Spiral, LM_Flashers_f37_SlotMachineToy, LM_Flashers_f37_sw63, LM_Flashers_f37_sw78)
Dim BL_Flashers_f38: BL_Flashers_f38=Array(LM_Flashers_f38_BS1, LM_Flashers_f38_Clock_Color, LM_Flashers_f38_Clock_White, LM_Flashers_f38_ClockToy, LM_Flashers_f38_FlipperL1, LM_Flashers_f38_FlipperR1, LM_Flashers_f38_FlipperR1U, LM_Flashers_f38_FlipperSpL1, LM_Flashers_f38_FlipperSpR1, LM_Flashers_f38_FlipperSpR1U, LM_Flashers_f38_GMKnob, LM_Flashers_f38_Gumballs, LM_Flashers_f38_Over1, LM_Flashers_f38_Over2, LM_Flashers_f38_Over3, LM_Flashers_f38_PF, LM_Flashers_f38_PF_Upper, LM_Flashers_f38_Parts, LM_Flashers_f38_Piano, LM_Flashers_f38_Pyramid, LM_Flashers_f38_RDiv, LM_Flashers_f38_RSling1, LM_Flashers_f38_RSling2, LM_Flashers_f38_Rails, LM_Flashers_f38_Rails_DT, LM_Flashers_f38_Robot, LM_Flashers_f38_SideMod, LM_Flashers_f38_Sign_Spiral, LM_Flashers_f38_SlotMachineToy, LM_Flashers_f38_sw56)
Dim BL_Flashers_f39: BL_Flashers_f39=Array(LM_Flashers_f39_Clock_White, LM_Flashers_f39_ClockToy, LM_Flashers_f39_FlipperL1, LM_Flashers_f39_FlipperR, LM_Flashers_f39_FlipperR1, LM_Flashers_f39_FlipperR1U, LM_Flashers_f39_FlipperRU, LM_Flashers_f39_FlipperSpL1, LM_Flashers_f39_FlipperSpR, LM_Flashers_f39_FlipperSpR1, LM_Flashers_f39_FlipperSpR1U, LM_Flashers_f39_FlipperSpRU, LM_Flashers_f39_GMKnob, LM_Flashers_f39_Gumballs, LM_Flashers_f39_Over1, LM_Flashers_f39_Over2, LM_Flashers_f39_Over3, LM_Flashers_f39_Over4, LM_Flashers_f39_PF, LM_Flashers_f39_PF_Upper, LM_Flashers_f39_Parts, LM_Flashers_f39_Piano, LM_Flashers_f39_Pyramid, LM_Flashers_f39_RDiv, LM_Flashers_f39_RSling1, LM_Flashers_f39_RSling2, LM_Flashers_f39_Rails, LM_Flashers_f39_Rails_DT, LM_Flashers_f39_Robot, LM_Flashers_f39_SLING1, LM_Flashers_f39_SideMod, LM_Flashers_f39_Sign_Spiral, LM_Flashers_f39_SlotMachineToy, LM_Flashers_f39_sw12, LM_Flashers_f39_sw56)
Dim BL_Flashers_f40: BL_Flashers_f40=Array(LM_Flashers_f40_ClockToy, LM_Flashers_f40_FlipperL1, LM_Flashers_f40_GMKnob, LM_Flashers_f40_Gumballs, LM_Flashers_f40_Over2, LM_Flashers_f40_Over3, LM_Flashers_f40_PF, LM_Flashers_f40_PF_Upper, LM_Flashers_f40_Parts, LM_Flashers_f40_RDiv, LM_Flashers_f40_Rails, LM_Flashers_f40_Robot, LM_Flashers_f40_SideMod, LM_Flashers_f40_Sign_Spiral)
Dim BL_Flashers_f41: BL_Flashers_f41=Array(LM_Flashers_f41_Clock_Color, LM_Flashers_f41_Clock_White, LM_Flashers_f41_ClockToy, LM_Flashers_f41_DiverterP, LM_Flashers_f41_DiverterP1, LM_Flashers_f41_FlipperL1, LM_Flashers_f41_FlipperR1, LM_Flashers_f41_FlipperR1U, LM_Flashers_f41_FlipperRU, LM_Flashers_f41_FlipperSpL1, LM_Flashers_f41_FlipperSpR1, LM_Flashers_f41_FlipperSpR1U, LM_Flashers_f41_GMKnob, LM_Flashers_f41_Over1, LM_Flashers_f41_Over2, LM_Flashers_f41_Over3, LM_Flashers_f41_Over4, LM_Flashers_f41_PF, LM_Flashers_f41_PF_Upper, LM_Flashers_f41_Parts, LM_Flashers_f41_Piano, LM_Flashers_f41_Pyramid, LM_Flashers_f41_RDiv, LM_Flashers_f41_Rails, LM_Flashers_f41_Rails_DT, LM_Flashers_f41_Robot, LM_Flashers_f41_RocketToy, LM_Flashers_f41_SideMod, LM_Flashers_f41_SlotMachineToy, LM_Flashers_f41_TVtoy, LM_Flashers_f41_URMagnet, LM_Flashers_f41_sw54p)
Dim BL_Flashers_l74: BL_Flashers_l74=Array(LM_Flashers_l74_Clock_Color, LM_Flashers_l74_Clock_White, LM_Flashers_l74_ClockShort, LM_Flashers_l74_ClockToy, LM_Flashers_l74_Over1, LM_Flashers_l74_Over2, LM_Flashers_l74_Over3, LM_Flashers_l74_PF, LM_Flashers_l74_Parts, LM_Flashers_l74_Piano, LM_Flashers_l74_Robot, LM_Flashers_l74_sw64, LM_Flashers_l74_sw66)
Dim BL_GI_Clock_l102: BL_GI_Clock_l102=Array(LM_GI_Clock_l102_B_Caps_Bot_Red, LM_GI_Clock_l102_B_Caps_Top_Amb, LM_GI_Clock_l102_B_Caps_Top_Yel, LM_GI_Clock_l102_Clock_Color, LM_GI_Clock_l102_Clock_White, LM_GI_Clock_l102_ClockLarge, LM_GI_Clock_l102_ClockShort, LM_GI_Clock_l102_ClockToy, LM_GI_Clock_l102_DiverterP, LM_GI_Clock_l102_FlipperL1, LM_GI_Clock_l102_FlipperR1U, LM_GI_Clock_l102_Over1, LM_GI_Clock_l102_Over2, LM_GI_Clock_l102_Over3, LM_GI_Clock_l102_PF, LM_GI_Clock_l102_Parts, LM_GI_Clock_l102_Piano, LM_GI_Clock_l102_Rails, LM_GI_Clock_l102_SideMod, LM_GI_Clock_l102_Sign_Spiral, LM_GI_Clock_l102_SlotMachineToy, LM_GI_Clock_l102_sw63, LM_GI_Clock_l102_sw64, LM_GI_Clock_l102_sw64m, LM_GI_Clock_l102_sw65, LM_GI_Clock_l102_sw65a, LM_GI_Clock_l102_sw65am, LM_GI_Clock_l102_sw65m)
Dim BL_GI_Left_l100: BL_GI_Left_l100=Array(LM_GI_Left_l100_B_Caps_Bot_Red, LM_GI_Left_l100_B_Caps_Top_Ambe, LM_GI_Left_l100_B_Caps_Top_Red, LM_GI_Left_l100_BR1, LM_GI_Left_l100_BR2, LM_GI_Left_l100_BS1, LM_GI_Left_l100_BS2, LM_GI_Left_l100_BS3, LM_GI_Left_l100_BumperPegs, LM_GI_Left_l100_ClockToy, LM_GI_Left_l100_FlipperR1U, LM_GI_Left_l100_GMKnob, LM_GI_Left_l100_Gumballs, LM_GI_Left_l100_Over1, LM_GI_Left_l100_Over2, LM_GI_Left_l100_Over3, LM_GI_Left_l100_Over4, LM_GI_Left_l100_Over5, LM_GI_Left_l100_PF, LM_GI_Left_l100_PF_Upper, LM_GI_Left_l100_Parts, LM_GI_Left_l100_RDiv, LM_GI_Left_l100_Rails, LM_GI_Left_l100_Robot, LM_GI_Left_l100_SideMod, LM_GI_Left_l100_Sign_Spiral, LM_GI_Left_l100_SlotMachineToy, LM_GI_Left_l100_TownSquarePost, LM_GI_Left_l100_sw36, LM_GI_Left_l100_sw37, LM_GI_Left_l100_sw38, LM_GI_Left_l100_sw53p, LM_GI_Left_l100_sw56, LM_GI_Left_l100_sw68, LM_GI_Left_l100_sw77)
Dim BL_GI_Left_l100a: BL_GI_Left_l100a=Array(LM_GI_Left_l100a_B_Caps_Top_Yel, LM_GI_Left_l100a_ClockToy, LM_GI_Left_l100a_FlipperL, LM_GI_Left_l100a_FlipperLU, LM_GI_Left_l100a_FlipperR, LM_GI_Left_l100a_FlipperR1, LM_GI_Left_l100a_FlipperRU, LM_GI_Left_l100a_FlipperSpL, LM_GI_Left_l100a_FlipperSpLU, LM_GI_Left_l100a_FlipperSpR, LM_GI_Left_l100a_FlipperSpRU, LM_GI_Left_l100a_LSling1, LM_GI_Left_l100a_LSling2, LM_GI_Left_l100a_MysticSeerToy, LM_GI_Left_l100a_Outlanes, LM_GI_Left_l100a_PF, LM_GI_Left_l100a_Parts, LM_GI_Left_l100a_RSling1, LM_GI_Left_l100a_RSling2, LM_GI_Left_l100a_Rails, LM_GI_Left_l100a_RocketToy, LM_GI_Left_l100a_SLING2, LM_GI_Left_l100a_sw36, LM_GI_Left_l100a_sw38, LM_GI_Left_l100a_sw48, LM_GI_Left_l100a_sw48m, LM_GI_Left_l100a_sw77)
Dim BL_GI_Left_l100b: BL_GI_Left_l100b=Array(LM_GI_Left_l100b_B_Caps_Top_Amb, LM_GI_Left_l100b_BR2, LM_GI_Left_l100b_BS3, LM_GI_Left_l100b_ClockToy, LM_GI_Left_l100b_FlipperL, LM_GI_Left_l100b_FlipperLU, LM_GI_Left_l100b_FlipperR, LM_GI_Left_l100b_FlipperR1U, LM_GI_Left_l100b_FlipperSpL, LM_GI_Left_l100b_FlipperSpLU, LM_GI_Left_l100b_LSling1, LM_GI_Left_l100b_LSling2, LM_GI_Left_l100b_MysticSeerToy, LM_GI_Left_l100b_Outlanes, LM_GI_Left_l100b_PF, LM_GI_Left_l100b_PF_Upper, LM_GI_Left_l100b_Parts, LM_GI_Left_l100b_SLING2, LM_GI_Left_l100b_SideMod, LM_GI_Left_l100b_SlotMachineToy, LM_GI_Left_l100b_TownSquarePost, LM_GI_Left_l100b_sw36, LM_GI_Left_l100b_sw37, LM_GI_Left_l100b_sw38, LM_GI_Left_l100b_sw48, LM_GI_Left_l100b_sw48m, LM_GI_Left_l100b_sw77)
Dim BL_GI_Left_l100c: BL_GI_Left_l100c=Array(LM_GI_Left_l100c_B_Caps_Top_Amb, LM_GI_Left_l100c_B_Caps_Top_Red, LM_GI_Left_l100c_B_Caps_Top_Yel, LM_GI_Left_l100c_BR3, LM_GI_Left_l100c_FlipperLU, LM_GI_Left_l100c_FlipperR, LM_GI_Left_l100c_FlipperSpLU, LM_GI_Left_l100c_PF, LM_GI_Left_l100c_Parts, LM_GI_Left_l100c_SLING2, LM_GI_Left_l100c_SideMod, LM_GI_Left_l100c_sw36, LM_GI_Left_l100c_sw37, LM_GI_Left_l100c_sw38, LM_GI_Left_l100c_sw48)
Dim BL_GI_Left_l100d: BL_GI_Left_l100d=Array(LM_GI_Left_l100d_B_Caps_Top_Amb, LM_GI_Left_l100d_B_Caps_Top_Yel, LM_GI_Left_l100d_FlipperLU, LM_GI_Left_l100d_FlipperR, LM_GI_Left_l100d_FlipperRU, LM_GI_Left_l100d_FlipperSpLU, LM_GI_Left_l100d_PF, LM_GI_Left_l100d_Parts, LM_GI_Left_l100d_SLING2, LM_GI_Left_l100d_sw36, LM_GI_Left_l100d_sw37, LM_GI_Left_l100d_sw38, LM_GI_Left_l100d_sw48, LM_GI_Left_l100d_sw48m)
Dim BL_GI_Left_l100e: BL_GI_Left_l100e=Array(LM_GI_Left_l100e_B_Caps_Top_Yel, LM_GI_Left_l100e_BR3, LM_GI_Left_l100e_FlipperL, LM_GI_Left_l100e_FlipperLU, LM_GI_Left_l100e_FlipperR, LM_GI_Left_l100e_FlipperRU, LM_GI_Left_l100e_FlipperSpL, LM_GI_Left_l100e_FlipperSpLU, LM_GI_Left_l100e_FlipperSpR, LM_GI_Left_l100e_FlipperSpRU, LM_GI_Left_l100e_Outlanes, LM_GI_Left_l100e_PF, LM_GI_Left_l100e_Parts, LM_GI_Left_l100e_RocketToy, LM_GI_Left_l100e_SLING2, LM_GI_Left_l100e_sw38, LM_GI_Left_l100e_sw48)
Dim BL_GI_MiniPF_l101: BL_GI_MiniPF_l101=Array(LM_GI_MiniPF_l101_GMKnob, LM_GI_MiniPF_l101_Gumballs, LM_GI_MiniPF_l101_PF, LM_GI_MiniPF_l101_PF_Upper, LM_GI_MiniPF_l101_Parts, LM_GI_MiniPF_l101_Pyramid, LM_GI_MiniPF_l101_SideMod, LM_GI_MiniPF_l101_Sign_Spiral, LM_GI_MiniPF_l101_SpiralToy)
Dim BL_GI_Right_l104: BL_GI_Right_l104=Array(LM_GI_Right_l104_B_Caps_Top_Amb, LM_GI_Right_l104_BR2, LM_GI_Right_l104_BS1, LM_GI_Right_l104_Clock_Color, LM_GI_Right_l104_Clock_White, LM_GI_Right_l104_ClockToy, LM_GI_Right_l104_DiverterP, LM_GI_Right_l104_FlipperL1, LM_GI_Right_l104_FlipperR1, LM_GI_Right_l104_FlipperR1U, LM_GI_Right_l104_FlipperSpL1, LM_GI_Right_l104_FlipperSpR, LM_GI_Right_l104_FlipperSpR1, LM_GI_Right_l104_FlipperSpR1U, LM_GI_Right_l104_FlipperSpRU, LM_GI_Right_l104_Gate1, LM_GI_Right_l104_InvaderToy, LM_GI_Right_l104_Over1, LM_GI_Right_l104_Over2, LM_GI_Right_l104_Over3, LM_GI_Right_l104_PF, LM_GI_Right_l104_PF_Upper, LM_GI_Right_l104_Parts, LM_GI_Right_l104_Piano, LM_GI_Right_l104_Rails, LM_GI_Right_l104_Robot, LM_GI_Right_l104_RocketToy, LM_GI_Right_l104_SideMod, LM_GI_Right_l104_Sign_Spiral, LM_GI_Right_l104_SlotMachineToy, LM_GI_Right_l104_TVtoy, LM_GI_Right_l104_URMagnet, LM_GI_Right_l104_sw47, LM_GI_Right_l104_sw61, LM_GI_Right_l104_sw62, LM_GI_Right_l104_sw63, LM_GI_Right_l104_sw64m, _
  LM_GI_Right_l104_sw65, LM_GI_Right_l104_sw65a, LM_GI_Right_l104_sw65am, LM_GI_Right_l104_sw65m, LM_GI_Right_l104_sw66, LM_GI_Right_l104_sw66m, LM_GI_Right_l104_sw67, LM_GI_Right_l104_sw67m, LM_GI_Right_l104_sw68m, LM_GI_Right_l104_sw78)
Dim BL_GI_Right_l104a: BL_GI_Right_l104a=Array(LM_GI_Right_l104a_B_Caps_Top_Re, LM_GI_Right_l104a_B_Caps_Top_Ye, LM_GI_Right_l104a_BR3, LM_GI_Right_l104a_ClockToy, LM_GI_Right_l104a_FlipperL, LM_GI_Right_l104a_FlipperLU, LM_GI_Right_l104a_FlipperR, LM_GI_Right_l104a_FlipperR1U, LM_GI_Right_l104a_FlipperRU, LM_GI_Right_l104a_FlipperSpL, LM_GI_Right_l104a_FlipperSpLU, LM_GI_Right_l104a_FlipperSpR, LM_GI_Right_l104a_FlipperSpR1U, LM_GI_Right_l104a_FlipperSpRU, LM_GI_Right_l104a_InvaderToy, LM_GI_Right_l104a_LSling1, LM_GI_Right_l104a_LSling2, LM_GI_Right_l104a_Outlanes, LM_GI_Right_l104a_Over3, LM_GI_Right_l104a_PF, LM_GI_Right_l104a_Parts, LM_GI_Right_l104a_RSling1, LM_GI_Right_l104a_RSling2, LM_GI_Right_l104a_Rails, LM_GI_Right_l104a_RocketToy, LM_GI_Right_l104a_SLING1, LM_GI_Right_l104a_sw11, LM_GI_Right_l104a_sw12, LM_GI_Right_l104a_sw48)
Dim BL_GI_Right_l104b: BL_GI_Right_l104b=Array(LM_GI_Right_l104b_ClockToy, LM_GI_Right_l104b_PF, LM_GI_Right_l104b_Parts, LM_GI_Right_l104b_Robot, LM_GI_Right_l104b_Sign_Spiral, LM_GI_Right_l104b_SlotMachineTo, LM_GI_Right_l104b_sw65m, LM_GI_Right_l104b_sw66, LM_GI_Right_l104b_sw66m, LM_GI_Right_l104b_sw67, LM_GI_Right_l104b_sw67m)
Dim BL_GI_Right_l104c: BL_GI_Right_l104c=Array(LM_GI_Right_l104c_B_Caps_Top_Am, LM_GI_Right_l104c_B_Caps_Top_Ye, LM_GI_Right_l104c_BR2, LM_GI_Right_l104c_BR3, LM_GI_Right_l104c_FlipperL, LM_GI_Right_l104c_FlipperLU, LM_GI_Right_l104c_FlipperR, LM_GI_Right_l104c_FlipperRU, LM_GI_Right_l104c_FlipperSpL, LM_GI_Right_l104c_FlipperSpLU, LM_GI_Right_l104c_FlipperSpR, LM_GI_Right_l104c_FlipperSpRU, LM_GI_Right_l104c_Outlanes, LM_GI_Right_l104c_PF, LM_GI_Right_l104c_Parts, LM_GI_Right_l104c_Rails, LM_GI_Right_l104c_SLING1, LM_GI_Right_l104c_sw11)
Dim BL_GI_Right_l104d: BL_GI_Right_l104d=Array(LM_GI_Right_l104d_FlipperL, LM_GI_Right_l104d_FlipperLU, LM_GI_Right_l104d_FlipperR, LM_GI_Right_l104d_FlipperR1U, LM_GI_Right_l104d_FlipperRU, LM_GI_Right_l104d_FlipperSpL, LM_GI_Right_l104d_FlipperSpLU, LM_GI_Right_l104d_FlipperSpR, LM_GI_Right_l104d_FlipperSpR1U, LM_GI_Right_l104d_FlipperSpRU, LM_GI_Right_l104d_InvaderToy, LM_GI_Right_l104d_Outlanes, LM_GI_Right_l104d_PF, LM_GI_Right_l104d_Parts, LM_GI_Right_l104d_Rails, LM_GI_Right_l104d_RocketToy, LM_GI_Right_l104d_SLING1, LM_GI_Right_l104d_sw11, LM_GI_Right_l104d_sw12)
Dim BL_Inserts_l11: BL_Inserts_l11=Array(LM_Inserts_l11_B_Caps_Top_Yello, LM_Inserts_l11_FlipperLU, LM_Inserts_l11_FlipperRU, LM_Inserts_l11_LSling1, LM_Inserts_l11_LSling2, LM_Inserts_l11_Parts, LM_Inserts_l11_RSling1)
Dim BL_Inserts_l12: BL_Inserts_l12=Array(LM_Inserts_l12_B_Caps_Top_Yello, LM_Inserts_l12_FlipperLU, LM_Inserts_l12_LSling1, LM_Inserts_l12_LSling2, LM_Inserts_l12_Parts, LM_Inserts_l12_SLING2)
Dim BL_Inserts_l13: BL_Inserts_l13=Array(LM_Inserts_l13_B_Caps_Top_Amber, LM_Inserts_l13_B_Caps_Top_Yello, LM_Inserts_l13_LSling1, LM_Inserts_l13_LSling2, LM_Inserts_l13_Parts, LM_Inserts_l13_sw48)
Dim BL_Inserts_l14: BL_Inserts_l14=Array(LM_Inserts_l14_B_Caps_Top_Red, LM_Inserts_l14_B_Caps_Top_Yello, LM_Inserts_l14_LSling2, LM_Inserts_l14_Parts, LM_Inserts_l14_sw48, LM_Inserts_l14_sw48m, LM_Inserts_l14_sw77)
Dim BL_Inserts_l15: BL_Inserts_l15=Array(LM_Inserts_l15_B_Caps_Top_Yello, LM_Inserts_l15_PF, LM_Inserts_l15_PF_Upper, LM_Inserts_l15_Parts, LM_Inserts_l15_sw48, LM_Inserts_l15_sw48m, LM_Inserts_l15_sw77)
Dim BL_Inserts_l16: BL_Inserts_l16=Array(LM_Inserts_l16_B_Caps_Top_Amber, LM_Inserts_l16_B_Caps_Top_Red, LM_Inserts_l16_B_Caps_Top_Yello, LM_Inserts_l16_PF, LM_Inserts_l16_PF_Upper, LM_Inserts_l16_Parts, LM_Inserts_l16_sw77, LM_Inserts_l16_sw77m, LM_Inserts_l16_sw78)
Dim BL_Inserts_l17: BL_Inserts_l17=Array(LM_Inserts_l17_B_Caps_Bot_Amber, LM_Inserts_l17_B_Caps_Top_Amber, LM_Inserts_l17_B_Caps_Top_Red, LM_Inserts_l17_B_Caps_Top_Yello, LM_Inserts_l17_BR2, LM_Inserts_l17_BS2, LM_Inserts_l17_PF_Upper, LM_Inserts_l17_Parts, LM_Inserts_l17_sw77, LM_Inserts_l17_sw77m, LM_Inserts_l17_sw78, LM_Inserts_l17_sw78m)
Dim BL_Inserts_l18: BL_Inserts_l18=Array(LM_Inserts_l18_B_Caps_Top_Amber, LM_Inserts_l18_PF_Upper, LM_Inserts_l18_Parts, LM_Inserts_l18_sw77, LM_Inserts_l18_sw78)
Dim BL_Inserts_l21: BL_Inserts_l21=Array(LM_Inserts_l21_B_Caps_Top_Yello, LM_Inserts_l21_Parts)
Dim BL_Inserts_l22: BL_Inserts_l22=Array(LM_Inserts_l22_B_Caps_Top_Yello, LM_Inserts_l22_FlipperRU, LM_Inserts_l22_Parts, LM_Inserts_l22_RSling1, LM_Inserts_l22_RSling2)
Dim BL_Inserts_l23: BL_Inserts_l23=Array(LM_Inserts_l23_B_Caps_Top_Yello, LM_Inserts_l23_Parts, LM_Inserts_l23_RSling1, LM_Inserts_l23_RSling2, LM_Inserts_l23_SLING1)
Dim BL_Inserts_l24: BL_Inserts_l24=Array(LM_Inserts_l24_B_Caps_Top_Amber, LM_Inserts_l24_B_Caps_Top_Yello, LM_Inserts_l24_Parts, LM_Inserts_l24_RSling2)
Dim BL_Inserts_l25: BL_Inserts_l25=Array(LM_Inserts_l25_B_Caps_Top_Amber, LM_Inserts_l25_B_Caps_Top_Yello, LM_Inserts_l25_FlipperR1, LM_Inserts_l25_FlipperR1U, LM_Inserts_l25_PF, LM_Inserts_l25_Parts)
Dim BL_Inserts_l26: BL_Inserts_l26=Array(LM_Inserts_l26_B_Caps_Top_Amber, LM_Inserts_l26_B_Caps_Top_Red, LM_Inserts_l26_FlipperR1U, LM_Inserts_l26_FlipperSpR1U, LM_Inserts_l26_Parts)
Dim BL_Inserts_l27: BL_Inserts_l27=Array(LM_Inserts_l27_B_Caps_Top_Amber, LM_Inserts_l27_FlipperR1U, LM_Inserts_l27_FlipperSpR1U, LM_Inserts_l27_PF, LM_Inserts_l27_Parts)
Dim BL_Inserts_l28: BL_Inserts_l28=Array(LM_Inserts_l28_B_Caps_Top_Amber, LM_Inserts_l28_B_Caps_Top_Yello, LM_Inserts_l28_Parts, LM_Inserts_l28_SlotMachineToy)
Dim BL_Inserts_l31: BL_Inserts_l31=Array(LM_Inserts_l31_B_Caps_Top_Red, LM_Inserts_l31_B_Caps_Top_Yello, LM_Inserts_l31_Outlanes, LM_Inserts_l31_Parts, LM_Inserts_l31_SideMod, LM_Inserts_l31_sw36)
Dim BL_Inserts_l32: BL_Inserts_l32=Array(LM_Inserts_l32_B_Caps_Top_Amber, LM_Inserts_l32_B_Caps_Top_Yello, LM_Inserts_l32_Parts)
Dim BL_Inserts_l33: BL_Inserts_l33=Array(LM_Inserts_l33_B_Caps_Top_Yello, LM_Inserts_l33_MysticSeerToy, LM_Inserts_l33_Outlanes, LM_Inserts_l33_PF, LM_Inserts_l33_Parts, LM_Inserts_l33_sw48)
Dim BL_Inserts_l34: BL_Inserts_l34=Array(LM_Inserts_l34_B_Caps_Top_Red, LM_Inserts_l34_B_Caps_Top_Yello, LM_Inserts_l34_PF, LM_Inserts_l34_Parts, LM_Inserts_l34_sw48, LM_Inserts_l34_sw77)
Dim BL_Inserts_l35: BL_Inserts_l35=Array(LM_Inserts_l35_B_Caps_Top_Yello, LM_Inserts_l35_Outlanes, LM_Inserts_l35_PF, LM_Inserts_l35_Parts, LM_Inserts_l35_sw48, LM_Inserts_l35_sw48m)
Dim BL_Inserts_l36: BL_Inserts_l36=Array(LM_Inserts_l36_B_Caps_Top_Red, LM_Inserts_l36_B_Caps_Top_Yello, LM_Inserts_l36_Parts, LM_Inserts_l36_sw77, LM_Inserts_l36_sw78)
Dim BL_Inserts_l37: BL_Inserts_l37=Array(LM_Inserts_l37_B_Caps_Top_Yello, LM_Inserts_l37_BR3, LM_Inserts_l37_BS3, LM_Inserts_l37_Outlanes, LM_Inserts_l37_PF, LM_Inserts_l37_Parts, LM_Inserts_l37_sw48, LM_Inserts_l37_sw48m, LM_Inserts_l37_sw77)
Dim BL_Inserts_l38: BL_Inserts_l38=Array(LM_Inserts_l38_B_Caps_Top_Amber, LM_Inserts_l38_BR3, LM_Inserts_l38_BS3, LM_Inserts_l38_BumperPegs, LM_Inserts_l38_PF_Upper, LM_Inserts_l38_Parts, LM_Inserts_l38_TownSquarePost, LM_Inserts_l38_sw48m, LM_Inserts_l38_sw77, LM_Inserts_l38_sw77m)
Dim BL_Inserts_l41: BL_Inserts_l41=Array(LM_Inserts_l41_FlipperL, LM_Inserts_l41_FlipperLU, LM_Inserts_l41_FlipperSpL, LM_Inserts_l41_FlipperSpLU, LM_Inserts_l41_LSling1, LM_Inserts_l41_LSling2, LM_Inserts_l41_PF, LM_Inserts_l41_Parts, LM_Inserts_l41_SLING2)
Dim BL_Inserts_l42: BL_Inserts_l42=Array(LM_Inserts_l42_B_Caps_Top_Yello, LM_Inserts_l42_FlipperL, LM_Inserts_l42_FlipperLU, LM_Inserts_l42_FlipperR, LM_Inserts_l42_FlipperRU, LM_Inserts_l42_FlipperSpL, LM_Inserts_l42_FlipperSpLU, LM_Inserts_l42_FlipperSpRU, LM_Inserts_l42_LSling1, LM_Inserts_l42_LSling2, LM_Inserts_l42_PF, LM_Inserts_l42_Parts)
Dim BL_Inserts_l43: BL_Inserts_l43=Array(LM_Inserts_l43_B_Caps_Top_Yello, LM_Inserts_l43_FlipperL, LM_Inserts_l43_FlipperLU, LM_Inserts_l43_FlipperR, LM_Inserts_l43_FlipperRU, LM_Inserts_l43_FlipperSpL, LM_Inserts_l43_FlipperSpLU, LM_Inserts_l43_FlipperSpRU, LM_Inserts_l43_LSling1, LM_Inserts_l43_LSling2, LM_Inserts_l43_PF, LM_Inserts_l43_Parts, LM_Inserts_l43_RSling1)
Dim BL_Inserts_l44: BL_Inserts_l44=Array(LM_Inserts_l44_B_Caps_Top_Yello, LM_Inserts_l44_FlipperL, LM_Inserts_l44_FlipperLU, LM_Inserts_l44_FlipperR, LM_Inserts_l44_FlipperRU, LM_Inserts_l44_FlipperSpLU, LM_Inserts_l44_FlipperSpR, LM_Inserts_l44_FlipperSpRU, LM_Inserts_l44_LSling1, LM_Inserts_l44_PF, LM_Inserts_l44_Parts, LM_Inserts_l44_RSling1, LM_Inserts_l44_RSling2)
Dim BL_Inserts_l45: BL_Inserts_l45=Array(LM_Inserts_l45_B_Caps_Top_Yello, LM_Inserts_l45_FlipperL, LM_Inserts_l45_FlipperLU, LM_Inserts_l45_FlipperR, LM_Inserts_l45_FlipperRU, LM_Inserts_l45_FlipperSpLU, LM_Inserts_l45_FlipperSpR, LM_Inserts_l45_FlipperSpRU, LM_Inserts_l45_PF, LM_Inserts_l45_Parts, LM_Inserts_l45_RSling1, LM_Inserts_l45_RSling2, LM_Inserts_l45_SLING1)
Dim BL_Inserts_l46: BL_Inserts_l46=Array(LM_Inserts_l46_FlipperR, LM_Inserts_l46_FlipperRU, LM_Inserts_l46_FlipperSpR, LM_Inserts_l46_FlipperSpRU, LM_Inserts_l46_PF, LM_Inserts_l46_Parts)
Dim BL_Inserts_l47: BL_Inserts_l47=Array(LM_Inserts_l47_FlipperL, LM_Inserts_l47_FlipperLU, LM_Inserts_l47_FlipperR, LM_Inserts_l47_FlipperRU, LM_Inserts_l47_FlipperSpL, LM_Inserts_l47_FlipperSpLU, LM_Inserts_l47_FlipperSpR, LM_Inserts_l47_FlipperSpRU, LM_Inserts_l47_PF, LM_Inserts_l47_Parts)
Dim BL_Inserts_l48: BL_Inserts_l48=Array(LM_Inserts_l48_FlipperR1U, LM_Inserts_l48_Outlanes, LM_Inserts_l48_PF, LM_Inserts_l48_Parts, LM_Inserts_l48_RSling2, LM_Inserts_l48_RocketToy)
Dim BL_Inserts_l51: BL_Inserts_l51=Array(LM_Inserts_l51_B_Caps_Top_Amber, LM_Inserts_l51_Clock_Color, LM_Inserts_l51_Clock_White, LM_Inserts_l51_ClockToy, LM_Inserts_l51_Parts, LM_Inserts_l51_Robot, LM_Inserts_l51_Sign_Spiral, LM_Inserts_l51_sw47m, LM_Inserts_l51_sw68)
Dim BL_Inserts_l52: BL_Inserts_l52=Array(LM_Inserts_l52_B_Caps_Bot_Amber, LM_Inserts_l52_B_Caps_Top_Amber, LM_Inserts_l52_Clock_Color, LM_Inserts_l52_Clock_White, LM_Inserts_l52_ClockToy, LM_Inserts_l52_FlipperL1, LM_Inserts_l52_PF_Upper, LM_Inserts_l52_Parts, LM_Inserts_l52_Sign_Spiral, LM_Inserts_l52_sw68)
Dim BL_Inserts_l53: BL_Inserts_l53=Array(LM_Inserts_l53_B_Caps_Bot_Amber, LM_Inserts_l53_B_Caps_Top_Amber, LM_Inserts_l53_Clock_Color, LM_Inserts_l53_Clock_White, LM_Inserts_l53_ClockToy, LM_Inserts_l53_FlipperL1, LM_Inserts_l53_Over3, LM_Inserts_l53_PF_Upper, LM_Inserts_l53_Parts, LM_Inserts_l53_Sign_Spiral)
Dim BL_Inserts_l54: BL_Inserts_l54=Array(LM_Inserts_l54_Clock_White, LM_Inserts_l54_ClockToy, LM_Inserts_l54_PF, LM_Inserts_l54_PF_Upper, LM_Inserts_l54_Parts, LM_Inserts_l54_Sign_Spiral)
Dim BL_Inserts_l55: BL_Inserts_l55=Array(LM_Inserts_l55_B_Caps_Top_Amber, LM_Inserts_l55_Clock_Color, LM_Inserts_l55_Clock_White, LM_Inserts_l55_ClockToy, LM_Inserts_l55_FlipperL1, LM_Inserts_l55_FlipperSpL1, LM_Inserts_l55_PF_Upper, LM_Inserts_l55_Parts, LM_Inserts_l55_Sign_Spiral)
Dim BL_Inserts_l56: BL_Inserts_l56=Array(LM_Inserts_l56_B_Caps_Bot_Amber, LM_Inserts_l56_Clock_Color, LM_Inserts_l56_Clock_White, LM_Inserts_l56_ClockToy, LM_Inserts_l56_Over3, LM_Inserts_l56_PF, LM_Inserts_l56_Parts, LM_Inserts_l56_Piano, LM_Inserts_l56_Sign_Spiral, LM_Inserts_l56_sw64, LM_Inserts_l56_sw68)
Dim BL_Inserts_l57: BL_Inserts_l57=Array(LM_Inserts_l57_Over2, LM_Inserts_l57_Over3, LM_Inserts_l57_PF, LM_Inserts_l57_Parts, LM_Inserts_l57_sw64)
Dim BL_Inserts_l58: BL_Inserts_l58=Array(LM_Inserts_l58_Clock_Color, LM_Inserts_l58_Clock_White, LM_Inserts_l58_Over2, LM_Inserts_l58_Parts)
Dim BL_Inserts_l64: BL_Inserts_l64=Array(LM_Inserts_l64_B_Caps_Bot_Red, LM_Inserts_l64_B_Caps_Top_Amber, LM_Inserts_l64_B_Caps_Top_Yello, LM_Inserts_l64_BR2, LM_Inserts_l64_BS2, LM_Inserts_l64_BumperPegs, LM_Inserts_l64_PF, LM_Inserts_l64_PF_Upper, LM_Inserts_l64_Parts, LM_Inserts_l64_TownSquarePost, LM_Inserts_l64_sw77, LM_Inserts_l64_sw77m, LM_Inserts_l64_sw78m)
Dim BL_Inserts_l65: BL_Inserts_l65=Array(LM_Inserts_l65_B_Caps_Top_Amber, LM_Inserts_l65_B_Caps_Top_Red, LM_Inserts_l65_B_Caps_Top_Yello, LM_Inserts_l65_BR2, LM_Inserts_l65_BS2, LM_Inserts_l65_PF, LM_Inserts_l65_PF_Upper, LM_Inserts_l65_Parts, LM_Inserts_l65_sw77, LM_Inserts_l65_sw77m, LM_Inserts_l65_sw78, LM_Inserts_l65_sw78m)
Dim BL_Inserts_l66: BL_Inserts_l66=Array(LM_Inserts_l66_Outlanes, LM_Inserts_l66_Parts, LM_Inserts_l66_RocketToy, LM_Inserts_l66_sw11)
Dim BL_Inserts_l67: BL_Inserts_l67=Array(LM_Inserts_l67_Clock_Color, LM_Inserts_l67_Clock_White, LM_Inserts_l67_Over1, LM_Inserts_l67_Parts, LM_Inserts_l67_SlotMachineToy)
Dim BL_Inserts_l68: BL_Inserts_l68=Array(LM_Inserts_l68_Clock_Color, LM_Inserts_l68_Clock_White, LM_Inserts_l68_Over1, LM_Inserts_l68_Over3, LM_Inserts_l68_Parts)
Dim BL_Inserts_l71: BL_Inserts_l71=Array(LM_Inserts_l71_B_Caps_Top_Amber, LM_Inserts_l71_B_Caps_Top_Yello, LM_Inserts_l71_FlipperR1U, LM_Inserts_l71_FlipperSpR1U, LM_Inserts_l71_PF, LM_Inserts_l71_Parts, LM_Inserts_l71_SlotMachineToy, LM_Inserts_l71_sw47, LM_Inserts_l71_sw66, LM_Inserts_l71_sw67, LM_Inserts_l71_sw67m, LM_Inserts_l71_sw68m)
Dim BL_Inserts_l72: BL_Inserts_l72=Array(LM_Inserts_l72_B_Caps_Top_Amber, LM_Inserts_l72_Clock_Color, LM_Inserts_l72_Clock_White, LM_Inserts_l72_ClockToy, LM_Inserts_l72_PF, LM_Inserts_l72_Parts, LM_Inserts_l72_Sign_Spiral, LM_Inserts_l72_sw66, LM_Inserts_l72_sw67, LM_Inserts_l72_sw68m)
Dim BL_Inserts_l73: BL_Inserts_l73=Array(LM_Inserts_l73_B_Caps_Bot_Amber, LM_Inserts_l73_B_Caps_Top_Amber, LM_Inserts_l73_Clock_Color, LM_Inserts_l73_Clock_White, LM_Inserts_l73_ClockToy, LM_Inserts_l73_PF, LM_Inserts_l73_Parts, LM_Inserts_l73_Sign_Spiral, LM_Inserts_l73_sw66, LM_Inserts_l73_sw66m, LM_Inserts_l73_sw67, LM_Inserts_l73_sw67m)
Dim BL_Inserts_l75: BL_Inserts_l75=Array(LM_Inserts_l75_B_Caps_Bot_Amber, LM_Inserts_l75_B_Caps_Top_Amber, LM_Inserts_l75_Clock_Color, LM_Inserts_l75_Clock_White, LM_Inserts_l75_Over3, LM_Inserts_l75_PF, LM_Inserts_l75_Parts, LM_Inserts_l75_Piano, LM_Inserts_l75_sw64, LM_Inserts_l75_sw64m)
Dim BL_Inserts_l76: BL_Inserts_l76=Array(LM_Inserts_l76_PF_Upper, LM_Inserts_l76_Parts, LM_Inserts_l76_Pyramid)
Dim BL_Inserts_l77: BL_Inserts_l77=Array(LM_Inserts_l77_PF_Upper, LM_Inserts_l77_Parts)
Dim BL_Inserts_l78: BL_Inserts_l78=Array(LM_Inserts_l78_PF_Upper, LM_Inserts_l78_Parts, LM_Inserts_l78_Pyramid)
Dim BL_Inserts_l81: BL_Inserts_l81=Array(LM_Inserts_l81_B_Caps_Bot_Amber, LM_Inserts_l81_ClockToy, LM_Inserts_l81_FlipperL1, LM_Inserts_l81_FlipperSpL1, LM_Inserts_l81_Over2, LM_Inserts_l81_PF, LM_Inserts_l81_PF_Upper, LM_Inserts_l81_Parts, LM_Inserts_l81_Sign_Spiral)
Dim BL_Inserts_l82: BL_Inserts_l82=Array(LM_Inserts_l82_B_Caps_Top_Amber, LM_Inserts_l82_Clock_Color, LM_Inserts_l82_Clock_White, LM_Inserts_l82_ClockToy, LM_Inserts_l82_Over3, LM_Inserts_l82_PF, LM_Inserts_l82_Parts, LM_Inserts_l82_Piano, LM_Inserts_l82_Robot, LM_Inserts_l82_Sign_Spiral, LM_Inserts_l82_sw47m, LM_Inserts_l82_sw64, LM_Inserts_l82_sw68)
Dim BL_Inserts_l83: BL_Inserts_l83=Array(LM_Inserts_l83_Clock_Color, LM_Inserts_l83_Clock_White, LM_Inserts_l83_ClockToy, LM_Inserts_l83_Over1, LM_Inserts_l83_Over2, LM_Inserts_l83_Over3, LM_Inserts_l83_PF, LM_Inserts_l83_Parts, LM_Inserts_l83_Piano, LM_Inserts_l83_sw64, LM_Inserts_l83_sw64m)
Dim BL_Inserts_l84: BL_Inserts_l84=Array(LM_Inserts_l84_Clock_Color, LM_Inserts_l84_Clock_White, LM_Inserts_l84_ClockLarge, LM_Inserts_l84_ClockShort, LM_Inserts_l84_ClockToy, LM_Inserts_l84_Over1, LM_Inserts_l84_Over2, LM_Inserts_l84_Over3, LM_Inserts_l84_PF, LM_Inserts_l84_Parts, LM_Inserts_l84_Piano, LM_Inserts_l84_sw64, LM_Inserts_l84_sw64m)
Dim BL_Inserts_l85: BL_Inserts_l85=Array(LM_Inserts_l85_Clock_Color, LM_Inserts_l85_Clock_White, LM_Inserts_l85_FlipperR1U, LM_Inserts_l85_Over2, LM_Inserts_l85_Over3, LM_Inserts_l85_PF, LM_Inserts_l85_Parts, LM_Inserts_l85_RocketToy, LM_Inserts_l85_SlotMachineToy, LM_Inserts_l85_sw65m, LM_Inserts_l85_sw67m)
Dim BL_Inserts_l86: BL_Inserts_l86=Array(LM_Inserts_l86_FlipperR1, LM_Inserts_l86_FlipperR1U, LM_Inserts_l86_FlipperSpR1, LM_Inserts_l86_FlipperSpR1U, LM_Inserts_l86_InvaderToy, LM_Inserts_l86_Over1, LM_Inserts_l86_Over3, LM_Inserts_l86_PF, LM_Inserts_l86_Parts, LM_Inserts_l86_SlotMachineToy)
Dim BL_Mod_l105: BL_Mod_l105=Array(LM_Mod_l105_InvaderToy, LM_Mod_l105_Over1, LM_Mod_l105_Parts)
Dim BL_Mod_l106: BL_Mod_l106=Array(LM_Mod_l106_FlipperR1, LM_Mod_l106_FlipperR1U, LM_Mod_l106_FlipperSpR1, LM_Mod_l106_FlipperSpR1U, LM_Mod_l106_InvaderToy, LM_Mod_l106_Over1, LM_Mod_l106_PF, LM_Mod_l106_Parts, LM_Mod_l106_SlotMachineToy)
Dim BL_Mod_l107: BL_Mod_l107=Array(LM_Mod_l107_FlipperR1, LM_Mod_l107_FlipperR1U, LM_Mod_l107_FlipperSpR1, LM_Mod_l107_FlipperSpR1U, LM_Mod_l107_InvaderToy, LM_Mod_l107_PF, LM_Mod_l107_Parts)
Dim BL_Mod_l108: BL_Mod_l108=Array(LM_Mod_l108_PF, LM_Mod_l108_Parts, LM_Mod_l108_sw68m)
Dim BL_Mod_l109: BL_Mod_l109=Array(LM_Mod_l109_B_Caps_Bot_Amber, LM_Mod_l109_B_Caps_Bot_Red, LM_Mod_l109_B_Caps_Bot_Yellow, LM_Mod_l109_B_Caps_Top_Amber, LM_Mod_l109_B_Caps_Top_Yellow, LM_Mod_l109_BR1, LM_Mod_l109_BR2, LM_Mod_l109_BR3, LM_Mod_l109_BS2, LM_Mod_l109_BumperPegs, LM_Mod_l109_PF, LM_Mod_l109_PF_Upper, LM_Mod_l109_Parts, LM_Mod_l109_TownSquarePost, LM_Mod_l109_sw48m, LM_Mod_l109_sw77, LM_Mod_l109_sw77m, LM_Mod_l109_sw78)
Dim BL_Mod_l110: BL_Mod_l110=Array(LM_Mod_l110_B_Caps_Bot_Amber, LM_Mod_l110_BR2, LM_Mod_l110_BS2, LM_Mod_l110_Camera, LM_Mod_l110_PF, LM_Mod_l110_PF_Upper, LM_Mod_l110_Parts)
Dim BL_Mod_l111: BL_Mod_l111=Array(LM_Mod_l111_BR1, LM_Mod_l111_FlipperR1, LM_Mod_l111_FlipperR1U, LM_Mod_l111_FlipperSpR1, LM_Mod_l111_FlipperSpR1U, LM_Mod_l111_Outlanes, LM_Mod_l111_PF, LM_Mod_l111_Parts, LM_Mod_l111_RSling1, LM_Mod_l111_RSling2, LM_Mod_l111_RocketToy, LM_Mod_l111_SLING1)
Dim BL_Room: BL_Room=Array(BM_B_Caps_Bot_Amber, BM_B_Caps_Bot_Red, BM_B_Caps_Bot_Yellow, BM_B_Caps_Top_Amber, BM_B_Caps_Top_Red, BM_B_Caps_Top_Yellow, BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_BumperPegs, BM_Camera, BM_Clock_Color, BM_Clock_White, BM_ClockLarge, BM_ClockShort, BM_ClockToy, BM_DiverterP, BM_DiverterP1, BM_FlipperL, BM_FlipperL1, BM_FlipperLU, BM_FlipperR, BM_FlipperR1, BM_FlipperR1U, BM_FlipperRU, BM_FlipperSpL, BM_FlipperSpL1, BM_FlipperSpLU, BM_FlipperSpR, BM_FlipperSpR1, BM_FlipperSpR1U, BM_FlipperSpRU, BM_GMKnob, BM_Gate1, BM_Gate2, BM_Gumballs, BM_InvaderToy, BM_LSling1, BM_LSling2, BM_MysticSeerToy, BM_Outlanes, BM_Over1, BM_Over2, BM_Over3, BM_Over4, BM_Over5, BM_PF, BM_PF_Upper, BM_Parts, BM_Piano, BM_Pyramid, BM_RDiv, BM_RSling1, BM_RSling2, BM_Rails, BM_Rails_DT, BM_Robot, BM_RocketToy, BM_SLING1, BM_SLING2, BM_ShooterDiv, BM_SideMod, BM_Sign_Spiral, BM_SlotMachineToy, BM_SpiralToy, BM_TVtoy, BM_TownSquarePost, BM_URMagnet, BM_sw11, BM_sw12, BM_sw27, BM_sw36, BM_sw37, _
  BM_sw38, BM_sw47, BM_sw47m, BM_sw48, BM_sw48m, BM_sw53p, BM_sw54p, BM_sw56, BM_sw61, BM_sw62, BM_sw63, BM_sw64, BM_sw64m, BM_sw65, BM_sw65a, BM_sw65am, BM_sw65m, BM_sw66, BM_sw66m, BM_sw67, BM_sw67m, BM_sw68, BM_sw68m, BM_sw77, BM_sw77m, BM_sw78, BM_sw78m)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_B_Caps_Bot_Amber, BM_B_Caps_Bot_Red, BM_B_Caps_Bot_Yellow, BM_B_Caps_Top_Amber, BM_B_Caps_Top_Red, BM_B_Caps_Top_Yellow, BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_BumperPegs, BM_Camera, BM_Clock_Color, BM_Clock_White, BM_ClockLarge, BM_ClockShort, BM_ClockToy, BM_DiverterP, BM_DiverterP1, BM_FlipperL, BM_FlipperL1, BM_FlipperLU, BM_FlipperR, BM_FlipperR1, BM_FlipperR1U, BM_FlipperRU, BM_FlipperSpL, BM_FlipperSpL1, BM_FlipperSpLU, BM_FlipperSpR, BM_FlipperSpR1, BM_FlipperSpR1U, BM_FlipperSpRU, BM_GMKnob, BM_Gate1, BM_Gate2, BM_Gumballs, BM_InvaderToy, BM_LSling1, BM_LSling2, BM_MysticSeerToy, BM_Outlanes, BM_Over1, BM_Over2, BM_Over3, BM_Over4, BM_Over5, BM_PF, BM_PF_Upper, BM_Parts, BM_Piano, BM_Pyramid, BM_RDiv, BM_RSling1, BM_RSling2, BM_Rails, BM_Rails_DT, BM_Robot, BM_RocketToy, BM_SLING1, BM_SLING2, BM_ShooterDiv, BM_SideMod, BM_Sign_Spiral, BM_SlotMachineToy, BM_SpiralToy, BM_TVtoy, BM_TownSquarePost, BM_URMagnet, BM_sw11, BM_sw12, BM_sw27, BM_sw36, BM_sw37, _
  BM_sw38, BM_sw47, BM_sw47m, BM_sw48, BM_sw48m, BM_sw53p, BM_sw54p, BM_sw56, BM_sw61, BM_sw62, BM_sw63, BM_sw64, BM_sw64m, BM_sw65, BM_sw65a, BM_sw65am, BM_sw65m, BM_sw66, BM_sw66m, BM_sw67, BM_sw67m, BM_sw68, BM_sw68m, BM_sw77, BM_sw77m, BM_sw78, BM_sw78m)
Dim BG_Lightmap: BG_Lightmap=Array(LM_Bumpers_l61_B_Caps_Bot_Amber, LM_Bumpers_l61_B_Caps_Bot_Red, LM_Bumpers_l61_B_Caps_Bot_Yello, LM_Bumpers_l61_B_Caps_Top_Amber, LM_Bumpers_l61_B_Caps_Top_Red, LM_Bumpers_l61_B_Caps_Top_Yello, LM_Bumpers_l61_BR1, LM_Bumpers_l61_BR2, LM_Bumpers_l61_BR3, LM_Bumpers_l61_BS1, LM_Bumpers_l61_BS2, LM_Bumpers_l61_BS3, LM_Bumpers_l61_BumperPegs, LM_Bumpers_l61_ClockToy, LM_Bumpers_l61_MysticSeerToy, LM_Bumpers_l61_Over3, LM_Bumpers_l61_PF, LM_Bumpers_l61_PF_Upper, LM_Bumpers_l61_Parts, LM_Bumpers_l61_Rails, LM_Bumpers_l61_SideMod, LM_Bumpers_l61_SlotMachineToy, LM_Bumpers_l61_sw36, LM_Bumpers_l61_sw37, LM_Bumpers_l62_B_Caps_Bot_Yello, LM_Bumpers_l62_B_Caps_Top_Amber, LM_Bumpers_l62_B_Caps_Top_Red, LM_Bumpers_l62_B_Caps_Top_Yello, LM_Bumpers_l62_BR1, LM_Bumpers_l62_BR2, LM_Bumpers_l62_BR3, LM_Bumpers_l62_BS1, LM_Bumpers_l62_BS2, LM_Bumpers_l62_BS3, LM_Bumpers_l62_BumperPegs, LM_Bumpers_l62_MysticSeerToy, LM_Bumpers_l62_Outlanes, LM_Bumpers_l62_Over3, LM_Bumpers_l62_PF, _
  LM_Bumpers_l62_PF_Upper, LM_Bumpers_l62_Parts, LM_Bumpers_l62_RocketToy, LM_Bumpers_l62_SideMod, LM_Bumpers_l62_SlotMachineToy, LM_Bumpers_l62_sw48, LM_Bumpers_l62_sw48m, LM_Bumpers_l62_sw77, LM_Bumpers_l62_sw77m, LM_Bumpers_l63_B_Caps_Bot_Amber, LM_Bumpers_l63_B_Caps_Bot_Red, LM_Bumpers_l63_B_Caps_Bot_Yello, LM_Bumpers_l63_B_Caps_Top_Amber, LM_Bumpers_l63_B_Caps_Top_Red, LM_Bumpers_l63_B_Caps_Top_Yello, LM_Bumpers_l63_BR1, LM_Bumpers_l63_BR2, LM_Bumpers_l63_BR3, LM_Bumpers_l63_BS1, LM_Bumpers_l63_BS2, LM_Bumpers_l63_BS3, LM_Bumpers_l63_BumperPegs, LM_Bumpers_l63_Camera, LM_Bumpers_l63_Clock_Color, LM_Bumpers_l63_Clock_White, LM_Bumpers_l63_FlipperL1, LM_Bumpers_l63_FlipperSpL1, LM_Bumpers_l63_Over3, LM_Bumpers_l63_PF, LM_Bumpers_l63_PF_Upper, LM_Bumpers_l63_Parts, LM_Bumpers_l63_Rails, LM_Bumpers_l63_Sign_Spiral, LM_Bumpers_l63_SlotMachineToy, LM_Bumpers_l63_TownSquarePost, LM_Bumpers_l63_sw77, LM_Bumpers_l63_sw77m, LM_Bumpers_l63_sw78, LM_Bumpers_l63_sw78m, LM_Flashers_f17_B_Caps_Bot_Ambe, _
  LM_Flashers_f17_B_Caps_Bot_Red, LM_Flashers_f17_B_Caps_Bot_Yell, LM_Flashers_f17_B_Caps_Top_Ambe, LM_Flashers_f17_B_Caps_Top_Red, LM_Flashers_f17_B_Caps_Top_Yell, LM_Flashers_f17_BR1, LM_Flashers_f17_BR2, LM_Flashers_f17_BR3, LM_Flashers_f17_BS1, LM_Flashers_f17_BS2, LM_Flashers_f17_BS3, LM_Flashers_f17_BumperPegs, LM_Flashers_f17_Camera, LM_Flashers_f17_Clock_Color, LM_Flashers_f17_Clock_White, LM_Flashers_f17_ClockToy, LM_Flashers_f17_FlipperLU, LM_Flashers_f17_FlipperR1, LM_Flashers_f17_FlipperR1U, LM_Flashers_f17_FlipperRU, LM_Flashers_f17_FlipperSpL, LM_Flashers_f17_FlipperSpLU, LM_Flashers_f17_MysticSeerToy, LM_Flashers_f17_Outlanes, LM_Flashers_f17_Over1, LM_Flashers_f17_Over3, LM_Flashers_f17_PF, LM_Flashers_f17_PF_Upper, LM_Flashers_f17_Parts, LM_Flashers_f17_Rails, LM_Flashers_f17_RocketToy, LM_Flashers_f17_SideMod, LM_Flashers_f17_SlotMachineToy, LM_Flashers_f17_TownSquarePost, LM_Flashers_f17_sw36, LM_Flashers_f17_sw37, LM_Flashers_f17_sw38, LM_Flashers_f17_sw47, LM_Flashers_f17_sw48, _
  LM_Flashers_f17_sw48m, LM_Flashers_f17_sw77, LM_Flashers_f17_sw77m, LM_Flashers_f17_sw78, LM_Flashers_f17_sw78m, LM_Flashers_f18_B_Caps_Bot_Ambe, LM_Flashers_f18_B_Caps_Bot_Yell, LM_Flashers_f18_B_Caps_Top_Red, LM_Flashers_f18_B_Caps_Top_Yell, LM_Flashers_f18_Clock_Color, LM_Flashers_f18_Clock_White, LM_Flashers_f18_ClockToy, LM_Flashers_f18_DiverterP, LM_Flashers_f18_FlipperL, LM_Flashers_f18_FlipperL1, LM_Flashers_f18_FlipperLU, LM_Flashers_f18_FlipperR1, LM_Flashers_f18_FlipperR1U, LM_Flashers_f18_FlipperRU, LM_Flashers_f18_FlipperSpL, LM_Flashers_f18_FlipperSpL1, LM_Flashers_f18_FlipperSpLU, LM_Flashers_f18_FlipperSpR1, LM_Flashers_f18_FlipperSpR1U, LM_Flashers_f18_FlipperSpRU, LM_Flashers_f18_GMKnob, LM_Flashers_f18_Gate1, LM_Flashers_f18_Gumballs, LM_Flashers_f18_InvaderToy, LM_Flashers_f18_LSling1, LM_Flashers_f18_LSling2, LM_Flashers_f18_Over1, LM_Flashers_f18_Over2, LM_Flashers_f18_Over3, LM_Flashers_f18_PF, LM_Flashers_f18_PF_Upper, LM_Flashers_f18_Parts, LM_Flashers_f18_Piano, _
  LM_Flashers_f18_Pyramid, LM_Flashers_f18_Rails, LM_Flashers_f18_Robot, LM_Flashers_f18_RocketToy, LM_Flashers_f18_SideMod, LM_Flashers_f18_Sign_Spiral, LM_Flashers_f18_SlotMachineToy, LM_Flashers_f18_TVtoy, LM_Flashers_f18_TownSquarePost, LM_Flashers_f18_sw47, LM_Flashers_f18_sw63, LM_Flashers_f18_sw64, LM_Flashers_f18_sw65am, LM_Flashers_f18_sw65m, LM_Flashers_f18_sw67m, LM_Flashers_f18_sw68, LM_Flashers_f19_BR2, LM_Flashers_f19_BR3, LM_Flashers_f19_BS3, LM_Flashers_f19_Camera, LM_Flashers_f19_Clock_Color, LM_Flashers_f19_Clock_White, LM_Flashers_f19_ClockLarge, LM_Flashers_f19_ClockToy, LM_Flashers_f19_FlipperL1, LM_Flashers_f19_FlipperR1, LM_Flashers_f19_FlipperR1U, LM_Flashers_f19_FlipperSpL1, LM_Flashers_f19_FlipperSpR1, LM_Flashers_f19_FlipperSpR1U, LM_Flashers_f19_GMKnob, LM_Flashers_f19_Gumballs, LM_Flashers_f19_Over1, LM_Flashers_f19_Over2, LM_Flashers_f19_Over3, LM_Flashers_f19_PF, LM_Flashers_f19_PF_Upper, LM_Flashers_f19_Parts, LM_Flashers_f19_Piano, LM_Flashers_f19_Pyramid, LM_Flashers_f19_RDiv, _
  LM_Flashers_f19_Rails, LM_Flashers_f19_Robot, LM_Flashers_f19_RocketToy, LM_Flashers_f19_SLING2, LM_Flashers_f19_SideMod, LM_Flashers_f19_Sign_Spiral, LM_Flashers_f19_SlotMachineToy, LM_Flashers_f19_SpiralToy, LM_Flashers_f19_sw38, LM_Flashers_f19_sw47, LM_Flashers_f19_sw56, LM_Flashers_f19_sw64, LM_Flashers_f19_sw64m, LM_Flashers_f19_sw65, LM_Flashers_f19_sw65a, LM_Flashers_f19_sw65am, LM_Flashers_f19_sw65m, LM_Flashers_f19_sw66, LM_Flashers_f19_sw66m, LM_Flashers_f19_sw67, LM_Flashers_f19_sw67m, LM_Flashers_f20_B_Caps_Bot_Ambe, LM_Flashers_f20_BS2, LM_Flashers_f20_Camera, LM_Flashers_f20_Clock_Color, LM_Flashers_f20_Clock_White, LM_Flashers_f20_ClockToy, LM_Flashers_f20_DiverterP1, LM_Flashers_f20_FlipperL, LM_Flashers_f20_FlipperL1, LM_Flashers_f20_FlipperLU, LM_Flashers_f20_FlipperR, LM_Flashers_f20_FlipperR1, LM_Flashers_f20_FlipperR1U, LM_Flashers_f20_FlipperRU, LM_Flashers_f20_FlipperSpL1, LM_Flashers_f20_FlipperSpR1U, LM_Flashers_f20_FlipperSpRU, LM_Flashers_f20_GMKnob, LM_Flashers_f20_Gumballs, _
  LM_Flashers_f20_InvaderToy, LM_Flashers_f20_Over1, LM_Flashers_f20_Over2, LM_Flashers_f20_Over3, LM_Flashers_f20_Over4, LM_Flashers_f20_Over5, LM_Flashers_f20_PF, LM_Flashers_f20_PF_Upper, LM_Flashers_f20_Parts, LM_Flashers_f20_Piano, LM_Flashers_f20_Pyramid, LM_Flashers_f20_RDiv, LM_Flashers_f20_RSling1, LM_Flashers_f20_Rails, LM_Flashers_f20_Rails_DT, LM_Flashers_f20_Robot, LM_Flashers_f20_SideMod, LM_Flashers_f20_Sign_Spiral, LM_Flashers_f20_SlotMachineToy, LM_Flashers_f20_SpiralToy, LM_Flashers_f20_TVtoy, LM_Flashers_f20_sw53p, LM_Flashers_f20_sw56, LM_Flashers_f20_sw65am, LM_Flashers_f20_sw78, LM_Flashers_f28_B_Caps_Bot_Ambe, LM_Flashers_f28_B_Caps_Top_Ambe, LM_Flashers_f28_B_Caps_Top_Yell, LM_Flashers_f28_BR2, LM_Flashers_f28_Camera, LM_Flashers_f28_Clock_Color, LM_Flashers_f28_Clock_White, LM_Flashers_f28_ClockLarge, LM_Flashers_f28_ClockToy, LM_Flashers_f28_DiverterP, LM_Flashers_f28_DiverterP1, LM_Flashers_f28_FlipperL, LM_Flashers_f28_FlipperL1, LM_Flashers_f28_FlipperLU, LM_Flashers_f28_FlipperR, _
  LM_Flashers_f28_FlipperR1, LM_Flashers_f28_FlipperR1U, LM_Flashers_f28_FlipperRU, LM_Flashers_f28_FlipperSpL, LM_Flashers_f28_FlipperSpL1, LM_Flashers_f28_FlipperSpLU, LM_Flashers_f28_FlipperSpR, LM_Flashers_f28_FlipperSpR1, LM_Flashers_f28_FlipperSpR1U, LM_Flashers_f28_FlipperSpRU, LM_Flashers_f28_GMKnob, LM_Flashers_f28_Gumballs, LM_Flashers_f28_LSling1, LM_Flashers_f28_Over1, LM_Flashers_f28_Over2, LM_Flashers_f28_Over3, LM_Flashers_f28_Over4, LM_Flashers_f28_Over5, LM_Flashers_f28_PF, LM_Flashers_f28_PF_Upper, LM_Flashers_f28_Parts, LM_Flashers_f28_Piano, LM_Flashers_f28_Pyramid, LM_Flashers_f28_RDiv, LM_Flashers_f28_RSling1, LM_Flashers_f28_RSling2, LM_Flashers_f28_Rails, LM_Flashers_f28_Robot, LM_Flashers_f28_RocketToy, LM_Flashers_f28_SideMod, LM_Flashers_f28_Sign_Spiral, LM_Flashers_f28_SlotMachineToy, LM_Flashers_f28_TVtoy, LM_Flashers_f28_sw11, LM_Flashers_f28_sw38, LM_Flashers_f28_sw47m, LM_Flashers_f28_sw53p, LM_Flashers_f28_sw63, LM_Flashers_f28_sw64, LM_Flashers_f28_sw64m, LM_Flashers_f28_sw65, _
  LM_Flashers_f28_sw65a, LM_Flashers_f28_sw65am, LM_Flashers_f28_sw65m, LM_Flashers_f28_sw66, LM_Flashers_f28_sw67, LM_Flashers_f28_sw68, LM_Flashers_f37_B_Caps_Bot_Ambe, LM_Flashers_f37_B_Caps_Top_Ambe, LM_Flashers_f37_B_Caps_Top_Yell, LM_Flashers_f37_Clock_Color, LM_Flashers_f37_Clock_White, LM_Flashers_f37_ClockToy, LM_Flashers_f37_FlipperR1, LM_Flashers_f37_FlipperR1U, LM_Flashers_f37_FlipperSpR1, LM_Flashers_f37_FlipperSpR1U, LM_Flashers_f37_InvaderToy, LM_Flashers_f37_Over1, LM_Flashers_f37_PF, LM_Flashers_f37_Parts, LM_Flashers_f37_Rails, LM_Flashers_f37_RocketToy, LM_Flashers_f37_Sign_Spiral, LM_Flashers_f37_SlotMachineToy, LM_Flashers_f37_sw63, LM_Flashers_f37_sw78, LM_Flashers_f38_BS1, LM_Flashers_f38_Clock_Color, LM_Flashers_f38_Clock_White, LM_Flashers_f38_ClockToy, LM_Flashers_f38_FlipperL1, LM_Flashers_f38_FlipperR1, LM_Flashers_f38_FlipperR1U, LM_Flashers_f38_FlipperSpL1, LM_Flashers_f38_FlipperSpR1, LM_Flashers_f38_FlipperSpR1U, LM_Flashers_f38_GMKnob, LM_Flashers_f38_Gumballs, _
  LM_Flashers_f38_Over1, LM_Flashers_f38_Over2, LM_Flashers_f38_Over3, LM_Flashers_f38_PF, LM_Flashers_f38_PF_Upper, LM_Flashers_f38_Parts, LM_Flashers_f38_Piano, LM_Flashers_f38_Pyramid, LM_Flashers_f38_RDiv, LM_Flashers_f38_RSling1, LM_Flashers_f38_RSling2, LM_Flashers_f38_Rails, LM_Flashers_f38_Rails_DT, LM_Flashers_f38_Robot, LM_Flashers_f38_SideMod, LM_Flashers_f38_Sign_Spiral, LM_Flashers_f38_SlotMachineToy, LM_Flashers_f38_sw56, LM_Flashers_f39_Clock_White, LM_Flashers_f39_ClockToy, LM_Flashers_f39_FlipperL1, LM_Flashers_f39_FlipperR, LM_Flashers_f39_FlipperR1, LM_Flashers_f39_FlipperR1U, LM_Flashers_f39_FlipperRU, LM_Flashers_f39_FlipperSpL1, LM_Flashers_f39_FlipperSpR, LM_Flashers_f39_FlipperSpR1, LM_Flashers_f39_FlipperSpR1U, LM_Flashers_f39_FlipperSpRU, LM_Flashers_f39_GMKnob, LM_Flashers_f39_Gumballs, LM_Flashers_f39_Over1, LM_Flashers_f39_Over2, LM_Flashers_f39_Over3, LM_Flashers_f39_Over4, LM_Flashers_f39_PF, LM_Flashers_f39_PF_Upper, LM_Flashers_f39_Parts, LM_Flashers_f39_Piano, _
  LM_Flashers_f39_Pyramid, LM_Flashers_f39_RDiv, LM_Flashers_f39_RSling1, LM_Flashers_f39_RSling2, LM_Flashers_f39_Rails, LM_Flashers_f39_Rails_DT, LM_Flashers_f39_Robot, LM_Flashers_f39_SLING1, LM_Flashers_f39_SideMod, LM_Flashers_f39_Sign_Spiral, LM_Flashers_f39_SlotMachineToy, LM_Flashers_f39_sw12, LM_Flashers_f39_sw56, LM_Flashers_f40_ClockToy, LM_Flashers_f40_FlipperL1, LM_Flashers_f40_GMKnob, LM_Flashers_f40_Gumballs, LM_Flashers_f40_Over2, LM_Flashers_f40_Over3, LM_Flashers_f40_PF, LM_Flashers_f40_PF_Upper, LM_Flashers_f40_Parts, LM_Flashers_f40_RDiv, LM_Flashers_f40_Rails, LM_Flashers_f40_Robot, LM_Flashers_f40_SideMod, LM_Flashers_f40_Sign_Spiral, LM_Flashers_f41_Clock_Color, LM_Flashers_f41_Clock_White, LM_Flashers_f41_ClockToy, LM_Flashers_f41_DiverterP, LM_Flashers_f41_DiverterP1, LM_Flashers_f41_FlipperL1, LM_Flashers_f41_FlipperR1, LM_Flashers_f41_FlipperR1U, LM_Flashers_f41_FlipperRU, LM_Flashers_f41_FlipperSpL1, LM_Flashers_f41_FlipperSpR1, LM_Flashers_f41_FlipperSpR1U, LM_Flashers_f41_GMKnob, _
  LM_Flashers_f41_Over1, LM_Flashers_f41_Over2, LM_Flashers_f41_Over3, LM_Flashers_f41_Over4, LM_Flashers_f41_PF, LM_Flashers_f41_PF_Upper, LM_Flashers_f41_Parts, LM_Flashers_f41_Piano, LM_Flashers_f41_Pyramid, LM_Flashers_f41_RDiv, LM_Flashers_f41_Rails, LM_Flashers_f41_Rails_DT, LM_Flashers_f41_Robot, LM_Flashers_f41_RocketToy, LM_Flashers_f41_SideMod, LM_Flashers_f41_SlotMachineToy, LM_Flashers_f41_TVtoy, LM_Flashers_f41_URMagnet, LM_Flashers_f41_sw54p, LM_Flashers_l74_Clock_Color, LM_Flashers_l74_Clock_White, LM_Flashers_l74_ClockShort, LM_Flashers_l74_ClockToy, LM_Flashers_l74_Over1, LM_Flashers_l74_Over2, LM_Flashers_l74_Over3, LM_Flashers_l74_PF, LM_Flashers_l74_Parts, LM_Flashers_l74_Piano, LM_Flashers_l74_Robot, LM_Flashers_l74_sw64, LM_Flashers_l74_sw66, LM_GI_Clock_l102_B_Caps_Bot_Red, LM_GI_Clock_l102_B_Caps_Top_Amb, LM_GI_Clock_l102_B_Caps_Top_Yel, LM_GI_Clock_l102_Clock_Color, LM_GI_Clock_l102_Clock_White, LM_GI_Clock_l102_ClockLarge, LM_GI_Clock_l102_ClockShort, LM_GI_Clock_l102_ClockToy, _
  LM_GI_Clock_l102_DiverterP, LM_GI_Clock_l102_FlipperL1, LM_GI_Clock_l102_FlipperR1U, LM_GI_Clock_l102_Over1, LM_GI_Clock_l102_Over2, LM_GI_Clock_l102_Over3, LM_GI_Clock_l102_PF, LM_GI_Clock_l102_Parts, LM_GI_Clock_l102_Piano, LM_GI_Clock_l102_Rails, LM_GI_Clock_l102_SideMod, LM_GI_Clock_l102_Sign_Spiral, LM_GI_Clock_l102_SlotMachineToy, LM_GI_Clock_l102_sw63, LM_GI_Clock_l102_sw64, LM_GI_Clock_l102_sw64m, LM_GI_Clock_l102_sw65, LM_GI_Clock_l102_sw65a, LM_GI_Clock_l102_sw65am, LM_GI_Clock_l102_sw65m, LM_GI_Left_l100_B_Caps_Bot_Red, LM_GI_Left_l100_B_Caps_Top_Ambe, LM_GI_Left_l100_B_Caps_Top_Red, LM_GI_Left_l100_BR1, LM_GI_Left_l100_BR2, LM_GI_Left_l100_BS1, LM_GI_Left_l100_BS2, LM_GI_Left_l100_BS3, LM_GI_Left_l100_BumperPegs, LM_GI_Left_l100_ClockToy, LM_GI_Left_l100_FlipperR1U, LM_GI_Left_l100_GMKnob, LM_GI_Left_l100_Gumballs, LM_GI_Left_l100_Over1, LM_GI_Left_l100_Over2, LM_GI_Left_l100_Over3, LM_GI_Left_l100_Over4, LM_GI_Left_l100_Over5, LM_GI_Left_l100_PF, LM_GI_Left_l100_PF_Upper, LM_GI_Left_l100_Parts, _
  LM_GI_Left_l100_RDiv, LM_GI_Left_l100_Rails, LM_GI_Left_l100_Robot, LM_GI_Left_l100_SideMod, LM_GI_Left_l100_Sign_Spiral, LM_GI_Left_l100_SlotMachineToy, LM_GI_Left_l100_TownSquarePost, LM_GI_Left_l100_sw36, LM_GI_Left_l100_sw37, LM_GI_Left_l100_sw38, LM_GI_Left_l100_sw53p, LM_GI_Left_l100_sw56, LM_GI_Left_l100_sw68, LM_GI_Left_l100_sw77, LM_GI_Left_l100a_B_Caps_Top_Yel, LM_GI_Left_l100a_ClockToy, LM_GI_Left_l100a_FlipperL, LM_GI_Left_l100a_FlipperLU, LM_GI_Left_l100a_FlipperR, LM_GI_Left_l100a_FlipperR1, LM_GI_Left_l100a_FlipperRU, LM_GI_Left_l100a_FlipperSpL, LM_GI_Left_l100a_FlipperSpLU, LM_GI_Left_l100a_FlipperSpR, LM_GI_Left_l100a_FlipperSpRU, LM_GI_Left_l100a_LSling1, LM_GI_Left_l100a_LSling2, LM_GI_Left_l100a_MysticSeerToy, LM_GI_Left_l100a_Outlanes, LM_GI_Left_l100a_PF, LM_GI_Left_l100a_Parts, LM_GI_Left_l100a_RSling1, LM_GI_Left_l100a_RSling2, LM_GI_Left_l100a_Rails, LM_GI_Left_l100a_RocketToy, LM_GI_Left_l100a_SLING2, LM_GI_Left_l100a_sw36, LM_GI_Left_l100a_sw38, LM_GI_Left_l100a_sw48, _
  LM_GI_Left_l100a_sw48m, LM_GI_Left_l100a_sw77, LM_GI_Left_l100b_B_Caps_Top_Amb, LM_GI_Left_l100b_BR2, LM_GI_Left_l100b_BS3, LM_GI_Left_l100b_ClockToy, LM_GI_Left_l100b_FlipperL, LM_GI_Left_l100b_FlipperLU, LM_GI_Left_l100b_FlipperR, LM_GI_Left_l100b_FlipperR1U, LM_GI_Left_l100b_FlipperSpL, LM_GI_Left_l100b_FlipperSpLU, LM_GI_Left_l100b_LSling1, LM_GI_Left_l100b_LSling2, LM_GI_Left_l100b_MysticSeerToy, LM_GI_Left_l100b_Outlanes, LM_GI_Left_l100b_PF, LM_GI_Left_l100b_PF_Upper, LM_GI_Left_l100b_Parts, LM_GI_Left_l100b_SLING2, LM_GI_Left_l100b_SideMod, LM_GI_Left_l100b_SlotMachineToy, LM_GI_Left_l100b_TownSquarePost, LM_GI_Left_l100b_sw36, LM_GI_Left_l100b_sw37, LM_GI_Left_l100b_sw38, LM_GI_Left_l100b_sw48, LM_GI_Left_l100b_sw48m, LM_GI_Left_l100b_sw77, LM_GI_Left_l100c_B_Caps_Top_Amb, LM_GI_Left_l100c_B_Caps_Top_Red, LM_GI_Left_l100c_B_Caps_Top_Yel, LM_GI_Left_l100c_BR3, LM_GI_Left_l100c_FlipperLU, LM_GI_Left_l100c_FlipperR, LM_GI_Left_l100c_FlipperSpLU, LM_GI_Left_l100c_PF, LM_GI_Left_l100c_Parts, _
  LM_GI_Left_l100c_SLING2, LM_GI_Left_l100c_SideMod, LM_GI_Left_l100c_sw36, LM_GI_Left_l100c_sw37, LM_GI_Left_l100c_sw38, LM_GI_Left_l100c_sw48, LM_GI_Left_l100d_B_Caps_Top_Amb, LM_GI_Left_l100d_B_Caps_Top_Yel, LM_GI_Left_l100d_FlipperLU, LM_GI_Left_l100d_FlipperR, LM_GI_Left_l100d_FlipperRU, LM_GI_Left_l100d_FlipperSpLU, LM_GI_Left_l100d_PF, LM_GI_Left_l100d_Parts, LM_GI_Left_l100d_SLING2, LM_GI_Left_l100d_sw36, LM_GI_Left_l100d_sw37, LM_GI_Left_l100d_sw38, LM_GI_Left_l100d_sw48, LM_GI_Left_l100d_sw48m, LM_GI_Left_l100e_B_Caps_Top_Yel, LM_GI_Left_l100e_BR3, LM_GI_Left_l100e_FlipperL, LM_GI_Left_l100e_FlipperLU, LM_GI_Left_l100e_FlipperR, LM_GI_Left_l100e_FlipperRU, LM_GI_Left_l100e_FlipperSpL, LM_GI_Left_l100e_FlipperSpLU, LM_GI_Left_l100e_FlipperSpR, LM_GI_Left_l100e_FlipperSpRU, LM_GI_Left_l100e_Outlanes, LM_GI_Left_l100e_PF, LM_GI_Left_l100e_Parts, LM_GI_Left_l100e_RocketToy, LM_GI_Left_l100e_SLING2, LM_GI_Left_l100e_sw38, LM_GI_Left_l100e_sw48, LM_GI_MiniPF_l101_GMKnob, LM_GI_MiniPF_l101_Gumballs, _
  LM_GI_MiniPF_l101_PF, LM_GI_MiniPF_l101_PF_Upper, LM_GI_MiniPF_l101_Parts, LM_GI_MiniPF_l101_Pyramid, LM_GI_MiniPF_l101_SideMod, LM_GI_MiniPF_l101_Sign_Spiral, LM_GI_MiniPF_l101_SpiralToy, LM_GI_Right_l104_B_Caps_Top_Amb, LM_GI_Right_l104_BR2, LM_GI_Right_l104_BS1, LM_GI_Right_l104_Clock_Color, LM_GI_Right_l104_Clock_White, LM_GI_Right_l104_ClockToy, LM_GI_Right_l104_DiverterP, LM_GI_Right_l104_FlipperL1, LM_GI_Right_l104_FlipperR1, LM_GI_Right_l104_FlipperR1U, LM_GI_Right_l104_FlipperSpL1, LM_GI_Right_l104_FlipperSpR, LM_GI_Right_l104_FlipperSpR1, LM_GI_Right_l104_FlipperSpR1U, LM_GI_Right_l104_FlipperSpRU, LM_GI_Right_l104_Gate1, LM_GI_Right_l104_InvaderToy, LM_GI_Right_l104_Over1, LM_GI_Right_l104_Over2, LM_GI_Right_l104_Over3, LM_GI_Right_l104_PF, LM_GI_Right_l104_PF_Upper, LM_GI_Right_l104_Parts, LM_GI_Right_l104_Piano, LM_GI_Right_l104_Rails, LM_GI_Right_l104_Robot, LM_GI_Right_l104_RocketToy, LM_GI_Right_l104_SideMod, LM_GI_Right_l104_Sign_Spiral, LM_GI_Right_l104_SlotMachineToy, _
  LM_GI_Right_l104_TVtoy, LM_GI_Right_l104_URMagnet, LM_GI_Right_l104_sw47, LM_GI_Right_l104_sw61, LM_GI_Right_l104_sw62, LM_GI_Right_l104_sw63, LM_GI_Right_l104_sw64m, LM_GI_Right_l104_sw65, LM_GI_Right_l104_sw65a, LM_GI_Right_l104_sw65am, LM_GI_Right_l104_sw65m, LM_GI_Right_l104_sw66, LM_GI_Right_l104_sw66m, LM_GI_Right_l104_sw67, LM_GI_Right_l104_sw67m, LM_GI_Right_l104_sw68m, LM_GI_Right_l104_sw78, LM_GI_Right_l104a_B_Caps_Top_Re, LM_GI_Right_l104a_B_Caps_Top_Ye, LM_GI_Right_l104a_BR3, LM_GI_Right_l104a_ClockToy, LM_GI_Right_l104a_FlipperL, LM_GI_Right_l104a_FlipperLU, LM_GI_Right_l104a_FlipperR, LM_GI_Right_l104a_FlipperR1U, LM_GI_Right_l104a_FlipperRU, LM_GI_Right_l104a_FlipperSpL, LM_GI_Right_l104a_FlipperSpLU, LM_GI_Right_l104a_FlipperSpR, LM_GI_Right_l104a_FlipperSpR1U, LM_GI_Right_l104a_FlipperSpRU, LM_GI_Right_l104a_InvaderToy, LM_GI_Right_l104a_LSling1, LM_GI_Right_l104a_LSling2, LM_GI_Right_l104a_Outlanes, LM_GI_Right_l104a_Over3, LM_GI_Right_l104a_PF, LM_GI_Right_l104a_Parts, _
  LM_GI_Right_l104a_RSling1, LM_GI_Right_l104a_RSling2, LM_GI_Right_l104a_Rails, LM_GI_Right_l104a_RocketToy, LM_GI_Right_l104a_SLING1, LM_GI_Right_l104a_sw11, LM_GI_Right_l104a_sw12, LM_GI_Right_l104a_sw48, LM_GI_Right_l104b_ClockToy, LM_GI_Right_l104b_PF, LM_GI_Right_l104b_Parts, LM_GI_Right_l104b_Robot, LM_GI_Right_l104b_Sign_Spiral, LM_GI_Right_l104b_SlotMachineTo, LM_GI_Right_l104b_sw65m, LM_GI_Right_l104b_sw66, LM_GI_Right_l104b_sw66m, LM_GI_Right_l104b_sw67, LM_GI_Right_l104b_sw67m, LM_GI_Right_l104c_B_Caps_Top_Am, LM_GI_Right_l104c_B_Caps_Top_Ye, LM_GI_Right_l104c_BR2, LM_GI_Right_l104c_BR3, LM_GI_Right_l104c_FlipperL, LM_GI_Right_l104c_FlipperLU, LM_GI_Right_l104c_FlipperR, LM_GI_Right_l104c_FlipperRU, LM_GI_Right_l104c_FlipperSpL, LM_GI_Right_l104c_FlipperSpLU, LM_GI_Right_l104c_FlipperSpR, LM_GI_Right_l104c_FlipperSpRU, LM_GI_Right_l104c_Outlanes, LM_GI_Right_l104c_PF, LM_GI_Right_l104c_Parts, LM_GI_Right_l104c_Rails, LM_GI_Right_l104c_SLING1, LM_GI_Right_l104c_sw11, LM_GI_Right_l104d_FlipperL, _
  LM_GI_Right_l104d_FlipperLU, LM_GI_Right_l104d_FlipperR, LM_GI_Right_l104d_FlipperR1U, LM_GI_Right_l104d_FlipperRU, LM_GI_Right_l104d_FlipperSpL, LM_GI_Right_l104d_FlipperSpLU, LM_GI_Right_l104d_FlipperSpR, LM_GI_Right_l104d_FlipperSpR1U, LM_GI_Right_l104d_FlipperSpRU, LM_GI_Right_l104d_InvaderToy, LM_GI_Right_l104d_Outlanes, LM_GI_Right_l104d_PF, LM_GI_Right_l104d_Parts, LM_GI_Right_l104d_Rails, LM_GI_Right_l104d_RocketToy, LM_GI_Right_l104d_SLING1, LM_GI_Right_l104d_sw11, LM_GI_Right_l104d_sw12, LM_Inserts_l11_B_Caps_Top_Yello, LM_Inserts_l11_FlipperLU, LM_Inserts_l11_FlipperRU, LM_Inserts_l11_LSling1, LM_Inserts_l11_LSling2, LM_Inserts_l11_Parts, LM_Inserts_l11_RSling1, LM_Inserts_l12_B_Caps_Top_Yello, LM_Inserts_l12_FlipperLU, LM_Inserts_l12_LSling1, LM_Inserts_l12_LSling2, LM_Inserts_l12_Parts, LM_Inserts_l12_SLING2, LM_Inserts_l13_B_Caps_Top_Amber, LM_Inserts_l13_B_Caps_Top_Yello, LM_Inserts_l13_LSling1, LM_Inserts_l13_LSling2, LM_Inserts_l13_Parts, LM_Inserts_l13_sw48, LM_Inserts_l14_B_Caps_Top_Red, _
  LM_Inserts_l14_B_Caps_Top_Yello, LM_Inserts_l14_LSling2, LM_Inserts_l14_Parts, LM_Inserts_l14_sw48, LM_Inserts_l14_sw48m, LM_Inserts_l14_sw77, LM_Inserts_l15_B_Caps_Top_Yello, LM_Inserts_l15_PF, LM_Inserts_l15_PF_Upper, LM_Inserts_l15_Parts, LM_Inserts_l15_sw48, LM_Inserts_l15_sw48m, LM_Inserts_l15_sw77, LM_Inserts_l16_B_Caps_Top_Amber, LM_Inserts_l16_B_Caps_Top_Red, LM_Inserts_l16_B_Caps_Top_Yello, LM_Inserts_l16_PF, LM_Inserts_l16_PF_Upper, LM_Inserts_l16_Parts, LM_Inserts_l16_sw77, LM_Inserts_l16_sw77m, LM_Inserts_l16_sw78, LM_Inserts_l17_B_Caps_Bot_Amber, LM_Inserts_l17_B_Caps_Top_Amber, LM_Inserts_l17_B_Caps_Top_Red, LM_Inserts_l17_B_Caps_Top_Yello, LM_Inserts_l17_BR2, LM_Inserts_l17_BS2, LM_Inserts_l17_PF_Upper, LM_Inserts_l17_Parts, LM_Inserts_l17_sw77, LM_Inserts_l17_sw77m, LM_Inserts_l17_sw78, LM_Inserts_l17_sw78m, LM_Inserts_l18_B_Caps_Top_Amber, LM_Inserts_l18_PF_Upper, LM_Inserts_l18_Parts, LM_Inserts_l18_sw77, LM_Inserts_l18_sw78, LM_Inserts_l21_B_Caps_Top_Yello, LM_Inserts_l21_Parts, _
  LM_Inserts_l22_B_Caps_Top_Yello, LM_Inserts_l22_FlipperRU, LM_Inserts_l22_Parts, LM_Inserts_l22_RSling1, LM_Inserts_l22_RSling2, LM_Inserts_l23_B_Caps_Top_Yello, LM_Inserts_l23_Parts, LM_Inserts_l23_RSling1, LM_Inserts_l23_RSling2, LM_Inserts_l23_SLING1, LM_Inserts_l24_B_Caps_Top_Amber, LM_Inserts_l24_B_Caps_Top_Yello, LM_Inserts_l24_Parts, LM_Inserts_l24_RSling2, LM_Inserts_l25_B_Caps_Top_Amber, LM_Inserts_l25_B_Caps_Top_Yello, LM_Inserts_l25_FlipperR1, LM_Inserts_l25_FlipperR1U, LM_Inserts_l25_PF, LM_Inserts_l25_Parts, LM_Inserts_l26_B_Caps_Top_Amber, LM_Inserts_l26_B_Caps_Top_Red, LM_Inserts_l26_FlipperR1U, LM_Inserts_l26_FlipperSpR1U, LM_Inserts_l26_Parts, LM_Inserts_l27_B_Caps_Top_Amber, LM_Inserts_l27_FlipperR1U, LM_Inserts_l27_FlipperSpR1U, LM_Inserts_l27_PF, LM_Inserts_l27_Parts, LM_Inserts_l28_B_Caps_Top_Amber, LM_Inserts_l28_B_Caps_Top_Yello, LM_Inserts_l28_Parts, LM_Inserts_l28_SlotMachineToy, LM_Inserts_l31_B_Caps_Top_Red, LM_Inserts_l31_B_Caps_Top_Yello, LM_Inserts_l31_Outlanes, _
  LM_Inserts_l31_Parts, LM_Inserts_l31_SideMod, LM_Inserts_l31_sw36, LM_Inserts_l32_B_Caps_Top_Amber, LM_Inserts_l32_B_Caps_Top_Yello, LM_Inserts_l32_Parts, LM_Inserts_l33_B_Caps_Top_Yello, LM_Inserts_l33_MysticSeerToy, LM_Inserts_l33_Outlanes, LM_Inserts_l33_PF, LM_Inserts_l33_Parts, LM_Inserts_l33_sw48, LM_Inserts_l34_B_Caps_Top_Red, LM_Inserts_l34_B_Caps_Top_Yello, LM_Inserts_l34_PF, LM_Inserts_l34_Parts, LM_Inserts_l34_sw48, LM_Inserts_l34_sw77, LM_Inserts_l35_B_Caps_Top_Yello, LM_Inserts_l35_Outlanes, LM_Inserts_l35_PF, LM_Inserts_l35_Parts, LM_Inserts_l35_sw48, LM_Inserts_l35_sw48m, LM_Inserts_l36_B_Caps_Top_Red, LM_Inserts_l36_B_Caps_Top_Yello, LM_Inserts_l36_Parts, LM_Inserts_l36_sw77, LM_Inserts_l36_sw78, LM_Inserts_l37_B_Caps_Top_Yello, LM_Inserts_l37_BR3, LM_Inserts_l37_BS3, LM_Inserts_l37_Outlanes, LM_Inserts_l37_PF, LM_Inserts_l37_Parts, LM_Inserts_l37_sw48, LM_Inserts_l37_sw48m, LM_Inserts_l37_sw77, LM_Inserts_l38_B_Caps_Top_Amber, LM_Inserts_l38_BR3, LM_Inserts_l38_BS3, LM_Inserts_l38_BumperPegs, _
  LM_Inserts_l38_PF_Upper, LM_Inserts_l38_Parts, LM_Inserts_l38_TownSquarePost, LM_Inserts_l38_sw48m, LM_Inserts_l38_sw77, LM_Inserts_l38_sw77m, LM_Inserts_l41_FlipperL, LM_Inserts_l41_FlipperLU, LM_Inserts_l41_FlipperSpL, LM_Inserts_l41_FlipperSpLU, LM_Inserts_l41_LSling1, LM_Inserts_l41_LSling2, LM_Inserts_l41_PF, LM_Inserts_l41_Parts, LM_Inserts_l41_SLING2, LM_Inserts_l42_B_Caps_Top_Yello, LM_Inserts_l42_FlipperL, LM_Inserts_l42_FlipperLU, LM_Inserts_l42_FlipperR, LM_Inserts_l42_FlipperRU, LM_Inserts_l42_FlipperSpL, LM_Inserts_l42_FlipperSpLU, LM_Inserts_l42_FlipperSpRU, LM_Inserts_l42_LSling1, LM_Inserts_l42_LSling2, LM_Inserts_l42_PF, LM_Inserts_l42_Parts, LM_Inserts_l43_B_Caps_Top_Yello, LM_Inserts_l43_FlipperL, LM_Inserts_l43_FlipperLU, LM_Inserts_l43_FlipperR, LM_Inserts_l43_FlipperRU, LM_Inserts_l43_FlipperSpL, LM_Inserts_l43_FlipperSpLU, LM_Inserts_l43_FlipperSpRU, LM_Inserts_l43_LSling1, LM_Inserts_l43_LSling2, LM_Inserts_l43_PF, LM_Inserts_l43_Parts, LM_Inserts_l43_RSling1, _
  LM_Inserts_l44_B_Caps_Top_Yello, LM_Inserts_l44_FlipperL, LM_Inserts_l44_FlipperLU, LM_Inserts_l44_FlipperR, LM_Inserts_l44_FlipperRU, LM_Inserts_l44_FlipperSpLU, LM_Inserts_l44_FlipperSpR, LM_Inserts_l44_FlipperSpRU, LM_Inserts_l44_LSling1, LM_Inserts_l44_PF, LM_Inserts_l44_Parts, LM_Inserts_l44_RSling1, LM_Inserts_l44_RSling2, LM_Inserts_l45_B_Caps_Top_Yello, LM_Inserts_l45_FlipperL, LM_Inserts_l45_FlipperLU, LM_Inserts_l45_FlipperR, LM_Inserts_l45_FlipperRU, LM_Inserts_l45_FlipperSpLU, LM_Inserts_l45_FlipperSpR, LM_Inserts_l45_FlipperSpRU, LM_Inserts_l45_PF, LM_Inserts_l45_Parts, LM_Inserts_l45_RSling1, LM_Inserts_l45_RSling2, LM_Inserts_l45_SLING1, LM_Inserts_l46_FlipperR, LM_Inserts_l46_FlipperRU, LM_Inserts_l46_FlipperSpR, LM_Inserts_l46_FlipperSpRU, LM_Inserts_l46_PF, LM_Inserts_l46_Parts, LM_Inserts_l47_FlipperL, LM_Inserts_l47_FlipperLU, LM_Inserts_l47_FlipperR, LM_Inserts_l47_FlipperRU, LM_Inserts_l47_FlipperSpL, LM_Inserts_l47_FlipperSpLU, LM_Inserts_l47_FlipperSpR, LM_Inserts_l47_FlipperSpRU, _
  LM_Inserts_l47_PF, LM_Inserts_l47_Parts, LM_Inserts_l48_FlipperR1U, LM_Inserts_l48_Outlanes, LM_Inserts_l48_PF, LM_Inserts_l48_Parts, LM_Inserts_l48_RSling2, LM_Inserts_l48_RocketToy, LM_Inserts_l51_B_Caps_Top_Amber, LM_Inserts_l51_Clock_Color, LM_Inserts_l51_Clock_White, LM_Inserts_l51_ClockToy, LM_Inserts_l51_Parts, LM_Inserts_l51_Robot, LM_Inserts_l51_Sign_Spiral, LM_Inserts_l51_sw47m, LM_Inserts_l51_sw68, LM_Inserts_l52_B_Caps_Bot_Amber, LM_Inserts_l52_B_Caps_Top_Amber, LM_Inserts_l52_Clock_Color, LM_Inserts_l52_Clock_White, LM_Inserts_l52_ClockToy, LM_Inserts_l52_FlipperL1, LM_Inserts_l52_PF_Upper, LM_Inserts_l52_Parts, LM_Inserts_l52_Sign_Spiral, LM_Inserts_l52_sw68, LM_Inserts_l53_B_Caps_Bot_Amber, LM_Inserts_l53_B_Caps_Top_Amber, LM_Inserts_l53_Clock_Color, LM_Inserts_l53_Clock_White, LM_Inserts_l53_ClockToy, LM_Inserts_l53_FlipperL1, LM_Inserts_l53_Over3, LM_Inserts_l53_PF_Upper, LM_Inserts_l53_Parts, LM_Inserts_l53_Sign_Spiral, LM_Inserts_l54_Clock_White, LM_Inserts_l54_ClockToy, LM_Inserts_l54_PF, _
  LM_Inserts_l54_PF_Upper, LM_Inserts_l54_Parts, LM_Inserts_l54_Sign_Spiral, LM_Inserts_l55_B_Caps_Top_Amber, LM_Inserts_l55_Clock_Color, LM_Inserts_l55_Clock_White, LM_Inserts_l55_ClockToy, LM_Inserts_l55_FlipperL1, LM_Inserts_l55_FlipperSpL1, LM_Inserts_l55_PF_Upper, LM_Inserts_l55_Parts, LM_Inserts_l55_Sign_Spiral, LM_Inserts_l56_B_Caps_Bot_Amber, LM_Inserts_l56_Clock_Color, LM_Inserts_l56_Clock_White, LM_Inserts_l56_ClockToy, LM_Inserts_l56_Over3, LM_Inserts_l56_PF, LM_Inserts_l56_Parts, LM_Inserts_l56_Piano, LM_Inserts_l56_Sign_Spiral, LM_Inserts_l56_sw64, LM_Inserts_l56_sw68, LM_Inserts_l57_Over2, LM_Inserts_l57_Over3, LM_Inserts_l57_PF, LM_Inserts_l57_Parts, LM_Inserts_l57_sw64, LM_Inserts_l58_Clock_Color, LM_Inserts_l58_Clock_White, LM_Inserts_l58_Over2, LM_Inserts_l58_Parts, LM_Inserts_l64_B_Caps_Bot_Red, LM_Inserts_l64_B_Caps_Top_Amber, LM_Inserts_l64_B_Caps_Top_Yello, LM_Inserts_l64_BR2, LM_Inserts_l64_BS2, LM_Inserts_l64_BumperPegs, LM_Inserts_l64_PF, LM_Inserts_l64_PF_Upper, LM_Inserts_l64_Parts, _
  LM_Inserts_l64_TownSquarePost, LM_Inserts_l64_sw77, LM_Inserts_l64_sw77m, LM_Inserts_l64_sw78m, LM_Inserts_l65_B_Caps_Top_Amber, LM_Inserts_l65_B_Caps_Top_Red, LM_Inserts_l65_B_Caps_Top_Yello, LM_Inserts_l65_BR2, LM_Inserts_l65_BS2, LM_Inserts_l65_PF, LM_Inserts_l65_PF_Upper, LM_Inserts_l65_Parts, LM_Inserts_l65_sw77, LM_Inserts_l65_sw77m, LM_Inserts_l65_sw78, LM_Inserts_l65_sw78m, LM_Inserts_l66_Outlanes, LM_Inserts_l66_Parts, LM_Inserts_l66_RocketToy, LM_Inserts_l66_sw11, LM_Inserts_l67_Clock_Color, LM_Inserts_l67_Clock_White, LM_Inserts_l67_Over1, LM_Inserts_l67_Parts, LM_Inserts_l67_SlotMachineToy, LM_Inserts_l68_Clock_Color, LM_Inserts_l68_Clock_White, LM_Inserts_l68_Over1, LM_Inserts_l68_Over3, LM_Inserts_l68_Parts, LM_Inserts_l71_B_Caps_Top_Amber, LM_Inserts_l71_B_Caps_Top_Yello, LM_Inserts_l71_FlipperR1U, LM_Inserts_l71_FlipperSpR1U, LM_Inserts_l71_PF, LM_Inserts_l71_Parts, LM_Inserts_l71_SlotMachineToy, LM_Inserts_l71_sw47, LM_Inserts_l71_sw66, LM_Inserts_l71_sw67, LM_Inserts_l71_sw67m, _
  LM_Inserts_l71_sw68m, LM_Inserts_l72_B_Caps_Top_Amber, LM_Inserts_l72_Clock_Color, LM_Inserts_l72_Clock_White, LM_Inserts_l72_ClockToy, LM_Inserts_l72_PF, LM_Inserts_l72_Parts, LM_Inserts_l72_Sign_Spiral, LM_Inserts_l72_sw66, LM_Inserts_l72_sw67, LM_Inserts_l72_sw68m, LM_Inserts_l73_B_Caps_Bot_Amber, LM_Inserts_l73_B_Caps_Top_Amber, LM_Inserts_l73_Clock_Color, LM_Inserts_l73_Clock_White, LM_Inserts_l73_ClockToy, LM_Inserts_l73_PF, LM_Inserts_l73_Parts, LM_Inserts_l73_Sign_Spiral, LM_Inserts_l73_sw66, LM_Inserts_l73_sw66m, LM_Inserts_l73_sw67, LM_Inserts_l73_sw67m, LM_Inserts_l75_B_Caps_Bot_Amber, LM_Inserts_l75_B_Caps_Top_Amber, LM_Inserts_l75_Clock_Color, LM_Inserts_l75_Clock_White, LM_Inserts_l75_Over3, LM_Inserts_l75_PF, LM_Inserts_l75_Parts, LM_Inserts_l75_Piano, LM_Inserts_l75_sw64, LM_Inserts_l75_sw64m, LM_Inserts_l76_PF_Upper, LM_Inserts_l76_Parts, LM_Inserts_l76_Pyramid, LM_Inserts_l77_PF_Upper, LM_Inserts_l77_Parts, LM_Inserts_l78_PF_Upper, LM_Inserts_l78_Parts, LM_Inserts_l78_Pyramid, _
  LM_Inserts_l81_B_Caps_Bot_Amber, LM_Inserts_l81_ClockToy, LM_Inserts_l81_FlipperL1, LM_Inserts_l81_FlipperSpL1, LM_Inserts_l81_Over2, LM_Inserts_l81_PF, LM_Inserts_l81_PF_Upper, LM_Inserts_l81_Parts, LM_Inserts_l81_Sign_Spiral, LM_Inserts_l82_B_Caps_Top_Amber, LM_Inserts_l82_Clock_Color, LM_Inserts_l82_Clock_White, LM_Inserts_l82_ClockToy, LM_Inserts_l82_Over3, LM_Inserts_l82_PF, LM_Inserts_l82_Parts, LM_Inserts_l82_Piano, LM_Inserts_l82_Robot, LM_Inserts_l82_Sign_Spiral, LM_Inserts_l82_sw47m, LM_Inserts_l82_sw64, LM_Inserts_l82_sw68, LM_Inserts_l83_Clock_Color, LM_Inserts_l83_Clock_White, LM_Inserts_l83_ClockToy, LM_Inserts_l83_Over1, LM_Inserts_l83_Over2, LM_Inserts_l83_Over3, LM_Inserts_l83_PF, LM_Inserts_l83_Parts, LM_Inserts_l83_Piano, LM_Inserts_l83_sw64, LM_Inserts_l83_sw64m, LM_Inserts_l84_Clock_Color, LM_Inserts_l84_Clock_White, LM_Inserts_l84_ClockLarge, LM_Inserts_l84_ClockShort, LM_Inserts_l84_ClockToy, LM_Inserts_l84_Over1, LM_Inserts_l84_Over2, LM_Inserts_l84_Over3, LM_Inserts_l84_PF, _
  LM_Inserts_l84_Parts, LM_Inserts_l84_Piano, LM_Inserts_l84_sw64, LM_Inserts_l84_sw64m, LM_Inserts_l85_Clock_Color, LM_Inserts_l85_Clock_White, LM_Inserts_l85_FlipperR1U, LM_Inserts_l85_Over2, LM_Inserts_l85_Over3, LM_Inserts_l85_PF, LM_Inserts_l85_Parts, LM_Inserts_l85_RocketToy, LM_Inserts_l85_SlotMachineToy, LM_Inserts_l85_sw65m, LM_Inserts_l85_sw67m, LM_Inserts_l86_FlipperR1, LM_Inserts_l86_FlipperR1U, LM_Inserts_l86_FlipperSpR1, LM_Inserts_l86_FlipperSpR1U, LM_Inserts_l86_InvaderToy, LM_Inserts_l86_Over1, LM_Inserts_l86_Over3, LM_Inserts_l86_PF, LM_Inserts_l86_Parts, LM_Inserts_l86_SlotMachineToy, LM_Mod_l105_InvaderToy, LM_Mod_l105_Over1, LM_Mod_l105_Parts, LM_Mod_l106_FlipperR1, LM_Mod_l106_FlipperR1U, LM_Mod_l106_FlipperSpR1, LM_Mod_l106_FlipperSpR1U, LM_Mod_l106_InvaderToy, LM_Mod_l106_Over1, LM_Mod_l106_PF, LM_Mod_l106_Parts, LM_Mod_l106_SlotMachineToy, LM_Mod_l107_FlipperR1, LM_Mod_l107_FlipperR1U, LM_Mod_l107_FlipperSpR1, LM_Mod_l107_FlipperSpR1U, LM_Mod_l107_InvaderToy, LM_Mod_l107_PF, _
  LM_Mod_l107_Parts, LM_Mod_l108_PF, LM_Mod_l108_Parts, LM_Mod_l108_sw68m, LM_Mod_l109_B_Caps_Bot_Amber, LM_Mod_l109_B_Caps_Bot_Red, LM_Mod_l109_B_Caps_Bot_Yellow, LM_Mod_l109_B_Caps_Top_Amber, LM_Mod_l109_B_Caps_Top_Yellow, LM_Mod_l109_BR1, LM_Mod_l109_BR2, LM_Mod_l109_BR3, LM_Mod_l109_BS2, LM_Mod_l109_BumperPegs, LM_Mod_l109_PF, LM_Mod_l109_PF_Upper, LM_Mod_l109_Parts, LM_Mod_l109_TownSquarePost, LM_Mod_l109_sw48m, LM_Mod_l109_sw77, LM_Mod_l109_sw77m, LM_Mod_l109_sw78, LM_Mod_l110_B_Caps_Bot_Amber, LM_Mod_l110_BR2, LM_Mod_l110_BS2, LM_Mod_l110_Camera, LM_Mod_l110_PF, LM_Mod_l110_PF_Upper, LM_Mod_l110_Parts, LM_Mod_l111_BR1, LM_Mod_l111_FlipperR1, LM_Mod_l111_FlipperR1U, LM_Mod_l111_FlipperSpR1, LM_Mod_l111_FlipperSpR1U, LM_Mod_l111_Outlanes, LM_Mod_l111_PF, LM_Mod_l111_Parts, LM_Mod_l111_RSling1, LM_Mod_l111_RSling2, LM_Mod_l111_RocketToy, LM_Mod_l111_SLING1)
Dim BG_All: BG_All=Array(BM_B_Caps_Bot_Amber, BM_B_Caps_Bot_Red, BM_B_Caps_Bot_Yellow, BM_B_Caps_Top_Amber, BM_B_Caps_Top_Red, BM_B_Caps_Top_Yellow, BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_BumperPegs, BM_Camera, BM_Clock_Color, BM_Clock_White, BM_ClockLarge, BM_ClockShort, BM_ClockToy, BM_DiverterP, BM_DiverterP1, BM_FlipperL, BM_FlipperL1, BM_FlipperLU, BM_FlipperR, BM_FlipperR1, BM_FlipperR1U, BM_FlipperRU, BM_FlipperSpL, BM_FlipperSpL1, BM_FlipperSpLU, BM_FlipperSpR, BM_FlipperSpR1, BM_FlipperSpR1U, BM_FlipperSpRU, BM_GMKnob, BM_Gate1, BM_Gate2, BM_Gumballs, BM_InvaderToy, BM_LSling1, BM_LSling2, BM_MysticSeerToy, BM_Outlanes, BM_Over1, BM_Over2, BM_Over3, BM_Over4, BM_Over5, BM_PF, BM_PF_Upper, BM_Parts, BM_Piano, BM_Pyramid, BM_RDiv, BM_RSling1, BM_RSling2, BM_Rails, BM_Rails_DT, BM_Robot, BM_RocketToy, BM_SLING1, BM_SLING2, BM_ShooterDiv, BM_SideMod, BM_Sign_Spiral, BM_SlotMachineToy, BM_SpiralToy, BM_TVtoy, BM_TownSquarePost, BM_URMagnet, BM_sw11, BM_sw12, BM_sw27, BM_sw36, BM_sw37, BM_sw38, _
  BM_sw47, BM_sw47m, BM_sw48, BM_sw48m, BM_sw53p, BM_sw54p, BM_sw56, BM_sw61, BM_sw62, BM_sw63, BM_sw64, BM_sw64m, BM_sw65, BM_sw65a, BM_sw65am, BM_sw65m, BM_sw66, BM_sw66m, BM_sw67, BM_sw67m, BM_sw68, BM_sw68m, BM_sw77, BM_sw77m, BM_sw78, BM_sw78m, LM_Bumpers_l61_B_Caps_Bot_Amber, LM_Bumpers_l61_B_Caps_Bot_Red, LM_Bumpers_l61_B_Caps_Bot_Yello, LM_Bumpers_l61_B_Caps_Top_Amber, LM_Bumpers_l61_B_Caps_Top_Red, LM_Bumpers_l61_B_Caps_Top_Yello, LM_Bumpers_l61_BR1, LM_Bumpers_l61_BR2, LM_Bumpers_l61_BR3, LM_Bumpers_l61_BS1, LM_Bumpers_l61_BS2, LM_Bumpers_l61_BS3, LM_Bumpers_l61_BumperPegs, LM_Bumpers_l61_ClockToy, LM_Bumpers_l61_MysticSeerToy, LM_Bumpers_l61_Over3, LM_Bumpers_l61_PF, LM_Bumpers_l61_PF_Upper, LM_Bumpers_l61_Parts, LM_Bumpers_l61_Rails, LM_Bumpers_l61_SideMod, LM_Bumpers_l61_SlotMachineToy, LM_Bumpers_l61_sw36, LM_Bumpers_l61_sw37, LM_Bumpers_l62_B_Caps_Bot_Yello, LM_Bumpers_l62_B_Caps_Top_Amber, LM_Bumpers_l62_B_Caps_Top_Red, LM_Bumpers_l62_B_Caps_Top_Yello, LM_Bumpers_l62_BR1, LM_Bumpers_l62_BR2, _
  LM_Bumpers_l62_BR3, LM_Bumpers_l62_BS1, LM_Bumpers_l62_BS2, LM_Bumpers_l62_BS3, LM_Bumpers_l62_BumperPegs, LM_Bumpers_l62_MysticSeerToy, LM_Bumpers_l62_Outlanes, LM_Bumpers_l62_Over3, LM_Bumpers_l62_PF, LM_Bumpers_l62_PF_Upper, LM_Bumpers_l62_Parts, LM_Bumpers_l62_RocketToy, LM_Bumpers_l62_SideMod, LM_Bumpers_l62_SlotMachineToy, LM_Bumpers_l62_sw48, LM_Bumpers_l62_sw48m, LM_Bumpers_l62_sw77, LM_Bumpers_l62_sw77m, LM_Bumpers_l63_B_Caps_Bot_Amber, LM_Bumpers_l63_B_Caps_Bot_Red, LM_Bumpers_l63_B_Caps_Bot_Yello, LM_Bumpers_l63_B_Caps_Top_Amber, LM_Bumpers_l63_B_Caps_Top_Red, LM_Bumpers_l63_B_Caps_Top_Yello, LM_Bumpers_l63_BR1, LM_Bumpers_l63_BR2, LM_Bumpers_l63_BR3, LM_Bumpers_l63_BS1, LM_Bumpers_l63_BS2, LM_Bumpers_l63_BS3, LM_Bumpers_l63_BumperPegs, LM_Bumpers_l63_Camera, LM_Bumpers_l63_Clock_Color, LM_Bumpers_l63_Clock_White, LM_Bumpers_l63_FlipperL1, LM_Bumpers_l63_FlipperSpL1, LM_Bumpers_l63_Over3, LM_Bumpers_l63_PF, LM_Bumpers_l63_PF_Upper, LM_Bumpers_l63_Parts, LM_Bumpers_l63_Rails, _
  LM_Bumpers_l63_Sign_Spiral, LM_Bumpers_l63_SlotMachineToy, LM_Bumpers_l63_TownSquarePost, LM_Bumpers_l63_sw77, LM_Bumpers_l63_sw77m, LM_Bumpers_l63_sw78, LM_Bumpers_l63_sw78m, LM_Flashers_f17_B_Caps_Bot_Ambe, LM_Flashers_f17_B_Caps_Bot_Red, LM_Flashers_f17_B_Caps_Bot_Yell, LM_Flashers_f17_B_Caps_Top_Ambe, LM_Flashers_f17_B_Caps_Top_Red, LM_Flashers_f17_B_Caps_Top_Yell, LM_Flashers_f17_BR1, LM_Flashers_f17_BR2, LM_Flashers_f17_BR3, LM_Flashers_f17_BS1, LM_Flashers_f17_BS2, LM_Flashers_f17_BS3, LM_Flashers_f17_BumperPegs, LM_Flashers_f17_Camera, LM_Flashers_f17_Clock_Color, LM_Flashers_f17_Clock_White, LM_Flashers_f17_ClockToy, LM_Flashers_f17_FlipperLU, LM_Flashers_f17_FlipperR1, LM_Flashers_f17_FlipperR1U, LM_Flashers_f17_FlipperRU, LM_Flashers_f17_FlipperSpL, LM_Flashers_f17_FlipperSpLU, LM_Flashers_f17_MysticSeerToy, LM_Flashers_f17_Outlanes, LM_Flashers_f17_Over1, LM_Flashers_f17_Over3, LM_Flashers_f17_PF, LM_Flashers_f17_PF_Upper, LM_Flashers_f17_Parts, LM_Flashers_f17_Rails, LM_Flashers_f17_RocketToy, _
  LM_Flashers_f17_SideMod, LM_Flashers_f17_SlotMachineToy, LM_Flashers_f17_TownSquarePost, LM_Flashers_f17_sw36, LM_Flashers_f17_sw37, LM_Flashers_f17_sw38, LM_Flashers_f17_sw47, LM_Flashers_f17_sw48, LM_Flashers_f17_sw48m, LM_Flashers_f17_sw77, LM_Flashers_f17_sw77m, LM_Flashers_f17_sw78, LM_Flashers_f17_sw78m, LM_Flashers_f18_B_Caps_Bot_Ambe, LM_Flashers_f18_B_Caps_Bot_Yell, LM_Flashers_f18_B_Caps_Top_Red, LM_Flashers_f18_B_Caps_Top_Yell, LM_Flashers_f18_Clock_Color, LM_Flashers_f18_Clock_White, LM_Flashers_f18_ClockToy, LM_Flashers_f18_DiverterP, LM_Flashers_f18_FlipperL, LM_Flashers_f18_FlipperL1, LM_Flashers_f18_FlipperLU, LM_Flashers_f18_FlipperR1, LM_Flashers_f18_FlipperR1U, LM_Flashers_f18_FlipperRU, LM_Flashers_f18_FlipperSpL, LM_Flashers_f18_FlipperSpL1, LM_Flashers_f18_FlipperSpLU, LM_Flashers_f18_FlipperSpR1, LM_Flashers_f18_FlipperSpR1U, LM_Flashers_f18_FlipperSpRU, LM_Flashers_f18_GMKnob, LM_Flashers_f18_Gate1, LM_Flashers_f18_Gumballs, LM_Flashers_f18_InvaderToy, LM_Flashers_f18_LSling1, _
  LM_Flashers_f18_LSling2, LM_Flashers_f18_Over1, LM_Flashers_f18_Over2, LM_Flashers_f18_Over3, LM_Flashers_f18_PF, LM_Flashers_f18_PF_Upper, LM_Flashers_f18_Parts, LM_Flashers_f18_Piano, LM_Flashers_f18_Pyramid, LM_Flashers_f18_Rails, LM_Flashers_f18_Robot, LM_Flashers_f18_RocketToy, LM_Flashers_f18_SideMod, LM_Flashers_f18_Sign_Spiral, LM_Flashers_f18_SlotMachineToy, LM_Flashers_f18_TVtoy, LM_Flashers_f18_TownSquarePost, LM_Flashers_f18_sw47, LM_Flashers_f18_sw63, LM_Flashers_f18_sw64, LM_Flashers_f18_sw65am, LM_Flashers_f18_sw65m, LM_Flashers_f18_sw67m, LM_Flashers_f18_sw68, LM_Flashers_f19_BR2, LM_Flashers_f19_BR3, LM_Flashers_f19_BS3, LM_Flashers_f19_Camera, LM_Flashers_f19_Clock_Color, LM_Flashers_f19_Clock_White, LM_Flashers_f19_ClockLarge, LM_Flashers_f19_ClockToy, LM_Flashers_f19_FlipperL1, LM_Flashers_f19_FlipperR1, LM_Flashers_f19_FlipperR1U, LM_Flashers_f19_FlipperSpL1, LM_Flashers_f19_FlipperSpR1, LM_Flashers_f19_FlipperSpR1U, LM_Flashers_f19_GMKnob, LM_Flashers_f19_Gumballs, LM_Flashers_f19_Over1, _
  LM_Flashers_f19_Over2, LM_Flashers_f19_Over3, LM_Flashers_f19_PF, LM_Flashers_f19_PF_Upper, LM_Flashers_f19_Parts, LM_Flashers_f19_Piano, LM_Flashers_f19_Pyramid, LM_Flashers_f19_RDiv, LM_Flashers_f19_Rails, LM_Flashers_f19_Robot, LM_Flashers_f19_RocketToy, LM_Flashers_f19_SLING2, LM_Flashers_f19_SideMod, LM_Flashers_f19_Sign_Spiral, LM_Flashers_f19_SlotMachineToy, LM_Flashers_f19_SpiralToy, LM_Flashers_f19_sw38, LM_Flashers_f19_sw47, LM_Flashers_f19_sw56, LM_Flashers_f19_sw64, LM_Flashers_f19_sw64m, LM_Flashers_f19_sw65, LM_Flashers_f19_sw65a, LM_Flashers_f19_sw65am, LM_Flashers_f19_sw65m, LM_Flashers_f19_sw66, LM_Flashers_f19_sw66m, LM_Flashers_f19_sw67, LM_Flashers_f19_sw67m, LM_Flashers_f20_B_Caps_Bot_Ambe, LM_Flashers_f20_BS2, LM_Flashers_f20_Camera, LM_Flashers_f20_Clock_Color, LM_Flashers_f20_Clock_White, LM_Flashers_f20_ClockToy, LM_Flashers_f20_DiverterP1, LM_Flashers_f20_FlipperL, LM_Flashers_f20_FlipperL1, LM_Flashers_f20_FlipperLU, LM_Flashers_f20_FlipperR, LM_Flashers_f20_FlipperR1, _
  LM_Flashers_f20_FlipperR1U, LM_Flashers_f20_FlipperRU, LM_Flashers_f20_FlipperSpL1, LM_Flashers_f20_FlipperSpR1U, LM_Flashers_f20_FlipperSpRU, LM_Flashers_f20_GMKnob, LM_Flashers_f20_Gumballs, LM_Flashers_f20_InvaderToy, LM_Flashers_f20_Over1, LM_Flashers_f20_Over2, LM_Flashers_f20_Over3, LM_Flashers_f20_Over4, LM_Flashers_f20_Over5, LM_Flashers_f20_PF, LM_Flashers_f20_PF_Upper, LM_Flashers_f20_Parts, LM_Flashers_f20_Piano, LM_Flashers_f20_Pyramid, LM_Flashers_f20_RDiv, LM_Flashers_f20_RSling1, LM_Flashers_f20_Rails, LM_Flashers_f20_Rails_DT, LM_Flashers_f20_Robot, LM_Flashers_f20_SideMod, LM_Flashers_f20_Sign_Spiral, LM_Flashers_f20_SlotMachineToy, LM_Flashers_f20_SpiralToy, LM_Flashers_f20_TVtoy, LM_Flashers_f20_sw53p, LM_Flashers_f20_sw56, LM_Flashers_f20_sw65am, LM_Flashers_f20_sw78, LM_Flashers_f28_B_Caps_Bot_Ambe, LM_Flashers_f28_B_Caps_Top_Ambe, LM_Flashers_f28_B_Caps_Top_Yell, LM_Flashers_f28_BR2, LM_Flashers_f28_Camera, LM_Flashers_f28_Clock_Color, LM_Flashers_f28_Clock_White, _
  LM_Flashers_f28_ClockLarge, LM_Flashers_f28_ClockToy, LM_Flashers_f28_DiverterP, LM_Flashers_f28_DiverterP1, LM_Flashers_f28_FlipperL, LM_Flashers_f28_FlipperL1, LM_Flashers_f28_FlipperLU, LM_Flashers_f28_FlipperR, LM_Flashers_f28_FlipperR1, LM_Flashers_f28_FlipperR1U, LM_Flashers_f28_FlipperRU, LM_Flashers_f28_FlipperSpL, LM_Flashers_f28_FlipperSpL1, LM_Flashers_f28_FlipperSpLU, LM_Flashers_f28_FlipperSpR, LM_Flashers_f28_FlipperSpR1, LM_Flashers_f28_FlipperSpR1U, LM_Flashers_f28_FlipperSpRU, LM_Flashers_f28_GMKnob, LM_Flashers_f28_Gumballs, LM_Flashers_f28_LSling1, LM_Flashers_f28_Over1, LM_Flashers_f28_Over2, LM_Flashers_f28_Over3, LM_Flashers_f28_Over4, LM_Flashers_f28_Over5, LM_Flashers_f28_PF, LM_Flashers_f28_PF_Upper, LM_Flashers_f28_Parts, LM_Flashers_f28_Piano, LM_Flashers_f28_Pyramid, LM_Flashers_f28_RDiv, LM_Flashers_f28_RSling1, LM_Flashers_f28_RSling2, LM_Flashers_f28_Rails, LM_Flashers_f28_Robot, LM_Flashers_f28_RocketToy, LM_Flashers_f28_SideMod, LM_Flashers_f28_Sign_Spiral, _
  LM_Flashers_f28_SlotMachineToy, LM_Flashers_f28_TVtoy, LM_Flashers_f28_sw11, LM_Flashers_f28_sw38, LM_Flashers_f28_sw47m, LM_Flashers_f28_sw53p, LM_Flashers_f28_sw63, LM_Flashers_f28_sw64, LM_Flashers_f28_sw64m, LM_Flashers_f28_sw65, LM_Flashers_f28_sw65a, LM_Flashers_f28_sw65am, LM_Flashers_f28_sw65m, LM_Flashers_f28_sw66, LM_Flashers_f28_sw67, LM_Flashers_f28_sw68, LM_Flashers_f37_B_Caps_Bot_Ambe, LM_Flashers_f37_B_Caps_Top_Ambe, LM_Flashers_f37_B_Caps_Top_Yell, LM_Flashers_f37_Clock_Color, LM_Flashers_f37_Clock_White, LM_Flashers_f37_ClockToy, LM_Flashers_f37_FlipperR1, LM_Flashers_f37_FlipperR1U, LM_Flashers_f37_FlipperSpR1, LM_Flashers_f37_FlipperSpR1U, LM_Flashers_f37_InvaderToy, LM_Flashers_f37_Over1, LM_Flashers_f37_PF, LM_Flashers_f37_Parts, LM_Flashers_f37_Rails, LM_Flashers_f37_RocketToy, LM_Flashers_f37_Sign_Spiral, LM_Flashers_f37_SlotMachineToy, LM_Flashers_f37_sw63, LM_Flashers_f37_sw78, LM_Flashers_f38_BS1, LM_Flashers_f38_Clock_Color, LM_Flashers_f38_Clock_White, LM_Flashers_f38_ClockToy, _
  LM_Flashers_f38_FlipperL1, LM_Flashers_f38_FlipperR1, LM_Flashers_f38_FlipperR1U, LM_Flashers_f38_FlipperSpL1, LM_Flashers_f38_FlipperSpR1, LM_Flashers_f38_FlipperSpR1U, LM_Flashers_f38_GMKnob, LM_Flashers_f38_Gumballs, LM_Flashers_f38_Over1, LM_Flashers_f38_Over2, LM_Flashers_f38_Over3, LM_Flashers_f38_PF, LM_Flashers_f38_PF_Upper, LM_Flashers_f38_Parts, LM_Flashers_f38_Piano, LM_Flashers_f38_Pyramid, LM_Flashers_f38_RDiv, LM_Flashers_f38_RSling1, LM_Flashers_f38_RSling2, LM_Flashers_f38_Rails, LM_Flashers_f38_Rails_DT, LM_Flashers_f38_Robot, LM_Flashers_f38_SideMod, LM_Flashers_f38_Sign_Spiral, LM_Flashers_f38_SlotMachineToy, LM_Flashers_f38_sw56, LM_Flashers_f39_Clock_White, LM_Flashers_f39_ClockToy, LM_Flashers_f39_FlipperL1, LM_Flashers_f39_FlipperR, LM_Flashers_f39_FlipperR1, LM_Flashers_f39_FlipperR1U, LM_Flashers_f39_FlipperRU, LM_Flashers_f39_FlipperSpL1, LM_Flashers_f39_FlipperSpR, LM_Flashers_f39_FlipperSpR1, LM_Flashers_f39_FlipperSpR1U, LM_Flashers_f39_FlipperSpRU, LM_Flashers_f39_GMKnob, _
  LM_Flashers_f39_Gumballs, LM_Flashers_f39_Over1, LM_Flashers_f39_Over2, LM_Flashers_f39_Over3, LM_Flashers_f39_Over4, LM_Flashers_f39_PF, LM_Flashers_f39_PF_Upper, LM_Flashers_f39_Parts, LM_Flashers_f39_Piano, LM_Flashers_f39_Pyramid, LM_Flashers_f39_RDiv, LM_Flashers_f39_RSling1, LM_Flashers_f39_RSling2, LM_Flashers_f39_Rails, LM_Flashers_f39_Rails_DT, LM_Flashers_f39_Robot, LM_Flashers_f39_SLING1, LM_Flashers_f39_SideMod, LM_Flashers_f39_Sign_Spiral, LM_Flashers_f39_SlotMachineToy, LM_Flashers_f39_sw12, LM_Flashers_f39_sw56, LM_Flashers_f40_ClockToy, LM_Flashers_f40_FlipperL1, LM_Flashers_f40_GMKnob, LM_Flashers_f40_Gumballs, LM_Flashers_f40_Over2, LM_Flashers_f40_Over3, LM_Flashers_f40_PF, LM_Flashers_f40_PF_Upper, LM_Flashers_f40_Parts, LM_Flashers_f40_RDiv, LM_Flashers_f40_Rails, LM_Flashers_f40_Robot, LM_Flashers_f40_SideMod, LM_Flashers_f40_Sign_Spiral, LM_Flashers_f41_Clock_Color, LM_Flashers_f41_Clock_White, LM_Flashers_f41_ClockToy, LM_Flashers_f41_DiverterP, LM_Flashers_f41_DiverterP1, _
  LM_Flashers_f41_FlipperL1, LM_Flashers_f41_FlipperR1, LM_Flashers_f41_FlipperR1U, LM_Flashers_f41_FlipperRU, LM_Flashers_f41_FlipperSpL1, LM_Flashers_f41_FlipperSpR1, LM_Flashers_f41_FlipperSpR1U, LM_Flashers_f41_GMKnob, LM_Flashers_f41_Over1, LM_Flashers_f41_Over2, LM_Flashers_f41_Over3, LM_Flashers_f41_Over4, LM_Flashers_f41_PF, LM_Flashers_f41_PF_Upper, LM_Flashers_f41_Parts, LM_Flashers_f41_Piano, LM_Flashers_f41_Pyramid, LM_Flashers_f41_RDiv, LM_Flashers_f41_Rails, LM_Flashers_f41_Rails_DT, LM_Flashers_f41_Robot, LM_Flashers_f41_RocketToy, LM_Flashers_f41_SideMod, LM_Flashers_f41_SlotMachineToy, LM_Flashers_f41_TVtoy, LM_Flashers_f41_URMagnet, LM_Flashers_f41_sw54p, LM_Flashers_l74_Clock_Color, LM_Flashers_l74_Clock_White, LM_Flashers_l74_ClockShort, LM_Flashers_l74_ClockToy, LM_Flashers_l74_Over1, LM_Flashers_l74_Over2, LM_Flashers_l74_Over3, LM_Flashers_l74_PF, LM_Flashers_l74_Parts, LM_Flashers_l74_Piano, LM_Flashers_l74_Robot, LM_Flashers_l74_sw64, LM_Flashers_l74_sw66, _
  LM_GI_Clock_l102_B_Caps_Bot_Red, LM_GI_Clock_l102_B_Caps_Top_Amb, LM_GI_Clock_l102_B_Caps_Top_Yel, LM_GI_Clock_l102_Clock_Color, LM_GI_Clock_l102_Clock_White, LM_GI_Clock_l102_ClockLarge, LM_GI_Clock_l102_ClockShort, LM_GI_Clock_l102_ClockToy, LM_GI_Clock_l102_DiverterP, LM_GI_Clock_l102_FlipperL1, LM_GI_Clock_l102_FlipperR1U, LM_GI_Clock_l102_Over1, LM_GI_Clock_l102_Over2, LM_GI_Clock_l102_Over3, LM_GI_Clock_l102_PF, LM_GI_Clock_l102_Parts, LM_GI_Clock_l102_Piano, LM_GI_Clock_l102_Rails, LM_GI_Clock_l102_SideMod, LM_GI_Clock_l102_Sign_Spiral, LM_GI_Clock_l102_SlotMachineToy, LM_GI_Clock_l102_sw63, LM_GI_Clock_l102_sw64, LM_GI_Clock_l102_sw64m, LM_GI_Clock_l102_sw65, LM_GI_Clock_l102_sw65a, LM_GI_Clock_l102_sw65am, LM_GI_Clock_l102_sw65m, LM_GI_Left_l100_B_Caps_Bot_Red, LM_GI_Left_l100_B_Caps_Top_Ambe, LM_GI_Left_l100_B_Caps_Top_Red, LM_GI_Left_l100_BR1, LM_GI_Left_l100_BR2, LM_GI_Left_l100_BS1, LM_GI_Left_l100_BS2, LM_GI_Left_l100_BS3, LM_GI_Left_l100_BumperPegs, LM_GI_Left_l100_ClockToy, _
  LM_GI_Left_l100_FlipperR1U, LM_GI_Left_l100_GMKnob, LM_GI_Left_l100_Gumballs, LM_GI_Left_l100_Over1, LM_GI_Left_l100_Over2, LM_GI_Left_l100_Over3, LM_GI_Left_l100_Over4, LM_GI_Left_l100_Over5, LM_GI_Left_l100_PF, LM_GI_Left_l100_PF_Upper, LM_GI_Left_l100_Parts, LM_GI_Left_l100_RDiv, LM_GI_Left_l100_Rails, LM_GI_Left_l100_Robot, LM_GI_Left_l100_SideMod, LM_GI_Left_l100_Sign_Spiral, LM_GI_Left_l100_SlotMachineToy, LM_GI_Left_l100_TownSquarePost, LM_GI_Left_l100_sw36, LM_GI_Left_l100_sw37, LM_GI_Left_l100_sw38, LM_GI_Left_l100_sw53p, LM_GI_Left_l100_sw56, LM_GI_Left_l100_sw68, LM_GI_Left_l100_sw77, LM_GI_Left_l100a_B_Caps_Top_Yel, LM_GI_Left_l100a_ClockToy, LM_GI_Left_l100a_FlipperL, LM_GI_Left_l100a_FlipperLU, LM_GI_Left_l100a_FlipperR, LM_GI_Left_l100a_FlipperR1, LM_GI_Left_l100a_FlipperRU, LM_GI_Left_l100a_FlipperSpL, LM_GI_Left_l100a_FlipperSpLU, LM_GI_Left_l100a_FlipperSpR, LM_GI_Left_l100a_FlipperSpRU, LM_GI_Left_l100a_LSling1, LM_GI_Left_l100a_LSling2, LM_GI_Left_l100a_MysticSeerToy, _
  LM_GI_Left_l100a_Outlanes, LM_GI_Left_l100a_PF, LM_GI_Left_l100a_Parts, LM_GI_Left_l100a_RSling1, LM_GI_Left_l100a_RSling2, LM_GI_Left_l100a_Rails, LM_GI_Left_l100a_RocketToy, LM_GI_Left_l100a_SLING2, LM_GI_Left_l100a_sw36, LM_GI_Left_l100a_sw38, LM_GI_Left_l100a_sw48, LM_GI_Left_l100a_sw48m, LM_GI_Left_l100a_sw77, LM_GI_Left_l100b_B_Caps_Top_Amb, LM_GI_Left_l100b_BR2, LM_GI_Left_l100b_BS3, LM_GI_Left_l100b_ClockToy, LM_GI_Left_l100b_FlipperL, LM_GI_Left_l100b_FlipperLU, LM_GI_Left_l100b_FlipperR, LM_GI_Left_l100b_FlipperR1U, LM_GI_Left_l100b_FlipperSpL, LM_GI_Left_l100b_FlipperSpLU, LM_GI_Left_l100b_LSling1, LM_GI_Left_l100b_LSling2, LM_GI_Left_l100b_MysticSeerToy, LM_GI_Left_l100b_Outlanes, LM_GI_Left_l100b_PF, LM_GI_Left_l100b_PF_Upper, LM_GI_Left_l100b_Parts, LM_GI_Left_l100b_SLING2, LM_GI_Left_l100b_SideMod, LM_GI_Left_l100b_SlotMachineToy, LM_GI_Left_l100b_TownSquarePost, LM_GI_Left_l100b_sw36, LM_GI_Left_l100b_sw37, LM_GI_Left_l100b_sw38, LM_GI_Left_l100b_sw48, LM_GI_Left_l100b_sw48m, _
  LM_GI_Left_l100b_sw77, LM_GI_Left_l100c_B_Caps_Top_Amb, LM_GI_Left_l100c_B_Caps_Top_Red, LM_GI_Left_l100c_B_Caps_Top_Yel, LM_GI_Left_l100c_BR3, LM_GI_Left_l100c_FlipperLU, LM_GI_Left_l100c_FlipperR, LM_GI_Left_l100c_FlipperSpLU, LM_GI_Left_l100c_PF, LM_GI_Left_l100c_Parts, LM_GI_Left_l100c_SLING2, LM_GI_Left_l100c_SideMod, LM_GI_Left_l100c_sw36, LM_GI_Left_l100c_sw37, LM_GI_Left_l100c_sw38, LM_GI_Left_l100c_sw48, LM_GI_Left_l100d_B_Caps_Top_Amb, LM_GI_Left_l100d_B_Caps_Top_Yel, LM_GI_Left_l100d_FlipperLU, LM_GI_Left_l100d_FlipperR, LM_GI_Left_l100d_FlipperRU, LM_GI_Left_l100d_FlipperSpLU, LM_GI_Left_l100d_PF, LM_GI_Left_l100d_Parts, LM_GI_Left_l100d_SLING2, LM_GI_Left_l100d_sw36, LM_GI_Left_l100d_sw37, LM_GI_Left_l100d_sw38, LM_GI_Left_l100d_sw48, LM_GI_Left_l100d_sw48m, LM_GI_Left_l100e_B_Caps_Top_Yel, LM_GI_Left_l100e_BR3, LM_GI_Left_l100e_FlipperL, LM_GI_Left_l100e_FlipperLU, LM_GI_Left_l100e_FlipperR, LM_GI_Left_l100e_FlipperRU, LM_GI_Left_l100e_FlipperSpL, LM_GI_Left_l100e_FlipperSpLU, _
  LM_GI_Left_l100e_FlipperSpR, LM_GI_Left_l100e_FlipperSpRU, LM_GI_Left_l100e_Outlanes, LM_GI_Left_l100e_PF, LM_GI_Left_l100e_Parts, LM_GI_Left_l100e_RocketToy, LM_GI_Left_l100e_SLING2, LM_GI_Left_l100e_sw38, LM_GI_Left_l100e_sw48, LM_GI_MiniPF_l101_GMKnob, LM_GI_MiniPF_l101_Gumballs, LM_GI_MiniPF_l101_PF, LM_GI_MiniPF_l101_PF_Upper, LM_GI_MiniPF_l101_Parts, LM_GI_MiniPF_l101_Pyramid, LM_GI_MiniPF_l101_SideMod, LM_GI_MiniPF_l101_Sign_Spiral, LM_GI_MiniPF_l101_SpiralToy, LM_GI_Right_l104_B_Caps_Top_Amb, LM_GI_Right_l104_BR2, LM_GI_Right_l104_BS1, LM_GI_Right_l104_Clock_Color, LM_GI_Right_l104_Clock_White, LM_GI_Right_l104_ClockToy, LM_GI_Right_l104_DiverterP, LM_GI_Right_l104_FlipperL1, LM_GI_Right_l104_FlipperR1, LM_GI_Right_l104_FlipperR1U, LM_GI_Right_l104_FlipperSpL1, LM_GI_Right_l104_FlipperSpR, LM_GI_Right_l104_FlipperSpR1, LM_GI_Right_l104_FlipperSpR1U, LM_GI_Right_l104_FlipperSpRU, LM_GI_Right_l104_Gate1, LM_GI_Right_l104_InvaderToy, LM_GI_Right_l104_Over1, LM_GI_Right_l104_Over2, LM_GI_Right_l104_Over3, _
  LM_GI_Right_l104_PF, LM_GI_Right_l104_PF_Upper, LM_GI_Right_l104_Parts, LM_GI_Right_l104_Piano, LM_GI_Right_l104_Rails, LM_GI_Right_l104_Robot, LM_GI_Right_l104_RocketToy, LM_GI_Right_l104_SideMod, LM_GI_Right_l104_Sign_Spiral, LM_GI_Right_l104_SlotMachineToy, LM_GI_Right_l104_TVtoy, LM_GI_Right_l104_URMagnet, LM_GI_Right_l104_sw47, LM_GI_Right_l104_sw61, LM_GI_Right_l104_sw62, LM_GI_Right_l104_sw63, LM_GI_Right_l104_sw64m, LM_GI_Right_l104_sw65, LM_GI_Right_l104_sw65a, LM_GI_Right_l104_sw65am, LM_GI_Right_l104_sw65m, LM_GI_Right_l104_sw66, LM_GI_Right_l104_sw66m, LM_GI_Right_l104_sw67, LM_GI_Right_l104_sw67m, LM_GI_Right_l104_sw68m, LM_GI_Right_l104_sw78, LM_GI_Right_l104a_B_Caps_Top_Re, LM_GI_Right_l104a_B_Caps_Top_Ye, LM_GI_Right_l104a_BR3, LM_GI_Right_l104a_ClockToy, LM_GI_Right_l104a_FlipperL, LM_GI_Right_l104a_FlipperLU, LM_GI_Right_l104a_FlipperR, LM_GI_Right_l104a_FlipperR1U, LM_GI_Right_l104a_FlipperRU, LM_GI_Right_l104a_FlipperSpL, LM_GI_Right_l104a_FlipperSpLU, LM_GI_Right_l104a_FlipperSpR, _
  LM_GI_Right_l104a_FlipperSpR1U, LM_GI_Right_l104a_FlipperSpRU, LM_GI_Right_l104a_InvaderToy, LM_GI_Right_l104a_LSling1, LM_GI_Right_l104a_LSling2, LM_GI_Right_l104a_Outlanes, LM_GI_Right_l104a_Over3, LM_GI_Right_l104a_PF, LM_GI_Right_l104a_Parts, LM_GI_Right_l104a_RSling1, LM_GI_Right_l104a_RSling2, LM_GI_Right_l104a_Rails, LM_GI_Right_l104a_RocketToy, LM_GI_Right_l104a_SLING1, LM_GI_Right_l104a_sw11, LM_GI_Right_l104a_sw12, LM_GI_Right_l104a_sw48, LM_GI_Right_l104b_ClockToy, LM_GI_Right_l104b_PF, LM_GI_Right_l104b_Parts, LM_GI_Right_l104b_Robot, LM_GI_Right_l104b_Sign_Spiral, LM_GI_Right_l104b_SlotMachineTo, LM_GI_Right_l104b_sw65m, LM_GI_Right_l104b_sw66, LM_GI_Right_l104b_sw66m, LM_GI_Right_l104b_sw67, LM_GI_Right_l104b_sw67m, LM_GI_Right_l104c_B_Caps_Top_Am, LM_GI_Right_l104c_B_Caps_Top_Ye, LM_GI_Right_l104c_BR2, LM_GI_Right_l104c_BR3, LM_GI_Right_l104c_FlipperL, LM_GI_Right_l104c_FlipperLU, LM_GI_Right_l104c_FlipperR, LM_GI_Right_l104c_FlipperRU, LM_GI_Right_l104c_FlipperSpL, _
  LM_GI_Right_l104c_FlipperSpLU, LM_GI_Right_l104c_FlipperSpR, LM_GI_Right_l104c_FlipperSpRU, LM_GI_Right_l104c_Outlanes, LM_GI_Right_l104c_PF, LM_GI_Right_l104c_Parts, LM_GI_Right_l104c_Rails, LM_GI_Right_l104c_SLING1, LM_GI_Right_l104c_sw11, LM_GI_Right_l104d_FlipperL, LM_GI_Right_l104d_FlipperLU, LM_GI_Right_l104d_FlipperR, LM_GI_Right_l104d_FlipperR1U, LM_GI_Right_l104d_FlipperRU, LM_GI_Right_l104d_FlipperSpL, LM_GI_Right_l104d_FlipperSpLU, LM_GI_Right_l104d_FlipperSpR, LM_GI_Right_l104d_FlipperSpR1U, LM_GI_Right_l104d_FlipperSpRU, LM_GI_Right_l104d_InvaderToy, LM_GI_Right_l104d_Outlanes, LM_GI_Right_l104d_PF, LM_GI_Right_l104d_Parts, LM_GI_Right_l104d_Rails, LM_GI_Right_l104d_RocketToy, LM_GI_Right_l104d_SLING1, LM_GI_Right_l104d_sw11, LM_GI_Right_l104d_sw12, LM_Inserts_l11_B_Caps_Top_Yello, LM_Inserts_l11_FlipperLU, LM_Inserts_l11_FlipperRU, LM_Inserts_l11_LSling1, LM_Inserts_l11_LSling2, LM_Inserts_l11_Parts, LM_Inserts_l11_RSling1, LM_Inserts_l12_B_Caps_Top_Yello, LM_Inserts_l12_FlipperLU, _
  LM_Inserts_l12_LSling1, LM_Inserts_l12_LSling2, LM_Inserts_l12_Parts, LM_Inserts_l12_SLING2, LM_Inserts_l13_B_Caps_Top_Amber, LM_Inserts_l13_B_Caps_Top_Yello, LM_Inserts_l13_LSling1, LM_Inserts_l13_LSling2, LM_Inserts_l13_Parts, LM_Inserts_l13_sw48, LM_Inserts_l14_B_Caps_Top_Red, LM_Inserts_l14_B_Caps_Top_Yello, LM_Inserts_l14_LSling2, LM_Inserts_l14_Parts, LM_Inserts_l14_sw48, LM_Inserts_l14_sw48m, LM_Inserts_l14_sw77, LM_Inserts_l15_B_Caps_Top_Yello, LM_Inserts_l15_PF, LM_Inserts_l15_PF_Upper, LM_Inserts_l15_Parts, LM_Inserts_l15_sw48, LM_Inserts_l15_sw48m, LM_Inserts_l15_sw77, LM_Inserts_l16_B_Caps_Top_Amber, LM_Inserts_l16_B_Caps_Top_Red, LM_Inserts_l16_B_Caps_Top_Yello, LM_Inserts_l16_PF, LM_Inserts_l16_PF_Upper, LM_Inserts_l16_Parts, LM_Inserts_l16_sw77, LM_Inserts_l16_sw77m, LM_Inserts_l16_sw78, LM_Inserts_l17_B_Caps_Bot_Amber, LM_Inserts_l17_B_Caps_Top_Amber, LM_Inserts_l17_B_Caps_Top_Red, LM_Inserts_l17_B_Caps_Top_Yello, LM_Inserts_l17_BR2, LM_Inserts_l17_BS2, LM_Inserts_l17_PF_Upper, _
  LM_Inserts_l17_Parts, LM_Inserts_l17_sw77, LM_Inserts_l17_sw77m, LM_Inserts_l17_sw78, LM_Inserts_l17_sw78m, LM_Inserts_l18_B_Caps_Top_Amber, LM_Inserts_l18_PF_Upper, LM_Inserts_l18_Parts, LM_Inserts_l18_sw77, LM_Inserts_l18_sw78, LM_Inserts_l21_B_Caps_Top_Yello, LM_Inserts_l21_Parts, LM_Inserts_l22_B_Caps_Top_Yello, LM_Inserts_l22_FlipperRU, LM_Inserts_l22_Parts, LM_Inserts_l22_RSling1, LM_Inserts_l22_RSling2, LM_Inserts_l23_B_Caps_Top_Yello, LM_Inserts_l23_Parts, LM_Inserts_l23_RSling1, LM_Inserts_l23_RSling2, LM_Inserts_l23_SLING1, LM_Inserts_l24_B_Caps_Top_Amber, LM_Inserts_l24_B_Caps_Top_Yello, LM_Inserts_l24_Parts, LM_Inserts_l24_RSling2, LM_Inserts_l25_B_Caps_Top_Amber, LM_Inserts_l25_B_Caps_Top_Yello, LM_Inserts_l25_FlipperR1, LM_Inserts_l25_FlipperR1U, LM_Inserts_l25_PF, LM_Inserts_l25_Parts, LM_Inserts_l26_B_Caps_Top_Amber, LM_Inserts_l26_B_Caps_Top_Red, LM_Inserts_l26_FlipperR1U, LM_Inserts_l26_FlipperSpR1U, LM_Inserts_l26_Parts, LM_Inserts_l27_B_Caps_Top_Amber, LM_Inserts_l27_FlipperR1U, _
  LM_Inserts_l27_FlipperSpR1U, LM_Inserts_l27_PF, LM_Inserts_l27_Parts, LM_Inserts_l28_B_Caps_Top_Amber, LM_Inserts_l28_B_Caps_Top_Yello, LM_Inserts_l28_Parts, LM_Inserts_l28_SlotMachineToy, LM_Inserts_l31_B_Caps_Top_Red, LM_Inserts_l31_B_Caps_Top_Yello, LM_Inserts_l31_Outlanes, LM_Inserts_l31_Parts, LM_Inserts_l31_SideMod, LM_Inserts_l31_sw36, LM_Inserts_l32_B_Caps_Top_Amber, LM_Inserts_l32_B_Caps_Top_Yello, LM_Inserts_l32_Parts, LM_Inserts_l33_B_Caps_Top_Yello, LM_Inserts_l33_MysticSeerToy, LM_Inserts_l33_Outlanes, LM_Inserts_l33_PF, LM_Inserts_l33_Parts, LM_Inserts_l33_sw48, LM_Inserts_l34_B_Caps_Top_Red, LM_Inserts_l34_B_Caps_Top_Yello, LM_Inserts_l34_PF, LM_Inserts_l34_Parts, LM_Inserts_l34_sw48, LM_Inserts_l34_sw77, LM_Inserts_l35_B_Caps_Top_Yello, LM_Inserts_l35_Outlanes, LM_Inserts_l35_PF, LM_Inserts_l35_Parts, LM_Inserts_l35_sw48, LM_Inserts_l35_sw48m, LM_Inserts_l36_B_Caps_Top_Red, LM_Inserts_l36_B_Caps_Top_Yello, LM_Inserts_l36_Parts, LM_Inserts_l36_sw77, LM_Inserts_l36_sw78, _
  LM_Inserts_l37_B_Caps_Top_Yello, LM_Inserts_l37_BR3, LM_Inserts_l37_BS3, LM_Inserts_l37_Outlanes, LM_Inserts_l37_PF, LM_Inserts_l37_Parts, LM_Inserts_l37_sw48, LM_Inserts_l37_sw48m, LM_Inserts_l37_sw77, LM_Inserts_l38_B_Caps_Top_Amber, LM_Inserts_l38_BR3, LM_Inserts_l38_BS3, LM_Inserts_l38_BumperPegs, LM_Inserts_l38_PF_Upper, LM_Inserts_l38_Parts, LM_Inserts_l38_TownSquarePost, LM_Inserts_l38_sw48m, LM_Inserts_l38_sw77, LM_Inserts_l38_sw77m, LM_Inserts_l41_FlipperL, LM_Inserts_l41_FlipperLU, LM_Inserts_l41_FlipperSpL, LM_Inserts_l41_FlipperSpLU, LM_Inserts_l41_LSling1, LM_Inserts_l41_LSling2, LM_Inserts_l41_PF, LM_Inserts_l41_Parts, LM_Inserts_l41_SLING2, LM_Inserts_l42_B_Caps_Top_Yello, LM_Inserts_l42_FlipperL, LM_Inserts_l42_FlipperLU, LM_Inserts_l42_FlipperR, LM_Inserts_l42_FlipperRU, LM_Inserts_l42_FlipperSpL, LM_Inserts_l42_FlipperSpLU, LM_Inserts_l42_FlipperSpRU, LM_Inserts_l42_LSling1, LM_Inserts_l42_LSling2, LM_Inserts_l42_PF, LM_Inserts_l42_Parts, LM_Inserts_l43_B_Caps_Top_Yello, _
  LM_Inserts_l43_FlipperL, LM_Inserts_l43_FlipperLU, LM_Inserts_l43_FlipperR, LM_Inserts_l43_FlipperRU, LM_Inserts_l43_FlipperSpL, LM_Inserts_l43_FlipperSpLU, LM_Inserts_l43_FlipperSpRU, LM_Inserts_l43_LSling1, LM_Inserts_l43_LSling2, LM_Inserts_l43_PF, LM_Inserts_l43_Parts, LM_Inserts_l43_RSling1, LM_Inserts_l44_B_Caps_Top_Yello, LM_Inserts_l44_FlipperL, LM_Inserts_l44_FlipperLU, LM_Inserts_l44_FlipperR, LM_Inserts_l44_FlipperRU, LM_Inserts_l44_FlipperSpLU, LM_Inserts_l44_FlipperSpR, LM_Inserts_l44_FlipperSpRU, LM_Inserts_l44_LSling1, LM_Inserts_l44_PF, LM_Inserts_l44_Parts, LM_Inserts_l44_RSling1, LM_Inserts_l44_RSling2, LM_Inserts_l45_B_Caps_Top_Yello, LM_Inserts_l45_FlipperL, LM_Inserts_l45_FlipperLU, LM_Inserts_l45_FlipperR, LM_Inserts_l45_FlipperRU, LM_Inserts_l45_FlipperSpLU, LM_Inserts_l45_FlipperSpR, LM_Inserts_l45_FlipperSpRU, LM_Inserts_l45_PF, LM_Inserts_l45_Parts, LM_Inserts_l45_RSling1, LM_Inserts_l45_RSling2, LM_Inserts_l45_SLING1, LM_Inserts_l46_FlipperR, LM_Inserts_l46_FlipperRU, _
  LM_Inserts_l46_FlipperSpR, LM_Inserts_l46_FlipperSpRU, LM_Inserts_l46_PF, LM_Inserts_l46_Parts, LM_Inserts_l47_FlipperL, LM_Inserts_l47_FlipperLU, LM_Inserts_l47_FlipperR, LM_Inserts_l47_FlipperRU, LM_Inserts_l47_FlipperSpL, LM_Inserts_l47_FlipperSpLU, LM_Inserts_l47_FlipperSpR, LM_Inserts_l47_FlipperSpRU, LM_Inserts_l47_PF, LM_Inserts_l47_Parts, LM_Inserts_l48_FlipperR1U, LM_Inserts_l48_Outlanes, LM_Inserts_l48_PF, LM_Inserts_l48_Parts, LM_Inserts_l48_RSling2, LM_Inserts_l48_RocketToy, LM_Inserts_l51_B_Caps_Top_Amber, LM_Inserts_l51_Clock_Color, LM_Inserts_l51_Clock_White, LM_Inserts_l51_ClockToy, LM_Inserts_l51_Parts, LM_Inserts_l51_Robot, LM_Inserts_l51_Sign_Spiral, LM_Inserts_l51_sw47m, LM_Inserts_l51_sw68, LM_Inserts_l52_B_Caps_Bot_Amber, LM_Inserts_l52_B_Caps_Top_Amber, LM_Inserts_l52_Clock_Color, LM_Inserts_l52_Clock_White, LM_Inserts_l52_ClockToy, LM_Inserts_l52_FlipperL1, LM_Inserts_l52_PF_Upper, LM_Inserts_l52_Parts, LM_Inserts_l52_Sign_Spiral, LM_Inserts_l52_sw68, LM_Inserts_l53_B_Caps_Bot_Amber, _
  LM_Inserts_l53_B_Caps_Top_Amber, LM_Inserts_l53_Clock_Color, LM_Inserts_l53_Clock_White, LM_Inserts_l53_ClockToy, LM_Inserts_l53_FlipperL1, LM_Inserts_l53_Over3, LM_Inserts_l53_PF_Upper, LM_Inserts_l53_Parts, LM_Inserts_l53_Sign_Spiral, LM_Inserts_l54_Clock_White, LM_Inserts_l54_ClockToy, LM_Inserts_l54_PF, LM_Inserts_l54_PF_Upper, LM_Inserts_l54_Parts, LM_Inserts_l54_Sign_Spiral, LM_Inserts_l55_B_Caps_Top_Amber, LM_Inserts_l55_Clock_Color, LM_Inserts_l55_Clock_White, LM_Inserts_l55_ClockToy, LM_Inserts_l55_FlipperL1, LM_Inserts_l55_FlipperSpL1, LM_Inserts_l55_PF_Upper, LM_Inserts_l55_Parts, LM_Inserts_l55_Sign_Spiral, LM_Inserts_l56_B_Caps_Bot_Amber, LM_Inserts_l56_Clock_Color, LM_Inserts_l56_Clock_White, LM_Inserts_l56_ClockToy, LM_Inserts_l56_Over3, LM_Inserts_l56_PF, LM_Inserts_l56_Parts, LM_Inserts_l56_Piano, LM_Inserts_l56_Sign_Spiral, LM_Inserts_l56_sw64, LM_Inserts_l56_sw68, LM_Inserts_l57_Over2, LM_Inserts_l57_Over3, LM_Inserts_l57_PF, LM_Inserts_l57_Parts, LM_Inserts_l57_sw64, _
  LM_Inserts_l58_Clock_Color, LM_Inserts_l58_Clock_White, LM_Inserts_l58_Over2, LM_Inserts_l58_Parts, LM_Inserts_l64_B_Caps_Bot_Red, LM_Inserts_l64_B_Caps_Top_Amber, LM_Inserts_l64_B_Caps_Top_Yello, LM_Inserts_l64_BR2, LM_Inserts_l64_BS2, LM_Inserts_l64_BumperPegs, LM_Inserts_l64_PF, LM_Inserts_l64_PF_Upper, LM_Inserts_l64_Parts, LM_Inserts_l64_TownSquarePost, LM_Inserts_l64_sw77, LM_Inserts_l64_sw77m, LM_Inserts_l64_sw78m, LM_Inserts_l65_B_Caps_Top_Amber, LM_Inserts_l65_B_Caps_Top_Red, LM_Inserts_l65_B_Caps_Top_Yello, LM_Inserts_l65_BR2, LM_Inserts_l65_BS2, LM_Inserts_l65_PF, LM_Inserts_l65_PF_Upper, LM_Inserts_l65_Parts, LM_Inserts_l65_sw77, LM_Inserts_l65_sw77m, LM_Inserts_l65_sw78, LM_Inserts_l65_sw78m, LM_Inserts_l66_Outlanes, LM_Inserts_l66_Parts, LM_Inserts_l66_RocketToy, LM_Inserts_l66_sw11, LM_Inserts_l67_Clock_Color, LM_Inserts_l67_Clock_White, LM_Inserts_l67_Over1, LM_Inserts_l67_Parts, LM_Inserts_l67_SlotMachineToy, LM_Inserts_l68_Clock_Color, LM_Inserts_l68_Clock_White, LM_Inserts_l68_Over1, _
  LM_Inserts_l68_Over3, LM_Inserts_l68_Parts, LM_Inserts_l71_B_Caps_Top_Amber, LM_Inserts_l71_B_Caps_Top_Yello, LM_Inserts_l71_FlipperR1U, LM_Inserts_l71_FlipperSpR1U, LM_Inserts_l71_PF, LM_Inserts_l71_Parts, LM_Inserts_l71_SlotMachineToy, LM_Inserts_l71_sw47, LM_Inserts_l71_sw66, LM_Inserts_l71_sw67, LM_Inserts_l71_sw67m, LM_Inserts_l71_sw68m, LM_Inserts_l72_B_Caps_Top_Amber, LM_Inserts_l72_Clock_Color, LM_Inserts_l72_Clock_White, LM_Inserts_l72_ClockToy, LM_Inserts_l72_PF, LM_Inserts_l72_Parts, LM_Inserts_l72_Sign_Spiral, LM_Inserts_l72_sw66, LM_Inserts_l72_sw67, LM_Inserts_l72_sw68m, LM_Inserts_l73_B_Caps_Bot_Amber, LM_Inserts_l73_B_Caps_Top_Amber, LM_Inserts_l73_Clock_Color, LM_Inserts_l73_Clock_White, LM_Inserts_l73_ClockToy, LM_Inserts_l73_PF, LM_Inserts_l73_Parts, LM_Inserts_l73_Sign_Spiral, LM_Inserts_l73_sw66, LM_Inserts_l73_sw66m, LM_Inserts_l73_sw67, LM_Inserts_l73_sw67m, LM_Inserts_l75_B_Caps_Bot_Amber, LM_Inserts_l75_B_Caps_Top_Amber, LM_Inserts_l75_Clock_Color, LM_Inserts_l75_Clock_White, _
  LM_Inserts_l75_Over3, LM_Inserts_l75_PF, LM_Inserts_l75_Parts, LM_Inserts_l75_Piano, LM_Inserts_l75_sw64, LM_Inserts_l75_sw64m, LM_Inserts_l76_PF_Upper, LM_Inserts_l76_Parts, LM_Inserts_l76_Pyramid, LM_Inserts_l77_PF_Upper, LM_Inserts_l77_Parts, LM_Inserts_l78_PF_Upper, LM_Inserts_l78_Parts, LM_Inserts_l78_Pyramid, LM_Inserts_l81_B_Caps_Bot_Amber, LM_Inserts_l81_ClockToy, LM_Inserts_l81_FlipperL1, LM_Inserts_l81_FlipperSpL1, LM_Inserts_l81_Over2, LM_Inserts_l81_PF, LM_Inserts_l81_PF_Upper, LM_Inserts_l81_Parts, LM_Inserts_l81_Sign_Spiral, LM_Inserts_l82_B_Caps_Top_Amber, LM_Inserts_l82_Clock_Color, LM_Inserts_l82_Clock_White, LM_Inserts_l82_ClockToy, LM_Inserts_l82_Over3, LM_Inserts_l82_PF, LM_Inserts_l82_Parts, LM_Inserts_l82_Piano, LM_Inserts_l82_Robot, LM_Inserts_l82_Sign_Spiral, LM_Inserts_l82_sw47m, LM_Inserts_l82_sw64, LM_Inserts_l82_sw68, LM_Inserts_l83_Clock_Color, LM_Inserts_l83_Clock_White, LM_Inserts_l83_ClockToy, LM_Inserts_l83_Over1, LM_Inserts_l83_Over2, LM_Inserts_l83_Over3, LM_Inserts_l83_PF, _
  LM_Inserts_l83_Parts, LM_Inserts_l83_Piano, LM_Inserts_l83_sw64, LM_Inserts_l83_sw64m, LM_Inserts_l84_Clock_Color, LM_Inserts_l84_Clock_White, LM_Inserts_l84_ClockLarge, LM_Inserts_l84_ClockShort, LM_Inserts_l84_ClockToy, LM_Inserts_l84_Over1, LM_Inserts_l84_Over2, LM_Inserts_l84_Over3, LM_Inserts_l84_PF, LM_Inserts_l84_Parts, LM_Inserts_l84_Piano, LM_Inserts_l84_sw64, LM_Inserts_l84_sw64m, LM_Inserts_l85_Clock_Color, LM_Inserts_l85_Clock_White, LM_Inserts_l85_FlipperR1U, LM_Inserts_l85_Over2, LM_Inserts_l85_Over3, LM_Inserts_l85_PF, LM_Inserts_l85_Parts, LM_Inserts_l85_RocketToy, LM_Inserts_l85_SlotMachineToy, LM_Inserts_l85_sw65m, LM_Inserts_l85_sw67m, LM_Inserts_l86_FlipperR1, LM_Inserts_l86_FlipperR1U, LM_Inserts_l86_FlipperSpR1, LM_Inserts_l86_FlipperSpR1U, LM_Inserts_l86_InvaderToy, LM_Inserts_l86_Over1, LM_Inserts_l86_Over3, LM_Inserts_l86_PF, LM_Inserts_l86_Parts, LM_Inserts_l86_SlotMachineToy, LM_Mod_l105_InvaderToy, LM_Mod_l105_Over1, LM_Mod_l105_Parts, LM_Mod_l106_FlipperR1, LM_Mod_l106_FlipperR1U, _
  LM_Mod_l106_FlipperSpR1, LM_Mod_l106_FlipperSpR1U, LM_Mod_l106_InvaderToy, LM_Mod_l106_Over1, LM_Mod_l106_PF, LM_Mod_l106_Parts, LM_Mod_l106_SlotMachineToy, LM_Mod_l107_FlipperR1, LM_Mod_l107_FlipperR1U, LM_Mod_l107_FlipperSpR1, LM_Mod_l107_FlipperSpR1U, LM_Mod_l107_InvaderToy, LM_Mod_l107_PF, LM_Mod_l107_Parts, LM_Mod_l108_PF, LM_Mod_l108_Parts, LM_Mod_l108_sw68m, LM_Mod_l109_B_Caps_Bot_Amber, LM_Mod_l109_B_Caps_Bot_Red, LM_Mod_l109_B_Caps_Bot_Yellow, LM_Mod_l109_B_Caps_Top_Amber, LM_Mod_l109_B_Caps_Top_Yellow, LM_Mod_l109_BR1, LM_Mod_l109_BR2, LM_Mod_l109_BR3, LM_Mod_l109_BS2, LM_Mod_l109_BumperPegs, LM_Mod_l109_PF, LM_Mod_l109_PF_Upper, LM_Mod_l109_Parts, LM_Mod_l109_TownSquarePost, LM_Mod_l109_sw48m, LM_Mod_l109_sw77, LM_Mod_l109_sw77m, LM_Mod_l109_sw78, LM_Mod_l110_B_Caps_Bot_Amber, LM_Mod_l110_BR2, LM_Mod_l110_BS2, LM_Mod_l110_Camera, LM_Mod_l110_PF, LM_Mod_l110_PF_Upper, LM_Mod_l110_Parts, LM_Mod_l111_BR1, LM_Mod_l111_FlipperR1, LM_Mod_l111_FlipperR1U, LM_Mod_l111_FlipperSpR1, LM_Mod_l111_FlipperSpR1U, _
  LM_Mod_l111_Outlanes, LM_Mod_l111_PF, LM_Mod_l111_Parts, LM_Mod_l111_RSling1, LM_Mod_l111_RSling2, LM_Mod_l111_RocketToy, LM_Mod_l111_SLING1)
' VLM  Arrays - End



'**********************************
'   General Math Functions
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

Function RndNumO(min, max)
    RndNumO = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function


'*******************************************
'  Options & Mods
'*******************************************

Dim TVMod, SpiralMod, StagedFlipperMod, CameraMod, FlipMod
Dim LightLevel : LightLevel = -1
Dim InsertLevel : InsertLevel = -1
Dim VolumeDial : VolumeDial = 0.8           'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5   'Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5   'Level of ramp rolling volume. Value between 0 and 1
Dim VRRoomChoice : VRRoomChoice = 2     '1 - Minimal Room,  2 - Mega Room,  3 - Ultra Minimal
Dim OutlaneDifficulty : OutlaneDifficulty = 1 'Easy - 0, Medium - 1 (default), Hard - 2

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y, v

    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

    ' Room Brightness
  v = (NightDay^2.2)/(100.0^2.2)
  SetRoomBrightness v

    ' Insert Power
  Dim inserts: inserts = Array(BL_Inserts_l11, BL_Inserts_l12, BL_Inserts_l13, BL_Inserts_l14, BL_Inserts_l15, BL_Inserts_l16, BL_Inserts_l17, BL_Inserts_l18, _
    BL_Inserts_l21, BL_Inserts_l22, BL_Inserts_l23, BL_Inserts_l24, BL_Inserts_l25, BL_Inserts_l26, BL_Inserts_l27, BL_Inserts_l28, BL_Inserts_l31, BL_Inserts_l32, _
    BL_Inserts_l33, BL_Inserts_l34, BL_Inserts_l35, BL_Inserts_l36, BL_Inserts_l37, BL_Inserts_l38, BL_Inserts_l41, BL_Inserts_l42, BL_Inserts_l43, BL_Inserts_l44, _
    BL_Inserts_l45, BL_Inserts_l46, BL_Inserts_l47, BL_Inserts_l48, BL_Inserts_l51, BL_Inserts_l52, BL_Inserts_l53, BL_Inserts_l54, BL_Inserts_l55, BL_Inserts_l56, _
    BL_Inserts_l57, BL_Inserts_l58, BL_Inserts_l64, BL_Inserts_l65, BL_Inserts_l66, BL_Inserts_l67, BL_Inserts_l68, BL_Inserts_l71, BL_Inserts_l72, BL_Inserts_l73, _
    BL_Inserts_l75, BL_Inserts_l76, BL_Inserts_l77, BL_Inserts_l78, BL_Inserts_l81, BL_Inserts_l82, BL_Inserts_l83, BL_Inserts_l84, BL_Inserts_l85, BL_Inserts_l86)
    v = Int(255 * (Table1.Option("Insert Power", 0.2, 1, 0.01, 0.5, 1) ^ 2))
    For Each y in inserts: For Each x in y: x.Color = RGB(v, v, v): Next: Next

  'Difficulty
  OutlaneDifficulty = Table1.Option("Outlane Difficulty", 0, 2, 1, 1, 0, Array("Easy", "Medium", "Hard"))
  SetOLDifficulty OutlaneDifficulty

    ' Show Rails
    v = Table1.Option("Show Rails", 0, 1, 1, 1, 0, Array("Hide", "Show"))
    For Each x in BP_Rails: x.Visible = v: Next

    ' Bumper Post Mod
    v = Table1.Option("Mod: Bumper Post", 0, 1, 1, 0, 0, Array("Off", "On"))
    BumperMod1.collidable = v
  BumperMod2.IsDropped = 1 - v
    BumperMod3.collidable = v
    For Each x in BP_BumperPegs: x.Visible = v: Next

    ' Camera Mod
    CameraMod = Table1.Option("Mod: Camera", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_Camera: x.Visible = CameraMod: Next

    ' Black & White Clock Mod
    v = Table1.Option("Mod: B&W Clock", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_Clock_Color: x.Visible = 1 - v: Next
    For Each x in BP_Clock_White: x.Visible = v: Next

    ' Mini Clock Mod
    v = Table1.Option("Mod: Mini Clock", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_ClockToy: x.Visible = v: Next

    ' Flipper Mod
    v = Table1.Option("Mod: Flippers", 0, 1, 1, 0, 0, Array("Off", "On"))
  FlipMod = v
    For Each x in BP_FlipperL: x.Visible = 1 - v: Next
    For Each x in BP_FlipperL1: x.Visible = 1 - v: Next
    For Each x in BP_FlipperR: x.Visible = 1 - v: Next
    For Each x in BP_FlipperR1: x.Visible = 1 - v: Next
    For Each x in BP_FlipperLU: x.Visible = 1 - v: Next
    For Each x in BP_FlipperRU: x.Visible = 1 - v: Next
    For Each x in BP_FlipperR1U: x.Visible = 1 - v: Next
    For Each x in BP_FlipperSpL: x.Visible = v: Next
    For Each x in BP_FlipperSpL1: x.Visible = v: Next
    For Each x in BP_FlipperSpR: x.Visible = v: Next
    For Each x in BP_FlipperSpR1: x.Visible = v: Next
    For Each x in BP_FlipperSpLU: x.Visible = v: Next
    For Each x in BP_FlipperSpRU: x.Visible = v: Next
    For Each x in BP_FlipperSpR1U: x.Visible = v: Next
  LeftFlipper_Animate
  LeftFlipper1_Animate
  RightFlipper_Animate
  RightFlipper1_Animate

    ' Gumball Mod
    v = Table1.Option("Mod: Gumball", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_Gumballs: x.Visible = v: Next

    ' Invader Mod
    v = Table1.Option("Mod: Invader", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_InvaderToy: x.Visible = v: Next
  InvaderTimer.Enabled = v
  l105.State = v
    l106.State = v
    l107.STate = v

    ' Mystic Seer Mod
    v = Table1.Option("Mod: Mystic Seer", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_MysticSeerToy: x.Visible = v: Next

    ' Piano Mod
    v = Table1.Option("Mod: Piano", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_Piano: x.Visible = v: Next

    ' Pyramid Mod
    v = Table1.Option("Mod: Pyramid", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_Pyramid: x.Visible = v: Next

    ' Robot Mod
    v = Table1.Option("Mod: Robot", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_Robot: x.Visible = v: Next

    ' Rocket Mod
    v = Table1.Option("Mod: Rocket", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_RocketToy: x.Visible = v: Next

    ' Scoop light Mod
    v = Table1.Option("Mod: Scoop light", 0, 1, 1, 0, 0, Array("Off", "On"))
    l108.State = v

    ' Sidewall Mod
    v = Table1.Option("Mod: Sidewalls", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_SideMod: x.Visible = v: Next

    ' Spiral Mod
    SpiralMod = Table1.Option("Mod: Spiral", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_SpiralToy: x.Visible = SpiralMod: Next

    ' Spiral Sign Mod
    v = Table1.Option("Mod: Spiral Sign", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_Sign_Spiral: x.Visible = v: Next

    ' Slot Machine Mod
    v = Table1.Option("Mod: Slot Machine", 0, 1, 1, 0, 0, Array("Off", "On"))
    swSlotReel.enabled = v
    SlotReel.visible = v
    For Each x in BP_SlotMachineToy: x.Visible = v: Next

    ' Target Mod
    v = Table1.Option("Mod: Targets", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_sw47: x.Visible = 1 - v: Next
    For Each x in BP_sw48: x.Visible = 1 - v: Next
    For Each x in BP_sw64: x.Visible = 1 - v: Next
    For Each x in BP_sw65: x.Visible = 1 - v: Next
    For Each x in BP_sw65a: x.Visible = 1 - v: Next
    For Each x in BP_sw66: x.Visible = 1 - v: Next
    For Each x in BP_sw67: x.Visible = 1 - v: Next
    For Each x in BP_sw68: x.Visible = 1 - v: Next
    For Each x in BP_sw77: x.Visible = 1 - v: Next
    For Each x in BP_sw78: x.Visible = 1 - v: Next
    For Each x in BP_sw47m: x.Visible = v: Next
    For Each x in BP_sw48m: x.Visible = v: Next
    For Each x in BP_sw64m: x.Visible = v: Next
    For Each x in BP_sw65m: x.Visible = v: Next
    For Each x in BP_sw65am: x.Visible = v: Next
    For Each x in BP_sw66m: x.Visible = v: Next
    For Each x in BP_sw67m: x.Visible = v: Next
    For Each x in BP_sw68m: x.Visible = v: Next
    For Each x in BP_sw77m: x.Visible = v: Next
    For Each x in BP_sw78m: x.Visible = v: Next

    ' Town Square Post Mod
    v = Table1.Option("Mod: Town Square Post", 0, 1, 1, 0, 0, Array("Off", "On"))
    For Each x in BP_TownSquarePost: x.Visible = v: Next
  l109.State = v

    ' TV Mod
    TVMod = Table1.Option("Mod: TV", 0, 1, 1, 0, 0, Array("Off", "On"))
    Frame.visible = TVMod
    For Each x in BP_TVtoy: x.Visible = TVMod: Next

    ' Extra Magnet Mod
    v = Table1.Option("Mod: Extra Magnet", 0, 1, 1, 0, 0, Array("Off", "On"))
    sw82.enabled = v
    sw82_help.enabled = v
    UpperRightMagnet.enabled = v
    For Each x in BP_URMagnet: x.Visible = v: Next

    ' Staged Flipper
    StagedFlipperMod = Table1.Option("Staged Flipper", 0, 1, 1, 0, 0, Array("Off", "On"))
    If StagedFlipperMod = 1 Then
        keyStagedFlipperL = KeyUpperLeft
        keyStagedFlipperR = KeyUpperRight
    Else
        keyStagedFlipperL = LeftFlipperKey
        keyStagedFlipperR = RightFlipperKey
    End If

    ' Refractions
  v = Table1.Option("Refraction Setting", 0, 2, 1, 2, 0, Array("Min Refraction (best performance)", "Sharp Refractions (improved performance)", "Rough Refractions (best visuals)"))
  SetRefractionProbes v

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    RampRollVolume = Table1.Option("Ramp Volume", 0, 1, 0.01, 0.5, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

  VRRoomChoice = Table1.Option("VR Room", 0, 3, 1, 2, 0, Array("Off", "Minimal Room", "360", "Ultra Minimal"))
  LoadVRRoom

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Sub swSlotReel_Hit:SlotReelTimer.enabled = 1:End Sub

Dim SlotPic : SlotPic = 0
Sub SlotReelTimer_Timer()
    SlotPic = SlotPic + 1
    if SlotPic > 10 Then SlotReelTimer.enabled = 0:ResetSlot.enabled = 1
    SlotReel.imageA = "slot_" & SlotPic
End Sub

Sub ResetSlot_Timer()
    SlotReel.imageA = "slot_10":vpmtimer.addtimer 150, "SlotReel.imageA = ""slot_0"" '":SlotPic = 0
    If SlotPic = 0 then ResetSlot.enabled = 0
End Sub

Dim TVPic : TVPic = 0
Sub TVTimer_Timer()
    TVPic = TVPic + 1
    if TVPic = 34 Then TVPic = 2
    Frame.imageA = "tv_" & TVPic
End Sub

Sub SpiralMove_Timer()
    dim a, x : a = BM_SpiralToy.RotZ + 10
    For Each x in BP_SpiralToy: x.RotZ = a: Next
End Sub

Dim InvLR, InvC : InvLR = 0 : InvC = 0
Sub InvaderTimer_Timer
    InvLR = InvLR + 1
    InvC = InvC + 1
    If InvLR >= 4 Then InvLR = 0
    If InvC >= 10 Then InvC = 0
    l105.State = InvC < 5
    l106.State = InvLR < 2
    l107.STate = InvLR >= 2
End SUb

Sub SetOLDifficulty(Opt)
'BM_Outlanes
  Dim BP, lvl

  zCol_Rubber_Post_L_Easy.collidable = 0
  zCol_Rubber_Post_R_Easy.collidable = 0

  zCol_Rubber_Post_L_Medium.collidable = 0
  zCol_Rubber_Post_R_Medium.collidable = 0

  zCol_Rubber_Post_L_Hard.collidable = 0
  zCol_Rubber_Post_R_Hard.collidable = 0


  Select Case Opt
    Case 0:
      lvl = 1432
      zCol_Rubber_Post_L_Easy.collidable = 1
      zCol_Rubber_Post_R_Easy.collidable = 1
    Case 1:
      lvl = 1418.5
      zCol_Rubber_Post_L_Medium.collidable = 1
      zCol_Rubber_Post_R_Medium.collidable = 1
    Case 2:
      lvl = 1405
      zCol_Rubber_Post_L_Hard.collidable = 1
      zCol_Rubber_Post_R_Hard.collidable = 1
  End Select

  For Each BP in BP_Outlanes
    BP.Y = lvl
  Next
End Sub

Sub SetRefractionProbes(Opt)
  On Error Resume Next
    Select Case Opt
      Case 0:
        BM_B_Caps_Bot_Amber.RefractionProbe = "Refraction 1 Sharp"
        BM_B_Caps_Bot_Red.RefractionProbe = "Refraction 1 Sharp"
        BM_B_Caps_Bot_Yellow.RefractionProbe = "Refraction 1 Sharp"
        BM_Over3.RefractionProbe = ""
'       BM_B_Caps_Top_Amber.RefractionProbe = "Refraction 2"
'       BM_B_Caps_Top_Red.RefractionProbe = "Refraction 2"
'       BM_B_Caps_Top_Yellow.RefractionProbe = "Refraction 2"
      Case 1:
        BM_B_Caps_Bot_Amber.RefractionProbe = "Refraction 1 Sharp"
        BM_B_Caps_Bot_Red.RefractionProbe = "Refraction 1 Sharp"
        BM_B_Caps_Bot_Yellow.RefractionProbe = "Refraction 1 Sharp"
        BM_Over3.RefractionProbe = "Refraction 1 Sharp"
'       BM_B_Caps_Top_Amber.RefractionProbe = "Refraction 2"
'       BM_B_Caps_Top_Red.RefractionProbe = "Refraction 2"
'       BM_B_Caps_Top_Yellow.RefractionProbe = "Refraction 2"
      Case 2:
        BM_B_Caps_Bot_Amber.RefractionProbe = "Refraction 1"
        BM_B_Caps_Bot_Red.RefractionProbe = "Refraction 1"
        BM_B_Caps_Bot_Yellow.RefractionProbe = "Refraction 1"
        BM_Over3.RefractionProbe = "Refraction 1"
'       BM_B_Caps_Top_Amber.RefractionProbe = "Refraction 2"
'       BM_B_Caps_Top_Red.RefractionProbe = "Refraction 2"
'       BM_B_Caps_Top_Yellow.RefractionProbe = "Refraction 2"
    End Select
  On Error Goto 0
End Sub


'****************************
'   ZRBR: Room Brightness
'****************************

'This code only applies to lightmapped tables. It is here for reference.
'NOTE: Objects bightness will be affected by the Day/Night slider only if their blenddisablelighting property is less than 1.
'      Lightmapped table primitives have their blenddisablelighting equal to 1, therefore we need this SetRoomBrightness sub
'      to handle updating their effective ambient brighness.

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","Playfield")

Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0

  ' Lighting level
  Dim v: v=(lvl * 245 + 10)/255

  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    ModulateMaterialBaseColor RoomBrightnessMtlArray(i), i, v
  Next
End Sub

Dim SavedMtlColorArray
SaveMtlColors
Sub SaveMtlColors
  ReDim SavedMtlColorArray(UBound(RoomBrightnessMtlArray))
  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    SaveMaterialBaseColor RoomBrightnessMtlArray(i), i
  Next
End Sub

Sub SaveMaterialBaseColor(name, idx)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  SavedMtlColorArray(idx) = round(base,0)
End Sub


Sub ModulateMaterialBaseColor(name, idx, val)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  Dim red, green, blue, saved_base, new_base

  'First get the existing material properties
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

  'Get saved color
  saved_base = SavedMtlColorArray(idx)

  'Next extract the r,g,b values from the base color
  red = saved_base And &HFF
  green = (saved_base \ &H100) And &HFF
  blue = (saved_base \ &H10000) And &HFF
  'msgbox red & " " & green & " " & blue

  'Create new color scaled down by 'val', and update the material
  new_base = RGB(red*val, green*val, blue*val)
  UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, new_base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub


'*******************************************
'  Timers
'*******************************************

Dim Bumpers : Bumpers = Array(Bumper1, Bumper2, Bumper3)
Dim FrameTime, LastFrameTime: LastFrameTime = 0

' The frame timer interval is -1, so executes at the display frame rate
Sub FrameTimer_Timer()
    FrameTime = GameTime - LastFrameTime
    LastFrameTime = GameTime

  ' Core tween updates
  UpdateAllTweens FrameTime

  ' Update rolling sounds
    RollingUpdate

  ' Handle stand up target animations
    DoSTAnim

  'debug.print plunger.position
  VR_Primary_plunger.Y = -8.430555 + (5* Plunger.Position)

    ' Ball Brightness
    Dim b, x, v
    For b = 0 to UBound(TZBalls)
        If CheckGumball(TZBalls(b)) Or CheckGumball2(TZBalls(b)) Or CheckGumball3(TZBalls(b)) Then
            ' Inside Gumball machine => only get the dimmed general light level
            v = Int(16 + 0.75 * (128 - 16) * (NightDay / 100))
      TZBalls(b).BulbIntensityScale = 0.2
    ElseIf TZBalls(b).x > 917 And TZBalls(b).y > 1250 Then
      ' Inside plunger lanes: No GI
            v = Int(16 + 0.75 * (128 - 16) * (NightDay / 100.0))
      TZBalls(b).BulbIntensityScale = 0.0
        Else
            ' On playfield => global light level, with a blend of GI depending on position
            x = TZBalls(b).x / 1096
            If x < 0 Then x = 0
            If x > 1 Then x = 1
            v = Int(16 + 0.75 * 128*0.01*((1-x)*LM_GI_Left_l100_PF.Opacity + x*LM_GI_Right_l104_PF.Opacity) + (128 - 16) * (NightDay / 100.0))
      TZBalls(b).BulbIntensityScale = 0.6
        End If
    If TZBalls(b).id = PowerBallID Then v = cint(v / 2)
        If v < 0 Then v = 0
        If v > 255 Then v = 255
        TZBalls(b).color = RGB(v, v, v)
    Next

    DynamicBSUpdate

    UpdateClock

    If BIP > 0 AND SpiralMod = 1 Then SpiralMove.enabled = l68.state > 0.5

    ' Animate Bumper switch
  If False Then
    ' 83 = Bumper skirt radius + 25 => start pressing
    ' 63 = Bumper radius + 25 => fully pressed
    Dim i, nearest, y, z, s
    For i = 0 To 2
      nearest = 10000.
      For s = 0 to UBound(TZBalls)
        If TZBalls(s).z < 30 Then ' Ball on playfield
          x = Bumpers(i).x - TZBalls(s).x
          y = Bumpers(i).y - TZBalls(s).y
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
      If i = 0 And BM_BS1.z <> z Then For Each x in BP_BS1: x.Z = z: Next
      If i = 1 And BM_BS2.z <> z Then For Each x in BP_BS2: x.Z = z: Next
      If i = 2 And BM_BS3.z <> z Then For Each x in BP_BS3: x.Z = z: Next
    Next
  End If
End Sub

'************************************************************************
'                        INIT TABLE
'************************************************************************

Dim TZBall1, TZBall2, TZBall3, TZBall4, TZBall5, TZBall6, TZBalls
Dim bsSlot, bsAutoPlunger, bsRocket, mLeftMini, mRightMini, mslot, mLeftMagnet, mLowerRightMagnet, mUpperRightMagnet

Sub Table1_Init
    vpmInit Me
  vpmMapLights AllLamps ' Map all lamps to the corresponding ROM output suing the value of TimerInterval of each light object
    With Controller
        .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Twilight Zone (Bally 1993)" & vbNewLine & "VPW"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 1' + 2 'Gumball + Clock  'just clock
    End With
  On Error Resume Next
'        .Run GetPlayerHWnd
  Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

    ' Init switches
    Controller.Switch(22) = 1 'close coin door

    '************   Nudging   **************************

    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 4
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Leftslingshot, Rightslingshot)

    '**************** Slot Machine Kickout ****************
    slotkick_vel = 50           'velocity
    slotkick_vel_variance = 0.5 'velocity variance
    slotkick_angle = Loopy_New.objrotz  'adjust objrotz on loop mesh to adjust kickout direction
    slotkick_angle_variance = 1 'Angle variance

    '****************   Magnets   ******************

    Set mLeftMini = New cvpmMagnet
    With mLeftMini
        .InitMagnet TLMiniFlip, 60      ' Left Powerfield Real Magnet Strength (adjust to taste) - 70
        .GrabCenter = False: .Size = 195
        .CreateEvents "mLeftMini"
    End With

    Set mRightMini = New cvpmMagnet
    With mRightMini
        .InitMagnet TRMiniFlip, 60      ' Right Powerfield Real Magnet Strength (adjust to taste) - 70
        .GrabCenter = False: .Size = 195
        .CreateEvents "mRightMini"
    End With

    Set mLeftMagnet = New cvpmMagnet
    With mLeftMagnet
        .InitMagnet LeftMagnet, 67
        .CreateEvents "mLeftMagnet"
        .GrabCenter = True
    End With

    Set mUpperRightMagnet = New cvpmMagnet
    With mUpperRightMagnet
        .InitMagnet UpperRightMagnet, 63
        .CreateEvents "mUpperRightMagnet"
        .GrabCenter = True
    End With

    Set mLowerRightMagnet = New cvpmMagnet
    With mLowerRightMagnet
        .InitMagnet LowerRightMagnet, 63
        .CreateEvents "mLowerRightMagnet"
        .GrabCenter = True
    End With

    '****************   Init GI   ******************
    dim x

    UpdateGI 0, 1:UpdateGI 1, 1:UpdateGI 2, 1:UpdateGI 3, 1:UpdateGI 4, 1

    'for ballsearch
     for each x in Array(sw84,sw85,sw88,sw58,sw18,sw25,sw16,sw15,sw17)
        x.UserValue = cInt(mid(x.name, 3, 2))
    Next

    GumballPopper.uservalue = 74

    '************  Trough   **************************
    Set TZBall6 = sw17.CreateSizedballWithMass(Ballsize/2,Ballmass)
    Set TZBall5 = sw16.CreateSizedballWithMass(Ballsize/2,Ballmass)
    Set TZBall4 = sw15.CreateSizedballWithMass(Ballsize/2,Ballmass)
    Set TZBall3 = FreezeKicker2.CreateSizedballWithMass(Ballsize/2,Ballmass)
    Set TZBall2 = FreezeKicker1.CreateSizedballWithMass(Ballsize/2,Ballmass)
    Set TZBall1 = FreezeKicker0.CreateSizedballWithMass(Ballsize/2,Ballmass)
    TZBalls = Array(TZBall1, TZBall2, TZBall3, TZBall4, TZBall5, TZBall6)

    Controller.Switch(17) = 1
    Controller.Switch(16) = 1
    Controller.Switch(15) = 1
    Controller.Switch(26) = 1

'   Freezekicker2.kick 0,0,0
'   Freezekicker1.kick 0,0,0
'   Freezekicker0.kick 0,0,0

    If PowerballStart = 0 then PowerballStart = RndNumO(1, 6) end if    ' 0 = completely random powerball start
    If PowerballStart = 7 then PowerballStart = RndNumO(1, 3) end if    ' 7 = Random start in the gumball machine

    If powerballstart = 1 Then SetPowerBall TZBall1
    If powerballstart = 2 Then SetPowerBall TZBall2
    If powerballstart = 3 Then SetPowerBall TZBall3
    If powerballstart = 4 Then SetPowerBall TZBall4:Controller.switch(26) = 0
    If powerballstart = 5 Then SetPowerBall TZBall5
    If powerballstart = 6 Then SetPowerBall TZBall6

    'SpawnBalls 'spawn all balls in gumball machine and in trough

    '************  FSS & VR Backglass **************************

    If RenderingMode = 2 or VRDesktopSim Then ' Enable VR stuff if run in VR mode

        For each x in VRbackglass
            x.visible = True
        Next
        Setbackglass
        BGTimer.enabled = true : BGTimer2.enabled = true : BGTimer3.enabled = true
    ElseIf Table1.ShowFSS Then ' Enable FSS
        Setbackglass
    Else
        for each x in FSS: x.visible = false: Next
        BGTimer.enabled = false : BGTimer2.enabled = false : BGTimer3.enabled = false
    End If
End Sub



'******************************************************
'                       KEYS
'******************************************************

Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftFlipperKey Then
        FlipperActivate LeftFlipper, LFPress
    PinCab_LeftFlipperButton.x = PinCab_LeftFlipperButton.x + 10
        If StagedFlipperMod <> 1 Then FlipperActivate2 LeftFlipper1, LFPress1
    End If
    If keycode = RightFlipperKey Then
    PinCab_RightFlipperButton.x = PinCab_RightFlipperButton.x - 10
        FlipperActivate RightFlipper, RFPress
        If StagedFlipperMod <> 1 Then FlipperActivate RightFlipper1, RFPress1
    End If
    If StagedFlipperMod = 1 Then
        If keycode = KeyUpperLeft Then FlipperActivate2 LeftFlipper1, LFPress1
        If keycode = KeyUpperRight Then FlipperActivate RightFlipper1, RFPress1
    End If

    If keycode = LeftTiltKey Then Nudge 90, 2:SoundNudgeLeft()
    If keycode = RightTiltKey Then Nudge 270, 2:SoundNudgeRight()
    If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter()
    If keycode = MechanicalTilt Then SoundNudgeCenter() : Controller.Switch(14) = 1
    If Keycode = KeyFront Then Controller.Switch(23) = 1 ' Buy in

    If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
        Select Case Int(rnd*3)
            Case 0: PlaySoundAtLevelStatic ("Coin_In_1"), CoinSoundLevel, sw18
            Case 1: PlaySoundAtLevelStatic ("Coin_In_2"), CoinSoundLevel, sw18
            Case 2: PlaySoundAtLevelStatic ("Coin_In_3"), CoinSoundLevel, sw18
        End Select
    End If

    If KeyCode = PlungerKey Then Plunger.Pullback:SoundPlungerPull()
    If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If keycode = LeftFlipperKey Then
        FlipperDeActivate LeftFlipper, LFPress
    PinCab_LeftFlipperButton.x = PinCab_LeftFlipperButton.x - 10
        If StagedFlipperMod <> 1 Then
            FlipperDeActivate2 LeftFlipper1, LFPress1
        End If
    End If
    If keycode = RightFlipperKey Then
        FlipperDeActivate RightFlipper, RFPress
    PinCab_RightFlipperButton.x = PinCab_RightFlipperButton.x + 10
        If StagedFlipperMod <> 1 Then
            FlipperDeActivate RightFlipper1, RFPress1
        End If
    End If

    If StagedFlipperMod = 1 Then
        If keycode = KeyUpperLeft Then
            FlipperDeActivate2 LeftFlipper1, LFPress1
        End If
        If keycode = KeyUpperRight Then
            FlipperDeActivate RightFlipper1, RFPress1
        End If
    End If

    If keycode = MechanicalTilt Then Controller.Switch(14) = 0
    If KeyCode = KeyFront Then Controller.Switch(23) = 0

    If KeyCode = PlungerKey Then
        Plunger.Fire
        If BIPL Then
            SoundPlungerReleaseBall()           'Plunger release sound when there is a ball in shooter lane
        Else
            SoundPlungerReleaseNoBall()         'Plunger release sound when there is no ball in shooter lane
        End If
    End If
    If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'************************************************
'*********   BallInit ***************************
'************************************************

Dim PowerBall, PowerBallID

Sub BallSearch() 'on hard pinmame reset check all these triggers and kickers
    dim x : for each x in Array(GumballPopper, sw84, sw85, sw88, sw58,sw18,sw25,sw16,sw15,sw17)
        if x.ballcntover then controller.Switch(x.uservalue) = True
    Next
    for each x in TZBalls : if x.y > 2500 then x.x = 204 : x.y = 519 : x.velx = 0 : x.vely = 0 : x.z = -30 : end if

    if sw15.ballcntover then    'Fix the trough powerball detecting switch
        if sw15.lastcapturedball.id <> PowerBallID then controller.Switch(26) = 1
    end if

: Next  'reset balls that have fallen off the table
End Sub

Sub SetPowerBall(ball)
    With ball
        .image = "PowerBall-GIOff-Env100-Clamped"
        .Mass = 0.8*Ballmass
        .BulbIntensityScale = 0.05
    .PlayfieldReflectionScale = 0.1
    End With
    PowerBallID = ball.id
    Set PowerBall = ball
End Sub


'**************************************************************
' SOLENOIDS
'**************************************************************

'     (*) - only in prototype, supported by rom 9.4
'    (**) - the additional GUM and BALL flashers were removed ro reduce cost
'   (***) - Gumball and Clock Mechanics are handled by vpm classes

'standard coils
SolCallback(1) = "SlotMachineKickout"                                   '(01) Slot Kickout
SolCallback(2) = "SolRocket"                                            '(02) Rocket Kicker
SolCallback(3) = "SolAutoKicker"                                        '(03) Auto-Fire Kicker
SolCallback(4) = "SolGumballPopper"                                     '(04) Gumball Popper
SolCallback(5) = "SolRightRampDiverter"                                 '(05) Right Ramp Diverter
SolCallback(6) = "SolGumballDiverter"                                   '(06) Gumball Diverter
SolCallback(7) = "SolKnocker"                                           '(07) Knocker
SolCallback(8) = "SolOuthole"                                           '(08) Outhole
SolCallback(9) = "SolBallRelease"                                       '(09) Ball Release
'SolCallback(10) = "SolRightSling"                                      '(10) Right Slingshot
'SolCallback(11) = "SolLeftSling"                                       '(11) Left Slingshot
'SolCallback(12) = "SolLowerBumper"                                     '(12) Lower Jet Bumper
'SolCallback(13) = "SolLeftBumper"                                      '(13) Left Jet Bumper
'SolCallback(14) = "SolRightBumper"                                     '(14) Right Jet Bumper
SolCallback(15) = "LockKickout"                                         '(15) Lock Release nf
SolCallback(16) = "SolShootDiverter"                                    '(16) Shooter Diverter
SolModCallback(17) = "UpdateF17"                                        '(17) Flasher bumpers x2
SolModCallback(18) = "FlashPWM 18, f18, BL_Flashers_f18, "              '(18) Flasher Power Payoff x2
SolModCallback(19) = "FlashPWM 19, f19, BL_Flashers_f19, "              '(19) Flasher Mini-Playfield x2
SolModCallback(20) = "FlashPWM 20, f20, BL_Flashers_f20, "              '(20) Flasher Upper Left Ramp x2 (**)
SolCallback(21) = "SolLeftMagnet"                                       '(21) Left Magnet
SolCallback(22) = "SolUpperRightMagnet"                                 '(22) Upper Right Magnet (*)
SolCallback(23) = "SolLowerRightMagnet"                                 '(23) Lower Right Magnet
SolCallback(24) = "SolGumballMotor"                                     '(24) Gumball Motor
SolCallback(25) = "SolMiniMagnet mLeftMini,"                            '(25) Left Mini-Playfield Magnet
SolCallback(26) = "SolMiniMagnet mRightMini,"                           '(26) Right Mini-Playfield Magnet
SolCallback(27) = "SolLeftRampDiverter"                                 '(27) Left Ramp Diverter
SolModCallback(28) = "FlashPWM 28, f28, BL_Flashers_f28, "              '(28) Flasher Inside Ramp
'aux board coils
SolModCallback(51) = "FlashPWM 37, f37, BL_Flashers_f37, "              '(37) Flasher Upper Right Flipper
SolModCallback(52) = "FlashPWM 38, f38, BL_Flashers_f38, "              '(38) Flasher Gumball Machine Higher
SolModCallback(53) = "FlashPWM 39, f39, BL_Flashers_f39, "              '(39) Flasher Gumball Machine Middle
SolModCallback(54) = "FlashPWM 40, f40, BL_Flashers_f40, "              '(40) Flasher Gumball Machine Lower
SolModCallback(55) = "FlashPWM 41, f41, BL_Flashers_f41, "              '(41) Flasher Upper Right Ramp x2 (**)
'SolCallback(56) = ""                                                   '(42) Clock Reverse (***)
'SolCallback(57) = ""                                                   '(43) Clock Forward (***)
'SolCallback(58) = ""                                                   '(44) Clock Switch Strobe (***)
'SolCallback(59) = "SolGumRelease"                                      'Fake Gumball Release (***) 'pinmame hack unreliable with solmodcallbacks

'fliptronic board
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sURFlipper) = "SolURFlipper"
SolCallback(sULFlipper) = "SolULFlipper"

Sub AdjustBulbTint(light, lightmaps)
  Dim p : p = light.GetInPlayIntensity / (light.Intensity * light.IntensityScale)
  Dim r : r = 1.0
  Dim g : g = p^0.15
  Dim b : b = (p - 0.04) / 0.96 : If b < 0 Then b = 0
  Dim i : i = .2126 * r^2.2 + .7152 * g^2.2 + .0722 * b^2.2
  Dim v : v = RGB(Int(255 * r), Int(255 * g), Int(255 * b))
  dim x : For Each x in lightmaps : x.color = v : x.opacity = 100.0 / i : Next
End Sub

Sub UpdateF17(v) : f17.State = v : f17b.State = v : AdjustBulbTint f17, BL_Flashers_f17 : End Sub

Sub FlashPWM(idx, light, lightmaps, pwm)
    light.State = pwm
    if idx = 18 Then l111.State = pwm
    AdjustBulbTint light, lightmaps
End Sub


'******************************************************
'                   KNOCKER
'******************************************************
Sub SolKnocker(Enabled)
    If enabled Then
        KnockerSolenoid
    End If
End Sub


'***********   Rocket   ********************************
Sub RocketKicker_Hit
    PlaysoundAtBallVol "fx_power", 0.3
    Controller.Switch(28) = 1
End Sub

Sub SolRocket(Enabled)
    If enabled Then
        If RocketKicker.BallCntOver = 0 Then
            PlaySoundAt SoundFX(SSolenoidOn,DOFContactors), RocketKicker
        Else
            PlaySoundAtVol SoundFX("fx_rocket_exit",DOFContactors), RocketKicker, 1
        End If
        RocketKicker.kick 304 + (Rnd*6), 45 + (Rnd* 20)
        Controller.Switch(28) = 0
    End If
End Sub

'***********   Autoplunger   ********************************
Sub AutoPlungerKicker_Hit
    Controller.Switch(72) = 1
    PlaysoundAt "fx_Lock_enter", AutoPlungerKicker
End Sub

Sub SolAutoKicker(Enabled)
    If enabled Then
        If AutoPlungerKicker.BallCntOver = 0 Then
            PlaySoundAt SoundFX("fx_AutoPlunger",DOFContactors), AutoPlungerKicker
        Else
            PlaySoundAtVol SoundFX("fx_launch",DOFContactors), AutoPlungerKicker, 0.3
        End If
        AutoPlungerKicker.kick 0, 40 + (Rnd* 10)
        Controller.Switch(72) = 0
    End If
End Sub

'***********   Gumball Popper   ******************************
Dim BIK:BIK=0       'ball in kicker

sub GumballPopper_Hit()
    Controller.Switch(74) = 1
    BIK=BIK+1
    PlaySoundAtBall "fx_kicker_catch"
end sub

sub GumballPopper_UnHit()
    Controller.Switch(74) = 0
    BIK=BIK-1
end sub

Sub GumballPopperHole_Hit
    PlaySoundAtVol "fx_Hole", GumballPopperHole, 0.05
    vpmTimer.PulseSw 51
End Sub

Sub SolGumballPopper(enabled)   'VUK
    If enabled Then
        BallSearch
        GumballPopper.Kick 0, 65, 1.5
        If BIK = 0 Then
            PlaySoundAt SoundFX(SSolenoidOn,DOFContactors), GumballPopper
        Else
            PlaySoundAt SoundFX("fx_GumPop",DOFContactors), GumballPopper
        End If
    End If
End Sub

'**************   Drain and Release   ************************************

Dim BIP:BIP = 0             'Balls In Play

sub sw18_hit()
    RandomSoundDrain sw18
    Controller.Switch(18) = 1
    BIP = BIP - 1
    If TVMod = 1 then TVTimer.enabled = 0:Frame.imageA = "tv_gameover"
end sub

sub sw18_unhit():controller.Switch(18) = 0:end sub

Sub SolOuthole(enabled)
    If Enabled Then
        'BallSearch
        sw18.kick 60, 9
        Updatetrough
        Playsoundat SoundFX(SSolenoidOn,DOFContactors), sw18
    End If
end sub

Sub SolBallrelease(enabled)
    If Enabled Then
        sw15.kick 60, 9
        Updatetrough                    'this is important to reset trough intervals
        BIP = BIP + 1
    End If
End sub

'******************************************************
'                       TROUGH
'******************************************************

sub sw25_hit():controller.Switch(25) = 1:updatetrough:end sub
sub sw25_unhit():controller.Switch(25) = 0:updatetrough:end sub

sub sw17_hit():controller.Switch(17) = 1:updatetrough:end sub
sub sw17_unhit():controller.Switch(17) = 0:updatetrough:end sub

sub sw16_hit():controller.Switch(16) = 1:updatetrough:end sub
sub sw16_unhit():controller.Switch(16) = 0:updatetrough:end sub

sub sw15_hit()
    if activeball.id = PowerBallID then     'opto handler
        controller.Switch(26) = 0   'if powerball
    Else
        controller.Switch(26) = 1   'if regular ball
    end if
    controller.Switch(15) = 1
    updatetrough:
end sub
sub sw15_unhit()
    controller.Switch(15) = 0
    controller.switch(26) = 0
    RandomSoundBallRelease sw15
    updatetrough
end sub

sub Updatetrough()
    updatetroughTimer.interval = 300
    updatetroughTimer.enabled = 1
end sub

sub updatetroughTimer_timer()
    if sw15.BallCntOver = 0 then sw16.kick 58, 8 end If
    if sw16.BallCntOver = 0 then sw17.kick 58, 8 end If
    if sw17.BallCntOver = 0 then sw25.Kick 58, 8 end If
    me.enabled = 0
end sub

'*******************   Lock   ******************

sub sw85_hit():controller.Switch(85) = 1:PlaysoundAt "fx_Lock_enter", sw85:updatelock:end sub
sub sw85_unhit():controller.Switch(85) = 0:end sub
sub sw84_hit():controller.Switch(84) = 1:updatelock:end sub
sub sw84_unhit():controller.Switch(84) = 0:end sub
sub sw88_hit():controller.Switch(88) = 1:updatelock:end sub
sub sw88_unhit():controller.Switch(88) = 0:end sub

sub lockramp_hit
    PlaySoundAtBallVolME "fx_metal_ramp_hit", 0.5
end sub

sub updatelock
    updatelocktimer.interval = 32
    updatelocktimer.enabled = 1
end sub

sub updatelocktimer_timer()
    if sw88.BallCntOver = 0 then sw84.kick 180, 2 end If
    if sw84.BallCntOver = 0 then sw85.kick 180, 2 end if
    me.enabled = 0
end sub

sub LockKickout(enabled)
    If enabled then
    sw88.kick 88, 10
    Playsoundat SoundFX("fx_Lock_exit",DOFContactors), sw88
    updatelock
    End If
end sub

'******************************************************************
'** SUBWAY, SHOOTER LANE, SLOTMACHINE; CAMERA, PIANO, DEAD END  ***
'******************************************************************

Sub SubwaySound(dummy)
    PlaySoundat "fx_subway", sw57
End sub

'********   Slot Machine    ****************************************

Sub SlotMachine_Hit()
    PlaySoundat "fx_SlotM_enter", slotMachine
End Sub

Sub sw57_Hit()                                          'submarine switch, Tslot proximity
    'debug.print activeball.id  & " " &  powerballid
    if activeball.id <> PowerBallID then
        vpmTimer.PulseSw 57
    end if
end Sub

Sub Sw58_Hit()
    Controller.Switch(58) = 1
    SlotKickerOverflow.Enabled = True
    Playsoundat "fx_kicker_catch", sw58
End Sub

Sub Sw58_UnHit()
    Controller.Switch(58) = 0
    SlotKickerOverflow.Enabled = False
End Sub

dim slotkick_vel, slotkick_vel_variance
dim slotkick_angle, slotkick_angle_variance

sub slotmachinekickout(enabled)
    if enabled Then
        Playsoundat SoundFX("fx_SlotM_exit",DOFContactors), sw58
        'If SlotKickerOverflow.ballcntover > 0 Then
            SlotKickerOverflow.Kick KickoutVariance(slotkick_angle,slotkick_angle_variance), KickoutVariance(slotkick_vel, slotkick_vel_variance)
        'else
            sw58.Kick KickoutVariance(slotkick_angle,slotkick_angle_variance), KickoutVariance(slotkick_vel, slotkick_vel_variance)
        'end if
    end if
end sub

Function KickoutVariance(aNumber, aVariance)    'strength, variance
    KickoutVariance = aNumber + ((Rnd*2)-1)*aVariance
End Function





'********   Shooter Lane    ***************************************

Sub ShooterLaneKicker_Hit
    PlaySoundAtVol "fx_hole",  ShooterLaneKicker, 0.15
    vpmtimer.addtimer 100, "SubwaySound"
End Sub

'********   Dead End   ********************************************

Sub DeadEnd_Hit
    vpmTimer.PulseSw 41
    PlaySoundAtVol "fx_DeadEnd", DeadEnd, 0.3
    vpmtimer.addtimer 100, "SubwaySound"
End Sub

'********   Camera   ***********************************************

Sub CameraKicker_Hit
    PlaySoundAtVol "fx_hole", CameraKicker, 0.15
    vpmtimer.addtimer 100, "SubwaySound"
End Sub

Sub sw42_Hit():vpmTimer.PulseSw 42:end Sub      'submarine switch, camera / upper playfield

Sub Hitch001_hit():PlaySoundAtBallVol "fx_lr2", 0.5:End Sub
Sub Hitch002_hit():PlaySoundAtBallVol "fx_lr3", 0.5:End Sub
Sub Hitch003_hit():PlaySoundAtBallVol "fx_lr4", 0.5:End Sub
Sub Hitch004_hit():If activeball.vely < 0 then:PlaySoundAtBallVol "fx_lr5", 0.5:End If:End Sub

'********  Piano   *************************************************

Sub Piano_Hit()
    PlaySoundAtVol "fx_Piano", Piano, 0.1
    vpmtimer.addtimer 100, "SubwaySound"
End Sub

Sub sw43_Hit():vpmTimer.PulseSw 43:end Sub      'submarine switch, piano

'*********************************************************************
'*****************       SLINGSHOTS       ****************************
'*********************************************************************

Dim LStep: LStep = 4: LeftSlingShot_Timer
Dim RStep: RStep = 4: RightSlingShot_Timer

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(Activeball)
  RandomSoundSlingshotLeft BM_SLING2
  LStep = -1 : LeftSlingShot_Timer ' Initialize Step to 0
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 10
    vpmTimer.PulseSw 34
End Sub

Sub LeftSlingShot_Timer
  Dim x, x1, x2, y: x1 = True:x2 = False:y = -20
    Select Case LStep
        Case 3:x1 = False:x2 = True:y = -10
        Case 4:x1 = False:x2 = False:y = 0:LeftSlingShot.TimerEnabled = 0
    End Select
  For Each x in BP_LSling1: x.Visible = x1: Next
  For Each x in BP_LSling2: x.Visible = x2: Next
  For Each x in BP_SLING2: x.transy = y: Next
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  RandomSoundSlingshotRight BM_SLING1
  RStep = -1 : RightSlingShot_Timer ' Initialize Step to 0
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 10
    vpmTimer.PulseSw 35
End Sub

Sub RightSlingShot_Timer
  Dim x, x1, x2, y: x1 = True:x2 = False:y = -20
    Select Case RStep
        Case 3:x1 = False:x2 = True:y = -10
        Case 4:x1 = False:x2 = False:y = 0:RightSlingShot.TimerEnabled = 0
    End Select
  For Each x in BP_RSling1: x.Visible = x1: Next
  For Each x in BP_RSling2: x.Visible = x2: Next
  For Each x in BP_SLING1: x.transy = y: Next
    RStep = RStep + 1
End Sub



'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************
' To add these slingshot corrections:
'   - On the table, add the endpoint primitives that define the two ends of the Slingshot
'   - Initialize the SlingshotCorrection objects in InitSlingCorrection
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
  AddSlingsPt 0, 0.00,   -5
  AddSlingsPt 1, 0.45,   -8
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  8
  AddSlingsPt 5, 1.00,  5
End Sub

Sub AddSlingsPt(idx, aX, aY)    'debugger wrapper for adjusting flipper script In-game
  Dim a
  a = Array(LS, RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub

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


''************************************************************************************
''*****************               Bumpers                 ****************************
''************************************************************************************

Sub Bumper1_Hit : vpmTimer.PulseSw(31) : RandomSoundBumper Bumper1 : End Sub
Sub Bumper1_Animate: Dim a, x: a = Bumper1.CurrentRingOffset: For Each x in BP_BR1: x.transz = a: Next: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(32) : RandomSoundBumper Bumper2 : End Sub
Sub Bumper2_Animate: Dim a, x: a = Bumper2.CurrentRingOffset: For Each x in BP_BR2: x.transz = a: Next: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(33) : RandomSoundBumper Bumper3 : End Sub
Sub Bumper3_Animate: Dim a, x: a = Bumper3.CurrentRingOffset: For Each x in BP_BR3: x.transz = a: Next: End Sub


'*******************************************************************
'****************   Diverters   ************************************
'*******************************************************************

Tween(TW_SHOOTER_DIV).Callback = GetRef("PrimArrayRotZUpdate")
Tween(TW_SHOOTER_DIV).CallbackParam = BP_ShooterDiv
SolShootDiverter False
Sub SolShootDiverter(Enabled)
    If Enabled Then
    PlaysoundAt SoundFX("fx_DivSS",DOFContactors), ShooterDiverter
    Tween(TW_SHOOTER_DIV).Clear().TargetLength(180, 100.0).Done
    ShooterDiverter.RotateToEnd
    Else
    Tween(TW_SHOOTER_DIV).Clear().TargetLength(118, 100.0).Done
    ShooterDiverter.RotateToStart
    End If
End Sub

Tween(TW_GUMBALL_DIV).Callback = GetRef("PrimArrayRotZUpdate")
Tween(TW_GUMBALL_DIV).CallbackParam = BP_DiverterP1
SolGumballDiverter False
Sub SolGumballDiverter(enabled)
    If Enabled Then
    PlaysoundAt SoundFX("fx_DivGM",DOFContactors), GumballDiverter
    Tween(TW_GUMBALL_DIV).Clear().TargetLength(39, 100.0).Done
    GumballDiverter.RotateToEnd
    Else
    Tween(TW_GUMBALL_DIV).Clear().TargetLength(75, 100.0).Done
    GumballDiverter.RotateToStart
    End If
End Sub


'********  Left Ramp Diverter   **********************

Tween(TW_LEFT_RAMP_DIV).Callback = GetRef("PrimArrayRotZUpdate")
Tween(TW_LEFT_RAMP_DIV).CallbackParam = BP_DiverterP
SolLeftRampDiverter False
Sub SolLeftRampDiverter(enabled)
    If Enabled Then
        PlaysoundAt SoundFX("fx_DivLR",DOFContactors), Plunger1
        RampDivWall.IsDropped=1
        Plunger1.pullback
    Tween(TW_LEFT_RAMP_DIV).Clear().TargetLength(0, 100.0).Done
    Else
        RampDivWall.IsDropped=0
        Plunger1.fire
    Tween(TW_LEFT_RAMP_DIV).Clear().TargetLength(22, 100.0).Done
    End If
End Sub


'********  Right Ramp Diverter   **********************

Dim KickerBall:Kickerball = Empty
Sub divTrig_Hit() : Set KickerBall = Activeball : End Sub
Sub divTrig_unHit() : KickerBall = Empty : End Sub

Sub divWall_Hit()
    WireRampOff
    'StopSound "fx_metalrolling"
    PlaySoundAtBallVol "fx_metalHit", 0.1
    If activeball.velx > 6 then activeball.velx = 1
End Sub

SolRightRampDiverter False
Sub SolRightRampDiverter(enabled)
    If enabled Then
        Playsoundat SoundFX("fx_DivRR",DOFContactors), DivTrig
        if Not IsEmpty(Kickerball) Then Kickball Kickerball, -10, 10, 0, 50
        Kickerball = Empty
    Tween(TW_RIGHT_RAMP_DIV).Clear().TargetLength(90-14, 200.0).Done
    Else
    Tween(TW_RIGHT_RAMP_DIV).Clear().TargetLength(  -14, 200.0).Done
    End If
    DivWall.collidable = not enabled
End Sub

Tween(TW_RIGHT_RAMP_DIV).Callback = GetRef("RightRampDiverter_Update")
Sub RightRampDiverter_Update(id, a)
  For each x in BP_RDiv : x.RotX = a : Next
  For each x in BP_SpiralToy : x.RotX = a : Next
End Sub

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
    dim rangle
    rangle = PI * (kangle - 90) / 180
    kball.z = kball.z + kzlift
    kball.velz = kvelz
    kball.velx = cos(rangle)*kvel
    kball.vely = sin(rangle)*kvel
End Sub


'*******************************************************************
'*************************       Targets        ********************
'*******************************************************************

Sub sw47_Hit:STHit 47:End Sub
Sub sw48_Hit:STHit 48:End Sub
Sub sw64_Hit:STHit 64:End Sub
Sub sw65_Hit:STHit 65:End Sub
Sub sw65a_Hit:STHit 165:End Sub
Sub sw66_Hit:STHit 66:End Sub
Sub sw67_Hit:STHit 67:End Sub
Sub sw68_Hit:STHit 68:End Sub
Sub sw77_Hit:STHit 77:End Sub
Sub sw78_Hit:STHit 78:End Sub

Sub UpdateStandupTargets
  dim BP, tx, px, py, pz

  pz = 82

'47 & 68 Swapped
  tx = BM_sw47.transx
  px = sw47.x
  py = sw47.y
  For each BP in BP_sw47: BP.transx = tx: Next

  tx = BM_sw48.transx
  px = sw48.x
  py = sw48.y
  For each BP in BP_sw48: BP.transx = tx: Next

  tx = BM_sw64.transx
  px = sw64.x
  py = sw64.y
  For each BP in BP_sw64: BP.transx = tx: Next

  tx = BM_sw65.transx
  px = sw65.x
  py = sw65.y
  For each BP in BP_sw65: BP.transx = tx: Next

  tx = BM_sw65a.transx
  px = sw65a.x
  py = sw65a.y
  For each BP in BP_sw65a: BP.transx = tx: Next

  tx = BM_sw66.transx
  px = sw66.x
  py = sw66.y
  For each BP in BP_sw66: BP.transx = tx: Next

  tx = BM_sw67.transx
  px = sw67.x
  py = sw67.y
  For each BP in BP_sw67: BP.transx = tx: Next

'68 & 47 Swapped
  tx = BM_sw68.transx
  px = sw68.x
  py = sw68.y
  For each BP in BP_sw68: BP.transx = tx: Next

  tx = BM_sw77.transx
  px = sw77.x
  py = sw77.y
  For each BP in BP_sw77: BP.transx = tx: Next

  tx = BM_sw78.transx
  px = sw78.x
  py = sw78.y
  For each BP in BP_sw78: BP.transx = tx: Next

End Sub

'******************************************************
'***************   Mini PF Switches *******************
'******************************************************

Sub sw44_Hit:vpmTimer.PulseSw 44: End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45 : End Sub
Sub sw45a_Hit:vpmTimer.PulseSw 45 : End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46 : End Sub
Sub sw46a_Hit:vpmTimer.PulseSw 46 : End Sub
Sub sw75_Hit:vpmTimer.PulseSw 75 : End Sub
Sub sw75_UnHit
    if activeball.vely < 0 Then  PlaySoundat "fx_power", sw75
End Sub
Sub sw76_Hit:vpmTimer.PulseSw 76: End Sub


'******************************************************
'***************  Ramps Switches **********************
'******************************************************

Sub sw53_Hit:vpmTimer.PulseSw 53:SoundPlayfieldGate: End Sub
Sub sw54_Hit:vpmTimer.PulseSw 54 :SoundPlayfieldGate: LRampSw.rotatetoend : End Sub
Sub sw54_UnHit: LRampSw.rotatetostart : End Sub
Sub sw73_Hit ' Enter wire ramp to mini playfield
    vpmTimer.PulseSw 73
    WireRampOn False
    'Playsoundatball "fx_metalrolling"
End Sub

Sub LRampSw_Animate: Dim a, x: a = -LRampSw.CurrentAngle: For Each x in BP_sw54p: x.RotX = a: Next: End Sub
Sub LRampG_Animate: Dim a, x: a = LRampG.CurrentAngle: For Each x in BP_sw53p: x.RotX = -a: Next: End Sub


'**************************************************************
'***************  Rollover Switches   *************************
'**************************************************************

Sub sw11_Hit: vpmTimer.PulseSw 11: End Sub
Sub sw11_Animate: Dim a, x: a = sw11.CurrentAnimOffset: For Each x in BP_sw11: x.transz = a: Next: End Sub
Sub sw12_Hit: vpmTimer.PulseSw 12: End Sub
Sub sw12_Animate: Dim a, x: a = sw12.CurrentAnimOffset: For Each x in BP_sw12: x.transz = a: Next: End Sub

Sub sw27_Animate: Dim a, x: a = sw27.CurrentAnimOffset: For Each x in BP_sw27: x.transz = a: Next: End Sub

Sub sw36_Hit: vpmTimer.PulseSw 36: End Sub
Sub sw36_Animate: Dim a, x: a = sw36.CurrentAnimOffset: For Each x in BP_sw36: x.transz = a: Next: End Sub
Sub sw37_Hit: vpmTimer.PulseSw 37: End Sub
Sub sw37_Animate: Dim a, x: a = sw37.CurrentAnimOffset: For Each x in BP_sw37: x.transz = a: Next: End Sub
Sub sw38_Hit: vpmTimer.PulseSw 38: End Sub
Sub sw38_Animate: Dim a, x: a = sw38.CurrentAnimOffset: For Each x in BP_sw38: x.transz = a: Next: End Sub
Sub sw52_Hit: vpmTimer.PulseSw 52: End Sub
'Sub sw52_Animate: Dim a, x: a = sw52.CurrentAnimOffset: For Each x in BP_sw52: x.transz = a: Next: End Sub ' Not baked since not visible
Sub sw56_Hit: vpmTimer.PulseSw 56: End Sub
Sub sw56_Animate: Dim a, x: a = sw56.CurrentAnimOffset: For Each x in BP_sw56: x.transz = a: Next: End Sub
Sub sw61_Hit: vpmTimer.PulseSw 61: End Sub
Sub sw61_Animate: Dim a, x: a = sw61.CurrentAnimOffset: For Each x in BP_sw61: x.transz = a: Next: End Sub
Sub sw62_Hit: vpmTimer.PulseSw 62: End Sub
Sub sw62_Animate: Dim a, x: a = sw62.CurrentAnimOffset: For Each x in BP_sw62: x.transz = a: Next: End Sub
Sub sw63_Hit: vpmTimer.PulseSw 63: End Sub
Sub sw63_Animate: Dim a, x: a = sw63.CurrentAnimOffset: For Each x in BP_sw63: x.transz = a: Next: End Sub

Sub Gate1_Animate: Dim a, x: a = Gate1.CurrentAngle: For Each x in BP_Gate1: x.RotX = -a: Next: End Sub
Sub Gate2_Animate: Dim a, x: a = Gate2.CurrentAngle: For Each x in BP_Gate2: x.RotX = -a: Next: End Sub

'******************************************************
'***************  Opto Switches   *********************
'******************************************************

Sub sw81_Hit:Controller.Switch(81) = 1:End Sub
Sub sw81_UnHit
    Controller.Switch(81) = 0
    If mLowerRightMagnet.MagnetOn and activeball.id <> PowerBallID then
        activeball.vely = 0
        activeball.velx = 0
    End If
End Sub

Sub sw82_Hit:Controller.Switch(82) = 1:End Sub
Sub sw82_UnHit
    Controller.Switch(82) = 0
    If mUpperRightMagnet.MagnetOn and activeball.id <> PowerBallID then
        activeball.vely = 0
        activeball.velx = 0
    End If
End Sub

Sub sw83_Hit:Controller.Switch(83) = 1:End Sub
Sub sw83_UnHit
    Controller.Switch(83) = 0
    If mLeftMagnet.MagnetOn and activeball.id <> PowerBallID then
        activeball.vely = 0
        activeball.velx = 0
    End If
End Sub

'Clock Pass Opto (only in prototypes,supported by rom version 9.4)
Sub sw86_Hit:Controller.Switch(86) = 1:End Sub
Sub sw86_UnHit:Controller.Switch(86) = 0:End Sub
'Gumball entry opto
sub sw87_hit():controller.switch(87) = 1:end Sub
sub sw87_unhit():controller.switch(87) = 0:end Sub
''Autoplunger 2nd Opto (only in prototypes,supported by rom version 9.4)
'Sub sw71_Hit:Controller.Switch(71) = 1:End Sub
'Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

'*************************************************************
'***************   Shooting lane   ***************************
'*************************************************************

Sub sw27_Hit()
    controller.switch(27) = 1
    BIPL = True
    If TVMod = 1 Then
        TVTimer.enabled = 1
        Frame.imageA = "tv_1"
    End If
End Sub

Sub sw27_UnHit():controller.switch(27) = 0:BIPL = False:End Sub


'******************************************************
'                   FLIPPERS
'******************************************************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
    If Enabled Then
        LF.Fire
        If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
            RandomSoundReflipUpLeft LeftFlipper
        Else
            SoundFlipperUpAttackLeft LeftFlipper
            RandomSoundFlipperUpLeft LeftFlipper
        End If
    Else
        LeftFlipper.RotateToStart
        If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
            RandomSoundFlipperDownLeft LeftFlipper
        End If
        FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        RF.Fire
        If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
            RandomSoundReflipUpRight RightFlipper
        Else
            SoundFlipperUpAttackRight RightFlipper
            RandomSoundFlipperUpRight RightFlipper
        End If
    Else
        RightFlipper.RotateToStart
        If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
            RandomSoundFlipperDownRight RightFlipper
        End If
        FlipperRightHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub SolULFlipper(Enabled)
    If Enabled Then
        LeftFlipper1.RotateToEnd
        If leftflipper1.currentangle < leftflipper1.endangle + ReflipAngle Then
            If StagedFlipperMod = 1 Then RandomSoundReflipUpLeft LeftFlipper1
        Else
            If StagedFlipperMod = 1 Then SoundFlipperUpAttackLeft LeftFlipper1
            If StagedFlipperMod = 1 Then RandomSoundFlipperUpLeft LeftFlipper1
        End If
    Else
        LeftFlipper1.RotateToStart
        If LeftFlipper1.currentangle < LeftFlipper1.startAngle - 5 Then
            If StagedFlipperMod = 1 Then RandomSoundFlipperDownLeft LeftFlipper1
        End If
        FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub SolURFlipper(Enabled)
    If Enabled Then
        RightFlipper1.RotateToEnd
        If Rightflipper1.currentangle < Rightflipper1.endangle + ReflipAngle Then
            If StagedFlipperMod = 1 Then RandomSoundReflipUpLeft RightFlipper1
        Else
            If StagedFlipperMod = 1 Then SoundFlipperUpAttackLeft RightFlipper1
            If StagedFlipperMod = 1 Then RandomSoundFlipperUpLeft RightFlipper1
        End If
    Else
        RightFlipper1.RotateToStart
        If RightFlipper1.currentangle < RightFlipper1.startAngle - 5 Then
            If StagedFlipperMod = 1 Then RandomSoundFlipperDownLeft RightFlipper1
        End If
        FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub UpdateFlipper(FlipperObj, FlipperSh, BM_Normal, BP_Normal, BM_NormalU, BP_NormalU, BM_Mod, BP_Mod, BM_ModU, BP_ModU)
  Dim a, v, LM : a = FlipperObj.CurrentAngle - 90
  FlipperSh.RotZ = a + 90
  v = CInt(100.0 * (FlipperObj.StartAngle - FlipperObj.CurrentAngle) / (FlipperObj.StartAngle - FlipperObj.EndAngle))
  If FlipMod = 0 Then
    For each LM in BP_Normal : LM.Rotz = a : LM.Opacity = 100 - v : Next
    For each LM in BP_NormalU : LM.Rotz = a : LM.Opacity = v : Next
    BM_Normal.visible  = v < 0.5
    BM_NormalU.visible = v > 0.5
  Else
    For each LM in BP_Mod : LM.Rotz = a : LM.Opacity = 100 - v : Next
    For each LM in BP_ModU : LM.Rotz = a : LM.Opacity = v : Next
    BM_Mod.visible  = v < 0.5
    BM_ModU.visible = v > 0.5
  End If
End Sub

Sub LeftFlipper_Animate : UpdateFlipper LeftFlipper, FlipperLSh, BM_FlipperL, BP_FlipperL, BM_FlipperLU, BP_FlipperLU, BM_FlipperSpL, BP_FlipperSpL, BM_FlipperSpLU, BP_FlipperSpLU : End Sub
Sub RightFlipper_Animate : UpdateFlipper RightFlipper, FlipperRSh, BM_FlipperR, BP_FlipperR, BM_FlipperRU, BP_FlipperRU, BM_FlipperSpR, BP_FlipperSpR, BM_FlipperSpRU, BP_FlipperSpRU : End Sub
Sub RightFlipper1_Animate : UpdateFlipper RightFlipper1, FlipperR1Sh, BM_FlipperR1, BP_FlipperR1, BM_FlipperR1U, BP_FlipperR1U, BM_FlipperSpR1, BP_FlipperSpR1, BM_FlipperSpR1U, BP_FlipperSpR1U : End Sub

Sub LeftFlipper1_Animate
  Dim a, LM : a = LeftFlipper1.CurrentAngle - 90
    FlipperL1Sh.RotZ = a + 90
  If FlipMod = 0 Then
    For each LM in BP_FlipperL1 : LM.Rotz = a : Next
  Else
    For each LM in BP_FlipperSpL1 : LM.Rotz = a : Next
  End If
End Sub


Sub LeftFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
    LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
    RightFlipperCollide parm
End Sub

Sub LeftFlipper1_Collide(parm)
    LeftFlipperCollide parm
End Sub

Sub RightFlipper1_Collide(parm)
    RightFlipperCollide parm
End Sub


'********************************************************************************
'******************  NFOZZY'S GUMBALL MACHINE  2 ********************************
'********************** UPDATED BY ROTHBAUERW ***********************************
'********************************************************************************

Sub GumKickout()    'unfreeze balls in gumball machine trough
    Freezekicker2.enabled = false
    Freezekicker1.enabled = false
    Freezekicker0.enabled = false

    Freezekicker2.kick 0,0,0
    Freezekicker1.kick 0,0,0
    Freezekicker0.kickz 0,0,0,-25

    FreezeKicker0.timerenabled = true
End Sub

FreezeKicker0.TimerInterval= 80 'interval for kickout, how long the dropwall stays down. Adjust me if it kicks out 2, or none.

Sub FreezeKicker0_Timer()   'repop gumball floor after a short delay
    Freezekicker0.enabled = true
    If Not FreezeKicker1.enabled then
        If CheckGumball(TZBall1) or CheckGumball(TZBall2) or CheckGumball(TZBall3) or CheckGumball(TZBall4) or CheckGumball(TZBall5) or CheckGumball(TZBall6) then
            Freezekicker1.enabled = true
        End If
    Elseif Not FreezeKicker2.enabled Then
        If CheckGumball2(TZBall1) or CheckGumball2(TZBall2) or CheckGumball2(TZBall3) or CheckGumball2(TZBall4) or CheckGumball2(TZBall5) or CheckGumball2(TZBall6) then
            Freezekicker2.enabled = true
            me.timerenabled = 0
        End If
    End If
End Sub

Function CheckGumball(ball)
    If Int(ball.x) = 211 and Int(ball.y) = 276 and Int(ball.z) = 171 Then
        CheckGumball = True
    Else
        CheckGumball = False
    End If
End Function

Function CheckGumball2(ball)
    If Int(ball.x) = 199 and Int(ball.y) = 236 and Int(ball.z) = 198 Then
        CheckGumball2 = True
    Else
        CheckGumball2 = False
    End If
End Function

Function CheckGumball3(ball)
    If Int(ball.x) = 191 and Int(ball.y) = 199 and Int(ball.z) = 231 Then
        CheckGumball3 = True
    Else
        CheckGumball3 = False
    End If
End Function

Sub SolGumRelease(enabled)  'this is a pinmame hack, will not work with solmodcallbacks. Called from motor sol instead.
    If enabled Then
        GumKickout                  'new
        vpmtimer.PulseSw 55 'Geneva switch
    End If
End Sub

Sub SolGumballMotor(aOn)
    if aOn then PlaySoundat SoundFX("fx_GumMachine",DOFGear), FreezeKicker0 : vpmtimer.addtimer 1400, "GumKnobTimer.enabled = 1'" : vpmtimer.addtimer 1700, "SolGumRelease 1'"
End Sub

GumKnobTimer.Interval = -1
Sub GumKnobTimer_Timer()    'prior 20ms period
    Dim a, x : a = BM_GMKnob.RotY + 1 * frametime
    If a >  360 then GumKnobTimer.enabled = 0 : a = 0
  For Each x in BP_GMKnob: x.RotY = a: Next
End Sub

'********************************************************************
'*************************   CLOCK   ********************************
'********************************************************************

Dim LastTime : LastTime = 0
dim LastClockIndex

Sub UpdateClock()
    Dim x, Time, Min, Hour, temp
    'Time = CInt(Controller.GetMech(0) )
    Time = Controller.GetMech(0)
    If Time <> LastTime Then
        Min = (Time Mod 60) * 6
        Hour = Int(Time / 2) - 45
    For Each x in BP_ClockLarge: x.RotY = min: Next
    For Each x in BP_ClockShort: x.RotY = hour: Next
        LastTime = Time '10.4 playsound args - name,loopcount,volume,pan,randompitch,pitch,UseExisting,Restart,Fade
        PlaySound SoundFXDOF("fx_motor",101,DOFPulse,DOFGear), -1, 1, 0.05, 0, 0, 1, 0
        LastClockIndex = 0
    Elseif LastClockIndex <=2 Then  'wait an update before stopping motor sound
        LastClockIndex = LastClockIndex + 1
    Elseif LastClockIndex > 2 then
        Stopsound "fx_motor"
    End If
End Sub

'**********************************************************************
'**************   POWER FIELD MAGNETS *********************************
'**********************************************************************

Sub SolMiniMagnet(aMag, enabled)
    If enabled Then
        PlaySoundat SoundFX("fx_magnet",DOFShaker), sw76
        With aMag
            .removeball PowerBall
            .MagnetOn = True
            .Update
            .MagnetOn = False
        End With
    End If
End Sub

'**********************************************************************
' SPECIAL CODE BY NFOZZY TO HANDLE MAGNET TRIGGERS
' based on the code by KIEFERSKUNK/DORSOLA
' Method: on extra triggers unhit, kill the velocity of the
' ball if the magnet is on, helping the magnet catch the ball.
'**********************************************************************

Sub sw81_help_unhit
    If activeball.vely > 28 then  activeball.vely = RndNumO (26,27)         '-ninuzzu- Let's slow down the ball a bit so the magnets can
    If activeball.vely < - 28 then  activeball.vely = - RndNumO (26,27)     'catch the ball
    If mLowerRightMagnet.MagnetOn = 1 and activeball.id <> PowerBallID then
        activeball.vely = activeball.vely * -0.2
        activeball.velx = activeball.velx * -0.2
    End If
End Sub

Sub LowerRightMagnet_hit()
    If mLowerRightMagnet.MagnetOn = 1 and activeball.id <> PowerBallID then
        activeball.vely = activeball.vely/10
        activeball.velx = activeball.velx/10
    End If
End Sub

Sub sw82_help_unhit
    If mUpperRightMagnet.MagnetOn = 1 and activeball.id <> PowerBallID then
        activeball.vely = activeball.vely * -0.2
        activeball.velx = activeball.velx * -0.2
    End If
End Sub

Sub UpperRightMagnet_hit()
    If mUpperRightMagnet.MagnetOn = 1 and activeball.id <> PowerBallID then
        activeball.vely = activeball.vely/10
        activeball.velx = activeball.velx/10
    End If
End Sub

Sub sw83_help_unhit
    If mLeftMagnet.MagnetOn = 1 and activeball.id <> PowerBallID then
        activeball.vely = activeball.vely * -0.2
        activeball.velx = activeball.velx * -0.2
    End If
End Sub

Sub LeftMagnet_hit()
    If mLeftMagnet.MagnetOn = 1 and activeball.id <> PowerBallID then
        activeball.vely = activeball.vely/10
        activeball.velx = activeball.velx/10
    End If
End Sub

Sub SolLeftMagnet(enabled)
    If enabled Then
        mLeftMagnet.MagnetOn = 1
        mleftmagnet.removeball PowerBall
        PlaySoundat SoundFX("fx_magnet_catch",DOFShaker), sw83
    Else
        mLeftMagnet.MagnetOn = 0
    End If
End Sub

Sub SolUpperRightMagnet(enabled)
    If enabled Then
        mUpperRightMagnet.MagnetOn = 1
        mUpperRightMagnet.removeball PowerBall
        PlaySoundat SoundFX("fx_magnet_catch",DOFShaker), sw82
    Else
        mUpperRightMagnet.MagnetOn = 0
    End If
End Sub

Sub SolLowerRightMagnet(enabled)
    If enabled Then
        mLowerRightMagnet.MagnetOn = 1
        mLowerRightMagnet.removeball PowerBall
        PlaySoundat SoundFX("fx_magnet_catch",DOFShaker), sw81
    Else
        mLowerRightMagnet.MagnetOn = 0
    End If
End Sub




'******************************************************
'       FLIPPER CORRECTION INITIALIZATION
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

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
    x.AddPt "Polarity", 1, 0.05, - 5
    x.AddPt "Polarity", 2, 0.16, - 5
    x.AddPt "Polarity", 3, 0.22, - 0
    x.AddPt "Polarity", 4, 0.25, - 0
    x.AddPt "Polarity", 5, 0.3, - 2
    x.AddPt "Polarity", 6, 0.4, - 3
    x.AddPt "Polarity", 7, 0.5, - 4.0
    x.AddPt "Polarity", 8, 0.7, - 3.5
    x.AddPt "Polarity", 9, 0.75, - 3.0
    x.AddPt "Polarity", 10, 0.8, - 2.5
    x.AddPt "Polarity", 11, 0.85, - 2.0
    x.AddPt "Polarity", 12, 0.9, - 1.5
    x.AddPt "Polarity", 13, 0.95, - 1.0
    x.AddPt "Polarity", 14, 1, - 0.5
    x.AddPt "Polarity", 15, 1.1, 0
    x.AddPt "Polarity", 16, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.15, 0.85
    x.AddPt "Velocity", 2, 0.2, 0.9
    x.AddPt "Velocity", 3, 0.23, 0.95
    x.AddPt "Velocity", 4, 0.41, 0.95
    x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
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

sub RightFlipper_timer()
    FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
    FlipperTricks2 LeftFlipper1, LFPress1, LFCount1, LFEndAngle1, LFState1, 1
    FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
    FlipperTricks RightFlipper1, RFPress1, RFCount1, RFEndAngle1, RFState1

    FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
    FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
  Dim BOT
  BOT = TZBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
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

dim LFPress, RFPress, LFPress1, RFPress1, LFCount, LFCount1, RFCount, RFCount1
dim LFState, LFState1, RFState, RFState1
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim EOST2, EOSA2,FReturn2
dim RFEndAngle, RFEndAngle1, LFEndAngle, LFEndAngle1

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
SOSRampup = 2.5

LFEndAngle = Leftflipper.endangle
LFEndAngle1 = Leftflipper1.endangle
RFEndAngle = RightFlipper.endangle
RFEndAngle1 = RightFlipper1.endangle

EOST2 = leftflipper1.eostorque
EOSA2 = leftflipper1.eostorqueangle
FReturn2 = leftflipper1.return

Const EOSTnew2 = 1.2 'EM

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn2 = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

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
    Dim b, BOT
    BOT = TZBalls

    For b = 0 To UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= - 0.4 Then BOT(b).vely =  - 0.4
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



Sub FlipperActivate2(Flipper, FlipperPress)
    FlipperPress = 1
    Flipper.Elasticity = FElasticity

    Flipper.eostorque = EOST2
    Flipper.eostorqueangle = EOSA2
End Sub

Sub FlipperDeactivate2(Flipper, FlipperPress)
    FlipperPress = 0
    Flipper.eostorqueangle = EOSA2
    Flipper.eostorque = EOST2*EOSReturn2/FReturn2


    If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
        Dim b, BOT
        BOT = TZBalls

        For b = 0 to UBound(BOT)
            If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
                If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
            End If
        Next
    End If
End Sub

Sub FlipperTricks2 (Flipper, FlipperPress, FCount, FEndAngle, FState, LiveCatchUpper)
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

        If LiveCatchUpper = 1 Then
            if GameTime - FCount < LiveCatch/2 Then
                Flipper.Elasticity = LiveElasticity
            elseif GameTime - FCount < LiveCatch Then
                Flipper.Elasticity = 0.1
            Else
                Flipper.Elasticity = FElasticity
            end if
        End If

        If FState <> 2 Then
            Flipper.eostorqueangle = EOSAnew
            Flipper.eostorque = EOSTnew
            Flipper.rampup = EOSRampup
            Flipper.endangle = FEndAngle
            FState = 2
        End If
    Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
        If FState <> 3 Then
            Flipper.eostorque = EOST2
            Flipper.eostorqueangle = EOSA2
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

Sub dPostsNoBounce_Hit(idx)
  RubbersD.dampen ActiveBall
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

dim cor : set cor = New CoRTracker

' The game timer interval must be 10 ms for the time being
CorTimer.Interval = 10
Sub CorTimer_timer() : cor.Update : End Sub

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
' ZRST: STAND-UP TARGETS by Rothbauerw
'******************************************************

Class StandupTarget
  Private m_primary, m_prims, m_sw, m_animate

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Prims(): Prims = m_prims: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public default Function init(primary, prims, sw, animate)
    Set m_primary = primary
    m_prims = prims
    m_sw = sw
    m_animate = animate

    Set Init = Me
  End Function
End Class

'Define a variable for each stand-up target
Dim ST47, ST48, ST64, ST65, ST65a, ST66, ST67, ST68, ST77, ST78

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:  vp target to determine target hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          transy must be used to offset the target animation
'   switch:  ROM switch number
'   animate:  Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST47 = (new StandupTarget)(sw47, BP_sw68, 47, 0)
Set ST48 = (new StandupTarget)(sw48, BP_sw48, 48, 0)
Set ST64 = (new StandupTarget)(sw64, BP_sw64, 64, 0)
Set ST65 = (new StandupTarget)(sw65, BP_sw65, 65, 0)
Set ST65a = (new StandupTarget)(sw65a, BP_sw65a, 65, 0)
Set ST66 = (new StandupTarget)(sw66, BP_sw66, 66, 0)
Set ST67 = (new StandupTarget)(sw67, BP_sw67, 67, 0)
Set ST68 = (new StandupTarget)(sw68, BP_sw47, 68, 0)
Set ST77 = (new StandupTarget)(sw77, BP_sw77, 77, 0)
Set ST78 = (new StandupTarget)(sw78, BP_sw78, 78, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST47, ST48, ST64, ST65, ST65a, ST66, ST67, ST68, ST77, ST78)

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.2    'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
  Dim i
  i = STArrayID(switch)

  PlayTargetSound
  STArray(i).animate = STCheckHit(ActiveBall,STArray(i).primary)

  If STArray(i).animate <> 0 Then
    DTBallPhysics ActiveBall, STArray(i).primary.orientation, STMass
  End If
  DoSTAnim
End Sub

Function STArrayID(switch)
  Dim i
  For i = 0 To UBound(STArray)
    If STArray(i).sw = switch Then
      STArrayID = i
      Exit Function
    End If
  Next
End Function

Function STCheckHit(aBall, target) 'Check if target is hit on it's face
  Dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
  rangle = (target.orientation - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
  paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

  perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

  If perpvel > 0 And  perpvelafter <= 0 Then
    STCheckHit = 1
  ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
    STCheckHit = 1
  Else
    STCheckHit = 0
  End If
End Function

Sub DoSTAnim()
  Dim i
  For i = 0 To UBound(STArray)
    STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prims,STArray(i).sw,STArray(i).animate)
  Next
End Sub

Function STAnimate(primary, prims, switch,  animate)
  Dim animtime, a, y

  STAnimate = animate

  If animate = 0  Then
    primary.uservalue = 0
    STAnimate = 0
    Exit Function
  ElseIf primary.uservalue = 0 Then
    primary.uservalue = GameTime
  End If

  animtime = GameTime - primary.uservalue

  If animate = 1 Then
    primary.collidable = 0
    For Each a In prims: a.transy =  - STMaxOffset: Next
    vpmTimer.PulseSw switch
    STAnimate = 2
  ElseIf animate = 2 Then
    y = prims(0).transy + STAnimStep
    If y >= 0 Then
      y = 0
      primary.collidable = 1
      STAnimate = 0
    Else
      STAnimate = 2
    End If
    For Each a In prims: a.transy =  y: Next
  End If
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


'******************************************************
'****   END STAND-UP TARGETS
'******************************************************


'***************************************************************
'   BALL SHADOWS
'***************************************************************


Const fovY                  = 0     'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const AmbientBSFactor       = 0.8   '0 to 1, higher is darker
Const AmbientMovement       = 2     '1 to 4, higher means more movement as the ball moves left and right

'****** Part C:  The Magic ******
Dim sourcenames, DSSources(30), DSGISide(30)
sourcenames = Array ("","","","","","","","","","","","")

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objBallShadow(12)

DynamicBSInit

sub DynamicBSInit()
    Dim iii
    for iii = 0 to tnob                                 'Prepares the shadow objects before play begins
        Set objBallShadow(iii) = Eval("BallShadow" & iii)
        objBallShadow(iii).material = "BallShadow" & iii
        UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
        objBallShadow(iii).Z = 1 + iii/1000 + 0.04
        objBallShadow(iii).visible = 0
    Next
end sub


Sub DynamicBSUpdate
    Dim BOT: BOT = TZBalls ' All balls, use GetBalls() if not available through an array
    Dim s, LSd, iii
    'Hide shadow of deleted balls
    For s = UBound(BOT) + 1 to tnob
        objBallShadow(s).visible = 0
    Next
    If UBound(BOT) < lob Then Exit Sub      'No balls in play, exit
    For s = lob to UBound(BOT)
    If BOT(s).Z > 30 Then                           'The flasher follows the ball up ramps while the primitive is on the pf
      If BOT(s).X < tablewidth/2 Then
        objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
      Else
        objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
      End If
      objBallShadow(s).Y = BOT(s).Y + BallSize/10 + fovY
      objBallShadow(s).visible = 1

    Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then    'On pf, primitive only
      objBallShadow(s).visible = 1
      If BOT(s).X < tablewidth/2 Then
        objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
      Else
        objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
      End If
      objBallShadow(s).Y = BOT(s).Y + fovY
    Else                                            'Under pf, no shadows
      objBallShadow(s).visible = 0
    end if
    Next
End Sub
'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************


'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
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


'******************************************************
'****  BALL ROLLING AND DROP SOUNDS
'******************************************************
'
' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 To tnob
    rolling(i) = False
  Next
End Sub

Sub RollingUpdate()
  Dim b
  Dim BOT
  BOT = TZBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 To tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(BOT)
    If BallVel(BOT(b)) > 1 And BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If BOT(b).VelZ <  - 1 And BOT(b).z < 55 And BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz >  - 7 Then
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


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************



'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'     * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'     * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'     * Create a Timer called RampRoll, that is enabled, with a interval of 100
'     * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'     * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'     * To stop tracking ball
'        * call WireRampOff
'        * Otherwise, the ball will auto remove if it's below 30 vp units
'

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
' Setup:  Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RampBalls(6,2)
Dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(6)

Sub WireRampOn(input)
  Waddball ActiveBall, input
  RampRollUpdate
End Sub

Sub WireRampOff()
  WRemoveBall ActiveBall.ID
End Sub

' WaddBall (Active Ball, Boolean)
Sub Waddball(input, RampInput) 'This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  Dim x
  For x = 1 To UBound(RampBalls)  'Check, don't add balls twice
    If RampBalls(x, 1) = input.id Then
      If Not IsEmpty(RampBalls(x,1) ) Then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 To UBound(RampBalls)
    If IsEmpty(RampBalls(x, 1)) Then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      Exit Sub
    End If
    If x = UBound(RampBalls) Then  'debug
      Debug.print "WireRampOn error, ball queue Is full: " & vbNewLine & _
      RampBalls(0, 0) & vbNewLine & _
      TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbNewLine & _
      TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbNewLine & _
      TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbNewLine & _
      TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbNewLine & _
      TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbNewLine & _
      " "
    End If
  Next
End Sub

' WRemoveBall (BallId)
Sub WRemoveBall(ID) 'This subroutine is called from the RampRollUpdate subroutine and is used to remove and stop the ball rolling sounds
  '   Debug.Print "In WRemoveBall() + Remove ball from loop array"
  Dim ballcount
  ballcount = 0
  Dim x
  For x = 1 To UBound(RampBalls)
    If ID = RampBalls(x, 1) Then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    If Not IsEmpty(Rampballs(x,1)) Then ballcount = ballcount + 1
  Next
  If BallCount = 0 Then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer()
  RampRollUpdate
End Sub

Sub RampRollUpdate()  'Timer update
  Dim x
  For x = 1 To UBound(RampBalls)
    If Not IsEmpty(RampBalls(x,1) ) Then
      If BallVel(RampBalls(x,0) ) > 1 Then ' if ball is moving, play rolling sound
        If RampType(x) Then
          PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      End If
      If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
  Next
  If Not RampBalls(0,0) Then RampRoll.enabled = 0
End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  Me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbNewLine & _
  "1 " & TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbNewLine & _
  "2 " & TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbNewLine & _
  "3 " & TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbNewLine & _
  "4 " & TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbNewLine & _
  "5 " & TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbNewLine & _
  "6 " & TypeName(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbNewLine & _
  " "
End Sub

Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
  BallPitch = pSlope(BallVel(ball), 1, - 1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, - 4000, 60, 7000)
End Function



'******************
'  Ramp Triggers
'******************

Sub LREnter_Hit() ' Enter left plastic ramp
    If ActiveBall.VelY < 0 Then PlaySoundAtBallVol "fx_rlenter", Vol(Activeball)*VolumeDial*10
    WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub WirerampSound1_Hit()
  WireRampOff
End Sub

Sub WirerampSound1_UnHit() ' Enter wire part after plastic (to left side)
    WireRampOn False
End Sub

Sub Balldrop1_Hit() ' Drop from ramp on left side
    WireRampOff
End Sub

Sub WirerampSound2_Hit()
  WireRampOff
End Sub

Sub WirerampSound2_UnHit() ' Enter wire part after plastic (to right side)
    WireRampOn False
End Sub

Sub Balldrop3_Hit() ' Drop from ramp on right side
    WireRampOff
End Sub

Sub RREnter_Hit() ' Enter right metal ramp
    If ActiveBall.VelY < 0 Then PlaySoundAtBallVol "fx_metal_ramp_hit", Vol(Activeball)
    WireRampOn False
End Sub

Sub WirerampSoundStop_Hit() ' Enter mini playfield after wire ramp
    WireRampOff
End Sub



'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************





'******************************************************
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

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
RubberFlipperSoundFactor = 0.075 / 5      'volume multiplier; must not be zero
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
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBallVolME (Soundname, aVol)
    Playsound soundname, 1,aVol, AudioPan(ActiveBall), 0,0,1,0, AudioFade(ActiveBall)
End Sub



'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / tableheight - 1

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
  tmp = tableobj.x * 2 / tablewidth - 1

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

Sub RandomSoundBumper(Bump)
  Select Case Int(Rnd * 3) + 1
    Case 1: RandomSoundBumperTop Bump
    Case 2: RandomSoundBumperMiddle Bump
    Case 3: RandomSoundBumperBottom Bump
  End Select
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
  TargetBouncer ActiveBall, 1
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

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////

Sub OnBallBallCollision(ball1, ball2, velocity)

  FlipperCradleCollision ball1, ball2, velocity

  Dim snd
  Select Case Int(Rnd * 7) + 1
    Case 1
      snd = "Ball_Collide_1"
    Case 2
      snd = "Ball_Collide_2"
    Case 3
      snd = "Ball_Collide_3"
    Case 4
      snd = "Ball_Collide_4"
    Case 5
      snd = "Ball_Collide_5"
    Case 6
      snd = "Ball_Collide_6"
    Case 7
      snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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
'***   VR Backglass Code
'******************************************************

dim Fl1lvl,Fl1On, Fl2lvl,Fl2On, Fl3lvl, Fl3On, Fl4lvl, Fl4On, Fl5lvl, Fl5On, Fl6lvl, Fl6On

Sub InitLampsNF()
    ' FIXME finish conversion and remove
    ' Backglass flashers
    'FlSol19,FlL16a,FlL16b,FlL24,FlL44a,FlL44b,FlL53a,FlL53b,FlL53c,FlL53d,FlL53f, FlL54,FlL61,FlL62,FlL63, FlL74a,FlL74b,FlL85
    Lampz.MassAssign(16) = FlL16a
    Lampz.MassAssign(16) = FlL16b
    Lampz.MassAssign(24) = FlL24
    Lampz.MassAssign(44) = FlL44a
    Lampz.MassAssign(44) = FlL44b
    Lampz.MassAssign(53) = FlL53a
    Lampz.MassAssign(53) = FlL53b
    Lampz.MassAssign(53) = FlL53c
    Lampz.MassAssign(53) = FlL53d
    Lampz.MassAssign(53) = FlL53f
    Lampz.MassAssign(54) = FlL54
    Lampz.MassAssign(61) = FlL61
    Lampz.MassAssign(62) = FlL62
    Lampz.MassAssign(63) = FlL63
    Lampz.MassAssign(74) = FlL74a
    Lampz.MassAssign(74) = FlL74b
    Lampz.MassAssign(85) = FlL85
    Lampz.MassAssign(119) = FlSol19
End Sub


'backglass positioning
Sub SetBackglass()
    Dim obj

    For Each obj In VRBackglass
        obj.x = obj.x + 25
        obj.height = - obj.y + 263
        obj.y = -45 'adjusts the distance from the backglass towards the user
        obj.rotx=-87.5
    Next

End Sub

'Blinking lights
'Fade code
Sub BGTimer_timer
    dim blinklight
    If VRBGGI009.opacity = 100 Then 'If Main GI is off, sets blinking lights to off
        For each blinklight in VRBGBlink : Blinklight.visible = True : Next
    Else
        For each blinklight in VRBGBlink : Blinklight.visible = False : Next
    End If

    VRBGFL015.opacity = 65 * Fl1lvl^1.5
    VRBGFL015_2.opacity = 65 * Fl1lvl^2
    VRBGFL001.opacity = 50 * Fl2lvl^1.5
    VRBGFL001_2.opacity = 25 * Fl2lvl^2
    VRBGFL005.opacity = 30 * Fl3lvl^1.5
    VRBGFL005_2.opacity = 40 * Fl3lvl^2
    VRBGFL004.opacity = 50 * Fl4lvl^1.5
    VRBGFL004_2.opacity = 35 * Fl4lvl^2
    VRBGFL006.opacity = 65 * Fl5lvl^1.5
    VRBGFL006_2.opacity = 65 * Fl5lvl^2
    VRBGFL007.opacity = 65 * Fl6lvl^1.5
    VRBGFL007_2.opacity = 50 * Fl6lvl^2

    If Fl1on = True Then
        Fl1lvl = Fl1lvl * 0.85 - 0.01
        if Fl1lvl < 0 Then
            Fl1lvl = 0
            Fl1on = False
        End If
    End If
    If Fl2on = True Then
        Fl2lvl = Fl2lvl * 0.85 - 0.01
        if Fl2lvl < 0 then
            Fl2lvl = 0
            Fl2On = False
        End If
    End If
    If Fl3On = True Then
        Fl3lvl = Fl3lvl * 0.85 - 0.01
        if Fl3lvl < 0 then
            Fl3lvl = 0
            Fl3On = False
        End If
    End If
    If Fl4On = True Then
        Fl4lvl = Fl4lvl * 0.85 - 0.01
        if Fl4lvl < 0 then
            Fl4lvl = 0
            Fl4On = False
        End If
    End If
    If Fl5On = True Then
        Fl5lvl = Fl5lvl * 0.85 - 0.01
        if Fl5lvl < 0 then
            Fl5lvl = 0
            Fl5On = False
        End If
    End If
    If Fl6On = True Then
        Fl6lvl = Fl6lvl * 0.85 - 0.01
        if Fl6lvl < 0 then
            Fl6lvl = 0
            Fl6On = False
        End If
    End If
end Sub

Dim Flasherseq1
Dim Flasherseq2

'faster blink
Sub BGTimer2_Timer
    Select Case Flasherseq1 'Int(Rnd*6)+1
        Case 1:  FL1lvl = 1 : Case 2:  FL1On = True
        Case 15: Fl4lvl = 1 : Case 16: Fl4On = True
        Case 30: Fl3lvl = 1 : Case 31: Fl3On = True
        Case 41: Fl1lvl = 1 : Case 42: Fl1On = True
        Case 55: Fl4lvl = 1 : Case 56: Fl4On = True
        Case 60: Fl3lvl = 1 : Case 61: Fl3On = True
        Case 81: Fl1lvl = 1 : Case 82: Fl1On = True
        Case 90: Fl3lvl = 1 : Case 91: Fl3On = True
        Case 95: Fl4lvl = 1 : Case 96: Fl4On = True

    End Select
    Flasherseq1 = Flasherseq1 + 1
    If Flasherseq1 > 120 Then
    Flasherseq1 = 1
    End if
End Sub

'Slower Blink
Sub BGTimer3_Timer
    Select Case Flasherseq2 'Int(Rnd*6)+1
        Case 1: Fl2lvl = 1 : Case 2: Fl2On = True
        Case 10: Fl5lvl = 1 : Case 11: Fl5On = True
        Case 30: Fl6lvl = 1 : Case 31: Fl6On = True
        Case 41:Fl2lvl = 1 : Case 42: Fl2On = True
        Case 45: Fl5lvl = 1 : Case 46: Fl5On = True
        Case 60: Fl6lvl = 1 : Case 61: Fl6On = True
        Case 75: Fl5lvl = 1 : Case 76: Fl5On = True
        Case 81:Fl2lvl = 1 : Case 82: Fl2On = True
        Case 90: Fl6lvl = 1 : Case 91: Fl6On = True
        Case 105: Fl5lvl = 1 : Case 106: Fl5On = True
        Case 120: Fl6lvl = 1 : Case 121: Fl6On = True
    End Select
    Flasherseq2 = Flasherseq2 + 1
    If Flasherseq2 > 123 Then
    Flasherseq2 = 1
    End if
End Sub

'*********************
' ZVRR - VR Room
'*********************

Dim VRRoom
Dim Stuff

LoadVRRoom

Sub LoadVRRoom

  For Each Stuff in VRBackglass: Stuff.Visible = 1: Next
  For each Stuff in VR_Min: Stuff.visible = 0: Next
  For each Stuff in VR_360: Stuff.visible = 0: Next
  For each Stuff in VR_Cab: Stuff.visible = 0: Next

  If RenderingMode = 2 or VRDesktopSim Then
    VRRoom = VRRoomChoice
    For each Stuff in bmrails: Stuff.visible = 0: Next
  Else
    VRRoom = 0
  End If

  If VRRoom = 1 Then
    For each Stuff in VR_Cab: Stuff.visible = 1: Next
    For each Stuff in VR_Min: Stuff.visible = 1: Next
  End If

  If VRRoom = 2 Then
    For each Stuff in VR_Cab: Stuff.visible = 1: Next
    For each Stuff in VR_360: Stuff.visible = 1: Next
  End If

  If VRRoom = 3 Then
    For each Stuff in VR_Cab: Stuff.visible = 1: Next
  End If

End Sub

' VPW Changelog
' =============
' 2018-07-24 Thalamus - Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions", Changed UseSolenoids=1 to 2
' 2018-08-09 Wob - Added vpmInit Me to table init and cSingleLFlip
' 2018-08-18 DJRobX - Remove CSingleLFlip - there is an ULFlip sub, Add more SSF and VPMModSol code.
' 2018-10-09 nf  - rescripted gumball machine, slot machine kickout, clock sfx, added ballsearch on hard pinmame reset / find lost balls
' 2022-07-17 rothbauerw - updated physics to latest including adjusting the slot kickout, slings, adding stand up target code, flipper updates, dampener update, added VR plunger and a few other minor VR tweaks
' 2022-07-22 Niwak  - Complete bake/lightmap of the table, fix ramp height and small adjustmennts, interactive UI for the options
' 2022-07-25 leojreimrc - VR/FSS Backglass
'
' 2.3.7 - Sixtoe - Significant rebuilding, refactoring and tidying of table and layout, physics changes and removal of significant number of redundant assets.
' 2.3.8 - Sixtoe - Continuing work
' 2.3.9 - Sixtoe - Fixed and tuned tons of little things, now playable again.
' 2.4.0 - Sixtoe - Realigned most things to new dimensions and playfield.
' 2.4.1 - Niwak - New renders, new script for 10.8
' 2.4.2 - Sixtoe - Fixed some physical issues, still some work to do but works better than it did.
' 2.4.3 - Niwak - New batch, double bake for flippers
' 2.4.4 - Niwak - Script cleanup, fix mods animations, fix bumper cap materials,...
' 2.4.5 - Niwak - Use latest PWM output in PinMame 3.6 beta - Part I
' 2.4.6 - Niwak - Add rules, Use latest PWM output in PinMame 3.6 beta - Part II
' 2.4.7 - Niwak - Fix powerball rendering
' 2.4.8 - Niwak - Adjust ball reflections and brightness
' 2.4.9 - Sixtoe - Added blocker wall under shooter lane because of weird powerball things
' 2.5.0 - Sixtoe - Made blocker wall cover flipper area (thanks RothbauerW)
' 2.5.1 - Niwak - Cleanup timers
' 2.5.2 - Niwak - More cleanups and optimizations. Do not use core.vbs solenoid handling as it uses 'Execute' which can be very slow
' 2.5.3 - Niwak - Uses new core.vbs solenoid callback that support FastFlips and avoids stutters
' 2.5.4 - Niwak - Revert SolCallback2 as it now has a better solution integrated in VPX
' 2.5.5 - leojreimrc - VR FSS Backglass
' 2.5.6 - Niwak - New bake with overall improvments and some refractions
' 2.5.7 - Sixtoe - Middle Ramp and Pop Bumper area ball traps fixed
' 2.5.8 - Sixtoe - Fixed double sling assignments, fixed mixed up pop bumper animations, hid physical objects, turned down power of hidden AutoPlungerKicker and adjusted it's orbit feed to stop it bricking.
' 2.5.9 - Sixtoe - Hooked up VR Room to options, tweaked entrance to upper playfield,
' 2.5.10 - apophis - Physics script updates. Flipper trigger updates. Fleep script and sounds updates. Ball/ramp rolling sound updates. Fixed ball shadows. Fixed PF reflections.
' 2.5.11 - apophis - Fixed flipper mod issue (no more z-fighting). Added sidewall mod option that was missing. Added refraction probe option. Updated room brightness sub. Updated DisableStaticPreRendering usage to be compatible with 10.8.1
' 2.5.12 - mcarter78 - Replace VR logo text, add desktop VR sim, add VRStuff collection to prevent errors
' rc1 - Sixtoe - Adjusted rails and rail lightmaps (thanks Cliffy), updated POV and changed to camera, updated sling posts, reinforced layout a bit more, refactored drain area,
' rc2 - leojreimroc - VR Backglass fixes, changed size to fix new backbox, adjusted blinking lights
' rc3 - Sixtoe - Rebuilt entire plunger / diverter area, Twilight Zone 360 room text made bigger and centered, created bmrails collection and fixed some vr room options, moved the trigger heights for dead end and camera playfield holes, added audio to rollovers
' rc4 - DGrimmReaper - VR FlipperButtons animated and changed to Yellow
' rc5 - apophis - Updated flipper triggers. More plunger lane fortifications. Set "VR_TX 3D" object to toy.
' rc6 - apophis - Dumb fix for ball falling thru plunger lane.
' rc7 - apophis - Added "Show Rails" option. Removed targetbouncer from bottom sling posts.
' rc8 - Sixtoe - Updated playfield mesh with piano indent, added sw45&sw45a&sw46&sw46a to upper playfield (doh), corrected sideblades, corrected 360 floor, adjusted several audio triggers, fixed incorrect diverter animation causing it to stick out side of cabinet, fixed gate animations, added plunger blocker wall, adjusted upper playfield feed to stop trapped balls in fringe cases, adjusted gate1 position, adjusted flipper sizes and trigger areas, numerous other small tweaks
' rc9 - Sixtoe - Rebuilt the plunger lane area, twice, still didn't work, reconfigured the plunger object and it seems to have fixed it, tweaked flipper triggers, added a safety ramp to feed ball from the autoplunger, turned down piano, added plunger lane scoop guide, script cleanup
' rc10 - Sixtoe - Turned down flippers to match their irl coils (they use oranges not blues), levelled pf mesh, did some more work to try and prevent stuck balls, made upper playfield ejector more sticky, reduced audio on several shots
' rc11 - Sixtoe - Edited upper PF nestmap to make entrance transparent, adjusted glass position, adjusted GI level on clock hands, fixed script for rocket (thanks Niwak), fixed bumper socket animation, fixed mixed up animated targets, fixed switch in plunger lane.
' rc12 - mcarter78 - fix ramp sound transitions, remove textbox
' rc13 - Sixtoe - Made blocker wall full table size (thanks rothbauerw) should hopefully fix powerball issues, added movable outlane difficulty posts to options
' rc14 - apophis - Added new outlane post prims to dPosts. Minor script fixes. Updated Table Info description.
' rc15 - apophis - Fixed stuck ball issue in subway near sw64 and sw65a.
' rc16 - apophis - Slighlty adjusted inlane cradle geometry for better post passes.
' rc17 - Sixtoe - Visual blocker walls, script cleanup, added option for tz_94f rom,
' rc18 - apophis - BM_PF depth bias set to -500 (thanks SG1bsoN!). Updated table info description.
' 1.0.1 - Niwak - Replace diverter animation by tweens, fix right ramp diverter, fix right flipper anim, fix slot mod reels depth bias

'
'**Requires Re-Render;
'Transparent plastic at the end of shooter lane
'Render Table Lighter
'Layer Mask Upper Playfield
'Clock Upper Target Brighten
'Fix Inserts
'Unwrap Sidewalls

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub

