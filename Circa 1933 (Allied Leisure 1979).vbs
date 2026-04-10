Option Explicit
LoadVPM "01560000","ALI.VBS",3.1

Sub LoadVPM(VPMver,VBSfile,VBSver)
 On Error Resume Next
    If ScriptEngineMajorVersion<5 Then MsgBox"VB Script Engine 5.0 or higher required"
    ExecuteGlobal GetTextFile(VBSfile)
    If Err Then MsgBox"Unable to open "&VBSfile&". Ensure that it is in the same folder as this table."&vbNewLine&Err.Description:Err.Clear
   Set Controller = CreateObject("B2S.Server")
    If Err Then MsgBox"Can't Load VPinMAME."&vbNewLine&Err.Description
    If VPMver>"" Then If Controller.Version<VPMver Or Err Then MsgBox"VPinMAME ver "&VPMver& " required.":Err.Clear
   If VPinMAMEDriverVer<VBSver Or Err Then MsgBox VBSFile&" ver "&VBSver&" or higher required."
  On Error Goto 0
End Sub

' Thalamus - used twice - last will be used anyway
' Sub Table1_Exit
' Controller.Stop
' End Sub

Const cGameName="circa33",cCredits="Circa 1933",UseSolenoids=1,UseLamps=1,UseGI=0,UseSync=1
Const SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",sCoin="Coin3"

SolCallback(1)="vpmSolSound ""Jet3"","         'Sol1
SolCallback(2)="vpmSolSound ""Jet3"","         'Sol2
SolCallback(3)="vpmSolSound ""Jet3"","         'Sol3
SolCallback(6)="vpmSolSound ""Jet3"","         'Sol4
SolCallback(7)="vpmSolSound ""Jet3"","         'Sol5
SolCallback(8)="bsTrough.SolOut"           'Sol6
SolCallback(9)="dtDrop.SolUnhit 1,"            'Sol7
SolCallback(10)="dtDrop.SolUnhit 2,"         'Sol8
SolCallback(11)="dtDrop.SolUnhit 3,"         'Sol9
SolCallback(12)="dtDrop.SolUnhit 4,"         'Sol10
'17: Total play counter
'18: total replay counter
SolCallback(19)="vpmSolSound ""ALBell1000"","     'Sol13  1,000pts
SolCallback(20)="vpmSolSound ""ALBell100"","      'Sol14  100pts
SolCallback(21)="vpmSolSound ""ALBell10"","       'Sol15  10pts
SolCallback(22)="vpmSolSound ""Knocker"","       'Sol16
SolCallback(25)="vpmNudge.SolGameOn"          'Sol17
'26: Gate one (opens with switch #17, closes with #19)
'27: Gate two (opens with switch #18, closes with #20)
SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"

Sub Table1_Paused:Controller.Pause=True:End Sub
Sub Table1_unPaused:Controller.Pause=False:End Sub

Dim bsTrough,dtDrop,Obj
'*************************** CUSTOM DIP CODE ***************************
Dim D0,D1,D2,D3,D4,D5,D6,D7,D8
Dim Setting(72)

Sub Table1_Init
   T1a.IsDropped=1:T2a.IsDropped=1:T3a.IsDropped=1
 D0=LoadValue("circa33","D00")
 If D0="" Then 'Data doesn't exist, create it
    D0=56
   D1=153
    D2=17
   D3=1
    D4=32
   D5=1
    D6=19
   D7=64
   D8=1
    SaveValue "circa33","D00",D0
    SaveValue "circa33","D01",D1
    SaveValue "circa33","D02",D2
    SaveValue "circa33","D03",D3
    SaveValue "circa33","D04",D4
    SaveValue "circa33","D05",D5
    SaveValue "circa33","D06",D6
    SaveValue "circa33","D07",D7
    SaveValue "circa33","D08",D8
  Else
    D0=LoadValue("circa33","D00")
   D1=LoadValue("circa33","D01")
   D2=LoadValue("circa33","D02")
   D3=LoadValue("circa33","D03")
   D4=LoadValue("circa33","D04")
   D5=LoadValue("circa33","D05")
   D6=LoadValue("circa33","D06")
   D7=LoadValue("circa33","D07")
   D8=LoadValue("circa33","D08")
 End If
  FigureDips 'Convert dips to usable settings
 On Error Resume Next
    With Controller
   .GameName=cGameName
   If Err Then MsgBox"Can't start Game "&cGameName&vbNewLine&Err.Description:Exit Sub
    .SplashInfoLine=cCredits
    .HandleMechanics=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .Hidden=1
   .Run
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
'  Controller.Dip(0)=(0*1 + 0*2 + 0*4 + 1*8 + 1*16 + 1*32 + 0*64 + 0*128)'2 players/replay/disable match/set 3 balls
'  Controller.Dip(1)=(1*1 + 0*2 + 0*4 + 1*8 + 1*16 + 0*32 + 0*64 + 1*128)'set max 99 credits
'  Controller.Dip(2)=(1*1 + 0*2 + 0*4 + 0*8 + 1*16 + 0*32 + 0*64 + 0*128)'coin1/'coin2 - 1 coin/1 credit
'  Controller.Dip(3)=(1*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128)'coin3 - 1 coin/1 credit  -- Credit Option 1=0
' Controller.Dip(4)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 1*32 + 0*64 + 0*128)'replay 1 120,000 (3 dipsets)
'  Controller.Dip(5)=(1*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128)
' Controller.Dip(6)=(1*1 + 1*2 + 0*4 + 0*8 + 1*16 + 0*32 + 0*64 + 0*128)'replay 2 130,000 (3 dipsets)
'  Controller.Dip(7)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 1*64 + 0*128)
' Controller.Dip(8)=(1*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128)'replay 3 140,000 (3 dipsets)

  SetActualDips

 PinMAMETimer.Interval=PinMAMEInterval
 PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch=31
  vpmNudge.Sensitivity=3
  vpmNudge.TiltObj=Array(Sling4,RightSlingshot,Bumper1,Bumper2,Bumper3)

 Set bsTrough=New cvpmBallStack
  bsTrough.InitNoTrough BallRelease,8,88,6
  bsTrough.InitExitSnd "BallRel","SolOn"


  Set dtDrop=New cvpmDropTarget
 dtDrop.InitDrop Array(Array(DT1,DT1A),Array(DT2,DT2A),Array(DT3,DT3A),Array(DT4,DT4A)),Array(25,26,27,28)
 dtDrop.InitSnd "FlapClos","FlapOpen"

End Sub

Sub SetActualDips
  D0=Setting(1)+Setting(2)*2+Setting(3)*4+Setting(4)*8+Setting(5)*16+Setting(6)*32+Setting(7)*64+Setting(8)*128
 D1=Setting(9)+Setting(10)*2+Setting(11)*4+Setting(12)*8+Setting(13)*16+Setting(14)*32+Setting(15)*64+Setting(16)*128
  D2=Setting(17)+Setting(18)*2+Setting(19)*4+Setting(20)*8+Setting(21)*16+Setting(22)*32+Setting(23)*64+Setting(24)*128
 D3=Setting(25)+Setting(26)*2+Setting(27)*4+Setting(28)*8+Setting(29)*16+Setting(30)*32+Setting(31)*64+Setting(32)*128
 D4=Setting(33)+Setting(34)*2+Setting(35)*4+Setting(36)*8+Setting(37)*16+Setting(38)*32+Setting(39)*64+Setting(40)*128
 D5=Setting(41)+Setting(42)*2+Setting(43)*4+Setting(44)*8+Setting(45)*16+Setting(46)*32+Setting(47)*64+Setting(48)*128
 D6=Setting(49)+Setting(50)*2+Setting(51)*4+Setting(52)*8+Setting(53)*16+Setting(54)*32+Setting(55)*64+Setting(56)*128
 D7=Setting(57)+Setting(58)*2+Setting(59)*4+Setting(60)*8+Setting(61)*16+Setting(62)*32+Setting(63)*64+Setting(64)*128
 D8=Setting(65)+Setting(66)*2+Setting(67)*4+Setting(68)*8+Setting(69)*16+Setting(70)*32+Setting(71)*64+Setting(72)*128
 Controller.Dip(0)=D0
  Controller.Dip(1)=D1
  Controller.Dip(2)=D2
  Controller.Dip(3)=D3
  Controller.Dip(4)=D4
  Controller.Dip(5)=D5
  Controller.Dip(6)=D6
  Controller.Dip(7)=D7
  Controller.Dip(8)=D8
End Sub

Sub FigureDips
'*Bank0
  If D0>=128 Then
   Setting(8)=1
    D0=D0-128
 Else
    Setting(8)=0
  End If
  If D0>=64 Then
    Setting(7)=1
    D0=D0-64
  Else
    Setting(7)=0
  End If
  If D0>=32 Then
    Setting(6)=1
    D0=D0-32
  Else
    Setting(6)=0
  End If
  If D0>=16 Then
    Setting(5)=1
    D0=D0-16
  Else
    Setting(5)=0
  End If
  If D0>=8 Then
   Setting(4)=1
    D0=D0-8
 Else
    Setting(4)=0
  End If
  If D0>=4 Then
   Setting(3)=1
    D0=D0-4
 Else
    Setting(3)=0
  End If
  If D0>=2 Then
   Setting(2)=1
    D0=D0-2
 Else
    Setting(2)=0
  End If
  If D0>=1 Then
   Setting(1)=1
    D0=D0-1
 Else
    Setting(1)=0
  End If

'*Bank1
 If D1>=128 Then
   Setting(16)=1
   D1=D1-128
 Else
    Setting(16)=0
 End If
  If D1>=64 Then
    Setting(15)=1
   D1=D1-64
  Else
    Setting(15)=0
 End If
  If D1>=32 Then
    Setting(14)=1
   D1=D1-32
  Else
    Setting(14)=0
 End If
  If D1>=16 Then
    Setting(13)=1
   D1=D1-16
  Else
    Setting(13)=0
 End If
  If D1>=8 Then
   Setting(12)=1
   D1=D1-8
 Else
    Setting(12)=0
 End If
  If D1>=4 Then
   Setting(11)=1
   D1=D1-4
 Else
    Setting(11)=0
 End If
  If D1>=2 Then
   Setting(10)=1
   D1=D1-2
 Else
    Setting(10)=0
 End If
  If D1>=1 Then
   Setting(9)=1
    D1=D1-1
 Else
    Setting(9)=0
  End If

'*Bank2
 If D2>=128 Then
   Setting(24)=1
   D2=D2-128
 Else
    Setting(24)=0
 End If
  If D2>=64 Then
    Setting(23)=1
   D2=D2-64
  Else
    Setting(23)=0
 End If
  If D2>=32 Then
    Setting(22)=1
   D2=D2-32
  Else
    Setting(22)=0
 End If
  If D2>=16 Then
    Setting(21)=1
   D2=D2-16
  Else
    Setting(21)=0
 End If
  If D2>=8 Then
   Setting(20)=1
   D2=D2-8
 Else
    Setting(20)=0
 End If
  If D2>=4 Then
   Setting(19)=1
   D2=D2-4
 Else
    Setting(19)=0
 End If
  If D2>=2 Then
   Setting(18)=1
   D2=D2-2
 Else
    Setting(18)=0
 End If
  If D2>=1 Then
   Setting(17)=1
   D2=D2-1
 Else
    Setting(17)=0
 End If

'*Bank3
 If D3>=128 Then
   Setting(32)=1
   D3=D3-128
 Else
    Setting(32)=0
 End If
  If D3>=64 Then
    Setting(31)=1
   D3=D3-64
  Else
    Setting(31)=0
 End If
  If D3>=32 Then
    Setting(30)=1
   D3=D3-32
  Else
    Setting(30)=0
 End If
  If D3>=16 Then
    Setting(29)=1
   D3=D3-16
  Else
    Setting(29)=0
 End If
  If D3>=8 Then
   Setting(28)=1
   D3=D3-8
 Else
    Setting(28)=0
 End If
  If D3>=4 Then
   Setting(27)=1
   D3=D3-4
 Else
    Setting(27)=0
 End If
  If D3>=2 Then
   Setting(26)=1
   D3=D3-2
 Else
    Setting(26)=0
 End If
  If D3>=1 Then
   Setting(25)=1
   D3=D3-1
 Else
    Setting(25)=0
 End If

'*Bank4
 If D4>=128 Then
   Setting(40)=1
   D4=D4-128
 Else
    Setting(40)=0
 End If
  If D4>=64 Then
    Setting(39)=1
   D4=D4-64
  Else
    Setting(39)=0
 End If
  If D4>=32 Then
    Setting(38)=1
   D4=D4-32
  Else
    Setting(38)=0
 End If
  If D4>=16 Then
    Setting(37)=1
   D4=D4-16
  Else
    Setting(37)=0
 End If
  If D4>=8 Then
   Setting(36)=1
   D4=D4-8
 Else
    Setting(36)=0
 End If
  If D4>=4 Then
   Setting(35)=1
   D4=D4-4
 Else
    Setting(35)=0
 End If
  If D4>=2 Then
   Setting(34)=1
   D4=D4-2
 Else
    Setting(34)=0
 End If
  If D4>=1 Then
   Setting(33)=1
   D4=D4-1
 Else
    Setting(33)=0
 End If

'*Bank5
 If D5>=128 Then
   Setting(48)=1
   D5=D5-128
 Else
    Setting(48)=0
 End If
  If D5>=64 Then
    Setting(47)=1
   D5=D5-64
  Else
    Setting(47)=0
 End If
  If D5>=32 Then
    Setting(46)=1
   D5=D5-32
  Else
    Setting(46)=0
 End If
  If D5>=16 Then
    Setting(45)=1
   D5=D5-16
  Else
    Setting(45)=0
 End If
  If D5>=8 Then
   Setting(44)=1
   D5=D5-8
 Else
    Setting(44)=0
 End If
  If D5>=4 Then
   Setting(43)=1
   D5=D5-4
 Else
    Setting(43)=0
 End If
  If D5>=2 Then
   Setting(42)=1
   D5=D5-2
 Else
    Setting(42)=0
 End If
  If D5>=1 Then
   Setting(41)=1
   D5=D5-1
 Else
    Setting(41)=0
 End If

'*Bank6
 If D6>=128 Then
   Setting(56)=1
   D6=D6-128
 Else
    Setting(56)=0
 End If
  If D6>=64 Then
    Setting(55)=1
   D6=D6-64
  Else
    Setting(55)=0
 End If
  If D6>=32 Then
    Setting(54)=1
   D6=D6-32
  Else
    Setting(54)=0
 End If
  If D6>=16 Then
    Setting(53)=1
   D6=D6-16
  Else
    Setting(53)=0
 End If
  If D6>=8 Then
   Setting(52)=1
   D6=D6-8
 Else
    Setting(52)=0
 End If
  If D6>=4 Then
   Setting(51)=1
   D6=D6-4
 Else
    Setting(51)=0
 End If
  If D6>=2 Then
   Setting(50)=1
   D6=D6-2
 Else
    Setting(50)=0
 End If
  If D6>=1 Then
   Setting(49)=1
   D6=D6-1
 Else
    Setting(49)=0
 End If

'*Bank7
 If D7>=128 Then
   Setting(64)=1
   D7=D7-128
 Else
    Setting(64)=0
 End If
  If D7>=64 Then
    Setting(63)=1
   D7=D7-64
  Else
    Setting(63)=0
 End If
  If D7>=32 Then
    Setting(62)=1
   D7=D7-32
  Else
    Setting(62)=0
 End If
  If D7>=16 Then
    Setting(61)=1
   D7=D7-16
  Else
    Setting(61)=0
 End If
  If D7>=8 Then
   Setting(60)=1
   D7=D7-8
 Else
    Setting(60)=0
 End If
  If D7>=4 Then
   Setting(59)=1
   D7=D7-4
 Else
    Setting(59)=0
 End If
  If D7>=2 Then
   Setting(58)=1
   D7=D7-2
 Else
    Setting(58)=0
 End If
  If D7>=1 Then
   Setting(57)=1
   D7=D7-1
 Else
    Setting(57)=0
 End If

'*Bank8
 If D8>=128 Then
   Setting(72)=1
   D8=D8-128
 Else
    Setting(72)=0
 End If
  If D8>=64 Then
    Setting(71)=1
   D8=D8-64
  Else
    Setting(71)=0
 End If
  If D8>=32 Then
    Setting(70)=1
   D8=D8-32
  Else
    Setting(70)=0
 End If
  If D8>=16 Then
    Setting(69)=1
   D8=D8-16
  Else
    Setting(69)=0
 End If
  If D8>=8 Then
   Setting(68)=1
   D8=D8-8
 Else
    Setting(68)=0
 End If
  If D8>=4 Then
   Setting(67)=1
   D8=D8-4
 Else
    Setting(67)=0
 End If
  If D8>=2 Then
   Setting(66)=1
   D8=D8-2
 Else
    Setting(66)=0
 End If
  If D8>=1 Then
   Setting(65)=1
   D8=D8-1
 Else
    Setting(65)=0
 End If
End Sub

Dim DipsShown,DipsPage,CurSelect,SelectMax(18)
CurSelect=0
DipsShown=False
SelectMax(0)=1
SelectMax(1)=1
SelectMax(2)=1
SelectMax(3)=4
SelectMax(4)=9
SelectMax(5)=9
SelectMax(6)=9
SelectMax(7)=9
SelectMax(8)=9
SelectMax(9)=9
SelectMax(10)=9
SelectMax(11)=9
SelectMax(12)=9
SelectMax(13)=9
SelectMax(14)=9
SelectMax(15)=9
SelectMax(16)=9
SelectMax(17)=9
SelectMax(18)=9

Sub PlungerTimer_Timer
  PlungerR.Y = 1465 + (5* Plunger.Position) -20
End Sub

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode=keyShowDips Then
   ToggleDips
    Exit Sub
  End If
  If DipsShown=False Then
     If KeyCode=PlungerKey Then Plunger.PullBack
     If vpmKeyDown(KeyCode) Then
       If KeyCode=keyReset Then
          Controller.Hidden=1
         SetActualDips
         Exit Sub
        End If
      End If
  Else
    If KeyCode=keyJoyLeft Then
      StoreSetting(DipsPage)
      If DipsPage>0 Then
        DipsPage=DipsPage-1
     Else
        DipsPage=18
     End If
      Display(DipsPage)
   End If
    If KeyCode=keyJoyRight Then
     StoreSetting(DipsPage)
      If DipsPage<18 Then
       DipsPage=DipsPage+1
     Else
        DipsPage=0
      End If
      Display(DipsPage)
   End If
    If KeyCode=keyJoyUp Then
      StoreSetting(DipsPage)
      If CurSelect>0 Then
       CurSelect=CurSelect-1
       StoreSetting(DipsPage)
      End If
      Display(DipsPage)
   End If
    If KeyCode=keyJoyDown Then
      StoreSetting(DipsPage)
      If CurSelect<SelectMax(DipsPage) Then
       CurSelect=CurSelect+1
       StoreSetting(DipsPage)
      End If
      Display(DipsPage)
   End If
  End If
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode=keyShowDips Then Exit Sub
  If DipsShown=False Then
     If KeyCode=PlungerKey Then:Plunger.Fire:PlaySoundAtVol "Plunger", Plunger, 1:End If
     If vpmKeyUp(KeyCode) Then Exit Sub
  End If
End Sub

Sub ToggleDips
 LeftFlipper.RotateToStart
 RightFlipper.RotateToStart
  Plunger.Fire
  If DipsShown=False Then
   Controller.Pause=True
   EMReel1.SetValue 1
    If Setting(2)=0 Then EMReel2.SetValue 18
    If Setting(2)=1 Then EMReel2.SetValue 19
    CurSelect=Setting(2)
    DipsPage=0
    DipsShown=True
  Else
    StoreSetting(DipsPage)
    EMReel1.SetValue 0
    EMReel2.SetValue 0
    Controller.Pause=False
    SetActualDips
   DipsShown=False
 End If
End Sub

Sub Table1_Exit
  SetActualDips
 SaveValue "circa33","D00",D0
  SaveValue "circa33","D01",D1
  SaveValue "circa33","D02",D2
  SaveValue "circa33","D03",D3
  SaveValue "circa33","D04",D4
  SaveValue "circa33","D05",D5
  SaveValue "circa33","D06",D6
  SaveValue "circa33","D07",D7
  SaveValue "circa33","D08",D8
End Sub

'*************************  SWITCHES  *************************
Sub Bumper1_Hit:vpmTimer.PulseSw 1:End Sub                                        '10
Sub Bumper2_Hit:vpmTimer.PulseSw 2:End Sub                                       '20
Sub Bumper3_Hit:vpmTimer.PulseSw 3:End Sub                                       '30
Sub T1_Hit:T1.TimerEnabled=1:vpmTimer.PulseSw 4:T1.IsDropped=1:T1a.IsDropped=0:End Sub                 '40
Sub T1_Timer:T1.TimerEnabled=0:T1a.IsDropped=1:T1.IsDropped=0:End Sub
Sub T2_Hit:T2.TimerEnabled=1:vpmTimer.PulseSw 5:T2.IsDropped=1:T2a.IsDropped=0:End Sub                  '50
Sub T2_Timer:T2.TimerEnabled=0:T2a.IsDropped=1:T2.IsDropped=0:End Sub
Sub Sling4_Slingshot:vpmTimer.PulseSw 6:End Sub                                     '60
Sub RightSlingshot_Slingshot:vpmTimer.PulseSw 7:End Sub                                  '70
Sub Drain_Hit:PlaySoundAtVol "ALBallDrain", Drain, 1:bsTrough.AddBall Me:End Sub                                       '80
Sub T3_Hit:T3.TimerEnabled=1:vpmTimer.PulseSw 14:T3.IsDropped=1:T3a.IsDropped=0:End Sub                  '140
Sub T3_Timer:T3.TimerEnabled=0:T3a.IsDropped=1:T3.IsDropped=0:End Sub
Sub Sling1_Hit:vpmTimer.PulseSw 15:End Sub                                       '150
Sub Sling2_Hit:vpmTimer.PulseSw 15:End Sub
Sub Sling3_Hit:vpmTimer.PulseSw 15:End Sub
Sub LeftSlingshot_Hit:vpmTimer.PulseSw 15:End Sub
Sub Trigger5_Hit:Controller.Switch(16)=1:Trigger5.TimerEnabled=0:Roll5.IsDropped=1::End Sub                '160
Sub Trigger5_unHit:Controller.Switch(16)=0:Trigger5.TimerEnabled=1:End Sub
Sub Trigger5_Timer:Trigger5.TimerEnabled=0:Roll5.IsDropped=0:End Sub
Sub Trigger8_Hit:Controller.Switch(17)=1:Trigger8.TimerEnabled=0:Roll8.IsDropped=1::DT1a.isdropped=0:PlaysoundAtVol "Targetup", ActiveBall, 1:End Sub               '170
Sub Trigger8_unHit:Controller.Switch(17)=0:Trigger8.TimerEnabled=1:End Sub
Sub Trigger8_Timer:Trigger8.TimerEnabled=0:Roll8.IsDropped=0:End Sub
Sub Trigger9_Hit:Controller.Switch(18)=1:Trigger9.TimerEnabled=0:Roll9.IsDropped=1:DT4a.isdropped=0:PlaysoundAtVol "Targetup", ActiveBall, 1:End Sub                '180
Sub Trigger9_unHit:Controller.Switch(18)=0:Trigger9.TimerEnabled=1:End Sub
Sub Trigger9_Timer:Trigger9.TimerEnabled=0:Roll9.IsDropped=0:End Sub
Sub Trigger4_Hit:Controller.Switch(21)=1:Trigger4.TimerEnabled=0:Roll4.IsDropped=1:End Sub                '210
Sub Trigger4_unHit:Controller.Switch(21)=0:Trigger4.TimerEnabled=1:End Sub
Sub Trigger4_Timer:Trigger4.TimerEnabled=0:Roll4.IsDropped=0:End Sub
Sub Trigger6_Hit:Controller.Switch(21)=1:Trigger6.TimerEnabled=0:Roll6.IsDropped=1:End Sub
Sub Trigger6_unHit:Controller.Switch(21)=0:Trigger6.TimerEnabled=1:End Sub
Sub Trigger6_Timer:Trigger6.TimerEnabled=0:Roll6.IsDropped=0:End Sub
Sub Trigger7_Hit:Controller.Switch(21)=1:Trigger7.TimerEnabled=0:Roll7.IsDropped=1:End Sub
Sub Trigger7_unHit:Controller.Switch(21)=0:Trigger7.TimerEnabled=1:End Sub
Sub Trigger7_Timer:Trigger7.TimerEnabled=0:Roll7.IsDropped=0:End Sub
Sub Trigger1_Hit:Controller.Switch(22)=1:Trigger1.TimerEnabled=0:Roll1.IsDropped=1::DT2a.isdropped=0:PlaysoundAtVol "Targetup", ActiveBall, 1:End Sub               '220
Sub Trigger1_unHit:Controller.Switch(22)=0:Trigger1.TimerEnabled=1:End Sub
Sub Trigger1_Timer:Trigger1.TimerEnabled=0:Roll1.IsDropped=0:End Sub
Sub Trigger3_Hit:Controller.Switch(23)=1:Trigger3.TimerEnabled=0:Roll3.IsDropped=1::DT3a.isdropped=0:PlaysoundAtVol "Targetup", ActiveBall, 1:End Sub               '230
Sub Trigger3_unHit:Controller.Switch(23)=0:Trigger3.TimerEnabled=1:End Sub
Sub Trigger3_Timer:Trigger3.TimerEnabled=0:Roll3.IsDropped=0:End Sub
Sub Trigger2_Hit:Controller.Switch(24)=1:Trigger2.TimerEnabled=0:Roll2.IsDropped=1:End Sub                '240
Sub Trigger2_unHit:Controller.Switch(24)=0:Trigger2.TimerEnabled=1:End Sub
Sub Trigger2_Timer:Trigger2.TimerEnabled=0:Roll2.IsDropped=0:End Sub

Sub DT1a_dropped
dtDrop.Hit 1
PlaysoundAtVol "Targetdown", ActiveBall, 1
DTALL
End Sub
                                          '250
Sub DT2a_dropped
dtDrop.Hit 2
PlaysoundAtVol "Targetdown", ActiveBall, 1
DTALL
End Sub
                                          '260
Sub DT3a_dropped
PlaysoundAtVol "Targetdown", ActiveBall, 1
dtDrop.Hit 3
DTALL
End Sub
                                          '270
Sub DT4a_dropped
dtDrop.Hit 4
PlaysoundAtVol "Targetdown", ActiveBall, 1
DTALL
End Sub

Sub DTALL() If DT1a.isdropped=1 and DT2a.isdropped=1 and DT3a.isdropped=1 and DT4a.isdropped=1 Then                                       '280
DT1a.isdropped=0 :DT2a.isdropped=0 :DT3a.isdropped=0 :DT4a.isdropped=0
End If
End Sub

Sub Gate1_Hit
  PlaysoundAtVol "Targetreset", ActiveBall, 1                                   '310 TILT
DT1a.isdropped=0 :DT2a.isdropped=0 :DT3a.isdropped=0 :DT4a.isdropped=0                                         '320 START GAME
End Sub                                          '330 SLAM TILT                                                            '340 COIN1
                                            '350 COIN                                                           '360 COIN3
'*************************  SEGMENTED DISPLAYS  *************************
Dim Digits(25)
Digits(0)=Array(L1,L2,L3,L4,L5,L6,L7)
Digits(1)=Array(L8,L9,L10,L11,L12,L13,L14)
Digits(2)=Array(L15,L16,L17,L18,L19,L20,L21)
Digits(3)=Array(L22,L23,L24,L25,L26,L27,L28)
Digits(4)=Array(L29,L30,L31,L32,L33,L34,L35)
Digits(5)=Array(L36,L37,L38,L39,L40,L41,L42)

Digits(6)=Array(L43,L44,L45,L46,L47,L48,L49)
Digits(7)=Array(L50,L51,L52,L53,L54,L55,L56)
Digits(8)=Array(L57,L58,L59,L60,L61,L62,L63)
Digits(9)=Array(L64,L65,L66,L67,L68,L69,L70)
Digits(10)=Array(L71,L72,L73,L74,L75,L76,L77)
Digits(11)=Array(L78,L79,L80,L81,L82,L83,L84)

'Player3 Score displays use 12/13/14/15/16/17
Digits(12)=Array(L99,L100,L101,L102,L103,L104,L105)
Digits(13)=Array(L106,L107,L108,L109,L110,L111,L112)
Digits(14)=Array(L113,L114,L115,L116,L117,L118,L119)
Digits(15)=Array(L120,L121,L122,L123,L124,L125,L126)
Digits(16)=Array(L127,L128,L129,L130,L131,L132,L133)
Digits(17)=Array(L134,L135,L136,L137,L138,L139,L140)
'Player4 Score displays use 18/19/20/21/22/23
Digits(18)=Array(L141,L142,L143,L144,L145,L146,L147)
Digits(19)=Array(L148,L149,L150,L151,L152,L153,L154)
Digits(20)=Array(L155,L156,L157,L158,L159,L160,L161)
Digits(21)=Array(L162,L163,L164,L165,L166,L167,L168)
Digits(22)=Array(L169,L170,L171,L172,L173,L174,L175)
Digits(23)=Array(L176,L177,L178,L179,L180,L181,L182)

Digits(24)=Array(L85,L86,L87,L88,L89,L90,L91)
Digits(25)=Array(L92,L93,L94,L95,L96,L97,L98)

Sub DisplayTimer_Timer
 Dim ChgLED,ii,num,chg,stat,obj
  ChgLED=Controller.ChangedLEDs(&Hffffffff,&Hffffffff)
  If Not IsEmpty(ChgLED) Then
   For ii=0 To UBound(chgLED)
      num=chgLED(ii,0):chg=chgLED(ii,1):stat=chgLED(ii,2)
     If Num<12 Or Num>23 Then
        For Each obj In Digits(num)
         If chg And 1 Then obj.State=stat And 1
          chg=chg\2:stat=stat\2
       Next
      End If
    Next
  End If
End Sub

'*************************   LAMPS  *************************
'2=BONUS 1
'3=BONUS 2
'4=BONUS 3
'5=BONUS 4
'6=BONUS 5
'7=BONUS 6
'8=BONUS 7
'9=BONUS 8
'10=BONUS 9
'11=BONUS 10
'12=TILT
'14=GAME OVER
'16=Slingshot Illumination
'17=Left 1,000 pts
'18=Left 2,000 pts
'19=Left 3,000 pts
'20=Left 4,000 pts
'23=Double bonus
'24=Triple bonus
'25=Right 1,000 pts
'26=Right 2,000 pts
'27=Right 3,0000 pts
'28=Right 4,000 pts
'29=Special when lit
'31=Extra ball when lit
'33=Ball 1
'34=Ball 2
'35=Ball 3
'36=Ball 4
'37=Ball 5
'38=Same player shoots again
'41=Player1 Game
'42=Player2 Game
'43=Player3 Game
'44=Player4 Game

Set Lights(12)=Lamp12
Set Lights(14)=Lamp14
Set Lights(33)=Lamp33
Set Lights(34)=Lamp34
Set Lights(35)=Lamp35
Set Lights(36)=Lamp36
Set Lights(37)=Lamp37
Set Lights(41)=Lamp41
Set Lights(42)=Lamp42

Set LampCallback=GetRef("UpdateMultipleLamps")
Sub UpdateMultipleLamps
  WL2.IsDropped=Controller.Lamp(2)
  WL3.IsDropped=Controller.Lamp(3)
  WL4.IsDropped=Controller.Lamp(4)
  WL5.IsDropped=Controller.Lamp(5)
  WL6.IsDropped=Controller.Lamp(6)
  WL7.IsDropped=Controller.Lamp(7)
  WL8.IsDropped=Controller.Lamp(8)
  WL9.IsDropped=Controller.Lamp(9)
  WL10.IsDropped=Controller.Lamp(10)
  WL11.IsDropped=Controller.Lamp(11)
  Light1.state= Controller.Lamp(16)
  Light2.state= Controller.Lamp(16)
  Light3.state= Controller.Lamp(16)
  Light4.state= Controller.Lamp(16)
  Light5.state= Controller.Lamp(16)
  Light6.state= Controller.Lamp(16)
  Light7.state= Controller.Lamp(16)
  Light8.state= Controller.Lamp(16)
  Light9.state= Controller.Lamp(16)
  Light10.state= Controller.Lamp(16)
  WL17.IsDropped=Controller.Lamp(17)
  WL18.IsDropped=Controller.Lamp(18)
  WL19.IsDropped=Controller.Lamp(19)
  WL20.IsDropped=Controller.Lamp(20)
  WL23.IsDropped=Controller.Lamp(23)
  WL24.IsDropped=Controller.Lamp(24)
  WL25.IsDropped=Controller.Lamp(25)
  WL26.IsDropped=Controller.Lamp(26)
  WL27.IsDropped=Controller.Lamp(27)
  WL28.IsDropped=Controller.Lamp(28)
  WL29.IsDropped=Controller.Lamp(29)
  WL31.IsDropped=Controller.Lamp(31)
  WL38.IsDropped=Controller.Lamp(38)
End Sub

Sub Display(DP)
  EMReel1.SetValue DP+1
 Select Case DP
    Case 0:CurSelect=Setting(2)
       If CurSelect=0 Then EMReel2.SetValue 18
       If CurSelect=1 Then EMReel2.SetValue 19
   Case 1:CurSelect=Setting(3)
       If CurSelect=0 Then EMReel2.SetValue 16
       If CurSelect=1 Then EMReel2.SetValue 17
   Case 2:CurSelect=Setting(4)
       If CurSelect=0 Then EMReel2.SetValue 16
       If CurSelect=1 Then EMReel2.SetValue 17
   Case 3:If Setting(5)=1 And Setting(6)=0 And Setting(7)=0 Then CurSelect=0
       If Setting(5)=0 And Setting(6)=1 And Setting(7)=0 Then CurSelect=1
        If Setting(5)=1 And Setting(6)=1 And Setting(7)=0 Then CurSelect=2
        If Setting(5)=0 And Setting(6)=0 And Setting(7)=1 Then CurSelect=3
        If Setting(5)=1 And Setting(6)=0 And Setting(7)=1 Then CurSelect=4
        EMReel2.SetValue 11+CurSelect
   Case 4:If Setting(9)=0 And Setting(10)=0 And Setting(11)=0 And Setting(12)=0 Then CurSelect=0
       If Setting(9)=1 And Setting(10)=0 And Setting(11)=0 And Setting(12)=0 Then CurSelect=1
        If Setting(9)=0 And Setting(10)=1 And Setting(11)=0 And Setting(12)=0 Then CurSelect=2
        If Setting(9)=1 And Setting(10)=1 And Setting(11)=0 And Setting(12)=0 Then CurSelect=3
        If Setting(9)=0 And Setting(10)=0 And Setting(11)=1 And Setting(12)=0 Then CurSelect=4
        If Setting(9)=1 And Setting(10)=0 And Setting(11)=1 And Setting(12)=0 Then CurSelect=5
        If Setting(9)=0 And Setting(10)=1 And Setting(11)=1 And Setting(12)=0 Then CurSelect=6
        If Setting(9)=1 And Setting(10)=1 And Setting(11)=1 And Setting(12)=0 Then CurSelect=7
        If Setting(9)=0 And Setting(10)=0 And Setting(11)=0 And Setting(12)=1 Then CurSelect=8
        If Setting(9)=1 And Setting(10)=0 And Setting(11)=0 And Setting(12)=1 Then CurSelect=9
        EMReel2.SetValue 1+CurSelect
    Case 5:If Setting(13)=0 And Setting(14)=0 And Setting(15)=0 And Setting(16)=0 Then CurSelect=0
        If Setting(13)=1 And Setting(14)=0 And Setting(15)=0 And Setting(16)=0 Then CurSelect=1
       If Setting(13)=0 And Setting(14)=1 And Setting(15)=0 And Setting(16)=0 Then CurSelect=2
       If Setting(13)=1 And Setting(14)=1 And Setting(15)=0 And Setting(16)=0 Then CurSelect=3
       If Setting(13)=0 And Setting(14)=0 And Setting(15)=1 And Setting(16)=0 Then CurSelect=4
       If Setting(13)=1 And Setting(14)=0 And Setting(15)=1 And Setting(16)=0 Then CurSelect=5
       If Setting(13)=0 And Setting(14)=1 And Setting(15)=1 And Setting(16)=0 Then CurSelect=6
       If Setting(13)=1 And Setting(14)=1 And Setting(15)=1 And Setting(16)=0 Then CurSelect=7
       If Setting(13)=0 And Setting(14)=0 And Setting(15)=0 And Setting(16)=1 Then CurSelect=8
       If Setting(13)=1 And Setting(14)=0 And Setting(15)=0 And Setting(16)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 6:If Setting(17)=0 And Setting(18)=0 And Setting(19)=0 And Setting(20)=0 Then CurSelect=0
        If Setting(17)=1 And Setting(18)=0 And Setting(19)=0 And Setting(20)=0 Then CurSelect=1
       If Setting(17)=0 And Setting(18)=1 And Setting(19)=0 And Setting(20)=0 Then CurSelect=2
       If Setting(17)=1 And Setting(18)=1 And Setting(19)=0 And Setting(20)=0 Then CurSelect=3
       If Setting(17)=0 And Setting(18)=0 And Setting(19)=1 And Setting(20)=0 Then CurSelect=4
       If Setting(17)=1 And Setting(18)=0 And Setting(19)=1 And Setting(20)=0 Then CurSelect=5
       If Setting(17)=0 And Setting(18)=1 And Setting(19)=1 And Setting(20)=0 Then CurSelect=6
       If Setting(17)=1 And Setting(18)=1 And Setting(19)=1 And Setting(20)=0 Then CurSelect=7
       If Setting(17)=0 And Setting(18)=0 And Setting(19)=0 And Setting(20)=1 Then CurSelect=8
       If Setting(17)=1 And Setting(18)=0 And Setting(19)=0 And Setting(20)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 7:If Setting(21)=0 And Setting(22)=0 And Setting(23)=0 And Setting(24)=0 Then CurSelect=0
        If Setting(21)=1 And Setting(22)=0 And Setting(23)=0 And Setting(24)=0 Then CurSelect=1
       If Setting(21)=0 And Setting(22)=1 And Setting(23)=0 And Setting(24)=0 Then CurSelect=2
       If Setting(21)=1 And Setting(22)=1 And Setting(23)=0 And Setting(24)=0 Then CurSelect=3
       If Setting(21)=0 And Setting(22)=0 And Setting(23)=1 And Setting(24)=0 Then CurSelect=4
       If Setting(21)=1 And Setting(22)=0 And Setting(23)=1 And Setting(24)=0 Then CurSelect=5
       If Setting(21)=0 And Setting(22)=1 And Setting(23)=1 And Setting(24)=0 Then CurSelect=6
       If Setting(21)=1 And Setting(22)=1 And Setting(23)=1 And Setting(24)=0 Then CurSelect=7
       If Setting(21)=0 And Setting(22)=0 And Setting(23)=0 And Setting(24)=1 Then CurSelect=8
       If Setting(21)=1 And Setting(22)=0 And Setting(23)=0 And Setting(24)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 8:If Setting(25)=0 And Setting(26)=0 And Setting(27)=0 And Setting(28)=0 Then CurSelect=0
        If Setting(25)=1 And Setting(26)=0 And Setting(27)=0 And Setting(28)=0 Then CurSelect=1
       If Setting(25)=0 And Setting(26)=1 And Setting(27)=0 And Setting(28)=0 Then CurSelect=2
       If Setting(25)=1 And Setting(26)=1 And Setting(27)=0 And Setting(28)=0 Then CurSelect=3
       If Setting(25)=0 And Setting(26)=0 And Setting(27)=1 And Setting(28)=0 Then CurSelect=4
       If Setting(25)=1 And Setting(26)=0 And Setting(27)=1 And Setting(28)=0 Then CurSelect=5
       If Setting(25)=0 And Setting(26)=1 And Setting(27)=1 And Setting(28)=0 Then CurSelect=6
       If Setting(25)=1 And Setting(26)=1 And Setting(27)=1 And Setting(28)=0 Then CurSelect=7
       If Setting(25)=0 And Setting(26)=0 And Setting(27)=0 And Setting(28)=1 Then CurSelect=8
       If Setting(25)=1 And Setting(26)=0 And Setting(27)=0 And Setting(28)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 9:If Setting(29)=0 And Setting(30)=0 And Setting(31)=0 And Setting(32)=0 Then CurSelect=0
        If Setting(29)=1 And Setting(30)=0 And Setting(31)=0 And Setting(32)=0 Then CurSelect=1
       If Setting(29)=0 And Setting(30)=1 And Setting(31)=0 And Setting(32)=0 Then CurSelect=2
       If Setting(29)=1 And Setting(30)=1 And Setting(31)=0 And Setting(32)=0 Then CurSelect=3
       If Setting(29)=0 And Setting(30)=0 And Setting(31)=1 And Setting(32)=0 Then CurSelect=4
       If Setting(29)=1 And Setting(30)=0 And Setting(31)=1 And Setting(32)=0 Then CurSelect=5
       If Setting(29)=0 And Setting(30)=1 And Setting(31)=1 And Setting(32)=0 Then CurSelect=6
       If Setting(29)=1 And Setting(30)=1 And Setting(31)=1 And Setting(32)=0 Then CurSelect=7
       If Setting(29)=0 And Setting(30)=0 And Setting(31)=0 And Setting(32)=1 Then CurSelect=8
       If Setting(29)=1 And Setting(30)=0 And Setting(31)=0 And Setting(32)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 10:If Setting(33)=0 And Setting(34)=0 And Setting(35)=0 And Setting(36)=0 Then CurSelect=0
       If Setting(33)=1 And Setting(34)=0 And Setting(35)=0 And Setting(36)=0 Then CurSelect=1
       If Setting(33)=0 And Setting(34)=1 And Setting(35)=0 And Setting(36)=0 Then CurSelect=2
       If Setting(33)=1 And Setting(34)=1 And Setting(35)=0 And Setting(36)=0 Then CurSelect=3
       If Setting(33)=0 And Setting(34)=0 And Setting(35)=1 And Setting(36)=0 Then CurSelect=4
       If Setting(33)=1 And Setting(34)=0 And Setting(35)=1 And Setting(36)=0 Then CurSelect=5
       If Setting(33)=0 And Setting(34)=1 And Setting(35)=1 And Setting(36)=0 Then CurSelect=6
       If Setting(33)=1 And Setting(34)=1 And Setting(35)=1 And Setting(36)=0 Then CurSelect=7
       If Setting(33)=0 And Setting(34)=0 And Setting(35)=0 And Setting(36)=1 Then CurSelect=8
       If Setting(33)=1 And Setting(34)=0 And Setting(35)=0 And Setting(36)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 11:If Setting(37)=0 And Setting(38)=0 And Setting(39)=0 And Setting(40)=0 Then CurSelect=0
       If Setting(37)=1 And Setting(38)=0 And Setting(39)=0 And Setting(40)=0 Then CurSelect=1
       If Setting(37)=0 And Setting(38)=1 And Setting(39)=0 And Setting(40)=0 Then CurSelect=2
       If Setting(37)=1 And Setting(38)=1 And Setting(39)=0 And Setting(40)=0 Then CurSelect=3
       If Setting(37)=0 And Setting(38)=0 And Setting(39)=1 And Setting(40)=0 Then CurSelect=4
       If Setting(37)=1 And Setting(38)=0 And Setting(39)=1 And Setting(40)=0 Then CurSelect=5
       If Setting(37)=0 And Setting(38)=1 And Setting(39)=1 And Setting(40)=0 Then CurSelect=6
       If Setting(37)=1 And Setting(38)=1 And Setting(39)=1 And Setting(40)=0 Then CurSelect=7
       If Setting(37)=0 And Setting(38)=0 And Setting(39)=0 And Setting(40)=1 Then CurSelect=8
       If Setting(37)=1 And Setting(38)=0 And Setting(39)=0 And Setting(40)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 12:If Setting(41)=0 And Setting(42)=0 And Setting(43)=0 And Setting(44)=0 Then CurSelect=0
       If Setting(41)=1 And Setting(42)=0 And Setting(43)=0 And Setting(44)=0 Then CurSelect=1
       If Setting(41)=0 And Setting(42)=1 And Setting(43)=0 And Setting(44)=0 Then CurSelect=2
       If Setting(41)=1 And Setting(42)=1 And Setting(43)=0 And Setting(44)=0 Then CurSelect=3
       If Setting(41)=0 And Setting(42)=0 And Setting(43)=1 And Setting(44)=0 Then CurSelect=4
       If Setting(41)=1 And Setting(42)=0 And Setting(43)=1 And Setting(44)=0 Then CurSelect=5
       If Setting(41)=0 And Setting(42)=1 And Setting(43)=1 And Setting(44)=0 Then CurSelect=6
       If Setting(41)=1 And Setting(42)=1 And Setting(43)=1 And Setting(44)=0 Then CurSelect=7
       If Setting(41)=0 And Setting(42)=0 And Setting(43)=0 And Setting(44)=1 Then CurSelect=8
       If Setting(41)=1 And Setting(42)=0 And Setting(43)=0 And Setting(44)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 13:If Setting(45)=0 And Setting(46)=0 And Setting(47)=0 And Setting(48)=0 Then CurSelect=0
       If Setting(45)=1 And Setting(46)=0 And Setting(47)=0 And Setting(48)=0 Then CurSelect=1
       If Setting(45)=0 And Setting(46)=1 And Setting(47)=0 And Setting(48)=0 Then CurSelect=2
       If Setting(45)=1 And Setting(46)=1 And Setting(47)=0 And Setting(48)=0 Then CurSelect=3
       If Setting(45)=0 And Setting(46)=0 And Setting(47)=1 And Setting(48)=0 Then CurSelect=4
       If Setting(45)=1 And Setting(46)=0 And Setting(47)=1 And Setting(48)=0 Then CurSelect=5
       If Setting(45)=0 And Setting(46)=1 And Setting(47)=1 And Setting(48)=0 Then CurSelect=6
       If Setting(45)=1 And Setting(46)=1 And Setting(47)=1 And Setting(48)=0 Then CurSelect=7
       If Setting(45)=0 And Setting(46)=0 And Setting(47)=0 And Setting(48)=1 Then CurSelect=8
       If Setting(45)=1 And Setting(46)=0 And Setting(47)=0 And Setting(48)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 14:If Setting(49)=0 And Setting(50)=0 And Setting(51)=0 And Setting(52)=0 Then CurSelect=0
       If Setting(49)=1 And Setting(50)=0 And Setting(51)=0 And Setting(52)=0 Then CurSelect=1
       If Setting(49)=0 And Setting(50)=1 And Setting(51)=0 And Setting(52)=0 Then CurSelect=2
       If Setting(49)=1 And Setting(50)=1 And Setting(51)=0 And Setting(52)=0 Then CurSelect=3
       If Setting(49)=0 And Setting(50)=0 And Setting(51)=1 And Setting(52)=0 Then CurSelect=4
       If Setting(49)=1 And Setting(50)=0 And Setting(51)=1 And Setting(52)=0 Then CurSelect=5
       If Setting(49)=0 And Setting(50)=1 And Setting(51)=1 And Setting(52)=0 Then CurSelect=6
       If Setting(49)=1 And Setting(50)=1 And Setting(51)=1 And Setting(52)=0 Then CurSelect=7
       If Setting(49)=0 And Setting(50)=0 And Setting(51)=0 And Setting(52)=1 Then CurSelect=8
       If Setting(49)=1 And Setting(50)=0 And Setting(51)=0 And Setting(52)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 15:If Setting(53)=0 And Setting(54)=0 And Setting(55)=0 And Setting(56)=0 Then CurSelect=0
       If Setting(53)=1 And Setting(54)=0 And Setting(55)=0 And Setting(56)=0 Then CurSelect=1
       If Setting(53)=0 And Setting(54)=1 And Setting(55)=0 And Setting(56)=0 Then CurSelect=2
       If Setting(53)=1 And Setting(54)=1 And Setting(55)=0 And Setting(56)=0 Then CurSelect=3
       If Setting(53)=0 And Setting(54)=0 And Setting(55)=1 And Setting(56)=0 Then CurSelect=4
       If Setting(53)=1 And Setting(54)=0 And Setting(55)=1 And Setting(56)=0 Then CurSelect=5
       If Setting(53)=0 And Setting(54)=1 And Setting(55)=1 And Setting(56)=0 Then CurSelect=6
       If Setting(53)=1 And Setting(54)=1 And Setting(55)=1 And Setting(56)=0 Then CurSelect=7
       If Setting(53)=0 And Setting(54)=0 And Setting(55)=0 And Setting(56)=1 Then CurSelect=8
       If Setting(53)=1 And Setting(54)=0 And Setting(55)=0 And Setting(56)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 16:If Setting(57)=0 And Setting(58)=0 And Setting(59)=0 And Setting(60)=0 Then CurSelect=0
       If Setting(57)=1 And Setting(58)=0 And Setting(59)=0 And Setting(60)=0 Then CurSelect=1
       If Setting(57)=0 And Setting(58)=1 And Setting(59)=0 And Setting(60)=0 Then CurSelect=2
       If Setting(57)=1 And Setting(58)=1 And Setting(59)=0 And Setting(60)=0 Then CurSelect=3
       If Setting(57)=0 And Setting(58)=0 And Setting(59)=1 And Setting(60)=0 Then CurSelect=4
       If Setting(57)=1 And Setting(58)=0 And Setting(59)=1 And Setting(60)=0 Then CurSelect=5
       If Setting(57)=0 And Setting(58)=1 And Setting(59)=1 And Setting(60)=0 Then CurSelect=6
       If Setting(57)=1 And Setting(58)=1 And Setting(59)=1 And Setting(60)=0 Then CurSelect=7
       If Setting(57)=0 And Setting(58)=0 And Setting(59)=0 And Setting(60)=1 Then CurSelect=8
       If Setting(57)=1 And Setting(58)=0 And Setting(59)=0 And Setting(60)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 17:If Setting(61)=0 And Setting(62)=0 And Setting(63)=0 And Setting(64)=0 Then CurSelect=0
       If Setting(61)=1 And Setting(62)=0 And Setting(63)=0 And Setting(64)=0 Then CurSelect=1
       If Setting(61)=0 And Setting(62)=1 And Setting(63)=0 And Setting(64)=0 Then CurSelect=2
       If Setting(61)=1 And Setting(62)=1 And Setting(63)=0 And Setting(64)=0 Then CurSelect=3
       If Setting(61)=0 And Setting(62)=0 And Setting(63)=1 And Setting(64)=0 Then CurSelect=4
       If Setting(61)=1 And Setting(62)=0 And Setting(63)=1 And Setting(64)=0 Then CurSelect=5
       If Setting(61)=0 And Setting(62)=1 And Setting(63)=1 And Setting(64)=0 Then CurSelect=6
       If Setting(61)=1 And Setting(62)=1 And Setting(63)=1 And Setting(64)=0 Then CurSelect=7
       If Setting(61)=0 And Setting(62)=0 And Setting(63)=0 And Setting(64)=1 Then CurSelect=8
       If Setting(61)=1 And Setting(62)=0 And Setting(63)=0 And Setting(64)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
    Case 18:If Setting(65)=0 And Setting(66)=0 And Setting(67)=0 And Setting(68)=0 Then CurSelect=0
       If Setting(65)=1 And Setting(66)=0 And Setting(67)=0 And Setting(68)=0 Then CurSelect=1
       If Setting(65)=0 And Setting(66)=1 And Setting(67)=0 And Setting(68)=0 Then CurSelect=2
       If Setting(65)=1 And Setting(66)=1 And Setting(67)=0 And Setting(68)=0 Then CurSelect=3
       If Setting(65)=0 And Setting(66)=0 And Setting(67)=1 And Setting(68)=0 Then CurSelect=4
       If Setting(65)=1 And Setting(66)=0 And Setting(67)=1 And Setting(68)=0 Then CurSelect=5
       If Setting(65)=0 And Setting(66)=1 And Setting(67)=1 And Setting(68)=0 Then CurSelect=6
       If Setting(65)=1 And Setting(66)=1 And Setting(67)=1 And Setting(68)=0 Then CurSelect=7
       If Setting(65)=0 And Setting(66)=0 And Setting(67)=0 And Setting(68)=1 Then CurSelect=8
       If Setting(65)=1 And Setting(66)=0 And Setting(67)=0 And Setting(68)=1 Then CurSelect=9
       EMReel2.SetValue 1+CurSelect
  End Select
End Sub

Sub StoreSetting(DP)
 Select Case DP
    Case 0:Setting(2)=CurSelect
   Case 1:Setting(3)=CurSelect
   Case 2:Setting(4)=CurSelect
   Case 3:If CurSelect=0 Then:Setting(5)=1:Setting(6)=0:Setting(7)=0:End If
        If CurSelect=1 Then:Setting(5)=0:Setting(6)=1:Setting(7)=0:End If
       If CurSelect=2 Then:Setting(5)=1:Setting(6)=1:Setting(7)=0:End If
       If CurSelect=3 Then:Setting(5)=0:Setting(6)=0:Setting(7)=1:End If
       If CurSelect=4 Then:Setting(5)=1:Setting(6)=0:Setting(7)=1:End If
   Case 4:If CurSelect=0 Then:Setting(9)=0:Setting(10)=0:Setting(11)=0:Setting(12)=0:End If
        If CurSelect=1 Then:Setting(9)=1:Setting(10)=0:Setting(11)=0:Setting(12)=0:End If
       If CurSelect=2 Then:Setting(9)=0:Setting(10)=1:Setting(11)=0:Setting(12)=0:End If
       If CurSelect=3 Then:Setting(9)=1:Setting(10)=1:Setting(11)=0:Setting(12)=0:End If
       If CurSelect=4 Then:Setting(9)=0:Setting(10)=0:Setting(11)=1:Setting(12)=0:End If
       If CurSelect=5 Then:Setting(9)=1:Setting(10)=0:Setting(11)=1:Setting(12)=0:End If
       If CurSelect=6 Then:Setting(9)=0:Setting(10)=1:Setting(11)=1:Setting(12)=0:End If
       If CurSelect=7 Then:Setting(9)=1:Setting(10)=1:Setting(11)=1:Setting(12)=0:End If
       If CurSelect=8 Then:Setting(9)=0:Setting(10)=0:Setting(11)=0:Setting(12)=1:End If
       If CurSelect=9 Then:Setting(9)=1:Setting(10)=0:Setting(11)=0:Setting(12)=1:End If
   Case 5:If CurSelect=0 Then:Setting(13)=0:Setting(14)=0:Setting(15)=0:Setting(16)=0:End If
       If CurSelect=1 Then:Setting(13)=1:Setting(14)=0:Setting(15)=0:Setting(16)=0:End If
        If CurSelect=2 Then:Setting(13)=0:Setting(14)=1:Setting(15)=0:Setting(16)=0:End If
        If CurSelect=3 Then:Setting(13)=1:Setting(14)=1:Setting(15)=0:Setting(16)=0:End If
        If CurSelect=4 Then:Setting(13)=0:Setting(14)=0:Setting(15)=1:Setting(16)=0:End If
        If CurSelect=5 Then:Setting(13)=1:Setting(14)=0:Setting(15)=1:Setting(16)=0:End If
        If CurSelect=6 Then:Setting(13)=0:Setting(14)=1:Setting(15)=1:Setting(16)=0:End If
        If CurSelect=7 Then:Setting(13)=1:Setting(14)=1:Setting(15)=1:Setting(16)=0:End If
        If CurSelect=8 Then:Setting(13)=0:Setting(14)=0:Setting(15)=0:Setting(16)=1:End If
        If CurSelect=9 Then:Setting(13)=1:Setting(14)=0:Setting(15)=0:Setting(16)=1:End If
    Case 6:If CurSelect=0 Then:Setting(17)=0:Setting(18)=0:Setting(19)=0:Setting(20)=0:End If
       If CurSelect=1 Then:Setting(17)=1:Setting(18)=0:Setting(19)=0:Setting(20)=0:End If
        If CurSelect=2 Then:Setting(17)=0:Setting(18)=1:Setting(19)=0:Setting(20)=0:End If
        If CurSelect=3 Then:Setting(17)=1:Setting(18)=1:Setting(19)=0:Setting(20)=0:End If
        If CurSelect=4 Then:Setting(17)=0:Setting(18)=0:Setting(19)=1:Setting(20)=0:End If
        If CurSelect=5 Then:Setting(17)=1:Setting(18)=0:Setting(19)=1:Setting(20)=0:End If
        If CurSelect=6 Then:Setting(17)=0:Setting(18)=1:Setting(19)=1:Setting(20)=0:End If
        If CurSelect=7 Then:Setting(17)=1:Setting(18)=1:Setting(19)=1:Setting(20)=0:End If
        If CurSelect=8 Then:Setting(17)=0:Setting(18)=0:Setting(19)=0:Setting(20)=1:End If
        If CurSelect=9 Then:Setting(17)=1:Setting(18)=0:Setting(19)=0:Setting(20)=1:End If
    Case 7:If CurSelect=0 Then:Setting(21)=0:Setting(22)=0:Setting(23)=0:Setting(24)=0:End If
       If CurSelect=1 Then:Setting(21)=1:Setting(22)=0:Setting(23)=0:Setting(24)=0:End If
        If CurSelect=2 Then:Setting(21)=0:Setting(22)=1:Setting(23)=0:Setting(24)=0:End If
        If CurSelect=3 Then:Setting(21)=1:Setting(22)=1:Setting(23)=0:Setting(24)=0:End If
        If CurSelect=4 Then:Setting(21)=0:Setting(22)=0:Setting(23)=1:Setting(24)=0:End If
        If CurSelect=5 Then:Setting(21)=1:Setting(22)=0:Setting(23)=1:Setting(24)=0:End If
        If CurSelect=6 Then:Setting(21)=0:Setting(22)=1:Setting(23)=1:Setting(24)=0:End If
        If CurSelect=7 Then:Setting(21)=1:Setting(22)=1:Setting(23)=1:Setting(24)=0:End If
        If CurSelect=8 Then:Setting(21)=0:Setting(22)=0:Setting(23)=0:Setting(24)=1:End If
        If CurSelect=9 Then:Setting(21)=1:Setting(22)=0:Setting(23)=0:Setting(24)=1:End If
    Case 8:If CurSelect=0 Then:Setting(25)=0:Setting(26)=0:Setting(27)=0:Setting(28)=0:End If
       If CurSelect=1 Then:Setting(25)=1:Setting(26)=0:Setting(27)=0:Setting(28)=0:End If
        If CurSelect=2 Then:Setting(25)=0:Setting(26)=1:Setting(27)=0:Setting(28)=0:End If
        If CurSelect=3 Then:Setting(25)=1:Setting(26)=1:Setting(27)=0:Setting(28)=0:End If
        If CurSelect=4 Then:Setting(25)=0:Setting(26)=0:Setting(27)=1:Setting(28)=0:End If
        If CurSelect=5 Then:Setting(25)=1:Setting(26)=0:Setting(27)=1:Setting(28)=0:End If
        If CurSelect=6 Then:Setting(25)=0:Setting(26)=1:Setting(27)=1:Setting(28)=0:End If
        If CurSelect=7 Then:Setting(25)=1:Setting(26)=1:Setting(27)=1:Setting(28)=0:End If
        If CurSelect=8 Then:Setting(25)=0:Setting(26)=0:Setting(27)=0:Setting(28)=1:End If
        If CurSelect=9 Then:Setting(25)=1:Setting(26)=0:Setting(27)=0:Setting(28)=1:End If
    Case 9:If CurSelect=0 Then:Setting(29)=0:Setting(30)=0:Setting(31)=0:Setting(32)=0:End If
       If CurSelect=1 Then:Setting(29)=1:Setting(30)=0:Setting(31)=0:Setting(32)=0:End If
        If CurSelect=2 Then:Setting(29)=0:Setting(30)=1:Setting(31)=0:Setting(32)=0:End If
        If CurSelect=3 Then:Setting(29)=1:Setting(30)=1:Setting(31)=0:Setting(32)=0:End If
        If CurSelect=4 Then:Setting(29)=0:Setting(30)=0:Setting(31)=1:Setting(32)=0:End If
        If CurSelect=5 Then:Setting(29)=1:Setting(30)=0:Setting(31)=1:Setting(32)=0:End If
        If CurSelect=6 Then:Setting(29)=0:Setting(30)=1:Setting(31)=1:Setting(32)=0:End If
        If CurSelect=7 Then:Setting(29)=1:Setting(30)=1:Setting(31)=1:Setting(32)=0:End If
        If CurSelect=8 Then:Setting(29)=0:Setting(30)=0:Setting(31)=0:Setting(32)=1:End If
        If CurSelect=9 Then:Setting(29)=1:Setting(30)=0:Setting(31)=0:Setting(32)=1:End If
    Case 10:If CurSelect=0 Then:Setting(33)=0:Setting(34)=0:Setting(35)=0:Setting(36)=0:End If
        If CurSelect=1 Then:Setting(33)=1:Setting(34)=0:Setting(35)=0:Setting(36)=0:End If
        If CurSelect=2 Then:Setting(33)=0:Setting(34)=1:Setting(35)=0:Setting(36)=0:End If
        If CurSelect=3 Then:Setting(33)=1:Setting(34)=1:Setting(35)=0:Setting(36)=0:End If
        If CurSelect=4 Then:Setting(33)=0:Setting(34)=0:Setting(35)=1:Setting(36)=0:End If
        If CurSelect=5 Then:Setting(33)=1:Setting(34)=0:Setting(35)=1:Setting(36)=0:End If
        If CurSelect=6 Then:Setting(33)=0:Setting(34)=1:Setting(35)=1:Setting(36)=0:End If
        If CurSelect=7 Then:Setting(33)=1:Setting(34)=1:Setting(35)=1:Setting(36)=0:End If
        If CurSelect=8 Then:Setting(33)=0:Setting(34)=0:Setting(35)=0:Setting(36)=1:End If
        If CurSelect=9 Then:Setting(33)=1:Setting(34)=0:Setting(35)=0:Setting(36)=1:End If
    Case 11:If CurSelect=0 Then:Setting(37)=0:Setting(38)=0:Setting(39)=0:Setting(40)=0:End If
        If CurSelect=1 Then:Setting(37)=1:Setting(38)=0:Setting(39)=0:Setting(40)=0:End If
        If CurSelect=2 Then:Setting(37)=0:Setting(38)=1:Setting(39)=0:Setting(40)=0:End If
        If CurSelect=3 Then:Setting(37)=1:Setting(38)=1:Setting(39)=0:Setting(40)=0:End If
        If CurSelect=4 Then:Setting(37)=0:Setting(38)=0:Setting(39)=1:Setting(40)=0:End If
        If CurSelect=5 Then:Setting(37)=1:Setting(38)=0:Setting(39)=1:Setting(40)=0:End If
        If CurSelect=6 Then:Setting(37)=0:Setting(38)=1:Setting(39)=1:Setting(40)=0:End If
        If CurSelect=7 Then:Setting(37)=1:Setting(38)=1:Setting(39)=1:Setting(40)=0:End If
        If CurSelect=8 Then:Setting(37)=0:Setting(38)=0:Setting(39)=0:Setting(40)=1:End If
        If CurSelect=9 Then:Setting(37)=1:Setting(38)=0:Setting(39)=0:Setting(40)=1:End If
    Case 12:If CurSelect=0 Then:Setting(41)=0:Setting(42)=0:Setting(43)=0:Setting(44)=0:End If
        If CurSelect=1 Then:Setting(41)=1:Setting(42)=0:Setting(43)=0:Setting(44)=0:End If
        If CurSelect=2 Then:Setting(41)=0:Setting(42)=1:Setting(43)=0:Setting(44)=0:End If
        If CurSelect=3 Then:Setting(41)=1:Setting(42)=1:Setting(43)=0:Setting(44)=0:End If
        If CurSelect=4 Then:Setting(41)=0:Setting(42)=0:Setting(43)=1:Setting(44)=0:End If
        If CurSelect=5 Then:Setting(41)=1:Setting(42)=0:Setting(43)=1:Setting(44)=0:End If
        If CurSelect=6 Then:Setting(41)=0:Setting(42)=1:Setting(43)=1:Setting(44)=0:End If
        If CurSelect=7 Then:Setting(41)=1:Setting(42)=1:Setting(43)=1:Setting(44)=0:End If
        If CurSelect=8 Then:Setting(41)=0:Setting(42)=0:Setting(43)=0:Setting(44)=1:End If
        If CurSelect=9 Then:Setting(41)=1:Setting(42)=0:Setting(43)=0:Setting(44)=1:End If
    Case 13:If CurSelect=0 Then:Setting(45)=0:Setting(46)=0:Setting(47)=0:Setting(48)=0:End If
        If CurSelect=1 Then:Setting(45)=1:Setting(46)=0:Setting(47)=0:Setting(48)=0:End If
        If CurSelect=2 Then:Setting(45)=0:Setting(46)=1:Setting(47)=0:Setting(48)=0:End If
        If CurSelect=3 Then:Setting(45)=1:Setting(46)=1:Setting(47)=0:Setting(48)=0:End If
        If CurSelect=4 Then:Setting(45)=0:Setting(46)=0:Setting(47)=1:Setting(48)=0:End If
        If CurSelect=5 Then:Setting(45)=1:Setting(46)=0:Setting(47)=1:Setting(48)=0:End If
        If CurSelect=6 Then:Setting(45)=0:Setting(46)=1:Setting(47)=1:Setting(48)=0:End If
        If CurSelect=7 Then:Setting(45)=1:Setting(46)=1:Setting(47)=1:Setting(48)=0:End If
        If CurSelect=8 Then:Setting(45)=0:Setting(46)=0:Setting(47)=0:Setting(48)=1:End If
        If CurSelect=9 Then:Setting(45)=1:Setting(46)=0:Setting(47)=0:Setting(48)=1:End If
    Case 14:If CurSelect=0 Then:Setting(49)=0:Setting(50)=0:Setting(51)=0:Setting(52)=0:End If
        If CurSelect=1 Then:Setting(49)=1:Setting(50)=0:Setting(51)=0:Setting(52)=0:End If
        If CurSelect=2 Then:Setting(49)=0:Setting(50)=1:Setting(51)=0:Setting(52)=0:End If
        If CurSelect=3 Then:Setting(49)=1:Setting(50)=1:Setting(51)=0:Setting(52)=0:End If
        If CurSelect=4 Then:Setting(49)=0:Setting(50)=0:Setting(51)=1:Setting(52)=0:End If
        If CurSelect=5 Then:Setting(49)=1:Setting(50)=0:Setting(51)=1:Setting(52)=0:End If
        If CurSelect=6 Then:Setting(49)=0:Setting(50)=1:Setting(51)=1:Setting(52)=0:End If
        If CurSelect=7 Then:Setting(49)=1:Setting(50)=1:Setting(51)=1:Setting(52)=0:End If
        If CurSelect=8 Then:Setting(49)=0:Setting(50)=0:Setting(51)=0:Setting(52)=1:End If
        If CurSelect=9 Then:Setting(49)=1:Setting(50)=0:Setting(51)=0:Setting(52)=1:End If
    Case 15:If CurSelect=0 Then:Setting(53)=0:Setting(54)=0:Setting(55)=0:Setting(56)=0:End If
        If CurSelect=1 Then:Setting(53)=1:Setting(54)=0:Setting(55)=0:Setting(56)=0:End If
        If CurSelect=2 Then:Setting(53)=0:Setting(54)=1:Setting(55)=0:Setting(56)=0:End If
        If CurSelect=3 Then:Setting(53)=1:Setting(54)=1:Setting(55)=0:Setting(56)=0:End If
        If CurSelect=4 Then:Setting(53)=0:Setting(54)=0:Setting(55)=1:Setting(56)=0:End If
        If CurSelect=5 Then:Setting(53)=1:Setting(54)=0:Setting(55)=1:Setting(56)=0:End If
        If CurSelect=6 Then:Setting(53)=0:Setting(54)=1:Setting(55)=1:Setting(56)=0:End If
        If CurSelect=7 Then:Setting(53)=1:Setting(54)=1:Setting(55)=1:Setting(56)=0:End If
        If CurSelect=8 Then:Setting(53)=0:Setting(54)=0:Setting(55)=0:Setting(56)=1:End If
        If CurSelect=9 Then:Setting(53)=1:Setting(54)=0:Setting(55)=0:Setting(56)=1:End If
    Case 16:If CurSelect=0 Then:Setting(57)=0:Setting(58)=0:Setting(59)=0:Setting(60)=0:End If
        If CurSelect=1 Then:Setting(57)=1:Setting(58)=0:Setting(59)=0:Setting(60)=0:End If
        If CurSelect=2 Then:Setting(57)=0:Setting(58)=1:Setting(59)=0:Setting(60)=0:End If
        If CurSelect=3 Then:Setting(57)=1:Setting(58)=1:Setting(59)=0:Setting(60)=0:End If
        If CurSelect=4 Then:Setting(57)=0:Setting(58)=0:Setting(59)=1:Setting(60)=0:End If
        If CurSelect=5 Then:Setting(57)=1:Setting(58)=0:Setting(59)=1:Setting(60)=0:End If
        If CurSelect=6 Then:Setting(57)=0:Setting(58)=1:Setting(59)=1:Setting(60)=0:End If
        If CurSelect=7 Then:Setting(57)=1:Setting(58)=1:Setting(59)=1:Setting(60)=0:End If
        If CurSelect=8 Then:Setting(57)=0:Setting(58)=0:Setting(59)=0:Setting(60)=1:End If
        If CurSelect=9 Then:Setting(57)=1:Setting(58)=0:Setting(59)=0:Setting(60)=1:End If
    Case 17:If CurSelect=0 Then:Setting(61)=0:Setting(62)=0:Setting(63)=0:Setting(64)=0:End If
        If CurSelect=1 Then:Setting(61)=1:Setting(62)=0:Setting(63)=0:Setting(64)=0:End If
        If CurSelect=2 Then:Setting(61)=0:Setting(62)=1:Setting(63)=0:Setting(64)=0:End If
        If CurSelect=3 Then:Setting(61)=1:Setting(62)=1:Setting(63)=0:Setting(64)=0:End If
        If CurSelect=4 Then:Setting(61)=0:Setting(62)=0:Setting(63)=1:Setting(64)=0:End If
        If CurSelect=5 Then:Setting(61)=1:Setting(62)=0:Setting(63)=1:Setting(64)=0:End If
        If CurSelect=6 Then:Setting(61)=0:Setting(62)=1:Setting(63)=1:Setting(64)=0:End If
        If CurSelect=7 Then:Setting(61)=1:Setting(62)=1:Setting(63)=1:Setting(64)=0:End If
        If CurSelect=8 Then:Setting(61)=0:Setting(62)=0:Setting(63)=0:Setting(64)=1:End If
        If CurSelect=9 Then:Setting(61)=1:Setting(62)=0:Setting(63)=0:Setting(64)=1:End If
    Case 18:If CurSelect=0 Then:Setting(65)=0:Setting(66)=0:Setting(67)=0:Setting(68)=0:End If
        If CurSelect=1 Then:Setting(65)=1:Setting(66)=0:Setting(67)=0:Setting(68)=0:End If
        If CurSelect=2 Then:Setting(65)=0:Setting(66)=1:Setting(67)=0:Setting(68)=0:End If
        If CurSelect=3 Then:Setting(65)=1:Setting(66)=1:Setting(67)=0:Setting(68)=0:End If
        If CurSelect=4 Then:Setting(65)=0:Setting(66)=0:Setting(67)=1:Setting(68)=0:End If
        If CurSelect=5 Then:Setting(65)=1:Setting(66)=0:Setting(67)=1:Setting(68)=0:End If
        If CurSelect=6 Then:Setting(65)=0:Setting(66)=1:Setting(67)=1:Setting(68)=0:End If
        If CurSelect=7 Then:Setting(65)=1:Setting(66)=1:Setting(67)=1:Setting(68)=0:End If
        If CurSelect=8 Then:Setting(65)=0:Setting(66)=0:Setting(67)=0:Setting(68)=1:End If
        If CurSelect=9 Then:Setting(65)=1:Setting(66)=0:Setting(67)=0:Setting(68)=1:End If
  End Select
End Sub

Sub Trigger10_Hit
  If ActiveBall.VelY<-20 Then
   StopSound"ALBallLaunch"
   PlaySoundAtVol "ALBallLaunch", ActiveBall, 1
 End If
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
' Thalamus, AudioFade - Patched
  If tmp > 0 Then
    AudioFade = Csng(tmp ^5) 'was 10
  Else
    AudioFade = Csng(-((- tmp) ^5) ) 'was 10
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
' Thalamus, AudioPan - Patched
  If tmp > 0 Then
    AudioPan = Csng(tmp ^5) 'was 10
  Else
    AudioPan = Csng(-((- tmp) ^5) ) 'was 10
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
' Thalamus, Pan - Patched
  If tmp > 0 Then
    Pan = Csng(tmp ^5) 'was 10
  Else
    Pan = Csng(-((- tmp) ^5) ) ' was 10
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

