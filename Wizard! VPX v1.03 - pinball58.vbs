'****************************************************************************************
' WIZARD!
' Bally 1975
' version 1.03
' VPX EM Table Recreation by pinball58
' Thanks to Ninuzzu for the help
' Thanks to Tom Tower for the monsterbashpincab site
'****************************************************************************************

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

Const cGameName = "Wizard_1975"

'************************************* TABLE OPTIONS **********************************************
'--------------------------------------------------------------------------------------------------
'**************************************************************************************************


LightMod = 1	     'Static Top Lights = 0       Flashing Top Lights = 1

ApronColor = 3       'Grey/Yellow Apron = 1	      Black/Yellow Apron = 2	   Random Apron = 3

FlippersColor = 3    'Yellow rubbers = 1          Red rubbers = 2              Random Rubbers = 3

BallsNumber = 0      '5 Ball per Game = 1         3 Balls per Game = 0

Special = 1          'Special Add a Credit = 1    Special Add a Ball = 0


'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
'**************************************************************************************************

Dim Bonus
Dim Credit
Dim Game
Dim LightMod
Dim BallN,BallN2,BallN3,BallN4
Dim Replay(4)
Dim ReplayScore1(4)
Dim ReplayScore2(4)
Dim ReplayScore3(4)
Dim TiltCount
Dim Tilt
Dim ApronColor
Dim FlippersColor
Dim BNumber
Dim BallsNumber
Dim Special
Dim FlashingLEB
Dim MatchNum
Dim Match
Dim MatchP(4)
Dim Player
Dim PlayerN
Dim AllowStart
Dim Init
Dim BallsNTot
Dim x
Dim Score(4)
Dim EMR (4)
Dim Score100k(4)
Dim BallIP
Dim PlayEB
Dim EB
Dim PUP

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

If Wizard.ShowDT = false then
    EMReel1.Visible = false
	EMReel2.Visible = false
	EMReel3.Visible = false
	EMReel4.Visible = false
	P1up.Visible = false
	P2up.Visible = false
	P3up.Visible = false
	P4up.Visible = false
	Score100k1.Visible = false
	Score100k2.Visible = false
	Score100k3.Visible = false
	Score100k4.Visible = false
	CreditsOsd.Visible = false
	CreditCount.Visible = false
	PlayerOsd.Visible = false
	PlayerPlay.Visible = false
	CanPlay.Visible = false
	PlayerNum.Visible = false
	BallsOsd.Visible = false
	BallCount.Visible = false
	m0.Visible = false : m1.Visible = false : m2.Visible = false : m3.Visible = false : m4.Visible = false : m5.Visible = false : m6.Visible = false : m7.Visible = false : m8.Visible = false : m9.Visible = false
	SAtb.Visible = false : SA2tb.Visible = false
	TILTtb.Visible = false
	GameOvertb.Visible = false
End If

'********* Table Init *************************************************

Sub Wizard_Init
	LoadEM
	Kicker1.timerenabled = 0
	Bonus = 1
	Replay(1) = 1 : Replay(2) = 1 : Replay(3) = 1 : Replay(4) = 1
	Game = 0
	AllowStart = 1
	PlayEB = 0
	EB = 0
	Set EMR(1) = EMReel1
	Set EMR(2) = EMReel2
	Set EMR(3) = EMReel3
	Set EMR(4) = EMReel4
	Set Score100k(1) = Score100k1
	Set Score100k(2) = Score100k2
	Set Score100k(3) = Score100k3
	Set Score100k(4) = Score100k4
	If BallsNumber = 0 Then
		BNumber = 3
		ReplayScore1(1) = 65000 : ReplayScore2(1) = 95000 : ReplayScore3(3) = 115000
		ReplayScore1(2) = 65000 : ReplayScore2(2) = 95000 : ReplayScore3(2) = 115000
	    ReplayScore1(3) = 65000 : ReplayScore2(3) = 95000 : ReplayScore3(3) = 115000
        ReplayScore1(4) = 65000 : ReplayScore2(4) = 95000 : ReplayScore3(4) = 115000
	End If
	If BallsNumber = 1 Then
	BNumber = 5
		ReplayScore1(1) = 95000 : ReplayScore2(1) = 120000 : ReplayScore3(1) = 144000
		ReplayScore1(2) = 95000 : ReplayScore2(2) = 120000 : ReplayScore3(2) = 144000
		ReplayScore1(3) = 95000 : ReplayScore2(3) = 120000 : ReplayScore3(3) = 144000
		ReplayScore1(4) = 95000 : ReplayScore2(4) = 120000 : ReplayScore3(4) = 144000
	End If
	If Special = 1 And BallsNumber = 1 Then ScoreCard.image = "scorecard5b" : InstructionCard.image = "instructioncardReplay" :  End If
	If Special = 1 And BallsNumber = 0 Then ScoreCard.image = "scorecard3b" : InstructionCard.image = "instructioncardReplay" :  End If
	If Special = 0 And BallsNumber = 1 Then ScoreCard.image = "scorecard5bNoRep" : InstructionCard.image = "instructioncardExtraBall" :  End If
	If Special = 0 And BallsNumber = 0 Then ScoreCard.image = "scorecard3bNoRep" : InstructionCard.image = "instructioncardExtraBall" :  End If
End Sub

'**********************************************************************

'********* Add Score & Replay ****************************************

Sub AddScore(points)

If Player = 1 Then x = 1 End If
If Player = 2 Then x = 2 End If
If Player = 3 Then x = 3 End If
If Player = 4 Then x = 4 End If

Score(x) = Score(x) + points

EMR(x).setvalue(score(x))

If B2SOn Then
Controller.B2SSetScorePlayer 1, Score(1)
Controller.B2SSetScorePlayer 2, Score(2)
Controller.B2SSetScorePlayer 3, Score(3)
Controller.B2SSetScorePlayer 4, Score(4)
End If

If Score(x)>99990 Then
	Score100k(x).text = "100.000"
	If B2SOn Then
	Controller.B2SSetScoreRollover 24+Player, Player
	End If
	Else
	Score100k(x).text = ""
End If

If points = +10 Then MatchP(x) = MatchP(x) + 10 End If
If MatchP(x) = 100 Then MatchP(x) = 0 End If

If Score(x) >= ReplayScore1(x) And Replay(x) = 1 And Special = 1 Then
	If Credit <25 Then
	Credit = Credit + 1
	DOF 119, DOFOn
	CreditCount.text= (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	LightCredit.state = 1
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	Replay(x)=2
End If

If Score(x) >= (ReplayScore1(x)*2) And Replay(x) = 2 And Special = 1 Then
	If Credit <25 Then
	Credit = Credit + 1
	DOF 119, DOFOn
	CreditCount.text= (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	LightCredit.state = 1
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	Replay(x)=3
End If

If Score(x) >= (ReplayScore1(x)*3) And Replay(x) = 3 And Special = 1 Then
	If Credit <25 Then
	Credit = Credit + 1
	DOF 119, DOFOn
	CreditCount.text= (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	LightCredit.state = 1
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	Replay(x)=4
End If

If Score(x) >= (ReplayScore2(x)) And Replay(x) = 5 And Special = 1  Then
	If Credit <25 Then
	Credit = Credit + 1
	DOF 119, DOFOn
	CreditCount.text= (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	LightCredit.state = 1
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	Replay(x)=6
End If

If Score(x) >= (ReplayScore2(x)*2) And Replay(x) = 6 And Special = 1 Then
	If Credit <25 Then
	Credit = Credit + 1
	DOF 119, DOFOn
	CreditCount.text= (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	LightCredit.state = 1
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	Replay(x)=7
End If

If Score(x) >= (ReplayScore2(x)*3) And Replay(x) = 7 And Special = 1 Then
	If Credit <25 Then
	Credit = Credit + 1
	DOF 119, DOFOn
	CreditCount.text= (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	LightCredit.state = 1
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	Replay(x)=8
End If

If Score(x) >= (ReplayScore3(x)) And Replay(x) = 10 And Special = 1 Then
	If Credit <25 Then
	Credit = Credit + 1
	DOF 119, DOFOn
	CreditCount.text= (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	LightCredit.state = 1
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	Replay(x)=11
End If

If Score(x) >= (ReplayScore3(x)*2) And Replay(x) = 11 And Special = 1 Then
	If Credit <25 Then
	Credit = Credit + 1
	DOF 119, DOFOn
	CreditCount.text= (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	LightCredit.state = 1
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	Replay(x)=12
End If

If Score(x) >= (ReplayScore3(x)*3) And Replay(x) = 12 And Special = 1 Then
	If Credit <25 Then
	Credit = Credit + 1
	DOF 119, DOFOn
	CreditCount.text= (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	LightCredit.state = 1
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	Replay(x)=0
End If

'**** Chimes Sound Score based ***

If points = 10 And (Score(x) MOD 100)\10 = 0 Then
	PlaySound SoundFXDOF("Chimes100",142,DOFPulse,DOFChimes)
    ElseIf points = 100 And(Score(x) MOD 1000)\10 = 0 Then
	PlaySound SoundFXDOF("Chimes1000",143,DOFPulse,DOFChimes)
    ElseIf points = 1000 Then
	PlaySound SoundFXDOF("Chimes1000",143,DOFPulse,DOFChimes)
	ElseIf points = 100 Then
	PlaySound SoundFXDOF("Chimes100",142,DOFPulse,DOFChimes)
	ElseIf points = 10 Then
	PlaySound SoundFXDOF("Chimes10",141,DOFPulse,DOFChimes)
End If

'*********************************

End Sub

'******* 300-3000-500-5000 Score & Chimes Sounds **********

Sub P300_timer()
	AddScore (100)
	P300.enabled = 0
	Ch100Timer.enabled = 1
End Sub

Sub P3000_timer()
	AddScore (1000)
	P3000.enabled = 0
	Ch1000Timer.enabled = 1
End Sub

Sub Ch100Timer_timer()
	AddScore (100)
	Ch100Timer.enabled = 0
End Sub

Sub Ch1000Timer_timer()
	AddScore (1000)
	Ch1000Timer.enabled = 0
End Sub

Sub P500_timer()
	AddScore (100)
	P500.enabled = 0
	Ch100Timer.enabled = 1
	Ch100Timer2.enabled = 1
	Ch100Timer3.enabled = 1
End Sub

Sub P5000_timer()
	AddScore (1000)
	P5000.enabled = 0
	Ch1000Timer.enabled = 1
	Ch1000Timer2.enabled = 1
	Ch1000Timer3.enabled = 1
End Sub

Sub Ch100Timer2_timer()
	AddScore (100)
	Ch100Timer2.enabled = 0
End Sub

Sub Ch100Timer3_timer()
	AddScore (100)
	Ch100Timer3.enabled = 0
End Sub

Sub Ch1000Timer2_timer()
	AddScore (1000)
	Ch1000Timer2.enabled = 0
End Sub

Sub Ch1000Timer3_timer()
	AddScore (1000)
	Ch1000Timer3.enabled = 0
End Sub

'************************************************************

'********* Match **************

Sub MatchTimer_timer()

MatchNum = Int(Rnd*10)+1
Select Case MatchNum
    Case 1 : Match = 0
    Case 2 : Match = 10
	Case 3 : Match = 20
	Case 4 : Match = 30
	Case 5 : Match = 40
	Case 6 : Match = 50
	Case 7 : Match = 60
	Case 8 : Match = 70
	Case 9 : Match = 80
	Case 10 : Match = 90
End Select

If Match = 0 Then m0.text = "00" End If
If Match = 10 Then m1.text = "10" End If
If Match = 20 Then m2.text = "20" End If
If Match = 30 Then m3.text = "30" End If
If Match = 40 Then m4.text = "40" End If
If Match = 50 Then m5.text = "50" End If
If Match = 60 Then m6.text = "60" End If
If Match = 70 Then m7.text = "70" End If
If Match = 80 Then m8.text = "80" End If
If Match = 90 Then m9.text = "90" End If
If B2SOn Then
If Match = 0 Then
Controller.B2SSetMatch 34, 100
Else
Controller.B2SSetMatch 34, Match
End If
End If
MatchTimer.enabled = 0
CheckMatch
P1up.text = "" : P2up.text = "" : P3up.text = "" : P4up.text = ""
If B2SOn Then
Controller.B2SSetPlayerUp 30, 0
End If

End Sub

Sub CheckMatch

If PlayerN = 1 Then
	If Match = MatchP(1) Then
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	If Credit<25 Then
	Credit = Credit+1
	DOF 119, DOFOn
	LightCredit.state = 1
	CreditCount.text = (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	End If
	GameOverTimer.enabled = 1
	PlayerN = 0
	Player = 0
	Init= 0
End If

If PlayerN = 2 Then
	If Match = MatchP(1) Then
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	If Credit<25 Then
	Credit = Credit+1
	DOF 119, DOFOn
	LightCredit.state = 1
	CreditCount.text = (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	End If
	MatchTimer2.enabled = 1
End If

If PlayerN = 3 Then
	If Match = MatchP(1) Then
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	If Credit<25 Then
	Credit = Credit+1
	DOF 119, DOFOn
	LightCredit.state = 1
	CreditCount.text = (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	End If
	MatchTimer2.enabled = 1
	MatchTimer3.enabled = 1
End If


If PlayerN = 4 Then
	If Match = MatchP(1) Then
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	If Credit<25 Then
	Credit = Credit+1
	DOF 119, DOFOn
	LightCredit.state = 1
	CreditCount.text = (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	End If
	MatchTimer2.enabled = 1
	MatchTimer3.enabled = 1
	MatchTimer4.enabled = 1
End If

End Sub

Sub MatchTimer2_timer()
	If Match = MatchP(2) Then
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	If Credit<25 Then
	Credit = Credit+1
	DOF 119, DOFOn
	LightCredit.state = 1
	CreditCount.text = (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	End If
	MatchTimer2.enabled = 0
	If PlayerN = 2 Then
	GameOverTimer2.enabled = 1
	PlayerN = 0
	Player = 0
	Init= 0
	End If
End Sub

Sub MatchTimer3_timer()
	If Match = MatchP(3) Then
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	If Credit<25 Then
	DOF 119, DOFOn
	Credit = Credit+1
	LightCredit.state = 1
	CreditCount.text = (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	End If
	MatchTimer3.enabled = 0
	If PlayerN = 3 Then
	GameOverTimer2.enabled = 1
	PlayerN = 0
	Player = 0
	Init= 0
	End If
End Sub

Sub MatchTimer4_timer()
	If Match = MatchP(4) Then
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	If Credit<25 Then
	Credit = Credit+1
	DOF 119, DOFOn
	LightCredit.state = 1
	CreditCount.text = (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	End If
	MatchTimer4.enabled = 0
	GameOverTimer2.enabled = 1
	PlayerN = 0
	Player = 0
	Init= 0
End Sub

'******************************

'********* Advance Bonus Lights & Sounds ***************************

Sub AddBonus(points)
Bonus = Bonus + points

If Bonus = 2 Then L1.state = 0 : L2.state = 1 : L3.state = 0 : L4.state = 0 : L5.state = 0 : L6.state = 0 : L7.state = 0 : L8.state = 0 : L9.state = 0 : L10.state = 0 End If
If Bonus = 3 Then L1.state = 0 : L2.state = 0 : L3.state = 1 : L4.state = 0 : L5.state = 0 : L6.state = 0 : L7.state = 0 : L8.state = 0 : L9.state = 0 : L10.state = 0 End If
If Bonus = 4 Then L1.state = 0 : L2.state = 0 : L3.state = 0 : L4.state = 1 : L5.state = 0 : L6.state = 0 : L7.state = 0 : L8.state = 0 : L9.state = 0 : L10.state = 0 End If
If Bonus = 5 Then L1.state = 0 : L2.state = 0 : L3.state = 0 : L4.state = 0 : L5.state = 1 : L6.state = 0 : L7.state = 0 : L8.state = 0 : L9.state = 0 : L10.state = 0 End If
If Bonus = 6 Then L1.state = 0 : L2.state = 0 : L3.state = 0 : L4.state = 0 : L5.state = 0 : L6.state = 1 : L7.state = 0 : L8.state = 0 : L9.state = 0 : L10.state = 0 End If
If Bonus = 7 Then L1.state = 0 : L2.state = 0 : L3.state = 0 : L4.state = 0 : L5.state = 0 : L6.state = 0 : L7.state = 1 : L8.state = 0 : L9.state = 0 : L10.state = 0 End If
If Bonus = 8 Then L1.state = 0 : L2.state = 0 : L3.state = 0 : L4.state = 0 : L5.state = 0 : L6.state = 0 : L7.state = 0 : L8.state = 1 : L9.state = 0 : L10.state = 0 End If
If Bonus = 9 Then L1.state = 0 : L2.state = 0 : L3.state = 0 : L4.state = 0 : L5.state = 0 : L6.state = 0 : L7.state = 0 : L8.state = 0 : L9.state = 1 : L10.state = 0 End If
If Bonus = 10 Then L1.state = 0 : L2.state = 0 : L3.state = 0 : L4.state = 0 : L5.state = 0 : L6.state = 0 : L7.state = 0 : L8.state = 0 : L9.state = 0 : L10.state = 1 End If
If Bonus = 11 Then L1.state = 1 : L2.state = 0 : L3.state = 0 : L4.state = 0 : L5.state = 0 : L6.state = 0 : L7.state = 0 : L8.state = 0 : L9.state = 0 : L10.state = 1 End If
If Bonus = 12 Then L1.state = 0 : L2.state = 1 : L3.state = 0 : L4.state = 0 : L5.state = 0 : L6.state = 0 : L7.state = 0 : L8.state = 0 : L9.state = 0 : L10.state = 1 End If
If Bonus = 13 Then L1.state = 0 : L2.state = 0 : L3.state = 1 : L4.state = 0 : L5.state = 0 : L6.state = 0 : L7.state = 0 : L8.state = 0 : L9.state = 0 : L10.state = 1 End If
If Bonus = 14 Then L1.state = 0 : L2.state = 0 : L3.state = 0 : L4.state = 1 : L5.state = 0 : L6.state = 0 : L7.state = 0 : L8.state = 0 : L9.state = 0 : L10.state = 1 End If
If Bonus = 15 Then L1.state = 0 : L2.state = 0 : L3.state = 0 : L4.state = 0 : L5.state = 1 : L6.state = 0 : L7.state = 0 : L8.state = 0 : L9.state = 0 : L10.state = 1 End If
If Bonus = 16 Then L1.state = 0 : L2.state = 0 : L3.state = 0 : L4.state = 0 : L5.state = 0 : L6.state = 1 : L7.state = 0 : L8.state = 0 : L9.state = 0 : L10.state = 1 End If
If Bonus = 17 Then L1.state = 0 : L2.state = 0 : L3.state = 0 : L4.state = 0 : L5.state = 0 : L6.state = 0 : L7.state = 1 : L8.state = 0 : L9.state = 0 : L10.state = 1 End If
If Bonus = 18 Then L1.state = 0 : L2.state = 0 : L3.state = 0 : L4.state = 0 : L5.state = 0 : L6.state = 0 : L7.state = 0 : L8.state = 1 : L9.state = 0 : L10.state = 1 End If
If Bonus = 19 Then
L1.state = 0
L2.state = 0
L3.state = 0
L4.state = 0
L5.state = 0
L6.state = 0
L7.state = 0
L8.state = 0
L9.state = 1
L10.state = 1
If Special = 0 And EB = 0 Then
LightExtraBall.state = 1
LightSpecial.state = 1
End If
If Special = 1 Then
If EB = 0 Then LightExtraBall.state = 1 End If
LightSpecial.state = 1
End If
End If

End Sub

'************************************************************

'********* Table Keys ***************************************************************

Sub Wizard_KeyDown(ByVal keycode)

	If keycode = AddCreditKey And Credit<25 Then
	PlaySound "CoinIn"
	Credit = Credit+1
	DOF 119, DOFOn
	LightCredit.state = 1
	CreditCount.text= (Credit)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	If keycode = AddCreditKey And Credit >=25 Then
	PlaySound "CoinIn"
	End If

	If keycode = StartGameKey And Credit>0 And BallsNumber <2 And PlayerN <4 And AllowStart = 1 Then
	PlayerN = PlayerN + 1
	Credit = Credit-1
	DOF 119, DOFOff
	Init = Init + 1
	If Credit <1 Then LightCredit.state = 0 End If
	If Init = 1 Then
	FirstStartTimer.enabled = 1
	StartupTimer.enabled = 1
	End If
	PlaySound "start"
	CreditCount.text= (Credit)
	Game = 1
	BallN= BNumber
	BallN2= BNumber
	BallN3= BNumber
	BallN4= BNumber
	If BallsNumber = 0 Then
	If PlayerN = 1 Then
	BallsNTot = 3
	If Wizard.ShowDT = true Then
	EMReel2.visible = false
	EMReel3.visible = false
	EMReel4.visible = false
	End If
	End If
	If PlayerN = 2 Then
	BallsNTot = 6
	If Wizard.ShowDT = true then
	EMReel2.visible = true
	EMReel3.visible = false
	EMReel4.visible = false
	End If
	End If
	If PlayerN = 3 Then
	BallsNTot = 9
	If Wizard.ShowDT = true then
	EMReel2.visible = true
	EMReel3.visible = true
	EMReel4.visible = false
	End If
	End If
	If PlayerN = 4 Then
	BallsNTot = 12
	If Wizard.ShowDT = true then
	EMReel2.visible = true
	EMReel3.visible = true
	EMReel4.visible = true
	End If
	End If
	End If
	If BallsNumber = 1 Then
	If PlayerN = 1 Then
	BallsNTot = 5
	If Wizard.ShowDT = true then
	EMReel2.visible = false
	EMReel3.visible = false
	EMReel4.visible = false
	End If
	End If
	If PlayerN = 2 Then
	BallsNTot = 10
	If Wizard.ShowDT = true then
	EMReel2.visible = true
	EMReel3.visible = false
	EMReel4.visible = false
	End If
	End If
	If PlayerN = 3 Then
	BallsNTot = 15
	If Wizard.ShowDT = true then
	EMReel2.visible = true
	EMReel3.visible = true
	EMReel4.visible = false
	End If
	End If
	If PlayerN = 4 Then
	BallsNTot = 20
	If Wizard.ShowDT = true then
	EMReel2.visible = true
	EMReel3.visible = true
	EMReel4.visible = true
	End If
	End If
	End If
	m0.text = "" : m1.text = "" : m2.text = "" : m3.text = "" : m4.text = "" : m5.text = "" : m6.text = "" : m7.text = "" : m8.text = "" : m9.text = ""
	MatchP(1) = 0
	MatchP(2) = 0
	MatchP(3) = 0
	MatchP(4) = 0
	If Replay(1) >1 Or Replay(2) >1 Or Replay(3) >1 Or Replay(4) >1 Then
	Replay(1) = 5 : Replay(2) = 5 : Replay(3) = 5 : Replay(4) = 5
	End If
	If Replay(1) >5 Or Replay(2) >5 Or Replay(3) >5 Or Replay(4) >5 Then
	Replay(1) = 10 : Replay(2) = 10 : Replay(3) = 10 : Replay(4) = 10
	End If
	Kicker1.timerenabled = 0
	GameOvertb.text = ""
	Tilttb.text= ""
	PlayerNum.text = (PlayerN)
	If Player = 0 Then
	PlayerPlay.text = ""
	Else
	PlayerPlay.text = (Player)
	End If
	If B2SOn Then
	Controller.B2SSetCredits Credit
	Controller.B2SSetCanPlay 31, PlayerN
	Controller.B2SSetGameOver 0
	Controller.B2SSetMatch 34,0
	End If
	End If

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If

	If keycode = LeftFlipperKey And Game = 1 And Tilt = False Then
		LeftFlipper.RotateToEnd
		PlaySound SoundFXDOF("flup",101,DOFOn,DOFFlippers), 0, .67, -0.05, 0.05
		PlaySound "buzz", -1 , .60, -1
		If LightMod = 1 Then
		Light3.state = 2
		Light3.blinkpattern = 10
		Light3.blinkinterval = 100
		Light4.state = 2
		Light4.blinkpattern = 10
		Light4.blinkinterval = 50
		End If
	End If

	If keycode = RightFlipperKey And Game = 1 And Tilt = False  Then
		RightFlipper.RotateToEnd
		PlaySound SoundFXDOF("flup",102,DOFOn,DOFFlippers), 0, .67, 0.05, 0.05
		Playsound "buzz1", -1 , .60, 1
		If LightMod = 1 Then
		Light3.state = 2
		Light3.blinkpattern = 10
		Light3.blinkinterval = 50
		Light4.state = 2
		Light4.blinkpattern = 10
		Light4.blinkinterval = 100
		End If
	End If

	If keycode = LeftTiltKey Then
		Nudge 90, 2
		Playsound SoundFX("fx_nudge",0)
		CheckTilt
	End If

	If keycode = RightTiltKey Then
		Nudge 270, 2
		Playsound SoundFX("fx_nudge",0)
		CheckTilt
	End If

	If keycode = CenterTiltKey Then
		Nudge 0, 2
		Playsound SoundFX("fx_nudge",0)
		CheckTilt
	End If

End Sub

Sub Wizard_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "fx_plunger",0,1,0.25,0.25

	End If

	If keycode = LeftFlipperKey And Game = 1 Then
		LeftFlipper.RotateToStart
		PlaySound SoundFXDOF("fldown",101,DOFOff,DOFFlippers), 0, 0.10, -0.05, 0.05
		StopSound "buzz"
		If Tilt = true Then
		FlipOff
		End If
		If LightMod = 1 And Tilt = False Then
		Light3.state = 2
		Light3.blinkpattern = 10
		Light3.blinkinterval = 550
		Light4.state = 2
		Light4.blinkpattern = 10
		Light4.blinkinterval = 500
		End If
	End If

	If keycode = RightFlipperKey And Game = 1 Then
		RightFlipper.RotateToStart
		PlaySound SoundFXDOF("fldown",102,DOFOff,DOFFlippers), 0, 0.10, 0.05, 0.05
		StopSound "buzz1"
		If Tilt = true Then
		FlipOff
		End If
		If LightMod = 1 And Tilt = False Then
		Light3.state = 2
		Light3.blinkpattern = 10
		Light3.blinkinterval = 550
		Light4.state = 2
		Light4.blinkpattern = 10
		Light4.blinkinterval = 500
		End If
	End If

End Sub

'**********************************************************************

'**************** Drain *************************

Sub Drain_Hit()
	If Player = 1 Then BallN = BallN -1 End If
	If Player = 2 Then BallN2 = BallN2 -1 End If
	If Player = 3 Then BallN3 = BallN3 -1 End If
	If Player = 4 Then BallN4 = BallN4 -1 End If
	BallsNTot = BallsNTot - 1
	DOF 121, DOFPulse
	PlaySound "drain",0,1,0,0.25
	Drain.DestroyBall
	AllowStart = 0
	EB = EB - 1
	If EB = -1 Then EB = 0

	If BallsNTot = 0 Then
	If Tilt = false Then
	CollectBonusTimer.enabled = 1
	End If
	If Tilt = true Then
	ResetBonusTimer.enabled = 1
	End If
	FlipOff
	Game = 0.5
	End If

	If BallsNtot >0 Then
	If Tilt = false Then
	CollectBonusTimer.enabled = 1
	Light1.state = 1
	Light3.state = 1
	Light4.state = 1
	Light2.state = 1
	End If
	If Tilt = true Then
	ResetBonusTimer.enabled = 1
	End If
	Kicker1.timerenabled = 0
	End If
End Sub

'*****************************************************

'*************** Start & Drain Sequential Events *******************

Sub FirstStartTimer_timer()
	PlaySound "wiz_init"
	FirstStartTimer.enabled = 0
End Sub

Sub Trigger10_hit
	DOF 122, DOFPulse
End Sub

Sub	StartupTimer_timer()
	Light21.state = 0
	Light22.state = 0
	Light31.state = 0
	LightDoubleBonus.state = 0
	Light49.state = 0
	Light51.state = 1
	Light52.state = 1
	Light47.state = 1
	Light46.state = 1
	Light45.state = 1
	Light33.state = 0
	Light34.state = 0
	Light36.state = 0
	Light37.state = 0
	L2.state = 0
	L3.state = 0
	L4.state = 0
	L5.state = 0
	L6.state = 0
	L7.state = 0
	L8.state = 0
	L9.state = 0
	L10.state = 0
	LightSpecial.state = 0
    LightExtraball.state = 0
	LightShootAgain.state = 0
	Light48.state = 1
	Light28.state = 0
	Light27.state = 0
	Light26.state = 0
	Light25.state = 0
	Bonus = 1
	EMR(1).ResetToZero()
	EMR(2).ResetToZero()
	EMR(3).ResetToZero()
	EMR(4).ResetToZero()
	Score(1) = 0
	Score(2) = 0
	Score(3) = 0
	Score(4) = 0
	Score100k1.text = ""
	Score100k2.text = ""
	Score100k3.text = ""
	Score100k4.text = ""
	If B2SOn Then
	Controller.B2SSetScorePlayer 1, Score(1)
	Controller.B2SSetScorePlayer 2, Score(2)
	Controller.B2SSetScorePlayer 3, Score(3)
	Controller.B2SSetScorePlayer 4, Score(4)
	Controller.B2SSetScoreRollover 25, 0
	Controller.B2SSetScoreRollover 26, 0
	Controller.B2SSetScoreRollover 27, 0
	Controller.B2SSetScoreRollover 28, 0
	End If
	StartupTimer.enabled = 0
	SequenceTimer.enabled = 1
	End Sub

Sub	DrainTimer_timer()
	Light21.state = 0
	Light22.state = 0
	Light31.state = 0
	LightDoubleBonus.state = 0
	Light49.state = 0
	Light51.state = 1
	Light52.state = 1
	Light47.state = 1
	Light46.state = 1
	Light45.state = 1
	Light33.state = 0
	Light34.state = 0
	Light36.state = 0
	Light37.state = 0
	L2.state = 0
	L3.state = 0
	L4.state = 0
	L5.state = 0
	L6.state = 0
	L7.state = 0
	L8.state = 0
	L9.state = 0
	L10.state = 0
	LightSpecial.state = 0
    LightExtraball.state = 0
	LightShootAgain.state = 0
	Light48.state = 1
	Light28.state = 0
	Light27.state = 0
	Light26.state = 0
	Light25.state = 0
	Bonus = 1
	DrainTimer.enabled = 0
	SequenceTimer.enabled = 1
	End Sub

Sub SequenceTimer_timer()
	verso = 1
	If Flag1.RotZ = -180 Then Flag1RTimer.enabled = 1
	If Flag2.RotZ = -180 Then Flag2RTimer.enabled = 1
	If Flag3.RotZ = -180 Then Flag3RTimer.enabled = 1
	If Flag4.RotZ = -180 Then Flag4RTimer.enabled = 1
	SequenceTimer.enabled = 0
	Sequence2Timer.enabled = 1
End Sub

Sub Flag1RTimer_timer()
	Const speed = 180
	Flag1.RotZ = Flag1.RotZ + speed*verso
	If Flag1.RotZ >=0 Then
	Flag1.RotZ = 0
	Flag1RTimer.enabled = 0
	End If
	PlaySound SoundFXDOF("flag",127,DOFPulse,DOFContactors)
End Sub

Sub Flag2RTimer_timer()
	Const speed = 180
	Flag2.RotZ = Flag2.RotZ + speed*verso
	If Flag2.RotZ >=0 Then
	Flag2.RotZ = 0
	Flag2RTimer.enabled = 0
	End If
	PlaySound SoundFXDOF("flag",127,DOFPulse,DOFContactors)
End Sub

Sub Flag3RTimer_timer()
	Const speed = 180
	Flag3.RotZ = Flag3.RotZ + speed*verso
	If Flag3.RotZ >=0 Then
	Flag3.RotZ = 0
	Flag3RTimer.enabled = 0
	End If
	PlaySound SoundFXDOF("flag",127,DOFPulse,DOFContactors)
End Sub

Sub Flag4RTimer_timer()
	Const speed = 180
	Flag4.RotZ = Flag4.RotZ + speed*verso
	If Flag4.RotZ >=0 Then
	Flag4.RotZ = 0
	Flag4RTimer.enabled = 0
	End If
	PlaySound SoundFXDOF("flag",127,DOFPulse,DOFContactors)
End Sub

Sub Sequence2Timer_timer()
	L1.state = 1
	If FlashingLEB = 1 Then LEBTimer.enabled = 1 : FlashingLEB = 0 End If
	Sequence2Timer.enabled = 0
	Sequence3Timer.enabled = 1
End Sub

Sub LEBTimer_timer()
	LightShootAgain.state = 2 : LEBTimer.enabled = 0 : LEBoffTimer.enabled = 1
End Sub

Sub LEBoffTimer_timer()
	LightShootAgain.state = 0 : LEBoffTimer.enabled = 0 : SAtb.text = "" : SA2tb.text = ""
	If B2SOn Then
	Controller.B2SSetShootAgain 36,0
	End If
End Sub

Sub Sequence3Timer_timer()
	BallRelease.CreateBall
	BallRelease.Kick 90, 7
	PlaySound SoundFXDOF("ballrelease",120,DOFPulse,DOFContactors)
	Bumper1.force = 12
	Bumper2.force = 12
	Bumper3.force = 12
	Tilt = false
	Tilttb.text = ""
	If B2SOn Then
	Controller.B2SSetTilt 33,0
	End If
	If LightMod = 1 Then
	Light1.state = 2
	Light1.blinkpattern = 10
	Light1.blinkinterval = 550
	Light3.state = 2
	Light3.blinkpattern = 10
	Light3.blinkinterval = 600
	Light4.state = 2
	Light4.blinkpattern = 10
	Light4.blinkinterval = 600
	Light2.state = 2
	Light2.blinkpattern = 10
	Light2.blinkinterval = 500
	End If
	If LightMod = 0 Then
	Light1.state = 1
	Light3.state = 1
	Light4.state = 1
	Light2.state = 1
	End If
	If PlayEB = 0 Then
	If PlayerN = 1 Or Player=PlayerN Then
	Player = 1
	Else
	Player = Player + 1
	End If
	End If
	PlayerPlay.text = (Player)
	If BallsNumber = 0 Then
	If Player = 1 And BallN = 3 Then BallIP = 1 : P1up.text = "1UP" : P2up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 1 End If
	If Player = 1 And BallN = 2 Then BallIP = 2 : P1up.text = "1UP" : P2up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 1 End If
	If Player = 1 And BallN = 1 Then BallIP = 3 : P1up.text = "1UP" : P2up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 1 End If
	If Player = 2 And BallN2 = 3 Then BallIP = 1 : P2up.text = "2UP" : P1up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 2 End If
	If Player = 2 And BallN2 = 2 Then BallIP = 2 : P2up.text = "2UP" : P1up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 2 End If
	If Player = 2 And BallN2 = 1 Then BallIP = 3 : P2up.text = "2UP" : P1up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 2 End If
	If Player = 3 And BallN3 = 3 Then BallIP = 1 : P3up.text = "3UP" : P1up.text = "" : P2up.text = "" : P4up.text = "" : PUP = 3 End If
	If Player = 3 And BallN3 = 2 Then BallIP = 2 : P3up.text = "3UP" : P1up.text = "" : P2up.text = "" : P4up.text = "" : PUP = 3 End If
	If Player = 3 And BallN3 = 1 Then BallIP = 3 : P3up.text = "3UP" : P1up.text = "" : P2up.text = "" : P4up.text = "" : PUP = 3 End If
	If Player = 4 And BallN4 = 3 Then BallIP = 1 : P4up.text = "4UP" : P1up.text = "" : P2up.text = "" : P3up.text = "" : PUP = 4 End If
	If Player = 4 And BallN4 = 2 Then BallIP = 2 : P4up.text = "4UP" : P1up.text = "" : P2up.text = "" : P3up.text = "" : PUP = 4 End If
	If Player = 4 And BallN4 = 1 Then BallIP = 3 : P4up.text = "4UP" : P1up.text = "" : P2up.text = "" : P3up.text = "" : PUP = 4 End If
	End If
	If BallsNumber = 1 Then
	If Player = 1 And BallN = 5 Then BallIP = 1 : P1up.text = "1UP" : P2up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 1 End If
	If Player = 1 And BallN = 4 Then BallIP = 2 : P1up.text = "1UP" : P2up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 1 End If
	If Player = 1 And BallN = 3 Then BallIP = 3 : P1up.text = "1UP" : P2up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 1 End If
	If Player = 1 And BallN = 2 Then BallIP = 4 : P1up.text = "1UP" : P2up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 1 End If
	If Player = 1 And BallN = 1 Then BallIP = 5 : P1up.text = "1UP" : P2up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 1 End If
	If Player = 2 And BallN2 = 5 Then BallIP = 1 : P2up.text = "2UP" : P1up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 2 End If
	If Player = 2 And BallN2 = 4 Then BallIP = 2 : P2up.text = "2UP" : P1up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 2 End If
	If Player = 2 And BallN2 = 3 Then BallIP = 3 : P2up.text = "2UP" : P1up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 2 End If
	If Player = 2 And BallN2 = 2 Then BallIP = 4 : P2up.text = "2UP" : P1up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 2 End If
	If Player = 2 And BallN2 = 1 Then BallIP = 5 : P2up.text = "2UP" : P1up.text = "" : P3up.text = "" : P4up.text = "" : PUP = 2 End If
	If Player = 3 And BallN3 = 5 Then BallIP = 1 : P3up.text = "3UP" : P1up.text = "" : P2up.text = "" : P4up.text = "" : PUP = 3 End If
	If Player = 3 And BallN3 = 4 Then BallIP = 2 : P3up.text = "3UP" : P1up.text = "" : P2up.text = "" : P4up.text = "" : PUP = 3 End If
	If Player = 3 And BallN3 = 3 Then BallIP = 3 : P3up.text = "3UP" : P1up.text = "" : P2up.text = "" : P4up.text = "" : PUP = 3 End If
	If Player = 3 And BallN3 = 2 Then BallIP = 4 : P3up.text = "3UP" : P1up.text = "" : P2up.text = "" : P4up.text = "" : PUP = 3 End If
	If Player = 3 And BallN3 = 1 Then BallIP = 5 : P3up.text = "3UP" : P1up.text = "" : P2up.text = "" : P4up.text = "" : PUP = 3 End If
	If Player = 4 And BallN4 = 5 Then BallIP = 1 : P4up.text = "4UP" : P1up.text = "" : P2up.text = "" : P3up.text = "" : PUP = 4 End If
	If Player = 4 And BallN4 = 4 Then BallIP = 2 : P4up.text = "4UP" : P1up.text = "" : P2up.text = "" : P3up.text = "" : PUP = 4 End If
	If Player = 4 And BallN4 = 3 Then BallIP = 3 : P4up.text = "4UP" : P1up.text = "" : P2up.text = "" : P3up.text = "" : PUP = 4 End If
	If Player = 4 And BallN4 = 2 Then BallIP = 4 : P4up.text = "4UP" : P1up.text = "" : P2up.text = "" : P3up.text = "" : PUP = 4 End If
	If Player = 4 And BallN4 = 1 Then BallIP = 5 : P4up.text = "4UP" : P1up.text = "" : P2up.text = "" : P3up.text = "" : PUP = 4 End If
	End If
	BallCount.text = (BallIP)
	If B2SOn Then
	Controller.B2SSetPlayerUp 30, PUP
	Controller.B2SSetBallinPlay 32, BallIP
	End If
	PlayEB = 0
	Sequence3Timer.enabled = 0
End Sub

Sub CollectBonusTimer_timer()
	If Bonus = 1 Then L1Timer.enabled = 1 End If
	If Bonus = 2 Then L2Timer.enabled = 1 End If
	If Bonus = 3 Then L3Timer.enabled = 1 End If
	If Bonus = 4 Then L4Timer.enabled = 1 End If
	If Bonus = 5 Then L5Timer.enabled = 1 End If
	If Bonus = 6 Then L6Timer.enabled = 1 End If
	If Bonus = 7 Then L7Timer.enabled = 1 End If
	If Bonus = 8 Then L8Timer.enabled = 1 End If
	If Bonus = 9 Then L9Timer.enabled = 1 End If
	If Bonus = 10 Then L10Timer.enabled = 1 End If
	If Bonus = 11 Then L11Timer.enabled = 1 End If
	If Bonus = 12 Then L12Timer.enabled = 1 End If
	If Bonus = 13 Then L13Timer.enabled = 1 End If
	If Bonus = 14 Then L14Timer.enabled = 1 End If
	If Bonus = 15 Then L15Timer.enabled = 1 End If
	If Bonus = 16 Then L16Timer.enabled = 1 End If
	If Bonus = 17 Then L17Timer.enabled = 1 End If
	If Bonus = 18 Then L18Timer.enabled = 1 End If
	If Bonus = 19 Then L19Timer.enabled = 1 End If
	CollectBonusTimer.enabled = 0
End Sub

Sub GameOverTimer_timer()
GameOvertb.text = "GAME OVER"
Game = 0
AllowStart = 1
If B2SOn Then
Controller.B2SSetGameOver 35,1
End If
GameOverTimer.enabled = 0
End Sub

Sub GameOverTimer2_timer()
GameOvertb.text = "GAME OVER"
Game = 0
AllowStart = 1
If B2SOn Then
Controller.B2SSetGameOver 35,1
End If
GameOverTimer2.enabled = 0
End Sub

Sub DBTimer_timer()
	AddScore (1000)
	DBTimer.enabled = 0
End Sub

Sub L1Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L1.state = 0 : L1Timer.enabled = 0
	If BallsNTot > 0 Then DrainTimer.enabled = 1 End If
	If Game = 0.5 And Special = 1 Then MatchTimer.enabled = 1
	If BallsNtoT = 0 And PlayerN = 1 Then GameOverTimer.enabled = 1 End If
End Sub

Sub L2Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L2.state = 0 : L1.state = 1 : L2Timer.enabled = 0 : L1Timer.enabled = 1
End Sub

Sub L3Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L3.state = 0 : L2.state = 1 : L3Timer.enabled = 0 : L2Timer.enabled = 1
End Sub

Sub L4Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
    L4.state = 0 : L3.state = 1 : L4Timer.enabled = 0 : L3Timer.enabled = 1
End Sub

Sub L5Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L5.state = 0 : L4.state = 1 : L5Timer.enabled = 0 : L4Timer.enabled = 1
End Sub

Sub L6Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L6.state = 0 : L5.state = 1 : L6Timer.enabled = 0 : L5Timer.enabled = 1
End Sub

Sub L7Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L7.state = 0 : L6.state = 1 : L7Timer.enabled = 0 : L6Timer.enabled = 1
End Sub

Sub L8Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L8.state = 0 : L7.state = 1 : L8Timer.enabled = 0 : L7Timer.enabled = 1
End Sub

Sub L9Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L9.state = 0 : L8.state = 1 : L9Timer.enabled = 0 : L8Timer.enabled = 1
End Sub

Sub L10Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L10.state = 0 : L9.state = 1 : L10Timer.enabled = 0 : L9Timer.enabled = 1
End Sub

Sub L11Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L1.state = 0  : L11Timer.enabled = 0 : L10Timer.enabled = 1
End Sub

Sub L12Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L2.state = 0 : L1.state = 1 :L12Timer.enabled = 0 : L11Timer.enabled = 1
End Sub

Sub L13Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L3.state = 0 : L2.state = 1 : L13Timer.enabled = 0 : L12Timer.enabled = 1
End Sub

Sub L14Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L4.state = 0 : L3.state = 1 : L14Timer.enabled = 0 : L13Timer.enabled = 1
End Sub

Sub L15Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L5.state = 0 : L4.state = 1 : L15Timer.enabled = 0 : L14Timer.enabled = 1
End Sub

Sub L16Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L6.state = 0 : L5.state = 1 : L16Timer.enabled = 0 : L15Timer.enabled = 1
End Sub

Sub L17Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L7.state = 0 : L6.state = 1 : L17Timer.enabled = 0 : L16Timer.enabled = 1
End Sub

Sub L18Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L8.state = 0 : L7.state = 1 : L18Timer.enabled = 0 : L17Timer.enabled = 1
End Sub

Sub L19Timer_timer()
	PlaySound "bonus"
	AddScore (1000)
	If LightDoubleBonus.state = 1 Then DBTimer.enabled = 1 End If
	L9.state = 0 : L8.state = 1 : L19Timer.enabled = 0 : L18Timer.enabled = 1
End Sub

Sub ResetBonusTimer_timer()
	If Bonus = 1 Then L1rTimer.enabled = 1 End If
	If Bonus = 2 Then L2rTimer.enabled = 1 End If
	If Bonus = 3 Then L3rTimer.enabled = 1 End If
	If Bonus = 4 Then L4rTimer.enabled = 1 End If
	If Bonus = 5 Then L5rTimer.enabled = 1 End If
	If Bonus = 6 Then L6rTimer.enabled = 1 End If
	If Bonus = 7 Then L7rTimer.enabled = 1 End If
	If Bonus = 8 Then L8rTimer.enabled = 1 End If
	If Bonus = 9 Then L9rTimer.enabled = 1 End If
	If Bonus = 10 Then L10rTimer.enabled = 1 End If
	If Bonus = 11 Then L11rTimer.enabled = 1 End If
	If Bonus = 12 Then L12rTimer.enabled = 1 End If
	If Bonus = 13 Then L13rTimer.enabled = 1 End If
	If Bonus = 14 Then L14rTimer.enabled = 1 End If
	If Bonus = 15 Then L15rTimer.enabled = 1 End If
	If Bonus = 16 Then L16rTimer.enabled = 1 End If
	If Bonus = 17 Then L17rTimer.enabled = 1 End If
	If Bonus = 18 Then L18rTimer.enabled = 1 End If
	If Bonus = 19 Then L19rTimer.enabled = 1 End If
	ResetBonusTimer.enabled = 0
End Sub

Sub L1rTimer_timer()
	PlaySound "bonus"
	L1.state = 0 : L1rTimer.enabled = 0
	If BallsNTot > 0 Then DrainTimer.enabled = 1 End If
	If Game = 0.5 And Special = 1 Then MatchTimer.enabled = 1
	If BallsNtoT = 0 And PlayerN = 1 Then GameOverTimer.enabled = 1 End If
End Sub

Sub L2rTimer_timer()
	PlaySound "bonus"
	L2.state = 0 : L1.state = 1 : L2rTimer.enabled = 0 : L1rTimer.enabled = 1
End Sub

Sub L3rTimer_timer()
	PlaySound "bonus"
	L3.state = 0 : L2.state = 1 : L3rTimer.enabled = 0 : L2rTimer.enabled = 1
End Sub

Sub L4rTimer_timer()
	PlaySound "bonus"
    L4.state = 0 : L3.state = 1 : L4rTimer.enabled = 0 : L3rTimer.enabled = 1
End Sub

Sub L5rTimer_timer()
	PlaySound "bonus"
	L5.state = 0 : L4.state = 1 : L5rTimer.enabled = 0 : L4rTimer.enabled = 1
End Sub

Sub L6rTimer_timer()
	PlaySound "bonus"
	L6.state = 0 : L5.state = 1 : L6rTimer.enabled = 0 : L5rTimer.enabled = 1
End Sub

Sub L7rTimer_timer()
	PlaySound "bonus"
	L7.state = 0 : L6.state = 1 : L7rTimer.enabled = 0 : L6rTimer.enabled = 1
End Sub

Sub L8rTimer_timer()
	PlaySound "bonus"
	L8.state = 0 : L7.state = 1 : L8rTimer.enabled = 0 : L7rTimer.enabled = 1
End Sub

Sub L9rTimer_timer()
	PlaySound "bonus"
	L9.state = 0 : L8.state = 1 : L9rTimer.enabled = 0 : L8rTimer.enabled = 1
End Sub

Sub L10rTimer_timer()
	PlaySound "bonus"
	L10.state = 0 : L9.state = 1 : L10rTimer.enabled = 0 : L9rTimer.enabled = 1
End Sub

Sub L11rTimer_timer()
	PlaySound "bonus"
	L1.state = 0  : L11rTimer.enabled = 0 : L10rTimer.enabled = 1
End Sub

Sub L12rTimer_timer()
	PlaySound "bonus"
	L2.state = 0 : L1.state = 1 :L12rTimer.enabled = 0 : L11rTimer.enabled = 1
End Sub

Sub L13rTimer_timer()
	PlaySound "bonus"
	L3.state = 0 : L2.state = 1 : L13rTimer.enabled = 0 : L12rTimer.enabled = 1
End Sub

Sub L14rTimer_timer()
	PlaySound "bonus"
	L4.state = 0 : L3.state = 1 : L14rTimer.enabled = 0 : L13rTimer.enabled = 1
End Sub

Sub L15rTimer_timer()
	PlaySound "bonus"
	L5.state = 0 : L4.state = 1 : L15rTimer.enabled = 0 : L14rTimer.enabled = 1
End Sub

Sub L16rTimer_timer()
	PlaySound "bonus"
	L6.state = 0 : L5.state = 1 : L16rTimer.enabled = 0 : L15rTimer.enabled = 1
End Sub

Sub L17rTimer_timer()
	PlaySound "bonus"
	L7.state = 0 : L6.state = 1 : L17rTimer.enabled = 0 : L16rTimer.enabled = 1
End Sub

Sub L18rTimer_timer()
	PlaySound "bonus"
	L8.state = 0 : L7.state = 1 : L18rTimer.enabled = 0 : L17rTimer.enabled = 1
End Sub

Sub L19rTimer_timer()
	PlaySound "bonus"
	L9.state = 0 : L8.state = 1 : L19rTimer.enabled = 0 : L18rTimer.enabled = 1
End Sub

'***********************************************************************

'************ Tilt **********************************

Sub TiltTimer_Timer()
	TiltTimer.Enabled = False
End Sub

Sub CheckTilt
	If TiltTimer.Enabled = True And Game = 1 Then
	TiltCount = TiltCount + 1
	If TiltCount = 3 Then
	Bumper1.force = 0
	Bumper2.force = 0
	Bumper3.force = 0
	Tilt = True
	PlaySound "tilt"
	FlipOff
	Tilttb.text= "TILT"
	If B2SOn Then
	Controller.B2SSetTilt 33,1
	End If
	End If
	Else
	TiltCount = 0
	TiltTimer.Enabled = True
	End If
End Sub

Sub FlipOff
	LeftFlipper.rotatetostart
	RightFlipper.rotatetostart
	StopSound "buzz"
	StopSound "fldown"
	Light1.state = 0
	Light2.state = 0
	Light3.state = 0
	Light4.state = 0
End Sub

'****************************************************

'********** Sling Shot Animations **********************************************

Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySound SoundFXDOF("slingR",104,DOFPulse,DOFContactors), 0, 0.70, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	If Tilt = false Then
	Addscore (10)
	End If

End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFXDOF("slingL",103,DOFPulse,DOFContactors),0,0.70,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	If Tilt = false Then
	Addscore (10)
	End If

End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1

End Sub

'*******************************************************************************

'********** Bumpers **********************************************

Dim dirRing : dirRing = -1
Dim dirRing2 : dirRing2 = -1
Dim dirRing3 : dirRing3 = -1

Sub Bumper1_Hit
If Tilt = false Then
Me.TimerEnabled = 1
	PlaySound SoundFXDOF("fx_bumper2",106,DOFPulse,DOFContactors),0,0.70
	AddScore (100)
End If

End Sub

Sub Bumper1_timer()
	If Tilt = false Then
	BumperRing1.Z = BumperRing1.Z + (5 * dirRing)
	If BumperRing1.Z <= 0 Then dirRing = 1
	If BumperRing1.Z >= 40 Then
		dirRing = -1
		BumperRing1.Z = 40
		Me.TimerEnabled = 0
	End If
	End If

End Sub

Sub Bumper2_Hit
If Tilt = false Then
Me.TimerEnabled = 1
	PlaySound SoundFXDOF("fx_bumper2",105,DOFPulse,DOFContactors),0,0.70
	If Light21.state = 0 Then
	Addscore (10)
	End If
	If Light21.state = 1 Then
	Addscore (100)
End If
End If

End Sub

Sub Bumper2_timer()
	If Tilt = false Then
	BumperRing2.Z = BumperRing2.Z + (5 * dirRing2)
	If BumperRing2.Z <= 0 Then dirRing2 = 1
	If BumperRing2.Z >= 40 Then
		dirRing2 = -1
		BumperRing2.Z = 40
		Me.TimerEnabled = 0
	End If
	End If

End Sub

Sub Bumper3_Hit
If Tilt = false Then
Me.TimerEnabled = 1
	PlaySound SoundFXDOF("fx_bumper2",107,DOFPulse,DOFContactors),0,0.70
	If Light22.state = 0 Then
	Addscore (10)
	End If
	If Light22.state = 1 Then
	Addscore (100)
	End If
	End If

End Sub

Sub Bumper3_timer()
	If Tilt = false Then
	BumperRing3.Z = BumperRing3.Z + (5 * dirRing3)
	If BumperRing3.Z <= 0 Then dirRing3 = 1
	If BumperRing3.Z >= 40 Then
		dirRing3 = -1
		BumperRing3.Z = 40
		Me.TimerEnabled = 0
	End If
	End If

End Sub

'*****************************************************************************

'******** Saucer *************************************************************

Sub Kicker1_hit
	me.timerenabled=1
	PlaySound "kicker_enter_center"
	SaucerTimer.enabled = 1
End Sub

Sub SaucerTimer_timer()
	If Tilt = false Then
	AddScore (1000)
	P3000.enabled = 1 '3000 points
	If Bonus <19 Then
	AddBonus (+1)
	TimerKickBonus.enabled = 1
	End If
	End If
	SaucerTimer.enabled = 0
End Sub

Sub Kicker1_timer
	PlaySound SoundFXDOF("popper_ball",112,DOFPulse,DOFContactors)
	Kicker1.kick 120,15
	Kicker1.timerenabled=0
End Sub

Sub TimerKickBonus_Timer
	If Bonus <19 Then
	AddBonus (+1)
	End If
	TimerKickBonus.enabled = 0
	TimerKickBonus2.enabled = 1
End Sub

Sub TimerKickBonus2_Timer
	If Bonus <19 Then
	AddBonus (+1)
	End If
	TimerKickBonus2.enabled = 0
End Sub

'*****************************************************************************

'********** Hit Events, Scores & Sounds ******************************************

Sub LeftOutlane_Hit
	DOF 114, DOFPulse
	PlaySound "fx_trigger"
	If Tilt = false Then
	Addscore (1000)
	End If
End Sub

Sub RightOutlane_Hit
	DOF 117, DOFPulse
	PlaySound "fx_trigger"
	If Tilt = false Then
	Addscore (1000)
	End If
End Sub

Sub Leftlane_Hit
	DOF 115, DOFPulse
	PlaySound "fx_trigger"
	If Tilt = false Then
	If Bonus <19 Then AddBonus (+1) End If
	Addscore (100) '500 points
	P500.enabled = 1
	End If
End Sub

Sub Rightlane_Hit
	DOF 116, DOFPulse
	PlaySound "fx_trigger"
	If Tilt = false Then
	If Bonus <19 Then AddBonus (+1) End If
	Addscore (100) '500 points
	P500.enabled = 1
	End If
End Sub

Sub Gate1_Hit
	PlaySound "gate"
End Sub

Sub Gate3_Hit
	PlaySound "gate"
End Sub

Sub TargetF1_Hit
	PlaySound SoundFXDOF("target",108,DOFPulse,DOFTargets)
	If Tilt = false Then
	If Light47.state = 1 Then
	If Bonus <19 Then AddBonus (+1) End If
	Flag1Timer.enabled = 1
	verso = -1
	End If
	Addscore (100)
	Light28.state = 1
	Light51.state = 0
	Light47.state = 0
	End If
End Sub

Sub TargetF2_Hit
	PlaySound SoundFXDOF("target",108,DOFPulse,DOFTargets)
	If Tilt = false Then
	If Light46.state = 1 Then
	If Bonus <19 Then AddBonus (+1) End If
	Flag2Timer.enabled = 1
	verso = -1
	End If
	Addscore (100)
	Light27.state = 1
	Light46.state = 0
	End If
End Sub

Sub TargetF3_Hit
	PlaySound SoundFXDOF("target",108,DOFPulse,DOFTargets)
	If Tilt = false Then
	If Light45.state = 1 Then
	If Bonus <19 Then AddBonus (+1) End If
	Flag3Timer.enabled = 1
	verso = -1
	End If
	Addscore (100)
	Light26.state = 1
	Light45.state = 0
	End If
End Sub

Sub TargetF4_Hit
	PlaySound SoundFXDOF("target",109,DOFPulse,DOFTargets)
	If Tilt = false Then
	If Light48.state = 1 Then
	If Bonus <19 Then AddBonus (+1) End If
	Flag4Timer.enabled = 1
	verso = -1
	End If
	Addscore (100)
	Light25.state = 1
	Light48.state = 0
	Light52.state = 0
	End If

End Sub

Sub TargetSpecial_Hit
	PlaySound SoundFXDOF("target",109,DOFPulse,DOFTargets)
	If LightSpecial.state=0 And Tilt = False Then
	Addscore (1000) '5000 points
	P5000.enabled = 1
	End If
	If LightSpecial.state=1 And Tilt = False Then
	PlaySound "target"
	Addscore (1000) '5000 points
	P5000.enabled = 1
	If Special = 1 Then
	If Credit <25 Then
	Credit= Credit+1
	DOF 119, DOFOn
	CreditCount.text= (Credit)
	LightCredit.state = 1
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
	End If
	End If
	If Special = 0 Then
	If Player = 1 Then BallN = BallN + 1 End If
	If Player = 2 Then BallN2 = BallN2 + 1 End If
	If Player = 3 Then BallN3 = BallN3 + 1 End If
	If Player = 4 Then BallN4 = BallN4 + 1 End If
	BallsNTot = BallsNTot + 1
	PlayEB = 1
	EB = 2
	LightExtraBall.state = 0
	LightShootAgain.state = 1
	SAtb.text = " SAME PLAYER"
	SA2tb.text = "SHOOTS AGAIN"
	If B2SOn Then
	Controller.B2SSetShootAgain 36,1
	End If
	FlashingLEB = 1
	End If
	LightSpecial.state= 0
	End If

End Sub

Sub Target300_Hit
	PlaySound SoundFXDOF("target",110,DOFPulse,DOFTargets)
	If Tilt = false Then
	If Bonus <19 Then
	AddBonus (+1)
	End If
	If Light49.state = 0 Then
	Addscore (100) '300 points
	P300.enabled = 1
	End If
	If Light49.state=1 Then
	Addscore (1000) '3000 points
	P3000.enabled = 1
	End If
	End If
End Sub

Sub Target1000d_Hit
	PlaySound SoundFXDOF("target",111,DOFPulse,DOFTargets)
	If Tilt = false Then
	If Bonus <19 Then AddBonus (+1) End If
	Addscore (1000)
	End If
End Sub

Sub Target1000u_Hit
	PlaySound SoundFXDOF("target",111,DOFPulse,DOFTargets)
	If Tilt = false Then
	If Bonus <19 Then AddBonus (+1) End If
	Addscore (1000)
	End If
End Sub

Sub Gate2_Hit
	PlaySound "gate"
End Sub

Sub Spinner1_Spin()
	DOF 113, DOFPulse
	PlaySound "fx_spinner"
	If Tilt = false Then
	If Light31.state = 0 Then
	Addscore (10)
	End If
	If Light31.state = 1 Then
	Addscore (100)
End If
End If
End Sub

Sub Trigger2_Hit
	PlaySound "fx_trigger"
	If Tilt = false Then
	Addscore (10)
	End If
End Sub

Sub Trigger3_Hit
	PlaySound "fx_trigger"
	If Tilt = false Then
	Addscore (10)
	End If
End Sub

Sub Trigger4_Hit
	PlaySound "fx_trigger"
	If Tilt = false Then
	If Light51.state = 1 Then
	If Bonus <19 Then AddBonus (+1)	End If
	Flag1Timer.enabled = 1
	verso = -1
	End If
	Light51.state = 0
	Light47.state = 0
	Light28.state = 1
	Addscore (100)
	End If
End Sub

Sub Trigger5_Hit
	PlaySound "fx_trigger"
	If Tilt = false Then
	If Light52.state = 1 Then
	If Bonus <19 Then AddBonus (+1) End If
	Flag4Timer.enabled = 1
	verso = -1
	End If
	Light52.state = 0
	Light48.state = 0
	Light25.state = 1
	Addscore (100)
	End If
End Sub

Sub TriggerF1_Hit
	PlaySound "fx_trigger"
	DOF 126, DOFPulse
	If Tilt = false Then
	If Light28.state = 1 Then
	Light21.state = 1
	Light22.state = 1
	Light33.state = 1
	Light34.state = 1
	Light36.state = 1
	Light37.state = 1
	Flag1Timer.enabled = 1
	verso = 1
	If Bonus <19 Then AddBonus (+1) End If
	Addscore (100)
	End If
	Light28.state = 0
	Light51.state = 1
	Light47.state = 1
	If Light28.state = 0 Then
	Addscore (100)
	End If
	End If

End Sub

Sub TriggerF2_Hit
	PlaySound "fx_trigger"
	DOF 125, DOFPulse
	If Tilt = false Then
	If Light27.state = 1 Then
	Light49.state = 1
	Flag2Timer.enabled = 1
	verso = 1
	If Bonus <19 Then AddBonus (+1) End If
	Addscore (100)
	End If
	Light27.state = 0
	Light46.state = 1
	If Light28.state = 0 Then
	Addscore (100)
	End If
	End If

End Sub

Sub TriggerF3_Hit
	PlaySound "fx_trigger"
	DOF 124, DOFPulse
	If Tilt = false Then
	If Light26.state = 1 Then
	LightDoubleBonus.state = 1
	Flag3Timer.enabled = 1
	verso = 1
	If Bonus <19 Then AddBonus (+1) End If
	Addscore (100)
	End If
	Light26.state = 0
	Light45.state = 1
	If Light28.state = 0 Then
	Addscore (100)
	End If
	End If

End Sub

Sub TriggerF4_Hit
	PlaySound "fx_trigger"
	DOF 123, DOFPulse
	If Tilt = false Then
	If Light25.state = 1 Then
	Flag4Timer.enabled = 1
	verso = 1
	Light31.state = 1
	If Bonus <19 Then AddBonus (+1) End If
	Addscore (100)
	End If
	Light25.state = 0
	Light52.state = 1
	Light48.state = 1
	If Light28.state = 0 Then
	Addscore (100)
	End If
	End If

End Sub

Sub extraball_Hit
	If LightExtraBall.state = 1 And Tilt = False Then
	FlashingLEB = 1
	If Player = 1 Then BallN = BallN + 1 End If
	If Player = 2 Then BallN2 = BallN2 + 1 End If
	If Player = 3 Then BallN3 = BallN3 + 1 End If
	If Player = 4 Then BallN4 = BallN4 + 1 End If
	BallsNTot = BallsNTot + 1
	PlayEB = 1
	EB = 2
	PlaySound SoundfXDOF("knocker",118,DOFPulse,DOFKnocker)
	LightExtraBall.state = 0
	LightShootAgain.state = 1
	SAtb.text = " SAME PLAYER" : SA2tb.text = "SHOOTS AGAIN"
	If B2SOn Then
	Controller.B2SSetShootAgain 36,1
	End If
	If Special = 0 Then LightSpecial.state = 0 End If
	End If

End Sub

'*************************************************

'************ Flags Animation *****************************

Const speed = 20
Dim verso

Sub Flag1Timer_timer()
	Flag1.RotZ = Flag1.RotZ + speed*verso
	If Flag1.RotZ <= -180 Then
	Flag1.RotZ = -180
	Flag1Timer.enabled = 0
	Playsound SoundFXDOF("flag",127,DOFPulse,DOFContactors)
	End If
	If Flag1.RotZ >=0 Then
	Flag1.RotZ = 0
	Flag1Timer.enabled = 0
	Playsound SoundFXDOF("flag",127,DOFPulse,DOFContactors)
	End If
End Sub

Sub Flag2Timer_timer()
	Flag2.RotZ = Flag2.RotZ + speed*verso
	If Flag2.RotZ <= -180 Then
	Flag2.RotZ = -180
	Flag2Timer.enabled = 0
	Playsound SoundFXDOF("flag",127,DOFPulse,DOFContactors)
	End If
	If Flag2.RotZ >= 0 Then
	Flag2.RotZ = 0
	Flag2Timer.enabled = 0
	Playsound SoundFXDOF("flag",127,DOFPulse,DOFContactors)
	End If

End Sub

Sub Flag3Timer_timer()
	Flag3.RotZ = Flag3.RotZ + speed*verso
	If Flag3.RotZ <= -180 Then
	Flag3.RotZ = -180
	Flag3Timer.enabled = 0
	Playsound SoundFXDOF("flag",127,DOFPulse,DOFContactors)
	End If
	If Flag3.RotZ >= 0 Then
	Flag3.RotZ = 0
	Flag3Timer.enabled = 0
	Playsound SoundFXDOF("flag",127,DOFPulse,DOFContactors)
	End If
End Sub

Sub Flag4Timer_timer()
	Flag4.RotZ = Flag4.RotZ + speed*verso
	If Flag4.RotZ <= -180 Then
	Flag4.RotZ = -180
	Flag4Timer.enabled = 0
	Playsound SoundFXDOF("flag",127,DOFPulse,DOFContactors)
	End If
	If Flag4.RotZ >= 0 Then
	Flag4.RotZ = 0
	Flag4Timer.enabled = 0
	Playsound SoundFXDOF("flag",127,DOFPulse,DOFContactors)
	End If
End Sub

'**************************************************************************************

'**************** Apron Color *********************

If ApronColor = 3 then ApronColor = Int(Rnd*2)+1 End If
Select Case ApronColor
    Case 1 : Apron.image = "apron wiz" : Plungcover.image = "apron wiz"
    Case 2 : Apron.image = "apron wiz2" : Plungcover.image = "apron wiz2"
End Select

'**************************************************

'**************** Flippers Color ******************

If FlippersColor = 3 then FlippersColor = Int(Rnd*2)+1 End If
Select Case FlippersColor
    Case 1 : LeftFlipper.image = "yellowflipper" : RightFlipper.image = "yellowflipper"
    Case 2 : LeftFlipper.image = "redflipper" : RightFlipper.image = "redflipper"
End Select

'**************************************************

'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc

'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.

'************** Others Table Sounds *****************

Sub Rubbers_Hit(idx):PlaySound "rubber_hit_1", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Rubberpegs_Hit (idx):PlaySound "rubber_hit_3", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Wireguides_Hit (idx):PlaySound "metalhit_thin", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Flatmetalguides_Hit (idx):PlaySound "metalhit_medium", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Trigger6_Hit:PlaySound "metalhit_medium" End Sub
Sub Trigger7_Hit:PlaySound "metalhit_medium" End Sub
Sub Trigger8_Hit:PlaySound "metalhit_medium" End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber2", 0, parm / 10, -0.1, 0.15
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber2", 0, parm / 10, 0.1, 0.15
End Sub

Sub Wall6_Hit
PlaySound "metalhit_thin"
End Sub

Sub Wall7_Hit
PlaySound "metalhit_thin"
End Sub

Sub Wall35_Hit
PlaySound "metalhit_thin"
End Sub

Sub Wall75_Hit
PlaySound "metalhit_thin"
End Sub

Sub Wall27_Hit
PlaySound "metalhit_thin"
End Sub

Sub Wall67_Hit
PlaySound "metalhit_thin"
End Sub

Sub Wall58_Hit
PlaySound "metalhit_thin"
End Sub

Sub Wall70_Hit
PlaySound "metalhit_thin"
End Sub

Sub Wall71_Hit
PlaySound "metalhit_thin"
End Sub

Sub Wall31_Hit
PlaySound "metalhit_medium"
End Sub

Sub Wall1_Hit
PlaySound "metalhit_thin"
End Sub

Sub Wall59_Hit
PlaySound "metalhit_thin"
End Sub

Sub Trigger9_Hit
If BallsNTot = 1 Then
FlipOff
End If
End Sub

Sub Wall72_Hit
PlaySound "rubber_hit_1"
End Sub

Sub Wall73_Hit
PlaySound "rubber_hit_1"
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Wizard" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Wizard.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Wizard" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Wizard.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Wizard" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Wizard.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Wizard.height-1
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
  If Wizard.VersionMinor > 3 OR Wizard.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub Wizard_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

