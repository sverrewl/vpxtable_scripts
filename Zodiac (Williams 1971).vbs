' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Willimas' "Zodiac" (1971)
' Version: 1.0.0
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

ExecuteGlobal GetTextFile("core.vbs") : If Err Then MsgBox "Can't open ""core.vbs"""


Const cGameName = "WilliamsZodiac"


'***********  Game Options  **********

Const ballspergame = 5


Dim Replay1:Replay1 = 68000
Dim Replay2:Replay2 = 81000
Dim Replay3:Replay3 = 94000

ReplayCard.image = "SC51"

instructCard.image="IC1"

Ballmass = 1.2      'was 1.0

Dim twoposts
twoposts = false            ' single post disc (False) or two posts (True)

'******************************************


Dim BIP, Balls
Dim PlayerScores(2)
dim prevscore(2)
Dim currentp, Players, Credits
Dim Objekt, object

Dim Tilt, tiltsens
Dim moveballs
Dim ballangle(15), direction(15)

Dim hiscore

dim FileObj
dim ScoreFile
dim TextStr
dim temp1, temp2, temp3, temp4, temp5, temp6, temp7, temp8, temp9, temp10, temp11, temp12

dim Score(2)

dim player1100k, player2100k
dim point
dim scn, scn1, scn2, scn3
dim bell
dim currpl, rst

Dim tenk(4), onek(4), hun(4), tens(4)
Dim starsign

Dim BallSize: BallSize  = 22

Dim BallMass
Dim mMagnetS, mMagnetM

Dim SBonus, SBonusMultiplier
Dim MBonus , MBonusMultiplier

Dim curPos, lastPos         'for spinning disc




'**** set backglass flashers

Sub SetBackglass()
  Dim obj

  For Each obj In Backglass_flashers
    obj.x = obj.x - 10
    obj.height = - obj.y -100
    obj.y = 40 'adjusts the distance from the backglass towards the user
  Next

End Sub

' ***************************************************************************

' BASIC FSS(EM) 1-4 player 5x drums, 1 credit drum CORE CODE

' ****************************************************************************

 '**********  global variables and constants

Dim ix, np,scores(4,2)

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen , objs


xoff =  463
yoff = 0
zoff =  735
xrot = -90

Const USEEMS = 2 ' 1-4 set between 1 to 4 based on number of players

const idx_emp1r1 =0 'player 1
const idx_emp2r1 =5 'player 2
const idx_emp3r1 =10 'player 3
const idx_emp4r1 =15 'player 4
const idx_emp4r6 =20 'credits

Dim BGObjEM(1)

Dim EMMODE

EMMODE = 1  'set to 1 if players score is tracked outside of the EM reels subroutine. set to 0 if a specific score broken out by digit is to be displayed

Dim nplayer,playr,value,curscr,curplayr


' ********************* POSITION EM REEL DRUMS ON BACKGLASS *************************

if USEEMS = 1 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
elseif USEEMS = 2 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
elseif USEEMS = 3 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
  emp3r1, emp3r2, emp3r3, emp3r4, emp3r5, _
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
elseif USEEMS = 4 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
  emp3r1, emp3r2, emp3r3, emp3r4, emp3r5, _
  emp4r1, emp4r2, emp4r3, emp4r4, emp4r5, _
  emp4r6) ' credits
end If


Sub center_objects_em()

Dim cnt,ii, xx, yy, yfact, xfact

  zscale = 0.0000001
  xcen =(960 /2) - (17 / 2)
  ycen = (1065 /2 ) + (313 /2)
  yfact = -25
  xfact = 0
  cnt =0

For Each objs In BGObjEM(0)
  If Not IsEmpty(objs) then
    if objs.name = emp4r6.name then
      yoff = -15 ' credit drum is 60% smaller
  Else
      yoff = -50
  end if

xx =objs.x

objs.x = (xoff - xcen) + xx + xfact

yy = objs.y

objs.y =yoff

If yy < 0 then
  yy = yy * -1
end if

objs.z = (zoff - ycen) + yy - (yy * zscale) + yfact

end if

  cnt = cnt + 1

Next

end sub


' ********************* UPDATE EM REEL DRUMS  *************************


 'reset scores used to track reel display to defaults

for np =0 to 3

  scores(np,0 ) = 0
  scores(np,1 ) = 0

Next


' ********************* sets localised score used to display players score on reels*********************

Sub SetScore(player, ndx , val)


Dim ncnt

  if val > 0 then
      If(ndx = 0)then ncnt = val * 10000
      If(ndx = 1)then ncnt = val * 1000
      If(ndx = 2)Then ncnt = val * 100
      If(ndx = 3)Then ncnt = val * 10
      If(ndx = 4)Then ncnt = val
      scores(player, 0) = scores(player, 0) + ncnt
  end if

End Sub

Sub SetDrum(player, drum , val)

Dim cnt
objs =BGObjEM(0)


Select case player

Case -1: ' the credit drum

  'emp4r6.ObjrotX = emp4r6.ObjrotX + val

If Not IsEmpty(objs(idx_emp4r6)) then

  objs(idx_emp4r6).ObjrotX = val

end if

Case 0:

Select Case drum

Case 1: If Not IsEmpty(objs(idx_emp1r1)) then: objs(idx_emp1r1).ObjrotX= val: end if
Case 2: If Not IsEmpty(objs(idx_emp1r1+1)) then: objs(idx_emp1r1+1).ObjrotX= val: end if
Case 3: If Not IsEmpty(objs(idx_emp1r1+2)) then: objs(idx_emp1r1+2).ObjrotX= val: end if
Case 4: If Not IsEmpty(objs(idx_emp1r1+3)) then: objs(idx_emp1r1+3).ObjrotX= val: end if
Case 5: If Not IsEmpty(objs(idx_emp1r1+4)) then: objs(idx_emp1r1+4).ObjrotX= val: end if

End Select

Case 1:

Select Case drum

Case 1: If Not IsEmpty(objs(idx_emp2r1)) then: objs(idx_emp2r1).ObjrotX= val: end if
Case 2: If Not IsEmpty(objs(idx_emp2r1+1)) then: objs(idx_emp2r1+1).ObjrotX= val: end if
Case 3: If Not IsEmpty(objs(idx_emp2r1+2)) then: objs(idx_emp2r1+2).ObjrotX= val: end if
Case 4: If Not IsEmpty(objs(idx_emp2r1+3)) then: objs(idx_emp2r1+3).ObjrotX= val: end if
Case 5: If Not IsEmpty(objs(idx_emp2r1+4)) then: objs(idx_emp2r1+4).ObjrotX= val: end if

End Select

Case 2:

Select Case drum

Case 1: If Not IsEmpty(objs(idx_emp3r1)) then: objs(idx_emp3r1).ObjrotX= val: end if
Case 2: If Not IsEmpty(objs(idx_emp3r1+1)) then: objs(idx_emp3r1+1).ObjrotX= val: end if
Case 3: If Not IsEmpty(objs(idx_emp3r1+2)) then: objs(idx_emp3r1+2).ObjrotX= val: end if
Case 4: If Not IsEmpty(objs(idx_emp3r1+3)) then: objs(idx_emp3r1+3).ObjrotX= val: end if
Case 5: If Not IsEmpty(objs(idx_emp3r1+4)) then: objs(idx_emp3r1+4).ObjrotX= val: end if

End Select

Case 3:

Select Case drum

Case 1: If Not IsEmpty(objs(idx_emp4r1)) then: objs(idx_emp4r1).ObjrotX= val: end if
Case 2: If Not IsEmpty(objs(idx_emp4r1+1)) then: objs(idx_emp4r1+1).ObjrotX= val: end if
Case 3: If Not IsEmpty(objs(idx_emp4r1+2)) then: objs(idx_emp4r1+2).ObjrotX= val: end if
Case 4: If Not IsEmpty(objs(idx_emp4r1+3)) then: objs(idx_emp4r1+3).ObjrotX= val: end if
Case 5: If Not IsEmpty(objs(idx_emp4r1+4)) then: objs(idx_emp4r1+4).ObjrotX= val: end if

End Select

End Select

'end if

End Sub


Sub SetReel(player, drum, val)


Select Case drum

Case -1: ' credits drum

SetDrum -1,0, -val*(360/16)


Case 1:


SetDrum player,drum, -val*36

Case 2:

SetDrum player,drum, -val*36

Case 3:


SetDrum player,drum, -val*36

Case 4:


SetDrum player,drum, -val*36

Case 5:

SetDrum player,drum, -val*36

End Select

End Sub



'*****************  routine to display players score ************

' if EMMODE = 1 then this routine only needs the following inputs in this format

'ex: UpdateReels 0-3, 0-3, 0-x99999, n/a, n/a, n/a, n/a, n/a, n/a

' if EMMODE = 0 then this routine only needs the following inputs in this format. This allows individual reels to be displayed

'ex: UpdateReels 0-3, 0-3, n/a, 0-99999 ,0-9999, 0-999, 0-99, 0-9

'eg if player 1 is 1 outside the sub. the sub will change to 0, if display on reel 1 outside the script, the sub will make it 0

Sub UpdateReels (Player,nReels ,nScore, Score10000 ,Score1000,Score100,Score10,Score1)

  value =nScore     ' this is the score to display on the reels for the active player

  If value => 100000 then       ' this routine ensures the score is between 0 and 99,999
    do until value < 100000
    value=value-100000
    loop
  end if


  nplayer = Player - 1  ' if player 1 is 1 outside the sub, this makes it 0 for use within the sub


  curscr = value
  curplayr = nplayer

' "sccores(x,x)  is used to store the display value for each reel.

  scores(0,1) = scores(0,0)
  scores(0,0) = 0
  scores(1,1) = scores(1,0)
  scores(1,0) = 0
  scores(2,1) = scores(2,0)
  scores(2,0) = 0
  scores(3,1) = scores(3,0)
  scores(3,0) = 0



' Drums are reset to zero every time a score is updated. The sub re-applies the score every execution

  For  ix =0 to 4

  SetDrum ix, 1 , 0
  SetDrum ix, 2 , 0
  SetDrum ix, 3 , 0
  SetDrum ix, 4 , 0
  SetDrum ix, 5 , 0

  Next

For playr =0 to 3

  if EMMODE = 0 then
    If curplayr = playr Then

    nplayer = playr

    SetReel nplayer, 1 , Score10000 : SetScore nplayer,0,Score10000
    SetReel nplayer, 2 , Score1000 : SetScore nplayer,1,Score1000
    SetReel nplayer, 3 , Score100 : SetScore nplayer,2,Score100
    SetReel nplayer, 4 , Score10 : SetScore nplayer,3,Score10
    SetReel nplayer, 5 , 0 : SetScore nplayer,4,0 ' assumes ones position is always zero

    else

    nplayer = playr
    value =scores(nplayer, 1)

  ' do ten thousands
    if(value >= 90000)  then:  SetReel nplayer, 1 , 9 : SetScore nplayer,0,9 : value = value - 90000: end if
    if(value >= 80000)  then:  SetReel nplayer, 1 , 8 : SetScore nplayer,0,8 : value = value - 80000: end if
    if(value >= 70000)  then:  SetReel nplayer, 1 , 7 : SetScore nplayer,0,7 : value = value - 70000: end if
    if(value >= 60000)  then:  SetReel nplayer, 1 , 6 : SetScore nplayer,0,6 : value = value - 60000: end if
    if(value >= 50000)  then:  SetReel nplayer, 1 , 5 : SetScore nplayer,0,5 : value = value - 50000: end if
    if(value >= 40000)  then:  SetReel nplayer, 1 , 4 : SetScore nplayer,0,4 : value = value - 40000: end if
    if(value >= 30000)  then:  SetReel nplayer, 1 , 3 : SetScore nplayer,0,3 : value = value - 30000: end if
    if(value >= 20000)  then:  SetReel nplayer, 1 , 2 : SetScore nplayer,0,2 : value = value - 20000: end if
    if(value >= 10000)  then:  SetReel nplayer, 1 , 1 : SetScore nplayer,0,1 : value = value - 10000: end if


  ' do thousands
    if(value >= 9000)  then:  SetReel nplayer, 2 , 9 : SetScore nplayer,1,9 : value = value - 9000: end if
    if(value >= 8000)  then:  SetReel nplayer, 2 , 8 : SetScore nplayer,1,8 : value = value - 8000: end if
    if(value >= 7000)  then:  SetReel nplayer, 2 , 7 : SetScore nplayer,1,7 : value = value - 7000: end if
    if(value >= 6000)  then:  SetReel nplayer, 2 , 6 : SetScore nplayer,1,6 : value = value - 6000: end if
    if(value >= 5000)  then:  SetReel nplayer, 2 , 5 : SetScore nplayer,1,5 : value = value - 5000: end if
    if(value >= 4000)  then:  SetReel nplayer, 2 , 4 : SetScore nplayer,1,4 : value = value - 4000: end if
    if(value >= 3000)  then:  SetReel nplayer, 2 , 3 : SetScore nplayer,1,3 : value = value - 3000: end if
    if(value >= 2000)  then:  SetReel nplayer, 2 , 2 : SetScore nplayer,1,2 : value = value - 2000: end if
    if(value >= 1000)  then:  SetReel nplayer, 2 , 1 : SetScore nplayer,1,1 : value = value - 1000: end if

    'do hundreds

    if(value >= 900) then: SetReel nplayer, 3 , 9 : SetScore nplayer,2,9 : value = value - 900: end if
    if(value >= 800) then: SetReel nplayer, 3 , 8 : SetScore nplayer,2,8 : value = value - 800: end if
    if(value >= 700) then: SetReel nplayer, 3 , 7 : SetScore nplayer,2,7 : value = value - 700: end if
    if(value >= 600) then: SetReel nplayer, 3 , 6 : SetScore nplayer,2,6 : value = value - 600: end if
    if(value >= 500) then: SetReel nplayer, 3 , 5 : SetScore nplayer,2,5 : value = value - 500: end if
    if(value >= 400) then: SetReel nplayer, 3 , 4 : SetScore nplayer,2,4 : value = value - 400: end if
    if(value >= 300) then: SetReel nplayer, 3 , 3 : SetScore nplayer,2,3 : value = value - 300: end if
    if(value >= 200) then: SetReel nplayer, 3 , 2 : SetScore nplayer,2,2 : value = value - 200: end if
    if(value >= 100) then: SetReel nplayer, 3 , 1 : SetScore nplayer,2,1 : value = value - 100: end if

    'do tens
    if(value >= 90) then: SetReel nplayer, 4 , 9 : SetScore nplayer,3,9 : value = value - 90: end if
    if(value >= 80) then: SetReel nplayer, 4 , 8 : SetScore nplayer,3,8 : value = value - 80: end if
    if(value >= 70) then: SetReel nplayer, 4 , 7 : SetScore nplayer,3,7 : value = value - 70: end if
    if(value >= 60) then: SetReel nplayer, 4 , 6 : SetScore nplayer,3,6 : value = value - 60: end if
    if(value >= 50) then: SetReel nplayer, 4 , 5 : SetScore nplayer,3,5 : value = value - 50: end if
    if(value >= 40) then: SetReel nplayer, 4 , 4 : SetScore nplayer,3,4 : value = value - 40: end if
    if(value >= 30) then: SetReel nplayer, 4 , 3 : SetScore nplayer,3,3 : value = value - 30: end if
    if(value >= 20) then: SetReel nplayer, 4 , 2 : SetScore nplayer,3,2 : value = value - 20: end if
    if(value >= 10) then: SetReel nplayer, 4 , 1 : SetScore nplayer,3,1 : value = value - 10: end if

    'do ones
    if(value >= 9) then: SetReel nplayer, 5 , 9 : SetScore nplayer,4,9 : value = value - 9: end if
    if(value >= 8) then: SetReel nplayer, 5 , 8 : SetScore nplayer,4,8 : value = value - 8: end if
    if(value >= 7) then: SetReel nplayer, 5 , 7 : SetScore nplayer,4,7 : value = value - 7: end if
    if(value >= 6) then: SetReel nplayer, 5 , 6 : SetScore nplayer,4,6 : value = value - 6: end if
    if(value >= 5) then: SetReel nplayer, 5 , 5 : SetScore nplayer,4,5 : value = value - 5: end if
    if(value >= 4) then: SetReel nplayer, 5 , 4 : SetScore nplayer,4,4 : value = value - 4: end if
    if(value >= 3) then: SetReel nplayer, 5 , 3 : SetScore nplayer,4,3 : value = value - 3: end if
    if(value >= 2) then: SetReel nplayer, 5 , 2 : SetScore nplayer,4,2 : value = value - 2: end if
    if(value >= 1) then: SetReel nplayer, 5 , 1 : SetScore nplayer,4,1 : value = value - 1: end if

    end if

  Else
    If curplayr = playr Then
      nplayer = curplayr
      value = curscr
    else
      value =scores(playr, 1) ' previous score for other players
      nplayer = playr
    end if

  ' do ten thousands
    if(value >= 90000)  then:  SetReel nplayer, 1 , 9 : SetScore nplayer,0,9 : value = value - 90000: end if
    if(value >= 80000)  then:  SetReel nplayer, 1 , 8 : SetScore nplayer,0,8 : value = value - 80000: end if
    if(value >= 70000)  then:  SetReel nplayer, 1 , 7 : SetScore nplayer,0,7 : value = value - 70000: end if
    if(value >= 60000)  then:  SetReel nplayer, 1 , 6 : SetScore nplayer,0,6 : value = value - 60000: end if
    if(value >= 50000)  then:  SetReel nplayer, 1 , 5 : SetScore nplayer,0,5 : value = value - 50000: end if
    if(value >= 40000)  then:  SetReel nplayer, 1 , 4 : SetScore nplayer,0,4 : value = value - 40000: end if
    if(value >= 30000)  then:  SetReel nplayer, 1 , 3 : SetScore nplayer,0,3 : value = value - 30000: end if
    if(value >= 20000)  then:  SetReel nplayer, 1 , 2 : SetScore nplayer,0,2 : value = value - 20000: end if
    if(value >= 10000)  then:  SetReel nplayer, 1 , 1 : SetScore nplayer,0,1 : value = value - 10000: end if


  ' do thousands
    if(value >= 9000)  then:  SetReel nplayer, 2 , 9 : SetScore nplayer,1,9 : value = value - 9000: end if
    if(value >= 8000)  then:  SetReel nplayer, 2 , 8 : SetScore nplayer,1,8 : value = value - 8000: end if
    if(value >= 7000)  then:  SetReel nplayer, 2 , 7 : SetScore nplayer,1,7 : value = value - 7000: end if
    if(value >= 6000)  then:  SetReel nplayer, 2 , 6 : SetScore nplayer,1,6 : value = value - 6000: end if
    if(value >= 5000)  then:  SetReel nplayer, 2 , 5 : SetScore nplayer,1,5 : value = value - 5000: end if
    if(value >= 4000)  then:  SetReel nplayer, 2 , 4 : SetScore nplayer,1,4 : value = value - 4000: end if
    if(value >= 3000)  then:  SetReel nplayer, 2 , 3 : SetScore nplayer,1,3 : value = value - 3000: end if
    if(value >= 2000)  then:  SetReel nplayer, 2 , 2 : SetScore nplayer,1,2 : value = value - 2000: end if
    if(value >= 1000)  then:  SetReel nplayer, 2 , 1 : SetScore nplayer,1,1 : value = value - 1000: end if

    'do hundreds

    if(value >= 900) then: SetReel nplayer, 3 , 9 : SetScore nplayer,2,9 : value = value - 900: end if
    if(value >= 800) then: SetReel nplayer, 3 , 8 : SetScore nplayer,2,8 : value = value - 800: end if
    if(value >= 700) then: SetReel nplayer, 3 , 7 : SetScore nplayer,2,7 : value = value - 700: end if
    if(value >= 600) then: SetReel nplayer, 3 , 6 : SetScore nplayer,2,6 : value = value - 600: end if
    if(value >= 500) then: SetReel nplayer, 3 , 5 : SetScore nplayer,2,5 : value = value - 500: end if
    if(value >= 400) then: SetReel nplayer, 3 , 4 : SetScore nplayer,2,4 : value = value - 400: end if
    if(value >= 300) then: SetReel nplayer, 3 , 3 : SetScore nplayer,2,3 : value = value - 300: end if
    if(value >= 200) then: SetReel nplayer, 3 , 2 : SetScore nplayer,2,2 : value = value - 200: end if
    if(value >= 100) then: SetReel nplayer, 3 , 1 : SetScore nplayer,2,1 : value = value - 100: end if

    'do tens
    if(value >= 90) then: SetReel nplayer, 4 , 9 : SetScore nplayer,3,9 : value = value - 90: end if
    if(value >= 80) then: SetReel nplayer, 4 , 8 : SetScore nplayer,3,8 : value = value - 80: end if
    if(value >= 70) then: SetReel nplayer, 4 , 7 : SetScore nplayer,3,7 : value = value - 70: end if
    if(value >= 60) then: SetReel nplayer, 4 , 6 : SetScore nplayer,3,6 : value = value - 60: end if
    if(value >= 50) then: SetReel nplayer, 4 , 5 : SetScore nplayer,3,5 : value = value - 50: end if
    if(value >= 40) then: SetReel nplayer, 4 , 4 : SetScore nplayer,3,4 : value = value - 40: end if
    if(value >= 30) then: SetReel nplayer, 4 , 3 : SetScore nplayer,3,3 : value = value - 30: end if
    if(value >= 20) then: SetReel nplayer, 4 , 2 : SetScore nplayer,3,2 : value = value - 20: end if
    if(value >= 10) then: SetReel nplayer, 4 , 1 : SetScore nplayer,3,1 : value = value - 10: end if

    'do ones
    if(value >= 9) then: SetReel nplayer, 5 , 9 : SetScore nplayer,4,9 : value = value - 9: end if
    if(value >= 8) then: SetReel nplayer, 5 , 8 : SetScore nplayer,4,8 : value = value - 8: end if
    if(value >= 7) then: SetReel nplayer, 5 , 7 : SetScore nplayer,4,7 : value = value - 7: end if
    if(value >= 6) then: SetReel nplayer, 5 , 6 : SetScore nplayer,4,6 : value = value - 6: end if
    if(value >= 5) then: SetReel nplayer, 5 , 5 : SetScore nplayer,4,5 : value = value - 5: end if
    if(value >= 4) then: SetReel nplayer, 5 , 4 : SetScore nplayer,4,4 : value = value - 4: end if
    if(value >= 3) then: SetReel nplayer, 5 , 3 : SetScore nplayer,4,3 : value = value - 3: end if
    if(value >= 2) then: SetReel nplayer, 5 , 2 : SetScore nplayer,4,2 : value = value - 2: end if
    if(value >= 1) then: SetReel nplayer, 5 , 1 : SetScore nplayer,4,1 : value = value - 1: end if

    end if

  Next

End Sub


'*****    Table Initialisation  *******

Sub Table1_Init()
     Set Controller = CreateObject("B2S.Server")
    Controller.B2SName = "Zodiac"  ' *****!!!!This must match the name of your directb2s file!!!!
    Controller.Run()
  LoadEM
  loadhs
  updatepostit
  SetBackglass
  center_objects_em
  Players = 0
  currentp=0
  EndScoresTimer.Enabled = false
  LightSunAdv.state=1
  Sbonus=1
  Mbonus=1
  UpdateSBonusLights
  UpdateMBonusLights

  if PlayerScores(1)="" then PlayerScores(1)=1000
  if PlayerScores(2)="" then PlayerScores(2)=2000
  if player1100k="" then player1100k=0
  if player2100k="" then player2100k=0
  If Credits="" then Credits=0
  If hiscore="" then hiscore=1000
  if starsign="" then starsign=1
    Controller.B2ssetdata 1, starsign
  if lastpos="" then lastpos=2      '(starting angle of spinner is 10*lastpos in deg.)

  advancestarsign


  If twoposts  = true then discposts.visible = 1: discposts1.visible = 0 : else : discposts.visible = 0 : discposts1.visible = 1 :  end if

  InitAngles
  InitLamp

  for each Object in ColFlGameOver : object.visible = 1 : next


  UpdateReels 1 ,1 ,PlayerScores(1), -1,-1,-1,-1,-1
  UpdateReels 2 ,2 ,PlayerScores(2), -1,-1,-1,-1,-1



  Tilt=False
    Controller.B2SSetTilt 33,0
  If Credits > 0 Then CreditLight.State = 1 'should be 1
  If Credits = 0 Then CreditLight.State = 0 'should be 0



  SetReel 4,-1,  Credits



    Dim obj

' ******* to test flashers and lights *********************
' for each Object in Backglass_flashers : object.visible = 1 : next
' for each object in playfieldlighttest : object.state=1  : next

End Sub


  Set mMagnetS=New cvpmMagnet
    With mMagnetS
    .InitMagnet sw23, 26
    .MagnetOn = 0
        .CreateEvents "mMagnetS"
    .GrabCenter=False
    End With

  Set mMagnetM=New cvpmMagnet
    With mMagnetM
    .InitMagnet sw24, 26
    .MagnetOn = 0
        .CreateEvents "mMagnetM"
    .GrabCenter=False
    End With







'***********KeyCodes
Sub Table1_KeyDown(ByVal keycode)

  If keycode = LeftFlipperKey and BIP = 1 and tilt=False Then
    LeftFlipper.RotateToEnd
    LeftFlipper001.RotateToEnd
  leftSlingShot.Disabled =0
  RightSlingShot.Disabled =0
    PlaySoundAtVol "flip1up", LeftFlipper, 1
    PlaySoundAtVol "flip1up", LeftFlipper001, 1
    PlayLoopSoundAtVol "buzzL", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey and BIP = 1 and tilt=False Then
    RightFlipper.RotateToEnd
    RightFlipper001.RotatetoEnd
  leftSlingShot.Disabled =0
  RightSlingShot.Disabled =0
    PlaySoundAtVol "flip1up", RightFlipper, 1
    PlaySoundAtVol "flip1up", RightFlipper001, 1
    PlayLoopSoundAtVol "buzz", RightFlipper, 1
  End If


  If keycode = rightmagnasave then
    'Plunger.Fire
        launcher.Z = launcher.z -5
    If Balls > 0 and currentp > 0 and BIP = 0 Then
      PlaySound "BallLaunch"
      if Shootagain.state=1 Then Shootagain.state=0 end if
      BallRelease.CreateSizedBallWithMass BallSize, BallSize^3/15625
      BallRelease.Kick 1, 25
      BIP = 1
    End If
  End If

  If keycode = LeftTiltKey and tilt=false Then
    Nudge 90, 2
    checktilt
  End If

  If keycode = RightTiltKey and tilt=false Then
    Nudge 270, 2
    checktilt
  End If

  If keycode = CenterTiltKey and tilt=false Then
    Nudge 0, 2
    checktilt
  End If
End Sub

Sub CheckTilt
  If Tilttimer.Enabled = True Then
   TiltSens = TiltSens + 1
   if TiltSens = 3 Then
    Tilt = True
    for each Object in ColFlTilt : object.visible = 1 : next
    TiltSens = 0
      playsound "hit1"
    turnoff
        Controller.B2SSetTilt 33,1
   End If
  Else
   TiltSens = 0
   Tilttimer.Enabled = True
  End If
End Sub

Sub Tilttimer_Timer()
  Tilttimer.Enabled = False
End Sub

Sub turnoff
  LeftFlipper.RotateToStart
  LeftFlipper001.RotateToStart
  RightFlipper.RotateToStart
  RightFlipper001.RotateToStart
  leftSlingShot.Disabled =1
  RightSlingShot.Disabled =1
end sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey then
    launcher.Z = launcher.z +5
  end if

  If Keycode = AddCreditKey Then
    PlaySoundAtVol "coinin", Drain, 1
    coindelay.enabled=true
        Controller.B2SSetCredits Credits
  End If

  If keycode = StartGameKey and Credits > 0 and Players =0  Then
    Players = Players + 1
    PlaySoundAtVol  "Start_of_Game", drain, 1
        Controller.B2SSetGameOver 35,0
        Controller.B2SSetBallInPlay 32, 1
        Controller.B2SSetCanPlay 31, 1
    tilt=False
    for each Object in ColFlTilt : object.visible = 0 : next
    Controller.B2SSetTilt 33,0
    credits = Credits - 1
        Controller.B2SSetCredits Credits
    If Credits > 0 Then CreditLight.State = 1 'should be 1
    If Credits = 0 Then CreditLight.State = 0 'should be 0
    SetReel 4,-1,  Credits
    currentp = 1
    for each Object in ColFlPlayer1
    object.visible = 1
    lightplayer1up.state=1
        Controller.B2SSetPlayerUp 30,Players
    Next
    for each Object in ColFlPlayer2
    object.visible = 0
    lightplayer2up.state=0
    Next
    Balls = ballspergame
    BIP = 0
    If players = 1 then
      resetGame
      FlBIP1A.visible = 1
      rst=0
      resettimer.enabled=true
    end if

    for each Object in ColFlGameOver : object.visible = 0 : next

    If Players = 1 Then FlPLR1.visible = 1 : FlPLR1A.visible = 1: Else  FlPLR1.visible = 0 : FlPLR1A.visible = 0: End If
    If Players = 2 Then FlPLR2.visible = 1 : FlPLR2A.visible = 1 :  Else FlPLR2.visible = 0 : FlPLR2A.visible = 0: End If

  elseif keycode = StartGameKey and Credits>0 and Players = 1 and playerScores(1)=0 then
    Players = Players + 1
    PlaySound "click",0,1,0.25,0.25

    credits = Credits - 1
        Controller.B2SSetCredits Credits
    If Credits > 0 Then CreditLight.State = 1 'should be 1
    If Credits = 0 Then CreditLight.State = 0 'should be 0
    SetReel 4,-1,  Credits
    If Players = 1 Then FlPLR1.visible = 1 : FlPLR1A.visible = 1: Else  FlPLR1.visible = 0 : FlPLR1A.visible = 0: End If
    If Players = 2 Then FlPLR2.visible = 1 : FlPLR2A.visible = 1 :  Else FlPLR2.visible = 0 : FlPLR2A.visible = 0: End If
        If Players=2 Then Controller.B2SSetCanPlay 31, 2
  End If

  if BIP=1 Then
  If keycode = LeftFlipperKey and tilt=False Then
    LeftFlipper.RotateToStart
    LeftFlipper001.RotateToStart
    PlaySoundAtVol "flip1down", LeftFlipper, 1
    PlaySoundAtVol "flip1down", LeftFlipper001, 1
    StopSound "buzzL"
  End If

  If keycode = RightFlipperKey and tilt=False Then
    RightFlipper.RotateToStart
    RightFlipper001.RotateToStart
    PlaySoundAtVol "flip1down", RightFlipper, 1
    PlaySoundAtVol "flip1down", RightFlipper001, 1
    StopSound "buzz"
  End If
  end if

End Sub


sub coindelay_timer
    credits=Credits+1
        Controller.B2SSetCredits Credits
    if credits>15 then credits=15
    If Credits > 0 Then CreditLight.State = 1 'should be 1
    If Credits = 0 Then CreditLight.State = 0 'should be 0
        Controller.B2SSetCredits Credits

    SetReel 4,-1,  Credits

  coindelay.enabled=false
End sub

Dim Pop, pop1, pop2, pop3


Sub AddSpecial(pop)

if pop =3 then pop2=3
if pop =2 then pop2=2
if pop =1 then pop2=1

if pop = 3 then pop3=0 : addspecialtimer.enabled=true : end if
if pop = 2 then pop3=0 : addspecialtimer.enabled=true : end if
if pop = 1 then pop3=0 : addspecialtimer.enabled=true : end if

End Sub

sub Addspecialtimer_timer

pop3=pop3 + 1

  PlaySound"knocker"
  credits=Credits+1
    Controller.B2SSetCredits Credits
  if credits>15 then credits=15
  If Credits > 0 Then CreditLight.State = 1 'should be 1
  If Credits = 0 Then CreditLight.State = 0 'should be 0
      Controller.B2SSetCredits Credits

  SetReel 4,-1,  Credits

  if pop3=pop2 then addspecialtimer.enabled=false
End sub

Sub CheckFreeGame
  If PrevScore(currentp) < Replay1 And playerScores(currentp) >= Replay1 Then addspecial(1)
  If PrevScore(currentp) < Replay2 And playerScores(currentp) >= Replay2 Then addspecial(1)
  If PrevScore(currentp) < Replay3 And playerScores(currentp) >= Replay3 Then addspecial(1)
End Sub

Sub EndScoresTimer_Timer()

  if Shootagain.state=1 Then
    EndScoresTimer.Enabled = false
    exit Sub
  end if

If Players = 2 then

if currentp = 1 Then

  currentp = 2
  for each Object in ColFlPlayer2
  object.visible = 1
  lightplayer2up.state=1
  Next

  for each Object in ColFlPlayer1
  object.visible = 0
  lightplayer1up.state=0
     Controller.B2SSetPlayerUp 30,currentp
  Next

else

  currentp = 1

  for each Object in ColFlPlayer1
  object.visible = 1
  lightplayer1up.state=1
  Next

  for each Object in ColFlPlayer2
  object.visible = 0
  lightplayer2up.state=0
    Controller.B2SSetPlayerUp 30,currentp
  Next
  Balls = Balls - 1
  end if

end if

If players = 1 Then
  Balls = Balls - 1
end if

  If Balls = 5 then FlBIP1A.visible = 1 Else  FlBIP1A.visible = 0 End If
  If Balls = 4 then FlBIP2A.visible = 1 Else  FlBIP2A.visible = 0 End If
  If Balls = 3 then FlBIP3A.visible = 1 Else FlBIP3A.visible = 0 End If
  If Balls = 2 then FlBIP4A.visible = 1 Else FlBIP4A.visible = 0 End If
  If Balls = 1 then FlBIP5A.visible = 1 Else FlBIP5A.visible = 0 End If
    If Balls=5 Then Controller.B2SSetBallInPlay 32, 1
    If Balls=4 Then Controller.B2SSetBallInPlay 32, 2
    If Balls=3 Then Controller.B2SSetBallInPlay 32, 3
    If Balls=2 Then Controller.B2SSetBallInPlay 32, 4
    If Balls=1 Then Controller.B2SSetBallInPlay 32, 5

If Balls = 0 Then GameOver: end if

If Balls <= 5 then EndScoresTimer.Enabled = false  end if


End Sub


Sub Drain_Hit()
  PlaySoundAtVol "Drains", Drain, 1
  Drain.DestroyBall
  StopSound "buzzL"
  StopSound "buzz"
  BIP = 0
  if tilt=true then
  tilt=false
  for each Object in ColFlTilt : object.visible = 0 : next
  end if
  EndScoresTimer.Enabled = true
  turnoff
  resetball
End Sub

'***********************
'     Flipper shadows
'***********************

Sub UpdateFlipperLogos_Timer

  FlipperLSh1.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

  FlipperLSh001.RotZ = LeftFlipper001.currentangle
  FlipperRSh2.RotZ = RightFlipper001.currentangle
End Sub



Sub GameOver
    Controller.B2SSetGameOver 35,1
  PlaySound "GameOver",0,1,0.25,0.25
  stopsound "motor"
  Players = 0
  currentp = 0
  Balls = 0
    Controller.B2SSetPlayerUp 30,0
  for each Object in ColFlGameOver : object.visible = 1 : next
    Controller.B2SSetCanPlay 31, 0
    Controller.B2SSetBallInPlay 32, 0


  ' ------------------
  ' check for highscore

  if PlayerScores(1)+(player1100k*100000)>hiscore then hiscore=PlayerScores(1)+(player1100k*100000)
  if PlayerScores(2)+(player2100k*100000)>hiscore then hiscore=PlayerScores(2)+(player2100k*100000)
  savehs
  updatePostIt

End Sub

Sub Triggers_hit(idx)
  dim rotball
  rotball=idx+1
  if moveballs=1 Then
    direction(rotball)=Int(Rnd*2)-1
    if direction(rotball)=0 then direction(rotball)=1
    if ballangle(rotball)=0 then
      ballangle(rotball)=(12*direction(rotball))
    end if
  end If


End Sub

Sub BallTurnTimer_Timer()
  dim x
  for x=1 to 15
    if ballangle(x)<>0 Then
      ballangle(x)=ballangle(x)+(12*direction(x))
      if (ballangle(x) > 359) or (ballangle(x) < -359) Then
        ballangle(x) = 0
      End If
    end if
  next
End Sub

Sub addbonus
    if LightSunAdv.state=1 then AddSBonus
    if LightMoonadv.state=1 then AddMbonus
end sub


Sub AddSBonus
'    If Tilted = False Then
    addscore(100)
        SBonus = SBonus + 1
        If SBonus > 20 Then SBonus = 20
        UpdateSBonusLights
'    End if
End Sub






Sub AddMBonus
'    If Tilt = False Then
    addscore(100)
        MBonus = MBonus + 1
        If MBonus > 20 Then MBonus = 20
        UpdateMBonusLights
'    End if
End Sub


'**********************************************************************************************************
'Spinning Disc Animation
'**********************************************************************************************************

Dim LampPosition            ' In Degrees (0-360  '180 for 2 posts)
Dim SpinSpeed             ' SpinSpeed is in Deg/Sec
Dim LegNum
Dim ee



const pi = 3.1415926535897932
const SpeedMultiplier = 18        ' was 18 - Affects speed transfer to object (deg/sec)
const Friction = .75          ' Friction coefficient (deg/sec/sec)
const MinSpeed = 25           ' Object stops at this speed. (deg/sec)

const ObjectRadius = 65         ' Distance from center of circle to center of post



' ***********    "discposts" or "discposts1".  needs to be in the centre of the leg's . Its co-ordinates is used for calculating hits to the disc


dim ObjectX, ObjectY
ObjectX = discposts.X
ObjectY = discposts.Y


Dim AngleRef(36), PostX(36), PostY(36), LegPos

LegPos = Array( Leg1, Leg2, Leg3, Leg4, Leg5, Leg6, Leg7, Leg8, Leg9, Leg10, Leg11, Leg12, Leg13, Leg14, Leg15, Leg16, Leg17, Leg18, Leg19, Leg20, Leg21, Leg22, Leg23, Leg24, Leg25, Leg26, Leg27, Leg28, Leg29, Leg30, Leg31, Leg32, Leg33, Leg34, Leg35, Leg36)

'

'--------------------------------
' Construction notes:
' With respect to the grid at standard size:
'   wall-centers are 1.5 units from center of circle.
'   Approach triggers are 2.75 units out.
'--------------------------------

Sub InitLamp()
  Dim obj

  for each obj in LegPos : obj.IsDropped = true : next

  LampPosition = lastpos*10
  SpinSpeed = 0
  LegNum = lastpos
  LegPos(legnum).IsDropped = false
  if twoposts=true then
  if legnum > 17 then LegPos(legnum - 18).IsDropped = false else LegPos(legnum + 18).IsDropped = false  end if
  end if
  DisplayLamp
End Sub

Sub InitAngles
  Dim ee

  for ee = 0 to 35

  AngleRef(ee) = ((10*ee) * pi/180)
  PostX(ee) = ObjectX + ObjectRadius*cos(AngleRef(ee))
  PostY(ee) = ObjectY - ObjectRadius*sin(AngleRef(ee))
  next
End Sub

Function GetAngle( x, y )
  ' Angle quadrants:
  '  Q2 (-x,-y)  |  Q1 (+x,-y)
  '--------------+--------------
  '  Q3 (-x,+y)  |  Q4 (+x,+y)

  if x = 0 And y < 0 then
    GetAngle = pi / 180             ' -90 degrees
  elseif x = 0 And y > 0 then
    GetAngle = -pi / 180            ' 90 degrees
  else
    GetAngle = atn( -y / x )
    if x < 0 then GetAngle = GetAngle + pi    ' Add 180 deg if in quadrant 2 or 3
  End if

  if GetAngle < 0 then GetAngle = GetAngle + 2*pi
  if GetAngle > 2*pi then GetAngle = GetAngle - 2*pi
End Function



Sub DisplayLamp()

  curPos = CInt(LampPosition / 10) Mod 36

  if curPos <> lastPos then

    LegPos(lastPos).IsDropped = true
    LegPos(curPos).IsDropped = false

  if twoposts=true then
    if lastpos > 17 then LegPos(lastpos - 18).IsDropped = true else LegPos(lastpos + 18).IsDropped = true end if
    if curpos > 17 then LegPos(curpos - 18).IsDropped = false else LegPos(curpos + 18).IsDropped = false  end if
  end if

    if curPos = 0 then addbonus ' trip lamp switch

    if curPos = 18 then addbonus  ' trip lamp switch  (180deg from curpos 0)

    lastPos = curPos
  end if

End Sub



Sub SpinTimer_Timer()
  if Abs(SpinSpeed) < MinSpeed then SpinSpeed = 0     ' Stop lamp below minimum speed.
  ComputePosition
    discposts.ObjRotZ = -lampposition - 180
    discposts1.ObjRotZ = -lampposition - 180
    discbase.ObjRotZ =  -lampposition + 55
  DisplayLamp
' DebugBox2.Text = SpinSpeed
' DebugBox2.Text = curPos
End Sub

' Definitions:
' SpinSpeed = Rotational speed in degrees per second.
' Diff/Position = degrees + SpinSpeed / nIntervals p.sec
' Friction applied: Speed - Speed*Frict/ nIntervals

Sub ComputePosition()
  ' Theory: Position is expressed in degrees.
  ' Add/Subtract to degrees with speed value.

  LampPosition = LampPosition + SpinSpeed / (1000 / SpinTimer.Interval)

  while LampPosition < 0.0
    LampPosition = LampPosition + 360     '180.0 if 2 posts
  wend

  while LampPosition > 360            '180  if 2 posts
    LampPosition = LampPosition - 360     '180  if 2 posts
  wend

  SpinSpeed = SpinSpeed - SpinSpeed * Friction / (1000/SpinTimer.Interval)
End Sub

'--------------------------
' Construction notes:
'   Angles are counter-clockwise, 0 degrees = right, 90 degrees = up
'--------------------------


Sub Leg1_Hit() : ComputeLegHit 0 , ActiveBall : End Sub
Sub Leg2_Hit() : ComputeLegHit 1 , ActiveBall : End Sub
Sub Leg3_Hit() : ComputeLegHit 2 , ActiveBall : End Sub
Sub Leg4_Hit() : ComputeLegHit 3 , ActiveBall : End Sub
Sub Leg5_Hit() : ComputeLegHit 4 , ActiveBall : End Sub
Sub Leg6_Hit() : ComputeLegHit 5 , ActiveBall : End Sub
Sub Leg7_Hit() : ComputeLegHit 6 , ActiveBall : End Sub
Sub Leg8_Hit() : ComputeLegHit 7 , ActiveBall : End Sub
Sub Leg9_Hit() : ComputeLegHit 8 , ActiveBall : End Sub
Sub Leg10_Hit() : ComputeLegHit 9 , ActiveBall : End Sub
Sub Leg11_Hit() : ComputeLegHit 10 , ActiveBall : End Sub
Sub Leg12_Hit() : ComputeLegHit 11 , ActiveBall : End Sub
Sub Leg13_Hit() : ComputeLegHit 12 , ActiveBall : End Sub
Sub Leg14_Hit() : ComputeLegHit 13 , ActiveBall : End Sub
Sub Leg15_Hit() : ComputeLegHit 14 , ActiveBall : End Sub
Sub Leg16_Hit() : ComputeLegHit 15 , ActiveBall : End Sub

Sub Leg17_Hit() : ComputeLegHit 16 , ActiveBall : End Sub
Sub Leg18_Hit() : ComputeLegHit 17 , ActiveBall : End Sub
Sub Leg19_Hit() : ComputeLegHit 18, ActiveBall : End Sub
Sub Leg20_Hit() : ComputeLegHit 19, ActiveBall : End Sub
Sub Leg21_Hit() : ComputeLegHit 20, ActiveBall : End Sub
Sub Leg22_Hit() : ComputeLegHit 21, ActiveBall : End Sub
Sub Leg23_Hit() : ComputeLegHit 22, ActiveBall : End Sub
Sub Leg24_Hit() : ComputeLegHit 23, ActiveBall : End Sub
Sub Leg25_Hit() : ComputeLegHit 24 , ActiveBall : End Sub
Sub Leg26_Hit() : ComputeLegHit 25 , ActiveBall : End Sub
Sub Leg27_Hit() : ComputeLegHit 26, ActiveBall : End Sub
Sub Leg28_Hit() : ComputeLegHit 27, ActiveBall : End Sub
Sub Leg29_Hit() : ComputeLegHit 28, ActiveBall : End Sub
Sub Leg30_Hit() : ComputeLegHit 29, ActiveBall : End Sub
Sub Leg31_Hit() : ComputeLegHit 30, ActiveBall : End Sub
Sub Leg32_Hit() : ComputeLegHit 31, ActiveBall : End Sub
Sub Leg33_Hit() : ComputeLegHit 32, ActiveBall : End Sub
Sub Leg34_Hit() : ComputeLegHit 33, ActiveBall : End Sub
Sub Leg35_Hit() : ComputeLegHit 34, ActiveBall : End Sub
Sub Leg36_Hit() : ComputeLegHit 35, ActiveBall : End Sub

Sub ComputeLegHit( LegNum, BallObj )

  ' Method: Determine coordinates of ball relative to leg and
  ' use those to determine the angle the ball hit the leg, then
  ' compute force applied to whole body.
  ' Construction notes: Items are all arranged CCW.

  ' Step 1: Determine angle of post to center, angle of hit point to
  ' center of post, and velocity/angle of ball travel.

  Dim postAngle, hitAngle, velAngle, velocity, ballAngle
  Dim dX, dY
  Dim tX, tY, lX, lY

  postAngle = AngleRef(LegNum)

  dX = BallObj.X - PostX(LegNum)
  dY = BallObj.Y - PostY(LegNum)
  hitAngle = GetAngle( dX, dY )

  dX = BallObj.X - ObjectX
  dY = BallObj.Y - ObjectY
  ballAngle = GetAngle( dX, dY )

  dX = BallObj.VelX
  dY = BallObj.VelY
  velocity = Sqr( dX*dX + dY*dY )
  velAngle = GetAngle( dX, dY )

  ' Step 2: Compute amount of force transferred to leg at hit point.
  ' This is determined by how much of the ball's velocity faces toward
  ' the center of the post.

  Dim speed, forceRatio, dAngle
  dAngle = velAngle - hitAngle
  forceRatio = cos(dAngle)    ' How close to a direct-line course is the ball on with the post?

  speed = SpeedMultiplier * 10 * forceRatio * velocity

  ' Step 3: Now, take the hit angle and compare it to the post's angle
  ' relative to the center of the object, and use that to determine the
  ' amount of force transferred to the object as a whole.

  dAngle = hitAngle - postAngle
  forceRatio = -sin(dAngle)
  speed = speed * forceRatio

' DebugBox3.Text = _
'   "vel= " & velocity & vbNewLine &_
'   "dirAng=" & velAngle & vbNewLine &_
'   "hitAng=" & hitAngle * 180 / pi & vbNewLine &_
'   "postAng=" & postAngle * 180 / pi & vbNewLine &_
'   "ballAng=" & ballAngle * 180 / pi & vbNewLine &_
'   "(Diff)=" & Abs(postAngle - ballAngle) * 180 / pi & vbNewLine &_
'   "frat=" & forceRatio & vbNewLine &_
'   "spd=" & speed

  SpinSpeed = SpinSpeed + speed
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

' Const Pi = 3.1415927

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


Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySoundAtVol "fx_spinner", Spinner, 1
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub LeftFlipper001_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper001_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


'*** High Score Handling ***

sub loadhs


Set FileObj=CreateObject("Scripting.FileSystemObject")
If Not FileObj.FolderExists(UserDirectory) then
Exit Sub
End if
If Not FileObj.FileExists(UserDirectory & "WilliamsZodiac.txt") then
Exit Sub
End if
Set ScoreFile=FileObj.GetFile(UserDirectory & "WilliamsZodiac.txt")
Set TextStr=ScoreFile.OpenAsTextStream(1,0)
If (TextStr.AtEndOfStream=True) then
Exit Sub
End if
temp1=TextStr.ReadLine
temp2=textstr.readline
temp3=textstr.readline
temp4=Textstr.ReadLine
temp5=textstr.readline
temp6=Textstr.ReadLine
temp7=TextStr.ReadLine
temp8=textstr.readline
'temp9=textstr.readline
'temp10=Textstr.ReadLine
'temp11=textstr.readline
'temp12=Textstr.ReadLine
TextStr.Close

credits = cdbl(temp1)
Playerscores(1) = CDbl(temp2)
player1100k = CDbl(temp3)
Playerscores(2) = CDbl(temp4)
player2100k = CDbl(temp5)
hiscore = cdbl(temp6)
starsign = cdbl(temp7)
lastpos = cdbl(temp8)
'horse3pos = cdbl(temp9)
'horse4pos = cdbl(temp10)
'horse5pos = cdbl(temp11)
'horse6pos = cdbl(temp12)
Set ScoreFile=Nothing
Set FileObj=Nothing
end sub

sub savehs

Dim FileObj
Dim ScoreFile
Set FileObj=CreateObject("Scripting.FileSystemObject")
If Not FileObj.FolderExists(UserDirectory) then
Exit Sub
End if
Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "WilliamsZodiac.txt",True)
scorefile.writeline credits
ScoreFile.WriteLine Playerscores(1)
ScoreFile.WriteLine Player1100k
ScoreFile.WriteLine Playerscores(2)
ScoreFile.WriteLine Player2100k
scorefile.writeline hiscore
scorefile.writeline starsign
scorefile.writeline lastpos

'scorefile.writeline horse3pos
'scorefile.writeline horse4pos
'scorefile.writeline horse5pos
'scorefile.writeline horse6pos
ScoreFile.Close
Set ScoreFile=Nothing
Set FileObj=Nothing
end sub






'*****************************************
'     BALL SHADOW
'*****************************************




Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)



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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/12))+2    '7 originally
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/12))-2    '7 originally
        End If

    ballShadow(b).Y = BOT(b).Y + 10

        If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
    Next

End Sub



'***********************
' slingshots (adapted from loserman's "Big Deal" table)
'***

Dim LStep, RStep

Sub RightSlingShot_slingshot
    PlaySoundAtVol "right_slingshot", ActiveBall, 1
  addscore(10)
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -25
    RStep = 0
    RightSlingshot.TimerEnabled = 1

End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling0.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_slingshot
    PlaySoundAtVol "left_slingshot", ActiveBall, 1
  addscore(10)
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -25
    LStep = 0
    LeftSlingShot.TimerEnabled = 1


End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing0.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub



Sub Dright1_hit
    playsound"rubber_hit_1", 0.05,0,0,1
  addscore(10)
End Sub


Sub Bleft1_hit
    playsound"rubber_hit_1", 0.05,0,0,1
  addscore(10)
End Sub


Sub Dright2_slingshot
    PlaySoundAtVol "right_slingshot", ActiveBall, 1
  addscore(10)
    rubber50.Visible = 0
    rubber51.Visible = 1
    sling4.TransZ = -25
    RStep = 0
    Dright2.TimerEnabled = 1

End Sub

Sub Dright2_Timer
    Select Case RStep
        Case 3:rubber51.Visible = 0:rubber52.Visible = 1:sling4.TransZ = -10
        Case 4:rubber52.Visible = 0:rubber50.Visible = 1:sling4.TransZ = 0:Dright2.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub Bleft2_slingshot
    playsound"left_slingshot"
  addscore(10)
    rubber40.Visible = 0
    rubber41.Visible = 1
    sling3.TransZ = -25
    LStep = 0
    Bleft2.TimerEnabled = 1


End Sub

Sub Bleft2_Timer
    Select Case LStep
        Case 3:rubber41.Visible = 0:rubber42.Visible = 1:sling3.TransZ = -10
        Case 4:rubber42.Visible = 0:rubber40.Visible = 1:sling3.TransZ = 0:Bleft2.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'**********************
'Scoring
'**********************



Sub LeftOutTrigger_Hit
  If tilt=false then
  addscore(1000)
  if lightspecialleft.state=1 then addspecial(1)
  end if
End Sub


Sub LeftTriggerA2_Hit
  If tilt=false then
  addscore(1000)
  end if
End Sub

Sub LeftTriggerA1_Hit
  If tilt=false then
  addscore(1000)
  lightA.state=1
  LeftAdvanceH.state=1
  if lightA.state=1 and lightB.state=1 Then
    shootagain.state=1
  end if
  if sl20.State = 1 then lightspecialleft.state=1
  end if
End Sub

Sub LeftTriggerUranus_Hit
  If tilt=false then
  if LeftAdvanceH.state=1 then
    starsign = starsign+1
    advancestarsign
    Controller.B2ssetdata 1, starsign
  end if
  addscore(1000)
  end if
End Sub





Sub RightOutTrigger_Hit
  If tilt=false then
  addscore(1000)
  if lightspecialright.state=1 then addspecial(1)
  end if
End Sub

Sub RightTriggerB2_Hit
  If tilt=false then
  addscore(1000)
  end if
End Sub

Sub RightTriggerB1_Hit
  If tilt=false then
  addscore(1000)
  lightB.state=1
  RightAdvanceH.state=1
  if lightA.state=1 and lightB.state=1 Then
    shootagain.state=1
  end if
  if ml20.State = 1 then lightspecialright.state=1
  end if
End Sub

Sub RightTriggerPluto_Hit
  If tilt=false then
  if RightAdvanceH.state=1 then
    starsign = starsign+1
    Controller.B2ssetdata 1, starsign
    advancestarsign
  end if
  addscore(1000)
  end if
End Sub



Sub TrigJupiter_Hit
  If tilt=false then
  addscore(100)
  lightjupiter.state=1
  if lightjupiter.state=1 and lightsaturn.state=1 and lightneptune.state=1 Then LeftBonus.state=1 end if
  end if
End Sub

Sub TrigSaturn_Hit
  If tilt=false then
  addscore(100)
  lightsaturn.state=1
  if lightjupiter.state=1 and lightsaturn.state=1 and lightneptune.state=1 Then LeftBonus.state=1 end if
  end if
End Sub

Sub TrigNepture_Hit
  If tilt=false then
  addscore(100)
  lightneptune.state=1
  if lightjupiter.state=1 and lightsaturn.state=1 and lightneptune.state=1 Then LeftBonus.state=1 end if
  end if
End Sub






Sub TrigVenus_Hit
  If tilt=false then
  addscore(100)
  lightvenus.state=1
  if lightvenus.state=1 and lightmars.state=1 and lightmercury.state=1 then RightBonus.state=1 end if
  end if
End Sub

Sub TrigMars_Hit
  If tilt=false then
  addscore(100)
  lightmars.state=1
  if lightvenus.state=1 and lightmars.state=1 and lightmercury.state=1 then RightBonus.state=1 end if
  end if
End Sub

Sub TrigMercury_Hit
  If tilt=false then
  addscore(100)
  lightmercury.state=1
  if lightvenus.state=1 and lightmars.state=1 and lightmercury.state=1 then RightBonus.state=1 end if
  end if
End Sub


Sub ButtonSun_Hit()
  If tilt=false then
'   PlaySound "sensor"
  AddScore(10)
  if LightMoonadv.state=1 then
    LightSunAdv.state=1
    LightMoonadv.state=0
  Else
    LightSunAdv.state=1
  end if
  end if
End Sub

Sub ButtonMoon_Hit()
  If tilt=false then
'   PlaySound "sensor"
    AddScore(10)
  if LightSunAdv.state=1 then
    LightMoonadv.state=1
    LightSunAdv.state=0
  Else
    LightMoonadv.state=1
  end if
  end if
End Sub

Sub advancestarsign
  if starsign = 13 then starsign = 1
      Select Case starsign
        Case 1:flaries.visible = 1:Fltaurus.visible = 0:Flgemini.visible = 0:Flcancer.visible = 0:Flleo.visible = 0:Flvergo.visible = 0:Fllibra.visible = 0:Flscorpio.visible = 0:Flsagittarius.visible = 0:Flcapricorn.visible = 0:Flaquarius.visible = 0:Flpisces.visible = 0
    flaries001.visible = 1:Fltaurus001.visible = 0:Flgemini001.visible = 0:Flcancer001.visible = 0:Flleo001.visible = 0:Flvergo001.visible = 0:Fllibra001.visible = 0:Flscorpio001.visible = 0:Flsagittarius001.visible = 0:Flcapricorn001.visible = 0:Flaquarius001.visible = 0:Flpisces001.visible = 0
        Case 2:flaries.visible = 0:Fltaurus.visible = 1:Flgemini.visible = 0:Flcancer.visible = 0:Flleo.visible = 0:Flvergo.visible = 0:Fllibra.visible = 0:Flscorpio.visible = 0:Flsagittarius.visible = 0:Flcapricorn.visible = 0:Flaquarius.visible = 0:Flpisces.visible = 0
    flaries001.visible = 0:Fltaurus001.visible = 1:Flgemini001.visible = 0:Flcancer001.visible = 0:Flleo001.visible = 0:Flvergo001.visible = 0:Fllibra001.visible = 0:Flscorpio001.visible = 0:Flsagittarius001.visible = 0:Flcapricorn001.visible = 0:Flaquarius001.visible = 0:Flpisces001.visible = 0
        Case 3:flaries.visible = 0:Fltaurus.visible = 0:Flgemini.visible = 1:Flcancer.visible = 0:Flleo.visible = 0:Flvergo.visible = 0:Fllibra.visible = 0:Flscorpio.visible = 0:Flsagittarius.visible = 0:Flcapricorn.visible = 0:Flaquarius.visible = 0:Flpisces.visible = 0
    flaries001.visible = 0:Fltaurus001.visible = 0:Flgemini001.visible = 1:Flcancer001.visible = 0:Flleo001.visible = 0:Flvergo001.visible = 0:Fllibra001.visible = 0:Flscorpio001.visible = 0:Flsagittarius001.visible = 0:Flcapricorn001.visible = 0:Flaquarius001.visible = 0:Flpisces001.visible = 0
        Case 4:flaries.visible = 0:Fltaurus.visible = 0:Flgemini.visible = 0:Flcancer.visible = 1:Flleo.visible = 0:Flvergo.visible = 0:Fllibra.visible = 0:Flscorpio.visible = 0:Flsagittarius.visible = 0:Flcapricorn.visible = 0:Flaquarius.visible = 0:Flpisces.visible = 0
    flaries001.visible = 0:Fltaurus001.visible = 0:Flgemini001.visible = 0:Flcancer001.visible = 1:Flleo001.visible = 0:Flvergo001.visible = 0:Fllibra001.visible = 0:Flscorpio001.visible = 0:Flsagittarius001.visible = 0:Flcapricorn001.visible = 0:Flaquarius001.visible = 0:Flpisces001.visible = 0
        Case 5:flaries.visible = 0:Fltaurus.visible = 0:Flgemini.visible = 0:Flcancer.visible = 0:Flleo.visible = 1:Flvergo.visible = 0:Fllibra.visible = 0:Flscorpio.visible = 0:Flsagittarius.visible = 0:Flcapricorn.visible = 0:Flaquarius.visible = 0:Flpisces.visible = 0
    flaries001.visible = 0:Fltaurus001.visible = 0:Flgemini001.visible = 0:Flcancer001.visible = 0:Flleo001.visible = 1:Flvergo001.visible = 0:Fllibra001.visible = 0:Flscorpio001.visible = 0:Flsagittarius001.visible = 0:Flcapricorn001.visible = 0:Flaquarius001.visible = 0:Flpisces001.visible = 0
        Case 6:flaries.visible = 0:Fltaurus.visible = 0:Flgemini.visible = 0:Flcancer.visible = 0:Flleo.visible = 0:Flvergo.visible = 1:Fllibra.visible = 0:Flscorpio.visible = 0:Flsagittarius.visible = 0:Flcapricorn.visible = 0:Flaquarius.visible = 0:Flpisces.visible = 0
    flaries001.visible = 0:Fltaurus001.visible = 0:Flgemini001.visible = 0:Flcancer001.visible = 0:Flleo001.visible = 0:Flvergo001.visible = 1:Fllibra001.visible = 0:Flscorpio001.visible = 0:Flsagittarius001.visible = 0:Flcapricorn001.visible = 0:Flaquarius001.visible = 0:Flpisces001.visible = 0
        Case 7:flaries.visible = 0:Fltaurus.visible = 0:Flgemini.visible = 0:Flcancer.visible = 0:Flleo.visible = 0:Flvergo.visible = 0:Fllibra.visible = 1:Flscorpio.visible = 0:Flsagittarius.visible = 0:Flcapricorn.visible = 0:Flaquarius.visible = 0:Flpisces.visible = 0
    flaries001.visible = 0:Fltaurus001.visible = 0:Flgemini001.visible = 0:Flcancer001.visible = 0:Flleo001.visible = 0:Flvergo001.visible = 0:Fllibra001.visible = 1:Flscorpio001.visible = 0:Flsagittarius001.visible = 0:Flcapricorn001.visible = 0:Flaquarius001.visible = 0:Flpisces001.visible = 0
        Case 8:flaries.visible = 0:Fltaurus.visible = 0:Flgemini.visible = 0:Flcancer.visible = 0:Flleo.visible = 0:Flvergo.visible = 0:Fllibra.visible = 0:Flscorpio.visible = 1:Flsagittarius.visible = 0:Flcapricorn.visible = 0:Flaquarius.visible = 0:Flpisces.visible = 0
    flaries001.visible = 0:Fltaurus001.visible = 0:Flgemini001.visible = 0:Flcancer001.visible = 0:Flleo001.visible = 0:Flvergo001.visible = 0:Fllibra001.visible = 0:Flscorpio001.visible = 1:Flsagittarius001.visible = 0:Flcapricorn001.visible = 0:Flaquarius001.visible = 0:Flpisces001.visible = 0
    Case 9:flaries.visible = 0:Fltaurus.visible = 0:Flgemini.visible = 0:Flcancer.visible = 0:Flleo.visible = 0:Flvergo.visible = 0:Fllibra.visible = 0:Flscorpio.visible = 0:Flsagittarius.visible = 1:Flcapricorn.visible = 0:Flaquarius.visible = 0:Flpisces.visible = 0
    flaries001.visible = 0:Fltaurus001.visible = 0:Flgemini001.visible = 0:Flcancer001.visible = 0:Flleo001.visible = 0:Flvergo001.visible = 0:Fllibra001.visible = 0:Flscorpio001.visible = 0:Flsagittarius001.visible = 1:Flcapricorn001.visible = 0:Flaquarius001.visible = 0:Flpisces001.visible = 0
        Case 10:flaries.visible = 0:Fltaurus.visible = 0:Flgemini.visible = 0:Flcancer.visible = 0:Flleo.visible = 0:Flvergo.visible = 0:Fllibra.visible = 0:Flscorpio.visible = 0:Flsagittarius.visible = 0:Flcapricorn.visible = 1:Flaquarius.visible = 0:Flpisces.visible = 0
    flaries001.visible = 0:Fltaurus001.visible = 0:Flgemini001.visible = 0:Flcancer001.visible = 0:Flleo001.visible = 0:Flvergo001.visible = 0:Fllibra001.visible = 0:Flscorpio001.visible = 0:Flsagittarius001.visible = 0:Flcapricorn001.visible = 1:Flaquarius001.visible = 0:Flpisces001.visible = 0
        Case 11:flaries.visible = 0:Fltaurus.visible = 0:Flgemini.visible = 0:Flcancer.visible = 0:Flleo.visible = 0:Flvergo.visible = 0:Fllibra.visible = 0:Flscorpio.visible = 0:Flsagittarius.visible = 0:Flcapricorn.visible = 0:Flaquarius.visible = 1:Flpisces.visible = 0
    flaries001.visible = 0:Fltaurus001.visible = 0:Flgemini001.visible = 0:Flcancer001.visible = 0:Flleo001.visible = 0:Flvergo001.visible = 0:Fllibra001.visible = 0:Flscorpio001.visible = 0:Flsagittarius001.visible = 0:Flcapricorn001.visible = 0:Flaquarius001.visible = 1:Flpisces001.visible = 0
        Case 12:flaries.visible = 0:Fltaurus.visible = 0:Flgemini.visible = 0:Flcancer.visible = 0:Flleo.visible = 0:Flvergo.visible = 0:Fllibra.visible = 0:Flscorpio.visible = 0:Flsagittarius.visible = 0:Flcapricorn.visible = 0:Flaquarius.visible = 0:Flpisces.visible = 1
    flaries001.visible = 0:Fltaurus001.visible = 0:Flgemini001.visible = 0:Flcancer001.visible = 0:Flleo001.visible = 0:Flvergo001.visible = 0:Fllibra001.visible = 0:Flscorpio001.visible = 0:Flsagittarius001.visible = 0:Flcapricorn001.visible = 0:Flaquarius001.visible = 0:Flpisces001.visible = 1
End Select
  if starsign = 12 then addspecial(3)
    Controller.B2ssetdata 1, starsign

End sub



'**********************
'Switches Magnets
'**********************


Sub suntrigger_Hit
  if leftBonus.state=1 then
  mMagnetS.MagnetOn = 1
  collectsbonus
  else addscore(1000)
  end if
End Sub

Sub MBAntiGrab1_timer()
  mMagnetS.MagnetOn = 0
  MBAntiGrab1.enabled = 0
End Sub



Sub moontrigger_Hit
  if RightBonus.state=1 then
  mMagnetM.MagnetOn = 1
  collectmbonus
  else addscore(1000)
  end if
End Sub



Sub MBAntiGrab2_timer()
  mMagnetM.MagnetOn = 0
  MBAntiGrab2.enabled = 0
End Sub




'***** reset score


Dim i

sub resettimer_timer

      rst=rst+1

      for i=1 to 2

      if Playerscores(i)>0 then

      tenk(i) = Int(Playerscores(i)/10000)
      onek(i) = Int((Playerscores(i)-tenk(i)*10000)/1000)
      hun(i) = Int((Playerscores(i)-tenk(i)*10000-onek(i)*1000)/100)
      tens(i) = Int((Playerscores(i)-tenk(i)*10000-onek(i)*1000-hun(i)*100)/10)

      end if

      if tenk(i)>0 then tenk(i)=tenk(i)-1 else tenk(i)=0: end if
      if onek(i)>0 then onek(i)=onek(i)-1 else onek(i)=0: end if
      if hun(i)>0 then hun(i)=hun(i)-1 else hun(i)=0: end if
      if tens(i)>0 then tens(i)=tens(i)-1 else tens(i)=0: end if

      Playerscores(i)=tenk(i)*10000+onek(i)*1000+hun(i)*100+tens(i)*10


      UpdateReels i ,i ,Playerscores(i), -1 , -1, -1, -1, -1


      next


if rst=14 then
'playsound "ballrelease"
end if

if rst=18 then
resettimer.enabled=false
end if
end sub



'**********************
'End of ball clean up
'**********************

Sub Resetball

lightA.state=0
lightB.state=0
lightvenus.state=0
lightmars.state=0
lightmercury.state=0
RightBonus.state=0
lightjupiter.state=0
lightsaturn.state=0
lightneptune.state=0
LeftBonus.state=0
LeftAdvanceH.state=0
RightAdvanceH.state=0
LightSpecialLeft.state=0
LightSpecialright.state=0
Sbonus=1
Mbonus=1
UpdateSBonusLights
UpdateMBonusLights
if starsign => 12 then starsign=1 :advancestarsign : end if
Controller.B2SSetTilt 33,0
    Controller.B2ssetdata 1, starsign
end sub

'**********************
'End of game clean up
'**********************

Sub ResetGame

Shootagain.state=0

if starsign => 12 then starsign=1 :advancestarsign : end if
    Controller.B2ssetdata 1, starsign
player1100k=0: player2100k=0

updatepostit

Sbonus=1
Mbonus=1
UpdateSBonusLights
UpdateMBonusLights


  flasher7.visible=1
  flasher8.visible=1
  flasher9.visible=1    'jupiter
  flasher10.visible=1   'earth
  flasher11.visible=1
  flasher12.visible=1   'saturn
  flasher13.visible=1
  flasher14.visible=1

for each Object in ColFlGameOver : object.visible = 0 : next
PlayerScores(1)=0
PlayerScores(2)=0
Controller.B2ssetscore 1, playerScores(1)
Controller.B2ssetscore 2, playerScores(2)

  End Sub



'*** General points scoring ***

sub addscore(point)

if tilt=false then

bell=0
Controller.B2ssetscore 1, playerScores(1)
Controller.B2ssetscore 2, playerScores(2)

if point = 10 or point = 100 or point = 1000 or point = 10000 then scn=1
if point = 50000 or point = 5000 or point = 500 or point = 50 then scn2=5
'if point = 30000 then scn2=3

'if point = 50000 then
'scn3=5
'bell=10000
'end if

'if point = 30000 then
'scn3=3
'bell=10000
'end if

if point = 10000 then
scn1=1
bell=10000
end if

if point = 5000 then
scn3=5
bell=1000
end if

if point = 500 then
scn3=5
bell=100
end if

if point = 50 then
scn3=5
bell=10
end if

if point = 1000 then
bell=1000
scn1=1
end if

if point = 100 then
bell=100
scn1=1
end if

if point = 10 then
bell=10
scn1=1
end if

if point = 10 or point = 100 or point = 1000 or point = 10000 then scn1=0 : scntimer.enabled=true : end if
'if point = 50000 then scn3=0 : scntimer1.enabled=true : end if
'if point = 30000 then scn3=0 : scntimer1.enabled=true : end if
if point = 5000 then scn3=0 : scntimer1.enabled=true : end if
if point = 500 then scn3=0 : scntimer1.enabled=true : end if
if point = 50 then scn3=0 : scntimer1.enabled=true : end if

end if
end sub

sub scntimer1_timer

prevscore(currentp)=playerScores(currentp)

scn3=scn3 + 1


if bell=1000 then playsound "1000_Point_Bell",0, 0.45, 0, 0: PlayerScores(currentp)=PlayerScores(currentp)+1000:checkfreegame:end if
if bell=100 then playsound "100_Point_Bell",0, 0.3, 0, 0: PlayerScores(currentp)=PlayerScores(currentp)+100:checkfreegame:end if
if bell=10 then playsound "10_Point_Bell",0, 0.15, 0, 0 : PlayerScores(currentp)=PlayerScores(currentp)+10:checkfreegame:end if

  if playerscores(1) >99999 then
  playerscores(1) = playerscores(1) - 100000
  player1100k=1
  end if
  if playerscores(2) >99999 then
  playerscores(2) = playerscores(2) - 100000
  player2100k=1
  end if




  UpdateReels currentp,currentp ,playerscores(currentp), -1,-1,-1,-1,-1


  if scn3=scn2 then scntimer1.enabled=false
end sub

sub scntimer_timer

prevscore(currentp)=playerScores(currentp)

scn1=scn1 + 1

if bell=1000 then playsound "1000_Point_Bell",0, 0.45, 0, 0: PlayerScores(currentp)=PlayerScores(currentp)+1000:checkfreegame:end if
if bell=100 then playsound "100_Point_Bell",0, 0.3, 0, 0: PlayerScores(currentp)=PlayerScores(currentp)+100:checkfreegame:end if
if bell=10 then playsound "10_Point_Bell",0, 0.15, 0, 0 : PlayerScores(currentp)=PlayerScores(currentp)+10:checkfreegame:end if

  if playerscores(1) >99999 then
  playerscores(1) = playerscores(1) - 100000
  player1100k=1
  end if
  if playerscores(2) >99999 then
  playerscores(2) = playerscores(2) - 100000
  player2100k=1
  end if




  UpdateReels currentp,currentp ,playerscores(currentp), -1,-1,-1,-1,-1


if scn1=scn then scntimer.enabled=false
end sub



'**** flashing backGlassOn

DIM min
Dim max
Dim disce1
Dim disce2
Dim disce3
Dim disce4
DIM disce5
DIM disce6
Dim zeit


Sub licht_Timer
min =0
max =6
disce1=(Int((max-min+1)*Rnd+min))
min =0
max =6
disce2=(Int((max-min+1)*Rnd+min))
min =0
max =6
disce3=(Int((max-min+1)*Rnd+min))
min =0
max =6
disce4=(Int((max-min+1)*Rnd+min))
min =0
max =6
disce5=(Int((max-min+1)*Rnd+min))
min =0
max =6
disce6=(Int((max-min+1)*Rnd+min))


If disce1 >0 then: FL_Z.visible=1:FL_O.visible=1:FL_D.visible=1: end if
If disce1 =0 then: FL_Z.visible=0:FL_O.visible=0:FL_D.visible=0: end if
If disce2 >0 then: FL_I.visible=1:FL_A.visible=1:FL_C.visible=1: end if
If disce2 =0 then: FL_I.visible=0:FL_A.visible=0:FL_C.visible=0: end if
If disce3 >0 then flasher3.visible=1
If disce3 =0 then flasher3.visible=0
If disce4 >0 then flasher4.visible=1
If disce4 =0 then flasher4.visible=0
If disce5 >0 then flasher5.visible=1
If disce5 =0 then flasher5.visible=0
If disce6 >0 then flasher6.visible=1
If disce6 =0 then flasher6.visible=0

min =250
max =500
zeit=(Int((max-min+1)*Rnd+min))
licht.Interval =zeit
End Sub




'************************************************Post It Note Section**************************************************************************
'***************Static Post It Note Update
Dim  hsY, shift, scoreMil, score100K, score10K, scoreK, score100, score10, scoreUnit
Dim hsArray: hsArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma")

Sub updatePostIt
  scoreMil = Int(hiscore/1000000)
  score100K = Int( (hiscore - (scoreMil*1000000) ) / 100000)
  score10K = Int( (hiscore - (scoreMil*1000000) - (score100K*100000) ) / 10000)
  scoreK = Int( (hiscore - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) ) / 1000)
  score100 = Int( (hiscore - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) ) / 100)
  score10 = Int( (hiscore - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) - (score100*100) ) / 10)
  scoreUnit = (hiscore - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) - (score100*100) - (score10*10) )

  pScore6.image = hsArray(scoreMil):If hiscore < 1000000 Then pScore6.image = hsArray(10)
  pScore5.image = hsArray(score100K):If hiscore < 100000 Then pScore5.image = hsArray(10)
  pScore4.image = hsArray(score10K):If hiscore < 10000 Then pScore4.image = hsArray(10)
  pScore3.image = hsArray(scoreK):If hiscore < 1000 Then pScore3.image = hsArray(10)
  pScore2.image = hsArray(score100):If hiscore < 100 Then pScore2.image = hsArray(10)
  pScore1.image = hsArray(score10):If hiscore < 10 Then pScore1.image = hsArray(10)
  pScore0.image = hsArray(scoreUnit):If hiscore < 1 Then pScore0.image = hsArray(10)
  If hiscore < 1000 Then
    PComma.image = hsArray(10)
  Else
    pComma.image = hsArray(11)
  End If
' If highScore(0) < 1000000 Then
'   pComma1.image = hsArray(10)
' Else
'   pComma1.image = hsArray(11)
' End If
  If hiscore > 999999 Then shift = 0 :pComma.transx = 0
  If hiscore < 1000000 Then shift = 1:pComma.transx = -10
  If hiscore < 100000 Then shift = 2:pComma.transx = -20
  If hiscore < 10000 Then shift = 3:pComma.transx = -30
  For hsY = 0 to 6
    EVAL("Pscore" & hsY).transx = (-10 * shift)
  Next
End Sub




'************************************************Bonus Section**************************************************************************

'*******************
'     BONUS SUN
'*******************




Sub CollectSBonus
If leftbonus.state = 1 then SBonusMultiplier=1 end if
    Select Case SBonusMultiplier
        Case 1:SBonusCountTimer.Interval = 250
        Case 2:SBonusCountTimer.Interval = 500
    End Select
    SBonusCountTimer.Enabled = 1
End Sub

Sub SBonusCountTimer_Timer
    If SBonus > 0 Then
        AddScore(1000 * SBonusMultiplier)
        UpdateSBonusLights
        SBonus = SBonus -1
    Else
        SBonusCountTimer.Enabled = 0
        UpdateSBonusLights
        SBonus = 1:UpdateSBonusLights
        MBAntiGrab1.enabled = 1
    End If
End Sub

Sub UpdateSBonusLights
    Select Case SBonus
        Case 0:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 0:sl20.State = 0
        Case 1:sl1.State = 1:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 0:sl20.State = 0
        Case 2:sl1.State = 0:sl2.State = 1:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 0:sl20.State = 0
        Case 3:sl1.State = 0:sl2.State = 0:sl3.State = 1:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 0:sl20.State = 0
        Case 4:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 1:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 0:sl20.State = 0
        Case 5:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 1:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 0:sl20.State = 0
        Case 6:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 1:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 0:sl20.State = 0
        Case 7:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 1:sl8.State = 0:sl9.State = 0:sl10.State = 0:sl20.State = 0
        Case 8:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 1:sl9.State = 0:sl10.State = 0:sl20.State = 0
        Case 9:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 1:sl10.State = 0:sl20.State = 0
        Case 10:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 1:sl20.State = 0
        Case 11:sl1.State = 1:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 1:sl20.State = 0
        Case 12:sl1.State = 0:sl2.State = 1:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 1:sl20.State = 0
        Case 13:sl1.State = 0:sl2.State = 0:sl3.State = 1:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 1:sl20.State = 0
        Case 14:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 1:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 1:sl20.State = 0
        Case 15:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 1:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 1:sl20.State = 0
        Case 16:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 1:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 1:sl20.State = 0
        Case 17:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 1:sl8.State = 0:sl9.State = 0:sl10.State = 1:sl20.State = 0
        Case 18:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 1:sl9.State = 0:sl10.State = 1:sl20.State = 0
        Case 19:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 1:sl10.State = 1:sl20.State = 0
        Case 20:sl1.State = 0:sl2.State = 0:sl3.State = 0:sl4.State = 0:sl5.State = 0:sl6.State = 0:sl7.State = 0:sl8.State = 0:sl9.State = 0:sl10.State = 0:sl20.State = 1
 End Select

if sl20.State = 1 and lightA.state=1 then   lightspecialleft.state=1

End Sub

'*******************
'     BONUS Moon
'*******************




Sub CollectMBonus
If rightbonus.state = 1 then MBonusMultiplier=1 end if
    Select Case MBonusMultiplier
        Case 1:MBonusCountTimer.Interval = 250
        Case 2:MBonusCountTimer.Interval = 500
    End Select
    MBonusCountTimer.Enabled = 1
End Sub

Sub MBonusCountTimer_Timer
    If MBonus > 0 Then
        AddScore(1000 * MBonusMultiplier)
        UpdateMBonusLights
        MBonus = MBonus -1
    Else
        MBonusCountTimer.Enabled = 0
        UpdateMBonusLights
        MBonus = 1:UpdateMBonusLights
        MBAntiGrab2.enabled = 1
    End If
End Sub

Sub UpdateMBonusLights
    Select Case MBonus
        Case 0:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 0:ml20.State = 0
        Case 1:ml1.State = 1:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 0:ml20.State = 0
        Case 2:ml1.State = 0:ml2.State = 1:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 0:ml20.State = 0
        Case 3:ml1.State = 0:ml2.State = 0:ml3.State = 1:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 0:ml20.State = 0
        Case 4:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 1:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 0:ml20.State = 0
        Case 5:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 1:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 0:ml20.State = 0
        Case 6:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 1:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 0:ml20.State = 0
        Case 7:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 1:ml8.State = 0:ml9.State = 0:ml10.State = 0:ml20.State = 0
        Case 8:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 1:ml9.State = 0:ml10.State = 0:ml20.State = 0
        Case 9:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 1:ml10.State = 0:ml20.State = 0
        Case 10:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 1:ml20.State = 0
        Case 11:ml1.State = 1:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 1:ml20.State = 0
        Case 12:ml1.State = 0:ml2.State = 1:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 1:ml20.State = 0
        Case 13:ml1.State = 0:ml2.State = 0:ml3.State = 1:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 1:ml20.State = 0
        Case 14:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 1:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 1:ml20.State = 0
        Case 15:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 1:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 1:ml20.State = 0
        Case 16:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 1:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 1:ml20.State = 0
        Case 17:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 1:ml8.State = 0:ml9.State = 0:ml10.State = 1:ml20.State = 0
        Case 18:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 1:ml9.State = 0:ml10.State = 1:ml20.State = 0
        Case 19:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 1:ml10.State = 1:ml20.State = 0
        Case 20:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0:ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0:ml9.State = 0:ml10.State = 0:ml20.State = 1
 End Select

if ml20.State = 1 and lightB.state=1 then   lightspecialright.state=1

End Sub

Sub Table1_Exit()
  turnoff
  savehs
End Sub
