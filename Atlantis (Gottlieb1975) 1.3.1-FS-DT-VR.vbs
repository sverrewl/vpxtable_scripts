'***********************************************************************
'*        Gottlieb's Atlantis (1975)                                   *
'*        Table primary build/scripted by Loserman!                    *
'*        Table images/lighting/rebuild with primitives by Bodydump    *
'*                                                                     *
'*    EM Underdogs update version 1.3 June 2025                      *
'*      Made with VPX 10.8.0 Final                                     *
'*      Press F12 in game for list of updates or see Table Info        *
'***********************************************************************


Option Explicit
SetLocale 1033
Randomize
ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "Atlantis_1975"


Const ShadowFlippersOn = true
Const ShadowBallOn = true
Const ShadowConfigFile = false
Const VRTest = 0

'----- VR Options -----
Dim VRRoomChoice : VRRoomChoice = 1       ' 1 = 360 Room 2 = Minimal Room

'----- VR Room Auto-Detect -----
Dim VRRoom, VR_Obj, VRMode

If RenderingMode = 2 or VRTest = 1 Then
  VRMode = True
  Ramp17.visible = 0
  PLeftSideRail.visible = 0
  PRightSideRail.visible = 0
  For Each VR_Obj in VRCab : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in VRMinRoom : VR_Obj.Visible = 1 : Next
  If VRRoomChoice = 1 Then
    For Each VR_Obj in VRMinRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRSphere : VR_Obj.Visible = 1 : Next
    VRBackRightLeg.z = 50
    VRBackleftLeg.z = 35
    VRLeftFrontLeg.z = 15
  Else
    For Each VR_Obj in VRMinRoom : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in VRSphere : VR_Obj.Visible = 0 : Next
    VRBackRightLeg.z = 165
    VRBackleftLeg.z = 165
    VRLeftFrontLeg.z = 40
  End If
Else
  VRMode = False
  For Each VR_Obj in VRCab : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRMinRoom : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRSphere : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in Backglass_flashers : VR_Obj.Visible = 0 : Next
End If


'******************************************************
'*******  Set Up Backglass Flashers *******
'******************************************************

Sub SetBackglass()
  Dim obj

  For Each obj In Backglass_flashers
    obj.x = obj.x - 10
    obj.height = - obj.y + 138
    obj.y = 36 'adjusts the distance from the backglass towards the user
  Next
End Sub


' ***************************************************************************
'          BASIC FSS(EM) 1-4 player 5x drums, 1 credit drum CORE CODE
' ****************************************************************************


' ********************* POSITION EM REEL DRUMS ON BACKGLASS *************************

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen
Dim inx
xoff =463
yoff = 0
zoff =825
xrot = -90

Const USEEMS = 1 ' 1-4 set between 1 to 4 based on number of players

const idx_emp1r1 =0 'player 1
const idx_emp2r1 =5 'player 2
const idx_emp3r1 =10 'player 3
const idx_emp4r1 =15 'player 4
const idx_emp4r6 =20 'credits


Dim BGObjEM(1)
if USEEMS = 1 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
'elseif USEEMS = 2 then
' BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
' emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
' Empty,Empty,Empty,Empty,Empty,_
' Empty,Empty,Empty,Empty,Empty,_
' emp4r6) ' credits
'elseif USEEMS = 3 then
' BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
' emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
' emp3r1, emp3r2, emp3r3, emp3r4, emp3r5, _
' Empty,Empty,Empty,Empty,Empty,_
' emp4r6) ' credits
'elseif USEEMS = 4 then
' BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
' emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
' emp3r1, emp3r2, emp3r3, emp3r4, emp3r5, _
' emp4r1, emp4r2, emp4r3, emp4r4, emp4r5, _
' emp4r6) ' credits
end If

Sub center_objects_em()
Dim cnt,ii, xx, yy, yfact, xfact, objs
'exit sub
yoff = -150
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
    yoff = -78
    end if

  xx =objs.x

  objs.x = (xoff - xcen) + xx + xfact
  yy = objs.y
  objs.y =yoff

    If yy < 0 then
    yy = yy * -1
    end if

  objs.z = (zoff - ycen) + yy - (yy * zscale) + yfact

  'objs.rotx = xrot
  end if
  cnt = cnt + 1
  Next

end sub



' ********************* UPDATE EM REEL DRUMS CORE LIB *************************

Dim cred,ix, np,npp, reels(5, 7), scores(6,2)

'reset scores to defaults
for np =0 to 5
scores(np,0 ) = 0
scores(np,1 ) = 0
Next

'reset EM drums to defaults
For np =0 to 3
  For  npp =0 to 6
  reels(np, npp) =0 ' default to zero
  Next
Next


Sub SetScore(player, ndx , val)

Dim ncnt

  if player = 5 or player = 6 then
    if val > 0 then
      If(ndx = 0)Then ncnt = val * 10
      If(ndx = 1)Then ncnt = val

      scores(player, 0) = scores(player, 0) + ncnt
    end if
  else
    if val > 0 then

    If(ndx = 0)then ncnt = val * 10000
    If(ndx = 1)then ncnt = val * 1000
    If(ndx = 2)Then ncnt = val * 100
    If(ndx = 3)Then ncnt = val * 10
    If(ndx = 4)Then ncnt = val

    scores(player, 0) = scores(player, 0) + ncnt
    'scores(player, 0) + ncnt

    end if
  end if
End Sub


Sub SetDrum(player, drum , val)
Dim cnt
Dim objs : objs =BGObjEM(0)

  If val = 0 then
    Select case player
    case -1: ' the credit drum
    If Not IsEmpty(objs(idx_emp4r6)) then
    objs(idx_emp4r6).ObjrotX = 0 ' 285
    'cnt =objs(idx_emp4r6).ObjrotX
    end if
    Case 0:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp1r1)) then: objs(idx_emp1r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp1r1+1)) then: objs(idx_emp1r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp1r1+2)) then: objs(idx_emp1r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp1r1+3)) then: objs(idx_emp1r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp1r1+4)) then: objs(idx_emp1r1+4).ObjrotX=0: end if
    End Select
    Case 1:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp2r1)) then: objs(idx_emp2r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp2r1+1)) then: objs(idx_emp2r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp2r1+2)) then: objs(idx_emp2r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp2r1+3)) then: objs(idx_emp2r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp2r1+4)) then: objs(idx_emp2r1+4).ObjrotX=0: end if
    End Select
    Case 2:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp3r1)) then: objs(idx_emp3r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp3r1+1)) then: objs(idx_emp3r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp3r1+2)) then: objs(idx_emp3r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp3r1+3)) then: objs(idx_emp3r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp3r1+4)) then: objs(idx_emp3r1+4).ObjrotX=0: end if
    End Select
    Case 3:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp4r1)) then: objs(idx_emp4r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp4r1+1)) then: objs(idx_emp4r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp4r1+2)) then: objs(idx_emp4r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp4r1+3)) then: objs(idx_emp4r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp4r1+4)) then: objs(idx_emp4r1+4).ObjrotX=0: end if
    End Select
  End Select

  else
  Select case player

    Case -1: ' the credit drum
    'emp4r6.ObjrotX = emp4r6.ObjrotX + val
    If Not IsEmpty(objs(idx_emp4r6)) then
    objs(idx_emp4r6).ObjrotX = objs(idx_emp4r6).ObjrotX + val
    end if

    Case 0:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp1r1)) then: objs(idx_emp1r1).ObjrotX= objs(idx_emp1r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp1r1+1)) then: objs(idx_emp1r1+1).ObjrotX= objs(idx_emp1r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp1r1+2)) then: objs(idx_emp1r1+2).ObjrotX= objs(idx_emp1r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp1r1+3)) then: objs(idx_emp1r1+3).ObjrotX= objs(idx_emp1r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp1r1+4)) then: objs(idx_emp1r1+4).ObjrotX= objs(idx_emp1r1+4).ObjrotX + val: end if
    End Select
    Case 1:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp2r1)) then: objs(idx_emp2r1).ObjrotX= objs(idx_emp2r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp2r1+1)) then: objs(idx_emp2r1+1).ObjrotX= objs(idx_emp2r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp2r1+2)) then: objs(idx_emp2r1+2).ObjrotX= objs(idx_emp2r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp2r1+3)) then: objs(idx_emp2r1+3).ObjrotX= objs(idx_emp2r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp2r1+4)) then: objs(idx_emp2r1+4).ObjrotX= objs(idx_emp2r1+4).ObjrotX + val: end if
    End Select
    Case 2:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp3r1)) then: objs(idx_emp3r1).ObjrotX= objs(idx_emp3r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp3r1+1)) then: objs(idx_emp3r1+1).ObjrotX= objs(idx_emp3r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp3r1+2)) then: objs(idx_emp3r1+2).ObjrotX= objs(idx_emp3r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp3r1+3)) then: objs(idx_emp3r1+3).ObjrotX= objs(idx_emp3r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp3r1+4)) then: objs(idx_emp3r1+4).ObjrotX= objs(idx_emp3r1+4).ObjrotX + val: end if
    End Select
    Case 3:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp4r1)) then: objs(idx_emp4r1).ObjrotX= objs(idx_emp4r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp4r1+1)) then: objs(idx_emp4r1+1).ObjrotX= objs(idx_emp4r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp4r1+2)) then: objs(idx_emp4r1+2).ObjrotX= objs(idx_emp4r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp4r1+3)) then: objs(idx_emp4r1+3).ObjrotX= objs(idx_emp4r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp4r1+4)) then: objs(idx_emp4r1+4).ObjrotX= objs(idx_emp4r1+4).ObjrotX + val: end if
    End Select

  End Select
  end if
End Sub


Sub SetReel(player, drum, val)

'TextBox1.text = "playr:" & player +1 & " drum:" & drum & "val:" & val

Dim  inc , cur, dif, fix, fval

inc = 33.5
fval = -5 ' graphic seam between 5 & 6 fix value, easier to fix here than photoshop

If  (player <= 3) or (drum = -1) then

  If drum = -1 then drum = 0

  cur =reels(player, drum)

  If val <> cur then ' something has changed
  Select Case drum

    Case 0: ' credits drum

      if val > cur then
        dif =val - cur
        fix =0
          If cur < 5 and cur+dif > 5 then
          fix = fix- fval
          end if
        dif = dif * inc

        dif = dif-fix

        SetDrum -1,0,  -dif
      Else
        if val = 0 Then
        SetDrum -1,0,  0' reset the drum to abs. zero
        Else
        dif = 11 - cur
        dif = dif + val

        dif = dif * inc
        dif = dif-fval

        SetDrum -1,0,   -dif
        end if
      end if
    Case 1:
    'TB1.text = val
    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc

      dif = dif-fix

      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val

      dif = dif * inc
      dif = dif-fval

      SetDrum player,drum,   -dif
      end if

    end if
    reels(player, drum) = val

    Case 2:
    'TB2.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if
    end if
    reels(player, drum) = val

    Case 3:
    'TB3.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix

      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    reels(player, drum) = val

    Case 4:
    'TB4.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    reels(player, drum) = val

    Case 5:
    'TB5.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    reels(player, drum) = val
   End Select

  end if
end if
End Sub

Dim EMMODE: EMMODE = 0
'Dim Score10,Score100,Score1000,Score10000,Score100000
Dim Score1000,Score10000,Score100000, ActivePLayer
Dim nplayer,playr,value,curscr,curplayr

' EMMODE = 1
'ex: UpdateReels 0-3, 0-3, 0-199999, n/a, n/a, n/a, n/a, n/a, n/a
' EMMODE = 0
'ex: UpdateReels 0-3, 0-3, n/a, 0-1,0-99999 ,0-9999, 0-999, 0-99, 0-9

Sub UpdateReels (Player,nReels ,nScore, n100K, Score10000 ,Score1000,Score100,Score10,Score1)

' to-do find out if player is one or zero based, if 1 based subtract 1.
value =nScore'Score(Player)
  nplayer = Player -1

  curscr = value
  curplayr = nplayer


scores(0,1) = scores(0,0)
  scores(0,0) = 0
  scores(1,1) = scores(1,0)
  scores(1,0) = 0
  scores(2,1) = scores(2,0)
  scores(2,0) = 0
  scores(3,1) = scores(3,0)
  scores(3,0) = 0

  For  ix =0 to 6
    reels(0, ix) =0
    reels(1, ix) =0
    reels(2, ix) =0
    reels(3, ix) =0
  Next

  For  ix =0 to 4

  SetDrum ix, 1 , 0
  SetDrum ix, 2 , 0
  SetDrum ix, 3 , 0
  SetDrum ix, 4 , 0
  SetDrum ix, 5 , 0

  Next

  For playr =0 to nReels

    if EMMODE = 0 then
    If (ActivePLayer) = playr Then
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
      value =scores(playr, 1) ' store score
      nplayer = playr
      end if

    scores(playr, 0)  = 0 ' reset score

    if(value >= 100000) then

      'if nplayer = 0 then: FL100K1.visible = 1
      'if nplayer = 1 then: FL100K2.visible = 1
      'if nplayer = 2 then: FL100K3.visible = 1
      'if nplayer = 3 then: FL100K4.visible = 1

    value = value - 100000

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

'******************* VR Plunger **********************

Sub TimerVRPlunger_Timer

  If VR_Primary_plunger.Y < 2142.833 then
       VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
  End If
End Sub

Sub TimerVRPlunger2_Timer
  VR_Primary_plunger.Y = 2071.833 + (5* Plunger.Position) -20
End Sub

'****************************************************************

Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed
Dim B2SOn   'True/False if want backglass

Const HSFileName="Atlantis_75VPX.txt"
Const B2STableName="Atlantis_1975"
Const LMEMTableConfig="LMEMTables.txt"
Const LMEMShadowConfig="LMEMShadows.txt"
Dim EnableBallShadow
Dim EnableFlipperShadow

Const DOF_FLIPPER_LEFT = 101
Const DOF_FLIPPER_RIGHT = 102
Const DOF_BUMPER_1 = 106
Const DOF_BUMPER_2 = 105
Const DOF_TARGET_BASE = 108
Const DOF_TRIGGER_BASE = 240
Const DOF_BALL_RELEASE = 107
Const DOF_GATE_RELATED = 121
Const DOF_BUTTON_BASE = 270
Const DOF_START_BUTTON = 116
Const DOF_KNOCKER = 115
Const DOF_DRAIN_RELATED = 215
Const DOF_CHIME_UNIT_HIGH = 128
Const DOF_CHIME_UNIT_MID = 129
Const DOF_CHIME_UNIT_LOW = 130
Const DOF_SLINGSHOT_LEFT = 103
Const DOF_SLINGSHOT_RIGHT = 104

'* this value adjusts score motor behavior - 0 allows you to continue scoring while the score motor is running - 1 sets score motor to behave more like a real EM
Const ScoreMotorAdjustment=1

'* this is a debug setting to use an older scoring routine vs a newer score routine - don't change this value
Const ScoreAdditionAdjustment=1

'* this controls whether you hear bells (0) or chimes (1) when scoring
Const ChimesOn=1


dim ScoreChecker
dim CheckAllScores
dim sortscores(4)
dim sortplayers(4)
Dim TextStr,TextStr2,TiltEndsGame
Dim i
Dim obj
Dim bgpos
Dim kgpos
Dim dooralreadyopen
Dim kgdooralreadyopen
Dim TargetSpecialLit
Dim Points210counter
Dim Points500counter
Dim Points1000counter
Dim Points2000counter
Dim BallsPerGame
Dim InProgress
Dim BallInPlay
Dim BIPL
Dim BallInPlungerLane
Dim CreditsPerCoin
Dim Score100K(4)
Dim Score(4)
Dim ScoreDisplay(4)
Dim HighScorePaid(4)
Dim HighScore
Dim HighScoreReward
Dim BonusMultiplier
Dim Credits
Dim Match
Dim Replay1
Dim Replay2
Dim Replay3
Dim Replay1Paid(4)
Dim Replay2Paid(4)
Dim Replay3Paid(4)
Dim TableTilted
Dim TiltCount

Dim OperatorMenu

Dim BonusBooster
Dim BonusBoosterCounter
Dim BonusCounter
Dim HoleCounter

Dim Ones
Dim Tens
Dim Hundreds
Dim Thousands

Dim Player
Dim Players

Dim rst
Dim bonuscountdown
Dim TempMultiCounter
dim TempPlayerup
dim RotatorTemp

Dim bump1
Dim bump2
Dim bump3

Dim LastChime10
Dim LastChime100
Dim LastChime1000

Dim Score10
Dim Score100

Dim targettempscore
Dim SpecialLightCounter
Dim SpecialLightFlag
Dim HorseshoeCounter
Dim DropTargetCounter

Dim LeftTargetCounter
Dim RightTargetCounter

Dim MotorRunning
Dim Replay1Table(15)
Dim Replay2Table(15)
Dim Replay3Table(15)
Dim ReplayTableSet
Dim ReplayLevel
Dim ReplayTableMax
Dim BonusSpecialThreshold
Dim TargetLightsOn
Dim AdvanceLightCounter
dim bonustempcounter

Dim Bumper1State,Bumper2State

Dim xx
Dim LStep, RStep
dim Object

Sub Table1_Init()
  If Table1.ShowDT = false then
    For each obj in DesktopCrap
      obj.visible=False
    next
  End If

  If VRMode = True Then
    SetBackglass
    center_objects_em
  End If

  EMSInit Table1
  LoadEM
  LoadLMEMConfig2
  HideOptions
  SetupReplayTables
  PlasticsOff
  BumpersOff
  OperatorMenu=0
  HighScore=0
  MotorRunning=0
  HighScoreReward=1
  Credits=0
  BallsPerGame=5
  ReplayLevel=1
  BonusSpecialThreshold=1
  SpecialLightFlag=2
  TiltEndsGame=1
  loadhs
  if HighScore=0 then HighScore=50000


  TableTilted=false

  Match=int(Rnd*10)*10
  MatchReel.SetValue((Match/10)+1)

  CanPlayReel.SetValue(0)
  GameOverReel.SetValue(1)
  TiltReel.SetValue(0)
  For each obj in PlayerHuds
    obj.SetValue(0)
  next
  For each obj in PlayerScores
    obj.ResetToZero
  next

  for each obj in bottgate
    obj.isdropped=true
    next

  for each obj in Bonus
    obj.state=0
  next
  for each obj in TargetSpecialLights
    obj.state=0
  next
  for each obj in TargetAdvanceLights
    obj.state=0
  next
  for each obj in RolloverAdvanceLights
    obj.state=0
  next

  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)

  BonusCounter=0
  HoleCounter=0
    bgpos=6
  kgpos=0

  dooralreadyopen=0
  kgdooralreadyopen=0

  InstructCard.image="IC"+FormatNumber(BallsPerGame,0)

  RefreshReplayCard


  Bumper1Light.state=0
  BumperLight1.state=0
  Bumper2Light.state=0
  BumperLight2.state=0

  TargetSpecialLit = 0
  Points210counter=0
  Points500counter=0
  Points1000counter=0
  Points2000counter=0

  AdvanceLightCounter=0

  BonusBooster=3
  BonusBoosterCounter=0
  Players=0
  RotatorTemp=1
  InProgress=false
  TargetLightsOn=false

  ScoreText.text=HighScore


  If B2SOn Then


    Controller.B2SSetMatch Match
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScoreRolloverPlayer2 0
    Controller.B2SSetScoreRolloverPlayer3 0
    Controller.B2SSetScoreRolloverPlayer4 0

    'Controller.B2SSetScore 3,HighScore
    Controller.B2SSetTilt 0
    Controller.B2SSetCredits Credits
    Controller.B2SSetGameOver 1
  End If

  for i=1 to 4
    player=i
    If B2SOn Then
      Controller.B2SSetScorePlayer player, 0
    End If
  next
  bump1=1
  bump2=1
  bump3=1
  OperatorMenuBackdrop.image = "PostitBL"
  For XOpt = 1 to MaxOption
    Eval("OperatorOption"&XOpt).image = "PostitBL"
  next

  For XOpt = 1 to 256
    Eval("Option"&XOpt).image = "PostItBL"
  next
  InitPauser5.enabled=true
  If Credits > 0 Then DOF 116, DOFOn
  If VRMode = True Then
    for each Object in ColFlGameOver : object.visible = 1 : next
  End If
End Sub

Sub Table1_exit()
  savehs
  SaveLMEMConfig
  SaveLMEMConfig2
  Controller.Stop
end sub


Sub Table1_KeyDown(ByVal keycode)

  If keycode = LeftFlipperKey Then
    Primary_flipper_button_left.X = Primary_flipper_button_left.X + 4
  End If

  If keycode = RightFlipperKey Then
    Primary_flipper_button_right.X = Primary_flipper_button_right.X - 4
  End If

  If keycode = PlungerKey Then
    TimerVRPlunger.Enabled = True
        TimerVRPlunger2.Enabled = False
  End If

  if keycode=StartGameKey Then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y - 4
  End if

  ' GNMOD
  if EnteringInitials then
    CollectInitials(keycode)
    exit sub
  end if


  if EnteringOptions then
    CollectOptions(keycode)
    exit sub
  end if



  If keycode = PlungerKey Then
    Plunger.PullBack
    PlungerPulled = 1
    EMSPlayPlungerPullSound Plunger
  End If

  if keycode = LeftFlipperKey and InProgress = false then
    OperatorMenuTimer.Enabled = true
  end if
  ' END GNMOD

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    LeftFlipper.RotateToEnd
    PlaySound "buzzL",-1
    DOF DOF_FLIPPER_LEFT, DOFOn
    EMSPlayLeftFlipperActivateSound LeftFlipper
  End If

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToEnd
    PlaySound "buzz",-1
    DOF DOF_FLIPPER_RIGHT, DOFOn
    EMSPlayRightFlipperActivateSound RightFlipper
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
    EMSPlayNudgeSound
    TiltIt
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
    EMSPlayNudgeSound
    TiltIt
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
    EMSPlayNudgeSound
    TiltIt
  End If

  If keycode = MechanicalTilt Then
    TiltCount=2
    TiltIt
  End If

  If keycode = AddCreditKey or keycode = 4 then
    If B2SOn Then
      'Controller.B2SSetScorePlayer6 HighScore

    End If
    EMSPlayCoinSound
    AddSpecial2
  end if

   if keycode = 5 then
    EMSPlayCoinSound
    AddSpecial2
    keycode= StartGameKey
  end if



  if keycode=StartGameKey and Credits>0 and InProgress=false and Players=0 and EnteringOptions = 0 then
'GNMOD
    OperatorMenuTimer.Enabled = false
'END GNMOD
    Credits=Credits-1
    If Credits < 1 Then DOF 116, DOFOff
    CreditsReel.SetValue(Credits)
    Players=1
    CanPlayReel.SetValue(Players)
    MatchReel.SetValue(0)
    Player=1
    EMSPlayStartupSound
    TempPlayerUp=Player
    rst=0
    BallInPlay=1
    InProgress=true
    resettimer.enabled=true
    BonusMultiplier=1
    If B2SOn Then
      Controller.B2SSetTilt 0
      Controller.B2SSetGameOver 0
      Controller.B2SSetMatch 0
      Controller.B2SSetCredits Credits
      Controller.B2SSetCanPlay 1
      Controller.B2SSetPlayerUp 1
'     Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetScoreRolloverPlayer1 0
    End If
    For each obj in PlayerScores
'     obj.ResetToZero
    next
    For each obj in PlayerHuds
      obj.SetValue(0)
    next
    PlayerHuds(Player-1).SetValue(1)
    GameOverReel.SetValue(0)
    If VRMode = True Then
      for each Object in ColFlPlayer1 : object.visible = 1 : next
      for each Object in ColFlTilt : object.visible = 0 : next
      for each Object in ColFlGameOver : object.visible = 0 : next
      for each Object in ColFlMatch : object.visible = 0 : next

      cred =reels(4, 0)
      reels(4, 0) = 0
      SetDrum -1,0,  0

      SetReel 0,-1,  Credits
      reels(4, 0) = Credits
    End If
  end if



End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = LeftFlipperKey Then
    Primary_flipper_button_left.X = Primary_flipper_button_left.X - 4
  End If

  If keycode = RightFlipperKey Then
    Primary_flipper_button_right.X = Primary_flipper_button_right.X + 4
  End If


  If keycode = PlungerKey Then
    TimerVRPlunger.Enabled = False
        TimerVRPlunger2.Enabled = True
  End If

  if keycode=StartGameKey Then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y + 4
  End if

  ' GNMOD
  if EnteringInitials then
    exit sub
  end if

  If keycode = PlungerKey Then
    if PlungerPulled = 0 then
      exit sub
    end if
    EMSPlayPlungerReleaseBallSound Plunger
    Plunger.Fire
    If BallInPlungerLane = 1 Then
    EMSPlayPlungerReleaseBallSound Plunger 'Plunger release sound when there is a ball in shooter lane
    Else
    EMSPlayPlungerReleaseNoBallSound Plunger 'Plunger release sound when there is no ball in shooter lane
    End If
  End If

  if keycode = LeftFlipperKey then
    OperatorMenuTimer.Enabled = false
  end if

  ' END GNMOD

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    LeftFlipper.RotateToStart
    EMSPlayLeftFlipperDeactivateSound LeftFlipper
    StopSound "buzzL"
    DOF DOF_FLIPPER_LEFT, DOFOff
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToStart
        EMSPlayRightFlipperDeactivateSound RightFlipper
        StopSound "buzz"
        DOF DOF_FLIPPER_RIGHT, DOFOff
  End If

End Sub

Sub BIPL_hit()
  bBallInPlungerLane = True
End Sub

Sub BIPL_unhit()
  bBallInPlungerLane = False
  EMSPlayPlungerReleaseNoBallSound Plunger
End Sub

Sub Drain_Hit()

  Drain.DestroyBall
  EMSPlayDrainSound Drain
    DOF DOF_DRAIN_RELATED, DOFPulse
  Pause4Bonustimer.enabled=true

End Sub

Sub Trigger001_Hit()
  DOF 110, DOFPulse

End Sub

Sub Pause4Bonustimer_timer
  Pause4Bonustimer.enabled=0
  NextBallDelay.enabled=true

End Sub
'***********************
'     Flipper Logos
'***********************

Sub UpdateFlipperLogos_Timer
  LFlip.RotZ = LeftFlipper.CurrentAngle
  RFlip.RotZ = RightFlipper.CurrentAngle
  PGate.Rotz = (Gate.CurrentAngle*.75) + 25
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
End Sub

'***********************
' slingshots
'

Sub RightSlingShot_Slingshot
  EMSPlayRightSlingshotSound Sling1
  DOF 104, DOFPulse
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  AddScore 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling0.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  EMSPlayLeftSlingshotSound Sling2
  DOF 103, DOFPulse
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  AddScore 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing0.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'***********************************
' Walls
'***********************************
Sub RubberWall24_Hit()

  If TableTilted=false then
    AddScore(100)
  End if

End Sub
Sub RubberWall14_Hit()

  If TableTilted=false then
    AddScore(100)
  End if

End Sub

Sub RubberWall13_Hit()

  If TableTilted=false then
    AddScore(100)
  End if
End Sub

Sub RubberWall11_Hit()

  If TableTilted=false then
    SetMotor(500)
  End if
End Sub

'***********************************
' Bumpers
'***********************************


Sub Bumper1_Hit
  If TableTilted=false then

    EMSPlayMiddleBumperSound Bumper1
        DOF DOF_BUMPER_1, DOFPulse
    bump1 = 1
    AddScore(100)
    end if

End Sub

Sub Bumper2_Hit
  If TableTilted=false then

    EMSPlayMiddleBumperSound Bumper2
        DOF DOF_BUMPER_2, DOFPulse
    bump2 = 1
    AddScore(100)

    end if

End Sub


'****************************
' Stationary targets
'****************************

Sub TargetRight1_Hit()
  If TableTilted=false then
    AddScore(100)
    If AdvanceLight1.state=1 then
      IncreaseBonus
    end if
    if SpecialLight1.state=1 then
      AddSpecial
    end if
  end if
  EMSPlayTargetHitSound
  'PlaySound SoundFXDOF("target",126,DOFPulse,DOFTargets)
end Sub


Sub TargetRight2_Hit()
  If TableTilted=false then
    AddScore(100)
    If AdvanceLight2.state=1 then
      IncreaseBonus
    end if
    if SpecialLight2.state=1 then
      AddSpecial
    end if
  end if
  EMSPlayTargetHitSound
  'PlaySound SoundFXDOF("target",126,DOFPulse,DOFTargets)
end Sub


Sub TargetRight3_Hit()
  If TableTilted=false then
    AddScore(100)
    If AdvanceLight3.state=1 then
      IncreaseBonus
    end if
    if SpecialLight3.state=1 then
      AddSpecial
    end if
  end if
  EMSPlayTargetHitSound
  'PlaySound SoundFXDOF("target",127,DOFPulse,DOFTargets)
end Sub


'************************************
'  Rollover lanes
'************************************
Sub TriggerTop1_Hit()
  If TableTilted=false then
    If UpperLight1.state=1 then
      SetMotor(3000)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub

Sub TriggerTop2_Hit()
  If TableTilted=false then
    If UpperLight2.state=1 then
      SetMotor(3000)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub

Sub TriggerTop3_Hit()
  If TableTilted=false then
    If UpperLight3.state=1 then
      SetMotor(3000)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub

Sub Trigger1_Hit()
  If TableTilted=false then
    If RolloverLight1.state=1 then
      SetMotor(3000)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub

Sub Trigger2_Hit()
  If TableTilted=false then
    If RolloverLight2.state=1 then
      SetMotor(3000)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub

Sub Trigger3_Hit()
  If TableTilted=false then
    If RolloverLight3.state=1 then
      SetMotor(3000)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub

Sub Trigger4_Hit()
  If TableTilted=false then
    If RolloverLight4.state=1 then
      SetMotor(3000)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub

Sub Trigger5_Hit()
  If TableTilted=false then
    If RolloverLight5.state=1 then
      SetMotor(3000)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub

Sub Trigger6_Hit()
  If TableTilted=false then
    If RolloverLight6.state=1 then
      SetMotor(3000)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub

Sub Trigger7_Hit()
  If TableTilted=false then
    If RolloverLight7.state=1 then
      SetMotor(3000)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub

Sub Trigger8_Hit()
  If TableTilted=false then
    If RolloverLight8.state=1 then
      SetMotor(3000)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub

Sub Trigger9_Hit()
  If TableTilted=false then
    If RolloverLight9.state=1 then
      SetMotor(3000)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub

Sub TriggerSpecial_Hit()
  If TableTilted=false then
    If OutlaneSpecialLight.state=1 then
      AddSpecial
      SetMotor(1000)
    else
      SetMotor(1000)
    end if
  end if
    EMSPlaySensorSound
        DOF DOF_TRIGGER_BASE, DOFPulse
End Sub


Sub CenterTrigger1_Hit()
  If TableTilted=false then
    Addscore(10)
    TargetLightsOn=False
    for each obj in TargetSpecialLights
      obj.state=0
    next
    for each obj in TargetAdvanceLights
      obj.state=0
    next
  end if
  EMSPlaySensorSound
    DOF DOF_TRIGGER_BASE, DOFPulse
End sub

Sub CenterTrigger2_Hit()
  If TableTilted=false then
    Addscore(10)
    TargetLightsOn=True
    if BonusCounter<9 then
      TargetAdvanceLights(AdvanceLightCounter).state=1
      if BallsPerGame=3 and AdvanceLightCounter<>1 then
        TargetAdvanceLights(0).state=1
        TargetAdvanceLights(2).state=1
      end if
    else
      TargetSpecialLights(AdvanceLightCounter).state=1
      if BallsPerGame=3 and AdvanceLightCounter<>1 then
        TargetSpecialLights(0).state=1
        TargetSpecialLights(2).state=1
      end if
    end if
  end if
  EMSPlaySensorSound
    DOF DOF_TRIGGER_BASE, DOFPulse
End sub

'********************************
' drop targets
'********************************

Sub DropTarget1_Hit()
  DOF 122, DOFPulse
  if TableTilted=false then
    if DropTarget1.IsDropped=0 then
      DropTargetCounter=DropTargetCounter+1
      DropTarget1.IsDropped=1
      Double1.IsDropped=1
      Shadow_DropTarget1.visible=False
      EMSPlayDropTargetDropSound DropTarget1
      SetMotor(500)
      CheckAllDrops
    end if
  end if
end sub

Sub DropTarget2_Hit()
  DOF 122, DOFPulse
  if TableTilted=false then
    if DropTarget2.IsDropped=0 then
      DropTargetCounter=DropTargetCounter+1
      DropTarget2.IsDropped=1
      Double2.IsDropped=1
      Double1.IsDropped=1
      Shadow_DropTarget2.visible=False
      EMSPlayDropTargetDropSound DropTarget2
      SetMotor(500)
      CheckAllDrops
    end if
  end if
end sub

Sub DropTarget3_Hit()
  DOF 122, DOFPulse
  if TableTilted=false then
    if DropTarget3.IsDropped=0 then
      DropTargetCounter=DropTargetCounter+1
      DropTarget3.IsDropped=1
      Double3.IsDropped=1
      Double2.IsDropped=1
      Shadow_DropTarget3.visible=False
      EMSPlayDropTargetDropSound DropTarget3
      SetMotor(500)
      CheckAllDrops
    end if
  end if
end sub

Sub DropTarget4_Hit()
  DOF 122, DOFPulse
  if TableTilted=false then
    if DropTarget4.IsDropped=0 then
      DropTargetCounter=DropTargetCounter+1
      DropTarget4.IsDropped=1
      Double4.IsDropped=1
      Double3.IsDropped=1
      Shadow_DropTarget4.visible=False
      EMSPlayDropTargetDropSound DropTarget4
      SetMotor(500)
      CheckAllDrops
    end if
  end if
end sub

Sub DropTarget5_Hit()
  DOF 122, DOFPulse
  if TableTilted=false then
    if DropTarget5.IsDropped=0 then
      DropTargetCounter=DropTargetCounter+1
      DropTarget5.IsDropped=1
      Double5.IsDropped=1
      Double4.IsDropped=1
      EMSPlayDropTargetDropSound DropTarget5
      SetMotor(500)
      CheckAllDrops
    end if
  end if
end sub

Sub DropTarget6_Hit()
  DOF 123, DOFPulse
  if TableTilted=false then
    if DropTarget6.IsDropped=0 then
      DropTargetCounter=DropTargetCounter+1
      DropTarget6.IsDropped=1
      Double6.IsDropped=1
      Double5.IsDropped=1
      Shadow_DropTarget6.visible=False
      EMSPlayDropTargetDropSound DropTarget6
      SetMotor(500)
      CheckAllDrops
    end if
  end if
end sub

Sub DropTarget7_Hit()
  DOF 123, DOFPulse
  if TableTilted=false then
    if DropTarget7.IsDropped=0 then
      DropTargetCounter=DropTargetCounter+1
      DropTarget7.IsDropped=1
      Double7.IsDropped=1
      Double6.IsDropped=1
      Shadow_DropTarget7.visible=False
      EMSPlayDropTargetDropSound DropTarget7
      SetMotor(500)
      CheckAllDrops
    end if
  end if
end sub

Sub DropTarget8_Hit()
  DOF 123, DOFPulse
  if TableTilted=false then
    if DropTarget8.IsDropped=0 then
      DropTargetCounter=DropTargetCounter+1
      DropTarget8.IsDropped=1
      Double8.IsDropped=1
      Double7.IsDropped=1
      Shadow_DropTarget8.visible=False
      EMSPlayDropTargetDropSound DropTarget8
      SetMotor(500)
      CheckAllDrops
    end if
  end if
end sub

Sub DropTarget9_Hit()
  DOF 123, DOFPulse
  if TableTilted=false then
    if DropTarget9.IsDropped=0 then
      DropTargetCounter=DropTargetCounter+1
      DropTarget9.IsDropped=1
      Double9.IsDropped=1
      Double8.IsDropped=1
      Shadow_DropTarget9.visible=False
      EMSPlayDropTargetDropSound DropTarget9
      SetMotor(500)
      CheckAllDrops
    end if
  end if
end sub

Sub DropTarget10_Hit()
  DOF 123, DOFPulse
  if TableTilted=false then
    if DropTarget10.IsDropped=0 then
      DropTargetCounter=DropTargetCounter+1
      DropTarget10.IsDropped=1
      Shadow_DropTarget10.visible=False
      EMSPlayDropTargetDropSound DropTarget10
      Double9.IsDropped=1
      SetMotor(500)
      CheckAllDrops
    end if
  end if
end sub

Sub Double1_Hit()
  DOF 122, DOFPulse
  if TableTilted=false then
    if Double1.IsDropped=0 then
      Double1.IsDropped=1
      Double2.IsDropped=1
      DropTarget1.IsDropped=1
      Shadow_DropTarget1.visible=False
      DropTarget2.IsDropped=1
      Shadow_DropTarget2.visible=False
      DropTargetCounter=DropTargetCounter+2
      DropTargetLight.state=1
      SetMotor(5000)
      CheckAllDrops
    end if
  end if
end Sub

Sub Double2_Hit()
  DOF 122, DOFPulse
  if TableTilted=false then
    if Double2.IsDropped=0 then
      Double2.IsDropped=1
      Double1.IsDropped=1
      Double3.IsDropped=1
      DropTarget2.IsDropped=1
      Shadow_DropTarget2.visible=False
      DropTarget3.IsDropped=1
      Shadow_DropTarget3.visible=False
      DropTargetCounter=DropTargetCounter+2
      DropTargetLight.state=1
      SetMotor(5000)
      CheckAllDrops
    end if
  end if
end Sub

Sub Double3_Hit()
  DOF 122, DOFPulse
  if TableTilted=false then
    if Double3.IsDropped=0 then
      Double3.IsDropped=1
      Double2.IsDropped=1
      Double4.IsDropped=1
      DropTarget3.IsDropped=1
      Shadow_DropTarget3.visible=False
      DropTarget4.IsDropped=1
      Shadow_DropTarget4.visible=False
      DropTargetCounter=DropTargetCounter+2
      DropTargetLight.state=1
      SetMotor(5000)
      CheckAllDrops
    end if
  end if
end Sub

Sub Double4_Hit()
  DOF 122, DOFPulse
  if TableTilted=false then
    if Double4.IsDropped=0 then
      Double4.IsDropped=1
      Double3.IsDropped=1
      Double5.IsDropped=1
      DropTarget4.IsDropped=1
      Shadow_DropTarget4.visible=False
      DropTarget5.IsDropped=1
      DropTargetCounter=DropTargetCounter+2
      DropTargetLight.state=1
      SetMotor(5000)
      CheckAllDrops
    end if
  end if
end Sub

Sub Double5_Hit()
  DOF 123, DOFPulse
  if TableTilted=false then
    if Double5.IsDropped=0 then
      Double5.IsDropped=1
      Double6.IsDropped=1
      Double4.IsDropped=1
      DropTarget5.IsDropped=1
      DropTarget6.IsDropped=1
      Shadow_DropTarget6.visible=False
      DropTargetCounter=DropTargetCounter+2
      DropTargetLight.state=1
      SetMotor(5000)
      CheckAllDrops
    end if
  end if
end Sub

Sub Double6_Hit()
  DOF 123, DOFPulse
  if TableTilted=false then
    if Double6.IsDropped=0 then
      Double6.IsDropped=1
      Double5.IsDropped=1
      Double7.IsDropped=1
      DropTarget6.IsDropped=1
      Shadow_DropTarget6.visible=False
      DropTarget7.IsDropped=1
      Shadow_DropTarget7.visible=False
      DropTargetCounter=DropTargetCounter+2
      DropTargetLight.state=1
      SetMotor(5000)
      CheckAllDrops
    end if
  end if
end Sub

Sub Double7_Hit()
  DOF 123, DOFPulse
  if TableTilted=false then
    if Double7.IsDropped=0 then
      Double7.IsDropped=1
      Double6.IsDropped=1
      Double8.IsDropped=1
      DropTarget7.IsDropped=1
      Shadow_DropTarget7.visible=False
      DropTarget8.IsDropped=1
      Shadow_DropTarget8.visible=False
      DropTargetCounter=DropTargetCounter+2
      DropTargetLight.state=1
      SetMotor(5000)
      CheckAllDrops
    end if
  end if
end Sub

Sub Double8_Hit()
  DOF 123, DOFPulse
  if TableTilted=false then
    if Double8.IsDropped=0 then
      Double8.IsDropped=1
      Double9.IsDropped=1
      Double7.IsDropped=1
      DropTarget8.IsDropped=1
      Shadow_DropTarget8.visible=False
      DropTarget9.IsDropped=1
      Shadow_DropTarget9.visible=False
      DropTargetCounter=DropTargetCounter+2
      DropTargetLight.state=1
      SetMotor(5000)
      CheckAllDrops
    end if
  end if
end Sub

Sub Double9_Hit()
  DOF 123, DOFPulse
  if TableTilted=false then
    if Double9.IsDropped=0 then
      Double9.IsDropped=1
      Double8.IsDropped=1
      DropTarget9.IsDropped=1
      Shadow_DropTarget9.visible=False
      DropTarget10.IsDropped=1
      Shadow_DropTarget10.visible=False
      DropTargetCounter=DropTargetCounter+2
      DropTargetLight.state=1
      SetMotor(5000)
      CheckAllDrops
    end if
  end if
end Sub


Sub ResetDrops
  for each obj in AllDropTargets
    obj.IsDropped=0
    Shadow_DropTarget1.visible=True
    Shadow_DropTarget2.visible=True
    Shadow_DropTarget3.visible=True
    Shadow_DropTarget4.visible=True
    Shadow_DropTarget6.visible=True
    Shadow_DropTarget7.visible=True
    Shadow_DropTarget8.visible=True
    Shadow_DropTarget9.visible=True
    Shadow_DropTarget10.visible=True
  next
  EMSPlayDropTargetResetSound AllDropTargets
  for each obj in Doubles
    obj.IsDropped=0
  next
  EMSPlayDropTargetResetSound Doubles
  DropTargetCounter=0
  DOF 124, DOFPulse
  DOF 125, DOFPulse
end Sub

Sub CheckAllDrops

  If DropTargetCounter=10 then
    SetMotor(5000)
    if TargetsSpecial.state=1 then
      AddSpecial
    end if
    ResetDropsTimer.enabled=1
    IncreaseBonus
  end if
end Sub

sub ResetDropsTimer_timer
    ResetDropsTimer.enabled=0
    ResetDrops
end sub

'**************************************
'**************************************


Sub CloseGateTrigger_Hit()
  if dooralreadyopen=1 then
    closeg.enabled=true
  end if
End Sub


Sub AddSpecial()
  EMSPlayKnockerSound
    DOF DOF_KNOCKER, DOFPulse
  Credits=Credits+1
  if Credits>15 then Credits=15

    cred =reels(4, 0)
    reels(4, 0) = 0
    SetDrum -1,0,  0

    SetReel 0,-1,  Credits
    reels(4, 0) = Credits

  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddSpecial2()
  EMSPlayClickSound
  DOF DOF_START_BUTTON, DOFOn
  Credits=Credits+1
  if Credits>15 then Credits=15

    cred =reels(4, 0)
    reels(4, 0) = 0
    SetDrum -1,0,  0

    SetReel 0,-1,  Credits
    reels(4, 0) = Credits

  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddBonus()
  bonuscountdown=bonuscounter
  ScoreBonus.enabled=true
End Sub

Sub ToggleBumper
  if UpperLight2.state=1 then
    for each obj in RolloverLights2
      obj.state=1
    next
    for each obj in RolloverLights1
      obj.state=0
    next
  else
    for each obj in RolloverLights2
      obj.state=0
    next
    for each obj in RolloverLights1
      obj.state=1
    next
  end if
end sub

sub ToggleAdvance
  for each obj in TargetSpecialLights
    obj.state=0
  next
  for each obj in TargetAdvanceLights
    obj.state=0
  next
  for each obj in RolloverAdvanceLights
    obj.state=0
  next


  AdvanceLightCounter=AdvanceLightCounter+1
  if AdvanceLightCounter>2 then
    AdvanceLightCounter=0
  end if
  RolloverAdvanceLights(AdvanceLightCounter).state=1
  If BallsPerGame=3 then
    if AdvanceLightCounter=0 or AdvanceLightCounter=2 then
      RolloverAdvanceLights(0).state=1
      RolloverAdvanceLights(2).state=1
    end if
  end if

  If TargetLightsOn=True then
    if BonusCounter<9 then
      TargetAdvanceLights(AdvanceLightCounter).state=1
      if BallsPerGame=3 then
        if AdvanceLightCounter=0 or AdvanceLightCounter=2 then
          TargetAdvanceLights(0).state=1
          TargetAdvanceLights(2).state=1
        end if
      end if
    else
      TargetSpecialLights(AdvanceLightCounter).state=1
      if BallsPerGame=3 then
        if AdvanceLightCounter=0 or AdvanceLightCounter=2 then
          TargetSpecialLights(0).state=1
          TargetSpecialLights(2).state=1
        end if
      end if
    end if
  end if

end sub

sub ToggleSpecial

  SpecialLightCounter=SpecialLightCounter+1
  If SpecialLightCounter>5 then
    SpecialLightCounter=0
  end if
  DropTargetLight.state=0
  Select Case SpecialLightCounter
    Case 0:
      DropTargetLight.state=0
    Case 1:
      if SpecialLightFlag=2 or SpecialLightFlag=3 then
        DropTargetLight.state=1
      end if
    Case 2:
      if SpecialLightFlag=1 or SpecialLightFlag=3 then
        DropTargetLight.state=1
      end if
    Case 3:
      if SpecialLightFlag=2 then
        DropTargetLight.state=1
      end if
    Case 4:
      if SpecialLightFlag=3 then
        DropTargetLight.state=1
      end if
    Case 5:
      DropTargetLight.state=1

  End Select

end sub

sub IncreaseHorseshoe

  HorseshoeCounter=HorseshoeCounter+1
  for each obj in HorseshoeLights
    obj.state=0
  next
  if HorseshoeCounter>9 then
    HorseshoeCounter=0
    IncreaseBonus
  else
    HorseshoeLights(HorseshoeCounter).state=1
  end if
end sub

Sub ResetBallDrops




  HoleCounter=0

End Sub


Sub LightsOut
  for each obj in Bonus
    obj.state=0
  next

  BonusCounter=0
  HoleCounter=0
  Bumper1Light.state=0
  BumperLight1.state=0
  Bumper2Light.state=0
  BumperLight2.state=0

  if dooralreadyopen=1 then
    closeg.enabled=true
  end if

end sub

Sub ResetBalls()

  TempMultiCounter=BallsPerGame-BallInPlay

  ResetBallDrops
  BonusMultiplier=1
  TableTilted=false
  TiltReel.SetValue(0)
  If VRMode = True Then
    for each Object in ColFlTilt : object.visible = 0 : next
  End If
  If B2Son then
    Controller.B2SSetTilt 0
  end if
  PlasticsOn
  BumpersOn
  'CreateBallID BallRelease
  Ballrelease.CreateSizedBall 25
    Ballrelease.Kick 40,7
  DOF DOF_BALL_RELEASE, DOFPulse
  EMSPlayBallReleaseSound 'BallRelease
  BallInPlayReel.SetValue(BallInPlay)
' InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+"_"+FormatNumber(BallInPlay,0)

End Sub

sub delaykgclose_timer
  delaykgclose.enabled=false
  closekg.enabled=true

end sub

 sub openg_timer

 end sub

sub closeg_timer
    closeg.interval=10
    bottgate(bgpos).isdropped=true
    bgpos=bgpos+1
    bottgate(bgpos).isdropped=false

  if bgpos=6 then

    closeg.enabled=false
    dooralreadyopen=0
  end if
end sub


sub resettimer_timer
    rst=rst+1

  if rst>3 and rst<14 then
    ResetReelsToZero(1)
  end if


    if rst=16 then
    EMSPlayStartBallSound 0
    end if
    if rst=18 then
    newgame
    resettimer.enabled=false
    end if
end sub

Sub ResetReelsToZero(reelzeroflag)
  dim d1(5)
  dim d2(5)
  dim scorestring1, scorestring2

  If reelzeroflag=1 then
    scorestring1=CStr(Score(1))
    scorestring1=right("00000" & scorestring1,5)
    for i=0 to 4
      d1(i)=CInt(mid(scorestring1,i+1,1))
    next
    for i=0 to 4
      if d1(i)>0 then
        d1(i)=d1(i)+1
        if d1(i)>9 then d1(i)=0
      end if

    next
    Score(1)=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)
    If B2SOn Then
      Controller.B2SSetScorePlayer 1, Score(1)
    End If
    PlayerScores(0).SetValue(Score(1))

  end if
  If reelzeroflag=2 then
    scorestring1=CStr(Score(2))
    scorestring1=right("00000" & scorestring1,5)
    for i=0 to 4
      d1(i)=CInt(mid(scorestring1,i+1,1))
    next
    for i=0 to 4
      if d1(i)>0 then
        d1(i)=d1(i)+1
        if d1(i)>9 then d1(i)=0
      end if

    next
    Score(2)=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)
    If B2SOn Then
      Controller.B2SSetScorePlayer 2, Score(2)
    End If
    PlayerScores(1).SetValue(Score(2))


  end if

end sub


sub ResetHorseshoeLights_timer
  if HorseshoeCounter=0 then
    ResetHorseshoeLights.enabled=0
    NextBallDelay.enabled=true
    exit sub
  end if
  HorseshoeLights(HorseshoeCounter).state=0
  HorseshoeCounter=HorseShoeCounter+1
  If HorseshoeCounter>9 then
    HorseshoeCounter=0
  end if
  HorseshoeLights(HorseshoeCounter).state=1
end sub

sub ScoreBonus_timer

  if bonuscountdown<=0 then
    ScoreBonus.enabled=false
    ScoreBonus.interval=600
    ResetHorseshoeLights.enabled=true
    EMSPlayMotorRunningSound
    exit sub
  end if
  if BonusMultiplier=1 then
    Playsound "BonusScore1x"
  else
    Playsound "BonusScore2x"
  end if
  if BonusCountdown>10 then
    bonustempcounter=1
    BonusScorer.enabled=true
'   SetMotor2(1000*BonusMultiplier)
    Bonus(bonuscountdown-10).state=0
  else
'   SetMotor2(1000*BonusMultiplier)
    bonustempcounter=1
    BonusScorer.enabled=true
    Bonus(bonuscountdown).state=0
  end if
  If BonusMultiplier=2 then
    Bonus(0).state=1
  end if
  If BonusMultiplier=1 then
    ScoreBonus.interval=400
  Else
    ScoreBonus.interval=400
  end if
  bonuscountdown=bonuscountdown-1

  if bonuscountdown>10 then
    Bonus(bonuscountdown-10).state=1
    exit sub
  end if
  if bonuscountdown>0 then
    Bonus(bonuscountdown).state=1
  end if

end sub

sub BonusScorer_timer

  select case bonustempcounter
    case 0:
      BonusScorer.enabled=false
      if BonusMultiplier=2 then
        AddBonusScore(1000)
      end if
    case 1:
      bonustempcounter=bonustempcounter-1
      AddBonusScore(1000)

  end select
end sub

sub NextBallDelay_timer()
  NextBallDelay.enabled=false
  nextball

end sub

sub newgame
  InProgress=true
  queuedscore=0
  for i = 1 to 4
    Score(i)=0
    Score100K(1)=0
    HighScorePaid(i)=false
    Replay1Paid(i)=false
    Replay2Paid(i)=false
    Replay3Paid(i)=false
  next
  If VRMode = True Then
    for each Object in ColFlTilt : object.visible = 0 : next
    for each Object in ColFlGameOver : object.visible = 0 : next
    for each Object in ColFlMatch : object.visible = 0 : next'
    FlBIP1A.visible = 1
  End If

  If B2SOn Then
    Controller.B2SSetTilt 0
    Controller.B2SSetGameOver 0
    Controller.B2SSetMatch 0
'   Controller.B2SSetScorePlayer1 0
'   Controller.B2SSetScorePlayer2 0
'   Controller.B2SSetScorePlayer3 0
'   Controller.B2SSetScorePlayer4 0
    Controller.B2SSetBallInPlay BallInPlay
  End if
  for each obj in TargetSpecialLights
    obj.state=0
  next
  for each obj in RolloverAdvanceLights
    obj.state=0
  next
  for each obj in TargetAdvanceLights
    obj.state=0
  next
  Bumper1Light.state=1
  BumperLight1.state=1
  Bumper2Light.state=1
  BumperLight2.state=1
  CenterTriggerLight1.state=1
  CenterTriggerLight2.state=1
  SpecialLightCounter=0
  DropTargetLight.state=0
  TargetsSpecial.state=0
  OutlaneSpecialLight.state=0
  AdvanceLightCounter=0
  RolloverAdvanceLights(AdvanceLightCounter).state=1
  if BallsPerGame=3 then
    RolloverAdvanceLights(2).state=1
  end if
  ResetDrops
  BonusCounter=0
  IncreaseBonus
  ToggleBumper
  ResetBalls
End sub

sub nextball

  Player=Player+1
  If Player>Players Then
    BallInPlay=BallInPlay+1
    If BallInPlay>BallsPerGame then
      EMSPlayMotorLeerSound
      InProgress=false

      If VRMode = True Then
        for each Object in ColFlGameOver : object.visible = 1 : next
        for each Object in ColFlPlayer1 : object.visible = 0 : next
        for each Object in ColFlBallinPlay : object.visible = 0 : next
      End If
      If B2SOn Then
        Controller.B2SSetGameOver 1
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetCanPlay 0
      End If
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      GameOverReel.SetValue(1)
'     InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+"_0"

      BallInPlayReel.SetValue(0)
      CanPlayReel.SetValue(0)
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      LightsOut
      BumpersOff
      PlasticsOff
      checkmatch
      CheckHighScore
      Players=0
      HighScoreTimer.interval=100
      HighScoreTimer.enabled=true
    Else
      Player=1
      If VRMode = True Then
        If BallInPlay = 2 Then FlBIP2A.visible = 1 : FlBIP1A.visible = 0 Else  FlBIP2A.visible = 0 End If
        If BallInPlay = 3 Then FlBIP3A.visible = 1 Else FlBIP3A.visible = 0 End If
        If BallInPlay = 4 Then FlBIP4A.visible = 1 Else FlBIP4A.visible = 0 End If
        If BallInPlay = 5 Then FlBIP5A.visible = 1 Else FlBIP5A.visible = 0 End If
      End If
      If B2SOn Then
        Controller.B2SSetPlayerUp Player
        Controller.B2SSetBallInPlay BallInPlay

      End If
'     EMSPlayRotateThroughPlayersSound
      TempPlayerUp=Player
'     PlayerUpRotator.enabled=true
      PlayStartBall.enabled=true
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      PlayerHuds(Player-1).SetValue(1)

      ResetBalls
    End If
  Else
    If VRMode = True Then
      If BallInPlay = 2 Then FlBIP2A.visible = 1 : FlBIP1A.visible = 0 Else  FlBIP2A.visible = 0 End If
      If BallInPlay = 3 Then FlBIP3A.visible = 1 Else FlBIP3A.visible = 0 End If
      If BallInPlay = 4 Then FlBIP4A.visible = 1 Else FlBIP4A.visible = 0 End If
      If BallInPlay = 5 Then FlBIP5A.visible = 1 Else FlBIP5A.visible = 0 End If
    End If
    If B2SOn Then
      Controller.B2SSetPlayerUp Player
      Controller.B2SSetBallInPlay BallInPlay
    End If
'   PlaySound("RotateThruPlayers")
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=true
    PlayStartBall.enabled=true
    For each obj in PlayerHuds
      obj.SetValue(0)
    next
    PlayerHuds(Player-1).SetValue(1)
    ResetBalls
  End If

End sub

sub CheckHighScore
  Dim playertops
    dim si
  dim sj
  dim stemp
  dim stempplayers
  for i=1 to 4
    sortscores(i)=0
    sortplayers(i)=0
  next
  playertops=0
  for i = 1 to Players
    sortscores(i)=Score(i)
    sortplayers(i)=i
  next

  for si = 1 to Players
    for sj = 1 to Players-1
      if sortscores(sj)>sortscores(sj+1) then
        stemp=sortscores(sj+1)
        stempplayers=sortplayers(sj+1)
        sortscores(sj+1)=sortscores(sj)
        sortplayers(sj+1)=sortplayers(sj)
        sortscores(sj)=stemp
        sortplayers(sj)=stempplayers
      end if
    next
  next
  ScoreChecker=4
  CheckAllScores=1
  NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)
  savehs
end sub


sub checkmatch
  Dim tempmatch
  if TableTilted=true and TiltEndsGame=1 then
    exit sub
  end if
  tempmatch=Int(Rnd*10)
  Match=tempmatch*10
  MatchReel.SetValue(tempmatch+1)

  If VRMode = True Then
    If Match = 0 Then FlM00.visible = 1 : FlM00A.visible = 1 : FlM00B.visible = 1 Else FlM00.visible = 0 : FlM00A.visible = 0 : FlM00B.visible = 0  End If
    If Match = 10 Then FlM10.visible = 1 : FlM10A.visible = 1 : FlM10B.visible = 1 Else FlM10.visible = 0 : FlM10A.visible = 0 : FlM10B.visible = 0 End If
    If Match = 20 Then FlM20.visible = 1 : FlM20A.visible = 1 : FlM20B.visible = 1 Else FlM20.visible = 0 : FlM20A.visible = 0 : FlM20B.visible = 0 End If
    If Match = 30 Then FlM30.visible = 1 : FlM30A.visible = 1 : FlM30B.visible = 1 Else FlM30.visible = 0 : FlM30A.visible = 0 : FlM30B.visible = 0 End If
    If Match = 40 Then FlM40.visible = 1 : FlM40A.visible = 1 : FlM40B.visible = 1 Else FlM40.visible = 0 : FlM40A.visible = 0 : FlM40B.visible = 0 End If
    If Match = 50 Then FlM50.visible = 1 : FlM50A.visible = 1 : FlM50B.visible = 1 Else FlM50.visible = 0 : FlM50A.visible = 0 : FlM50B.visible = 0 End If
    If Match = 60 Then FlM60.visible = 1 : FlM60A.visible = 1 : FlM60B.visible = 1 Else FlM60.visible = 0 : FlM60A.visible = 0 : FlM60B.visible = 0 End If
    If Match = 70 Then FlM70.visible = 1 : FlM70A.visible = 1 : FlM70B.visible = 1 Else FlM70.visible = 0 : FlM70A.visible = 0 : FlM70B.visible = 0 End If
    If Match = 80 Then FlM80.visible = 1 : FlM80A.visible = 1 : FlM80B.visible = 1 Else FlM80.visible = 0 : FlM80A.visible = 0 : FlM80B.visible = 0 End If
    If Match = 90 Then FlM90.visible = 1 : FlM90A.visible = 1 : FlM90B.visible = 1 Else FlM90.visible = 0 : FlM90A.visible = 0 : FlM90B.visible = 0 End If
  End If

  If B2SOn Then
    If Match = 0 Then
      Controller.B2SSetMatch 100
    Else
      Controller.B2SSetMatch Match
    End If
  End if
  for i = 1 to Players
    if Match=(Score(i) mod 100) then
      AddSpecial
    end if
  next
end sub

Sub TiltTimer_Timer()
  if TiltCount > 0 then TiltCount = TiltCount - 1
  if TiltCount = 0 then
    TiltTimer.Enabled = False
  end if
end sub

Sub TiltIt()
    TiltCount = TiltCount + 1
    if TiltCount = 3 then
      TableTilted=True
      TiltReel.SetValue(1)
      If TiltEndsGame=1 then
        BallInPlay=BallsPerGame
        If VRMode = True Then
          for each Object in ColFlGameOver : object.visible = 1 : next
          for each Object in ColFlBallinPlay : object.visible = 0 : next
          for each Object in ColFlPlayer1 : object.visible = 0 : next
        End If
        If B2SOn Then
          Controller.B2SSetGameOver 1
          Controller.B2SSetPlayerUp 0
          Controller.B2SSetBallInPlay 0
          Controller.B2SSetCanPlay 0
        End If
        For each obj in PlayerHuds
          obj.SetValue(0)
        next
        GameOverReel.SetValue(1)
  '     InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+"_0"

        BallInPlayReel.SetValue(0)
      end if

      PlasticsOff
      BumpersOff
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      If VRMode = True Then
        for each Object in ColFlTilt : object.visible = 1 : next
      End If
      If B2Son then
        Controller.B2SSetTilt 1
      end if
    else
      TiltTimer.Interval = 500
      TiltTimer.Enabled = True
    end if

end sub

Sub IncreaseBonus()


  If BonusCounter=10 then
    exit sub
  end if
  If BonusCounter<10 then
    Bonus(BonusCounter).state=0
  end if
  BonusCounter=BonusCounter+1

  Bonus(BonusCounter).State=1
  If BonusCounter=10 then

    TargetsSpecial.state=1
    OutlaneSpecialLight.state=1
  end if

' LeftOutlaneLight.state=0
' RightOutlaneLight.state=0

' if (BonusSpecialThreshold=0) then
'   if (BonusCounter=8) or (BonusCounter=12) then
'     LeftOutlaneLight.state=1
'     RightOutlaneLight.state=1
'   end if
' else
'   if (BonusCounter=10) or (BonusCounter=15) then
'     LeftOutlaneLight.state=1
'     RightOutlaneLight.state=1
'   end if
' end if
End Sub


Sub BonusBoost_Timer()
  IncreaseBonus
  BonusBoosterCounter=BonusBoosterCounter-1
  If BonusBoosterCounter=0 then
    BonusBoost.enabled=false
  end if

end sub

Sub CheckForLightSpecial()

  if (TopLightA.state=0) and (TopLightB.state=0) and (TopLightC.state=0) then
    TopRightTargetLight.State=1
    TopLeftTargetLight.State=1
  end if
end sub

Sub PlayStartBall_timer()

  PlayStartBall.enabled=false
  EMSPlayStartBallSound 1
end sub

Sub PlayerUpRotator_timer()
    If RotatorTemp<5 then
      TempPlayerUp=TempPlayerUp+1
      If TempPlayerUp>4 then
        TempPlayerUp=1
      end if
      If B2SOn Then
        Controller.B2SSetPlayerUp TempPlayerUp
      End If

    else
      if B2SOn then
        Controller.B2SSetPlayerUp Player
      end if
      PlayerUpRotator.enabled=false
      RotatorTemp=1
    end if
    RotatorTemp=RotatorTemp+1


end sub

sub savehs
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HSFileName,True)
    ScoreFile.WriteLine 0
    ScoreFile.WriteLine Credits
    scorefile.writeline BallsPerGame
    scorefile.writeline SpecialLightFlag & TiltEndsGame
    scorefile.writeline ReplayLevel

    for xx=1 to 5
      scorefile.writeline HSScore(xx)
    next
    for xx=1 to 5
      scorefile.writeline HSName(xx)
    next

    ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub

sub loadhs
    ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
    dim temp1
    dim temp2
  dim temp3
  dim temp4
  dim temp5
  dim temp6
  dim temp7
  dim temp8
  dim temp9
  dim temp10
  dim temp11
  dim temp12
  dim temp13
  dim temp14
  dim temp15

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & HSFileName) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & HSFileName)
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    temp1=TextStr.ReadLine
    temp2=textstr.readline
    temp3=textstr.readline
    temp4=textstr.readline
    temp5=textstr.readline
    HighScore=cdbl(temp1)
    if HighScore<1 then
      temp6=textstr.readline
      temp7=textstr.readline
      temp8=textstr.readline
      temp9=textstr.readline
      temp10=textstr.readline
      temp11=textstr.readline
      temp12=textstr.readline
      temp13=textstr.readline
      temp14=textstr.readline
      temp15=textstr.readline
    end if

    TextStr.Close
      Credits=cdbl(temp2)
    If VRMode = True Then

      cred =reels(4, 0)
      reels(4, 0) = 0
      SetDrum -1,0,  0

      SetReel 0,-1,  Credits
      reels(4, 0) = Credits
    End If
    BallsPerGame=cdbl(temp3)
    If Len(temp4)<2 then
      SpecialLightFlag=2
      TiltEndsGame=1
    else
      SpecialLightFlag=Mid(cdbl(temp4),1,1)
      TiltEndsGame=Mid(cdbl(temp4),2,1)
    end if
    ReplayLevel=cdbl(temp5)
    if HighScore<1 then
      HSScore(1) = int(temp6)
      HSScore(2) = int(temp7)
      HSScore(3) = int(temp8)
      HSScore(4) = int(temp9)
      HSScore(5) = int(temp10)

      HSName(1) = temp11
      HSName(2) = temp12
      HSName(3) = temp13
      HSName(4) = temp14
      HSName(5) = temp15
    end if
    Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub

sub SaveLMEMConfig
  Dim FileObj
  Dim LMConfig
  dim temp1
  dim tempb2s
  tempb2s=0
  if B2SOn=true then
    tempb2s=1
  else
    tempb2s=0
  end if
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set LMConfig=FileObj.CreateTextFile(UserDirectory & LMEMTableConfig,True)
  LMConfig.WriteLine tempb2s
  LMConfig.Close
  Set LMConfig=Nothing
  Set FileObj=Nothing

end Sub

sub LoadLMEMConfig
  Dim FileObj
  Dim LMConfig
  dim tempC
  dim tempb2s

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & LMEMTableConfig) then
    Exit Sub
  End if
  Set LMConfig=FileObj.GetFile(UserDirectory & LMEMTableConfig)
  Set TextStr2=LMConfig.OpenAsTextStream(1,0)
  If (TextStr2.AtEndOfStream=True) then
    Exit Sub
  End if
  tempC=TextStr2.ReadLine
  TextStr2.Close
  tempb2s=cdbl(tempC)
  if tempb2s=0 then
    B2SOn=false
  else
    B2SOn=true
  end if
  Set LMConfig=Nothing
  Set FileObj=Nothing
end sub

sub SaveLMEMConfig2
  If ShadowConfigFile=false then exit sub
  Dim FileObj
  Dim LMConfig2
  dim temp1
  dim temp2
  dim tempBS
  dim tempFS

  if EnableBallShadow=true then
    tempBS=1
  else
    tempBS=0
  end if
  if EnableFlipperShadow=true then
    tempFS=1
  else
    tempFS=0
  end if

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set LMConfig2=FileObj.CreateTextFile(UserDirectory & LMEMShadowConfig,True)
  LMConfig2.WriteLine tempBS
  LMConfig2.WriteLine tempFS
  LMConfig2.Close
  Set LMConfig2=Nothing
  Set FileObj=Nothing

end Sub

sub LoadLMEMConfig2
  If ShadowConfigFile=false then
    EnableBallShadow = ShadowBallOn
    BallShadowUpdate.enabled = ShadowBallOn
    EnableFlipperShadow = ShadowFlippersOn
    FlipperLSh.visible = ShadowFlippersOn
    FlipperRSh.visible = ShadowFlippersOn
    exit sub
  end if
  Dim FileObj
  Dim LMConfig2
  dim tempC
  dim tempD
  dim tempFS
  dim tempBS

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & LMEMShadowConfig) then
    Exit Sub
  End if
  Set LMConfig2=FileObj.GetFile(UserDirectory & LMEMShadowConfig)
  Set TextStr2=LMConfig2.OpenAsTextStream(1,0)
  If (TextStr2.AtEndOfStream=True) then
    Exit Sub
  End if
  tempC=TextStr2.ReadLine
  tempD=TextStr2.Readline
  TextStr2.Close
  tempBS=cdbl(tempC)
  tempFS=cdbl(tempD)
  if tempBS=0 then
    EnableBallShadow=false
    BallShadowUpdate.enabled=false
  else
    EnableBallShadow=true
  end if
  if tempFS=0 then
    EnableFlipperShadow=false
    FlipperLSh.visible=false
    FLipperRSh.visible=false
  else
    EnableFlipperShadow=true
  end if
  Set LMConfig2=Nothing
  Set FileObj=Nothing

end sub


Sub DisplayHighScore


end sub

sub InitPauser5_timer
    If B2SOn Then
      'Controller.B2SSetScore 3,HighScore
    End If
    DisplayHighScore
    CreditsReel.SetValue(Credits)
    InitPauser5.enabled=false
end sub



sub BumpersOff
  Bumper1Light.Visible=false
  BumperLight1.visible=False
  Bumper2Light.Visible=false
  BumperLight2.visible=False
end sub

sub BumpersOn
  Bumper1Light.Visible=true
  BumperLight1.visible=True
  Bumper2Light.Visible=true
  BumperLight2.visible=True
end sub

Sub PlasticsOn
  Shadow1.visible=True
  For each obj in Flashers
    obj.state=1
  Next
end sub

Sub PlasticsOff
  Shadow1.visible=False
  Shadow_DropTarget1.visible=False
  Shadow_DropTarget2.visible=False
  Shadow_DropTarget3.visible=False
  Shadow_DropTarget4.visible=False
  Shadow_DropTarget6.visible=False
  Shadow_DropTarget7.visible=False
  Shadow_DropTarget8.visible=False
  Shadow_DropTarget9.visible=False
  Shadow_DropTarget10.visible=False
  For each obj in Flashers
    obj.state=0
  Next
end sub

Sub SetupReplayTables

  Replay1Table(1)=37000
  Replay1Table(2)=39000
  Replay1Table(3)=41000
  Replay1Table(4)=43000
  Replay1Table(5)=45000
  Replay1Table(6)=47000
  Replay1Table(7)=50000
  Replay1Table(8)=51000
  Replay1Table(9)=54000
  Replay1Table(10)=58000
  Replay1Table(11)=61000
  Replay1Table(12)=63000
  Replay1Table(13)=65000
  Replay1Table(14)=67000
  Replay1Table(15)=70000

  Replay2Table(1)=51000
  Replay2Table(2)=53000
  Replay2Table(3)=55000
  Replay2Table(4)=57000
  Replay2Table(5)=59000
  Replay2Table(6)=61000
  Replay2Table(7)=64000
  Replay2Table(8)=65000
  Replay2Table(9)=68000
  Replay2Table(10)=72000
  Replay2Table(11)=75000
  Replay2Table(12)=77000
  Replay2Table(13)=79000
  Replay2Table(14)=81000
  Replay2Table(15)=84000

  Replay3Table(1)=59000
  Replay3Table(2)=61000
  Replay3Table(3)=63000
  Replay3Table(4)=65000
  Replay3Table(5)=67000
  Replay3Table(6)=69000
  Replay3Table(7)=72000
  Replay3Table(8)=73000
  Replay3Table(9)=76000
  Replay3Table(10)=80000
  Replay3Table(11)=83000
  Replay3Table(12)=85000
  Replay3Table(13)=87000
  Replay3Table(14)=89000
  Replay3Table(15)=92000

  ReplayTableMax=15

end sub

Sub RefreshReplayCard
  Dim tempst1
  Dim tempst2

  tempst1=FormatNumber(BallsPerGame,0)
  tempst2=FormatNumber(ReplayLevel,0)

  ReplayCard.image = "SC" + tempst2
  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
end sub


'****************************************
'  SCORE MOTOR
'****************************************

ScoreMotorTimer.Enabled = 1
ScoreMotorTimer.Interval = 135 '135
AddScoreTimer.Enabled = 1
AddScoreTimer.Interval = 135

Dim queuedscore
Dim MotorMode
Dim MotorPosition

Sub SetMotor(y)
  Select Case ScoreMotorAdjustment
    Case 0:
      queuedscore=queuedscore+y
    Case 1:
      If MotorRunning<>1 And InProgress=true then
        queuedscore=queuedscore+y
      end if
    end Select
end sub

Sub SetMotor2(x)
  If MotorRunning<>1 And InProgress=true then
    MotorRunning=1

    Select Case x
      Case 10:
        AddScore(10)
        MotorRunning=0
        BumpersOn

      Case 20:
        MotorMode=10
        MotorPosition=2
        BumpersOff
      Case 30:
        MotorMode=10
        MotorPosition=3
        BumpersOff
      Case 40:
        MotorMode=10
        MotorPosition=4
        BumpersOff
      Case 50:
        MotorMode=10
        MotorPosition=5
        BumpersOff
      Case 100:
        AddScore(100)
        MotorRunning=0
        BumpersOn
      Case 200:
        MotorMode=100
        MotorPosition=2
        BumpersOff
      Case 300:
        MotorMode=100
        MotorPosition=3
        BumpersOff
      Case 400:
        MotorMode=100
        MotorPosition=4
        BumpersOff
      Case 500:
        MotorMode=100
        MotorPosition=5
        BumpersOff
        ToggleAdvance
      Case 1000:
        AddScore(1000)
        MotorRunning=0
        BumpersOn
      Case 2000:
        MotorMode=1000
        MotorPosition=2
        BumpersOff
      Case 3000:
        MotorMode=1000
        MotorPosition=3
        BumpersOff
      Case 4000:
        MotorMode=1000
        MotorPosition=4
        BumpersOff
      Case 5000:
        MotorMode=1000
        MotorPosition=5
        BumpersOff
    End Select
  End If
End Sub

Sub AddScoreTimer_Timer
  Dim tempscore


  If MotorRunning<>1 And InProgress=true then
    if queuedscore>=5000 then
      tempscore=5000
      queuedscore=queuedscore-5000
      SetMotor2(5000)
      exit sub
    end if
    if queuedscore>=4000 then
      tempscore=4000
      queuedscore=queuedscore-4000
      SetMotor2(4000)
      exit sub
    end if

    if queuedscore>=3000 then
      tempscore=3000
      queuedscore=queuedscore-3000
      SetMotor2(3000)
      exit sub
    end if

    if queuedscore>=2000 then
      tempscore=2000
      queuedscore=queuedscore-2000
      SetMotor2(2000)
      exit sub
    end if

    if queuedscore>=1000 then
      tempscore=1000
      queuedscore=queuedscore-1000
      SetMotor2(1000)
      exit sub
    end if

    if queuedscore>=500 then
      tempscore=500
      queuedscore=queuedscore-500
      SetMotor2(500)
      exit sub
    end if
    if queuedscore>=400 then
      tempscore=400
      queuedscore=queuedscore-400
      SetMotor2(400)
      exit sub
    end if
    if queuedscore>=300 then
      tempscore=300
      queuedscore=queuedscore-300
      SetMotor2(300)
      exit sub
    end if
    if queuedscore>=200 then
      tempscore=200
      queuedscore=queuedscore-200
      SetMotor2(200)
      exit sub
    end if
    if queuedscore>=100 then
      tempscore=100
      queuedscore=queuedscore-100
      SetMotor2(100)
      exit sub
    end if

    if queuedscore>=50 then
      tempscore=50
      queuedscore=queuedscore-50
      SetMotor2(50)
      exit sub
    end if
    if queuedscore>=40 then
      tempscore=40
      queuedscore=queuedscore-40
      SetMotor2(40)
      exit sub
    end if
    if queuedscore>=30 then
      tempscore=30
      queuedscore=queuedscore-30
      SetMotor2(30)
      exit sub
    end if
    if queuedscore>=20 then
      tempscore=20
      queuedscore=queuedscore-20
      SetMotor2(20)
      exit sub
    end if
    if queuedscore>=10 then
      tempscore=10
      queuedscore=queuedscore-10
      SetMotor2(10)
      exit sub
    end if


  End If


end Sub

Sub ScoreMotorTimer_Timer
  If MotorPosition > 0 Then
    Select Case MotorPosition
      Case 5,4,3,2:
        If MotorMode=1000 Then
          AddScore(1000)
        end if
        if MotorMode=100 then
          AddScore(100)
        End If
        if MotorMode=10 then
          AddScore(10)
        End if
        MotorPosition=MotorPosition-1
      Case 1:
        If MotorMode=1000 Then
          AddScore(1000)
          DropTargetLight.state=0
        end if
        If MotorMode=100 then
          AddScore(100)
        End If
        if MotorMode=10 then
          AddScore(10)
        End if
        MotorPosition=0:MotorRunning=0:BumpersOn
    End Select
  End If

End Sub


Sub AddScore(x)
  If TableTilted=true then exit sub
  Select Case ScoreAdditionAdjustment
    Case 0:
      AddScore1(x)
    Case 1:
      AddScore2(x)
  end Select

end sub


Sub AddScore1(x)
' debugtext.text=score
  Select Case x
    Case 1:
      PlayChime(10)
      Score(Player)=Score(Player)+1

    Case 10:
      PlayChime(10)
      Score(Player)=Score(Player)+10
'     debugscore=debugscore+10

    Case 100:
      PlayChime(100)
      Score(Player)=Score(Player)+100
'     debugscore=debugscore+100

    Case 1000:
      PlayChime(1000)
      Score(Player)=Score(Player)+1000
'     debugscore=debugscore+1000
  End Select
  PlayerScores(Player-1).AddValue(x)
  If ScoreDisplay(Player)<100000 then
    ScoreDisplay(Player)=Score(Player)
  Else
    Score100K(Player)=Int(Score(Player)/100000)
    ScoreDisplay(Player)=Score(Player)-100000
  End If
  if Score(Player)=>100000 then
    If B2SOn Then
      If Player=1 Then
        Controller.B2SSetScoreRolloverPlayer1 Score100K(Player)
      End If
      If Player=2 Then
        Controller.B2SSetScoreRolloverPlayer2 Score100K(Player)
      End If

      If Player=3 Then
        Controller.B2SSetScoreRolloverPlayer3 Score100K(Player)
      End If

      If Player=4 Then
        Controller.B2SSetScoreRolloverPlayer4 Score100K(Player)
      End If
    End If
  End If
  If B2SOn Then
    Controller.B2SSetScorePlayer Player, ScoreDisplay(Player)
  End If
  If Score(Player)>Replay1 and Replay1Paid(Player)=false then
    Replay1Paid(Player)=True
    AddSpecial
  End If
  If Score(Player)>Replay2 and Replay2Paid(Player)=false then
    Replay2Paid(Player)=True
    AddSpecial
  End If
  If Score(Player)>Replay3 and Replay3Paid(Player)=false then
    Replay3Paid(Player)=True
    AddSpecial
  End If
' ScoreText.text=debugscore
End Sub

Sub AddScore2(x)
  Dim OldScore, NewScore, OldTestScore, NewTestScore
    OldScore = Score(Player)

  Select Case x
        Case 1:
            Score(Player)=Score(Player)+1
    Case 10:
      Score(Player)=Score(Player)+10
    Case 100:
      Score(Player)=Score(Player)+100
    Case 1000:
      Score(Player)=Score(Player)+1000
  End Select
  NewScore = Score(Player)
  if Score(Player)=>100000 then
    If B2SOn Then
      If Player=1 Then
        Controller.B2SSetScoreRolloverPlayer1 1
      End If
      If Player=2 Then
        Controller.B2SSetScoreRolloverPlayer2 1
      End If

      If Player=3 Then
        Controller.B2SSetScoreRolloverPlayer3 1
      End If

      If Player=4 Then
        Controller.B2SSetScoreRolloverPlayer4 1
      End If
    End If
  End If

  OldTestScore = OldScore
  NewTestScore = NewScore
  Do
    if OldTestScore < Replay1 and NewTestScore >= Replay1 then
      AddSpecial()
      NewTestScore = 0
    Elseif OldTestScore < Replay2 and NewTestScore >= Replay2 then
      AddSpecial()
      NewTestScore = 0
    Elseif OldTestScore < Replay3 and NewTestScore >= Replay3 then
      AddSpecial()
      NewTestScore = 0
    End if
    NewTestScore = NewTestScore - 100000
    OldTestScore = OldTestScore - 100000
  Loop While NewTestScore > 0

    OldScore = int(OldScore / 10) ' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
    NewScore = int(NewScore / 10) ' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore&", OldScore Mod 10="&OldScore Mod 10 & ", NewScore % 10="&NewScore Mod 10)

    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(10)

    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(100)

    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(1000)

    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(1000)
    end if

  If B2SOn Then
    Controller.B2SSetScorePlayer Player, Score(Player)
  End If

  if Score(Player) > 0 then
  EMMODE = 1

  UpdateReels  Player,Player ,Score(Player), 0, -1,-1,-1,-1,-1
  EMMODE = 0 ' restore EM mode

  end if

' EMReel1.SetValue Score(Player)
  PlayerScores(Player-1).AddValue(x)
End Sub



Sub PlayChime(x)
  if ChimesOn=0 then
    Select Case x
      Case 10
        If LastChime10=1 Then
          EMSPlayChimeSound (0)
          DOF DOF_CHIME_UNIT_HIGH,DOFPulse
          'PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",128,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          EMSPlayChimeSound (0)
          DOF DOF_CHIME_UNIT_HIGH,DOFPulse
          'PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",128,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          EMSPlayChimeSound (1)
          DOF DOF_CHIME_UNIT_MID,DOFPulse
          'PlaySound SoundFXDOF("SpinACard_1_100_Point_Bell",129,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          EMSPlayChimeSound (1)
          DOF DOF_CHIME_UNIT_MID,DOFPulse
          'PlaySound SoundFXDOF("SpinACard_1_100_Point_Bell",129,DOFPulse,DOFChimes)
          LastChime100=1
        End If

    End Select
  else
    Select Case x
      Case 10
        If LastChime10=1 Then
          EMSPlayChimeSound (0)
          DOF DOF_CHIME_UNIT_HIGH,DOFPulse
          'PlaySound SoundFXDOF("SJ_Chime_10a",128,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          EMSPlayChimeSound (0)
          DOF DOF_CHIME_UNIT_HIGH,DOFPulse
          'PlaySound SoundFXDOF("SJ_Chime_10b",128,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          EMSPlayChimeSound (1)
          DOF DOF_CHIME_UNIT_MID,DOFPulse
          'PlaySound SoundFXDOF("SJ_Chime_100a",129,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          EMSPlayChimeSound (1)
          DOF DOF_CHIME_UNIT_MID,DOFPulse
          'PlaySound SoundFXDOF("SJ_Chime_100b",129,DOFPulse,DOFChimes)
          LastChime100=1
        End If
      Case 1000
        If LastChime1000=1 Then
          EMSPlayChimeSound (2)
          DOF DOF_CHIME_UNIT_LOW,DOFPulse
          'PlaySound SoundFXDOF("SJ_Chime_1000a",130,DOFPulse,DOFChimes)
          LastChime1000=0
        Else
          EMSPlayChimeSound (2)
          DOF DOF_CHIME_UNIT_LOW,DOFPulse
          'PlaySound SoundFXDOF("SJ_Chime_1000b",130,DOFPulse,DOFChimes)
          LastChime1000=1
        End If
    End Select
  end if
End Sub


Sub HideOptions()

end sub


'==================================================================================================
'                        EM Sounds
'==================================================================================================

Const MASTER_VOLUME = 1.0         'value no greater than 1.
Const PAN_TRANSFORM_EXPONENT = 10     'smaller value, less "spread" of audio channels
Const FADE_TRANSFORM_EXPONENT = 10      'smaller value, less "spread" of audio channels
Const MIN_SOUND_PAN_LEFT = -1.0
Const SOUND_PAN_LEFT = -0.9
Const SOUND_PAN_CENTER = 0.0
Const SOUND_PAN_RIGHT = 0.9
Const MAX_SOUND_PAN_RIGHT = 1.0
Const MIN_SOUND_FADE_NEAR_BACKGLASS = -1.0
Const SOUND_FADE_NEAR_BACKGLASS = -0.9
Const SOUND_FADE_CENTER = 0.0
Const SOUND_FADE_NEAR_PLAYER = 0.9
Const MAX_SOUND_FADE_NEAR_PLAYER = 1.0
Const DONT_LOOP_SOUND = 0
Const LOOP_UNTIL_STOPPED = -1
Const NO_RANDOM_PITCH = 0
Const NO_PITCH_CHANGE = 0
Const DONT_USE_EXISTING_SOUND = 0
Const USE_EXISTING_SOUND = 1
Const DONT_RESTART_SOUND = 0
Const RESTART_SOUND = 1

Dim TableWidth, TableHeight
TableWidth = 2024             'reasonable defaults,
TableHeight = 4096              'but please call EMSInit() to get real values assigned

' Pass table object, call before any sound is played
Sub EMSInit (TableObj)
  TableWidth = TableObj.Width
  TableHeight = TableObj.Height
End Sub

'---------------------------Play Positional Sound-------------------------------------------------

' Calculates the volume of the sound based on the ball speed
Function EMSVolumeForBall (Ball)
    EMSVolumeForBall = CSng (EMSBallVelocity (Ball) ^2)
End Function

' Scale pan so there is some bleed-over of left/right channels.
' ex: with headphones you don't want the plunger only in the right ear.
Function EMSTransformPan (Pan)
  If Pan > 0 Then
    EMSTransformPan = CSng (Pan ^ PAN_TRANSFORM_EXPONENT)
  Else
    EMSTransformPan = -CSng (Pan ^ PAN_TRANSFORM_EXPONENT)
  End If
End Function

' Determines pan based on X position on table (-1 is table-left, 0 is table-center, 1 is table-right).
Function EMSPanForTableX (X)
    EMSPanForTableX = CSng ((X * 2 / TableWidth) - 1)
  If EMSPanForTableX < MIN_SOUND_PAN_LEFT Then
    EMSPanForTableX = MIN_SOUND_PAN_LEFT
  ElseIf EMSPanForTableX > MAX_SOUND_PAN_RIGHT Then
    EMSPanForTableX = MAX_SOUND_PAN_RIGHT
  End If
End Function

' Scale fade so there is some bleed-over between front/back channels.
Function EMSTransformFade (Fade)
  If Fade > 0 Then
    EMSTransformFade = CSng (Fade ^ FADE_TRANSFORM_EXPONENT)
  Else
    EMSTransformFade = -CSng (Fade ^ FADE_TRANSFORM_EXPONENT)
  End If
End Function

' Determines fade based on Y position on table (-1 is table-far, 0 is table-center, 1 is table-near).
Function EMSFadeForTableY (Y)
    EMSFadeForTableY = CSng ((Y * 2 / TableHeight) - 1)
  If EMSFadeForTableY < MIN_SOUND_FADE_NEAR_BACKGLASS Then
    EMSFadeForTableY = MIN_SOUND_FADE_NEAR_BACKGLASS
  ElseIf EMSFadeForTableY > MAX_SOUND_FADE_NEAR_PLAYER Then
    EMSFadeForTableY = MAX_SOUND_FADE_NEAR_PLAYER
  End If
End Function

' Plays sound SoundName at level Volume specifying pan of Pan and fade of Fade.
Sub EMSPlaySoundAtVolumePanAndFade (SoundName, Volume, Pan, Fade)
  PlaySound SoundName, DONT_LOOP_SOUND, Volume * MASTER_VOLUME, Pan, NO_RANDOM_PITCH, NO_PITCH_CHANGE, DONT_USE_EXISTING_SOUND, DONT_RESTART_SOUND, Fade
End Sub

' Plays sound SoundName at level Volume specifying pan and fade based on the position of ForObj on the table.
Sub EMSPlaySoundAtVolumeForObject (SoundName, Volume, ForObj)
  Dim Pan
  Pan = EMSTransformPan (EMSPanForTableX (ForObj.X))
  Dim fade
  Fade = EMSTransformFade (EMSFadeForTableY (ForObj.Y))
  EMSPlaySoundAtVolumePanAndFade SoundName, Volume, Pan, Fade
End Sub

' Plays sound SoundName at level Volume specifying pan and fade based on the position of the active ball on the table.
Sub EMSPlaySoundAtVolumeForActiveBall (SoundName, Volume)
  Dim Pan
  Pan = EMSTransformPan (EMSPanForTableX (ActiveBall.X))
  Dim Fade
  Fade = EMSTransformFade (EMSFadeForTableY (ActiveBall.Y))
  EMSPlaySoundAtVolumePanAndFade SoundName, Volume, Pan, Fade
End Sub

' Plays an existing sound SoundName at level Volume specifying pan of Pan and fade of Fade.
Sub EMSPlaySoundExistingAtVolumePanAndFade (SoundName, Volume, Pan, Fade)
  PlaySound soundName, DONT_LOOP_SOUND, Volume * MASTER_VOLUME, Pan, NO_RANDOM_PITCH, NO_PITCH_CHANGE, USE_EXISTING_SOUND, DONT_RESTART_SOUND, Fade
End Sub

' Plays an existing sound SoundName at level Volume specifying pan and fade based on the position of the active ball on the table.
Sub EMSPlaySoundExistingAtVolumeForActiveBall (SoundName, Volume)
  Dim Pan
  Pan = EMSTransformPan (EMSPanForTableX (ActiveBall.X))
  Dim Fade
  Fade = EMSTransformFade (EMSFadeForTableY (ActiveBall.Y))
  EMSPlaySoundExistingAtVolumePanAndFade SoundName, Volume, Pan, Fade
End Sub

' Loops a sound SoundName LoopCount times at level Volume specifying pan and fade based on the position of ForObj on the table.
Sub EMSPlaySoundLoopedAtVolumeForObject (SoundName, LoopCount, Volume, ForObj)
  Dim Pan
  Pan = EMSTransformPan (EMSPanForTableX (ForObj.X))
  Dim Fade
  Fade = EMSTransformFade (EMSFadeForTableY (ForObj.Y))
  PlaySound SoundName, LoopCount, Volume * MASTER_VOLUME, Pan, NO_RANDOM_PITCH, NO_PITCH_CHANGE, DONT_USE_EXISTING_SOUND, RESTART_SOUND, Fade
End Sub

'-------------------------------Ball Rolling------------------------------------------------------

'Calculates the ball speed.
Function EMSBallVelocity (Ball)
    EMSBallVelocity = Int (Sqr ((Ball.VelX ^ 2) + (Ball.VelY ^ 2)))
End Function

Const ROLLING_SOUND_SCALAR = 0.22

' Calculates the roll volume of the sound based on the ball speed.
Function EMSVolumePlayfieldRoll (Ball)
  EMSVolumePlayfieldRoll = ROLLING_SOUND_SCALAR * 0.0005 * CSng (EMSBallVelocity (Ball) ^ 3)
End Function

' Calculates the roll pitch of the sound based on the ball speed.
Function EMSPitchPlayfieldRoll (Ball)
  EMSPitchPlayfieldRoll = EMSBallVelocity (Ball) ^ 2 * 15
End Function

' Calculates the pitch of the sound based on the ball speed.
Function EMSPitch (Ball)
    EMSPitch = EMSBallVelocity (Ball) * 20
End Function

Const NumberOfBalls = 5
Const NumberOfLockedBalls = 0

ReDim Rolling (NumberOfBalls)
InitRolling

Sub InitRolling
  Dim I
  For I = 0 To NumberOfBalls
    Rolling(I) = False
  Next
End Sub

Sub RollingSoundTimer_Timer ()
  Dim BOT, B
  BOT = GetBalls

  ' Stop the sound of deleted balls.
  For B = UBound (BOT) + 1 To NumberOfBalls
    Rolling(B) = False
    StopSound ("BallRoll_" & B)
  Next

  ' Exit the Sub If no balls on the table.
  If UBound (BOT) = -1 Then Exit Sub

  ' Play the rolling sound For each ball.
  For B = 0 To UBound (BOT)
    If EMSBallVelocity (BOT(B)) > 1 AND BOT(B).z < 30 Then
      Rolling(B) = True
      Dim Pan
      Pan = EMSTransformPan (EMSPanForTableX (BOT(B).X))
      Dim Fade
      Fade = EMSTransformFade (EMSFadeForTableY (BOT(B).Y))
      PlaySound ("BallRoll_" & B), LOOP_UNTIL_STOPPED, EMSVolumePlayfieldRoll (BOT(B)) * 1.1 * MASTER_VOLUME, Pan, NO_RANDOM_PITCH, EMSPitchPlayfieldRoll (BOT(B)), USE_EXISTING_SOUND, DONT_RESTART_SOUND, Fade
    Else
      If Rolling(b) = True Then
        StopSound ("BallRoll_" & B)
        Rolling(B) = False
      End If
    End If
  Next
End Sub

Sub OnBallBallCollision (Ball1, Ball2, Velocity)
  PlaySound ("BallCollide"), 0, CSng (Velocity) ^ 2 / 2000, AudioPan (Ball1), 0, Pitch (Ball1), 0, 0, AudioFade (Ball1)
End Sub

'-------------------------------Start Sounds------------------------------------------------------

Const COIN_VOLUME = 1.0           'volume level; range [0, 1]

Sub EMSPlayCoinSound ()
  EMSPlaySoundAtVolumePanAndFade ("Coin_In_" & Int (Rnd * 3) + 1), COIN_VOLUME, SOUND_PAN_CENTER, EMSTransformFade(SOUND_FADE_NEAR_PLAYER)
End Sub

Const STARTUP_VOLUME = 1.0          'volume level; range [0, 1]

Sub EMSPlayStartupSound ()
  EMSPlaySoundAtVolumePanAndFade "startup_norm", STARTUP_VOLUME, SOUND_PAN_CENTER, EMSTransformFade (SOUND_FADE_NEAR_BACKGLASS)
End Sub

Const PLUNGER_PULL_VOLUME = 0.8       'volume level; range [0, 1]

Sub EMSPlayPlungerPullSound (PlungerObj)
  EMSPlaySoundAtVolumeForObject "Plunger_Pull_1", PLUNGER_PULL_VOLUME, PlungerObj
End Sub

Const PLUNGER_RELEASE_VOLUME = 0.8      'volume level; range [0, 1]

Sub EMSPlayPlungerReleaseBallSound (PlungerObj)
  EMSPlaySoundAtVolumeForObject "Plunger_Release_Ball", PLUNGER_RELEASE_VOLUME, PlungerObj
End Sub

Sub EMSPlayPlungerReleaseNoBallSound (PlungerObj)
  EMSPlaySoundAtVolumeForObject "Plunger_Release_No_Ball", PLUNGER_RELEASE_VOLUME, PlungerObj
End Sub

Const START_BALL_VOLUME = 0.4       'volume level; range [0, 1]

Sub EMSPlayStartBallSound (Which)
  Dim SoundName
  If Which = 0 Then
    SoundName = "StartBall1"
  Else
    SoundName = "StartBall2-5"
  End If
  EMSPlaySoundAtVolumePanAndFade SoundName, START_BALL_VOLUME, SOUND_PAN_CENTER, EMSTransformFade (SOUND_FADE_NEAR_PLAYER)
End Sub

Const BALL_RELEASE_VOLUME = 1.0       'volume level; range [0, 1]

Sub EMSPlayBallReleaseSound ()
  EMSPlaySoundAtVolumePanAndFade ("BallRelease" & Int (Rnd * 7) + 1), BALL_RELEASE_VOLUME, EMSTransformPan (SOUND_PAN_RIGHT), EMSTransformFade (SOUND_FADE_NEAR_PLAYER)
End Sub

Const ROTATE_THROUGH_PLAYERS_VOLUME = 0.8 'volume level; range [0, 1]

Sub EMSPlayRotateThroughPlayersSound ()
  EMSPlaySoundAtVolumePanAndFade "RotateThruPlayers", ROTATE_THROUGH_PLAYERS_VOLUME, SOUND_PAN_CENTER, EMSTransformFade (SOUND_FADE_NEAR_PLAYER)
End Sub

'-----------------------------Flipper Sounds------------------------------------------------------

Const FLIPPER_UP_ATTACK_MIN_VOLUME = 0.010  'volume level; range [0, 1]
Const FLIPPER_UP_ATTACK_MAX_VOLUME = 0.635  'volume level; range [0, 1]
Const FLIPPER_UP_VOLUME = 1.0       'volume level; range [0, 1]
Const FLIPPER_BUZZ_VOLUME = 0.05      'volume level; range [0, 1]

Dim FlipperLeftHitParm, FlipperRightHitParm
FlipperLeftHitParm = FLIPPER_UP_VOLUME    'sound helper; not configurable
FlipperRightHitParm = FLIPPER_UP_VOLUME   'sound helper; not configurable

Sub EMSPlayLeftFlipperUpAttackSound (FlipperObj)
  Dim SoundLevel
  SoundLevel = Rnd () * (FLIPPER_UP_ATTACK_MAX_VOLUME - FLIPPER_UP_ATTACK_MIN_VOLUME) + FLIPPER_UP_ATTACK_MIN_VOLUME
  EMSPlaySoundAtVolumeForObject SoundFX ("Flipper_Attack-L01", DOFFlippers), SoundLevel, FlipperObj
End Sub

Sub EMSPlayRightFlipperUpAttackSound (FlipperObj)
  Dim SoundLevel
  SoundLevel = Rnd () * (FLIPPER_UP_ATTACK_MAX_VOLUME - FLIPPER_UP_ATTACK_MIN_VOLUME) + FLIPPER_UP_ATTACK_MIN_VOLUME
  EMSPlaySoundAtVolumeForObject SoundFX ("Flipper_Attack-R01", DOFFlippers), SoundLevel, FlipperObj
End Sub

Sub EMSPlayLeftFlipperUpSound (FlipperObj)
  EMSPlaySoundAtVolumeForObject SoundFX ("Flipper_L0" & Int (Rnd * 9) + 1, DOFFlippers), FlipperLeftHitParm, FlipperObj
End Sub

Sub EMSPlayRightFlipperUpSound (FlipperObj)
  EMSPlaySoundAtVolumeForObject SoundFX("Flipper_R0" & Int (Rnd * 9) + 1, DOFFlippers), FlipperRightHitParm, FlipperObj
End Sub

Const FLIPPER_REFLIP_VOLUME = 0.8     'volume level; range [0, 1]

Sub EMSPlayLeftFlipperReflipSound (FlipperObj)
  EMSPlaySoundAtVolumeForObject SoundFX ("Flipper_ReFlip_L0" & Int (Rnd * 3) + 1, DOFFlippers), FLIPPER_REFLIP_VOLUME, FlipperObj
End Sub

Sub EMSPlayRightFlipperReflipSound (FlipperObj)
  EMSPlaySoundAtVolumeForObject SoundFX ("Flipper_ReFlip_R0" & Int (Rnd * 3) + 1, DOFFlippers), FLIPPER_REFLIP_VOLUME, FlipperObj
End Sub

Const REFLIP_ANGLE = 20

Sub EMSPlayLeftFlipperActivateSound (FlipperObj)
  If FlipperObj.Currentangle < FlipperObj.EndAngle + REFLIP_ANGLE Then
    EMSPlayLeftFlipperReflipSound FlipperObj
  Else
    EMSPlayLeftFlipperUpAttackSound FlipperObj
    EMSPlayLeftFlipperUpSound FlipperObj
  End If
End Sub

Sub EMSPlayRightFlipperActivateSound (FlipperObj)
  If FlipperObj.Currentangle < FlipperObj.EndAngle + REFLIP_ANGLE Then
    EMSPlayRightFlipperReflipSound FlipperObj
  Else
    EMSPlayRightFlipperUpAttackSound FlipperObj
    EMSPlayRightFlipperUpSound FlipperObj
  End If
End Sub

Const FLIPPER_DEACTIVATE_VOLUME = 0.45      'volume level; range [0, 1]

Sub EMSPlayLeftFlipperDeactivateSound (FlipperObj)
  EMSPlaySoundAtVolumeForObject SoundFX ("Flipper_Left_Down_" & Int (Rnd * 7) + 1, DOFFlippers), FLIPPER_DEACTIVATE_VOLUME, FlipperObj
  StopSound "buzzL"
End Sub

Sub EMSPlayRightFlipperDeactivateSound (FlipperObj)
  EMSPlaySoundAtVolumeForObject SoundFX ("Flipper_Right_Down_" & Int (Rnd * 8) + 1, DOFFlippers), FLIPPER_DEACTIVATE_VOLUME, FlipperObj
  StopSound "buzz"
End Sub

Const RUBBER_FLIPPER_SOUND_SCALAR = 0.015 'volume multiplier; must not be zero

Sub EMSPlayLeftFlipperCollideSound (Parm)
  FlipperLeftHitParm = Parm / 10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FLIPPER_UP_VOLUME * FlipperLeftHitParm
  EMSPlaySoundAtVolumeForActiveBall ("Flipper_Rubber_" & Int (Rnd * 7) + 1), Parm  * RUBBER_FLIPPER_SOUND_SCALAR
End Sub

Sub EMSPlayRightFlipperCollideSound (Parm)
  FlipperRightHitParm = Parm / 10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FLIPPER_UP_VOLUME * FlipperRightHitParm
  EMSPlaySoundAtVolumeForActiveBall ("Flipper_Rubber_" & Int (Rnd * 7) + 1), Parm  * RUBBER_FLIPPER_SOUND_SCALAR
End Sub

'-------------------------------Playfield Sounds--------------------------------------------------

Const METAL_IMPACT_SOUND_SCALAR = 0.25

Sub EMSPlayMetalHitSound ()
  EMSPlaySoundAtVolumeForActiveBall ("Metal_Touch_" & Int (Rnd * 13) + 1), EMSVolumeForBall (ActiveBall) * METAL_IMPACT_SOUND_SCALAR
End Sub

Const GATE_HIT_VOLUME = 1         'volume level; range [0, 1]

Sub EMSPlayGateHitSound ()
  EMSPlaySoundAtVolumeForActiveBall ("Gate_1"), GATE_HIT_VOLUME
End Sub

Const DRAIN_VOLUME = 0.8          'volume level; range [0, 1]

Sub EMSPlayDrainSound (DrainObj)
  EMSPlaySoundAtVolumeForObject ("Drain_" & Int (Rnd * 11) + 1), DRAIN_VOLUME, DrainObj
End Sub

Const RUBBER_STRONG_SOUND_SCALAR = 0.011  'volume multiplier; must not be zero
Const RUBBER_WEAK_SOUND_SCALAR = 0.015    'volume multiplier; must not be zero

Sub EMSPlayRubberHitSound ()
  Dim FinalSpeedSquared
  FinalSpeedSquared = (ActiveBall.VelX * ActiveBall.VelX) + (ActiveBall.VelY * ActiveBall.VelY)
  If FinalSpeedSquared > 25 Then      'strong
    Dim SoundIndex
    SoundIndex = Int (Rnd * 10) + 1
    If SoundIndex > 9 Then
      EMSPlaySoundAtVolumeForActiveBall "Rubber_1_Hard", EMSVolumeForBall (ActiveBall) * RUBBER_STRONG_SOUND_SCALAR * 0.6
    Else
      EMSPlaySoundAtVolumeForActiveBall ("Rubber_Strong_" & Int (SoundIndex)), EMSVolumeForBall (ActiveBall) * RUBBER_STRONG_SOUND_SCALAR
    End If
  Else                  'weak
    EMSPlaySoundAtVolumeForActiveBall ("Rubber_" & Int (Rnd * 9) + 1), EMSVolumeForBall (ActiveBall) * RUBBER_WEAK_SOUND_SCALAR
  End If
End Sub

Const SLINGSHOT_VOLUME = 0.95       'volume level; range [0, 1]

Sub EMSPlayLeftSlingshotSound (Sling)
  EMSPlaySoundAtVolumeForObject SoundFX ("Sling_L" & Int(Rnd * 10) + 1, DOFContactors), SLINGSHOT_VOLUME, Sling
End Sub

Sub EMSPlayRightSlingshotSound (Sling)
  EMSPlaySoundAtVolumeForObject SoundFX ("Sling_R" & Int(Rnd * 8) + 1, DOFContactors), SLINGSHOT_VOLUME, Sling
End Sub

Const BUMPER_SOUND_SCALAR = 6.0       'volume multiplier; must not be zero

Sub EMSPlayTopBumperSound (BumperObj)
  EMSPlaySoundAtVolumeForObject SoundFX ("Bumpers_Top_" & Int (Rnd * 5) + 1, DOFContactors), EMSVolumeForBall (ActiveBall) * BUMPER_SOUND_SCALAR, BumperObj
End Sub

Sub EMSPlayMiddleBumperSound (BumperObj)
  EMSPlaySoundAtVolumeForObject SoundFX ("Bumpers_Middle_" & Int (Rnd * 5) + 1, DOFContactors), EMSVolumeForBall (ActiveBall) * BUMPER_SOUND_SCALAR, BumperObj
End Sub

Sub EMSPlayBottomBumperSound (BumperObj)
  EMSPlaySoundAtVolumeForObject SoundFX ("Bumpers_Bottom_" & Int (Rnd * 5) + 1, DOFContactors), EMSVolumeForBall (ActiveBall) * BUMPER_SOUND_SCALAR, BumperObj
End Sub

Const SENSOR_VOLUME = 1.0         'volume level; range [0, 1]

Sub EMSPlaySensorSound ()
  EMSPlaySoundAtVolumeForActiveBall "sensor", SENSOR_VOLUME
End Sub

Const TARGET_SOUND_SCALAR = 0.025     'volume multiplier; must not be zero

Sub EMSPlayTargetHitSound ()
  Dim FinalSpeedSquared
  FinalSpeedSquared = (Activeball.VelX * Activeball.VelX + Activeball.VelY * Activeball.VelY)
  If FinalSpeedSquared > 100 Then
    EMSPlaySoundAtVolumeForActiveBall SoundFX ("Target_Hit_" & Int (Rnd * 4) + 5, DOFTargets), EMSVolumeForBall (ActiveBall) * 0.45 * TARGET_SOUND_SCALAR
  Else
    EMSPlaySoundAtVolumeForActiveBall SoundFX ("Target_Hit_" & Int (Rnd * 4) + 1, DOFTargets), EMSVolumeForBall (ActiveBall) * TARGET_SOUND_SCALAR
  End If
End Sub

'///////////////////////////  DROP TARGETS  ///////////////////////////
Const DropTargetDropSound = "DropTargetDropped" 'Drop Target Drop sound
Const DropTargetResetSound = "DropTarget_Up" 'Drop Target reset sound
Const DropTargetHitSound = "Target_Hit_1"        'Drop Target Hit sound
Const DROP_TARGET_RESET_VOLUME = 1.0    'volume level; range [0, 1]
Const DROP_TARGET_HIT_VOLUME = 1.0
Const DROP_TARGET_DROP_VOLUME = 1.0

'Sub EMSPlayDropTargetHitSound (obj)
' EMSPlaySoundAtVolumePanAndFade "Target_Hit_", DROP_TARGET_HIT_VOLUME, SOUND_PAN_CENTER, EMSTransformFade (SOUND_FADE_NEAR_BACKGLASS)
'End Sub

Sub EMSPlayDropTargetDropSound (obj)
  EMSPlaySoundAtVolumePanAndFade "DropTargetDropped", DROP_TARGET_DROP_VOLUME, SOUND_PAN_CENTER, EMSTransformFade (SOUND_FADE_NEAR_BACKGLASS)
End Sub

Sub EMSPlayDropTargetResetSound(obj)
    EMSPlaySoundAtVolumePanAndFade "DropTarget_Up", DROP_TARGET_RESET_VOLUME, SOUND_PAN_CENTER, EMSTransformFade (SOUND_FADE_NEAR_BACKGLASS)
End Sub

'--------------------------------Other Sounds-----------------------------------------------------

Const CLICK_VOLUME = 1.0          'volume level; range [0, 1]

Sub EMSPlayClickSound ()
  EMSPlaySoundAtVolumePanAndFade "click", CLICK_VOLUME, SOUND_PAN_CENTER, SOUND_FADE_CENTER
End Sub

Const CHIME_VOLUME = 1.0          'volume level; range [0, 1]
Const CHIME_PAN = 0.5

Sub EMSPlayChimeSound (ChimeNum)
  Dim SoundName
  Select Case ChimeNum
    Case 0
      If (Rnd * 2) > 0 Then
        SoundName = "SJ_Chime_10a"
      Else
        SoundName = "SJ_Chime_10b"
      End If
    Case 1
      If (Rnd * 2) > 0 Then
        SoundName = "SJ_Chime_100a"
      Else
        SoundName = "SJ_Chime_100b"
      End If
    Case 2
      If (Rnd * 2) > 0 Then
        SoundName = "SJ_Chime_1000a"
      Else
        SoundName = "SJ_Chime_1000b"
      End If
    Case Else
      Exit Sub
  End Select
  EMSPlaySoundAtVolumePanAndFade SoundName, CHIME_VOLUME, EMSTransformPan (CHIME_PAN), SOUND_FADE_CENTER
End Sub

Const KNOCKER_VOLUME = 1.0          'volume level; range [0, 1]

Sub EMSPlayKnockerSound ()
  EMSPlaySoundAtVolumePanAndFade "Knocker_1", KNOCKER_VOLUME, SOUND_PAN_CENTER, EMSTransformFade (SOUND_FADE_NEAR_BACKGLASS)
End Sub

Const MOTOR_LEER_VOLUME = 1.0       'volume level; range [0, 1]

Sub EMSPlayMotorLeerSound ()
  EMSPlaySoundAtVolumePanAndFade "MotorLeer", MOTOR_LEER_VOLUME, SOUND_PAN_CENTER, EMSTransformFade (SOUND_FADE_NEAR_BACKGLASS)
End Sub

Const NUDGE_VOLUME = 1.0        'volume level; range [0, 1]

Sub EMSPlayNudgeSound ()
  EMSPlaySoundAtVolumePanAndFade "Nudge_" & Int (Rnd * 3) + 1, NUDGE_VOLUME, SOUND_PAN_CENTER, SOUND_FADE_CENTER
End Sub

'==================================================================================================
'                      End EM Sounds
'==================================================================================================



'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5)

Sub BallShadowUpdate_timer ()
  Dim BallsOnTable, B
  BallsOnTable = GetBalls

  ' Hide shadow of deleted balls.
  If UBound(BallsOnTable) < (NumberOfBalls - 1) Then
    For B = (UBound(BallsOnTable) + 1) To (NumberOfBalls - 1)
      BallShadow(B).Visible = 0
    Next
  End If

  ' Exit if no balls on the table.
  If UBound(BallsOnTable) = -1 Then Exit Sub

  ' Render the shadow for each ball.
  For B = 0 To UBound(BallsOnTable)
    If BallsOnTable(B).X < TableWidth / 2 Then
      BallShadow(B).X = ((BallsOnTable(B).X) - (Ballsize / 16) + ((BallsOnTable(B).X - (TableWidth / 2)) / 17)) ' + 13
    Else
      BallShadow(B).X = ((BallsOnTable(B).X) + (Ballsize / 16) + ((BallsOnTable(B).X - (TableWidth / 2)) / 17)) ' - 13
    End If
    BallShadow(B).Y = BallsOnTable(B).Y + 10
    If BallsOnTable(B).Z > 20 Then
      BallShadow(B).Visible = 1
    Else
      BallShadow(B).Visible = 0
    End If
  Next
End Sub

'*****************************************
' Object sounds
'*****************************************
Sub Targets_Hit (idx)
  EMSPlayTargetHitSound
End Sub

Sub Metals_Thin_Hit (idx)
  EMSPlayMetalHitSound
End Sub

Sub Metals_Medium_Hit (idx)
  EMSPlayMetalHitSound
End Sub

Sub Metals2_Hit (idx)
  EMSPlayMetalHitSound
End Sub

Sub Gates_Hit (idx)
  EMSPlayGateHitSound
End Sub

Sub Rubbers_Hit(idx)
  EMSPlayRubberHitSound
End Sub

Sub Posts_Hit(idx)
  EMSPlayPostHitSound
End Sub

Sub LeftFlipper_Collide(parm)
  EMSPlayLeftFlipperCollideSound Parm
End Sub

Sub RightFlipper_Collide(parm)
  EMSPlayRightFlipperCollideSound Parm
End Sub

' ============================================================================================
' GNMOD - Multiple High Score Display and Collection
' ============================================================================================
Dim EnteringInitials    ' Normally zero, set to non-zero to enter initials
EnteringInitials = 0

Dim PlungerPulled
PlungerPulled = 0

Dim SelectedChar      ' character under the "cursor" when entering initials

Dim HSTimerCount      ' Pass counter for HS timer, scores are cycled by the timer
HSTimerCount = 5      ' Timer is initially enabled, it'll wrap from 5 to 1 when it's displayed

Dim InitialString     ' the string holding the player's initials as they're entered

Dim AlphaString       ' A-Z, 0-9, space (_) and backspace (<)
Dim AlphaStringPos      ' pointer to AlphaString, move forward and backward with flipper keys
AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"

Dim HSNewHigh       ' The new score to be recorded

Dim HSScore(5)        ' High Scores read in from config file
Dim HSName(5)       ' High Score Initials read in from config file

' default high scores, remove this when the scores are available from the config file
HSScore(1) = 75000
HSScore(2) = 70000
HSScore(3) = 60000
HSScore(4) = 55000
HSScore(5) = 50000

HSName(1) = "AAA"
HSName(2) = "ZZZ"
HSName(3) = "XXX"
HSName(4) = "ABC"
HSName(5) = "BBB"

Sub HighScoreTimer_Timer
  if EnteringInitials then
    if HSTimerCount = 1 then
      SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
      HSTimerCount = 2
    else
      SetHSLine 3, InitialString
      HSTimerCount = 1
    end if
  elseif InProgress then
    SetHSLine 1, "HIGH SCORE1"
    SetHSLine 2, HSScore(1)
    SetHSLine 3, HSName(1)
    HSTimerCount = 5  ' set so the highest score will show after the game is over
    HighScoreTimer.enabled=false
  elseif CheckAllScores then
    NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)

  else
    ' cycle through high scores
    HighScoreTimer.interval=2000
    HSTimerCount = HSTimerCount + 1
    if HsTimerCount > 5 then
      HSTimerCount = 1
    End If
    SetHSLine 1, "HIGH SCORE"+FormatNumber(HSTimerCount,0)
    SetHSLine 2, HSScore(HSTimerCount)
    SetHSLine 3, HSName(HSTimerCount)
  end if
End Sub

Function GetHSChar(String, Index)
  dim ThisChar
  dim FileName
  ThisChar = Mid(String, Index, 1)
  FileName = "PostIt"
  if ThisChar = " " or ThisChar = "" then
    FileName = FileName & "BL"
  elseif ThisChar = "<" then
    FileName = FileName & "LT"
  elseif ThisChar = "_" then
    FileName = FileName & "SP"
  else
    FileName = FileName & ThisChar
  End If
  GetHSChar = FileName
End Function

Sub SetHsLine(LineNo, String)
  dim Letter
  dim ThisDigit
  dim ThisChar
  dim StrLen
  dim LetterLine
  dim Index
  dim StartHSArray
  dim EndHSArray
  dim LetterName
  dim xfor
  StartHSArray=array(0,1,12,22)
  EndHSArray=array(0,11,21,31)
  StrLen = len(string)
  Index = 1

  for xfor = StartHSArray(LineNo) to EndHSArray(LineNo)
    Eval("HS"&xfor).image = GetHSChar(String, Index)
    Index = Index + 1
  next

End Sub

Sub NewHighScore(NewScore, PlayNum)
  if NewScore > HSScore(5) then
    HighScoreTimer.interval = 500
    HSTimerCount = 1
    AlphaStringPos = 1    ' start with first character "A"
    EnteringInitials = 1  ' intercept the control keys while entering initials
    InitialString = ""    ' initials entered so far, initialize to empty
    SetHSLine 1, "PLAYER "+FormatNumber(PlayNum,0)
    SetHSLine 2, "ENTER NAME"
    SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
    HSNewHigh = NewScore
    For xx=1 to HighScoreReward
      AddSpecial
    next
  End if
  ScoreChecker=ScoreChecker-1
  if ScoreChecker=0 then
    CheckAllScores=0
  end if
End Sub

Sub CollectInitials(keycode)
  If keycode = LeftFlipperKey Then
    ' back up to previous character
    AlphaStringPos = AlphaStringPos - 1
    if AlphaStringPos < 1 then
      AlphaStringPos = len(AlphaString)   ' handle wrap from beginning to end
      if InitialString = "" then
        ' Skip the backspace if there are no characters to backspace over
        AlphaStringPos = AlphaStringPos - 1
      End if
    end if
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTargetDropped"
  elseif keycode = RightFlipperKey Then
    ' advance to next character
    AlphaStringPos = AlphaStringPos + 1
    if AlphaStringPos > len(AlphaString) or (AlphaStringPos = len(AlphaString) and InitialString = "") then
      ' Skip the backspace if there are no characters to backspace over
      AlphaStringPos = 1
    end if
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTargetDropped"
  elseif keycode = StartGameKey or keycode = PlungerKey Then
    SelectedChar = MID(AlphaString, AlphaStringPos, 1)
    if SelectedChar = "_" then
      InitialString = InitialString & " "
      PlaySound("Ding10")
    elseif SelectedChar = "<" then
      InitialString = MID(InitialString, 1, len(InitialString) - 1)
      if len(InitialString) = 0 then
        ' If there are no more characters to back over, don't leave the < displayed
        AlphaStringPos = 1
      end if
      PlaySound("Ding100")
    else
      InitialString = InitialString & SelectedChar
      PlaySound("Ding10")
    end if
    if len(InitialString) < 3 then
      SetHSLine 3, InitialString & SelectedChar
    End If
  End If
  if len(InitialString) = 3 then
    ' save the score
    for i = 5 to 1 step -1
      if i = 1 or (HSNewHigh > HSScore(i) and HSNewHigh <= HSScore(i - 1)) then
        ' Replace the score at this location
        if i < 5 then
' MsgBox("Moving " & i & " to " & (i + 1))
          HSScore(i + 1) = HSScore(i)
          HSName(i + 1) = HSName(i)
        end if
' MsgBox("Saving initials " & InitialString & " to position " & i)
        EnteringInitials = 0
        HSScore(i) = HSNewHigh
        HSName(i) = InitialString
        HSTimerCount = 5
        HighScoreTimer_Timer
        HighScoreTimer.interval = 2000
        PlaySound("Ding1000")
        exit sub
      elseif i < 5 then
        ' move the score in this slot down by 1, it's been exceeded by the new score
' MsgBox("Moving " & i & " to " & (i + 1))
        HSScore(i + 1) = HSScore(i)
        HSName(i + 1) = HSName(i)
      end if
    next
  End If

End Sub
' END GNMOD
' ============================================================================================
' GNMOD - New Options menu
' ============================================================================================
Dim EnteringOptions
Dim CurrentOption
Dim OptionCHS
Dim MaxOption
Dim OptionHighScorePosition
Dim XOpt
Dim StartingArray
Dim EndingArray

StartingArray=Array(0,1,2,30,33,61,89,117,145,173,201,229)
EndingArray=Array(0,1,29,32,60,88,116,144,172,200,228,256)
EnteringOptions = 0
MaxOption = 9
OptionCHS = 0
OptionHighScorePosition = 0
Const OptionLinesToMark="111100011"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4="Wedgehead Tilt Setting"
Const OptionLine5=""
Const OptionLine6=""
Const OptionLine7=""
Const OptionLine8="" 'do not use this line
Const OptionLine9="" 'do not use this line

Sub OperatorMenuTimer_Timer
  EnteringOptions = 1
  OperatorMenuTimer.enabled=false
  ShowOperatorMenu
end sub

sub ShowOperatorMenu
  OperatorMenuBackdrop.image = "OperatorMenu"

  OptionCHS = 0
  CurrentOption = 1
  DisplayAllOptions
  OperatorOption1.image = "BluePlus"
  SetHighScoreOption

End Sub

Sub DisplayAllOptions
  dim linecounter
  dim tempstring
  For linecounter = 1 to MaxOption
    tempstring=Eval("OptionLine"&linecounter)
    Select Case linecounter
      Case 1:
        tempstring=tempstring + FormatNumber(BallsPerGame,0)
        SetOptLine 1,tempstring
      Case 2:
        if Replay3Table(ReplayLevel)=999000 then
          tempstring = tempstring + FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0)
        else
          tempstring = tempstring + FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0) + "/" + FormatNumber(Replay3Table(ReplayLevel),0)
        end if
        SetOptLine 2,tempstring
      Case 3:
        If OptionCHS=0 then
          tempstring = "NO"
        else
          tempstring = "YES"
        end if
        SetOptLine 3,tempstring
      Case 4:
        SetOptLine 4, tempstring
        If TiltEndsGame=0 then
          tempstring="Tilt loses ball in play"
        else
          tempstring="Tilt ends game"
        end if
        SetOptLine 5, tempstring
      Case 5:
        SetOptLine 6, tempstring
        SetOptLine 7, tempstring

      Case 6:
        SetOptLine 8, tempstring
        SetOptLine 9, tempstring

      Case 7:
        SetOptLine 10, tempstring
        SetOptLine 11, tempstring

      Case 8:

      Case 9:


    End Select

  next
end sub

sub MoveArrow
  do
    CurrentOption = CurrentOption + 1
    If CurrentOption>Len(OptionLinesToMark) then
      CurrentOption=1
    end if
  loop until Mid(OptionLinesToMark,CurrentOption,1)="1"
end sub

sub CollectOptions(ByVal keycode)
  if Keycode = LeftFlipperKey then
    PlaySound "DropTargetDropped"
    For XOpt = 1 to MaxOption
      Eval("OperatorOption"&XOpt).image = "PostitBL"
    next
    MoveArrow
    if CurrentOption<8 then
      Eval("OperatorOption"&CurrentOption).image = "BluePlus"
    elseif CurrentOption=8 then
      Eval("OperatorOption"&CurrentOption).image = "GreenCheck"
    else
      Eval("OperatorOption"&CurrentOption).image = "RedX"
    end if

  elseif Keycode = RightFlipperKey then
    PlaySound "DropTargetDropped"
    if CurrentOption = 1 then
      If BallsPerGame = 3 then
        BallsPerGame = 5
      else
        BallsPerGame = 3
      end if
      DisplayAllOptions
    elseif CurrentOption = 2 then
      ReplayLevel=ReplayLevel+1
      If ReplayLevel>ReplayTableMax then
        ReplayLevel=1
      end if
      DisplayAllOptions
    elseif CurrentOption = 3 then
      if OptionCHS = 0 then
        OptionCHS = 1

      else
        OptionCHS = 0

      end if
      DisplayAllOptions
    elseif CurrentOption = 4 then
      if TiltEndsGame=0 then
        TiltEndsGame=1
      else
        TiltEndsGame=0
      end if
      DisplayAllOptions

    elseif CurrentOption = 8 or CurrentOption = 9 then
        if OptionCHS=1 then
          HSScore(1) = 75000
          HSScore(2) = 70000
          HSScore(3) = 60000
          HSScore(4) = 55000
          HSScore(5) = 50000

          HSName(1) = "AAA"
          HSName(2) = "ZZZ"
          HSName(3) = "XXX"
          HSName(4) = "ABC"
          HSName(5) = "BBB"
        end if

        if CurrentOption = 8 then
          savehs
        else
          loadhs
        end if
        OperatorMenuBackdrop.image = "PostitBL"
        For XOpt = 1 to MaxOption
          Eval("OperatorOption"&XOpt).image = "PostitBL"
        next

        For XOpt = 1 to 256
          Eval("Option"&XOpt).image = "PostItBL"
        next
        RefreshReplayCard
        InstructCard.image="IC"+FormatNumber(BallsPerGame,0)
        EnteringOptions = 0

    end if
  end if
End Sub

Sub SetHighScoreOption

End Sub

Function GetOptChar(String, Index)
  dim ThisChar
  dim FileName
  ThisChar = Mid(String, Index, 1)
  FileName = "PostIt"
  if ThisChar = " " or ThisChar = "" then
    FileName = FileName & "BL"
  elseif ThisChar = "<" then
    FileName = FileName & "LT"
  elseif ThisChar = "_" then
    FileName = FileName & "SP"
  elseif ThisChar = "/" then
    FileName = FileName & "SL"
  elseif ThisChar = "," then
    FileName = FileName & "CM"
  else
    FileName = FileName & ThisChar
  End If
  GetOptChar = FileName
End Function

dim LineLengths(22) ' maximum number of lines
Sub SetOptLine(LineNo, String)
  Dim DispLen
    Dim StrLen
  dim xfor
  dim Letter
  dim ThisDigit
  dim ThisChar
  dim LetterLine
  dim Index
  dim LetterName
  StrLen = len(string)
  Index = 1

  StrLen = len(String)
    DispLen = StrLen
    if (DispLen < LineLengths(LineNo)) Then
        DispLen = LineLengths(LineNo)
    end If

  for xfor = StartingArray(LineNo) to StartingArray(LineNo) + DispLen
    Eval("Option"&xfor).image = GetOptChar(string, Index)
    Index = Index + 1
  next
  LineLengths(LineNo) = StrLen

End Sub

'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. an object or point to tell the script where the tip of the flipper is at rest (EndPointLp, EndPointRp)
' 4. and, special scripting
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
' | EOS Torque         | 0.3            | 0.3                   | 0.275                  | 0.275              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

InitPolarity


'*******************************************
' Late 70's to early 80's

Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 2.7
    x.AddPt "Polarity", 2, 0.33, - 2.7
    x.AddPt "Polarity", 3, 0.37, - 2.7
    x.AddPt "Polarity", 4, 0.41, - 2.7
    x.AddPt "Polarity", 5, 0.45, - 2.7
    x.AddPt "Polarity", 6, 0.576, - 2.7
    x.AddPt "Polarity", 7, 0.66, - 1.8
    x.AddPt "Polarity", 8, 0.743, - 0.5
    x.AddPt "Polarity", 9, 0.81, - 0.5
    x.AddPt "Polarity", 10, 0.88, 0

    x.AddPt "Velocity", 0, 0, 1
    x.AddPt "Velocity", 1, 0.16, 1.06
    x.AddPt "Velocity", 2, 0.41, 1.05
    x.AddPt "Velocity", 3, 0.53, 1 '0.982
    x.AddPt "Velocity", 4, 0.702, 0.968
    x.AddPt "Velocity", 5, 0.95,  0.968
    x.AddPt "Velocity", 6, 1.03, 0.945
  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
    LF.SetObjects "LF", LeftFlipper, TriggerLF
    RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub

'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 80
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 3.7
'   x.AddPt "Polarity", 2, 0.33, - 3.7
'   x.AddPt "Polarity", 3, 0.37, - 3.7
'   x.AddPt "Polarity", 4, 0.41, - 3.7
'   x.AddPt "Polarity", 5, 0.45, - 3.7
'   x.AddPt "Polarity", 6, 0.576,- 3.7
'   x.AddPt "Polarity", 7, 0.66, - 2.3
'   x.AddPt "Polarity", 8, 0.743, - 1.5
'   x.AddPt "Polarity", 9, 0.81, - 1
'   x.AddPt "Polarity", 10, 0.88, 0
'
'   x.AddPt "Velocity", 0, 0, 1
'   x.AddPt "Velocity", 1, 0.16, 1.06
'   x.AddPt "Velocity", 2, 0.41, 1.05
'   x.AddPt "Velocity", 3, 0.53, 1 '0.982
'   x.AddPt "Velocity", 4, 0.702, 0.968
'   x.AddPt "Velocity", 5, 0.95,  0.968
'   x.AddPt "Velocity", 6, 1.03, 0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
''*******************************************
''  Late 80's early 90's
'
'Sub InitPolarity()
' dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 5
'   x.AddPt "Polarity", 2, 0.4, - 5
'   x.AddPt "Polarity", 3, 0.6, - 4.5
'   x.AddPt "Polarity", 4, 0.65, - 4.0
'   x.AddPt "Polarity", 5, 0.7, - 3.5
'   x.AddPt "Polarity", 6, 0.75, - 3.0
'   x.AddPt "Polarity", 7, 0.8, - 2.5
'   x.AddPt "Polarity", 8, 0.85, - 2.0
'   x.AddPt "Polarity", 9, 0.9, - 1.5
'   x.AddPt "Polarity", 10, 0.95, - 1.0
'   x.AddPt "Polarity", 11, 1, - 0.5
'   x.AddPt "Polarity", 12, 1.1, 0
'   x.AddPt "Polarity", 13, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 1
'   x.AddPt "Velocity", 1, 0.16, 1.06
'   x.AddPt "Velocity", 2, 0.41, 1.05
'   x.AddPt "Velocity", 3, 0.53, 1 '0.982
'   x.AddPt "Velocity", 4, 0.702, 0.968
'   x.AddPt "Velocity", 5, 0.95,  0.968
'   x.AddPt "Velocity", 6, 1.03,  0.945
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
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
      If Not IsEmpty(balls(x) ) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x) ) Then
        balldata(x).Data = balls(x)
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef
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
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
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
Const EOSTnew = 1 'EM's to late 80's
'Const EOSTnew = 0.8 '90's and later
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
   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
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

Const LiveDistanceMin = 30  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle)  '-1 for Right Flipper
  Dim LiveCatchBounce                                                           'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime
  CatchTime = GameTime - FCount

  If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
    If CatchTime <= LiveCatch * 0.5 Then                        'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    Else
      LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)    'Partial catch when catch happens a bit late
    End If

    If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx = 0
    ball.angmomy = 0
    ball.angmomz = 0
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
'
Sub RDampen_Timer
  Cor.Update
End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************
'******************************************************
'****  VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
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
