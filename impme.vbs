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

' Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
'   Vol = Csng(BallVel(ball) ^2 / 2000)
' End Function
'
' Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
'   Pitch = BallVel(ball) * 20
' End Function
'
' Function BallVel(ball) 'Calculates the ball speed
'   BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
' End Function

' --------------------------------------------------------------------------------

' From Rusty on Facebook
'
' Enter the first 3 sections of the script below as one block
' Do not overwrite the existing sections of regular BallRolling script that preceeds this new code.
' Ensure that < 30 is correct against the the original rolling ball code in the tabble. If not, change 30 to match.
' Ensure that the table's existing Ball Collison script is overwritten by this code (or has otherwise been removed)
' Table1 = TableName of the table being modified. Update as required.
'********************************************************************

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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

' --------------------------------------------------------------------------------

' Optional code

' The following code is OPTIONAL and is a simplified Rubber Bump code for a
' "collection"... Bump sounds are random and rely on the normal sound parameters
' for volume, pitch etc...
'
' ** Note:
' x is a variable that can be use to adjust volume. A value of 1 should be used here for no volume adjustment.
' Decimals (.01 - .99) will reduce volume.
' Whole numbers (1-10) will increase volume. (I'm not sure where this maxes out, I generally go no further than 3, but I have seen some tables use up to 10 or more)
'
' As with all the sound code contained in this sheet. A VPX version call has been included for 10.3 backwards compatibility.

' *******************************************************************************************************
' Random Rubber Bumps - Use instead of the existing velocity based random rubber script
' x is a volume variable. Decimals decrease volume. Whole Numbers increase volume
' ******************************************************************************************************

Sub Rubbers_Hit(idx)
  PlaySoundAtBallVol "fx_rubber_hit_" & Int(Rnd*3)+1, x
End Sub

' **************************************************************************************
' Random Flipper Hits - Use instead of the existing velocity based random flipper script
' x is a volume variable. Decimals decrease volume. Whole Numbers increase volume
' **************************************************************************************

Sub LeftFlipper_Collide(parm)
  PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, x
End Sub

Sub RightFlipper_Collide(parm)
  PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, x
End Sub

' The following is an different way of adding raised ramp bumps. It works best
' when used in conjunction with the Raised Ramp Rolling Ball Script, but may also
' be used to compliment regular ramp sounds. Simply add switches to ramps in key
' bend locations and call them as part of the collection.

'**************************************************************************************************************
' Random Ramp Bumps  - Best used to compliment the Raised Ramp RollingBall Script
' No switches are required. Make ramps have "Hit Event" and create a "Collection" for each of the following.
' Adjust Vol (eg .5) & Pitch: 		RandomBump3 .5, Pitch(ActiveBall)+5
' Ramp physics also effect bump sounds. Friction will change ball speed, so sounds will react different.
' Hit Threshold will will effect how much force is required to create a bump. Low hit threshold + more bumps.
' Other adjustments avialable in the script as noted.
' Note that specific sounds are required in sound manager as noted. Export/Import these from an existing table with this code.
'**************************************************************************************************************


Dim NextOrbitHit:NextOrbitHit = 0

Sub WireRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump3 .5, Pitch(ActiveBall)+5
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .2 + (Rnd * .2)
  end if
End Sub

Sub PlasticRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump 2, Pitch(ActiveBall)
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .1 + (Rnd * .2)
  end if
End Sub


Sub MetalGuideBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump2 2, Pitch(ActiveBall)
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .2 + (Rnd * .2)
  end if
End Sub

Sub MetalWallBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump 2, 20000 'Increased pitch to simulate metal wall
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .2 + (Rnd * .2)
  end if
End Sub


' Requires rampbump1 to 7 in Sound Manager

Sub RandomBump(voladj, freq)
  dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Requires metalguidebump1 to 2 in Sound Manager

Sub RandomBump2(voladj, freq)
  dim BumpSnd:BumpSnd= "metalguidebump" & CStr(Int(Rnd*2)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Requires WireRampBump1 to 5 in Sound Manager

Sub RandomBump3(voladj, freq)
  dim BumpSnd:BumpSnd= "WireRampBump" & CStr(Int(Rnd*5)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub



' Stop Bump Sounds - Place triggers for these at the end of the ramps, to ensure that no bump sounds are played after the ball leaves the ramp.

Sub BumpSTOP1_Hit ()
  dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
  NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOP2_Hit ()
  dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
  NextOrbitHit = Timer + 1
End Sub

