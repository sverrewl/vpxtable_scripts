'                      :#@@@@@-                                     :+:
'     *@@@@@@@@@*   .%@@@@@@@@@@@- ..@@@@@. .@@@@@  #@@@@@@@+ ::@@@@@@*@@@@@@@@@@@@@@#        =#@@@@@@@@@@=   =@@@@@@@@%.  +*%%@+   .-++    -@@@@@
'   #@@@@@@@@@@@@@=.@@@@@+--=@@@@@=-:@@@@@.-:@@@@@.-%@@@@@@@@+--@@@@@@#@@@@@@@@@@@@@@-     ---@@@@@@@@@@@@@#-@@@@@@@@@@@@%-*@@@@@-.:#@@@%.-:%@@@@-=@@@@@@@+ -:@@@@@@+
'  =@@@@%.---.-:.::#@@@+------.@@@@.:@@@@@:-:@@@@@:-+@@@@@@@@@.-@@@@@@-----.@@@@@:::       ---@@@@#=..*@@@@@*@@@-----:@@@@@:@@@@@@::@@@@@--:@@@@@:+@@@@@@@@=-:@@@@@@+
'  %@@@+-::--:---:#@%@@::   :--*@@@#.@@@@@:-:@@@@%--=@@@@@@@@@@:@@@@@@:  --:@@@@@=   ::::: .-:@@@@.----.@@@@%#@:..:----@@@@*.@@@@@+:@@@@@@-@@@@@+-=@@@@@@@@@.:@@@@@@+
' .@@@@-        @@@@ @@-     ..@@@@-=@@@@#--:@@@@#--:@@@@=#@@@@@@@@@@@*  ---@@@@@= -+@@@@@@=--+@@@# .---#@@@@=@     ---%@@@*-=@@@@@@@@-:@@@@@@@@:--@@@@@@@@@@+@@@@@@=
' -:@@@@:    -:@@@@%+@@@@=  .%@@@@#::@@@@@%%@@@@@:--.@@@@+-+@@@@@@@@@@=  ---#@@@@:.- =+++++ --.@@@@   -.@@@@%%@%    -:%@@@@--:@@@@@@@#--#@@@@@@.--=@@@@:@@@@@@@@@@@@=
' --*@@@@@%*%@@@@@@::%@@@@@@@@@@@=.--:@@@@@@@@@@----.@@@@+--+@@@@@@@@@-  ---+@@@@: :------. ---@@@@*+#@@@@@@-@@@@@@@@@@@@@ ---.@@@@@@:--:@@@@@%.---@@@@--%@@@@@@@@@@=
' :==.@@@@@@@@@@@: -==:+@@@@@@%.-: -==-.%@@@@@- .-==.@@@%----.:-------   ---+@@@@:           ---@@@@@@@@@@%.:.*@@@@@@@@@-  .---+@@@#: ---..... .--:@@@@---#@@@@@@@@@=
'  .=+== -%@@+ :-   .==========-    .========:  .--.====   .------::     -: -===             -=-@@@@@*: :--:-----:   :--    :------   .:-----  .-----:  -----------
'     =+++++++:         .--:.                                                                ..-=====--.     :=======:
'
'                                    Count-Down (Gottlieb 1979)
'                               https://www.ipdb.org/machine.cgi?id=573

' v1 - version by JPSalas 2024, version 5.5.0
' ...
' v4  - FrankEnstein - Blender noodling until pretty comes out.
' v5  - mcarter78 - Replace all posts/rubbers, add nFozzy/Fleep, set up vpm mapped lights, update code to latest VPW standards, add physical kicker/trough
' v6  - FrankEnstein -  added inserts edge details, better light core, added bump to rail walls, fixed the apron model, bumper correct size and added protector under, bunch of other visual tuneups
' v7  - FrankEnstein - 2k batch, animate flippers, fixed positions and proportions that were off,
' v8  - mcarter78 - implement drop targets (broken), animate bumper/rollovers, light controlled drop resets, tune rubbers physics, add chimes
' v9  - FrankEnstein - Added raytraced shadows, added option for reflections control (off, static, lighting), added options for apron light mod, removed ghost reflection objects
' v10 - FrankEnstein - Added custom instruction card in F12 menu, back to 4k bake. rendered multiple sling positions. separated rails and lockbar and added options to disable.
' v11 - FrankEnstein - Fixed drop targets (thanks to Apophis's wisdom)
' v12 - bthlonewolf - Added animations for VUK and slings
' v13 - FrankEnstein - Actually fixed all the drop targed, moved the sling timers to Timers layer, added shadow flashers over dropped targets to darken below playfield
' v14 - bthlonewolf - Added InitRubberState to set default animation state for moveables.Added animations for rubbers
' v15 - DGrimmReaper - VR Hybrid Support Added (Thanks to Ext2k, Apophis and RajoJoey for VR Assets)
' v16 - FrankEnstein - Rebalanced apron light and ball brightness, updated playfield mesh for less saucer wiggle, added TOC, added clean or rough reflections option
' v17 - FrankEnstein - added post adjustement options in F12 (conservative and liberal, per manuals wording), added outlane bottom edge pins.
' v18 - FrankEnstein - Matched flippers gap to real table, updated UV maps on areas with issues, fixed apron light on launch not picking up correct value, Bumper skirt is movable.
' v19 - bthlonewolf - Added handlers for each sling difficulty (sw53j_RL / sw53j_RC for right, sw53i_LL / sw53i_LC); fixed right sling visibility; reduced zCol_RubberBands_Middles elasticity to .95
' v20 - FrankEnstein - Rebalanced flippers power based on reference videos. Fixed the rougue ball reflection from the apron light,
' v21 - FrankEnstein - Fixed and improved flipper routines.
' v22 - FrankEnstein - Fixed GI lights height/RT shadows. Fixed lower and mid rubbers that weren't firing correctly. Adjusted the rubber knob to match reference videos. Added skirt tilt animation. Dialed ring reponse. Updated desktop BG
' v23 - FrankEnstein - Added dynamic drop target shadows.
' v24 - DGrimmReaper - Swapped VR Backglass assets with Hauntfreaks
' v25 - FrankEnstein - Playfield image overhaul with lots of the questionable areas patched with better source, added variation to insert bulbs positions, better rubber knob, flippers have better UVs
' v26 - FrankEnstein - fixed blue insert ring. This is looking a lot like an RC1 now
' v27 - mcarter78 - Narrow scoring area in rubbers to near leaf switches, add rubbers for liberal/conservative difficulty
' v28 - mcarter78 - Remove non-scoring rubbers from Rubbers collection to disable sounds, animate rubbers on non-scoring hits
' v29 - FrankEnstein - fixed some pf mapping. lower sets of drop target shadows played a bit stronger
' 1.0.0 Release
' 1.0.1 - FrankEnstein - Adjusted flipper distance to correct/exact measurements. Added F12 option to set flipper distance back to casual mode.
' 1.0.2 - FrankEnstein - Fixed visual inconsistency with the top rail wire and replaced it with metal panel, repainted the lower playfield from better sources. Added metal prims to drop ball from getting stuck on edge of flipper
' 1.1.1 - FrankEnstein - Fixed apron light affecting the apron
' 1.1.2 - bthlonewolf - Added intial Post-it high score.
' 1.1.3 - bthlonewolf - Added tweak menu for Post-it
' 1.1.4 - bthlonewolf - Fixed post-it for VR
' 1.1.5 - bthlonewolf - Fixed HS bug for right-flipper+start (game restart) - cancels post-it if high score
' 1.1.6 - FrankEnstein - Fixed the missing plunger texture and rebalanced postit lighting

' TABLE OF CONTENTS
' -----------------
'   ZVAR: Constants and Global Variables
'   ZTIM: Timers
'   ZINI: Table Initialization and Exiting
'   ZOPT: User Options
'   ZMAT: General Math Functions
'   ZBBR: Ball Brightness
'   ZRBR: Room Brightness
' ZKEY: Key Press Handling
' ZSOL: Solenoids & Flashers
' ZDRN: Drain, Trough, and Ball Release
' ZFLP: Flippers
' ZANI: Misc Animations
' ZSLG: Slingshot Animations
' ZBMP: Bumpers
' ZSWI: Switches
' ZFRG: Frogs
' ZVUK: VUKs and Kickers
' ZSHA: Ambient ball shadows
' ZBRL: Ball Rolling and Drop Sounds
'   ZRRL: Ramp Rolling SFX
'   ZFLE: Fleep Mechanical Sounds
' ZNFF: Flipper Corrections and Tricks
'   ZDMP: Rubber Dampeners
' ZRST: Stand Up Targets
'   ZBOU: VPW TargetBouncer
'   ZVBG: VR Backglass
'   ZVRR: VR Room

Option Explicit
Randomize
SetLocale 1033      'Forces VBS to use english to stop crashes.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Bumper_Ring: BP_Bumper_Ring=Array(BM_Bumper_Ring, LM_GI_RT_GI_10_Bumper_Ring, LM_GI_RT_GI_5_Bumper_Ring, LM_GI_RT_GI_7_Bumper_Ring, LM_GI_RT_GI_8_Bumper_Ring, LM_GI_Bumper_Ring)
Dim BP_Bumper_Socket: BP_Bumper_Socket=Array(BM_Bumper_Socket, LM_GI_RT_GI_10_Bumper_Socket, LM_GI_RT_GI_11_Bumper_Socket, LM_GI_RT_GI_16_Bumper_Socket, LM_GI_RT_GI_7_Bumper_Socket, LM_GI_RT_GI_8_Bumper_Socket, LM_GI_RT_GI_9_Bumper_Socket, LM_GI_Bumper_Socket)
Dim BP_Gate2_Wire: BP_Gate2_Wire=Array(BM_Gate2_Wire)
Dim BP_GateFlap: BP_GateFlap=Array(BM_GateFlap, LM_GI_RT_GI_16_GateFlap, LM_GI_GateFlap)
Dim BP_LFlip: BP_LFlip=Array(BM_LFlip, LM_GI_RT_GI_1_LFlip, LM_GI_RT_GI_5_LFlip, LM_GI_RT_GI_2_LFlip, LM_GI_RT_GI_99_LFlip, LM_GI_LFlip, LM_LI_L23_LFlip)
Dim BP_LFlip1: BP_LFlip1=Array(BM_LFlip1, LM_GI_RT_GI_1_LFlip1, LM_GI_RT_GI_10_LFlip1, LM_GI_RT_GI_5_LFlip1, LM_GI_RT_GI_4_LFlip1, LM_GI_LFlip1)
Dim BP_LFlip1U: BP_LFlip1U=Array(BM_LFlip1U, LM_GI_RT_GI_1_LFlip1U, LM_GI_RT_GI_10_LFlip1U, LM_GI_RT_GI_5_LFlip1U, LM_GI_RT_GI_4_LFlip1U, LM_GI_LFlip1U)
Dim BP_LFlipU: BP_LFlipU=Array(BM_LFlipU, LM_GI_RT_GI_1_LFlipU, LM_GI_RT_GI_5_LFlipU, LM_GI_RT_GI_2_LFlipU, LM_GI_RT_GI_99_LFlipU, LM_GI_LFlipU, LM_LI_L23_LFlipU, LM_LI_L24_LFlipU)
Dim BP_Overlay: BP_Overlay=Array(BM_Overlay, LM_GI_RT_GI_1_Overlay, LM_GI_RT_GI_10_Overlay, LM_GI_RT_GI_11_Overlay, LM_GI_RT_GI_5_Overlay, LM_GI_RT_GI_16_Overlay, LM_GI_RT_GI_2_Overlay, LM_GI_RT_GI_4_Overlay, LM_GI_RT_GI_7_Overlay, LM_GI_RT_GI_8_Overlay, LM_GI_RT_GI_9_Overlay, LM_GI_Overlay, LM_LI_L35_Overlay, LM_LI_L36_Overlay, LM_LI_L8_Overlay)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_GI_RT_GI_1_Parts, LM_GI_RT_GI_10_Parts, LM_GI_RT_GI_11_Parts, LM_GI_RT_GI_5_Parts, LM_GI_RT_GI_16_Parts, LM_GI_RT_GI_2_Parts, LM_GI_RT_GI_4_Parts, LM_GI_RT_GI_7_Parts, LM_GI_RT_GI_8_Parts, LM_GI_RT_GI_9_Parts, LM_GI_RT_GI_99_Parts, LM_GI_Parts, LM_LI_L13_Parts, LM_LI_L14_Parts, LM_LI_L15_Parts, LM_LI_L16_Parts, LM_LI_L35_Parts, LM_LI_L36_Parts, LM_LI_L5_Parts, LM_LI_L6_Parts, LM_LI_L8_Parts)
Dim BP_PegPlasticSlingL1C: BP_PegPlasticSlingL1C=Array(BM_PegPlasticSlingL1C, LM_GI_RT_GI_1_PegPlasticSlingL1, LM_GI_RT_GI_5_PegPlasticSlingL1, LM_GI_PegPlasticSlingL1C, LM_LI_L5_PegPlasticSlingL1C, LM_LI_L6_PegPlasticSlingL1C)
Dim BP_PegPlasticSlingLC: BP_PegPlasticSlingLC=Array(BM_PegPlasticSlingLC, LM_GI_RT_GI_1_PegPlasticSlingLC, LM_GI_RT_GI_5_PegPlasticSlingLC, LM_GI_PegPlasticSlingLC)
Dim BP_PegPlasticSlingR1C: BP_PegPlasticSlingR1C=Array(BM_PegPlasticSlingR1C, LM_GI_RT_GI_2_PegPlasticSlingR1, LM_GI_RT_GI_4_PegPlasticSlingR1, LM_GI_PegPlasticSlingR1C, LM_LI_L35_PegPlasticSlingR1C, LM_LI_L36_PegPlasticSlingR1C)
Dim BP_PegPlasticSlingRC: BP_PegPlasticSlingRC=Array(BM_PegPlasticSlingRC, LM_GI_RT_GI_2_PegPlasticSlingRC, LM_GI_PegPlasticSlingRC, LM_LI_L36_PegPlasticSlingRC)
Dim BP_PkickarmR: BP_PkickarmR=Array(BM_PkickarmR, LM_GI_PkickarmR)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GI_RT_GI_1_Playfield, LM_GI_RT_GI_10_Playfield, LM_GI_RT_GI_11_Playfield, LM_GI_RT_GI_5_Playfield, LM_GI_RT_GI_16_Playfield, LM_GI_RT_GI_2_Playfield, LM_GI_RT_GI_4_Playfield, LM_GI_RT_GI_7_Playfield, LM_GI_RT_GI_8_Playfield, LM_GI_RT_GI_9_Playfield, LM_GI_RT_GI_99_Playfield, LM_GI_Playfield, LM_LI_L10_Playfield, LM_LI_L11_Playfield, LM_LI_L12_Playfield, LM_LI_L13_Playfield, LM_LI_L14_Playfield, LM_LI_L15_Playfield, LM_LI_L16_Playfield, LM_LI_L19_Playfield, LM_LI_L20_Playfield, LM_LI_L21_Playfield, LM_LI_L22_Playfield, LM_LI_L23_Playfield, LM_LI_L24_Playfield, LM_LI_L25_Playfield, LM_LI_L26_Playfield, LM_LI_L27_Playfield, LM_LI_L28_Playfield, LM_LI_L29_Playfield, LM_LI_L30_Playfield, LM_LI_L31_Playfield, LM_LI_L32_Playfield, LM_LI_L33_Playfield, LM_LI_L35_Playfield, LM_LI_L36_Playfield, LM_LI_L4_Playfield, LM_LI_L5_Playfield, LM_LI_L6_Playfield, LM_LI_L7_Playfield, LM_LI_L8_Playfield, LM_LI_L9_Playfield)
Dim BP_PostPlasticCapL1C: BP_PostPlasticCapL1C=Array(BM_PostPlasticCapL1C, LM_GI_RT_GI_1_PostPlasticCapL1C, LM_GI_RT_GI_5_PostPlasticCapL1C, LM_GI_PostPlasticCapL1C)
Dim BP_PostPlasticCapLC: BP_PostPlasticCapLC=Array(BM_PostPlasticCapLC, LM_GI_RT_GI_1_PostPlasticCapLC, LM_GI_RT_GI_5_PostPlasticCapLC, LM_GI_PostPlasticCapLC)
Dim BP_PostPlasticCapR1C: BP_PostPlasticCapR1C=Array(BM_PostPlasticCapR1C, LM_GI_RT_GI_2_PostPlasticCapR1C, LM_GI_PostPlasticCapR1C)
Dim BP_PostPlasticCapRC: BP_PostPlasticCapRC=Array(BM_PostPlasticCapRC, LM_GI_RT_GI_2_PostPlasticCapRC, LM_GI_RT_GI_4_PostPlasticCapRC, LM_GI_PostPlasticCapRC)
Dim BP_RFlip: BP_RFlip=Array(BM_RFlip, LM_GI_RT_GI_1_RFlip, LM_GI_RT_GI_2_RFlip, LM_GI_RT_GI_4_RFlip, LM_GI_RT_GI_99_RFlip, LM_GI_RFlip, LM_LI_L23_RFlip)
Dim BP_RFlip1: BP_RFlip1=Array(BM_RFlip1, LM_GI_RT_GI_11_RFlip1, LM_GI_RT_GI_5_RFlip1, LM_GI_RT_GI_4_RFlip1, LM_GI_RFlip1)
Dim BP_RFlip1U: BP_RFlip1U=Array(BM_RFlip1U, LM_GI_RT_GI_11_RFlip1U, LM_GI_RT_GI_5_RFlip1U, LM_GI_RT_GI_4_RFlip1U, LM_GI_RFlip1U)
Dim BP_RFlipU: BP_RFlipU=Array(BM_RFlipU, LM_GI_RT_GI_1_RFlipU, LM_GI_RT_GI_2_RFlipU, LM_GI_RT_GI_4_RFlipU, LM_GI_RT_GI_99_RFlipU, LM_GI_RFlipU, LM_LI_L23_RFlipU, LM_LI_L25_RFlipU)
Dim BP_RailL: BP_RailL=Array(BM_RailL, LM_GI_RT_GI_99_RailL)
Dim BP_RailR: BP_RailR=Array(BM_RailR, LM_GI_RT_GI_99_RailR)
Dim BP_Rubber10: BP_Rubber10=Array(BM_Rubber10, LM_GI_RT_GI_1_Rubber10, LM_GI_RT_GI_5_Rubber10, LM_GI_RT_GI_2_Rubber10, LM_GI_RT_GI_4_Rubber10, LM_GI_Rubber10, LM_LI_L5_Rubber10, LM_LI_L6_Rubber10)
Dim BP_Rubber10a: BP_Rubber10a=Array(BM_Rubber10a, LM_GI_RT_GI_1_Rubber10a, LM_GI_RT_GI_5_Rubber10a, LM_GI_RT_GI_2_Rubber10a, LM_GI_RT_GI_4_Rubber10a, LM_GI_Rubber10a, LM_LI_L5_Rubber10a, LM_LI_L6_Rubber10a)
Dim BP_Rubber10b: BP_Rubber10b=Array(BM_Rubber10b, LM_GI_RT_GI_1_Rubber10b, LM_GI_RT_GI_5_Rubber10b, LM_GI_RT_GI_2_Rubber10b, LM_GI_RT_GI_4_Rubber10b, LM_GI_Rubber10b, LM_LI_L5_Rubber10b, LM_LI_L6_Rubber10b)
Dim BP_Rubber2: BP_Rubber2=Array(BM_Rubber2, LM_GI_RT_GI_10_Rubber2, LM_GI_RT_GI_11_Rubber2, LM_GI_RT_GI_5_Rubber2, LM_GI_RT_GI_16_Rubber2, LM_GI_RT_GI_7_Rubber2, LM_GI_RT_GI_8_Rubber2, LM_GI_RT_GI_9_Rubber2, LM_GI_Rubber2)
Dim BP_Rubber2a: BP_Rubber2a=Array(BM_Rubber2a, LM_GI_RT_GI_10_Rubber2a, LM_GI_RT_GI_11_Rubber2a, LM_GI_RT_GI_5_Rubber2a, LM_GI_RT_GI_16_Rubber2a, LM_GI_RT_GI_7_Rubber2a, LM_GI_RT_GI_8_Rubber2a, LM_GI_RT_GI_9_Rubber2a, LM_GI_Rubber2a)
Dim BP_Rubber2b: BP_Rubber2b=Array(BM_Rubber2b, LM_GI_RT_GI_10_Rubber2b, LM_GI_RT_GI_11_Rubber2b, LM_GI_RT_GI_5_Rubber2b, LM_GI_RT_GI_16_Rubber2b, LM_GI_RT_GI_7_Rubber2b, LM_GI_RT_GI_8_Rubber2b, LM_GI_RT_GI_9_Rubber2b, LM_GI_Rubber2b)
Dim BP_Rubber2c: BP_Rubber2c=Array(BM_Rubber2c, LM_GI_RT_GI_10_Rubber2c, LM_GI_RT_GI_11_Rubber2c, LM_GI_RT_GI_5_Rubber2c, LM_GI_RT_GI_16_Rubber2c, LM_GI_RT_GI_7_Rubber2c, LM_GI_RT_GI_8_Rubber2c, LM_GI_RT_GI_9_Rubber2c, LM_GI_Rubber2c)
Dim BP_Rubber2d: BP_Rubber2d=Array(BM_Rubber2d, LM_GI_RT_GI_10_Rubber2d, LM_GI_RT_GI_11_Rubber2d, LM_GI_RT_GI_5_Rubber2d, LM_GI_RT_GI_16_Rubber2d, LM_GI_RT_GI_7_Rubber2d, LM_GI_RT_GI_8_Rubber2d, LM_GI_RT_GI_9_Rubber2d, LM_GI_Rubber2d)
Dim BP_Rubber4: BP_Rubber4=Array(BM_Rubber4, LM_GI_RT_GI_10_Rubber4, LM_GI_RT_GI_11_Rubber4, LM_GI_RT_GI_5_Rubber4, LM_GI_RT_GI_16_Rubber4, LM_GI_RT_GI_4_Rubber4, LM_GI_RT_GI_7_Rubber4, LM_GI_RT_GI_8_Rubber4, LM_GI_RT_GI_9_Rubber4, LM_GI_Rubber4)
Dim BP_Rubber4a: BP_Rubber4a=Array(BM_Rubber4a, LM_GI_RT_GI_10_Rubber4a, LM_GI_RT_GI_11_Rubber4a, LM_GI_RT_GI_5_Rubber4a, LM_GI_RT_GI_16_Rubber4a, LM_GI_RT_GI_4_Rubber4a, LM_GI_RT_GI_7_Rubber4a, LM_GI_RT_GI_8_Rubber4a, LM_GI_RT_GI_9_Rubber4a, LM_GI_Rubber4a)
Dim BP_Rubber4b: BP_Rubber4b=Array(BM_Rubber4b, LM_GI_RT_GI_10_Rubber4b, LM_GI_RT_GI_11_Rubber4b, LM_GI_RT_GI_5_Rubber4b, LM_GI_RT_GI_16_Rubber4b, LM_GI_RT_GI_4_Rubber4b, LM_GI_RT_GI_7_Rubber4b, LM_GI_RT_GI_8_Rubber4b, LM_GI_RT_GI_9_Rubber4b, LM_GI_Rubber4b)
Dim BP_Rubber4c: BP_Rubber4c=Array(BM_Rubber4c, LM_GI_RT_GI_10_Rubber4c, LM_GI_RT_GI_11_Rubber4c, LM_GI_RT_GI_5_Rubber4c, LM_GI_RT_GI_16_Rubber4c, LM_GI_RT_GI_4_Rubber4c, LM_GI_RT_GI_7_Rubber4c, LM_GI_RT_GI_8_Rubber4c, LM_GI_RT_GI_9_Rubber4c, LM_GI_Rubber4c)
Dim BP_Rubber4d: BP_Rubber4d=Array(BM_Rubber4d, LM_GI_RT_GI_10_Rubber4d, LM_GI_RT_GI_11_Rubber4d, LM_GI_RT_GI_5_Rubber4d, LM_GI_RT_GI_16_Rubber4d, LM_GI_RT_GI_4_Rubber4d, LM_GI_RT_GI_7_Rubber4d, LM_GI_RT_GI_8_Rubber4d, LM_GI_RT_GI_9_Rubber4d, LM_GI_Rubber4d)
Dim BP_Rubber7: BP_Rubber7=Array(BM_Rubber7, LM_GI_RT_GI_11_Rubber7, LM_GI_RT_GI_4_Rubber7, LM_GI_RT_GI_9_Rubber7, LM_GI_Rubber7, LM_LI_L14_Rubber7)
Dim BP_Rubber7a: BP_Rubber7a=Array(BM_Rubber7a, LM_GI_RT_GI_11_Rubber7a, LM_GI_RT_GI_4_Rubber7a, LM_GI_RT_GI_9_Rubber7a, LM_GI_Rubber7a, LM_LI_L14_Rubber7a)
Dim BP_Rubber7b: BP_Rubber7b=Array(BM_Rubber7b, LM_GI_RT_GI_11_Rubber7b, LM_GI_RT_GI_4_Rubber7b, LM_GI_RT_GI_9_Rubber7b, LM_GI_Rubber7b, LM_LI_L14_Rubber7b)
Dim BP_Rubber8: BP_Rubber8=Array(BM_Rubber8, LM_GI_RT_GI_10_Rubber8, LM_GI_RT_GI_5_Rubber8, LM_GI_RT_GI_8_Rubber8, LM_GI_Rubber8)
Dim BP_Rubber8a: BP_Rubber8a=Array(BM_Rubber8a, LM_GI_RT_GI_10_Rubber8a, LM_GI_RT_GI_5_Rubber8a, LM_GI_RT_GI_8_Rubber8a, LM_GI_Rubber8a)
Dim BP_Rubber8b: BP_Rubber8b=Array(BM_Rubber8b, LM_GI_RT_GI_10_Rubber8b, LM_GI_RT_GI_5_Rubber8b, LM_GI_RT_GI_8_Rubber8b, LM_GI_Rubber8b)
Dim BP_Rubber9: BP_Rubber9=Array(BM_Rubber9, LM_GI_RT_GI_1_Rubber9, LM_GI_RT_GI_10_Rubber9, LM_GI_RT_GI_11_Rubber9, LM_GI_RT_GI_5_Rubber9, LM_GI_RT_GI_2_Rubber9, LM_GI_RT_GI_4_Rubber9, LM_GI_Rubber9, LM_LI_L35_Rubber9, LM_LI_L36_Rubber9)
Dim BP_Rubber9a: BP_Rubber9a=Array(BM_Rubber9a, LM_GI_RT_GI_1_Rubber9a, LM_GI_RT_GI_10_Rubber9a, LM_GI_RT_GI_11_Rubber9a, LM_GI_RT_GI_5_Rubber9a, LM_GI_RT_GI_2_Rubber9a, LM_GI_RT_GI_4_Rubber9a, LM_GI_Rubber9a, LM_LI_L35_Rubber9a, LM_LI_L36_Rubber9a)
Dim BP_Rubber9b: BP_Rubber9b=Array(BM_Rubber9b, LM_GI_RT_GI_1_Rubber9b, LM_GI_RT_GI_10_Rubber9b, LM_GI_RT_GI_11_Rubber9b, LM_GI_RT_GI_5_Rubber9b, LM_GI_RT_GI_2_Rubber9b, LM_GI_RT_GI_4_Rubber9b, LM_GI_Rubber9b, LM_LI_L35_Rubber9b, LM_LI_L36_Rubber9b)
Dim BP_RubberLSling: BP_RubberLSling=Array(BM_RubberLSling, LM_GI_RT_GI_1_RubberLSling, LM_GI_RT_GI_11_RubberLSling, LM_GI_RubberLSling, LM_LI_L6_RubberLSling)
Dim BP_RubberLSlinga: BP_RubberLSlinga=Array(BM_RubberLSlinga, LM_GI_RT_GI_1_RubberLSlinga, LM_GI_RT_GI_11_RubberLSlinga, LM_GI_RubberLSlinga, LM_LI_L6_RubberLSlinga)
Dim BP_RubberLSlingb: BP_RubberLSlingb=Array(BM_RubberLSlingb, LM_GI_RT_GI_1_RubberLSlingb, LM_GI_RT_GI_11_RubberLSlingb, LM_GI_RubberLSlingb, LM_LI_L6_RubberLSlingb)
Dim BP_RubberRSling: BP_RubberRSling=Array(BM_RubberRSling, LM_GI_RT_GI_2_RubberRSling, LM_GI_RT_GI_4_RubberRSling, LM_GI_RubberRSling, LM_LI_L36_RubberRSling)
Dim BP_RubberRSlinga: BP_RubberRSlinga=Array(BM_RubberRSlinga, LM_GI_RT_GI_2_RubberRSlinga, LM_GI_RT_GI_4_RubberRSlinga, LM_GI_RubberRSlinga, LM_LI_L36_RubberRSlinga)
Dim BP_RubberRSlingb: BP_RubberRSlingb=Array(BM_RubberRSlingb, LM_GI_RT_GI_2_RubberRSlingb, LM_GI_RT_GI_4_RubberRSlingb, LM_GI_RubberRSlingb, LM_LI_L36_RubberRSlingb)
Dim BP_lockdownbar: BP_lockdownbar=Array(BM_lockdownbar, LM_GI_RT_GI_99_lockdownbar)
Dim BP_sw10: BP_sw10=Array(BM_sw10, LM_GI_RT_GI_1_sw10, LM_GI_RT_GI_5_sw10, LM_GI_sw10)
Dim BP_sw11: BP_sw11=Array(BM_sw11, LM_GI_RT_GI_1_sw11, LM_GI_RT_GI_5_sw11, LM_GI_sw11)
Dim BP_sw13: BP_sw13=Array(BM_sw13, LM_GI_RT_GI_2_sw13, LM_GI_sw13)
Dim BP_sw14: BP_sw14=Array(BM_sw14, LM_GI_RT_GI_2_sw14, LM_GI_sw14)
Dim BP_sw20: BP_sw20=Array(BM_sw20, LM_GI_RT_GI_1_sw20, LM_GI_RT_GI_5_sw20, LM_GI_sw20)
Dim BP_sw21: BP_sw21=Array(BM_sw21, LM_GI_RT_GI_1_sw21, LM_GI_RT_GI_5_sw21, LM_GI_RT_GI_4_sw21, LM_GI_sw21)
Dim BP_sw23: BP_sw23=Array(BM_sw23, LM_GI_RT_GI_1_sw23, LM_GI_RT_GI_5_sw23, LM_GI_RT_GI_4_sw23, LM_GI_sw23)
Dim BP_sw24: BP_sw24=Array(BM_sw24, LM_GI_RT_GI_1_sw24, LM_GI_RT_GI_5_sw24, LM_GI_RT_GI_2_sw24, LM_GI_RT_GI_4_sw24, LM_GI_sw24)
Dim BP_sw30: BP_sw30=Array(BM_sw30, LM_GI_RT_GI_10_sw30, LM_GI_RT_GI_7_sw30, LM_GI_RT_GI_8_sw30, LM_GI_sw30)
Dim BP_sw31: BP_sw31=Array(BM_sw31, LM_GI_RT_GI_10_sw31, LM_GI_RT_GI_7_sw31, LM_GI_RT_GI_8_sw31, LM_GI_sw31)
Dim BP_sw33: BP_sw33=Array(BM_sw33, LM_GI_RT_GI_10_sw33, LM_GI_RT_GI_7_sw33, LM_GI_RT_GI_8_sw33, LM_GI_sw33)
Dim BP_sw34: BP_sw34=Array(BM_sw34, LM_GI_RT_GI_10_sw34, LM_GI_RT_GI_7_sw34, LM_GI_RT_GI_8_sw34, LM_GI_sw34)
Dim BP_sw40: BP_sw40=Array(BM_sw40, LM_GI_RT_GI_1_sw40, LM_GI_RT_GI_10_sw40, LM_GI_RT_GI_5_sw40, LM_GI_RT_GI_4_sw40, LM_GI_sw40, LM_LI_L12_sw40)
Dim BP_sw43: BP_sw43=Array(BM_sw43, LM_GI_RT_GI_10_sw43, LM_GI_sw43)
Dim BP_sw44: BP_sw44=Array(BM_sw44, LM_GI_RT_GI_11_sw44)
Dim BP_sw50: BP_sw50=Array(BM_sw50, LM_GI_RT_GI_1_sw50, LM_GI_RT_GI_10_sw50, LM_GI_RT_GI_11_sw50, LM_GI_RT_GI_5_sw50, LM_GI_RT_GI_2_sw50, LM_GI_RT_GI_4_sw50, LM_GI_sw50, LM_LI_L11_sw50)
Dim BP_sw60: BP_sw60=Array(BM_sw60, LM_GI_RT_GI_10_sw60, LM_GI_RT_GI_11_sw60, LM_GI_RT_GI_5_sw60, LM_GI_RT_GI_16_sw60, LM_GI_RT_GI_4_sw60, LM_GI_RT_GI_9_sw60, LM_GI_sw60)
Dim BP_sw61: BP_sw61=Array(BM_sw61, LM_GI_RT_GI_10_sw61, LM_GI_RT_GI_11_sw61, LM_GI_RT_GI_5_sw61, LM_GI_RT_GI_16_sw61, LM_GI_RT_GI_4_sw61, LM_GI_RT_GI_9_sw61, LM_GI_sw61, LM_LI_L8_sw61)
Dim BP_sw63: BP_sw63=Array(BM_sw63, LM_GI_RT_GI_10_sw63, LM_GI_RT_GI_11_sw63, LM_GI_RT_GI_16_sw63, LM_GI_RT_GI_4_sw63, LM_GI_RT_GI_9_sw63, LM_GI_sw63, LM_LI_L8_sw63)
Dim BP_sw64: BP_sw64=Array(BM_sw64, LM_GI_RT_GI_10_sw64, LM_GI_RT_GI_11_sw64, LM_GI_RT_GI_16_sw64, LM_GI_RT_GI_4_sw64, LM_GI_RT_GI_9_sw64, LM_GI_sw64)
Dim BP_sw70: BP_sw70=Array(BM_sw70, LM_GI_RT_GI_4_sw70, LM_GI_sw70)
Dim BP_sw71: BP_sw71=Array(BM_sw71, LM_GI_RT_GI_4_sw71, LM_GI_sw71)
Dim BP_sw73: BP_sw73=Array(BM_sw73, LM_GI_RT_GI_4_sw73, LM_GI_sw73)
Dim BP_sw74: BP_sw74=Array(BM_sw74, LM_GI_RT_GI_4_sw74, LM_GI_sw74, LM_LI_L35_sw74)
' Arrays per lighting scenario
Dim BL_GI: BL_GI=Array(LM_GI_Bumper_Ring, LM_GI_Bumper_Socket, LM_GI_GateFlap, LM_GI_LFlip, LM_GI_LFlip1, LM_GI_LFlip1U, LM_GI_LFlipU, LM_GI_Overlay, LM_GI_Parts, LM_GI_PegPlasticSlingL1C, LM_GI_PegPlasticSlingLC, LM_GI_PegPlasticSlingR1C, LM_GI_PegPlasticSlingRC, LM_GI_PkickarmR, LM_GI_Playfield, LM_GI_PostPlasticCapL1C, LM_GI_PostPlasticCapLC, LM_GI_PostPlasticCapR1C, LM_GI_PostPlasticCapRC, LM_GI_RFlip, LM_GI_RFlip1, LM_GI_RFlip1U, LM_GI_RFlipU, LM_GI_Rubber10, LM_GI_Rubber10a, LM_GI_Rubber10b, LM_GI_Rubber2, LM_GI_Rubber2a, LM_GI_Rubber2b, LM_GI_Rubber2c, LM_GI_Rubber2d, LM_GI_Rubber4, LM_GI_Rubber4a, LM_GI_Rubber4b, LM_GI_Rubber4c, LM_GI_Rubber4d, LM_GI_Rubber7, LM_GI_Rubber7a, LM_GI_Rubber7b, LM_GI_Rubber8, LM_GI_Rubber8a, LM_GI_Rubber8b, LM_GI_Rubber9, LM_GI_Rubber9a, LM_GI_Rubber9b, LM_GI_RubberLSling, LM_GI_RubberLSlinga, LM_GI_RubberLSlingb, LM_GI_RubberRSling, LM_GI_RubberRSlinga, LM_GI_RubberRSlingb, LM_GI_sw10, LM_GI_sw11, LM_GI_sw13, LM_GI_sw14, LM_GI_sw20, LM_GI_sw21, LM_GI_sw23, LM_GI_sw24, _
  LM_GI_sw30, LM_GI_sw31, LM_GI_sw33, LM_GI_sw34, LM_GI_sw40, LM_GI_sw43, LM_GI_sw50, LM_GI_sw60, LM_GI_sw61, LM_GI_sw63, LM_GI_sw64, LM_GI_sw70, LM_GI_sw71, LM_GI_sw73, LM_GI_sw74)
Dim BL_GI_RT_GI_1: BL_GI_RT_GI_1=Array(LM_GI_RT_GI_1_LFlip, LM_GI_RT_GI_1_LFlip1, LM_GI_RT_GI_1_LFlip1U, LM_GI_RT_GI_1_LFlipU, LM_GI_RT_GI_1_Overlay, LM_GI_RT_GI_1_Parts, LM_GI_RT_GI_1_PegPlasticSlingL1, LM_GI_RT_GI_1_PegPlasticSlingLC, LM_GI_RT_GI_1_Playfield, LM_GI_RT_GI_1_PostPlasticCapL1C, LM_GI_RT_GI_1_PostPlasticCapLC, LM_GI_RT_GI_1_RFlip, LM_GI_RT_GI_1_RFlipU, LM_GI_RT_GI_1_Rubber10, LM_GI_RT_GI_1_Rubber10a, LM_GI_RT_GI_1_Rubber10b, LM_GI_RT_GI_1_Rubber9, LM_GI_RT_GI_1_Rubber9a, LM_GI_RT_GI_1_Rubber9b, LM_GI_RT_GI_1_RubberLSling, LM_GI_RT_GI_1_RubberLSlinga, LM_GI_RT_GI_1_RubberLSlingb, LM_GI_RT_GI_1_sw10, LM_GI_RT_GI_1_sw11, LM_GI_RT_GI_1_sw20, LM_GI_RT_GI_1_sw21, LM_GI_RT_GI_1_sw23, LM_GI_RT_GI_1_sw24, LM_GI_RT_GI_1_sw40, LM_GI_RT_GI_1_sw50)
Dim BL_GI_RT_GI_10: BL_GI_RT_GI_10=Array(LM_GI_RT_GI_10_Bumper_Ring, LM_GI_RT_GI_10_Bumper_Socket, LM_GI_RT_GI_10_LFlip1, LM_GI_RT_GI_10_LFlip1U, LM_GI_RT_GI_10_Overlay, LM_GI_RT_GI_10_Parts, LM_GI_RT_GI_10_Playfield, LM_GI_RT_GI_10_Rubber2, LM_GI_RT_GI_10_Rubber2a, LM_GI_RT_GI_10_Rubber2b, LM_GI_RT_GI_10_Rubber2c, LM_GI_RT_GI_10_Rubber2d, LM_GI_RT_GI_10_Rubber4, LM_GI_RT_GI_10_Rubber4a, LM_GI_RT_GI_10_Rubber4b, LM_GI_RT_GI_10_Rubber4c, LM_GI_RT_GI_10_Rubber4d, LM_GI_RT_GI_10_Rubber8, LM_GI_RT_GI_10_Rubber8a, LM_GI_RT_GI_10_Rubber8b, LM_GI_RT_GI_10_Rubber9, LM_GI_RT_GI_10_Rubber9a, LM_GI_RT_GI_10_Rubber9b, LM_GI_RT_GI_10_sw30, LM_GI_RT_GI_10_sw31, LM_GI_RT_GI_10_sw33, LM_GI_RT_GI_10_sw34, LM_GI_RT_GI_10_sw40, LM_GI_RT_GI_10_sw43, LM_GI_RT_GI_10_sw50, LM_GI_RT_GI_10_sw60, LM_GI_RT_GI_10_sw61, LM_GI_RT_GI_10_sw63, LM_GI_RT_GI_10_sw64)
Dim BL_GI_RT_GI_11: BL_GI_RT_GI_11=Array(LM_GI_RT_GI_11_Bumper_Socket, LM_GI_RT_GI_11_Overlay, LM_GI_RT_GI_11_Parts, LM_GI_RT_GI_11_Playfield, LM_GI_RT_GI_11_RFlip1, LM_GI_RT_GI_11_RFlip1U, LM_GI_RT_GI_11_Rubber2, LM_GI_RT_GI_11_Rubber2a, LM_GI_RT_GI_11_Rubber2b, LM_GI_RT_GI_11_Rubber2c, LM_GI_RT_GI_11_Rubber2d, LM_GI_RT_GI_11_Rubber4, LM_GI_RT_GI_11_Rubber4a, LM_GI_RT_GI_11_Rubber4b, LM_GI_RT_GI_11_Rubber4c, LM_GI_RT_GI_11_Rubber4d, LM_GI_RT_GI_11_Rubber7, LM_GI_RT_GI_11_Rubber7a, LM_GI_RT_GI_11_Rubber7b, LM_GI_RT_GI_11_Rubber9, LM_GI_RT_GI_11_Rubber9a, LM_GI_RT_GI_11_Rubber9b, LM_GI_RT_GI_11_RubberLSling, LM_GI_RT_GI_11_RubberLSlinga, LM_GI_RT_GI_11_RubberLSlingb, LM_GI_RT_GI_11_sw44, LM_GI_RT_GI_11_sw50, LM_GI_RT_GI_11_sw60, LM_GI_RT_GI_11_sw61, LM_GI_RT_GI_11_sw63, LM_GI_RT_GI_11_sw64)
Dim BL_GI_RT_GI_16: BL_GI_RT_GI_16=Array(LM_GI_RT_GI_16_Bumper_Socket, LM_GI_RT_GI_16_GateFlap, LM_GI_RT_GI_16_Overlay, LM_GI_RT_GI_16_Parts, LM_GI_RT_GI_16_Playfield, LM_GI_RT_GI_16_Rubber2, LM_GI_RT_GI_16_Rubber2a, LM_GI_RT_GI_16_Rubber2b, LM_GI_RT_GI_16_Rubber2c, LM_GI_RT_GI_16_Rubber2d, LM_GI_RT_GI_16_Rubber4, LM_GI_RT_GI_16_Rubber4a, LM_GI_RT_GI_16_Rubber4b, LM_GI_RT_GI_16_Rubber4c, LM_GI_RT_GI_16_Rubber4d, LM_GI_RT_GI_16_sw60, LM_GI_RT_GI_16_sw61, LM_GI_RT_GI_16_sw63, LM_GI_RT_GI_16_sw64)
Dim BL_GI_RT_GI_2: BL_GI_RT_GI_2=Array(LM_GI_RT_GI_2_LFlip, LM_GI_RT_GI_2_LFlipU, LM_GI_RT_GI_2_Overlay, LM_GI_RT_GI_2_Parts, LM_GI_RT_GI_2_PegPlasticSlingR1, LM_GI_RT_GI_2_PegPlasticSlingRC, LM_GI_RT_GI_2_Playfield, LM_GI_RT_GI_2_PostPlasticCapR1C, LM_GI_RT_GI_2_PostPlasticCapRC, LM_GI_RT_GI_2_RFlip, LM_GI_RT_GI_2_RFlipU, LM_GI_RT_GI_2_Rubber10, LM_GI_RT_GI_2_Rubber10a, LM_GI_RT_GI_2_Rubber10b, LM_GI_RT_GI_2_Rubber9, LM_GI_RT_GI_2_Rubber9a, LM_GI_RT_GI_2_Rubber9b, LM_GI_RT_GI_2_RubberRSling, LM_GI_RT_GI_2_RubberRSlinga, LM_GI_RT_GI_2_RubberRSlingb, LM_GI_RT_GI_2_sw13, LM_GI_RT_GI_2_sw14, LM_GI_RT_GI_2_sw24, LM_GI_RT_GI_2_sw50)
Dim BL_GI_RT_GI_4: BL_GI_RT_GI_4=Array(LM_GI_RT_GI_4_LFlip1, LM_GI_RT_GI_4_LFlip1U, LM_GI_RT_GI_4_Overlay, LM_GI_RT_GI_4_Parts, LM_GI_RT_GI_4_PegPlasticSlingR1, LM_GI_RT_GI_4_Playfield, LM_GI_RT_GI_4_PostPlasticCapRC, LM_GI_RT_GI_4_RFlip, LM_GI_RT_GI_4_RFlip1, LM_GI_RT_GI_4_RFlip1U, LM_GI_RT_GI_4_RFlipU, LM_GI_RT_GI_4_Rubber10, LM_GI_RT_GI_4_Rubber10a, LM_GI_RT_GI_4_Rubber10b, LM_GI_RT_GI_4_Rubber4, LM_GI_RT_GI_4_Rubber4a, LM_GI_RT_GI_4_Rubber4b, LM_GI_RT_GI_4_Rubber4c, LM_GI_RT_GI_4_Rubber4d, LM_GI_RT_GI_4_Rubber7, LM_GI_RT_GI_4_Rubber7a, LM_GI_RT_GI_4_Rubber7b, LM_GI_RT_GI_4_Rubber9, LM_GI_RT_GI_4_Rubber9a, LM_GI_RT_GI_4_Rubber9b, LM_GI_RT_GI_4_RubberRSling, LM_GI_RT_GI_4_RubberRSlinga, LM_GI_RT_GI_4_RubberRSlingb, LM_GI_RT_GI_4_sw21, LM_GI_RT_GI_4_sw23, LM_GI_RT_GI_4_sw24, LM_GI_RT_GI_4_sw40, LM_GI_RT_GI_4_sw50, LM_GI_RT_GI_4_sw60, LM_GI_RT_GI_4_sw61, LM_GI_RT_GI_4_sw63, LM_GI_RT_GI_4_sw64, LM_GI_RT_GI_4_sw70, LM_GI_RT_GI_4_sw71, LM_GI_RT_GI_4_sw73, LM_GI_RT_GI_4_sw74)
Dim BL_GI_RT_GI_5: BL_GI_RT_GI_5=Array(LM_GI_RT_GI_5_Bumper_Ring, LM_GI_RT_GI_5_LFlip, LM_GI_RT_GI_5_LFlip1, LM_GI_RT_GI_5_LFlip1U, LM_GI_RT_GI_5_LFlipU, LM_GI_RT_GI_5_Overlay, LM_GI_RT_GI_5_Parts, LM_GI_RT_GI_5_PegPlasticSlingL1, LM_GI_RT_GI_5_PegPlasticSlingLC, LM_GI_RT_GI_5_Playfield, LM_GI_RT_GI_5_PostPlasticCapL1C, LM_GI_RT_GI_5_PostPlasticCapLC, LM_GI_RT_GI_5_RFlip1, LM_GI_RT_GI_5_RFlip1U, LM_GI_RT_GI_5_Rubber10, LM_GI_RT_GI_5_Rubber10a, LM_GI_RT_GI_5_Rubber10b, LM_GI_RT_GI_5_Rubber2, LM_GI_RT_GI_5_Rubber2a, LM_GI_RT_GI_5_Rubber2b, LM_GI_RT_GI_5_Rubber2c, LM_GI_RT_GI_5_Rubber2d, LM_GI_RT_GI_5_Rubber4, LM_GI_RT_GI_5_Rubber4a, LM_GI_RT_GI_5_Rubber4b, LM_GI_RT_GI_5_Rubber4c, LM_GI_RT_GI_5_Rubber4d, LM_GI_RT_GI_5_Rubber8, LM_GI_RT_GI_5_Rubber8a, LM_GI_RT_GI_5_Rubber8b, LM_GI_RT_GI_5_Rubber9, LM_GI_RT_GI_5_Rubber9a, LM_GI_RT_GI_5_Rubber9b, LM_GI_RT_GI_5_sw10, LM_GI_RT_GI_5_sw11, LM_GI_RT_GI_5_sw20, LM_GI_RT_GI_5_sw21, LM_GI_RT_GI_5_sw23, LM_GI_RT_GI_5_sw24, LM_GI_RT_GI_5_sw40, LM_GI_RT_GI_5_sw50, _
  LM_GI_RT_GI_5_sw60, LM_GI_RT_GI_5_sw61)
Dim BL_GI_RT_GI_7: BL_GI_RT_GI_7=Array(LM_GI_RT_GI_7_Bumper_Ring, LM_GI_RT_GI_7_Bumper_Socket, LM_GI_RT_GI_7_Overlay, LM_GI_RT_GI_7_Parts, LM_GI_RT_GI_7_Playfield, LM_GI_RT_GI_7_Rubber2, LM_GI_RT_GI_7_Rubber2a, LM_GI_RT_GI_7_Rubber2b, LM_GI_RT_GI_7_Rubber2c, LM_GI_RT_GI_7_Rubber2d, LM_GI_RT_GI_7_Rubber4, LM_GI_RT_GI_7_Rubber4a, LM_GI_RT_GI_7_Rubber4b, LM_GI_RT_GI_7_Rubber4c, LM_GI_RT_GI_7_Rubber4d, LM_GI_RT_GI_7_sw30, LM_GI_RT_GI_7_sw31, LM_GI_RT_GI_7_sw33, LM_GI_RT_GI_7_sw34)
Dim BL_GI_RT_GI_8: BL_GI_RT_GI_8=Array(LM_GI_RT_GI_8_Bumper_Ring, LM_GI_RT_GI_8_Bumper_Socket, LM_GI_RT_GI_8_Overlay, LM_GI_RT_GI_8_Parts, LM_GI_RT_GI_8_Playfield, LM_GI_RT_GI_8_Rubber2, LM_GI_RT_GI_8_Rubber2a, LM_GI_RT_GI_8_Rubber2b, LM_GI_RT_GI_8_Rubber2c, LM_GI_RT_GI_8_Rubber2d, LM_GI_RT_GI_8_Rubber4, LM_GI_RT_GI_8_Rubber4a, LM_GI_RT_GI_8_Rubber4b, LM_GI_RT_GI_8_Rubber4c, LM_GI_RT_GI_8_Rubber4d, LM_GI_RT_GI_8_Rubber8, LM_GI_RT_GI_8_Rubber8a, LM_GI_RT_GI_8_Rubber8b, LM_GI_RT_GI_8_sw30, LM_GI_RT_GI_8_sw31, LM_GI_RT_GI_8_sw33, LM_GI_RT_GI_8_sw34)
Dim BL_GI_RT_GI_9: BL_GI_RT_GI_9=Array(LM_GI_RT_GI_9_Bumper_Socket, LM_GI_RT_GI_9_Overlay, LM_GI_RT_GI_9_Parts, LM_GI_RT_GI_9_Playfield, LM_GI_RT_GI_9_Rubber2, LM_GI_RT_GI_9_Rubber2a, LM_GI_RT_GI_9_Rubber2b, LM_GI_RT_GI_9_Rubber2c, LM_GI_RT_GI_9_Rubber2d, LM_GI_RT_GI_9_Rubber4, LM_GI_RT_GI_9_Rubber4a, LM_GI_RT_GI_9_Rubber4b, LM_GI_RT_GI_9_Rubber4c, LM_GI_RT_GI_9_Rubber4d, LM_GI_RT_GI_9_Rubber7, LM_GI_RT_GI_9_Rubber7a, LM_GI_RT_GI_9_Rubber7b, LM_GI_RT_GI_9_sw60, LM_GI_RT_GI_9_sw61, LM_GI_RT_GI_9_sw63, LM_GI_RT_GI_9_sw64)
Dim BL_GI_RT_GI_99: BL_GI_RT_GI_99=Array(LM_GI_RT_GI_99_LFlip, LM_GI_RT_GI_99_LFlipU, LM_GI_RT_GI_99_Parts, LM_GI_RT_GI_99_Playfield, LM_GI_RT_GI_99_RFlip, LM_GI_RT_GI_99_RFlipU, LM_GI_RT_GI_99_RailL, LM_GI_RT_GI_99_RailR, LM_GI_RT_GI_99_lockdownbar)
Dim BL_LI_L10: BL_LI_L10=Array(LM_LI_L10_Playfield)
Dim BL_LI_L11: BL_LI_L11=Array(LM_LI_L11_Playfield, LM_LI_L11_sw50)
Dim BL_LI_L12: BL_LI_L12=Array(LM_LI_L12_Playfield, LM_LI_L12_sw40)
Dim BL_LI_L13: BL_LI_L13=Array(LM_LI_L13_Parts, LM_LI_L13_Playfield)
Dim BL_LI_L14: BL_LI_L14=Array(LM_LI_L14_Parts, LM_LI_L14_Playfield, LM_LI_L14_Rubber7, LM_LI_L14_Rubber7a, LM_LI_L14_Rubber7b)
Dim BL_LI_L15: BL_LI_L15=Array(LM_LI_L15_Parts, LM_LI_L15_Playfield)
Dim BL_LI_L16: BL_LI_L16=Array(LM_LI_L16_Parts, LM_LI_L16_Playfield)
Dim BL_LI_L19: BL_LI_L19=Array(LM_LI_L19_Playfield)
Dim BL_LI_L20: BL_LI_L20=Array(LM_LI_L20_Playfield)
Dim BL_LI_L21: BL_LI_L21=Array(LM_LI_L21_Playfield)
Dim BL_LI_L22: BL_LI_L22=Array(LM_LI_L22_Playfield)
Dim BL_LI_L23: BL_LI_L23=Array(LM_LI_L23_LFlip, LM_LI_L23_LFlipU, LM_LI_L23_Playfield, LM_LI_L23_RFlip, LM_LI_L23_RFlipU)
Dim BL_LI_L24: BL_LI_L24=Array(LM_LI_L24_LFlipU, LM_LI_L24_Playfield)
Dim BL_LI_L25: BL_LI_L25=Array(LM_LI_L25_Playfield, LM_LI_L25_RFlipU)
Dim BL_LI_L26: BL_LI_L26=Array(LM_LI_L26_Playfield)
Dim BL_LI_L27: BL_LI_L27=Array(LM_LI_L27_Playfield)
Dim BL_LI_L28: BL_LI_L28=Array(LM_LI_L28_Playfield)
Dim BL_LI_L29: BL_LI_L29=Array(LM_LI_L29_Playfield)
Dim BL_LI_L30: BL_LI_L30=Array(LM_LI_L30_Playfield)
Dim BL_LI_L31: BL_LI_L31=Array(LM_LI_L31_Playfield)
Dim BL_LI_L32: BL_LI_L32=Array(LM_LI_L32_Playfield)
Dim BL_LI_L33: BL_LI_L33=Array(LM_LI_L33_Playfield)
Dim BL_LI_L35: BL_LI_L35=Array(LM_LI_L35_Overlay, LM_LI_L35_Parts, LM_LI_L35_PegPlasticSlingR1C, LM_LI_L35_Playfield, LM_LI_L35_Rubber9, LM_LI_L35_Rubber9a, LM_LI_L35_Rubber9b, LM_LI_L35_sw74)
Dim BL_LI_L36: BL_LI_L36=Array(LM_LI_L36_Overlay, LM_LI_L36_Parts, LM_LI_L36_PegPlasticSlingR1C, LM_LI_L36_PegPlasticSlingRC, LM_LI_L36_Playfield, LM_LI_L36_Rubber9, LM_LI_L36_Rubber9a, LM_LI_L36_Rubber9b, LM_LI_L36_RubberRSling, LM_LI_L36_RubberRSlinga, LM_LI_L36_RubberRSlingb)
Dim BL_LI_L4: BL_LI_L4=Array(LM_LI_L4_Playfield)
Dim BL_LI_L5: BL_LI_L5=Array(LM_LI_L5_Parts, LM_LI_L5_PegPlasticSlingL1C, LM_LI_L5_Playfield, LM_LI_L5_Rubber10, LM_LI_L5_Rubber10a, LM_LI_L5_Rubber10b)
Dim BL_LI_L6: BL_LI_L6=Array(LM_LI_L6_Parts, LM_LI_L6_PegPlasticSlingL1C, LM_LI_L6_Playfield, LM_LI_L6_Rubber10, LM_LI_L6_Rubber10a, LM_LI_L6_Rubber10b, LM_LI_L6_RubberLSling, LM_LI_L6_RubberLSlinga, LM_LI_L6_RubberLSlingb)
Dim BL_LI_L7: BL_LI_L7=Array(LM_LI_L7_Playfield)
Dim BL_LI_L8: BL_LI_L8=Array(LM_LI_L8_Overlay, LM_LI_L8_Parts, LM_LI_L8_Playfield, LM_LI_L8_sw61, LM_LI_L8_sw63)
Dim BL_LI_L9: BL_LI_L9=Array(LM_LI_L9_Playfield)
Dim BL_World: BL_World=Array(BM_Bumper_Ring, BM_Bumper_Socket, BM_Gate2_Wire, BM_GateFlap, BM_LFlip, BM_LFlip1, BM_LFlip1U, BM_LFlipU, BM_Overlay, BM_Parts, BM_PegPlasticSlingL1C, BM_PegPlasticSlingLC, BM_PegPlasticSlingR1C, BM_PegPlasticSlingRC, BM_PkickarmR, BM_Playfield, BM_PostPlasticCapL1C, BM_PostPlasticCapLC, BM_PostPlasticCapR1C, BM_PostPlasticCapRC, BM_RFlip, BM_RFlip1, BM_RFlip1U, BM_RFlipU, BM_RailL, BM_RailR, BM_Rubber10, BM_Rubber10a, BM_Rubber10b, BM_Rubber2, BM_Rubber2a, BM_Rubber2b, BM_Rubber2c, BM_Rubber2d, BM_Rubber4, BM_Rubber4a, BM_Rubber4b, BM_Rubber4c, BM_Rubber4d, BM_Rubber7, BM_Rubber7a, BM_Rubber7b, BM_Rubber8, BM_Rubber8a, BM_Rubber8b, BM_Rubber9, BM_Rubber9a, BM_Rubber9b, BM_RubberLSling, BM_RubberLSlinga, BM_RubberLSlingb, BM_RubberRSling, BM_RubberRSlinga, BM_RubberRSlingb, BM_lockdownbar, BM_sw10, BM_sw11, BM_sw13, BM_sw14, BM_sw20, BM_sw21, BM_sw23, BM_sw24, BM_sw30, BM_sw31, BM_sw33, BM_sw34, BM_sw40, BM_sw43, BM_sw44, BM_sw50, BM_sw60, BM_sw61, BM_sw63, BM_sw64, BM_sw70, _
  BM_sw71, BM_sw73, BM_sw74)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper_Ring, BM_Bumper_Socket, BM_Gate2_Wire, BM_GateFlap, BM_LFlip, BM_LFlip1, BM_LFlip1U, BM_LFlipU, BM_Overlay, BM_Parts, BM_PegPlasticSlingL1C, BM_PegPlasticSlingLC, BM_PegPlasticSlingR1C, BM_PegPlasticSlingRC, BM_PkickarmR, BM_Playfield, BM_PostPlasticCapL1C, BM_PostPlasticCapLC, BM_PostPlasticCapR1C, BM_PostPlasticCapRC, BM_RFlip, BM_RFlip1, BM_RFlip1U, BM_RFlipU, BM_RailL, BM_RailR, BM_Rubber10, BM_Rubber10a, BM_Rubber10b, BM_Rubber2, BM_Rubber2a, BM_Rubber2b, BM_Rubber2c, BM_Rubber2d, BM_Rubber4, BM_Rubber4a, BM_Rubber4b, BM_Rubber4c, BM_Rubber4d, BM_Rubber7, BM_Rubber7a, BM_Rubber7b, BM_Rubber8, BM_Rubber8a, BM_Rubber8b, BM_Rubber9, BM_Rubber9a, BM_Rubber9b, BM_RubberLSling, BM_RubberLSlinga, BM_RubberLSlingb, BM_RubberRSling, BM_RubberRSlinga, BM_RubberRSlingb, BM_lockdownbar, BM_sw10, BM_sw11, BM_sw13, BM_sw14, BM_sw20, BM_sw21, BM_sw23, BM_sw24, BM_sw30, BM_sw31, BM_sw33, BM_sw34, BM_sw40, BM_sw43, BM_sw44, BM_sw50, BM_sw60, BM_sw61, BM_sw63, BM_sw64, BM_sw70, _
  BM_sw71, BM_sw73, BM_sw74)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_Bumper_Ring, LM_GI_Bumper_Socket, LM_GI_GateFlap, LM_GI_LFlip, LM_GI_LFlip1, LM_GI_LFlip1U, LM_GI_LFlipU, LM_GI_Overlay, LM_GI_Parts, LM_GI_PegPlasticSlingL1C, LM_GI_PegPlasticSlingLC, LM_GI_PegPlasticSlingR1C, LM_GI_PegPlasticSlingRC, LM_GI_PkickarmR, LM_GI_Playfield, LM_GI_PostPlasticCapL1C, LM_GI_PostPlasticCapLC, LM_GI_PostPlasticCapR1C, LM_GI_PostPlasticCapRC, LM_GI_RFlip, LM_GI_RFlip1, LM_GI_RFlip1U, LM_GI_RFlipU, LM_GI_Rubber10, LM_GI_Rubber10a, LM_GI_Rubber10b, LM_GI_Rubber2, LM_GI_Rubber2a, LM_GI_Rubber2b, LM_GI_Rubber2c, LM_GI_Rubber2d, LM_GI_Rubber4, LM_GI_Rubber4a, LM_GI_Rubber4b, LM_GI_Rubber4c, LM_GI_Rubber4d, LM_GI_Rubber7, LM_GI_Rubber7a, LM_GI_Rubber7b, LM_GI_Rubber8, LM_GI_Rubber8a, LM_GI_Rubber8b, LM_GI_Rubber9, LM_GI_Rubber9a, LM_GI_Rubber9b, LM_GI_RubberLSling, LM_GI_RubberLSlinga, LM_GI_RubberLSlingb, LM_GI_RubberRSling, LM_GI_RubberRSlinga, LM_GI_RubberRSlingb, LM_GI_sw10, LM_GI_sw11, LM_GI_sw13, LM_GI_sw14, LM_GI_sw20, LM_GI_sw21, LM_GI_sw23, _
  LM_GI_sw24, LM_GI_sw30, LM_GI_sw31, LM_GI_sw33, LM_GI_sw34, LM_GI_sw40, LM_GI_sw43, LM_GI_sw50, LM_GI_sw60, LM_GI_sw61, LM_GI_sw63, LM_GI_sw64, LM_GI_sw70, LM_GI_sw71, LM_GI_sw73, LM_GI_sw74, LM_GI_RT_GI_1_LFlip, LM_GI_RT_GI_1_LFlip1, LM_GI_RT_GI_1_LFlip1U, LM_GI_RT_GI_1_LFlipU, LM_GI_RT_GI_1_Overlay, LM_GI_RT_GI_1_Parts, LM_GI_RT_GI_1_PegPlasticSlingL1, LM_GI_RT_GI_1_PegPlasticSlingLC, LM_GI_RT_GI_1_Playfield, LM_GI_RT_GI_1_PostPlasticCapL1C, LM_GI_RT_GI_1_PostPlasticCapLC, LM_GI_RT_GI_1_RFlip, LM_GI_RT_GI_1_RFlipU, LM_GI_RT_GI_1_Rubber10, LM_GI_RT_GI_1_Rubber10a, LM_GI_RT_GI_1_Rubber10b, LM_GI_RT_GI_1_Rubber9, LM_GI_RT_GI_1_Rubber9a, LM_GI_RT_GI_1_Rubber9b, LM_GI_RT_GI_1_RubberLSling, LM_GI_RT_GI_1_RubberLSlinga, LM_GI_RT_GI_1_RubberLSlingb, LM_GI_RT_GI_1_sw10, LM_GI_RT_GI_1_sw11, LM_GI_RT_GI_1_sw20, LM_GI_RT_GI_1_sw21, LM_GI_RT_GI_1_sw23, LM_GI_RT_GI_1_sw24, LM_GI_RT_GI_1_sw40, LM_GI_RT_GI_1_sw50, LM_GI_RT_GI_10_Bumper_Ring, LM_GI_RT_GI_10_Bumper_Socket, LM_GI_RT_GI_10_LFlip1, LM_GI_RT_GI_10_LFlip1U, _
  LM_GI_RT_GI_10_Overlay, LM_GI_RT_GI_10_Parts, LM_GI_RT_GI_10_Playfield, LM_GI_RT_GI_10_Rubber2, LM_GI_RT_GI_10_Rubber2a, LM_GI_RT_GI_10_Rubber2b, LM_GI_RT_GI_10_Rubber2c, LM_GI_RT_GI_10_Rubber2d, LM_GI_RT_GI_10_Rubber4, LM_GI_RT_GI_10_Rubber4a, LM_GI_RT_GI_10_Rubber4b, LM_GI_RT_GI_10_Rubber4c, LM_GI_RT_GI_10_Rubber4d, LM_GI_RT_GI_10_Rubber8, LM_GI_RT_GI_10_Rubber8a, LM_GI_RT_GI_10_Rubber8b, LM_GI_RT_GI_10_Rubber9, LM_GI_RT_GI_10_Rubber9a, LM_GI_RT_GI_10_Rubber9b, LM_GI_RT_GI_10_sw30, LM_GI_RT_GI_10_sw31, LM_GI_RT_GI_10_sw33, LM_GI_RT_GI_10_sw34, LM_GI_RT_GI_10_sw40, LM_GI_RT_GI_10_sw43, LM_GI_RT_GI_10_sw50, LM_GI_RT_GI_10_sw60, LM_GI_RT_GI_10_sw61, LM_GI_RT_GI_10_sw63, LM_GI_RT_GI_10_sw64, LM_GI_RT_GI_11_Bumper_Socket, LM_GI_RT_GI_11_Overlay, LM_GI_RT_GI_11_Parts, LM_GI_RT_GI_11_Playfield, LM_GI_RT_GI_11_RFlip1, LM_GI_RT_GI_11_RFlip1U, LM_GI_RT_GI_11_Rubber2, LM_GI_RT_GI_11_Rubber2a, LM_GI_RT_GI_11_Rubber2b, LM_GI_RT_GI_11_Rubber2c, LM_GI_RT_GI_11_Rubber2d, LM_GI_RT_GI_11_Rubber4, LM_GI_RT_GI_11_Rubber4a, _
  LM_GI_RT_GI_11_Rubber4b, LM_GI_RT_GI_11_Rubber4c, LM_GI_RT_GI_11_Rubber4d, LM_GI_RT_GI_11_Rubber7, LM_GI_RT_GI_11_Rubber7a, LM_GI_RT_GI_11_Rubber7b, LM_GI_RT_GI_11_Rubber9, LM_GI_RT_GI_11_Rubber9a, LM_GI_RT_GI_11_Rubber9b, LM_GI_RT_GI_11_RubberLSling, LM_GI_RT_GI_11_RubberLSlinga, LM_GI_RT_GI_11_RubberLSlingb, LM_GI_RT_GI_11_sw44, LM_GI_RT_GI_11_sw50, LM_GI_RT_GI_11_sw60, LM_GI_RT_GI_11_sw61, LM_GI_RT_GI_11_sw63, LM_GI_RT_GI_11_sw64, LM_GI_RT_GI_16_Bumper_Socket, LM_GI_RT_GI_16_GateFlap, LM_GI_RT_GI_16_Overlay, LM_GI_RT_GI_16_Parts, LM_GI_RT_GI_16_Playfield, LM_GI_RT_GI_16_Rubber2, LM_GI_RT_GI_16_Rubber2a, LM_GI_RT_GI_16_Rubber2b, LM_GI_RT_GI_16_Rubber2c, LM_GI_RT_GI_16_Rubber2d, LM_GI_RT_GI_16_Rubber4, LM_GI_RT_GI_16_Rubber4a, LM_GI_RT_GI_16_Rubber4b, LM_GI_RT_GI_16_Rubber4c, LM_GI_RT_GI_16_Rubber4d, LM_GI_RT_GI_16_sw60, LM_GI_RT_GI_16_sw61, LM_GI_RT_GI_16_sw63, LM_GI_RT_GI_16_sw64, LM_GI_RT_GI_2_LFlip, LM_GI_RT_GI_2_LFlipU, LM_GI_RT_GI_2_Overlay, LM_GI_RT_GI_2_Parts, LM_GI_RT_GI_2_PegPlasticSlingR1, _
  LM_GI_RT_GI_2_PegPlasticSlingRC, LM_GI_RT_GI_2_Playfield, LM_GI_RT_GI_2_PostPlasticCapR1C, LM_GI_RT_GI_2_PostPlasticCapRC, LM_GI_RT_GI_2_RFlip, LM_GI_RT_GI_2_RFlipU, LM_GI_RT_GI_2_Rubber10, LM_GI_RT_GI_2_Rubber10a, LM_GI_RT_GI_2_Rubber10b, LM_GI_RT_GI_2_Rubber9, LM_GI_RT_GI_2_Rubber9a, LM_GI_RT_GI_2_Rubber9b, LM_GI_RT_GI_2_RubberRSling, LM_GI_RT_GI_2_RubberRSlinga, LM_GI_RT_GI_2_RubberRSlingb, LM_GI_RT_GI_2_sw13, LM_GI_RT_GI_2_sw14, LM_GI_RT_GI_2_sw24, LM_GI_RT_GI_2_sw50, LM_GI_RT_GI_4_LFlip1, LM_GI_RT_GI_4_LFlip1U, LM_GI_RT_GI_4_Overlay, LM_GI_RT_GI_4_Parts, LM_GI_RT_GI_4_PegPlasticSlingR1, LM_GI_RT_GI_4_Playfield, LM_GI_RT_GI_4_PostPlasticCapRC, LM_GI_RT_GI_4_RFlip, LM_GI_RT_GI_4_RFlip1, LM_GI_RT_GI_4_RFlip1U, LM_GI_RT_GI_4_RFlipU, LM_GI_RT_GI_4_Rubber10, LM_GI_RT_GI_4_Rubber10a, LM_GI_RT_GI_4_Rubber10b, LM_GI_RT_GI_4_Rubber4, LM_GI_RT_GI_4_Rubber4a, LM_GI_RT_GI_4_Rubber4b, LM_GI_RT_GI_4_Rubber4c, LM_GI_RT_GI_4_Rubber4d, LM_GI_RT_GI_4_Rubber7, LM_GI_RT_GI_4_Rubber7a, LM_GI_RT_GI_4_Rubber7b, _
  LM_GI_RT_GI_4_Rubber9, LM_GI_RT_GI_4_Rubber9a, LM_GI_RT_GI_4_Rubber9b, LM_GI_RT_GI_4_RubberRSling, LM_GI_RT_GI_4_RubberRSlinga, LM_GI_RT_GI_4_RubberRSlingb, LM_GI_RT_GI_4_sw21, LM_GI_RT_GI_4_sw23, LM_GI_RT_GI_4_sw24, LM_GI_RT_GI_4_sw40, LM_GI_RT_GI_4_sw50, LM_GI_RT_GI_4_sw60, LM_GI_RT_GI_4_sw61, LM_GI_RT_GI_4_sw63, LM_GI_RT_GI_4_sw64, LM_GI_RT_GI_4_sw70, LM_GI_RT_GI_4_sw71, LM_GI_RT_GI_4_sw73, LM_GI_RT_GI_4_sw74, LM_GI_RT_GI_5_Bumper_Ring, LM_GI_RT_GI_5_LFlip, LM_GI_RT_GI_5_LFlip1, LM_GI_RT_GI_5_LFlip1U, LM_GI_RT_GI_5_LFlipU, LM_GI_RT_GI_5_Overlay, LM_GI_RT_GI_5_Parts, LM_GI_RT_GI_5_PegPlasticSlingL1, LM_GI_RT_GI_5_PegPlasticSlingLC, LM_GI_RT_GI_5_Playfield, LM_GI_RT_GI_5_PostPlasticCapL1C, LM_GI_RT_GI_5_PostPlasticCapLC, LM_GI_RT_GI_5_RFlip1, LM_GI_RT_GI_5_RFlip1U, LM_GI_RT_GI_5_Rubber10, LM_GI_RT_GI_5_Rubber10a, LM_GI_RT_GI_5_Rubber10b, LM_GI_RT_GI_5_Rubber2, LM_GI_RT_GI_5_Rubber2a, LM_GI_RT_GI_5_Rubber2b, LM_GI_RT_GI_5_Rubber2c, LM_GI_RT_GI_5_Rubber2d, LM_GI_RT_GI_5_Rubber4, LM_GI_RT_GI_5_Rubber4a, _
  LM_GI_RT_GI_5_Rubber4b, LM_GI_RT_GI_5_Rubber4c, LM_GI_RT_GI_5_Rubber4d, LM_GI_RT_GI_5_Rubber8, LM_GI_RT_GI_5_Rubber8a, LM_GI_RT_GI_5_Rubber8b, LM_GI_RT_GI_5_Rubber9, LM_GI_RT_GI_5_Rubber9a, LM_GI_RT_GI_5_Rubber9b, LM_GI_RT_GI_5_sw10, LM_GI_RT_GI_5_sw11, LM_GI_RT_GI_5_sw20, LM_GI_RT_GI_5_sw21, LM_GI_RT_GI_5_sw23, LM_GI_RT_GI_5_sw24, LM_GI_RT_GI_5_sw40, LM_GI_RT_GI_5_sw50, LM_GI_RT_GI_5_sw60, LM_GI_RT_GI_5_sw61, LM_GI_RT_GI_7_Bumper_Ring, LM_GI_RT_GI_7_Bumper_Socket, LM_GI_RT_GI_7_Overlay, LM_GI_RT_GI_7_Parts, LM_GI_RT_GI_7_Playfield, LM_GI_RT_GI_7_Rubber2, LM_GI_RT_GI_7_Rubber2a, LM_GI_RT_GI_7_Rubber2b, LM_GI_RT_GI_7_Rubber2c, LM_GI_RT_GI_7_Rubber2d, LM_GI_RT_GI_7_Rubber4, LM_GI_RT_GI_7_Rubber4a, LM_GI_RT_GI_7_Rubber4b, LM_GI_RT_GI_7_Rubber4c, LM_GI_RT_GI_7_Rubber4d, LM_GI_RT_GI_7_sw30, LM_GI_RT_GI_7_sw31, LM_GI_RT_GI_7_sw33, LM_GI_RT_GI_7_sw34, LM_GI_RT_GI_8_Bumper_Ring, LM_GI_RT_GI_8_Bumper_Socket, LM_GI_RT_GI_8_Overlay, LM_GI_RT_GI_8_Parts, LM_GI_RT_GI_8_Playfield, LM_GI_RT_GI_8_Rubber2, _
  LM_GI_RT_GI_8_Rubber2a, LM_GI_RT_GI_8_Rubber2b, LM_GI_RT_GI_8_Rubber2c, LM_GI_RT_GI_8_Rubber2d, LM_GI_RT_GI_8_Rubber4, LM_GI_RT_GI_8_Rubber4a, LM_GI_RT_GI_8_Rubber4b, LM_GI_RT_GI_8_Rubber4c, LM_GI_RT_GI_8_Rubber4d, LM_GI_RT_GI_8_Rubber8, LM_GI_RT_GI_8_Rubber8a, LM_GI_RT_GI_8_Rubber8b, LM_GI_RT_GI_8_sw30, LM_GI_RT_GI_8_sw31, LM_GI_RT_GI_8_sw33, LM_GI_RT_GI_8_sw34, LM_GI_RT_GI_9_Bumper_Socket, LM_GI_RT_GI_9_Overlay, LM_GI_RT_GI_9_Parts, LM_GI_RT_GI_9_Playfield, LM_GI_RT_GI_9_Rubber2, LM_GI_RT_GI_9_Rubber2a, LM_GI_RT_GI_9_Rubber2b, LM_GI_RT_GI_9_Rubber2c, LM_GI_RT_GI_9_Rubber2d, LM_GI_RT_GI_9_Rubber4, LM_GI_RT_GI_9_Rubber4a, LM_GI_RT_GI_9_Rubber4b, LM_GI_RT_GI_9_Rubber4c, LM_GI_RT_GI_9_Rubber4d, LM_GI_RT_GI_9_Rubber7, LM_GI_RT_GI_9_Rubber7a, LM_GI_RT_GI_9_Rubber7b, LM_GI_RT_GI_9_sw60, LM_GI_RT_GI_9_sw61, LM_GI_RT_GI_9_sw63, LM_GI_RT_GI_9_sw64, LM_GI_RT_GI_99_LFlip, LM_GI_RT_GI_99_LFlipU, LM_GI_RT_GI_99_Parts, LM_GI_RT_GI_99_Playfield, LM_GI_RT_GI_99_RFlip, LM_GI_RT_GI_99_RFlipU, LM_GI_RT_GI_99_RailL, _
  LM_GI_RT_GI_99_RailR, LM_GI_RT_GI_99_lockdownbar, LM_LI_L10_Playfield, LM_LI_L11_Playfield, LM_LI_L11_sw50, LM_LI_L12_Playfield, LM_LI_L12_sw40, LM_LI_L13_Parts, LM_LI_L13_Playfield, LM_LI_L14_Parts, LM_LI_L14_Playfield, LM_LI_L14_Rubber7, LM_LI_L14_Rubber7a, LM_LI_L14_Rubber7b, LM_LI_L15_Parts, LM_LI_L15_Playfield, LM_LI_L16_Parts, LM_LI_L16_Playfield, LM_LI_L19_Playfield, LM_LI_L20_Playfield, LM_LI_L21_Playfield, LM_LI_L22_Playfield, LM_LI_L23_LFlip, LM_LI_L23_LFlipU, LM_LI_L23_Playfield, LM_LI_L23_RFlip, LM_LI_L23_RFlipU, LM_LI_L24_LFlipU, LM_LI_L24_Playfield, LM_LI_L25_Playfield, LM_LI_L25_RFlipU, LM_LI_L26_Playfield, LM_LI_L27_Playfield, LM_LI_L28_Playfield, LM_LI_L29_Playfield, LM_LI_L30_Playfield, LM_LI_L31_Playfield, LM_LI_L32_Playfield, LM_LI_L33_Playfield, LM_LI_L35_Overlay, LM_LI_L35_Parts, LM_LI_L35_PegPlasticSlingR1C, LM_LI_L35_Playfield, LM_LI_L35_Rubber9, LM_LI_L35_Rubber9a, LM_LI_L35_Rubber9b, LM_LI_L35_sw74, LM_LI_L36_Overlay, LM_LI_L36_Parts, LM_LI_L36_PegPlasticSlingR1C, _
  LM_LI_L36_PegPlasticSlingRC, LM_LI_L36_Playfield, LM_LI_L36_Rubber9, LM_LI_L36_Rubber9a, LM_LI_L36_Rubber9b, LM_LI_L36_RubberRSling, LM_LI_L36_RubberRSlinga, LM_LI_L36_RubberRSlingb, LM_LI_L4_Playfield, LM_LI_L5_Parts, LM_LI_L5_PegPlasticSlingL1C, LM_LI_L5_Playfield, LM_LI_L5_Rubber10, LM_LI_L5_Rubber10a, LM_LI_L5_Rubber10b, LM_LI_L6_Parts, LM_LI_L6_PegPlasticSlingL1C, LM_LI_L6_Playfield, LM_LI_L6_Rubber10, LM_LI_L6_Rubber10a, LM_LI_L6_Rubber10b, LM_LI_L6_RubberLSling, LM_LI_L6_RubberLSlinga, LM_LI_L6_RubberLSlingb, LM_LI_L7_Playfield, LM_LI_L8_Overlay, LM_LI_L8_Parts, LM_LI_L8_Playfield, LM_LI_L8_sw61, LM_LI_L8_sw63, LM_LI_L9_Playfield)
Dim BG_All: BG_All=Array(BM_Bumper_Ring, BM_Bumper_Socket, BM_Gate2_Wire, BM_GateFlap, BM_LFlip, BM_LFlip1, BM_LFlip1U, BM_LFlipU, BM_Overlay, BM_Parts, BM_PegPlasticSlingL1C, BM_PegPlasticSlingLC, BM_PegPlasticSlingR1C, BM_PegPlasticSlingRC, BM_PkickarmR, BM_Playfield, BM_PostPlasticCapL1C, BM_PostPlasticCapLC, BM_PostPlasticCapR1C, BM_PostPlasticCapRC, BM_RFlip, BM_RFlip1, BM_RFlip1U, BM_RFlipU, BM_RailL, BM_RailR, BM_Rubber10, BM_Rubber10a, BM_Rubber10b, BM_Rubber2, BM_Rubber2a, BM_Rubber2b, BM_Rubber2c, BM_Rubber2d, BM_Rubber4, BM_Rubber4a, BM_Rubber4b, BM_Rubber4c, BM_Rubber4d, BM_Rubber7, BM_Rubber7a, BM_Rubber7b, BM_Rubber8, BM_Rubber8a, BM_Rubber8b, BM_Rubber9, BM_Rubber9a, BM_Rubber9b, BM_RubberLSling, BM_RubberLSlinga, BM_RubberLSlingb, BM_RubberRSling, BM_RubberRSlinga, BM_RubberRSlingb, BM_lockdownbar, BM_sw10, BM_sw11, BM_sw13, BM_sw14, BM_sw20, BM_sw21, BM_sw23, BM_sw24, BM_sw30, BM_sw31, BM_sw33, BM_sw34, BM_sw40, BM_sw43, BM_sw44, BM_sw50, BM_sw60, BM_sw61, BM_sw63, BM_sw64, BM_sw70, BM_sw71, _
  BM_sw73, BM_sw74, LM_GI_Bumper_Ring, LM_GI_Bumper_Socket, LM_GI_GateFlap, LM_GI_LFlip, LM_GI_LFlip1, LM_GI_LFlip1U, LM_GI_LFlipU, LM_GI_Overlay, LM_GI_Parts, LM_GI_PegPlasticSlingL1C, LM_GI_PegPlasticSlingLC, LM_GI_PegPlasticSlingR1C, LM_GI_PegPlasticSlingRC, LM_GI_PkickarmR, LM_GI_Playfield, LM_GI_PostPlasticCapL1C, LM_GI_PostPlasticCapLC, LM_GI_PostPlasticCapR1C, LM_GI_PostPlasticCapRC, LM_GI_RFlip, LM_GI_RFlip1, LM_GI_RFlip1U, LM_GI_RFlipU, LM_GI_Rubber10, LM_GI_Rubber10a, LM_GI_Rubber10b, LM_GI_Rubber2, LM_GI_Rubber2a, LM_GI_Rubber2b, LM_GI_Rubber2c, LM_GI_Rubber2d, LM_GI_Rubber4, LM_GI_Rubber4a, LM_GI_Rubber4b, LM_GI_Rubber4c, LM_GI_Rubber4d, LM_GI_Rubber7, LM_GI_Rubber7a, LM_GI_Rubber7b, LM_GI_Rubber8, LM_GI_Rubber8a, LM_GI_Rubber8b, LM_GI_Rubber9, LM_GI_Rubber9a, LM_GI_Rubber9b, LM_GI_RubberLSling, LM_GI_RubberLSlinga, LM_GI_RubberLSlingb, LM_GI_RubberRSling, LM_GI_RubberRSlinga, LM_GI_RubberRSlingb, LM_GI_sw10, LM_GI_sw11, LM_GI_sw13, LM_GI_sw14, LM_GI_sw20, LM_GI_sw21, LM_GI_sw23, LM_GI_sw24, _
  LM_GI_sw30, LM_GI_sw31, LM_GI_sw33, LM_GI_sw34, LM_GI_sw40, LM_GI_sw43, LM_GI_sw50, LM_GI_sw60, LM_GI_sw61, LM_GI_sw63, LM_GI_sw64, LM_GI_sw70, LM_GI_sw71, LM_GI_sw73, LM_GI_sw74, LM_GI_RT_GI_1_LFlip, LM_GI_RT_GI_1_LFlip1, LM_GI_RT_GI_1_LFlip1U, LM_GI_RT_GI_1_LFlipU, LM_GI_RT_GI_1_Overlay, LM_GI_RT_GI_1_Parts, LM_GI_RT_GI_1_PegPlasticSlingL1, LM_GI_RT_GI_1_PegPlasticSlingLC, LM_GI_RT_GI_1_Playfield, LM_GI_RT_GI_1_PostPlasticCapL1C, LM_GI_RT_GI_1_PostPlasticCapLC, LM_GI_RT_GI_1_RFlip, LM_GI_RT_GI_1_RFlipU, LM_GI_RT_GI_1_Rubber10, LM_GI_RT_GI_1_Rubber10a, LM_GI_RT_GI_1_Rubber10b, LM_GI_RT_GI_1_Rubber9, LM_GI_RT_GI_1_Rubber9a, LM_GI_RT_GI_1_Rubber9b, LM_GI_RT_GI_1_RubberLSling, LM_GI_RT_GI_1_RubberLSlinga, LM_GI_RT_GI_1_RubberLSlingb, LM_GI_RT_GI_1_sw10, LM_GI_RT_GI_1_sw11, LM_GI_RT_GI_1_sw20, LM_GI_RT_GI_1_sw21, LM_GI_RT_GI_1_sw23, LM_GI_RT_GI_1_sw24, LM_GI_RT_GI_1_sw40, LM_GI_RT_GI_1_sw50, LM_GI_RT_GI_10_Bumper_Ring, LM_GI_RT_GI_10_Bumper_Socket, LM_GI_RT_GI_10_LFlip1, LM_GI_RT_GI_10_LFlip1U, _
  LM_GI_RT_GI_10_Overlay, LM_GI_RT_GI_10_Parts, LM_GI_RT_GI_10_Playfield, LM_GI_RT_GI_10_Rubber2, LM_GI_RT_GI_10_Rubber2a, LM_GI_RT_GI_10_Rubber2b, LM_GI_RT_GI_10_Rubber2c, LM_GI_RT_GI_10_Rubber2d, LM_GI_RT_GI_10_Rubber4, LM_GI_RT_GI_10_Rubber4a, LM_GI_RT_GI_10_Rubber4b, LM_GI_RT_GI_10_Rubber4c, LM_GI_RT_GI_10_Rubber4d, LM_GI_RT_GI_10_Rubber8, LM_GI_RT_GI_10_Rubber8a, LM_GI_RT_GI_10_Rubber8b, LM_GI_RT_GI_10_Rubber9, LM_GI_RT_GI_10_Rubber9a, LM_GI_RT_GI_10_Rubber9b, LM_GI_RT_GI_10_sw30, LM_GI_RT_GI_10_sw31, LM_GI_RT_GI_10_sw33, LM_GI_RT_GI_10_sw34, LM_GI_RT_GI_10_sw40, LM_GI_RT_GI_10_sw43, LM_GI_RT_GI_10_sw50, LM_GI_RT_GI_10_sw60, LM_GI_RT_GI_10_sw61, LM_GI_RT_GI_10_sw63, LM_GI_RT_GI_10_sw64, LM_GI_RT_GI_11_Bumper_Socket, LM_GI_RT_GI_11_Overlay, LM_GI_RT_GI_11_Parts, LM_GI_RT_GI_11_Playfield, LM_GI_RT_GI_11_RFlip1, LM_GI_RT_GI_11_RFlip1U, LM_GI_RT_GI_11_Rubber2, LM_GI_RT_GI_11_Rubber2a, LM_GI_RT_GI_11_Rubber2b, LM_GI_RT_GI_11_Rubber2c, LM_GI_RT_GI_11_Rubber2d, LM_GI_RT_GI_11_Rubber4, LM_GI_RT_GI_11_Rubber4a, _
  LM_GI_RT_GI_11_Rubber4b, LM_GI_RT_GI_11_Rubber4c, LM_GI_RT_GI_11_Rubber4d, LM_GI_RT_GI_11_Rubber7, LM_GI_RT_GI_11_Rubber7a, LM_GI_RT_GI_11_Rubber7b, LM_GI_RT_GI_11_Rubber9, LM_GI_RT_GI_11_Rubber9a, LM_GI_RT_GI_11_Rubber9b, LM_GI_RT_GI_11_RubberLSling, LM_GI_RT_GI_11_RubberLSlinga, LM_GI_RT_GI_11_RubberLSlingb, LM_GI_RT_GI_11_sw44, LM_GI_RT_GI_11_sw50, LM_GI_RT_GI_11_sw60, LM_GI_RT_GI_11_sw61, LM_GI_RT_GI_11_sw63, LM_GI_RT_GI_11_sw64, LM_GI_RT_GI_16_Bumper_Socket, LM_GI_RT_GI_16_GateFlap, LM_GI_RT_GI_16_Overlay, LM_GI_RT_GI_16_Parts, LM_GI_RT_GI_16_Playfield, LM_GI_RT_GI_16_Rubber2, LM_GI_RT_GI_16_Rubber2a, LM_GI_RT_GI_16_Rubber2b, LM_GI_RT_GI_16_Rubber2c, LM_GI_RT_GI_16_Rubber2d, LM_GI_RT_GI_16_Rubber4, LM_GI_RT_GI_16_Rubber4a, LM_GI_RT_GI_16_Rubber4b, LM_GI_RT_GI_16_Rubber4c, LM_GI_RT_GI_16_Rubber4d, LM_GI_RT_GI_16_sw60, LM_GI_RT_GI_16_sw61, LM_GI_RT_GI_16_sw63, LM_GI_RT_GI_16_sw64, LM_GI_RT_GI_2_LFlip, LM_GI_RT_GI_2_LFlipU, LM_GI_RT_GI_2_Overlay, LM_GI_RT_GI_2_Parts, LM_GI_RT_GI_2_PegPlasticSlingR1, _
  LM_GI_RT_GI_2_PegPlasticSlingRC, LM_GI_RT_GI_2_Playfield, LM_GI_RT_GI_2_PostPlasticCapR1C, LM_GI_RT_GI_2_PostPlasticCapRC, LM_GI_RT_GI_2_RFlip, LM_GI_RT_GI_2_RFlipU, LM_GI_RT_GI_2_Rubber10, LM_GI_RT_GI_2_Rubber10a, LM_GI_RT_GI_2_Rubber10b, LM_GI_RT_GI_2_Rubber9, LM_GI_RT_GI_2_Rubber9a, LM_GI_RT_GI_2_Rubber9b, LM_GI_RT_GI_2_RubberRSling, LM_GI_RT_GI_2_RubberRSlinga, LM_GI_RT_GI_2_RubberRSlingb, LM_GI_RT_GI_2_sw13, LM_GI_RT_GI_2_sw14, LM_GI_RT_GI_2_sw24, LM_GI_RT_GI_2_sw50, LM_GI_RT_GI_4_LFlip1, LM_GI_RT_GI_4_LFlip1U, LM_GI_RT_GI_4_Overlay, LM_GI_RT_GI_4_Parts, LM_GI_RT_GI_4_PegPlasticSlingR1, LM_GI_RT_GI_4_Playfield, LM_GI_RT_GI_4_PostPlasticCapRC, LM_GI_RT_GI_4_RFlip, LM_GI_RT_GI_4_RFlip1, LM_GI_RT_GI_4_RFlip1U, LM_GI_RT_GI_4_RFlipU, LM_GI_RT_GI_4_Rubber10, LM_GI_RT_GI_4_Rubber10a, LM_GI_RT_GI_4_Rubber10b, LM_GI_RT_GI_4_Rubber4, LM_GI_RT_GI_4_Rubber4a, LM_GI_RT_GI_4_Rubber4b, LM_GI_RT_GI_4_Rubber4c, LM_GI_RT_GI_4_Rubber4d, LM_GI_RT_GI_4_Rubber7, LM_GI_RT_GI_4_Rubber7a, LM_GI_RT_GI_4_Rubber7b, _
  LM_GI_RT_GI_4_Rubber9, LM_GI_RT_GI_4_Rubber9a, LM_GI_RT_GI_4_Rubber9b, LM_GI_RT_GI_4_RubberRSling, LM_GI_RT_GI_4_RubberRSlinga, LM_GI_RT_GI_4_RubberRSlingb, LM_GI_RT_GI_4_sw21, LM_GI_RT_GI_4_sw23, LM_GI_RT_GI_4_sw24, LM_GI_RT_GI_4_sw40, LM_GI_RT_GI_4_sw50, LM_GI_RT_GI_4_sw60, LM_GI_RT_GI_4_sw61, LM_GI_RT_GI_4_sw63, LM_GI_RT_GI_4_sw64, LM_GI_RT_GI_4_sw70, LM_GI_RT_GI_4_sw71, LM_GI_RT_GI_4_sw73, LM_GI_RT_GI_4_sw74, LM_GI_RT_GI_5_Bumper_Ring, LM_GI_RT_GI_5_LFlip, LM_GI_RT_GI_5_LFlip1, LM_GI_RT_GI_5_LFlip1U, LM_GI_RT_GI_5_LFlipU, LM_GI_RT_GI_5_Overlay, LM_GI_RT_GI_5_Parts, LM_GI_RT_GI_5_PegPlasticSlingL1, LM_GI_RT_GI_5_PegPlasticSlingLC, LM_GI_RT_GI_5_Playfield, LM_GI_RT_GI_5_PostPlasticCapL1C, LM_GI_RT_GI_5_PostPlasticCapLC, LM_GI_RT_GI_5_RFlip1, LM_GI_RT_GI_5_RFlip1U, LM_GI_RT_GI_5_Rubber10, LM_GI_RT_GI_5_Rubber10a, LM_GI_RT_GI_5_Rubber10b, LM_GI_RT_GI_5_Rubber2, LM_GI_RT_GI_5_Rubber2a, LM_GI_RT_GI_5_Rubber2b, LM_GI_RT_GI_5_Rubber2c, LM_GI_RT_GI_5_Rubber2d, LM_GI_RT_GI_5_Rubber4, LM_GI_RT_GI_5_Rubber4a, _
  LM_GI_RT_GI_5_Rubber4b, LM_GI_RT_GI_5_Rubber4c, LM_GI_RT_GI_5_Rubber4d, LM_GI_RT_GI_5_Rubber8, LM_GI_RT_GI_5_Rubber8a, LM_GI_RT_GI_5_Rubber8b, LM_GI_RT_GI_5_Rubber9, LM_GI_RT_GI_5_Rubber9a, LM_GI_RT_GI_5_Rubber9b, LM_GI_RT_GI_5_sw10, LM_GI_RT_GI_5_sw11, LM_GI_RT_GI_5_sw20, LM_GI_RT_GI_5_sw21, LM_GI_RT_GI_5_sw23, LM_GI_RT_GI_5_sw24, LM_GI_RT_GI_5_sw40, LM_GI_RT_GI_5_sw50, LM_GI_RT_GI_5_sw60, LM_GI_RT_GI_5_sw61, LM_GI_RT_GI_7_Bumper_Ring, LM_GI_RT_GI_7_Bumper_Socket, LM_GI_RT_GI_7_Overlay, LM_GI_RT_GI_7_Parts, LM_GI_RT_GI_7_Playfield, LM_GI_RT_GI_7_Rubber2, LM_GI_RT_GI_7_Rubber2a, LM_GI_RT_GI_7_Rubber2b, LM_GI_RT_GI_7_Rubber2c, LM_GI_RT_GI_7_Rubber2d, LM_GI_RT_GI_7_Rubber4, LM_GI_RT_GI_7_Rubber4a, LM_GI_RT_GI_7_Rubber4b, LM_GI_RT_GI_7_Rubber4c, LM_GI_RT_GI_7_Rubber4d, LM_GI_RT_GI_7_sw30, LM_GI_RT_GI_7_sw31, LM_GI_RT_GI_7_sw33, LM_GI_RT_GI_7_sw34, LM_GI_RT_GI_8_Bumper_Ring, LM_GI_RT_GI_8_Bumper_Socket, LM_GI_RT_GI_8_Overlay, LM_GI_RT_GI_8_Parts, LM_GI_RT_GI_8_Playfield, LM_GI_RT_GI_8_Rubber2, _
  LM_GI_RT_GI_8_Rubber2a, LM_GI_RT_GI_8_Rubber2b, LM_GI_RT_GI_8_Rubber2c, LM_GI_RT_GI_8_Rubber2d, LM_GI_RT_GI_8_Rubber4, LM_GI_RT_GI_8_Rubber4a, LM_GI_RT_GI_8_Rubber4b, LM_GI_RT_GI_8_Rubber4c, LM_GI_RT_GI_8_Rubber4d, LM_GI_RT_GI_8_Rubber8, LM_GI_RT_GI_8_Rubber8a, LM_GI_RT_GI_8_Rubber8b, LM_GI_RT_GI_8_sw30, LM_GI_RT_GI_8_sw31, LM_GI_RT_GI_8_sw33, LM_GI_RT_GI_8_sw34, LM_GI_RT_GI_9_Bumper_Socket, LM_GI_RT_GI_9_Overlay, LM_GI_RT_GI_9_Parts, LM_GI_RT_GI_9_Playfield, LM_GI_RT_GI_9_Rubber2, LM_GI_RT_GI_9_Rubber2a, LM_GI_RT_GI_9_Rubber2b, LM_GI_RT_GI_9_Rubber2c, LM_GI_RT_GI_9_Rubber2d, LM_GI_RT_GI_9_Rubber4, LM_GI_RT_GI_9_Rubber4a, LM_GI_RT_GI_9_Rubber4b, LM_GI_RT_GI_9_Rubber4c, LM_GI_RT_GI_9_Rubber4d, LM_GI_RT_GI_9_Rubber7, LM_GI_RT_GI_9_Rubber7a, LM_GI_RT_GI_9_Rubber7b, LM_GI_RT_GI_9_sw60, LM_GI_RT_GI_9_sw61, LM_GI_RT_GI_9_sw63, LM_GI_RT_GI_9_sw64, LM_GI_RT_GI_99_LFlip, LM_GI_RT_GI_99_LFlipU, LM_GI_RT_GI_99_Parts, LM_GI_RT_GI_99_Playfield, LM_GI_RT_GI_99_RFlip, LM_GI_RT_GI_99_RFlipU, LM_GI_RT_GI_99_RailL, _
  LM_GI_RT_GI_99_RailR, LM_GI_RT_GI_99_lockdownbar, LM_LI_L10_Playfield, LM_LI_L11_Playfield, LM_LI_L11_sw50, LM_LI_L12_Playfield, LM_LI_L12_sw40, LM_LI_L13_Parts, LM_LI_L13_Playfield, LM_LI_L14_Parts, LM_LI_L14_Playfield, LM_LI_L14_Rubber7, LM_LI_L14_Rubber7a, LM_LI_L14_Rubber7b, LM_LI_L15_Parts, LM_LI_L15_Playfield, LM_LI_L16_Parts, LM_LI_L16_Playfield, LM_LI_L19_Playfield, LM_LI_L20_Playfield, LM_LI_L21_Playfield, LM_LI_L22_Playfield, LM_LI_L23_LFlip, LM_LI_L23_LFlipU, LM_LI_L23_Playfield, LM_LI_L23_RFlip, LM_LI_L23_RFlipU, LM_LI_L24_LFlipU, LM_LI_L24_Playfield, LM_LI_L25_Playfield, LM_LI_L25_RFlipU, LM_LI_L26_Playfield, LM_LI_L27_Playfield, LM_LI_L28_Playfield, LM_LI_L29_Playfield, LM_LI_L30_Playfield, LM_LI_L31_Playfield, LM_LI_L32_Playfield, LM_LI_L33_Playfield, LM_LI_L35_Overlay, LM_LI_L35_Parts, LM_LI_L35_PegPlasticSlingR1C, LM_LI_L35_Playfield, LM_LI_L35_Rubber9, LM_LI_L35_Rubber9a, LM_LI_L35_Rubber9b, LM_LI_L35_sw74, LM_LI_L36_Overlay, LM_LI_L36_Parts, LM_LI_L36_PegPlasticSlingR1C, _
  LM_LI_L36_PegPlasticSlingRC, LM_LI_L36_Playfield, LM_LI_L36_Rubber9, LM_LI_L36_Rubber9a, LM_LI_L36_Rubber9b, LM_LI_L36_RubberRSling, LM_LI_L36_RubberRSlinga, LM_LI_L36_RubberRSlingb, LM_LI_L4_Playfield, LM_LI_L5_Parts, LM_LI_L5_PegPlasticSlingL1C, LM_LI_L5_Playfield, LM_LI_L5_Rubber10, LM_LI_L5_Rubber10a, LM_LI_L5_Rubber10b, LM_LI_L6_Parts, LM_LI_L6_PegPlasticSlingL1C, LM_LI_L6_Playfield, LM_LI_L6_Rubber10, LM_LI_L6_Rubber10a, LM_LI_L6_Rubber10b, LM_LI_L6_RubberLSling, LM_LI_L6_RubberLSlinga, LM_LI_L6_RubberLSlingb, LM_LI_L7_Playfield, LM_LI_L8_Overlay, LM_LI_L8_Parts, LM_LI_L8_Playfield, LM_LI_L8_sw61, LM_LI_L8_sw63, LM_LI_L9_Playfield)
' VLM  Arrays - End

Const DEBUG_POSTIT_SCORES = False

'******************************************************
'  ZVAR: Constants and Global Variables
'******************************************************
Const BallSize = 50   'Ball size must be 50
Const BallMass = 1    'Ball mass must be 1
Const tnob = 1      'Total number of balls on the playfield including captive balls.
Const lob = 0     'Total number of locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

'  Standard definitions
Const cGameName = "countdwn"    'PinMAME ROM name
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""      'Sound sample used for this, obsolete.
Const SSolenoidOff = ""     ' ^
Const SFlipperOn = ""     ' ^
Const SFlipperOff = ""      ' ^
Const SCoin = ""        ' ^
Const VRTest = 0

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
    'lrail.Visible = 0
    'rrail.Visible = 0
end if

if B2SOn = true then VarHidden = 1

'----- VR Room Auto-Detect -----
Dim VRRoom, VR_Obj, VRMode, BP

If RenderingMode = 2 or VRTest = 1 Then
  'Textbox1.visible = 0
  VarHidden = 0
  VRMode = True
  For Each BP in BP_RailL : BP.Visible = 0: Next
  For Each BP in BP_RailR: BP.Visible = 0: Next
  For Each BP in BP_lockdownbar : BP.Visible = 0: Next
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
Else
  VRMode = False
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
End If

LoadVPM "03060000", "gts1.vbs", 3.26


'******************************************************
'  ZTIM: Timers
'******************************************************

'The FrameTimer interval should be -1, so executes at the display frame rate
'The frame timer should be used to update anything visual, like some animations, shadows, etc.
'However, a lot of animations will be handled in their respective _animate subroutines.

Dim FrameTime, InitFrameTime
InitFrameTime = 0

FrameTimer.Interval = -1
Sub FrameTimer_Timer()
  FrameTime = gametime - InitFrameTime 'Calculate FrameTime as some animuations could use this
  InitFrameTime = gametime  'Count frametime
  GIUpdate
  'Add animation stuff here
  BSUpdate
  UpdateBallBrightness
  RollingUpdate       'Update rolling sounds
  DoDTAnim          'Drop target animations
  If VRMode = True Then
    VRDisplayTimer
  Else
    UpdateLeds
VRDisplayTimer
  End if
  AnimateDropTargets
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub


'************
' Table init.
'************

Dim x, CDBall, gBOT

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Count-Down - Gottlieb 1979"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
  End With
  On Error Resume Next
  Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  'Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object
  vpmMapLights AllLamps     'Make a collection called "AllLamps" and put all the light objects in it.

    ' Nudging
    vpmNudge.TiltSwitch = 4
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1)

  'Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  'Trough - Creates a ball in the kicker switch and gives that ball used an individual name.
  Set CDBall = Drain.CreateSizedballWithMass(Ballsize/2,Ballmass)

  'Forces the trough switches to "on" at table boot so the game logic knows there are balls in the trough.
  Controller.Switch(66) = 1

  '***Setting up a ball array (collection), must contain all the balls you create on the table.
  gBOT = Array(CDBall)


    ' Turn on Gi
    GiOn

    ' Start the RealTime timer
    'RealTime.Enabled = 1

    ' Load table color
    'LoadLut

  InitRubberState
  HighScoreLoad

  If VRMode = True Then
    setup_backglass
  End if
End Sub

Sub Table_Paused:Controller.Pause = 1:End Sub
Sub Table_unPaused:Controller.Pause = 0:End Sub

' Thalamus - table is named Table1 - not Table !
Sub Table1_exit()
  debug.print "Exiting"
  HighScoreSave
  Controller.Pause = False
  Controller.Stop
End Sub


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1


' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Dim ApronLightOption
Dim PostItNoteEnabled : PostItNoteEnabled = True

Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 0, 11, 1, 0, 0, _
    Array("Normal", "Vibrant", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 0 Then Table1.ColorGradeImage = "colorgradelut256x16_1to1"
  if ColorLUT = 1 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat100"
  if ColorLUT = 2 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat90"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat80"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat70"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat60"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat50"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat40"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat30"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat20"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat10"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat00"

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

  Dim v,r
  Dim BP
    ' Toggle Rails
  If RenderingMode <> 2 and VRMode = False Then
    v = Table1.Option("Show Rails", 0, 1, 1, 1, 0, Array("Hide", "Show"))
    For Each BP in BP_RailL : BP.Visible = v: Next
    For Each BP in BP_RailR: BP.Visible = v: Next
'   v = JunglePrincess.Option("Desktop background", 0, 2, 1, 1, 0, Array("Blank", "Brick Wall", "Arcade"))
'   Select Case v
'     Case 0 : JunglePrincess.BackdropImage = ""
'     Case 1 : JunglePrincess.BackdropImage = "JunglePrincessBGDT"
'     Case 2 : JunglePrincess.BackdropImage = "DT_BG_NextLevel"
'   End Select
  End If

    ' Toggle Lockdown Bar
  If RenderingMode <> 2 and VRMode = False Then
    v = Table1.Option("Show Lockbar", 0, 1, 1, 1, 0, Array("Hide", "Show"))
    For Each BP in BP_lockdownbar : BP.Visible = v: Next
  End If

    ' Toggle Apron Light
  Dim ApronLightCurrentState: ApronLightCurrentState = GI_99.State
  v = Table1.Option("Apron Light", 0, 1, 1, 1, 0, Array("Off (Authentic)", "On (Mod)"))
  ApronLightOption = v
  If ApronLightOption Then 'and gilvl Then
    GI_99.State = True
  Else
    GI_99.State = False
  End If

  ' Playfield Reflections
  v = Table1.Option("Playfield Reflections", 0, 1, 1, 1, 0, Array("Clean", "Rough"))
  Select Case v
    Case 0: playfield_mesh.ReflectionProbe = "Playfield Reflections": BM_Playfield.ReflectionProbe = "Playfield Reflections"
    Case 1: playfield_mesh.ReflectionProbe = "Playfield Reflections Rough": BM_Playfield.ReflectionProbe = "Playfield Reflections Rough"
  End Select

    ' Toggle Reflections Scope
  v = Table1.Option("Playfield Reflections Scope", 0, 2, 1, 1, 0, Array("None", "Partial", "Full"))
  ReflectionToggle(v)

    ' Outlane Difficulty
  v = Table1.Option("Outlanes difficulty", 0, 1, 1, 0, 0, Array("Conservative", "Liberal"))
  SetOutlaneDifficulty(v)

    ' Flipper Distance
  v = Table1.Option("Flipper distance", 0, 1, 1, 0, 0, Array("Correct", "Casual"))
  SetFlipperDistance(v)

  ' Display Post-it note for high scores
  v = Table1.Option("Post-It High Scores", 0, 1, 1, 1, 0, Array("Off", "On"))
  if v=1 then PostItNoteEnabled = True else PostItNoteEnabled = False
  For Each r in PostIt : r.Visible = PostItNoteEnabled: Next
  if PostItNoteEnabled then HighScoreTimer.enabled=True

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Function SetFlipperDistance(d)
  Dim BP, deltaVL, deltaVR, deltaTL, deltaTR, deltaFL, deltaFR, deltaSL, deltaSR
  Select Case d
    Case 0:
      deltaVR = 598.2261
      deltaVL = 259.6957
      deltaFR = 259
      deltaFL = 599
      deltaTR = 633.2383
      deltaTL = 221.9266
      deltaSL = 262.4177
      deltaSR = 595.8148
    Case 1:
      deltaVR = 594.2261
      deltaVL = 263.6957
      deltaFR = 263
      deltaFL = 595
      deltaTR = 629.2383
      deltaTL = 225.9266
      deltaSL = 267.4177
      deltaSR = 591.8148
  End Select

  For Each BP in BP_LFlip:  BP.x = deltaVL : Next
  For Each BP in BP_LFlipU: BP.x = deltaVL : Next
  For Each BP in BP_RFlip:  BP.x = deltaVR : Next
  For Each BP in BP_RFlipU: BP.x = deltaVR : Next
  LeftFlipper.x  = deltaFL
  RightFlipper.x = deltaFR
  TriggerLF.x    = deltaTL
  TriggerRF.x    = deltaTR
  FlipperLSh.x   = deltaSL
  FlipperRSh.x   = deltaSR
End Function

Function SetOutlaneDifficulty(d)
  Dim BP, v1, v2
  ' Right Side Phys: zCol_Rubber_Post_RC, zCol_Rubber_Post_RL, zCol_Rubber_Post_R1C. zCol_Rubber_Post_R1L, sw53j_RL, sw53j_RC
  ' Right Side BMs : BM_PegPlasticSlingRC, BM_PegPlasticSlingR1C (this is actually L)
  ' Left Side Phys : zCol_Rubber_Post_LC, zCol_Rubber_Post_LL, zCol_Rubber_Post_L1L, zCol_Rubber_Post_L1C, sw53i_LL, sw53i_LC
  ' Left Side BMs  : BM_PegPlasticSlingLC, BM_PegPlasticSlingL1C (this is actually C)
  Select Case d
    Case 0: ' Conservative (Hard)
      v1 = False: v2 =True
      For Each BP in BP_PegPlasticSlingRC  : BP.transz = 0 : Next
      For Each BP in BP_PegPlasticSlingR1C : BP.transz = 0 : Next
      For Each BP in BP_PegPlasticSlingLC  : BP.transz = 0 : Next
      For Each BP in BP_PegPlasticSlingL1C : BP.transz = 0 : Next
      For Each BP in BP_PostPlasticCapRC   : BP.transy = 0 : Next
      For Each BP in BP_PostPlasticCapR1C  : BP.transy = 0 : Next
      For Each BP in BP_PostPlasticCapLC   : BP.transy = 0 : Next
      For Each BP in BP_PostPlasticCapL1C  : BP.transy = 0 : Next
    Case 1: ' Liberal (Easy)
      v1 = True:  v2 = False
      For Each BP in BP_PegPlasticSlingRC  : BP.transz = -10 : Next
      For Each BP in BP_PegPlasticSlingR1C : BP.transz = 10  : Next
      For Each BP in BP_PegPlasticSlingLC  : BP.transz = -10 : Next
      For Each BP in BP_PegPlasticSlingL1C : BP.transz = 10  : Next
      For Each BP in BP_PostPlasticCapRC   : BP.transy = -10 : Next
      For Each BP in BP_PostPlasticCapR1C  : BP.transy = 10  : Next
      For Each BP in BP_PostPlasticCapLC   : BP.transy = -10 : Next
      For Each BP in BP_PostPlasticCapL1C  : BP.transy = 10  : Next

  End Select
  ' set the BPs (did not render multiple rubber positions so disabling)
' For Each BP in BP_RPostRubber_p1 : BP.Visible = v1: Next
' For Each BP in BP_RPostRubber_p2 : BP.Visible = v2: Next
' For Each BP in BP_LPostRubber_p1 : BP.Visible = v1: Next
' For Each BP in BP_LPostRubber_p2 : BP.Visible = v2: Next

  ' set the physics
  zCol_Rubber_Post_RC.collidable   = v2:  zCol_Rubber_Post_LC.collidable  = v2
  zCol_Rubber_Post_R1L.collidable  = v2:  zCol_Rubber_Post_L1C.collidable = v2
  sw53j_RL.collidable              = v2:  sw53i_LL.collidable             = v2
  RubberBand_RL.collidable     = v2:  RubberBand_LL.collidable    = v2
  RubberBand1_RL.collidable    = v2:  RubberBand1_LL.collidable   = v2

  zCol_Rubber_Post_RL.collidable   = v1:  zCol_Rubber_Post_LL.collidable  = v1
  zCol_Rubber_Post_R1C.collidable  = v1:  zCol_Rubber_Post_L1L.collidable = v1
  sw53j_RC.collidable              = v1:  sw53i_LC.collidable             = v1
  RubberBand_RC.collidable     = v1:  RubberBand_LC.collidable    = v1
  RubberBand1_RC.collidable    = v1:  RubberBand1_RC.collidable   = v1

  ' Debug
' zCol_Rubber_Post_RC.visible  = v2
' zCol_Rubber_Post_LC.visible  = v2
' zCol_Rubber_Post_R1L.visible = v2
' zCol_Rubber_Post_L1C.visible = v2
' zCol_Rubber_Post_RL.visible  = v1
' zCol_Rubber_Post_LL.visible  = v1
' zCol_Rubber_Post_R1C.visible = v1
' zCol_Rubber_Post_L1L.visible = v1

End Function

Function ReflectionToggle(state)
  Dim ReflObjectsArray: ReflObjectsArray = Array(BP_Parts, BP_Overlay, BP_LFlip, BP_LFlipU, BP_RFlip, BP_RFlipU, BP_Bumper_Ring, BP_sw70, BP_sw71, BP_sw73, BP_sw74, BP_sw20, BP_sw21, BP_sw23, BP_sw24, BP_sw30, BP_sw31, BP_sw33, BP_sw34, BP_sw60, BP_sw61, BP_sw63, BP_sw64)
  Dim ReflPrtObjsArray: ReflPrtObjsArray = Array(BM_Parts, BM_Overlay, BM_LFlip, BM_LFlipU, BM_RFlip, BM_RFlipU, BM_Bumper_Ring, BM_sw70, BM_sw71, BM_sw73, BM_sw74, BM_sw20, BM_sw21, BM_sw23, BM_sw24, BM_sw30, BM_sw31, BM_sw33, BM_sw34, BM_sw60, BM_sw61, BM_sw63, BM_sw64)
  Dim IBP, BP, v1, v2
  Select Case state
    Case 0: v1 = false: v2 = false  'No reflections
    Case 1: v1 = false: v2 = true   'Only reflect static prims
    Case 2: v1 = true:  v2 = true   'Reflect everything (that matters as defined in ReflObjectsArray)
  End Select

  For Each IBP in ReflObjectsArray
    For Each BP in IBP
      BP.ReflectionEnabled = v1
    Next
  Next

  For Each BP in ReflPrtObjsArray
    BP.ReflectionEnabled = v2
  Next

  BM_Parts.ReflectionEnabled = v2
  BM_Playfield.ReflectionEnabled = False
End Function


'**********************************
'   ZMAT: General Math Functions
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


'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 1000       'X position of punger lane left
Const PLRight = 1060      'X position of punger lane right
Const PLTop = 1225        'Y position of punger lane top
Const PLBottom = 1900       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135*gilvl ' orig was 120 and 70

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    d_w = b_base*(1 - (gBOT(s).z-25)/500)
    If d_w < 30 Then d_w = 30
    ' Handle plunger lane
    If InRect(gBOT(s).x,gBOT(s).y,PLLeft,PLBottom,PLLeft,PLTop,PLRight,PLTop,PLRight,PLBottom) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-PLBottom))
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


'****************************
'   ZRBR: Room Brightness
'****************************

'This code only applies to lightmapped tables. It is here for reference.
'NOTE: Objects bightness will be affected by the Day/Night slider only if their blenddisablelighting property is less than 1.
'      Lightmapped table primitives have their blenddisablelighting equal to 1, therefore we need this SetRoomBrightness sub
'      to handle updating their effective ambient brighness.

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid")

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


'******************************************************
' ZKEY: Key Press Handling
'******************************************************

Sub Table1_KeyDown(ByVal keycode) '***What to do when a button is pressed***

  If keycode = LeftFlipperKey Then
    VR_CabFlipperLeft.X = VR_CabFlipperLeft.X +10
  End If
    If keycode = RightFlipperKey Then
    VR_CabFlipperRight.X = VR_CabFlipperRight.X -10
  End If
  If Keycode = StartGameKey Then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y -5
  End If
  If keycode = AddCreditKey or keycode = 4 then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y -5
  End If

  ' GNMOD for post-it high scores
  if PostItNoteEnabled and EnteringInitials then
    CollectInitials(keycode)
    exit sub
  end if

  If Keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If Keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
  If keycode = PlungerKey Then
    Plunger.pullback
    SoundPlungerPull
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If
  If keycode = LeftTiltKey Then Nudge 90, 1 : SoundNudgeLeft      ' Sets the nudge angle and power
  If keycode = RightTiltKey Then Nudge 270, 1 : SoundNudgeRight   ' ^
  If keycode = CenterTiltKey Then Nudge 0, 1 : SoundNudgeCenter   ' ^
  If keycode = StartGameKey Then SoundStartButton
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)   '***What to do when a button is released***

  If keycode = LeftFlipperKey Then
    VR_CabFlipperLeft.X = VR_CabFlipperLeft.X -10
  End If
    If keycode = RightFlipperKey Then
    VR_CabFlipperRight.X = VR_CabFlipperRight.X +10
  End If
  If Keycode = StartGameKey Then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y +5
  End If
  If keycode = AddCreditKey or keycode = 4 then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y +5
  End If

  if PostItNoteEnabled and EnteringInitials then
    exit sub
  end if

  If Keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
  If Keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
  If keycode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
    VR_Primary_plunger.y = 2109.31
  End If
  If vpmKeyUp(keycode) Then Exit Sub
End Sub


'*********
'Solenoids
'*********

SolCallback(1) = "SolBallRelease"
SolCallback(2) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(3) = "VpmSolSound""10pts"","
SolCallback(4) = "VpmSolSound""100chime"","
SolCallback(5) = "VpmSolSound""1000chime"","
SolCallback(6) = "SolSaucerRelease"
SolCallback(7) = "ResetDropsGreen" 'Green droptargets
SolCallback(8) = "ResetDropsRed" 'Red droptargets
SolCallback(17) = "vpmNudge.SolGameOn"

SolCallback(sLRFlipper) = "SolRFlipper" 'Right Flipper
SolCallback(sLLFlipper) = "SolLFlipper" 'Left Flipper
SolCallback(sURFlipper) = "SolURFlipper"  'Right Flipper
SolCallback(sULFlipper) = "SolULFlipper"  'Left Flipper


Sub SolBallRelease(enabled)
  If enabled Then
    Drain.kick 57, 10
    Controller.Switch(66) = 0
    RandomSoundBallRelease Drain
  End If
End Sub

'*********
' Switches
'*********


Sub sw53f_Hit:vpmTimer.PulseSw 53: LeftUpperBottomRubber.enabled = true : End Sub

Sub sw53e_Hit:vpmTimer.PulseSw 53: RightUpperBottomRubber.enabled = true : End Sub

Sub sw53c_Hit:vpmTimer.PulseSw 53: LeftMidRubber.enabled = true :End Sub

Sub sw53d_Hit:vpmTimer.PulseSw 53: RightMidRubber.enabled = true :End Sub

Sub sw53h_Hit:vpmTimer.PulseSw 53: LeftLowerRubber.enabled = true :End Sub

Sub sw53g_Hit:vpmTimer.PulseSw 53: RightLowerRubber.enabled = true : End Sub

Sub sw53i_Hit:vpmTimer.PulseSw 53: LeftSlingshot.enabled = true : End Sub

Sub sw53i_LC_Hit:vpmTimer.PulseSw 53: LeftSlingshot.enabled = true : End Sub

Sub sw53i_LL_Hit:vpmTimer.PulseSw 53: LeftSlingshot.enabled = true : End Sub

Sub sw53j_Hit:vpmTimer.PulseSw 53: RightSlingshot.enabled = true :End Sub

Sub sw53j_RC_Hit:vpmTimer.PulseSw 53: RightSlingshot.enabled = true :End Sub

Sub sw53j_RL_Hit:vpmTimer.PulseSw 53: RightSlingshot.enabled = true :End Sub

Sub sw53a_Hit:vpmTimer.PulseSw 53: LeftUpperTopRubber.enabled = true : End Sub

Sub sw53b_Hit:vpmTimer.PulseSw 53: RightUpperTopRubber.enabled = true : End Sub

Sub RubberBand_Hit:LeftUpperTopRubber.enabled = true:End Sub
Sub RubberBand001_Hit:RightUpperTopRubber.enabled = true:End Sub
Sub RubberBand002_Hit:RightUpperBottomRubber.enabled = true:End Sub
Sub RubberBand003_Hit:LeftUpperBottomRubber.enabled = true:End Sub
Sub RubberBand004_Hit:LeftLowerRubber.enabled = true:End Sub
Sub RubberBand005_Hit:RightLowerRubber.enabled = true:End Sub
Sub RubberBand_LL_Hit:LeftSlingshot.enabled = true:End Sub
Sub RubberBand_LC_Hit:LeftSlingshot.enabled = true:End Sub
Sub RubberBand_RL_Hit:RightSlingshot.enabled = true:End Sub
Sub RubberBand_RC_Hit:RightSlingshot.enabled = true:End Sub





' Drain & Saucers
Sub Drain_Hit:Controller.Switch(66) = 1:RandomSoundDrain Drain:End Sub

' Rollovers
Sub sw10_Hit:Controller.Switch(10) = 1: AnimateWire BP_sw10, 1 : End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0: AnimateWire BP_sw10, 0 : End Sub

Sub sw11_Hit:Controller.Switch(11) = 1: AnimateWire BP_sw11, 1 : End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0: AnimateWire BP_sw11, 0 : End Sub

Sub sw13_Hit:Controller.Switch(13) = 1: AnimateWire BP_sw13, 1 : End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0: AnimateWire BP_sw13, 0 : End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:AnimateWire BP_sw14, 1 : End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:AnimateWire BP_sw14, 0 : End Sub

Sub sw40_Hit:Controller.Switch(40) = 1:AnimateWire BP_sw40, 1 : End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:AnimateWire BP_sw40, 0 : End Sub

Sub sw50_Hit:Controller.Switch(50) = 1:AnimateWire BP_sw50, 1 : End Sub
Sub sw50_UnHit:Controller.Switch(50) = 0:AnimateWire BP_sw50, 0 : End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:AnimateWire BP_sw43, 1 : End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:AnimateWire BP_sw43, 0 : End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:AnimateWire BP_sw44, 1 : End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:AnimateWire BP_sw44, 0 : End Sub

Sub AnimateWire(group, action) ' Action = 1 - to drop, 0 to raise)
  Dim BP
  If action = 1 Then
    For Each BP in group : BP.transz = -13 : Next
  Else
    For Each BP in group : BP.transz = 0 : Next
  End If
End Sub


' ZSLG: Slingshot Animations
'************************* Animated Rubbers **************************
dim LeftLowerRubberCounter : LeftLowerRubberCounter = 0
sub LeftLowerRubber_Timer
  'BP_Rubber10, BP_Rubber10a, BP_Rubber10b
  Dim BP

  Select Case LeftLowerRubberCounter
    Case 1:
      ' rubber in
      For Each BP in BP_Rubber10 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber10a : bp.Visible = 1 : Next
      For Each BP in BP_Rubber10b : bp.Visible = 0 : Next
    Case 3:
      ' rubber out
      For Each BP in BP_Rubber10 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber10a : bp.Visible = 0 : Next
      For Each BP in BP_Rubber10b : bp.Visible = 1 : Next
    Case Else:
      ' default rubber position
      For Each BP in BP_Rubber10 : bp.Visible = 1 : Next
      For Each BP in BP_Rubber10a : bp.Visible = 0 : Next
      For Each BP in BP_Rubber10b : bp.Visible = 0 : Next
  End Select

  if LeftLowerRubberCounter >= 4 Then
    LeftLowerRubberCounter = 0
    LeftLowerRubber.enabled = false
  Else
    LeftLowerRubberCounter = LeftLowerRubberCounter + 1
  end If
end Sub

dim RightLowerRubberCounter : RightLowerRubberCounter = 0
sub RightLowerRubber_Timer
  'BP_Rubber9, BP_Rubber9a, BP_Rubber9b
  Dim BP

  Select Case RightLowerRubberCounter
    Case 1:
      ' rubber in
      For Each BP in BP_Rubber9 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber9a : bp.Visible = 1 : Next
      For Each BP in BP_Rubber9b : bp.Visible = 0 : Next
    Case 3:
      ' rubber out
      For Each BP in BP_Rubber9 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber9a : bp.Visible = 0 : Next
      For Each BP in BP_Rubber9b : bp.Visible = 1 : Next
    Case Else:
      ' default rubber position
      For Each BP in BP_Rubber9 : bp.Visible = 1 : Next
      For Each BP in BP_Rubber9a : bp.Visible = 0 : Next
      For Each BP in BP_Rubber9b : bp.Visible = 0 : Next
  End Select

  if RightLowerRubberCounter >= 4 Then
    RightLowerRubberCounter = 0
    RightLowerRubber.enabled = false
  Else
    RightLowerRubberCounter = RightLowerRubberCounter + 1
  end If
end Sub

dim LeftMidRubberCounter : LeftMidRubberCounter = 0
sub LeftMidRubber_Timer
  'BP_Rubber8, BP_Rubber8a, BP_Rubber8b
  Dim BP

  Select Case LeftMidRubberCounter
    Case 1:
      ' rubber in
      For Each BP in BP_Rubber8 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber8a : bp.Visible = 1 : Next
      For Each BP in BP_Rubber8b : bp.Visible = 0 : Next
    Case 3:
      ' rubber out
      For Each BP in BP_Rubber8 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber8a : bp.Visible = 0 : Next
      For Each BP in BP_Rubber8b : bp.Visible = 1 : Next
    Case Else:
      ' default rubber position
      For Each BP in BP_Rubber8 : bp.Visible = 1 : Next
      For Each BP in BP_Rubber8a : bp.Visible = 0 : Next
      For Each BP in BP_Rubber8b : bp.Visible = 0 : Next
  End Select

  if LeftMidRubberCounter >= 4 Then
    LeftMidRubberCounter = 0
    LeftMidRubber.enabled = false
  Else
    LeftMidRubberCounter = LeftMidRubberCounter + 1
  end If
end Sub

dim RightMidRubberCounter : RightMidRubberCounter = 0
sub RightMidRubber_Timer
  'BP_Rubber7, BP_Rubber7a, BP_Rubber7b
  Dim BP

  Select Case RightMidRubberCounter
    Case 1:
      ' rubber in
      For Each BP in BP_Rubber7 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber7a : bp.Visible = 1 : Next
      For Each BP in BP_Rubber7b : bp.Visible = 0 : Next
    Case 3:
      ' rubber out
      For Each BP in BP_Rubber7 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber7a : bp.Visible = 0 : Next
      For Each BP in BP_Rubber7b : bp.Visible = 1 : Next
    Case Else:
      ' default rubber position
      For Each BP in BP_Rubber7 : bp.Visible = 1 : Next
      For Each BP in BP_Rubber7a : bp.Visible = 0 : Next
      For Each BP in BP_Rubber7b : bp.Visible = 0 : Next
  End Select

  if RightMidRubberCounter >= 4 Then
    RightMidRubberCounter = 0
    RightMidRubber.enabled = false
  Else
    RightMidRubberCounter = RightMidRubberCounter + 1
  end If
end Sub

dim LeftUpperBottomRubberCounter : LeftUpperBottomRubberCounter = 0
sub LeftUpperBottomRubber_Timer
  'BP_Rubber2, BP_Rubber2a, BP_Rubber2b
  Dim BP

  Select Case LeftUpperBottomRubberCounter
    Case 1:
      ' rubber in
      For Each BP in BP_Rubber2 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber2a : bp.Visible = 1 : Next
      For Each BP in BP_Rubber2b : bp.Visible = 0 : Next
    Case 3:
      ' rubber out
      For Each BP in BP_Rubber2 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber2a : bp.Visible = 0 : Next
      For Each BP in BP_Rubber2b : bp.Visible = 1 : Next
    Case Else:
      ' default rubber position
      For Each BP in BP_Rubber2 : bp.Visible = 1 : Next
      For Each BP in BP_Rubber2a : bp.Visible = 0 : Next
      For Each BP in BP_Rubber2b : bp.Visible = 0 : Next
  End Select

  if LeftUpperBottomRubberCounter >= 4 Then
    LeftUpperBottomRubberCounter = 0
    LeftUpperBottomRubber.enabled = false
  Else
    LeftUpperBottomRubberCounter = LeftUpperBottomRubberCounter + 1
  end If
end Sub

dim LeftUpperTopRubberCounter : LeftUpperTopRubberCounter = 0
sub LeftUpperTopRubber_Timer
  'BP_Rubber2, BP_Rubber2c, BP_Rubber2d
  Dim BP

  Select Case LeftUpperTopRubberCounter
    Case 1:
      ' rubber in
      For Each BP in BP_Rubber2 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber2c : bp.Visible = 1 : Next
      For Each BP in BP_Rubber2d : bp.Visible = 0 : Next
    Case 3:
      ' rubber out
      For Each BP in BP_Rubber2 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber2c : bp.Visible = 0 : Next
      For Each BP in BP_Rubber2d : bp.Visible = 1 : Next
    Case Else:
      ' default rubber position
      For Each BP in BP_Rubber2 : bp.Visible = 1 : Next
      For Each BP in BP_Rubber2c : bp.Visible = 0 : Next
      For Each BP in BP_Rubber2d : bp.Visible = 0 : Next
  End Select

  if LeftUpperTopRubberCounter >= 4 Then
    LeftUpperTopRubberCounter = 0
    LeftUpperTopRubber.enabled = false
  Else
    LeftUpperTopRubberCounter = LeftUpperTopRubberCounter + 1
  end If
end Sub


dim RightUpperBottomRubberCounter : RightUpperBottomRubberCounter = 0
sub RightUpperBottomRubber_Timer
  'BP_Rubber4, BP_Rubber4a, BP_Rubber4b
  Dim BP

  Select Case RightUpperBottomRubberCounter
    Case 1:
      ' rubber in
      For Each BP in BP_Rubber4 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber4a : bp.Visible = 1 : Next
      For Each BP in BP_Rubber4b : bp.Visible = 0 : Next
    Case 3:
      ' rubber out
      For Each BP in BP_Rubber4 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber4a : bp.Visible = 0 : Next
      For Each BP in BP_Rubber4b : bp.Visible = 1 : Next
    Case Else:
      ' default rubber position
      For Each BP in BP_Rubber4 : bp.Visible = 1 : Next
      For Each BP in BP_Rubber4a : bp.Visible = 0 : Next
      For Each BP in BP_Rubber4b : bp.Visible = 0 : Next
  End Select

  if RightUpperBottomRubberCounter >= 4 Then
    RightUpperBottomRubberCounter = 0
    RightUpperBottomRubber.enabled = false
  Else
    RightUpperBottomRubberCounter = RightUpperBottomRubberCounter + 1
  end If
end Sub

dim RightUpperTopRubberCounter : RightUpperTopRubberCounter = 0
sub RightUpperTopRubber_Timer
  'BP_Rubber4, BP_Rubber4c, BP_Rubber4d
  Dim BP

  Select Case RightUpperTopRubberCounter
    Case 1:
      ' rubber in
      For Each BP in BP_Rubber4 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber4c : bp.Visible = 1 : Next
      For Each BP in BP_Rubber4d : bp.Visible = 0 : Next
    Case 3:
      ' rubber out
      For Each BP in BP_Rubber4 : bp.Visible = 0 : Next
      For Each BP in BP_Rubber4c : bp.Visible = 0 : Next
      For Each BP in BP_Rubber4d : bp.Visible = 1 : Next
    Case Else:
      ' default rubber position
      For Each BP in BP_Rubber4 : bp.Visible = 1 : Next
      For Each BP in BP_Rubber4c : bp.Visible = 0 : Next
      For Each BP in BP_Rubber4d : bp.Visible = 0 : Next
  End Select

  if RightUpperTopRubberCounter >= 4 Then
    RightUpperTopRubberCounter = 0
    RightUpperTopRubber.enabled = false
  Else
    RightUpperTopRubberCounter = RightUpperTopRubberCounter + 1
  end If
end Sub

sub InitRubberState
  'configure default state for any animated objects

  dim BP

  'default state for left sling
  For Each BP in BP_RubberLSling : bp.Visible = 1 : Next
  For Each BP in BP_RubberLSlinga : bp.Visible = 0 : Next
  For Each BP in BP_RubberLSlingb : bp.Visible = 0 : Next

  'default state for right sling
  For Each BP in BP_RubberRSling : bp.Visible = 1 : Next
  For Each BP in BP_RubberRSlinga : bp.Visible = 0 : Next
  For Each BP in BP_RubberRSlingb : bp.Visible = 0 : Next

  'default state for left lower rubber
  For Each BP in BP_Rubber10 : bp.Visible = 1 : Next
  For Each BP in BP_Rubber10a : bp.Visible = 0 : Next
  For Each BP in BP_Rubber10b : bp.Visible = 0 : Next

  'default state for right lower rubber
  For Each BP in BP_Rubber9 : bp.Visible = 1 : Next
  For Each BP in BP_Rubber9a : bp.Visible = 0 : Next
  For Each BP in BP_Rubber9b : bp.Visible = 0 : Next

  'default state for left mid rubber
  For Each BP in BP_Rubber8 : bp.Visible = 1 : Next
  For Each BP in BP_Rubber8a : bp.Visible = 0 : Next
  For Each BP in BP_Rubber8b : bp.Visible = 0 : Next

  'default state for right mid rubber
  For Each BP in BP_Rubber7 : bp.Visible = 1 : Next
  For Each BP in BP_Rubber7a : bp.Visible = 0 : Next
  For Each BP in BP_Rubber7b : bp.Visible = 0 : Next

  'defaiult state for left upper rubber
  For Each BP in BP_Rubber2 : bp.Visible = 1 : Next
  For Each BP in BP_Rubber2a : bp.Visible = 0 : Next
  For Each BP in BP_Rubber2b : bp.Visible = 0 : Next
  For Each BP in BP_Rubber2c : bp.Visible = 0 : Next
  For Each BP in BP_Rubber2d : bp.Visible = 0 : Next

  'defaiult state for right upper rubber
  For Each BP in BP_Rubber4 : bp.Visible = 1 : Next
  For Each BP in BP_Rubber4a : bp.Visible = 0 : Next
  For Each BP in BP_Rubber4b : bp.Visible = 0 : Next
  For Each BP in BP_Rubber4c : bp.Visible = 0 : Next
  For Each BP in BP_Rubber4d : bp.Visible = 0 : Next

end sub

'************************* Slings **************************
dim LeftSlingCounter : LeftSlingCounter = 0
sub LeftSlingshot_Timer
  'BP_RubberLSling, BP_RubberLSlinga,BP_RubberLSlingb
  Dim BP

  Select Case LeftSlingCounter
    Case 1:
      ' rubber in
      For Each BP in BP_RubberLSling : bp.Visible = 0 : Next
      For Each BP in BP_RubberLSlinga : bp.Visible = 1 : Next
      For Each BP in BP_RubberLSlingb : bp.Visible = 0 : Next
    Case 3:
      ' rubber out
      For Each BP in BP_RubberLSling : bp.Visible = 0 : Next
      For Each BP in BP_RubberLSlinga : bp.Visible = 0 : Next
      For Each BP in BP_RubberLSlingb : bp.Visible = 1 : Next
    Case Else:
      ' default rubber position
      For Each BP in BP_RubberLSling : bp.Visible = 1 : Next
      For Each BP in BP_RubberLSlinga : bp.Visible = 0 : Next
      For Each BP in BP_RubberLSlingb : bp.Visible = 0 : Next
  End Select

  if LeftSlingCounter >= 4 Then
    LeftSlingCounter = 0
    LeftSlingshot.enabled = false
  Else
    LeftSlingCounter = LeftSlingCounter + 1
  end If
end Sub

dim RightSlingCounter : RightSlingCounter = 0
sub RightSlingshot_Timer
  'BP_RubberRSling, BP_RubberRSlinga,BP_RubberRSlingb
  Dim BP

  Select Case RightSlingCounter
    Case 1:
      ' rubber in
      For Each BP in BP_RubberRSling : bp.Visible = 0 : Next
      For Each BP in BP_RubberRSlinga : bp.Visible = 1 : Next
      For Each BP in BP_RubberRSlingb : bp.Visible = 0 : Next
    Case 3:
      ' rubber out
      For Each BP in BP_RubberRSling : bp.Visible = 0 : Next
      For Each BP in BP_RubberRSlinga : bp.Visible = 0 : Next
      For Each BP in BP_RubberRSlingb : bp.Visible = 1 : Next
    Case Else:
      ' default rubber position
      For Each BP in BP_RubberRSling : bp.Visible = 1 : Next
      For Each BP in BP_RubberRSlinga : bp.Visible = 0 : Next
      For Each BP in BP_RubberRSlingb : bp.Visible = 0 : Next
  End Select

  if RightSlingCounter >= 4 Then
    RightSlingCounter = 0
    RightSlingshot.enabled = false
  Else
    RightSlingCounter = RightSlingCounter + 1
  end If
end Sub

'*******************************************
'  ZBMP: Bumper Setups
'*******************************************

Sub Bumper1_Hit()
  RandomSoundBumperTop Bumper1
  vpmTimer.PulseSw 51

  Dim BP
  For Each BP in BP_Bumper_Socket
    BP.roty=skirtAY(me,Activeball)
    BP.rotx=skirtAX(me,Activeball)
  Next

  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub Bumper1_timer
  Dim BP
  For Each BP in BP_Bumper_Socket
    BP.roty=0
    BP.rotx=0
  Next
  me.timerenabled=0
end sub

Sub Bumper1_Animate
  Dim z: z = Bumper1.CurrentRingOffset

  Dim BP
  For Each BP in BP_Bumper_Ring : BP.transz = z : Next
End Sub


'******************************************************
'     SKIRT ANIMATION FUNCTIONS
'******************************************************
' NOTE: set bumper object timer to around 150-175 in order to be able
' to actually see the animaation, adjust to your liking

'Const PI = 3.1415926
Const SkirtTilt=5   'angle of skirt tilting in degrees

Function SkirtAX(bumper, bumperball)
  skirtAX=cos(skirtA(bumper,bumperball))*(SkirtTilt)    'x component of angle
  if (bumper.y<bumperball.y) then skirtAX=skirtAX*-1    'adjust for ball hit bottom half
End Function

Function SkirtAY(bumper, bumperball)
  skirtAY=sin(skirtA(bumper,bumperball))*(SkirtTilt)    'y component of angle
  if (bumper.x>bumperball.x) then skirtAY=skirtAY*-1    'adjust for ball hit left half
End Function

Function SkirtA(bumper, bumperball)
  dim hitx, hity, dx, dy
  hitx=bumperball.x
  hity=bumperball.y

  dy=Abs(hity-bumper.y)         'y offset ball at hit to center of bumper
  if dy=0 then dy=0.0000001
  dx=Abs(hitx-bumper.x)         'x offset ball at hit to center of bumper
  skirtA=(atn(dx/dy)) '/(PI/180)      'angle in radians to ball from center of Bumper1
End Function


'************************************************************
' ZVUK: VUKs and Kickers
'************************************************************

'******************* VUKs **********************

Dim KickerBall41

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)  'Defines how KickBall works
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

'Upper Left VUK
Sub sw41_Hit                    'Switch associated with Kicker
    set KickerBall41 = activeball
    Controller.Switch(41) = 1
    SoundSaucerLock
End Sub

Sub SolSaucerRelease(Enable)              'Solonoid name associated with kicker.
    If Enable then
    If Controller.Switch(41) <> 0 Then
      KickBall KickerBall41, 170, 10, 5, 10
      SoundSaucerKick 1, sw41
      Controller.Switch(41) = 0
      AnimateKickArm              'Trigger kickarm animation
    End If
  End If
End Sub

' VUK/saucer animation, uses sw41
Sub AnimateKickArm()
  Dim BP
  For Each BP in BP_PkickarmR : BP.rotx = -15 : Next
  VPMTimer.AddTimer 150, "DropKickArm '"
End Sub

Sub DropKickArm()
  Dim BP
  For Each BP in BP_PkickarmR : BP.rotx = 0 : Next
End Sub

'********************** Drop Targets ************************

Sub Sw20_Hit: DTHit 20: TargetBouncer Activeball, 1.5: End Sub
Sub Sw21_Hit: DTHit 21: TargetBouncer Activeball, 1.5: End Sub
Sub Sw23_Hit: DTHit 23: TargetBouncer Activeball, 1.5: End Sub
Sub Sw24_Hit: DTHit 24: TargetBouncer Activeball, 1.5: End Sub
Sub Sw30_Hit: DTHit 30: TargetBouncer Activeball, 1.5: End Sub
Sub Sw31_Hit: DTHit 31: TargetBouncer Activeball, 1.5: End Sub
Sub Sw33_Hit: DTHit 33: TargetBouncer Activeball, 1.5: End Sub
Sub Sw34_Hit: DTHit 34: TargetBouncer Activeball, 1.5: End Sub
Sub Sw60_Hit: DTHit 60: TargetBouncer Activeball, 1.5: End Sub
Sub Sw61_Hit: DTHit 61: TargetBouncer Activeball, 1.5: End Sub
Sub Sw63_Hit: DTHit 63: TargetBouncer Activeball, 1.5: End Sub
Sub Sw64_Hit: DTHit 64: TargetBouncer Activeball, 1.5: End Sub
Sub Sw70_Hit: DTHit 70: TargetBouncer Activeball, 1.5: End Sub
Sub Sw71_Hit: DTHit 71: TargetBouncer Activeball, 1.5: End Sub
Sub Sw73_Hit: DTHit 73: TargetBouncer Activeball, 1.5: End Sub
Sub Sw74_Hit: DTHit 74: TargetBouncer Activeball, 1.5: End Sub

Sub ResetDropsRed(enabled)
  if enabled then
    RandomSoundDropTargetReset BM_sw21
    DTRaise 20
    DTRaise 21
    DTRaise 23
    DTRaise 24
    DTShadows(0).visible = True
    DTShadows(1).visible = True
    DTShadows(2).visible = True
    DTShadows(3).visible = True
  end if
End Sub

Sub ResetDropsGreen(enabled)
  if enabled then
    RandomSoundDropTargetReset BM_sw31
    DTRaise 30
    DTRaise 31
    DTRaise 33
    DTRaise 34
    DTShadows(4).visible = True
    DTShadows(5).visible = True
    DTShadows(6).visible = True
    DTShadows(7).visible = True
  end if
End Sub

Sub ResetDropsYellow(enabled)
  if enabled then
    RandomSoundDropTargetReset BM_sw61
    DTRaise 60
    DTRaise 61
    DTRaise 63
    DTRaise 64
    DTShadows(8).visible = True
    DTShadows(9).visible = True
    DTShadows(10).visible = True
    DTShadows(11).visible = True
  end if
End Sub

Sub ResetDropsBlue(enabled)
  if enabled then
    RandomSoundDropTargetReset BM_sw71
    DTRaise 70
    DTRaise 71
    DTRaise 73
    DTRaise 74
    DTShadows(12).visible = True
    DTShadows(13).visible = True
    DTShadows(14).visible = True
    DTShadows(15).visible = True
  end if
End Sub

Set LampCallback=GetRef("UpdateMultipleLamps")

'Drop Target Reset with lamp
Dim N1,O1,P1,Q1
N1=0:O1=0:P1=0:Q1=0

Sub UpdateMultipleLamps
  N1=Controller.Lamp(17)
  If N1<>O1 Then
    If N1 Then ResetDropsYellow True
  End If
  P1=Controller.Lamp(18)
  If P1<>Q1 Then
    If P1 Then ResetDropsBlue True
  End If

  If VRMode = True Then
    If Controller.Lamp(1) = 0 Then
      VR_BGBIP.visible=0
      VR_BGGO.visible=1
    else
      VR_BGBIP.visible=1
      VR_BGGO.visible=0
    end if

    If Controller.Lamp(2) = 0 Then: VR_BGTilt.visible=0: else: VR_BGTilt.visible=1  'Tilt
    VR_BGSA.visible = l4.state  'Shoot Again
    If Controller.Lamp(3) = 0 Then: VR_BGHS.visible=0: else: VR_BGHS.visible=1  'High Score
  End If
End Sub

Sub AnimateDropTargets
  dim BP, tz, rx, ry

    tz = BM_sw20.transz
  rx = BM_sw20.rotx
  ry = BM_sw20.roty
  For each BP in BP_sw20 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw21.transz
  rx = BM_sw21.rotx
  ry = BM_sw21.roty
  For each BP in BP_sw21 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw23.transz
  rx = BM_sw23.rotx
  ry = BM_sw23.roty
  For each BP in BP_sw23 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw24.transz
  rx = BM_sw24.rotx
  ry = BM_sw24.roty
  For each BP in BP_sw24: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw30.transz
  rx = BM_sw30.rotx
  ry = BM_sw30.roty
  For each BP in BP_sw30: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw31.transz
  rx = BM_sw31.rotx
  ry = BM_sw31.roty
  For each BP in BP_sw31: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw33.transz
  rx = BM_sw33.rotx
  ry = BM_sw33.roty
  For each BP in BP_sw33: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw34.transz
  rx = BM_sw34.rotx
  ry = BM_sw34.roty
  For each BP in BP_sw34: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw60.transz
  rx = BM_sw60.rotx
  ry = BM_sw60.roty
  For each BP in BP_sw60: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw61.transz
  rx = BM_sw61.rotx
  ry = BM_sw61.roty
  For each BP in BP_sw61: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw63.transz
  rx = BM_sw63.rotx
  ry = BM_sw63.roty
  For each BP in BP_sw63: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw64.transz
  rx = BM_sw64.rotx
  ry = BM_sw64.roty
  For each BP in BP_sw64: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw70.transz
  rx = BM_sw70.rotx
  ry = BM_sw70.roty
  For each BP in BP_sw70: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw71.transz
  rx = BM_sw71.rotx
  ry = BM_sw71.roty
  For each BP in BP_sw71: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw73.transz
  rx = BM_sw73.rotx
  ry = BM_sw73.roty
  For each BP in BP_sw73: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw74.transz
  rx = BM_sw74.rotx
  ry = BM_sw74.roty
  For each BP in BP_sw74: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub

'******************************************************
' ZFLP: FLIPPERS
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
      RandomSoundReflipUpLeft LeftFlipper1
    Else
      SoundFlipperUpAttackLeft LeftFlipper1
      RandomSoundFlipperUpLeft LeftFlipper1
    End If
  Else
    LeftFlipper1.RotateToStart
    If LeftFlipper1.currentangle < LeftFlipper1.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper1
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolURFlipper(Enabled)
  If Enabled Then
    RightFlipper1.RotateToEnd
    If RightFlipper1.currentangle < RightFlipper1.endangle + ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper1
    Else
      SoundFlipperUpAttackRight RightFlipper1
      RandomSoundFlipperUpRight RightFlipper1
    End If
  Else
    RightFlipper1.RotateToStart
    If RightFlipper1.currentangle < RightFlipper1.startAngle - 5 Then
      RandomSoundFlipperDownRight RightFlipper1
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub


'Flipper collide subs
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

Sub LeftFlipper1_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper1, ULFCount, parm
  ULF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper1_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper1, URFCount, parm
  URF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub


'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************


'******************************************************
'   ZRDT:  DROP TARGETS by Rothbauerw
'******************************************************
' The Stand Up and Drop Target solutions improve the physics for targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full target animation, switch handling and deflection on hit. For drop targets there is also a slight lift when
' the drop targets raise, bricking, and popping the ball up if it's over the drop target when it raises.
'
' Add a Timers named DTAnim and STAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
' DTAnim.interval = 10
' DTAnim.enabled = True

' Sub DTAnim_Timer
'   DoDTAnim
' DoSTAnim
' End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.
'
' For each stand up target we'll use a vp target, a laid back collidable primitive, and one primitive for visuals and animation.
' The visual primitive should should have it's pivot point centered on the x and y axis and the z should be at or just below the playfield.
' The target should animate backwards using transy.
'
' To create visual target primitives that work with the stand up and drop target code, follow the below instructions:
' (Other methods will work as well, but this is easy for even non-blender users to do)
' 1) Open a new blank table. Delete everything off the table in editor.
' 2) Copy and paste the VP target from your table into this blank table.
' 3) Place the target at x = 0, y = 0  (upper left hand corner) with an orientation of 0 (target facing the front of the table)
' 4) Under the file menu, select Export "OBJ Mesh"
' 5) Go to "https://threejs.org/editor/". Here you can modify the exported obj file. When you export, it exports your target and also
'    the playfield mesh. You need to delete the playfield mesh here. Under the file menu, chose import, and select the obj you exported
'    from VPX. In the right hand panel, find the Playfield object and click on it and delete. Then use the file menu to Export OBJ.
' 6) In VPX, you can add a primitive and use "Import Mesh" to import the exported obj from the previous step. X,Y,Z scale should be 1.
'    The primitive will use the same target texture as the VP target object.
'
' * Note, each target must have a unique switch number. If they share a same number, add 100 to additional target with that number.
' For example, three targets with switch 32 would use 32, 132, 232 for their switch numbers.
' The 100 and 200 will be removed when setting the switch value for the target.

'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

Class DropTarget
  Private m_primary, m_secondary, m_prim, m_sw, m_animate, m_isDropped

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Secondary(): Set Secondary = m_secondary: End Property
  Public Property Let Secondary(input): Set m_secondary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
  Public Property Let IsDropped(input): m_isDropped = input: End Property

  Public default Function init(primary, secondary, prim, sw, animate, isDropped)
    Set m_primary = primary
    Set m_secondary = secondary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate
    m_isDropped = isDropped

    Set Init = Me
  End Function
End Class

'Define a variable for each drop target
Dim DT20,DT21,DT23,DT24,DT30,DT31,DT33,DT34,DT60,DT61,DT63,DT64,DT70,DT71,DT73,DT74

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
' animate:      Array slot for handling the animation instrucitons, set to 0
'           Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target
'   isDropped:      Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.

Set DT20 = (new DropTarget)(sw20, sw20a, BM_sw20, 20, 0, false)
Set DT21 = (new DropTarget)(sw21, sw21a, BM_sw21, 21, 0, false)
Set DT23 = (new DropTarget)(sw23, sw23a, BM_sw23, 23, 0, false)
Set DT24 = (new DropTarget)(sw24, sw24a, BM_sw24, 24, 0, false)
Set DT30 = (new DropTarget)(sw30, sw30a, BM_sw30, 30, 0, false)
Set DT31 = (new DropTarget)(sw31, sw31a, BM_sw31, 31, 0, false)
Set DT33 = (new DropTarget)(sw33, sw33a, BM_sw33, 33, 0, false)
Set DT34 = (new DropTarget)(sw34, sw34a, BM_sw34, 34, 0, false)
Set DT60 = (new DropTarget)(sw60, sw60a, BM_sw60, 60, 0, false)
Set DT61 = (new DropTarget)(sw61, sw61a, BM_sw61, 61, 0, false)
Set DT63 = (new DropTarget)(sw63, sw63a, BM_sw63, 63, 0, false)
Set DT64 = (new DropTarget)(sw64, sw64a, BM_sw64, 64, 0, false)
Set DT70 = (new DropTarget)(sw70, sw70a, BM_sw70, 70, 0, false)
Set DT71 = (new DropTarget)(sw71, sw71a, BM_sw71, 71, 0, false)
Set DT73 = (new DropTarget)(sw73, sw73a, BM_sw73, 73, 0, false)
Set DT74 = (new DropTarget)(sw74, sw74a, BM_sw74, 74, 0, false)


Dim DTArray
DTArray = Array(DT20,DT21,DT23,DT24,DT30,DT31,DT33,DT34,DT60,DT61,DT63,DT64,DT70,DT71,DT73,DT74)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 60 'VP units primitive drops so top of at or below the playfield
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
  DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
  If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
    DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate =  - 1
  DoDTAnim
End Sub

Sub DTDrop(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate = 1
  DoDTAnim
End Sub

Function DTArrayID(switch)
  Dim i
  For i = 0 To UBound(DTArray)
    If DTArray(i).sw = switch Then
      DTArrayID = i
      Exit Function
    End If
  Next
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

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  Dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - Tan(bangle) * aball.x + Tan(rangle2) * dtprim.x) / (Tan(rangle2) - Tan(bangle))
  Yintersect = Tan(rangle2) * Xintersect + (dtprim.y - Tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
  paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

  perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

  If perpvel > 0 And  perpvelafter <= 0 Then
    If DTEnableBrick = 1 And  perpvel > DTBrickVel And DTBrickVel <> 0 And cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
    DTCheckBrick = 4
  Else
    DTCheckBrick = 0
  End If
End Function

Sub DoDTAnim()
  Dim i
  For i = 0 To UBound(DTArray)
    DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
  Next
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
  Dim transz, switchid
  Dim animtime, rangle

  switchid = switch

  Dim ind
  ind = DTArrayID(switchid)

  rangle = prim.rotz * PI / 180

  DTAnimate = animate

  If animate = 0 Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  ElseIf primary.uservalue = 0 Then
    primary.uservalue = GameTime
  End If

  animtime = GameTime - primary.uservalue

  If (animate = 1 Or animate = 4) And animtime < DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    DTAnimate = animate
    Exit Function
  ElseIf (animate = 1 Or animate = 4) And animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 1 'If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0 'updated by rothbauerw to account for edge case
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
  End If

  If animate = 2 Then
    transz = (animtime - DTDropDelay) / DTDropSpeed * DTDropUnits *  - 1
    If prim.transz >  - DTDropUnits  Then
      prim.transz = transz
    End If

    prim.rotx = DTMaxBend * Cos(rangle) / 2
    prim.roty = DTMaxBend * Sin(rangle) / 2

    DTAction switchid

    If prim.transz <= - DTDropUnits Then
      prim.transz =  - DTDropUnits
      secondary.collidable = 0
      DTArray(ind).isDropped = True 'Mark target as dropped
      controller.Switch(Switchid mod 100) = 1
      primary.uservalue = 0
      DTAnimate = 0
      Exit Function
    Else
      DTAnimate = 2
      Exit Function
    End If
  End If

  If animate = 3 And animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
  ElseIf animate = 3 And animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  If animate =  - 1 Then
    transz = (1 - (animtime) / DTDropUpSpeed) * DTDropUnits *  - 1

    If prim.transz =  - DTDropUnits Then
      Dim b
      'Dim gBOT
      'gBOT = GetBalls

      For b = 0 To UBound(gBOT)
        If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < prim.z + DTDropUnits + 25 Then
          gBOT(b).velz = 20
        End If
      Next
    End If

    If prim.transz < 0 Then
      prim.transz = transz
    ElseIf transz > 0 Then
      prim.transz = transz
    End If

    If prim.transz > DTDropUpUnits Then
      DTAnimate =  - 2
      prim.transz = DTDropUpUnits
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = GameTime
    End If
    primary.collidable = 0
    secondary.collidable = 1
    DTArray(ind).isDropped = False 'Mark target as not dropped
    controller.Switch(Switchid mod 100) = 0
  End If

  If animate =  - 2 And animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay) / DTDropSpeed * DTDropUnits *  - 1 + DTDropUpUnits
    If prim.transz < 0 Then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    End If
  End If
End Function

Function DTDropped(switchid)
  Dim ind
  ind = DTArrayID(switchid)

  DTDropped = DTArray(ind).isDropped
End Function

Sub DTAction(switchid)
  ' Turn off the shadow of the target that got hit
  Select Case switchid
    Case 20: DTShadows(0).visible = False
    Case 21: DTShadows(1).visible = False
    Case 23: DTShadows(2).visible = False
    Case 24: DTShadows(3).visible = False

    Case 30: DTShadows(4).visible = False
    Case 31: DTShadows(5).visible = False
    Case 33: DTShadows(6).visible = False
    Case 34: DTShadows(7).visible = False

    Case 60: DTShadows(8).visible = False
    Case 61: DTShadows(9).visible = False
    Case 63: DTShadows(10).visible = False
    Case 64: DTShadows(11).visible = False

    Case 70: DTShadows(12).visible = False
    Case 71: DTShadows(13).visible = False
    Case 73: DTShadows(14).visible = False
    Case 74: DTShadows(15).visible = False
  End Select
End Sub
'******************************************************
'****  END DROP TARGETS
'******************************************************




'*****************
'   Gi Lights
'*****************
Dim gilvl
Dim OldGiState

OldGiState = -1 'start witht he Gi off

Sub GiON
    For each x in aGiLights
        x.State = 1
    Next
  gilvl = 1
  If VRMode = True Then
    VR_BGLit.visible=1
  End If
End Sub

Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
  gilvl = 0
  If VRMode = True Then
    VR_BGLit.visible=0
  End If
End Sub

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then
            GiOff
        Else
            GiOn
        End If
    End If
End Sub

Dim GiIntensity
GiIntensity = 1   'can be used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

' New LUT postit
Function GetHSChar(String, Index)
    Dim ThisChar
    Dim FileName
    ThisChar = Mid(String, Index, 1)
    FileName = "PostIt"
    If ThisChar = " " or ThisChar = "" then
        FileName = FileName & "BL"
    ElseIf ThisChar = "<" then
        FileName = FileName & "LT"
    ElseIf ThisChar = "_" then
        FileName = FileName & "SP"
    Else
        FileName = FileName & ThisChar
    End If
    GetHSChar = FileName
End Function

Sub SetLUTLine(String)
    Dim Index
    Dim xFor
    Index = 1
    LUBack.imagea="PostItNote"
    For xFor = 1 to 40
        Eval("LU" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub HideLUT
    SetLUTLine ""
    LUBack.imagea="PostitBL"
End Sub

'**********************************************************
'     JP's Lamp Fading for VPX and Vpinmame v4.0
' FadingStep used for all kind of lamps
' FlashLevel used for modulated flashers
' LampState keep the real lamp state in a array
'**********************************************************

Dim LampState(200), FadingStep(200), FlashLevel(200)

'InitLamps() ' turn off the lights and flashers and reset them to the default parameters

' vpinmame Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0)) = chgLamp(ii, 1)       'keep the real state in an array
            FadingStep(chgLamp(ii, 0)) = chgLamp(ii, 1)
        Next
    End If
    UpdateLeds
    UpdateLamps
End Sub

Sub UpdateLamps()
    'backdrop lights
    Lamp 1, li1   'game over
    Lampm 2, li2a 'Tilt
    Lamp 2, li2   'Tilt
    Lamp 3, li3   'High Game to date
    Lampm 4, li4a 'Same Player Shoots
    Lampm 4, li4b 'Same Player Shoots
    ' number to match
    ' ball in play

    Lamp 4, li4
    Lamp 5, li5
    Lamp 6, li6
    Lamp 7, li7
    Lamp 8, li8
    Lamp 9, li9
    Lamp 10, li10
    Lamp 11, li11
    Lamp 12, li12
    Lamp 13, li13
    Lamp 14, li14
    Lamp 15, li15
    Lamp 16, li16
    Lamp 19, li19
    Lamp 20, li20
    Lamp 21, li21
    Lamp 22, li22
    Lamp 23, li23
    Lamp 24, li24
    Lamp 25, li25
    Lamp 26, li26
    Lamp 27, li27
    Lamp 28, li28
    Lamp 29, li29
    Lamp 30, li30
    Lamp 31, li31
    Lamp 32, li32
    Lamp 33, li33
    Lamp 35, li35
    Lamp 36, li36
End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 10
    LampTimer.Enabled = 1
    For x = 0 to 200
        FadingStep(x) = 0
        FlashLevel(x) = 0
    Next
End Sub

Sub SetLamp(nr, value) ' 0 is off, 1 is on
    FadingStep(nr) = abs(value)
End Sub

' Lights: used for VPX standard lights, the fading is handled by VPX itself, they are here to be able to make them work together with the flashers

Sub Lamp(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.state = 1:FadingStep(nr) = -1
        Case 0:object.state = 0:FadingStep(nr) = -1
    End Select
End Sub

Sub Lampm(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingStep(nr)
        Case 1:object.state = 1
        Case 0:object.state = 0
    End Select
End Sub

' Flashers:  0 starts the fading until it is off

Sub Flash(nr, object)
    Dim tmp
    Select Case FadingStep(nr)
        Case 1:Object.IntensityScale = 1:FadingStep(nr) = -1
        Case 0
            tmp = Object.IntensityScale * 0.85 - 0.01
            If tmp > 0 Then
                Object.IntensityScale = tmp
            Else
                Object.IntensityScale = 0
                FadingStep(nr) = -1
            End If
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the fading state
    Dim tmp
    Select Case FadingStep(nr)
        Case 1:Object.IntensityScale = 1
        Case 0
            tmp = Object.IntensityScale * 0.85 - 0.01
            If tmp > 0 Then
                Object.IntensityScale = tmp
            Else
                Object.IntensityScale = 0
            End If
    End Select
End Sub

' Desktop Objects: Reels & texts

' Reels - 4 steps fading
Sub Reel(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1:FadingStep(nr) = -1
        Case 0:object.SetValue 2:FadingStep(nr) = 2
        Case 2:object.SetValue 3:FadingStep(nr) = 3
        Case 3:object.SetValue 0:FadingStep(nr) = -1
    End Select
End Sub

Sub Reelm(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1
        Case 0:object.SetValue 2
        Case 2:object.SetValue 3
        Case 3:object.SetValue 0
    End Select
End Sub

' Reels non fading
Sub NfReel(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1:FadingStep(nr) = -1
        Case 0:object.SetValue 0:FadingStep(nr) = -1
    End Select
End Sub

Sub NfReelm(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1
        Case 0:object.SetValue 0
    End Select
End Sub

'Texts

Sub Text(nr, object, message)
    Select Case FadingStep(nr)
        Case 1:object.Text = message:FadingStep(nr) = -1
        Case 0:object.Text = "":FadingStep(nr) = -1
    End Select
End Sub

Sub Textm(nr, object, message)
    Select Case FadingStep(nr)
        Case 1:object.Text = message
        Case 0:object.Text = ""
    End Select
End Sub

' Modulated Subs for the WPC tables

Sub SetModLamp(nr, level)
    FlashLevel(nr) = level / 150 'lights & flashers
End Sub

Sub LampMod(nr, object)          ' modulated lights used as flashers
    Object.IntensityScale = FlashLevel(nr)
    Object.State = 1             'in case it was off
End Sub

Sub FlashMod(nr, object)         'sets the flashlevel from the SolModCallback
    Object.IntensityScale = FlashLevel(nr)
End Sub

'Walls, flashers, ramps and Primitives used as 4 step fading images
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingStep(nr)
        Case 1:object.image = a:FadingStep(nr) = -1
        Case 0:object.image = b:FadingStep(nr) = 2
        Case 2:object.image = c:FadingStep(nr) = 3
        Case 3:object.image = d:FadingStep(nr) = -1
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingStep(nr)
        Case 1:object.image = a
        Case 0:object.image = b
        Case 2:object.image = c
        Case 3:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingStep(nr)
        Case 1:object.image = a:FadingStep(nr) = -1
        Case 0:object.image = b:FadingStep(nr) = -1
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingStep(nr)
        Case 1:object.image = a
        Case 0:object.image = b
    End Select
End Sub

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(32)
Dim Patterns(11)
Dim Patterns2(11)
Dim PatternsVal(11)
Dim Patterns2Val(11)

'Gottliebs
Patterns(0) = 0    'empty
Patterns(1) = 63   '0
Patterns(2) = 6    '1
Patterns(3) = 91   '2
Patterns(4) = 79   '3
Patterns(5) = 102  '4
Patterns(6) = 109  '5
Patterns(7) = 124  '6
Patterns(8) = 7    '7
Patterns(9) = 127  '8
Patterns(10) = 103 '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 768 '134  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 239 '9

'Assign 6-digit output to reels
Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5

Set Digits(6) = b0
Set Digits(7) = b1
Set Digits(8) = b2
Set Digits(9) = b3
Set Digits(10) = b4
Set Digits(11) = b5

Set Digits(12) = c0
Set Digits(13) = c1
Set Digits(14) = c2
Set Digits(15) = c3
Set Digits(16) = c4
Set Digits(17) = c5

Set Digits(18) = d0
Set Digits(19) = d1
Set Digits(20) = d2
Set Digits(21) = d3
Set Digits(22) = d4
Set Digits(23) = d5

Set Digits(24) = e0
Set Digits(25) = e1

Set Digits(26) = f0
Set Digits(27) = f1

Sub UpdateLeds
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat, prevGameState
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)

    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1)
            stat = chgLED(ii, 2)

      if PostItNoteEnabled then
        For jj = 0 to 10
          If stat = Patterns(jj) OR stat = Patterns2(jj) then
            Digits(chgLED(ii, 0)).SetValue jj
            ' if jj=0, that is "blank"/off -- we'll store jj-1 to correspond with decimal value

            'if chgLED(ii, 0) > 23 then debug.print chgLED(ii, 0) & " " & jj
            '25 is bip , 27 is credits
            if chgLED(ii, 0) = 25 then
              prevGameState = isGameActive
              if jj-1 > 0 then isGameActive = True else isGameActive = False
              if prevGameState and not isGameActive then GameEnded
              if not prevGameState and isGameActive then GameStarted
            end if

            if chgLED(ii, 0) >= 0  and chgLED(ii, 0) <=5  Then PlayerReels(0,chgLED(ii, 0)) = jj-1
            if chgLED(ii, 0) >= 6  and chgLED(ii, 0) <=11 Then PlayerReels(1,chgLED(ii, 0)-6) = jj-1
            if chgLED(ii, 0) >= 12 and chgLED(ii, 0) <=17 Then PlayerReels(2,chgLED(ii, 0)-12) = jj-1
            if chgLED(ii, 0) >= 18 and chgLED(ii, 0) <=23 Then PlayerReels(3,chgLED(ii, 0)-18) = jj-1
          End If
        Next
      end if
        Next
    CalculatePlayerScores
    End If
End Sub

'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. and, special scripting
'
' TriggerLF and RF should now be 27 vp units from the flippers. In addition, 3 degrees should be added to the end angle
' when creating these triggers.
'
' RF.ReProcessBalls Activeball and LF.ReProcessBalls Activeball must be added the flipper_collide subs.
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
' | EOS Torque         | 0.4            | 0.4                   | 0.375                  | 0.375              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity
Dim ULF : Set ULF = New FlipperPolarity
Dim URF : Set URF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
Sub InitPolarity()
   dim x, a : a = Array(LF, RF, ULF, URF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger


        x.AddPt "Polarity", 0, 0, 0
        x.AddPt "Polarity", 1, 0.05, - 2.7
        x.AddPt "Polarity", 2, 0.16, - 2.7
        x.AddPt "Polarity", 3, 0.22, - 0
        x.AddPt "Polarity", 4, 0.25, - 0
        x.AddPt "Polarity", 5, 0.3, - 1
        x.AddPt "Polarity", 6, 0.4, - 2
        x.AddPt "Polarity", 7, 0.5, - 2.7
        x.AddPt "Polarity", 8, 0.65, - 1.8
        x.AddPt "Polarity", 9, 0.75, - 0.5
        x.AddPt "Polarity", 10, 0.81, - 0.5
        x.AddPt "Polarity", 11, 0.88, 0
        x.AddPt "Polarity", 12, 1.3, 0

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
  ULF.SetObjects "LF", LeftFlipper1, TriggerULF
    URF.SetObjects "RF", RightFlipper1, TriggerURF
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

Sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
  FlipperTricks LeftFlipper1, ULFPress, ULFCount, ULFEndAngle, ULFState
  FlipperTricks RightFlipper1, URFPress, URFCount, URFEndAngle, URFState
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


Dim FCCDamping: FCCDamping = 0.35

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
  'Check left flipper
    If LeftFlipper1.currentangle = ULFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, LeftFlipper) OR FlipperTrigger(ball2.x, ball2.y, LeftFlipper) Then DoDamping = true
    End If
    'Check right flipper
    If RightFlipper1.currentangle = URFEndAngle Then
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

Dim LFPress, RFPress, ULFPress, URFPress, LFCount, RFCount, ULFCount, URFCount
Dim LFState, RFState, ULFState, URFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle, ULFEndAngle, URFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
ULFState = 1
URFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
Const EOSAnew = 1.5
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
Const EOSReturn = 0.045  'late 70's to mid 80's

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle
ULFEndAngle = Leftflipper1.endangle
URFEndAngle = Rightflipper1.endangle

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
Const LiveDistanceMax = 112 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)
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
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7-1.0

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

'****************************************************************
'****  END VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'****************************************************************




'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

' Put all the Post and Pin objects in dPosts collection. Make sure dPosts fires hit events.
Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

' This collection contains the bottom sling posts. They are not in the dPosts collection so that the TargetBouncer is not applied to them, but they should still have dampening applied
' If you experience airballs with posts or targets, consider adding them to this collection
Sub NoTargetBouncer_Hit
    RubbersD.dampen ActiveBall
End Sub

' Put all the Sleeve objects in dSleeves collection. Make sure dSleeves fires hit events.
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



'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************


'***************************************************************
' ZSHA: Ambient ball shadows
'***************************************************************

' For dynamic ball shadows, Check the "Raytraced ball shadows" box for the specific light.
' Also make sure the light's z position is around 25 (mid ball)

'Ambient (Room light source)
Const AmbientBSFactor = 0.9    '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(5)

'Initialization
BSInit

Sub BSInit()
  Dim iii
  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 3 + iii / 1000
    objBallShadow(iii).visible = 0
  Next
End Sub


Sub BSUpdate
  Dim s: For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow

    'Primitive shadow on playfield, flasher shadow in ramps
    '** If on main and upper pf
    If gBOT(s).Z > 20 And gBOT(s).Z < 30 Then
      objBallShadow(s).visible = 1
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
      'objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25

    '** No shadow if ball is off the main playfield (this may need to be adjusted per table)
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub


'******************************************************
' ZBRL: BALL ROLLING AND DROP SOUNDS
'******************************************************
'
' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

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
  Dim b', BOT
' BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ < -1 and gBOT(b).z < 55 and gBOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz > -7 Then
          RandomSoundBallBouncePlayfieldSoft gBOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard gBOT(b)
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
' Tutorial videos by Apophis
' Audio : Adding Fleep Part 1         https://youtu.be/rG35JVHxtx4?si=zdN9W4cZWEyXbOz_
' Audio : Adding Fleep Part 2         https://youtu.be/dk110pWMxGo?si=2iGMImXXZ0SFKVCh
' Audio : Adding Fleep Part 3         https://youtu.be/ESXWGJZY_EI?si=6D20E2nUM-xAw7xy


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

  If tmp > 0 Then
    AudioFade = CSng(tmp ^ 10)
  Else
    AudioFade = CSng( - (( - tmp) ^ 10) )
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

  If tmp > 0 Then
    AudioPan = CSng(tmp ^ 10)
  Else
    AudioPan = CSng( - (( - tmp) ^ 10) )
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

Sub Aprons_Hit (idx)
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
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************




'Gottlieb System 1 games with 3 tones only (or chimes).
'These are: Charlie's Angels, Cleopatra, Close Encounters, Count Down, Dragon, Joker Poker, Pinball Pool, Sinbad, Solar Ride.
'For the other System 1 games use the file called "Gottlieb System 1 with multi-mode sound.txt"
'Added by Inkochnito

Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "System 1 (3 tones) - DIP switches"
        .AddFrame 205, 0, 190, "Maximum credits", &H00030000, Array("5 credits", 0, "8 credits", &H00020000, "10 credits", &H00010000, "15 credits", &H00030000) 'dip 17&18
        .AddFrame 0, 0, 190, "Coin chute control", &H00040000, Array("seperate", 0, "same", &H00040000)                                                          'dip 19
        .AddFrame 0, 46, 190, "Game mode", &H00000400, Array("extra ball", 0, "replay", &H00000400)                                                              'dip 11
        .AddFrame 0, 92, 190, "High game to date awards", &H00200000, Array("no award", 0, "3 replays", &H00200000)                                              'dip 22
        .AddFrame 0, 138, 190, "Balls per game", &H00000100, Array("5 balls", 0, "3 balls", &H00000100)                                                          'dip 9
        .AddFrame 0, 184, 190, "Tilt effect", &H00000800, Array("game over", 0, "ball in play only", &H00000800)                                                 'dip 12
        .AddChk 205, 80, 190, Array("Match feature", &H00000200)                                                                                                 'dip 10
        .AddChk 205, 95, 190, Array("Credits displayed", &H00001000)                                                                                             'dip 13
        .AddChk 205, 110, 190, Array("Play credit button tune", &H00002000)                                                                                      'dip 14
        .AddChk 205, 125, 190, Array("Play tones when scoring", &H00080000)                                                                                      'dip 20
        .AddChk 205, 140, 190, Array("Play coin switch tune", &H00400000)                                                                                        'dip 23
        .AddChk 205, 155, 190, Array("High game to date displayed", &H00100000)                                                                                  'dip 21
        .AddLabel 50, 240, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")



' Flipper bake maps and shadows

Dim max_angle_RF, min_angle_RF, mid_angle_RF            ' min and max angles from the right flipper
max_angle_RF = RightFlipper.StartAngle              ' right flipper down angle
min_angle_RF = RightFlipper.EndAngle              ' right flipper up angle
mid_angle_RF = (max_angle_RF-min_angle_RF)/2 + min_angle_RF   ' right flipper bake map switch point angle

Dim max_angle_LF, min_angle_LF, mid_angle_LF            ' min and max angles from the left flipper
max_angle_LF = LeftFlipper.StartAngle             ' left flipper down angle
min_angle_LF = LeftFlipper.EndAngle               ' left flipper up angle
mid_angle_LF = (max_angle_LF-min_angle_LF)/2 + min_angle_LF   ' left flipper bake map switch point angle

Dim max_angle_R1F, min_angle_R1F, mid_angle_R1F           ' min and max angles from the right flipper
max_angle_R1F = RightFlipper1.StartAngle            ' right flipper down angle
min_angle_R1F = RightFlipper1.EndAngle              ' right flipper up angle
mid_angle_R1F = (max_angle_R1F-min_angle_R1F)/2 + min_angle_R1F ' right flipper bake map switch point angle

Dim max_angle_L1F, min_angle_L1F, mid_angle_L1F           ' min and max angles from the left flipper
max_angle_L1F = LeftFlipper1.StartAngle             ' left flipper down angle
min_angle_L1F = LeftFlipper1.EndAngle             ' left flipper up angle
mid_angle_L1F = (max_angle_L1F-min_angle_L1F)/2 + min_angle_L1F ' left flipper bake map switch point angle

Sub LeftFlipper_Animate
  Dim BP
  Dim a: a = LeftFlipper.CurrentAngle       ' store flipper angle in a
  FlipperLSh.RotZ = a
  For Each BP in BP_LFlip
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a > mid_angle_LF
  Next
  For Each BP in BP_LFlipU
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle_LF
  Next
End Sub

Sub LeftFlipper1_Animate
  Dim BP
  Dim a: a = LeftFlipper1.CurrentAngle        ' store flipper angle in a
  FlipperLSh001.rotz = a              ' set flipper shadow angle
  For Each BP in BP_LFlip1
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a > mid_angle_L1F
  Next
  For Each BP in BP_LFlip1U
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle_L1F
  Next
End Sub

Sub RightFlipper_Animate
  Dim BP
  Dim a: a = RightFlipper.CurrentAngle        ' store flipper angle in a
  FlipperRSh.RotZ = a
  For Each BP in BP_RFlip
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle_RF
  Next
  For Each BP in BP_RFlipU
    BP.visible = a > mid_angle_RF
    BP.RotZ = a                 ' rotate the maps
  Next
End Sub

Sub RightFlipper1_Animate
  Dim BP
  Dim a: a = RightFlipper1.CurrentAngle       ' store flipper angle in a
  FlipperRSh001.rotz = a              ' set flipper shadow angle
  For Each BP in BP_RFlip1
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle_R1F
  Next
  For Each BP in BP_RFlip1U
    BP.visible = a > mid_angle_R1F
    BP.RotZ = a                 ' rotate the maps
  Next
End Sub

'************************
' Gates Brick maps to follow the VP gates
'************************

Sub Gate_Animate
  Dim a: a = Gate.CurrentAngle
  Dim BP : For Each BP in BP_Pgate: BP.RotX = -a - 25 : Next
End Sub


'   ZVBG: VR Backglass
'******************************
' Setup Backglass
'******************************

Dim xoff,yoff1, yoff2, yoff3, yoff4, yoff5,zoff,xrot,zscale, xcen,ycen

Sub setup_backglass()

  xoff = -20
  yoff1 = 139 ' this is where you adjust the forward/backward position for player 1 score
  yoff2 = 139 ' this is where you adjust the forward/backward position for player 2 score
  yoff3 = 139 ' this is where you adjust the forward/backward position for player 3 score
  yoff4 = 139 ' this is where you adjust the forward/backward position for player 4 score
  yoff5 = 139 ' this is where you adjust the forward/backward position for credits and ball in play
  zoff = 699
  xrot = -90

  center_digits()

end sub


Sub center_digits()
  Dim ix, xx, yy, yfact, xfact, xobj

  zscale = 0.0000001

  xcen = (130 /2) - (92 / 2)
  ycen = (780 /2 ) + (203 /2)

  for ix = 0 to 5
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff1

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 6 to 11
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff2

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 12 to 17
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff3

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 18 to 23
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff4

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 24 to 27
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff5

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

end sub

'   ZVRR: VR Room
'********************************************
'              Display Output
'********************************************


Dim VRDigits(28)
VRDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6,led1x7,led1x8)
VRDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6,led2x7,led2x8)
VRDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6,led3x7,led3x8)
VRDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6,led4x7,led4x8)
VRDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6,led5x7,led5x8)
VRDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6,led6x7,led6x8)

VRDigits(6) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6,led8x7,led8x8)
VRDigits(7) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6,led9x7,led9x8)
VRDigits(8) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6,led10x7,led10x8)
VRDigits(9) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6,led11x7,led11x8)
VRDigits(10) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6,led12x7,led12x8)
VRDigits(11) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6,led13x7,led13x8)

VRDigits(12) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006,LED1x007,LED1x008)
VRDigits(13) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106,LED1x107,LED1x108)
VRDigits(14) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206,LED1x207,LED1x208)
VRDigits(15) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306,LED1x307,LED1x308)
VRDigits(16) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406,LED1x407,LED1x408)
VRDigits(17) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506,LED1x507,LED1x508)

VRDigits(18) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006,led2x007,led2x008)
VRDigits(19) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106,led2x107,led2x108)
VRDigits(20) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206,led2x207,led2x208)
VRDigits(21) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306,led2x307,led2x308)
VRDigits(22) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406,led2x407,led2x408)
VRDigits(23) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506,led2x507,led2x508)

'credit -- Ball In Play
VRDigits(24) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306,LEDax307,LEDax308)
VRDigits(25) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406,LEDbx407,LEDbx408)
VRDigits(26) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506,LEDcx507,LEDcx508)
VRDigits(27) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606,LEDdx607,LEDdx608)

dim DisplayColor
DisplayColor =  RGB(0,190,255)

Sub VRDisplayTimer
  on error resume next
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x, prevGameState
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
      num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)

      For Each obj In VRDigits(num)
'                   If chg And 1 Then obj.visible=stat And 1    'if you use the object color for off; turn the display object visible to not visible on the playfield, and uncomment this line out.
           If chg And 1 Then FadeDisplay obj, stat And 1
                   chg=chg\2 : stat=stat\2
            Next

      'if using post-it note high scores, track LEDs for scorekeeping
      if PostItNoteEnabled then
        For jj = 0 to 10
          If chgLED(ii, 2) = Patterns(jj) OR chgLED(ii, 2) = Patterns2(jj) then
            '25 is bip , 27 is credits
            if chgLED(ii, 0) = 25 then
              'debug.print "BIP: " & jj-1
              prevGameState = isGameActive
              '>0 balls in play
              if jj-1 > 0 then isGameActive = True else isGameActive = False
              if prevGameState and not isGameActive then GameEnded
              if not prevGameState and isGameActive then GameStarted
            end if

            if chgLED(ii, 0) >= 0  and chgLED(ii, 0) <=5  Then PlayerReels(0,chgLED(ii, 0)) = jj-1
            if chgLED(ii, 0) >= 6  and chgLED(ii, 0) <=11 Then PlayerReels(1,chgLED(ii, 0)-6) = jj-1
            if chgLED(ii, 0) >= 12 and chgLED(ii, 0) <=17 Then PlayerReels(2,chgLED(ii, 0)-12) = jj-1
            if chgLED(ii, 0) >= 18 and chgLED(ii, 0) <=23 Then PlayerReels(3,chgLED(ii, 0)-18) = jj-1
          End If
        Next
      end if
        Next
    CalculatePlayerScores
    End If
 End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
    Object.Opacity = 12
  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 6
  End If
End Sub


Sub InitDigits()
  dim tmp, x, obj
  for x = 0 to uBound(VRDigits)
    if IsArray(VRDigits(x) ) then
      For each obj in VRDigits(x)
        obj.height = obj.height + 18
        FadeDisplay obj, 0
      next
    end If
  Next
End Sub

InitDigits

Sub TimerPlunger_Timer

  If VR_Primary_plunger.Y < 2244.31 then
      VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
  VR_Primary_plunger.Y = 2109.31 + (5* Plunger.Position) -20
End Sub

' *** Score Handling

dim isGameActive : isGameActive = False
dim PlayerScores(3)  ' 4 players
dim PlayerReels(3,5) ' 6 reels per player
dim HighScore

Sub HighScoreSave
  'don't save HS if post-it is disabled, as there would be no names / etc
  if not PostItNoteEnabled then exit Sub

  dim i

  for i=1 to 5
    SaveValue cGameName, "HighScore" & i, HSScore(i)
    SaveValue cGameName, "Name" & i, HSName(i)
  next
end Sub

Sub HighScoreLoad
  dim i, val, name

  for i=1 to 5
    val = LoadValue(cGameName,"HighScore" & i)
    name = LoadValue(cGameName,"Name" & i)

    If isNumeric(val) Then
      HSScore(i)=Cdbl(val)
      HSName(i) = name
    end if
  next
end Sub

Sub ResetPlayerReelsAndScores
  dim ii, jj
  For ii = 0 To UBound(PlayerReels)
    For jj = 0 To UBound(PlayerReels,2)
      PlayerReels(ii,jj) = -1 'default to an "off" value
    Next
  Next
  For ii = 0 To UBound(PlayerReels)
    PlayerScores(ii) = 0
  Next
End Sub

function isReelOff(player)
  'determines if entire reel for a player is off (-1)
  Dim ii,jj
  Dim isOff: isOff = True

  For jj = 0 To UBound(PlayerReels,2) ' loop each reel for player
    if PlayerReels(player,jj) > -1 then isOff = False : exit For
  Next

  isReelOff = isOff
end Function

Sub CalculatePlayerScores
  Dim ii,jj, PrevScore

  'if game is not active, exit -- this prevents looking at high scores/calculations when not in game
  if not isGameActive then exit Sub

  For ii = 0 To UBound(PlayerReels) 'loop each player
    if not isReelOff(ii) then
      PrevScore = PlayerScores(ii)
      PlayerScores(ii) = 0

      For jj = 0 To UBound(PlayerReels,2) ' loop each reel for each player
        if PlayerReels(ii,jj) > 0 then
          PlayerScores(ii) = PlayerScores(ii) + (PlayerReels(ii,jj) * (10^(UBound(PlayerReels,2)-jj)))
        end if
      Next

      'determine if rollover
      'if PrevScore is > than current score, points rolled over
      'optionally can add qualifier like prevscore > 960,000 or thereabouts to gate rollover
      if PrevScore > PlayerScores(ii) and PrevScore > 950000 then
        PlayerScores(ii) = PlayerScores(ii) + (10^6)
        'reset the reels (score may be wrong momentarily, but will get corrected on update)
        'For jj = 0 To UBound(PlayerReels,2) ' loop each reel
        ' PlayerReels(ii,jj) = 0
        'Next
      end if
    end if
  Next

  'debug: show player 1 scores on post-it
  if DEBUG_POSTIT_SCORES then SetHSLine 3, "DBG " & PlayerScores(0)
End Sub

Sub PrintScores(printReels)
  Dim ii,jj
  For ii = 0 To UBound(PlayerScores)
    debug.Print "P" & ii+1 & ": " & PlayerScores(ii)
    if printReels then
      For jj = 0 To UBound(PlayerReels,2)
        debug.Print "  R" & jj & ": " & PlayerReels(ii,jj)
      Next
    end if
  Next
End Sub

sub GameEnded
  'do any cleanup/recalc after game has ended
  debug.print "Game has ended"
  PrintScores true

  'if not using Post-it note, exit, as rest of code
  'handles timer/feedback/display of HS on post-it
  if not PostItNoteEnabled then exit Sub
  CheckHighScore
  HighScoreTimer.interval=500
  HighScoreTimer.enabled=true
end Sub

sub GameStarted
  'do any cleanup/recalc at beginning of game
  debug.print "Game has started"

  'if game was reset (right flipper + game Start) make sure we're not capturing initials
  if (EnteringInitials = 1) then
    debug.print "Game reset with high score detected -- cancelling high score input"
    EnteringInitials = 0 ' take us out of entering initials mode
    HighScoreTimer_Timer ' trigger timer so we move off capture intials
  end if

  ResetPlayerReelsAndScores
  PrintScores true
end Sub

' *** End Score Handling

' ============================================================================================
' GNMOD - Multiple High Score Display and Collection
' ============================================================================================

dim ScoreChecker
dim CheckAllScores
dim sortscores(4)
dim sortplayers(4)

sub CheckHighScore
  'debug.print "CheckHighScore"
  Dim playertops
  dim i,si,sj
  dim stemp
  dim stempplayers

  for i=1 to 4
    sortscores(i)=0
    sortplayers(i)=0
  next

  playertops=0
  for i = 1 to 4
    sortscores(i)= PlayerScores(i-1)  'Score(i)
    sortplayers(i)=i
    'debug.print "sort scores " & i & " " & sortscores(i)
  next

  for si = 1 to 4
    for sj = 1 to 4-1
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

  'for i = 1 to 4
  ' debug.print sortscores(i)
  'next

  ScoreChecker=4
  CheckAllScores=1
  NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)

  HighScoreSave
end sub

'Dim InProgress

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
HSScore(1) = 499999
HSScore(2) = 388888
HSScore(3) = 277777
HSScore(4) = 161666
HSScore(5) = 80085

HSName(1) = "VPW"
HSName(2) = "M78"
HSName(3) = "BLW"
HSName(4) = "DGR"
HSName(5) = "FE"

Sub HighScoreTimer_Timer

  if not PostItNoteEnabled then
    debug.print "Post-it not enabled - disabling timer"
    HighScoreTimer.enabled=false
    exit Sub
  end If

  if EnteringInitials then
    if HSTimerCount = 1 then
      SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
      HSTimerCount = 2
    else
      SetHSLine 3, InitialString
      HSTimerCount = 1
    end if
  elseif isGameActive then
    SetHSLine 1, "HIGH SCORE1"
    SetHSLine 2, HSScore(1)
    SetHSLine 3, HSName(1)
    HSTimerCount = 5  ' set so the highest score will show after the game is over
    HighScoreTimer.enabled=false
  elseif CheckAllScores then
    NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)

  else
    'debug.print "cycling hs"
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

Function GetHSCharHS(String, Index)
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
  GetHSCharHS = FileName
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
    Eval("HS"&xfor).image = GetHSCharHS(String, Index)
    Index = Index + 1
  next

End Sub

Sub NewHighScore(NewScore, PlayNum)
  'debug.print "NewHighScore " & NewScore & " " & PlayNum
  if NewScore > HSScore(5) then
    'debug.print "newscore>5"
    HighScoreTimer.interval = 500
    HSTimerCount = 1
    AlphaStringPos = 1    ' start with first character "A"
    EnteringInitials = 1  ' intercept the control keys while entering initials
    InitialString = ""    ' initials entered so far, initialize to empty
    SetHSLine 1, "PLAYER "+FormatNumber(PlayNum,0)
    SetHSLine 2, "ENTER NAME"
    SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
    HSNewHigh = NewScore
    'For xx=1 to HighScoreReward
    ' AddSpecial
    'next
  End if
  ScoreChecker=ScoreChecker-1
  if ScoreChecker=0 then
    CheckAllScores=0
  end if
End Sub

Sub CollectInitials(keycode)
  dim i

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
    PlaySound "DropTarget_Down"
  elseif keycode = RightFlipperKey Then
    ' advance to next character
    AlphaStringPos = AlphaStringPos + 1
    if AlphaStringPos > len(AlphaString) or (AlphaStringPos = len(AlphaString) and InitialString = "") then
      ' Skip the backspace if there are no characters to backspace over
      AlphaStringPos = 1
    end if
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTarget_Down"
  elseif keycode = StartGameKey or keycode = PlungerKey Then
    SelectedChar = MID(AlphaString, AlphaStringPos, 1)
    if SelectedChar = "_" then
      InitialString = InitialString & " "
      PlaySound("10pts")
    elseif SelectedChar = "<" then
      InitialString = MID(InitialString, 1, len(InitialString) - 1)
      if len(InitialString) = 0 then
        ' If there are no more characters to back over, don't leave the < displayed
        AlphaStringPos = 1
      end if
      PlaySound("100chime")
    else
      InitialString = InitialString & SelectedChar
      PlaySound("10pts")
    end if
    if len(InitialString) < 3 then
      SetHSLine 3, InitialString & SelectedChar
    End If
  End If
  'debug.print "Length of initial string: " & len(InitialString)
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
        SetHSLine 3, InitialString
        HighScoreTimer_Timer
        'HighScoreTimer.interval = 2000
        PlaySound("1000chime")
        HighScoreSave
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
