##   ____           _                ____
##  / ___|__ _  ___| |_ _   _ ___   / ___|__ _ _ __  _   _  ___  _ __
## | |   / _` |/ __| __| | | / __| | |   / _` | '_ \| | | |/ _ \| '_ \
## | |__| (_| | (__| |_| |_| \__ \ | |__| (_| | | | | |_| | (_) | | | |
##  \____\__,_|\___|\__|\__,_|___/  \____\__,_|_| |_|\__, |\___/|_| |_|
##                                                   |___/
##           ___ ___  _  _ _____ ___ _  _ _   _ ___ ___
##          / __/ _ \| \| |_   _|_ _| \| | | | | __|   \
##         | (_| (_) | .` | | |  | || .` | |_| | _|| |) |
##          \___\___/|_|\_| |_| |___|_|\_|\___/|___|___/
##
## A P-ROC Project by Eric Priepke, Copyright 2012-2013
## Built on the PyProcGame Framework from Adam Preble and Gerry Stellenberg
## Original Cactus Canyon software by Matt Coriale
##
## This is a heavily modified copy of the trough code from pyprocgame - I added a bunch of "print"s to
## Help me understand what it was doing, and re-worked a bunch of it to make it make more sense
## to me - mostly in the 'checking swtiches' section.  And made changes to deal with balls
## bouncing back in to the trough after launch

#from procgame.game import Mode
import ep

class Trough(ep.EP_Mode):
    """Manages trough by providing the following functionality:

         - Keeps track of the number of balls in play
         - Keeps track of the number of balls in the trough
         - Launches one or more balls on request and calls a launch_callback when complete, if one exists.
         - Auto-launches balls while ball save is active (if linked to a ball save object
         - Identifies when balls drain and calls a registered drain_callback, if one exists.
         - Maintains a count of balls locked in playfield lock features (if externally incremented) and adjusts the count of number of balls in play appropriately.  This will help the drain_callback distinguish between a ball ending or simply a multiball ball draining.

     Parameters:

         'game': Parent game object.
         'position_switchnames': List of switchnames for each ball position in the trough.
         'eject_switchname': Name of switch in the ball position the feeds the shooter lane.
         'eject_coilname': Name of coil used to put a ball into the shooter lane.
         'early_save_switchnames': List of switches that will initiate a ball save before the draining ball reaches the trough (ie. Outlanes).
         'shooter_lane_switchname': Name of the switch in the shooter lane.  This is checked before a new ball is ejected.
         'drain_callback': Optional - Name of method to be called when a ball drains (and isn't saved).
     """
    def __init__(self, game, position_switchnames, eject_switchname, eject_coilname,
                 early_save_switchnames, shooter_lane_switchname, drain_callback=None):
        super(Trough, self).__init__(game, 90)
        self.myID = "Trough"
        self.position_switchnames = position_switchnames
        self.eject_switchname = eject_switchname
        self.eject_coilname = eject_coilname
        self.shooter_lane_switchname = shooter_lane_switchname
        self.drain_callback = drain_callback
        self.eject_sw_count = 0
        self.troughStrength = self.game.user_settings['Machine (Standard)']['Trough Eject Strength']

        # Install early ball_save switch handlers.
        for switch in early_save_switchnames:
            self.add_switch_handler(name=switch, event_type='active', delay=None, handler=self.early_save_switch_handler)

        # Reset variables
        self.num_balls_in_play = 0
        self.num_balls_locked = 0
        self.num_balls_to_launch = 0
        self.num_balls_to_stealth_launch = 0
        self.launch_in_progress = False

        # set externally if used
        self.launch_callback = None
        self.balls_to_autoplunge = 0

        self.ignore_next_drain = False
        self.num_balls_to_save = 1
        self.ball_save_callback = None
        self.ball_save_active = False
        self.ball_save_lamp = self.game.lamps.shootAgain
        self.ball_save_begin = 0
        self.ball_save_multiple_saves = False
        self.ball_save_timer = 0
        self.ball_save_hold = 0
        self.default_time = self.game.user_settings['Gameplay (Feature)']['Default Ball Save Time']
        self.lane_busy = False

    def tilted(self):
        pass

    # specific to CC - deliniating the shooter lane as delayed start switch
    def sw_shooterLane_inactive(self,sw):
        if self.ball_save_begin > 0:
            self.delay(delay=1,handler=self.ball_save_delayed_start_handler)

    #self.debug()

    ## New Method - Only calling a switch count if the 4th spot activates for some reason.
    def sw_troughBallFour_active(self,sw):
        if self.ignore_next_drain:
            self.ignore_next_drain = False
            #print "Ignoring this drain - Early Save Switch Activated"
            pass
        else:
            self.position_switch_handler(sw)

    def debug(self):
        self.game.set_status(str(self.num_balls_in_play) + "," + str(self.num_balls_locked))
        self.delay(name='launch', event_type=None, delay=1.0,handler=self.debug)

    def early_save_switch_handler(self, sw):
        if self.ball_save_active:
            # Only do an early ball save if a ball is ready to be launched.
            # Otherwise, let the trough switches take care of it.
            if self.game.switches[self.eject_switchname].is_active():
                self.save_ball()
                # and set a flag to ignore the next ball in the trough
                self.ignore_next_drain = True

    def mode_stopped(self):
        self.cancel_delayed('check_switches')
        self.disable_ball_save()

    # Switches will change states a lot as balls roll down the trough.
    # So don't go through all of the logic every time.  Keep resetting a
    # delay function when switches change state.  When they're all settled,
    # the delay will call the real handler (check_switches).
    def position_switch_handler(self, sw):
        self.cancel_delayed('check_switches')
        self.delay(name='check_switches', event_type=None, delay=1, handler=self.check_switches)

    def check_switches(self):
        #print "CHECKING SWITCHES - Balls in play: " + str(self.num_balls_in_play)
        if self.eject_sw_count > 1:
            #print "Passing check switches - eject count more than 1"
            pass
        # If we're currently launching balls - skip the check - we'll check at the end of the launch cycle
        elif self.launch_in_progress:
            #print "Passing check - Launch in progress"
            pass
        # If there's a ball currently in play
        elif self.num_balls_in_play > 0:
            # how many balls should the machine have
            num_current_machine_balls = self.game.num_balls_total
            # how many balls in in the trough now
            counted_balls_in_trough = self.num_balls()
            #print "Check Switches - launch status: " + str(self.launch_in_progress)
            #  Ball saver on situations
            if self.ball_save_active:
                #print "BALL SAVE IS ACTIVE"
                #print "TROUGH SHOULD SAVE " + str(self.num_balls_to_save) + " BALLS"
                # if the balls in play + the balls in the trough is higher than the total, we should save one
                if counted_balls_in_trough + self.num_balls_in_play > num_current_machine_balls:
                    #print "Check Switches wants to save a ball"
                    self.save_ball()
                else:
                    #print "Ball save is on, but balls in play + balls in trough line up"
                    return 'ignore'
            # if the ball save is NOT active, then we process it as a drain
            else:
                #print "BALL SAVE IS NOT ACTIVE"
                # The ball should end if all of the balls are in the trough.
                if counted_balls_in_trough == num_current_machine_balls:
                    self.num_balls_in_play = 0
                    if self.drain_callback:
                        #print "THE TROUGH IS FULL, BALL SAVE INACTIVE, ENDING BALL"
                        self.drain_callback()
                # otherwise there's thinking to do
                else:
                    #print "BALL DRAINED"
                    if self.drain_callback:
                        # Write off the drained balls
                        self.num_balls_in_play = num_current_machine_balls - counted_balls_in_trough
                        # call a drain - unless showdown or ambush is starting
                        if not self.game.showdown.startup and not self.game.ambush.startup:
                            self.drain_callback()
                        #print "BALLS NOW IN PLAY: " + str(self.num_balls_in_play)
        # if there aren't any balls in play
        else:
            if self.launch_in_progress:
                #print "WHAT THE - NO BALLS IN PLAY, and we're launching - try again?"
                # experiental condition for lanny's case where I don't think the balls settled
                if self.num_balls() == 4 or self.game.switches.troughEject.is_active():
                    #print "It fell back in, try again"
                    self.cancel_delayed("Bounce_Delay")
                    self.common_launch_code()

    # Count the number of balls in the trough by counting active trough switches.
    def num_balls(self):
        """Returns the number of balls in the trough."""
        ball_count = 0
        for switch in self.position_switchnames:
            if self.game.switches[switch].is_active():
                ball_count += 1
                #print "Active trough switch: " + str(switch)
        if self.game.switches.troughEject.is_active() and not self.game.fakePinProc:
            #print "There's a ball stacked up in the way of the eject opto"
            ball_count += 1
        #print "balls counted: " + str(ball_count)
        return ball_count

    def is_full(self):
        #print "Checking if trough is full"
        return self.num_balls() == self.game.num_balls_total

    # Either initiate a new launch or add another ball to the count of balls
    # being launched.  Make sure to keep a separate count for stealth launches
    # that should not increase num_balls_in_play.
    def launch_balls(self, num, callback=None, stealth=False):
        """Launches balls into play.

              'num': Number of balls to be launched.
              If ball launches are still pending from a previous request,
              this number will be added to the previously requested number.

              'callback': If specified, the callback will be called once
              all of the requested balls have been launched.

              'stealth': Set to true if the balls being launched should NOT
              be added to the number of balls in play.  For instance, if
              a ball is being locked on the playfield, and a new ball is
              being launched to keep only 1 active ball in play,
              stealth should be used.
          """
        #print "I SHOULD LAUNCH A BALL NOW"
        self.num_balls_to_launch += num
        if stealth:
            self.num_balls_to_stealth_launch += num
        if not self.launch_in_progress:
            self.launch_in_progress = True
            #print "Launch status: " + str(self.launch_in_progress)
            self.common_launch_code()

    # This is the part of the ball launch code that repeats for multiple launches.
    def common_launch_code(self):
        # Only kick out another ball if the last ball is gone from the
        # shooter lane.
        #print "Launch action loop"
        # set the trough eject switch count to zero
        # This is the switch a ball passes going in/out of the trough
        # Counting hits on it allows for detecting a ball bouncing back in
        self.eject_sw_count = 0

        # If the shooter lane is clear, huck a ball
        if self.game.switches[self.shooter_lane_switchname].is_inactive():
            self.game.coils[self.eject_coilname].pulse(self.troughStrength)
            # go to a hold pattern to wait for the shooter lane
            # if after 2 seconds the shooter lane hasn't been hit
            # and the eject switch hasn't triggered, we'll assume the launch worked
            #if not self.game.fakePinProc:
                #print "Trough - scheduling the Bounce_Delay"
            self.delay("Bounce_Delay",delay=3,handler=self.finish_launch)
            # if we are under fakepinproc, proceed immediately to ball in play
            #else:
                #print "Fakepinproc - Finishing Launch"
                #self.finish_launch()

        # Otherwise, wait 1 second before trying again.
        else:
            #print "Shooter lane busy - reschedule"
            self.delay(name='launch', event_type=None, delay=1.0,
                handler=self.common_launch_code)

    def finish_launch(self):
        #print "Finishing Launch"
        # set the eject hits to zero
        self.eject_sw_count = 0
        # tick down the balls to launch
        self.num_balls_to_launch -= 1
        # Set the lane as busy to hold additional launches
        self.lane_busy = True
        # Start a safety net delay to clear the lane busy in case nothing else works
        self.cancel_delayed("SafetyNet")
        self.delay(name = "SafetyNet",delay=15,handler=self.clear_lane_busy)

        #print "BALL LAUNCHED - left to launch: " +str(self.num_balls_to_launch)
        # Only increment num_balls_in_play if there are no more
        # stealth launches to complete.
        if self.num_balls_to_stealth_launch > 0:
            self.num_balls_to_stealth_launch -= 1
        else:
            self.num_balls_in_play += 1
        #print "IN PLAY: " + str(self.num_balls_in_play)
        # If more balls need to be launched, run the wait for clear loop
        if self.num_balls_to_launch > 0:
            #print "More balls to launch: " + str(self.num_balls_to_launch) + " - Waiting for clear"
            self.wait_for_clear()
        # If all the requested balls are launched, go to the cleanup check
        else:
            self.launch_in_progress = False
            if not self.game.fakePinProc:
                self.delay(delay=.5,handler=self.post_launch_check)
            else:
                #print "Fakepinproc - post launch check skipped"
                pass

    ## New routine to have more control over when to launch sucessive balls
    def wait_for_clear(self):
        if self.game.fakePinProc:
            self.common_launch_code()
        else:
            # a looping cycle to wait for the shooter lane to clear
            if self.lane_busy:
                self.delay(name="waiting",delay=1,handler=self.wait_for_clear)
            else:
                # delay the call back to common launch code to allow the bowl to clear
                self.delay(name="waiting",delay=2,handler=self.common_launch_code)

    def sw_shooterLane_inactive_for_4s(self,sw):
        # if the shooter lane opens for 5 seconds, and the busy flag is on, shut if off
        if self.lane_busy:
            #print "Shooter lane inactive is clearing the busy flag"
            self.lane_busy = False
            self.cancel_delayed("SafetyNet")

    def sw_skillBowl_active(self,sw):
        # if the skill bowl switch is hit, and the busy flag is on, turn it off.
        if self.lane_busy:
            #print "Skill bowl is clearing the busy flag"
            self.lane_busy = False
            self.cancel_delayed("SafetyNet")

    def clear_lane_busy(self):
        #print "Safety net is clearing the lane busy"
        self.lane_busy = False

    def post_launch_check(self):
        #print "Running Post Launch Check"
        # count the balls in the trough, and math out how many should be there
        actuallyInTrough = self.num_balls()
        shouldBeInTrough = self.game.num_balls_total - self.num_balls_in_play
        # if they line up - awesome
        if shouldBeInTrough == actuallyInTrough:
            #print "Everything Adds up at the end of the launch."
            pass
        # If we're short, there's extra balls on the playfield.  Sucks, but such is life.
        elif shouldBeInTrough > actuallyInTrough:
            #print "There aren't as many balls in the trough as there should be"
            pass
            #print "Ball in play: " + str(self.num_balls_in_play) + " Counted: " + str(actuallyInTrough)
        # the other option is we have too many balls in the trough
        # in that case, fix things up by stealth launching the difference
        else:
            #print "There are more balls in the trough than there should be"
            #print "Balls in play: " + str(self.num_balls_in_play) + " Counted: " + str(actuallyInTrough)
            #print "Stealth launch to fix that"
            num = actuallyInTrough - shouldBeInTrough
            #print "Launching: " + str(num)
            self.balls_to_autoplunge = num
            self.launch_balls(num,stealth=True)

    def sw_shooterLane_active_for_300ms(self,sw):
        #print "SOLID LAUNCH, GOOD TO GO"
        # if we're ejecting - process the launch
        if self.launch_in_progress:
            # kill the fallback loop
            self.cancel_delayed("Bounce_Delay")
            self.finish_launch()

    def sw_troughEject_active(self,sw):
        # tick up the count with each switch hit
        self.eject_sw_count += 1
        #print "Trough Eject switch count: " + str(self.eject_sw_count)
        # if we go to more than 1, the ball came back
        if self.eject_sw_count > 1:
            #print "We're over on eject count - ball fell back in"
            self.cancel_delayed('Bounce_Delay')
            # retry the ball launch
            self.delay("Retry",delay=1,handler=self.common_launch_code)
        else:
            #print "That's one"
            pass

    ## BALL SAVE BITS

    def start_ball_save_lamp(self):
        """Starts blinking the ball save lamp.  Oftentimes called externally to start blinking the lamp before a ball is plunged."""
        self.ball_save_lamp.schedule(schedule=0xFF00FF00, cycle_seconds=0, now=True)

    def update_lamps(self):
       if self.ball_save_timer > 5:
           self.ball_save_lamp.schedule(schedule=0xFF00FF00, cycle_seconds=0, now=True)
       elif self.ball_save_timer > 2:
           self.ball_save_lamp.schedule(schedule=0x55555555, cycle_seconds=0, now=True)
       else:
           self.ball_save_lamp.disable()

    def disable_ball_save(self):
        """Disables the ball save logic."""
        # kill the active flag
        if self.ball_save_active:
            self.ball_save_active = False
        # set the timer to 0
        self.ball_save_timer = 0
        # turn off the light
        self.ball_save_lamp.disable()

    def start_ball_save(self, num_balls_to_save=1, time=0, now=True, allow_multiple_saves=False):
        """Activates the ball save logic."""

        # If limited flip party mode is on - and player is at/over limit, don't turn on the ball save
        if self.game.user_settings['Gameplay (Feature)']['Party Mode'] == 'Flip Ct' \
        and self.game.show_tracking('Total Flips') >= self.game.user_settings['Gameplay (Feature)']['Party - Flip Count']:
            pass
        else:
            if time == 0:
                time = self.default_time
            #print "Starting Ball Save for " + str(time) + " seconds"
            self.allow_multiple_saves = allow_multiple_saves
            self.num_balls_to_save = num_balls_to_save
            if time > self.ball_save_timer: self.ball_save_timer = time
            self.update_lamps()
            # if starting right away, do that
            if now:
                self.cancel_delayed('ball_save_timer')
                self.delay(name='ball_save_timer', event_type=None, delay=1, handler=self.ball_save_countdown)
                self.ball_save_active = True
            # if not starting right away - stash the timer value and set the flag for delayed start
            else:
                self.ball_save_begin = 1
                self.timer_hold = time

    def ball_save_countdown(self):
        self.ball_save_timer -= 1
        self.update_lamps()
        if self.ball_save_timer >= 1:
            self.delay(name='ball_save_timer', event_type=None, delay=1, handler=self.ball_save_countdown)
        else:
            if self.game.current_player().extra_balls > 0:
                self.game.lamps.shootAgain.enable()
            self.disable_ball_save()

    def ball_save_is_active(self):
        #print "Ball Save active check"
        return self.ball_save_timer > 0

    def ball_save_delayed_start_handler(self, sw):
        if self.ball_save_begin:
            self.ball_save_timer = self.timer_hold
            self.ball_save_begin = 0
            self.update_lamps()
            self.cancel_delayed('ball_save_timer')
            self.delay(name='ball_save_timer', event_type=None, delay=1, handler=self.ball_save_countdown)
            self.ball_save_active = True

    def save_ball(self):
        #print "Saving Ball"
        # log in audits
        self.game.game_data['Feature']['Ball Saves'] += 1

        self.num_balls_to_save -= 1
        #print "Left to save: " + str(self.num_balls_to_save)
        self.ball_save_callback()
        self.balls_to_autoplunge += 1
        self.launch_balls(1,stealth=True)
        if self.num_balls_to_save == 0:
            #print "Last saved ball - Turning off ball save"
            self.disable_ball_save()
