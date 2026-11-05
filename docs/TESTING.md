# PowerPoint acceptance checks for 1.2.4

[← Developer guide](DEVELOPMENT.md) · Run terminal commands from the repository root.

The native smoke check has run on macOS. VBA compilation, Office events,
presenter-remote input and physical audio routing require PowerPoint testing.
The browser preview is not evidence for these checks.

Import both modules into a disposable `.pptm` with three slides: a timer on
slides 1 and 3, and ordinary content on slide 2. Set timer-slide transitions
to None. In the VBA editor, use **Debug > Compile** if available. Run
`Timer_RunSelfTests`: expect **59 checks passed**. Run `Timer_Shutdown`
before editing code, using breakpoints, or pressing VBE Reset/End.

| Scenario | Expected result |
| --- | --- |
| Follow the setup ZIP/START-HERE instructions in a fresh presentation | Both files import, duration/sound prompts work, and the timer starts from its own controls. |
| Expand the setup screenshots in START-HERE, then open the illustrated SETUP guide | All images load with clear captions; the finished-import image shows both modules, and Timer_Initialize appears as upgrade-only. |
| Duplicate an Action slide, insert the timer on the copy, then resize the whole group smaller and larger | The original stays available for instructions; the copy’s digits/buttons remain usable after resizing and moving. |
| Start a fresh show using PowerPoint’s normal command after reopening the deck, without running initialization | Timer stays Ready until its own play/pause or digits are clicked, then counts normally. |
| Press remote next, spacebar or right arrow on a Ready timer slide | Only ordinary navigation/animations advance; no timer starts or resumes and no extra start step exists. |
| Upgrade a deck with old click-start effects, then run `Timer_Initialize` in Normal view and save | Every old tagged start effect and shape is removed. Timer duration, sound choices, placement and user content/animation order remain. |
| Run upgrade cleanup twice; also run it on a deck that never had click-start cues | Cleanup is idempotent and does not add effects, shapes or extra clicks. |
| Create a new timer | No click-start shape or animation is added; its own play/digits controls start it. |
| Create a timer with `Timer_Insert`, choose a duration, then choose sound 2 / Glass | Glass is saved for this timer; `Timer_TestSound` and completion both play Glass. |
| Enter `p 3`, then choose Hero | One Hero preview plays; saving it makes completion/tests use Hero through the selected media output. |
| Choose a name with mixed case or a number from 1–14 | The correct canonical filename is used, including after PowerPoint uppercases saved tags. |
| Choose 0 / None during creation | Timer is created with SOUND OFF; no end alert or sound test plays. SOUND ON re-enables Ping. |
| Preview another sound with `Timer_ChooseSound`, then Cancel | Prior sound choice, mute preference and countdown state are preserved. Preview stops. |
| Cancel the sound choice while creating a new timer | No timer group is created. |
| Choose different sounds on timer slides 1 and 3, save and reopen | Each timer retains its own choice. |
| Duplicate a timer slide and rebuild with `Timer_RefreshStyle` | Sound choice and mute preference survive. |
| Enter an invalid sound number, fractional value or unlisted path | The chooser asks again; it does not save an invalid path or start a countdown. |
| Choose Funk, which is longer than the minimum repeat gap | Each automatic repeat finishes before the next starts. |
| Start 00:30, leave for slide 2 for 10 seconds, then return | About 00:20 remains; no reset and no jump back to the timer slide while away. |
| Start timers on slides 1 and 3 | Both retain independent deadlines, pause states and sound settings. |
| Pause immediately after play, before the first displayed second changes | Return and resume the fractional remaining time; no reset to full duration. |
| Enter an unstarted timer slide by navigating backwards | It stays Ready until the timer’s play/pause button or digits are clicked. |
| Pause, navigate away, wait, return | Same remaining time and PAUSED status; ordinary navigation never resumes it. |
| Press Esc while running, wait 5 seconds, start the show again | Countdown includes those 5 seconds. No extra click is needed to resume counting. |
| Switch apps for several seconds, then return | Clock catches up; no reset or deadline extension. |
| Finish while on another slide or in Normal view | One three-repeat alert sequence; returning shows zero / TIME'S UP without replaying the alert. |
| RESET, SET, +/− and SOUND while running, paused, and finished | Existing control behavior remains; reset/SET wait for play; adjustments preserve running/paused status. |
| Reset during an alert; mute during an alert; restart during an alert | Pending repeats cancel. Current audio may finish. |
| Select external media output; separately leave macOS alert output on MacBook speakers | `Timer_TestSound` and timer completion play over the selected media output. |
| Run `Timer_TestSound` in PowerPoint with the default Ping setting | afplay launches as a separate process; Ping plays three times and PowerPoint stays responsive, with no Objective-C/VBA sound calls. |
| Configure a readable custom AIFF, including a non-ASCII filename | afplay receives its literal UTF-8 path and playback follows media output. |
| Configure a missing or invalid audio file | afplay exits with an error; `Timer_CheckCurrent` reports it and PowerPoint stays usable. |
| Load/test sound, switch media output, then test again | Subsequent sound follows the newly selected media output. |
| Mute media output or disconnect the selected device | No fallback to system alert output; playback failure does not stop other countdowns. `Timer_CheckCurrent` reports any failure. |
| Use a nonconsecutive slide range or custom show | Redraw stays on the actual timer slide; it does not jump to the wrong deck index. |
| Use B/W to blank the show, then restore it | No timer redraw unblanks the show; the deadline continues. |
| Delete a background timer or its slide in Normal view | Remaining timers keep running; no unhandled callback error. |
| Close the deck while a timer is running or an alert is playing | Before-close event cancels native callbacks and terminates/reaps the owned afplay child before VBA unload; PowerPoint stays usable. |
| Attempt to close, then cancel the close | Timer remains paused; use its controls to resume it. |
| Close and reopen the deck, then start the show normally | Clicking play initializes the configured duration and background updates; no startup macro is needed. |
| Duplicate a timer slide | Each timer retains independent settings/state and has no click-start cue. |
| Upgrade an existing flat-layout 1.1.x timer and run `Timer_Initialize` | No cue is added. The timer stays in place with its duration/sound settings; its own controls work. |

For local native checks:

```sh
python3 tools/check_native_runtime.py --catalog-only  # Validate all listed sound files, no playback
python3 tools/check_native_runtime.py
python3 tools/check_native_runtime.py --play  # Optional: hear one sound
```

The catalog-only mode checks every chooser entry with afinfo without playing audio.
The full script checks import/paste parity, real Core Foundation main-thread timer
delivery and cancellation, monotonic time, posix_spawn argument handling,
afplay playback, error exit status, and owned-child termination/reaping. Default
audio tests use zero volume; `--play` makes one Ping audible. It does not compile
or execute the VBA code. Physical output selection still needs listening.
A restricted runner without an audio device will fail the playback check; this
must not be reported as a successful playback test.

When updating a deck containing 1.2.0/1.2.1, quit and reopen PowerPoint before
replacing SlideTimer.bas. Do not run the old sound test or shutdown macro: the
old shutdown also invokes the Objective-C sound bridge. Replace both updated
modules; SlideTimerEvents.cls no longer subscribes to slide-show clicks.
