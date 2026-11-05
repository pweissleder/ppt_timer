# Slide Timer for PowerPoint on Mac

A countdown timer that lives **on your slide**. Start, pause and resume by clicking the timer’s play/pause button or digits. At zero it shows **TIME'S UP** and plays your chosen macOS sound three times through the selected **media output**.

The timer keeps counting when you change slides, switch apps, or press Esc to leave the slide show briefly. Each timer keeps its own running, paused, ready or finished state while the deck stays open.

![Slide Timer showing 05:00 with SET, plus, minus and play/pause buttons](Timer-Preview.png)

## Features

- Native seven-segment PowerPoint shapes; no special fonts needed.
- Choose and preview an end sound during creation, including a No sound option. Each timer saves its own choice.
- Existing play/pause, SET presets, +/− one minute, reset and sound controls.
- Independent deadlines and pause states for different timer slides.
- Short updates scheduled on the macOS main run loop, so a countdown does not hold a macro open and block navigation.
- End sound follows **System Settings > Sound > Output** and media volume, including an external display, Bluetooth device or audio interface selected there.
- Durations from 1 second to 23:59:59, with an hours field when needed.
- Offline: no network, AppleScript, ActiveX, add-in or third-party runtime.

## Requirements

- Desktop **PowerPoint for Mac**, Microsoft 365 / Office 2019 or later, 64-bit, VBA7.
- Presentation saved as **.pptm**, with macros enabled.
- Both **SlideTimer** (standard module) and **SlideTimerEvents** (class module) imported into that presentation.

PowerPoint for Windows, the web, iPad/iPhone, Keynote, PDFs and exported videos are not supported. The module uses native macOS libraries. PowerPoint's event and animation behavior still needs checking on the Office version you present with; see [TESTING.md](TESTING.md).

## Setup

1. Download [`SlideTimer.bas`](SlideTimer.bas) and [`SlideTimerEvents.cls`](SlideTimerEvents.cls).
2. Open your presentation and save it as **PowerPoint Macro-Enabled Presentation (.pptm)**.
3. Open **Tools > Macro > Visual Basic Editor**. In **View > Project Explorer**, select your presentation's project.
4. Choose **File > Import File…** for **both** files. Import each only once. You should see the standard module **SlideTimer** and class module **SlideTimerEvents**.
   - If importing is unavailable, use **Insert > Module** and paste [`SlideTimer-Paste.txt`](SlideTimer-Paste.txt). Then use **Insert > Class Module**, set its **(Name)** in the Properties window to **SlideTimerEvents**, and paste [`SlideTimerEvents-Paste.txt`](SlideTimerEvents-Paste.txt). Do not paste the class into a standard module.
5. In PowerPoint's **Normal** view, select a slide and run **`Timer_Insert`** from **Tools > Macro > Macros…**. Enter a duration such as `5` or `1:30`, then choose the end sound by number or name. Type `p 2` to preview Glass; enter `2` to keep it. Choose `0` for no sound. Canceling either creation prompt leaves the slide without a new timer.
6. Save, start the show with PowerPoint’s usual **Slide Show** commands, then click the timer’s **play/pause** button or **digits**.

Start shows normally, including after reopening the deck. The timer controls initialize their background updates and close-event handling automatically; you do not need to run a setup macro before each show.

When upgrading from 1.2.0–1.2.3, run **`Timer_Initialize` once in Normal view** and save. This removes the old hidden click-start shapes and their animation steps from your existing slides.

## Using the timer

| Control | What it does |
| --- | --- |
| **Play/pause** or **digits** | Start, pause or resume. After TIME'S UP, starts a fresh countdown. |
| **SET** | Next preset: 1, 2, 3, 5, 10, 15, 20, 30, 45, 60 minutes, then 1. Resets to Ready. |
| **+** / **−** | Add/subtract one minute from the duration and remaining time; minimum 1 second. Works while running or paused. |
| **RESET** | Restore the configured duration and wait for play. |
| **SOUND ON / OFF** | Enable/mute this timer's end alert. Muting cancels pending repeats. |

Spacebar, arrow keys and presenter remotes keep their ordinary slide-navigation and animation behavior. Use the timer’s own controls to start or resume it. RESET and SET restore Ready state and wait for play; returning to a running, paused or finished timer preserves that state.

**Navigate, switch apps or press Esc whenever you need to.** A running timer continues, including while the show is closed and the deck is still open. The end alert can therefore play while you are on another slide or in Normal view. Use pause or `Timer_Shutdown` to stop counting.

At zero the digits turn muted red and show **00:00 / TIME'S UP**. Reset, restart or SOUND OFF cancels remaining alert repeats. An alert already playing can finish.

SET uses presets during the show because PowerPoint for Mac can put input dialogs behind the full-screen show. Use `Timer_Configure` in Normal view for an exact duration.

## Macros in Normal view

| Macro | Purpose |
| --- | --- |
| `Timer_Insert` | Add a timer with duration/sound choices, or configure an existing timer |
| `Timer_Present` | Optional shortcut to launch the configured show; normal show commands work too |
| `Timer_Initialize` | Optional one-time upgrade cleanup of obsolete click-start shapes/animations |
| `Timer_Configure` | Type an exact duration for the current timer |
| `Timer_ChooseSound` | Choose/preview the current timer’s end sound in Normal view |
| `Timer_Start` | Start/resume without toggling; useful for a shape's Run macro action |
| `Timer_Shutdown` | Pause timers, cancel repeats and stop native callbacks before editing VBA |
| `Timer_Remove` | Remove the current timer |
| `Timer_RefreshStyle` | Rebuild the current timer, keeping duration, sound, position and width |
| `Timer_CheckCurrent` | Check parts and report the last background error |
| `Timer_TestSound` | Play this timer’s selected sound three times through the current media output; muted timers stay silent |
| `Timer_RunSelfTests` | Run duration, state-transition and clock checks; expect **59 checks passed** |
| `Timer_Help` | Short usage help |

Duration input accepts `5` (5 minutes), `1:30` (90 seconds), `0:10` (10 seconds), or `1:02:03` (1 hour, 2 minutes, 3 seconds). Use whole numbers, in the range `0:01`–`23:59:59`.

## Placing and reusing timers

- Move/resize the timer as **one group**; keep its parts and names intact.
- Duplicate a timer slide, or run `Timer_Insert` on another slide. Each timer has its own duration, sound preference and in-memory countdown state.
- Import both modules into every deck that uses the timer. Copying shapes does not copy VBA.
- Set timer-slide transitions to **None**. The timer forces a redraw once per changed second, and a transition could replay.
- Do not animate the entire timer group to disappear; it must remain visible to operate its controls.

## Sound settings

`Timer_Insert` now asks which sound to use after the duration. The choices are **Ping, Glass, Hero, Submarine, Funk, Basso, Blow, Bottle, Frog, Morse, Pop, Purr, Sosumi and Tink**, plus **0 / None** for a silent timer. Enter a listed number or name. To listen before deciding, enter `p 2` (preview Glass), then enter `2` or `Glass` to save it. Previewing does not save a choice; Cancel keeps an existing timer unchanged or cancels creation of a new timer.

To change a timer later, select its slide in Normal view and run **`Timer_ChooseSound`**. Its sound choice is saved with that timer in the `.pptm`, survives slide duplication and `Timer_RefreshStyle`, and is used by both timer completion and `Timer_TestSound`. Older timers without a choice keep the existing default until you change them. Choosing None mutes the timer; the existing SOUND ON button can re-enable its previous choice (Ping for a newly silent timer).

The end sound uses Apple's built-in **`/usr/bin/afplay`** in a separate process. VBA passes the executable and sound path as a UTF-8 argument vector using `posix_spawn`; no shell, AppleScript or Objective-C sound calls are used. The audio player uses the normal media output and volume, rather than the separate system-alert output. Audio decoding and playback happen outside PowerPoint's process.

Choose the desired device in **System Settings > Sound > Output**, set media volume, and run `Timer_TestSound` on your actual presentation setup. Changing the output after the sound has loaded is supported; each repeat uses the default output again. A failed player launch is reported immediately; a nonzero player exit is recorded for `Timer_CheckCurrent`. Playback failures do not stop other countdowns. There is no fallback to the separate system alert route.

Constants near the top of `SlideTimer.bas`:

| Constant | Default | Meaning |
| --- | --- | --- |
| `PT_ALERT_FILE` | `/System/Library/Sounds/Ping.aiff` | Fallback for older timers without a saved sound choice and tests on a slide with no timer |
| `PT_ALERT_REPEATS` | `3` | Number of end alerts |
| `PT_ALERT_GAP` | `1.8` | Minimum seconds between alerts; the next repeat also waits until playback finishes |

The SET presets are in `NextPresetDuration`. A missing/unplayable sound produces a player error instead of silently playing through a different output. Automatic repeats wait for the current player to finish, so longer choices such as Funk are not cut off. A new manual test or preview replaces current playback. `Timer_Shutdown` terminates and reaps that child. The child inherits PowerPoint’s sandbox; restrictive organization or sandbox policies can block playback, which is reported without restoring the crashing sound bridge.

## How it works

- Shape tags store configured duration, sound choice and mute preference. Countdown state lives in memory per timer while its presentation stays open.
- `CACurrentMediaTime` gives a monotonic deadline. Remaining time is recalculated from that deadline after a delayed update, so app switches and slide navigation do not reset or extend it.
- A `CFRunLoopTimer` runs a brief update every 100 ms on PowerPoint's main thread. It is registered in common run-loop modes; there is no `DoEvents` countdown loop.
- The event class handles presentation closing and invalidates native callbacks **before the presentation’s VBA project closes**. It does not handle slide-show clicks. Timer controls attach these close handlers when used.
- Changed segments are updated each second. A visible timer slide is redrawn with `GotoSlide …, msoFalse`, using the deck's actual slide index; hidden timer updates never navigate back to it. A black/white/paused show is not forced back into its running view.
- The native callback catches errors and never opens a dialog. `Timer_CheckCurrent` reports a background failure.

API references: [Apple run-loop timers](https://developer.apple.com/library/archive/documentation/Cocoa/Conceptual/Multithreading/RunLoopManagement/RunLoopManagement.html), [Apple subprocess lifecycle guidance](https://developer.apple.com/documentation/appkit/managing-ongoing-background-processes-in-your-mac), [PresentationBeforeClose](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.presentationbeforeclose).

## Lifecycle and limitations

- **Run `Timer_Shutdown` before editing/removing either VBA module, pressing VBE Reset/End, or using breakpoints.** A scheduled native callback points into VBA code, so changing that code while it is scheduled is unsafe. Normal presentation closing uses `PresentationBeforeClose` to stop it automatically.
- Closing the deck ends its timer session. Reopening or restarting PowerPoint does not restore running deadlines. Saved shape appearances may still show their last display; initialization restores the in-memory Ready state when a timer is first observed.
- PowerPoint must remain running and responsive. If macOS or a modal operation delays updates, the next update catches up to the deadline; the alert can be delayed until then.
- Do not let the Mac sleep during a countdown.
- Use one slide show at a time and one timer per slide. Multiple timers within that deck can count independently.

## Troubleshooting

**An old timer slide still consumes an extra click.** Exit the show, run `Timer_Initialize` once in Normal view, then save. This deletes only the old tagged click-start shapes and effects. Start the timer with its play/pause button or digits.

**Play does nothing.** Enable macros in this deck and use Slide Show view. For Presenter View, the direct shape buttons are on the audience display.

**No sound or the wrong output.** Check SOUND ON and the selected **media output and volume**, then run `Timer_TestSound`. The separate alert-output preference is no longer used. If playback fails, run `Timer_CheckCurrent` for the last background error. A spawn error means the player could not launch; an afplay exit status means it launched and could not finish playing. Check that the selected output is connected and permitted by your organization’s policy.

**User-defined type not defined at `SlideTimerEvents`.** Import the class module too, or create a Class Module named exactly `SlideTimerEvents` and paste the class-only text file.

**Missing timer component or broken control.** Run `Timer_RefreshStyle` in Normal view. Play/pause and digits use `Timer_Toggle`; RESET uses `Timer_Reset`; SET uses `Timer_Configure`; +/− use `Timer_AddMinute` / `Timer_SubtractMinute`; SOUND uses `Timer_ToggleSound`. A custom start-only shape can use `Timer_Start`.

**Macros disappeared.** Saving as `.pptx` strips VBA. Import both modules again and save as `.pptm`.

## Updating from 1.1.x

1. **If replacing 1.2.0 or 1.2.1 after a sound failure/crash, quit and reopen PowerPoint first, and do not run its old sound test or shutdown macro.** The old shutdown also calls the unsafe sound bridge. For a working newer version, run `Timer_Shutdown` before editing. For 1.1.x, pause the countdown and exit the show.
2. In the Visual Basic Editor, remove the old **SlideTimer** standard module (decline exporting if you already have a backup). Also remove the old **SlideTimerEvents** class if present; this version removes its slide-show click handlers.
3. Import both updated files: **SlideTimer.bas** and **SlideTimerEvents.cls**.
4. Run **`Timer_Initialize`** once in Normal view. It removes obsolete hidden click-start shapes/animations; timer groups, duration/sound settings and user animations stay.
5. Save as `.pptm`, start the show normally, and click the timer’s play/pause button or digits. Use `Timer_RefreshStyle` only for an old/non-flat or damaged timer.

## Verification and browser preview

Run `Timer_RunSelfTests` in PowerPoint (expect 59 checks), `Timer_TestSound`, and the [presentation test matrix](TESTING.md). Validate the sound list without playing anything with `python3 tools/check_native_runtime.py --catalog-only`. The macOS binding smoke check is `python3 tools/check_native_runtime.py`; add `--play` to hear one alert. It checks native APIs and distribution parity, not VBA compilation or PowerPoint behavior.

[`START-HERE.html`](START-HERE.html) is a browser mock-up with a synthesized chime. It demonstrates the controls; it cannot validate PowerPoint event handling, routing or navigation.

## Changelog

- **1.2.4:** Remove remote/next-click start and its animation step. Start with the timer’s own controls, with no setup macro required per show. Optional one-time cleanup removes old start cues; sound choices and background state remain.
- **1.2.3:** Per-timer sound chooser during creation, preview with `p 2`, No sound option and `Timer_ChooseSound` for existing timers. Save choices in shape tags and use them for completion/tests; let long sounds finish before repeating.
- **1.2.2:** Remove the crashing Objective-C/VBA sound bridge entirely. Play Ping through a separate afplay process, track exit errors and clean up the owned child.
- **1.2.1:** Attempted Hero-loader repair; superseded by 1.2.2 after an Objective-C/VBA crash report. Do not use its sound path.
- **1.2.0:** End sound follows media output; remote/next-click start cue; scheduled updates replace the blocking loop; per-timer state survives navigation, app switches and temporary show exit. Adds the SlideTimerEvents class and shutdown lifecycle.
- **1.1.9:** Pausing ends the macro; play resumes remaining time.
- **1.1.8:** In-show SET presets; release the show after the finish alert.
- **1.1.7:** Hero sound, three repeats.
- **1.1.6:** AudioToolbox alert instead of VBA Beep.
- **1.1.5:** Redraw every second.
- **1.1.3:** Cached renderer.
- **1.1.2:** Flat LCD layout and style migration.

## License

[MIT](LICENSE) © 2026 Pascal Weissleder
