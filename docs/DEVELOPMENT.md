# Developer guide

[← Back to the overview](../README.md) · [PowerPoint acceptance checks](TESTING.md) · [Changelog](CHANGELOG.md)

The user-facing setup is in [SETUP.md](SETUP.md). This page covers repository layout, implementation and maintenance. The documentation reorganization does not change the 1.2.4 VBA runtime.

## Repository layout

```text
README.md                    Short overview and setup
START-HERE.html              Browser setup guide and interactive demonstration
powerpoint/
  SlideTimer.bas             Importable standard module
  SlideTimerEvents.cls       Importable class module for presentation-close events
  paste/                     Matching code without import-only headers
assets/                      Timer preview and setup screenshots
docs/                        User guides, developer notes, tests and changelog
downloads/                   Ready-to-extract setup ZIP
tools/                       Package builder and native runtime checks
```

The VBA files are the canonical runtime sources. The paste versions must remain identical to the corresponding code after stripping the standard module’s first Attribute line or the class export header before `Option Explicit`. Preserve the CRLF export format for `.bas`/`.cls` and normalized parity with the paste files.

## Build and check the setup package

From the repository root, using Python 3:

```sh
python3 tools/build_setup_package.py
python3 tools/build_setup_package.py --check
python3 tools/check_native_runtime.py --catalog-only
```

The builder creates **downloads/SlideTimer-Mac.zip** from an explicit file list. The archive includes both import files, the paste alternatives, browser guide, preview image, all 13 setup screenshots, documentation, maintenance scripts and license. It excludes Git data, local macOS files, caches and user attachments. Fixed ZIP timestamps keep repeat builds reproducible. The browser guide inside the archive opens directly at the setup steps instead of offering to download the same ZIP again.

Rebuild the archive whenever a packaged source, guide or image changes. `--check` reports a stale package without overwriting it. No extra Python packages are needed.

## Native checks

```sh
python3 tools/check_native_runtime.py --catalog-only
python3 tools/check_native_runtime.py
python3 tools/check_native_runtime.py --play
```

The catalog-only check validates all 14 listed system sound files with afinfo without playing audio. The full check exercises real Core Foundation main-thread timer delivery/cancellation, monotonic time, posix_spawn argument handling, afplay exit errors and owned-child termination/reaping. Default playback tests use zero volume; `--play` makes one Ping audible.

A restricted runner without access to an audio device may fail playback. Do not count that as successful audio verification. These scripts do **not** compile or run the VBA inside PowerPoint. Office event behavior, resize/repaint behavior, input and physical external audio routing need the [PowerPoint test checklist](TESTING.md). `Timer_RunSelfTests` contains **59 checks** to run in PowerPoint.

## Runtime architecture

- The timer is a group of native PowerPoint shapes. Digits use seven-segment geometry; no special font is needed. Its buttons use Action Settings to call public VBA macros.
- Shape tags store configured duration, selected sound and mute preference. In-memory state keeps each timer’s deadline, remaining time and status while the presentation stays open.
- `CACurrentMediaTime` supplies a monotonic deadline. Remaining time is recalculated rather than decremented by tick counts, so delayed callbacks do not extend a countdown.
- A `CFRunLoopTimer` makes brief updates every 100 ms on PowerPoint’s main thread in common run-loop modes. Common modes requires the exported sentinel pointer; an equal CFString is insufficient. There is no blocking DoEvents countdown loop.
- Only changed digits/status are painted. A visible timer slide is refreshed with `GotoSlide …, msoFalse`, using the actual deck slide index. Hidden timers never navigate the show back to themselves. Black/white/paused show states are not forced into running view.
- The class module handles `PresentationBeforeClose` and `PresentationClose`. It cancels native callbacks before the VBA project unloads. It has no slide-click handlers.
- Idle callbacks stop when no countdown, queued alert or audio child needs service. Timer controls attach close events and start callbacks themselves; normal show commands need no initialization macro.
- Old tagged click-start effects/shapes from 1.2.0–1.2.3 are removed in editing view by upgrade cleanup. No new start animations are created.

## Audio

VBA starts **/usr/bin/afplay** with a literal UTF-8 argument vector using `posix_spawn`. No shell, AppleScript, Objective-C sound bridge or third-party runtime is involved. Playback occurs in a separate process using the media output. The child inherits PowerPoint’s sandbox; a denied launch or failed playback is reported rather than bypassed or sent to the system-alert route.

A single owned child PID is polled with nonblocking `waitpid`. Shutdown terminates and reaps it. Automatic repeats wait for current playback to finish, so longer choices such as Funk are not clipped. Manual tests/previews replace current playback. The selected sound is stored per timer, canonicalized when reading PowerPoint’s uppercase tag values, and retained when duplicating slides or rebuilding style.

| Constant | Default | Purpose |
| --- | --- | --- |
| `PT_ALERT_FILE` | `/System/Library/Sounds/Ping.aiff` | Fallback for older timers without a saved choice and tests on a slide with no timer |
| `PT_AUDIO_PLAYER` | `/usr/bin/afplay` | macOS media player executable |
| `PT_ALERT_REPEATS` | `3` | Number of end alerts |
| `PT_ALERT_GAP` | `1.8` | Minimum interval; repeats also wait for the player to finish |

`SoundNames` holds the 14 chooser entries. `NextPresetDuration` holds SET presets. New code should preserve the pointer/UTF-8 argument handling and avoid restoring the Objective-C bridge implicated in the 1.2.0/1.2.1 crash report.

## Public macro reference

| Macro | Purpose |
| --- | --- |
| `Timer_Insert` | Create a timer with duration and sound prompts; configure duration if one exists |
| `Timer_Configure` | Exact duration in editing view; next preset during the show |
| `Timer_ChooseSound` | Sound chooser and preview in editing view |
| `Timer_Toggle` | Start, pause or resume from play/pause and digits |
| `Timer_Start` | Idempotent start/resume for a custom shape action |
| `Timer_Reset` | Restore configured duration and wait for play |
| `Timer_AddMinute` / `Timer_SubtractMinute` | Adjust duration and remaining time |
| `Timer_ToggleSound` | Mute/unmute this timer |
| `Timer_TestSound` | Test the selected timer’s sound; respect mute |
| `Timer_Remove` | Remove the current timer |
| `Timer_RefreshStyle` | Rebuild while keeping duration, sound, position and width |
| `Timer_Initialize` | Optional one-time upgrade cleanup of obsolete click-start effects/shapes |
| `Timer_Present` | Optional configured-show shortcut; ordinary show commands also work |
| `Timer_CheckCurrent` | Validate shape components and report the last background error |
| `Timer_RunSelfTests` | Duration parsing/formatting, state transitions, native clock and sound-choice checks |
| `Timer_Shutdown` | Pause timers, cancel alerts/callbacks and clean up the child player |
| `Timer_Help` | Short usage dialog |

## Lifecycle and limits

Run **Timer_Shutdown before editing/removing code, using breakpoints, or pressing VBA Reset/End**. A scheduled native callback points into VBA code. The 1.2.0/1.2.1 exception is to restart PowerPoint instead of invoking that old sound/shutdown code; see [UPDATING.md](UPDATING.md).

Normal closing performs cleanup. If closing is canceled, the countdown remains paused. Closing/reopening the presentation does not restore a running deadline. PowerPoint can save shape appearances changed during the show; RESET before saving restores the full-duration appearance.

Use one show at a time and one timer per slide. Multiple timer slides can keep independent state. Modal operations or an unresponsive app can delay updates/alerts; the clock catches up afterward. Do not let the Mac sleep during a countdown. Set timer-slide transitions to None, keep controls grouped and avoid animating the whole timer to disappear.

API references: [Apple run-loop timers](https://developer.apple.com/library/archive/documentation/Cocoa/Conceptual/Multithreading/RunLoopManagement/RunLoopManagement.html), [Apple subprocess lifecycle](https://developer.apple.com/documentation/appkit/managing-ongoing-background-processes-in-your-mac), [PowerPoint PresentationBeforeClose](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.presentationbeforeclose).
