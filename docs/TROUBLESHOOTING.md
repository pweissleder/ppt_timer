# Troubleshooting

[← Back to the overview](../README.md)

## I cannot find Import File

Use the [copy-and-paste setup](SETUP.md#if-import-file-is-unavailable). You need one normal Module and one Class Module, each with its matching text file.

## Clicking play does nothing

Check these in order:

1. Start **Slide Show** view. The timer’s buttons run during the show, not while editing.
2. Allow macros for this presentation if PowerPoint asks.
3. Make sure the presentation is saved as **.pptm**.
4. Check that **both timer files** were imported into this presentation. Copying a timer shape from another deck does not copy its code.

In Presenter View, the clickable timer is on the audience slide display. The remote and spacebar are for normal slide navigation, not for starting the timer.

## There is no sound, or it comes from the wrong speakers

1. Check that the timer says **SOUND ON**.
2. On the Mac, open **System Settings → Sound → Output**. Select the device you want and check media volume.
3. In PowerPoint editing view, select that timer’s slide and run **Timer_TestSound** from **Tools → Macro → Macros…**.
4. If it still fails, run **Timer_CheckCurrent** and note the error message.

The separate macOS sound-effects output setting is not used by this version. A player-launch or playback error may mean a disconnected device, unreadable sound file, or a policy that blocks playback. The countdown can continue even if sound fails.

## I want a different duration or sound

Select the timer’s slide in editing view and run **Timer_Configure** for the duration or **Timer_ChooseSound** for the sound. To add a new timer, use **Timer_Insert**.

## An old timer slide still needs an extra click

This can happen when upgrading from 1.2.0–1.2.3. Exit the show, run **Timer_Initialize once in editing view**, then save. It removes the old hidden start shapes and animation steps. After that, start shows normally and click the timer’s own play/pause button or digits.

## PowerPoint says “User-defined type not defined” at SlideTimerEvents

The second file is missing, or its module has the wrong type/name. Import **SlideTimerEvents.cls** as well as **SlideTimer.bas**. For the paste route, it must be a **Class Module** named exactly **SlideTimerEvents**.

## PowerPoint reports a duplicate name

The code may have been imported or pasted twice. Use the [update guide](UPDATING.md) to replace the existing pair of modules with one current pair. Keep a backup of your presentation first.

## The timer is broken after I ungrouped it

In editing view, select the slide and run **Timer_RefreshStyle**. This rebuilds the timer while keeping its configured duration, sound, position and width. You can move or resize the new timer as one group.

If PowerPoint reports a missing timer component, this is the same repair to try.

## The timer refresh replays a transition

Set that slide’s transition to **None**. The timer redraws its visible slide as the countdown changes.

## My timer opens showing PAUSED or TIME’S UP

PowerPoint can save appearance changes made during a show. Start the show, press RESET, then save if you want the slide to open showing the full configured duration. Reopening the deck starts a new timer session; it does not resume a countdown from the previous session.

## The macros disappeared after saving

Saving as **.pptx** removes the code. Save as **.pptm**, then add both timer files again using the [setup guide](SETUP.md).

## I used 1.2.0 or 1.2.1 and PowerPoint crashed when playing sound

Quit and reopen PowerPoint before replacing the code. Do not run that old version’s sound test or shutdown macro. The current version removed the Objective-C sound bridge that was implicated in the crash report. Follow the [update guide](UPDATING.md) to replace both files.

## I need to edit or replace the timer code

For 1.2.2 or later, run **Timer_Shutdown** before editing/removing either module or using the VBA editor’s Reset/End button or breakpoints. This stops background updates safely. The exception for 1.2.0/1.2.1 is described above.

Normal presentation closing performs cleanup automatically. If you are maintaining the code, see the [developer guide](DEVELOPMENT.md).
