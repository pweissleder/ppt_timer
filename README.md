# Slide Timer for PowerPoint on Mac

Put a countdown directly on your slide. Click its play/pause button to start, choose an end sound, and make it as large or small as your layout needs. The timer keeps counting when you move to another slide or briefly leave the show.

![The timer with a five-minute countdown and its controls](assets/timer-preview.png)

**For desktop PowerPoint on Mac.** Save your presentation as **.pptm** so it keeps the timer. No add-in or internet connection is needed while presenting.

## Start here

Follow the [setup walkthrough with real PowerPoint screenshots](docs/SETUP.md) if you prefer to see each click.

**[Download the setup ZIP](downloads/SlideTimer-Mac.zip?raw=true)**, extract it, and open **START-HERE.html** for the setup guide and an interactive preview.

Prefer to follow the steps here?

1. **Save your presentation as `.pptm`.** In PowerPoint, choose **File → Save As → PowerPoint Macro-Enabled Presentation**.
2. **Add the two timer files once per presentation.** Open **Tools → Macro → Visual Basic Editor**, select your presentation on the left, right-click its project entry and choose **Import File…** to import **both files** from the package’s **powerpoint** folder: `SlideTimer.bas` and `SlideTimerEvents.cls`. You do not need to edit their contents.
3. **Choose the slide for the timer.** Back in PowerPoint’s normal editing view, select that slide. See the Action-slide tip below.
4. **Create your timer.** Open **Tools → Macro → Macros…**, choose **`Timer_Insert`**, then click **Run**. Enter a duration, such as `3` for three minutes or `1:30` for ninety seconds. Next, choose an end sound; `0` means no sound.
5. **Position and resize it.** Select the whole timer and drag it into place. Drag a corner handle to make it larger or smaller. Keep its parts grouped together.
6. **Present normally.** Start the slide show and click the timer’s **play/pause button or digits**. Click again to pause or resume. Save your file.

Allow macros for this presentation if PowerPoint asks. If importing is unavailable, follow the [copy-and-paste setup](docs/SETUP.md#if-import-file-is-unavailable).

## A useful workshop setup

**Add the timer to a copy of your Action slide.** Keep the original slide for explaining the activity. Duplicate it, add the timer to the copy, then advance to that copy when the activity begins. Participants can still read the instructions while the countdown runs.

**Resize the timer however you like to fit the slide.** A small timer beneath the instructions works well; a larger one is easier to see from the back of the room. Move or resize it as one group so its buttons and digits stay together. Set the timer slide’s transition to **None**.

## The controls

| Control | Use it to… |
| --- | --- |
| Play/pause or the digits | Start, pause or resume |
| SET | Choose the next preset duration |
| + / − | Add or subtract one minute |
| RESET | Return to the chosen duration and wait for play |
| SOUND ON / OFF | Enable or mute the end sound |

To change the exact duration later, run **`Timer_Configure`** in editing view. To change or preview the end sound, run **`Timer_ChooseSound`**. Type `p 2` in the sound chooser to preview Glass, then `2` to keep it.

The end sound follows **System Settings → Sound → Output** on your Mac. Try a short countdown with your actual speakers or display before presenting. A running timer continues across slides, app switches and a brief exit from the show; pause it if you want it to stop counting. Closing the presentation ends that timer session.

## Help when you need it

- [Setup with screenshots, including the paste alternative](docs/SETUP.md)
- [Using, resizing and reusing timers](docs/USING.md)
- [Troubleshooting](docs/TROUBLESHOOTING.md)
- [Updating an existing presentation](docs/UPDATING.md)

Updating from **1.2.0–1.2.3**? After replacing both files, run **`Timer_Initialize` once in editing view** and save. This removes the old hidden click-start animations. After that, start shows normally.

## For contributors

The import files are in [powerpoint/](powerpoint/), user guides in [docs/](docs/), and maintenance scripts in [tools/](tools/). See the [developer guide](docs/DEVELOPMENT.md), [PowerPoint test checklist](docs/TESTING.md), and [changelog](docs/CHANGELOG.md).

[MIT license](LICENSE) · Pascal Weissleder
