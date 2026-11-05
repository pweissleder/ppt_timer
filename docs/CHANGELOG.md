# Changelog

[← Back to the overview](../README.md)

- **Documentation refresh:** Organized import/paste files, user guides and developer notes; added a reproducible setup ZIP, Action-slide copy workflow and resize guidance. Added a real PowerPoint screenshot walkthrough for setup, sound previews and resizing, with upgrade cleanup shown separately. Runtime remains 1.2.4.
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
