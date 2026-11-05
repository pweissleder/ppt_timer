#!/usr/bin/env python3
"""Build/check the self-contained setup ZIP without third-party dependencies."""

import argparse
import io
import re
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "downloads/SlideTimer-Mac.zip"
PACKAGE_ROOT = "SlideTimer-Mac"
DOWNLOAD_URL = "https://github.com/pweissleder/ppt_timer/raw/refs/heads/main/downloads/SlideTimer-Mac.zip"
FILES = (
    "README.md",
    "START-HERE.html",
    "LICENSE",
    "powerpoint/SlideTimer.bas",
    "powerpoint/SlideTimerEvents.cls",
    "powerpoint/paste/SlideTimer-Paste.txt",
    "powerpoint/paste/SlideTimerEvents-Paste.txt",
    "assets/timer-preview.png",
    "assets/setup/01-save-as-pptm.png",
    "assets/setup/02-open-visual-basic-editor.png",
    "assets/setup/03-import-file-menu.png",
    "assets/setup/04-import-slide-timer.png",
    "assets/setup/05-import-events.png",
    "assets/setup/06-modules-ready.png",
    "assets/setup/07-view-macros.png",
    "assets/setup/08-upgrade-cleanup.png",
    "assets/setup/09-run-timer-insert.png",
    "assets/setup/10-set-duration.png",
    "assets/setup/11-choose-sound.png",
    "assets/setup/12-preview-sound.png",
    "assets/setup/13-resize-timer.png",
    "docs/SETUP.md",
    "docs/USING.md",
    "docs/TROUBLESHOOTING.md",
    "docs/UPDATING.md",
    "docs/DEVELOPMENT.md",
    "docs/TESTING.md",
    "docs/CHANGELOG.md",
    "tools/build_setup_package.py",
    "tools/check_native_runtime.py",
)


def package_bytes():
    source = (ROOT / "powerpoint/SlideTimer.bas").read_text()
    events = (ROOT / "powerpoint/SlideTimerEvents.cls").read_text()
    if source.split("\n", 1)[1] != (ROOT / "powerpoint/paste/SlideTimer-Paste.txt").read_text():
        raise ValueError("The standard module and paste version differ.")
    if events[events.index("Option Explicit"):] != (ROOT / "powerpoint/paste/SlideTimerEvents-Paste.txt").read_text():
        raise ValueError("The class module and paste version differ.")
    version = re.search(r'PT_VERSION As String = "([^"]+)"', source).group(1)
    payloads = {}
    for relative in FILES:
        path = ROOT / relative
        data = path.read_bytes()
        if relative.endswith(".md"):
            # A package must not try to include/download a copy of itself.
            text = data.decode("utf-8")
            text = re.sub(r"(?:\.\./)?downloads/SlideTimer-Mac\.zip\?raw=true", DOWNLOAD_URL, text)
            data = text.encode("utf-8")
        elif relative == "START-HERE.html":
            text = data.decode("utf-8")
            text = text.replace('href="downloads/SlideTimer-Mac.zip" download', 'href="#setup"')
            text = text.replace(">Download setup ZIP</a>", ">Open setup steps</a>")
            text = text.replace("Download and extract the setup ZIP. You need desktop PowerPoint",
                                "You already have the setup files in this folder. You need desktop PowerPoint")
            data = text.encode("utf-8")
        payloads[relative] = data

    payloads["README.txt"] = (
        f"SLIDE TIMER {version} — POWERPOINT FOR MAC\n\n"
        "Start by opening START-HERE.html in this folder. The screenshot guide and preview work offline.\n\n"
        "The powerpoint folder contains BOTH files to import into your presentation:\n"
        "  SlideTimer.bas\n  SlideTimerEvents.cls\n\n"
        "Save your presentation as .pptm. Add both files once per presentation,\n"
        "then run Timer_Insert to choose a duration and sound. Start the show normally\n"
        "and click the timer's play/pause button or digits.\n\n"
        "Workshop tip: duplicate your Action slide and add the timer to the copy.\n"
        "Resize it as you like by selecting the whole group and dragging a corner.\n\n"
        "The powerpoint/paste folder has the alternative code for copy-and-paste setup.\n"
        "The docs folder contains setup, usage, troubleshooting and update guides.\n"
    ).encode("utf-8")

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for relative, data in sorted(payloads.items()):
            info = zipfile.ZipInfo(f"{PACKAGE_ROOT}/{relative}", (2026, 1, 1, 0, 0, 0))
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
    return buffer.getvalue(), len(payloads)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="Fail if the committed/downloadable ZIP is stale.")
    args = parser.parse_args()
    data, count = package_bytes()
    if args.check:
        if not OUTPUT.is_file() or OUTPUT.read_bytes() != data:
            raise SystemExit("Setup package is missing or stale. Run python3 tools/build_setup_package.py")
        print(f"Setup package is current: {count} files.")
    else:
        OUTPUT.parent.mkdir(parents=True, exist_ok=True)
        temporary = OUTPUT.with_suffix(".zip.tmp")
        temporary.write_bytes(data)
        temporary.replace(OUTPUT)
        print(f"Built {OUTPUT.relative_to(ROOT)}: {count} files, {len(data):,} bytes.")


if __name__ == "__main__":
    main()
