#!/usr/bin/env python3
"""Exercise the macOS APIs used by the VBA module (no PowerPoint automation).

Checks isolated afplay playback at zero volume, process cleanup and native
main-thread run-loop delivery. Use --play to hear one Ping over the current
media output. No Objective-C runtime calls are made. This cannot compile VBA or
verify PowerPoint's animation/event implementation.
"""

import ctypes as c
import re
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def bind(library, name, result, *arguments):
    function = getattr(library, name)
    function.restype = result
    function.argtypes = list(arguments)
    return function


def main():
    source = (ROOT / "powerpoint/SlideTimer.bas").read_text()
    assert source.split("\n", 1)[1] == (ROOT / "powerpoint/paste/SlideTimer-Paste.txt").read_text()
    events = (ROOT / "powerpoint/SlideTimerEvents.cls").read_text()
    assert events[events.index("Option Explicit"):] == (ROOT / "powerpoint/paste/SlideTimerEvents-Paste.txt").read_text()
    print("Import and paste distributions match.")

    if sys.platform != "darwin":
        raise SystemExit("Native checks require macOS.")
    catalog_source = source[source.index("Private Function SoundNames()") :]
    catalog_source = catalog_source[:catalog_source.index("End Function")]
    names = re.findall(r'"([A-Za-z]+)"', catalog_source)
    assert len(names) == 14 and len(set(names)) == len(names), "Invalid chooser catalog"
    for name in names:
        sound_path = Path("/System/Library/Sounds") / (name + ".aiff")
        assert sound_path.is_file(), f"Missing chooser sound: {name}"
        result = subprocess.run(["/usr/bin/afinfo", str(sound_path)], capture_output=True,
                                text=True, timeout=10)
        assert result.returncode == 0 and "estimated duration:" in result.stdout, (
            f"Cannot decode chooser sound {name}: {result.stderr}"
        )
    print(f"All {len(names)} chooser sounds exist and decode; no audio played by catalog checks.")
    if "--catalog-only" in sys.argv:
        return
    cf = c.CDLL("/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation")
    quartz = c.CDLL("/System/Library/Frameworks/QuartzCore.framework/QuartzCore")
    system = c.CDLL("/usr/lib/libSystem.B.dylib")
    pointer = c.c_void_p
    create_string = bind(cf, "CFStringCreateWithCString", pointer, pointer, c.c_char_p, c.c_uint32)
    release = bind(cf, "CFRelease", None, pointer)
    absolute_time = bind(cf, "CFAbsoluteTimeGetCurrent", c.c_double)
    monotonic = bind(quartz, "CACurrentMediaTime", c.c_double)
    main_loop = bind(cf, "CFRunLoopGetMain", pointer)
    callback_type = c.CFUNCTYPE(None, pointer, pointer)
    create_timer = bind(cf, "CFRunLoopTimerCreate", pointer, pointer, c.c_double,
                        c.c_double, c.c_ulong, c.c_long, callback_type, pointer)
    add_timer = bind(cf, "CFRunLoopAddTimer", None, pointer, pointer, pointer)
    invalidate = bind(cf, "CFRunLoopTimerInvalidate", None, pointer)
    run_loop = bind(cf, "CFRunLoopRunInMode", c.c_int32, pointer, c.c_double, c.c_ubyte)
    # Mirror VBA's dlsym + memmove dereference of the exported sentinel.
    find_symbol = bind(system, "dlsym", pointer, pointer, c.c_char_p)
    symbol = find_symbol(pointer(-2), b"kCFRunLoopCommonModes")
    assert symbol
    mode_pointer = pointer()
    c.memmove(c.byref(mode_pointer), symbol, c.sizeof(mode_pointer))
    common_mode = mode_pointer.value
    assert common_mode == pointer.in_dll(cf, "kCFRunLoopCommonModes").value
    default_mode = create_string(None, b"kCFRunLoopDefaultMode", 0x08000100)
    assert common_mode and default_mode
    received = []
    main_thread = threading.get_ident()

    @callback_type
    def tick(timer, context):
        received.append((threading.get_ident(), monotonic()))

    timer = create_timer(None, absolute_time() + 0.02, 0.05, 0, 0, tick, None)
    assert timer
    try:
        add_timer(main_loop(), timer, common_mode)
        run_loop(default_mode, 0.23, False)
        assert len(received) >= 3, f"Only {len(received)} callbacks arrived"
        assert all(thread == main_thread for thread, _ in received)
        before = monotonic()
        # Delayed UI delivery must still see elapsed monotonic time.
        time.sleep(0.15)
        assert monotonic() - before >= 0.14
        invalidate(timer)
        stopped_count = len(received)
        run_loop(default_mode, 0.12, False)
        assert len(received) == stopped_count, "Callback fired after invalidation"
        print(f"Main-thread timer: {stopped_count} callbacks; cancellation passed.")

        # Idle completion calls StopScheduler from inside the callback itself.
        stopped_inside = []

        @callback_type
        def stop_inside_callback(firing_timer, context):
            invalidate(firing_timer)
            release(firing_timer)
            stopped_inside.append(True)

        self_stopping = create_timer(None, absolute_time() + 0.01, 0.02, 0, 0,
                                     stop_inside_callback, None)
        assert self_stopping
        add_timer(main_loop(), self_stopping, common_mode)
        run_loop(default_mode, 0.1, False)
        if not stopped_inside:
            invalidate(self_stopping)
            release(self_stopping)
            raise AssertionError("Self-stopping callback did not fire")
        assert len(stopped_inside) == 1, "Self-stopping timer fired again"
        print("Invalidation and release inside the callback passed.")
    finally:
        invalidate(timer)
        release(timer)
        release(default_mode)

    assert "objc_msgSend" not in source and "libobjc" not in source
    assert "NSSound" not in source, "Unsafe in-process sound bridge remains"
    player = re.search(r'PT_AUDIO_PLAYER As String = "([^"]+)"', source).group(1)
    sound_file = re.search(r'PT_ALERT_FILE As String = "([^"]+)"', source).group(1)
    spawn = bind(system, "posix_spawn", c.c_int, c.POINTER(c.c_int), pointer,
                 pointer, pointer, pointer, pointer)
    wait_pid = bind(system, "waitpid", c.c_int, c.c_int, c.POINTER(c.c_int), c.c_int)
    kill = bind(system, "kill", c.c_int, c.c_int, c.c_int)

    def start_audio(path, *, audible=False):
        arguments = [player] + ([] if audible else ["-v", "0"]) + [str(path)]
        buffers = [c.create_string_buffer(argument.encode("utf-8")) for argument in arguments]
        argv = (pointer * (len(buffers) + 1))()
        for index, buffer in enumerate(buffers):
            argv[index] = c.addressof(buffer)
        environment = (pointer * 1)()  # Match VBA's empty, terminated envp.
        pid = c.c_int()
        result = spawn(c.byref(pid), c.addressof(buffers[0]), None, None, argv, environment)
        assert result == 0 and pid.value > 0, f"posix_spawn failed: {result}"
        return pid.value

    def finish_audio(pid):
        deadline = time.monotonic() + 6
        status = c.c_int()
        while time.monotonic() < deadline:
            result = wait_pid(pid, c.byref(status), 1)
            if result == pid:
                return status.value
            assert result == 0, f"waitpid failed for {pid}"
            time.sleep(0.02)
        kill(pid, 15)
        wait_pid(pid, c.byref(status), 0)
        raise AssertionError("afplay did not finish before the test timeout")

    status = finish_audio(start_audio(sound_file, audible="--play" in sys.argv))
    assert status == 0, f"afplay failed to play Ping (status {status})"
    print("Ping playback completed in a separate afplay process" +
          ("; confirm the physical output by listening." if "--play" in sys.argv else " at zero volume."))
    with tempfile.TemporaryDirectory(prefix="slide-timer-audio-") as directory:
        unicode_path = Path(directory) / "Timer Héro-音-🔔 'quoted'.aiff"
        unicode_path.write_bytes(Path(sound_file).read_bytes())
        assert finish_audio(start_audio(unicode_path)) == 0, "Literal UTF-8 filename did not play"
        invalid_path = Path(directory) / "invalid.aiff"
        invalid_path.write_bytes(b"Not an audio file.")
        assert finish_audio(start_audio(invalid_path)) != 0, "Invalid audio unexpectedly played"
    print("UTF-8 filename and invalid-audio exit-status checks passed.")
    pid = start_audio(sound_file)
    assert kill(pid, 15) == 0, "Cannot terminate the owned afplay child"
    assert finish_audio(pid) & 0x7F == 15, "Child did not exit after SIGTERM"
    status = c.c_int()
    assert wait_pid(pid, c.byref(status), 1) == -1, "Child was not reaped"
    print("Audio-child termination and reaping passed; no Objective-C sound calls remain.")


if __name__ == "__main__":
    main()
